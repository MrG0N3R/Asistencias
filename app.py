from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


REGULAR_ENTRY_LIMIT = time(8, 0)
REGULAR_EXIT_START = time(16, 0)
REGULAR_EXIT_LIMIT = time(17, 0)
FINISHED_GOODS_ENTRY = time(7, 0)
FINISHED_GOODS_EXIT = time(16, 30)
ARRIVAL_TOLERANCE_MINUTES = 10
ADMIN_SCHEDULED_HOURS = 9
ADMIN_MIN_WORK_HOURS = 8.7
PRODUCTION_MIN_WORK_HOURS = 8
MIN_PAIR_HOURS = 4
MAX_PAIR_HOURS = 14

PRODUCTION_SHIFTS = (
    {"name": "Produccion 06:00-14:00", "start": time(6, 0), "end": time(14, 0), "overnight": False},
    {"name": "Produccion 14:00-22:00", "start": time(14, 0), "end": time(22, 0), "overnight": False},
    {"name": "Produccion 22:00-06:00", "start": time(22, 0), "end": time(6, 0), "overnight": True},
)


@dataclass
class Mark:
    employee_id: str
    name: str
    department: str
    building: str
    timestamp: datetime


def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_timestamp(value) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time())
    if not value:
        return None
    text = clean(value)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def normalize_header(value: str) -> str:
    text = clean(value).upper().replace(".", "")
    return re.sub(r"[^A-Z0-9]+", "", text)


def find_header(ws) -> tuple[int, dict[str, int]]:
    required = {
        "NOTRAB": "employee_id",
        "NOMBRE": "name",
        "FECHA": "timestamp",
        "RELOJ": "clock",
        "DEPARTAMENTO": "department",
    }
    for row_idx, row in enumerate(ws.iter_rows(max_row=40, values_only=True), start=1):
        normalized = {normalize_header(str(v)): idx for idx, v in enumerate(row) if v is not None}
        if {"NOMBRE", "FECHA", "RELOJ", "DEPARTAMENTO"}.issubset(normalized):
            return row_idx, {target: normalized[source] for source, target in required.items() if source in normalized}
    raise ValueError("No se encontro el encabezado esperado: NO.TRAB, NOMBRE, FECHA, RELOJ, DEPARTAMENTO.")


def infer_building(clock: str) -> str:
    return "Oficinas" if "OFICINA" in clock.upper() else "Planta"


def normalized_department(department: str) -> str:
    normalized = department.upper().replace("Ó", "O")
    return re.sub(r"[^A-Z0-9]+", " ", normalized).strip()


def is_finished_goods(department: str) -> bool:
    normalized = normalized_department(department)
    return "ALMACEN" in normalized and "PRODUCTO TERMINADO" in normalized


def is_raw_materials(department: str) -> bool:
    normalized = normalized_department(department)
    return "ALMACEN" in normalized and "MATERIA PRIMA" in normalized


def is_production(department: str) -> bool:
    return "PRODUCCION" in normalized_department(department) or is_raw_materials(department)


def period_start(day: date) -> date:
    return day - timedelta(days=(day.weekday() - 3) % 7)


def period_label(day: date) -> str:
    start = period_start(day)
    end = start + timedelta(days=6)
    return f"{start:%Y-%m-%d} al {end:%Y-%m-%d}"


def combine(day: date, value: time) -> datetime:
    return datetime.combine(day, value)


def production_shift_for_entry(entry: datetime) -> dict:
    entry_time = entry.time()
    if time(5, 0) <= entry_time < time(13, 0):
        shift = PRODUCTION_SHIFTS[0]
        shift_day = entry.date()
    elif time(13, 0) <= entry_time < time(21, 0):
        shift = PRODUCTION_SHIFTS[1]
        shift_day = entry.date()
    elif entry_time >= time(21, 0):
        shift = PRODUCTION_SHIFTS[2]
        shift_day = entry.date()
    else:
        shift = PRODUCTION_SHIFTS[2]
        shift_day = entry.date() - timedelta(days=1)

    scheduled_start = combine(shift_day, shift["start"])
    scheduled_end = combine(shift_day, shift["end"])
    if shift["overnight"]:
        scheduled_end += timedelta(days=1)
    return {
        "shift": shift["name"],
        "operational_date": shift_day,
        "scheduled_start": scheduled_start,
        "scheduled_end": scheduled_end,
    }


def regular_shift_for_entry(entry: datetime) -> dict:
    operational_date = entry.date()
    return {
        "shift": "Administrativo 08:00-17:00",
        "operational_date": operational_date,
        "scheduled_start": combine(operational_date, REGULAR_ENTRY_LIMIT),
        "scheduled_end": combine(operational_date, REGULAR_EXIT_LIMIT),
    }


def finished_goods_shift_for_entry(entry: datetime) -> dict:
    operational_date = entry.date()
    return {
        "shift": "Almacen de producto terminado 07:00-16:30",
        "operational_date": operational_date,
        "scheduled_start": combine(operational_date, FINISHED_GOODS_ENTRY),
        "scheduled_end": combine(operational_date, FINISHED_GOODS_EXIT),
    }


def minutes_between(start: datetime | None, end: datetime | None) -> int:
    if not start or not end or end < start:
        return 0
    return int((end - start).total_seconds() // 60)


def can_pair(entry: datetime, exit_mark: datetime | None) -> bool:
    if not exit_mark:
        return False
    hours = (exit_mark - entry).total_seconds() / 3600
    return MIN_PAIR_HOURS <= hours <= MAX_PAIR_HOURS


def load_marks(file_bytes: bytes) -> list[Mark]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    marks: list[Mark] = []
    for ws in wb.worksheets:
        try:
            header_row, columns = find_header(ws)
        except ValueError:
            continue
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            employee_id = clean(row[columns["employee_id"]]) if columns.get("employee_id", -1) < len(row) else ""
            name = clean(row[columns["name"]]) if columns.get("name", -1) < len(row) else ""
            timestamp = parse_timestamp(row[columns["timestamp"]]) if columns.get("timestamp", -1) < len(row) else None
            clock = clean(row[columns["clock"]]) if columns.get("clock", -1) < len(row) else ""
            department = clean(row[columns["department"]]) if columns.get("department", -1) < len(row) else ""
            if employee_id and name and timestamp:
                marks.append(Mark(employee_id, name, department, infer_building(clock), timestamp))
    if not marks:
        raise ValueError("El archivo no contiene marcas validas con el formato esperado.")
    return marks


def build_work_sessions(marks: list[Mark]) -> list[dict]:
    by_employee: dict[str, list[Mark]] = defaultdict(list)
    for mark in marks:
        by_employee[mark.employee_id].append(mark)

    sessions: list[dict] = []
    for employee_marks in by_employee.values():
        employee_marks.sort(key=lambda mark: mark.timestamp)
        sample = employee_marks[0]
        production = is_production(sample.department)

        if production:
            index = 0
            while index < len(employee_marks):
                entry_mark = employee_marks[index]
                exit_mark = employee_marks[index + 1] if index + 1 < len(employee_marks) else None
                if exit_mark and can_pair(entry_mark.timestamp, exit_mark.timestamp):
                    sessions.append({"entry_mark": entry_mark, "exit_mark": exit_mark})
                    index += 2
                else:
                    sessions.append({"entry_mark": entry_mark, "exit_mark": None})
                    index += 1
            continue

        by_day: dict[date, list[Mark]] = defaultdict(list)
        for mark in employee_marks:
            by_day[mark.timestamp.date()].append(mark)
        for day_marks in by_day.values():
            day_marks.sort(key=lambda mark: mark.timestamp)
            sessions.append({
                "entry_mark": day_marks[0],
                "exit_mark": day_marks[-1] if len(day_marks) > 1 else None,
            })

    return sorted(sessions, key=lambda session: (session["entry_mark"].timestamp, session["entry_mark"].employee_id))


def analyze(file_bytes: bytes, file_name: str) -> dict:
    marks = load_marks(file_bytes)
    sessions = build_work_sessions(marks)
    attendance = []
    incidents = []
    employee_period_lates: Counter[tuple[str, str]] = Counter()

    for session in sessions:
        entry_mark: Mark = session["entry_mark"]
        exit_mark: Mark | None = session["exit_mark"]
        entry = entry_mark.timestamp
        exit_time = exit_mark.timestamp if exit_mark else None
        worked_minutes = minutes_between(entry, exit_time)
        worked_hours = round(worked_minutes / 60, 2)
        tagged_as_production = "PRODUCCION" in normalized_department(entry_mark.department)
        reclassified_as_administrative = tagged_as_production and worked_minutes > 8.5 * 60
        production = is_production(entry_mark.department) and not reclassified_as_administrative
        if is_finished_goods(entry_mark.department):
            shift_info = finished_goods_shift_for_entry(entry)
            schedule_type = "Almacen de producto terminado"
        elif production:
            shift_info = production_shift_for_entry(entry)
            schedule_type = "Produccion"
        else:
            shift_info = regular_shift_for_entry(entry)
            schedule_type = "Administrativo reclasificado" if reclassified_as_administrative else "Administrativo"
        operational_date = shift_info["operational_date"]
        period = period_label(operational_date)
        minimum_work_minutes = round(
            (PRODUCTION_MIN_WORK_HOURS if production else ADMIN_MIN_WORK_HOURS) * 60
        )
        minimum_work_hours = PRODUCTION_MIN_WORK_HOURS if production else ADMIN_MIN_WORK_HOURS
        late_limit = shift_info["scheduled_start"] + timedelta(minutes=ARRIVAL_TOLERANCE_MINUTES)

        # Calculo de minutos de retardo netos
        late_minutes = int((entry - late_limit).total_seconds() // 60) if entry > late_limit else 0
        late = late_minutes > 0

        missing_exit = exit_time is None
        short_shift = worked_minutes > 0 and worked_minutes < minimum_work_minutes
        exit_out_of_range = False
        if not production and exit_time:
            exit_out_of_range = (
                exit_time < shift_info["scheduled_end"]
                and worked_minutes < minimum_work_minutes
            )
        if production and exit_time:
            exit_out_of_range = exit_time < shift_info["scheduled_end"] and worked_minutes < minimum_work_minutes

        status = "Asistencia"
        if late:
            status = "Retardo"
        if missing_exit:
            status = "Salida faltante"
        elif short_shift:
            status = f"Jornada menor a {minimum_work_hours:g}h"

        record = {
            "employee_id": entry_mark.employee_id,
            "name": entry_mark.name,
            "department": entry_mark.department,
            "building": entry_mark.building,
            "date": operational_date.isoformat(),
            "entry_date": entry.date().isoformat(),
            "period": period,
            "shift": shift_info["shift"],
            "scheduled_entry": shift_info["scheduled_start"].strftime("%H:%M"),
            "entry_tolerance_limit": late_limit.strftime("%H:%M"),
            "scheduled_exit": shift_info["scheduled_end"].strftime("%H:%M"),
            "entry": entry.strftime("%H:%M"),
            "exit": exit_time.strftime("%H:%M") if exit_time else "",
            "marks": 2 if exit_time else 1,
            "worked_hours": worked_hours,
            "late_minutes": late_minutes,
            "production": production,
            "schedule_type": schedule_type,
            "reclassified_as_administrative": reclassified_as_administrative,
            "late": late,
            "short_shift": short_shift,
            "missing_exit": missing_exit,
            "exit_out_of_range": exit_out_of_range,
            "status": status,
        }
        attendance.append(record)

        if late:
            employee_period_lates[(entry_mark.employee_id, period)] += 1
            incidents.append({
                **record,
                "incident_type": "Retardo",
                "detail": f"Entrada a las {record['entry']} ({late_minutes} min tarde sobre tolerancia de {record['entry_tolerance_limit']})",
            })
        if missing_exit:
            incidents.append({**record, "incident_type": "Salida faltante", "detail": "Sin registro de checado de salida"})
        elif short_shift:
            incidents.append({
                **record,
                "incident_type": f"Jornada menor a {minimum_work_hours:g}h",
                "detail": f"Trabajó {worked_hours}h de {minimum_work_hours:g}h mínimas requeridas",
            })
        elif exit_out_of_range:
            incidents.append({**record, "incident_type": "Salida fuera de rango", "detail": "Registro de salida anticipado al horario asignado"})

    period_penalties = []
    late_lookup = dict(employee_period_lates)
    employee_meta = {r["employee_id"]: r for r in attendance}
    for (employee_id, period), late_count in sorted(late_lookup.items(), key=lambda item: (item[0][1], item[0][0])):
        if late_count >= 2:
            meta = employee_meta[employee_id]
            period_penalties.append({
                "employee_id": employee_id,
                "name": meta["name"],
                "department": meta["department"],
                "building": meta["building"],
                "period": period,
                "late_count": late_count,
                "equivalent_absence": 1,
                "bonus_lost": "Si",
            })

    for incident in incidents:
        late_count = late_lookup.get((incident["employee_id"], incident["period"]), 0)
        incident["period_late_count"] = late_count
        incident["bonus_lost"] = "Si" if late_count >= 2 else "No"

    periods = sorted({r["period"] for r in attendance})
    weekly = []
    for period in periods:
        rows = [r for r in attendance if r["period"] == period]
        weekly.append({
            "period": period,
            "attendance_days": len(rows),
            "late_count": sum(1 for r in rows if r["late"]),
            "short_shift_count": sum(1 for r in rows if r["short_shift"]),
            "penalties": sum(1 for p in period_penalties if p["period"] == period),
        })

    return {
        "file_name": file_name,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rules": {
            "arrival_tolerance": "Todo el personal tiene 10 minutos de tolerancia después del inicio de su turno.",
            "administrative": "Horario 08:00-17:00 (9 horas); se acepta desde 8.7 horas trabajadas sin incidencia.",
            "production": "Turnos 06:00-14:00, 14:00-22:00 y 22:00-06:00; mínimo de 8 horas.",
            "minimum_hours": "Administrativos: 8.7 horas válidas. Producción: 8 horas.",
            "finished_goods": "Almacén de producto terminado: horario fijo 07:00-16:30, con 10 minutos de tolerancia.",
            "raw_materials": "Almacén de materia prima comparte los turnos y reglas de Producción.",
            "production_reclassification": "Personal etiquetado como Producción con más de 8.5 horas trabajadas se evalúa como Administrativo.",
            "period": "Jueves a miércoles.",
            "penalty": "2 retardos en el mismo periodo = falta y pérdida de bono de puntualidad.",
        },
        "summary": {
            "employees": len({r["employee_id"] for r in attendance}),
            "attendance_days": len(attendance),
            "lates": sum(1 for r in attendance if r["late"]),
            "incidents": len(incidents),
            "penalties": len(period_penalties),
            "missing_exits": sum(1 for r in attendance if r["missing_exit"]),
            "short_shifts": sum(1 for r in attendance if r["short_shift"]),
            "buildings": dict(Counter(r["building"] for r in attendance)),
            "departments": dict(Counter(r["department"] for r in attendance)),
        },
        "weekly": weekly,
        "attendance": attendance,
        "incidents": incidents,
        "period_penalties": period_penalties,
    }


def set_widths(ws):
    for column in ws.columns:
        max_len = 0
        letter = get_column_letter(column[0].column)
        for cell in column:
            max_len = max(max_len, len(str(cell.value or "")))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 42)


def write_table(ws, headers: list[str], rows: list[dict], keys: list[str]):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5597")
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append([row.get(key, "") for key in keys])
    set_widths(ws)
    ws.freeze_panes = "A2"


def export_workbook(result: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws["A1"] = "Consolidado de asistencia"
    ws["A1"].font = Font(bold=True, size=16)
    ws.append(["Archivo", result["file_name"]])
    ws.append(["Generado", result["generated_at"]])
    ws.append([])
    for label, value in [
        ("Empleados", result["summary"]["employees"]),
        ("Días de asistencia", result["summary"]["attendance_days"]),
        ("Retardos", result["summary"]["lates"]),
        ("Incidencias", result["summary"]["incidents"]),
        ("Periodos con falta/bono perdido", result["summary"]["penalties"]),
        ("Salidas faltantes", result["summary"]["missing_exits"]),
        ("Jornadas menores a 8h", result["summary"]["short_shifts"]),
    ]:
        ws.append([label, value])
    ws.append([])
    ws.append(["Regla", "Valor"])
    for key, value in result["rules"].items():
        ws.append([key, value])
    set_widths(ws)

    attendance_keys = [
        "employee_id", "name", "department", "building", "date", "period", "shift",
        "scheduled_entry", "scheduled_exit", "entry", "exit", "worked_hours", "marks", "status",
    ]
    ws = wb.create_sheet("Asistencias")
    write_table(
        ws,
        ["No. Trab", "Nombre", "Departamento", "Edificio", "Fecha operativa", "Periodo", "Turno", "Entrada prog.", "Salida prog.", "Entrada", "Salida", "Horas", "Marcas", "Estatus"],
        result["attendance"],
        attendance_keys,
    )

    incident_keys = [
        "employee_id", "name", "department", "building", "date", "period", "shift",
        "scheduled_entry", "entry", "exit", "worked_hours", "late_minutes", "incident_type", "detail", "period_late_count", "bonus_lost",
    ]
    ws = wb.create_sheet("Incidencias")
    write_table(
        ws,
        ["No. Trab", "Nombre", "Departamento", "Edificio", "Fecha operativa", "Periodo", "Turno", "Entrada prog.", "Entrada", "Salida", "Horas", "Min Retardo", "Incidencia", "Detalle", "Retardos periodo", "Pierde bono"],
        result["incidents"],
        incident_keys,
    )

    ws = wb.create_sheet("Penalizaciones")
    write_table(
        ws,
        ["No. Trab", "Nombre", "Departamento", "Edificio", "Periodo", "Retardos", "Falta equivalente", "Pierde bono"],
        result["period_penalties"],
        ["employee_id", "name", "department", "building", "period", "late_count", "equivalent_absence", "bonus_lost"],
    )

    ws = wb.create_sheet("Resumen semanal")
    write_table(
        ws,
        ["Periodo", "Dias asistencia", "Retardos", "Jornadas <8h", "Penalizaciones"],
        result["weekly"],
        ["period", "attendance_days", "late_count", "short_shift_count", "penalties"],
    )

    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()


def to_frame(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows)


def apply_filters(
    df: pd.DataFrame,
    building: str,
    department: str,
    period: str,
    search: str,
    start_date: date | None = None,
    end_date: date | None = None,
) -> pd.DataFrame:
    if df.empty:
        return df
    filtered = df.copy()
    if building != "Todos":
        filtered = filtered[filtered["building"] == building]
    if department != "Todos":
        filtered = filtered[filtered["department"] == department]
    if period != "Todos":
        filtered = filtered[filtered["period"] == period]
    
    if "date" in filtered.columns and (start_date or end_date):
        row_dates = pd.to_datetime(filtered["date"], errors="coerce").dt.date
        if start_date:
            filtered = filtered[row_dates >= start_date]
            row_dates = pd.to_datetime(filtered["date"], errors="coerce").dt.date
        if end_date:
            filtered = filtered[row_dates <= end_date]

    if search:
        needle = search.lower()
        filtered = filtered[
            filtered["name"].str.lower().str.contains(needle, na=False)
            | filtered["employee_id"].astype(str).str.lower().str.contains(needle, na=False)
        ]
    return filtered


def summarize_incidents(incidents: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "employee_id", "name", "department", "building", "incident_count",
        "incident_types", "dates", "total_lates", "bonus_lost",
    ]
    if incidents.empty:
        return pd.DataFrame(columns=columns)

    def unique_join(values) -> str:
        return ", ".join(dict.fromkeys(str(value) for value in values if pd.notna(value) and str(value)))

    summary = (
        incidents.sort_values(["name", "date"])
        .groupby(["employee_id", "name", "department", "building"], as_index=False, dropna=False)
        .agg(
            incident_count=("incident_type", "size"),
            incident_types=("incident_type", unique_join),
            dates=("date", unique_join),
            total_lates=("incident_type", lambda values: int((values == "Retardo").sum())),
            bonus_lost=("bonus_lost", lambda values: "Si" if (values == "Si").any() else "No"),
        )
    )
    return summary[columns]


def recalculate_period_penalties(incidents: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    penalty_columns = [
        "employee_id", "name", "department", "building", "period",
        "late_count", "equivalent_absence", "bonus_lost",
    ]
    if incidents.empty:
        return incidents.copy(), pd.DataFrame(columns=penalty_columns)

    adjusted = incidents.copy()
    late_counts = (
        adjusted[adjusted["incident_type"] == "Retardo"]
        .groupby(["employee_id", "period"])
        .size()
    )
    adjusted["period_late_count"] = [
        int(late_counts.get((employee_id, period), 0))
        for employee_id, period in zip(adjusted["employee_id"], adjusted["period"])
    ]
    adjusted["bonus_lost"] = adjusted["period_late_count"].ge(2).map({True: "Si", False: "No"})

    penalties = (
        adjusted[adjusted["period_late_count"] >= 2]
        .drop_duplicates(["employee_id", "period"])
        [["employee_id", "name", "department", "building", "period", "period_late_count"]]
        .rename(columns={"period_late_count": "late_count"})
    )
    penalties["equivalent_absence"] = 1
    penalties["bonus_lost"] = "Si"
    return adjusted, penalties[penalty_columns]


def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="Dashboard de Asistencia - UGRPG", layout="wide")
    st.title("Dashboard de Asistencia e Incidencias")
    st.caption("Sistema de análisis de asistencias, Checadores Planta y Oficinas.")

    uploaded_file = st.file_uploader("Archivo de asistencia (.xlsx)", type=["xlsx"])
    if not uploaded_file:
        st.info("Sube un archivo Excel de checador para analizar la información.")
        return

    try:
        result = analyze(uploaded_file.getvalue(), uploaded_file.name)
    except Exception as exc:
        st.error(str(exc))
        return

    attendance_df = to_frame(result["attendance"])
    incidents_df = to_frame(result["incidents"])
    available_dates = pd.to_datetime(attendance_df["date"], errors="coerce").dt.date.dropna()
    min_date = available_dates.min()
    max_date = available_dates.max()

    with st.sidebar:
        st.header("Filtros")
        selected_dates = st.date_input(
            "Rango de fechas a evaluar",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY",
        )
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        elif isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 1:
            start_date = end_date = selected_dates[0]
        else:
            start_date = end_date = min_date, max_date
            
        building = st.selectbox("Edificio", ["Todos"] + sorted(attendance_df["building"].dropna().unique().tolist()))
        department = st.selectbox("Departamento", ["Todos"] + sorted(attendance_df["department"].dropna().unique().tolist()))
        period = st.selectbox("Periodo", ["Todos"] + sorted(attendance_df["period"].dropna().unique().tolist()))
        search = st.text_input("Buscar empleado (ID o Nombre)")
        st.divider()
        st.download_button(
            "Descargar Consolidado (Excel)",
            data=export_workbook(result),
            file_name=f"consolidado_asistencia_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    filtered_attendance = apply_filters(
        attendance_df, building, department, period, search, start_date, end_date
    )
    filtered_incidents = apply_filters(
        incidents_df, building, department, period, search, start_date, end_date
    )
    filtered_incidents, penalties_df = recalculate_period_penalties(filtered_incidents)

    # Tarjetas Métricas principales
    kpi_cols = st.columns(5)
    kpi_cols[0].metric("Empleados", filtered_attendance["employee_id"].nunique() if not filtered_attendance.empty else 0)
    kpi_cols[1].metric("Días Asistencia", len(filtered_attendance))
    kpi_cols[2].metric("Total Retardos", int(filtered_attendance["late"].sum()) if not filtered_attendance.empty else 0)
    kpi_cols[3].metric("Incidencias Totales", len(filtered_incidents))
    kpi_cols[4].metric("Bonos Perdidos", len(penalties_df))

    # Pestañas para desglose detallado
    tab_overview, tab_lates, tab_short_shifts, tab_employee_lookup = st.tabs([
        "📊 Vista General", 
        "⏰ Detalle de Retardos", 
        "⚠️ Salidas Faltantes / Jornadas Incompletas", 
        "👤 Ficha por Empleado"
    ])

    with tab_overview:
        chart_left, chart_right = st.columns([1.2, 0.8])
        with chart_left:
            st.subheader("Incidencias por Periodo Jueves-Miércoles")
            if filtered_attendance.empty:
                st.warning("Sin datos para los filtros actuales.")
            else:
                weekly = (
                    filtered_attendance.groupby("period", as_index=False)
                    .agg(Asistencias=("employee_id", "count"), Retardos=("late", "sum"), Jornadas_Incompletas=("short_shift", "sum"))
                    .set_index("period")
                )
                st.bar_chart(weekly[["Retardos", "Jornadas_Incompletas"]])
        with chart_right:
            st.subheader("Distribución por Edificio")
            if not filtered_attendance.empty:
                st.bar_chart(filtered_attendance["building"].value_counts())

        st.subheader("Resumen Concentrado de Incidencias")
        incident_summary = summarize_incidents(filtered_incidents)
        st.dataframe(incident_summary, use_container_width=True, hide_index=True)

    with tab_lates:
        st.subheader("Desglose Detallado de Retardos")
        lates_only = filtered_incidents[filtered_incidents["incident_type"] == "Retardo"]
        if lates_only.empty:
            st.success("No hay retardos registrados para los filtros seleccionados.")
        else:
            late_cols = ["employee_id", "name", "department", "date", "period", "shift", "scheduled_entry", "entry", "late_minutes", "period_late_count", "bonus_lost"]
            st.dataframe(
                lates_only[late_cols].rename(columns={
                    "employee_id": "No. Trab",
                    "name": "Nombre",
                    "department": "Departamento",
                    "date": "Fecha",
                    "period": "Periodo",
                    "shift": "Turno",
                    "scheduled_entry": "Entrada Prog.",
                    "entry": "Entrada Real",
                    "late_minutes": "Minutos Tarde",
                    "period_late_count": "Retardos Acum. Periodo",
                    "bonus_lost": "Pierde Bono"
                }),
                use_container_width=True,
                hide_index=True
            )

    with tab_short_shifts:
        st.subheader("Incidencias de Jornada Incompleta o Registro Faltante")
        non_lates = filtered_incidents[filtered_incidents["incident_type"] != "Retardo"]
        if non_lates.empty:
            st.success("Sin incidencias de salida para este rango.")
        else:
            short_cols = ["employee_id", "name", "department", "date", "shift", "entry", "exit", "worked_hours", "incident_type", "detail"]
            st.dataframe(
                non_lates[short_cols].rename(columns={
                    "employee_id": "No. Trab",
                    "name": "Nombre",
                    "department": "Departamento",
                    "date": "Fecha",
                    "shift": "Turno",
                    "entry": "Entrada",
                    "exit": "Salida",
                    "worked_hours": "Horas Trab.",
                    "incident_type": "Incidencia",
                    "detail": "Detalle"
                }),
                use_container_width=True,
                hide_index=True
            )

    with tab_employee_lookup:
        st.subheader("Consulta Individual de Empleado")
        emp_list = sorted(filtered_attendance["name"].unique().tolist()) if not filtered_attendance.empty else []
        if not emp_list:
            st.info("Sin registros de empleados.")
        else:
            selected_emp = st.selectbox("Selecciona Empleado", emp_list)
            emp_attendance = filtered_attendance[filtered_attendance["name"] == selected_emp]
            emp_incidents = filtered_incidents[filtered_incidents["name"] == selected_emp]

            col1, col2, col3 = st.columns(3)
            col1.metric("Días Trabajados", len(emp_attendance))
            col2.metric("Retardos Encontrados", int(emp_attendance["late"].sum()))
            col3.metric("Penalización Bono", "SÍ" if (emp_incidents["bonus_lost"] == "Si").any() else "NO")

            st.markdown("**Historial Completo de Marcaciones**")
            att_cols = ["date", "period", "shift", "entry", "exit", "worked_hours", "status"]
            st.dataframe(emp_attendance[att_cols], use_container_width=True, hide_index=True)


if __name__ == "__main__":
    run_streamlit()

