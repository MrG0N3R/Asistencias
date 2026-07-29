from __future__ import annotations

import hashlib
import io
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
SECURITY_FULL_SHIFT_HOURS = 12
SECURITY_PLANT_2_MIN_HOURS = 10.5
MIN_PAIR_HOURS = 4
MAX_PAIR_HOURS = 14

SCHEDULE_ADMINISTRATIVE = "administrative"
SCHEDULE_PRODUCTION = "production"
SCHEDULE_FINISHED_GOODS = "finished_goods"
SCHEDULE_SECURITY_PLANT_1 = "security_plant_1_12x24"
SCHEDULE_SECURITY_PLANT_2 = "security_plant_2"
SCHEDULE_SECURITY_PLANT_MIXED = "security_plant_1_and_2"
SCHEDULE_SECURITY_CEDIS = "security_cedis_24x12"
SCHEDULE_LABELS = {
    SCHEDULE_ADMINISTRATIVE: "Administrativo 08:00-17:00",
    SCHEDULE_PRODUCTION: "Produccion (turnos rotativos)",
    SCHEDULE_FINISHED_GOODS: "Almacen de producto terminado 07:00-16:30",
    SCHEDULE_SECURITY_PLANT_1: "Vigilancia 1 Planta 12x24 (07:00-19:00 / 19:00-07:00)",
    SCHEDULE_SECURITY_PLANT_2: "Vigilancia 2 Planta (07:00-17:30)",
    SCHEDULE_SECURITY_PLANT_MIXED: "Vigilancia Planta rol mixto (Vigilancia 1 y 2)",
    SCHEDULE_SECURITY_CEDIS: "Vigilancia CEDIS 24x12 (07:00-19:00)",
}
SCHEDULE_EXCEPTIONS_FILE = Path(__file__).with_name("schedule_exceptions.json")

PRODUCTION_SHIFTS = (
    {"name": "Produccion 06:00-14:00", "start": time(6, 0), "end": time(14, 0), "overnight": False},
    {"name": "Produccion 14:00-22:00", "start": time(14, 0), "end": time(22, 0), "overnight": False},
    {"name": "Produccion 22:00-06:00", "start": time(22, 0), "end": time(6, 0), "overnight": True},
)

SPANISH_MONTHS = {
    "ENE": 1,
    "FEB": 2,
    "MAR": 3,
    "ABR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AGO": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DIC": 12,
}

HR_INCIDENT_LABELS = {
    "FALTAS INJUSTIFICADAS": "Falta injustificada",
    "ACCIDENTE DE TRAYECTO": "Accidente de trayecto",
    "ENF. GRAL./ACC. FUERA TRAB.": "Enfermedad general / accidente fuera de trabajo",
}

HR_DEPARTMENT_ALIASES = {
    "PRODUCCIONN": "PRODUCCION",
    "PRODUTO TERMINADO": "PRODUCTO TERMINADO",
}


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


def normalize_employee_id(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return clean(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = clean(value)
    numeric_text = re.fullmatch(r"(\d+)\.0+", text)
    return numeric_text.group(1) if numeric_text else text


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
        normalized = {normalize_header(v): idx for idx, v in enumerate(row) if v is not None}
        if set(required).issubset(normalized):
            return row_idx, {target: normalized[source] for source, target in required.items() if source in normalized}
    raise ValueError("No se encontro el encabezado esperado: NO.TRAB, NOMBRE, FECHA, RELOJ, DEPARTAMENTO.")


def infer_building(clock: str) -> str:
    normalized = clock.upper()
    if "CEDIS" in normalized:
        return "CEDIS"
    return "Oficinas" if "OFICINA" in normalized else "Planta"


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


def security_plant_1_shift_for_entry(entry: datetime) -> dict:
    entry_time = entry.time()
    if time(5, 0) <= entry_time < time(13, 0):
        shift_day = entry.date()
        scheduled_start = combine(shift_day, time(7, 0))
        scheduled_end = combine(shift_day, time(19, 0))
        shift_name = "Vigilancia 1 Planta 07:00-19:00"
    elif entry_time >= time(13, 0):
        shift_day = entry.date()
        scheduled_start = combine(shift_day, time(19, 0))
        scheduled_end = combine(shift_day + timedelta(days=1), time(7, 0))
        shift_name = "Vigilancia 1 Planta 19:00-07:00"
    else:
        shift_day = entry.date() - timedelta(days=1)
        scheduled_start = combine(shift_day, time(19, 0))
        scheduled_end = combine(entry.date(), time(7, 0))
        shift_name = "Vigilancia 1 Planta 19:00-07:00"
    return {
        "shift": shift_name,
        "operational_date": shift_day,
        "scheduled_start": scheduled_start,
        "scheduled_end": scheduled_end,
    }


def security_plant_2_shift_for_entry(entry: datetime) -> dict:
    operational_date = entry.date()
    return {
        "shift": "Vigilancia 2 Planta 07:00-17:30",
        "operational_date": operational_date,
        "scheduled_start": combine(operational_date, time(7, 0)),
        "scheduled_end": combine(operational_date, time(17, 30)),
    }


def security_cedis_shift_for_entry(entry: datetime) -> dict:
    operational_date = entry.date()
    return {
        "shift": "Vigilancia CEDIS 07:00-19:00",
        "operational_date": operational_date,
        "scheduled_start": combine(operational_date, time(7, 0)),
        "scheduled_end": combine(operational_date, time(19, 0)),
    }


def security_mixed_shift_for_session(entry: datetime, exit_time: datetime | None) -> dict:
    entry_time = entry.time()
    if entry_time >= time(13, 0) or entry_time < time(5, 0):
        shift_info = security_plant_1_shift_for_entry(entry)
        shift_info["resolved_schedule"] = SCHEDULE_SECURITY_PLANT_1
        return shift_info

    worked_hours = (
        (exit_time - entry).total_seconds() / 3600
        if exit_time and exit_time >= entry
        else None
    )
    if worked_hours is not None and abs(worked_hours - SECURITY_PLANT_2_MIN_HOURS) <= abs(
        worked_hours - SECURITY_FULL_SHIFT_HOURS
    ):
        shift_info = security_plant_2_shift_for_entry(entry)
        shift_info["resolved_schedule"] = SCHEDULE_SECURITY_PLANT_2
        return shift_info

    shift_info = security_plant_1_shift_for_entry(entry)
    shift_info["resolved_schedule"] = SCHEDULE_SECURITY_PLANT_1
    return shift_info


def minutes_between(start: datetime | None, end: datetime | None) -> int:
    if not start or not end or end < start:
        return 0
    return int((end - start).total_seconds() // 60)


def can_pair(entry: datetime, exit_mark: datetime | None) -> bool:
    if not exit_mark:
        return False
    hours = (exit_mark - entry).total_seconds() / 3600
    return MIN_PAIR_HOURS <= hours <= MAX_PAIR_HOURS


def normalize_schedule_exceptions(schedule_exceptions: dict[str, str] | None) -> dict[str, str]:
    if not schedule_exceptions:
        return {}
    valid_schedules = set(SCHEDULE_LABELS)
    normalized = {}
    for employee_id, schedule in schedule_exceptions.items():
        clean_id = normalize_employee_id(employee_id)
        if not clean_id:
            continue
        if schedule not in valid_schedules:
            raise ValueError(f"Horario de excepcion no reconocido para el empleado {clean_id}.")
        normalized[clean_id] = schedule
    return normalized


def resolve_schedule_selections(
    selected_schedule_groups: dict[str, list[str]],
) -> tuple[dict[str, str], set[str]]:
    employee_schedule_selections: dict[str, set[str]] = defaultdict(set)
    for schedule, selected_ids in selected_schedule_groups.items():
        for employee_id in selected_ids:
            employee_schedule_selections[normalize_employee_id(employee_id)].add(schedule)

    allowed_mixed_schedules = {
        SCHEDULE_SECURITY_PLANT_1,
        SCHEDULE_SECURITY_PLANT_2,
    }
    conflicts = {
        employee_id
        for employee_id, selected_schedules in employee_schedule_selections.items()
        if len(selected_schedules) > 1 and selected_schedules != allowed_mixed_schedules
    }
    resolved = {
        employee_id: (
            SCHEDULE_SECURITY_PLANT_MIXED
            if selected_schedules == allowed_mixed_schedules
            else next(iter(selected_schedules))
        )
        for employee_id, selected_schedules in employee_schedule_selections.items()
        if employee_id and employee_id not in conflicts
    }
    return resolved, conflicts


def load_saved_schedule_exceptions(path: Path = SCHEDULE_EXCEPTIONS_FILE) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"No se pudo leer el catalogo de excepciones: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("El catalogo de excepciones debe contener un objeto por numero de empleado.")
    return normalize_schedule_exceptions(data)


def save_schedule_exceptions(
    schedule_exceptions: dict[str, str],
    path: Path = SCHEDULE_EXCEPTIONS_FILE,
) -> None:
    normalized = normalize_schedule_exceptions(schedule_exceptions)
    temporary_path = path.with_suffix(f"{path.suffix}.tmp")
    temporary_path.write_text(
        json.dumps(normalized, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temporary_path.replace(path)


def load_marks(file_bytes: bytes) -> list[Mark]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    marks: list[Mark] = []
    for ws in wb.worksheets:
        try:
            header_row, columns = find_header(ws)
        except ValueError:
            continue
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            def row_value(key: str):
                index = columns.get(key)
                return row[index] if index is not None and 0 <= index < len(row) else None

            employee_id = normalize_employee_id(row_value("employee_id"))
            name = clean(row_value("name"))
            timestamp = parse_timestamp(row_value("timestamp"))
            clock = clean(row_value("clock"))
            department = clean(row_value("department"))
            if employee_id and name and timestamp:
                marks.append(Mark(employee_id, name, department, infer_building(clock), timestamp))
    if not marks:
        raise ValueError("El archivo no contiene marcas validas con el formato esperado.")
    return marks


def build_work_sessions(
    marks: list[Mark],
    schedule_exceptions: dict[str, str] | None = None,
) -> list[dict]:
    schedule_exceptions = normalize_schedule_exceptions(schedule_exceptions)
    by_employee: dict[str, list[Mark]] = defaultdict(list)
    for mark in marks:
        by_employee[mark.employee_id].append(mark)

    sessions: list[dict] = []
    for employee_marks in by_employee.values():
        employee_marks.sort(key=lambda mark: mark.timestamp)
        sample = employee_marks[0]
        forced_schedule = schedule_exceptions.get(sample.employee_id)
        sequential_pairing = (
            forced_schedule in {
                SCHEDULE_PRODUCTION,
                SCHEDULE_SECURITY_PLANT_1,
                SCHEDULE_SECURITY_PLANT_MIXED,
            }
            if forced_schedule
            else is_production(sample.department)
        )

        if forced_schedule in {SCHEDULE_SECURITY_PLANT_1, SCHEDULE_SECURITY_PLANT_MIXED}:
            target_hours = (
                (SECURITY_FULL_SHIFT_HOURS, SECURITY_PLANT_2_MIN_HOURS)
                if forced_schedule == SCHEDULE_SECURITY_PLANT_MIXED
                else (SECURITY_FULL_SHIFT_HOURS,)
            )
            index = 0
            while index < len(employee_marks):
                entry_mark = employee_marks[index]
                candidate_indexes = [
                    candidate_index
                    for candidate_index in range(index + 1, len(employee_marks))
                    if can_pair(entry_mark.timestamp, employee_marks[candidate_index].timestamp)
                ]
                if candidate_indexes:
                    exit_index = min(
                        candidate_indexes,
                        key=lambda candidate_index: (
                            min(
                                abs(
                                    (
                                        employee_marks[candidate_index].timestamp
                                        - entry_mark.timestamp
                                    ).total_seconds()
                                    / 3600
                                    - target
                                )
                                for target in target_hours
                            ),
                            -candidate_index,
                        ),
                    )
                    sessions.append({
                        "entry_mark": entry_mark,
                        "exit_mark": employee_marks[exit_index],
                    })
                    index = exit_index + 1
                else:
                    sessions.append({"entry_mark": entry_mark, "exit_mark": None})
                    index += 1
            continue

        if sequential_pairing:
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


def analyze(
    file_bytes: bytes,
    file_name: str,
    schedule_exceptions: dict[str, str] | None = None,
) -> dict:
    schedule_exceptions = normalize_schedule_exceptions(schedule_exceptions)
    marks = load_marks(file_bytes)
    sessions = build_work_sessions(marks, schedule_exceptions)
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
        forced_schedule = schedule_exceptions.get(entry_mark.employee_id)
        tagged_as_production = "PRODUCCION" in normalized_department(entry_mark.department)
        reclassified_as_administrative = (
            not forced_schedule and tagged_as_production and worked_minutes > 8.5 * 60
        )
        if forced_schedule == SCHEDULE_ADMINISTRATIVE:
            production = False
            shift_info = regular_shift_for_entry(entry)
            schedule_type = "Administrativo"
        elif forced_schedule == SCHEDULE_PRODUCTION:
            production = True
            shift_info = production_shift_for_entry(entry)
            schedule_type = "Produccion"
        elif forced_schedule == SCHEDULE_FINISHED_GOODS:
            production = False
            shift_info = finished_goods_shift_for_entry(entry)
            schedule_type = "Almacen de producto terminado"
        elif forced_schedule == SCHEDULE_SECURITY_PLANT_1:
            production = False
            shift_info = security_plant_1_shift_for_entry(entry)
            schedule_type = "Vigilancia 1 Planta 12x24"
        elif forced_schedule == SCHEDULE_SECURITY_PLANT_2:
            production = False
            shift_info = security_plant_2_shift_for_entry(entry)
            schedule_type = "Vigilancia 2 Planta"
        elif forced_schedule == SCHEDULE_SECURITY_PLANT_MIXED:
            production = False
            shift_info = security_mixed_shift_for_session(entry, exit_time)
            schedule_type = (
                "Vigilancia 2 Planta (rol mixto)"
                if shift_info["resolved_schedule"] == SCHEDULE_SECURITY_PLANT_2
                else "Vigilancia 1 Planta (rol mixto)"
            )
        elif forced_schedule == SCHEDULE_SECURITY_CEDIS:
            production = False
            shift_info = security_cedis_shift_for_entry(entry)
            schedule_type = "Vigilancia CEDIS 24x12"
        else:
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
        schedule_source = (
            "Excepcion por numero de empleado"
            if forced_schedule
            else ("Regla automatica por horas" if reclassified_as_administrative else "Departamento")
        )
        operational_date = shift_info["operational_date"]
        period = period_label(operational_date)
        if forced_schedule in {SCHEDULE_SECURITY_PLANT_1, SCHEDULE_SECURITY_CEDIS}:
            minimum_work_hours = SECURITY_FULL_SHIFT_HOURS
        elif forced_schedule == SCHEDULE_SECURITY_PLANT_2:
            minimum_work_hours = SECURITY_PLANT_2_MIN_HOURS
        elif forced_schedule == SCHEDULE_SECURITY_PLANT_MIXED:
            minimum_work_hours = (
                SECURITY_PLANT_2_MIN_HOURS
                if shift_info["resolved_schedule"] == SCHEDULE_SECURITY_PLANT_2
                else SECURITY_FULL_SHIFT_HOURS
            )
        else:
            minimum_work_hours = PRODUCTION_MIN_WORK_HOURS if production else ADMIN_MIN_WORK_HOURS
        minimum_work_minutes = round(minimum_work_hours * 60)
        late_limit = shift_info["scheduled_start"] + timedelta(minutes=ARRIVAL_TOLERANCE_MINUTES)

        late = entry > late_limit
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

        evaluated_building = "CEDIS" if forced_schedule == SCHEDULE_SECURITY_CEDIS else entry_mark.building
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
            "building": evaluated_building,
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
            "production": production,
            "schedule_type": schedule_type,
            "schedule_source": schedule_source,
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
                "detail": f"Entrada posterior a {record['entry_tolerance_limit']} (incluye 10 min de tolerancia)",
            })
        if missing_exit:
            incidents.append({**record, "incident_type": "Salida faltante", "detail": "No se encontro salida emparejada"})
        elif short_shift:
            incidents.append({
                **record,
                "incident_type": f"Jornada menor a {minimum_work_hours:g}h",
                "detail": f"No cumple el minimo válido de {minimum_work_hours:g} horas",
            })
        elif exit_out_of_range:
            incidents.append({**record, "incident_type": "Salida fuera de rango", "detail": "Salida fuera del rango esperado"})

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
            "finished_goods": "Almacen de producto terminado: horario fijo 07:00-16:30, con 10 minutos de tolerancia.",
            "raw_materials": "Almacen de materia prima comparte los turnos y reglas de Produccion.",
            "production_reclassification": "Personal etiquetado como Produccion con mas de 8.5 horas trabajadas se evalua como Administrativo.",
            "employee_exceptions": "La excepcion asignada por numero de empleado tiene prioridad absoluta sobre departamento y regla automatica.",
            "security_plant_1": "Vigilancia 1 Planta: rol 12x24, turnos 07:00-19:00 y 19:00-07:00; jornada minima de 12 horas.",
            "security_plant_2": "Vigilancia 2 Planta: horario 07:00-17:30; jornada minima de 10.5 horas.",
            "security_mixed": "El rol mixto Vigilancia 1/2 resuelve cada jornada por entrada y duracion: 19:00-07:00, 07:00-17:30 o 07:00-19:00.",
            "security_cedis": "Vigilancia CEDIS: horario 07:00-19:00; jornada minima de 12 horas.",
            "security_rotation_scope": "El reporte valida las jornadas marcadas; las 24 horas de descanso no generan ausencias sin un rol con fecha ancla.",
            "period": "Jueves a miercoles.",
            "penalty": "2 retardos en el mismo periodo = falta y perdida de bono de puntualidad.",
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
        "schedule_exceptions": [
            {
                "employee_id": employee_id,
                "name": next((mark.name for mark in marks if mark.employee_id == employee_id), ""),
                "department": next((mark.department for mark in marks if mark.employee_id == employee_id), ""),
                "assigned_schedule": SCHEDULE_LABELS[schedule],
            }
            for employee_id, schedule in sorted(schedule_exceptions.items())
        ],
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
        ("Dias de asistencia", result["summary"]["attendance_days"]),
        ("Retardos", result["summary"]["lates"]),
        ("Incidencias", result["summary"]["incidents"]),
        ("Periodos con falta/bono perdido", result["summary"]["penalties"]),
        ("Salidas faltantes", result["summary"]["missing_exits"]),
        ("Jornadas menores a 8h", result["summary"]["short_shifts"]),
        ("Excepciones de horario activas", len(result.get("schedule_exceptions", []))),
    ]:
        ws.append([label, value])
    ws.append([])
    ws.append(["Regla", "Valor"])
    for key, value in result["rules"].items():
        ws.append([key, value])
    set_widths(ws)

    attendance_keys = [
        "employee_id", "name", "department", "schedule_type", "schedule_source", "building", "date", "period", "shift",
        "scheduled_entry", "scheduled_exit", "entry", "exit", "worked_hours", "marks", "status",
    ]
    ws = wb.create_sheet("Asistencias")
    write_table(
        ws,
        ["No. Trab", "Nombre", "Departamento archivo", "Horario aplicado", "Origen horario", "Edificio", "Fecha operativa", "Periodo", "Turno", "Entrada prog.", "Salida prog.", "Entrada", "Salida", "Horas", "Marcas", "Estatus"],
        result["attendance"],
        attendance_keys,
    )

    ws = wb.create_sheet("Excepciones horario")
    write_table(
        ws,
        ["No. Trab", "Nombre", "Departamento archivo", "Horario asignado"],
        result.get("schedule_exceptions", []),
        ["employee_id", "name", "department", "assigned_schedule"],
    )

    incident_keys = [
        "employee_id", "name", "department", "schedule_type", "schedule_source", "building", "date", "period", "shift",
        "scheduled_entry", "entry", "exit", "worked_hours", "incident_type", "detail", "period_late_count", "bonus_lost",
    ]
    ws = wb.create_sheet("Incidencias")
    write_table(
        ws,
        ["No. Trab", "Nombre", "Departamento archivo", "Horario aplicado", "Origen horario", "Edificio", "Fecha operativa", "Periodo", "Turno", "Entrada prog.", "Entrada", "Salida", "Horas", "Incidencia", "Detalle", "Retardos periodo", "Pierde bono"],
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


def parse_spanish_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = clean(value)
    match = re.fullmatch(r"(\d{1,2})/([A-Za-zÁÉÍÓÚÑáéíóúñ]{3})/(\d{4})", text)
    if match:
        month = SPANISH_MONTHS.get(match.group(2).upper())
        if month:
            try:
                return date(int(match.group(3)), month, int(match.group(1)))
            except ValueError:
                return None
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def friendly_hr_incident_type(value: str) -> str:
    text = re.sub(r"\s+\((?:Aus|Inc)\)\s*$", "", clean(value), flags=re.IGNORECASE)
    normalized = re.sub(r"\s+", " ", text.upper()).strip()
    return HR_INCIDENT_LABELS.get(normalized, text.title())


def hr_incident_category(value: str) -> str:
    text = clean(value).upper()
    if "(AUS)" in text:
        return "Ausencia"
    if "(INC)" in text:
        return "Incapacidad"
    return "Otra incidencia"


def friendly_hr_department(value: str) -> str:
    department = clean(value).upper()
    return HR_DEPARTMENT_ALIASES.get(department, department)


def extract_hr_report_period(frames: dict[str, pd.DataFrame]) -> tuple[date | None, date | None]:
    pattern = re.compile(
        r"REPORTE DE INCIDENCIAS DEL\s+(\d{1,2}/[A-ZÁÉÍÓÚÑ]{3}/\d{4})\s+AL\s+"
        r"(\d{1,2}/[A-ZÁÉÍÓÚÑ]{3}/\d{4})",
        re.IGNORECASE,
    )
    for frame in frames.values():
        for value in frame.fillna("").astype(str).to_numpy().ravel():
            match = pattern.search(clean(value))
            if match:
                return parse_spanish_date(match.group(1)), parse_spanish_date(match.group(2))
    return None, None


def load_hr_incidents(file_bytes: bytes, file_name: str) -> dict:
    extension = Path(file_name).suffix.lower()
    if extension not in {".xls", ".xlsx"}:
        raise ValueError("El archivo debe estar en formato .xls o .xlsx.")

    engine = "xlrd" if extension == ".xls" else "openpyxl"
    try:
        frames = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=None,
            header=None,
            dtype=object,
            engine=engine,
        )
    except ImportError as exc:
        raise ValueError(
            "No está instalado el componente para leer archivos .xls. "
            "Instala las dependencias del proyecto e intenta de nuevo."
        ) from exc
    except Exception as exc:
        raise ValueError(f"No se pudo leer el reporte de incidencias: {exc}") from exc

    report_start, report_end = extract_hr_report_period(frames)
    incidents: list[dict] = []
    employees_seen: set[str] = set()

    for sheet_name, frame in frames.items():
        current_employee: dict | None = None
        for row in frame.itertuples(index=False, name=None):
            values = [None if pd.isna(value) else value for value in row]
            first = clean(values[0]) if values else ""
            second = clean(values[1]) if len(values) > 1 else ""

            if re.fullmatch(r"\d{1,10}(?:\.0+)?", first) and second:
                employee_id = normalize_employee_id(values[0])
                current_employee = {
                    "employee_id": employee_id,
                    "name": second.replace("�", "Ñ"),
                    "department": friendly_hr_department(values[5]) if len(values) > 5 else "",
                    "source_department": clean(values[5]) if len(values) > 5 else "",
                }
                employees_seen.add(employee_id)
                continue

            incident_date = parse_spanish_date(values[0]) if values else None
            if not current_employee or not incident_date or not second:
                continue
            if second.upper().startswith("TOTAL "):
                continue

            incidents.append(
                {
                    **current_employee,
                    "date": incident_date,
                    "category": hr_incident_category(second),
                    "incident_type": friendly_hr_incident_type(second),
                    "source_description": second,
                    "sheet": sheet_name,
                }
            )

    if not incidents:
        raise ValueError(
            "No se encontraron incidencias. Verifica que sea el reporte "
            "'Listado de incidencias por empleado' de CONTPAQi."
        )

    incidents_df = pd.DataFrame(incidents).sort_values(
        ["date", "department", "name"], ascending=[True, True, True]
    )
    if report_start is None:
        report_start = incidents_df["date"].min()
    if report_end is None:
        report_end = incidents_df["date"].max()

    return {
        "file_name": file_name,
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "report_start": report_start,
        "report_end": report_end,
        "incidents": incidents_df,
        "employees_in_report": len(employees_seen),
    }


def summarize_hr_incidents(incidents: pd.DataFrame) -> dict:
    if incidents.empty:
        return {
            "employees": 0,
            "incident_days": 0,
            "absences": 0,
            "disabilities": 0,
            "departments": 0,
        }
    return {
        "employees": int(incidents["employee_id"].nunique()),
        "incident_days": int(len(incidents)),
        "absences": int((incidents["category"] == "Ausencia").sum()),
        "disabilities": int((incidents["category"] == "Incapacidad").sum()),
        "departments": int(incidents["department"].replace("", pd.NA).nunique()),
    }


def hr_employee_summary(incidents: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "employee_id",
        "name",
        "department",
        "incident_days",
        "absences",
        "disabilities",
        "incident_types",
        "first_date",
        "last_date",
    ]
    if incidents.empty:
        return pd.DataFrame(columns=columns)

    def joined_unique(values) -> str:
        return ", ".join(dict.fromkeys(clean(value) for value in values if clean(value)))

    summary = (
        incidents.groupby(["employee_id", "name", "department"], as_index=False, dropna=False)
        .agg(
            incident_days=("date", "size"),
            absences=("category", lambda values: int((values == "Ausencia").sum())),
            disabilities=("category", lambda values: int((values == "Incapacidad").sum())),
            incident_types=("incident_type", joined_unique),
            first_date=("date", "min"),
            last_date=("date", "max"),
        )
        .sort_values(["incident_days", "name"], ascending=[False, True])
    )
    return summary[columns]


def hr_department_summary(incidents: pd.DataFrame) -> pd.DataFrame:
    columns = ["department", "employees", "incident_days", "absences", "disabilities"]
    if incidents.empty:
        return pd.DataFrame(columns=columns)
    summary = (
        incidents.assign(department=incidents["department"].replace("", "Sin departamento"))
        .groupby("department", as_index=False)
        .agg(
            employees=("employee_id", "nunique"),
            incident_days=("date", "size"),
            absences=("category", lambda values: int((values == "Ausencia").sum())),
            disabilities=("category", lambda values: int((values == "Incapacidad").sum())),
        )
        .sort_values(["incident_days", "department"], ascending=[False, True])
    )
    return summary[columns]


def style_excel_header(row) -> None:
    for cell in row:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_dataframe_sheet(ws, dataframe: pd.DataFrame, headers: list[str], keys: list[str]) -> None:
    ws.append(headers)
    style_excel_header(ws[1])
    ws.row_dimensions[1].height = 24
    for record in dataframe.to_dict("records"):
        ws.append([record.get(key, "") for key in keys])
    separator = Side(style="thin", color="D9E3EC")
    for row_number, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            if isinstance(cell.value, date):
                cell.number_format = "dd/mm/yyyy"
            cell.alignment = Alignment(
                horizontal="center" if isinstance(cell.value, (int, float, date)) else "left",
                vertical="center",
            )
            cell.border = Border(bottom=separator)
            if row_number % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")
    set_widths(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False


def export_hr_incidents_workbook(result: dict, incidents: pd.DataFrame) -> bytes:
    summary = summarize_hr_incidents(incidents)
    employee_summary = hr_employee_summary(incidents)
    department_summary = hr_department_summary(incidents)
    type_summary = (
        incidents.groupby(["category", "incident_type"], as_index=False)
        .size()
        .rename(columns={"size": "incident_days"})
        .sort_values("incident_days", ascending=False)
        if not incidents.empty
        else pd.DataFrame(columns=["category", "incident_type", "incident_days"])
    )

    wb = Workbook()
    dashboard = wb.active
    dashboard.title = "Dashboard"
    dashboard.sheet_view.showGridLines = False
    dashboard.merge_cells("A1:H2")
    dashboard["A1"] = "Dashboard de incidencias verificadas por RH"
    dashboard["A1"].font = Font(bold=True, size=18, color="FFFFFF")
    dashboard["A1"].fill = PatternFill("solid", fgColor="17365D")
    dashboard["A1"].alignment = Alignment(horizontal="left", vertical="center")

    period_start = result["report_start"]
    period_end = result["report_end"]
    period_text = (
        f"{period_start:%d/%m/%Y} al {period_end:%d/%m/%Y}"
        if period_start and period_end
        else "Periodo no identificado"
    )
    dashboard["A4"] = "Periodo"
    dashboard["B4"] = period_text
    dashboard["E4"] = "Archivo"
    dashboard["F4"] = result["file_name"]
    dashboard["A5"] = "Generado"
    dashboard["B5"] = result["generated_at"]
    for cell in ("A4", "E4", "A5"):
        dashboard[cell].font = Font(bold=True, color="17365D")

    kpis = [
        ("Personas afectadas", summary["employees"], "A7:B7", "A8:B8"),
        ("Días de incidencia", summary["incident_days"], "D7:E7", "D8:E8"),
        ("Ausencias", summary["absences"], "G7:H7", "G8:H8"),
        ("Incapacidades", summary["disabilities"], "A10:B10", "A11:B11"),
        ("Departamentos", summary["departments"], "D10:E10", "D11:E11"),
    ]
    for label, value, label_range, value_range in kpis:
        for row in dashboard[label_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="EAF0F6")
        for row in dashboard[value_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="EAF0F6")
        dashboard.merge_cells(label_range)
        dashboard.merge_cells(value_range)
        label_cell = dashboard[label_range.split(":")[0]]
        value_cell = dashboard[value_range.split(":")[0]]
        label_cell.value = label
        value_cell.value = value
        label_cell.font = Font(bold=True, color="5B6573")
        value_cell.font = Font(bold=True, size=16, color="17365D")
        label_cell.alignment = Alignment(horizontal="center")
        value_cell.alignment = Alignment(horizontal="center")

    dashboard["A14"] = "Incidencias por tipo"
    dashboard["D14"] = "Departamentos con más días"
    dashboard["G14"] = "Personas con más días"
    for cell in ("A14", "D14", "G14"):
        dashboard[cell].font = Font(bold=True, color="17365D")

    dashboard["A15"] = "Tipo"
    dashboard["B15"] = "Días"
    dashboard["D15"] = "Departamento"
    dashboard["E15"] = "Días"
    dashboard["G15"] = "Empleado"
    dashboard["H15"] = "Días"
    for cell in ("A15", "B15", "D15", "E15", "G15", "H15"):
        dashboard[cell].font = Font(bold=True, color="FFFFFF")
        dashboard[cell].fill = PatternFill("solid", fgColor="0F766E")

    for row_number, record in enumerate(type_summary.head(10).to_dict("records"), start=16):
        dashboard.cell(row_number, 1, record["incident_type"])
        dashboard.cell(row_number, 2, record["incident_days"])
    for row_number, record in enumerate(department_summary.head(10).to_dict("records"), start=16):
        dashboard.cell(row_number, 4, record["department"])
        dashboard.cell(row_number, 5, record["incident_days"])
    for row_number, record in enumerate(employee_summary.head(10).to_dict("records"), start=16):
        dashboard.cell(row_number, 7, record["name"])
        dashboard.cell(row_number, 8, record["incident_days"])

    type_rows = len(type_summary.head(10))
    if type_rows:
        chart = PieChart()
        chart.title = "Distribución por tipo"
        chart.add_data(
            Reference(dashboard, min_col=2, min_row=15, max_row=15 + type_rows),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(dashboard, min_col=1, min_row=16, max_row=15 + type_rows)
        )
        chart.legend = None
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showLeaderLines = True
        chart.height = 7
        chart.width = 12
        dashboard.add_chart(chart, "A28")

    department_rows = len(department_summary.head(10))
    if department_rows:
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = "Días por departamento"
        chart.y_axis.title = "Departamento"
        chart.x_axis.title = "Días"
        chart.add_data(
            Reference(dashboard, min_col=5, min_row=15, max_row=15 + department_rows),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(dashboard, min_col=4, min_row=16, max_row=15 + department_rows)
        )
        chart.legend = None
        chart.height = 7
        chart.width = 14
        dashboard.add_chart(chart, "E28")

    for column, width in {"A": 35, "B": 14, "C": 3, "D": 28, "E": 14, "F": 3, "G": 38, "H": 14}.items():
        dashboard.column_dimensions[column].width = width
    dashboard.freeze_panes = "A7"

    detail = wb.create_sheet("Detalle de incidencias")
    detail_df = incidents.sort_values(["date", "name"]).copy()
    write_dataframe_sheet(
        detail,
        detail_df,
        [
            "No. Trab",
            "Nombre",
            "Departamento homologado",
            "Departamento origen",
            "Fecha",
            "Categoría",
            "Tipo de incidencia",
        ],
        [
            "employee_id",
            "name",
            "department",
            "source_department",
            "date",
            "category",
            "incident_type",
        ],
    )

    employees = wb.create_sheet("Resumen por empleado")
    write_dataframe_sheet(
        employees,
        employee_summary,
        [
            "No. Trab",
            "Nombre",
            "Departamento",
            "Días de incidencia",
            "Ausencias",
            "Incapacidades",
            "Tipos de incidencia",
            "Primera fecha",
            "Última fecha",
        ],
        [
            "employee_id",
            "name",
            "department",
            "incident_days",
            "absences",
            "disabilities",
            "incident_types",
            "first_date",
            "last_date",
        ],
    )

    departments = wb.create_sheet("Resumen por departamento")
    write_dataframe_sheet(
        departments,
        department_summary,
        ["Departamento", "Personas", "Días de incidencia", "Ausencias", "Incapacidades"],
        ["department", "employees", "incident_days", "absences", "disabilities"],
    )

    types = wb.create_sheet("Resumen por tipo")
    write_dataframe_sheet(
        types,
        type_summary,
        ["Categoría", "Tipo de incidencia", "Días"],
        ["category", "incident_type", "incident_days"],
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
    """Condensa todas las incidencias de una persona en una sola fila."""
    columns = [
        "employee_id", "name", "department", "schedule_types", "schedule_sources", "building", "incident_count",
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
            schedule_types=("schedule_type", unique_join),
            schedule_sources=("schedule_source", unique_join),
            incident_types=("incident_type", unique_join),
            dates=("date", unique_join),
            total_lates=("incident_type", lambda values: int((values == "Retardo").sum())),
            bonus_lost=("bonus_lost", lambda values: "Si" if (values == "Si").any() else "No"),
        )
    )
    return summary[columns]


def recalculate_period_penalties(incidents: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Recalcula retardos y penalizaciones usando únicamente el rango filtrado."""
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


def render_attendance_view(st):
    st.title("Dashboard de asistencia")
    st.caption("Carga archivos .xlsx con el formato de checador para Planta y Oficinas.")

    uploaded_file = st.file_uploader("Archivo de asistencia (.xlsx)", type=["xlsx"])
    if not uploaded_file:
        st.info("Sube un archivo para calcular asistencias, retardos, incidencias y penalizaciones por periodo jueves a miercoles.")
        return

    file_bytes = uploaded_file.getvalue()
    try:
        preview_marks = load_marks(file_bytes)
    except Exception as exc:
        st.error(str(exc))
        return

    employee_meta = {}
    for mark in preview_marks:
        employee_meta.setdefault(
            mark.employee_id,
            {"name": mark.name, "department": mark.department},
        )
    employee_ids = sorted(
        employee_meta,
        key=lambda value: (not value.isdigit(), int(value) if value.isdigit() else value),
    )

    def employee_label(employee_id: str) -> str:
        meta = employee_meta[employee_id]
        return f"{employee_id} - {meta['name']} ({meta['department']})"

    try:
        saved_schedule_exceptions = load_saved_schedule_exceptions()
        saved_exceptions_error = ""
    except ValueError as exc:
        saved_schedule_exceptions = {}
        saved_exceptions_error = str(exc)

    file_widget_key = hashlib.sha256(file_bytes).hexdigest()[:12]
    with st.sidebar:
        st.header("Excepciones de horario")
        st.caption(
            "La asignacion por No. Trab es estricta y reemplaza el departamento reportado en el archivo."
        )
        st.caption(
            "Para un rol mixto, selecciona al mismo empleado en Vigilancia 1 Planta y Vigilancia 2 Planta."
        )
        if saved_exceptions_error:
            st.warning(saved_exceptions_error)
        administrative_ids = st.multiselect(
            "Administrativo 08:00-17:00",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_ADMINISTRATIVE
            ],
            format_func=employee_label,
            key=f"administrative_exceptions_{file_widget_key}",
        )
        production_ids = st.multiselect(
            "Produccion (turnos rotativos)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_PRODUCTION
            ],
            format_func=employee_label,
            key=f"production_exceptions_{file_widget_key}",
        )
        finished_goods_ids = st.multiselect(
            "Producto terminado 07:00-16:30",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_FINISHED_GOODS
            ],
            format_func=employee_label,
            key=f"finished_goods_exceptions_{file_widget_key}",
        )
        security_plant_1_ids = st.multiselect(
            "Vigilancia 1 Planta 12x24 (07-19 / 19-07)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id)
                in {SCHEDULE_SECURITY_PLANT_1, SCHEDULE_SECURITY_PLANT_MIXED}
            ],
            format_func=employee_label,
            key=f"security_plant_1_exceptions_{file_widget_key}",
        )
        security_plant_2_ids = st.multiselect(
            "Vigilancia 2 Planta (07:00-17:30)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id)
                in {SCHEDULE_SECURITY_PLANT_2, SCHEDULE_SECURITY_PLANT_MIXED}
            ],
            format_func=employee_label,
            key=f"security_plant_2_exceptions_{file_widget_key}",
        )
        security_cedis_ids = st.multiselect(
            "Vigilancia CEDIS 24x12 (07:00-19:00)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_SECURITY_CEDIS
            ],
            format_func=employee_label,
            key=f"security_cedis_exceptions_{file_widget_key}",
        )

    selected_schedule_groups = {
        SCHEDULE_ADMINISTRATIVE: administrative_ids,
        SCHEDULE_PRODUCTION: production_ids,
        SCHEDULE_FINISHED_GOODS: finished_goods_ids,
        SCHEDULE_SECURITY_PLANT_1: security_plant_1_ids,
        SCHEDULE_SECURITY_PLANT_2: security_plant_2_ids,
        SCHEDULE_SECURITY_CEDIS: security_cedis_ids,
    }
    schedule_exceptions, conflicts = resolve_schedule_selections(selected_schedule_groups)
    if conflicts:
        conflict_names = ", ".join(employee_label(employee_id) for employee_id in sorted(conflicts))
        st.error(f"Cada empleado solo puede tener un horario de excepcion. Revisa: {conflict_names}")
        return

    with st.sidebar:
        st.caption(
            f"{len(schedule_exceptions)} aplicada(s) a este archivo; "
            f"{len(saved_schedule_exceptions)} guardada(s) en el catalogo."
        )
        if st.button("Guardar excepciones para futuros archivos", use_container_width=True):
            merged_exceptions = {
                employee_id: schedule
                for employee_id, schedule in saved_schedule_exceptions.items()
                if employee_id not in employee_meta
            }
            merged_exceptions.update(schedule_exceptions)
            try:
                save_schedule_exceptions(merged_exceptions)
                st.success("Catalogo de excepciones guardado.")
            except OSError as exc:
                st.error(f"No se pudo guardar el catalogo de excepciones: {exc}")

    try:
        result = analyze(file_bytes, uploaded_file.name, schedule_exceptions)
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
            start_date = end_date = selected_dates
        building = st.selectbox("Edificio", ["Todos"] + sorted(attendance_df["building"].dropna().unique().tolist()))
        department = st.selectbox("Departamento", ["Todos"] + sorted(attendance_df["department"].dropna().unique().tolist()))
        period = st.selectbox("Periodo", ["Todos"] + sorted(attendance_df["period"].dropna().unique().tolist()))
        search = st.text_input("Buscar empleado")
        st.divider()
        st.download_button(
            "Descargar consolidado",
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

    kpi_cols = st.columns(5)
    kpi_cols[0].metric("Empleados", filtered_attendance["employee_id"].nunique() if not filtered_attendance.empty else 0)
    kpi_cols[1].metric("Dias asistencia", len(filtered_attendance))
    kpi_cols[2].metric("Retardos", int(filtered_attendance["late"].sum()) if not filtered_attendance.empty else 0)
    kpi_cols[3].metric("Incidencias", len(filtered_incidents))
    kpi_cols[4].metric("Bonos perdidos", len(penalties_df))

    st.subheader("Reglas aplicadas")
    st.dataframe(pd.DataFrame([{"Regla": key, "Valor": value} for key, value in result["rules"].items()]), use_container_width=True, hide_index=True)
    if result["schedule_exceptions"]:
        with st.expander(f"Excepciones de horario activas ({len(result['schedule_exceptions'])})"):
            st.dataframe(
                pd.DataFrame(result["schedule_exceptions"]),
                use_container_width=True,
                hide_index=True,
            )

    chart_left, chart_right = st.columns([1.2, 0.8])
    with chart_left:
        st.subheader("Resumen por periodo")
        if filtered_attendance.empty:
            st.warning("Sin datos para los filtros actuales.")
        else:
            weekly = (
                filtered_attendance.groupby("period", as_index=False)
                .agg(Asistencias=("employee_id", "count"), Retardos=("late", "sum"), Jornadas_incompletas=("short_shift", "sum"))
                .set_index("period")
            )
            st.bar_chart(weekly[["Retardos", "Jornadas_incompletas"]])
    with chart_right:
        st.subheader("Asistencias por edificio")
        if not filtered_attendance.empty:
            st.bar_chart(filtered_attendance["building"].value_counts())

    st.subheader("Departamentos con mas incidencias")
    if filtered_incidents.empty:
        st.success("Sin incidencias para los filtros actuales.")
    else:
        st.bar_chart(filtered_incidents["department"].value_counts().head(10))

    st.subheader("Incidencias por empleado")
    incident_summary = summarize_incidents(filtered_incidents)
    st.dataframe(incident_summary, use_container_width=True, hide_index=True)

    st.subheader("Asistencias del personal")
    attendance_columns = [
        "employee_id", "name", "department", "schedule_type", "schedule_source",
        "building", "date", "period", "shift",
        "scheduled_entry", "scheduled_exit", "entry", "exit", "worked_hours", "status",
    ]
    st.dataframe(filtered_attendance[attendance_columns] if not filtered_attendance.empty else filtered_attendance, use_container_width=True, hide_index=True)


def render_hr_incidents_view(st):
    st.title("Incidencias verificadas por RH")
    st.caption(
        "Convierte el reporte consolidado de CONTPAQi en una lectura clara por persona, "
        "tipo de incidencia y departamento."
    )

    uploaded_file = st.file_uploader(
        "Reporte de incidencias (.xls o .xlsx)",
        type=["xls", "xlsx"],
        key="hr_incidents_file",
        help="Usa el reporte 'Listado de incidencias por empleado' revisado por Recursos Humanos.",
    )
    if not uploaded_file:
        st.info(
            "Sube el consolidado para identificar cuántas personas fueron afectadas, "
            "cuántos días corresponden a ausencias o incapacidades y dónde se concentran."
        )
        guide_columns = st.columns(3)
        guide_columns[0].markdown("**1. Sube el reporte**\n\nSe aceptan archivos `.xls` y `.xlsx`.")
        guide_columns[1].markdown("**2. Explora el dashboard**\n\nFiltra por fecha, área, tipo o persona.")
        guide_columns[2].markdown("**3. Descarga el resultado**\n\nObtén un Excel con dashboard y tablas de detalle.")
        return

    try:
        result = load_hr_incidents(uploaded_file.getvalue(), uploaded_file.name)
    except Exception as exc:
        st.error(str(exc))
        return

    incidents = result["incidents"]
    min_date = incidents["date"].min()
    max_date = incidents["date"].max()

    with st.sidebar:
        st.header("Filtros de incidencias RH")
        selected_dates = st.date_input(
            "Rango de fechas",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY",
            key="hr_date_range",
        )
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        elif isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 1:
            start_date = end_date = selected_dates[0]
        else:
            start_date = end_date = selected_dates

        departments = sorted(
            value for value in incidents["department"].dropna().unique().tolist() if clean(value)
        )
        department = st.selectbox(
            "Departamento",
            ["Todos"] + departments,
            key="hr_department",
        )
        categories = st.multiselect(
            "Categoría",
            sorted(incidents["category"].unique().tolist()),
            default=sorted(incidents["category"].unique().tolist()),
            key="hr_categories",
        )
        incident_types = st.multiselect(
            "Tipo de incidencia",
            sorted(incidents["incident_type"].unique().tolist()),
            default=sorted(incidents["incident_type"].unique().tolist()),
            key="hr_incident_types",
        )
        search = st.text_input("Buscar empleado", key="hr_employee_search")

    filtered = incidents[
        incidents["date"].between(start_date, end_date)
        & incidents["category"].isin(categories)
        & incidents["incident_type"].isin(incident_types)
    ].copy()
    if department != "Todos":
        filtered = filtered[filtered["department"] == department]
    if search:
        needle = search.casefold()
        filtered = filtered[
            filtered["name"].str.casefold().str.contains(needle, na=False)
            | filtered["employee_id"].astype(str).str.casefold().str.contains(needle, na=False)
        ]

    summary = summarize_hr_incidents(filtered)
    period_text = f"{result['report_start']:%d/%m/%Y} al {result['report_end']:%d/%m/%Y}"
    header_left, header_right = st.columns([1.5, 0.5], vertical_alignment="center")
    with header_left:
        st.markdown(f"**Periodo del reporte:** {period_text}  \n**Archivo:** {result['file_name']}")
        normalized_departments = incidents[
            incidents["department"] != incidents["source_department"].str.upper()
        ][["source_department", "department"]].drop_duplicates()
        if not normalized_departments.empty:
            changes = ", ".join(
                f"{row.source_department} → {row.department}"
                for row in normalized_departments.itertuples(index=False)
            )
            st.caption(f"Departamentos homologados para el análisis: {changes}.")
    with header_right:
        st.download_button(
            "Descargar dashboard en Excel",
            data=export_hr_incidents_workbook(result, filtered),
            file_name=f"dashboard_incidencias_RH_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=filtered.empty,
        )

    kpi_columns = st.columns(5)
    kpi_columns[0].metric("Personas afectadas", summary["employees"])
    kpi_columns[1].metric("Días de incidencia", summary["incident_days"])
    kpi_columns[2].metric("Ausencias", summary["absences"])
    kpi_columns[3].metric("Incapacidades", summary["disabilities"])
    kpi_columns[4].metric("Departamentos", summary["departments"])

    if filtered.empty:
        st.warning("No hay incidencias que coincidan con los filtros seleccionados.")
        return

    employee_summary = hr_employee_summary(filtered)
    department_summary = hr_department_summary(filtered)
    type_summary = (
        filtered.groupby("incident_type")
        .size()
        .sort_values(ascending=False)
        .rename("Días")
    )

    top_type = type_summary.index[0]
    top_department = department_summary.iloc[0]
    top_employee = employee_summary.iloc[0]
    st.info(
        f"Lectura rápida: **{top_type}** es la incidencia con más días ({int(type_summary.iloc[0])}). "
        f"**{top_department['department']}** concentra {int(top_department['incident_days'])} días y "
        f"**{top_employee['name']}** es la persona con mayor duración ({int(top_employee['incident_days'])} días)."
    )

    chart_left, chart_right = st.columns([1, 1])
    with chart_left:
        st.subheader("Días por tipo de incidencia")
        st.bar_chart(type_summary, color="#0F766E")
    with chart_right:
        st.subheader("Días por departamento")
        department_chart = department_summary.set_index("department")["incident_days"].head(10)
        st.bar_chart(department_chart, color="#D97706")

    st.subheader("Evolución diaria")
    daily = (
        filtered.assign(
            Ausencias=(filtered["category"] == "Ausencia").astype(int),
            Incapacidades=(filtered["category"] == "Incapacidad").astype(int),
        )
        .groupby("date")[["Ausencias", "Incapacidades"]]
        .sum()
    )
    st.bar_chart(daily, color=["#D97706", "#0F766E"])

    employee_tab, department_tab, detail_tab = st.tabs(
        ["Por empleado", "Por departamento", "Detalle diario"]
    )
    with employee_tab:
        employee_display = employee_summary.rename(
            columns={
                "employee_id": "No. Trab",
                "name": "Nombre",
                "department": "Departamento",
                "incident_days": "Días de incidencia",
                "absences": "Ausencias",
                "disabilities": "Incapacidades",
                "incident_types": "Tipos de incidencia",
                "first_date": "Primera fecha",
                "last_date": "Última fecha",
            }
        )
        st.dataframe(employee_display, use_container_width=True, hide_index=True)
    with department_tab:
        department_display = department_summary.rename(
            columns={
                "department": "Departamento",
                "employees": "Personas",
                "incident_days": "Días de incidencia",
                "absences": "Ausencias",
                "disabilities": "Incapacidades",
            }
        )
        st.dataframe(department_display, use_container_width=True, hide_index=True)
    with detail_tab:
        detail_display = filtered[
            [
                "date",
                "employee_id",
                "name",
                "department",
                "source_department",
                "category",
                "incident_type",
            ]
        ].rename(
            columns={
                "date": "Fecha",
                "employee_id": "No. Trab",
                "name": "Nombre",
                "department": "Departamento homologado",
                "source_department": "Departamento origen",
                "category": "Categoría",
                "incident_type": "Tipo de incidencia",
            }
        )
        st.dataframe(detail_display, use_container_width=True, hide_index=True)


def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="Panel de Recursos Humanos", page_icon="📊", layout="wide")
    st.markdown(
        """
        <style>
        [data-testid="stMetric"] {
            background: linear-gradient(145deg, #f7fafc 0%, #eef4f8 100%);
            border: 1px solid #d9e3ec;
            border-radius: 14px;
            padding: 16px;
        }
        [data-testid="stMetricValue"] { color: #17365d; }
        .stDownloadButton button {
            background: #17365d;
            color: white;
            border-radius: 10px;
            border: 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    with st.sidebar:
        st.markdown("### Panel de Recursos Humanos")
        selected_view = st.radio(
            "Vista",
            ["Asistencia de checador", "Incidencias verificadas RH"],
            label_visibility="collapsed",
        )
        st.divider()

    if selected_view == "Incidencias verificadas RH":
        render_hr_incidents_view(st)
    else:
        render_attendance_view(st)


if __name__ == "__main__":
    run_streamlit()
