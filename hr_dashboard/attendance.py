from __future__ import annotations

import io
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook

from .config import (
    ADMIN_MIN_WORK_HOURS,
    ARRIVAL_TOLERANCE_MINUTES,
    FINISHED_GOODS_ENTRY,
    FINISHED_GOODS_EXIT,
    MAX_PAIR_HOURS,
    MIN_PAIR_HOURS,
    PRODUCTION_MIN_WORK_HOURS,
    PRODUCTION_SHIFTS,
    REGULAR_ENTRY_LIMIT,
    REGULAR_EXIT_LIMIT,
    SCHEDULE_ADMINISTRATIVE,
    SCHEDULE_EXCEPTIONS_FILE,
    SCHEDULE_FINISHED_GOODS,
    SCHEDULE_LABELS,
    SCHEDULE_PRODUCTION,
    SCHEDULE_SECURITY_CEDIS,
    SCHEDULE_SECURITY_PLANT_1,
    SCHEDULE_SECURITY_PLANT_2,
    SCHEDULE_SECURITY_PLANT_MIXED,
    SECURITY_FULL_SHIFT_HOURS,
    SECURITY_PLANT_2_MIN_HOURS,
)
from .utils import clean, normalize_employee_id

@dataclass
class Mark:
    employee_id: str
    name: str
    department: str
    building: str
    timestamp: datetime

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
        "DEPARTAMENTO": "department",
    }
    for row_idx, row in enumerate(ws.iter_rows(max_row=40, values_only=True), start=1):
        normalized = {normalize_header(v): idx for idx, v in enumerate(row) if v is not None}
        if not set(required).issubset(normalized):
            continue

        columns = {target: normalized[source] for source, target in required.items()}
        if "RELOJ" in normalized:
            columns["clock"] = normalized["RELOJ"]
        if "ENTRADASALIDA" in normalized:
            columns["direction"] = normalized["ENTRADASALIDA"]
        return row_idx, columns

    raise ValueError(
        "No se encontro el encabezado esperado: NO.TRAB, NOMBRE, FECHA y DEPARTAMENTO."
    )


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
                marks.append(
                    Mark(
                        employee_id,
                        name,
                        department,
                        infer_building(clock),
                        timestamp,
                    )
                )
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
