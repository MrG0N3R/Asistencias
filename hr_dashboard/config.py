from __future__ import annotations

from datetime import time
from pathlib import Path

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
SCHEDULE_EXCEPTIONS_FILE = Path(__file__).resolve().parent.parent / "schedule_exceptions.json"

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
    "ALMACDN MATERIA PRIMA": "ALMACEN MATERIA PRIMA",
}
