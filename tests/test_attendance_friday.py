from __future__ import annotations

import io
import unittest
from datetime import datetime

from openpyxl import Workbook

from hr_dashboard.attendance import analyze
from hr_dashboard.config import (
    SCHEDULE_ADMINISTRATIVE,
    SCHEDULE_PRODUCTION,
    SCHEDULE_SECURITY_PLANT_2,
)


class AdministrativeFridayTests(unittest.TestCase):
    def analyze_session(self, day=25, entry="08:00", exit_time="16:30",
                        department="ADMINISTRACION", schedule=None):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["NO.TRAB", "NOMBRE", "FECHA", "DEPARTAMENTO"])
        for mark in (entry, exit_time):
            if mark is not None:
                timestamp = datetime.fromisoformat(f"2026-09-{day:02d}T{mark}:00")
                sheet.append([100, "Ana", timestamp, department])
        stream = io.BytesIO()
        workbook.save(stream)
        workbook.close()
        exceptions = {"100": schedule} if schedule else None
        return analyze(stream.getvalue(), "attendance.xlsx", exceptions)

    def test_friday_departure_at_1630_with_arrival_tolerance_is_valid(self):
        for entry in ("08:00", "08:10"):
            with self.subTest(entry=entry):
                result = self.analyze_session(entry=entry)
                record = result["attendance"][0]
                self.assertEqual(record["scheduled_exit"], "16:30")
                self.assertEqual(record["shift"], "Administrativo 08:00-16:30")
                self.assertEqual(record["status"], "Asistencia")
                self.assertFalse(record["exit_out_of_range"])
                self.assertEqual(result["incidents"], [])

    def test_friday_preserves_duration_tolerance_boundary(self):
        valid = self.analyze_session(exit_time="16:12")
        self.assertEqual(valid["incidents"], [])
        short = self.analyze_session(exit_time="16:11")
        self.assertTrue(short["attendance"][0]["short_shift"])
        self.assertEqual(short["incidents"][0]["incident_type"], "Jornada menor a 8.2h")

    def test_other_days_keep_regular_schedule_and_minimum(self):
        for day in (21, 22, 23, 24, 26, 27):
            with self.subTest(day=day):
                result = self.analyze_session(day=day)
                self.assertEqual(result["attendance"][0]["scheduled_exit"], "17:00")
                self.assertEqual(result["incidents"][0]["incident_type"], "Jornada menor a 8.7h")
        self.assertEqual(self.analyze_session(day=24, exit_time="16:42")["incidents"], [])

    def test_friday_still_detects_late_arrival_and_missing_exit(self):
        late = self.analyze_session(entry="08:11")
        self.assertEqual([i["incident_type"] for i in late["incidents"]], ["Retardo"])
        missing = self.analyze_session(exit_time=None)
        self.assertEqual([i["incident_type"] for i in missing["incidents"]], ["Salida faltante"])

    def test_assigned_and_reclassified_administratives_get_friday_schedule(self):
        assigned = self.analyze_session(department="PRODUCCION", schedule=SCHEDULE_ADMINISTRATIVE)
        self.assertEqual(assigned["incidents"], [])
        self.assertEqual(assigned["attendance"][0]["scheduled_exit"], "16:30")
        reclassified = self.analyze_session(department="PRODUCCION", entry="07:59")
        self.assertTrue(reclassified["attendance"][0]["reclassified_as_administrative"])
        self.assertEqual(reclassified["attendance"][0]["scheduled_exit"], "16:30")
        self.assertEqual(reclassified["incidents"], [])

    def test_non_administrative_roles_keep_their_minimum(self):
        cases = (
            ("PRODUCCION", SCHEDULE_PRODUCTION, "06:00", "13:30", "14:00", "Jornada menor a 8h"),
            ("ALMACEN MATERIA PRIMA", None, "06:00", "13:30", "14:00", "Jornada menor a 8h"),
            ("ALMACEN PRODUCTO TERMINADO", None, "07:00", "15:30", "16:30", "Jornada menor a 8.7h"),
            ("VIGILANCIA", SCHEDULE_SECURITY_PLANT_2, "07:00", "17:00", "17:30", "Jornada menor a 10.5h"),
        )
        for department, schedule, entry, exit_time, expected_exit, incident in cases:
            with self.subTest(department=department):
                result = self.analyze_session(department=department, schedule=schedule,
                                              entry=entry, exit_time=exit_time)
                self.assertEqual(result["attendance"][0]["scheduled_exit"], expected_exit)
                self.assertEqual([i["incident_type"] for i in result["incidents"]], [incident])


if __name__ == "__main__":
    unittest.main()
