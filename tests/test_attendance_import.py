from __future__ import annotations

import io
import unittest
from datetime import datetime

from openpyxl import Workbook

from hr_dashboard.attendance import find_header, load_marks


def workbook_bytes(headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["LISTADO DE ENTRADAS Y SALIDAS"])
    sheet.append([])
    sheet.append(["PERIODO DE CONSULTA"])
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


class AttendanceImportTests(unittest.TestCase):
    def test_accepts_entrada_salida_format_without_clock(self):
        data = workbook_bytes(
            ["NO.TRAB", "NOMBRE", "FECHA", "ENTRADA/SALIDA", "DEPARTAMENTO"],
            [
                [100, "Ana", datetime(2026, 8, 1, 8, 0), "E", "ADMINISTRACION"],
                [100, "Ana", datetime(2026, 8, 1, 17, 0), "S", "ADMINISTRACION"],
            ],
        )

        marks = load_marks(data)

        self.assertEqual(len(marks), 2)
        self.assertEqual(marks[0].employee_id, "100")
        self.assertEqual(marks[0].department, "ADMINISTRACION")
        self.assertEqual(marks[0].building, "Planta")

    def test_accepts_format_without_clock_or_direction(self):
        data = workbook_bytes(
            ["NO.TRAB", "NOMBRE", "FECHA", "DEPARTAMENTO"],
            [[150, "Alicia", datetime(2026, 8, 1, 8, 0), "CONTABILIDAD"]],
        )

        marks = load_marks(data)

        self.assertEqual(len(marks), 1)
        self.assertEqual(marks[0].department, "CONTABILIDAD")
        self.assertEqual(marks[0].building, "Planta")

    def test_preserves_clock_format_and_building_inference(self):
        data = workbook_bytes(
            ["NO.TRAB", "NOMBRE", "FECHA", "RELOJ", "DEPARTAMENTO"],
            [[200, "Beto", datetime(2026, 8, 1, 7, 0), "RELOJ CEDIS", "ALMACEN"]],
        )

        marks = load_marks(data)

        self.assertEqual(len(marks), 1)
        self.assertEqual(marks[0].building, "CEDIS")

    def test_header_reports_only_required_columns(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["NO.TRAB", "NOMBRE", "FECHA"])

        with self.assertRaisesRegex(ValueError, "DEPARTAMENTO"):
            find_header(sheet)


if __name__ == "__main__":
    unittest.main()
