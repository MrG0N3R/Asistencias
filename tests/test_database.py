from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import psycopg2

from hr_dashboard import database


class DatabaseConfigurationTests(unittest.TestCase):
    def setUp(self):
        self.directory = self.enterContext(TemporaryDirectory())
        self.path = Path(self.directory) / "activity_db.toml"
        self.enterContext(patch("hr_dashboard.database.CONFIG_PATH", self.path))
        self.connect = self.enterContext(patch("hr_dashboard.database.psycopg2.connect"))
        self.path.write_text('dbname = "test"\npassword = "contraseña-prueba"\n', encoding="utf-8")

    def test_utf8_password_is_preserved(self):
        database.get_connection()
        self.assertEqual(self.connect.call_args.kwargs["password"], "contraseña-prueba")
        self.assertEqual(self.connect.call_args.kwargs["client_encoding"], "utf8")

    def test_utf8_bom_from_windows_editor_is_accepted(self):
        self.path.write_text('dbname = "test"\n', encoding="utf-8-sig")
        database.get_connection()
        self.connect.assert_called_once()

    def test_invalid_file_encoding_fails_before_connecting(self):
        self.path.write_bytes('password = "contraseña"\n'.encode("cp1252"))
        with self.assertRaisesRegex(database.DatabaseConfigurationError, "Guarda.*UTF-8"):
            database.get_connection()
        self.connect.assert_not_called()

    def test_spanish_authentication_error_reports_credentials_not_codec(self):
        message = 'FATAL: la autentificación password falló para el usuario «test»'.encode("cp1252")
        self.connect.side_effect = UnicodeDecodeError("utf-8", message, 1, 2, "invalid")
        with self.assertRaisesRegex(database.DatabaseConfigurationError, "rechazó el usuario o la contraseña") as caught:
            database.get_connection()
        self.assertNotIn("contraseña-prueba", str(caught.exception))
        self.assertNotIn("«test»", str(caught.exception))

    def test_other_localized_connection_failure_is_not_reported_as_wrong_password(self):
        message = 'FATAL: la base de datos «test» no existe'.encode("cp1252")
        self.connect.side_effect = UnicodeDecodeError("utf-8", message, 1, 2, "invalid")
        with self.assertRaisesRegex(database.DatabaseConfigurationError, "Falló la conexión"):
            database.get_connection()

    def test_normal_operational_errors_still_propagate(self):
        self.connect.side_effect = psycopg2.OperationalError("connection refused")
        with self.assertRaises(psycopg2.OperationalError):
            database.get_connection()


if __name__ == "__main__":
    unittest.main()
