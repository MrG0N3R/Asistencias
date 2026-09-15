from pathlib import Path
import time
import unittest
from unittest.mock import patch

import psycopg2
from streamlit.testing.v1 import AppTest

from hr_dashboard.auth_store import User
from hr_dashboard.database import DatabaseConfigurationError


ROOT = Path(__file__).resolve().parents[1]


class LoginTests(unittest.TestCase):
    def setUp(self):
        self.user = User(1, "operador", "Operador de RH")
        self.authenticate = self.enterContext(patch("hr_dashboard.auth.authenticate", return_value=self.user))
        self.active_user = self.enterContext(patch("hr_dashboard.auth.get_active_user", return_value=self.user))
        self.register = self.enterContext(patch("hr_dashboard.tracker._register_new_access", return_value=True))
        self.update = self.enterContext(patch("hr_dashboard.tracker._update_last_activity"))
        self.views = [self.enterContext(patch("hr_dashboard.ui." + name)) for name in (
            "render_attendance_view", "render_hr_incidents_view", "render_monthly_consolidated_view"
        )]
        self.app = AppTest.from_file(str(ROOT / "app.py"), default_timeout=15)

    def login(self):
        self.app.run()
        self.app.text_input[0].set_value("operador")
        self.app.text_input[1].set_value("Clave-de-prueba-2026!")
        self.app.button[0].click().run()
        self.assertEqual(len(self.app.exception), 0)

    def test_anonymous_user_cannot_open_any_hr_view(self):
        self.app.run()
        self.assertIn("Iniciar sesión", self.app.title[0].value)
        self.assertEqual(len(self.app.radio), 0)
        for view in self.views:
            view.assert_not_called()
        self.register.assert_not_called()

    def test_login_enables_all_three_views_and_tracks_correct_app(self):
        self.login()
        self.views[0].assert_called_once()
        self.assertEqual(self.register.call_args.args[:2], ("Dashboard Checadores", "operador"))
        self.app.sidebar.radio[0].set_value("Incidencias verificadas RH").run()
        self.views[1].assert_called_once()
        self.app.sidebar.radio[0].set_value("Consolidado del Mes").run()
        self.views[2].assert_called_once()
        self.assertEqual(self.register.call_count, 1)
        self.assertEqual(self.update.call_count, 2)

    def test_logout_clears_user_data_and_rotates_next_tracking_session(self):
        self.login()
        first_session = self.app.session_state["tracker_session_id"]
        self.app.session_state["private_report"] = b"private HR report"
        self.app.sidebar.button[0].click().run()
        for key in ("auth_user_id", "username", "tracker_session_id", "private_report"):
            self.assertNotIn(key, self.app.session_state)
        self.assertIn("Iniciar sesión", self.app.title[0].value)
        self.login()
        self.assertNotEqual(first_session, self.app.session_state["tracker_session_id"])

    def test_wrong_password_keeps_reports_hidden(self):
        self.authenticate.return_value = None
        self.login()
        self.assertEqual(len(self.app.error), 1)
        for view in self.views:
            view.assert_not_called()
        self.register.assert_not_called()

    def test_missing_configuration_and_database_failure_keep_reports_hidden(self):
        for exc in (DatabaseConfigurationError("private config"), psycopg2.OperationalError("private connection")):
            self.authenticate.side_effect = exc
            self.login()
            self.assertEqual(len(self.app.error), 1)
            self.assertNotIn("private", self.app.error[0].value)
        for view in self.views:
            view.assert_not_called()

    def test_disabled_account_cannot_continue(self):
        self.login()
        self.active_user.return_value = None
        self.app.run()
        self.assertNotIn("auth_user_id", self.app.session_state)
        self.assertIn("Iniciar sesión", self.app.title[0].value)
        self.update.assert_not_called()

    def test_expired_session_requires_login(self):
        self.login()
        self.app.session_state["auth_last_activity"] = time.time() - 1801
        self.app.run()
        self.assertIn("Iniciar sesión", self.app.title[0].value)
        self.assertNotIn("auth_user_id", self.app.session_state)

    def test_activity_registration_retries_after_failure(self):
        self.register.return_value = False
        self.login()
        self.assertNotIn("tracker_session_id", self.app.session_state)
        self.register.return_value = True
        self.app.run()
        self.assertEqual(self.register.call_count, 2)
        self.assertIn("tracker_session_id", self.app.session_state)


if __name__ == "__main__":
    unittest.main()
