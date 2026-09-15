import os
import unittest
from unittest.mock import patch
import uuid

from hr_dashboard import auth_store, tracker
from hr_dashboard.database import get_connection

PASSWORD = "Clave-de-prueba-2026!"


class PasswordTests(unittest.TestCase):
    def test_password_hashes_are_salted_and_verified(self):
        first = auth_store.hash_password(PASSWORD)
        second = auth_store.hash_password(PASSWORD)
        self.assertNotEqual(first, second)
        self.assertNotIn(PASSWORD, first)
        self.assertTrue(auth_store.verify_password(PASSWORD, first))
        self.assertFalse(auth_store.verify_password("incorrecta", first))

    def test_rejects_short_passwords_and_invalid_hashes(self):
        with self.assertRaises(ValueError):
            auth_store.hash_password("corta")
        for encoded in (None, "", "texto-plano", "pbkdf2_sha256$999999999$a$b"):
            self.assertFalse(auth_store.verify_password(PASSWORD, encoded))


class TransactionConnection:
    """Evita commits del código probado; el test revierte toda su transacción."""

    def __init__(self, connection):
        self.connection = connection

    def cursor(self):
        return self.connection.cursor()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def close(self):
        pass


@unittest.skipUnless(os.environ.get("RUN_DB_TESTS") == "1", "Requiere PostgreSQL configurado")
class DatabaseLoginTests(unittest.TestCase):
    def setUp(self):
        self.conn = get_connection()
        self.addCleanup(self.conn.close)
        self.addCleanup(self.conn.rollback)
        self.enterContext(patch("hr_dashboard.auth_store.get_connection", return_value=TransactionConnection(self.conn)))
        auth_store.initialize_database()
        self.username = "test_" + uuid.uuid4().hex
        auth_store.create_user(self.username, "Usuario de prueba", PASSWORD)

    def test_valid_invalid_and_injection_credentials(self):
        user = auth_store.authenticate(" " + self.username.upper() + " ", PASSWORD)
        self.assertEqual(user.username, self.username)
        self.assertIsNone(auth_store.authenticate(self.username, "incorrecta"))
        self.assertIsNone(auth_store.authenticate("' OR 1=1 --", PASSWORD))
        self.assertIsNone(auth_store.authenticate("inexistente_" + self.username, PASSWORD))
        self.assertEqual(auth_store.get_active_user(user.id), user)

    def test_lockout_persists_then_expires_and_resets_counter(self):
        for _ in range(5):
            self.assertIsNone(auth_store.authenticate(self.username, "incorrecta"))
        self.assertIsNone(auth_store.authenticate(self.username, PASSWORD))
        with self.conn.cursor() as cur:
            cur.execute("""
                UPDATE public.checadores_users
                SET locked_until = CURRENT_TIMESTAMP - INTERVAL '1 second'
                WHERE username = %s
            """, (self.username,))
        # El primer error después del bloqueo empieza una cuenta nueva.
        self.assertIsNone(auth_store.authenticate(self.username, "incorrecta"))
        self.assertIsNotNone(auth_store.authenticate(self.username, PASSWORD))
        with self.conn.cursor() as cur:
            cur.execute("SELECT failed_attempts, locked_until FROM public.checadores_users WHERE username = %s", (self.username,))
            self.assertEqual(cur.fetchone(), (0, None))

    def test_disabled_user_cannot_login_or_resume(self):
        user = auth_store.authenticate(self.username, PASSWORD)
        with self.conn.cursor() as cur:
            cur.execute("UPDATE public.checadores_users SET is_active = FALSE WHERE id = %s", (user.id,))
        self.assertIsNone(auth_store.authenticate(self.username, PASSWORD))
        self.assertIsNone(auth_store.get_active_user(user.id))

    def test_activity_records_app_user_and_session_in_central_table(self):
        session_id = str(uuid.uuid4())
        app_name = "Dashboard Checadores"
        with patch("hr_dashboard.tracker._get_connection", return_value=TransactionConnection(self.conn)):
            self.assertTrue(tracker._register_new_access(app_name, self.username, session_id))
            tracker._update_last_activity(app_name, self.username, session_id)
        with self.conn.cursor() as cur:
            cur.execute("SELECT app_name, username, first_access, last_activity FROM public.app_user_activity WHERE session_id = %s", (session_id,))
            rows = cur.fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][:2], (app_name, self.username))
        self.assertGreaterEqual(rows[0][3], rows[0][2])

