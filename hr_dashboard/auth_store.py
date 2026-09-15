"""Usuarios de Dashboard Checadores almacenados en PostgreSQL."""

from contextlib import closing
from dataclasses import dataclass
import hashlib
import hmac
import secrets

from .database import get_connection


ITERATIONS = 600_000
MAX_FAILED_ATTEMPTS = 5
LOCK_MINUTES = 5


@dataclass(frozen=True)
class User:
    id: int
    username: str
    display_name: str


def normalize_username(username):
    return username.strip().lower()


def hash_password(password):
    if not 12 <= len(password) <= 1024:
        raise ValueError("La contraseña debe tener entre 12 y 1024 caracteres.")
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ITERATIONS)
    return f"pbkdf2_sha256${ITERATIONS}${salt.hex()}${digest.hex()}"


def verify_password(password, encoded):
    if not isinstance(password, str) or len(password) > 1024:
        return False
    try:
        algorithm, rounds, salt, expected = encoded.split("$")
        rounds = int(rounds)
        if algorithm != "pbkdf2_sha256" or not 100_000 <= rounds <= 2_000_000:
            return False
        digest = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), bytes.fromhex(salt), rounds
        )
        return hmac.compare_digest(digest, bytes.fromhex(expected))
    except (ValueError, TypeError, AttributeError):
        return False


# Mismo trabajo criptográfico para un nombre de usuario inexistente.
DUMMY_HASH = hash_password(secrets.token_urlsafe(32))


def initialize_database():
    with closing(get_connection()) as conn, conn, conn.cursor() as cur:
        # En la base central esta tabla ya existe; no se modifica su historial.
        cur.execute("""
            CREATE TABLE IF NOT EXISTS public.app_user_activity (
                id SERIAL PRIMARY KEY,
                app_name VARCHAR(255) NOT NULL,
                username VARCHAR(100) NOT NULL,
                session_id VARCHAR(100) NOT NULL,
                first_access TIMESTAMP NOT NULL,
                last_activity TIMESTAMP NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS public.checadores_users (
                id BIGSERIAL PRIMARY KEY,
                username VARCHAR(100) UNIQUE NOT NULL,
                display_name VARCHAR(150) NOT NULL,
                password_hash TEXT NOT NULL,
                is_active BOOLEAN NOT NULL DEFAULT TRUE,
                failed_attempts INTEGER NOT NULL DEFAULT 0,
                locked_until TIMESTAMPTZ,
                created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
                CHECK (username = lower(btrim(username)) AND username <> '')
            )
        """)


def create_user(username, display_name, password):
    username = normalize_username(username)
    display_name = display_name.strip()
    if not username or len(username) > 100 or any(c.isspace() for c in username):
        raise ValueError("El usuario debe tener de 1 a 100 caracteres, sin espacios.")
    if not display_name or len(display_name) > 150:
        raise ValueError("El nombre debe tener de 1 a 150 caracteres.")
    encoded = hash_password(password)
    with closing(get_connection()) as conn, conn, conn.cursor() as cur:
        cur.execute("""
            INSERT INTO public.checadores_users
                (username, display_name, password_hash)
            VALUES (%s, %s, %s)
        """, (username, display_name, encoded))


def authenticate(username, password):
    username = normalize_username(username)
    if not username or len(username) > 100 or not password or len(password) > 1024:
        return None
    with closing(get_connection()) as conn, conn, conn.cursor() as cur:
        cur.execute("""
            SELECT id, username, display_name, password_hash, is_active,
                   COALESCE(locked_until > CURRENT_TIMESTAMP, FALSE),
                   CASE WHEN locked_until <= CURRENT_TIMESTAMP THEN 0
                        ELSE failed_attempts END
            FROM public.checadores_users
            WHERE username = %s
            FOR UPDATE
        """, (username,))
        row = cur.fetchone()
        valid = verify_password(password, row[3] if row else DUMMY_HASH)
        if not row or not row[4] or row[5]:
            return None
        if not valid:
            failures = row[6] + 1
            cur.execute("""
                UPDATE public.checadores_users
                SET failed_attempts = %s,
                    locked_until = CASE WHEN %s >= %s
                        THEN CURRENT_TIMESTAMP + (%s * INTERVAL '1 minute')
                        ELSE NULL END
                WHERE id = %s
            """, (failures, failures, MAX_FAILED_ATTEMPTS, LOCK_MINUTES, row[0]))
            return None
        cur.execute("""
            UPDATE public.checadores_users
            SET failed_attempts = 0, locked_until = NULL WHERE id = %s
        """, (row[0],))
        return User(*row[:3])


def get_active_user(user_id):
    with closing(get_connection()) as conn, conn, conn.cursor() as cur:
        cur.execute("""
            SELECT id, username, display_name FROM public.checadores_users
            WHERE id = %s AND is_active = TRUE
        """, (user_id,))
        row = cur.fetchone()
        return User(*row) if row else None
