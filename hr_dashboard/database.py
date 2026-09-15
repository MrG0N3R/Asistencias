"""Conexión compartida para usuarios y registro de actividad."""

from pathlib import Path

import psycopg2
import toml


CONFIG_PATH = Path(__file__).resolve().parents[1] / ".streamlit" / "activity_db.toml"


class DatabaseConfigurationError(RuntimeError):
    pass


def get_connection():
    try:
        # UTF-8 con o sin BOM; no reinterpretar contraseñas con otra codificación.
        config = toml.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
    except UnicodeDecodeError as exc:
        raise DatabaseConfigurationError(
            "Guarda .streamlit/activity_db.toml con codificación UTF-8."
        ) from exc
    except (OSError, ValueError) as exc:
        raise DatabaseConfigurationError(
            "Revisa la configuración en .streamlit/activity_db.toml."
        ) from exc
    config.setdefault("connect_timeout", 5)
    config.setdefault("client_encoding", "utf8")
    try:
        return psycopg2.connect(**config)
    except UnicodeDecodeError as exc:
        # libpq puede devolver errores localizados en Windows-1252 antes de
        # establecer la conexión. No se modifica la codificación de los datos.
        message = exc.object.decode("cp1252", errors="replace").lower()
        if "password" in message and ("fall" in message or "fail" in message):
            detail = "PostgreSQL rechazó el usuario o la contraseña."
        else:
            detail = "Falló la conexión a PostgreSQL y su mensaje no se pudo leer como UTF-8."
        # Evita mostrar la respuesta cruda, que podría contener datos de conexión.
        raise DatabaseConfigurationError(
            f"{detail} Revisa user, password, host y dbname en "
            ".streamlit/activity_db.toml."
        ) from exc
