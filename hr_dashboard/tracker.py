from contextlib import closing
import uuid
from datetime import datetime
import streamlit as st
from .database import get_connection


def _get_connection():
    return get_connection()

def track_user_activity(app_name: str, username_override: str = None):
    """
    Registra el acceso y actualiza constantemente la última actividad del usuario.
    Funciona tanto para apps con login propio como para apps sin login.
    """
    # 1. Identificar al usuario
    if username_override:
        current_user = username_override
    elif "username" in st.session_state:
        current_user = st.session_state["username"]
    elif "user" in st.session_state:
        current_user = st.session_state["user"]
    else:
        current_user = "invitado_anonimo"

    # 2. Asignar un ID único a la sesión del navegador
    if "tracker_session_id" not in st.session_state:
        session_id = str(uuid.uuid4())
        if _register_new_access(app_name, current_user, session_id):
            st.session_state.tracker_session_id = session_id
    else:
        _update_last_activity(app_name, current_user, st.session_state.tracker_session_id)

def _register_new_access(app_name: str, username: str, session_id: str):
    """Inserta el primer acceso cuando abre la aplicación."""
    try:
        with closing(_get_connection()) as conn, conn, conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO public.app_user_activity (app_name, username, session_id, first_access, last_activity)
                VALUES (%s, %s, %s, %s, %s);
                """,
                (app_name, username, session_id, datetime.now(), datetime.now())
            )
        return True
    except Exception as e:
        print(f"[Tracker Error] No se pudo registrar acceso: {type(e).__name__}")
        return False

def _update_last_activity(app_name: str, username: str, session_id: str):
    """Actualiza el timestamp de última interacción activa."""
    try:
        with closing(_get_connection()) as conn, conn, conn.cursor() as cur:
            cur.execute(
                """
                UPDATE public.app_user_activity
                SET last_activity = %s, username = %s
                WHERE session_id = %s AND app_name = %s;
                """,
                (datetime.now(), username, session_id, app_name)
            )
    except Exception as e:
        print(f"[Tracker Error] No se pudo actualizar actividad: {type(e).__name__}")
