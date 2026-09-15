"""Pantalla de acceso y ciclo de sesión de Streamlit."""

import time

import psycopg2
import streamlit as st

from .auth_store import authenticate, get_active_user
from .database import DatabaseConfigurationError


SESSION_TIMEOUT_SECONDS = 30 * 60
INVALID_LOGIN = (
    "Usuario o contraseña incorrectos, o cuenta no disponible. "
    "Si acumulaste 5 intentos fallidos, espera 5 minutos."
)


def clear_session():
    for key in list(st.session_state):
        del st.session_state[key]


def require_login():
    now = time.time()
    user_id = st.session_state.get("auth_user_id")
    if user_id is not None:
        if now - st.session_state.get("auth_last_activity", 0) >= SESSION_TIMEOUT_SECONDS:
            clear_session()
            st.info("Tu sesión expiró por inactividad. Inicia sesión nuevamente.")
        else:
            try:
                user = get_active_user(user_id)
            except (psycopg2.Error, DatabaseConfigurationError):
                st.error("No se pudo verificar tu sesión. Intenta recargar la página.")
                st.stop()
            if user is not None:
                st.session_state["auth_last_activity"] = now
                st.session_state["username"] = user.username
                with st.sidebar:
                    st.caption("Sesión iniciada")
                    st.text(user.display_name)
                    st.caption(f"Usuario: {user.username}")
                    if st.button("Cerrar sesión", use_container_width=True):
                        clear_session()
                        st.rerun()
                    st.divider()
                return user
            clear_session()
            st.warning("Tu cuenta ya no está disponible. Contacta al administrador.")

    _, center, _ = st.columns([1, 2, 1])
    with center:
        st.title("🔐 Iniciar sesión")
        st.subheader("Dashboard Checadores · UGRPG")
        st.write("Ingresa tus datos para consultar asistencias, incidencias y reportes de Recursos Humanos.")
        with st.form("login", clear_on_submit=True):
            username = st.text_input("Usuario", max_chars=100)
            password = st.text_input("Contraseña", type="password", max_chars=1024)
            submitted = st.form_submit_button("Entrar", use_container_width=True)
        if submitted:
            if not username.strip() or not password:
                st.warning("Escribe tu usuario y contraseña.")
            else:
                try:
                    user = authenticate(username, password)
                except (psycopg2.Error, DatabaseConfigurationError):
                    st.error("No se pudo iniciar sesión. Contacta al administrador o inténtalo más tarde.")
                else:
                    if user is None:
                        st.error(INVALID_LOGIN)
                    else:
                        clear_session()
                        st.session_state["auth_user_id"] = user.id
                        st.session_state["auth_last_activity"] = time.time()
                        st.session_state["username"] = user.username
                        st.rerun()
        st.caption("Para obtener acceso o recuperar tu contraseña, contacta al administrador.")
    st.stop()
