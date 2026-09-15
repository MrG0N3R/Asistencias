from __future__ import annotations

from .views.attendance import render_attendance_view
from .views.hr_incidents import render_hr_incidents_view
from .views.monthly_consolidated import render_monthly_consolidated_view

def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="Panel de Recursos Humanos", page_icon="📊", layout="wide")
    from .auth import require_login
    from .tracker import track_user_activity

    current_user = require_login()
    track_user_activity(
        app_name="Dashboard Checadores",
        username_override=current_user.username,
    )
    st.markdown(
        """
        <style>
        [data-testid="stMetric"] {
            background: linear-gradient(145deg, #f7fafc 0%, #eef4f8 100%);
            border: 1px solid #d9e3ec;
            border-radius: 14px;
            padding: 16px;
        }
        [data-testid="stMetricLabel"],
        [data-testid="stMetricLabel"] p { color: #425466; }
        [data-testid="stMetricValue"] { color: #17365d; }
        .stDownloadButton button {
            background: #17365d;
            color: white;
            border-radius: 10px;
            border: 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    with st.sidebar:
        st.markdown("### Panel de Recursos Humanos")
        selected_view = st.radio(
            "Vista",
            [
                "Asistencia de checador",
                "Incidencias verificadas RH",
                "Consolidado del Mes",
            ],
            label_visibility="collapsed",
        )
        st.divider()

    if selected_view == "Incidencias verificadas RH":
        render_hr_incidents_view(st)
    elif selected_view == "Consolidado del Mes":
        render_monthly_consolidated_view(st)
    else:
        render_attendance_view(st)
