"""Streamlit views for the human resources dashboard."""

from .attendance import render_attendance_view
from .hr_incidents import render_hr_incidents_view

__all__ = ["render_attendance_view", "render_hr_incidents_view"]
