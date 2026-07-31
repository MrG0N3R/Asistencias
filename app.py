"""Streamlit entry point and compatibility facade.

New code should import from the focused modules under ``hr_dashboard``.
"""

from hr_dashboard.attendance import *
from hr_dashboard.attendance_export import *
from hr_dashboard.config import *
from hr_dashboard.dataframes import *
from hr_dashboard.excel_common import *
from hr_dashboard.hr_export import *
from hr_dashboard.hr_incidents import *
from hr_dashboard.ui import run_streamlit
from hr_dashboard.utils import *
from hr_dashboard.views.attendance import render_attendance_view
from hr_dashboard.views.hr_incidents import render_hr_incidents_view
from hr_dashboard.views.monthly_consolidated import render_monthly_consolidated_view


if __name__ == "__main__":
    run_streamlit()
