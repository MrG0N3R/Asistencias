from __future__ import annotations

import hashlib
from datetime import datetime

import pandas as pd

from ..attendance import (
    analyze,
    load_marks,
    load_saved_schedule_exceptions,
    resolve_schedule_selections,
    save_schedule_exceptions,
)
from ..attendance_export import export_workbook
from ..config import (
    SCHEDULE_ADMINISTRATIVE,
    SCHEDULE_FINISHED_GOODS,
    SCHEDULE_PRODUCTION,
    SCHEDULE_SECURITY_CEDIS,
    SCHEDULE_SECURITY_PLANT_1,
    SCHEDULE_SECURITY_PLANT_2,
    SCHEDULE_SECURITY_PLANT_MIXED,
)
from ..dataframes import (
    apply_filters,
    recalculate_period_penalties,
    summarize_incidents,
    to_frame,
)

def render_attendance_view(st):
    st.title("Dashboard de asistencia")
    st.caption(
        "Carga archivos .xlsx con columna RELOJ o con el formato ENTRADA/SALIDA."
    )

    uploaded_file = st.file_uploader(
        "Archivo de asistencia (.xlsx)",
        type=["xlsx"],
        help=(
            "Se aceptan los encabezados NO.TRAB, NOMBRE, FECHA, DEPARTAMENTO "
            "y una columna RELOJ o ENTRADA/SALIDA."
        ),
    )
    if not uploaded_file:
        st.info("Sube un archivo para calcular asistencias, retardos, incidencias y penalizaciones por periodo jueves a miercoles.")
        return

    file_bytes = uploaded_file.getvalue()
    try:
        preview_marks = load_marks(file_bytes)
    except Exception as exc:
        st.error(str(exc))
        return

    if preview_marks and all(mark.building == "Sin especificar" for mark in preview_marks):
        st.warning(
            "Este archivo no incluye la columna RELOJ. Las marcas se analizarán "
            "normalmente, pero el edificio aparecerá como 'Sin especificar'."
        )

    employee_meta = {}
    for mark in preview_marks:
        employee_meta.setdefault(
            mark.employee_id,
            {"name": mark.name, "department": mark.department},
        )
    employee_ids = sorted(
        employee_meta,
        key=lambda value: (not value.isdigit(), int(value) if value.isdigit() else value),
    )

    def employee_label(employee_id: str) -> str:
        meta = employee_meta[employee_id]
        return f"{employee_id} - {meta['name']} ({meta['department']})"

    try:
        saved_schedule_exceptions = load_saved_schedule_exceptions()
        saved_exceptions_error = ""
    except ValueError as exc:
        saved_schedule_exceptions = {}
        saved_exceptions_error = str(exc)

    file_widget_key = hashlib.sha256(file_bytes).hexdigest()[:12]
    with st.sidebar:
        st.header("Excepciones de horario")
        st.caption(
            "La asignacion por No. Trab es estricta y reemplaza el departamento reportado en el archivo."
        )
        st.caption(
            "Para un rol mixto, selecciona al mismo empleado en Vigilancia 1 Planta y Vigilancia 2 Planta."
        )
        if saved_exceptions_error:
            st.warning(saved_exceptions_error)
        administrative_ids = st.multiselect(
            "Administrativo 08:00-17:00",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_ADMINISTRATIVE
            ],
            format_func=employee_label,
            key=f"administrative_exceptions_{file_widget_key}",
        )
        production_ids = st.multiselect(
            "Produccion (turnos rotativos)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_PRODUCTION
            ],
            format_func=employee_label,
            key=f"production_exceptions_{file_widget_key}",
        )
        finished_goods_ids = st.multiselect(
            "Producto terminado 07:00-16:30",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_FINISHED_GOODS
            ],
            format_func=employee_label,
            key=f"finished_goods_exceptions_{file_widget_key}",
        )
        security_plant_1_ids = st.multiselect(
            "Vigilancia 1 Planta 12x24 (07-19 / 19-07)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id)
                in {SCHEDULE_SECURITY_PLANT_1, SCHEDULE_SECURITY_PLANT_MIXED}
            ],
            format_func=employee_label,
            key=f"security_plant_1_exceptions_{file_widget_key}",
        )
        security_plant_2_ids = st.multiselect(
            "Vigilancia 2 Planta (07:00-17:30)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id)
                in {SCHEDULE_SECURITY_PLANT_2, SCHEDULE_SECURITY_PLANT_MIXED}
            ],
            format_func=employee_label,
            key=f"security_plant_2_exceptions_{file_widget_key}",
        )
        security_cedis_ids = st.multiselect(
            "Vigilancia CEDIS 24x12 (07:00-19:00)",
            employee_ids,
            default=[
                employee_id
                for employee_id in employee_ids
                if saved_schedule_exceptions.get(employee_id) == SCHEDULE_SECURITY_CEDIS
            ],
            format_func=employee_label,
            key=f"security_cedis_exceptions_{file_widget_key}",
        )

    selected_schedule_groups = {
        SCHEDULE_ADMINISTRATIVE: administrative_ids,
        SCHEDULE_PRODUCTION: production_ids,
        SCHEDULE_FINISHED_GOODS: finished_goods_ids,
        SCHEDULE_SECURITY_PLANT_1: security_plant_1_ids,
        SCHEDULE_SECURITY_PLANT_2: security_plant_2_ids,
        SCHEDULE_SECURITY_CEDIS: security_cedis_ids,
    }
    schedule_exceptions, conflicts = resolve_schedule_selections(selected_schedule_groups)
    if conflicts:
        conflict_names = ", ".join(employee_label(employee_id) for employee_id in sorted(conflicts))
        st.error(f"Cada empleado solo puede tener un horario de excepcion. Revisa: {conflict_names}")
        return

    with st.sidebar:
        st.caption(
            f"{len(schedule_exceptions)} aplicada(s) a este archivo; "
            f"{len(saved_schedule_exceptions)} guardada(s) en el catalogo."
        )
        if st.button("Guardar excepciones para futuros archivos", use_container_width=True):
            merged_exceptions = {
                employee_id: schedule
                for employee_id, schedule in saved_schedule_exceptions.items()
                if employee_id not in employee_meta
            }
            merged_exceptions.update(schedule_exceptions)
            try:
                save_schedule_exceptions(merged_exceptions)
                st.success("Catalogo de excepciones guardado.")
            except OSError as exc:
                st.error(f"No se pudo guardar el catalogo de excepciones: {exc}")

    try:
        result = analyze(file_bytes, uploaded_file.name, schedule_exceptions)
    except Exception as exc:
        st.error(str(exc))
        return

    attendance_df = to_frame(result["attendance"])
    incidents_df = to_frame(result["incidents"])
    available_dates = pd.to_datetime(attendance_df["date"], errors="coerce").dt.date.dropna()
    min_date = available_dates.min()
    max_date = available_dates.max()

    with st.sidebar:
        st.header("Filtros")
        selected_dates = st.date_input(
            "Rango de fechas a evaluar",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY",
        )
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        elif isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 1:
            start_date = end_date = selected_dates[0]
        else:
            start_date = end_date = selected_dates
        building = st.selectbox("Edificio", ["Todos"] + sorted(attendance_df["building"].dropna().unique().tolist()))
        department = st.selectbox("Departamento", ["Todos"] + sorted(attendance_df["department"].dropna().unique().tolist()))
        period = st.selectbox("Periodo", ["Todos"] + sorted(attendance_df["period"].dropna().unique().tolist()))
        search = st.text_input("Buscar empleado")
        st.divider()
        st.download_button(
            "Descargar consolidado",
            data=export_workbook(result),
            file_name=f"consolidado_asistencia_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    filtered_attendance = apply_filters(
        attendance_df, building, department, period, search, start_date, end_date
    )
    filtered_incidents = apply_filters(
        incidents_df, building, department, period, search, start_date, end_date
    )
    filtered_incidents, penalties_df = recalculate_period_penalties(filtered_incidents)

    kpi_cols = st.columns(5)
    kpi_cols[0].metric("Empleados", filtered_attendance["employee_id"].nunique() if not filtered_attendance.empty else 0)
    kpi_cols[1].metric("Dias asistencia", len(filtered_attendance))
    kpi_cols[2].metric("Retardos", int(filtered_attendance["late"].sum()) if not filtered_attendance.empty else 0)
    kpi_cols[3].metric("Incidencias", len(filtered_incidents))
    kpi_cols[4].metric("Bonos perdidos", len(penalties_df))

    st.subheader("Reglas aplicadas")
    st.dataframe(pd.DataFrame([{"Regla": key, "Valor": value} for key, value in result["rules"].items()]), use_container_width=True, hide_index=True)
    if result["schedule_exceptions"]:
        with st.expander(f"Excepciones de horario activas ({len(result['schedule_exceptions'])})"):
            st.dataframe(
                pd.DataFrame(result["schedule_exceptions"]),
                use_container_width=True,
                hide_index=True,
            )

    chart_left, chart_right = st.columns([1.2, 0.8])
    with chart_left:
        st.subheader("Resumen por periodo")
        if filtered_attendance.empty:
            st.warning("Sin datos para los filtros actuales.")
        else:
            weekly = (
                filtered_attendance.groupby("period", as_index=False)
                .agg(Asistencias=("employee_id", "count"), Retardos=("late", "sum"), Jornadas_incompletas=("short_shift", "sum"))
                .set_index("period")
            )
            st.bar_chart(weekly[["Retardos", "Jornadas_incompletas"]])
    with chart_right:
        st.subheader("Asistencias por edificio")
        if not filtered_attendance.empty:
            st.bar_chart(filtered_attendance["building"].value_counts())

    st.subheader("Departamentos con mas incidencias")
    if filtered_incidents.empty:
        st.success("Sin incidencias para los filtros actuales.")
    else:
        st.bar_chart(filtered_incidents["department"].value_counts().head(10))

    st.subheader("Incidencias por empleado")
    incident_summary = summarize_incidents(filtered_incidents)
    st.dataframe(incident_summary, use_container_width=True, hide_index=True)

    st.subheader("Asistencias del personal")
    attendance_columns = [
        "employee_id", "name", "department", "schedule_type", "schedule_source",
        "building", "date", "period", "shift",
        "scheduled_entry", "scheduled_exit", "entry", "exit", "worked_hours", "status",
    ]
    st.dataframe(filtered_attendance[attendance_columns] if not filtered_attendance.empty else filtered_attendance, use_container_width=True, hide_index=True)
