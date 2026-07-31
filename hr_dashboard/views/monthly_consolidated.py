from __future__ import annotations

from datetime import datetime

from ..hr_export import export_monthly_consolidated_workbook
from ..hr_incidents import (
    hr_department_summary,
    hr_employee_summary,
    load_hr_incidents,
    monthly_hr_action_recommendations,
    summarize_hr_incidents,
)
from ..utils import clean


def _filter_monthly_incidents(
    incidents,
    start_date,
    end_date,
    department,
    categories,
    incident_types,
    search,
):
    filtered = incidents[
        incidents["date"].between(start_date, end_date)
        & incidents["category"].isin(categories)
        & incidents["incident_type"].isin(incident_types)
    ].copy()
    if department != "Todos":
        filtered = filtered[filtered["department"] == department]
    if search:
        needle = search.casefold()
        filtered = filtered[
            filtered["name"].str.casefold().str.contains(needle, na=False)
            | filtered["employee_id"].astype(str).str.casefold().str.contains(needle, na=False)
        ]
    return filtered


def _render_action_callout(st, recommendation) -> None:
    message = (
        f"**Prioridad {recommendation.priority}: {recommendation.finding}**\n\n"
        f"**Acción sugerida:** {recommendation.suggested_action}\n\n"
        f"**Alcance:** {recommendation.scope}"
    )
    if recommendation.priority == "Alta":
        st.warning(message)
    elif recommendation.priority == "Media":
        st.info(message)
    else:
        st.success(message)
    st.caption(f"Regla activada: {recommendation.rule}")


def render_monthly_consolidated_view(st):
    st.title("Consolidado del Mes")
    st.caption(
        "Analiza el consolidado mensual de incidencias, identifica concentraciones y "
        "genera sugerencias operativas de toma de acción."
    )

    uploaded_file = st.file_uploader(
        "Consolidado mensual de incidencias (.xls o .xlsx)",
        type=["xls", "xlsx"],
        key="monthly_consolidated_file",
        help="Usa el reporte mensual 'Listado de incidencias por empleado' de CONTPAQi.",
    )
    if not uploaded_file:
        st.info(
            "Sube el consolidado del mes para detectar casos prioritarios, áreas de concentración "
            "y acciones sugeridas para seguimiento de RH."
        )
        guide_columns = st.columns(3)
        guide_columns[0].markdown(
            "**1. Sube el consolidado**\n\nSe aceptan archivos `.xls` y `.xlsx`."
        )
        guide_columns[1].markdown(
            "**2. Revisa las alertas**\n\nIdentifica casos prolongados, recurrencia y concentración."
        )
        guide_columns[2].markdown(
            "**3. Ejecuta y documenta**\n\nDescarga el dashboard con la hoja de toma de acción."
        )
        return

    try:
        result = load_hr_incidents(uploaded_file.getvalue(), uploaded_file.name)
    except Exception as exc:
        st.error(str(exc))
        return

    incidents = result["incidents"]
    min_date = incidents["date"].min()
    max_date = incidents["date"].max()

    with st.sidebar:
        st.header("Filtros del consolidado")
        selected_dates = st.date_input(
            "Rango de fechas",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY",
            key="monthly_date_range",
        )
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        elif isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 1:
            start_date = end_date = selected_dates[0]
        else:
            start_date = end_date = selected_dates

        departments = sorted(
            value for value in incidents["department"].dropna().unique().tolist() if clean(value)
        )
        department = st.selectbox(
            "Departamento",
            ["Todos"] + departments,
            key="monthly_department",
        )
        categories = st.multiselect(
            "Categoría",
            sorted(incidents["category"].unique().tolist()),
            default=sorted(incidents["category"].unique().tolist()),
            key="monthly_categories",
        )
        incident_types = st.multiselect(
            "Tipo de incidencia",
            sorted(incidents["incident_type"].unique().tolist()),
            default=sorted(incidents["incident_type"].unique().tolist()),
            key="monthly_incident_types",
        )
        search = st.text_input("Buscar empleado", key="monthly_employee_search")

    filtered = _filter_monthly_incidents(
        incidents,
        start_date,
        end_date,
        department,
        categories,
        incident_types,
        search,
    )
    if filtered.empty:
        st.warning("No hay incidencias que coincidan con los filtros seleccionados.")
        return

    summary = summarize_hr_incidents(filtered)
    recommendations = monthly_hr_action_recommendations(filtered)
    high_priority_count = int((recommendations["priority"] == "Alta").sum())
    period_text = f"{result['report_start']:%d/%m/%Y} al {result['report_end']:%d/%m/%Y}"

    header_left, header_right = st.columns([2, 1], vertical_alignment="center")
    with header_left:
        st.markdown(
            f"**Periodo del consolidado:** {period_text}  \n"
            f"**Archivo:** {result['file_name']}"
        )
        normalized_departments = incidents[
            incidents["department"] != incidents["source_department"].str.upper()
        ][["source_department", "department"]].drop_duplicates()
        if not normalized_departments.empty:
            changes = ", ".join(
                f"{row.source_department} → {row.department}"
                for row in normalized_departments.itertuples(index=False)
            )
            st.caption(f"Departamentos homologados para el análisis: {changes}.")
    with header_right:
        st.download_button(
            "Descargar consolidado analizado",
            data=export_monthly_consolidated_workbook(
                result,
                filtered,
                recommendations,
            ),
            file_name=f"consolidado_mensual_RH_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    kpi_columns = st.columns(4)
    kpi_columns[0].metric("Personas afectadas", summary["employees"])
    kpi_columns[1].metric("Días de incidencia", summary["incident_days"])
    kpi_columns[2].metric("Incapacidades", summary["disabilities"])
    kpi_columns[3].metric("Ausencias", summary["absences"])

    st.subheader("Sugerencia de toma de acción")
    st.markdown(
        f"Se generaron **{len(recommendations)} recomendaciones**; "
        f"**{high_priority_count}** son de prioridad alta."
    )
    _render_action_callout(st, recommendations.iloc[0])
    st.caption(
        "Las sugerencias se recalculan con los filtros y deben validarse contra las políticas "
        "internas de RH, Seguridad e Higiene y la documentación de cada caso."
    )

    employee_summary = hr_employee_summary(filtered)
    department_summary = hr_department_summary(filtered)
    type_summary = (
        filtered.groupby("incident_type")
        .size()
        .sort_values(ascending=False)
        .rename("Días")
    )

    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.subheader("Días por tipo de incidencia")
        st.bar_chart(type_summary, color="#0F766E")
    with chart_right:
        st.subheader("Días por departamento")
        department_chart = department_summary.set_index("department")["incident_days"].head(10)
        st.bar_chart(department_chart, color="#D97706")

    st.subheader("Evolución diaria del mes")
    daily = (
        filtered.assign(
            Ausencias=(filtered["category"] == "Ausencia").astype(int),
            Incapacidades=(filtered["category"] == "Incapacidad").astype(int),
            Otras=(filtered["category"] == "Otra incidencia").astype(int),
        )
        .groupby("date")[["Ausencias", "Incapacidades", "Otras"]]
        .sum()
    )
    st.bar_chart(daily, color=["#D97706", "#0F766E", "#64748B"])

    action_tab, employee_tab, department_tab, detail_tab = st.tabs(
        ["Toma de acción", "Por empleado", "Por departamento", "Detalle diario"]
    )
    with action_tab:
        action_display = recommendations.rename(
            columns={
                "priority": "Prioridad",
                "finding": "Hallazgo",
                "suggested_action": "Acción sugerida",
                "scope": "Alcance",
                "affected_employees": "Personas",
                "incident_days": "Días",
                "rule": "Regla activada",
            }
        )
        st.dataframe(action_display, use_container_width=True, hide_index=True)
    with employee_tab:
        employee_display = employee_summary.rename(
            columns={
                "employee_id": "No. Trab",
                "name": "Nombre",
                "department": "Departamento",
                "incident_days": "Días de incidencia",
                "absences": "Ausencias",
                "disabilities": "Incapacidades",
                "incident_types": "Tipos de incidencia",
                "first_date": "Primera fecha",
                "last_date": "Última fecha",
            }
        )
        st.dataframe(employee_display, use_container_width=True, hide_index=True)
    with department_tab:
        department_display = department_summary.rename(
            columns={
                "department": "Departamento",
                "employees": "Personas",
                "incident_days": "Días de incidencia",
                "absences": "Ausencias",
                "disabilities": "Incapacidades",
            }
        )
        st.dataframe(department_display, use_container_width=True, hide_index=True)
    with detail_tab:
        detail_display = filtered[
            [
                "date",
                "employee_id",
                "name",
                "department",
                "source_department",
                "category",
                "incident_type",
            ]
        ].rename(
            columns={
                "date": "Fecha",
                "employee_id": "No. Trab",
                "name": "Nombre",
                "department": "Departamento homologado",
                "source_department": "Departamento origen",
                "category": "Categoría",
                "incident_type": "Tipo de incidencia",
            }
        )
        st.dataframe(detail_display, use_container_width=True, hide_index=True)
