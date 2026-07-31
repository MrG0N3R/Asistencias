from __future__ import annotations

from datetime import datetime

import pandas as pd

from ..hr_export import (
    export_hr_incidents_comparison_workbook,
    export_hr_incidents_workbook,
)
from ..hr_incidents import (
    combine_hr_incident_reports,
    compare_hr_incident_periods,
    hr_comparison_trend,
    hr_department_summary,
    hr_employee_summary,
    hr_period_summary,
    load_hr_incidents,
    summarize_hr_incidents,
)
from ..utils import clean


def _filter_text(incidents, department, categories, incident_types, search):
    filtered = incidents[
        incidents["category"].isin(categories)
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


def _render_department_normalizations(st, incidents):
    normalized_departments = incidents[
        incidents["department"] != incidents["source_department"].str.upper()
    ][["source_department", "department"]].drop_duplicates()
    if not normalized_departments.empty:
        changes = ", ".join(
            f"{row.source_department} → {row.department}"
            for row in normalized_departments.itertuples(index=False)
        )
        st.caption(f"Departamentos homologados para el análisis: {changes}.")


def _render_single_dashboard(st, result):
    incidents = result["incidents"]
    min_date = incidents["date"].min()
    max_date = incidents["date"].max()

    with st.sidebar:
        st.header("Filtros de incidencias RH")
        selected_dates = st.date_input(
            "Rango de fechas",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY",
            key="hr_date_range",
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
            key="hr_department",
        )
        categories = st.multiselect(
            "Categoría",
            sorted(incidents["category"].unique().tolist()),
            default=sorted(incidents["category"].unique().tolist()),
            key="hr_categories",
        )
        incident_types = st.multiselect(
            "Tipo de incidencia",
            sorted(incidents["incident_type"].unique().tolist()),
            default=sorted(incidents["incident_type"].unique().tolist()),
            key="hr_incident_types",
        )
        search = st.text_input("Buscar empleado", key="hr_employee_search")

    filtered = incidents[incidents["date"].between(start_date, end_date)].copy()
    filtered = _filter_text(filtered, department, categories, incident_types, search)

    summary = summarize_hr_incidents(filtered)
    period_text = f"{result['report_start']:%d/%m/%Y} al {result['report_end']:%d/%m/%Y}"
    header_left, header_right = st.columns([1.5, 0.5], vertical_alignment="center")
    with header_left:
        st.markdown(f"**Periodo del reporte:** {period_text}  \n**Archivo:** {result['file_name']}")
        _render_department_normalizations(st, incidents)
    with header_right:
        st.download_button(
            "Descargar dashboard en Excel",
            data=export_hr_incidents_workbook(result, filtered),
            file_name=f"dashboard_incidencias_RH_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=filtered.empty,
        )

    kpi_columns = st.columns(5)
    kpi_columns[0].metric("Personas afectadas", summary["employees"])
    kpi_columns[1].metric("Días de incidencia", summary["incident_days"])
    kpi_columns[2].metric("Ausencias", summary["absences"])
    kpi_columns[3].metric("Incapacidades", summary["disabilities"])
    kpi_columns[4].metric("Departamentos", summary["departments"])

    if filtered.empty:
        st.warning("No hay incidencias que coincidan con los filtros seleccionados.")
        return

    employee_summary = hr_employee_summary(filtered)
    department_summary = hr_department_summary(filtered)
    type_summary = (
        filtered.groupby("incident_type")
        .size()
        .sort_values(ascending=False)
        .rename("Días")
    )

    top_type = type_summary.index[0]
    top_department = department_summary.iloc[0]
    top_employee = employee_summary.iloc[0]
    st.info(
        f"Lectura rápida: **{top_type}** es la incidencia con más días "
        f"({int(type_summary.iloc[0])}). **{top_department['department']}** concentra "
        f"{int(top_department['incident_days'])} días y **{top_employee['name']}** es la "
        f"persona con mayor duración ({int(top_employee['incident_days'])} días)."
    )

    chart_left, chart_right = st.columns([1, 1])
    with chart_left:
        st.subheader("Días por tipo de incidencia")
        st.bar_chart(type_summary, color="#0F766E")
    with chart_right:
        st.subheader("Días por departamento")
        department_chart = department_summary.set_index("department")["incident_days"].head(10)
        st.bar_chart(department_chart, color="#D97706")

    st.subheader("Evolución diaria")
    daily = (
        filtered.assign(
            Ausencias=(filtered["category"] == "Ausencia").astype(int),
            Incapacidades=(filtered["category"] == "Incapacidad").astype(int),
        )
        .groupby("date")[["Ausencias", "Incapacidades"]]
        .sum()
    )
    st.bar_chart(daily, color=["#D97706", "#0F766E"])

    employee_tab, department_tab, detail_tab = st.tabs(
        ["Por empleado", "Por departamento", "Detalle diario"]
    )
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


def _render_comparison_dashboard(st, results, incidents):
    st.success(
        f"Modo comparativo activado: {len(results)} reportes ordenados por su fecha de cierre."
    )

    all_periods = [result["report_label"] for result in results]
    with st.sidebar:
        st.header("Filtros comparativos RH")
        selected_periods = st.multiselect(
            "Periodos incluidos",
            all_periods,
            default=all_periods,
            key="hr_comparison_periods",
            help="La comparación principal usa los dos periodos más recientes de esta selección.",
        )
        departments = sorted(
            value for value in incidents["department"].dropna().unique().tolist() if clean(value)
        )
        department = st.selectbox(
            "Departamento",
            ["Todos"] + departments,
            key="hr_comparison_department",
        )
        categories = st.multiselect(
            "Categoría",
            sorted(incidents["category"].unique().tolist()),
            default=sorted(incidents["category"].unique().tolist()),
            key="hr_comparison_categories",
        )
        incident_types = st.multiselect(
            "Tipo de incidencia",
            sorted(incidents["incident_type"].unique().tolist()),
            default=sorted(incidents["incident_type"].unique().tolist()),
            key="hr_comparison_incident_types",
        )
        search = st.text_input("Buscar empleado", key="hr_comparison_employee_search")

    if len(selected_periods) < 2:
        st.warning("Selecciona al menos dos periodos para mantener activo el comparativo.")
        return

    selected_results = sorted(
        (result for result in results if result["report_label"] in selected_periods),
        key=lambda result: result["report_order"],
    )
    filtered = incidents[incidents["report_label"].isin(selected_periods)].copy()
    filtered = _filter_text(filtered, department, categories, incident_types, search)
    if filtered.empty:
        st.warning("No hay incidencias que coincidan con los filtros seleccionados.")
        return

    period_rows = []
    for result in selected_results:
        period_incidents = filtered[
            filtered["report_order"] == result["report_order"]
        ]
        summary = summarize_hr_incidents(period_incidents)
        period_rows.append(
            {
                "report_order": result["report_order"],
                "report_label": result["report_label"],
                "report_start": result["report_start"],
                "report_end": result["report_end"],
                "source_file": result["file_name"],
                **summary,
            }
        )
    period_summary = pd.DataFrame(period_rows)
    previous = period_summary.iloc[-2]
    current = period_summary.iloc[-1]
    comparison = compare_hr_incident_periods(
        filtered,
        previous_order=int(previous["report_order"]),
        current_order=int(current["report_order"]),
    )
    recurring = int(
        ((comparison["previous_days"] > 0) & (comparison["current_days"] > 0)).sum()
    )
    improved = int(comparison["status"].isin(["Bajó", "Ya no aparece"]).sum())
    new_cases = int((comparison["status"] == "Nueva").sum())

    header_left, header_right = st.columns([1.5, 0.5], vertical_alignment="center")
    with header_left:
        st.markdown(
            f"**Comparación principal:** {previous['report_label']} → "
            f"{current['report_label']}  \n"
            f"**Histórico incluido:** {len(period_summary)} periodos · "
            f"{len(filtered):,} días de incidencia"
        )
        _render_department_normalizations(st, filtered)
    with header_right:
        st.download_button(
            "Descargar comparativo en Excel",
            data=export_hr_incidents_comparison_workbook(
                filtered,
                periods=period_summary,
                comparison=comparison,
            ),
            file_name=f"comparativo_incidencias_RH_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    kpi_columns = st.columns(5)
    kpi_columns[0].metric(
        "Días periodo actual",
        int(current["incident_days"]),
        int(current["incident_days"] - previous["incident_days"]),
        delta_color="inverse",
    )
    kpi_columns[1].metric(
        "Personas afectadas",
        int(current["employees"]),
        int(current["employees"] - previous["employees"]),
        delta_color="inverse",
    )
    kpi_columns[2].metric("Recurrentes", recurring)
    kpi_columns[3].metric("Bajaron o ya no aparecen", improved)
    kpi_columns[4].metric("Nuevas", new_cases)

    status_counts = comparison["status"].value_counts()
    main_direction = int(current["incident_days"] - previous["incident_days"])
    day_label = "día" if abs(main_direction) == 1 else "días"
    if main_direction < 0:
        direction_text = f"bajaron {abs(main_direction)} {day_label}"
    elif main_direction > 0:
        direction_text = f"subieron {main_direction} {day_label}"
    else:
        direction_text = "se mantuvieron sin cambio"
    st.info(
        f"Lectura rápida: las incidencias **{direction_text}** en el último periodo. "
        f"Hay **{recurring}** combinaciones persona/incidencia recurrentes, "
        f"**{int(status_counts.get('Bajó', 0))}** bajaron, "
        f"**{int(status_counts.get('Ya no aparece', 0))}** ya no aparecen y "
        f"**{new_cases}** aparecieron por primera vez."
    )

    trend = period_summary.set_index("report_label")[
        ["incident_days", "absences", "disabilities"]
    ].rename(
        columns={
            "incident_days": "Días totales",
            "absences": "Ausencias",
            "disabilities": "Incapacidades",
        }
    )
    st.subheader("Evolución entre periodos")
    st.line_chart(trend, color=["#17365D", "#D97706", "#0F766E"])

    trend_index = period_summary["report_label"].tolist()
    type_trend = hr_comparison_trend(filtered, "incident_type").reindex(
        trend_index,
        fill_value=0,
    )
    department_trend = hr_comparison_trend(filtered, "department").reindex(
        trend_index,
        fill_value=0,
    )
    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.subheader("Tendencia por tipo")
        top_types = (
            filtered["incident_type"].value_counts().head(8).index.tolist()
        )
        st.line_chart(type_trend.reindex(columns=top_types, fill_value=0))
    with chart_right:
        st.subheader("Tendencia por departamento")
        top_departments = (
            filtered["department"].replace("", "Sin departamento").value_counts().head(8).index.tolist()
        )
        department_chart = department_trend.copy()
        if "" in department_chart.columns:
            department_chart = department_chart.rename(columns={"": "Sin departamento"})
        st.line_chart(department_chart.reindex(columns=top_departments, fill_value=0))

    changes_tab, type_tab, department_tab, detail_tab = st.tabs(
        [
            "Cambios por persona",
            "Tendencia por tipo",
            "Tendencia por departamento",
            "Detalle consolidado",
        ]
    )
    with changes_tab:
        status_filter = st.multiselect(
            "Estados visibles",
            ["Subió", "Nueva", "Se mantiene", "Bajó", "Ya no aparece"],
            default=["Subió", "Nueva", "Se mantiene", "Bajó", "Ya no aparece"],
            key="hr_comparison_statuses",
        )
        change_display = comparison[comparison["status"].isin(status_filter)].rename(
            columns={
                "employee_id": "No. Trab",
                "name": "Nombre",
                "department": "Departamento",
                "category": "Categoría",
                "incident_type": "Tipo de incidencia",
                "previous_days": "Días anteriores",
                "current_days": "Días actuales",
                "change": "Variación",
                "change_percent": "Variación %",
                "status": "Estado",
            }
        )
        st.dataframe(change_display, use_container_width=True, hide_index=True)
    with type_tab:
        type_display = (
            filtered.groupby(
                ["report_order", "report_label", "category", "incident_type"],
                as_index=False,
            )
            .size()
            .rename(
                columns={
                    "report_label": "Periodo",
                    "category": "Categoría",
                    "incident_type": "Tipo de incidencia",
                    "size": "Días",
                }
            )
            .drop(columns="report_order")
        )
        st.dataframe(type_display, use_container_width=True, hide_index=True)
    with department_tab:
        department_display = (
            filtered.assign(
                department=filtered["department"].replace("", "Sin departamento")
            )
            .groupby(["report_order", "report_label", "department"], as_index=False)
            .agg(Personas=("employee_id", "nunique"), Días=("date", "size"))
            .rename(
                columns={
                    "report_label": "Periodo",
                    "department": "Departamento",
                }
            )
            .drop(columns="report_order")
        )
        st.dataframe(department_display, use_container_width=True, hide_index=True)
    with detail_tab:
        detail_display = filtered[
            [
                "report_label",
                "source_file",
                "date",
                "employee_id",
                "name",
                "department",
                "category",
                "incident_type",
            ]
        ].rename(
            columns={
                "report_label": "Periodo",
                "source_file": "Archivo",
                "date": "Fecha",
                "employee_id": "No. Trab",
                "name": "Nombre",
                "department": "Departamento",
                "category": "Categoría",
                "incident_type": "Tipo de incidencia",
            }
        )
        st.dataframe(detail_display, use_container_width=True, hide_index=True)


def render_hr_incidents_view(st):
    st.title("Incidencias verificadas por RH")
    st.caption(
        "Convierte los reportes consolidados de CONTPAQi en una lectura clara por persona, "
        "tipo de incidencia y departamento. Con dos o más archivos se activa el comparativo."
    )

    uploaded_files = st.file_uploader(
        "Reportes de incidencias (.xls o .xlsx)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key="hr_incidents_files",
        help=(
            "Usa reportes 'Listado de incidencias por empleado' revisados por Recursos Humanos. "
            "Carga dos o más periodos para analizar su evolución."
        ),
    )
    if not uploaded_files:
        st.info(
            "Sube un reporte para ver el dashboard individual o varios reportes para comparar "
            "qué incidencias subieron, bajaron, se mantienen o ya no aparecen."
        )
        guide_columns = st.columns(3)
        guide_columns[0].markdown(
            "**1. Sube uno o varios reportes**\n\nSe aceptan archivos `.xls` y `.xlsx`."
        )
        guide_columns[1].markdown(
            "**2. Explora el dashboard**\n\nCon 2+ archivos se activa la evolución histórica."
        )
        guide_columns[2].markdown(
            "**3. Descarga el resultado**\n\nObtén un Excel individual o comparativo."
        )
        return

    results = []
    errors = []
    with st.spinner("Procesando reportes de incidencias..."):
        for uploaded_file in uploaded_files:
            try:
                results.append(
                    load_hr_incidents(uploaded_file.getvalue(), uploaded_file.name)
                )
            except Exception as exc:
                errors.append(f"**{uploaded_file.name}:** {exc}")

    if errors:
        st.error(
            "No fue posible procesar todos los archivos. Corrige o retira los siguientes:\n\n"
            + "\n\n".join(f"- {error}" for error in errors)
        )
        return

    if len(results) == 1:
        _render_single_dashboard(st, results[0])
        return

    ordered_results, combined_incidents = combine_hr_incident_reports(results)
    _render_comparison_dashboard(st, ordered_results, combined_incidents)
