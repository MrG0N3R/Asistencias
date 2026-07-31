from __future__ import annotations

import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from .config import HR_DEPARTMENT_ALIASES, HR_INCIDENT_LABELS, SPANISH_MONTHS
from .utils import clean, normalize_employee_id

def parse_spanish_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = clean(value)
    match = re.fullmatch(r"(\d{1,2})/([A-Za-zÁÉÍÓÚÑáéíóúñ]{3})/(\d{4})", text)
    if match:
        month = SPANISH_MONTHS.get(match.group(2).upper())
        if month:
            try:
                return date(int(match.group(3)), month, int(match.group(1)))
            except ValueError:
                return None
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def friendly_hr_incident_type(value: str) -> str:
    text = re.sub(r"\s+\((?:Aus|Inc)\)\s*$", "", clean(value), flags=re.IGNORECASE)
    normalized = re.sub(r"\s+", " ", text.upper()).strip()
    return HR_INCIDENT_LABELS.get(normalized, text.title())


def hr_incident_category(value: str) -> str:
    text = clean(value).upper()
    if "(AUS)" in text:
        return "Ausencia"
    if "(INC)" in text:
        return "Incapacidad"
    return "Otra incidencia"


def friendly_hr_department(value: str) -> str:
    department = clean(value).upper()
    return HR_DEPARTMENT_ALIASES.get(department, department)


def extract_hr_report_period(frames: dict[str, pd.DataFrame]) -> tuple[date | None, date | None]:
    pattern = re.compile(
        r"REPORTE DE INCIDENCIAS DEL\s+(\d{1,2}/[A-ZÁÉÍÓÚÑ]{3}/\d{4})\s+AL\s+"
        r"(\d{1,2}/[A-ZÁÉÍÓÚÑ]{3}/\d{4})",
        re.IGNORECASE,
    )
    for frame in frames.values():
        for value in frame.fillna("").astype(str).to_numpy().ravel():
            match = pattern.search(clean(value))
            if match:
                return parse_spanish_date(match.group(1)), parse_spanish_date(match.group(2))
    return None, None


def load_hr_incidents(file_bytes: bytes, file_name: str) -> dict:
    extension = Path(file_name).suffix.lower()
    if extension not in {".xls", ".xlsx"}:
        raise ValueError("El archivo debe estar en formato .xls o .xlsx.")

    engine = "xlrd" if extension == ".xls" else "openpyxl"
    try:
        frames = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=None,
            header=None,
            dtype=object,
            engine=engine,
        )
    except ImportError as exc:
        raise ValueError(
            "No está instalado el componente para leer archivos .xls. "
            "Instala las dependencias del proyecto e intenta de nuevo."
        ) from exc
    except Exception as exc:
        raise ValueError(f"No se pudo leer el reporte de incidencias: {exc}") from exc

    report_start, report_end = extract_hr_report_period(frames)
    incidents: list[dict] = []
    employees_seen: set[str] = set()

    for sheet_name, frame in frames.items():
        current_employee: dict | None = None
        for row in frame.itertuples(index=False, name=None):
            values = [None if pd.isna(value) else value for value in row]
            first = clean(values[0]) if values else ""
            second = clean(values[1]) if len(values) > 1 else ""

            if re.fullmatch(r"\d{1,10}(?:\.0+)?", first) and second:
                employee_id = normalize_employee_id(values[0])
                current_employee = {
                    "employee_id": employee_id,
                    "name": second.replace("�", "Ñ"),
                    "department": friendly_hr_department(values[5]) if len(values) > 5 else "",
                    "source_department": clean(values[5]) if len(values) > 5 else "",
                }
                employees_seen.add(employee_id)
                continue

            incident_date = parse_spanish_date(values[0]) if values else None
            if not current_employee or not incident_date or not second:
                continue
            if second.upper().startswith("TOTAL "):
                continue

            incidents.append(
                {
                    **current_employee,
                    "date": incident_date,
                    "category": hr_incident_category(second),
                    "incident_type": friendly_hr_incident_type(second),
                    "source_description": second,
                    "sheet": sheet_name,
                }
            )

    if not incidents:
        raise ValueError(
            "No se encontraron incidencias. Verifica que sea el reporte "
            "'Listado de incidencias por empleado' de CONTPAQi."
        )

    incidents_df = pd.DataFrame(incidents).sort_values(
        ["date", "department", "name"], ascending=[True, True, True]
    )
    if report_start is None:
        report_start = incidents_df["date"].min()
    if report_end is None:
        report_end = incidents_df["date"].max()

    return {
        "file_name": file_name,
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "report_start": report_start,
        "report_end": report_end,
        "incidents": incidents_df,
        "employees_in_report": len(employees_seen),
    }


def summarize_hr_incidents(incidents: pd.DataFrame) -> dict:
    if incidents.empty:
        return {
            "employees": 0,
            "incident_days": 0,
            "absences": 0,
            "disabilities": 0,
            "departments": 0,
        }
    return {
        "employees": int(incidents["employee_id"].nunique()),
        "incident_days": int(len(incidents)),
        "absences": int((incidents["category"] == "Ausencia").sum()),
        "disabilities": int((incidents["category"] == "Incapacidad").sum()),
        "departments": int(incidents["department"].replace("", pd.NA).nunique()),
    }


def hr_employee_summary(incidents: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "employee_id",
        "name",
        "department",
        "incident_days",
        "absences",
        "disabilities",
        "incident_types",
        "first_date",
        "last_date",
    ]
    if incidents.empty:
        return pd.DataFrame(columns=columns)

    def joined_unique(values) -> str:
        return ", ".join(dict.fromkeys(clean(value) for value in values if clean(value)))

    summary = (
        incidents.groupby(["employee_id", "name", "department"], as_index=False, dropna=False)
        .agg(
            incident_days=("date", "size"),
            absences=("category", lambda values: int((values == "Ausencia").sum())),
            disabilities=("category", lambda values: int((values == "Incapacidad").sum())),
            incident_types=("incident_type", joined_unique),
            first_date=("date", "min"),
            last_date=("date", "max"),
        )
        .sort_values(["incident_days", "name"], ascending=[False, True])
    )
    return summary[columns]


def hr_department_summary(incidents: pd.DataFrame) -> pd.DataFrame:
    columns = ["department", "employees", "incident_days", "absences", "disabilities"]
    if incidents.empty:
        return pd.DataFrame(columns=columns)
    summary = (
        incidents.assign(department=incidents["department"].replace("", "Sin departamento"))
        .groupby("department", as_index=False)
        .agg(
            employees=("employee_id", "nunique"),
            incident_days=("date", "size"),
            absences=("category", lambda values: int((values == "Ausencia").sum())),
            disabilities=("category", lambda values: int((values == "Incapacidad").sum())),
        )
        .sort_values(["incident_days", "department"], ascending=[False, True])
    )
    return summary[columns]


def combine_hr_incident_reports(results: list[dict]) -> tuple[list[dict], pd.DataFrame]:
    """Ordena reportes por periodo y agrega metadatos para compararlos."""
    if not results:
        return [], pd.DataFrame()

    ordered_results = sorted(
        results,
        key=lambda result: (
            result.get("report_end") or date.min,
            result.get("report_start") or date.min,
            result.get("file_name", ""),
        ),
    )

    base_labels = []
    for result in ordered_results:
        start = result.get("report_start")
        end = result.get("report_end")
        if start and end:
            base_labels.append(f"{start:%d/%m/%Y} al {end:%d/%m/%Y}")
        else:
            base_labels.append(result.get("file_name", "Periodo sin identificar"))

    duplicated_labels = {label for label in base_labels if base_labels.count(label) > 1}
    combined_frames = []
    for order, (result, base_label) in enumerate(zip(ordered_results, base_labels), start=1):
        label = (
            f"{base_label} · {result['file_name']}"
            if base_label in duplicated_labels
            else base_label
        )
        result["report_order"] = order
        result["report_label"] = label

        frame = result["incidents"].copy()
        frame["source_file"] = result["file_name"]
        frame["report_start"] = result["report_start"]
        frame["report_end"] = result["report_end"]
        frame["report_order"] = order
        frame["report_label"] = label
        combined_frames.append(frame)

    combined = pd.concat(combined_frames, ignore_index=True)
    combined = combined.sort_values(
        ["report_order", "date", "department", "name"],
        ascending=[True, True, True, True],
    )
    return ordered_results, combined


def hr_period_summary(incidents: pd.DataFrame) -> pd.DataFrame:
    """Resume los indicadores de cada archivo/periodo cargado."""
    columns = [
        "report_order",
        "report_label",
        "report_start",
        "report_end",
        "source_file",
        "employees",
        "incident_days",
        "absences",
        "disabilities",
        "departments",
    ]
    if incidents.empty:
        return pd.DataFrame(columns=columns)

    summary = (
        incidents.groupby(
            ["report_order", "report_label", "report_start", "report_end", "source_file"],
            as_index=False,
            dropna=False,
        )
        .agg(
            employees=("employee_id", "nunique"),
            incident_days=("date", "size"),
            absences=("category", lambda values: int((values == "Ausencia").sum())),
            disabilities=("category", lambda values: int((values == "Incapacidad").sum())),
            departments=(
                "department",
                lambda values: int(values.replace("", pd.NA).nunique()),
            ),
        )
        .sort_values("report_order")
    )
    return summary[columns]


def compare_hr_incident_periods(
    incidents: pd.DataFrame,
    previous_order: int | None = None,
    current_order: int | None = None,
) -> pd.DataFrame:
    """Compara días por persona y tipo entre dos periodos.

    El estado describe el movimiento del periodo actual respecto al anterior:
    Nueva, Subió, Bajó, Se mantiene o Ya no aparece.
    """
    columns = [
        "employee_id",
        "name",
        "department",
        "category",
        "incident_type",
        "previous_days",
        "current_days",
        "change",
        "change_percent",
        "status",
    ]
    if incidents.empty or "report_order" not in incidents:
        return pd.DataFrame(columns=columns)

    available_orders = sorted(incidents["report_order"].dropna().astype(int).unique())
    if current_order is None:
        if not available_orders:
            return pd.DataFrame(columns=columns)
        current_order = available_orders[-1]
    if previous_order is None:
        earlier_orders = [order for order in available_orders if order < current_order]
        if not earlier_orders:
            return pd.DataFrame(columns=columns)
        previous_order = earlier_orders[-1]

    selected = incidents[incidents["report_order"].isin([previous_order, current_order])].copy()
    counts = (
        selected.groupby(["employee_id", "incident_type", "report_order"])
        .size()
        .unstack(fill_value=0)
    )
    for order in (previous_order, current_order):
        if order not in counts:
            counts[order] = 0
    counts = counts[[previous_order, current_order]].rename(
        columns={previous_order: "previous_days", current_order: "current_days"}
    )

    metadata = (
        selected.sort_values(["report_order", "date"])
        .groupby(["employee_id", "incident_type"], as_index=True)
        .agg(
            name=("name", "last"),
            department=("department", "last"),
            category=("category", "last"),
        )
    )
    comparison = metadata.join(counts, how="outer").reset_index()
    comparison["previous_days"] = comparison["previous_days"].fillna(0).astype(int)
    comparison["current_days"] = comparison["current_days"].fillna(0).astype(int)
    comparison["change"] = comparison["current_days"] - comparison["previous_days"]

    def status(row) -> str:
        if row["previous_days"] == 0:
            return "Nueva"
        if row["current_days"] == 0:
            return "Ya no aparece"
        if row["change"] > 0:
            return "Subió"
        if row["change"] < 0:
            return "Bajó"
        return "Se mantiene"

    comparison["status"] = comparison.apply(status, axis=1)
    comparison["change_percent"] = comparison.apply(
        lambda row: (
            None
            if row["previous_days"] == 0
            else round((row["change"] / row["previous_days"]) * 100, 1)
        ),
        axis=1,
    )
    status_order = {
        "Subió": 0,
        "Nueva": 1,
        "Se mantiene": 2,
        "Bajó": 3,
        "Ya no aparece": 4,
    }
    comparison["_status_order"] = comparison["status"].map(status_order)
    comparison = comparison.sort_values(
        ["_status_order", "change", "current_days", "name"],
        ascending=[True, False, False, True],
    ).drop(columns="_status_order")
    return comparison[columns]


def hr_comparison_trend(
    incidents: pd.DataFrame,
    dimension: str,
) -> pd.DataFrame:
    """Devuelve una matriz periodo × dimensión lista para graficar."""
    if incidents.empty or dimension not in incidents:
        return pd.DataFrame()
    return (
        incidents.groupby(["report_order", "report_label", dimension])
        .size()
        .rename("incident_days")
        .reset_index()
        .pivot_table(
            index=["report_order", "report_label"],
            columns=dimension,
            values="incident_days",
            fill_value=0,
        )
        .sort_index(level="report_order")
        .droplevel("report_order")
    )


def monthly_hr_action_recommendations(incidents: pd.DataFrame) -> pd.DataFrame:
    """Propone acciones operativas a partir de señales visibles en el consolidado.

    Las reglas son deliberadamente simples y auditables. No sustituyen las políticas
    internas de RH ni una valoración médica o legal.
    """
    columns = [
        "priority",
        "finding",
        "suggested_action",
        "scope",
        "affected_employees",
        "incident_days",
        "rule",
    ]
    if incidents.empty:
        return pd.DataFrame(columns=columns)

    recommendations: list[dict] = []

    def add_recommendation(
        priority: str,
        finding: str,
        suggested_action: str,
        scope: str,
        affected_employees: int,
        incident_days: int,
        rule: str,
    ) -> None:
        recommendations.append(
            {
                "priority": priority,
                "finding": finding,
                "suggested_action": suggested_action,
                "scope": scope,
                "affected_employees": int(affected_employees),
                "incident_days": int(incident_days),
                "rule": rule,
            }
        )

    missing_department = incidents[incidents["department"].fillna("").str.strip().eq("")]
    if not missing_department.empty:
        names = ", ".join(missing_department["name"].drop_duplicates().head(5))
        add_recommendation(
            "Alta",
            "Hay incidencias sin departamento asignado.",
            "Completar el departamento antes de distribuir resultados y corregir el catálogo maestro.",
            names,
            missing_department["employee_id"].nunique(),
            len(missing_department),
            "Se activa cuando existe al menos un registro sin departamento.",
        )

    disability_by_employee = (
        incidents[incidents["category"] == "Incapacidad"]
        .groupby(["employee_id", "name", "department"], as_index=False)
        .size()
        .rename(columns={"size": "days"})
    )
    prolonged = disability_by_employee[disability_by_employee["days"] >= 15].sort_values(
        "days", ascending=False
    )
    if not prolonged.empty:
        scope = ", ".join(
            f"{row.name} ({int(row.days)} días)"
            for row in prolonged.head(5).itertuples(index=False)
        )
        add_recommendation(
            "Alta",
            f"{len(prolonged)} persona(s) acumulan 15 o más días de incapacidad.",
            "Revisar documentación, seguimiento del caso y plan de reincorporación con RH y Seguridad e Higiene.",
            scope,
            len(prolonged),
            prolonged["days"].sum(),
            "Incapacidad acumulada por persona ≥ 15 días en el mes.",
        )

    absence_by_employee = (
        incidents[incidents["category"] == "Ausencia"]
        .groupby(["employee_id", "name", "department"], as_index=False)
        .size()
        .rename(columns={"size": "days"})
    )
    recurrent_absence = absence_by_employee[absence_by_employee["days"] >= 3].sort_values(
        "days", ascending=False
    )
    if not recurrent_absence.empty:
        scope = ", ".join(
            f"{row.name} ({int(row.days)} días)"
            for row in recurrent_absence.head(5).itertuples(index=False)
        )
        add_recommendation(
            "Alta",
            f"{len(recurrent_absence)} persona(s) registran 3 o más ausencias en el mes.",
            "Realizar entrevista de seguimiento, documentar causas y aplicar la política interna que corresponda.",
            scope,
            len(recurrent_absence),
            recurrent_absence["days"].sum(),
            "Ausencias acumuladas por persona ≥ 3 días en el mes.",
        )

    known_departments = incidents[incidents["department"].fillna("").str.strip().ne("")]
    if not known_departments.empty:
        department_counts = known_departments["department"].value_counts()
        top_department = department_counts.index[0]
        top_department_days = int(department_counts.iloc[0])
        department_share = top_department_days / len(known_departments)
        if top_department_days >= 5 and department_share >= 0.30:
            department_rows = known_departments[
                known_departments["department"] == top_department
            ]
            add_recommendation(
                "Media",
                f"{top_department} concentra {department_share:.0%} de los días con departamento identificado.",
                "Revisar con la jefatura del área si existe una causa común y acordar una acción preventiva focalizada.",
                top_department,
                department_rows["employee_id"].nunique(),
                top_department_days,
                "Un departamento concentra ≥ 30% y al menos 5 días de incidencia.",
            )

    type_counts = incidents["incident_type"].value_counts()
    if not type_counts.empty:
        top_type = type_counts.index[0]
        top_type_days = int(type_counts.iloc[0])
        type_share = top_type_days / len(incidents)
        if top_type_days >= 5 and type_share >= 0.40:
            type_rows = incidents[incidents["incident_type"] == top_type]
            add_recommendation(
                "Media",
                f"{top_type} representa {type_share:.0%} de los días de incidencia.",
                "Analizar expedientes y causas recurrentes de este tipo para definir una intervención específica.",
                top_type,
                type_rows["employee_id"].nunique(),
                top_type_days,
                "Un tipo de incidencia concentra ≥ 40% y al menos 5 días del mes.",
            )

    unclassified = incidents[incidents["category"] == "Otra incidencia"]
    if not unclassified.empty:
        incident_types = ", ".join(unclassified["incident_type"].drop_duplicates().head(5))
        add_recommendation(
            "Media",
            "Hay registros fuera de las categorías Ausencia e Incapacidad.",
            "Validar su clasificación para que los indicadores mensuales no oculten permisos u otros conceptos.",
            incident_types,
            unclassified["employee_id"].nunique(),
            len(unclassified),
            "Existe al menos un registro clasificado como Otra incidencia.",
        )

    if not recommendations:
        add_recommendation(
            "Baja",
            "No se activaron alertas con los umbrales actuales.",
            "Mantener el seguimiento mensual y comparar el próximo consolidado contra este periodo.",
            "Consolidado mensual",
            incidents["employee_id"].nunique(),
            len(incidents),
            "Sin incidencias que superen los umbrales definidos.",
        )

    result = pd.DataFrame(recommendations)
    priority_order = {"Alta": 0, "Media": 1, "Baja": 2}
    result["_priority_order"] = result["priority"].map(priority_order)
    result = result.sort_values(
        ["_priority_order", "incident_days", "finding"],
        ascending=[True, False, True],
    ).drop(columns="_priority_order")
    return result[columns].reset_index(drop=True)
