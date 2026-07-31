from __future__ import annotations

from datetime import date

import pandas as pd

def to_frame(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows)


def apply_filters(
    df: pd.DataFrame,
    building: str,
    department: str,
    period: str,
    search: str,
    start_date: date | None = None,
    end_date: date | None = None,
) -> pd.DataFrame:
    if df.empty:
        return df
    filtered = df.copy()
    if building != "Todos":
        filtered = filtered[filtered["building"] == building]
    if department != "Todos":
        filtered = filtered[filtered["department"] == department]
    if period != "Todos":
        filtered = filtered[filtered["period"] == period]
    if "date" in filtered.columns and (start_date or end_date):
        row_dates = pd.to_datetime(filtered["date"], errors="coerce").dt.date
        if start_date:
            filtered = filtered[row_dates >= start_date]
            row_dates = pd.to_datetime(filtered["date"], errors="coerce").dt.date
        if end_date:
            filtered = filtered[row_dates <= end_date]
    if search:
        needle = search.lower()
        filtered = filtered[
            filtered["name"].str.lower().str.contains(needle, na=False)
            | filtered["employee_id"].astype(str).str.lower().str.contains(needle, na=False)
        ]
    return filtered


def summarize_incidents(incidents: pd.DataFrame) -> pd.DataFrame:
    """Condensa todas las incidencias de una persona en una sola fila."""
    columns = [
        "employee_id", "name", "department", "schedule_types", "schedule_sources", "building", "incident_count",
        "incident_types", "dates", "total_lates", "bonus_lost",
    ]
    if incidents.empty:
        return pd.DataFrame(columns=columns)

    def unique_join(values) -> str:
        return ", ".join(dict.fromkeys(str(value) for value in values if pd.notna(value) and str(value)))

    summary = (
        incidents.sort_values(["name", "date"])
        .groupby(["employee_id", "name", "department", "building"], as_index=False, dropna=False)
        .agg(
            incident_count=("incident_type", "size"),
            schedule_types=("schedule_type", unique_join),
            schedule_sources=("schedule_source", unique_join),
            incident_types=("incident_type", unique_join),
            dates=("date", unique_join),
            total_lates=("incident_type", lambda values: int((values == "Retardo").sum())),
            bonus_lost=("bonus_lost", lambda values: "Si" if (values == "Si").any() else "No"),
        )
    )
    return summary[columns]


def recalculate_period_penalties(incidents: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Recalcula retardos y penalizaciones usando únicamente el rango filtrado."""
    penalty_columns = [
        "employee_id", "name", "department", "building", "period",
        "late_count", "equivalent_absence", "bonus_lost",
    ]
    if incidents.empty:
        return incidents.copy(), pd.DataFrame(columns=penalty_columns)

    adjusted = incidents.copy()
    late_counts = (
        adjusted[adjusted["incident_type"] == "Retardo"]
        .groupby(["employee_id", "period"])
        .size()
    )
    adjusted["period_late_count"] = [
        int(late_counts.get((employee_id, period), 0))
        for employee_id, period in zip(adjusted["employee_id"], adjusted["period"])
    ]
    adjusted["bonus_lost"] = adjusted["period_late_count"].ge(2).map({True: "Si", False: "No"})

    penalties = (
        adjusted[adjusted["period_late_count"] >= 2]
        .drop_duplicates(["employee_id", "period"])
        [["employee_id", "name", "department", "building", "period", "period_late_count"]]
        .rename(columns={"period_late_count": "late_count"})
    )
    penalties["equivalent_absence"] = 1
    penalties["bonus_lost"] = "Si"
    return adjusted, penalties[penalty_columns]
