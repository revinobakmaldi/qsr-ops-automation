from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass(frozen=True)
class SharePointConfig:
    drive_id: str | None = None
    output_folder: str | None = None
    folder_url: str | None = None


@dataclass(frozen=True)
class ExportValue:
    key: str
    label: str
    value: str | None = None
    filters: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ValueSource:
    table: str
    column: str
    key_prefix: str | None = None


@dataclass(frozen=True)
class DateFilter:
    table: str
    column: str
    template: str = "{table}/{column} eq {business_date}"
    value: str = "business_date"
    slicer_name: str | None = None


@dataclass(frozen=True)
class DateFilters:
    daily: DateFilter | None = None
    weekly: DateFilter | None = None
    monthly: DateFilter | None = None


@dataclass(frozen=True)
class ExportGroup:
    key: str
    page_name: str
    display_name: str | None = None
    output_filename: str | None = None
    filter_templates: list[str] = field(default_factory=list)
    filters: list[str] = field(default_factory=list)
    values: list[ExportValue] = field(default_factory=list)
    values_from: ValueSource | None = None


@dataclass(frozen=True)
class PowerBIReport:
    key: str
    report_name: str
    workspace_id: str
    report_id: str
    output_filename: str
    dataset_id: str | None = None
    filters: list[str] = field(default_factory=list)
    pages: list[dict[str, Any]] = field(default_factory=list)
    bookmark_state: str | None = None
    locale: str | None = None
    business_date_filter: DateFilter | None = None
    date_filters: DateFilters | None = None
    export_groups: list[ExportGroup] = field(default_factory=list)


@dataclass(frozen=True)
class ReportDelivery:
    recipient_upn: str
    message: str
    channel: str = "email"
    subject: str = "{report_name} PDF is ready"
    report_key: str | None = None
    export_key: str | None = None


@dataclass(frozen=True)
class AppConfig:
    sharepoint: SharePointConfig
    powerbi_reports: list[PowerBIReport]
    report_delivery: list[ReportDelivery]


def load_config(path: str | Path) -> AppConfig:
    raw = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("Config file must contain a YAML mapping.")

    sharepoint = raw.get("sharepoint") or {}
    reports = raw.get("powerbi_reports") or []
    deliveries = raw.get("report_delivery") or []

    return AppConfig(
        sharepoint=_parse_sharepoint(sharepoint),
        powerbi_reports=[_parse_report(report) for report in reports],
        report_delivery=[ReportDelivery(**delivery) for delivery in deliveries],
    )


def _required(data: dict[str, Any], label: str, key: str) -> str:
    value = data.get(key)
    if not value:
        raise ValueError(f"Missing required config value: {label}")
    return str(value)


def _parse_sharepoint(data: dict[str, Any]) -> SharePointConfig:
    if data.get("folder_url"):
        return SharePointConfig(folder_url=str(data["folder_url"]))
    return SharePointConfig(
        drive_id=_required(data, "sharepoint.drive_id", "drive_id"),
        output_folder=_required(data, "sharepoint.output_folder", "output_folder").strip("/"),
    )


def _parse_report(data: dict[str, Any]) -> PowerBIReport:
    groups = [
        ExportGroup(
            key=group["key"],
            page_name=group["page_name"],
            display_name=group.get("display_name"),
            output_filename=group.get("output_filename"),
            filter_templates=group.get("filter_templates", []),
            filters=group.get("filters", []),
            values=[ExportValue(**value) for value in group.get("values", [])],
            values_from=ValueSource(**group["values_from"]) if group.get("values_from") else None,
        )
        for group in data.get("export_groups", [])
    ]
    report_data = {
        key: value
        for key, value in data.items()
        if key not in {"export_groups", "business_date_filter", "date_filters"}
    }
    business_date_filter = (
        DateFilter(**data["business_date_filter"]) if data.get("business_date_filter") else None
    )
    date_filters = (
        DateFilters(
            daily=DateFilter(**data["date_filters"]["daily"]) if data["date_filters"].get("daily") else None,
            weekly=DateFilter(**data["date_filters"]["weekly"]) if data["date_filters"].get("weekly") else None,
            monthly=DateFilter(**data["date_filters"]["monthly"]) if data["date_filters"].get("monthly") else None,
        )
        if data.get("date_filters")
        else None
    )
    return PowerBIReport(
        **report_data,
        business_date_filter=business_date_filter,
        date_filters=date_filters,
        export_groups=groups,
    )
