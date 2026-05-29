from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

from .config import AppConfig, ExportGroup, ExportValue, PowerBIReport


@dataclass(frozen=True)
class ExportJob:
    export_key: str
    report: PowerBIReport
    group: ExportGroup | None
    value: ExportValue | None
    report_name: str
    output_filename: str
    pages: list[dict]
    filters: list[str]


def iter_export_jobs(
    config: AppConfig,
    run_date: str | None = None,
    business_date: str | None = None,
    dynamic_values: dict[tuple[str, str], list[ExportValue]] | None = None,
) -> list[ExportJob]:
    rendered_date = run_date or date.today().isoformat()
    rendered_business_date = business_date or rendered_date
    jobs: list[ExportJob] = []
    for report in config.powerbi_reports:
        base_filters = [*report.filters]
        if report.business_date_filter:
            base_filters.append(_render_business_date_filter(report.business_date_filter, rendered_business_date))
        if not report.export_groups:
            jobs.append(
                ExportJob(
                    export_key=report.key,
                    report=report,
                    group=None,
                    value=None,
                    report_name=report.report_name,
                    output_filename=render_filename(
                        report.output_filename,
                        report=report,
                        run_date=rendered_date,
                        business_date=rendered_business_date,
                    ),
                    pages=report.pages,
                    filters=base_filters,
                )
            )
            continue

        for group in report.export_groups:
            values = group.values or (dynamic_values or {}).get((report.key, group.key))
            values = values or [ExportValue(key=group.key, label=group.display_name or group.key)]
            for value in values:
                export_key = f"{report.key}.{group.key}.{value.key}"
                filters = [*base_filters, *group.filters, *value.filters]
                filters.extend(_render_filter_templates(group.filter_templates, value))
                output_template = group.output_filename or report.output_filename
                jobs.append(
                    ExportJob(
                        export_key=export_key,
                        report=report,
                        group=group,
                        value=value,
                        report_name=report.report_name,
                        output_filename=render_filename(
                            output_template,
                            report=report,
                            group=group,
                            value=value,
                            run_date=rendered_date,
                            business_date=rendered_business_date,
                            export_key=export_key,
                        ),
                        pages=[{"pageName": group.page_name}],
                        filters=filters,
                    )
                )
    return jobs


def render_filename(
    template: str,
    report: PowerBIReport,
    run_date: str,
    business_date: str | None = None,
    group: ExportGroup | None = None,
    value: ExportValue | None = None,
    export_key: str | None = None,
) -> str:
    rendered = template.format(
        report_key=report.key,
        export_key=export_key or report.key,
        report_name=report.report_name,
        level_key=group.key if group else "",
        level_name=group.display_name if group and group.display_name else group.key if group else "",
        value_key=value.key if value else "",
        value_label=value.label if value else "",
        value=value.value if value and value.value is not None else value.label if value else "",
        run_date=run_date,
        business_date=business_date or run_date,
    )
    if not rendered.lower().endswith(".pdf"):
        rendered += ".pdf"
    return _safe_filename(rendered)


def _render_filter_templates(templates: list[str], value: ExportValue) -> list[str]:
    rendered = []
    raw_value = value.value if value.value is not None else value.label
    for template in templates:
        rendered.append(
            template.format(
                value=escape_filter_value(raw_value),
                value_raw=raw_value,
                value_key=value.key,
                value_label=value.label,
            )
        )
    return rendered


def _safe_filename(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value)
    value = re.sub(r"\s+", " ", value).strip()
    if not value:
        raise ValueError("Rendered output filename is empty.")
    return value


def value_to_key(value: str, prefix: str | None = None) -> str:
    key = value.lower().strip()
    key = re.sub(r"[^a-z0-9]+", "_", key)
    key = key.strip("_") or "blank"
    return f"{prefix}_{key}" if prefix else key


def escape_filter_value(value: str) -> str:
    return value.replace("'", "''")


def _render_business_date_filter(date_filter, business_date: str) -> str:
    return date_filter.template.format(
        table=date_filter.table,
        column=date_filter.column,
        business_date=business_date,
        business_date_datetime=f"datetime'{business_date}T00:00:00'",
    )
