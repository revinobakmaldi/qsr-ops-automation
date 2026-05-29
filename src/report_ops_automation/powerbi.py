from __future__ import annotations

import time
from typing import Any

from .config import ExportGroup, ExportValue, PowerBIReport
from .exports import ExportJob
from .http import ApiClient


class PowerBIClient:
    def __init__(self, token: str):
        self.api = ApiClient(token, "https://api.powerbi.com/v1.0/myorg")

    def export_report_pdf(self, job: ExportJob, poll_seconds: int = 5, timeout_seconds: int = 900) -> bytes:
        report = job.report
        payload = _export_payload(job, self._page_name_map(report))
        export_job = self.api.post_json(
            f"/groups/{report.workspace_id}/reports/{report.report_id}/ExportTo",
            payload,
        )
        export_id = export_job["id"]
        status_path = f"/groups/{report.workspace_id}/reports/{report.report_id}/exports/{export_id}"

        started = time.monotonic()
        while time.monotonic() - started < timeout_seconds:
            status = self.api.get_json(status_path)
            state = status.get("status")
            if state == "Succeeded":
                response = self.api.request("GET", f"{status_path}/file")
                return response.content
            if state == "Failed":
                raise RuntimeError(f"Power BI export failed for {job.export_key}: {status}")
            print(f"{job.export_key}: export {state} ({status.get('percentComplete', 0)}%)")
            time.sleep(poll_seconds)

        raise TimeoutError(f"Timed out waiting for Power BI export: {job.export_key}")

    def get_report_dataset_id(self, report: PowerBIReport) -> str:
        payload = self.api.get_json(f"/groups/{report.workspace_id}/reports/{report.report_id}")
        dataset_id = payload.get("datasetId")
        if not dataset_id:
            raise RuntimeError(f"Power BI report {report.key} did not return a datasetId.")
        return dataset_id

    def get_distinct_values(self, report: PowerBIReport, group: ExportGroup) -> list[ExportValue]:
        if not group.values_from:
            return []
        source = group.values_from
        dataset_id = self.get_report_dataset_id(report)
        query = _distinct_values_dax(source.table, source.column)
        payload = self.api.post_json(
            f"/groups/{report.workspace_id}/datasets/{dataset_id}/executeQueries",
            {
                "queries": [{"query": query}],
                "serializerSettings": {"includeNulls": False},
            },
        )
        rows = (
            payload.get("results", [{}])[0]
            .get("tables", [{}])[0]
            .get("rows", [])
        )
        values = []
        for row in rows:
            value = str(row.get("[Value]", row.get("Value", ""))).strip()
            if value:
                values.append(
                    ExportValue(
                        key=_value_to_key(value, source.key_prefix),
                        label=value,
                        value=value,
                    )
                )
        return values

    def _page_name_map(self, report: PowerBIReport) -> dict[str, str]:
        pages = self.api.get_json(f"/groups/{report.workspace_id}/reports/{report.report_id}/pages").get("value", [])
        mapping: dict[str, str] = {}
        for page in pages:
            name = page.get("name")
            display_name = page.get("displayName")
            if name:
                mapping[name] = name
            if display_name and name:
                mapping[display_name] = name
        return mapping


def _export_payload(job: ExportJob, page_name_map: dict[str, str] | None = None) -> dict[str, Any]:
    report = job.report
    config: dict[str, Any] = {
        "settings": {"includeHiddenPages": False},
    }
    if job.filters:
        config["reportLevelFilters"] = [{"filter": value} for value in job.filters]
    if job.pages:
        config["pages"] = [_resolve_page(page, page_name_map or {}) for page in job.pages]
    if report.bookmark_state:
        config["defaultBookmark"] = {"state": report.bookmark_state}
    if report.locale:
        config["settings"]["locale"] = report.locale

    return {
        "format": "PDF",
        "powerBIReportConfiguration": config,
    }


def _resolve_page(page: dict[str, Any], page_name_map: dict[str, str]) -> dict[str, Any]:
    page_name = page.get("pageName")
    if not page_name:
        return page
    resolved = page_name_map.get(page_name)
    if not resolved:
        available = ", ".join(sorted(page_name_map)) or "no pages returned"
        raise ValueError(f"Could not resolve Power BI page '{page_name}'. Available pages: {available}")
    return {**page, "pageName": resolved}


def _distinct_values_dax(table: str, column: str) -> str:
    table_ref = table.replace("'", "''")
    column_ref = column.replace("]", "]]")
    full_ref = f"'{table_ref}'[{column_ref}]"
    return f"""
EVALUATE
SELECTCOLUMNS(
    FILTER(VALUES({full_ref}), NOT ISBLANK({full_ref})),
    "Value", {full_ref}
)
ORDER BY [Value]
""".strip()


def _value_to_key(value: str, prefix: str | None = None) -> str:
    from .exports import value_to_key

    return value_to_key(value, prefix)
