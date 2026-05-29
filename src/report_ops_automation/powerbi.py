from __future__ import annotations

import base64
import json
import time
from typing import Any

from .config import ExportGroup, ExportValue, PowerBIReport
from .exports import ExportJob, SlicerOverride
from .http import ApiClient


class PowerBIClient:
    def __init__(self, token: str):
        self.api = ApiClient(token, "https://api.powerbi.com/v1.0/myorg")

    def export_report_pdf(
        self,
        job: ExportJob,
        poll_seconds: int = 5,
        timeout_seconds: int = 900,
        filter_level: str = "report",
        debug: bool = False,
    ) -> bytes:
        report = job.report
        page_name_map = self._page_name_map(report)
        payload = _export_payload(job, page_name_map, filter_level=filter_level)
        if debug:
            import json
            print(f"[debug] ExportTo payload:\n{json.dumps(payload, indent=2)}")
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
        if report.dataset_id:
            return report.dataset_id
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
            f"/datasets/{dataset_id}/executeQueries",
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

    def list_bookmarks(self, report: PowerBIReport) -> list[dict]:
        return self.api.get_json(
            f"/groups/{report.workspace_id}/reports/{report.report_id}/bookmarks"
        ).get("value", [])

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


def _export_payload(
    job: ExportJob,
    page_name_map: dict[str, str] | None = None,
    filter_level: str = "report",
) -> dict[str, Any]:
    report = job.report
    config: dict[str, Any] = {
        "settings": {"includeHiddenPages": False},
    }
    if report.dataset_id:
        config["datasetToBind"] = report.dataset_id

    resolved_pages = None
    if job.pages:
        resolved_pages = [_resolve_page(page, page_name_map or {}) for page in job.pages]
        config["pages"] = resolved_pages

    if job.filters:
        filter_expr = " and ".join(job.filters)
        if filter_level == "page" and resolved_pages:
            config["pageLevelFilters"] = [
                {"pageName": p["pageName"], "filter": filter_expr}
                for p in resolved_pages
            ]
        else:
            config["reportLevelFilters"] = [{"filter": filter_expr}]

    if job.captured_bookmark_state:
        config["defaultBookmark"] = {"state": job.captured_bookmark_state}
    elif job.slicer_overrides:
        page_ids = [p["pageName"] for p in resolved_pages] if resolved_pages else list(set((page_name_map or {}).values()))
        config["defaultBookmark"] = {"state": _build_slicer_state(job.slicer_overrides, page_ids)}
    elif report.bookmark_state:
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


def _build_slicer_state(slicer_overrides: list[SlicerOverride], page_ids: list[str]) -> str:
    visual_containers = {
        override.slicer_name: {
            "filters": json.dumps([{
                "$schema": "http://powerbi.com/product/schema#basic",
                "target": {"table": override.table, "column": override.column},
                "operator": "In",
                "values": [override.value],
                "filterType": 1,
                "requireSingleSelection": False,
            }]),
            "singleVisualGroup": {},
        }
        for override in slicer_overrides
    }
    active = page_ids[0] if page_ids else ""
    state = {
        "explorationState": {
            "activeSection": active,
            "sections": {
                page_id: {"visualContainers": visual_containers}
                for page_id in page_ids
            },
        }
    }
    return base64.b64encode(json.dumps(state).encode()).decode()


def _value_to_key(value: str, prefix: str | None = None) -> str:
    from .exports import value_to_key

    return value_to_key(value, prefix)
