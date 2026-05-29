#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from report_ops_automation.auth import GRAPH_APP_SCOPE, POWERBI_SCOPE, get_app_token
from report_ops_automation.config import load_config
from report_ops_automation.env import load_dotenv, require_first_env
from report_ops_automation.exports import ExportValue, iter_export_jobs
from report_ops_automation.powerbi import PowerBIClient
from report_ops_automation.sharepoint import SharePointClient


def main() -> None:
    parser = argparse.ArgumentParser(description="Export Power BI PDFs and upload them to SharePoint.")
    parser.add_argument("--config", default="config/reports.yaml")
    parser.add_argument("--run-date", default=date.today().isoformat())
    parser.add_argument(
        "--business-date",
        default=(date.today() - timedelta(days=1)).isoformat(),
        help="Business date filter in YYYY-MM-DD format. Defaults to yesterday.",
    )
    parser.add_argument(
        "--export-key",
        action="append",
        default=[],
        help="Export only this generated key. Can be passed more than once.",
    )
    parser.add_argument("--max-jobs", type=int, default=0, help="Stop after this many export jobs.")
    args = parser.parse_args()

    load_dotenv()
    config = load_config(args.config)

    powerbi_tenant_id = require_first_env("POWERBI_TENANT_ID", "AZURE_TENANT_ID")
    powerbi_client_id = require_first_env("POWERBI_CLIENT_ID", "AZURE_CLIENT_ID")
    powerbi_client_secret = require_first_env("POWERBI_CLIENT_SECRET", "AZURE_CLIENT_SECRET")

    sharepoint_tenant_id = require_first_env("SHAREPOINT_TENANT_ID", "AZURE_TENANT_ID")
    sharepoint_client_id = require_first_env("SHAREPOINT_CLIENT_ID", "AZURE_CLIENT_ID")
    sharepoint_client_secret = require_first_env("SHAREPOINT_CLIENT_SECRET", "AZURE_CLIENT_SECRET")

    powerbi = PowerBIClient(
        get_app_token(powerbi_tenant_id, powerbi_client_id, powerbi_client_secret, POWERBI_SCOPE)
    )
    sharepoint = SharePointClient.from_config(
        get_app_token(sharepoint_tenant_id, sharepoint_client_id, sharepoint_client_secret, GRAPH_APP_SCOPE),
        config.sharepoint,
    )

    jobs = select_jobs(config, powerbi, args.run_date, args.business_date, set(args.export_key))
    if args.max_jobs:
        jobs = jobs[: args.max_jobs]
    if not jobs:
        raise ValueError("No export jobs matched the provided selector.")

    for job in jobs:
        print(f"Exporting {job.export_key}...")
        for filter_value in job.filters:
            print(f"  filter: {filter_value}")
        pdf = powerbi.export_report_pdf(job)
        item = sharepoint.upload_file(sharepoint.output_folder, job.output_filename, pdf)
        print(f"Uploaded {job.output_filename}: {item.get('webUrl', item.get('id'))}")


def select_jobs(config, powerbi: PowerBIClient, run_date: str, business_date: str, export_keys: set[str]):
    static_jobs = iter_export_jobs(config, run_date, business_date)
    if not export_keys:
        dynamic_values = discover_dynamic_values(config, powerbi)
        return iter_export_jobs(config, run_date, business_date, dynamic_values)

    static_by_key = {job.export_key: job for job in static_jobs}
    if export_keys <= set(static_by_key):
        return [static_by_key[key] for key in sorted(export_keys)]

    dynamic_values = discover_dynamic_values(config, powerbi)
    dynamic_by_key = {
        job.export_key: job for job in iter_export_jobs(config, run_date, business_date, dynamic_values)
    }
    missing = export_keys - set(dynamic_by_key)
    if missing:
        raise ValueError(f"Unknown export_key(s): {', '.join(sorted(missing))}")
    return [dynamic_by_key[key] for key in sorted(export_keys)]


def discover_dynamic_values(config, powerbi: PowerBIClient) -> dict[tuple[str, str], list[ExportValue]]:
    values: dict[tuple[str, str], list[ExportValue]] = {}
    for report in config.powerbi_reports:
        for group in report.export_groups:
            if group.values or not group.values_from:
                continue
            discovered = powerbi.get_distinct_values(report, group)
            values[(report.key, group.key)] = discovered
            print(f"Discovered {len(discovered)} values for {report.key}.{group.key}")
    return values


if __name__ == "__main__":
    main()
