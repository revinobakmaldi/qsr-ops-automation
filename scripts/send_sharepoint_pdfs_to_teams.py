#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from report_ops_automation.auth import GRAPH_APP_SCOPE, POWERBI_SCOPE, get_app_token, get_delegated_graph_token
from report_ops_automation.config import load_config
from report_ops_automation.email import EmailClient, format_email_message, format_email_subject
from report_ops_automation.env import load_dotenv, require_env, require_first_env
from report_ops_automation.exports import ExportValue, iter_export_jobs
from report_ops_automation.powerbi import PowerBIClient
from report_ops_automation.sharepoint import SharePointClient
from report_ops_automation.teams import TeamsClient, format_delivery_message


def main() -> None:
    parser = argparse.ArgumentParser(description="Send SharePoint PDF report links to mapped recipients.")
    parser.add_argument("--config", default="config/reports.yaml")
    parser.add_argument("--run-date", default=date.today().isoformat())
    parser.add_argument(
        "--business-date",
        default=(date.today() - timedelta(days=1)).isoformat(),
        help="Business date filter in YYYY-MM-DD format. Defaults to yesterday.",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    load_dotenv()
    config = load_config(args.config)

    powerbi_tenant_id = require_env("POWERBI_TENANT_ID")
    powerbi_client_id = require_env("POWERBI_CLIENT_ID")
    powerbi_client_secret = require_env("POWERBI_CLIENT_SECRET")
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

    dynamic_values = discover_dynamic_values(config, powerbi)
    jobs_by_export_key = {
        job.export_key: job for job in iter_export_jobs(config, args.run_date, args.business_date, dynamic_values)
    }
    jobs_by_report_key: dict[str, list] = {}
    for job in jobs_by_export_key.values():
        jobs_by_report_key.setdefault(job.report.key, []).append(job)
    files_by_name = {
        item["name"]: item
        for item in sharepoint.list_folder(sharepoint.output_folder)
        if item.get("file") and item.get("name", "").lower().endswith(".pdf")
    }

    delivery_clients = DeliveryClients(config.report_delivery)

    for delivery in config.report_delivery:
        job = None
        if delivery.export_key:
            job = jobs_by_export_key.get(delivery.export_key)
            if not job:
                raise ValueError(f"Delivery references unknown export_key: {delivery.export_key}")
        elif delivery.report_key:
            report_jobs = jobs_by_report_key.get(delivery.report_key, [])
            if not report_jobs:
                raise ValueError(f"Delivery references unknown report_key: {delivery.report_key}")
            if len(report_jobs) > 1:
                raise ValueError(
                    f"Delivery report_key {delivery.report_key} matches multiple PDFs. "
                    "Use export_key instead."
                )
            job = report_jobs[0]
        else:
            raise ValueError("Delivery must include either export_key or report_key.")

        item = files_by_name.get(job.output_filename)
        if not item:
            available = ", ".join(sorted(files_by_name)) or "no PDFs found"
            raise FileNotFoundError(
                f"Could not find {job.output_filename} in SharePoint folder. Available: {available}"
            )

        if args.dry_run:
            print(
                f"Would send {job.output_filename} to {delivery.recipient_upn} "
                f"via {delivery.channel}: {item['webUrl']}"
            )
            continue

        delivery_clients.send(delivery, job.report_name, job.output_filename, item["webUrl"])
        print(f"Sent {job.output_filename} to {delivery.recipient_upn} via {delivery.channel}")


class DeliveryClients:
    def __init__(self, deliveries):
        channels = {delivery.channel.lower() for delivery in deliveries}
        self.email: EmailClient | None = None
        self.teams: tuple[TeamsClient, str] | None = None

        if "email" in channels:
            email_tenant_id = require_first_env("EMAIL_TENANT_ID", "SHAREPOINT_TENANT_ID", "AZURE_TENANT_ID")
            email_client_id = require_first_env("EMAIL_CLIENT_ID", "SHAREPOINT_CLIENT_ID", "AZURE_CLIENT_ID")
            email_client_secret = require_first_env(
                "EMAIL_CLIENT_SECRET",
                "SHAREPOINT_CLIENT_SECRET",
                "AZURE_CLIENT_SECRET",
            )
            email_sender = require_env("EMAIL_SENDER_UPN")
            self.email = EmailClient(
                get_app_token(email_tenant_id, email_client_id, email_client_secret, GRAPH_APP_SCOPE),
                email_sender,
            )

        if "teams" in channels:
            tenant_id = require_env("AZURE_TENANT_ID")
            delegated_client_id = require_env("GRAPH_DELEGATED_CLIENT_ID")
            sender_upn = require_env("TEAMS_SENDER_UPN")
            graph_token = get_delegated_graph_token(tenant_id, delegated_client_id)
            self.teams = (TeamsClient(graph_token), sender_upn)

        unsupported = channels - {"email", "teams"}
        if unsupported:
            raise ValueError(f"Unsupported delivery channel(s): {', '.join(sorted(unsupported))}")

    def send(self, delivery, report_name: str, file_name: str, web_url: str) -> None:
        channel = delivery.channel.lower()
        if channel == "email":
            if not self.email:
                raise RuntimeError("Email client was not initialized.")
            self.email.send_message(
                delivery.recipient_upn,
                format_email_subject(delivery.subject, report_name, file_name),
                format_email_message(delivery.message, report_name, file_name, web_url),
            )
            return

        if channel == "teams":
            if not self.teams:
                raise RuntimeError("Teams client was not initialized.")
            teams, sender_upn = self.teams
            message = format_delivery_message(delivery.message, report_name, file_name, web_url)
            chat_id = teams.create_one_on_one_chat(sender_upn, delivery.recipient_upn)
            teams.send_message(chat_id, message)
            return

        raise ValueError(f"Unsupported delivery channel: {delivery.channel}")


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
