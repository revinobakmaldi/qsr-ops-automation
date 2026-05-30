#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import subprocess
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import replace
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
    parser.add_argument(
        "--export-key-file",
        default=None,
        help="Path to a file with one export key per line. Merged with --export-key. Use to retry a failed_*.txt file.",
    )
    parser.add_argument("--max-jobs", type=int, default=0, help="Stop after this many export jobs.")
    parser.add_argument(
        "--filter-mode",
        choices=["all", "value-only", "value-in", "value-url-encoded", "date-only", "none"],
        default="all",
        help="Limit filters for troubleshooting exports.",
    )
    parser.add_argument(
        "--filter-level",
        choices=["report", "page"],
        default="report",
        help="Apply filters at report level (default) or page level via pageLevelFilters.",
    )
    parser.add_argument("--workers", type=int, default=1, help="Number of parallel export workers.")
    parser.add_argument(
        "--on-missing-date",
        choices=["fail", "latest"],
        default="fail",
        help="What to do when a daily/weekly/monthly slicer value is not found in the dataset. "
             "'fail' aborts before exporting. 'latest' substitutes the most recent available value.",
    )
    parser.add_argument("--render-wait", type=float, default=4.0, help="Seconds to wait after render event before generating PDF. Default 4s; reduce to 2s if no HTML visuals.")
    parser.add_argument("--node-script", default=None, help="Path to capture_slicer_state.js. Required for slicer-based date filtering.")
    parser.add_argument("--debug", action="store_true", help="Print the ExportTo payload before sending.")
    parser.add_argument("--list-values", action="store_true", help="Discover and print actual column values from dataset, then exit.")
    parser.add_argument("--list-bookmarks", action="store_true", help="List saved bookmarks on each report, then exit.")
    args = parser.parse_args()

    load_dotenv()
    config = load_config(args.config)

    powerbi_tenant_id = require_first_env("POWERBI_TENANT_ID", "AZURE_TENANT_ID")
    powerbi_client_id = require_first_env("POWERBI_CLIENT_ID", "AZURE_CLIENT_ID")
    powerbi_client_secret = require_first_env("POWERBI_CLIENT_SECRET", "AZURE_CLIENT_SECRET")

    sharepoint_tenant_id = require_first_env("SHAREPOINT_TENANT_ID", "AZURE_TENANT_ID")
    sharepoint_client_id = require_first_env("SHAREPOINT_CLIENT_ID", "AZURE_CLIENT_ID")
    sharepoint_client_secret = require_first_env("SHAREPOINT_CLIENT_SECRET", "AZURE_CLIENT_SECRET")

    pbi_token_early = get_app_token(powerbi_tenant_id, powerbi_client_id, powerbi_client_secret, POWERBI_SCOPE)
    sp_token_early = get_app_token(sharepoint_tenant_id, sharepoint_client_id, sharepoint_client_secret, GRAPH_APP_SCOPE)
    powerbi = PowerBIClient(pbi_token_early)

    if args.list_values:
        _print_column_values(config, powerbi)
        return

    if args.list_bookmarks:
        _print_bookmarks(config, powerbi)
        return

    export_keys = list(args.export_key)
    if args.export_key_file:
        with open(args.export_key_file) as f:
            export_keys += [line.strip() for line in f if line.strip()]

    jobs = select_jobs(config, powerbi, args.run_date, args.business_date, set(export_keys))
    if args.max_jobs:
        jobs = jobs[: args.max_jobs]
    if not jobs:
        raise ValueError("No export jobs matched the provided selector.")

    jobs = [apply_filter_mode(job, args.filter_mode) for job in jobs]
    jobs = _validate_slicer_dates(jobs, powerbi, args.on_missing_date)

    pdf_script = Path(args.node_script).parent / "export_report_pdf.js" if args.node_script else None
    use_puppeteer = pdf_script and pdf_script.exists()

    if use_puppeteer:
        failed = _run_puppeteer_exports(
            jobs, str(pdf_script), config.sharepoint, sp_token_early,
            workers=args.workers,
            render_wait_ms=int(args.render_wait * 1000),
        )
        if failed:
            from datetime import datetime
            fname = f"failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            Path(fname).write_text("\n".join(failed) + "\n")
            print(f"{len(failed)} job(s) failed. Keys saved to {fname} — retry with --export-key-file {fname}")
            raise SystemExit(1)
        return

    if args.node_script:
        jobs = _inject_captured_states(jobs, args.node_script)

    def run_job(job):
        pbi = PowerBIClient(pbi_token_early)
        sp = SharePointClient.from_config(sp_token_early, config.sharepoint)
        print(f"Exporting {job.export_key}...")
        for f in job.filters:
            print(f"  filter: {f}")
        pdf = pbi.export_report_pdf(job, filter_level=args.filter_level, debug=args.debug)
        item = sp.upload_file(sp.output_folder, job.output_filename, pdf)
        url = item.get("webUrl", item.get("id"))
        print(f"Uploaded {job.output_filename}: {url}")
        return job.export_key, url

    failed = []
    if args.workers <= 1:
        for job in jobs:
            try:
                run_job(job)
            except Exception as e:
                print(f"FAILED {job.export_key}: {e}")
                failed.append(job.export_key)
    else:
        print(f"Running {len(jobs)} jobs with {args.workers} parallel workers...")
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = {executor.submit(run_job, job): job for job in jobs}
            for future in as_completed(futures):
                job = futures[future]
                try:
                    future.result()
                except Exception as e:
                    print(f"FAILED {job.export_key}: {e}")
                    failed.append(job.export_key)
    if failed:
        raise RuntimeError(f"{len(failed)} job(s) failed: {', '.join(failed)}")


def apply_filter_mode(job, filter_mode: str):
    if filter_mode == "all":
        return job
    if filter_mode == "none":
        return replace_job_filters(job, [], slicer_overrides=[])
    if filter_mode == "value-only":
        return replace_job_filters(job, job.filters[-1:] if job.value else [])
    if filter_mode == "value-in":
        return replace_job_filters(job, [eq_filter_to_in(job.filters[-1])] if job.value else [])
    if filter_mode == "value-url-encoded":
        return replace_job_filters(job, [url_encode_filter_value(job.filters[-1])] if job.value else [])
    if filter_mode == "date-only":
        return replace_job_filters(job, job.filters[:3])
    raise ValueError(f"Unsupported filter_mode: {filter_mode}")


def replace_job_filters(job, filters, slicer_overrides=None):
    kwargs = {"filters": filters}
    if slicer_overrides is not None:
        kwargs["slicer_overrides"] = slicer_overrides
    return replace(job, **kwargs)



def _latest_slicer_value(powerbi, report, table, column, current_value: str) -> str | None:
    """Return the most recent available value for a slicer column.
    Datetime columns are sorted by the datetime value (DESC in DAX).
    Text columns (e.g. WeekMonthYear) are parsed and sorted chronologically in Python.
    """
    is_datetime = "T" in current_value and len(current_value) >= 10
    limit = 1 if is_datetime else 104  # 2 years of weeks as safety margin
    values = powerbi.get_column_values(report, table, column, limit=limit)
    if not values:
        return None
    if is_datetime:
        return values[0]  # DAX TOPN DESC already gives most recent datetime first

    # Text column — sort chronologically by parsing "Week N Mon YYYY"
    def _week_sort_key(label: str):
        parts = label.split()
        if len(parts) == 4 and parts[0] == "Week":
            try:
                from datetime import datetime as _dt
                month = _dt.strptime(parts[2], "%b").month
                return (int(parts[3]), month, int(parts[1]))
            except (ValueError, IndexError):
                pass
        return (0, 0, 0)

    return max(values, key=_week_sort_key)


def _validate_slicer_dates(jobs: list, powerbi, behavior: str) -> list:
    """Check each report's date slicer values exist in the dataset before exporting.
    behavior='fail'   → raise ValueError listing which values are missing.
    behavior='latest' → substitute the most recent available value and continue.
    """
    from dataclasses import replace as dc_replace

    checked: dict[tuple, str | None] = {}  # (report_key, table, column, value) → resolved

    for job in jobs:
        for slicer in job.slicer_overrides:
            key = (job.report.key, slicer.table, slicer.column, slicer.value)
            if key in checked:
                continue
            exists = powerbi.value_exists_in_column(job.report, slicer.table, slicer.column, slicer.value)
            if exists:
                checked[key] = slicer.value
            elif behavior == "latest":
                latest = _latest_slicer_value(powerbi, job.report, slicer.table, slicer.column, slicer.value)
                if latest:
                    print(
                        f"WARNING: {slicer.table}/{slicer.column} = '{slicer.value}' not in dataset. "
                        f"Using latest: '{latest}'"
                    )
                    checked[key] = latest
                else:
                    checked[key] = None
            else:
                checked[key] = None

    missing = [(k, v) for k, v in checked.items() if v is None]
    if missing:
        details = "\n".join(f"  {k[1]}/{k[2]} = '{k[3]}'" for k, _ in missing)
        raise ValueError(
            f"Date slicer value(s) not found in dataset. "
            f"The data pipeline may not have run yet for this date.\n{details}\n"
            f"Rerun with --on-missing-date latest to use the most recent available value instead."
        )

    def resolve_slicers(job):
        new_slicers = [
            dc_replace(s, value=checked[(job.report.key, s.table, s.column, s.value)])
            if checked.get((job.report.key, s.table, s.column, s.value)) != s.value
            else s
            for s in job.slicer_overrides
        ]
        return dc_replace(job, slicer_overrides=new_slicers)

    return [resolve_slicers(job) for job in jobs]


def url_encode_filter_value(filter_value: str) -> str:
    from urllib.parse import quote

    prefix, sep, raw_value = filter_value.partition(" eq ")
    if sep != " eq " or not (raw_value.startswith("'") and raw_value.endswith("'")):
        return filter_value
    encoded = quote(raw_value[1:-1], safe="")
    return f"{prefix} eq '{encoded}'"


def eq_filter_to_in(filter_value: str) -> str:
    prefix, sep, raw_value = filter_value.partition(" eq ")
    if sep != " eq " or not (raw_value.startswith("'") and raw_value.endswith("'")):
        return filter_value
    return f"{prefix} in ({raw_value})"


def _run_puppeteer_exports(
    jobs, pdf_script: str, sharepoint_config, sp_token: str,
    workers: int = 1, render_wait_ms: int = 4000,
) -> None:
    import os
    from report_ops_automation.sharepoint import SharePointClient

    by_report: dict[str, list] = {}
    for job in jobs:
        by_report.setdefault(job.report.key, []).append(job)

    all_failed: list[str] = []

    for report_key, report_jobs in by_report.items():
        n = min(workers, len(report_jobs))
        chunk_size = math.ceil(len(report_jobs) / n)
        chunks = [report_jobs[i:i + chunk_size] for i in range(0, len(report_jobs), chunk_size)]
        print(f"Generating {len(report_jobs)} PDFs via Puppeteer for {report_key} ({n} worker(s))...")

        def run_chunk(chunk, worker_id):
            tag = f"[w{worker_id}]"
            first = chunk[0]
            node_config = {
                "workspaceId": first.report.workspace_id,
                "reportId": first.report.report_id,
                "renderWait": render_wait_ms,
                "slicers": [
                    {"table": s.table, "column": s.column, "value": s.value}
                    for s in first.slicer_overrides
                ],
                "jobs": [],
            }

            tmp_files = {}
            for job in chunk:
                parsed_filters = []
                for f in job.filters:
                    parts = f.split(" eq ")
                    if len(parts) == 2:
                        table_col = parts[0].strip()
                        value = parts[1].strip().strip("'")
                        if "/" in table_col:
                            table, col = table_col.split("/", 1)
                            parsed_filters.append({"table": table, "column": col, "values": [value]})
                tmp = tempfile.mktemp(suffix=".pdf")
                tmp_files[job.export_key] = (job, tmp)
                node_config["jobs"].append({
                    "exportKey": job.export_key,
                    "pageName": job.pages[0]["pageName"] if job.pages else "",
                    "outputFile": tmp,
                    "filters": parsed_filters,
                })

            with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
                json.dump(node_config, f)
                cfg_path = f.name

            # Stream results: upload each PDF immediately as Node.js finishes it.
            # This way a crash on job N doesn't lose jobs 1..(N-1) that already succeeded.
            proc = subprocess.Popen(
                ["node", pdf_script, "--config", cfg_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
            )

            def _pipe_stderr():
                for line in proc.stderr:
                    print(f"{tag} {line.rstrip()}")
            threading.Thread(target=_pipe_stderr, daemon=True).start()

            sp = SharePointClient.from_config(sp_token, sharepoint_config)
            failed = []
            for line in proc.stdout:
                line = line.strip()
                if not line:
                    continue
                try:
                    r = json.loads(line)
                except json.JSONDecodeError:
                    continue
                job, tmp = tmp_files[r["exportKey"]]
                if r.get("success") and os.path.exists(tmp):
                    with open(tmp, "rb") as fh:
                        pdf_bytes = fh.read()
                    item = sp.upload_file(sp.output_folder, job.output_filename, pdf_bytes)
                    print(f"{tag} Uploaded {job.output_filename}: {item.get('webUrl', item.get('id'))}")
                    os.unlink(tmp)
                else:
                    print(f"{tag} FAILED {r['exportKey']}: {r.get('error', 'unknown error')}")
                    failed.append(r["exportKey"])

            proc.wait()
            if proc.returncode != 0:
                raise RuntimeError(f"export_report_pdf.js failed (worker {worker_id})")
            return failed

        if n == 1:
            all_failed.extend(run_chunk(chunks[0], 0))
        else:
            with ThreadPoolExecutor(max_workers=n) as executor:
                futures = {executor.submit(run_chunk, chunk, i): i for i, chunk in enumerate(chunks)}
                for future in as_completed(futures):
                    worker_id = futures[future]
                    try:
                        all_failed.extend(future.result())
                    except Exception as e:
                        print(f"[w{worker_id}] crashed: {e}")
                        all_failed.append(f"worker-{worker_id}-crash")

    return all_failed


def _inject_captured_states(jobs, node_script: str):
    reports_with_slicers = {}
    for job in jobs:
        if job.slicer_overrides and job.report.key not in reports_with_slicers:
            reports_with_slicers[job.report.key] = {
                "key": job.report.key,
                "workspaceId": job.report.workspace_id,
                "reportId": job.report.report_id,
                "slicers": [
                    {"table": s.table, "column": s.column, "value": s.value}
                    for s in job.slicer_overrides
                ],
            }

    if not reports_with_slicers:
        return jobs

    node_config = {"reports": list(reports_with_slicers.values())}
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
        json.dump(node_config, f)
        config_path = f.name

    print(f"Capturing slicer states via Node.js for {len(reports_with_slicers)} report(s)...")
    result = subprocess.run(
        ["node", node_script, "--config", config_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"capture_slicer_state.js failed:\n{result.stderr}")

    if result.stderr:
        print(result.stderr.strip())

    states = json.loads(result.stdout)
    return [
        replace(job, captured_bookmark_state=states.get(job.report.key))
        if job.slicer_overrides else job
        for job in jobs
    ]


def _print_bookmarks(config, powerbi: PowerBIClient) -> None:
    for report in config.powerbi_reports:
        bookmarks = powerbi.list_bookmarks(report)
        print(f"\n{report.key} ({report.report_id}):")
        if not bookmarks:
            print("  (no bookmarks found)")
        for bm in bookmarks:
            print(f"  name={bm.get('name')!r}  displayName={bm.get('displayName')!r}  state={bm.get('state', '(none)')[:80]!r}")


def _print_column_values(config, powerbi: PowerBIClient) -> None:
    for report in config.powerbi_reports:
        for group in report.export_groups:
            if not group.values_from:
                continue
            src = group.values_from
            print(f"\n{report.key}.{group.key}  ({src.table}/{src.column}):")
            vals = powerbi.get_distinct_values(report, group)
            for v in vals:
                print(f"  key={v.key!r}  value={v.value!r}")
            if not vals:
                print("  (no values returned)")


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
