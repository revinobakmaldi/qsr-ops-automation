#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from report_ops_automation.auth import GRAPH_APP_SCOPE, get_app_token
from report_ops_automation.config import load_config
from report_ops_automation.env import load_dotenv, require_first_env
from report_ops_automation.sharepoint import SharePointClient


def main() -> None:
    parser = argparse.ArgumentParser(description="Upload a tiny PDF to the configured SharePoint folder.")
    parser.add_argument("--config", default="config/reports.yaml")
    parser.add_argument("--file-name", default="")
    args = parser.parse_args()

    load_dotenv()
    config = load_config(args.config)

    tenant_id = require_first_env("SHAREPOINT_TENANT_ID", "AZURE_TENANT_ID")
    client_id = require_first_env("SHAREPOINT_CLIENT_ID", "AZURE_CLIENT_ID")
    client_secret = require_first_env("SHAREPOINT_CLIENT_SECRET", "AZURE_CLIENT_SECRET")

    sharepoint = SharePointClient.from_config(
        get_app_token(tenant_id, client_id, client_secret, GRAPH_APP_SCOPE),
        config.sharepoint,
    )

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = args.file_name or f"sharepoint_upload_test_{stamp}.pdf"
    item = sharepoint.upload_file(sharepoint.output_folder, file_name, _minimal_pdf())
    print(f"Uploaded {file_name}")
    print(item.get("webUrl", item.get("id")))


def _minimal_pdf() -> bytes:
    return b"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] /Contents 4 0 R >>
endobj
4 0 obj
<< /Length 61 >>
stream
BT /F1 18 Tf 36 96 Td (SharePoint upload test) Tj ET
endstream
endobj
xref
0 5
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000204 00000 n
trailer
<< /Size 5 /Root 1 0 R >>
startxref
315
%%EOF
"""


if __name__ == "__main__":
    main()
