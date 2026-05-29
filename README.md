# Report Ops Automation

Two automations are included:

1. Export Power BI reports with configured filter state to PDF, rename them, and upload them to a SharePoint document library folder.
2. Find the generated PDFs in SharePoint and send each report link to the mapped Teams recipient.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
cp config/reports.example.yaml config/reports.yaml
```

Fill `.env` and `config/reports.yaml`.

For reports with multiple pages and value maps, use `export_groups` in `config/reports.yaml`.
Each generated PDF gets this key:

```text
{report_key}.{export_group_key}.{value_key}
```

Example: `sales_daily.regional.west`.

When an export group has `values_from`, the script queries the Power BI semantic model for distinct values from that table/column and generates one PDF per value.

## Azure permissions

The export script uses app credentials. Power BI and SharePoint can use different app registrations:

- Power BI REST API application permissions needed for report export.
- Microsoft Graph application permission such as `Sites.ReadWrite.All` for SharePoint upload.
- Power BI tenant settings must allow service principals to use Power BI APIs, and report export requires supported capacity/licensing.

Set these in `.env`:

- `POWERBI_TENANT_ID`, `POWERBI_CLIENT_ID`, `POWERBI_CLIENT_SECRET`
- `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`

The old shared names `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` are still accepted as fallbacks.

Email delivery uses Microsoft Graph application permissions:

- Add `Mail.Send` application permission to the email app registration and grant admin consent.
- Set `EMAIL_TENANT_ID`, `EMAIL_CLIENT_ID`, `EMAIL_CLIENT_SECRET`, and `EMAIL_SENDER_UPN`.
- If `EMAIL_*` credentials are omitted, the script falls back to `SHAREPOINT_*`, but that app must also have `Mail.Send`.

Teams delivery uses delegated device-code auth because normal Teams chat message sending is a user action:

- Microsoft Graph delegated permissions such as `Chat.ReadWrite`, `ChatMessage.Send`, `User.Read`, and SharePoint read permissions such as `Files.Read.All` or `Sites.Read.All`.
- The signed-in account must be the same user as `TEAMS_SENDER_UPN`.

## Run

```bash
python scripts/export_powerbi_to_sharepoint.py --config config/reports.yaml
python scripts/send_sharepoint_pdfs_to_teams.py --config config/reports.yaml
```

Use `--dry-run` on the delivery script to print planned deliveries without sending messages.

Delivery mapping supports `channel: email` or `channel: teams`. Email is the default.

```yaml
report_delivery:
  - export_key: "operation.regional.region_jakarta"
    channel: "email"
    recipient_upn: "person@company.com"
    subject: "{report_name} PDF is ready"
    message: "Hi, your {report_name} PDF is ready: <a href=\"{web_url}\">{file_name}</a>"
```

## Microsoft API references

- Power BI `ExportTo` starts an async export job and supports `reportLevelFilters`.
- Graph drive upload uses `PUT /drives/{drive-id}/root:/{path}:/content` for files up to 250 MB.
- Graph drive folder listing uses `GET /drives/{drive-id}/root:/{path}:/children`.
- Email delivery uses `POST /users/{sender}/sendMail`.
- Teams messages use `POST /chats/{chat-id}/messages`; one-on-one chats can be created with `POST /chats`.
