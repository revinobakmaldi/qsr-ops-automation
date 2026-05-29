from __future__ import annotations

from html import escape

from .http import ApiClient


class EmailClient:
    def __init__(self, graph_token: str, sender_upn: str):
        self.api = ApiClient(graph_token, "https://graph.microsoft.com/v1.0")
        self.sender_upn = sender_upn

    def send_message(self, recipient_email: str, subject: str, html_message: str) -> None:
        self.api.post_json(
            f"/users/{self.sender_upn}/sendMail",
            {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML",
                        "content": html_message,
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": recipient_email,
                            }
                        }
                    ],
                },
                "saveToSentItems": True,
            },
        )


def format_email_subject(template: str, report_name: str, file_name: str) -> str:
    return template.format(
        report_name=report_name,
        file_name=file_name,
    )


def format_email_message(template: str, report_name: str, file_name: str, web_url: str) -> str:
    return template.format(
        report_name=escape(report_name),
        file_name=escape(file_name),
        web_url=escape(web_url, quote=True),
    )
