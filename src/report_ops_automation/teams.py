from __future__ import annotations

from html import escape

from .http import ApiClient


class TeamsClient:
    def __init__(self, graph_token: str):
        self.api = ApiClient(graph_token, "https://graph.microsoft.com/v1.0")

    def create_one_on_one_chat(self, sender_upn: str, recipient_upn: str) -> str:
        payload = {
            "chatType": "oneOnOne",
            "members": [
                _chat_member(sender_upn),
                _chat_member(recipient_upn),
            ],
        }
        response = self.api.post_json("/chats", payload)
        chat_id = response.get("id")
        if not chat_id:
            raise RuntimeError(f"Graph did not return a chat id: {response}")
        return chat_id

    def send_message(self, chat_id: str, html_message: str) -> dict:
        return self.api.post_json(
            f"/chats/{chat_id}/messages",
            {"body": {"contentType": "html", "content": html_message}},
        )


def format_delivery_message(template: str, report_name: str, file_name: str, web_url: str) -> str:
    return template.format(
        report_name=escape(report_name),
        file_name=escape(file_name),
        web_url=escape(web_url, quote=True),
    )


def _chat_member(upn: str) -> dict:
    return {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["owner"],
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{upn}')",
    }
