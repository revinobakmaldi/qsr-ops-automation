from __future__ import annotations

from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

from .config import SharePointConfig
from .http import ApiClient


class SharePointClient:
    def __init__(self, graph_token: str, drive_id: str):
        self.drive_id = drive_id
        self.api = ApiClient(graph_token, "https://graph.microsoft.com/v1.0")

    def upload_file(self, folder_path: str, file_name: str, content: bytes) -> dict:
        graph_path = _drive_path(folder_path, file_name)
        response = self.api.request(
            "PUT",
            f"/drives/{self.drive_id}/root:/{graph_path}:/content",
            data=content,
            headers={"Content-Type": "application/pdf"},
        )
        return response.json()

    def list_folder(self, folder_path: str) -> list[dict]:
        graph_path = _drive_path(folder_path)
        url = f"/drives/{self.drive_id}/root:/{graph_path}:/children"
        items: list[dict] = []
        while url:
            payload = self.api.get_json(url)
            items.extend(payload.get("value", []))
            url = payload.get("@odata.nextLink")
        return items

    @classmethod
    def from_config(cls, graph_token: str, config: SharePointConfig) -> "ResolvedSharePointClient":
        api = ApiClient(graph_token, "https://graph.microsoft.com/v1.0")
        if config.folder_url:
            drive_id, output_folder = _resolve_folder_url(api, config.folder_url)
            return ResolvedSharePointClient(graph_token, drive_id, output_folder)
        if not config.drive_id or not config.output_folder:
            raise ValueError("SharePoint config requires either folder_url or drive_id + output_folder.")
        return ResolvedSharePointClient(graph_token, config.drive_id, config.output_folder)


class ResolvedSharePointClient(SharePointClient):
    def __init__(self, graph_token: str, drive_id: str, output_folder: str):
        super().__init__(graph_token, drive_id)
        self.output_folder = output_folder.strip("/")


def _drive_path(*parts: str) -> str:
    cleaned = "/".join(str(part).strip("/") for part in parts if part)
    return "/".join(quote(segment) for segment in Path(cleaned).parts if segment not in {"/", ""})


def _resolve_folder_url(api: ApiClient, folder_url: str) -> tuple[str, str]:
    parsed = urlparse(folder_url)
    folder_id = parse_qs(parsed.query).get("id", [None])[0]
    if not folder_id:
        raise ValueError("SharePoint folder_url must include an id query parameter.")

    decoded_path = unquote(folder_id)
    path_parts = [part for part in decoded_path.split("/") if part]
    if len(path_parts) < 4 or path_parts[0] != "sites":
        raise ValueError(f"Unsupported SharePoint folder path: {decoded_path}")

    site_path = f"/sites/{path_parts[1]}"
    library_name = path_parts[2]
    output_folder = "/".join(path_parts[3:])

    site = api.get_json(f"/sites/{parsed.hostname}:{site_path}")
    drives = api.get_json(f"/sites/{site['id']}/drives").get("value", [])
    drive = next(
        (
            item
            for item in drives
            if item.get("name") == library_name
            or item.get("webUrl", "").rstrip("/").endswith(f"/{quote(library_name)}")
        ),
        None,
    )
    if not drive:
        available = ", ".join(item.get("name", "") for item in drives)
        raise ValueError(f"Could not find SharePoint document library '{library_name}'. Available: {available}")
    return drive["id"], output_folder
