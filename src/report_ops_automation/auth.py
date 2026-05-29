from __future__ import annotations

from pathlib import Path

import msal


POWERBI_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
GRAPH_APP_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_DELEGATED_SCOPES = [
    "User.Read",
    "Files.Read.All",
    "Sites.Read.All",
    "Chat.ReadWrite",
    "ChatMessage.Send",
]


def get_app_token(tenant_id: str, client_id: str, client_secret: str, scopes: list[str]) -> str:
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    result = app.acquire_token_for_client(scopes=scopes)
    if "access_token" not in result:
        raise RuntimeError(f"Could not acquire app token: {result.get('error_description', result)}")
    return result["access_token"]


def get_delegated_graph_token(
    tenant_id: str,
    client_id: str,
    scopes: list[str] | None = None,
    cache_path: str | Path = ".token_cache.bin",
) -> str:
    cache_file = Path(cache_path)
    cache = msal.SerializableTokenCache()
    if cache_file.exists():
        cache.deserialize(cache_file.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )
    requested_scopes = scopes or GRAPH_DELEGATED_SCOPES
    accounts = app.get_accounts()
    result = app.acquire_token_silent(requested_scopes, account=accounts[0] if accounts else None)
    if not result:
        flow = app.initiate_device_flow(scopes=requested_scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not start device-code flow: {flow}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        cache_file.write_text(cache.serialize(), encoding="utf-8")

    if "access_token" not in result:
        raise RuntimeError(f"Could not acquire delegated token: {result.get('error_description', result)}")
    return result["access_token"]
