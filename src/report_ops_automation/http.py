from __future__ import annotations

import time
from typing import Any

import requests


class ApiClient:
    def __init__(self, token: str, base_url: str):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {token}"})

    def request(self, method: str, path: str, **kwargs: Any) -> requests.Response:
        url = path if path.startswith("https://") else f"{self.base_url}/{path.lstrip('/')}"
        last_exc: Exception | None = None
        for attempt in range(6):
            try:
                response = self.session.request(method, url, timeout=120, **kwargs)
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as exc:
                last_exc = exc
                delay = min(2 ** attempt, 30)
                print(f"Network error (attempt {attempt + 1}/6), retrying in {delay}s: {exc}")
                time.sleep(delay)
                continue

            if response.status_code not in {429, 500, 502, 503, 504}:
                self._raise_for_error(response)
                return response

            retry_after = response.headers.get("Retry-After")
            delay = int(retry_after) if retry_after and retry_after.isdigit() else min(2**attempt, 30)
            time.sleep(delay)

        if last_exc:
            raise last_exc
        self._raise_for_error(response)
        return response

    def get_json(self, path: str) -> dict[str, Any]:
        return self.request("GET", path).json()

    def post_json(self, path: str, payload: dict[str, Any]) -> dict[str, Any]:
        response = self.request("POST", path, json=payload)
        return response.json() if response.content else {}

    @staticmethod
    def _raise_for_error(response: requests.Response) -> None:
        if response.ok:
            return
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise RuntimeError(f"{response.request.method} {response.url} failed: {response.status_code} {detail}")
