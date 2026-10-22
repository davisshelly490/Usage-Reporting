"""Verify authenticated access to the OpenAI Audit Logs API.

Run: python audit_logs_api.py

Reads OPENAI_ADMIN_KEY from the environment or the .env beside this script.
Requires an OpenAI organization Admin API key with audit-log read access.
OPENAI_WORKSPACE_ID is not used by this endpoint.
Requires Python 3.10+; no packages to install.

Makes one GET request for at most one event, validates the response, and stops.
Event contents and credentials are never printed or saved.

Reference:
https://developers.openai.com/api/reference/resources/admin/subresources/organization/subresources/audit_logs/methods/list
"""

import json
import os
from pathlib import Path
import sys
from urllib.error import HTTPError, URLError
from urllib.request import HTTPRedirectHandler, Request, build_opener


AUDIT_LOGS_URL = "https://api.openai.com/v1/organization/audit_logs?limit=1"


def load_env(path: Path) -> None:
    """Load the admin key without replacing a process environment variable."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].strip()
        name, separator, value = line.partition("=")
        if not separator or name.strip() != "OPENAI_ADMIN_KEY":
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] in "\"'" and value[-1] == value[0]:
            value = value[1:-1]
        else:
            value = value.split(" #", 1)[0].rstrip()
        os.environ.setdefault("OPENAI_ADMIN_KEY", value)


class NoRedirects(HTTPRedirectHandler):
    """Keep the admin credential on the intended endpoint."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def verify_connection(key: str) -> dict:
    """Make one authenticated read and return only connection metadata."""
    request = Request(
        AUDIT_LOGS_URL,
        headers={"Authorization": f"Bearer {key}", "Accept": "application/json"},
        method="GET",
    )
    try:
        with build_opener(NoRedirects()).open(request, timeout=30) as response:
            status = response.status
            payload = json.load(response)
    except HTTPError as exc:
        hints = {
            401: "Credential rejected. Check OPENAI_ADMIN_KEY and its expiry; this endpoint requires an organization Admin API key.",
            403: "Access denied. Check that the Admin key can read audit logs and that audit logging is available and enabled for the organization.",
            404: "Audit Logs API unavailable. Check organization access and API availability.",
            429: "Rate limited. Try again later.",
        }
        hint = hints.get(exc.code, "Request failed; no automatic retry was attempted.")
        exc.close()
        raise RuntimeError(f"HTTP {exc.code}: {hint}") from None
    except (URLError, TimeoutError, OSError):
        raise RuntimeError("Connection failed. Check network access to api.openai.com.") from None
    except (ValueError, UnicodeError):
        raise RuntimeError("The API returned an unreadable JSON response.") from None

    if status != 200:
        raise RuntimeError(f"Unexpected HTTP {status}; connection not verified.")
    if (
        not isinstance(payload, dict)
        or payload.get("object") != "list"
        or not isinstance(payload.get("data"), list)
        or not isinstance(payload.get("has_more"), bool)
    ):
        raise RuntimeError("Unexpected audit-log response shape; connection not verified.")
    for event in payload["data"]:
        if not isinstance(event, dict) or not isinstance(event.get("id"), str) or not isinstance(event.get("type"), str):
            raise RuntimeError("Unexpected audit-log event shape; connection not verified.")
    return {
        "http_status": status,
        "events_returned": len(payload["data"]),
        "has_more": payload["has_more"],
    }


def main() -> int:
    try:
        load_env(Path(__file__).resolve().with_name(".env"))
        key = os.environ.get("OPENAI_ADMIN_KEY", "").strip()
        if not key:
            raise RuntimeError("Set OPENAI_ADMIN_KEY in .env or your environment.")
        if any(character.isspace() for character in key):
            raise RuntimeError("OPENAI_ADMIN_KEY contains whitespace; check its formatting.")
        result = verify_connection(key)
        print(f"Connected to the OpenAI Audit Logs API (HTTP {result['http_status']}).")
        print(f"Events returned: {result['events_returned']}; more available: {result['has_more']}.")
        print("Authenticated read verified. Connection check complete.")
        return 0
    except (RuntimeError, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
