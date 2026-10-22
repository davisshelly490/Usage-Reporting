"""Verify authenticated access to the OpenAI Audit Logs API.

Run: python audit_log_api.py

Reads OPENAI_ADMIN_KEY from the environment or the .env beside this script.
Uses the Haskell workspace ID defined directly in WORKSPACE_ID below.
Requires the ChatGPT workspace Admin key's audit/compliance log read access.
Requires Python 3.10+; no packages to install.

Makes one GET request for at most one AUDIT_LOG file's metadata from the
last 24 hours, validates the response, and stops. No files are downloaded.
Event contents and credentials are never printed or saved.

Reference:
https://developers.chatgpt.com/cookbook/examples/chatgpt/compliance_api/logs_platform
"""

import json
from datetime import datetime, timedelta, timezone
import os
from pathlib import Path
import re
import sys
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import HTTPRedirectHandler, Request, build_opener
from uuid import UUID


AUDIT_LOGS_BASE = "https://api.chatgpt.com/v1/compliance/workspaces"
WORKSPACE_ID = "2311451a-e495-405d-a90f-c9c27f7ceb3d"


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
        os.environ.setdefault(name.strip(), value)


class NoRedirects(HTTPRedirectHandler):
    """Keep the admin credential on the intended endpoint."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def verify_connection(key: str, workspace_id: str) -> dict:
    """Make one authenticated read and return only connection metadata."""
    workspace_id = str(UUID(workspace_id))
    after = (datetime.now(timezone.utc) - timedelta(days=1)).isoformat()
    query = urlencode({"limit": 1, "event_type": "AUDIT_LOG", "after": after})
    request = Request(
        f"{AUDIT_LOGS_BASE}/{workspace_id}/logs?{query}",
        headers={"Authorization": f"Bearer {key}", "Accept": "application/json"},
        method="GET",
    )
    try:
        with build_opener(NoRedirects()).open(request, timeout=30) as response:
            status = response.status
            payload = json.load(response)
    except HTTPError as exc:
        # Keep useful server diagnostics while removing credentials and tokens.
        diagnostics = []
        content_type = exc.headers.get("Content-Type", "") if exc.headers else ""
        if "html" in content_type.lower():
            diagnostics.append("response=HTML (may be an edge/proxy rejection)")
        if exc.headers and exc.headers.get("cf-mitigated"):
            diagnostics.append("edge_challenge=true")
        request_id = exc.headers.get("x-request-id", "") if exc.headers else ""
        if re.fullmatch(r"[A-Za-z0-9_-]{1,120}", request_id):
            diagnostics.append(f"request_id={request_id}")
        try:
            body = json.load(exc)
            error = body.get("error", body.get("detail", {})) if isinstance(body, dict) else {}
            if isinstance(error, dict):
                for field in ("type", "code"):
                    value = error.get(field)
                    if isinstance(value, str) and re.fullmatch(r"[a-z_]{1,80}", value):
                        diagnostics.append(f"{field}={value}")
                message = error.get("message")
            else:
                message = error if isinstance(error, str) else None
            if isinstance(message, str):
                message = message.replace(key, "[REDACTED]")
                message = re.sub(r"[A-Za-z0-9_-]{24,}", "[REDACTED]", message)
                message = " ".join(message.split())[:800]
                diagnostics.append(f"server_message={message}")
        except (ValueError, AttributeError, OSError):
            pass
        hints = {
            400: "The compliance endpoint rejected the request parameters.",
            401: "The ChatGPT compliance endpoint rejected the loaded credential; no automatic retry was attempted.",
            403: "The compliance endpoint denied this request. See server_message for the specific reason.",
            404: "Workspace compliance route unavailable. Check workspace access and API availability.",
            429: "Rate limited. Try again later.",
        }
        hint = hints.get(exc.code, "Request failed; no automatic retry was attempted.")
        exc.close()
        detail = f" ({', '.join(diagnostics)})" if diagnostics else ""
        raise RuntimeError(f"HTTP {exc.code}{detail}: {hint}") from None
    except (URLError, TimeoutError, OSError):
        raise RuntimeError("Connection failed. Check network access to api.chatgpt.com.") from None
    except (ValueError, UnicodeError):
        raise RuntimeError("The API returned an unreadable JSON response.") from None

    if status != 200:
        raise RuntimeError(f"Unexpected HTTP {status}; connection not verified.")
    if (
        not isinstance(payload, dict)
        or not isinstance(payload.get("data"), list)
        or not isinstance(payload.get("has_more"), bool)
    ):
        raise RuntimeError("Unexpected audit-log response shape; connection not verified.")
    for event in payload["data"]:
        if not isinstance(event, dict) or not isinstance(event.get("id"), str):
            raise RuntimeError("Unexpected audit-log file metadata; connection not verified.")
    return {
        "http_status": status,
        "files_returned": len(payload["data"]),
        "has_more": payload["has_more"],
        "after_utc": after,
    }


def main() -> int:
    try:
        load_env(Path(__file__).resolve().with_name(".env"))
        key = os.environ.get("OPENAI_ADMIN_KEY", "").strip()
        workspace_id = WORKSPACE_ID
        if not key:
            raise RuntimeError("Set OPENAI_ADMIN_KEY in .env or your environment.")
        if any(character.isspace() for character in key):
            raise RuntimeError("OPENAI_ADMIN_KEY contains whitespace; check its formatting.")
        result = verify_connection(key, workspace_id)
        print(f"Connected to ChatGPT workspace audit logs (HTTP {result['http_status']}).")
        print(f"Event type: AUDIT_LOG; after: {result['after_utc']} (UTC).")
        print(f"Log-file metadata rows returned: {result['files_returned']}; more available: {result['has_more']}.")
        print("This checks one page of file metadata. No event contents were downloaded.")
        print("Authenticated read verified. Connection check complete.")
        return 0
    except (RuntimeError, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
