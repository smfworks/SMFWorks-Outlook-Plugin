"""Outlook status API for Hermes Desktop.

Mounted at /api/plugins/hermes-outlook/.

The pane hosts Outlook on the web. This API only names the signed-in
account from Azure CLI's current Graph token. No passwords, no secrets
in git. Mail bodies are not read here.
"""
from __future__ import annotations

import json
import os
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

try:
    from fastapi import APIRouter
except Exception:  # pragma: no cover
    class APIRouter:  # type: ignore
        def get(self, *_a, **_k):
            return lambda fn: fn

GRAPH = "https://graph.microsoft.com/v1.0"
AZ_CANDIDATES = (
    os.environ.get("AZ_CMD") or "",
    r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    "az.cmd",
    "az",
)

router = APIRouter()


def _az_cmd() -> Optional[str]:
    for cand in AZ_CANDIDATES:
        if not cand:
            continue
        p = Path(cand)
        if p.is_file():
            return str(p)
        if cand in ("az", "az.cmd"):
            return cand
    return None


def graph_token() -> Tuple[Optional[str], Optional[str]]:
    """Return (access_token, error). Never log the token."""
    az = _az_cmd()
    if not az:
        return None, "azure_cli_missing"
    try:
        proc = subprocess.run(
            [az, "account", "get-access-token", "--resource", "https://graph.microsoft.com", "-o", "json"],
            capture_output=True,
            text=True,
            timeout=45,
            shell=False,
        )
    except FileNotFoundError:
        return None, "azure_cli_missing"
    except subprocess.TimeoutExpired:
        return None, "azure_cli_timeout"
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()[:240]
        return None, err or "azure_cli_failed"
    try:
        payload = json.loads(proc.stdout or "{}")
    except json.JSONDecodeError:
        return None, "azure_cli_bad_json"
    token = payload.get("accessToken")
    if not isinstance(token, str) or not token:
        return None, "azure_cli_no_token"
    return token, None


def graph_get(path: str, token: str, query: Optional[Dict[str, str]] = None) -> Tuple[int, Dict[str, Any]]:
    url = GRAPH + path
    if query:
        url += "?" + urllib.parse.urlencode(query)
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8") or "{}")
            return resp.status, body if isinstance(body, dict) else {"value": body}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            parsed = {"error": {"message": raw[:400]}}
        return e.code, parsed if isinstance(parsed, dict) else {"error": {"message": raw[:400]}}
    except Exception as e:  # noqa: BLE001
        return 599, {"error": {"message": str(e)}}


@router.get("/status")
def status() -> Dict[str, Any]:
    token, err = graph_token()
    if err or not token:
        return {
            "ok": False,
            "source": "azure-cli",
            "error": err or "no_token",
            "hint": "Outlook on the web still loads. Run `az login` to name the account.",
        }
    code, me = graph_get("/me", token, {"$select": "displayName,mail,userPrincipalName,id"})
    if code != 200:
        return {
            "ok": False,
            "source": "azure-cli",
            "error": ((me.get("error") or {}) if isinstance(me.get("error"), dict) else {}).get("message") or f"graph_{code}",
        }
    return {
        "ok": True,
        "source": "azure-cli",
        "me": {
            "id": me.get("id"),
            "displayName": me.get("displayName"),
            "mail": me.get("mail"),
            "userPrincipalName": me.get("userPrincipalName"),
        },
    }
