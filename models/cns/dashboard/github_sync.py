"""GitHub sync for committed actuals.

Streamlit Cloud filesystems are ephemeral — local writes to
``committed_actuals.json`` are wiped on container restart. To make committed
P&L and balance-sheet uploads durable, this module pushes the file back to the
repo via the GitHub Contents API. The next deploy (or container restart) then
reads the up-to-date file from the repo.

Configuration (st.secrets or environment):
    github_token      — fine-grained PAT with Contents:write on the target repo
    github_repo       — "<owner>/<name>", default "c-clemons/financial-modeling"
    github_branch     — default "main"

If no token is available (typical for local development), sync is silently
skipped and a callable hook can surface the state in the UI.
"""

from __future__ import annotations

import base64
import json
import os
import urllib.error
import urllib.request
from pathlib import Path
from typing import Optional

DEFAULT_REPO = "c-clemons/financial-modeling"
DEFAULT_BRANCH = "main"
REPO_FILE_PATH = "models/cns/dashboard/data/committed_actuals.json"


def _read_secret(key: str) -> Optional[str]:
    """Read from st.secrets if available, else from environment."""
    try:
        import streamlit as st  # noqa: WPS433  — optional dependency
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key.upper())


def sync_enabled() -> bool:
    return bool(_read_secret("github_token"))


def push_committed_file(local_path: Path, commit_message: str) -> dict:
    """PUT ``local_path`` to the configured GitHub repo path.

    Returns a status dict: {ok: bool, message: str, sha: str|None, url: str|None}.

    Does not raise — on failure, returns ok=False with a description, so the
    Streamlit page can surface the error without crashing the upload flow.
    """
    token = _read_secret("github_token")
    if not token:
        return {"ok": False, "message": "no token configured", "sha": None, "url": None}

    repo = _read_secret("github_repo") or DEFAULT_REPO
    branch = _read_secret("github_branch") or DEFAULT_BRANCH
    api = f"https://api.github.com/repos/{repo}/contents/{REPO_FILE_PATH}"

    try:
        content_bytes = local_path.read_bytes()
    except FileNotFoundError:
        return {"ok": False, "message": f"local file not found: {local_path}",
                "sha": None, "url": None}

    # Look up the current SHA so the PUT is treated as an update, not a create
    existing_sha = _get_existing_sha(api, branch, token)

    payload = {
        "message": commit_message,
        "content": base64.b64encode(content_bytes).decode("ascii"),
        "branch": branch,
    }
    if existing_sha:
        payload["sha"] = existing_sha

    req = urllib.request.Request(
        api,
        data=json.dumps(payload).encode("utf-8"),
        method="PUT",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        return {"ok": False, "message": f"HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:200]}",
                "sha": None, "url": None}
    except urllib.error.URLError as e:
        return {"ok": False, "message": f"network error: {e.reason}",
                "sha": None, "url": None}

    new_sha = body.get("content", {}).get("sha")
    commit_url = body.get("commit", {}).get("html_url")
    return {"ok": True, "message": "pushed to GitHub", "sha": new_sha, "url": commit_url}


def _get_existing_sha(api_url: str, branch: str, token: str) -> Optional[str]:
    req = urllib.request.Request(
        f"{api_url}?ref={branch}",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read()).get("sha")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None  # first write
        raise
