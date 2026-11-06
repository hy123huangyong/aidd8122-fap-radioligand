# -*- coding: utf-8 -*-
"""GitHub device-code authorization helper for src/publish_repo.py.

    <venv>/Scripts/python.exe src/authorize_device.py issue
    <venv>/Scripts/python.exe src/authorize_device.py wait

``issue`` asks GitHub for a fresh 15-minute device code and prints the
verification URL plus the user code.  Open the URL in a browser, sign in with
the GitHub account, enter the code and click Authorize.

``wait`` then exchanges the approved device code for an OAuth token and caches
it in %TEMP%/aidd_gh_token (mode 0600).  The token value is never printed.
Run ``src/publish_repo.py`` afterwards; it deletes the cache when done.

The campus proxy intermittently returns 502 on github.com, so every request
is retried.  Uses the same public OAuth client id as the official GitHub CLI.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
from pathlib import Path

CLIENT_ID = "178c6fc778ccc68e1d6a"       # official GitHub CLI OAuth app
SCOPE = "repo"
ROOT = Path(__file__).resolve().parents[1]
DEVICE_CACHE = Path(tempfile.gettempdir()) / "aidd_gh_device.json"
TOKEN_CACHE = Path(tempfile.gettempdir()) / "aidd_gh_token"
CODE_HINT = ROOT / "GITHUB_AUTH_CODE.txt"   # always holds the currently valid code


def post_form(url: str, params: dict, tries: int = 10) -> dict:
    data = urllib.parse.urlencode(params)
    last = ""
    for i in range(tries):
        p = subprocess.run(
            ["curl", "-sS", "--max-time", "20", "-X", "POST",
             "-H", "Accept: application/json",
             "-H", "Content-Type: application/x-www-form-urlencoded",
             "--data", data, url],
            capture_output=True, text=True)
        raw = (p.stdout or "").strip()
        if raw.startswith("{"):
            return json.loads(raw)
        if raw.startswith("access_token="):        # form-encoded fallback
            return dict(urllib.parse.parse_qsl(raw))
        last = (raw or p.stderr or "empty response").strip()[:120]
        time.sleep(3)
    raise RuntimeError(f"POST {url} failed after {tries} attempts: {last}")


def write_hint(code: str, minutes: int) -> None:
    CODE_HINT.write_text(
        "GitHub device-code authorization (regenerated automatically)\n"
        "===========================================================\n"
        f"Page : https://github.com/login/device\n"
        f"Code : {code}     (valid ~{minutes} min from generation)\n"
        "\nSign in with your GitHub account, enter the code, click Authorize.\n"
        "This file is local only and is excluded from the repository.\n"
        f"Updated: {time.strftime('%Y-%m-%d %H:%M:%S')}\n",
        encoding="utf-8")
    os.chmod(CODE_HINT, 0o600)


def fresh_code() -> dict:
    d = post_form("https://github.com/login/device/code",
                  {"client_id": CLIENT_ID, "scope": SCOPE})
    DEVICE_CACHE.write_text(json.dumps(d), encoding="utf-8")
    os.chmod(DEVICE_CACHE, 0o600)
    write_hint(d["user_code"], int(d["expires_in"]) // 60)
    print(f"  NEW CODE: {d['user_code']} (written to {CODE_HINT.name})", flush=True)
    return d


def issue() -> None:
    d = fresh_code()
    print("\nOpen this page and sign in with your GitHub account:")
    print(f"    {d['verification_uri']}")
    print(f"Enter this code: {d['user_code']}")
    print(f"Then click Authorize.  Valid for {int(d['expires_in']) // 60} minutes.")
    print("\nThe authorization page will name the official GitHub CLI OAuth app.")
    print("After authorizing, tell the agent to continue (or run the 'wait' command).")


def wait(timeout: int = 7200) -> None:
    if not DEVICE_CACHE.exists():
        raise SystemExit("No device code cached - run 'issue' first.")
    d = json.loads(DEVICE_CACHE.read_text(encoding="utf-8"))
    write_hint(d["user_code"], 15)
    print(f"  current code: {d['user_code']} (also in {CODE_HINT.name})", flush=True)
    deadline = time.time() + timeout
    interval = int(d.get("interval", 5))
    while time.time() < deadline:
        try:
            r = post_form("https://github.com/login/oauth/access_token", {
                "client_id": CLIENT_ID,
                "device_code": d["device_code"],
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            }, tries=4)
        except RuntimeError as e:
            print("  network retry:", str(e)[:90], flush=True)
            time.sleep(5)
            continue
        if r.get("access_token"):
            TOKEN_CACHE.write_text(r["access_token"], encoding="utf-8")
            os.chmod(TOKEN_CACHE, 0o600)
            DEVICE_CACHE.unlink(missing_ok=True)
            CODE_HINT.unlink(missing_ok=True)
            print("AUTHORIZED - token cached (value not shown).", flush=True)
            return
        err = r.get("error")
        if err in ("expired_token", "incorrect_device_code"):
            print("  device code expired - issuing a fresh one", flush=True)
            try:
                d = fresh_code()
            except RuntimeError as e:
                print("  re-issue failed, retrying:", str(e)[:80], flush=True)
                time.sleep(5)
                continue
            interval = int(d.get("interval", 5))
            continue
        if err == "access_denied":
            raise SystemExit("Authorization was denied in the browser.")
        if err == "slow_down":
            interval = min(interval + 5, 15)
        print(f"  waiting for browser authorization ... {int(deadline - time.time())}s left",
              flush=True)
        time.sleep(interval)
    raise SystemExit("Timed out waiting for authorization. Run 'issue' again.")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "issue"
    if mode == "issue":
        issue()
    elif mode == "wait":
        wait()
    else:
        raise SystemExit(__doc__)
