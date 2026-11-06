# -*- coding: utf-8 -*-
"""Create and populate the public GitHub repository, then close the report link.

This performs steps 2-5 of the submission checklist in one command::

    <venv>/Scripts/python.exe src/publish_repo.py

Authentication
--------------
A GitHub OAuth token (from the device-code flow) or a classic PAT with
``repo`` scope, supplied by either

  * the ``AIDD_GH_TOKEN`` environment variable, or
  * ``%TEMP%/aidd_gh_token`` (written by the device-code helper).

The token is never written to the repository, to ``.git/config``, or to any
log line.  Delete the temp file after publishing (the script does it for you).

Why not ``git push``
--------------------
The campus transparent proxy intermittently answers
"CONNECT tunnel failed, response 502" for ``github.com`` git endpoints, while
``api.github.com`` is consistently reachable.  All writes here therefore use
the REST API (blob -> tree -> commit -> ref), which needs no git transport.
The local commit is still created so the working tree stays reproducible, and
the remote is recorded for normal use once the network behaves.

Sequence
--------
1. Resolve the authenticated user; create the empty Public repo if absent.
2. Run ``20_set_repo_url.py`` so the delivered report carries the live link.
3. Commit the resulting changes locally.
4. Upload every tracked file via the REST API as one root commit on ``main``.
5. Verify anonymously that the repo and its key paths are readable.
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPO = "aidd8122-fap-radioligand"
STEM = "AIDD8122_Final_Report_HuangYong_P2600222_v2"
API = "https://api.github.com"
TOKEN_CACHE = Path(tempfile.gettempdir()) / "aidd_gh_token"


def read_token(explicit: str | None = None) -> str:
    tok = explicit or os.environ.get("AIDD_GH_TOKEN", "").strip()
    if not tok and TOKEN_CACHE.exists():
        tok = TOKEN_CACHE.read_text(encoding="utf-8").strip()
    if not tok:
        raise SystemExit(
            "No GitHub token. Set AIDD_GH_TOKEN, or finish the device-code\n"
            "authorization so %TEMP%/aidd_gh_token exists.")
    return tok


def api(method: str, path: str, token: str | None = None,
        body: dict | None = None, retries: int = 4) -> tuple[int, dict | list | None]:
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "AIDD8122-submission-publisher",
    }
    if token:
        headers["Authorization"] = "Bearer " + token
    if data is not None:
        headers["Content-Type"] = "application/json"
    last = None
    for attempt in range(retries):
        req = urllib.request.Request(API + path, data=data, method=method,
                                     headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                raw = r.read()
                return r.status, (json.loads(raw) if raw else None)
        except urllib.error.HTTPError as e:
            raw = e.read()
            try:
                payload = json.loads(raw)
            except Exception:
                payload = {"message": raw[:200].decode("utf-8", "replace")}
            if e.code in (404, 409, 422):
                return e.code, payload
            last = f"HTTP {e.code}: {str(payload)[:180]}"
        except Exception as e:
            last = f"{type(e).__name__}: {e}"
        time.sleep(2 + 2 * attempt)
    raise RuntimeError(f"API {method} {path} failed: {last}")


def git(*args: str) -> str:
    p = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True)
    if p.returncode:
        raise RuntimeError(f"git {' '.join(args)} failed: {p.stderr.strip()[:300]}")
    return p.stdout


def resolve_owner(token: str, override: str | None) -> str:
    if override:
        return override
    status, me = api("GET", "/user", token)
    if status != 200 or not isinstance(me, dict):
        raise RuntimeError(f"Could not resolve the authenticated user: {status} {me}")
    print(f"  authenticated as: {me['login']}")
    return me["login"]


def ensure_repo(owner: str, token: str) -> str:
    url = f"https://github.com/{owner}/{REPO}"
    status, info = api("GET", f"/repos/{owner}/{REPO}", token)
    if status == 200:
        print(f"  repo exists      : {url} (private={info.get('private')})")
        if info.get("private"):
            api("PATCH", f"/repos/{owner}/{REPO}", token, {"private": False})
            print("  set visibility  : Public")
        return url
    status, created = api("POST", "/user/repos", token, {
        "name": REPO,
        "description": "AIDD8122 final project - Trust Before Throughput: "
                       "FAP-selective radioligand warheads (Huang Yong, P2600222)",
        "private": False,
        "has_issues": False,
        "has_projects": False,
        "has_wiki": False,
        "auto_init": False,
    })
    if status not in (201, 422):
        raise RuntimeError(f"Could not create repository: {status} {created}")
    print(f"  repo created     : {url} (Public, empty)")
    return url


def write_report_link(url: str) -> None:
    print("  writing link into report / docs / submission package ...")
    subprocess.run([sys.executable, str(ROOT / "src" / "20_set_repo_url.py"), url],
                   cwd=ROOT, check=True)


def stage_and_commit(url: str) -> None:
    status = git("status", "--porcelain")
    if status.strip():
        git("add", "-A")
        git("commit", "-m", f"Publish public code repository: {url}")
        print("  local commit     : created")
    else:
        print("  local commit     : already clean")


def tracked_files() -> list[Path]:
    out = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT,
                         capture_output=True, check=True).stdout
    names = [n.decode("utf-8", "surrogateescape") for n in out.split(b"\0") if n]
    files = [ROOT / n for n in names if (ROOT / n).is_file()]
    return files


def upload(owner: str, token: str) -> None:
    files = tracked_files()
    print(f"  uploading        : {len(files)} tracked files via REST API")
    entries = []
    for i, f in enumerate(files, 1):
        rel = f.relative_to(ROOT).as_posix()
        content = base64.b64encode(f.read_bytes()).decode("ascii")
        status, blob = api("POST", f"/repos/{owner}/{REPO}/git/blobs", token,
                           {"content": content, "encoding": "base64"})
        if status not in (201, 200) or not isinstance(blob, dict):
            raise RuntimeError(f"blob upload failed for {rel}: {status} {blob}")
        entries.append({"path": rel, "mode": "100644",
                        "type": "blob", "sha": blob["sha"]})
        if i % 25 == 0 or i == len(files):
            print(f"    {i}/{len(files)} blobs")
    status, tree = api("POST", f"/repos/{owner}/{REPO}/git/trees", token,
                       {"tree": entries})
    if status not in (201, 200):
        raise RuntimeError(f"tree creation failed: {status} {tree}")
    status, commit = api("POST", f"/repos/{owner}/{REPO}/git/commits", token, {
        "message": "AIDD8122 final project: reproducible code, data, report and slides",
        "tree": tree["sha"],
        "parents": [],
    })
    if status not in (201, 200):
        raise RuntimeError(f"commit creation failed: {status} {commit}")
    status, ref = api("POST", f"/repos/{owner}/{REPO}/git/refs", token,
                      {"ref": "refs/heads/main", "sha": commit["sha"]})
    if status not in (201, 200):
        raise RuntimeError(f"ref creation failed: {status} {ref}")
    api("PATCH", f"/repos/{owner}/{REPO}", token, {"default_branch": "main"})
    print(f"  root commit      : {commit['sha'][:10]} on main")


def configure_remote(url: str) -> None:
    current = subprocess.run(["git", "remote"], cwd=ROOT,
                             capture_output=True, text=True).stdout.split()
    if "origin" in current:
        git("remote", "set-url", "origin", url + ".git")
    else:
        git("remote", "add", "origin", url + ".git")
    print(f"  remote origin    : {url}.git")


def verify(url: str, owner: str) -> None:
    status, info = api("GET", f"/repos/{owner}/{REPO}")          # anonymous
    ok_repo = status == 200 and isinstance(info, dict) and not info.get("private")
    print(f"  anonymous API    : {'OK' if ok_repo else 'FAIL'} "
          f"(http={status}, private={info.get('private') if isinstance(info, dict) else '?'})")
    for path in ("README.md", "src", "results", "docs", "report"):
        st, payload = api("GET", f"/repos/{owner}/{REPO}/contents/{path}")
        n = len(payload) if isinstance(payload, list) else 1
        print(f"    /{path:<12} : {'OK' if st == 200 else 'FAIL'} ({n} entries)")
    pdf = ROOT / "report" / f"{STEM}.pdf"
    raw = pdf.read_bytes() if pdf.exists() else b""
    print(f"  report PDF /URI  : {'OK' if url.encode() in raw else 'CHECK'} "
          f"({'link found in PDF' if url.encode() in raw else 'URL bytes not found'})")
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(ROOT / "report" / f"{STEM}.docx")
    ext = [r.target_ref for r in doc.part.rels.values() if r.is_external]
    print(f"  docx external rel: {ext}")
    print(f"  docx hyperlinks  : "
          f"{len(doc.element.body.findall('.//' + qn('w:hyperlink')))}")
    if not ok_repo:
        raise RuntimeError("Repository is not anonymously readable - do not submit yet.")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--owner", help="GitHub username (default: authenticated user)")
    ap.add_argument("--token", help="OAuth/PAT token (prefer AIDD_GH_TOKEN or temp cache)")
    ap.add_argument("--skip-link", action="store_true",
                    help="do not rewrite the report link (repo upload only)")
    args = ap.parse_args()

    token = read_token(args.token)
    owner = resolve_owner(token, args.owner)
    url = ensure_repo(owner, token)
    if not args.skip_link:
        write_report_link(url)
    stage_and_commit(url)
    configure_remote(url)
    status, _ = api("GET", f"/repos/{owner}/{REPO}/git/ref/heads/main", token)
    if status == 200:
        print("  main already has commits - skipping upload (idempotent)")
    else:
        upload(owner, token)
    verify(url, owner)
    if TOKEN_CACHE.exists():
        TOKEN_CACHE.unlink()
        print("  token cache      : removed")
    print(f"\nDONE. Public repository: {url}")


if __name__ == "__main__":
    main()
