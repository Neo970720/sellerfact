#!/usr/bin/env python3
"""셀러팩트 큐 소비기 — 접수 큐(Apps Script)에서 대기 작업을 받아 점검을 돌린다.

GitHub Actions 스케줄이 이 스크립트를 호출한다. AI 토큰 0 · 외부 유료 API 0.
환경변수 (없으면 조용히 건너뛴다):
  QUEUE_ENDPOINT  접수 큐 엔드포인트 (예: https://script.google.com/macros/s/.../exec)
  QUEUE_KEY       큐 공유 키 (Actions 시크릿)
"""

from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
JOBS = ROOT / "jobs"
REPORTS = ROOT / "docs" / "r"
TIMEOUT = 45


def call(endpoint: str, params: dict) -> dict:
    url = endpoint + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "sellerfact-runner/1.0"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:  # noqa: S310 (엔드포인트는 운영자 설정값)
        body = resp.read().decode("utf-8", errors="replace")
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return {"ok": False, "error": f"non-json: {body[:200]}"}


def already_done(job: str) -> bool:
    meta = REPORTS / job / "meta.json"
    if not meta.exists():
        return False
    try:
        return json.loads(meta.read_text("utf-8")).get("status") == "done"
    except Exception:  # noqa: BLE001
        return False


def main() -> int:
    endpoint = os.environ.get("QUEUE_ENDPOINT", "").strip()
    key = os.environ.get("QUEUE_KEY", "").strip()
    if not endpoint or not key:
        print("[queue] QUEUE_ENDPOINT/QUEUE_KEY 미설정 — 이번 실행은 건너뜁니다.")
        return 0

    res = call(endpoint, {"action": "queue", "key": key})
    if not res.get("ok"):
        print(f"[queue] 큐 조회 실패: {res.get('error')}")
        return 0
    jobs = res.get("jobs", [])
    print(f"[queue] 대기 {len(jobs)}건")
    processed = 0
    for item in jobs:
        job = str(item.get("job", "")).strip()
        if not job or already_done(job):
            continue
        payload = call(endpoint, {"action": "payload", "key": key, "job": job})
        if not payload.get("ok"):
            print(f"[queue] {job} 본문 조회 실패: {payload.get('error')}")
            call(endpoint, {"action": "ack", "key": key, "job": job,
                            "status": "error", "message": "payload fetch failed"})
            continue
        csv_bytes = base64.b64decode(payload["csv_b64"])
        if len(csv_bytes) > 400_000:
            call(endpoint, {"action": "ack", "key": key, "job": job,
                            "status": "error", "message": "input too large"})
            continue
        JOBS.mkdir(exist_ok=True)
        input_path = JOBS / f"{job}.csv"
        input_path.write_bytes(csv_bytes)
        client = str(payload.get("client") or item.get("client") or "셀러")[:40]

        cmd = [sys.executable, str(ROOT / "scripts" / "run_job.py"), "--job", job,
               "--input", str(input_path), "--client", client]
        if os.environ.get("SELLERFACT_OFFLINE") == "1":
            cmd.append("--no-network")
        proc = subprocess.run(cmd, capture_output=True, text=True)  # noqa: S603
        print(proc.stdout.strip())
        if proc.returncode != 0:
            print(proc.stderr.strip()[-1500:])
            call(endpoint, {"action": "ack", "key": key, "job": job,
                            "status": "error", "message": proc.stderr.strip()[-300:]})
            continue
        status = "done"
        url = f"https://neo970720.github.io/sellerfact/view.html?job={job}"
        call(endpoint, {"action": "ack", "key": key, "job": job, "status": status, "url": url})
        processed += 1
    print(f"[queue] 처리 완료 {processed}건")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
