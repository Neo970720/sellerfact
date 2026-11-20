#!/usr/bin/env python3
"""게시가 끝난 작업을 큐에 ack한다 — 고객 메일 발송은 ack 시점에 일어난다.

poll_queue.py가 남긴 jobs/_done.txt를 읽어 ack하고, 처리 후 파일을 비운다.
게시(커밋)가 실패하면 이 단계가 실행되지 않으므로, 작업은 큐에 남아
다음 주기에 다시 처리된다(리포트는 같은 경로에 덮어써진다).
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from poll_queue import call  # noqa: E402

BASE = os.environ.get("REPORT_BASE", "https://neo970720.github.io/sellerfact")


def main() -> int:
    endpoint = os.environ.get("QUEUE_ENDPOINT", "").strip()
    key = os.environ.get("QUEUE_KEY", "").strip()
    done_file = ROOT / "jobs" / "_done.txt"
    if not endpoint or not key or not done_file.exists():
        print("[ack] 처리할 작업 없음")
        return 0
    jobs = [j.strip() for j in done_file.read_text("utf-8").splitlines() if j.strip()]
    if not jobs:
        return 0
    for job in jobs:
        meta = ROOT / "docs" / "r" / job / "meta.json"
        ok = meta.exists() and '"done"' in meta.read_text("utf-8")
        res = call(endpoint, {"action": "ack", "key": key, "job": job,
                              "status": "done" if ok else "error",
                              "url": f"{BASE}/view.html?job={job}",
                              "message": "" if ok else "report missing after run"})
        print(f"[ack] {job} → {res.get('ok')}")
    done_file.write_text("", "utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
