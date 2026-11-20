#!/usr/bin/env python3
"""GitHub 이슈로 들어온 점검 요청을 처리한다.

이슈 제목 예:  [AUDIT] 내 셀러 10상품 점검
이슈 본문: ```csv 로 감싼 CSV 블록 + (선택) "셀러명: 홍길동"

처리 후 리포트 링크를 이슈에 댓글로 남긴다.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def extract_block(body: str) -> str:
    m = re.search(r"```(?:csv)?\s*\n(.*?)```", body, re.S)
    if m:
        return m.group(1).strip()
    # 코드펜스가 없으면 첫 줄부터 CSV 헤더가 있는 구간을 쓴다
    lines = body.splitlines()
    for i, line in enumerate(lines):
        if line.lower().startswith("sku_id"):
            return "\n".join(lines[i:]).strip()
    return ""


def main() -> int:
    number = os.environ.get("ISSUE_NUMBER", "").strip()
    body = os.environ.get("ISSUE_BODY", "") or ""
    title = os.environ.get("ISSUE_TITLE", "") or ""
    if not number:
        print("이슈 번호 없음")
        return 1
    job = f"issue-{number}"
    client = "셀러"
    m = re.search(r"셀러명\s*[:：]\s*(.+)", body)
    if m:
        client = m.group(1).strip().splitlines()[0][:40]

    csv_text = extract_block(body)
    if not csv_text:
        subprocess.run(["gh", "issue", "comment", number, "--body",
                        "CSV 블록을 찾지 못했습니다. ```csv 로 감싼 표를 붙여 주세요."],
                       check=False)  # noqa: S603,S607
        return 1
    path = ROOT / "jobs" / f"{job}.csv"
    path.parent.mkdir(exist_ok=True)
    path.write_text(csv_text, "utf-8")

    proc = subprocess.run([sys.executable, str(ROOT / "scripts" / "run_job.py"),
                           "--job", job, "--input", str(path), "--client", client],
                          capture_output=True, text=True)  # noqa: S603
    print(proc.stdout.strip())
    print(proc.stderr.strip()[-800:])
    meta_path = ROOT / "docs" / "r" / job / "meta.json"
    meta = json.loads(meta_path.read_text("utf-8")) if meta_path.exists() else {}
    url = f"https://neo970720.github.io/sellerfact/view.html?job={job}"
    if meta.get("status") == "done":
        grades = meta.get("counts", {}).get("grades", {})
        comment = (f"점검이 끝났습니다. **{meta.get('sku_count')}개 SKU**\n\n"
                   f"- 등급 분포: {grades}\n\n"
                   f"- 리포트 보기: {url}\n"
                   f"- PDF: {url.replace('view.html?job=' + job, 'r/' + job + '/report.pdf')}\n"
                   f"- 엑셀: {url.replace('view.html?job=' + job, 'r/' + job + '/report.xlsx')}")
    else:
        comment = f"점검을 끝내지 못했습니다: {meta.get('message', '알 수 없는 오류')}\n입력 형식을 확인해 주세요."
    subprocess.run(["gh", "issue", "comment", number, "--body", comment], check=False)  # noqa: S603,S607
    return 0 if meta.get("status") == "done" else 1


if __name__ == "__main__":
    raise SystemExit(main())
