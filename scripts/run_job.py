#!/usr/bin/env python3
"""셀러팩트 무인 점검 실행기 — 한 건(job)의 입력 CSV를 점검해 공개 리포트 폴더로 낸다.

이 스크립트는 AI 없이 순수 파이썬으로 동작한다(토큰 0). GitHub Actions가 호출한다.

사용:
  python3 scripts/run_job.py --job demo-001 --input jobs/demo-001.csv --client "샘플셀러"
산출:
  docs/r/<job>/report.xlsx   6시트 엑셀 리포트
  docs/r/<job>/report.pdf    11쪽 PDF 리포트
  docs/r/<job>/data.json     원본 판정 데이터
  docs/r/<job>/meta.json     상태·등급 요약(고객 페이지가 폴링)
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import sys
import tempfile
import traceback
from datetime import datetime, timezone, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

MAX_SKUS = 300
KST = timezone(timedelta(hours=9))

REQUIRED_COLS = ["sku_id", "product_name", "category"]
KNOWN_COLS = ["sku_id", "product_name", "category", "model_name", "price", "detail_text",
              "kc_cert_no", "is_purchase_agency", "origin", "disclosure_fields", "seller_note"]


def normalize_csv(src: Path, dst: Path) -> int:
    """인코딩·헤더를 정규화하고 SKU 수를 센다. 알 수 없는 열은 버린다."""
    raw = src.read_bytes()
    text = None
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")

    rows = list(csv.reader(text.splitlines()))
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        raise SystemExit("빈 파일입니다.")
    header = [re.sub(r"\s+", "", h) for h in rows[0]]
    idx = {}
    for i, name in enumerate(header):
        if name in KNOWN_COLS:
            idx[name] = i
    missing = [c for c in REQUIRED_COLS if c not in idx]
    if missing:
        raise SystemExit(f"필수 열이 없습니다: {', '.join(missing)}")
    out_rows = []
    for r in rows[1:]:
        if len(r) <= idx["product_name"] or not r[idx["product_name"]].strip():
            continue
        rec = {c: (r[i].strip() if i < len(r) else "") for c, i in idx.items()}
        out_rows.append(rec)
    if not out_rows:
        raise SystemExit("데이터 행이 없습니다.")
    if len(out_rows) > MAX_SKUS:
        raise SystemExit(f"한 번에 최대 {MAX_SKUS}개 SKU까지 점검합니다(입력 {len(out_rows)}개).")

    dst.parent.mkdir(parents=True, exist_ok=True)
    with dst.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=[c for c in KNOWN_COLS if c in idx])
        w.writeheader()
        for rec in out_rows:
            w.writerow(rec)
    return len(out_rows)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="셀러팩트 무인 점검 실행기")
    ap.add_argument("--job", required=True, help="작업 ID(영문·숫자·하이픈)")
    ap.add_argument("--input", required=True, help="입력 CSV 경로")
    ap.add_argument("--client", default="셀러", help="리포트에 표시할 셀러명")
    ap.add_argument("--out-root", default=str(ROOT / "docs" / "r"))
    ap.add_argument("--no-network", action="store_true", help="인증 실시간 조회 없이 실행")
    args = ap.parse_args(argv)

    job = re.sub(r"[^A-Za-z0-9_-]", "", args.job)[:60]
    if not job:
        raise SystemExit("작업 ID가 올바르지 않습니다.")
    out_dir = Path(args.out_root) / job
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    started = datetime.now(KST)
    meta = {
        "job": job,
        "client": args.client,
        "status": "running",
        "started_at": started.isoformat(timespec="seconds"),
    }
    (out_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            clean_csv = tmp_dir / "input.csv"
            n = normalize_csv(Path(args.input), clean_csv)

            import audit_engine  # noqa: E402

            buf = str(tmp_dir / "out")
            argv_engine = ["--input", str(clean_csv), "--out", buf,
                           "--client", args.client, "--brand-name", "셀러팩트"]
            if args.no_network:
                argv_engine.append("--no-network")
            audit_engine.main(argv_engine)

            produced = sorted(Path(buf).glob("*"))
            by_suffix = {}
            for p in produced:
                by_suffix[p.suffix.lower()] = p
            if ".xlsx" in by_suffix:
                shutil.copy2(by_suffix[".xlsx"], out_dir / "report.xlsx")
            if ".pdf" in by_suffix:
                shutil.copy2(by_suffix[".pdf"], out_dir / "report.pdf")
            if ".json" in by_suffix:
                shutil.copy2(by_suffix[".json"], out_dir / "data.json")

            data = json.loads((out_dir / "data.json").read_text("utf-8"))
            counts = data["counts"]
            top = [r for r in data["results"] if r.get("grade") in ("D", "E")]
            meta.update({
                "status": "done",
                "finished_at": datetime.now(KST).isoformat(timespec="seconds"),
                "sku_count": n,
                "counts": counts,
                "ruleset": {k: data["ruleset"][k] for k in ("id", "version", "law_basis_date")},
                "offline": bool(args.no_network),
                "urgent": [{"sku_id": r.get("sku_id"), "product_name": r.get("product_name"),
                            "grade": r.get("grade"), "score": r.get("score")} for r in top][:5],
                "files": {"pdf": "report.pdf", "xlsx": "report.xlsx", "json": "data.json"},
            })
    except SystemExit as exc:
        meta.update({"status": "error", "message": str(exc)})
        (out_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
        print(f"[job:{job}] 실패: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        meta.update({"status": "error", "message": f"{type(exc).__name__}: {exc}"})
        (out_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
        traceback.print_exc()
        print(f"[job:{job}] 예외: {exc}", file=sys.stderr)
        return 1

    (out_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
    print(json.dumps({"job": job, "status": "done", "sku": meta["sku_count"],
                      "grades": meta["counts"]["grades"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
