"""셀러팩트 감사 리포트 — 엑셀(.xlsx) 생성기.

수식은 xlsxwriter로 쓰면서 파이썬에서 계산한 값을 캐시값으로 함께 저장한다.
(엑셀 미설치 환경에서도 값이 보이고, tests/validate_outputs.py가 캐시값과
 독립 계산값을 대조해 수식 오류를 검증한다.)
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

import xlsxwriter

NAVY = "#0F1B2D"
ACCENT = "#F2A93B"
INK = "#1C2434"
GREY = "#5A6474"
LIGHT = "#F4F6FA"
BORDER = "#D8DEE9"
GRADE_FILL = {"A": "#D8F0E2", "B": "#E6F2D6", "C": "#FBEBCB", "D": "#FADFCB", "E": "#F7D2D6"}
GRADE_INK = {"A": "#1B7F5A", "B": "#4C8C3F", "C": "#8A5D00", "D": "#9C421A", "E": "#A61B29"}
SEV_FILL = {"치명": "#A61B29", "중대": "#C4571F", "보통": "#C98A13", "경미": "#5A6474", "정보": "#8A93A3"}

COLS_SKU = [
    ("SKU", 10), ("상품명", 34), ("카테고리", 18), ("상품군", 11), ("등급", 6), ("점수", 6),
    ("핵심 위험", 40), ("KC 대상 트랙", 20), ("인증번호", 18), ("인증 조회", 12),
    ("인증 상태", 10), ("리콜", 9), ("표시광고", 8), ("고시 누락", 8), ("구매대행", 9),
    ("다음 조치", 38),
]

STATUS_KO = {
    "verified": "확인", "not_found": "미확인", "unavailable": "조회불가",
    "format_error": "형식오류", "skipped": "미입력", "not_applicable": "해당없음",
}


def _est_height(values, widths, base=11.6, min_h=30, max_h=150):
    """열 폭(문자 수)과 텍스트 길이로 행 높이를 추정한다(한글은 1.7자 폭)."""
    lines = 1
    for v, w in zip(values, widths):
        if v is None:
            continue
        t = str(v)
        chars_per_line = max(6, int(w * 1.05))
        est = 0
        for part in t.split("\n"):
            korean = sum(1 for ch in part if ord(ch) > 0x3000)
            weight = len(part) + korean * 0.7
            est += max(1, int(weight / chars_per_line) + (1 if weight % chars_per_line else 0))
        lines = max(lines, est)
    return max(min_h, min(max_h, lines * base + 8))


def _fmt(wb, **kw):
    base = {"font_name": "맑은 고딕", "font_size": 10, "valign": "vcenter"}
    base.update(kw)
    return wb.add_format(base)


def build(data: dict, out_path: Path) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(out_path), {"nan_inf_to_errors": True})
    wb.set_properties({
        "title": "셀러팩트 감사 리포트",
        "subject": "KC·표시광고·고시정보 사전 점검",
        "author": "셀러팩트(SellerFact)",
        "company": "SellerFact",
        "comments": f"룰셋 {data['ruleset']['version']}",
    })

    f_title = _fmt(wb, bold=True, font_size=22, font_color=NAVY)
    f_sub = _fmt(wb, font_size=11, font_color=GREY)
    f_small = _fmt(wb, font_size=9, font_color=GREY, text_wrap=True, valign="top")
    f_section = _fmt(wb, bold=True, font_size=11, font_color="#FFFFFF", bg_color=NAVY, indent=1)
    f_label = _fmt(wb, bold=True, font_color=INK, bg_color=LIGHT, border=1, border_color=BORDER, indent=1)
    f_value = _fmt(wb, font_color=INK, border=1, border_color=BORDER, indent=1)
    f_head = _fmt(wb, bold=True, font_color="#FFFFFF", bg_color=NAVY, border=1, border_color=BORDER,
                  text_wrap=True, align="center")
    f_cell = _fmt(wb, border=1, border_color=BORDER, text_wrap=True, valign="top")
    f_cell_c = _fmt(wb, border=1, border_color=BORDER, text_wrap=True, align="center", valign="top")
    f_score = _fmt(wb, border=1, border_color=BORDER, align="center", bold=True)
    f_note = _fmt(wb, font_size=9, text_wrap=True, valign="top", font_color=GREY)
    f_kpi = _fmt(wb, font_size=26, bold=True, font_color=NAVY, align="center")
    f_kpi_lbl = _fmt(wb, font_size=9, font_color=GREY, align="center")

    results = data["results"]
    counts = {g["grade"]: 0 for g in [
        {"grade": "A"}, {"grade": "B"}, {"grade": "C"}, {"grade": "D"}, {"grade": "E"}]}
    for r in results:
        counts[r["grade"]] = counts.get(r["grade"], 0) + 1
    sev_counts: dict = {}
    for r in results:
        for f in r["findings"]:
            sev_counts[f["severity"]] = sev_counts.get(f["severity"], 0) + 1
    total_findings = sum(sev_counts.values())
    action_needed = sum(1 for r in results if r["grade"] in ("C", "D", "E"))

    # ---------------------------------------------------------------- 요약
    ws = wb.add_worksheet("요약")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 24)
    ws.set_column("C:C", 26)
    ws.set_column("D:D", 26)
    ws.set_column("E:E", 26)
    ws.set_column("F:F", 16)
    ws.set_column("G:G", 2)

    ws.merge_range("B2:E2", "셀러팩트", f_title)
    ws.set_row(1, 34)
    ws.merge_range("B3:E3", "SELLERFACT · 상품 규제 사전 점검 리포트", f_sub)
    ws.merge_range("B4:E4", f"점검 대상  {data.get('client','-')}", f_sub)

    ws.merge_range("B6:F6", "점검 개요", f_section)
    rows = [
        ("점검일시", data["generated_at"]),
        ("점검 SKU", f"{len(results)}개"),
        ("룰셋 버전", f"{data['ruleset']['id']} · {data['ruleset']['version']}"),
        ("법령 기준일", data["ruleset"]["law_basis_date"]),
        ("인증 조회", "제품안전정보센터(safetykorea.kr) 실시간 조회"),
    ]
    for i, (k, v) in enumerate(rows):
        ws.write(6 + i, 1, k, f_label)
        ws.merge_range(6 + i, 2, 6 + i, 4, v, f_value)
        ws.set_row(6 + i, 18)

    # KPI
    ws.merge_range("B13:F13", "핵심 지표", f_section)
    kpis = [("B14", len(results), "B15", "점검 SKU"),
            ("C14", action_needed, "C15", "조치 필요(C~E)"),
            ("D14", sev_counts.get("치명", 0), "D15", "치명 위험"),
            ("E14", total_findings, "E15", "총 발견 항목")]
    ws.set_row(13, 32)
    ws.set_row(14, 15)
    for cell, val, lbl_cell, label in kpis:
        ws.write(cell, val, f_kpi)
        ws.write(lbl_cell, label, f_kpi_lbl)

    # 등급 분포 (수식 + 캐시값)
    ws.merge_range("B17:F17", "위험 등급 분포", f_section)
    grade_labels = {"A": "A 양호", "B": "B 보완 권장", "C": "C 보완 필요", "D": "D 위험", "E": "E 긴급"}
    ws.set_row(17, 26)
    ws.set_row(18, 15)
    for i, g in enumerate(["A", "B", "C", "D", "E"]):
        col = chr(ord("B") + i)
        f_g = _fmt(wb, font_size=18, bold=True, align="center", font_color=GRADE_INK[g],
                   bg_color=GRADE_FILL[g], border=1, border_color=BORDER)
        ws.write_formula(f"{col}18", f'=COUNTIF(\'SKU 점검표\'!$E$5:$E${4 + len(results)},"{g}")',
                         f_g, counts.get(g, 0))
        ws.write(f"{col}19", grade_labels[g], _fmt(wb, font_size=9, align="center", font_color=GREY))

    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name": "SKU 수",
        "categories": ["요약", 18, 1, 18, 5],
        "values": ["요약", 17, 1, 17, 5],
        "fill": {"color": NAVY},
        "data_labels": {"value": True},
    })
    chart.set_title({"name": "등급 분포"})
    chart.set_legend({"none": True})
    chart.set_size({"width": 470, "height": 210})
    chart.set_chartarea({"border": {"none": True}})
    ws.insert_chart("B21", chart)

    # 우선 조치 Top
    top = []
    for r in results:
        for f in r["findings"]:
            if f["severity"] in ("치명", "중대"):
                top.append((r["sku_id"], r["product_name"], f))
    top.sort(key=lambda x: 0 if x[2]["severity"] == "치명" else 1)
    ws.merge_range("B34:F34", "우선 조치 항목 (치명·중대)", _fmt(wb, bold=True, font_color="#FFFFFF",
                                                       bg_color=NAVY, font_size=11, indent=1))
    tr = 34
    for sku, name, f in top[:10]:
        ws.merge_range(tr, 1, tr, 5, f"[{sku}] {f['title']}", _fmt(
            wb, font_size=9, text_wrap=True, valign="vcenter",
            font_color=SEV_FILL.get(f["severity"], INK)))
        ws.set_row(tr, 18)
        tr += 1

    # 면책
    ws.merge_range(tr + 1, 1, tr + 3, 5, "", f_note)
    ws.write(tr + 1, 1, "※ " + data["ruleset"]["disclaimer"], f_note)
    ws.write(tr + 5, 1, "셀러팩트 · SellerFact  |  본 리포트는 구매한 SKU 수량 범위 내에서만 유효합니다.", f_small)

    # ---------------------------------------------------------------- SKU 점검표
    ws2 = wb.add_worksheet("SKU 점검표")
    ws2.hide_gridlines(2)
    ws2.merge_range(0, 0, 0, len(COLS_SKU) - 1, f"SKU별 점검 결과 — {data.get('client','')}", f_title)
    ws2.set_row(0, 34)
    ws2.merge_range(1, 0, 1, len(COLS_SKU) - 1,
                    f"점검일 {data['generated_at']} · 룰셋 {data['ruleset']['version']} · 인증 조회: 제품안전정보센터 실시간",
                    f_sub)
    for c, (name, width) in enumerate(COLS_SKU):
        ws2.set_column(c, c, width)
        ws2.write(3, c, name, f_head)
    ws2.set_row(3, 30)

    for i, r in enumerate(results):
        row = 4 + i
        ad_cnt = sum(1 for f in r["findings"] if f["category"] == "표시광고")
        disc_cnt = sum(1 for f in r["findings"] if f["category"] == "고시정보")
        pa_cnt = sum(1 for f in r["findings"] if f["category"] == "구매대행")
        vals = [
            r["sku_id"], r["product_name"], r["category"], r["group"], r["grade"], r["score"],
            r["summary"], r["track"], r["cert_no"] or "-", STATUS_KO.get(r["cert_status"], r["cert_status"]),
            r["cert_state"] or "-", r["recall"], ad_cnt, disc_cnt, pa_cnt, r["next_action"],
        ]
        for c, v in enumerate(vals):
            if c == 4:
                fmt = _fmt(wb, border=1, border_color=BORDER, align="center", bold=True,
                           bg_color=GRADE_FILL.get(r["grade"], "#FFFFFF"),
                           font_color=GRADE_INK.get(r["grade"], INK))
                ws2.write(row, c, v, fmt)
            elif c == 5:
                ws2.write(row, c, v, f_score)
            elif c in (12, 13, 14):
                fmt = _fmt(wb, border=1, border_color=BORDER, align="center",
                           font_color="#A61B29" if v else INK, bold=bool(v))
                ws2.write(row, c, v, fmt)
            elif c == 6:
                ws2.write(row, c, v, _fmt(wb, border=1, border_color=BORDER, text_wrap=True, valign="top",
                                          font_color=GRADE_INK.get(r["grade"], INK)))
            else:
                ws2.write(row, c, v, f_cell)
        ws2.set_row(row, 32)

    ws2.autofilter(3, 0, 3 + len(results), len(COLS_SKU) - 1)
    ws2.freeze_panes(4, 2)
    ws2.set_landscape()
    ws2.fit_to_pages(1, 0)

    # ---------------------------------------------------------------- 상세 점검
    ws3 = wb.add_worksheet("상세 점검")
    ws3.hide_gridlines(2)
    head3 = ["SKU", "상품명", "구분", "심각도", "점검 항목", "판정 내용", "근거(조회·탐지)",
             "근거 조문", "예상 처분", "수정안·다음 조치", "신뢰도"]
    widths3 = [9, 26, 9, 7, 26, 46, 30, 28, 26, 44, 7]
    ws3.merge_range(0, 0, 0, len(head3) - 1, "상세 점검 항목 (SKU × 항목)", f_title)
    ws3.set_row(0, 30)
    ws3.merge_range(1, 0, 1, len(head3) - 1,
                    "심각도 순 · 각 항목은 근거 조문과 예상 처분을 함께 표기합니다.", f_sub)
    for c, (name, w) in enumerate(zip(head3, widths3)):
        ws3.set_column(c, c, w)
        ws3.write(3, c, name, f_head)
    ws3.set_row(3, 28)

    order = {"치명": 0, "중대": 1, "보통": 2, "경미": 3, "정보": 4}
    row = 4
    for r in results:
        fs = sorted(r["findings"], key=lambda f: order.get(f["severity"], 5))
        for f in fs:
            ws3.write(row, 0, r["sku_id"], f_cell_c)
            ws3.write(row, 1, r["product_name"], f_cell)
            ws3.write(row, 2, f["category"], f_cell_c)
            ws3.write(row, 3, f["severity"], _fmt(wb, border=1, border_color=BORDER, align="center",
                                                 bold=True, font_color="#FFFFFF",
                                                 bg_color=SEV_FILL.get(f["severity"], GREY)))
            ws3.write(row, 4, f["title"], _fmt(wb, border=1, border_color=BORDER, text_wrap=True,
                                              valign="top", bold=True, font_color=INK))
            ws3.write(row, 5, f["detail"], f_cell)
            ws3.write(row, 6, f["evidence"], f_cell)
            ws3.write(row, 7, f["law"], f_cell)
            ws3.write(row, 8, f["penalty_text"], f_cell)
            ws3.write(row, 9, (f["fix"] + "\n→ " + f["action"]).strip("→ "), f_cell)
            ws3.write(row, 10, f["confidence"], f_cell_c)
            ws3.set_row(row, _est_height(
                [r["sku_id"], r["product_name"], f["category"], f["severity"], f["title"], f["detail"],
                 f["evidence"], f["law"], f["penalty_text"], f["fix"] + f["action"], f["confidence"]],
                widths3, base=11.2))
            row += 1
    ws3.autofilter(3, 0, row - 1, len(head3) - 1)
    ws3.freeze_panes(4, 2)
    ws3.set_landscape()
    ws3.fit_to_pages(1, 0)

    # ---------------------------------------------------------------- 수정 문구
    ws4 = wb.add_worksheet("수정 문구 제안")
    ws4.hide_gridlines(2)
    head4 = ["SKU", "상품명", "위치", "문제 표현", "수정 문구 제안", "근거 조문"]
    widths4 = [9, 30, 12, 34, 62, 34]
    ws4.merge_range(0, 0, 0, len(head4) - 1, "수정 문구 제안 — 그대로 붙여 쓸 수 있는 대체안", f_title)
    ws4.set_row(0, 30)
    ws4.merge_range(1, 0, 1, len(head4) - 1,
                    "탐지된 표현별로 위법 소지를 없앤 문장을 제안합니다. 최종 문안은 사업자 판단으로 확정하세요.", f_sub)
    for c, (name, w) in enumerate(zip(head4, widths4)):
        ws4.set_column(c, c, w)
        ws4.write(3, c, name, f_head)
    row = 4
    for r in results:
        for f in r["findings"]:
            if f["category"] not in ("표시광고",) and not f.get("rule_id", "").startswith("COSM"):
                continue
            loc = "상품명" if any(t in r["product_name"] for t in re.findall(r"'([^']+)'", f["evidence"])) else "상세페이지"
            ws4.write(row, 0, r["sku_id"], f_cell_c)
            ws4.write(row, 1, r["product_name"], f_cell)
            ws4.write(row, 2, loc, f_cell_c)
            ws4.write(row, 3, f["evidence"].replace("탐지 표현: ", ""), f_cell)
            ws4.write(row, 4, f["fix"], _fmt(wb, border=1, border_color=BORDER, text_wrap=True,
                                            valign="top", bg_color="#FFFBF0"))
            ws4.write(row, 5, f["law"], f_cell)
            ws4.set_row(row, _est_height(
                [r["sku_id"], r["product_name"], loc, f["evidence"], f["fix"], f["law"]],
                widths4, base=11.6))
            row += 1
    ws4.freeze_panes(4, 2)
    ws4.set_landscape()

    # ---------------------------------------------------------------- 점검 기준
    ws5 = wb.add_worksheet("점검 기준")
    ws5.hide_gridlines(2)
    ws5.set_column("A:A", 3)
    ws5.set_column("B:B", 34)
    ws5.set_column("C:C", 90)
    ws5.merge_range("B2:C2", "점검 기준 · 출처", f_title)
    ws5.write("B4", "룰셋", _fmt(wb, bold=True, bg_color=LIGHT, border=1, border_color=BORDER))
    ws5.write("C4", f"{data['ruleset']['id']} v{data['ruleset']['version']} (법령 기준일 {data['ruleset']['law_basis_date']})", f_value)
    ws5.write("B5", "적용 법령", _fmt(wb, bold=True, bg_color=LIGHT, border=1, border_color=BORDER))
    ws5.write("C5", "\n".join(data["ruleset"]["laws"]), _fmt(wb, border=1, border_color=BORDER, text_wrap=True, valign="top"))
    ws5.set_row(4, 70)
    r = 6
    ws5.write(r, 1, "출처", _fmt(wb, bold=True, bg_color=LIGHT, border=1, border_color=BORDER))
    ws5.write(r, 2, "링크 · 확인일", _fmt(wb, bold=True, bg_color=LIGHT, border=1, border_color=BORDER))
    r += 1
    for s in data["ruleset"]["sources"]:
        ws5.write(r, 1, s["name"], f_cell)
        ws5.write(r, 2, f"{s['url']}\n({s['checked']}) {s.get('note','')}", _fmt(wb, border=1, border_color=BORDER, text_wrap=True, valign="top", font_size=9))
        r += 1
    r += 1
    ws5.write(r, 1, "확인하지 못한 사항", _fmt(wb, bold=True, bg_color=LIGHT, border=1, border_color=BORDER))
    ws5.write(r, 2, "\n".join("· " + u for u in data["ruleset"]["unverified"]),
              _fmt(wb, border=1, border_color=BORDER, text_wrap=True, valign="top"))
    ws5.set_row(r, 70)
    r += 2
    ws5.write(r, 1, "면책", _fmt(wb, bold=True, bg_color=LIGHT, border=1, border_color=BORDER))
    ws5.write(r, 2, data["ruleset"]["disclaimer"], _fmt(wb, border=1, border_color=BORDER, text_wrap=True, valign="top"))
    ws5.set_row(r, 60)

    # ---------------------------------------------------------------- 입력 데이터
    ws6 = wb.add_worksheet("입력 데이터")
    ws6.hide_gridlines(2)
    head6 = ["SKU", "상품명", "카테고리", "모델명", "가격", "KC 인증번호", "구매대행", "원산지", "고시정보 제공 항목", "상세설명"]
    widths6 = [9, 34, 18, 16, 10, 18, 8, 10, 50, 70]
    for c, (name, w) in enumerate(zip(head6, widths6)):
        ws6.set_column(c, c, w)
        ws6.write(0, c, name, f_head)
    ws6.set_row(0, 26)
    for i, r in enumerate(results):
        ws6.write(1 + i, 0, r["sku_id"], f_cell_c)
        ws6.write(1 + i, 1, r["product_name"], f_cell)
        ws6.write(1 + i, 2, r["category"], f_cell)
        ws6.write(1 + i, 3, r["model_name"], f_cell)
        ws6.write(1 + i, 4, r["price"], f_cell_c)
        ws6.write(1 + i, 5, r["cert_no"], f_cell_c)
        ws6.write(1 + i, 6, "Y" if r["is_purchase_agency"] else "N", f_cell_c)
        ws6.write(1 + i, 7, r["origin"], f_cell_c)
        ws6.write(1 + i, 8, r["disclosure_fields"], f_cell)
        ws6.write(1 + i, 9, r["detail_text"], f_cell)
        ws6.set_row(1 + i, 44)
    wb.close()
    return out_path
