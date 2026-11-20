"""셀러팩트 감사 리포트 — PDF 생성기 (reportlab + 나눔고딕)."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, Image, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle)

ROOT = Path(__file__).resolve().parent.parent
def _font_dir() -> Path:
    """나눔고딕이 있는 디렉터리 — 저장소 동봉 폰트 우선, 없으면 홈 폰트."""
    for cand in (ROOT / "assets" / "fonts", ROOT / "fonts", Path.home() / ".fonts"):
        if (cand / "NanumGothic-Regular.ttf").exists():
            return cand
    return Path.home() / ".fonts"


FONT_DIR = _font_dir()
FONT_R = FONT_DIR / "NanumGothic-Regular.ttf"
FONT_B = FONT_DIR / "NanumGothic-Bold.ttf"

NAVY = colors.HexColor("#0F1B2D")
INK = colors.HexColor("#1C2434")
GREY = colors.HexColor("#5A6474")
LIGHT = colors.HexColor("#F4F6FA")
LINE = colors.HexColor("#D8DEE9")
ACCENT = colors.HexColor("#F2A93B")
WHITE = colors.white
GRADE_C = {"A": "#1B7F5A", "B": "#4C8C3F", "C": "#A8720B", "D": "#C4571F", "E": "#A61B29"}
GRADE_BG = {"A": "#E9F5EF", "B": "#EFF6E7", "C": "#FDF3DF", "D": "#FBEAE1", "E": "#F9E2E4"}
SEV_C = {"치명": "#A61B29", "중대": "#C4571F", "보통": "#A8720B", "경미": "#5A6474", "정보": "#8A93A3"}

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm


def register_fonts():
    if "NanumGothic" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("NanumGothic", str(FONT_R)))
        pdfmetrics.registerFont(TTFont("NanumGothic-Bold", str(FONT_B)))
    try:
        from reportlab.pdfbase.pdfmetrics import registerFontFamily
        registerFontFamily("NanumGothic", normal="NanumGothic", bold="NanumGothic-Bold",
                           italic="NanumGothic", boldItalic="NanumGothic-Bold")
    except Exception:
        pass


def styles():
    return {
        "h1": ParagraphStyle("h1", fontName="NanumGothic-Bold", fontSize=23, leading=30, textColor=NAVY),
        "h2": ParagraphStyle("h2", fontName="NanumGothic-Bold", fontSize=13.5, leading=18, textColor=NAVY,
                             spaceBefore=2, spaceAfter=4),
        "h3": ParagraphStyle("h3", fontName="NanumGothic-Bold", fontSize=11, leading=15, textColor=INK),
        "body": ParagraphStyle("body", fontName="NanumGothic", fontSize=9.2, leading=13.6, textColor=INK),
        "small": ParagraphStyle("small", fontName="NanumGothic", fontSize=8, leading=11.6, textColor=GREY),
        "tiny": ParagraphStyle("tiny", fontName="NanumGothic", fontSize=7.2, leading=10, textColor=GREY),
        "cell": ParagraphStyle("cell", fontName="NanumGothic", fontSize=8, leading=11, textColor=INK),
        "cellb": ParagraphStyle("cellb", fontName="NanumGothic-Bold", fontSize=8, leading=11, textColor=INK),
        "cellw": ParagraphStyle("cellw", fontName="NanumGothic-Bold", fontSize=8, leading=11, textColor=WHITE),
        "right": ParagraphStyle("right", fontName="NanumGothic", fontSize=8, leading=11, textColor=GREY,
                                alignment=TA_RIGHT),
        "center": ParagraphStyle("center", fontName="NanumGothic", fontSize=9, leading=12, textColor=INK,
                                 alignment=TA_CENTER),
    }


def P(text: str, style) -> Paragraph:
    """본문 문단. 우리가 넣는 마크업(<b>, <br/>, <font>)은 살리고,
    원문 데이터의 앰퍼샌드만 안전하게 이스케이프한다."""
    s = str(text)
    s = re.sub(r"&(?!(amp|lt|gt|quot|apos|#\d+);)", "&amp;", s)
    s = s.replace("\n", "<br/>")
    return Paragraph(s, style)


class Report:
    def __init__(self, data: dict, out_path: Path, brand_name="셀러팩트", brand_en="SellerFact"):
        self.data = data
        self.out = Path(out_path)
        self.brand = brand_name
        self.brand_en = brand_en
        self.st = styles()
        self.section_title = ""
        register_fonts()

    # ---------------------------------------------------------------- chrome
    def _cover(self, canvas, doc):
        canvas.saveState()
        canvas.setFillColor(NAVY)
        canvas.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
        canvas.setFillColor(ACCENT)
        canvas.rect(0, PAGE_H - 6 * mm, PAGE_W, 6 * mm, stroke=0, fill=1)
        # brand mark
        canvas.setFillColor(WHITE)
        canvas.setFont("NanumGothic-Bold", 20)
        canvas.drawString(MARGIN, PAGE_H - 40 * mm, self.brand)
        canvas.setFont("NanumGothic", 8.5)
        canvas.setFillColor(colors.HexColor("#9FA9B8"))
        canvas.drawString(MARGIN, PAGE_H - 46 * mm, self.brand_en.upper() + "  ·  PRODUCT REGULATORY PRE-CHECK")
        canvas.setStrokeColor(colors.HexColor("#2B3A52"))
        canvas.setLineWidth(0.6)
        canvas.line(MARGIN, PAGE_H - 52 * mm, PAGE_W - MARGIN, PAGE_H - 52 * mm)
        canvas.setFont("NanumGothic", 7)
        canvas.setFillColor(colors.HexColor("#5C6879"))
        canvas.drawRightString(PAGE_W - MARGIN, 18 * mm,
                               "CONFIDENTIAL · 구매자에게만 제공되는 리포트입니다")
        canvas.restoreState()

    def _cover_grade_panel(self):
        results = self.data["results"]
        counts: dict = {}
        for r in results:
            counts[r["grade"]] = counts.get(r["grade"], 0) + 1
        rows = [[P("위험 등급 분포", ParagraphStyle("gp", fontName="NanumGothic", fontSize=7.5,
                                              textColor=colors.HexColor("#9FA9B8")))]]
        for g, label in [("A", "양호"), ("B", "보완 권장"), ("C", "보완 필요"), ("D", "위험"), ("E", "긴급")]:
            n = counts.get(g, 0)
            line = Table([[P(f"{g}", ParagraphStyle("gg", fontName="NanumGothic-Bold", fontSize=9,
                                                    textColor=colors.HexColor(GRADE_C[g]))),
                           P(label, ParagraphStyle("gl", fontName="NanumGothic", fontSize=8.5,
                                                   textColor=WHITE if n else colors.HexColor("#5C6879"))),
                           P(str(n), ParagraphStyle("gn", fontName="NanumGothic-Bold", fontSize=9.5,
                                                    textColor=WHITE if n else colors.HexColor("#5C6879"),
                                                    alignment=TA_RIGHT))]],
                         colWidths=[7 * mm, 27 * mm, 10 * mm])
            line.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                     ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
            rows.append([line])
        t = Table(rows, colWidths=[44 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#16243A")),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#2B3A52")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t


    def _body(self, canvas, doc):
        canvas.saveState()
        canvas.setFillColor(NAVY)
        canvas.rect(0, PAGE_H - 14 * mm, PAGE_W, 14 * mm, stroke=0, fill=1)
        canvas.setFillColor(WHITE)
        canvas.setFont("NanumGothic-Bold", 9)
        canvas.drawString(MARGIN, PAGE_H - 9.2 * mm, self.brand)
        canvas.setFont("NanumGothic", 7.5)
        canvas.setFillColor(colors.HexColor("#9FA9B8"))
        canvas.drawString(MARGIN + 22 * mm, PAGE_H - 9.2 * mm, self.brand_en.upper())
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 9.2 * mm,
                               f"{self.data.get('client','')} · 상품 규제 사전 점검 리포트")
        canvas.setStrokeColor(LINE)
        canvas.setLineWidth(0.6)
        canvas.line(MARGIN, 14 * mm, PAGE_W - MARGIN, 14 * mm)
        canvas.setFillColor(GREY)
        canvas.setFont("NanumGothic", 7)
        canvas.drawString(MARGIN, 9.5 * mm,
                          f"셀러팩트(SellerFact) · 룰셋 {self.data['ruleset']['version']} · {self.data['generated_at']}")
        canvas.drawRightString(PAGE_W - MARGIN, 9.5 * mm, f"{doc.page}")
        canvas.restoreState()

    # ---------------------------------------------------------------- blocks
    def _kpi_row(self, results, sev_counts, action_needed):
        cells = [
            ("점검 SKU", f"{len(results)}", "개"),
            ("조치 필요", f"{action_needed}", "개 (C~E)"),
            ("치명 위험", f"{sev_counts.get('치명', 0)}", "건"),
            ("발견 항목", f"{sum(sev_counts.values())}", "건"),
        ]
        row = []
        for label, val, unit in cells:
            inner = Table([[P(val, ParagraphStyle("n", fontName="NanumGothic-Bold", fontSize=22, leading=25,
                                                 textColor=NAVY, alignment=TA_CENTER))],
                           [P(label + "  " + unit, ParagraphStyle("u", fontName="NanumGothic", fontSize=7.5,
                                                                  leading=10, textColor=GREY,
                                                                  alignment=TA_CENTER))]],
                          colWidths=[(PAGE_W - 2 * MARGIN) / 4 - 4 * mm])
            inner.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                ("BOX", (0, 0), (-1, -1), 0.6, LINE),
                ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ]))
            row.append(inner)
        t = Table([row], colWidths=[(PAGE_W - 2 * MARGIN) / 4] * 4)
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
        return t

    def _grade_bars(self, counts, total):
        rows = []
        for g, label, desc in [("A", "양호", "즉시 조치 없음"), ("B", "보완 권장", "경미한 누락"),
                               ("C", "보완 필요", "과태료 소지"), ("D", "위험", "판매 전 수정"),
                               ("E", "긴급", "판매 중지 검토")]:
            n = counts.get(g, 0)
            bar_w = 0 if total == 0 else max(2.0, 46.0 * n / total)
            bar = Table([[""]], colWidths=[bar_w * mm], rowHeights=[5.2 * mm])
            bar.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(GRADE_C[g])),
                                     ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                     ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
            rows.append([P(f"<b>{g}</b>", ParagraphStyle("g", fontName="NanumGothic-Bold", fontSize=9,
                                                          textColor=colors.HexColor(GRADE_C[g]))),
                         P(f"{label}", self.st["cell"]), P(desc, self.st["small"]), bar,
                         P(f"{n}개", self.st["cellb"])])
        t = Table(rows, colWidths=[10 * mm, 22 * mm, 32 * mm, 52 * mm, 14 * mm])
        t.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (3, 0), (3, -1), "LEFT"),
            ("LINEBELOW", (0, 0), (-1, -2), 0.4, LINE),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    def _finding_table(self, findings):
        if not findings:
            t = Table([[P("탐지된 위험 항목이 없습니다.", self.st["cell"])]], colWidths=[PAGE_W - 2 * MARGIN])
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E9F5EF")),
                                   ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#B9DCC9")),
                                   ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
            return t
        head = [P("심각도", self.st["cellw"]), P("점검 항목", self.st["cellw"]),
                P("근거 · 조문", self.st["cellw"]), P("수정안 / 다음 조치", self.st["cellw"])]
        data = [head]
        for f in findings:
            evidence = f.get("evidence") or ""
            law = f.get("law") or ""
            pen = f.get("penalty_text") or ""
            basis = "<br/>".join(x for x in [evidence, law, pen] if x)
            fix = f.get("fix") or ""
            action = f.get("action") or ""
            fix_html = fix + (f"<br/><font color='#5A6474'>→ {action}</font>" if action else "")
            data.append([P(f["severity"], ParagraphStyle("s", fontName="NanumGothic-Bold", fontSize=8,
                                                          textColor=colors.HexColor(SEV_C.get(f["severity"], "#5A6474")))),
                         P(f["title"], self.st["cellb"]),
                         P(basis, self.st["tiny"]),
                         P(fix_html, self.st["cell"])])
        widths = [16 * mm, 42 * mm, 56 * mm, 60 * mm]
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.4, LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#FAFBFD")]),
        ]))
        return t

    def _sku_card(self, r):
        g = r["grade"]
        badge = Table([[P(f"<b>{g}</b>", ParagraphStyle("bg", fontName="NanumGothic-Bold", fontSize=15,
                                                        textColor=colors.HexColor(GRADE_C[g]),
                                                        alignment=TA_CENTER))],
                       [P(f"{r['score']}점", ParagraphStyle("bs", fontName="NanumGothic", fontSize=7,
                                                            textColor=GREY, alignment=TA_CENTER))]],
                      colWidths=[13 * mm], rowHeights=[8 * mm, 4.5 * mm])
        badge.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(GRADE_BG[g])),
                                   ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(GRADE_C[g])),
                                   ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                   ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))

        title = P(f"<b>{r['product_name']}</b>", self.st["h3"])
        meta = P(f"{r['sku_id']} · {r['category']} · 상품군 {r['group']} · 모델 {r['model_name'] or '-'}", self.st["small"])
        head = Table([[badge, [title, meta]]], colWidths=[15 * mm, PAGE_W - 2 * MARGIN - 15 * mm])
        head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                  ("LEFTPADDING", (0, 0), (0, 0), 0),
                                  ("LEFTPADDING", (1, 0), (1, 0), 4),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))

        cert_status_ko = {"verified": "확인됨", "not_found": "조회 결과 없음", "unavailable": "조회 불가",
                          "format_error": "형식 오류", "skipped": "미입력", "not_applicable": "해당 없음"}
        facts = [
            ("KC 1차 판정", f"{r['track']} · 신뢰도 {r['track_confidence']}"),
            ("판정 근거", r["track_basis"][:96]),
            ("인증번호", f"{r['cert_no'] or '-'} · {cert_status_ko.get(r['cert_status'], r['cert_status'])}"
                       + (f" · 상태 {r['cert_state']}" if r["cert_state"] else "")),
            ("리콜 조회", r["recall"] if r["cert_status"] in ("verified", "not_found") else "미조회(인증번호 없음/해당 없음)"),
            ("구매대행", "예" if r["is_purchase_agency"] else "아니오"),
        ]
        fact_rows = [[P(k, ParagraphStyle("fk", fontName="NanumGothic", fontSize=7.4, textColor=GREY)),
                      P(v, self.st["tiny"])] for k, v in facts if v]
        ft = Table(fact_rows, colWidths=[20 * mm, PAGE_W - 2 * MARGIN - 20 * mm])
        ft.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("TOPPADDING", (0, 0), (-1, -1), 1.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0)]))

        next_line = P(f"<b>다음 조치</b> · {r['next_action']}", self.st["small"])
        return KeepTogether([head, Spacer(1, 3), ft, Spacer(1, 3),
                             self._finding_table(sorted(r["findings"], key=lambda f: ["치명", "중대", "보통", "경미", "정보"].index(f["severity"]))),
                             Spacer(1, 2), next_line, Spacer(1, 7)])

    # ---------------------------------------------------------------- build
    def build(self) -> Path:
        data = self.data
        results = data["results"]
        counts: dict = {}
        sev_counts: dict = {}
        for r in results:
            counts[r["grade"]] = counts.get(r["grade"], 0) + 1
            for f in r["findings"]:
                sev_counts[f["severity"]] = sev_counts.get(f["severity"], 0) + 1
        action_needed = sum(1 for r in results if r["grade"] in ("C", "D", "E"))
        total = len(results)

        doc = BaseDocTemplate(str(self.out), pagesize=A4,
                              leftMargin=MARGIN, rightMargin=MARGIN,
                              topMargin=22 * mm, bottomMargin=20 * mm,
                              title=f"{self.brand} 상품 규제 사전 점검 리포트",
                              author=self.brand, subject="KC·표시광고·고시정보 사전 점검")
        frame_cover = Frame(MARGIN, 24 * mm, PAGE_W - 2 * MARGIN, PAGE_H - 80 * mm, id="cover",
                            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        frame_body = Frame(MARGIN, 18 * mm, PAGE_W - 2 * MARGIN, PAGE_H - 44 * mm, id="body",
                           leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        doc.addPageTemplates([
            PageTemplate(id="cover", frames=[frame_cover], onPage=self._cover),
            PageTemplate(id="body", frames=[frame_body], onPage=self._body),
        ])

        S: list = []
        st = self.st

        # ---------- cover
        S.append(Spacer(1, 28 * mm))
        S.append(P("상품 규제 사전 점검 리포트", ParagraphStyle("ct", fontName="NanumGothic-Bold", fontSize=30,
                                                              leading=38, textColor=WHITE)))
        S.append(Spacer(1, 4 * mm))
        S.append(P("KC 인증 · 표시광고 · 상품정보제공고시", ParagraphStyle("cs", fontName="NanumGothic", fontSize=12,
                                                                       leading=18,
                                                                       textColor=colors.HexColor("#9FA9B8"))))
        S.append(Spacer(1, 26 * mm))
        cover_rows = [
            ["점검 대상", data.get("client", "-")],
            ["점검 SKU", f"{total}개"],
            ["점검일", data["generated_at"]],
            ["룰셋 버전", data["ruleset"]["version"]],
            ["법령 기준일", data["ruleset"]["law_basis_date"]],
            ["조치 필요", f"{action_needed}개 SKU (등급 C~E)"],
        ]
        t = Table([[P(k, ParagraphStyle("ck", fontName="NanumGothic", fontSize=9, textColor=colors.HexColor("#9FA9B8"))),
                    P(v, ParagraphStyle("cv", fontName="NanumGothic-Bold", fontSize=10.5, textColor=WHITE))]
                   for k, v in cover_rows], colWidths=[26 * mm, 66 * mm])
        t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#2B3A52")),
                               ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                               ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
        outer = Table([[t, self._cover_grade_panel()]], colWidths=[96 * mm, 74 * mm])
        outer.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
        S.append(outer)
        S.append(Spacer(1, 14 * mm))
        S.append(P("담긴 내용", ParagraphStyle("cc", fontName="NanumGothic-Bold", fontSize=9,
                                          textColor=colors.HexColor("#F2A93B"))))
        S.append(Spacer(1, 2 * mm))
        items = [
            "KC 인증 대상·트랙 1차 판정 (2026-04-01 개정 반영)",
            "인증번호 유효성·인증상태·리콜 실시간 조회",
            "표시광고 금지·위험 표현 탐지와 수정 문구",
            "상품정보제공고시·구매대행 고지·원산지 점검",
            "위반별 근거 조문·처분 범위·다음 조치",
        ]
        for it in items:
            S.append(P("· " + it, ParagraphStyle("ci", fontName="NanumGothic", fontSize=8,
                                                  leading=13, textColor=colors.HexColor("#9FA9B8"))))
        S.append(Spacer(1, 22 * mm))
        S.append(P("본 리포트는 공개 법령·고시와 공공 인증정보를 기준으로 한 1차 스크리닝 결과입니다. "
                   "최종 인증 트랙 판정과 위법성 판단은 지정 시험·인증기관 및 관할 행정기관의 확인이 필요합니다.",
                   ParagraphStyle("cd", fontName="NanumGothic", fontSize=7.8, leading=12,
                                  textColor=colors.HexColor("#7C8798"))))
        S.append(NextPageTemplate("body"))
        S.append(PageBreak())

        # ---------- summary
        S.append(P("01  요약", st["h2"]))
        S.append(P("무엇을 먼저 고쳐야 하는지부터 정리했습니다.", st["small"]))
        S.append(Spacer(1, 3 * mm))
        S.append(self._kpi_row(results, sev_counts, action_needed))
        S.append(Spacer(1, 6 * mm))
        S.append(P("등급 분포", st["h3"]))
        S.append(Spacer(1, 1.5 * mm))
        S.append(self._grade_bars(counts, total))

        S.append(Spacer(1, 5 * mm))
        S.append(P("가장 먼저 고칠 항목", st["h3"]))
        S.append(Spacer(1, 1.5 * mm))
        prio = []
        for r in results:
            for f in r["findings"]:
                if f["severity"] in ("치명", "중대"):
                    prio.append((0 if f["severity"] == "치명" else 1, r, f))
        prio.sort(key=lambda x: x[0])
        pdata = [[P("SKU", st["cellw"]), P("심각도", st["cellw"]), P("항목", st["cellw"]),
                  P("근거 조문 · 예상 처분", st["cellw"])]]
        for _, r, f in prio[:6]:
            pdata.append([P(r["sku_id"], st["cell"]),
                          P(f["severity"], ParagraphStyle("ps", fontName="NanumGothic-Bold", fontSize=8,
                                                          textColor=colors.HexColor(SEV_C.get(f["severity"], "#5A6474")))),
                          P(f["title"], st["cell"]),
                          P(f"{(f['law'] or '-')}<br/><font color='#5A6474'>{f['penalty_text'] or ''}</font>",
                            st["tiny"])])
        pt = Table(pdata, colWidths=[14 * mm, 15 * mm, 66 * mm, 79 * mm], repeatRows=1)
        pt.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.4, LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#FAFBFD")]),
        ]))
        S.append(pt)
        S.append(Spacer(1, 5 * mm))
        # 총평
        crit = [f for r in results for f in r["findings"] if f["severity"] == "치명"]
        by_cat: dict = {}
        for f in crit:
            by_cat[f["category"]] = by_cat.get(f["category"], 0) + 1
        parts = [f"{v}건은 {k}" for k, v in sorted(by_cat.items(), key=lambda x: -x[1])]
        e_cnt = counts.get("E", 0)
        review = (
            f"점검한 {total}개 SKU 중 <b>{action_needed}개</b>에서 조치가 필요합니다. "
            + (f"치명 위험 {len(crit)}건은 " + ", ".join(parts) + "에서 나왔습니다. " if crit else "")
            + (f"등급 E {e_cnt}개는 지금 판매를 잠시 내리는 것이 가장 빠른 손실 방지책입니다. " if e_cnt else "")
            + "인증은 있는데 문구만 문제인 SKU는 문장 교체만으로 위험을 크게 낮출 수 있으니, "
              "03의 수정 문구를 먼저 반영하고 인증이 없는 SKU는 04 로드맵 순서로 처리하시면 됩니다."
        )
        S.append(P("총평", st["h3"]))
        box = Table([[P(review, st["body"])]], colWidths=[PAGE_W - 2 * MARGIN])
        box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FBF6EA")),
                                 ("LINEBEFORE", (0, 0), (0, -1), 2.2, ACCENT),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                                 ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
        S.append(box)
        S.append(Spacer(1, 5 * mm))
        S.append(P("이 리포트 읽는 법", st["h3"]))
        S.append(P("1) 등급 A~E는 발견 항목의 심각도 가중치 합계로 계산합니다(E가 가장 위험). "
                   "2) '인증 조회'는 제품안전정보센터 실시간 조회 결과이며, 조회 불가는 통과가 아닙니다. "
                   "3) 심각도 '정보'는 위반이 아니라 확인이 필요한 항목입니다. "
                   "4) 수정 문구는 그대로 붙여 쓸 수 있게 작성했고, 최종 문안 확정은 사업자 판단입니다.", st["small"]))
        S.append(PageBreak())

        # ---------- per SKU
        S.append(P("02  SKU별 상세 점검", st["h2"]))
        S.append(P("각 상품의 1차 판정과 실제 인증 조회 결과, 고칠 문장을 함께 담았습니다.", st["small"]))
        S.append(Spacer(1, 3 * mm))
        for r in results:
            S.append(self._sku_card(r))

        # ---------- 수정 문구
        S.append(PageBreak())
        S.append(P("03  수정 문구 제안", st["h2"]))
        S.append(P("탐지된 표현을 위법 소지가 없는 문장으로 바꾼 안입니다. 상세페이지에 그대로 반영할 수 있습니다.", st["small"]))
        S.append(Spacer(1, 3 * mm))
        rows = [[P("SKU", st["cellw"]), P("문제 표현", st["cellw"]), P("수정 문구 제안", st["cellw"]), P("근거", st["cellw"])]]
        has_any = False
        for r in results:
            for f in r["findings"]:
                if f["category"] != "표시광고":
                    continue
                has_any = True
                rows.append([P(r["sku_id"], st["cell"]),
                             P(f["evidence"].replace("탐지 표현: ", ""), st["cell"]),
                             P(f["fix"], st["cell"]), P(f["law"], st["tiny"])])
        if not has_any:
            rows.append([P("-", st["cell"]), P("탐지된 위험 표현 없음", st["cell"]), P("-", st["cell"]), P("-", st["cell"])])
        tt = Table(rows, colWidths=[14 * mm, 48 * mm, 76 * mm, 36 * mm], repeatRows=1)
        tt.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.4, LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#FFFBF2")]),
        ]))
        S.append(tt)

        # ---------- 로드맵
        S.append(Spacer(1, 7 * mm))
        S.append(P("04  조치 로드맵", st["h2"]))
        S.append(P("위험도가 높은 항목부터 영업일 기준으로 정리했습니다.", st["small"]))
        S.append(Spacer(1, 2.5 * mm))
        roadmap = [
            ("D+0", "판매 중지·문구 차단", "치명 등급 SKU는 판매를 잠시 내리고, 인증번호·효능 표현을 상세페이지에서 내립니다."),
            ("D+1~3", "인증 확인·정정", "보유 인증서 원본으로 인증번호·모델명을 대조하고, 인증기관에 트랙·모델 포함 여부를 확인합니다."),
            ("D+3~7", "고시정보·고지 보완", "상품정보제공고시 표와 구매대행 고지, 원산지 표시를 채웁니다."),
            ("D+7~30", "신규 인증·재판매", "인증이 없는 SKU는 시험·신고를 진행하고, 통과 후 인증정보와 함께 재판매합니다."),
            ("상시", "월간 모니터링", "신규 상품·리콜·법령 개정을 월 1회 점검해 리포트로 갱신합니다."),
        ]
        rdata = [[P(k, ParagraphStyle("rk", fontName="NanumGothic-Bold", fontSize=9, textColor=WHITE)),
                  P(v, st["cellb"]), P(d, st["cell"])] for k, v, d in roadmap]
        rt = Table(rdata, colWidths=[18 * mm, 40 * mm, 116 * mm])
        rt.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), NAVY), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        S.append(rt)

        # ---------- 기준·출처
        S.append(PageBreak())
        S.append(P("05  점검 기준과 출처", st["h2"]))
        S.append(P(f"룰셋 {data['ruleset']['id']} v{data['ruleset']['version']} · 법령 기준일 {data['ruleset']['law_basis_date']}", st["small"]))
        S.append(Spacer(1, 2 * mm))
        S.append(P("적용 법령", st["h3"]))
        for law in data["ruleset"]["laws"]:
            S.append(P("· " + law, st["tiny"]))
        S.append(Spacer(1, 3 * mm))
        S.append(P("출처·조회 채널", st["h3"]))
        for s in data["ruleset"]["sources"]:
            S.append(P(f"· {s['name']} — <font color='#5A6474'>{s['url']}</font> ({s['checked']})", st["tiny"]))
        S.append(Spacer(1, 3 * mm))
        S.append(P("확인하지 못한 사항", st["h3"]))
        for u in data["ruleset"]["unverified"]:
            S.append(P("· " + u, st["tiny"]))
        S.append(Spacer(1, 4 * mm))
        box = Table([[P("<b>면책</b>", st["cellb"])], [P(data["ruleset"]["disclaimer"], st["small"])]],
                    colWidths=[PAGE_W - 2 * MARGIN])
        box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("BOX", (0, 0), (-1, -1), 0.5, LINE),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                                 ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        S.append(box)

        doc.build(S)
        return self.out


def build(data: dict, out_path: Path, brand_name="셀러팩트", brand_en="SellerFact") -> Path:
    return Report(data, out_path, brand_name, brand_en).build()
