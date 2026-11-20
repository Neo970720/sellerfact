"""
셀러팩트(SellerFact) 감사 엔진 — KC 인증·표시광고·고시정보 사전 점검

입력 : 셀러 상품 목록 CSV/XLSX  (sku_id, product_name, category, model_name, detail_text,
       kc_cert_no, is_purchase_agency, origin, disclosure_fields, price)
처리 : ① KC 대상·트랙 1차 판정(룰셋 버전 관리)
       ② 인증번호 유효성·리콜 실측 조회(제품안전정보센터 공개 조회, 실패 시 상태 표기)
       ③ 표시광고 금지·위험 표현 탐지
       ④ 전자상거래 고시정보 누락·구매대행 고지 점검
       ⑤ 위반별 근거 조문·처분 범위·수정 문구·다음 조치 생성
출력 : 결과 JSON + 엑셀 리포트 + PDF 리포트

사용:
  python3 src/audit_engine.py --input sample/sample_sellers_10sku.csv --out output \
      --client "샘플 셀러(주)" --brand-name "셀러팩트"
  python3 src/audit_engine.py ... --no-network   # 조회 없이(오프라인) 실행
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional

SRC = Path(__file__).resolve().parent
sys.path.insert(0, str(SRC))

from kc_client import SafetyKoreaClient  # noqa: E402

RULESET_PATH = SRC / "rules" / "ruleset.json"

# 이 유형이 하나라도 나오면 해당 SKU는 사실상 판매 불가 상태로 본다(등급 E)
SALE_BLOCKING = {"CERT-MISSING", "CERT-STATE", "CERT-RECALL", "PA-CHILD", "COSM-FUNC",
                 "RF-MISSING", "AD-DISEASE", "AD-DRUG", "AD-SUPPLEMENT"}

GROUP_BY_CATEGORY = [
    (("유아", "아기", "아동", "키즈", "어린이", "신생아"), "어린이제품"),
    (("화장품", "스킨", "코스메틱", "뷰티"), "화장품"),
    (("식품", "건강", "홍삼", "영양"), "식품"),
    (("가전", "전기", "충전", "배터리", "조명", "디지털"), "전기용품"),
    (("주방", "생활", "용기", "잡화"), "생활용품"),
]


def load_ruleset(path: Path = RULESET_PATH) -> dict:
    return json.loads(path.read_text("utf-8"))


@dataclass
class Finding:
    category: str          # KC인증 | 표시광고 | 고시정보 | 구매대행 | 원산지 | 일반
    severity: str          # 치명 | 중대 | 보통 | 경미 | 정보
    title: str
    detail: str
    evidence: str = ""
    law: str = ""
    penalty_text: str = ""
    penalty_max: Optional[int] = None
    fix: str = ""
    action: str = ""
    confidence: str = "중간"
    rule_id: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class SkuResult:
    sku_id: str
    product_name: str
    category: str
    group: str
    model_name: str = ""
    price: str = ""
    detail_text: str = ""
    origin: str = ""
    disclosure_fields: str = ""
    rule_id: str = ""
    rule_label: str = ""
    track: str = ""
    track_basis: str = ""
    track_confidence: str = ""
    track_alternatives: list = field(default_factory=list)
    cert_no: str = ""
    cert_status: str = "skipped"
    cert_state: str = ""
    cert_scope: str = ""
    cert_agency: str = ""
    cert_date: str = ""
    cert_item: str = ""
    cert_model: str = ""
    cert_maker: str = ""
    cert_maker_country: str = ""
    recall: str = "미조회"
    is_purchase_agency: bool = False
    findings: list = field(default_factory=list)
    score: int = 100
    grade: str = "A"
    grade_label: str = ""
    summary: str = ""
    next_action: str = ""
    fee_note: str = ""

    def to_dict(self) -> dict:
        d = asdict(self)
        d["findings"] = [f if isinstance(f, dict) else f.to_dict() for f in self.findings]
        return d


class AuditEngine:
    def __init__(self, ruleset: Optional[dict] = None, client: Optional[SafetyKoreaClient] = None,
                 offline: bool = False):
        self.rs = ruleset or load_ruleset()
        self.client = client or (None if offline else SafetyKoreaClient())
        self.offline = offline
        self.weights = {k: v["weight"] for k, v in self.rs["severity"].items()}

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def detect_group(category: str, name: str = "", rule_group: str = "") -> str:
        blob = f"{category} {name}"
        for keys, grp in GROUP_BY_CATEGORY:
            if any(k in blob for k in keys):
                return grp
        return rule_group or "기타"

    def match_product_rule(self, row: dict) -> dict:
        blob = " ".join([str(row.get("product_name", "")), str(row.get("category", "")),
                         str(row.get("detail_text", ""))[:400]])
        cat_group = self.detect_group(str(row.get("category", "")), str(row.get("product_name", "")))
        best, best_score = None, 0
        for rule in self.rs["product_rules"]:
            score = 0
            for kw in rule.get("keywords", []):
                if not kw:
                    continue
                for m in re.finditer(re.escape(kw), blob):
                    # 부정 문맥("무선충전 미지원")은 키워드로 보지 않는다
                    tail = blob[m.end():m.end() + 6]
                    if re.match(r"\s*(미지원|없음|불가|제외|아님|미탑재)", tail):
                        continue
                    score += len(kw)
                    break
            if rule.get("group") == cat_group and cat_group != "기타":
                score += 8  # 카테고리와 같은 군의 룰을 우선
            if score > best_score:
                best, best_score = rule, score
        return best or self.rs["product_rules"][-1]

    def _penalty(self, key: str) -> dict:
        return self.rs["penalties"].get(key, {})

    def _fee(self, track: str) -> str:
        fees = self.rs["fees"]
        flat = track.replace(" ", "")
        for k in fees:
            if k != "note" and k.replace(" ", "") in flat:
                f = fees[k]
                fee = f.get("gov_fee")
                gov = f"정부 수수료 {fee:,}원" if fee else "정부 수수료 별도"
                return f"{gov} + 시험비 별도 · 예상 {f.get('expected_days','-')}"
        return "인증기관 견적 확인 필요"

    # ------------------------------------------------------------------ evaluation
    def evaluate(self, row: dict) -> SkuResult:
        rs = self.rs
        product_name = str(row.get("product_name", "")).strip()
        category = str(row.get("category", "")).strip()
        model_name = str(row.get("model_name", "")).strip()
        detail = str(row.get("detail_text", "")).strip()
        cert_no = str(row.get("kc_cert_no", "") or "").strip()
        is_pa = str(row.get("is_purchase_agency", "")).strip().upper() in ("Y", "YES", "TRUE", "1", "예")
        origin = str(row.get("origin", "") or "").strip()
        fields_raw = str(row.get("disclosure_fields", "") or "")
        provided = {p.strip() for p in re.split(r"[;|]", fields_raw) if p.strip()}

        rule = self.match_product_rule(row)
        group = self.detect_group(category, product_name, rule.get("group", ""))
        if rule.get("id") == "GENERIC" and group == "기타":
            group = self.detect_group(category, product_name)

        primary = (rule.get("tracks") or [{}])[0]
        track = primary.get("track", "해당 없음")
        res = SkuResult(
            sku_id=str(row.get("sku_id", "")).strip(), product_name=product_name, category=category,
            group=group, model_name=model_name, price=str(row.get("price", "")).strip(),
            detail_text=detail, origin=origin, disclosure_fields=fields_raw,
            rule_id=rule.get("id", ""), rule_label=rule.get("label", ""),
            track=track, track_basis=primary.get("basis", ""),
            track_confidence=primary.get("confidence", "낮음"),
            track_alternatives=[t.get("track") for t in rule.get("tracks", [])[1:]],
            cert_no=cert_no, is_purchase_agency=is_pa,
        )
        findings: list[Finding] = []

        # ② 인증번호 조회 (전파 트랙은 조회 채널이 다름)
        is_rf = "전파" in track
        is_kc = track.startswith(("안전인증", "안전확인", "공급자적합성", "어린이", "전파"))
        if is_kc and not is_rf and not cert_no:
            pass
        if cert_no and is_kc and not is_rf:
            if self.offline or not self.client:
                d = type("D", (), {"status": "unavailable", "message": "오프라인 모드(조회 생략)",
                                   "cert_state": "", "cert_scope": "", "agency": "", "cert_date": "",
                                   "item_name": "", "model_name": "", "maker": "", "maker_country": "",
                                   "recall_hits": [], "recall_in_cert": ""})()
            else:
                d = self.client.lookup(cert_no, model_name)
            res.cert_status, res.cert_state = d.status, d.cert_state or ""
            res.cert_scope, res.cert_agency = d.cert_scope or "", d.agency or ""
            res.cert_date = d.cert_date or ""
            res.cert_item, res.cert_model = d.item_name or "", d.model_name or ""
            res.cert_maker, res.cert_maker_country = d.maker or "", d.maker_country or ""
            hits = d.recall_hits or []
            res.recall = f"{len(hits)}건" if hits else ("이력 없음" if d.status == "verified" else "미확인")

            # ① 인증 트랙 대조
            if d.status == "verified":
                state = (d.cert_state or "").strip()
                if state and state != "적합":
                    p = self._penalty("missing_safety_cert")
                    findings.append(Finding(
                        category="KC인증", severity="치명",
                        title=f"인증 상태가 '{state}' — 해당 인증으로 판매할 수 없습니다",
                        detail=f"{cert_no}는 제품안전정보센터에서 인증상태 '{state}'로 확인됩니다. 인증이 반납·취소된 상태로 보이며, 이 번호를 표시한 판매는 미인증 판매와 같은 위험이 있습니다.",
                        evidence=f"인증구분: {d.cert_scope} / 인증일자: {d.cert_date} / 인증기관: {d.agency}",
                        law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                        fix="판매 중지 후 올바른 인증번호(동일 모델)를 확인해 교체하고, 없으면 신규 인증을 진행하세요.",
                        action=f"제품안전정보센터에서 동일 모델 인증 재검색 → 인증기관 문의({self._fee(track)})",
                        confidence="높음", rule_id="CERT-STATE"))
                scope = d.cert_scope or ""
                if "어린이제품" in scope and group != "어린이제품":
                    findings.append(Finding(
                        category="KC인증", severity="보통",
                        title="인증번호의 인증구분이 상품군과 다릅니다",
                        detail=f"조회된 인증은 '{scope}'입니다. 판매 상품군({group})과 인증구분이 일치하는지 확인이 필요합니다.",
                        evidence=f"품목명: {d.item_name}", law="전기용품 및 생활용품 안전관리법 제9조제4항",
                        penalty_text="500만원 이하 과태료(인증정보 오표시 시)", penalty_max=5000000,
                        fix="인증서와 상품의 품목·모델을 대조해 정확한 인증정보를 표기하세요.",
                        action="인증서 사본 확인", confidence="중간", rule_id="CERT-SCOPE-GROUP"))
                if model_name and d.model_name:
                    if self._norm_model(model_name) != self._norm_model(d.model_name) and \
                       self._norm_model(model_name) not in self._norm_model(d.model_name):
                        findings.append(Finding(
                            category="KC인증", severity="중대",
                            title="등록 모델명과 인증 모델명이 다릅니다",
                            detail=f"입력 모델명 '{model_name}' / 인증된 모델명 '{d.model_name}'. 인증은 모델 단위로 효력이 있어, 다른 모델에 인증번호를 쓰면 인증정보 오표시가 됩니다.",
                            evidence=f"인증번호 {cert_no} · 품목 {d.item_name}",
                            law="전기용품 및 생활용품 안전관리법 제9조제2항·제4항",
                            penalty_text="500만원 이하 과태료 또는 인증표시 위반(3년 이하 징역·3천만원 이하 벌금)", penalty_max=5000000,
                            fix="인증받은 모델명을 정확히 표기하고, 파생모델이면 인증서의 파생모델란에 포함되어 있는지 확인하세요.",
                            action="인증서 파생모델 확인 또는 변경인증 신청", confidence="높음", rule_id="CERT-MODEL"))
                if hits or (d.recall_in_cert and d.recall_in_cert not in ("", "-")):
                    findings.append(Finding(
                        category="KC인증", severity="치명",
                        title="리콜 이력이 확인됩니다",
                        detail=f"리콜정보 검색에서 {len(hits)}건이 확인되었습니다. 리콜 제품은 판매·보관이 제한되고, 판매자는 회수·교환 조치 의무가 있습니다.",
                        evidence="; ".join(f"{h.get('product')}/{h.get('model')}({h.get('recall_type')}, {h.get('public_date')})" for h in hits[:3]) or f"인증정보 리콜현황: {d.recall_in_cert}",
                        law="전기용품 및 생활용품 안전관리법 제40조(수거·파기 등 명령), 어린이제품법 제10조",
                        penalty_text="회수·교환·환불 명령 불이행 시 3년 이하 징역 또는 3천만원 이하 벌금", penalty_max=30000000,
                        fix="즉시 판매 중지 및 구매자 회수·환불 공지.",
                        action="제품안전정보센터 리콜 상세 확인 → 판매중지 및 회수 절차", confidence="높음", rule_id="CERT-RECALL"))
                if state == "적합" and not hits:
                    res.score = res.score  # 인증 자체는 정상
            elif d.status == "not_found":
                p = self._penalty(primary.get("penalty_key", "missing_safety_cert"))
                findings.append(Finding(
                    category="KC인증", severity="중대",
                    title="입력한 인증번호로 인증정보를 찾을 수 없습니다",
                    detail=f"제품안전정보센터 인증정보 검색에서 '{cert_no}' 결과가 없습니다. 오타·구형 번호·타사 번호·임의 기재 중 하나일 수 있으며, 그대로 표시하면 허위 인증정보 표시가 됩니다.",
                    evidence=f"조회 {datetime.now():%Y-%m-%d %H:%M} · safetykorea.kr 인증정보 검색",
                    law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                    fix="인증서 원본의 인증번호·모델명을 그대로 옮겨 다시 확인하고, 확인될 때까지 인증 관련 문구를 내리세요.",
                    action="인증서 재확인 → 필요 시 인증기관 조회 의뢰", confidence="높음", rule_id="CERT-NOT-FOUND"))
            elif d.status == "format_error":
                findings.append(Finding(
                    category="KC인증", severity="중대",
                    title="인증번호 형식이 KC 인증번호 규격과 다릅니다",
                    detail=f"입력값 '{cert_no}' — KC 인증번호는 보통 'SU071827-26003', 'CB015R1432-26029'처럼 인증구분 접두어+기관코드+연도+일련번호 구조입니다.",
                    evidence="입력값 형식 검사", law="전기용품 및 생활용품 안전관리법 제9조제4항",
                    penalty_text="500만원 이하 과태료(인증정보 오표시)", penalty_max=5000000,
                    fix="인증서에 인쇄된 번호를 그대로 입력하세요(공백·전각문자 제거).",
                    action="인증서 원본 대조", confidence="높음", rule_id="CERT-FORMAT"))
            elif d.status == "unavailable":
                findings.append(Finding(
                    category="KC인증", severity="정보",
                    title="인증정보 실시간 조회 불가 — 수동 확인 필요",
                    detail=f"{cert_no}: {d.message} 리포트에는 '확인 불가'로 표시했으며, 유효하다고 가정하지 않았습니다.",
                    evidence="safetykorea.kr 응답 없음/차단",
                    fix="네트워크 복구 후 재조회하거나, 인증서 사본을 받아 대조하세요.",
                    action="인증서 사본 요청", confidence="낮음", rule_id="CERT-UNAVAILABLE"))
        elif cert_no and is_rf:
            findings.append(Finding(
                category="KC인증", severity="정보",
                title="전파 적합성평가는 제품안전정보센터가 아닌 RRA에서 확인",
                detail=f"입력 인증번호 '{cert_no}'. 방송통신기자재 적합성평가(적합인증·적합등록·자기적합확인)는 국립전파연구원 전자파적합등록 조회에서 확인합니다.",
                fix="RRA 적합성평가 조회에서 인증번호·모델명을 확인하고, 제품·포장에 적합성평가 표시(KC마크·관리번호·모델명)를 부착하세요.",
                action="RRA(rra.go.kr) 적합성평가 조회", confidence="높음", rule_id="RF-CHANNEL"))

        # ① 인증 대상인데 번호 없음 / 트랙 자체가 면제·확인 필요
        needs_cert = is_kc and primary.get("penalty_key", "").startswith(("missing_", "rf_")) and \
            "해당 없음" not in track and "확인 필요" not in track
        if needs_cert and not cert_no:
            p = self._penalty(primary.get("penalty_key", "missing_safety_cert"))
            if is_rf:
                findings.append(Finding(
                    category="KC인증", severity="중대",
                    title="무선기능 제품 — 전파 적합성평가 정보가 없습니다",
                    detail="무선(Wi-Fi·블루투스·무선충전) 기능이 있는 기자재는 적합성평가를 받고 표시해야 판매할 수 있습니다. 적합성평가 없이 판매하면 형사처벌 대상입니다.",
                    evidence=f"판정 근거: {primary.get('basis','')}",
                    law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                    fix="적합성평가(RRA) 취득 후 KC마크·관리번호·모델명을 제품과 상세페이지에 표시하세요.",
                    action="RRA 적합성평가 조회·신청 안내", confidence="중간", rule_id="RF-MISSING"))
            else:
                low = primary.get("confidence") != "높음"
                findings.append(Finding(
                    category="KC인증", severity="치명",
                    title=(f"KC 인증정보 없음 — {track} 대상으로 판정됩니다" if not low
                           else f"KC 인증정보 없음 — {track} 대상일 수 있습니다(확인 필요)"),
                    detail=f"'{product_name}'은(는) 1차 판정상 {track} 대상입니다. 인증번호가 없으면 미인증 제조·수입·판매로 처벌 대상이 될 수 있습니다.",
                    evidence=f"판정 근거: {primary.get('basis','')} (신뢰도 {primary.get('confidence','낮음')})",
                    law=p.get("article", ""),
                    penalty_text=("대상 품목일 경우 " if low else "") + p.get("text", ""),
                    penalty_max=p.get("max_fine"),
                    fix="동일 모델의 인증번호를 확인해 상세페이지·상품정보에 표시하고, 없으면 인증 취득 후 판매하세요. 그 전까지는 판매를 중지하는 것이 안전합니다.",
                    action=f"인증 취득 문의 — {self._fee(track)}", confidence=primary.get("confidence", "낮음"),
                    rule_id="CERT-MISSING"))

        if "확인 필요" in track or rule.get("id") == "GENERIC":
            findings.append(Finding(
                category="KC인증", severity="경미",
                title="KC 대상 품목 여부 확인이 필요합니다",
                detail=f"입력 정보로는 {rule.get('label')}으로 분류되어 KC 대상 여부를 단정하지 못했습니다. 대상이 아니면 조치 없음, 대상이면 인증정보 표시가 필요합니다.",
                evidence=primary.get("basis", ""), law="전기용품 및 생활용품 안전관리법 시행규칙 별표 3~6",
                penalty_text="대상 품목일 경우 미인증 판매 시 3년 이하 징역 또는 3천만원 이하 벌금",
                penalty_max=30000000,
                fix="품목명·재질·전압·대상연령을 기준으로 인증기관에 대상 여부를 1회 확인하세요.",
                action="인증기관 대상 여부 문의", confidence="낮음", rule_id="KC-TARGET-UNSURE"))

        if not is_kc and "확인 필요" not in track:
            findings.append(Finding(
                category="KC인증", severity="정보",
                title="KC 인증 대상 품목이 아닙니다 — 개별법 기준으로 점검",
                detail=f"{rule.get('label')}은(는) 전기·생활용품/어린이제품 KC 인증 대상이 아니라 {primary.get('law','개별법')} 기준으로 표시·광고·표시사항을 점검합니다.",
                evidence=primary.get("basis", ""),
                law=primary.get("law", ""),
                fix="KC 인증번호 대신 개별법상 표시사항(화장품: 전성분·사용기한, 식품: 원재료·소비기한)을 확인하세요.",
                action="개별법 표시사항 점검", confidence="중간", rule_id="NOT-KC"))


        # ③ 표시광고 표현 탐지
        scope_keys = ["general"]
        if group == "식품":
            scope_keys.append("food")
        if group == "화장품":
            scope_keys.append("cosmetics")
        if group == "어린이제품":
            scope_keys.append("children")
        scan = f"{product_name} {detail}"
        for rule_ad in self.rs["ad_rules"]:
            if not set(rule_ad["scopes"]) & set(scope_keys):
                continue
            hits = []
            for pat in rule_ad["patterns"]:
                if self._ad_hit(pat, scan):
                    hits.append(pat)
            if not hits:
                continue
            p = self._penalty(rule_ad["penalty_key"])
            terms = ", ".join(f"'{h}'" for h in hits[:4])
            findings.append(Finding(
                category="표시광고", severity=rule_ad["severity"],
                title=f"{rule_ad['label']} — {terms}",
                detail=f"상세페이지·상품명에서 {len(hits)}건의 위험 표현이 탐지되었습니다. 소비자 오인 우려가 있어 표시광고법·개별법 위반 소지가 있습니다.",
                evidence=f"탐지 표현: {terms}",
                law=rule_ad["law"], penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                fix=rule_ad["replacement"].replace("{원문}", hits[0]),
                action=rule_ad["action"], confidence="중간", rule_id=rule_ad["id"]))

        # 화장품 기능성 표방
        if group == "화장품":
            func_terms = [t for t in ["미백", "주름", "자외선", "탄력", "개선", "기능성"] if t in scan]
            has_functional_label = "기능성화장품" in fields_raw or "기능성 화장품" in scan
            if func_terms and not has_functional_label:
                p = self._penalty("cosmetic_functional_unapproved")
                findings.append(Finding(
                    category="표시광고", severity="치명",
                    title="기능성 표현 — 기능성화장품 심사·표시 확인 필요",
                    detail=f"'{', '.join(func_terms)}' 등 기능성 표방 표현이 있으나 기능성화장품 심사(보고) 및 '기능성화장품' 표시가 확인되지 않습니다. 미심사 기능성화장품 판매는 형사처벌 대상입니다.",
                    evidence=f"탐지 표현: {', '.join(func_terms)}",
                    law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                    fix="1) 기능성화장품 심사·보고 후 '기능성화장품' 문구와 심사 결과 범위의 효능만 표시하거나, 2) 기능성 표현을 삭제하고 일반 화장품으로 표시하세요.",
                    action="식약처 기능성화장품 심사·보고 절차 확인(책임판매업 등록 필요)", confidence="중간", rule_id="COSM-FUNC"))

        # ④ 고시정보 누락
        required = list(self.rs["disclosure_sets"].get(group, self.rs["disclosure_sets"]["기타"]))
        required += [dict(c) for c in self.rs["disclosure_common"]]
        missing = [r["label"] for r in required if r["key"] not in provided]
        if missing:
            p = self._penalty("disclosure_missing")
            sev = "보통" if len(missing) >= 3 else "경미"
            findings.append(Finding(
                category="고시정보", severity=sev,
                title=f"상품정보제공고시 누락 {len(missing)}건",
                detail="계약 전 표시해야 하는 품목별 정보·거래조건 정보가 확인되지 않았습니다. 표·항목을 채우지 않으면 과태료 대상입니다.",
                evidence="누락: " + ", ".join(missing[:10]) + (" 외" if len(missing) > 10 else ""),
                law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                fix="상세페이지 하단 '상품정보제공고시' 표에 누락 항목을 채우세요. 모르는 항목은 비우지 말고 사유와 그에 준하는 정보를 적습니다.",
                action="상품정보제공고시 표 보완(플랫폼 상품정보 입력란 포함)", confidence="높음", rule_id="DISCLOSURE"))
        if group in ("전기용품", "어린이제품", "전파") and cert_no and "KC인증번호" not in provided:
            p = self._penalty("no_online_cert_info")
            findings.append(Finding(
                category="고시정보", severity="보통",
                title="온라인 인증정보 게시 누락",
                detail="인증을 받은 제품은 온라인 판매 시 인증번호 등 인증정보를 소비자가 볼 수 있게 게시해야 합니다.",
                evidence=f"보유 인증번호: {cert_no} / 상품정보 제공 항목에 미포함",
                law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                fix="상품정보 입력란의 '인증정보'에 인증유형·인증번호를 문자로 표기(이미지 대신 텍스트)하세요.",
                action="플랫폼 인증정보 입력란 보완", confidence="높음", rule_id="CERT-DISPLAY"))

        # ⑤ 구매대행
        if is_pa:
            pa = self.rs["purchase_agency_notice"]
            if group == "어린이제품":
                p = self._penalty("children_purchase_agency")
                findings.append(Finding(
                    category="구매대행", severity="치명",
                    title="어린이제품은 구매대행·중개 판매가 금지됩니다",
                    detail="어린이제품은 구매대행 특례 대상이 아니며, 안전인증·안전확인·공급자적합성확인 표시가 없는 어린이제품의 판매 중개·구매대행은 금지됩니다.",
                    evidence=f"입력값 is_purchase_agency=Y / 상품군 {group}",
                    law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                    fix="구매대행 판매를 중단하고, 국내 인증을 보유한 동일 제품으로 전환하거나 직접 수입·인증하세요.",
                    action="판매 방식 변경(직수입+인증) 검토", confidence="높음", rule_id="PA-CHILD"))
            elif group == "전파":
                findings.append(Finding(
                    category="구매대행", severity="중대",
                    title="구매대행 무선기기 — 적합성평가 확인 필요",
                    detail="전파법상 적합성평가를 받지 않은 기자재는 판매·수입이 금지되며, 구매대행에도 예외가 없습니다.",
                    law="전파법 제58조의2·제84조", penalty_text="3년 이하 징역 또는 3천만원 이하 벌금", penalty_max=30000000,
                    fix="적합성평가(적합등록 등) 확인 후 판매하세요.",
                    action="RRA 조회 후 판매 여부 결정", confidence="중간", rule_id="PA-RF"))
            missing_notice = [i for i in pa["required_items"]
                              if not any(k in detail for k in [i[:4], "구매대행", "개인통관"])]
            if len(missing_notice) >= 2:
                p = self._penalty("no_purchase_agency_notice")
                findings.append(Finding(
                    category="구매대행", severity="중대",
                    title="구매대행 고지사항 누락",
                    detail="구매대행 판매 시 (1) 구매대행 상품임 (2) 안전관리대상 제품 해당 여부 (3) 안전인증·안전확인 표시에 관한 사항을 인터넷 홈페이지에 게시해야 합니다.",
                    evidence="고지 문구 미확인(상세페이지 텍스트 기준)",
                    law=p.get("article", ""), penalty_text=p.get("text", ""), penalty_max=p.get("max_fine"),
                    fix="상세페이지 상단에 '본 상품은 해외 구매대행 상품으로, KC 인증 면제 대상(구매대행)임을 고지합니다. 개인통관고유부호가 필요합니다.' 문구를 넣으세요.",
                    action="구매대행 고지 문구 삽입", confidence="중간", rule_id="PA-NOTICE"))

        # 원산지
        if not origin:
            p = self._penalty("origin_missing")
            findings.append(Finding(
                category="원산지", severity="경미" if not is_pa else "보통",
                title="원산지 정보가 없습니다",
                detail="원산지 표시는 수입·구매대행 상품에서 특히 문제가 됩니다. 상품정보·상세페이지에 원산지를 표시하세요.",
                evidence="입력값 origin 공란", law=p.get("article", ""), penalty_text=p.get("text", ""),
                fix=f"실제 원산지를 표기하세요(예: '원산지: {res.cert_maker_country or '○○'}').",
                action="원산지 표시 추가", confidence="중간", rule_id="ORIGIN"))
        elif origin and re.search(r"국산|국내산|한국", detail) and not re.search(r"국산|국내산|한국", origin):
            findings.append(Finding(
                category="원산지", severity="중대",
                title="상세페이지의 국산 표현이 입력 원산지와 다릅니다",
                detail=f"입력 원산지는 '{origin}'인데 상세페이지에 국산·국내산 표현이 있습니다.",
                evidence=f"원산지 입력: {origin}", law="대외무역법 제33조, 표시광고법 제3조제1항",
                penalty_text="시정명령·과태료 및 표시광고법상 거짓광고", fix="국산 표현 삭제 또는 실제 원산지로 정정.",
                action="원산지 표현 정정", confidence="높음", rule_id="ORIGIN-CLAIM"))

        # 점수·등급
        if not is_kc and not cert_no:
            res.cert_status = "not_applicable"
        res.findings = findings
        total = sum(self.weights.get(f.severity, 0) for f in findings)
        res.score = max(0, min(100, 100 - total))
        has_critical = any(f.severity == "치명" for f in findings)
        grade = next((g for g in self.rs["grades"] if res.score >= g["min_score"]), self.rs["grades"][-1])
        # 판매 자체가 막히는 유형(미인증·인증효력 없음·리콜·어린이 구매대행·미심사 기능성·중대 효능표현)은 긴급
        if any(f.rule_id in SALE_BLOCKING for f in findings):
            grade = next(g for g in self.rs["grades"] if g["grade"] == "E")
        elif has_critical and grade["grade"] in ("A", "B", "C"):
            grade = next(g for g in self.rs["grades"] if g["grade"] == ("E" if res.score < 30 else "D"))
        res.grade, res.grade_label = grade["grade"], grade["label"]
        res.summary = self._summary(res)
        res.next_action = self._next_action(res)
        res.fee_note = self._fee(track) if needs_cert else "-"
        return res

    @staticmethod
    def _ad_hit(pat: str, text: str) -> bool:
        """표현 탐지. '면 100%' 같은 소재 함량 표기는 과장으로 보지 않는다."""
        t, p = text.lower(), pat.lower()
        if p != "100%":
            return p in t
        material = r"(면|코튼|울|캐시미어|모달|린넨|실리콘|스테인리스|천연|오가닉|폴리에스터|재생|국내산|한국산)\s*$"
        for m in re.finditer(re.escape(p), t):
            if re.search(material, t[max(0, m.start() - 8):m.start()]):
                continue
            return True
        return False

    @staticmethod
    def _norm_model(s: str) -> str:
        return re.sub(r"[^0-9a-z]", "", str(s).lower())

    def _summary(self, res: SkuResult) -> str:
        if not res.findings:
            return "탐지된 위험 항목 없음"
        crit = [f for f in res.findings if f.severity == "치명"]
        if crit:
            return crit[0].title
        sev_order = {"중대": 3, "보통": 2, "경미": 1, "정보": 0}
        top = sorted(res.findings, key=lambda f: -sev_order.get(f.severity, 0))[0]
        return top.title

    def _next_action(self, res: SkuResult) -> str:
        for sev in ("치명", "중대", "보통", "경미"):
            for f in res.findings:
                if f.severity == sev and f.action:
                    return f.action
        return "정기 모니터링 유지"

    # ------------------------------------------------------------------ batch
    def run(self, rows: list[dict]) -> dict:
        results = [self.evaluate(r) for r in rows]
        counts: dict = {}
        for r in results:
            counts[r.grade] = counts.get(r.grade, 0) + 1
        sev_counts: dict = {}
        for r in results:
            for f in r.findings:
                sev_counts[f.severity] = sev_counts.get(f.severity, 0) + 1
        return {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "ruleset": {
                "id": self.rs["meta"]["id"], "version": self.rs["meta"]["version"],
                "law_basis_date": self.rs["meta"]["law_basis_date"],
                "disclaimer": self.rs["meta"]["disclaimer"],
                "unverified": self.rs["meta"]["unverified"],
                "sources": self.rs["meta"]["sources"],
                "laws": self.rs["meta"]["laws"],
            },
            "counts": {"total": len(results), "grades": counts, "severities": sev_counts},
            "results": [r.to_dict() for r in results],
        }


# ---------------------------------------------------------------------- IO
def load_rows(path: Path) -> list[dict]:
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header = [str(h).strip() if h is not None else "" for h in rows[0]]
        out = []
        for r in rows[1:]:
            if all(v is None or str(v).strip() == "" for v in r):
                continue
            out.append({header[i]: ("" if r[i] is None else str(r[i]).strip()) for i in range(len(header))})
        return out
    with path.open(encoding="utf-8-sig", newline="") as f:
        return [{k.strip(): ("" if v is None else str(v).strip()) for k, v in row.items()} for row in csv.DictReader(f)]


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="셀러팩트 감사 엔진")
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", default="output")
    ap.add_argument("--client", default="샘플 셀러")
    ap.add_argument("--brand-name", default="셀러팩트")
    ap.add_argument("--brand-en", default="SellerFact")
    ap.add_argument("--no-network", action="store_true")
    ap.add_argument("--no-pdf", action="store_true")
    ap.add_argument("--no-xlsx", action="store_true")
    args = ap.parse_args(argv)

    rows = load_rows(Path(args.input))
    engine = AuditEngine(offline=args.no_network)
    data = engine.run(rows)
    data["client"] = args.client
    data["input_file"] = str(args.input)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d")
    slug = re.sub(r"[^0-9A-Za-z가-힣]+", "", args.client) or "audit"
    json_path = out_dir / f"{slug}_audit_{stamp}.json"
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
    print(f"[engine] JSON  {json_path}")

    paths = {"json": str(json_path)}
    if not args.no_xlsx:
        import report_xlsx
        p = report_xlsx.build(data, out_dir / f"{slug}_감사리포트_{stamp}.xlsx")
        paths["xlsx"] = str(p)
        print(f"[engine] XLSX  {p}")
    if not args.no_pdf:
        import report_pdf
        p = report_pdf.build(data, out_dir / f"{slug}_감사리포트_{stamp}.pdf",
                             brand_name=args.brand_name, brand_en=args.brand_en)
        paths["pdf"] = str(p)
        print(f"[engine] PDF   {p}")
    print(json.dumps({"grades": data["counts"]["grades"], "files": paths}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
