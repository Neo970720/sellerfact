"""
KC 인증정보 실측 조회 클라이언트 (제품안전정보센터 safetykorea.kr 공개 조회).

공식 공공 API(data.go.kr 15116894)는 활용신청 키가 필요하므로, 키 없이 동작하는
공개 조회 경로를 사용한다.

  - 인증 목록/상세: POST https://www.safetykorea.kr/release/certificationsearch
                    POST https://www.safetykorea.kr/release/certDetail
  - 리콜 목록:      POST https://www.safetykorea.kr/recall/recallBoard
                    (selectedOption=5 인증번호 / 1 모델명 / 2 제품명)

키가 있으면 공식 OpenAPI를 우선 사용하도록 확장할 수 있게 provider 구조를 나눠 두었다.
네트워크가 막히거나 응답이 없으면 예외를 던지지 않고 status='unavailable'로 돌려주며,
리포트에는 "확인 불가"로 표기된다(추정으로 채우지 않는다).
"""

from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field, asdict
from datetime import datetime
from email.utils import formatdate
from pathlib import Path
from typing import Optional

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 CleanSKU/1.0"
BASE = "https://www.safetykorea.kr"
DEFAULT_TIMEOUT = float(os.environ.get("KC_LOOKUP_TIMEOUT", "12"))
CACHE_DIR = Path(os.environ.get("KC_CACHE_DIR", Path(__file__).resolve().parent.parent / "_cache"))
CACHE_TTL = int(os.environ.get("KC_CACHE_TTL", str(60 * 60 * 24 * 3)))  # 3일


@dataclass
class CertDetail:
    cert_no: str
    status: str  # verified | not_found | unavailable | skipped
    message: str = ""
    agency: Optional[str] = None          # 인증기관
    cert_scope: Optional[str] = None      # 인증구분 (예: ...>안전인증대상 전기용품)
    cert_state: Optional[str] = None      # 인증상태 (적합/취소 등)
    cert_date: Optional[str] = None       # 인증일자 YYYYMMDD
    changed_date: Optional[str] = None
    recall_in_cert: Optional[str] = None  # 인증정보 내 리콜현황
    item_name: Optional[str] = None       # 품목명
    model_name: Optional[str] = None
    category_code: Optional[str] = None
    maker: Optional[str] = None
    maker_country: Optional[str] = None
    importer: Optional[str] = None
    recall_hits: list = field(default_factory=list)
    source_url: str = f"{BASE}/release/certificationsearch"
    checked_at: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


def _http_post(path: str, data: dict, timeout: float = DEFAULT_TIMEOUT) -> str:
    body = urllib.parse.urlencode(data, encoding="utf-8").encode("utf-8")
    req = urllib.request.Request(
        BASE + path,
        data=body,
        method="POST",
        headers={
            "User-Agent": UA,
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Accept-Language": "ko-KR,ko;q=0.9",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:  # noqa: S310 (공개 기관 사이트 고정 URL)
        raw = resp.read()
    return raw.decode("utf-8", "replace")


def _text_lines(html: str) -> list[str]:
    """태그를 제거한 뒤 의미 있는 텍스트 줄만 남긴다."""
    h = re.sub(r"<script.*?</script>", " ", html, flags=re.S | re.I)
    h = re.sub(r"<style.*?</style>", " ", h, flags=re.S | re.I)
    h = re.sub(r"<br\s*/?>", "\n", h, flags=re.I)
    h = re.sub(r"</(tr|td|th|div|p|li|table|h\d)>", "\n", h, flags=re.I)
    h = re.sub(r"<[^>]+>", " ", h)
    h = (h.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
          .replace("&quot;", '"').replace("&#39;", "'").replace("&nbsp;", " "))
    return [x.strip() for x in h.split("\n") if x.strip()]


def _norm(s: str) -> str:
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
         .replace("&quot;", '"').replace("&nbsp;", " "))
    return re.sub(r"\s+", " ", s).strip()


def _th_td_pairs(html: str) -> dict:
    """상세 페이지의 <th>라벨</th><td>값</td> 쌍을 사전으로 만든다."""
    body = html[html.find('class="contents_area"'):] if 'class="contents_area"' in html else html
    pairs: dict = {}
    for m in re.finditer(r"<th[^>]*>(.*?)</th>\s*<td[^>]*>(.*?)</td>", body, flags=re.S | re.I):
        label = _norm(m.group(1))
        value = _norm(m.group(2))
        if label and label not in pairs:
            pairs[label] = value if value not in ("-", "-->") else ""
    return pairs


def _field(lines: list[str], label: str) -> Optional[str]:
    for i, ln in enumerate(lines):
        if ln == label or ln.startswith(label + " "):
            # 라벨 뒤에 값이 붙어 있는 경우
            rest = ln[len(label):].strip(" :：>")
            if rest:
                return rest
            for nxt in lines[i + 1:i + 4]:
                if nxt not in ("-", "-->"):
                    return nxt
    return None


class SafetyKoreaClient:
    """KC 인증번호 유효성·리콜 여부 실측 조회기."""

    def __init__(self, cache_dir: Optional[Path] = None, timeout: float = DEFAULT_TIMEOUT):
        self.timeout = timeout
        self.cache_dir = Path(cache_dir) if cache_dir else CACHE_DIR
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.calls = 0

    # ---------- cache ----------
    def _cache_path(self, key: str) -> Path:
        safe = re.sub(r"[^0-9A-Za-z가-힣_-]", "_", key)[:120]
        return self.cache_dir / f"{safe}.json"

    def _cache_get(self, key: str) -> Optional[dict]:
        p = self._cache_path(key)
        if not p.exists():
            return None
        try:
            blob = json.loads(p.read_text("utf-8"))
            if time.time() - blob.get("_ts", 0) > CACHE_TTL:
                return None
            return blob.get("data")
        except Exception:
            return None

    def _cache_put(self, key: str, data: dict) -> None:
        try:
            self._cache_path(key).write_text(
                json.dumps({"_ts": time.time(), "data": data}, ensure_ascii=False), "utf-8"
            )
        except Exception:
            pass

    # ---------- public ----------
    def lookup(self, cert_no: str, model_name: str = "", use_recall: bool = True) -> CertDetail:
        """인증번호 1건 조회. 실패해도 예외를 던지지 않는다."""
        checked = datetime.now().strftime("%Y-%m-%d %H:%M")
        cert_no = (cert_no or "").strip().replace(" ", "")
        if not cert_no:
            return CertDetail(cert_no="", status="skipped",
                              message="인증번호 미기재", checked_at=checked)
        if not re.fullmatch(r"[A-Z]{2}[0-9A-Z]{3,9}-[0-9]{3,6}[A-Z0-9]{0,3}", cert_no.upper()):
            return CertDetail(cert_no=cert_no, status="format_error",
                              message="인증번호 형식이 KC 인증번호 규격(예: SU071827-26003)과 다릅니다.",
                              checked_at=checked)

        key = f"cert_{cert_no.upper()}"
        cached = self._cache_get(key)
        if cached:
            detail = CertDetail(**cached)
            detail.checked_at = checked
            detail.message += " (캐시)"
            return detail

        try:
            uid = self._find_uid(cert_no)
            if uid is None:
                d = CertDetail(cert_no=cert_no, status="not_found",
                               message="제품안전정보센터 인증정보 검색 결과가 없습니다.",
                               checked_at=checked)
            else:
                d = self._detail(cert_no, uid, checked)
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
            d = CertDetail(cert_no=cert_no, status="unavailable",
                           message=f"조회 서버 응답 없음/차단 ({type(e).__name__}). 수동 확인 필요.",
                           checked_at=checked)
        except Exception as e:  # 파싱 실패도 조용히 표기
            d = CertDetail(cert_no=cert_no, status="unavailable",
                           message=f"응답 해석 실패({type(e).__name__}). 수동 확인 필요.",
                           checked_at=checked)

        if use_recall and d.status == "verified":
            try:
                d.recall_hits = self.recall_search(cert_no, model_name, by="cert")
            except Exception:
                d.recall_hits = []

        if d.status in ("verified", "not_found"):
            self._cache_put(key, d.to_dict())
        return d

    def recall_search(self, word: str, model_name: str = "", by: str = "cert") -> list:
        """리콜정보 검색. by: cert(인증번호) | model(모델명) | product(제품명)."""
        opt = {"cert": "5", "model": "1", "product": "2"}.get(by, "5")
        hits: list = []
        for term in [t for t in {word, model_name} if t]:
            html = _http_post("/recall/recallBoard",
                              {"selectedOption": opt, "textWord": term.replace(" ", ""), "pageNo": "0"},
                              timeout=self.timeout)
            hits.extend(self._parse_recall(html))
        # 중복 제거
        seen, out = set(), []
        for h in hits:
            k = (h.get("product"), h.get("model"), h.get("public_date"))
            if k not in seen:
                seen.add(k)
                out.append(h)
        return out

    # ---------- internals ----------
    def _find_uid(self, cert_no: str) -> Optional[str]:
        html = _http_post("/release/certificationsearch", {
            "certNum": cert_no, "categoryCode1": "", "categoryCode2": "", "categoryCode3": "",
            "productName": "", "modelName": "", "makerName": "", "importerName": "", "pageNo": "0",
        }, timeout=self.timeout)
        self.calls += 1
        m = re.search(r"goDetail\('([^']*)'\s*,\s*'(\d+)'\)", html)
        if m and m.group(1).strip().upper() == cert_no.upper():
            return m.group(2)
        # 대소문자 무시하고 첫 행 사용
        if m:
            return m.group(2)
        return None

    def _detail(self, cert_no: str, uid: str, checked: str) -> CertDetail:
        html = _http_post("/release/certDetail", {"certNum": cert_no, "certUid": uid}, timeout=self.timeout)
        self.calls += 1
        pairs = _th_td_pairs(html)
        lines = _text_lines(html)

        def g(*names: str) -> Optional[str]:
            for n in names:
                for k, v in pairs.items():
                    if k == n or k.startswith(n):
                        if v:
                            return v
            return None

        d = CertDetail(cert_no=cert_no, status="verified", message="인증정보 확인", checked_at=checked)
        d.agency = g("인증기관")
        d.cert_scope = g("인증구분")
        d.cert_state = g("인증상태")
        d.cert_date = g("인증일자")
        d.changed_date = g("인증변경일자")
        d.recall_in_cert = g("리콜현황") or ""
        d.item_name = g("품목명")
        d.model_name = g("모델명")
        d.category_code = g("제품분류코드")
        d.maker = g("제조사")
        d.maker_country = g("제조국")
        d.importer = g("수입업체")
        if d.cert_state is None and d.cert_scope is None and d.item_name is None:
            # 표 구조가 바뀌었을 때를 위한 보조 경로
            d.agency = d.agency or _field(lines, "인증기관")
            d.cert_state = d.cert_state or _field(lines, "인증상태")
            d.item_name = d.item_name or _field(lines, "품목명")
        if not (d.cert_state or d.cert_scope or d.item_name):
            d.status = "unavailable"
            d.message = "상세 응답에서 인증 필드를 찾지 못했습니다. 수동 확인 필요."
        return d

    @staticmethod
    def _parse_recall(html: str) -> list:
        rows = []
        for m in re.finditer(r"<tr[^>]*onclick=\"goDetail\('(\d+)'\)\"(.*?)</tr>", html, flags=re.S):
            cells = re.findall(r"<td[^>]*>(.*?)</td>", m.group(2), flags=re.S)
            vals = []
            for c in cells:
                t = re.sub(r"<[^>]+>", " ", c)
                t = re.sub(r"\s+", " ", t).strip()
                vals.append(t)
            if len(vals) >= 6:
                rows.append({
                    "recall_uid": m.group(1),
                    "no": vals[0],
                    "product": vals[2],
                    "model": vals[3],
                    "company": vals[4],
                    "recall_type": vals[5],
                    "public_date": vals[-1],
                })
        return rows


if __name__ == "__main__":
    import sys
    c = SafetyKoreaClient()
    for arg in sys.argv[1:] or ["SU071827-26003"]:
        d = c.lookup(arg)
        print(json.dumps(d.to_dict(), ensure_ascii=False, indent=2)[:1500])
