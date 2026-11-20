# 셀러팩트 (SellerFact) — 무인 자동 점검 파이프라인

상품 목록 CSV 하나로 **KC 인증·표시광고·상품정보 고시**를 점검해
**엑셀 6시트 + PDF 11쪽 리포트**를 자동으로 만들어 주는 파이프라인입니다.

- **AI 토큰 0**: 점검은 순수 파이썬(`src/`)이 실행합니다. 모델 호출이 없습니다.
- **초기 비용 0**: GitHub Actions(공개 저장소) + GitHub Pages + Google Apps Script 무료 한도만 사용합니다.
- **사람 개입 0**: 접수 → 점검 → 결과 게시가 전부 자동입니다.

## 구성

| 단계 | 담당 | 파일 |
|---|---|---|
| 접수·라이선스 확인 | Google Apps Script 웹앱(무료, 슬립 없음) | 저장소 밖 (아래 "접수 큐" 참고) |
| 점검 실행 | GitHub Actions (`ubuntu-latest`, 파이썬 3.12) | `.github/workflows/audit.yml` |
| 점검 엔진 | 순수 파이썬 — 룰셋 기반 판정 + 공공 인증정보 실측 조회 | `src/audit_engine.py`, `src/kc_client.py` |
| 리포트 생성 | xlsxwriter(6시트) + reportlab(11쪽, 나눔고딕 동봉) | `src/report_xlsx.py`, `src/report_pdf.py` |
| 결과 게시 | Pages (`docs/r/<job>/`) | `docs/view.html` |
| 고객 화면 | 정적 페이지(가입 불필요) | `docs/index.html` |

## 실행 흐름

```
고객 → docs/index.html (CSV + 라이선스 키)
     → Apps Script 웹앱: Gumroad 라이선스 검증 → 큐(스프레드시트)에 적재
     → GitHub Actions (10분마다 스케줄): scripts/poll_queue.py가 큐를 소비
     → scripts/run_job.py: 점검 + 리포트 생성 → docs/r/<job>/ 커밋
     → Pages가 즉시 서비스 → 고객 페이지가 meta.json을 폴링해 다운로드 링크 표시
```

## 수동 실행 (테스트·재처리)

```bash
python3 scripts/run_job.py --job test-001 --input jobs/demo-001.csv --client "테스트 셀러"
# 또는 GitHub에서
gh workflow run audit.yml -f job=demo-001 -f client="샘플 셀러"
```

## 접수 큐 (Apps Script)

`QUEUE_ENDPOINT` / `QUEUE_KEY` 두 개의 저장소 시크릿이 설정되어 있으면 10분마다 큐를 확인합니다.
설정이 없으면 스케줄 실행은 아무 일도 하지 않고 정상 종료합니다.

큐 프로토콜(JSON):

| action | 필수 파라미터 | 응답 |
|---|---|---|
| `queue` | `key` | `{ok, jobs:[{job, client, skus}]}` — 아직 결과가 없는 작업 |
| `payload` | `key`, `job` | `{ok, job, client, csv_b64}` |
| `ack` | `key`, `job`, `status`, `url?`, `message?` | `{ok}` |
| `status` | `job` | `{state: queued|done|error, message?, hint?}` (공개) |

## 입력·출력 규격

- 입력: `sku_id, product_name, category`(필수) + `model_name, price, detail_text, kc_cert_no,
  is_purchase_agency, origin, disclosure_fields, seller_note`
- 제한: 1회 300 SKU / 40KB. 인코딩은 UTF-8·CP949 자동 판별.
- 출력: `docs/r/<job>/report.xlsx`, `report.pdf`, `data.json`, `meta.json`

## 룰셋 갱신

`src/rules/ruleset.json`의 `meta.version`·`law_basis_date`를 올리고, 바뀐 조문·별표를 `sources`에 추가합니다.
품목은 `product_rules`, 표현은 `ad_rules`, 처분은 `penalties`에 반영하면 리포트가 자동 반영합니다.

## 한계

- 공개 법령·공시 기반 **1차 스크리닝**입니다. 최종 인증 트랙·적합성은 지정 시험·인증기관 확인이 필요합니다.
- 전파 적합성평가(RRA)는 제품안전정보센터 조회 대상이 아니므로 리포트에 조회 채널만 안내합니다.
- 인증 신청·대행, 시험 의뢰, 법률 자문은 하지 않습니다.
