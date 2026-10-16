# 손해율 월말 집계 및 워드 리포트 자동화

```bash
pip install -r requirements.txt
python 1_generate_dummy.py          # (Phase 1) 더미 마스터 545건 + 로데이터 생성
python main_pipeline.py             # 기본: 마감 25.09, UY 계약 25.06~25.09
python main_pipeline.py --base 25.08 --uy-start 25.06
```

산출물 (`output/`)
- `monthly_loss_report.docx` — 워드 리포트 (전사 요약 → UY 누적 / CY 당월별 심사유형·대급부·Top 20)
- `by_code_UY_<기준월>.csv`, `by_code_CY_<기준월>.csv` — 담보별 집계(검증용)

## 계산 정의
| 항목 | 정의 |
|---|---|
| 손해액 | 지급보험금합계 + 지급비금변동분 합계 (`LOSS_COMPONENTS`) |
| 손해율 | 손해액 ÷ 위험보험료 × 100 |
| UY 누적 | 계약년월 ∈ [UY 시작, 기준월] 이고 마감년월 ≤ 기준월인 실적 누적 |
| CY 당월 | 마감년월 = 기준월인 실적 (계약년월 무관) |
| 전년 동기 | 위 조건을 12개월 앞당긴 구간 (예: 24.06~24.09 계약 / 24.09 마감) |
| 위험보험료 증감률 | (당해 ÷ 전년 − 1) × 100 |
| 손해율 증감(p) | 당해 손해율 − 전년 손해율 |
| 손해율 Top 20 | 위험보험료 비중 0.05% 이상 담보 중 손해율 상위 (`LR_RANK_MIN_PREMIUM_SHARE`, 0이면 제한 없음) |

마스터에 없는 담보코드는 `(미매핑)`으로 표시되고 실행 시 경고가 출력됩니다.

## Phase 2 (내부망) 전환
1. `1_generate_dummy.py`는 사용하지 않음 (삭제 가능)
2. 경로: `main_pipeline.py` 상단 CONFIG 수정 또는 환경변수
   `LR_DATA_DIR`, `LR_OUTPUT_DIR`, `LR_RAW_FILE`, `LR_MASTER_FILE`
3. 실제 컬럼명이 다르면 `COL` 매핑만 수정
4. 로데이터에는 **전년 동기 구간**(예: 24.06~24.09)도 포함되어야 증감이 계산됨
5. CSV 인코딩은 UTF-8 / CP949 자동 판별, 년월은 `25.06`, `2025.06`, `202506`, `2025-06` 모두 허용
