"""
[외부망 전용] 더미 데이터 생성기 — 내부망 보고서 레이아웃 검증용

생성 파일
  data/code_master.csv : 담보 마스터 545건
      담보코드 | 담보명(=급부명, 보고서 행 단위) | 대급부명(표3 '급부' 구분) | 구분(표1·2 '구분', 빈칸이면 대급부명) | 고당대통(Y)
  data/raw_data.csv    : 로데이터 (계약 202506~, 마감 ~202608)
      원본 10개 컬럼 + 건수(11번째, 선택)

※ Phase 2(내부망)에서는 이 파일을 사용하지 않는다.
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
N_CODES = 545
DATA_DIR = Path(__file__).resolve().parent / "data"

LAUNCH = (2025, 6)       # 최초 판매월
LAST_CLOSE = (2026, 8)   # 마지막 마감월
UW_TYPES = ["일반", "간편"]

# 급부명: (대급부명=표3 구분, 구분=표1·2 구분, 고당대통, 보험료 가중치, 기대 손해율)
BENEFITS = {
    "암 외 기타질병 사망": ("사망", "사망", "", 3.0, 0.9),
    "일반사망": ("사망", "사망", "", 2.0, 0.5),
    "일반암 진단": ("암진단", "진단", "", 4.0, 1.1),
    "림프절전이암 진단": ("암진단", "진단", "", 0.6, 2.4),
    "제자리암 진단": ("암진단", "진단", "", 0.5, 2.2),
    "갑상선암 진단": ("암진단", "진단", "", 0.7, 1.3),
    "뇌혈관질환 진단": ("암이외진단", "진단", "", 3.0, 1.5),
    "뇌혈관질환 산정특례대상보장": ("암이외진단", "진단", "", 0.8, 1.4),
    "급여 뇌심장질환 혈관조영술 검사지원": ("암이외진단", "진단", "", 0.7, 1.3),
    "뇌졸중 진단": ("암이외진단", "진단", "", 1.2, 1.2),
    "부정맥 진단": ("암이외진단", "진단", "", 0.6, 1.25),
    "심장질환 산정특례대상보장": ("암이외진단", "진단", "", 0.7, 1.2),
    "허혈성심장질환 진단": ("암이외진단", "진단", "", 2.5, 2.0),
    "대상포진 진단": ("암이외진단", "진단", "Y", 3.5, 1.6),
    "대상포진눈병 진단": ("암이외진단", "진단", "Y", 0.5, 1.4),
    "통풍 진단": ("암이외진단", "진단", "Y", 2.2, 1.5),
    "당뇨병 진단": ("암이외진단", "진단", "Y", 2.0, 1.7),
    "고혈압(원발성) 약물치료": ("암이외진단", "진단", "Y", 1.5, 1.3),
    "질병장해_50%이상": ("장해", "장해", "", 0.8, 2.3),
    "질병장해_3~100%": ("장해", "장해", "", 2.4, 2.1),
    "일반재해장해_3~100%": ("장해", "장해", "", 2.0, 2.0),
    "관상동맥우회수술": ("수술", "수술", "", 0.5, 2.6),
    "1종 수술": ("수술", "수술", "", 1.0, 1.3),
    "2종 수술": ("수술", "수술", "", 2.2, 1.35),
    "3종 수술": ("수술", "수술", "", 1.0, 0.9),
    "4종 수술": ("수술", "수술", "", 0.9, 2.1),
    "5종 수술": ("수술", "수술", "", 2.6, 0.95),
    "갑상선암 수술": ("수술", "수술", "", 0.5, 2.2),
    "질병II(백내장및대장용종제외)수술": ("수술", "수술", "", 1.2, 2.0),
    "질병II(특정경증질환제외) 수술": ("수술", "수술", "", 0.9, 1.3),
    "3대기관 양성신생물 및 폴립 수술": ("수술", "수술", "", 0.8, 2.0),
    "허혈성심장질환 수술": ("수술", "수술", "", 0.6, 2.1),
    "재해수술": ("수술", "수술", "", 1.0, 2.0),
    "질병II(백내장및대장용종제외)종합병원 수술": ("수술", "수술", "", 0.7, 1.95),
    "허혈성심장질환 입원": ("입원", "입원", "", 0.6, 2.0),
    "간병인사용입원(요양병원제외)": ("입원", "입원", "", 2.4, 2.1),
    "첫날부터종합병원입원": ("입원", "입원", "", 2.4, 2.05),
    "첫날부터입원특약(120일)": ("입원", "입원", "", 1.1, 2.0),
    "첫날부터입원특약(30일)": ("입원", "입원", "", 1.0, 1.3),
    "간호간병통합서비스입원(요양병원제외)": ("입원", "입원", "", 0.8, 1.95),
    "재해골절치료": ("기타", "기타", "", 2.2, 0.9),
    "깁스치료": ("기타", "기타", "", 2.0, 0.8),
    "급여 부목치료": ("기타", "기타", "", 1.8, 0.85),
    "가족일상생활배상책임": ("기타", "기타", "", 0.6, 0.4),
    "운전자벌금": ("기타", "기타", "", 0.5, 0.3),
}


def month_list():
    y, m = LAUNCH
    out = []
    while (y, m) <= LAST_CLOSE:
        out.append(y * 100 + m)
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def make_master(rng: np.random.Generator) -> pd.DataFrame:
    names = list(BENEFITS)
    w = np.array([BENEFITS[b][3] for b in names])
    # 모든 급부가 최소 1개 담보를 갖도록 한 뒤 나머지는 가중 랜덤 배정
    picks = names + list(rng.choice(names, size=N_CODES - len(names), p=w / w.sum()))
    rng.shuffle(picks)
    rows = []
    for i, b in enumerate(picks, start=1):
        big, grp, godang, _, _ = BENEFITS[b]
        rows.append({"담보코드": f"CV{i:04d}", "담보명": b, "대급부명": big,
                     "구분": grp if grp != big else "", "고당대통": godang})
    return pd.DataFrame(rows)


def make_raw(master: pd.DataFrame, rng: np.random.Generator) -> pd.DataFrame:
    n = len(master)
    scale = rng.lognormal(mean=11.5, sigma=1.0, size=n) * master["담보명"].map(
        {k: v[3] for k, v in BENEFITS.items()}).to_numpy()
    exp_lr = master["담보명"].map({k: v[4] for k, v in BENEFITS.items()}).to_numpy()
    exp_lr = exp_lr * rng.lognormal(0, 0.2, size=n)
    main_prod = rng.choice([f"MP{i:03d}" for i in range(1, 16)], size=n)
    rider_prod = np.array([f"RP{rng.integers(1, 200):04d}" for _ in range(n)])

    months = month_list()
    frames = []
    for i, cm in enumerate(months):                    # 계약월
        new_biz = 0.6 + 0.05 * i                       # 출시 후 판매 증가
        lr_drift = rng.normal(1.0, 0.08, size=n)       # 계약월별 손해율 변동
        for j, clm in enumerate(months[i:]):           # 마감월(계약월 이후)
            persist = 0.97 ** j                        # 유지율
            maturity = min(1.3, 0.4 + 0.15 * j)        # 초기 경과월 사고 적음
            for uw in UW_TYPES:
                uw_mult = 1.0 if uw == "일반" else 0.45
                uw_lr = 1.0 if uw == "일반" else 1.18
                prem = scale * new_biz * persist * uw_mult * rng.lognormal(0, 0.15, size=n)
                lr = exp_lr * lr_drift * uw_lr * maturity * rng.lognormal(0, 0.35, size=n)
                claim = prem * lr
                reserve_chg = claim * rng.normal(0.08, 0.10, size=n)
                cnt = rng.poisson(np.maximum(claim / 400_000, 0.05))
                frames.append(pd.DataFrame({
                    "계약년월": cm,
                    "마감년월": clm,
                    "주상품코드": main_prod,
                    "특약상품코드": rider_prod,
                    "경과개월수": j + 1,
                    "심사유형(일반.간편)": uw,
                    "담보코드": master["담보코드"].to_numpy(),
                    "위험보험료 합계": prem.round(0).astype(np.int64),
                    "지급보험금합계": claim.round(0).astype(np.int64),
                    "지급비금변동분 합계": reserve_chg.round(0).astype(np.int64),
                    "건수": cnt,
                }))
    raw = pd.concat(frames, ignore_index=True)
    # 마스터에 없는 담보코드 몇 건 섞기 (미매핑 처리 검증용)
    orphan = raw.sample(5, random_state=SEED).copy()
    orphan["담보코드"] = ["ZZ9001", "ZZ9002", "ZZ9003", "ZZ9004", "ZZ9005"]
    return pd.concat([raw, orphan], ignore_index=True)


def main() -> None:
    rng = np.random.default_rng(SEED)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    master = make_master(rng)
    raw = make_raw(master, rng)
    master.to_csv(DATA_DIR / "code_master.csv", index=False, encoding="utf-8-sig")
    raw.to_csv(DATA_DIR / "raw_data.csv", index=False, encoding="utf-8-sig")
    print(f"[OK] code_master.csv : {len(master):,}건 (급부 {master['담보명'].nunique()}종)")
    print(f"[OK] raw_data.csv    : {len(raw):,}건 (계약 {raw['계약년월'].min()}~, 마감 ~{raw['마감년월'].max()})")


if __name__ == "__main__":
    main()
