"""
[Phase 1 전용] 더미 데이터 생성기

생성 파일
  data/code_master.csv : 담보 마스터 545건 (담보코드, 담보명, 대급부명)
  data/raw_data.csv    : 로데이터 (출시 202505 ~ 마감 202609)

※ Phase 2(내부망)에서는 이 파일을 사용하지 않는다.
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
N_CODES = 545
DATA_DIR = Path(__file__).resolve().parent / "data"

# 생성 구간: 출시월 ~ 마감월 (계약년월 · 마감년월 모두 이 구간)
LAUNCH = (2025, 5)
LAST_CLOSE = (2026, 9)
UW_TYPES = ["일반", "간편"]

# 대급부명: (담보명 예시, 위험보험료 규모 가중치, 기대 손해율)
BENEFITS = {
    "사망":     (["일반사망", "질병사망", "상해사망", "재해사망"], 3.0, 0.55),
    "암":       (["암진단비", "유사암진단비", "고액암진단비", "암수술비", "암직접치료비"], 2.5, 0.80),
    "뇌혈관":   (["뇌졸중진단비", "뇌출혈진단비", "뇌혈관질환진단비"], 1.5, 0.75),
    "심장":     (["급성심근경색진단비", "허혈성심장질환진단비", "심혈관질환진단비"], 1.5, 0.72),
    "수술":     (["질병수술비", "상해수술비", "1~5종수술비", "특정질병수술비"], 1.2, 0.95),
    "입원":     (["질병입원일당", "상해입원일당", "중환자실입원일당"], 1.0, 1.05),
    "상해":     (["상해후유장해", "골절진단비", "화상진단비", "깁스치료비"], 0.8, 0.85),
    "실손의료": (["질병입원실손", "질병통원실손", "상해입원실손", "상해통원실손"], 2.0, 1.15),
    "치아":     (["보존치료비", "보철치료비", "임플란트치료비"], 0.5, 0.90),
    "기타":     (["배상책임", "운전자벌금", "교통사고처리지원금", "가족일상생활배상"], 0.4, 0.60),
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
    weights = np.array([BENEFITS[b][1] for b in names])
    picks = rng.choice(names, size=N_CODES, p=weights / weights.sum())
    rows = []
    for i, ben in enumerate(picks, start=1):
        base = rng.choice(BENEFITS[ben][0])
        variant = rng.choice(["", "(갱신형)", "(비갱신형)", "(간편)", "Ⅱ", "(20년만기)"])
        rows.append({"담보코드": f"CV{i:04d}", "담보명": f"{base}{variant}", "대급부명": ben})
    return pd.DataFrame(rows)


def make_raw(master: pd.DataFrame, rng: np.random.Generator) -> pd.DataFrame:
    n = len(master)
    # 담보별 고유 특성 (연도와 무관하게 고정 → 전년 대비 비교가 의미 있도록)
    scale = rng.lognormal(mean=12.5, sigma=1.2, size=n)               # 월 위험보험료 규모(원)
    exp_lr = master["대급부명"].map({k: v[2] for k, v in BENEFITS.items()}).to_numpy()
    exp_lr = exp_lr * rng.lognormal(0, 0.25, size=n)                   # 담보별 손해율 편차
    main_prod = rng.choice([f"MP{i:03d}" for i in range(1, 16)], size=n)
    rider_prod = np.array([f"RP{rng.integers(1, 200):04d}" for _ in range(n)])

    months = month_list()
    frames = []
    for i, cm in enumerate(months):                    # 계약월
        new_biz = 0.6 + 0.05 * i                       # 출시 후 판매 증가
        lr_drift = rng.normal(1.0, 0.06, size=n)       # 계약월별 손해율 변동
        for j, clm in enumerate(months[i:]):           # 마감월(계약월 이후)
            elapsed = j + 1
            persist = 0.97 ** j                        # 유지율
            maturity = min(1.0, 0.35 + 0.15 * j)       # 초기 경과월 사고 적음(손해 지연)
            for uw in UW_TYPES:
                uw_mult = 1.0 if uw == "일반" else 0.45
                uw_lr = 1.0 if uw == "일반" else 1.18
                prem = scale * new_biz * persist * uw_mult * rng.lognormal(0, 0.15, size=n)
                lr = exp_lr * lr_drift * uw_lr * maturity * rng.lognormal(0, 0.35, size=n)
                claim = prem * lr
                reserve_chg = claim * rng.normal(0.08, 0.10, size=n)   # 지급비금 변동분(±)
                frames.append(pd.DataFrame({
                    "계약년월": cm,
                    "마감년월": clm,
                    "주상품코드": main_prod,
                    "특약상품코드": rider_prod,
                    "경과개월수": elapsed,
                    "심사유형(일반.간편)": uw,
                    "담보코드": master["담보코드"].to_numpy(),
                    "위험보험료 합계": prem.round(0).astype(np.int64),
                    "지급보험금합계": claim.round(0).astype(np.int64),
                    "지급비금변동분 합계": reserve_chg.round(0).astype(np.int64),
                }))
    raw = pd.concat(frames, ignore_index=True)
    # 마스터에 없는 담보코드 몇 건 섞기 (Left Join 미매핑 처리 검증용)
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
    print(f"[OK] code_master.csv : {len(master):,}건")
    print(f"[OK] raw_data.csv    : {len(raw):,}건 "
          f"(계약년월 {sorted(raw['계약년월'].unique())})")


if __name__ == "__main__":
    main()
