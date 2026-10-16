"""
[외부망 전용] 엑셀 집계 템플릿 생성기 — 내부망에는 결과 .xlsx / .bas 만 반입한다.

  excel/손해율_집계.xlsx : 수식 기반 집계 템플릿 (더미 데이터 포함)
  excel/report_macro.bas  : 워드 리포트 생성 + 수식 채우기 VBA (CP949)

모든 집계는 엑셀 수식(VLOOKUP / SUMIFS / RANK / INDEX·MATCH)으로 되어 있어
내부망에서는 파이썬 없이 엑셀만으로 동작한다.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

BASE = Path(__file__).resolve().parent
OUT = BASE / "excel" / "손해율_집계.xlsx"

FONT = "맑은 고딕"
NAVY = "1F3864"
F_BASE = Font(name=FONT, size=9)
F_HDR = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=11, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=8.5, color="595959")
F_INPUT = Font(name=FONT, size=10, bold=True, color="0000FF")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_ZEBRA = PatternFill("solid", fgColor="F2F5FA")
FILL_HELPER = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="A6B3C9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LEFT, RIGHT, CENTER = (Alignment(horizontal=h, vertical="center") for h in ("left", "right", "center"))

NF_AMT = '#,##0,,;-#,##0,,;"-"'          # 원 → 백만원 표시
NF_PCT = '0.0%;-0.0%;"-"'
NF_PCT_S = '+0.0%;-0.0%;0.0%'
NF_PT = '[Red]+0.0%"p";[Blue]-0.0%"p";0.0%"p"'   # 손해율 증감: 악화 빨강 / 개선 파랑
NF_YM = '00\\.00'

RAW = "raw_data"
MST = "code_master"
AGG = "담보별집계"
R = f"{RAW}!"


def hdr_row(ws, row, col, headers, widths=None):
    for i, h in enumerate(headers):
        c = ws.cell(row, col + i, h)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, Alignment(
            horizontal="center", vertical="center", wrap_text=True), BORDER
        if widths:
            ws.column_dimensions[get_column_letter(col + i)].width = widths[i]


def body(c, fmt=None, align=LEFT, zebra=False):
    c.font, c.alignment, c.border = F_BASE, align, BORDER
    if fmt:
        c.number_format = fmt
    if zebra:
        c.fill = FILL_ZEBRA


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ---------------------------------------------------------------------------
def sheet_guide(wb):
    ws = wb.active
    ws.title = "사용법"
    ws.column_dimensions["A"].width = 110
    lines = [
        ("손해율 월말 집계 템플릿 (엑셀 수식 + VBA 워드 출력)", F_TITLE),
        ("", None),
        ("[매월 작업 순서]", F_SUB),
        ("1. '설정' 시트의 노란 칸(기준월, UY 계약 시작월)을 입력한다.  예) 25.09 / 25.06", F_BASE),
        ("2. 'raw_data' 시트 A~J열(원본 10개 컬럼)에 로데이터를 붙여넣는다. 당해 + 전년 동기 구간이 모두 있어야 증감이 계산된다.", F_BASE),
        ("   ※ K~S열(연두색 헤더)은 수식 열이다. 붙여넣은 행 수만큼 수식을 채운다 → 매크로 'PrepareRawData' 실행 (또는 K2:S2 선택 후 채우기 핸들 더블클릭).", F_BASE),
        ("   ※ 이전 달 데이터보다 행이 줄었으면 남은 아래쪽 행은 삭제한다 (매크로는 자동 정리).", F_BASE),
        ("3. 담보 마스터가 바뀌었으면 'code_master' 시트 A~C열을 교체한다. (담보별집계는 마스터 2~546행 = 545개 기준, 늘어나면 행을 아래로 복사)", F_BASE),
        ("4. '요약', 'Top20' 시트에서 결과를 확인한다. (수식이 자동 계산됨, 수동계산 모드라면 F9)", F_BASE),
        ("5. 워드 리포트: 매크로 'MakeWordReport' 실행 → 이 파일과 같은 폴더에 손해율_결산리포트_YY.MM.docx 생성.", F_BASE),
        ("   매크로가 막혀 있으면 '요약'/'Top20' 시트의 표를 복사해 워드에 붙여넣기(원본 서식 유지)하면 동일한 표가 된다.", F_BASE),
        ("", None),
        ("[매크로 최초 1회 설치]", F_SUB),
        ("Alt+F11 → 파일 → 파일 가져오기 → report_macro.bas 선택 → 닫고, 이 파일을 '엑셀 매크로 사용 통합 문서(*.xlsm)'로 다른 이름 저장.", F_BASE),
        ("이후 Alt+F8 → PrepareRawData / MakeWordReport 실행.", F_BASE),
        ("", None),
        ("[계산 정의]", F_SUB),
        ("손해액 = 지급보험금합계 + 지급비금변동분 합계,  손해율 = 손해액 ÷ 위험보험료", F_BASE),
        ("UY 누적 = 계약년월이 [UY 시작월 ~ 기준월]이고 마감년월 ≤ 기준월인 실적의 누적", F_BASE),
        ("CY 당월 = 마감년월 = 기준월인 실적 (계약년월 무관)", F_BASE),
        ("전년 동기 = 위 조건을 12개월 앞당긴 구간,  손해율 증감(p) = 당해 손해율 − 전년 손해율", F_BASE),
        ("손해율 Top 20 = 위험보험료 비중이 '설정'의 최소 비중 이상인 담보 중 손해율 상위 (0이면 제한 없음)", F_BASE),
        ("마스터에 없는 담보코드는 '(미매핑)'으로 표시되며, 전사 합계에는 포함되고 담보별 Top 20에서는 제외된다.", F_BASE),
        ("", None),
        ("[주의]", F_SUB),
        ("년월은 25.06 또는 2025.06 형식(숫자/텍스트 모두 가능). 202506처럼 점이 없는 형식이면 raw_data N·O열 수식을 수정해야 한다.", F_BASE),
        ("담보코드가 숫자로 저장된 경우 raw_data와 code_master의 형식(텍스트/숫자)을 반드시 일치시킨다 (VLOOKUP 미매칭 원인).", F_BASE),
        ("엑셀 최대 행 수는 1,048,576행. 로데이터가 수십만 행 이상이면 계산이 수십 초 걸릴 수 있다.", F_BASE),
    ]
    for i, (t, f) in enumerate(lines, start=1):
        c = ws.cell(i, 1, t)
        c.font = f or F_BASE


def sheet_settings(wb):
    ws = wb.create_sheet("설정")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 60
    ws["A1"], ws["A1"].font = "기본 설정", F_TITLE
    ws["A2"], ws["A2"].font = "노란 칸만 입력하세요. 나머지는 자동 계산.", F_NOTE
    rows = [
        ("마감 기준월", 25.09, NF_YM, True, "CY 당월 = 이 달, UY 누적 = 이 달 마감까지  (예: 25.09)"),
        ("UY 계약 시작월", 25.06, NF_YM, True, "UY 누적 대상 계약년월 시작  (예: 25.06)"),
        ("손해율 Top20 최소 보험료 비중", 0.0005, "0.00%", True, "보험료가 너무 작은 담보의 극단값 제외 (0이면 제한 없음)"),
        ("기준월 (YYMM)", "=MOD(ROUND(B4*100,0),10000)", "0", False, "자동"),
        ("UY 시작월 (YYMM)", "=MOD(ROUND(B5*100,0),10000)", "0", False, "자동"),
        ("전년 기준월 (YYMM)", "=B7-100", "0", False, "자동"),
        ("전년 UY 시작월 (YYMM)", "=B8-100", "0", False, "자동"),
    ]
    for i, (label, val, fmt, is_input, desc) in enumerate(rows, start=4):
        ws.cell(i, 1, label).font = F_BASE
        c = ws.cell(i, 2, val)
        c.number_format, c.border, c.alignment = fmt, BORDER, RIGHT
        c.font = F_INPUT if is_input else F_BASE
        if is_input:
            c.fill = FILL_INPUT
        ws.cell(i, 3, desc).font = F_NOTE
    for nm, ref in [("LR_MIN_SHARE", "$B$6"), ("BASE_YM", "$B$7"), ("UY_START", "$B$8"),
                    ("BASE_YM_PY", "$B$9"), ("UY_START_PY", "$B$10")]:
        name(wb, nm, f"설정!{ref}")
    ws["B4"].comment = Comment("숫자로 25.09 입력 (10월은 25.10 → 표시 25.10)", "template")


RAW_COLS = ["계약년월", "마감년월", "주상품코드", "특약상품코드", "경과개월수", "심사유형(일반.간편)",
            "담보코드", "위험보험료 합계", "지급보험금합계", "지급비금변동분 합계"]
HELPERS = [  # (헤더, 수식 템플릿{r})
    ("담보명", '=IFERROR(VLOOKUP($G{r},code_master!$A:$C,2,FALSE),"(미매핑)")'),
    ("대급부명", '=IFERROR(VLOOKUP($G{r},code_master!$A:$C,3,FALSE),"(미매핑)")'),
    ("손해액", "=I{r}+J{r}"),
    ("계약YM", "=MOD(ROUND(VALUE(A{r})*100,0),10000)"),
    ("마감YM", "=MOD(ROUND(VALUE(B{r})*100,0),10000)"),
    ("UY당해", "=IF(AND(N{r}>=UY_START,N{r}<=BASE_YM,O{r}<=BASE_YM),1,0)"),
    ("UY전년", "=IF(AND(N{r}>=UY_START_PY,N{r}<=BASE_YM_PY,O{r}<=BASE_YM_PY),1,0)"),
    ("CY당해", "=IF(O{r}=BASE_YM,1,0)"),
    ("CY전년", "=IF(O{r}=BASE_YM_PY,1,0)"),
]


def sheet_raw(wb, raw: pd.DataFrame):
    ws = wb.create_sheet(RAW)
    hdr_row(ws, 1, 1, RAW_COLS + [h for h, _ in HELPERS],
            [8, 8, 10, 11, 8, 12, 9, 14, 13, 14, 22, 10, 13, 8, 8, 7, 7, 7, 7])
    for j in range(len(RAW_COLS) + 1, len(RAW_COLS) + len(HELPERS) + 1):
        ws.cell(1, j).fill = PatternFill("solid", fgColor="548235")
    for r, rec in enumerate(raw.itertuples(index=False), start=2):
        for j, v in enumerate(rec, start=1):
            ws.cell(r, j, v)
        for j, (_, f) in enumerate(HELPERS, start=len(RAW_COLS) + 1):
            ws.cell(r, j, f.format(r=r))
    ws.freeze_panes = "A2"


def sheet_master(wb, master: pd.DataFrame):
    ws = wb.create_sheet(MST)
    hdr_row(ws, 1, 1, list(master.columns), [10, 28, 12])
    for r, rec in enumerate(master.itertuples(index=False), start=2):
        for j, v in enumerate(rec, start=1):
            ws.cell(r, j, v)
    ws.freeze_panes = "A2"


def sheet_agg(wb, n_codes: int):
    ws = wb.create_sheet(AGG)
    last = n_codes + 1
    blk = ["위험보험료", "손해액", "손해율", "전년 위험보험료", "전년 손해액", "전년 손해율",
           "위험보험료 증감률", "손해율 증감(p)"]
    heads = (["담보코드", "대급부명", "담보명"] + [f"UY {h}" for h in blk] + [f"CY {h}" for h in blk]
             + ["UY 보험료순위", "UY 손해율키", "UY 손해율순위", "CY 보험료순위", "CY 손해율키", "CY 손해율순위"])
    hdr_row(ws, 1, 1, heads, [9, 10, 24] + [12] * 22)
    ws.row_dimensions[1].height = 30
    fmts = [NF_AMT, NF_AMT, NF_PCT, NF_AMT, NF_AMT, NF_PCT, NF_PCT_S, NF_PT]
    for r in range(2, last + 1):
        ws.cell(r, 1, f"={MST}!A{r}")
        ws.cell(r, 2, f"={MST}!C{r}")
        ws.cell(r, 3, f"={MST}!B{r}")
        for start, cur, prev in ((4, "P", "Q"), (12, "R", "S")):
            L = [get_column_letter(start + k) for k in range(8)]
            f = [
                f"=SUMIFS({R}$H:$H,{R}$G:$G,$A{r},{R}${cur}:${cur},1)",
                f"=SUMIFS({R}$M:$M,{R}$G:$G,$A{r},{R}${cur}:${cur},1)",
                f'=IF({L[0]}{r}=0,"",{L[1]}{r}/{L[0]}{r})',
                f"=SUMIFS({R}$H:$H,{R}$G:$G,$A{r},{R}${prev}:${prev},1)",
                f"=SUMIFS({R}$M:$M,{R}$G:$G,$A{r},{R}${prev}:${prev},1)",
                f'=IF({L[3]}{r}=0,"",{L[4]}{r}/{L[3]}{r})',
                f'=IF({L[3]}{r}=0,"",{L[0]}{r}/{L[3]}{r}-1)',
                f'=IF(OR({L[2]}{r}="",{L[5]}{r}=""),"",{L[2]}{r}-{L[5]}{r})',
            ]
            for k in range(8):
                body(ws.cell(r, start + k, f[k]), fmts[k], RIGHT)
        # 순위 (동점은 위쪽 행 우선 → 고유 순위)
        for base_col, p, lr in ((20, "D", "F"), (23, "L", "N")):
            rk, key, lrk = (get_column_letter(base_col + k) for k in range(3))
            ws.cell(r, base_col, f'=IF({p}{r}>0,RANK({p}{r},{p}$2:{p}${last})+COUNTIF({p}$2:{p}{r},{p}{r})-1,"")')
            ws.cell(r, base_col + 1,
                    f"=IF(AND({p}{r}>0,{p}{r}>=LR_MIN_SHARE*SUM({p}$2:{p}${last})),{lr}{r},-1)")
            ws.cell(r, base_col + 2,
                    f'=IF({key}{r}<0,"",RANK({key}{r},{key}$2:{key}${last})+COUNTIF({key}$2:{key}{r},{key}{r})-1)')
            for k in range(3):
                body(ws.cell(r, base_col + k), "0.0000" if k == 1 else "0", RIGHT)
        for c in range(1, 4):
            body(ws.cell(r, c))
    ws.freeze_panes = "D2"
    return last


def period_texts():
    uy = ('=TEXT(UY_START,"00\\.00")&"~"&TEXT(BASE_YM,"00\\.00")&" 계약, "&TEXT(BASE_YM,"00\\.00")&" 마감 누적"')
    cy = '=TEXT(BASE_YM,"00\\.00")&" 마감월"'
    return uy, cy


def sheet_summary(wb, benefits):
    ws = wb.create_sheet("요약")
    widths = [16, 34, 14, 14, 11, 11, 14, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws["A1"] = '=TEXT(BASE_YM,"00\\.00")&" 기준 손해율 결산 리포트"'
    ws["A1"].font = F_TITLE
    ws["A2"] = "금액 단위: 백만원  |  손해율 = (지급보험금 + 지급비금변동분) ÷ 위험보험료  |  증감(p): 빨강 악화 / 파랑 개선"
    ws["A2"].font = F_NOTE
    uy_txt, cy_txt = period_texts()

    def sums(flag, extra=""):
        return f"SUMIFS({R}$H:$H,{R}${flag}:${flag},1{extra})", f"SUMIFS({R}$M:$M,{R}${flag}:${flag},1{extra})"

    def metric_cells(ws, r, c0, cur, prev, extra=""):
        pc, lc = sums(cur, extra)
        pp, lp = sums(prev, extra)
        P, Lc = get_column_letter(c0), get_column_letter(c0 + 1)
        vals = [f"={pc}", f"={lc}", f'=IF({P}{r}=0,"",{Lc}{r}/{P}{r})',
                f'=IFERROR({lp}/{pp},"")', f'=IFERROR({P}{r}/{pp}-1,"")',
                f'=IF(OR({get_column_letter(c0+2)}{r}="",{get_column_letter(c0+3)}{r}=""),"",'
                f'{get_column_letter(c0+2)}{r}-{get_column_letter(c0+3)}{r})']
        for k, (v, fmt) in enumerate(zip(vals, [NF_AMT, NF_AMT, NF_PCT, NF_PCT, NF_PCT_S, NF_PT])):
            body(ws.cell(r, c0 + k, v), fmt, RIGHT, zebra=(r % 2 == 1))

    def section(row, title):
        ws.cell(row, 1, title).font = F_SUB
        return row + 1

    # 1. 전사 요약
    row = section(4, "1. 핵심 요약 (전사)")
    hdr_row(ws, row, 1, ["구분", "기간", "위험보험료", "손해액", "손해율", "전년 손해율",
                         "위험보험료 증감률", "손해율 증감(p)"])
    top = row
    for label, txt, cur, prev in (("UY 누적", uy_txt, "P", "Q"), ("CY 당월", cy_txt, "R", "S")):
        row += 1
        body(ws.cell(row, 1, label), align=CENTER, zebra=(row % 2 == 1))
        body(ws.cell(row, 2, txt), zebra=(row % 2 == 1))
        metric_cells(ws, row, 3, cur, prev)
    name(wb, "RPT_SUMMARY", f"요약!$A${top}:$H${row}")
    row += 1
    ws.cell(row, 1, '="※ 전년 동기: UY "&TEXT(UY_START_PY,"00\\.00")&"~"&TEXT(BASE_YM_PY,"00\\.00")'
                    '&" 계약, "&TEXT(BASE_YM_PY,"00\\.00")&" 마감 누적 / CY "&TEXT(BASE_YM_PY,"00\\.00")&" 마감월"').font = F_NOTE

    # 2·3. UY / CY 별 심사유형·대급부
    for no, (lab, cur, prev, rng_key) in enumerate((("UY 누적", "P", "Q", "UY"), ("CY 당월", "R", "S", "CY")), start=2):
        row = section(row + 2, f"{no}. {lab} 손해율")
        for sub, col_letter, keys, nm in ((f"{no}-1. 심사유형별", "F", ["일반", "간편"], "UW"),
                                          (f"{no}-2. 대급부별", "L", benefits, "BEN")):
            ws.cell(row, 1, sub).font = Font(name=FONT, size=10, bold=True, color=NAVY)
            row += 1
            hdr_row(ws, row, 1, ["심사유형" if nm == "UW" else "대급부명", "", "위험보험료", "손해액", "손해율",
                                 "전년 손해율", "위험보험료 증감률", "손해율 증감(p)"])
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            top = row
            for k in keys:
                row += 1
                body(ws.cell(row, 1, k), zebra=(row % 2 == 1))
                body(ws.cell(row, 2), zebra=(row % 2 == 1))
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
                metric_cells(ws, row, 3, cur, prev, f',{R}${col_letter}:${col_letter},$A{row}')
            name(wb, f"RPT_{rng_key}_{nm}", f"요약!$A${top}:$H${row}")
            row += 1
    ws.sheet_view.showGridLines = False


def sheet_top(wb, last):
    ws = wb.create_sheet("Top20")
    widths = [6, 10, 11, 28, 13, 14, 11, 14, 4, 6]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions["J"].hidden = True    # MATCH 결과(행 위치) 보조열
    ws["A1"] = "담보별 Top 20  (금액 단위: 백만원)"
    ws["A1"].font = F_TITLE
    A = f"'{AGG}'!"
    blocks = [  # 제목, 순위열, 보험료, 증감률, 손해율, 증감p, 이름
        ("UY 누적 · 위험보험료 상위 20 담보", "T", "D", "J", "F", "K", "RPT_UY_TOPP"),
        ("UY 누적 · 손해율 상위 20 담보", "V", "D", "J", "F", "K", "RPT_UY_TOPL"),
        ("CY 당월 · 위험보험료 상위 20 담보", "W", "L", "R", "N", "S", "RPT_CY_TOPP"),
        ("CY 당월 · 손해율 상위 20 담보", "Y", "L", "R", "N", "S", "RPT_CY_TOPL"),
    ]
    row = 3
    for title, rk, prem, pchg, lr, lrchg, nm in blocks:
        ws.cell(row, 1, title).font = F_SUB
        row += 1
        hdr_row(ws, row, 1, ["순위", "담보코드", "대급부명", "담보명", "위험보험료", "위험보험료 증감률",
                             "손해율(%)", "전동기 대비 증감(p)"])
        top = row
        for n in range(1, 21):
            row += 1
            z = row % 2 == 1
            ws.cell(row, 10, f"=IFERROR(MATCH(A{row},{A}${rk}$2:${rk}${last},0),\"\")")
            body(ws.cell(row, 1, n), align=CENTER, zebra=z)
            for c, col, fmt, al in ((2, "A", None, LEFT), (3, "B", None, LEFT), (4, "C", None, LEFT),
                                    (5, prem, NF_AMT, RIGHT), (6, pchg, NF_PCT_S, RIGHT),
                                    (7, lr, NF_PCT, RIGHT), (8, lrchg, NF_PT, RIGHT)):
                body(ws.cell(row, c, f'=IF($J{row}="","",INDEX({A}${col}$2:${col}${last},$J{row}))'),
                     fmt, al, zebra=z)
        name(wb, nm, f"Top20!$A${top}:$H${row}")
        row += 1
        if "손해율" in title:
            ws.cell(row, 1, '="※ 위험보험료 비중 "&TEXT(LR_MIN_SHARE,"0.00%")&" 이상 담보 대상"').font = F_NOTE
        row += 2
    ws.sheet_view.showGridLines = False


def main():
    raw = pd.read_csv(BASE / "data" / "raw_data.csv", dtype={"계약년월": str, "마감년월": str, "담보코드": str})
    master = pd.read_csv(BASE / "data" / "code_master.csv", dtype=str)
    benefits = sorted(master["대급부명"].unique()) + ["(미매핑)"]
    wb = Workbook()
    sheet_guide(wb)
    sheet_settings(wb)
    sheet_summary(wb, benefits)
    last = len(master) + 1
    sheet_top(wb, last)
    sheet_agg(wb, len(master))
    sheet_raw(wb, raw)
    sheet_master(wb, master)
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"[OK] {OUT}")


if __name__ == "__main__":
    main()
