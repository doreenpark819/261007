"""
[외부망 전용] 엑셀 집계 템플릿 생성기 — 내부망에는 결과 .xlsx / .bas 만 반입한다.

  excel/손해율_집계.xlsx : 수식 기반 집계 템플릿 (더미 데이터 포함)
  excel/report_macro.bas  : 데이터 불러오기 + 워드 리포트 VBA (CP949)

설계 원칙 (내부망에서 손이 안 가고 깨지지 않게)
  - 사용자가 입력하는 칸은 '실행' 시트의 기준월·UY 시작월 2개뿐
  - raw_data 에는 원본 10개 컬럼만 붙여넣는다 (보조 수식 열 없음 → 행 수가 바뀌어도 그대로 동작)
  - 담보 마스터는 최대 MAX_CODES 개까지 자동 반영 (545 → 늘어나도 수정 불필요)
  - 대급부 목록도 마스터에서 자동 추출 (대급부명이 바뀌어도 수정 불필요)
  - '실행' 시트의 데이터 점검표가 년월 형식·텍스트 숫자·미매핑·중복 등을 자동 검사
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

BASE = Path(__file__).resolve().parent
OUT = BASE / "excel" / "손해율_집계.xlsx"

MAX_CODES = 1000       # 담보 마스터 수용 한도
MAX_BENEFITS = 20      # 대급부 종류 수용 한도
TOP_N = 20

FONT = "맑은 고딕"
NAVY = "1F3864"
F_BASE = Font(name=FONT, size=9)
F_BOLD = Font(name=FONT, size=9, bold=True)
F_HDR = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=11, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=8.5, color="595959")
F_INPUT = Font(name=FONT, size=12, bold=True, color="0000FF")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_ZEBRA = PatternFill("solid", fgColor="F2F5FA")
THIN = Side(style="thin", color="A6B3C9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LEFT, RIGHT, CENTER = (Alignment(horizontal=h, vertical="center") for h in ("left", "right", "center"))
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

NF_AMT = '#,##0,,;-#,##0,,;"-"'                      # 원 → 백만원 표시
NF_PCT = '0.0%;-0.0%;"-"'
NF_PCT_S = '+0.0%;-0.0%;0.0%'
NF_PT = '[Red]+0.0%"p";[Blue]-0.0%"p";0.0%"p"'       # 손해율 증감: 악화 빨강 / 개선 파랑

RAW, MST, AGG = "raw_data", "code_master", "담보별집계"
R = f"{RAW}!"
A_ = f"'{AGG}'!"
LAST = MAX_CODES + 1          # 담보별집계 마지막 행

# 집계 구간 조건 (SUMIFS 조건 인자) — 년월은 202609 같은 숫자
#   C = UY 누적, M = UY 당월, A = 전기간(출시 이후) 누적 / py=True 는 비교기간
CRIT = {
    ("C", False): f'{R}$A:$A,">="&UY_START,{R}$A:$A,"<="&BASE_YM,{R}$B:$B,"<="&BASE_YM',
    ("C", True): f'{R}$A:$A,">="&CMP_START,{R}$A:$A,"<="&CMP_END,{R}$B:$B,"<="&CMP_CLOSE',
    ("M", False): f'{R}$A:$A,">="&UY_START,{R}$A:$A,"<="&BASE_YM,{R}$B:$B,BASE_YM',
    ("M", True): f'{R}$A:$A,">="&CMP_START,{R}$A:$A,"<="&CMP_END,{R}$B:$B,CMP_CLOSE',
    ("A", False): f'{R}$A:$A,">="&LAUNCH_YM,{R}$B:$B,"<="&BASE_YM',
    ("A", True): f'{R}$A:$A,">="&LAUNCH_YM,{R}$A:$A,"<="&CMP_END,{R}$B:$B,"<="&CMP_CLOSE',
}
LABEL = {"C": "LABEL_C", "M": "LABEL_M"}


def prem_f(win, py, extra=""):
    return f"SUMIFS({R}$H:$H,{CRIT[(win, py)]}{extra})"


def loss_f(win, py, extra=""):
    return f"(SUMIFS({R}$I:$I,{CRIT[(win, py)]}{extra})+SUMIFS({R}$J:$J,{CRIT[(win, py)]}{extra}))"


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def hdr_row(ws, row, col, headers):
    for i, h in enumerate(headers):
        c = ws.cell(row, col + i, h)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def body(c, fmt=None, align=LEFT, zebra=False):
    c.font, c.alignment, c.border = F_BASE, align, BORDER
    if fmt:
        c.number_format = fmt
    if zebra:
        c.fill = FILL_ZEBRA


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# 실행 시트: 입력 2칸 + 작업 순서 + 데이터 점검표
# ---------------------------------------------------------------------------
def sheet_run(wb):
    ws = wb.active
    ws.title = "실행"
    widths(ws, [3, 34, 16, 12, 70])
    ws.sheet_view.showGridLines = False
    ws["B1"], ws["B1"].font = "손해율 월말 집계", F_TITLE

    ws["B3"], ws["B3"].font = "① 입력 (노란 칸만, 매월 기준월 하나만 바꾸면 됨)", F_SUB
    inputs = [
        (4, "마감 기준월", 202609, "예) 202609"),
        (5, "출시월 (전기간 시작)", 202505, "상품 출시월. 보통 바꿀 일 없음"),
        (6, "손해율 Top20 최소 보험료 비중", 0.0005, "보험료가 아주 작은 담보 제외용 (보통 그대로, 0이면 제한 없음)"),
    ]
    for r, label, val, desc in inputs:
        ws.cell(r, 2, label).font = F_BOLD
        c = ws.cell(r, 3, val)
        c.number_format, c.border, c.alignment = ("0.00%" if r == 6 else "0"), BORDER, CENTER
        c.font, c.fill = (F_INPUT, FILL_INPUT) if r == 4 else (F_BASE, PatternFill("solid", fgColor="FFF2CC"))
        ws.cell(r, 5, desc).font = F_NOTE
    grey = Font(name=FONT, size=8, color="A6A6A6")
    ws["D3"], ws["D3"].font = "자동값", grey
    prev = lambda x: f"IF(MOD({x},100)=1,{x}-89,{x}-1)"          # YYYYMM 한 달 전
    uy = lambda s_, e: (f'"UY"&MOD(INT({s_}/100),100)&"."&MOD({s_},100)&"~"'
                        f'&IF(INT({s_}/100)=INT({e}/100),"",MOD(INT({e}/100),100)&".")&MOD({e},100)')
    internal = [  # 셀, 수식, 이름
        ("D4", "=IFERROR(--TRIM(C4),0)", "BASE_YM"),
        ("D5", "=IFERROR(--TRIM(C5),0)", "LAUNCH_YM"),
        ("D7", "=INT(D4/100)*100+1", "UY_START"),          # 기준연도 1월
        ("D8", f"={prev('D7')}", "CMP_START"),             # 비교: 계약 시작 1개월 전
        ("D9", f"={prev('D4')}", "CMP_END"),               # 비교: 계약 끝 1개월 전
        ("D10", "=D9", "CMP_CLOSE"),                       # 비교: 마감월 1개월 전
    ]
    hidden = Font(name=FONT, size=8, color="FFFFFF")              # 구간 계산값은 흰 글씨(숨김)
    for ref, f, nm in internal:
        ws[ref] = f
        ws[ref].font, ws[ref].number_format = (grey if ref in ("D4", "D5") else hidden), "0"
        name(wb, nm, f"실행!${ref[0]}${ref[1:]}")
    name(wb, "LR_MIN_SHARE", "실행!$C$6")
    auto = [  # 행, 항목, 표시(C), 설명(E), 이름
        (7, "UY 누적 (자동)", f'={uy("D7", "D4")}&" 누적"',
         '="계약 "&D7&"~"&D4&" 의 "&D4&" 마감까지 누적 손해율"', "LABEL_C"),
        (8, "UY 당월 (자동)", '="UY"&MOD(INT(D4/100),100)&"."&MOD(D4,100)&" 당월"',
         '="계약 "&D7&"~"&D4&" 의 "&D4&" 한 달간 위험보험료·손해액"', "LABEL_M"),
        (9, "비교기간 (자동, 1개월 전)", f"={uy('D8', 'D9')}",
         '="계약 "&D8&"~"&D9&", 마감 "&D10&" 기준 (누적·당월 각각 같은 방식)"', "LABEL_CMP"),
        (10, "전기간 누적 (자동)", '="출시 이후 누적"',
         '="계약 "&D5&"~"&D4&" 의 "&D4&" 마감까지 누적 (비교: 계약 ~"&D9&", "&D10&" 마감)"', None),
    ]
    for r, label, disp, desc, nm in auto:
        ws.cell(r, 2, label).font = F_BASE
        c = ws.cell(r, 3, disp)
        c.font, c.alignment, c.border = F_BOLD, CENTER, BORDER
        ws.cell(r, 5, desc).font = F_NOTE
        if nm:
            name(wb, nm, f"실행!$C${r}")

    ws["B12"], ws["B12"].font = "② 매월 작업 순서", F_SUB
    steps = [
        "1) 위 노란 칸에 마감 기준월(예: 202609)만 입력 → UY 구간·비교기간은 자동 계산",
        "2) 로데이터 넣기 — 매크로 [LoadRawData] 실행 → 파일 선택 (기존 데이터는 자동 삭제)",
        "     매크로를 못 쓰면: raw_data 시트 전체 삭제 → A1에 헤더 포함 10개 열 붙여넣기 (값 붙여넣기)",
        "     ※ 로데이터는 출시월(202505)부터 기준월 마감까지 전체를 넣는다 (비교기간 계산에 지난달 마감분 필요)",
        "3) 아래 ③ 데이터 점검의 '종합 상태'가 정상인지 확인",
        "4) 워드 리포트 — 매크로 [MakeWordReport] 실행 → 이 파일과 같은 폴더에 docx 생성",
        "     매크로를 못 쓰면: '요약', 'Top20' 시트의 표를 복사해 워드에 붙여넣기(원본 서식 유지)",
        "담보 마스터가 바뀐 경우에만: code_master 시트 A~C열(담보코드·담보명·대급부명) 교체 (최대 1,000개 자동 반영)",
    ]
    for i, t in enumerate(steps, start=13):
        ws.cell(i, 2, t).font = F_BASE

    # ③ 데이터 점검
    top = 22
    ws.cell(top, 2, "③ 데이터 점검 (자동)").font = F_SUB
    hdr_row(ws, top + 1, 2, ["점검 항목", "결과", "상태", "문제일 때 조치"])
    rows_ = f"(COUNTA({R}$G:$G)-1)"
    agg_cnt = f"SUM({A_}$AD$2:$AD${LAST})"
    checks = [
        ("기준월 입력 형식", "=BASE_YM",
         "=IF(AND(BASE_YM>=190001,BASE_YM<=209912,MOD(BASE_YM,100)>=1,MOD(BASE_YM,100)<=12),\"정상\",\"오류\")",
         "202509처럼 6자리 숫자로 입력"),
        ("출시월 형식", "=LAUNCH_YM",
         "=IF(AND(LAUNCH_YM>=190001,MOD(LAUNCH_YM,100)>=1,MOD(LAUNCH_YM,100)<=12,LAUNCH_YM<=BASE_YM),\"정상\",\"오류\")",
         "6자리 숫자, 기준월보다 같거나 이전이어야 함"),
        ("로데이터 행 수", f"={rows_}",
         f"=IF({rows_}>0,\"정상\",\"오류\")", "raw_data 시트에 데이터가 없음 → 2) 단계 실행"),
        ("로데이터 열 수 (10개)", f"=COUNTA({R}$1:$1)",
         f"=IF(COUNTA({R}$1:$1)=10,\"정상\",\"확인\")", "원본 10개 열 순서대로 A~J열에 있어야 함"),
        ("년월이 숫자가 아닌 행", f"=2*{rows_}-COUNT({R}$A:$A)-COUNT({R}$B:$B)",
         f"=IF(2*{rows_}-COUNT({R}$A:$A)-COUNT({R}$B:$B)=0,\"정상\",\"오류\")",
         "텍스트로 저장된 년월 → [LoadRawData]/[MakeWordReport] 매크로가 자동 변환 (수동: 데이터→텍스트 나누기→마침)"),
        ("금액이 숫자가 아닌 칸", f"=3*{rows_}-COUNT({R}$H:$H)-COUNT({R}$I:$I)-COUNT({R}$J:$J)",
         f"=IF(3*{rows_}-COUNT({R}$H:$H)-COUNT({R}$I:$I)-COUNT({R}$J:$J)=0,\"정상\",\"오류\")",
         "위와 같음 (빈칸이면 0 입력)"),
        ("년월 범위 (열 순서 확인)", f"=MIN({R}$A:$B)&\" ~ \"&MAX({R}$A:$B)",
         f"=IF(AND(MIN({R}$A:$B)>=190001,MAX({R}$A:$B)<=209912),\"정상\",\"오류\")",
         "A열=계약년월, B열=마감년월이 맞는지 확인"),
        ("기준월 마감 데이터 행 수", f"=COUNTIF({R}$B:$B,BASE_YM)",
         f"=IF(COUNTIF({R}$B:$B,BASE_YM)>0,\"정상\",\"오류\")", "기준월 입력값 또는 로데이터 기간 확인"),
        ("비교기간 마감월 데이터 행 수", f"=COUNTIF({R}$B:$B,CMP_CLOSE)",
         f"=IF(COUNTIF({R}$B:$B,CMP_CLOSE)>0,\"정상\",\"확인\")", "없으면 동기간 증감이 공란 → 지난달 마감분까지 포함해 추출"),
        ("출시월 이전 계약 행", f"=COUNTIF({R}$A:$A,\"<\"&LAUNCH_YM)",
         f"=IF(COUNTIF({R}$A:$A,\"<\"&LAUNCH_YM)=0,\"정상\",\"확인\")", "출시월 입력값 확인 (전기간 누적에서 제외됨)"),
        ("기준월 이후 마감 행 (집계 제외)", f"=COUNTIF({R}$B:$B,\">\"&BASE_YM)",
         f"=IF(COUNTIF({R}$B:$B,\">\"&BASE_YM)=0,\"정상\",\"확인\")", "기준월을 잘못 입력했는지 확인 (과거 월 재작성이면 정상)"),
        ("심사유형이 일반/간편 외 값인 행", f"={rows_}-COUNTIF({R}$F:$F,\"일반\")-COUNTIF({R}$F:$F,\"간편\")",
         f"=IF({rows_}-COUNTIF({R}$F:$F,\"일반\")-COUNTIF({R}$F:$F,\"간편\")=0,\"정상\",\"확인\")",
         "심사유형별 표에서 빠짐 (전사 합계에는 포함)"),
        ("마스터에 없는 담보코드 행", f"={rows_}-{agg_cnt}",
         f"=IF({rows_}-{agg_cnt}=0,\"정상\",\"확인\")",
         "전사 합계엔 포함, 담보별 Top20엔 빠짐 → code_master에 담보 추가"),
        ("마스터 담보코드 중복", f"=SUM({A_}$AE$2:$AE${LAST})",
         f"=IF(SUM({A_}$AE$2:$AE${LAST})=0,\"정상\",\"오류\")", "code_master에서 중복 코드 제거"),
        ("마스터 담보 수 (최대 1,000)", f"=COUNTA({MST}!$A:$A)-1",
         f"=IF(COUNTA({MST}!$A:$A)-1<={MAX_CODES},\"정상\",\"오류\")", "1,000개 초과 시 템플릿 확장 필요"),
        ("대급부 종류 수 (최대 20)", f"=COUNT({A_}$AC$2:$AC${LAST})",
         f"=IF(COUNT({A_}$AC$2:$AC${LAST})<={MAX_BENEFITS},\"정상\",\"오류\")", "20종 초과 시 템플릿 확장 필요"),
        ("손해율 Top20 대상 담보 수", f"=COUNT({A_}$V$2:$V${LAST})",
         f"=IF(COUNT({A_}$V$2:$V${LAST})>={TOP_N},\"정상\",\"확인\")", "20개 미만이면 최소 보험료 비중을 낮추기"),
    ]
    r = top + 2
    first = r
    for label, val, status, fix in checks:
        body(ws.cell(r, 2, label))
        body(ws.cell(r, 3, val), "0" if label.endswith("형식") else "#,##0", RIGHT)
        body(ws.cell(r, 4, status), align=CENTER)
        body(ws.cell(r, 5, fix), align=WRAP)
        r += 1
    last = r - 1
    ws.cell(r + 1, 2, "종합 상태").font = Font(name=FONT, size=12, bold=True)
    c = ws.cell(r + 1, 3, f'=IF(COUNTIF(D{first}:D{last},"오류")>0,"오류",'
                          f'IF(COUNTIF(D{first}:D{last},"확인")>0,"확인 필요","정상"))')
    c.font, c.alignment, c.border = Font(name=FONT, size=12, bold=True), CENTER, BORDER
    ws.cell(r + 1, 5, "정상이면 바로 보고서 생성, '확인 필요'는 내용 확인 후 진행 가능, '오류'는 반드시 조치").font = F_NOTE
    name(wb, "CHECK_STATUS", f"실행!$C${r + 1}")
    for rng in (f"D{first}:D{last}", f"C{r + 1}"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"정상"'],
                                      fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"오류"'],
                                      fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(f"C{r + 1}", CellIsRule(operator="equal", formula=['"확인 필요"'],
                                  fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(color="9C5700", bold=True)))
    ws.conditional_formatting.add(f"D{first}:D{last}", CellIsRule(operator="equal", formula=['"확인"'],
                                  fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(color="9C5700", bold=True)))
    ws.cell(r + 3, 2, "※ 매크로 최초 1회 설치: Alt+F11 → 파일 → 파일 가져오기 → report_macro.bas → 닫기 → "
                      "'Excel 매크로 사용 통합 문서(*.xlsm)'로 다른 이름 저장. 이후 Alt+F8로 실행.").font = F_NOTE
    ws.cell(r + 4, 2, "※ 손해율 = (지급보험금 + 지급비금변동분) ÷ 위험보험료.  증감(p) = 당기 손해율 − 비교기간 손해율.  "
                      "비교기간 = 계약 구간과 마감월을 모두 1개월 앞당긴 동기간 (예: UY26.1~9 ↔ UY25.12~26.8, 202608 마감).").font = F_NOTE


# ---------------------------------------------------------------------------
# 담보별 집계 (마스터 기준, 최대 MAX_CODES 행)
# ---------------------------------------------------------------------------
def sheet_agg(wb):
    ws = wb.create_sheet(AGG)
    blk = ["위험보험료", "손해액", "손해율", "비교 위험보험료", "비교 손해액", "비교 손해율",
           "위험보험료 증감률", "손해율 증감(p)"]
    heads = (["담보코드", "대급부명", "담보명"] + [f"UY누적 {h}" for h in blk] + [f"UY당월 {h}" for h in blk]
             + ["누적 보험료순위", "누적 손해율키", "누적 손해율순위", "당월 보험료순위", "당월 손해율키", "당월 손해율순위",
                "", "대급부 첫등장", "대급부 UY보험료", "대급부 순위", "로데이터 행수", "중복"])
    hdr_row(ws, 1, 1, heads)
    widths(ws, [9, 10, 24] + [12] * 28)
    ws.row_dimensions[1].height = 30
    fmts = [NF_AMT, NF_AMT, NF_PCT, NF_AMT, NF_AMT, NF_PCT, NF_PCT_S, NF_PT]
    for r in range(2, LAST + 1):
        ws.cell(r, 1, f'=IF({MST}!$A{r}="","",{MST}!$A{r})')
        ws.cell(r, 2, f'=IF($A{r}="","",{MST}!$C{r}&"")')
        ws.cell(r, 3, f'=IF($A{r}="","",{MST}!$B{r}&"")')
        code = f",{R}$G:$G,$A{r}"
        for start, win in ((4, "C"), (12, "M")):
            L = [get_column_letter(start + k) for k in range(8)]
            f = [
                f'=IF($A{r}="","",{prem_f(win, False, code)})',
                f'=IF($A{r}="","",{loss_f(win, False, code)})',
                f'=IF(N({L[0]}{r})=0,"",{L[1]}{r}/{L[0]}{r})',
                f'=IF($A{r}="","",{prem_f(win, True, code)})',
                f'=IF($A{r}="","",{loss_f(win, True, code)})',
                f'=IF(N({L[3]}{r})=0,"",{L[4]}{r}/{L[3]}{r})',
                f'=IF(N({L[3]}{r})=0,"",{L[0]}{r}/{L[3]}{r}-1)',
                f'=IF(OR({L[2]}{r}="",{L[5]}{r}=""),"",{L[2]}{r}-{L[5]}{r})',
            ]
            for k in range(8):
                body(ws.cell(r, start + k, f[k]), fmts[k], RIGHT)
        # 순위 (동점은 위쪽 행 우선 → 고유 순위)
        for base_col, p, lr in ((20, "D", "F"), (23, "L", "N")):
            key = get_column_letter(base_col + 1)
            ws.cell(r, base_col, f'=IF(N({p}{r})>0,RANK({p}{r},{p}$2:{p}${LAST})+COUNTIF({p}$2:{p}{r},{p}{r})-1,"")')
            ws.cell(r, base_col + 1, f"=IF(N({p}{r})>0,IF({p}{r}>=LR_MIN_SHARE*SUM({p}$2:{p}${LAST}),{lr}{r},-1),-1)")
            ws.cell(r, base_col + 2,
                    f'=IF({key}{r}<0,"",RANK({key}{r},{key}$2:{key}${LAST})+COUNTIF({key}$2:{key}{r},{key}{r})-1)')
        # 대급부 목록 자동 추출 (UY 보험료 큰 순)
        ws.cell(r, 27, f'=IF($A{r}="","",IF(COUNTIF($B$2:$B{r},$B{r})=1,1,""))')
        ws.cell(r, 28, f'=IF(AA{r}=1,SUMIF($B$2:$B${LAST},$B{r},$D$2:$D${LAST}),"")')
        ws.cell(r, 29, f'=IF(AA{r}=1,RANK(AB{r},$AB$2:$AB${LAST})+COUNTIF($AB$2:$AB{r},AB{r})-1,"")')
        ws.cell(r, 30, f'=IF($A{r}="","",COUNTIF({R}$G:$G,$A{r}))')
        ws.cell(r, 31, f'=IF($A{r}="","",IF(COUNTIF($A$2:$A${LAST},$A{r})>1,1,0))')
    ws.freeze_panes = "D2"


# ---------------------------------------------------------------------------
# 요약 시트 (모든 표 7열 A~G)
# ---------------------------------------------------------------------------
def sheet_summary(wb):
    ws = wb.create_sheet("요약")
    widths(ws, [22, 15, 15, 11, 13, 15, 14])
    ws.sheet_view.showGridLines = False
    ws["A1"] = '=TEXT(BASE_YM,"0000\\.00")&" 기준 손해율 결산 리포트"'
    ws["A1"].font = F_TITLE
    ws["A2"] = ("금액 단위: 백만원  |  손해율 = (지급보험금 + 지급비금변동분) ÷ 위험보험료  |  "
                "증감 = 당기 − 비교기간(1개월 전 동기간), 빨강 악화 / 파랑 개선")
    ws["A2"].font = F_NOTE
    head = ["위험보험료", "손해액", "손해율", "비교기간 손해율", "위험보험료 증감률", "손해율 증감(p)"]

    def metrics(r, prem, loss, prem_py, loss_py, blank_if=None):
        z = r % 2 == 1
        g = (lambda f: f'=IF({blank_if},"",{f})') if blank_if else (lambda f: f"={f}")
        vals = [g(prem), g(loss), f'=IF(N(B{r})=0,"",C{r}/B{r})',
                g(f'IF(N({prem_py})=0,"",{loss_py}/{prem_py})'),
                g(f'IF(N({prem_py})=0,"",B{r}/{prem_py}-1)'),
                f'=IF(OR(D{r}="",E{r}=""),"",D{r}-E{r})']
        for k, (v, fmt) in enumerate(zip(vals, [NF_AMT, NF_AMT, NF_PCT, NF_PCT, NF_PCT_S, NF_PT])):
            body(ws.cell(r, 2 + k, v), fmt, RIGHT, zebra=z)

    def label(r, text, align=LEFT):
        body(ws.cell(r, 1, text), align=align, zebra=(r % 2 == 1))

    r = 4
    ws.cell(r, 1, "1. 핵심 요약 (전사)").font = F_SUB
    r += 1
    hdr_row(ws, r, 1, ["구분"] + head)
    top = r
    for lab, win in (("=LABEL_C", "C"), ("=LABEL_M", "M"), ("전기간 누적 (출시 이후)", "A")):
        r += 1
        label(r, lab, CENTER)
        metrics(r, prem_f(win, False), loss_f(win, False), prem_f(win, True), loss_f(win, True))
    name(wb, "RPT_SUMMARY", f"요약!$A${top}:$G${r}")
    r += 1
    ws.cell(r, 1, '="※ 비교기간: "&LABEL_CMP&" 계약, "&TEXT(CMP_CLOSE,"0000\\.00")&" 마감 기준 (누적·당월 각각).  '
                  '전기간 비교: 출시~"&TEXT(CMP_END,"0000\\.00")&" 계약, "&TEXT(CMP_CLOSE,"0000\\.00")&" 마감 누적"').font = F_NOTE
    name(wb, "NOTE_CMP", f"요약!$A${r}")

    for no, win in ((2, "C"), (3, "M")):
        r += 2
        ws.cell(r, 1, f'="{no}. "&{LABEL[win]}&" 손해율 (비교: "&LABEL_CMP&")"').font = F_SUB
        # 심사유형별
        r += 1
        ws.cell(r, 1, f"{no}-1. 심사유형별").font = Font(name=FONT, size=10, bold=True, color=NAVY)
        r += 1
        hdr_row(ws, r, 1, ["심사유형"] + head)
        top = r
        for uw in ("일반", "간편"):
            r += 1
            label(r, uw)
            ex = f',{R}$F:$F,$A{r}'
            metrics(r, prem_f(win, False, ex), loss_f(win, False, ex), prem_f(win, True, ex), loss_f(win, True, ex))
        name(wb, f"RPT_{win}_UW", f"요약!$A${top}:$G${r}")
        # 대급부별 (마스터에서 자동 추출, 빈 칸은 워드 출력 시 자동 삭제)
        r += 2
        ws.cell(r, 1, f"{no}-2. 대급부별").font = Font(name=FONT, size=10, bold=True, color=NAVY)
        r += 1
        hdr_row(ws, r, 1, ["대급부명"] + head)
        top = r
        p0, p1 = ("D", "G") if win == "C" else ("L", "O")      # 당기/비교 보험료 열
        l0, l1 = ("E", "H") if win == "C" else ("M", "P")
        for k in range(1, MAX_BENEFITS + 1):
            r += 1
            label(r, f'=IFERROR(INDEX({A_}$B$2:$B${LAST},MATCH({k},{A_}$AC$2:$AC${LAST},0)),"")')
            s = lambda col: f'SUMIF({A_}$B$2:$B${LAST},$A{r},{A_}${col}$2:${col}${LAST})'
            metrics(r, s(p0), s(l0), s(p1), s(l1), blank_if=f'$A{r}=""')
        # 마스터 미매핑분 = 전사 - 담보별 합계 (없으면 빈 행)
        r += 1
        tot = lambda col: f"SUM({A_}${col}$2:${col}${LAST})"
        um = [f"({prem_f(win, False)}-{tot(p0)})", f"({loss_f(win, False)}-{tot(l0)})",
              f"({prem_f(win, True)}-{tot(p1)})", f"({loss_f(win, True)}-{tot(l1)})"]
        label(r, f'=IF(AND(ROUND({um[0]},0)=0,ROUND({um[2]},0)=0),"","(미매핑)")')
        metrics(r, *um, blank_if=f'$A{r}=""')
        name(wb, f"RPT_{win}_BEN", f"요약!$A${top}:$G${r}")


# ---------------------------------------------------------------------------
# Top20 시트
# ---------------------------------------------------------------------------
def sheet_top(wb):
    ws = wb.create_sheet("Top20")
    widths(ws, [6, 10, 11, 28, 13, 14, 11, 14, 3, 6, 4])
    ws.column_dimensions["J"].hidden = True    # MATCH 결과 (집계 행 위치)
    ws.column_dimensions["K"].hidden = True    # 순위 번호
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A1"].font = "담보별 Top 20  (금액 단위: 백만원)", F_TITLE
    blocks = [  # 창, 종류, 순위열, 보험료, 증감률, 손해율, 증감p, 이름
        ("C", "위험보험료", "T", "D", "J", "F", "K", "RPT_C_TOPP"),
        ("C", "손해율", "V", "D", "J", "F", "K", "RPT_C_TOPL"),
        ("M", "위험보험료", "W", "L", "R", "N", "S", "RPT_M_TOPP"),
        ("M", "손해율", "Y", "L", "R", "N", "S", "RPT_M_TOPL"),
    ]
    r = 3
    for win, kind, rk, prem, pchg, lr, lrchg, nm in blocks:
        ws.cell(r, 1, f'={LABEL[win]}&" · {kind} 상위 20 담보 (비교: "&LABEL_CMP&")"').font = F_SUB
        r += 1
        hdr_row(ws, r, 1, ["순위", "담보코드", "대급부명", "담보명", "위험보험료", "위험보험료 증감률",
                           "손해율(%)", "동기간 대비 증감(p)"])
        top = r
        for n in range(1, TOP_N + 1):
            r += 1
            z = r % 2 == 1
            ws.cell(r, 11, n)
            ws.cell(r, 10, f'=IFERROR(MATCH(K{r},{A_}${rk}$2:${rk}${LAST},0),"")')
            body(ws.cell(r, 1, f'=IF($J{r}="","",K{r})'), align=CENTER, zebra=z)
            for c, col, fmt, al in ((2, "A", None, LEFT), (3, "B", None, LEFT), (4, "C", None, LEFT),
                                    (5, prem, NF_AMT, RIGHT), (6, pchg, NF_PCT_S, RIGHT),
                                    (7, lr, NF_PCT, RIGHT), (8, lrchg, NF_PT, RIGHT)):
                body(ws.cell(r, c, f'=IF($J{r}="","",INDEX({A_}${col}$2:${col}${LAST},$J{r}))'), fmt, al, zebra=z)
        name(wb, nm, f"Top20!$A${top}:$H${r}")
        r += 1
        if kind == "손해율":
            ws.cell(r, 1, '="※ 위험보험료 비중 "&TEXT(LR_MIN_SHARE,"0.00%")&" 이상 담보 대상"').font = F_NOTE
        r += 2


# ---------------------------------------------------------------------------
RAW_COLS = ["계약년월", "마감년월", "주상품코드", "특약상품코드", "경과개월수", "심사유형(일반.간편)",
            "담보코드", "위험보험료 합계", "지급보험금합계", "지급비금변동분 합계"]


def sheet_raw(wb, raw: pd.DataFrame):
    ws = wb.create_sheet(RAW)
    hdr_row(ws, 1, 1, RAW_COLS)
    widths(ws, [9, 9, 10, 11, 8, 12, 10, 14, 13, 14])
    for r, rec in enumerate(raw.itertuples(index=False), start=2):
        for j, v in enumerate(rec, start=1):
            ws.cell(r, j, v)
    for col in "CDG":                                  # 코드 열은 텍스트 (앞자리 0 보존)
        for (c,) in ws.iter_rows(min_row=2, min_col=ord(col) - 64, max_col=ord(col) - 64):
            c.number_format = "@"
    ws.freeze_panes = "A2"


def sheet_master(wb, master: pd.DataFrame):
    ws = wb.create_sheet(MST)
    hdr_row(ws, 1, 1, ["담보코드", "담보명", "대급부명"])
    widths(ws, [10, 28, 12])
    for r, rec in enumerate(master.itertuples(index=False), start=2):
        for j, v in enumerate(rec, start=1):
            ws.cell(r, j, v)
        ws.cell(r, 1).number_format = "@"
    ws.freeze_panes = "A2"


def main():
    raw = pd.read_csv(BASE / "data" / "raw_data.csv",
                      dtype={"주상품코드": str, "특약상품코드": str, "담보코드": str})
    master = pd.read_csv(BASE / "data" / "code_master.csv", dtype=str)[["담보코드", "담보명", "대급부명"]]
    wb = Workbook()
    sheet_run(wb)
    sheet_summary(wb)
    sheet_top(wb)
    sheet_raw(wb, raw[RAW_COLS])
    sheet_master(wb, master)
    sheet_agg(wb)
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"[OK] {OUT}")


if __name__ == "__main__":
    main()
