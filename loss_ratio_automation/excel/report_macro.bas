Attribute VB_Name = "LossRatioReport"
Option Explicit
' =====================================================================
'  손해율 집계 템플릿 매크로
'   - PrepareRawData : raw_data K~S열 수식을 데이터 마지막 행까지 채움
'   - MakeWordReport : 요약/Top20 표를 워드(.docx)로 출력
'  Word 참조 설정 불필요 (Late Binding)
' =====================================================================

Private Const RAW_SHEET As String = "raw_data"

Public Sub PrepareRawData()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "raw_data 시트에 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 이전 달보다 행이 줄어든 경우 남은 수식 정리
    ws.Range("K" & lastRow + 1 & ":S" & ws.Rows.Count).ClearContents

    ws.Range("K2:K" & lastRow).Formula = "=IFERROR(VLOOKUP($G2,code_master!$A:$C,2,FALSE),""(미매핑)"")"
    ws.Range("L2:L" & lastRow).Formula = "=IFERROR(VLOOKUP($G2,code_master!$A:$C,3,FALSE),""(미매핑)"")"
    ws.Range("M2:M" & lastRow).Formula = "=I2+J2"
    ws.Range("N2:N" & lastRow).Formula = "=MOD(ROUND(VALUE(A2)*100,0),10000)"
    ws.Range("O2:O" & lastRow).Formula = "=MOD(ROUND(VALUE(B2)*100,0),10000)"
    ws.Range("P2:P" & lastRow).Formula = "=IF(AND(N2>=UY_START,N2<=BASE_YM,O2<=BASE_YM),1,0)"
    ws.Range("Q2:Q" & lastRow).Formula = "=IF(AND(N2>=UY_START_PY,N2<=BASE_YM_PY,O2<=BASE_YM_PY),1,0)"
    ws.Range("R2:R" & lastRow).Formula = "=IF(O2=BASE_YM,1,0)"
    ws.Range("S2:S" & lastRow).Formula = "=IF(O2=BASE_YM_PY,1,0)"

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.ScreenUpdating = True

    Dim unmapped As Long
    unmapped = Application.WorksheetFunction.CountIf(ws.Range("K2:K" & lastRow), "(미매핑)")
    MsgBox "수식 채우기 완료: " & Format(lastRow - 1, "#,##0") & "행" & vbCrLf & _
           "마스터 미매핑 행: " & Format(unmapped, "#,##0"), vbInformation
End Sub

Public Sub MakeWordReport()
    Dim wdApp As Object, doc As Object
    Dim baseTxt As String, outPath As String
    Dim blocks As Variant, i As Long

    Application.CalculateFull
    Dim ym As Long
    ym = ThisWorkbook.Worksheets("설정").Range("B7").Value      ' 기준월 YYMM (예: 2509)
    baseTxt = Format(ym \ 100, "00") & "." & Format(ym Mod 100, "00")

    ' 제목, 이름 정의(표 범위), 앞에서 페이지 나눔 여부
    blocks = Array( _
        Array("1. 핵심 요약 (전사)", "RPT_SUMMARY", False), _
        Array("2. UY 누적 손해율 - 심사유형별", "RPT_UY_UW", False), _
        Array("2. UY 누적 손해율 - 대급부별", "RPT_UY_BEN", False), _
        Array("2. UY 누적 - 위험보험료 상위 20 담보", "RPT_UY_TOPP", True), _
        Array("2. UY 누적 - 손해율 상위 20 담보", "RPT_UY_TOPL", True), _
        Array("3. CY 당월 손해율 - 심사유형별", "RPT_CY_UW", True), _
        Array("3. CY 당월 손해율 - 대급부별", "RPT_CY_BEN", False), _
        Array("3. CY 당월 - 위험보험료 상위 20 담보", "RPT_CY_TOPP", True), _
        Array("3. CY 당월 - 손해율 상위 20 담보", "RPT_CY_TOPL", True))

    On Error GoTo Fail
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set doc = wdApp.Documents.Add

    With doc.PageSetup
        .Orientation = 1                                   ' 가로
        .LeftMargin = wdApp.CentimetersToPoints(1.6)
        .RightMargin = wdApp.CentimetersToPoints(1.6)
        .TopMargin = wdApp.CentimetersToPoints(1.5)
        .BottomMargin = wdApp.CentimetersToPoints(1.5)
    End With

    AddText doc, "20" & baseTxt & " 기준 손해율 결산 리포트", 20, True, 1, False
    AddText doc, "마감 기준월 20" & baseTxt & "  |  작성일 " & Format(Date, "yyyy.mm.dd") & _
                 "  |  금액 단위 백만원", 9, False, 1, False
    AddText doc, "※ 손해율 = (지급보험금 + 지급비금변동분) ÷ 위험보험료,  증감(p) = 당해 손해율 - 전년 동기 손해율 (빨강: 악화 / 파랑: 개선)", _
            8, False, 0, False

    For i = LBound(blocks) To UBound(blocks)
        AddText doc, CStr(blocks(i)(0)), 12, True, 0, CBool(blocks(i)(2))
        PasteRange doc, ThisWorkbook.Names(CStr(blocks(i)(1))).RefersToRange
    Next i
    Application.CutCopyMode = False

    outPath = ThisWorkbook.Path & Application.PathSeparator & "손해율_결산리포트_" & baseTxt & ".docx"
    doc.SaveAs2 outPath, 16                                ' wdFormatDocumentDefault
    MsgBox "워드 리포트 생성 완료:" & vbCrLf & outPath, vbInformation
    Exit Sub
Fail:
    Application.CutCopyMode = False
    MsgBox "워드 리포트 생성 중 오류: " & Err.Description, vbCritical
End Sub

' 문서 끝에 문단 추가 (align: 0=왼쪽, 1=가운데)
Private Sub AddText(doc As Object, txt As String, sz As Single, isBold As Boolean, _
                    align As Long, pageBreak As Boolean)
    Dim r As Object
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertAfter txt
    r.Font.Name = "맑은 고딕"
    r.Font.Size = sz
    r.Font.Bold = isBold
    If sz >= 12 Then r.Font.Color = RGB(31, 56, 100) Else r.Font.Color = RGB(89, 89, 89)
    r.ParagraphFormat.Alignment = align
    r.ParagraphFormat.SpaceBefore = IIf(sz >= 12, 10, 0)
    r.ParagraphFormat.SpaceAfter = 4
    r.ParagraphFormat.PageBreakBefore = pageBreak
    r.ParagraphFormat.KeepWithNext = True
    r.InsertParagraphAfter
    ResetLastParagraph doc
End Sub

' 새로 생긴 마지막 빈 문단이 제목 서식(페이지 나눔 등)을 물려받지 않도록 초기화
Private Sub ResetLastParagraph(doc As Object)
    With doc.Paragraphs(doc.Paragraphs.Count).Format
        .PageBreakBefore = False
        .KeepWithNext = False
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub

' 엑셀 범위를 서식 그대로 워드 표로 붙여넣기
Private Sub PasteRange(doc As Object, src As Range)
    Dim r As Object
    src.Copy
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.PasteExcelTable False, False, False
    With doc.Tables(doc.Tables.Count)
        .Rows.AllowBreakAcrossPages = False
        .Rows(1).HeadingFormat = True
        .AutoFitBehavior 2                                 ' 페이지 폭에 맞춤
    End With
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertParagraphAfter
    ResetLastParagraph doc
End Sub
