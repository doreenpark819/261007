Attribute VB_Name = "LossRatioReport"
Option Explicit
' =====================================================================
'  손해율 집계 템플릿 매크로 (Word 참조 설정 불필요)
'   LoadRawData    : 로데이터 파일 선택 → raw_data 시트 교체 + 형식 자동 정리
'   MakeWordReport : 형식 정리 → 데이터 점검 확인 → 워드(.docx) 리포트 생성
' =====================================================================

Private Const RAW_SHEET As String = "raw_data"
Private Const N_COLS As Long = 10

' ---------------------------------------------------------------------
' 1) 로데이터 불러오기
' ---------------------------------------------------------------------
Public Sub LoadRawData()
    Dim fn As Variant, src As Workbook, data As Variant
    Dim ws As Worksheet, nRows As Long

    fn = Application.GetOpenFilename("로데이터 (*.xlsx;*.xlsb;*.xls;*.csv;*.txt),*.xlsx;*.xlsb;*.xls;*.csv;*.txt", , _
                                     "로데이터 파일 선택 (1행 헤더, A~J열 10개 컬럼)")
    If VarType(fn) = vbBoolean Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    If LCase(Right(fn, 4)) = ".csv" Or LCase(Right(fn, 4)) = ".txt" Then
        ' 코드 열(C,D,G)은 텍스트로 읽어 앞자리 0 보존
        Workbooks.OpenText Filename:=fn, DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
            Comma:=(LCase(Right(fn, 4)) = ".csv"), Tab:=(LCase(Right(fn, 4)) = ".txt"), Local:=True, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 2), Array(5, 1), _
                             Array(6, 1), Array(7, 2), Array(8, 1), Array(9, 1), Array(10, 1))
        Set src = ActiveWorkbook
    Else
        Set src = Workbooks.Open(Filename:=fn, ReadOnly:=True, UpdateLinks:=0)
    End If

    With src.Worksheets(1)
        nRows = .Cells(.Rows.Count, 7).End(xlUp).Row          ' G열(담보코드) 기준 마지막 행
        data = .Range("A1").Resize(nRows, N_COLS).Value
    End With
    src.Close SaveChanges:=False

    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    Application.Calculation = xlCalculationManual
    ws.Cells.ClearContents
    ws.Range("C:D,G:G").NumberFormat = "@"
    ws.Range("A1").Resize(nRows, N_COLS).Value = data
    Erase data

    CleanRawData
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ThisWorkbook.Worksheets("실행").Activate
    MsgBox Format(nRows - 1, "#,##0") & "행을 불러왔습니다." & vbCrLf & _
           "데이터 점검 종합 상태: " & ThisWorkbook.Names("CHECK_STATUS").RefersToRange.Value, vbInformation
    Exit Sub
Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "불러오기 실패: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------
' raw_data 형식 정리: 년월·금액 텍스트 → 숫자, 코드·심사유형 공백 제거
' ---------------------------------------------------------------------
Private Sub CleanRawData()
    Dim ws As Worksheet, nRows As Long, c As Variant
    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    nRows = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If nRows < 2 Then Exit Sub

    For Each c In Array(1, 2)                 ' 계약년월, 마감년월 → YYYYMM 숫자
        FixColumn ws.Range(ws.Cells(2, c), ws.Cells(nRows, c)), True
    Next c
    For Each c In Array(5, 8, 9, 10)          ' 경과개월수, 금액 → 숫자 (빈칸은 0)
        FixColumn ws.Range(ws.Cells(2, c), ws.Cells(nRows, c)), False
    Next c
    For Each c In Array(3, 4, 6, 7)           ' 코드·심사유형 → 앞뒤 공백 제거
        TrimColumn ws.Range(ws.Cells(2, c), ws.Cells(nRows, c))
    Next c
End Sub

Private Sub FixColumn(rng As Range, isYM As Boolean)
    Dim v As Variant, i As Long, s As String, k As Long, d As String
    v = rng.Value
    If Not IsArray(v) Then v = OneCell(v)                   ' 데이터가 1행뿐인 경우
    For i = 1 To UBound(v, 1)
        If IsDate(v(i, 1)) And VarType(v(i, 1)) = vbDate Then
            v(i, 1) = Year(v(i, 1)) * 100 + Month(v(i, 1))   ' 날짜로 바뀐 년월
        ElseIf VarType(v(i, 1)) = vbString Then
            s = Trim(Replace(v(i, 1), ",", ""))
            If isYM Then                                       ' 2025.06 / 2025-06 / '202506 → 202506
                d = ""
                For k = 1 To Len(s)
                    If Mid(s, k, 1) Like "#" Then d = d & Mid(s, k, 1)
                Next k
                If Len(d) = 5 Then d = Left(d, 4) & "0" & Right(d, 1)   ' 2025.6 → 202506
                If Len(d) = 6 Then v(i, 1) = CLng(d)
            ElseIf s = "" Or s = "-" Then
                v(i, 1) = 0
            ElseIf IsNumeric(s) Then
                v(i, 1) = CDbl(s)
            End If
        ElseIf IsEmpty(v(i, 1)) And Not isYM Then
            v(i, 1) = 0
        End If
    Next i
    rng.NumberFormat = "General"
    rng.Value = v
End Sub

Private Function OneCell(x As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    a(1, 1) = x
    OneCell = a
End Function

Private Sub TrimColumn(rng As Range)
    Dim v As Variant, i As Long
    v = rng.Value
    If Not IsArray(v) Then v = OneCell(v)                   ' 데이터가 1행뿐인 경우
    For i = 1 To UBound(v, 1)
        If VarType(v(i, 1)) = vbString Then v(i, 1) = Trim(v(i, 1))
    Next i
    rng.Value = v
End Sub

' ---------------------------------------------------------------------
' 2) 워드 리포트 생성
' ---------------------------------------------------------------------
Public Sub MakeWordReport()
    Dim wdApp As Object, doc As Object
    Dim baseTxt As String, outPath As String, status As String
    Dim blocks As Variant, i As Long, ym As Long
    Dim lc As String, lm As String, cmp As String

    Application.ScreenUpdating = False
    CleanRawData
    Application.CalculateFull
    Application.ScreenUpdating = True

    status = ThisWorkbook.Names("CHECK_STATUS").RefersToRange.Value
    If status = "오류" Then
        ThisWorkbook.Worksheets("실행").Activate
        MsgBox "데이터 점검에서 '오류'가 있습니다. 실행 시트의 ③ 데이터 점검을 확인하세요.", vbExclamation
        Exit Sub
    ElseIf status <> "정상" Then
        If MsgBox("데이터 점검에 '확인' 항목이 있습니다. 그래도 보고서를 만들까요?", _
                  vbQuestion + vbYesNo) = vbNo Then
            ThisWorkbook.Worksheets("실행").Activate
            Exit Sub
        End If
    End If

    ym = ThisWorkbook.Names("BASE_YM").RefersToRange.Value
    baseTxt = Format(ym \ 100, "0000") & "." & Format(ym Mod 100, "00")

    lc = NameValue("LABEL_C")       ' 예: UY26.1~9 누적
    lm = NameValue("LABEL_M")       ' 예: UY26.9 당월
    cmp = NameValue("LABEL_CMP")    ' 예: UY25.12~26.8

    ' 제목, 표 범위 이름, 앞에서 페이지 나눔
    blocks = Array( _
        Array("1. 핵심 요약 (전사)", "RPT_SUMMARY", False), _
        Array("2-1. " & lc & " · 심사유형별 (비교: " & cmp & ")", "RPT_C_UW", False), _
        Array("2-2. " & lc & " · 대급부별 (비교: " & cmp & ")", "RPT_C_BEN", False), _
        Array("2-3. " & lc & " · 위험보험료 상위 20 담보", "RPT_C_TOPP", True), _
        Array("2-4. " & lc & " · 손해율 상위 20 담보", "RPT_C_TOPL", True), _
        Array("3-1. " & lm & " · 심사유형별 (비교: " & cmp & ")", "RPT_M_UW", True), _
        Array("3-2. " & lm & " · 대급부별 (비교: " & cmp & ")", "RPT_M_BEN", False), _
        Array("3-3. " & lm & " · 위험보험료 상위 20 담보", "RPT_M_TOPP", True), _
        Array("3-4. " & lm & " · 손해율 상위 20 담보", "RPT_M_TOPL", True))

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

    AddText doc, baseTxt & " 기준 손해율 결산 리포트", 20, True, 1, False
    AddText doc, "마감 기준월 " & baseTxt & "  |  작성일 " & Format(Date, "yyyy.mm.dd") & _
                 "  |  금액 단위 백만원", 9, False, 1, False
    AddText doc, lc & "  /  " & lm & "  /  비교기간 " & cmp, 9, False, 1, False
    AddText doc, "※ 손해율 = (지급보험금 + 지급비금변동분) ÷ 위험보험료,  증감(p) = 당기 손해율 - 비교기간 손해율 (빨강: 악화 / 파랑: 개선)", _
            8, False, 0, False
    AddText doc, NameValue("NOTE_CMP"), 8, False, 0, False

    For i = LBound(blocks) To UBound(blocks)
        AddText doc, CStr(blocks(i)(0)), 12, True, 0, CBool(blocks(i)(2))
        PasteRange doc, ThisWorkbook.Names(CStr(blocks(i)(1))).RefersToRange
    Next i
    Application.CutCopyMode = False

    outPath = ThisWorkbook.Path & Application.PathSeparator & "손해율_결산리포트_" & Format(ym, "0") & ".docx"
    doc.SaveAs2 outPath, 16                                ' wdFormatDocumentDefault
    MsgBox "워드 리포트 생성 완료:" & vbCrLf & outPath, vbInformation
    Exit Sub
Fail:
    Application.CutCopyMode = False
    MsgBox "워드 리포트 생성 중 오류: " & Err.Description & vbCrLf & _
           "(같은 이름의 docx가 열려 있으면 닫고 다시 실행하세요)", vbCritical
End Sub

Private Function NameValue(nm As String) As String
    NameValue = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
End Function

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

' 엑셀 범위를 서식 그대로 워드 표로 붙여넣고, 첫 칸이 빈 행은 삭제
Private Sub PasteRange(doc As Object, src As Range)
    Dim r As Object, t As Object, i As Long, s As String
    src.Copy
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.PasteExcelTable False, False, False
    Set t = doc.Tables(doc.Tables.Count)
    For i = t.Rows.Count To 2 Step -1
        s = t.Cell(i, 1).Range.Text
        s = Replace(Replace(s, Chr(13), ""), Chr(7), "")
        If Len(Trim(s)) = 0 Then t.Rows(i).Delete
    Next i
    t.Rows.AllowBreakAcrossPages = False
    t.Rows(1).HeadingFormat = True
    t.AutoFitBehavior 2                                     ' 페이지 폭에 맞춤
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertParagraphAfter
    ResetLastParagraph doc
End Sub

' 마지막 빈 문단이 제목 서식(페이지 나눔 등)을 물려받지 않도록 초기화
Private Sub ResetLastParagraph(doc As Object)
    With doc.Paragraphs(doc.Paragraphs.Count).Format
        .PageBreakBefore = False
        .KeepWithNext = False
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub
