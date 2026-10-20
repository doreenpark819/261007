Attribute VB_Name = "LossRatioReport"
Option Explicit
' =====================================================================
' =====================================================================

Private Const RAW_SHEET As String = "raw_data"
Private Const MST_SHEET As String = "code_master"
Private Const N_COLS As Long = 11
Private wd As Object

' Decode "\uXXXX" escapes (keeps this file ASCII-only so encoding never breaks)
Private Function T(ByVal s As String) As String
    Dim i As Long, o As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" Then
            o = o & ChrW(CLng("&H" & Mid$(s, i + 2, 4) & "&"))
            i = i + 6
        Else
            o = o & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    T = o
End Function

Private Function RPT_SHEET() As String
    RPT_SHEET = T("\uBCF4\uACE0\uC11C")
End Function

Private Function FONT_NAME() As String
    FONT_NAME = T("\uB9D1\uC740 \uACE0\uB515")
End Function


' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
Public Sub LoadRawData()
    Dim fn As Variant, src As Workbook, data As Variant
    Dim ws As Worksheet, nRows As Long, ext As String

    fn = Application.GetOpenFilename(T("\uB85C\uB370\uC774\uD130 (*.xlsx;*.xlsb;*.xls;*.csv;*.txt),*.xlsx;*.xlsb;*.xls;*.csv;*.txt"), , _
                                     T("\uB85C\uB370\uC774\uD130 \uD30C\uC77C \uC120\uD0DD (1\uD589 \uD5E4\uB354, A~J 10\uAC1C \uC5F4 + K\uC5F4 \uAC74\uC218)"))
    If VarType(fn) = vbBoolean Then Exit Sub
    ext = LCase(Right(fn, 4))

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    If ext = ".csv" Or ext = ".txt" Then
        Workbooks.OpenText Filename:=fn, DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
            Comma:=(ext = ".csv"), Tab:=(ext = ".txt"), Local:=True, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), _
                             Array(7, 2), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1))
        Set src = ActiveWorkbook
    Else
        Set src = Workbooks.Open(Filename:=fn, ReadOnly:=True, UpdateLinks:=0)
    End If

    With src.Worksheets(1)
        nRows = .Cells(.Rows.Count, 7).End(xlUp).Row
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

    ThisWorkbook.Worksheets(T("\uC2E4\uD589")).Activate
    MsgBox Format(nRows - 1, "#,##0") & T("\uD589\uC744 \uBD88\uB7EC\uC654\uC2B5\uB2C8\uB2E4.") & vbCrLf & _
           T("\uB370\uC774\uD130 \uC810\uAC80 \uC885\uD569 \uC0C1\uD0DC: ") & ThisWorkbook.Names("CHECK_STATUS").RefersToRange.Value, vbInformation
    Exit Sub
Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox T("\uBD88\uB7EC\uC624\uAE30 \uC2E4\uD328: ") & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
Public Sub LoadCodeMaster()
    Dim fn As Variant, src As Workbook, data As Variant
    Dim ws As Worksheet, nRows As Long, ext As String, i As Long, j As Long

    fn = Application.GetOpenFilename(T("\uB2F4\uBCF4 \uB9C8\uC2A4\uD130 (*.xlsx;*.xlsb;*.xls;*.csv),*.xlsx;*.xlsb;*.xls;*.csv"), , _
                                     T("\uB2F4\uBCF4 \uB9C8\uC2A4\uD130 \uD30C\uC77C \uC120\uD0DD (1\uD589 \uD5E4\uB354, A \uB2F4\uBCF4\uCF54\uB4DC / B \uB2F4\uBCF4\uBA85 / C \uB300\uAE09\uBD80\uBA85 / D \uAD6C\uBD84 / E \uACE0\uB2F9\uB300\uD1B5)"))
    If VarType(fn) = vbBoolean Then Exit Sub
    ext = LCase(Right(fn, 4))

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    If ext = ".csv" Then
        Workbooks.OpenText Filename:=fn, DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
            Comma:=True, Local:=True, _
            FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2))
        Set src = ActiveWorkbook
    Else
        Set src = Workbooks.Open(Filename:=fn, ReadOnly:=True, UpdateLinks:=0)
    End If

    With src.Worksheets(1)
        nRows = .Cells(.Rows.Count, 1).End(xlUp).Row
        data = .Range("A1").Resize(nRows, 5).Value
    End With
    src.Close SaveChanges:=False

    For i = 2 To nRows
        For j = 1 To 5
            If VarType(data(i, j)) = vbString Then data(i, j) = Trim(data(i, j))
        Next j
    Next i

    Set ws = ThisWorkbook.Worksheets(MST_SHEET)
    Application.Calculation = xlCalculationManual
    ws.Range("A2:E" & ws.Rows.Count).ClearContents
    ws.Range("A:A").NumberFormat = "@"
    If nRows >= 2 Then
        ws.Range("A2").Resize(nRows - 1, 5).Value = SubArray(data, 2, nRows, 5)
    End If
    Erase data

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ThisWorkbook.Worksheets(T("\uC2E4\uD589")).Activate
    MsgBox T("\uB2F4\uBCF4 \uB9C8\uC2A4\uD130 ") & Format(nRows - 1, "#,##0") & T("\uAC74\uC744 \uBD88\uB7EC\uC654\uC2B5\uB2C8\uB2E4.") & vbCrLf & _
           T("\uB370\uC774\uD130 \uC810\uAC80 \uC885\uD569 \uC0C1\uD0DC: ") & ThisWorkbook.Names("CHECK_STATUS").RefersToRange.Value, vbInformation
    Exit Sub
Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox T("\uB9C8\uC2A4\uD130 \uBD88\uB7EC\uC624\uAE30 \uC2E4\uD328: ") & Err.Description, vbCritical
End Sub

Private Function SubArray(a As Variant, r1 As Long, r2 As Long, nC As Long) As Variant
    Dim b() As Variant, i As Long, j As Long
    ReDim b(1 To r2 - r1 + 1, 1 To nC)
    For i = r1 To r2
        For j = 1 To nC
            b(i - r1 + 1, j) = a(i, j)
        Next j
    Next i
    SubArray = b
End Function

' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
Private Sub CleanRawData()
    Dim ws As Worksheet, nRows As Long, c As Variant
    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    nRows = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If nRows < 2 Then Exit Sub

    For Each c In Array(1, 2)
        FixColumn ws.Range(ws.Cells(2, c), ws.Cells(nRows, c)), True, False
    Next c
    For Each c In Array(5, 8, 9, 10)
        FixColumn ws.Range(ws.Cells(2, c), ws.Cells(nRows, c)), False, True
    Next c
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(2, 11), ws.Cells(nRows, 11))) > 0 Then
        FixColumn ws.Range(ws.Cells(2, 11), ws.Cells(nRows, 11)), False, True
    End If
    For Each c In Array(3, 4, 6, 7)
        TrimColumn ws.Range(ws.Cells(2, c), ws.Cells(nRows, c))
    Next c
End Sub

Private Sub FixColumn(rng As Range, isYM As Boolean, blankToZero As Boolean)
    Dim v As Variant, i As Long, s As String, k As Long, d As String
    v = rng.Value
    If Not IsArray(v) Then v = OneCell(v)
    For i = 1 To UBound(v, 1)
        If VarType(v(i, 1)) = vbDate Then
            v(i, 1) = Year(v(i, 1)) * 100 + Month(v(i, 1))
        ElseIf VarType(v(i, 1)) = vbString Then
            s = Trim(Replace(v(i, 1), ",", ""))
            If isYM Then
                d = ""
                For k = 1 To Len(s)
                    If Mid(s, k, 1) Like "#" Then d = d & Mid(s, k, 1)
                Next k
                If Len(d) = 5 Then d = Left(d, 4) & "0" & Right(d, 1)
                If Len(d) = 6 Then v(i, 1) = CLng(d)
            ElseIf s = "" Or s = "-" Then
                v(i, 1) = 0
            ElseIf IsNumeric(s) Then
                v(i, 1) = CDbl(s)
            End If
        ElseIf IsEmpty(v(i, 1)) And blankToZero Then
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
    If Not IsArray(v) Then v = OneCell(v)
    For i = 1 To UBound(v, 1)
        If VarType(v(i, 1)) = vbString Then v(i, 1) = Trim(v(i, 1))
    Next i
    rng.Value = v
End Sub

' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
Public Sub MakeWordReport()
    Dim doc As Object, ws As Worksheet
    Dim status As String, outPath As String, tag As String, txt As String
    Dim parts As Variant, r As Long, lastRow As Long, pb As Boolean, st As Long

    Application.ScreenUpdating = False
    CleanRawData
    Application.CalculateFull

    status = ThisWorkbook.Names("CHECK_STATUS").RefersToRange.Value
    If status = T("\uC624\uB958") Then
        Application.ScreenUpdating = True
        ThisWorkbook.Worksheets(T("\uC2E4\uD589")).Activate
        MsgBox T("\uB370\uC774\uD130 \uC810\uAC80\uC5D0 '\uC624\uB958'\uAC00 \uC788\uC2B5\uB2C8\uB2E4. \uC2E4\uD589 \uC2DC\uD2B8\uC758 \u2463 \uB370\uC774\uD130 \uC810\uAC80\uC744 \uD655\uC778\uD558\uC138\uC694."), vbExclamation
        Exit Sub
    ElseIf status <> T("\uC815\uC0C1") Then
        Application.ScreenUpdating = True
        If MsgBox(T("\uB370\uC774\uD130 \uC810\uAC80\uC5D0 '\uD655\uC778' \uD56D\uBAA9\uC774 \uC788\uC2B5\uB2C8\uB2E4. \uADF8\uB798\uB3C4 \uBCF4\uACE0\uC11C\uB97C \uB9CC\uB4E4\uAE4C\uC694?"), vbQuestion + vbYesNo) = vbNo Then
            ThisWorkbook.Worksheets(T("\uC2E4\uD589")).Activate
            Exit Sub
        End If
        Application.ScreenUpdating = False
    End If

    On Error GoTo Fail
    Set wd = CreateObject("Word.Application")
    Set doc = wd.Documents.Add
    With doc.PageSetup
        .Orientation = 0
        .PageWidth = wd.CentimetersToPoints(21)
        .PageHeight = wd.CentimetersToPoints(29.7)
        .LeftMargin = wd.CentimetersToPoints(2)
        .RightMargin = wd.CentimetersToPoints(2)
        .TopMargin = wd.CentimetersToPoints(2)
        .BottomMargin = wd.CentimetersToPoints(2)
    End With

    Set ws = ThisWorkbook.Worksheets(RPT_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastRow
        tag = Trim(CStr(ws.Cells(r, 1).Value))
        If tag = "END" Then Exit For
        If tag <> "" Then
            parts = Split(tag, ",")
            Select Case parts(0)
                Case "H1", "H2", "BUL", "CAP", "UNIT", "NOTE"
                    AddPara doc, CellText(ws.Cells(r, 2)), CStr(parts(0)), pb
                    pb = False
                Case "LEG"
                    txt = CellText(ws.Cells(r, 2))
                    st = AddPara(doc, txt, "LEG", pb)
                    pb = False
                    ColorLegend doc, st, txt
                Case "PB"
                    pb = True
                Case "TBL"
                    PutTable doc, ws, r, parts
            End Select
        End If
    Next r

    Application.ScreenUpdating = True
    wd.Visible = True
    outPath = ThisWorkbook.Path & Application.PathSeparator & T("\uC190\uD574\uC728_\uBCF4\uACE0\uC11C_") & _
              Format(ThisWorkbook.Names("BASE_YM").RefersToRange.Value, "0") & ".docx"
    doc.SaveAs2 outPath, 16                                ' wdFormatDocumentDefault
    MsgBox T("\uC6CC\uB4DC \uBCF4\uACE0\uC11C \uC0DD\uC131 \uC644\uB8CC:") & vbCrLf & outPath, vbInformation
    Exit Sub
Fail:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not wd Is Nothing Then wd.Visible = True
    MsgBox T("\uC6CC\uB4DC \uBCF4\uACE0\uC11C \uC0DD\uC131 \uC911 \uC624\uB958 (") & r & T("\uD589 \uCC98\uB9AC \uC911): ") & Err.Description & vbCrLf & _
           T("(\uAC19\uC740 \uC774\uB984\uC758 docx\uAC00 \uC5F4\uB824 \uC788\uC73C\uBA74 \uB2EB\uACE0 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694)"), vbCritical
End Sub

Private Function CellText(c As Range) As String
    Dim v As Variant
    v = c.Value2
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        CellText = v
        Exit Function
    End If
    On Error Resume Next
    CellText = Application.WorksheetFunction.Text(v, c.NumberFormat)
    If Err.Number <> 0 Then CellText = CStr(v)
    On Error GoTo 0
End Function

Private Function AddPara(doc As Object, txt As String, tag As String, pb As Boolean) As Long
    Dim r As Object
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    AddPara = r.Start
    r.InsertAfter txt
    r.Font.Name = FONT_NAME
    r.Font.Bold = False
    r.Font.Color = 0
    r.Font.Size = 10
    With r.ParagraphFormat
        .Alignment = 0
        .LeftIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 2
        .PageBreakBefore = pb
        .KeepWithNext = True
    End With
    Select Case tag
        Case "H1"
            r.Font.Size = 15: r.Font.Bold = True
            r.ParagraphFormat.SpaceBefore = 6: r.ParagraphFormat.SpaceAfter = 6
        Case "H2"
            r.Font.Size = 13: r.Font.Bold = True
            r.ParagraphFormat.LeftIndent = wd.CentimetersToPoints(0.4)
            r.ParagraphFormat.SpaceAfter = 6
        Case "BUL"
            r.Font.Size = 11.5
            r.ParagraphFormat.LeftIndent = wd.CentimetersToPoints(0.8)
            r.ParagraphFormat.SpaceBefore = 6: r.ParagraphFormat.SpaceAfter = 6
        Case "CAP"
            r.Font.Size = 11.5: r.Font.Bold = True
            r.ParagraphFormat.Alignment = 1
            r.ParagraphFormat.SpaceBefore = 10
        Case "UNIT"
            r.Font.Size = 9
            r.ParagraphFormat.Alignment = 2
        Case "LEG"
            r.Font.Size = 9: r.Font.Color = RGB(89, 89, 89)
        Case "NOTE"
            r.Font.Size = 9
            r.ParagraphFormat.LeftIndent = wd.CentimetersToPoints(0.3)
            r.ParagraphFormat.KeepWithNext = False
    End Select
    r.InsertParagraphAfter
    ResetLastParagraph doc
End Function

Private Sub ColorLegend(doc As Object, st As Long, txt As String)
    Dim p As Long, s0 As Long
    p = InStr(txt, " / ")
    If p = 0 Then Exit Sub
    s0 = IIf(Left(txt, 2) = "* ", 2, 0)
    doc.Range(st + s0, st + p - 1).Font.Color = RGB(192, 0, 0)
    doc.Range(st + p + 2, st + Len(txt)).Font.Color = RGB(0, 112, 192)
End Sub

Private Sub PutTable(doc As Object, ws As Worksheet, r As Long, parts As Variant)
    Dim nHead As Long, nBody As Long, nFoot As Long, nC As Long, gCol As Long, cCol As Long
    Dim w As Variant, tmp As Worksheet, outR As Long, bStart As Long, i As Long, rr As Long, j As Long
    Dim code As Variant, clr As Long, t As Object, wr As Object

    nHead = CLng(parts(1)): nBody = CLng(parts(2)): nFoot = CLng(parts(3))
    nC = CLng(parts(4)): gCol = CLng(parts(5)): cCol = CLng(parts(6))
    w = Split(parts(7), ";")

    Application.DisplayAlerts = False
    Set tmp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    tmp.Cells.Font.Name = FONT_NAME

    outR = 1
    CopyBlock ws.Range(ws.Cells(r, 2), ws.Cells(r + nHead - 1, 1 + nC)), tmp, outR
    outR = outR + nHead
    bStart = outR
    For i = 0 To nBody - 1
        rr = r + nHead + i
        If Trim(CellText(ws.Cells(rr, 3))) = "" Then Exit For
        CopyBlock ws.Range(ws.Cells(rr, 2), ws.Cells(rr, 1 + nC)), tmp, outR
        If cCol > 0 Then
            code = ws.Cells(rr, 1 + cCol).Value
            clr = -1
            If CStr(code) = "1" Then clr = RGB(192, 0, 0)
            If CStr(code) = "2" Then clr = RGB(0, 112, 192)
            If clr <> -1 Then tmp.Range(tmp.Cells(outR, 2), tmp.Cells(outR, nC)).Font.Color = clr
        End If
        outR = outR + 1
    Next i
    If gCol > 0 And outR - 1 > bStart Then MergeGroups tmp, bStart, outR - 1, gCol
    If nFoot > 0 Then
        CopyBlock ws.Range(ws.Cells(r + nHead + nBody, 2), ws.Cells(r + nHead + nBody + nFoot - 1, 1 + nC)), tmp, outR
        outR = outR + nFoot
    End If
    For j = 0 To UBound(w)
        tmp.Columns(j + 1).ColumnWidth = CDbl(w(j)) * 1.5
    Next j

    tmp.Range(tmp.Cells(1, 1), tmp.Cells(outR - 1, nC)).Copy
    Set wr = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    wr.PasteExcelTable False, False, False
    Application.CutCopyMode = False
    tmp.Delete
    Application.DisplayAlerts = True

    Set t = doc.Tables(doc.Tables.Count)
    On Error Resume Next
    t.AutoFitBehavior 2
    t.Rows.AllowBreakAcrossPages = False
    t.Range.Font.Name = FONT_NAME
    On Error GoTo 0
    Set wr = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    wr.InsertParagraphAfter
    ResetLastParagraph doc
End Sub

Private Sub CopyBlock(src As Range, tmp As Worksheet, outRow As Long)
    Dim c As Range, d As Range, i As Long
    src.Copy
    tmp.Cells(outRow, 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    For Each c In src.Cells
        If c.MergeArea.Cells(1, 1).Address = c.Address Then
            Set d = tmp.Cells(outRow + c.Row - src.Row, 1 + c.Column - src.Column)
            d.NumberFormat = "@"
            d.Value = CellText(c)
        End If
    Next c
    For i = 1 To src.Rows.Count
        tmp.Rows(outRow + i - 1).RowHeight = src.Rows(i).RowHeight
    Next i
End Sub

Private Sub MergeGroups(ws As Worksheet, r1 As Long, r2 As Long, col As Long)
    Dim s As Long, r As Long
    s = r1
    For r = r1 + 1 To r2 + 1
        If r > r2 Then
            DoMerge ws, s, r - 1, col
        ElseIf CStr(ws.Cells(r, col).Value) <> CStr(ws.Cells(s, col).Value) Then
            DoMerge ws, s, r - 1, col
            s = r
        End If
    Next r
End Sub

Private Sub DoMerge(ws As Worksheet, a As Long, b As Long, col As Long)
    If b <= a Then Exit Sub
    ws.Range(ws.Cells(a + 1, col), ws.Cells(b, col)).ClearContents
    With ws.Range(ws.Cells(a, col), ws.Cells(b, col))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub ResetLastParagraph(doc As Object)
    With doc.Paragraphs(doc.Paragraphs.Count).Format
        .PageBreakBefore = False
        .KeepWithNext = False
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LeftIndent = 0
        .Alignment = 0
    End With
End Sub
