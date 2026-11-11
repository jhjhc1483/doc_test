Attribute VB_Name = "HcExecSetup"
Option Explicit

' ===========================================================================
' Setup: builds every sheet (run once in a new workbook)
' ===========================================================================
Public Sub Setup()
    Dim ws As Object
    Set ws = SheetByName(SH_IN)
    If Not ws Is Nothing Then
        If Not Confirm(U("\uC774\uBBF8 \uB9CC\uB4E4\uC5B4\uC9C4 \uD30C\uC77C\uC785\uB2C8\uB2E4. \uCC98\uC74C \uC0C1\uD0DC(\uC608\uC2DC \uB0B4\uC6A9)\uB85C \uB2E4\uC2DC \uB9CC\uB4E4\uAE4C\uC694? \uC9C0\uAE08 \uB0B4\uC6A9\uC740 \uC9C0\uC6CC\uC9D1\uB2C8\uB2E4.")) Then Exit Sub
    End If
    BeginWork
    SetupAll
    EndWork
    Info U("\uC784\uC6D0 \uC77C\uC77C \uC608\uC815\uC0AC\uD56D \uD30C\uC77C\uC744 \uB9CC\uB4E4\uC5C8\uC2B5\uB2C8\uB2E4. [\uB2E4\uB978 \uC774\uB984\uC73C\uB85C \uC800\uC7A5]\uC5D0\uC11C '\uB9E4\uD06C\uB85C \uC0AC\uC6A9 \uBB38\uC11C(.xlsm)'\uB85C \uC800\uC7A5\uD558\uC138\uC694.")
End Sub

' used by tools/build.py (LibreOffice)
Public Function BuildRelease() As String
    gQuiet = True
    On Error GoTo fail
    SetupAll
    gQuiet = False
    BuildRelease = "ok " & gDbg
    Exit Function
fail:
    gQuiet = False
    BuildRelease = "ERROR " & Err.Number & ": " & Err.Description & " [step " & gStep & "] [" & gTrace & "]"
End Function

Private Sub Lbl(ByVal ws As Object, ByVal addr As String, ByVal txt As String, Optional ByVal bold As Boolean = False, Optional ByVal size As Double = 10, Optional ByVal color As Long = 0)
    With ws.Range(addr)
        If .Cells.Count > 1 Then .Merge
        .Cells(1, 1).Value = txt
        .Font.Bold = bold
        .Font.Size = size
        .Font.Color = color
        .VerticalAlignment = XL_CENTER
    End With
End Sub

Private Sub Inp(ByVal ws As Object, ByVal addr As String, ByVal v As Variant, ByVal nm As String, Optional ByVal textFmt As Boolean = True)
    With ws.Range(addr)
        If .Cells.Count > 1 Then .Merge
        If textFmt Then .NumberFormat = "@"
        .Cells(1, 1).Value = v
        .Interior.Color = 16777215
        .Borders.LineStyle = XL_CONT
        .Borders.Color = COL_INLINE
        .VerticalAlignment = XL_CENTER
        .Font.Size = 10
    End With
    If nm <> "" Then DefName nm, ws, ws.Range(addr).Cells(1, 1).Address
End Sub

Private Sub ListValid(ByVal ws As Object, ByVal addr As String, ByVal items As String)
    On Error Resume Next
    With ws.Range(addr).Validation
        .Delete
        .Add 3, 1, 1, items
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Private Sub SheetBase(ByVal ws As Object)
    ws.Cells.Font.Name = U("\uB9D1\uC740 \uACE0\uB515")
    ws.Cells.Font.Size = 10
    HideGrid ws
End Sub

Private Sub SetupAll()
    Dim wb As Object, ws As Object, k As Long
    Set wb = HostBook()
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Styles("Normal").Font.Name = U("\uB9D1\uC740 \uACE0\uB515")
    wb.Styles("Normal").Font.Size = 10
    On Error GoTo 0
    '
    gStep = "setup start"
    Set ws = wb.Worksheets(1)
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    ClearSheet ws
    On Error Resume Next
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
    gStep = "rename input"
    ws.Name = SH_IN
    Dim a4 As Object, prev As Object
    Set a4 = wb.Worksheets.Add(, ws)
    a4.Name = SH_A4
    Set prev = a4
    Dim samp As Variant
    For k = 1 To NX
        Set ws = wb.Worksheets.Add(, prev)
        gStep = "exec sheet " & k
        samp = SampleExec(k)
        SetupExecSheet ws, samp(0), samp(1), samp(2)
        ws.Name = SafeSheetName(samp(0) & " " & samp(1))
        Set prev = ws
    Next
    gStep = "tables"
    Set ws = wb.Worksheets.Add(, prev): ws.Name = SH_TBL: SetupTables ws: Set prev = ws
    Set ws = wb.Worksheets.Add(, prev): ws.Name = SH_FMT: SetupFormat ws: Set prev = ws
    Set ws = wb.Worksheets.Add(, prev): ws.Name = SH_HELP: SetupHelp ws: Set prev = ws
    Set ws = wb.Worksheets.Add(, prev): ws.Name = SH_PASTE: SetupPaste ws: Set prev = ws
    gStep = "scratch"
    ScratchSheet SH_MS
    ScratchSheet SH_BAK
    gStep = "input"
    SetupInput wb.Worksheets(SH_IN)
    gStep = "a4"
    SetupA4 a4
    Application.DisplayAlerts = True
    gStep = "build"
    DoBuild
    wb.Worksheets(SH_IN).Activate
End Sub

Private Function SampleExec(ByVal k As Long) As Variant
    Select Case k
        Case 1: SampleExec = Array(U("\uBD80\uC0AC\uC7A5"), U("\uC774\uC815\uD6C8"), Array( _
            U("09:00|\uC778\uC0AC\uC704\uC6D0\uD68C @\uC18C\uD68C\uC758\uC2E4|\uD558\uBC18\uAE30 \uC2B9\uC9C4 \uC2EC\uC0AC|16:00 \uC774\uD6C4 \uC0AC\uB0B4 \uB300\uAE30"), _
            U("10:00|\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4||"), U("13:30|\uAC10\uC0AC\uD300 \uC5C5\uBB34\uBCF4\uACE0||"), U("15:00|\uC678\uBD80 \uAC10\uC0AC\uC778 \uBBF8\uD305 @\uC811\uACAC\uC2E4||")))
        Case 2: SampleExec = Array(U("\uBD80\uC0AC\uC7A5"), U("\uBC15\uC11C\uC5F0"), Array( _
            U("08:30|\uC784\uC6D0 \uC870\uCC2C \uAC04\uB2F4\uD68C @20\uCE35 \uB77C\uC6B4\uC9C0||\uD604\uC7A5 \uC810\uAC80 \uD6C4 \uC9C1\uD1F4"), U("10:00|\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4||"), _
            U("14:00~17:00|\uC601\uC5C5 \uD604\uC7A5 \uC810\uAC80 (\uBD80\uC0B0\uC9C0\uC0AC)|\uB3D9\uD589: \uC601\uC5C5\uBCF8\uBD80\uC7A5, \uC601\uB0A8\uC9C0\uC5ED\uD300\uC7A5|")))
        Case 3: SampleExec = Array(U("\uC804\uBB34"), U("\uCD5C\uC601\uD638"), Array( _
            U("10:00|\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4||"), U("11:00|\uC2E0\uC0AC\uC5C5 TF \uD68C\uC758|2027\uB144 \uC2E0\uADDC \uC9C4\uCD9C \uC2DC\uC7A5 \uAC80\uD1A0|"), U("15:00|\uD611\uB825\uC0AC \uB300\uD45C \uBA74\uB2F4 @\uC811\uACAC\uC2E4||")))
        Case 4: SampleExec = Array(U("\uB300\uD45C\uC774\uC0AC"), U("\uAE40\uBBFC\uC218"), Array( _
            U("08:30|\uC784\uC6D0 \uC870\uCC2C \uAC04\uB2F4\uD68C @20\uCE35 \uB77C\uC6B4\uC9C0|\uCC38\uC11D: \uBD80\uC0AC\uC7A5 \uC774\uD558 \uC784\uC6D0 \uC804\uC6D0|11/10(\uD654) \uC2F1\uAC00\uD3EC\uB974 \uCD9C\uC7A5 (08:00 \uC778\uCC9C \uCD9C\uBC1C)"), _
            U("10:00|\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4^- \uC548\uAC74: 4\uBD84\uAE30 \uC2E4\uC801 \uC804\uB9DD, 2027\uB144 \uC0AC\uC5C5\uACC4\uD68D \uBC29\uD5A5||"), _
            U("11:30|\uC2E0\uADDC \uACE0\uAC1D\uC0AC \uB300\uD45C \uC811\uACAC @\uC811\uACAC\uC2E4||"), U("12:00|\uC624\uCC2C (\uACE0\uAC1D\uC0AC \uB300\uD45C)||"), _
            U("14:00|\uB178\uC0AC\uD611\uC758\uD68C \uC815\uAE30\uD68C\uC758 @3\uCE35 \uD68C\uC758\uC2E4|\uB178\uCE21 6\uBA85, \uC0AC\uCE21 6\uBA85|"), _
            U("16:00|\uD22C\uC790\uC2EC\uC758\uC704\uC6D0\uD68C @\uB300\uD68C\uC758\uC2E4|\uC548\uAC74: \uC81C2\uACF5\uC7A5 \uC124\uBE44 \uC99D\uC124 (!!\uC2B9\uC778 \uD544\uC694!!)|"), U("17:30|\uACB0\uC7AC||"), U("19:00|\uC5C5\uACC4 CEO \uB9CC\uCC2C||")))
        Case 5: SampleExec = Array(U("\uC804\uBB34"), U("\uC815\uB2E4\uC740"), Array( _
            U("09:30|\uC7AC\uBB34\uD300 \uC790\uAE08 \uD604\uD669 \uBCF4\uACE0||\uC624\uD6C4 \uC678\uBD80 \uC77C\uC815 \uD6C4 \uBCF5\uADC0"), U("10:00|\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4||"), U("14:00|\uC8FC\uAC70\uB798 \uC740\uD589 \uBBF8\uD305 @\uC678\uBD80||")))
        Case 6: SampleExec = Array(U("\uC0C1\uBB34"), U("\uAC15\uD0DC\uC900"), Array( _
            U("\uC885\uC77C|\uD574\uC678 \uBC14\uC774\uC5B4 \uC0C1\uB2F4 (\uC11C\uC6B8\uC0AC\uBB34\uC18C)|\uBCA0\uD2B8\uB0A8 \uBC14\uC774\uC5B4 3\uAC1C\uC0AC|10:00 \uACBD\uC601\uD68C\uC758 \uBD88\uCC38 (\uC815\uB2E4\uC740 \uC804\uBB34 \uB300\uB9AC \uBCF4\uACE0)")))
        Case 7: SampleExec = Array(U("\uC0C1\uBB34"), U("\uC724\uC9C0\uD6C8"), Array( _
            U("10:00|\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4||"), U("13:00|\uC5F0\uAD6C\uC18C \uACFC\uC81C \uC810\uAC80 @\uAE30\uC220\uC5F0\uAD6C\uC18C|\uD558\uBC18\uAE30 \uC911\uC810\uACFC\uC81C 5\uAC74|"), U("16:30|\uD488\uC9C8 \uD68C\uC758||")))
        Case 8: SampleExec = Array(U("\uC0C1\uBB34"), U("\uD55C\uC18C\uC601"), Array( _
            U("09:00|\uD64D\uBCF4 \uC804\uB7B5 \uD68C\uC758||\uB0B4\uC77C \uC624\uC804 \uBC18\uCC28"), U("11:00|\uC5B8\uB860\uC0AC \uC778\uD130\uBDF0 @\uD64D\uBCF4\uC2E4|\uC8FC\uC81C: \uCE5C\uD658\uACBD \uACBD\uC601|"), U("15:00|\uC0AC\uD68C\uACF5\uD5CC \uC704\uC6D0\uD68C||")))
    End Select
End Function

Public Sub SetupExecSheet(ByVal ws As Object, ByVal pos As String, ByVal nm As String, ByVal rows As Variant)
    gStep = "exec sheet clear"
    ClearSheet ws
    gStep = "exec sheet base"
    SheetBase ws
    gStep = "exec sheet widths"
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 54
    ws.Columns(3).ColumnWidth = 32
    ws.Columns(4).ColumnWidth = 34
    Dim cc As Variant
    gStep = "exec sheet labels"
    ws.Range("A1").Value = U("\uC9C1\uC704")
    ws.Range("C1").Value = U("\uC131\uBA85")
    For Each cc In Array("A1", "C1")
        ws.Range(cc).Font.Bold = True
        ws.Range(cc).Interior.Color = COL_HEAD
        ws.Range(cc).HorizontalAlignment = XL_CENTER
    Next
    For Each cc In Array("B1", "D1")
        ws.Range(cc).NumberFormat = "@"
        ws.Range(cc).Font.Bold = True
    Next
    ws.Range("B1").Value = pos
    ws.Range("D1").Value = nm
    ws.Range("A1:D1").Borders.LineStyle = XL_CONT
    ws.Range("A1:D1").Borders.Color = COL_INLINE
    ws.Rows(1).RowHeight = 22
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = U("\uD55C \uC904\uC5D0 \uC77C\uC815 \uD558\uB098 \u00B7 \uC2DC\uAC04\uC740 09:00, 14:00~16:00, \uC885\uC77C \u00B7 \uB0B4\uC6A9 \uB05D\uC5D0 @\uC7A5\uC18C \u00B7 \uB0B4\uC6A9\uC774 * \uB85C \uC2DC\uC791\uD558\uBA74 \uC704 \uC77C\uC815\uC758 \uC791\uC740 \uAE00\uC528 \u00B7 \u203B \uCC38\uACE0\uB294 \uCE78 \uB9E8 \uBC11 \u00B7 \uC2DC\uD2B8 \uC21C\uC11C = \uC790\uB9AC \uC21C\uC11C")
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = COL_GRAY
    ws.Range("A2").WrapText = True
    ws.Rows(2).RowHeight = 28
    With ws.Range("A3:D3")
        .Value = Array(U("\uC2DC\uAC04"), U("\uB0B4\uC6A9"), U("\uBE44\uACE0 (\uC791\uC740 \uAE00\uC528 *)"), U("\u203B \uCC38\uACE0"))
        .Font.Bold = True
        .Interior.Color = COL_HEAD
        .HorizontalAlignment = XL_CENTER
        .Borders.LineStyle = XL_CONT
        .Borders.Color = COL_INLINE
    End With
    With ws.Range("A4:D200")
        .NumberFormat = "@"
        .WrapText = True
        .VerticalAlignment = XL_TOP
    End With
    ws.Range("A4:A200").HorizontalAlignment = XL_CENTER
    gStep = "exec sheet rows"
    Dim i As Long, f As Variant
    If IsArray(rows) Then
        For i = 0 To UBound(rows)
            f = Split(CStr(rows(i)), "|")
            ws.Cells(4 + i, 1).Value = f(0)
            ws.Cells(4 + i, 2).Value = Replace(f(1), "^", vbLf)
            ws.Cells(4 + i, 3).Value = f(2)
            ws.Cells(4 + i, 4).Value = f(3)
        Next
    End If
    On Error Resume Next
    Dim cur As Object
    Set cur = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("B4").Select
    If Not cur Is Nothing Then cur.Activate
    On Error GoTo 0
    AddButton ws, ws.Columns(5).Left + 12, 4, 96, 24, U("A4 \uB9CC\uB4E4\uAE30"), "BuildA4"
    AddButton ws, ws.Columns(5).Left + 12, 32, 96, 24, U("\uD615\uAD11\uD39C"), "HighlightSel"
    AddButton ws, ws.Columns(5).Left + 12, 60, 96, 24, U("\uC785\uB825\uC73C\uB85C"), "GoInput"
End Sub

Private Sub SetupInput(ByVal ws As Object)
    Dim c As Long
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 2
    For c = 2 To 13
        ws.Columns(c).ColumnWidth = 7
    Next
    ws.Columns(14).ColumnWidth = 2
    Lbl ws, "B1:I1", U("\uC784\uC6D0 \uC77C\uC77C \uC608\uC815\uC0AC\uD56D"), True, 15
    Lbl ws, "J1:M1", U("8\uBA85 \u00B7 A4 \uD55C \uC7A5 \uC790\uB3D9 \uB9DE\uCDA4"), False, 9, COL_GRAY
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 60
    ws.Range("B3:M3").Merge
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").WrapText = True
    ws.Rows(3).RowHeight = 30
    DefName "d_status", ws, "$B$3"
    '
    Lbl ws, "B5:M5", U("\uBB38\uC11C \uC815\uBCF4"), True, 11
    Lbl ws, "B6:C6", U("\uC81C\uBAA9"), False, 9, COL_GRAY: Inp ws, "D6:M6", U("\uC784\uC6D0 \uC77C\uC77C \uC608\uC815\uC0AC\uD56D"), "d_title"
    Lbl ws, "B7:C7", U("\uB0A0\uC9DC"), False, 9, COL_GRAY: Inp ws, "D7:F7", Date, "d_date", False
    ws.Range("D7").NumberFormat = "yyyy-mm-dd"
    ws.Range("D7").HorizontalAlignment = XL_LEFT
    Lbl ws, "G7:H7", U("\uC67C\uCABD \uC704 \uBB38\uAD6C"), False, 9, COL_GRAY: Inp ws, "I7:M7", "", "d_sub"
    Lbl ws, "B8:C8", U("\uD558\uB2E8 \uBB38\uAD6C"), False, 9, COL_GRAY: Inp ws, "D8:M8", "", "d_foot"
    ws.Range("D6:M8").Rows.RowHeight = 20
    '
    Lbl ws, "B10:M10", U("\uBC30\uCE58 (\uC784\uC6D0 \uC2DC\uD2B8\uC758 \uD0ED \uC21C\uC11C = \uC790\uB9AC \uC21C\uC11C 1~8)"), True, 11
    Dim r As Long
    For r = 11 To 14
        ws.Rows(r).RowHeight = 34
    Next
    DefName "d_map", ws, "$B$11"
    ws.Range("B15:M15").Merge
    ws.Rows(15).RowHeight = 20
    ws.Range("B16:M16").Merge
    ws.Range("B16").Value = U("\uCE78\uC744 \uB204\uB974\uBA74 \uADF8 \uC784\uC6D0 \uC2DC\uD2B8\uB85C \uAC11\uB2C8\uB2E4. \uC790\uB9AC\uB97C \uBC14\uAFB8\uB824\uBA74 [\uC790\uB9AC \uBC14\uAFB8\uAE30]\uB97C \uB204\uB974\uAC70\uB098 \uC784\uC6D0 \uC2DC\uD2B8 \uD0ED\uC744 \uB04C\uC5B4 \uC21C\uC11C\uB97C \uBC14\uAFB8\uC138\uC694. \uC9C1\uC704\u00B7\uC131\uBA85\uC740 \uAC01 \uC784\uC6D0 \uC2DC\uD2B8 \uB9E8 \uC704\uC5D0\uC11C \uACE0\uCE69\uB2C8\uB2E4.")
    ws.Range("B16").Font.Size = 9
    ws.Range("B16").Font.Color = COL_GRAY
    ws.Range("B16").WrapText = True
    ws.Rows(16).RowHeight = 28
    '
    Lbl ws, "B18:M18", U("\uB9E4\uC77C \uD558\uB294 \uC77C"), True, 11
    Dim steps As Variant, i As Long
    steps = Array(U("1. \uAC01 \uC784\uC6D0 \uC2DC\uD2B8\uC5D0 \uC77C\uC815\uC744 \uC501\uB2C8\uB2E4(\uB610\uB294 [\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30]\uB85C \uB450\uB808\uC774 \uACF5\uB3D9\uD3B8\uC9D1 \uD30C\uC77C\uC744 \uAC00\uC838\uC635\uB2C8\uB2E4)."), _
        U("2. [A4 \uB9CC\uB4E4\uAE30]\uB97C \uB204\uB974\uBA74 A4 \uC2DC\uD2B8\uC5D0 \uD55C \uC7A5\uC9DC\uB9AC \uBB38\uC11C\uAC00 \uB9CC\uB4E4\uC5B4\uC9D1\uB2C8\uB2E4(\uAE00\uC790\uD06C\uAE30 \uC790\uB3D9 \uB9DE\uCDA4)."), _
        U("3. [\uC778\uC1C4 \uBBF8\uB9AC\uBCF4\uAE30] \uB610\uB294 A4 \uC2DC\uD2B8\uC5D0\uC11C Ctrl+P. PDF\uB85C \uC800\uC7A5\uB3C4 \uB429\uB2C8\uB2E4."), _
        U("\u203B \uB2E8\uCD94\uAC00 \uC548 \uBCF4\uC774\uBA74 Alt+F8(\uB9E4\uD06C\uB85C)\uC5D0\uC11C BuildA4(A4 \uB9CC\uB4E4\uAE30), ImportExcel(\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30), UndoLast(\uB418\uB3CC\uB9AC\uAE30) \uB4F1\uC744 \uC2E4\uD589\uD558\uC138\uC694."))
    For i = 0 To UBound(steps)
        ws.Range(ws.Cells(19 + i, 2), ws.Cells(19 + i, 13)).Merge
        ws.Cells(19 + i, 2).Value = steps(i)
        ws.Cells(19 + i, 2).Font.Size = 9
        ws.Cells(19 + i, 2).WrapText = True
        ws.Rows(19 + i).RowHeight = 26
    Next
    ' buttons
    Dim y As Double, x As Double, bw As Double
    y = ws.Rows(2).Top + 4
    x = ws.Columns(2).Left
    bw = 102
    AddButton ws, x, y, bw, 24, U("A4 \uB9CC\uB4E4\uAE30"), "BuildA4"
    AddButton ws, x + (bw + 6), y, bw, 24, U("\uC778\uC1C4 \uBBF8\uB9AC\uBCF4\uAE30"), "PrintPreviewA4"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, U("\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30"), "ImportExcel"
    AddButton ws, x + 3 * (bw + 6), y, bw, 24, U("\uBD99\uC5EC\uB123\uC740 \uD45C \uAC00\uC838\uC624\uAE30"), "ImportPasted"
    AddButton ws, x + 4 * (bw + 6), y, bw, 24, U("\uB418\uB3CC\uB9AC\uAE30"), "UndoLast"
    y = y + 28
    AddButton ws, x, y, bw, 24, U("\uC0C8 \uBB38\uC11C"), "NewDoc"
    AddButton ws, x + (bw + 6), y, bw, 24, U("\uC790\uB9AC \uBC14\uAFB8\uAE30"), "SwapSeats"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, U("\uD615\uAD11\uD39C"), "HighlightSel"
    AddButton ws, x + 3 * (bw + 6), y, bw * 2 + 6, 24, U("\uACF5\uB3D9\uD3B8\uC9D1\uC6A9 \uC5D1\uC140 \uC591\uC2DD \uB9CC\uB4E4\uAE30"), "MakeTemplate"
End Sub

Private Sub SetupA4(ByVal ws As Object)
    ClearSheet ws
    HideGrid ws
    On Error Resume Next
    Dim tb As Object
    Set tb = ws.Shapes.AddTextbox(1, 600, 6, 260, 70)
    If Not tb Is Nothing Then
        tb.Name = "a4status"
        tb.TextFrame.Characters.Text = ""
        tb.TextFrame.Characters.Font.Size = 9
        tb.Line.Visible = False
    End If
    On Error GoTo 0
    AddButton ws, 600, 90, 120, 26, U("\uB2E4\uC2DC \uB9CC\uB4E4\uAE30"), "BuildA4"
    AddButton ws, 600, 122, 120, 26, U("\uC778\uC1C4 \uBBF8\uB9AC\uBCF4\uAE30"), "PrintPreviewA4"
    AddButton ws, 600, 154, 120, 26, U("\uC785\uB825\uC73C\uB85C"), "GoInput"
End Sub

Private Sub SetupTables(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 6
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 70
    Lbl ws, "A1:D1", U("\uB9E8 \uBC11 \uD45C 4\uAC1C (\uC778\uC811\uD68C\uC0AC\uD65C\uB3D9 \u00B7 \uD734\uAC00 \u00B7 \uCD9C\uC7A5/\uBC29\uBB38 \u00B7 \uADFC\uBB34)"), True, 12
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = U("\uD45C \uC774\uB984\u00B7\uCE78 \uC81C\uBAA9\uC740 \uACE0\uCCD0 \uC368\uB3C4 \uB429\uB2C8\uB2E4. \uB0B4\uC6A9 \uCE78\uC5D0\uB294 \uC5EC\uB7EC \uC904(Alt+Enter)\uC744 \uC4F8 \uC218 \uC788\uACE0, '10:00 \uD68C\uC758'\uCC98\uB7FC \uC2DC\uAC04\uC73C\uB85C \uC2DC\uC791\uD558\uBA74 \uC2DC\uAC04\uC774 \uAD75\uAC8C, '* '\uB85C \uC2DC\uC791\uD558\uBA74 \uC791\uC740 \uAE00\uC528\uC785\uB2C8\uB2E4. \uADFC\uBB34\uC758 \uCE78 \uC81C\uBAA9\uC744 \uBE44\uC6B0\uBA74 \uC81C\uBAA9 \uC904 \uC5C6\uC774 \uCC0D\uD799\uB2C8\uB2E4.")
    ws.Range("A2").WrapText = True
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = COL_GRAY
    ws.Rows(2).RowHeight = 40
    With ws.Range("A3:D3")
        .Value = Array(U("\uD45C \uC774\uB984"), U("\uCE78"), U("\uCE78 \uC81C\uBAA9"), U("\uB0B4\uC6A9"))
        .Font.Bold = True
        .Interior.Color = COL_HEAD
        .HorizontalAlignment = XL_CENTER
    End With
    DefName "t_top", ws, "$A$4"
    Dim k As Long, j As Long, r As Long, heads As Variant, texts As Variant
    For k = 1 To 4
        r = 4 + TblRow0(k)
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = Choose(k, U("\uC778\uC811\uD68C\uC0AC\uD65C\uB3D9"), U("\uD734\uAC00"), U("\uCD9C\uC7A5/\uBC29\uBB38"), U("\uADFC\uBB34"))
        ws.Cells(r, 1).Font.Bold = True
        Select Case k
            Case 1
                heads = Array(U("A\uC0AC"), U("B\uC0AC"), U("C\uC0AC"), U("D\uC0AC"), U("E\uC0AC"))
                texts = Array(U("10:00 \uC0AC\uC7A5 \uC8FC\uAC04\uD68C\uC758^14:00 \uB178\uC0AC\uD611\uC758\uD68C"), U("\uC784\uC6D0 \uC6CC\uD06C\uC20D (1\uBC15 2\uC77C)"), U("\uD2B9\uC774\uC0AC\uD56D \uC5C6\uC74C"), U("15:00 \uC2E0\uC81C\uD488 \uBC1C\uD45C\uD68C^* \uBCF8\uC0AC \uAC15\uB2F9"), U("11:00 \uB300\uD45C \uC5B8\uB860 \uC778\uD130\uBDF0"))
            Case 2
                heads = Array("")
                texts = Array(U("\uC0C1\uBB34 \uD55C\uC18C\uC601 (11/10 \uC624\uC804 \uBC18\uCC28) \u00B7 \uAE30\uD68D\uD300\uC7A5 (11/9~11/11 \uC5F0\uCC28)"))
            Case 3
                heads = Array("")
                texts = Array(U("\uBD80\uC0AC\uC7A5 \uBC15\uC11C\uC5F0: \uBD80\uC0B0\uC9C0\uC0AC \uD604\uC7A5 \uC810\uAC80 (14:00~17:00) \u00B7 \uC0C1\uBB34 \uAC15\uD0DC\uC900: \uC11C\uC6B8\uC0AC\uBB34\uC18C (\uC885\uC77C)"))
            Case 4
                heads = Array(U("\uB2F9\uC9C1"), U("\uBE44\uC0C1\uB300\uAE30"))
                texts = Array(U("\uCD1D\uBB34\uD300 \uACFC\uC7A5 (18:00~\uC775\uC77C 09:00)"), U("\uBE44\uC11C\uC2E4\uC7A5"))
        End Select
        For j = 1 To TblN(k)
            ws.Cells(r + j - 1, 2).Value = j
            ws.Cells(r + j - 1, 2).HorizontalAlignment = XL_CENTER
            ws.Range(ws.Cells(r + j - 1, 3), ws.Cells(r + j - 1, 4)).NumberFormat = "@"
            If TblHasHeads(k) Then
                ws.Cells(r + j - 1, 3).Value = heads(j - 1)
            Else
                ws.Cells(r + j - 1, 3).Value = U("(\uC81C\uBAA9 \uC5C6\uC74C)")
                ws.Cells(r + j - 1, 3).Font.Color = COL_GRAY
                ws.Cells(r + j - 1, 3).Interior.Color = COL_SOFT
            End If
            ws.Cells(r + j - 1, 4).Value = Replace(texts(j - 1), "^", vbLf)
        Next
        With ws.Range(ws.Cells(r, 1), ws.Cells(r + TblN(k) - 1, 4))
            .Borders.LineStyle = XL_CONT
            .Borders.Color = COL_INLINE
            .VerticalAlignment = XL_TOP
            .WrapText = True
        End With
    Next
    AddButton ws, ws.Columns(5).Left + 12, 4, 96, 24, U("A4 \uB9CC\uB4E4\uAE30"), "BuildA4"
    AddButton ws, ws.Columns(5).Left + 12, 32, 96, 24, U("\uD615\uAD11\uD39C"), "HighlightSel"
    AddButton ws, ws.Columns(5).Left + 12, 60, 96, 24, U("\uC785\uB825\uC73C\uB85C"), "GoInput"
End Sub

Private Sub SetupFormat(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 12
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 12
    ws.Columns(6).ColumnWidth = 40
    Lbl ws, "B1:F1", U("\uC11C\uC2DD"), True, 14
    ws.Rows(1).RowHeight = 28
    Dim r As Long
    r = 3
    Lbl ws, "B" & r & ":F" & r, U("\uC6A9\uC9C0\u00B7\uBC30\uCE58"), True, 11: r = r + 1
    Lbl ws, "B" & r, U("\uC6A9\uC9C0 \uBC29\uD5A5"): Inp ws, "C" & r & ":D" & r, U("A4 \uC138\uB85C"), "s_orient": ListValid ws, "C" & r, U("A4 \uC138\uB85C,A4 \uAC00\uB85C"): r = r + 1
    Lbl ws, "B" & r, U("\uC5EC\uBC31 (mm, 5~25)"): Inp ws, "C" & r, 10, "s_margin", False: r = r + 1
    Lbl ws, "B" & r, U("\uAE00\uAF34"): Inp ws, "C" & r & ":D" & r, U("\uB9D1\uC740 \uACE0\uB515"), "s_font": ListValid ws, "C" & r, U("\uB9D1\uC740 \uACE0\uB515,\uB098\uB214\uACE0\uB515,\uB3CB\uC6C0,\uAD74\uB9BC,\uBC14\uD0D5"): r = r + 1
    r = r + 1
    Lbl ws, "B" & r & ":F" & r, U("\uCE78 \uD06C\uAE30 (\uBE44\uC728)"), True, 11: r = r + 1
    Lbl ws, "B" & r, U("\uC717\uC904 \uCE78 \uD3ED (\uC67C\uCABD\u00B7\uAC00\uC6B4\uB370\u00B7\uC624\uB978\uCABD)"): Inp ws, "C" & r, 1, "s_c1a", False: Inp ws, "D" & r, 1, "s_c1b", False: Inp ws, "E" & r, 1, "s_c1c", False: r = r + 1
    Lbl ws, "B" & r, U("\uAC00\uC6B4\uB370 \uC904 \uD3ED (\uD070 \uCE78\u00B7\uC624\uB978\uCABD)"): Inp ws, "C" & r, 2, "s_c2a", False: Inp ws, "D" & r, 1, "s_c2b", False: r = r + 1
    Lbl ws, "B" & r, U("\uC544\uB7AB\uC904 \uCE78 \uD3ED (\uC67C\uCABD\u00B7\uC624\uB978\uCABD)"): Inp ws, "C" & r, 1, "s_c3a", False: Inp ws, "D" & r, 1, "s_c3b", False: r = r + 1
    Lbl ws, "B" & r, U("\uC904 \uB192\uC774 (\uC717\uC904\u00B7\uAC00\uC6B4\uB370\u00B7\uC544\uB7AB\uC904)"): Inp ws, "C" & r, "", "s_h1", False: Inp ws, "D" & r, "", "s_h2", False: Inp ws, "E" & r, "", "s_h3", False
    Lbl ws, "F" & r, U("\uBE44\uC6B0\uBA74 \uC790\uB3D9 (\uB0B4\uC6A9\uC5D0 \uB9DE\uCDA4)"), False, 9, COL_GRAY: r = r + 1
    Lbl ws, "B" & r, U("\uAC00\uC6B4\uB370 \uC904 \uC624\uB978\uCABD \uC704 \uCE78 \uB192\uC774"): Inp ws, "C" & r, "", "s_split", False
    Lbl ws, "D" & r & ":F" & r, U("0.1~0.9 (\uBE44\uC6B0\uBA74 \uC790\uB3D9, 0.5 = \uC704\uC544\uB798 \uAC19\uAC8C)"), False, 9, COL_GRAY: r = r + 1
    Lbl ws, "B" & r, U("\uC9C0\uAE08 \uD06C\uAE30"), False, 10, COL_GRAY
    ws.Range("C" & r & ":F" & r).Merge
    ws.Range("C" & r).WrapText = True
    ws.Range("C" & r).Font.Size = 9
    ws.Rows(r).RowHeight = 30
    DefName "s_sizes", ws, "$C$" & r
    r = r + 1
    Dim btnY As Double
    ws.Rows(r).RowHeight = 34
    btnY = ws.Rows(r).Top + 5
    AddButton ws, ws.Columns(2).Left, btnY, 118, 24, U("\uAC00\uC6B4\uB370 \uD070 \uCE78 2/3"), "PresetBig23"
    AddButton ws, ws.Columns(2).Left + 124, btnY, 118, 24, U("\uAC00\uC6B4\uB370 \uD070 \uCE78 1/2"), "PresetBig12"
    AddButton ws, ws.Columns(2).Left + 248, btnY, 118, 24, U("\uB192\uC774\uB9CC \uC790\uB3D9\uC73C\uB85C"), "PresetAutoH"
    AddButton ws, ws.Columns(2).Left + 372, btnY, 118, 24, U("\uBAA8\uB450 \uCC98\uC74C\uB300\uB85C"), "PresetReset"
    r = r + 2
    Lbl ws, "B" & r & ":F" & r, U("\uAE00\uC790\uD06C\uAE30 (\uC790\uB3D9 \uB9DE\uCDA4)"), True, 11: r = r + 1
    Lbl ws, "B" & r, U("\uCD5C\uB300 (pt)"): Inp ws, "C" & r, 12, "s_max", False: r = r + 1
    Lbl ws, "B" & r, U("\uCD5C\uC18C (pt)"): Inp ws, "C" & r, 7, "s_min", False: r = r + 1
    ws.Range("B" & r & ":F" & r).Merge
    ws.Range("B" & r).Value = U("8\uCE78 \uBAA8\uB450 \uAC19\uC740 \uAE00\uC790\uD06C\uAE30\uB85C, A4 \uD55C \uC7A5\uC744 \uAF49 \uCC44\uC6B0\uB294 \uAC00\uC7A5 \uD070 \uD06C\uAE30\uB97C \uC774 \uBC94\uC704 \uC548\uC5D0\uC11C \uACE0\uB985\uB2C8\uB2E4(0.5pt \uB2E8\uC704). \uB0A8\uB294 \uB192\uC774\uB294 \uCE78\uC744 \uB298\uB824 \uCC44\uC6B0\uACE0, \u203B \uCC38\uACE0\uB294 \uB298 \uCE78\uC758 \uB9E8 \uBC11\uC5D0 \uBD99\uC2B5\uB2C8\uB2E4.")
    ws.Range("B" & r).WrapText = True
    ws.Range("B" & r).Font.Size = 9
    ws.Range("B" & r).Font.Color = COL_GRAY
    ws.Rows(r).RowHeight = 30
    r = r + 2
    Lbl ws, "B" & r & ":F" & r, U("\uAE30\uD0C0"), True, 11: r = r + 1
    Lbl ws, "B" & r, U("\uB9E8 \uBC11 \uD45C 4\uAC1C \uB123\uAE30"): Inp ws, "C" & r, U("\uC608"), "s_tables": ListValid ws, "C" & r, U("\uC608,\uC544\uB2C8\uC624"): r = r + 1
    Lbl ws, "B" & r, U("\uC624\uB978\uCABD \uC704 \uBE68\uAC04 \uC0C1\uC790"): Inp ws, "C" & r, U("\uC608"), "s_secret": ListValid ws, "C" & r, U("\uC608,\uC544\uB2C8\uC624")
    Inp ws, "D" & r & ":E" & r, U("\uBE44\uACF5\uAC1C"), "s_secretText": r = r + 1
    Lbl ws, "B" & r, U("\uC77C\uC815\uC774 \uC5C6\uB294 \uCE78\uC5D0 \uC4F8 \uBB38\uAD6C"): Inp ws, "C" & r & ":E" & r, U("\uC77C\uC815 \uC5C6\uC74C"), "s_empty"
    Lbl ws, "F" & r, U("\uBE44\uC6CC \uB450\uBA74 \uBE48\uCE78"), False, 9, COL_GRAY: r = r + 2
    Lbl ws, "B" & r & ":F" & r, U("\uBB38\uC11C \uC554\uD638"), True, 11: r = r + 1
    ws.Range("B" & r & ":F" & r).Merge
    ws.Range("B" & r).Value = U("\uC774 \uD30C\uC77C\uC744 \uC5F4 \uB54C \uC554\uD638\uB97C \uBB3B\uAC8C \uD558\uB824\uBA74 \uD55C\uC140\uC758 \uBB38\uC11C \uC554\uD638\uB97C \uC4F0\uC138\uC694(\uB0B4\uC6A9\uC774 \uC554\uD638\uD654\uB429\uB2C8\uB2E4). \uC544\uB798 \uB2E8\uCD94\uB85C \uC815\uD558\uAC70\uB098, \uD55C\uC140 \uBA54\uB274\uC758 \uBB38\uC11C \uC554\uD638\uC5D0\uC11C \uC815\uD569\uB2C8\uB2E4. \uC554\uD638\uB97C \uC78A\uC73C\uBA74 \uC5F4 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.")
    ws.Range("B" & r).WrapText = True
    ws.Range("B" & r).Font.Size = 9
    ws.Rows(r).RowHeight = 30
    r = r + 1
    ws.Rows(r).RowHeight = 34
    AddButton ws, ws.Columns(2).Left, ws.Rows(r).Top + 5, 140, 24, U("\uBB38\uC11C \uC554\uD638 \uC815\uD558\uAE30"), "SetDocPassword"
    AddButton ws, ws.Columns(2).Left + 146, ws.Rows(r).Top + 5, 118, 24, U("A4 \uB9CC\uB4E4\uAE30"), "BuildA4"
End Sub

Private Sub SetupPaste(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Range("A1:H1").Merge
    ws.Range("A1").Value = U("\uC5D1\uC140(\uB450\uB808\uC774)\uC5D0\uC11C \uBCF5\uC0AC\uD55C \uD45C\uB97C 3\uBC88\uC9F8 \uC904 A\uCE78\uBD80\uD130 \uBD99\uC5EC\uB123\uACE0 [\uBD99\uC5EC\uB123\uC740 \uD45C \uAC00\uC838\uC624\uAE30]\uB97C \uB204\uB974\uC138\uC694. \uBA38\uB9AC\uAE00(\uC2DC\uAC04\u00B7\uB0B4\uC6A9\u2026)\uC774 \uC788\uC5B4\uC57C \uD558\uACE0, \uC5EC\uB7EC \uC784\uC6D0\uC774\uBA74 '\uC131\uBA85' \uC5F4\uC744 \uB450\uC138\uC694.")
    ws.Range("A1").WrapText = True
    ws.Range("A1").Font.Size = 9
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 34
    AddButton ws, 4, ws.Rows(2).Top + 5, 140, 24, U("\uBD99\uC5EC\uB123\uC740 \uD45C \uAC00\uC838\uC624\uAE30"), "ImportPasted"
    AddButton ws, 150, ws.Rows(2).Top + 5, 96, 24, U("\uC785\uB825\uC73C\uB85C"), "GoInput"
End Sub

Private Sub SetupHelp(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 44
    ws.Columns(2).ColumnWidth = 60
    Dim rows As Collection, i As Long, f As Variant, k As Long, ln As String
    Set rows = New Collection
    rows.Add U("#\uC77C\uC815 \uCE78 \u2014 \uD55C \uC904\uC5D0 \uC77C\uC815 \uD558\uB098")
    rows.Add U("\uC785\uB825|\uC778\uC1C4\uBB3C")
    rows.Add U("\uC2DC\uAC04 \uCE78 09:00, \uB0B4\uC6A9 \uCE78 \uC8FC\uAC04 \uC5C5\uBB34\uD68C\uC758|\uC2DC\uAC04 09:00 (\uAD75\uAC8C) / \uB0B4\uC6A9")
    rows.Add U("14:00~16:00|\uC2DC\uAC04 \uBC94\uC704 (~ \uB4A4\uC5D0\uC11C \uC904 \uBC14\uAFC8)")
    rows.Add U("9:30, 9\uC2DC \uBC18, \uC624\uD6C4 2\uC2DC|09:30, 09:30, 14:00 \uCC98\uB7FC \uB9DE\uCDB0 \uC90C")
    rows.Add U("\uC885\uC77C, \uC624\uC804 (\uC2DC\uAC04 \uCE78\uC5D0)|\uC2DC\uAC04 \uCE78\uC5D0 \uC885\uC77C / \uC624\uC804")
    rows.Add U("\uC2DC\uAC04 \uCE78\uC744 \uBE44\uC6B0\uACE0 \uB0B4\uC6A9\uC5D0 09:00 \uD68C\uC758|\uC2DC\uAC04\uACFC \uB0B4\uC6A9\uC744 \uC54C\uC544\uC11C \uB098\uB214")
    rows.Add U("\uB0B4\uC6A9 \uB05D\uC5D0 @\uC18C\uD68C\uC758\uC2E4|\uB0B4\uC6A9 (\uC18C\uD68C\uC758\uC2E4)")
    rows.Add U("\uBE44\uACE0 \uCE78, \uB610\uB294 \uB0B4\uC6A9\uC774 * \uB85C \uC2DC\uC791\uD558\uB294 \uC904|\uBC14\uB85C \uC704 \uC77C\uC815 \uBC11\uC5D0 \uC791\uC740 \uAE00\uC528 \uBD80\uAC00 \uC124\uBA85")
    rows.Add U("\uB0B4\uC6A9 \uCE78 \uC548\uC5D0\uC11C Alt+Enter \uD6C4 - \uC548\uAC74|\uBC14\uB85C \uC704 \uC77C\uC815 \uBC11\uC5D0 \uD558\uC704 \uD56D\uBAA9")
    rows.Add U("\u203B \uCC38\uACE0 \uCE78, \uB610\uB294 \uB0B4\uC6A9\uC774 \u203B \uB85C \uC2DC\uC791|\uADF8 \uC784\uC6D0 \uCE78 \uB9E8 \uBC11 \u203B \uCC38\uACE0")
    rows.Add U("**\uAD75\uAC8C**, !!\uBE68\uAC04 \uAC15\uC870!!, __\uBC11\uC904__, ==\uD615\uAD11\uD39C==|\uAE00\uC790 \uAFB8\uBC08 (\uD615\uAD11\uD39C\uC740 \uADF8 \uC904\uC758 \uCE78 \uC804\uCCB4\uAC00 \uB178\uB797\uAC8C)")
    rows.Add ""
    rows.Add U("#\uB2E8\uCD94")
    rows.Add U("A4 \uB9CC\uB4E4\uAE30|A4 \uC2DC\uD2B8\uC5D0 \uD55C \uC7A5\uC9DC\uB9AC \uBB38\uC11C\uB97C \uC0C8\uB85C \uB9CC\uB4ED\uB2C8\uB2E4(\uAE00\uC790\uD06C\uAE30 \uC790\uB3D9 \uB9DE\uCDA4)")
    rows.Add U("\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30|\uB450\uB808\uC774\uC5D0\uC11C \uB0B4\uB824\uBC1B\uC740 \uACF5\uB3D9\uD3B8\uC9D1 \uD30C\uC77C(.xlsx\u00B7.csv)\uC744 \uACE8\uB77C \uC784\uC6D0\uBCC4\uB85C \uB123\uC2B5\uB2C8\uB2E4")
    rows.Add U("\uBD99\uC5EC\uB123\uC740 \uD45C \uAC00\uC838\uC624\uAE30|\uBD99\uC5EC\uB123\uAE30 \uC2DC\uD2B8\uC5D0 \uBD99\uC778 \uD45C\uB97C \uC784\uC6D0\uBCC4\uB85C \uB123\uC2B5\uB2C8\uB2E4")
    rows.Add U("\uB418\uB3CC\uB9AC\uAE30|\uB9C8\uC9C0\uB9C9 [\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30]\u00B7[\uBD99\uC5EC\uB123\uC740 \uD45C \uAC00\uC838\uC624\uAE30]\u00B7[\uC0C8 \uBB38\uC11C] \uC804\uC73C\uB85C")
    rows.Add U("\uC0C8 \uBB38\uC11C|\uC77C\uC815\u00B7\u203B \uCC38\uACE0\u00B7\uB9E8 \uBC11 \uD45C \uB0B4\uC6A9\uB9CC \uBE44\uC6C0 (\uC9C1\uC704\u00B7\uC131\uBA85\u00B7\uC11C\uC2DD\uC740 \uADF8\uB300\uB85C, \uB0A0\uC9DC\uB294 \uC624\uB298)")
    rows.Add U("\uC790\uB9AC \uBC14\uAFB8\uAE30|\uB450 \uC790\uB9AC\uC758 \uC784\uC6D0\uC744 \uB9DE\uBC14\uAFC8 (\uC784\uC6D0 \uC2DC\uD2B8 \uD0ED\uC744 \uB04C\uC5B4 \uC21C\uC11C\uB97C \uBC14\uAFD4\uB3C4 \uB428)")
    rows.Add U("\uD615\uAD11\uD39C|\uACE0\uB978 \uCE78\uC758 \uAE00\uC790\uB97C ==\u2026==\uB85C \uAC10\uC308 (\uB2E4\uC2DC \uB204\uB974\uBA74 \uC9C0\uC6C0, Ctrl+Shift+H)")
    rows.Add U("\uACF5\uB3D9\uD3B8\uC9D1\uC6A9 \uC5D1\uC140 \uC591\uC2DD \uB9CC\uB4E4\uAE30|\uC784\uC6D0\uB9C8\uB2E4 \uC2DC\uD2B8\uAC00 \uC788\uB294 .xlsx \u2014 \uB450\uB808\uC774\uC5D0 \uC62C\uB824 \uD568\uAED8 \uC791\uC131")
    rows.Add ""
    rows.Add U("#\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30 \uADDC\uCE59")
    rows.Add U("\uC784\uC6D0 \uCC3E\uAE30|\uC2DC\uD2B8 \uC774\uB984\uC5D0 \uB4E0 \uC131\uBA85\uC73C\uB85C (\uC5C6\uC73C\uBA74 \uACB9\uCE58\uC9C0 \uC54A\uB294 \uC9C1\uC704\uB85C). \uBABB \uCC3E\uC73C\uBA74 \uC790\uB9AC \uBC88\uD638\uB97C \uBB3C\uC5B4\uBD04")
    rows.Add U("\uD55C \uC2DC\uD2B8\uC5D0 \uC5EC\uB7EC \uC784\uC6D0|'\uC131\uBA85' \uC5F4\uC744 \uB450\uC138\uC694 (\uBCD1\uD569\uB41C \uCE78\uB3C4 \uB428)")
    rows.Add U("\uB0A0\uC9DC \uC5F4|\uBB38\uC11C \uB0A0\uC9DC\uC758 \uC904\uB9CC \uAC00\uC838\uC634. \uBB38\uC11C \uB0A0\uC9DC\uC5D0 \uC5C6\uC73C\uBA74 \uC77C\uC815\uC774 \uAC00\uC7A5 \uB9CE\uC740 \uB0A0\uC9DC\uB97C \uAD8C\uD558\uACE0, \uAC00\uC838\uC624\uBA74 \uBB38\uC11C \uB0A0\uC9DC\uB3C4 \uBC14\uB01C")
    rows.Add U("\uBE7C\uB294 \uC2DC\uD2B8|\uBE48 \uC2DC\uD2B8, \uC791\uC131\uBC29\uBC95 \uAC19\uC740 \uC548\uB0B4 \uC2DC\uD2B8, \uC228\uAE34 \uC2DC\uD2B8")
    rows.Add ""
    rows.Add U("#\uBB38\uC11C \uC554\uD638")
    rows.Add U("\uC5F4 \uB54C \uC554\uD638 \uBB3B\uAE30|\uC11C\uC2DD \uC2DC\uD2B8\uC758 [\uBB38\uC11C \uC554\uD638 \uC815\uD558\uAE30] \uB610\uB294 \uD55C\uC140 \uBA54\uB274\uC758 \uBB38\uC11C \uC554\uD638. \uD30C\uC77C \uB0B4\uC6A9\uC774 \uC554\uD638\uD654\uB429\uB2C8\uB2E4")
    rows.Add ""
    rows.Add U("#\uB2E8\uCD94\uAC00 \uC548 \uBCF4\uC774\uAC70\uB098 \uB20C\uB9AC\uC9C0 \uC54A\uC744 \uB54C")
    rows.Add U("Alt+F8 (\uB9E4\uD06C\uB85C \uBAA9\uB85D)|BuildA4, ImportExcel, ImportPasted, UndoLast, NewDoc, SwapSeats, HighlightSel, MakeTemplate, PrintPreviewA4, SetDocPassword")
    For k = 1 To rows.Count
        i = k - 1
        ln = rows(k)
        If Left$(ln, 1) = "#" Then
            ws.Cells(i + 1, 1).Value = Mid$(ln, 2)
            ws.Cells(i + 1, 1).Font.Bold = True
            ws.Cells(i + 1, 1).Font.Size = 12
        ElseIf ln <> "" Then
            f = Split(ln, "|")
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).NumberFormat = "@"
            ws.Cells(i + 1, 1).Value = f(0)
            If UBound(f) >= 1 Then ws.Cells(i + 1, 2).Value = f(1)
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).Borders.LineStyle = XL_CONT
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).Borders.Color = COL_INLINE
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).WrapText = True
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).VerticalAlignment = XL_TOP
            If f(0) = U("\uC785\uB825") Then ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).Interior.Color = COL_HEAD
        End If
    Next
End Sub

' ---- tests (LibreOffice) -------------------------------------------------
Public Function TestImport() As String
    gQuiet = True
    gDbg = ""
    On Error GoTo fail
    TestImport = ImportFiles(Array(Environ$("HC_TEST_XLSX"))) & " | " & gDbg
    gQuiet = False
    Exit Function
fail:
    gQuiet = False
    TestImport = "ERROR " & Err.Number & ": " & Err.Description & " [" & gStep & "] [" & gTrace & "] " & gDbg
End Function

Public Function TestBuild() As String
    gQuiet = True
    DoBuild
    gQuiet = False
    TestBuild = "fs=" & gFS & " " & gTrace
End Function

' exercises the buttons without dialogs (LibreOffice test)
Public Function TestAll() As String
    Dim r As String, ws As Object, p As String
    gQuiet = True
    gDbg = ""
    On Error GoTo fail
    gStep = "swap"
    LoadExecs
    SwapSeatsIJ 1, 4
    LoadExecs
    r = r & "swap1=" & exName(1) & "/" & exName(4) & "; "
    SwapSeatsIJ 1, 4
    LoadExecs
    r = r & "back=" & exName(1) & "; "
    gStep = "highlight"
    r = r & "mark=" & MarkLines(U("09:00 \uD68C\uC758 @\uB300\uD68C\uC758\uC2E4") & vbLf & U("* \uCC38\uC11D: \uC2E4\uC7A5") & vbLf & U("\u203B \uBA54\uBAA8")) & "; "
    gStep = "newdoc"
    NewDoc
    LoadExecs
    r = r & "afterNew items1=" & exItems(1).Count & "; "
    gStep = "undo"
    UndoLast
    LoadExecs
    r = r & "afterUndo items1=" & exItems(1).Count & "; "
    gStep = "paste"
    Set ws = HostBook().Worksheets(SH_PASTE)
    ws.Range("A3:C3").Value = Array(U("\uC131\uBA85"), U("\uC2DC\uAC04"), U("\uB0B4\uC6A9"))
    ws.Range("A4:C4").Value = Array(U("\uD55C\uC18C\uC601"), "10:00", U("\uBD99\uC5EC\uB123\uAE30 \uC2DC\uD5D8"))
    ImportPasted
    LoadExecs
    r = r & "pasted8=" & exItems(8).Count & " pasteLeft=" & ws.Range("A4").Value & "; "
    UndoLast
    gStep = "template"
    p = Environ$("HC_TEST_OUT")
    If p <> "" Then r = r & "template=" & MakeTemplateTo(p) & "; "
    gStep = "landscape"
    NSet "s_orient", U("A4 \uAC00\uB85C")
    NSet "s_tables", U("\uC544\uB2C8\uC624")
    NSet "s_secret", U("\uC544\uB2C8\uC624")
    DoBuild
    r = r & "landscape status=" & NStr("d_status") & "; "
    NSet "s_orient", U("A4 \uC138\uB85C")
    NSet "s_tables", U("\uC608")
    NSet "s_secret", U("\uC608")
    NSet "s_h1", 1: NSet "s_h2", 1: NSet "s_h3", 1
    DoBuild
    r = r & "rows set status=" & NStr("d_status") & "; "
    PresetAutoH
    r = r & "auto status=" & NStr("d_status")
    gQuiet = False
    TestAll = r
    Exit Function
fail:
    gQuiet = False
    TestAll = r & " ERROR " & Err.Number & ": " & Err.Description & " [" & gStep & "] [" & gTrace & "] " & gDbg
End Function

'
Public Sub CheckHcExecSetup()
    MsgBox "HcExecSetup OK"
End Sub
