Attribute VB_Name = "HcDaily"
Option Explicit
' ===========================================================================
' HcDaily - 부서별 일일 예정사항 (A4 한 장 자동 맞춤) for Hancell
'   sheets: 입력, A4, one sheet per department (tab order = print order),
'           서식, 입력방법, 붙여넣기, hidden _측정 / _백업
'   department sheet: row 1  부서 | name | 인쇄 | 예/아니오 | 상태 | (import info)
'                     row 2  시간 | 내용 | 장소 | 비고 | 설명 | 링크
'                     row 3+ one schedule per row; tables after a blank row
' ===========================================================================

Public Const COL_LINE As Long = 2236962       ' #222222
Public Const COL_HEAD As Long = 15526375      ' #e7e9ec
Public Const COL_DEPT As Long = 16184820      ' #f4f5f6
Public Const COL_GRID As Long = 9079434       ' #8a8a8a
Public Const COL_HALF As Long = 12895428      ' #c4c4c4
Public Const COL_QLINE As Long = 10066329     ' #999999
Public Const COL_QTEXT As Long = 3355443      ' #333333
Public Const COL_EMPTY As Long = 7829367      ' #777777
Public Const COL_TBORDER As Long = 10526880   ' #a0a0a0
Public Const COL_THEAD As Long = 15987184     ' #f0f1f3
Public Const COL_APPROVAL As Long = 15921906  ' #f2f2f2
Public Const COL_MARKD As Long = 8447743      ' #ffe680
Public Const COL_LINK As Long = 12076824      ' #1847b8
Public Const COL_GRAY As Long = 9079434
Public Const COL_SOFT As Long = 16381942
Public Const COL_INLINE As Long = 14011590
Public Const COL_ACCENT As Long = 12932898
Public Const COL_OK As Long = 4422942
Public Const COL_BAD As Long = 2832066
Public Const CELLPAD As Double = 1.5

Public Function SH_IN() As String
    SH_IN = "입력"
End Function
Public Function SH_A4() As String
    SH_A4 = "A4"
End Function
Public Function SH_FMT() As String
    SH_FMT = "서식"
End Function
Public Function SH_HELP() As String
    SH_HELP = "입력방법"
End Function
Public Function SH_PASTE() As String
    SH_PASTE = "붙여넣기"
End Function
Public Function SH_MS() As String
    SH_MS = "_측정"
End Function
Public Function SH_BAK() As String
    SH_BAK = "_백업"
End Function

' ---- settings --------------------------------------------------------------
Public gLayout As String          ' table | table2 | block | block2
Public gHalves As Boolean, gLandscape As Boolean, gMargin As Double, gFont As String
Public gAuto As Boolean, gMinF As Double, gMaxF As Double, gManFont As Double, gManGap As Double
Public gBullet As String, gTimeBold As Boolean, gFill As Boolean
Public gColHead As Boolean, gColH1 As String, gColH2 As String
Public gEmptyHide As Boolean, gEmptyText As String, gApproval As Boolean, gApLabels As String
Public gTitle As String, gDate As String, gSub As String, gFoot As String
Public gW As Double, gH As Double

' ---- departments -----------------------------------------------------------
Public dN As Long
Public dWs() As Object
Public dName() As String, dPrint() As Boolean, dInfo() As String, dLines() As Long
Public dBlocks() As Object, dDetails() As Object

' ---- current fit -------------------------------------------------------------
Public gFS As Double, gU As Double, gGapPt As Double, gPadV As Double, gPadH As Double

' ===========================================================================
' Buttons
' ===========================================================================
Public Sub BuildA4()
    On Error GoTo fail
    BeginWork
    DoBuild
    EndWork
    On Error Resume Next
    HostBook().Worksheets(SH_A4).Activate
    Exit Sub
fail:
    EndWork
    Info "A4를 만들다가 문제가 생겼습니다." & vbLf & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub GoInput()
    On Error Resume Next
    HostBook().Worksheets(SH_IN).Activate
End Sub

Public Sub PrintPreviewA4()
    On Error Resume Next
    HostBook().Worksheets(SH_A4).Activate
    HostBook().Worksheets(SH_A4).PrintPreview
    If Err.Number <> 0 Then Info "인쇄 미리보기를 열 수 없습니다. A4 시트에서 Ctrl+P로 인쇄하세요."
End Sub

' ===========================================================================
' Settings
' ===========================================================================
Public Sub LoadSettings()
    Dim s As String
    s = NStr("s_layout")
    If s = "표 2단" Then
        gLayout = "table2"
    ElseIf s = "제목형" Then
        gLayout = "block"
    ElseIf s = "제목형 2단" Then
        gLayout = "block2"
    Else
        gLayout = "table"
    End If
    gHalves = (NStr("s_split") <> "한 칸")
    gLandscape = (NStr("s_orient") = "A4 가로")
    gMargin = NNum("s_margin", 12, 3, 30)
    gFont = NStr("s_font")
    If gFont = "" Then gFont = "맑은 고딕"
    gAuto = NBool("s_auto", True)
    gMinF = NNum("s_min", 7, 4, 16)
    gMaxF = NNum("s_max", 13, 6, 24)
    gManFont = NNum("s_mfont", 10.5, 4, 24)
    gManGap = NNum("s_mgap", 0.25, 0, 1.5)
    s = NStr("s_bullet")
    If InStr(s, "□") > 0 Then
        gBullet = "report2"
    ElseIf InStr(s, "•") > 0 Then
        gBullet = "dot"
    ElseIf InStr(s, "그대로") > 0 Then
        gBullet = "keep"
    Else
        gBullet = "report"
    End If
    gTimeBold = NBool("s_timebold", True)
    gFill = NBool("s_fill", True)
    gColHead = NBool("s_colhead", True)
    gColH1 = NStr("s_colh1")
    gColH2 = NStr("s_colh2")
    gEmptyHide = (NStr("s_emptymode") = "숨기기")
    gEmptyText = NStr("s_empty")
    gApproval = NBool("s_approval", False)
    gApLabels = NStr("s_aplabels")
    gTitle = NStr("d_title")
    gDate = DateValueIso(NVal("d_date"))
    If gDate = "" Then gDate = IsoOf(Date)
    gSub = NStr("d_sub")
    gFoot = NStr("d_foot")
    If gLandscape Then
        gW = (297 - 2 * gMargin) * PT_MM
        gH = (210 - 2 * gMargin) * PT_MM
    Else
        gW = (210 - 2 * gMargin) * PT_MM
        gH = (297 - 2 * gMargin) * PT_MM
    End If
End Sub

' ===========================================================================
' Department sheets
' ===========================================================================
Public Function IsDeptSheet(ByVal ws As Object) As Boolean
    On Error Resume Next
    IsDeptSheet = (CStr(ws.Cells(1, 1).Value) = "부서" And CStr(ws.Cells(1, 3).Value) = "인쇄")
End Function

Public Function DeptSheetList() As Collection
    Dim c As New Collection, ws As Object
    For Each ws In HostBook().Worksheets
        If IsDeptSheet(ws) Then c.Add ws
    Next
    Set DeptSheetList = c
End Function

Public Function UsedLast(ByVal ws As Object, ByRef lastR As Long, ByRef lastC As Long) As Boolean
    Dim ur As Object
    Set ur = ws.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    UsedLast = True
End Function

Public Function NormName(ByVal s As String) As String
    NormName = LCase$(RemoveChars(s, ""))
End Function

' index (1-based) of the department with this name, 0 when none
Public Function FindDept(ByVal nm As String) As Long
    Dim k As String, i As Long
    k = NormName(nm)
    If k = "" Then Exit Function
    For i = 1 To dN
        If NormName(dName(i)) = k Then
            FindDept = i
            Exit Function
        End If
    Next
End Function

' department whose name appears inside the text (longest name wins)
Public Function FindDeptInText(ByVal s As String) As Long
    Dim k As String, i As Long, n As String, best As Long, bestLen As Long
    k = NormName(s)
    For i = 1 To dN
        n = NormName(dName(i))
        If Len(n) >= 2 And InStr(k, n) > 0 And Len(n) > bestLen Then
            best = i
            bestLen = Len(n)
        End If
    Next
    FindDeptInText = best
End Function

Public Sub LoadDepts(ByVal parse As Boolean)
    Dim c As Collection, i As Long, ws As Object, lines As Collection, res As Variant, bc As Object, dc As Object
    Set c = DeptSheetList()
    dN = c.Count
    If dN = 0 Then Exit Sub
    ReDim dWs(1 To dN): ReDim dName(1 To dN): ReDim dPrint(1 To dN): ReDim dInfo(1 To dN): ReDim dLines(1 To dN)
    ReDim dBlocks(1 To dN): ReDim dDetails(1 To dN)
    For i = 1 To dN
        Set ws = c(i)
        Set dWs(i) = ws
        dName(i) = CellText(ws.Cells(1, 2).Value)
        dPrint(i) = Not (CellText(ws.Cells(1, 4).Value) = "아니오" Or CellText(ws.Cells(1, 4).Value) = "아니요")
        dInfo(i) = CellText(ws.Cells(1, 6).Value)
        If parse Then
            gStep = "dept read " & i
            Set lines = DeptLines(ws)
            dLines(i) = lines.Count
            gStep = "dept parse " & i
            res = ParseBlocks(lines)
            Set bc = res(0)
            Set dc = res(1)
            Set dBlocks(i) = bc
            Set dDetails(i) = dc
        End If
    Next
End Sub

Public Function UniqueSheetName(ByVal want As String, ByVal self As Object) As String
    Dim nm As String, n As Long, ws As Object, clash As Boolean, selfName As String
    If Not self Is Nothing Then selfName = self.Name
    want = SafeSheetName(want)
    If want = "" Then want = "부서"
    nm = want
    n = 2
    Do
        clash = False
        For Each ws In HostBook().Worksheets
            If ws.Name <> selfName Then
                If LCase$(ws.Name) = LCase$(nm) Then clash = True
            End If
        Next
        If Not clash Then Exit Do
        nm = Left$(want, 27) & "(" & n & ")"
        n = n + 1
    Loop
    UniqueSheetName = nm
End Function

Public Sub SyncDeptNames()
    Dim i As Long, want As String
    On Error Resume Next
    For i = 1 To dN
        want = dName(i)
        If want = "" Then want = "부서" & i
        want = UniqueSheetName(want, dWs(i))
        If dWs(i).Name <> want Then dWs(i).Name = want
    Next
    On Error GoTo 0
End Sub

' cell text that keeps leading spaces (markdown indentation)
Public Function CellTextKeepLead(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Replace(NormNL(CStr(v)), ChrW(160), " ")
        CellTextKeepLead = RTrimX(s)
    Else
        CellTextKeepLead = CellText(v)
    End If
End Function

' department sheet -> markdown lines
Public Function DeptLines(ByVal ws As Object) As Collection
    Dim out As New Collection, lastR As Long, lastC As Long, v As Variant, r As Long, c As Long, afterBlank As Boolean
    Dim cells() As String, heads() As String, anyV As Boolean, tbEnd As Long, l As Variant
    UsedLast ws, lastR, lastC
    If lastC < 6 Then lastC = 6
    If lastR < 3 Then
        Set DeptLines = out
        Exit Function
    End If
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim heads(0 To lastC - 1)
    For c = 1 To lastC
        heads(c - 1) = CellText(v(2, c))
    Next
    r = 3
    Do While r <= lastR
        ReDim cells(0 To lastC - 1)
        anyV = False
        For c = 1 To lastC
            If c = 2 Then
                cells(c - 1) = CellTextKeepLead(v(r, c))
            Else
                cells(c - 1) = CellText(v(r, c), c = 1)
            End If
            If TrimX(cells(c - 1)) <> "" Then anyV = True
        Next
        If Not anyV Then
            afterBlank = True
        Else
            tbEnd = 0
            If afterBlank Then tbEnd = TableBlockLines(v, r, lastR, lastC, out)
            afterBlank = False
            If tbEnd > 0 Then
                r = tbEnd - 1
            Else
                RecordToLines cells, heads, out
            End If
        End If
        r = r + 1
    Loop
    Set DeptLines = out
End Function

' ---- a table inside a department sheet: [title], header row, body rows up to the next blank row
Private Function RowFilled(v As Variant, ByVal r As Long, ByVal lastC As Long) As Long
    Dim c As Long, n As Long
    For c = 1 To lastC
        If CellText(v(r, c)) <> "" Then n = n + 1
    Next
    RowFilled = n
End Function

Private Function LooksTableHeader(v As Variant, ByVal r As Long, ByVal lastR As Long, ByVal lastC As Long) As Boolean
    Dim c As Long, n As Long, t As String
    If r > lastR Then Exit Function
    If CellText(v(r, 1)) = "" Then Exit Function
    If DTimeLen(CellText(v(r, 1))) > 0 Then Exit Function
    For c = 1 To lastC
        t = CellText(v(r, c))
        If t <> "" Then
            n = n + 1
            If Len(t) > 30 Or InStr(t, vbLf) > 0 Then Exit Function
        End If
    Next
    LooksTableHeader = (n >= 2)
End Function

' returns the row after the table (0 when no table starts at r)
Private Function TableBlockLines(v As Variant, ByVal r As Long, ByVal lastR As Long, ByVal lastC As Long, ByVal out As Collection) As Long
    Dim h As Long, e As Long, title As String, maxC As Long, rr As Long, c As Long, ln As String
    h = r
    If RowFilled(v, r, lastC) = 1 And CellText(v(r, 1)) <> "" Then
        If LooksTableHeader(v, r + 1, lastR, lastC) Then
            title = CellText(v(r, 1))
            h = r + 1
        End If
    End If
    If Not LooksTableHeader(v, h, lastR, lastC) Then Exit Function
    e = h + 1
    Do While e <= lastR
        If RowFilled(v, e, lastC) = 0 Then Exit Do
        e = e + 1
    Loop
    If e - h < 2 Then Exit Function
    For rr = h To e - 1
        For c = 1 To lastC
            If CellText(v(rr, c)) <> "" And c > maxC Then maxC = c
        Next
    Next
    If title <> "" Then out.Add "### " & title
    For rr = h To e - 1
        ln = "|"
        For c = 1 To maxC
            ln = ln & " " & Replace(Replace(CellText(v(rr, c)), "|", "/"), vbLf, " ") & " |"
        Next
        out.Add ln
        If rr = h Then
            ln = "|"
            For c = 1 To maxC
                ln = ln & "---|"
            Next
            out.Add ln
        End If
    Next
    TableBlockLines = e
End Function

' does a 내용 cell hold markdown that should be kept as it is?
Private Function IsRawMarkdown(ByVal first As String) As Boolean
    Dim t As String, kind As String, mk As String, rank As Double, txt As String, raw As String
    If Len(first) = 0 Then Exit Function
    If IsWs(Left$(first, 1)) Then
        IsRawMarkdown = True
        Exit Function
    End If
    t = TrimX(first)
    If Left$(t, 1) = "#" Or Left$(t, 1) = ">" Or Left$(t, 1) = "|" Or Left$(t, 1) = "》" Then
        IsRawMarkdown = True
    ElseIf IsHrLine(t) Then
        IsRawMarkdown = True
    ElseIf MatchListItem(t, kind, mk, rank, txt, raw) Then
        IsRawMarkdown = True
    End If
End Function

Public Function SafeHref(ByVal uu As String) As String
    Dim l As String
    uu = TrimX(uu)
    l = LCase$(uu)
    If StartsWith(l, "www.") Then
        uu = "http://" & uu
        l = "http://" & l
    End If
    If StartsWith(l, "http:") Or StartsWith(l, "https:") Or StartsWith(l, "file:") Or StartsWith(l, "ftp:") Or StartsWith(l, "mailto:") Then
        SafeHref = uu
    ElseIf StartsWith(uu, "\\") Then
        SafeHref = "file:" & Replace(uu, "\", "/")
    ElseIf Len(uu) > 2 And Mid$(uu, 2, 1) = ":" And (Mid$(uu, 3, 1) = "\" Or Mid$(uu, 3, 1) = "/") Then
        SafeHref = "file:///" & Replace(uu, "\", "/")
    End If
End Function

' one row (시간 내용 장소 비고 설명 링크 + extra columns) -> markdown lines (port of recordToLines)
Public Sub RecordToLines(cells() As String, heads() As String, ByVal out As Collection)
    Dim tm As String, content As String, place As String, link As String, detail As String
    Dim extras As New Collection, c As Long, l As Variant, body As Variant, first As String, i As Long
    Dim url As String, ln As String, plain As Boolean, lab As String, ev As Variant, parts As Variant
    tm = DNormTime(TrimX(cells(0)))
    content = cells(1)
    place = Replace(TrimX(cells(2)), vbLf, " ")
    If TrimX(cells(3)) <> "" Then extras.Add Array("", cells(3))
    detail = TrimX(cells(4))
    link = TrimX(cells(5))
    For c = 6 To UBound(cells)
        If TrimX(cells(c)) <> "" Then extras.Add Array(heads(c), cells(c))
    Next
    If tm = "" And TrimX(content) = "" And place = "" And extras.Count = 0 And detail = "" And link = "" Then Exit Sub
    ' markdown typed straight into 내용 (no other columns): keep as it is
    If tm = "" And place = "" And extras.Count = 0 And detail = "" And link = "" Then
        body = Split(content, vbLf)
        If IsRawMarkdown(CStr(body(0))) Then
            For i = 0 To UBound(body)
                out.Add RTrimX(CStr(body(i)))
            Next
            Exit Sub
        End If
    End If
    body = Split(TrimX(content), vbLf)
    first = ""
    If UBound(body) >= 0 Then first = TrimX(CStr(body(0)))
    If link <> "" Then
        parts = Split(CollapseWs(link), " ")
        url = parts(0)
        If SafeHref(url) = "" Then
            extras.Add Array("링크", link)
            url = ""
        End If
    End If
    If Left$(first, 1) = "*" And TrimX(Mid$(first, 2)) <> "" Then
        ln = " * " & TrimX(tm & " " & TrimX(Mid$(first, 2)))
        If url <> "" Then ln = ln & " " & url
        out.Add ln
    ElseIf StartsWith(first, ">>") Or Left$(first, 1) = "》" Then
        out.Add "  >> " & TrimX(Mid$(first, IIf(Left$(first, 1) = "》", 2, 3)))
    Else
        first = StripLeadBullet(first)
        ln = "- " & TrimX(tm & " " & first)
        If url <> "" Then ln = ln & " " & url
        If place <> "" Then ln = ln & " @" & place
        out.Add ln
    End If
    For Each l In Split(detail, vbLf)
        If TrimX(CStr(l)) <> "" Then out.Add "  >> " & TrimX(CStr(l))
    Next
    For i = 1 To UBound(body)
        If TrimX(CStr(body(i))) <> "" Then out.Add "  " & TrimX(CStr(body(i)))
    Next
    For Each ev In extras
        lab = ev(0)
        For Each l In Split(CStr(ev(1)), vbLf)
            If TrimX(CStr(l)) <> "" Then
                plain = (lab = "" Or lab = "비고" Or lab = "참고" Or lab = "기타" Or LCase$(lab) = "note" Or LCase$(lab) = "notes" Or StartsWith(TrimX(CStr(l)), lab))
                If plain Then
                    out.Add " * " & TrimX(CStr(l))
                Else
                    out.Add " * " & lab & ": " & TrimX(CStr(l))
                End If
            End If
        Next
    Next
End Sub

Private Function StripLeadBullet(ByVal s As String) As String
    If FirstIn(s, "-+•·ㆍ∙◦▪▫○●◎◇◆▶▷►□■") Then s = TrimX(Mid$(s, 2))
    StripLeadBullet = s
End Function

' 9:30 -> 09:30, fractions of a day -> hh:mm
Public Function DNormTime(ByVal v As String) As String
    v = Replace(TrimX(v), vbLf, " ")
    If v = "" Then Exit Function
    If IsNumeric(v) And InStr(v, ":") = 0 Then
        If Val(v) >= 0 And Val(v) < 1 And InStr(v, ".") > 0 Then
            DNormTime = Format$(CDate(Val(v)), "hh:mm")
            Exit Function
        End If
    End If
    If IsDigitAt(v, 1) And Mid$(v, 2, 1) = ":" And IsDigitAt(v, 3) And IsDigitAt(v, 4) Then v = "0" & v
    DNormTime = v
End Function

' length of a leading time in a line (port of RE_TIME in the HTML), 0 when none
Public Function DTimeLen(ByVal s As String) As Long
    Dim c1 As Long, e As Long, q As Long, c2 As Long, q2 As Long
    c1 = DTimePart(s, 1)
    If c1 > 0 Then
        q = SkipWs(s, c1)
        If FirstIn(Mid$(s, q, 1), "~∼〜-–—") Then
            q2 = SkipWs(s, q + 1)
            c2 = DTimePart(s, q2)
            If c2 > 0 And Not DTimeStop(s, c2) Then
                DTimeLen = c2 - 1
                Exit Function
            End If
            If FirstIn(Mid$(s, q, 1), "~∼〜") And Not DTimeStop(s, q + 1) Then
                DTimeLen = q
                Exit Function
            End If
        End If
        If Not DTimeStop(s, c1) Then
            DTimeLen = c1 - 1
            Exit Function
        End If
    End If
    If StartsWith(s, "종일") Then
        If Not DTimeStop(s, 3) Then DTimeLen = 2
        Exit Function
    End If
    If StartsWith(s, "오전") Or StartsWith(s, "오후") Then
        q = SkipWs(s, 3)
        If Mid$(s, q, 1) = "중" Then
            If Not DTimeStop(s, q + 1) Then DTimeLen = q
        End If
    End If
End Function

' next char is a digit or ':' (the (?![\d:]) look-ahead)
Private Function DTimeStop(ByVal s As String, ByVal p As Long) As Boolean
    If p > Len(s) Then Exit Function
    DTimeStop = IsDigitAt(s, p) Or Mid$(s, p, 1) = ":"
End Function

' one time token at p; returns the position after it (0 when none)
Private Function DTimePart(ByVal s As String, ByVal p As Long) As Long
    Dim q As Long, d As Long, q2 As Long, t As String
    q = p
    t = Mid$(s, q, 2)
    If t = "오전" Or t = "오후" Then
        q = SkipWs(s, q + 2)
        If Not IsDigitAt(s, q) Then q = p
    End If
    If Not IsDigitAt(s, q) Then Exit Function
    d = q + 1
    If IsDigitAt(s, d) Then d = d + 1
    If Mid$(s, d, 1) = ":" Then
        If IsDigitAt(s, d + 1) And IsDigitAt(s, d + 2) Then DTimePart = d + 3
        Exit Function
    End If
    If Mid$(s, d, 1) = "시" And Mid$(s, d + 1, 1) <> "간" Then
        q2 = SkipWs(s, d + 1)
        If IsDigitAt(s, q2) Then
            If IsDigitAt(s, q2 + 1) And Mid$(s, q2 + 2, 1) = "분" Then
                DTimePart = q2 + 3
                Exit Function
            ElseIf Mid$(s, q2 + 1, 1) = "분" Then
                DTimePart = q2 + 2
                Exit Function
            End If
        ElseIf Mid$(s, q2, 1) = "반" Then
            DTimePart = q2 + 1
            Exit Function
        End If
        DTimePart = d + 1
    End If
End Function

' ===========================================================================
' Units (one department) in the table / block layouts
' ===========================================================================
Private Function DeptW() As Double
    Dim i As Long, m As Double
    If gColHead Then m = TextWidthEm(gColH1)
    For i = 1 To dN
        If dPrint(i) Then m = MaxD(m, TextWidthEm(dName(i)))
    Next
    DeptW = (Clamp(m, 2.4, 6.2) + 1.3) * gFS
End Function

Private Function ShowDept(ByVal d As Long) As Boolean
    Dim bl As Collection, dt As Collection
    If Not dPrint(d) Then Exit Function
    Set bl = dBlocks(d)
    Set dt = dDetails(d)
    If gEmptyHide And bl.Count = 0 And dt.Count = 0 Then Exit Function
    ShowDept = True
End Function

' one table row. returns natural height
Private Function LayRowTable(ByVal d As Long, ByVal x As Double, ByVal y As Double, ByVal tw As Double, ByVal halves As Boolean, _
        ByVal rowH As Double, ByVal doRender As Boolean) As Double
    Dim dw As Double, nm As String, nh As Double, cw As Double, ch As Double, h As Double, idx As Long, note As String, d0 As Variant
    Dim dt As Collection
    dw = DeptW()
    nm = dName(d)
    If nm = "" Then nm = "(이름 없음)"
    nh = MeasureH(nm, "", dw - 2 * gPadH, gFont, gFS, True)
    cw = tw - dw
    If halves Then
        ch = LayContent(d, x + dw + gPadH - CELLPAD, y + gPadV, cw - 2 * gPadH + 2 * CELLPAD, True, 2 * gPadH, x + dw + cw / 2, rowH, False)
    Else
        ch = LayContent(d, x + dw + gPadH - CELLPAD, y + gPadV, cw - 2 * gPadH + 2 * CELLPAD, False, 0, 0, rowH, False)
    End If
    h = MaxD(nh, ch) + 2 * gPadV
    If doRender Then
        If rowH < h Then rowH = h
        idx = ElText(x, y, dw, rowH, nm, "", gFS, True, XL_CENTER, XL_CENTER)
        ElFill idx, COL_DEPT
        ElBorder idx, 3, XL_CONT, XL_THIN, COL_GRID
        Set dt = dDetails(d)
        For Each d0 In dt
            note = note & IIf(note = "", "", vbLf) & PlainInline(CStr(d0))
        Next
        If note <> "" Then ElNote idx, note
        idx = ElBox(x, y, tw, rowH)
        ElBorder idx, 1, XL_CONT, XL_THIN, COL_GRID
        If halves Then
            LayContent d, x + dw + gPadH - CELLPAD, y + gPadV, cw - 2 * gPadH + 2 * CELLPAD, True, 2 * gPadH, x + dw + cw / 2, rowH, True
        Else
            LayContent d, x + dw + gPadH - CELLPAD, y + gPadV, cw - 2 * gPadH + 2 * CELLPAD, False, 0, 0, rowH, True
        End If
    End If
    LayRowTable = h
End Function

Private Function LayHeadRow(ByVal x As Double, ByVal y As Double, ByVal tw As Double, ByVal doRender As Boolean) As Double
    Dim dw As Double, h As Double, idx As Long
    If Not gColHead Then Exit Function
    dw = DeptW()
    h = MaxD(MeasureH(IIf(gColH1 = "", "가", gColH1), "", dw - 2 * gPadH, gFont, gFS, True), _
        MeasureH(IIf(gColH2 = "", "가", gColH2), "", tw - dw - 2 * gPadH, gFont, gFS, True)) + 2 * gPadV
    If doRender Then
        idx = ElText(x, y, dw, h, gColH1, "", gFS, True, XL_CENTER, XL_CENTER)
        ElFill idx, COL_HEAD
        ElBorder idx, 3, XL_CONT, XL_THIN, COL_GRID
        ElBorder idx, 1, XL_CONT, XL_MEDIUM, 0
        idx = ElText(x + dw, y, tw - dw, h, gColH2, "", gFS, True, XL_CENTER, XL_CENTER)
        ElFill idx, COL_HEAD
        ElBorder idx, 1, XL_CONT, XL_MEDIUM, 0
    End If
    LayHeadRow = h
End Function

' one department in the 제목형 layout. returns natural height
Private Function LayBlockUnit(ByVal d As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal halves As Boolean, _
        ByVal doRender As Boolean) As Double
    Dim nm As String, nfs As Double, nh As Double, idx As Long, yy As Double, ch As Double, cx As Double, cw As Double, note As String
    Dim dt As Collection, d0 As Variant
    nm = dName(d)
    If nm = "" Then nm = "(이름 없음)"
    nfs = R5(gFS * 1.06)
    nh = MeasureH("■ " & nm, "", w, gFont, nfs, True) + 0.1 * gFS
    If doRender Then
        RBReset
        RBAdd "■ ", True, , , R5(nfs * 0.7)
        RBAdd nm, True
        idx = ElText(x, y, w, nh, RBText(), RBRuns(), nfs, True, XL_LEFT, XL_BOTTOM)
        ElBorder idx, 1, XL_CONT, XL_MEDIUM, COL_LINE
        Set dt = dDetails(d)
        For Each d0 In dt
            note = note & IIf(note = "", "", vbLf) & PlainInline(CStr(d0))
        Next
        If note <> "" Then ElNote idx, note
    End If
    yy = y + nh + 0.15 * gFS + gGapPt
    cx = x + 0.2 * gFS
    cw = w - 0.2 * gFS
    If halves Then
        ch = LayContent(d, cx, yy, cw, True, 1.4 * gFS, cx + (cw - 1.4 * gFS) / 2 + 0.7 * gFS, 0, doRender)
    Else
        ch = LayContent(d, cx, yy, cw, False, 0, 0, 0, doRender)
    End If
    LayBlockUnit = (yy - y) + ch
End Function

' ===========================================================================
' Page
' ===========================================================================
Private Function TitleFs() As Double
    TitleFs = R5(Clamp(gFS * 1.75, 15, 22))
End Function

Private Function MetaFs() As Double
    MetaFs = R5(Clamp(gFS * 0.92, 8.5, 11))
End Function

Private Function ApprovalLabels() As Variant
    Dim s As String, p As Variant, out As New Collection, ch As Variant
    s = gApLabels
    For Each ch In Array("，", "、", "/")
        s = Replace(s, CStr(ch), ",")
    Next
    For Each p In Split(s, ",")
        If TrimX(CStr(p)) <> "" And out.Count < 6 Then out.Add TrimX(CStr(p))
    Next
    ApprovalLabels = ColToArray(out)
End Function

Private Function LayHead(ByVal doRender As Boolean) As Double
    Dim apW As Double, apH As Double, labs As Variant, n As Long, j As Long, idx As Long, tW As Double, tH As Double, y As Double
    Dim mfs As Double, mH As Double, dTxt As String, rowH As Double
    mfs = MetaFs()
    If gApproval Then
        labs = ApprovalLabels()
        n = ArrLen(labs)
    End If
    If n > 0 Then
        apW = (6 + 17 * n) * PT_MM
        apH = 20 * PT_MM
    End If
    tW = gW - IIf(n > 0, apW + 4 * PT_MM, 0)
    If gTitle <> "" Then tH = MeasureH(gTitle, "", tW, gFont, TitleFs(), True)
    rowH = MaxD(tH, apH)
    If doRender Then
        If gTitle <> "" Then ElText 0, (rowH - tH) / 2, tW, tH, gTitle, "", TitleFs(), True, XL_CENTER, XL_CENTER
        If n > 0 Then
            idx = ElText(gW - apW, 0, 6 * PT_MM, apH, "결" & vbLf & "재", "", 9, False, XL_CENTER, XL_CENTER)
            ElFill idx, COL_APPROVAL
            ElBorderAll idx, XL_CONT, XL_THIN, 0
            For j = 0 To n - 1
                idx = ElText(gW - apW + (6 + 17 * j) * PT_MM, 0, 17 * PT_MM, 6 * PT_MM, CStr(labs(j)), "", 9, False, XL_CENTER, XL_CENTER)
                ElFill idx, COL_APPROVAL
                ElBorderAll idx, XL_CONT, XL_THIN, 0
                idx = ElBox(gW - apW + (6 + 17 * j) * PT_MM, 6 * PT_MM, 17 * PT_MM, 14 * PT_MM)
                ElBorderAll idx, XL_CONT, XL_THIN, 0
            Next
        End If
    End If
    y = rowH + 2 * PT_MM
    dTxt = FmtDateK(gDate)
    mH = MaxD(MeasureH(IIf(gSub = "", "가", gSub), "", gW * 0.62, gFont, mfs), MeasureH(dTxt, "", gW * 0.38, gFont, mfs, True))
    If doRender Then
        If gSub <> "" Then ElText 0, y, gW * 0.62, mH, gSub, "", mfs, False, XL_LEFT, XL_BOTTOM
        ElText gW * 0.62, y, gW * 0.38, mH, dTxt, "", mfs, True, XL_RIGHT, XL_BOTTOM
    End If
    y = y + mH
    If gLayout = "block" Or gLayout = "block2" Then
        y = y + 2 * PT_MM
        If doRender Then
            idx = ElBox(0, 0, gW, y)
            ElBorder idx, 1, XL_CONT, XL_MEDIUM, 0
        End If
    End If
    LayHead = y + 3 * PT_MM
End Function

Private Function LayFoot(ByVal y As Double, ByVal doRender As Boolean) As Double
    Dim h As Double
    If gFoot = "" Then Exit Function
    h = MeasureH(gFoot, "", gW, gFont, MetaFs())
    If doRender Then ElText 0, y + 2.5 * PT_MM, gW, h, gFoot, "", MetaFs()
    LayFoot = h + 2.5 * PT_MM
End Function

Private Sub SetVars(ByVal fs As Double, ByVal lo As Double, ByVal hi As Double)
    Dim g As Double
    If Not gAuto Then
        gFS = R5(gManFont)
        g = gManGap
        gGapPt = g * gFS
        gPadV = (0.18 + g * 0.6) * gFS
        gPadH = 0.55 * gFS
        gU = 1
        Exit Sub
    End If
    gFS = fs
    If hi > lo Then
        gU = Clamp((fs - lo) / (hi - lo), 0, 1)
    Else
        gU = 1
    End If
    gGapPt = (0.05 + 0.3 * gU) * fs
    gPadV = (0.16 + 0.34 * gU) * fs
    gPadH = (0.45 + 0.25 * gU) * fs
End Sub

' body height for the current variables; split = first department of the right column (2단)
Private Function BodyHeight(ByRef split As Long, ByRef hL As Double, ByRef hR As Double) As Double
    Dim i As Long, hs() As Double, n As Long, tot As Double, acc As Double, best As Double, k As Long, colW As Double
    Dim idxs() As Long, gapB As Double, h As Double
    ReDim hs(1 To dN + 1): ReDim idxs(1 To dN + 1)
    For i = 1 To dN
        If ShowDept(i) Then
            n = n + 1
            idxs(n) = i
        End If
    Next
    split = 0
    gapB = 0.55 * gFS + 2 * gGapPt
    Select Case gLayout
        Case "table"
            h = LayHeadRow(0, 0, gW, False)
            For k = 1 To n
                h = h + LayRowTable(idxs(k), 0, 0, gW, gHalves, 0, False)
            Next
            BodyHeight = h
        Case "block"
            For k = 1 To n
                If k > 1 Then h = h + gapB
                h = h + LayBlockUnit(idxs(k), 0, 0, gW, gHalves, False)
            Next
            BodyHeight = h
        Case Else
            colW = (gW - 5 * PT_MM) / 2
            For k = 1 To n
                If gLayout = "table2" Then
                    hs(k) = LayRowTable(idxs(k), 0, 0, colW, False, 0, False)
                Else
                    hs(k) = LayBlockUnit(idxs(k), 0, 0, colW, False, False)
                    If k > 1 Then hs(k) = hs(k) + gapB
                End If
                tot = tot + hs(k)
            Next
            best = 1E+30
            split = n + 1
            If n > 1 Then
                acc = 0
                For k = 2 To n
                    acc = acc + hs(k - 1)
                    If MaxD(acc, tot - acc) <= best Then
                        best = MaxD(acc, tot - acc)
                        split = k
                    End If
                Next
            End If
            hL = 0: hR = 0
            For k = 1 To n
                If k < split Then
                    hL = hL + hs(k)
                Else
                    hR = hR + hs(k)
                End If
            Next
            If gLayout = "block2" And split <= n Then hR = hR - gapB
            If gLayout = "table2" Then
                hL = hL + LayHeadRow(0, 0, colW, False)
                If split <= n Then hR = hR + LayHeadRow(0, 0, colW, False)
            End If
            BodyHeight = MaxD(hL, hR)
            If split > n Then
                split = dN + 1
            Else
                split = idxs(split)
            End If
    End Select
End Function

Private Sub FitEval(ByVal fs As Double, ByVal lo As Double, ByVal hi As Double, ByRef ok As Boolean, ByRef need As Double, _
        ByRef avail As Double, ByRef headH As Double, ByRef footH As Double, ByRef split As Long)
    Dim hL As Double, hR As Double
    SetVars fs, lo, hi
    headH = LayHead(False)
    footH = LayFoot(0, False)
    avail = gH - headH - footH
    need = BodyHeight(split, hL, hR)
    ok = (need <= avail - 3)
End Sub

Public Sub DoBuild()
    Dim lo As Double, hi As Double, fs As Double, ok As Boolean, need As Double, avail As Double, headH As Double, footH As Double
    Dim split As Long, a As Long, b As Long, md As Long, i As Long, n As Long, y As Double, colW As Double, extra As Double
    Dim ws As Object, lastR As Long, lastC As Long, st As String, k As Long, yL As Double, yR As Double, x As Double
    Dim units() As Long, nU As Long, gapB As Double, rowH As Double, nL As Long, nR2 As Long, tableTop As Double
    Dim xi As Long, colTop As Double, colY As Double, colNeed As Double, cnt As Long, cx As Double
    gStep = "settings"
    LoadSettings
    gStep = "departments"
    LoadDepts True
    SyncDeptNames
    gStep = "measure init"
    SetScreen True
    MsInit SH_MS
    gStep = "fit"
    lo = R5(MinD(gMinF, gMaxF)): hi = R5(MaxD(gMinF, gMaxF))
    If Not gAuto Then
        FitEval gManFont, lo, hi, ok, need, avail, headH, footH, split
    Else
        FitEval hi, lo, hi, ok, need, avail, headH, footH, split
        If Not ok Then
            FitEval lo, lo, hi, ok, need, avail, headH, footH, split
            If ok Then
                a = CLng(lo * 2): b = CLng(hi * 2)
                Do While b - a > 1
                    md = (a + b) \ 2
                    FitEval md / 2, lo, hi, ok, need, avail, headH, footH, split
                    If ok Then
                        a = md
                    Else
                        b = md
                    End If
                Loop
                FitEval a / 2, lo, hi, ok, need, avail, headH, footH, split
            End If
        End If
    End If
    fs = gFS
    ' draw
    gStep = "render"
    SetScreen False
    ElReset
    y = LayHead(True)
    ReDim units(1 To dN + 1)
    For i = 1 To dN
        If ShowDept(i) Then
            nU = nU + 1
            units(nU) = i
        End If
    Next
    gapB = 0.55 * gFS + 2 * gGapPt
    Select Case gLayout
        Case "table"
            extra = 0
            If ok And gFill And nU > 0 Then extra = MaxD(0, (avail - 1 - need) / nU)
            tableTop = y
            y = y + LayHeadRow(0, y, gW, True)
            For k = 1 To nU
                rowH = LayRowTable(units(k), 0, y, gW, gHalves, 0, False) + extra
                LayRowTable units(k), 0, y, gW, gHalves, rowH, True
                y = y + rowH
            Next
            If nU > 0 Then TableFrame 0, tableTop, gW, y - tableTop
        Case "block"
            For k = 1 To nU
                If k > 1 Then y = y + gapB
                y = y + LayBlockUnit(units(k), 0, y, gW, gHalves, True)
            Next
        Case Else
            colW = (gW - 5 * PT_MM) / 2
            yL = y: yR = y
            For k = 1 To nU
                If units(k) < split Then
                    nL = nL + 1
                Else
                    nR2 = nR2 + 1
                End If
            Next
            For xi = 0 To 1
                If xi = 0 Then
                    n = nL
                Else
                    n = nR2
                End If
                cx = xi * (colW + 5 * PT_MM)
                If n > 0 Then
                    If gLayout = "table2" Then
                        colTop = y
                        colY = y + LayHeadRow(cx, y, colW, True)
                        colNeed = LayHeadRow(0, 0, colW, False)
                        For k = 1 To nU
                            If (xi = 0 And units(k) < split) Or (xi = 1 And units(k) >= split) Then colNeed = colNeed + LayRowTable(units(k), 0, 0, colW, False, 0, False)
                        Next
                        extra = 0
                        If ok And gFill Then extra = MaxD(0, (avail - 1 - colNeed) / n)
                        For k = 1 To nU
                            If (xi = 0 And units(k) < split) Or (xi = 1 And units(k) >= split) Then
                                rowH = LayRowTable(units(k), cx, colY, colW, False, 0, False) + extra
                                LayRowTable units(k), cx, colY, colW, False, rowH, True
                                colY = colY + rowH
                            End If
                        Next
                        TableFrame cx, colTop, colW, colY - colTop
                    Else
                        colY = y
                        cnt = 0
                        For k = 1 To nU
                            If (xi = 0 And units(k) < split) Or (xi = 1 And units(k) >= split) Then
                                If cnt > 0 Then colY = colY + gapB
                                colY = colY + LayBlockUnit(units(k), cx, colY, colW, False, True)
                                cnt = cnt + 1
                            End If
                        Next
                    End If
                End If
            Next
            y = y + need
    End Select
    If nU = 0 Then ElText 0, y + 20 * PT_MM, gW, 20, "표시할 부서가 없습니다. 부서 시트를 추가하세요.", "", 12, False, XL_CENTER, XL_TOP, COL_EMPTY
    If ok And gFoot <> "" Then y = gH - footH
    LayFoot y, True
    Set ws = HostBook().Worksheets(SH_A4)
    gStep = "realize"
    Realize ws, gFont, lastR, lastC
    gStep = "page setup"
    PageSetupA4 ws, gLandscape, gMargin, lastR, lastC
    gStep = "status"
    If ok Then
        If gAuto Then
            st = "A4 한 장 자동 맞춤 · 글자 " & NumStr(fs) & "pt"
        Else
            st = "A4 한 장 (수동) · 글자 " & NumStr(fs) & "pt"
        End If
        If gHalves And (gLayout = "table" Or gLayout = "block") Then st = st & " · 예정사항 좌우 2칸"
        If gAuto And fs >= hi Then st = st & " · 최대 글자크기 · 여유 " & Format$((avail - need) / PT_MM, "0") & "mm"
    Else
        st = "A4 한 장을 " & Format$(-Int(-(need - avail) / PT_MM), "0") & "mm 넘습니다 (글자 " & NumStr(fs) & "pt). "
        If gAuto Then
            st = st & "내용을 줄이거나, 서식 시트에서 최소 글자크기를 낮추거나, 2단·가로로 바꿔 보세요. 지금 인쇄하면 한 장에 맞게 줄여 찍힙니다."
        Else
            st = st & "글자크기·항목 간격을 줄이거나 자동 맞춤을 켜세요."
        End If
    End If
    ShowStatus st, ok
    gStep = "list"
    UpdateDeptList
    PlaceA4Buttons ws, lastC
    gStep = "done"
End Sub

' thick lines above and below a table
Private Sub TableFrame(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim idx As Long
    idx = ElBox(x, y, w, h)
    ElBorder idx, 0, XL_CONT, XL_MEDIUM, 0
    ElBorder idx, 1, XL_CONT, XL_MEDIUM, 0
End Sub

Private Sub ShowStatus(ByVal st As String, ByVal ok As Boolean)
    Dim r As Object, ws As Object, shp As Object
    Set r = NameRange("d_status")
    If Not r Is Nothing Then
        r.Cells(1, 1).Value = IIf(ok, "맞춤 완료: ", "넘침: ") & st
        r.Cells(1, 1).Font.Color = IIf(ok, COL_OK, COL_BAD)
    End If
    On Error Resume Next
    Set ws = HostBook().Worksheets(SH_A4)
    Set shp = Nothing
    Set shp = ws.Shapes("a4status")
    If Not shp Is Nothing Then
        shp.TextFrame.Characters.Text = IIf(ok, "맞춤 완료: ", "넘침: ") & st
        shp.TextFrame.Characters.Font.Color = IIf(ok, COL_OK, COL_BAD)
    End If
    On Error GoTo 0
End Sub

Private Sub PlaceA4Buttons(ByVal ws As Object, ByVal lastC As Long)
    ' by name: walking every shape fails in some hosts once cells have notes
    Dim x As Double, y As Double, nm As Variant, shp As Object
    On Error Resume Next
    x = ws.Cells(1, lastC + 1).Left + 18
    Set shp = Nothing
    Set shp = ws.Shapes("a4status")
    If Not shp Is Nothing Then
        shp.Left = x
        shp.Top = 6
    End If
    y = 86
    For Each nm In Array("btn_BuildA4", "btn_PrintPreviewA4", "btn_SavePdf", "btn_GoInput")
        Set shp = Nothing
        Set shp = ws.Shapes(CStr(nm))
        If Not shp Is Nothing Then
            shp.Left = x
            shp.Top = y
            y = y + shp.Height + 6
        End If
    Next
    On Error GoTo 0
End Sub

' department list and input status on the 입력 sheet
Public Sub UpdateDeptList()
    Dim r As Object, ws As Object, i As Long, got As Long, self As Long, miss As String, nVis As Long, row As Long, st As String
    Dim lastRow As Long, bl As Collection
    Set r = NameRange("d_list")
    If r Is Nothing Then Exit Sub
    Set ws = r.Worksheet
    On Error Resume Next
    lastRow = r.Row + 200
    ws.Range(ws.Cells(r.Row, r.Column), ws.Cells(lastRow, r.Column + 4)).Hyperlinks.Delete
    ws.Range(ws.Cells(r.Row, r.Column), ws.Cells(lastRow, r.Column + 4)).Clear
    On Error GoTo 0
    For i = 1 To dN
        row = r.Row + i - 1
        Set bl = dBlocks(i)
        If dInfo(i) <> "" Then
            st = dInfo(i)
        ElseIf dLines(i) > 0 Then
            st = "직접 입력"
        Else
            st = "미입력"
        End If
        If dPrint(i) Then
            nVis = nVis + 1
            If dInfo(i) <> "" Then
                got = got + 1
            ElseIf dLines(i) > 0 Then
                self = self + 1
            Else
                miss = miss & IIf(miss = "", "", ", ") & IIf(dName(i) = "", "(이름 없음)", dName(i))
            End If
        End If
        ws.Cells(row, r.Column).Value = i
        ws.Cells(row, r.Column + 1).NumberFormat = "@"
        ws.Cells(row, r.Column + 1).Value = dName(i)
        AddLink ws, ws.Cells(row, r.Column + 1), "#'" & dWs(i).Name & "'!B3", COL_ACCENT
        ws.Cells(row, r.Column + 2).Value = IIf(dPrint(i), "인쇄", "인쇄 안 함")
        ws.Cells(row, r.Column + 3).Value = dLines(i) & "줄"
        ws.Cells(row, r.Column + 4).NumberFormat = "@"
        ws.Cells(row, r.Column + 4).Value = st
        ws.Cells(row, r.Column + 4).Font.Color = IIf(st = "미입력", COL_BAD, IIf(st = "직접 입력", 0, COL_OK))
        ws.Range(ws.Cells(row, r.Column), ws.Cells(row, r.Column + 4)).Font.Size = 9
        ws.Range(ws.Cells(row, r.Column), ws.Cells(row, r.Column + 4)).Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONT
        ws.Range(ws.Cells(row, r.Column), ws.Cells(row, r.Column + 4)).Borders(XL_EDGE_BOTTOM).Color = COL_INLINE
    Next
    NSet "d_recv", "입력 현황: 엑셀에서 " & got & " · 직접 입력 " & self & " · 비어 있음 " & (nVis - got - self) & " / 전체 " & nVis & "곳"
    If miss = "" Then
        NSet "d_miss", IIf(nVis > 0, "모든 부서 내용이 들어왔습니다.", "")
    Else
        NSet "d_miss", FmtDateK(gDate) & " 일일 예정사항 엑셀 미입력: " & miss
    End If
End Sub

' 이 모듈을 한셀이 읽었는지 확인 (실행했을 때 창이 뜨면 정상)
Public Sub CheckHcDaily()
    MsgBox "HcDaily OK"
End Sub
