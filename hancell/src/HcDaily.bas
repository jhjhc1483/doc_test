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
Private Const CELLPAD As Double = 1.5

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
    ThisWorkbook.Worksheets(SH_A4).Activate
    Exit Sub
fail:
    EndWork
    Info "A4를 만들다가 문제가 생겼습니다." & vbLf & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub GoInput()
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_IN).Activate
End Sub

Public Sub PrintPreviewA4()
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_A4).Activate
    ThisWorkbook.Worksheets(SH_A4).PrintPreview
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
    For Each ws In ThisWorkbook.Worksheets
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
        For Each ws In ThisWorkbook.Worksheets
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

Private Function SafeHref(ByVal u As String) As String
    Dim l As String
    u = TrimX(u)
    l = LCase$(u)
    If StartsWith(l, "www.") Then
        u = "http://" & u
        l = "http://" & l
    End If
    If StartsWith(l, "http:") Or StartsWith(l, "https:") Or StartsWith(l, "file:") Or StartsWith(l, "ftp:") Or StartsWith(l, "mailto:") Then
        SafeHref = u
    ElseIf StartsWith(u, "\\") Then
        SafeHref = "file:" & Replace(u, "\", "/")
    ElseIf Len(u) > 2 And Mid$(u, 2, 1) = ":" And (Mid$(u, 3, 1) = "\" Or Mid$(u, 3, 1) = "/") Then
        SafeHref = "file:///" & Replace(u, "\", "/")
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
' Markdown (port of parseBlocks)
'   block = Array(type, level, marker, kind, "", lines, details, head, align, body)
' ===========================================================================
Public Const BT As Long = 0
Public Const BLV As Long = 1
Public Const BMK As Long = 2
Public Const BKD As Long = 3
Public Const BLN As Long = 5
Public Const BDT As Long = 6
Public Const BHD As Long = 7
Public Const BAL As Long = 8
Public Const BBD As Long = 9
Public Const SEP2 As String = "|~|"

Private Function MkBlock(ByVal t As String, ByVal level As Long, ByVal marker As String, ByVal kind As String) As Variant
    Dim c1 As Collection, c2 As Collection, c3 As Collection, a As Variant
    Set c1 = New Collection
    Set c2 = New Collection
    Set c3 = New Collection
    a = Array(t, level, marker, kind, "", c1, c2, "", "", c3)
    MkBlock = a
End Function

Public Function IsHrLine(ByVal t As String) As Boolean
    Dim k As String, ch As String
    k = RemoveChars(t, "")
    If Len(k) < 3 Then Exit Function
    ch = Left$(k, 1)
    If ch <> "-" And ch <> "*" And ch <> "_" And ch <> "=" Then Exit Function
    IsHrLine = (Replace(k, ch, "") = "")
End Function

Private Function SymRank(ByVal mk As String) As Double
    If InStr("□■", mk) > 0 Then
        SymRank = 0
    ElseIf InStr("○●◎◇◆▶▷►", mk) > 0 Then
        SymRank = 2
    ElseIf InStr("☐☑✓✔", mk) > 0 Then
        SymRank = 3
    ElseIf InStr("·ㆍ•∙◦▪▫→⇒", mk) > 0 Then
        SymRank = 4
    ElseIf mk = "※" Then
        SymRank = 5
    Else
        SymRank = 3
    End If
End Function

' list item at the start of a trimmed line (port of matchListItem)
Public Function MatchListItem(ByVal t As String, ByRef kind As String, ByRef marker As String, ByRef rank As Double, _
        ByRef txt As String, ByRef raw As String) As Boolean
    Dim c As String, p As Long, q As Long, n As Long, code As Long
    c = Left$(t, 1)
    If (c = "-" Or c = "*" Or c = "+") And IsWs(Mid$(t, 2, 1)) Then
        q = SkipWs(t, 2)
        If Mid$(t, q, 1) = "[" And Mid$(t, q + 2, 1) = "]" And InStr(" xXvV", Mid$(t, q + 1, 1)) > 0 And IsWs(Mid$(t, q + 3, 1)) Then
            kind = "check": rank = 3
            If InStr("xXvV", Mid$(t, q + 1, 1)) > 0 Then
                marker = "☑"
            Else
                marker = "☐"
            End If
            txt = Mid$(t, SkipWs(t, q + 3))
            MatchListItem = True
            Exit Function
        End If
        If c = "*" Then
            kind = "small": marker = "*": rank = 3.5
        Else
            kind = "md": raw = c: marker = c: rank = 3
        End If
        txt = Mid$(t, q)
        MatchListItem = True
        Exit Function
    End If
    ' 1. 1) (1)
    p = 1
    If c = "(" Then p = 2
    n = 0
    Do While IsDigitAt(t, p + n) And n < 3
        n = n + 1
    Loop
    If n > 0 Then
        q = p + n
        If (c = "(" And Mid$(t, q, 1) = ")") Or (c <> "(" And (Mid$(t, q, 1) = "." Or Mid$(t, q, 1) = ")")) Then
            If IsWs(Mid$(t, q + 1, 1)) Then
                kind = "num": marker = Left$(t, q): rank = 1
                txt = Mid$(t, SkipWs(t, q + 1))
                MatchListItem = True
                Exit Function
            End If
        End If
    End If
    ' 가. 가) (가)
    If c = "(" Then
        If InStr("가나다라마바사아자차카타파하", Mid$(t, 2, 1)) > 0 And Mid$(t, 3, 1) = ")" And IsWs(Mid$(t, 4, 1)) Then
            kind = "num": marker = Left$(t, 3): rank = 1.5
            txt = Mid$(t, SkipWs(t, 4))
            MatchListItem = True
            Exit Function
        End If
    ElseIf InStr("가나다라마바사아자차카타파하", c) > 0 And (Mid$(t, 2, 1) = "." Or Mid$(t, 2, 1) = ")") And IsWs(Mid$(t, 3, 1)) Then
        kind = "num": marker = Left$(t, 2): rank = 1.5
        txt = Mid$(t, SkipWs(t, 3))
        MatchListItem = True
        Exit Function
    End If
    code = CodeOf(c)
    If code >= 9312 And code <= 9331 Then
        kind = "num": marker = c: rank = 1.5
        txt = Mid$(t, SkipWs(t, 2))
        MatchListItem = True
        Exit Function
    End If
    If InStr("□■○●◎◇◆▶▷►·ㆍ•∙◦▪▫※☐☑✓✔→⇒", c) > 0 And c <> "" Then
        kind = "sym": marker = c: rank = SymRank(c)
        txt = Mid$(t, SkipWs(t, 2))
        MatchListItem = True
    End If
End Function

Private Function BulletSeq(ByVal style As String) As Variant
    If style = "report2" Then
        BulletSeq = Array("□", "○", "-", "·")
    ElseIf style = "dot" Then
        BulletSeq = Array("•", "◦", "▪", "·")
    Else
        BulletSeq = Array("○", "-", "·", "·")
    End If
End Function

Private Function AutoMarker(ByVal raw As String, ByVal level As Long, ByVal hasParent As Boolean, ByVal parentMk As String) As String
    Dim seq As Variant, i As Long, kor As Variant
    If gBullet = "keep" Then
        If raw = "-" Then
            AutoMarker = "-"
        Else
            AutoMarker = "•"
        End If
        Exit Function
    End If
    seq = BulletSeq(gBullet)
    If Not hasParent Then
        AutoMarker = seq(0)
        Exit Function
    End If
    For i = 0 To 3
        If seq(i) = parentMk Then
            AutoMarker = seq(MinD(i + 1, 3))
            Exit Function
        End If
    Next
    kor = Array("□", "○", "-", "·")
    For i = 0 To 3
        If kor(i) = parentMk Then
            AutoMarker = kor(MinD(i + 1, 3))
            Exit Function
        End If
    Next
    AutoMarker = seq(MinD(level, 3))
End Function

Private Function SplitRowCells(ByVal r As String) As Variant
    Dim cs As Variant, i As Long
    r = TrimX(r)
    If Left$(r, 1) = "|" Then r = Mid$(r, 2)
    If Right$(r, 1) = "|" Then r = Left$(r, Len(r) - 1)
    cs = Split(r, "|")
    For i = 0 To UBound(cs)
        cs(i) = TrimX(CStr(cs(i)))
    Next
    SplitRowCells = cs
End Function

Private Function IsAlignCell(ByVal c As String) As Boolean
    Dim k As String
    k = c
    If Left$(k, 1) = ":" Then k = Mid$(k, 2)
    If Right$(k, 1) = ":" Then k = Left$(k, Len(k) - 1)
    IsAlignCell = (Len(k) > 0 And Replace(k, "-", "") = "")
End Function

' lines -> Array(blocks As Collection, deptDetails As Collection)
Public Function ParseBlocks(ByVal srcLines As Collection) As Variant
    Dim out As New Collection, deptDet As New Collection
    Dim ls() As String, n As Long, i As Long, l As Variant, ln As String, t As String, indent As Long
    Dim stI() As Long, stR() As Double, stM() As String, nSt As Long
    Dim lastItem As Long, para As Long, target As Long, b As Variant, k As Long, j As Long
    Dim kind As String, mk As String, rank As Double, txt As String, raw As String, rows As Collection, cs As Variant
    Dim headS As String, alignS As String, level As Long, hasParent As Boolean, parentMk As String, c As Collection
    ReDim ls(1 To srcLines.Count + 1)
    For Each l In srcLines
        n = n + 1
        ls(n) = CStr(l)
    Next
    ReDim stI(1 To 64): ReDim stR(1 To 64): ReDim stM(1 To 64)
    i = 1
    Do While i <= n
        ln = Replace(ls(i), vbTab, "    ")
        j = 1
        Do While j <= Len(ln)
            If Mid$(ln, j, 1) = "　" Then
                ln = Left$(ln, j - 1) & "  " & Mid$(ln, j + 1)
                j = j + 2
            ElseIf Mid$(ln, j, 1) = " " Then
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        t = TrimX(ln)
        indent = 0
        Do While Mid$(ln, indent + 1, 1) = " "
            indent = indent + 1
        Loop
        If t = "" Then
            If out.Count > 0 Then
                b = out(out.Count)
                If b(BT) <> "sp" Then out.Add MkBlock("sp", 0, "", "")
            End If
            lastItem = 0: para = 0: target = 0
            i = i + 1
            GoTo nextLine
        End If
        If StartsWith(t, ">>") Or Left$(t, 1) = "》" Then
            AddDetail out, deptDet, target, DetailText(t)
            i = i + 1
            GoTo nextLine
        End If
        If Left$(t, 1) = "|" Then
            Set rows = New Collection
            Do While i <= n
                If Left$(TrimX(ls(i)), 1) <> "|" Then Exit Do
                rows.Add SplitRowCells(ls(i))
                i = i + 1
            Loop
            headS = "": alignS = ""
            b = MkBlock("table", 0, "", "")
            Set c = b(BBD)
            k = 1
            If rows.Count >= 2 Then
                cs = rows(2)
                If AllAlign(cs) Then
                    headS = Join(rows(1), SEP2)
                    For j = 0 To UBound(cs)
                        If Left$(cs(j), 1) = ":" And Right$(cs(j), 1) = ":" Then
                            cs(j) = "c"
                        ElseIf Right$(cs(j), 1) = ":" Then
                            cs(j) = "r"
                        Else
                            cs(j) = ""
                        End If
                    Next
                    alignS = Join(cs, SEP2)
                    k = 3
                End If
            End If
            For j = k To rows.Count
                c.Add Join(rows(j), SEP2)
            Next
            b(BHD) = headS
            b(BAL) = alignS
            out.Add b
            nSt = 0: lastItem = 0: para = 0: target = 0
            GoTo nextLine
        End If
        If IsHrLine(t) Then
            out.Add MkBlock("hr", 0, "", "")
            nSt = 0: lastItem = 0: para = 0: target = 0
            i = i + 1
            GoTo nextLine
        End If
        If Left$(t, 1) = "#" Then
            k = 0
            Do While Mid$(t, k + 1, 1) = "#" And k < 6
                k = k + 1
            Loop
            If IsWs(Mid$(t, k + 1, 1)) Then
                txt = TrimX(Mid$(t, k + 1))
                Do While Right$(txt, 1) = "#"
                    txt = RTrimX(Left$(txt, Len(txt) - 1))
                Loop
                b = MkBlock("h", k, "", "")
                Set c = b(BLN)
                c.Add txt
                out.Add b
                nSt = 0: lastItem = 0: para = 0
                target = out.Count
                i = i + 1
                GoTo nextLine
            End If
        End If
        If Left$(t, 1) = ">" Then
            txt = Mid$(t, 2)
            If Left$(txt, 1) = " " Then txt = Mid$(txt, 2)
            If out.Count > 0 Then
                b = out(out.Count)
                If b(BT) = "q" Then
                    Set c = b(BLN)
                    c.Add txt
                    lastItem = 0: para = 0: target = out.Count
                    i = i + 1
                    GoTo nextLine
                End If
            End If
            b = MkBlock("q", 0, "", "")
            Set c = b(BLN)
            c.Add txt
            out.Add b
            lastItem = 0: para = 0: target = out.Count
            i = i + 1
            GoTo nextLine
        End If
        If MatchListItem(t, kind, mk, rank, txt, raw) Then
            If StartsWith(TrimX(txt), ">>") Or Left$(TrimX(txt), 1) = "》" Then
                AddDetail out, deptDet, target, DetailText(TrimX(txt))
                i = i + 1
                GoTo nextLine
            End If
            ' place in the stack of open list levels
            Do While nSt > 0
                If indent < stI(nSt) Or (indent = stI(nSt) And rank < stR(nSt)) Then
                    nSt = nSt - 1
                Else
                    Exit Do
                End If
            Loop
            If nSt > 0 Then
                If stI(nSt) = indent And stR(nSt) = rank Then nSt = nSt - 1
            End If
            hasParent = (nSt > 0)
            If hasParent Then
                parentMk = stM(nSt)
            Else
                parentMk = ""
            End If
            If nSt < 63 Then nSt = nSt + 1
            stI(nSt) = indent: stR(nSt) = rank
            level = nSt - 1
            If kind = "md" Then mk = AutoMarker(raw, level, hasParent, parentMk)
            stM(nSt) = mk
            b = MkBlock("li", level, mk, kind)
            Set c = b(BLN)
            c.Add txt
            out.Add b
            lastItem = out.Count: para = 0: target = out.Count
            i = i + 1
            GoTo nextLine
        End If
        If lastItem > 0 And indent > 0 Then
            b = out(lastItem)
            Set c = b(BLN)
            c.Add t
            i = i + 1
            GoTo nextLine
        End If
        If para > 0 Then
            b = out(para)
            Set c = b(BLN)
            c.Add t
        Else
            b = MkBlock("p", 0, "", "")
            Set c = b(BLN)
            c.Add t
            out.Add b
            para = out.Count
        End If
        lastItem = 0: target = para: nSt = 0
        i = i + 1
nextLine:
    Loop
    ' no spacer at the start or the end
    Do While out.Count > 0
        b = out(1)
        If b(BT) <> "sp" Then Exit Do
        out.Remove 1
    Loop
    Do While out.Count > 0
        b = out(out.Count)
        If b(BT) <> "sp" Then Exit Do
        out.Remove out.Count
    Loop
    ParseBlocks = Array(out, deptDet)
End Function

Private Function AllAlign(cs As Variant) As Boolean
    Dim j As Long
    For j = 0 To UBound(cs)
        If Not IsAlignCell(CStr(cs(j))) Then Exit Function
    Next
    AllAlign = (UBound(cs) >= 0)
End Function

Private Function DetailText(ByVal t As String) As String
    If Left$(t, 1) = "》" Then
        t = Mid$(t, 2)
    Else
        t = Mid$(t, 3)
    End If
    If Left$(t, 1) = " " Then t = Mid$(t, 2)
    DetailText = t
End Function

Private Sub AddDetail(ByVal out As Collection, ByVal deptDet As Collection, ByVal target As Long, ByVal txt As String)
    Dim b As Variant, c As Collection
    If target > 0 Then
        b = out(target)
        Set c = b(BDT)
        c.Add txt
    Else
        deptDet.Add txt
    End If
End Sub

' ===========================================================================
' One line of text -> rich text (links, @place, bold time, inline marks)
' ===========================================================================
Private Function IsUrlStart(ByVal s As String, ByVal p As Long) As Boolean
    Dim l As String
    l = LCase$(Mid$(s, p, 8))
    IsUrlStart = StartsWith(l, "http://") Or StartsWith(l, "https://") Or StartsWith(l, "file://") Or StartsWith(l, "ftp://") Or StartsWith(l, "www.")
End Function

' removes bare addresses from the text (they are not printed); md links stay
Private Function SplitUrls(ByVal s As String, ByVal urls As Collection) As String
    Dim i As Long, n As Long, out As String, e As Long, u As String, found As Boolean, q As Long, ch As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "[" Then
            q = InStr(i, s, "](")
            If q > 0 Then
                e = InStr(q, s, ")")
                If e > 0 And InStr(Mid$(s, i, q - i), vbLf) = 0 Then
                    out = out & Mid$(s, i, e - i + 1)
                    i = e + 1
                    GoTo nextCh
                End If
            End If
        End If
        If IsUrlStart(s, i) Then
            If i = 1 Or Not IsUrlStart(s, i - 1) Then
                e = i
                Do While e <= n
                    ch = Mid$(s, e, 1)
                    If IsWs(ch) Or ch = "<" Or ch = ">" Or ch = """" Or ch = "'" Then Exit Do
                    e = e + 1
                Loop
                u = Mid$(s, i, e - i)
                Do While Len(u) > 0 And InStr(".,;:!?)]", Right$(u, 1)) > 0
                    u = Left$(u, Len(u) - 1)
                Loop
                If Len(u) > 4 Then
                    urls.Add u
                    found = True
                    i = i + Len(u)
                    GoTo nextCh
                End If
            End If
        End If
        out = out & Mid$(s, i, 1)
        i = i + 1
nextCh:
    Loop
    If found Then
        out = Replace(Replace(Replace(out, "( )", ""), "()", ""), "[]", "")
        Do While InStr(out, "  ") > 0
            out = Replace(out, "  ", " ")
        Loop
        out = TrimX(out)
    End If
    SplitUrls = out
End Function

' "내용 @장소" -> "내용 (장소)"
Private Function AtPlace(ByVal s As String) As String
    Dim p As Long, place As String
    p = InStrRev(s, "@")
    If p > 1 Then
        If IsWs(Mid$(s, p - 1, 1)) Then
            place = TrimX(Mid$(s, p + 1))
            If place <> "" And InStr(place, vbLf) = 0 Then
                AtPlace = RTrimX(Left$(s, p - 1)) & " (" & place & ")"
                Exit Function
            End If
        End If
    End If
    AtPlace = s
End Function

' [text](url) -> text; first usable url returned through linkUrl
Private Function MdLinks(ByVal s As String, ByRef linkUrl As String) As String
    Dim i As Long, q As Long, e As Long, out As String, txt As String, u As String, e2 As Long
    i = 1
    Do
        q = InStr(i, s, "[")
        If q = 0 Then Exit Do
        e = InStr(q, s, "](")
        If e = 0 Then Exit Do
        e2 = InStr(e + 2, s, ")")
        If e2 = 0 Then Exit Do
        txt = Mid$(s, q + 1, e - q - 1)
        u = TrimX(Replace(Replace(Mid$(s, e + 2, e2 - e - 2), "<", ""), ">", ""))
        If SafeHref(u) <> "" And InStr(txt, "]") = 0 Then
            out = out & Mid$(s, i, q - i) & txt
            If linkUrl = "" Then linkUrl = SafeHref(u)
        Else
            out = out & Mid$(s, i, e2 - i + 1)
        End If
        i = e2 + 1
    Loop
    MdLinks = out & Mid$(s, i)
End Function

' appends one line to the rich text builder. returns the url of a link (first one), extra urls into urls
Public Sub RBLine(ByVal lineTxt As String, ByVal size As Double, ByRef linkUrl As String, ByVal urls As Collection, Optional ByVal noTime As Boolean = False)
    Dim body As String, tl As Long, myUrls As New Collection, u As Variant, first As Boolean
    body = SplitUrls(lineTxt, myUrls)
    body = AtPlace(body)
    body = MdLinks(Replace(body, "`", ""), linkUrl)
    first = True
    For Each u In myUrls
        If linkUrl = "" And first Then
            linkUrl = SafeHref(CStr(u))
        Else
            urls.Add CStr(u)
        End If
        first = False
    Next
    If body = "" And myUrls.Count > 0 Then body = "관련 링크"
    tl = 0
    If gTimeBold And Not noTime Then tl = DTimeLen(body)
    If tl > 0 Then
        RBAdd Left$(body, tl), True, , , size
        RBInline Mid$(body, tl + 1), , size, , True
    Else
        RBInline body, , size, , True
    End If
End Sub

' text of several lines -> builder (returns link + note)
Private Sub RBLines(ByVal lines As Collection, ByVal size As Double, ByRef linkUrl As String, ByRef note As String, ByVal details As Collection)
    Dim l As Variant, n As Long, urls As New Collection, d As Variant
    RBReset
    For Each l In lines
        If n > 0 Then RBAdd vbLf
        n = n + 1
        RBLine CStr(l), size, linkUrl, urls
    Next
    note = ""
    If Not details Is Nothing Then
        For Each d In details
            note = note & IIf(note = "", "", vbLf) & PlainInline(CStr(d))
        Next
    End If
    For Each d In urls
        note = note & IIf(note = "", "", vbLf) & "↗ " & d
    Next
End Sub

' ===========================================================================
' Layout of blocks
' ===========================================================================
Private Function SmallFs(ByVal fs As Double) As Double
    SmallFs = R5(MaxD(MaxD(fs - 2, 0.75 * fs), 6))
End Function

' height of one block at width w; draws it when doRender
Private Function LayBlock(b As Variant, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal doRender As Boolean) As Double
    Dim fs As Double, lk As String, note As String, h As Double, idx As Long, sz As Double, lines As Collection, det As Collection
    Dim ind As Double, mkW As Double, mfs As Double, txt As String, runs As String, mark As Boolean
    fs = gFS
    Select Case b(BT)
        Case "sp"
            LayBlock = 0.12 * fs + 0.6 * gGapPt
        Case "hr"
            If doRender Then
                idx = ElBox(x, y + 0.75, w, 0.75)
                ElBorder idx, 0, XL_DASH, XL_HAIR, COL_QLINE
            End If
            LayBlock = 1.5
        Case "table"
            LayBlock = LayMdTable(b, x, y, w, doRender)
        Case Else
            Set lines = b(BLN)
            Set det = b(BDT)
            sz = fs
            If b(BT) = "h" Then
                If b(BLV) = 1 Then
                    sz = R5(fs * 1.1)
                ElseIf b(BLV) = 2 Then
                    sz = R5(fs * 1.05)
                End If
            End If
            If b(BT) = "li" Then
                If b(BKD) = "small" Then sz = SmallFs(fs)
                ind = b(BLV) * 1.15 * fs
                mfs = sz
                mkW = MaxD(MeasureW(CStr(b(BMK)), gFont, mfs), 0.9 * mfs + 3) + 0.15 * mfs
                lk = ""
                RBLines lines, sz, lk, note, det
                txt = RBText(): runs = RBRuns(): mark = rbMark
                h = MeasureH(txt, runs, w - ind - mkW, gFont, sz, False)
                If doRender Then
                    ElText x + ind, y, mkW, h, CStr(b(BMK)), "", mfs
                    idx = ElText(x + ind + mkW, y, w - ind - mkW, h, txt, runs, sz)
                    FinishText idx, lk, note, mark
                End If
                LayBlock = h
                Exit Function
            End If
            lk = ""
            RBLines lines, sz, lk, note, det
            txt = RBText(): runs = RBRuns(): mark = rbMark
            If b(BT) = "q" Then
                h = MeasureH(txt, runs, w - 0.5 * fs - 2, gFont, sz, False)
                If doRender Then
                    idx = ElBox(x, y, w, h)
                    ElBorder idx, 2, XL_CONT, XL_MEDIUM, COL_QLINE
                    idx = ElText(x + 0.5 * fs + 2, y, w - 0.5 * fs - 2, h, txt, runs, sz, False, XL_LEFT, XL_TOP, COL_QTEXT)
                    FinishText idx, lk, note, mark
                End If
            Else
                h = MeasureH(txt, runs, w, gFont, sz, (b(BT) = "h"))
                If doRender Then
                    idx = ElText(x, y, w, h, txt, runs, sz, (b(BT) = "h"))
                    FinishText idx, lk, note, mark
                End If
            End If
            LayBlock = h
    End Select
End Function

Private Sub FinishText(ByVal idx As Long, ByVal lk As String, ByVal note As String, ByVal mark As Boolean)
    If mark Then ElFill idx, COL_MARKD
    If lk <> "" Then ElLink idx, lk
    If note <> "" Then ElNote idx, note
End Sub

' markdown table
Private Function LayMdTable(b As Variant, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal doRender As Boolean) As Double
    Dim body As Collection, head As Variant, al As Variant, nC As Long, r As Variant, cs As Variant, j As Long, nR As Long
    Dim wts() As Double, cw() As Double, tot As Double, rowsTxt() As String, k As Long, fs As Double, ph As Double, pv As Double
    Dim yy As Double, rh As Double, hh() As Double, lk As String, note As String, urls As New Collection, idx As Long, cx As Double
    Dim txts() As String, runsA() As String, links() As String, isHead As Boolean, hal As Long, rIdx As Long
    fs = gFS
    ph = 0.35 * fs
    pv = 0.08 * fs
    Set body = b(BBD)
    If b(BHD) <> "" Then head = Split(b(BHD), SEP2)
    If b(BAL) <> "" Then al = Split(b(BAL), SEP2)
    If IsArray(head) Then nC = UBound(head) + 1
    For Each r In body
        cs = Split(r, SEP2)
        If UBound(cs) + 1 > nC Then nC = UBound(cs) + 1
    Next
    If nC = 0 Then Exit Function
    nR = body.Count
    If IsArray(head) Then nR = nR + 1
    ReDim wts(0 To nC - 1): ReDim cw(0 To nC - 1)
    ReDim txts(1 To nR, 0 To nC - 1): ReDim runsA(1 To nR, 0 To nC - 1): ReDim links(1 To nR, 0 To nC - 1)
    ReDim hh(1 To nR)
    rIdx = 0
    If IsArray(head) Then
        rIdx = 1
        For j = 0 To nC - 1
            If j <= UBound(head) Then txts(1, j) = head(j)
        Next
    End If
    For Each r In body
        rIdx = rIdx + 1
        cs = Split(r, SEP2)
        For j = 0 To nC - 1
            If j <= UBound(cs) Then txts(rIdx, j) = cs(j)
        Next
    Next
    For k = 1 To nR
        For j = 0 To nC - 1
            wts(j) = MaxD(wts(j), MinD(TextWidthEm(PlainInline(txts(k, j))), 18))
        Next
    Next
    For j = 0 To nC - 1
        wts(j) = MaxD(wts(j), 2)
        tot = tot + wts(j)
    Next
    For j = 0 To nC - 1
        cw(j) = w * wts(j) / tot
    Next
    ' measure rows
    For k = 1 To nR
        isHead = (IsArray(head) And k = 1)
        For j = 0 To nC - 1
            lk = ""
            RBReset
            RBLine txts(k, j), fs, lk, urls, isHead
            txts(k, j) = RBText(): runsA(k, j) = RBRuns(): links(k, j) = lk
            hh(k) = MaxD(hh(k), MeasureH(IIf(RBText() = "", "가", RBText()), runsA(k, j), cw(j) - 2 * ph + 2 * CELLPAD, gFont, fs, isHead))
        Next
        hh(k) = hh(k) + 2 * pv
    Next
    yy = y
    For k = 1 To nR
        isHead = (IsArray(head) And k = 1)
        If doRender Then
            cx = x
            For j = 0 To nC - 1
                idx = ElBox(cx, yy, cw(j), hh(k))
                ElBorderAll idx, XL_CONT, XL_THIN, COL_TBORDER
                If isHead Then ElFill idx, COL_THEAD
                hal = XL_LEFT
                If isHead Then
                    hal = XL_CENTER
                ElseIf IsArray(al) Then
                    If j <= UBound(al) Then
                        If al(j) = "c" Then
                            hal = XL_CENTER
                        ElseIf al(j) = "r" Then
                            hal = XL_RIGHT
                        End If
                    End If
                End If
                If txts(k, j) <> "" Then
                    idx = ElText(cx + ph - CELLPAD, yy + pv, cw(j) - 2 * ph + 2 * CELLPAD, hh(k) - 2 * pv, txts(k, j), runsA(k, j), fs, isHead, hal)
                    If links(k, j) <> "" Then ElLink idx, links(k, j)
                End If
                cx = cx + cw(j)
            Next
        End If
        yy = yy + hh(k)
    Next
    LayMdTable = yy - y
End Function

' blocks i0..i1 of a list; first block has no gap above
Private Function LayBlocks(ByVal bl As Collection, ByVal i0 As Long, ByVal i1 As Long, ByVal x As Double, ByVal y As Double, _
        ByVal w As Double, ByVal doRender As Boolean) As Double
    Dim i As Long, yy As Double, b As Variant, first As Boolean
    yy = y
    first = True
    For i = i0 To i1
        b = bl(i)
        If Not first Then
            If b(BT) = "h" Then
                yy = yy + 2 * gGapPt + 0.25 * gFS
            Else
                yy = yy + gGapPt
            End If
        End If
        yy = yy + LayBlock(b, x, yy, w, doRender)
        first = False
    Next
    LayBlocks = yy - y
End Function

' groups for the left/right split (port of groupBlocks). returns starts and ends (block indices)
Private Function MakeGroups(ByVal bl As Collection, ByRef gs() As Long, ByRef ge() As Long) As Long
    Dim i As Long, b As Variant, n As Long, cur As Boolean, afterHead As Boolean, startsNew As Boolean
    ' gs / ge are sized by the caller (resizing an array passed ByRef is unsafe in some hosts)
    For i = 1 To bl.Count
        b = bl(i)
        If b(BT) = "sp" Then
            cur = False
            afterHead = False
        ElseIf b(BT) = "h" Then
            n = n + 1
            gs(n) = i: ge(n) = i
            cur = True
            afterHead = True
        Else
            If b(BT) = "li" Then
                startsNew = (b(BLV) = 0 And b(BKD) <> "small")
            Else
                startsNew = (b(BT) <> "q")
            End If
            If Not cur Or (startsNew And Not afterHead) Then
                n = n + 1
                gs(n) = i: ge(n) = i
                cur = True
            Else
                ge(n) = i
            End If
            afterHead = False
        End If
    Next
    MakeGroups = n
End Function

' groups gi..gj (each a block range) one under another
Private Function LayGroups(ByVal bl As Collection, gs() As Long, ge() As Long, ByVal gi As Long, ByVal gj As Long, _
        ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal doRender As Boolean) As Double
    Dim g As Long, yy As Double
    yy = y
    For g = gi To gj
        If g > gi Then yy = yy + gGapPt
        yy = yy + LayBlocks(bl, gs(g), ge(g), x, yy, w, doRender)
    Next
    LayGroups = yy - y
End Function

Private Function HasWideTable(ByVal bl As Collection) As Boolean
    Dim b As Variant, body As Collection, r As Variant, n As Long
    For Each b In bl
        If b(BT) = "table" Then
            n = 0
            If b(BHD) <> "" Then n = UBound(Split(b(BHD), SEP2)) + 1
            Set body = b(BBD)
            For Each r In body
                If UBound(Split(r, SEP2)) + 1 > n Then n = UBound(Split(r, SEP2)) + 1
            Next
            If n > 4 Then
                HasWideTable = True
                Exit Function
            End If
        End If
    Next
End Function

' content of department d inside width w. halfGap: space between the halves (0 = one column)
' returns natural height; draws when doRender
Private Function LayContent(ByVal d As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal useHalves As Boolean, _
        ByVal halfGap As Double, ByVal divX As Double, ByVal rowH As Double, ByVal doRender As Boolean) As Double
    Dim bl As Collection, gs() As Long, ge() As Long, ng As Long, hw As Double, k As Long, best As Double, bestK As Long
    Dim L() As Double, R() As Double, j As Long, mx As Double, hL As Double, hR As Double, idx As Long
    Set bl = dBlocks(d)
    If bl.Count = 0 Then
        If gEmptyText <> "" Then
            hL = MeasureH(gEmptyText, "", w, gFont, gFS)
            If doRender Then ElText x, y, w, hL, gEmptyText, "", gFS, False, XL_LEFT, XL_TOP, COL_EMPTY
        End If
        LayContent = hL
        Exit Function
    End If
    ng = 0
    ReDim gs(1 To bl.Count + 1): ReDim ge(1 To bl.Count + 1)
    If useHalves And Not HasWideTable(bl) Then ng = MakeGroups(bl, gs, ge)
    If ng < 2 Then
        LayContent = LayBlocks(bl, 1, bl.Count, x, y, w, doRender)
        Exit Function
    End If
    hw = (w - halfGap) / 2
    ReDim L(1 To ng): ReDim R(1 To ng)
    For j = 1 To ng
        L(j) = LayGroups(bl, gs, ge, 1, j, x, y, hw, False)
        R(j) = LayGroups(bl, gs, ge, j, ng, x, y, hw, False)
    Next
    best = 1E+30
    bestK = ng
    For j = 1 To ng - 1
        mx = MaxD(L(j), R(j + 1))
        If mx < best - 0.5 Or (mx <= best + 0.5 And L(j) >= R(j + 1)) Then
            best = MinD(best, mx)
            bestK = j
        End If
    Next
    hL = L(bestK)
    If bestK < ng Then hR = R(bestK + 1)
    If doRender Then
        LayGroups bl, gs, ge, 1, bestK, x, y, hw, True
        If bestK < ng Then LayGroups bl, gs, ge, bestK + 1, ng, x + hw + halfGap, y, hw, True
        If divX > 0 Then
            idx = ElBox(divX - 4, y - IIf(gLayout = "table", gPadV, 0), 4, IIf(gLayout = "table", rowH, MaxD(hL, hR)))
            ElBorder idx, 3, XL_CONT, XL_THIN, COL_HALF
        End If
    End If
    LayContent = MaxD(hL, hR)
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
    Set ws = ThisWorkbook.Worksheets(SH_A4)
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
    Set ws = ThisWorkbook.Worksheets(SH_A4)
    Set shp = ws.Shapes("a4status")
    If Not shp Is Nothing Then
        shp.TextFrame.Characters.Text = IIf(ok, "맞춤 완료: ", "넘침: ") & st
        shp.TextFrame.Characters.Font.Color = IIf(ok, COL_OK, COL_BAD)
    End If
    On Error GoTo 0
End Sub

Private Sub PlaceA4Buttons(ByVal ws As Object, ByVal lastC As Long)
    Dim x As Double, shp As Object, y As Double
    On Error Resume Next
    x = ws.Cells(1, lastC + 1).Left + 18
    y = 6
    For Each shp In ws.Shapes
        If shp.Name = "a4status" Then
            shp.Left = x
            shp.Top = 6
        Else
            shp.Left = x
            shp.Top = 80 + y
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
        ws.Cells(row, r.Column + 3).NumberFormat = "@"
        ws.Cells(row, r.Column + 3).Value = st
        ws.Cells(row, r.Column + 4).Value = dLines(i) & "줄"
        ws.Cells(row, r.Column + 3).Font.Color = IIf(st = "미입력", COL_BAD, IIf(st = "직접 입력", 0, COL_OK))
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
