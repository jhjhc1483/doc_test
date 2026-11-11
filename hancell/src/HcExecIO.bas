Attribute VB_Name = "HcExecIO"
Option Explicit

' ===========================================================================
' Import (엑셀 불러오기 / 붙여넣은 표 가져오기)
' ===========================================================================
Private Const SEP1 As String = "|~|"
Private gN As Long
Private gLabel() As String, gTitleG() As String, gSource() As String, gHasExec() As Boolean, gHasRoles() As Boolean
Private gTarget() As Long, gRowsC() As Object, gRoles() As Variant, gLabels() As Variant

Private Sub GroupsReset()
    gN = 0
    ReDim gLabel(1 To 1): ReDim gTitleG(1 To 1): ReDim gSource(1 To 1): ReDim gHasExec(1 To 1): ReDim gHasRoles(1 To 1)
    ReDim gTarget(1 To 1): ReDim gRowsC(1 To 1): ReDim gRoles(1 To 1): ReDim gLabels(1 To 1)
End Sub

Private Function GroupAdd(ByVal label As String, ByVal title As String, ByVal source As String, ByVal hasExec As Boolean, _
        ByVal hasRoles As Boolean, roles As Variant, labels As Variant) As Long
    gN = gN + 1
    ReDim Preserve gLabel(1 To gN): ReDim Preserve gTitleG(1 To gN): ReDim Preserve gSource(1 To gN)
    ReDim Preserve gHasExec(1 To gN): ReDim Preserve gHasRoles(1 To gN): ReDim Preserve gTarget(1 To gN)
    ReDim Preserve gRowsC(1 To gN): ReDim Preserve gRoles(1 To gN): ReDim Preserve gLabels(1 To gN)
    gLabel(gN) = label: gTitleG(gN) = title: gSource(gN) = source: gHasExec(gN) = hasExec: gHasRoles(gN) = hasRoles
    gTarget(gN) = -1
    Set gRowsC(gN) = New Collection
    ' stored as joined text: arrays kept in variants are shared by reference in some hosts
    gRoles(gN) = Join(roles, SEP1): gLabels(gN) = Join(labels, SEP1)
    GroupAdd = gN
End Function

Private Function SkipSheetName(ByVal nm As String) As Boolean
    Dim k As String, p As Variant
    k = LCase$(RemoveChars(nm, ""))
    For Each p In Array("작성방법", "작성요령", "작성예시", "안내", "도움말", "사용법", "예시", "목록", "설정", "readme", "guide", "help")
        If StartsWith(k, CStr(p)) Then
            SkipSheetName = True
            Exit Function
        End If
    Next
End Function

' a sheet (values) -> groups (port of sheetGroups)
Private Sub SheetGroups(ByVal sheetName As String, v As Variant, ByVal r0 As Long, ByVal source As String)
    Dim nR As Long, nC As Long, r As Long, c As Long, hr As Long, i As Long
    If Not IsArray(v) Then Exit Sub
    nR = UBound(v, 1): nC = UBound(v, 2)
    Dim rowS() As String
    ReDim rowS(0 To nC - 1)
    Dim roles() As String, labels() As String, hasRoles As Boolean
    ReDim roles(0 To nC - 1): ReDim labels(0 To nC - 1)
    hr = 0
    For r = r0 To MinD(nR, r0 + 19)
        For c = 1 To nC
            rowS(c - 1) = CellText(v(r, c))
        Next
        If IsHeaderRow(rowS) Then
            hr = r
            Exit For
        End If
    Next
    Dim ce As Long, cp As Long, cd As Long, title As String, rowTxt As String
    ce = -1: cp = -1: cd = -1
    If hr > 0 Then
        hasRoles = True
        For c = 1 To nC
            HeadRole CellText(v(hr, c)), roles(c - 1), labels(c - 1)
            If roles(c - 1) = "exec" And ce < 0 Then ce = c - 1
            If roles(c - 1) = "pos" And cp < 0 Then cp = c - 1
            If roles(c - 1) = "date" And cd < 0 Then cd = c - 1
        Next
        For r = r0 To hr - 1
            rowTxt = ""
            For c = 1 To nC
                If CellText(v(r, c)) <> "" Then rowTxt = TrimX(rowTxt & " " & CellText(v(r, c)))
            Next
            If rowTxt <> "" Then title = TrimX(title & " " & rowTxt)
        Next
    Else
        hr = r0 - 1
    End If
    Dim curExec As String, curPos As String, curDate As String, anyV As Boolean, key As String, g As Long, y0 As Long
    Dim keys() As String, keyG() As Long, nKeys As Long
    ReDim keys(1 To nR + 1): ReDim keyG(1 To nR + 1)
    y0 = CLng(Left$(gDate, 4))
    For r = hr + 1 To nR
        anyV = False
        ReDim rowS(0 To nC - 1)
        For c = 1 To nC
            rowS(c - 1) = CellText(v(r, c), hasRoles And roles(c - 1) = "time")
            If hasRoles Then
                If roles(c - 1) = "date" And rowS(c - 1) <> "" Then rowS(c - 1) = DateValueIso(v(r, c), y0)
            End If
            If rowS(c - 1) <> "" Then anyV = True
        Next
        If anyV Then
            If Not IsHeaderRow(rowS) Then
                If ce >= 0 Then
                    If rowS(ce) <> "" Then
                        curExec = rowS(ce)
                        If cp >= 0 Then curPos = rowS(cp)
                    ElseIf cp >= 0 Then
                        If rowS(cp) <> "" Then curPos = rowS(cp)
                    End If
                ElseIf cp >= 0 Then
                    If rowS(cp) <> "" Then curPos = rowS(cp)
                End If
                If cd >= 0 Then
                    If rowS(cd) <> "" Then
                        If IsIso(rowS(cd)) Then
                            curDate = rowS(cd)
                        Else
                            curDate = IIf(ParseLooseDate(rowS(cd), y0) <> "", ParseLooseDate(rowS(cd), y0), curDate)
                        End If
                    End If
                End If
                If ce >= 0 Then
                    key = "e:" & curPos & "|" & curExec
                Else
                    key = "s"
                End If
                g = FindKey(keys, nKeys, key)
                If g > 0 Then g = keyG(g)
                If g = 0 Then
                    If ce >= 0 Then
                        g = GroupAdd(TrimX(curPos & " " & curExec), title, source & IIf(TrimX(curPos & " " & curExec) <> "", " › " & TrimX(curPos & " " & curExec), ""), True, hasRoles, roles, labels)
                    Else
                        g = GroupAdd(sheetName, title, source, False, hasRoles, roles, labels)
                    End If
                    nKeys = nKeys + 1
                    keys(nKeys) = key
                    keyG(nKeys) = g
                End If
                gRowsC(g).Add Array(IIf(cd >= 0, curDate, ""), Join(rowS, SEP1))
            End If
        End If
    Next
End Sub

Private Function NormKey(ByVal s As String) As String
    NormKey = LCase$(RemoveChars(RemoveParens(s), "·.,_-"))
End Function

' find the seat by a label (sheet name / 성명 column). -1 when unknown or ambiguous
Private Function MatchExec(ByVal label As String) As Long
    Dim L As String, i As Long, n As String, p As String, sc As Long, best As Long, bestScore As Long, tie As Boolean
    L = NormKey(label)
    best = -1
    If L = "" Then
        MatchExec = -1
        Exit Function
    End If
    For i = 1 To NX
        n = NormKey(exName(i)): p = NormKey(exPos(i))
        sc = 0
        If n <> "" And InStr(L, n) > 0 Then
            sc = 10
            If p <> "" And InStr(L, p) > 0 Then sc = 12
        ElseIf p <> "" And L = p Then
            sc = 5
        ElseIf p <> "" And InStr(L, p) > 0 Then
            sc = 3
        End If
        If sc > bestScore Then
            bestScore = sc: best = i: tie = False
        ElseIf sc > 0 And sc = bestScore Then
            tie = True
        End If
    Next
    If tie Then
        MatchExec = -1
    Else
        MatchExec = best
    End If
End Function

Private Function ExecLabel(ByVal i As Long) As String
    ExecLabel = TrimX(exPos(i) & " " & exName(i))
    If ExecLabel = "" Then ExecLabel = i & "번 칸"
End Function

' rows of group g that belong to date D -> lines (for counting / no-header groups)
Private Sub GroupLines(ByVal g As Long, ByVal D As String, ByVal sched As Collection, ByVal refc As Collection)
    Dim rw As Variant, cells() As String, roles() As String, labels() As String
    roles = Split(gRoles(g), SEP1): labels = Split(gLabels(g), SEP1)
    For Each rw In gRowsC(g)
        If rw(0) = "" Or D = "" Or rw(0) = D Then
            cells = Split(rw(1), SEP1)
            RowLines cells, roles, labels, gHasRoles(g), sched, refc
        End If
    Next
End Sub

Public Sub ImportExcel()
    Dim f As Variant
    On Error GoTo fail
    f = PickFiles("두레이에서 내려받은 엑셀 파일을 고르세요", "엑셀 파일 (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", True)
    If IsArray(f) Then ImportFiles f
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Function ImportFiles(files As Variant) As String
    Dim p As Variant, wb As Object, ws As Object, v As Variant, names As String, fn As String
    LoadSettings
    LoadExecs
    gStep = "import: reset"
    GroupsReset
    For Each p In files
        If LCase$(CStr(p)) = LCase$(HostBook().FullName) Then
            Info "지금 열려 있는 이 파일은 불러올 수 없습니다. 두레이에서 내려받은 공동편집 파일을 고르세요."
            Exit Function
        End If
        fn = Mid$(CStr(p), InStrRev(Replace(CStr(p), "\", "/"), "/") + 1)
        gStep = "import: open " & fn
        On Error Resume Next
        Set wb = Nothing
        Set wb = Application.Workbooks.Open(CStr(p), 0, True)
        On Error GoTo 0
        If wb Is Nothing Then
            Info fn & ": 파일을 열 수 없습니다."
        Else
            If names <> "" Then names = names & ", "
            names = names & fn
            For Each ws In wb.Worksheets
                gStep = "import: sheet " & ws.Name
                If ws.Visible = XL_SHEET_VISIBLE And Not SkipSheetName(ws.Name) Then
                    v = SheetValues(ws)
                    gStep = "import: groups " & ws.Name
                    SheetGroups ws.Name, v, 1, fn & " › " & ws.Name
                End If
            Next
            wb.Close False
        End If
    Next
    If names = "" Then Exit Function
    gStep = "import: open dialog"
    ImportFiles = OpenImport(names)
End Function

Private Function SheetValues(ByVal ws As Object) As Variant
    Dim ur As Object, v As Variant, a(1 To 1, 1 To 1) As Variant
    Set ur = ws.UsedRange
    Set ur = ws.Range(ws.Cells(1, 1), ws.Cells(ur.Row + ur.Rows.Count - 1, ur.Column + ur.Columns.Count - 1))
    v = ur.Value
    If Not IsArray(v) Then
        a(1, 1) = v
        v = a
    End If
    SheetValues = v
End Function

Public Sub ImportPasted()
    Dim ws As Object, v As Variant, lastR As Long, lastC As Long
    On Error GoTo fail
    Set ws = SheetByName(SH_PASTE)
    If ws Is Nothing Then Exit Sub
    UsedLast ws, lastR, lastC
    If lastR < 3 Then
        Info "붙여넣기 시트의 3번째 줄부터 엑셀에서 복사한 표를 붙여넣은 뒤 다시 누르세요."
        ws.Activate
        Exit Sub
    End If
    LoadSettings
    LoadExecs
    GroupsReset
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, MaxD(lastC, 2))).Value
    SheetGroups "붙여넣은 내용", v, 3, "붙여넣은 내용"
    If OpenImport("붙여넣은 내용") = "ok" Then ws.Range(ws.Cells(3, 1), ws.Cells(lastR, MaxD(lastC, 2))).ClearContents
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

' choose date / seats, confirm, apply. returns "ok" when applied
Private Function OpenImport(ByVal srcLabel As String) As String
    Dim g As Long, rw As Variant, D As String, k As Long, c As Long
    Dim bestC As Long, msg As String, ans As String, defIdx As Long
    Dim dates() As String, cnt() As Long, nD As Long, t As String, j As Long, tc As Long
    ReDim dates(1 To 1): ReDim cnt(1 To 1)
    If gN = 0 Then
        Info "가져올 일정을 찾지 못했습니다. 머리글(시간·내용)이 있는지 확인하세요."
        Exit Function
    End If
    For g = 1 To gN
        gTarget(g) = MatchExec(gLabel(g))
        If gTarget(g) < 0 And Not gHasExec(g) And gTitleG(g) <> "" Then gTarget(g) = MatchExec(gTitleG(g))
        For Each rw In gRowsC(g)
            If rw(0) <> "" Then
                c = FindKey(dates, nD, rw(0))
                If c = 0 Then
                    nD = nD + 1
                    ReDim Preserve dates(1 To nD): ReDim Preserve cnt(1 To nD)
                    dates(nD) = rw(0)
                    cnt(nD) = 1
                Else
                    cnt(c) = cnt(c) + 1
                End If
            End If
        Next
    Next
    ' sort dates
    For k = 1 To nD - 1
        For j = k + 1 To nD
            If dates(j) < dates(k) Then
                t = dates(k): dates(k) = dates(j): dates(j) = t
                tc = cnt(k): cnt(k) = cnt(j): cnt(j) = tc
            End If
        Next
    Next
    ' date
    If nD > 0 Then
        If FindKey(dates, nD, gDate) > 0 Then
            D = gDate
        Else
            For k = 1 To nD
                If cnt(k) > bestC Then
                    bestC = cnt(k)
                    D = dates(k)
                End If
            Next
        End If
        If nD > 1 Or D <> gDate Then
            msg = "엑셀에 날짜가 여러 개 있습니다. 가져올 날짜의 번호를 쓰세요." & vbLf
            If D <> gDate Then msg = "문서 날짜(" & FmtDateK(gDate) & ")에는 일정이 없습니다. 가져올 날짜의 번호를 쓰세요. 가져오면 문서 날짜도 바뀝니다." & vbLf
            For k = 1 To nD
                msg = msg & vbLf & k & ") " & FmtDateK(dates(k)) & " (" & cnt(k) & "줄)"
                If dates(k) = D Then defIdx = k
            Next
            msg = msg & vbLf & vbLf & "(날짜가 없는 줄은 늘 가져옵니다)"
            ans = Ask(msg, CStr(defIdx))
            If TrimX(ans) = "" Then Exit Function
            If Not IsNumeric(ans) Then Exit Function
            If CLng(ans) < 1 Or CLng(ans) > nD Then Exit Function
            D = dates(CLng(ans))
        End If
    End If
    ' entries
    gStep = "import: entries"
    Dim on_() As Boolean, nSched() As Long, nRef() As Long, empties As Long, sched As Collection, refc As Collection, res As Variant
    ReDim on_(1 To gN): ReDim nSched(1 To gN): ReDim nRef(1 To gN)
    For g = 1 To gN
        Set sched = New Collection: Set refc = New Collection
        GroupLines g, D, sched, refc
        res = ParseSched(sched)
        nSched(g) = res(0).Count
        nRef(g) = refc.Count + res(1).Count
        on_(g) = (sched.Count + refc.Count > 0)
        If Not on_(g) Then empties = empties + 1
    Next
    gStep = "import: seats"
    ' seats for groups not found
    Dim seatList As String
    For k = 1 To NX
        seatList = seatList & vbLf & k & ". " & ExecLabel(k)
    Next
    For g = 1 To gN
        If on_(g) And gTarget(g) < 0 Then
            ans = Ask("'" & gSource(g) & "' — 맞는 임원을 찾지 못했습니다." & vbLf & "넣을 자리 번호(1~8)를 쓰세요. 비워 두면 넣지 않습니다." & vbLf & seatList, "")
            If IsNumeric(ans) Then
                If CLng(ans) >= 1 And CLng(ans) <= NX Then gTarget(g) = CLng(ans)
            End If
            If gTarget(g) < 0 Then on_(g) = False
        End If
    Next
    gStep = "import: summary"
    ' summary
    Dim used(1 To 8) As Long, n As Long, lines As String, miss As String
    For g = 1 To gN
        If on_(g) Then
            If used(gTarget(g)) = 0 Then n = n + 1
            used(gTarget(g)) = used(gTarget(g)) + 1
            lines = lines & vbLf & "· " & ExecLabel(gTarget(g)) & " ← " & gSource(g) & " (일정 " & nSched(g) & "건" & IIf(nRef(g) > 0, " · ※ " & nRef(g) & "줄", "") & ")"
        End If
    Next
    If n = 0 Then
        Info "이 날짜에는 가져올 일정이 없습니다."
        Exit Function
    End If
    For k = 1 To NX
        If used(k) = 0 Then miss = miss & IIf(miss = "", "", ", ") & ExecLabel(k)
    Next
    msg = srcLabel & vbLf & n & "명에게 넣습니다. 고른 임원의 일정·※ 참고는 엑셀 내용으로 바뀝니다." & vbLf & lines
    If miss <> "" Then msg = msg & vbLf & vbLf & "그대로 두는 임원: " & miss
    If empties > 0 Then msg = msg & vbLf & "비어 있는 시트 " & empties & "개는 뺐습니다."
    If D <> "" Then msg = msg & vbLf & "날짜: " & FmtDateK(D)
    msg = msg & vbLf & vbLf & "가져올까요?"
    If Len(msg) > 1000 Then msg = Left$(msg, 990) & "…"
    If Not Confirm(msg) Then Exit Function
    ' apply
    gStep = "import: backup"
    SaveBackup "엑셀 불러오기"
    gStep = "import: write"
    BeginWork
    For k = 1 To NX
        If used(k) > 0 Then ClearExecRows exWs(k)
    Next
    Dim nextRow(1 To 8) As Long
    For k = 1 To NX
        nextRow(k) = 4
    Next
    For g = 1 To gN
        If on_(g) Then WriteGroupRows g, D, exWs(gTarget(g)), nextRow(gTarget(g))
    Next
    If D <> "" And D <> gDate Then NSet "d_date", IsoToDate(D)
    DoBuild
    EndWork
    Info "엑셀에서 " & n & "명의 일정을 가져왔습니다." & vbLf & "잘못 가져왔으면 입력 시트의 [되돌리기]를 누르세요."
    OpenImport = "ok"
End Function

Private Sub ClearExecRows(ByVal ws As Object)
    Dim lastR As Long, lastC As Long
    UsedLast ws, lastR, lastC
    If lastR < 4 Then lastR = 4
    If lastC < 4 Then lastC = 4
    ws.Range(ws.Cells(4, 1), ws.Cells(lastR, lastC)).ClearContents
End Sub

' write rows of group g into an executive sheet (A 시간, B 내용, C 비고, D ※ 참고)
Private Sub WriteGroupRows(ByVal g As Long, ByVal D As String, ByVal ws As Object, ByRef rowN As Long)
    Dim rw As Variant, cells() As String, roles() As String, labels() As String, i As Long
    Dim tm As String, content As String, place As String, stars As String, refs As String
    Dim sched As Collection, refc As Collection, l As Variant
    roles = Split(gRoles(g), SEP1): labels = Split(gLabels(g), SEP1)
    For Each rw In gRowsC(g)
        If rw(0) = "" Or D = "" Or rw(0) = D Then
            cells = Split(rw(1), SEP1)
            If gHasRoles(g) Then
                tm = "": content = "": place = "": stars = "": refs = ""
                For i = LBound(cells) To UBound(cells)
                    If cells(i) <> "" Then
                        Select Case roles(i)
                            Case "time": tm = cells(i)
                            Case "content": content = content & IIf(content = "", "", vbLf) & cells(i)
                            Case "place": place = cells(i)
                            Case "ref": refs = refs & IIf(refs = "", "", vbLf) & cells(i)
                            Case "star"
                                If labels(i) <> "" Then
                                    stars = stars & IIf(stars = "", "", vbLf) & labels(i) & ": " & cells(i)
                                Else
                                    stars = stars & IIf(stars = "", "", vbLf) & cells(i)
                                End If
                        End Select
                    End If
                Next
                If place <> "" Then content = AttachPlace(content, place)
                WriteExecRow ws, rowN, tm, content, stars, refs
            Else
                Set sched = New Collection: Set refc = New Collection
                RowLines cells, roles, labels, False, sched, refc
                For Each l In sched
                    WriteExecRow ws, rowN, "", CStr(l), "", ""
                Next
                For Each l In refc
                    WriteExecRow ws, rowN, "", "", "", CStr(l)
                Next
            End If
        End If
    Next
End Sub

Private Function AttachPlace(ByVal content As String, ByVal place As String) As String
    Dim ls As Variant, i As Long
    ls = Split(content, vbLf)
    For i = 0 To UBound(ls)
        If TrimX(CStr(ls(i))) <> "" Then
            If FirstIn(TrimX(CStr(ls(i))), "*＊※") Then
                AttachPlace = "@" & place & vbLf & content
            Else
                ls(i) = CStr(ls(i)) & " @" & place
                AttachPlace = Join(ls, vbLf)
            End If
            Exit Function
        End If
    Next
    AttachPlace = "@" & place
End Function

Private Sub WriteExecRow(ByVal ws As Object, ByRef rowN As Long, ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String)
    If a = "" And b = "" And c = "" And d = "" Then Exit Sub
    ws.Range(ws.Cells(rowN, 1), ws.Cells(rowN, 4)).NumberFormat = "@"
    ws.Cells(rowN, 1).Value = a
    ws.Cells(rowN, 2).Value = b
    ws.Cells(rowN, 3).Value = c
    ws.Cells(rowN, 4).Value = d
    rowN = rowN + 1
End Sub

' ===========================================================================
' Backup / undo
' ===========================================================================
Private Sub SaveBackup(ByVal reason As String)
    Dim b As Object, k As Long, ws As Object, lastR As Long, lastC As Long, v As Variant, tr As Object
    Set b = ScratchSheet(SH_BAK)
    b.Cells.Clear
    b.Cells(1, 1).Value = reason
    b.Cells(1, 2).Value = Now
    b.Cells(2, 1).Value = NVal("d_date")
    For k = 1 To NX
        Set ws = exWs(k)
        UsedLast ws, lastR, lastC
        If lastC < 4 Then lastC = 4
        If lastC > 10 Then lastC = 10
        If lastR < 4 Then lastR = 4
        b.Cells(3, (k - 1) * 11 + 1).Value = lastR
        v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
        b.Range(b.Cells(4, (k - 1) * 11 + 1), b.Cells(4 + lastR - 1, (k - 1) * 11 + lastC)).NumberFormat = "@"
        b.Range(b.Cells(4, (k - 1) * 11 + 1), b.Cells(4 + lastR - 1, (k - 1) * 11 + lastC)).Value = v
    Next
    Set tr = NameRange("t_top")
    If Not tr Is Nothing Then
        v = tr.Resize(9, 4).Value
        b.Range(b.Cells(4, 89), b.Cells(12, 92)).NumberFormat = "@"
        b.Range(b.Cells(4, 89), b.Cells(12, 92)).Value = v
    End If
End Sub

Public Sub UndoLast()
    Dim b As Object, k As Long, ws As Object, lastR As Long, v As Variant, tr As Object
    On Error GoTo fail
    Set b = SheetByName(SH_BAK)
    If b Is Nothing Then
        Info "되돌릴 내용이 없습니다."
        Exit Sub
    End If
    If CStr(b.Cells(1, 1).Value) = "" Then
        Info "되돌릴 내용이 없습니다."
        Exit Sub
    End If
    If Not Confirm("'" & b.Cells(1, 1).Value & "' 하기 전으로 되돌릴까요?") Then Exit Sub
    BeginWork
    LoadExecs
    For k = 1 To NX
        Set ws = exWs(k)
        lastR = CLng(b.Cells(3, (k - 1) * 11 + 1).Value)
        If lastR >= 4 Then
            ClearExecRows ws
            v = b.Range(b.Cells(4, (k - 1) * 11 + 1), b.Cells(4 + lastR - 1, (k - 1) * 11 + 10)).Value
            ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 10)).NumberFormat = "@"
            ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 10)).Value = v
        End If
    Next
    Set tr = NameRange("t_top")
    If Not tr Is Nothing Then
        tr.Resize(9, 4).ClearContents
        tr.Resize(9, 4).Value = b.Range(b.Cells(4, 89), b.Cells(12, 92)).Value
    End If
    If Not IsEmpty(b.Cells(2, 1).Value) Then NSet "d_date", b.Cells(2, 1).Value
    b.Cells(1, 1).Value = ""
    DoBuild
    EndWork
    Info "되돌렸습니다."
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

' ===========================================================================
' New document, template, highlight, seats, presets, password
' ===========================================================================
Public Sub NewDoc()
    On Error GoTo fail
    If Not Confirm("8명 모두의 일정·※ 참고 내용과 맨 밑 표 내용을 비웁니다." & vbLf & "(직위·성명·회사명·표 이름·서식은 그대로 두고, 날짜는 오늘로 바뀝니다)") Then Exit Sub
    BeginWork
    LoadSettings
    LoadExecs
    SaveBackup "새 문서"
    Dim k As Long, tr As Object
    For k = 1 To NX
        ClearExecRows exWs(k)
    Next
    Set tr = NameRange("t_top")
    If Not tr Is Nothing Then tr.Offset(0, 3).Resize(9, 1).ClearContents
    NSet "d_date", Date
    DoBuild
    EndWork
    Info "내용을 비웠습니다. 잘못 눌렀으면 [되돌리기]를 누르세요."
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub MakeTemplate()
    Dim p As String
    On Error GoTo fail
    p = PickSavePath("임원일일예정_공동편집양식.xlsx", "엑셀 파일 (*.xlsx),*.xlsx")
    If p = "" Then Exit Sub
    MakeTemplateTo p
    Exit Sub
fail:
    Application.DisplayAlerts = True
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Function MakeTemplateTo(ByVal path As String) As String
    Dim wb As Object, ws As Object, k As Long, guide As Variant, i As Long
    LoadSettings
    LoadExecs
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Set ws = wb.Worksheets(1)
    ws.Name = "작성방법"
    guide = Array("임원 일일 예정사항 — 공동편집 양식", "", _
        "1. 임원마다 시트가 하나씩 있습니다. 시트 이름(직위 성명)으로 임원을 찾으니 바꾸지 마세요.", _
        "2. 한 줄에 일정 하나: 시간 / 내용을 각자 칸에 씁니다. 시간은 09:00 또는 14:00~15:30처럼 씁니다(종일·오전도 됩니다).", _
        "3. 비고 칸에 쓴 것은 그 일정 밑에 작은 글씨 부가 설명(*)으로 붙습니다.", _
        "4. 내용 칸 안에서 줄을 바꿔(Alt+Enter) ""* ""로 시작하면 부가 설명, ""- ""로 시작하면 하위 항목입니다. 내용이 ""* ""로 시작하는 줄은 바로 위 일정에 붙습니다.", _
        "5. ※ 참고 칸에 쓴 것은 그 임원 칸 맨 밑에 ※ 표시로 인쇄됩니다. 어느 줄에 써도 됩니다.", _
        "6. 여러 날을 한 시트에 쓰려면 맨 앞에 ""날짜"" 열을 넣으세요. 문서 날짜에 맞는 줄만 가져옵니다.", _
        "7. 다 쓰면 담당자가 파일을 내려받아 한셀 파일의 [엑셀 불러오기]로 가져옵니다.", "", "예시")
    For i = 0 To UBound(guide)
        ws.Cells(i + 1, 1).Value = guide(i)
    Next
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(UBound(guide) + 1, 1).Font.Bold = True
    TemplateHeader ws, UBound(guide) + 2
    Dim ex As Variant
    ex = Array(Array("08:30", "임원 조찬 간담회 @20층 라운지", "참석: 부사장 이하 임원 전원", "11/10(화) 싱가포르 출장"), _
        Array("10:00", "주간 경영회의 @대회의실" & vbLf & "- 안건: 4분기 실적 전망", "", ""), _
        Array("14:00~16:00", "노사협의회 정기회의", "노측 6명, 사측 6명", ""), _
        Array("", "* 회의 후 현장 점검", "", ""))
    For i = 0 To UBound(ex)
        ws.Range(ws.Cells(UBound(guide) + 3 + i, 1), ws.Cells(UBound(guide) + 3 + i, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(UBound(guide) + 3 + i, 1), ws.Cells(UBound(guide) + 3 + i, 4)).Value = ex(i)
        ws.Range(ws.Cells(UBound(guide) + 3 + i, 1), ws.Cells(UBound(guide) + 3 + i, 4)).Borders.LineStyle = XL_CONT
        ws.Range(ws.Cells(UBound(guide) + 3 + i, 1), ws.Cells(UBound(guide) + 3 + i, 4)).WrapText = True
    Next
    TemplateWidths ws
    Dim prev As Object, nm As String
    Set prev = ws
    For k = 1 To NX
        Set ws = wb.Worksheets.Add(, prev)
        nm = SafeSheetName(ExecLabel(k))
        On Error Resume Next
        ws.Name = nm
        If Err.Number <> 0 Then ws.Name = Left$(nm, 27) & "(" & k & ")"
        On Error GoTo 0
        TemplateHeader ws, 1
        ws.Range("A2:D300").NumberFormat = "@"
        ws.Range("A2:D300").WrapText = True
        ws.Range("A2:D300").VerticalAlignment = XL_TOP
        TemplateWidths ws
        Set prev = ws
    Next
    wb.Worksheets(2).Activate
    On Error Resume Next
    wb.SaveAs path, 51
    If Err.Number <> 0 Then
        Err.Clear
        wb.SaveAs path
    End If
    If Err.Number <> 0 Then
        Application.DisplayAlerts = True
        Info "양식을 새 창으로 만들었습니다. 그 창에서 [다른 이름으로 저장]으로 .xlsx 로 저장하세요."
        Exit Function
    End If
    wb.Close False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Info "공동편집용 엑셀 양식을 저장했습니다." & vbLf & path & vbLf & "두레이에 올려 함께 작성하세요."
    MakeTemplateTo = "ok"
End Function

Private Sub TemplateHeader(ByVal ws As Object, ByVal r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Value = Array("시간", "내용", "비고 (작은 글씨 *)", "※ 참고")
        .Font.Bold = True
        .Interior.Color = COL_HEAD
        .HorizontalAlignment = XL_CENTER
        .Borders.LineStyle = XL_CONT
    End With
End Sub

Private Sub TemplateWidths(ByVal ws As Object)
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 54
    ws.Columns(3).ColumnWidth = 32
    ws.Columns(4).ColumnWidth = 34
End Sub

' ==text== on the selected cells (again: remove)
Public Sub HighlightSel()
    Dim c As Object, s As String, n As Long
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    On Error GoTo 0
    For Each c In Selection.Cells
        If Not IsEmpty(c.Value) And c.Column > 1 Then
            s = CStr(c.Value)
            If InStr(s, "==") > 0 Then
                s = Replace(s, "==", "")
            Else
                s = MarkLines(s)
            End If
            c.NumberFormat = "@"
            c.Value = s
            n = n + 1
        End If
        If n > 500 Then Exit For
    Next
    If n = 0 Then Info "칠할 칸(내용·비고·※ 참고·맨 밑 표 내용)을 먼저 고르세요. 칠한 칸을 다시 고르고 누르면 지워집니다."
End Sub

Public Function MarkLines(ByVal s As String) As String
    Dim ls As Variant, i As Long, a As Long, b As Long, ln As String
    ls = SplitLines(s)
    For i = 0 To UBound(ls)
        ln = CStr(ls(i))
        ContentRange ln, a, b
        Do While a < b
            If Not IsWs(Mid$(ln, a + 1, 1)) Then Exit Do
            a = a + 1
        Loop
        Do While b > a
            If Not IsWs(Mid$(ln, b, 1)) Then Exit Do
            b = b - 1
        Loop
        If b > a Then ls(i) = Left$(ln, a) & "==" & Mid$(ln, a + 1, b - a) & "==" & Mid$(ln, b + 1)
    Next
    MarkLines = Join(ls, vbLf)
End Function

' 0-based [a, b) of the part of a line that may be highlighted (port of contentRange)
Private Sub ContentRange(ByVal ln As String, ByRef a As Long, ByRef b As Long)
    Dim t As String, k As Long, sp As String, rest As String, p As Long
    a = SkipWs(ln, 1) - 1
    t = Mid$(ln, a + 1)
    k = InStr(t, vbTab)
    If k > 1 And k <= 17 And Not IsBullet(t) Then
        a = a + k
    ElseIf Left$(t, 1) = "※" Then
        a = a + (SkipWs(t, 2) - 1)
    ElseIf FirstIn(t, "*＊") Then
        p = 1
        Do While FirstIn(Mid$(t, p), "*＊")
            p = p + 1
        Loop
        a = a + (SkipWs(t, p) - 1)
    Else
        If IsBullet(t) Then
            p = SkipWs(t, 2)
            a = a + p - 1
            t = Mid$(t, p)
        End If
        If SplitTime(t, sp, rest) Then a = a + Len(RTrimX(t)) - Len(rest)
    End If
    b = Len(RTrimX(ln))
    p = InStrRev(Left$(ln, b), "@")
    If p > 1 Then
        If IsWs(Mid$(ln, p - 1, 1)) And TrimX(Mid$(ln, p + 1, b - p)) <> "" Then
            Do While p > 1
                If Not IsWs(Mid$(ln, p - 1, 1)) Then Exit Do
                p = p - 1
            Loop
            If p - 1 >= a Then b = p - 1
        End If
    End If
    If b < a Then b = a
End Sub

Public Sub SwapSeats()
    Dim ans As String, parts As Variant, i As Long, j As Long
    On Error GoTo fail
    LoadExecs
    Dim seatList As String, k As Long
    For k = 1 To NX
        seatList = seatList & vbLf & k & ". " & SlotLabel(k) & ": " & ExecLabel(k)
    Next
    ans = Ask("자리를 바꿀 두 칸의 번호를 쓰세요 (예: 1 4)" & vbLf & seatList, "")
    ans = TrimX(Replace(Replace(ans, ",", " "), "-", " "))
    If ans = "" Then Exit Sub
    parts = Split(CollapseWs(ans), " ")
    If UBound(parts) < 1 Then Exit Sub
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Sub
    i = CLng(parts(0)): j = CLng(parts(1))
    If i < 1 Or i > NX Or j < 1 Or j > NX Or i = j Then Exit Sub
    BeginWork
    SwapSeatsIJ i, j
    DoBuild
    EndWork
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub SwapSeatsIJ(ByVal i As Long, ByVal j As Long)
    Dim a As Object, b As Object, ra As Long, rb As Long, ca As Long, cb As Long, n As Long, va As Variant, vb As Variant, na As String, nb As String
    Set a = exWs(i): Set b = exWs(j)
    UsedLast a, ra, ca
    UsedLast b, rb, cb
    n = MaxD(ra, rb): If n < 4 Then n = 4
    If ca < cb Then ca = cb
    If ca < 4 Then ca = 4
    va = a.Range(a.Cells(1, 1), a.Cells(n, ca)).Value
    vb = b.Range(b.Cells(1, 1), b.Cells(n, ca)).Value
    ' clear first: writing an array with empty items does not clear cells in every host
    a.Range(a.Cells(4, 1), a.Cells(n, ca)).ClearContents
    b.Range(b.Cells(4, 1), b.Cells(n, ca)).ClearContents
    a.Range(a.Cells(1, 1), a.Cells(n, ca)).Value = vb
    b.Range(b.Cells(1, 1), b.Cells(n, ca)).Value = va
    na = a.Name: nb = b.Name
    a.Name = "~swap~"
    b.Name = na
    a.Name = nb
End Sub

Private Sub SetCols(ByVal c1a As Variant, ByVal c1b As Variant, ByVal c1c As Variant, ByVal c2a As Variant, ByVal c2b As Variant, ByVal c3a As Variant, ByVal c3b As Variant)
    If Not IsEmpty(c1a) Then NSet "s_c1a", c1a: NSet "s_c1b", c1b: NSet "s_c1c", c1c
    If Not IsEmpty(c2a) Then NSet "s_c2a", c2a: NSet "s_c2b", c2b
    If Not IsEmpty(c3a) Then NSet "s_c3a", c3a: NSet "s_c3b", c3b
End Sub

Public Sub PresetBig23()
    SetCols Empty, Empty, Empty, 2, 1, Empty, Empty
    BuildA4
End Sub

Public Sub PresetBig12()
    SetCols Empty, Empty, Empty, 1, 1, Empty, Empty
    BuildA4
End Sub

Public Sub PresetAutoH()
    NSet "s_h1", "": NSet "s_h2", "": NSet "s_h3", "": NSet "s_split", ""
    BuildA4
End Sub

Public Sub PresetReset()
    SetCols 1, 1, 1, 2, 1, 1, 1
    NSet "s_h1", "": NSet "s_h2", "": NSet "s_h3", "": NSet "s_split", ""
    BuildA4
End Sub

Public Sub SetDocPassword()
    Dim p1 As String, p2 As String
    p1 = Ask("열 때 물어볼 문서 암호를 쓰세요. (비워 두면 암호를 없앱니다)", "")
    If p1 <> "" Then
        p2 = Ask("확인을 위해 한 번 더 쓰세요.", "")
        If p1 <> p2 Then
            Info "두 번 쓴 암호가 서로 다릅니다. 다시 해 주세요."
            Exit Sub
        End If
    End If
    On Error Resume Next
    HostBook().Password = p1
    If Err.Number = 0 Then HostBook().Save
    If Err.Number <> 0 Then
        Info "매크로로 문서 암호를 정하지 못했습니다." & vbLf & "한셀 메뉴의 [문서 암호](보안·검토 메뉴나 파일 → 다른 이름으로 저장의 도구)에서 정하세요."
    ElseIf p1 = "" Then
        Info "문서 암호를 없앴습니다."
    Else
        Info "문서 암호를 정하고 저장했습니다. 다음에 열 때부터 암호를 묻습니다." & vbLf & "암호를 잊으면 열 수 없으니 따로 적어 두세요."
    End If
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "^+h", "HighlightSel"
    Application.OnKey "^+H", "HighlightSel"
End Sub

' 이 모듈을 한셀이 읽었는지 확인 (실행했을 때 창이 뜨면 정상)
Public Sub CheckHcExecIO()
    MsgBox "HcExecIO OK"
End Sub
