Attribute VB_Name = "HcExecIO"
Option Explicit

' ===========================================================================
'
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
    For Each p In Array(U("\uC791\uC131\uBC29\uBC95"), U("\uC791\uC131\uC694\uB839"), U("\uC791\uC131\uC608\uC2DC"), U("\uC548\uB0B4"), U("\uB3C4\uC6C0\uB9D0"), U("\uC0AC\uC6A9\uBC95"), U("\uC608\uC2DC"), U("\uBAA9\uB85D"), U("\uC124\uC815"), "readme", "guide", "help")
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
                        g = GroupAdd(TrimX(curPos & " " & curExec), title, source & IIf(TrimX(curPos & " " & curExec) <> "", U(" \u203A ") & TrimX(curPos & " " & curExec), ""), True, hasRoles, roles, labels)
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
    NormKey = LCase$(RemoveChars(RemoveParens(s), U("\u00B7.,_-")))
End Function

'
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
    If ExecLabel = "" Then ExecLabel = i & U("\uBC88 \uCE78")
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
    On Error Resume Next
    f = Application.GetOpenFilename(U("\uC5D1\uC140 \uD30C\uC77C (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv"), 1, U("\uB450\uB808\uC774\uC5D0\uC11C \uB0B4\uB824\uBC1B\uC740 \uC5D1\uC140 \uD30C\uC77C\uC744 \uACE0\uB974\uC138\uC694"), , True)
    If Err.Number <> 0 Then
        Err.Clear
        f = Application.GetOpenFilename(U("\uC5D1\uC140 \uD30C\uC77C (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv"))
    End If
    On Error GoTo fail
    If IsArray(f) Then
        ImportFiles f
    ElseIf VarType(f) = vbString Then
        If f <> "" And LCase$(f) <> "false" Then ImportFiles Array(f)
    End If
    Exit Sub
fail:
    EndWork
    Info U("\uBB38\uC81C\uAC00 \uC0DD\uACBC\uC2B5\uB2C8\uB2E4: ") & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Function ImportFiles(files As Variant) As String
    Dim p As Variant, wb As Object, ws As Object, v As Variant, names As String, fn As String
    LoadSettings
    LoadExecs
    gStep = "import: reset"
    GroupsReset
    For Each p In files
        If LCase$(CStr(p)) = LCase$(ThisWorkbook.FullName) Then
            Info U("\uC9C0\uAE08 \uC5F4\uB824 \uC788\uB294 \uC774 \uD30C\uC77C\uC740 \uBD88\uB7EC\uC62C \uC218 \uC5C6\uC2B5\uB2C8\uB2E4. \uB450\uB808\uC774\uC5D0\uC11C \uB0B4\uB824\uBC1B\uC740 \uACF5\uB3D9\uD3B8\uC9D1 \uD30C\uC77C\uC744 \uACE0\uB974\uC138\uC694.")
            Exit Function
        End If
        fn = Mid$(CStr(p), InStrRev(Replace(CStr(p), "\", "/"), "/") + 1)
        gStep = "import: open " & fn
        On Error Resume Next
        Set wb = Nothing
        Set wb = Application.Workbooks.Open(CStr(p), 0, True)
        On Error GoTo 0
        If wb Is Nothing Then
            Info fn & U(": \uD30C\uC77C\uC744 \uC5F4 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.")
        Else
            If names <> "" Then names = names & ", "
            names = names & fn
            For Each ws In wb.Worksheets
                gStep = "import: sheet " & ws.Name
                If ws.Visible = XL_SHEET_VISIBLE And Not SkipSheetName(ws.Name) Then
                    v = SheetValues(ws)
                    gStep = "import: groups " & ws.Name
                    SheetGroups ws.Name, v, 1, fn & U(" \u203A ") & ws.Name
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
        Info U("\uBD99\uC5EC\uB123\uAE30 \uC2DC\uD2B8\uC758 3\uBC88\uC9F8 \uC904\uBD80\uD130 \uC5D1\uC140\uC5D0\uC11C \uBCF5\uC0AC\uD55C \uD45C\uB97C \uBD99\uC5EC\uB123\uC740 \uB4A4 \uB2E4\uC2DC \uB204\uB974\uC138\uC694.")
        ws.Activate
        Exit Sub
    End If
    LoadSettings
    LoadExecs
    GroupsReset
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, MaxD(lastC, 2))).Value
    SheetGroups U("\uBD99\uC5EC\uB123\uC740 \uB0B4\uC6A9"), v, 3, U("\uBD99\uC5EC\uB123\uC740 \uB0B4\uC6A9")
    If OpenImport(U("\uBD99\uC5EC\uB123\uC740 \uB0B4\uC6A9")) = "ok" Then ws.Range(ws.Cells(3, 1), ws.Cells(lastR, MaxD(lastC, 2))).ClearContents
    Exit Sub
fail:
    EndWork
    Info U("\uBB38\uC81C\uAC00 \uC0DD\uACBC\uC2B5\uB2C8\uB2E4: ") & Err.Description & vbLf & "(" & gStep & ")"
End Sub

' choose date / seats, confirm, apply. returns "ok" when applied
Private Function OpenImport(ByVal srcLabel As String) As String
    Dim g As Long, rw As Variant, D As String, k As Long, c As Long
    Dim bestC As Long, msg As String, ans As String, defIdx As Long
    Dim dates() As String, cnt() As Long, nD As Long, t As String, j As Long, tc As Long
    ReDim dates(1 To 1): ReDim cnt(1 To 1)
    If gN = 0 Then
        Info U("\uAC00\uC838\uC62C \uC77C\uC815\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4. \uBA38\uB9AC\uAE00(\uC2DC\uAC04\u00B7\uB0B4\uC6A9)\uC774 \uC788\uB294\uC9C0 \uD655\uC778\uD558\uC138\uC694.")
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
            msg = U("\uC5D1\uC140\uC5D0 \uB0A0\uC9DC\uAC00 \uC5EC\uB7EC \uAC1C \uC788\uC2B5\uB2C8\uB2E4. \uAC00\uC838\uC62C \uB0A0\uC9DC\uC758 \uBC88\uD638\uB97C \uC4F0\uC138\uC694.") & vbLf
            If D <> gDate Then msg = U("\uBB38\uC11C \uB0A0\uC9DC(") & FmtDateK(gDate) & U(")\uC5D0\uB294 \uC77C\uC815\uC774 \uC5C6\uC2B5\uB2C8\uB2E4. \uAC00\uC838\uC62C \uB0A0\uC9DC\uC758 \uBC88\uD638\uB97C \uC4F0\uC138\uC694. \uAC00\uC838\uC624\uBA74 \uBB38\uC11C \uB0A0\uC9DC\uB3C4 \uBC14\uB01D\uB2C8\uB2E4.") & vbLf
            For k = 1 To nD
                msg = msg & vbLf & k & ") " & FmtDateK(dates(k)) & " (" & cnt(k) & U("\uC904)")
                If dates(k) = D Then defIdx = k
            Next
            msg = msg & vbLf & vbLf & U("(\uB0A0\uC9DC\uAC00 \uC5C6\uB294 \uC904\uC740 \uB298 \uAC00\uC838\uC635\uB2C8\uB2E4)")
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
            ans = Ask("'" & gSource(g) & U("' \u2014 \uB9DE\uB294 \uC784\uC6D0\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.") & vbLf & U("\uB123\uC744 \uC790\uB9AC \uBC88\uD638(1~8)\uB97C \uC4F0\uC138\uC694. \uBE44\uC6CC \uB450\uBA74 \uB123\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4.") & vbLf & seatList, "")
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
            lines = lines & vbLf & U("\u00B7 ") & ExecLabel(gTarget(g)) & U(" \u2190 ") & gSource(g) & U(" (\uC77C\uC815 ") & nSched(g) & U("\uAC74") & IIf(nRef(g) > 0, U(" \u00B7 \u203B ") & nRef(g) & U("\uC904"), "") & ")"
        End If
    Next
    If n = 0 Then
        Info U("\uC774 \uB0A0\uC9DC\uC5D0\uB294 \uAC00\uC838\uC62C \uC77C\uC815\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.")
        Exit Function
    End If
    For k = 1 To NX
        If used(k) = 0 Then miss = miss & IIf(miss = "", "", ", ") & ExecLabel(k)
    Next
    msg = srcLabel & vbLf & n & U("\uBA85\uC5D0\uAC8C \uB123\uC2B5\uB2C8\uB2E4. \uACE0\uB978 \uC784\uC6D0\uC758 \uC77C\uC815\u00B7\u203B \uCC38\uACE0\uB294 \uC5D1\uC140 \uB0B4\uC6A9\uC73C\uB85C \uBC14\uB01D\uB2C8\uB2E4.") & vbLf & lines
    If miss <> "" Then msg = msg & vbLf & vbLf & U("\uADF8\uB300\uB85C \uB450\uB294 \uC784\uC6D0: ") & miss
    If empties > 0 Then msg = msg & vbLf & U("\uBE44\uC5B4 \uC788\uB294 \uC2DC\uD2B8 ") & empties & U("\uAC1C\uB294 \uBE90\uC2B5\uB2C8\uB2E4.")
    If D <> "" Then msg = msg & vbLf & U("\uB0A0\uC9DC: ") & FmtDateK(D)
    msg = msg & vbLf & vbLf & U("\uAC00\uC838\uC62C\uAE4C\uC694?")
    If Len(msg) > 1000 Then msg = Left$(msg, 990) & U("\u2026")
    If Not Confirm(msg) Then Exit Function
    ' apply
    gStep = "import: backup"
    SaveBackup U("\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30")
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
    Info U("\uC5D1\uC140\uC5D0\uC11C ") & n & U("\uBA85\uC758 \uC77C\uC815\uC744 \uAC00\uC838\uC654\uC2B5\uB2C8\uB2E4.") & vbLf & U("\uC798\uBABB \uAC00\uC838\uC654\uC73C\uBA74 \uC785\uB825 \uC2DC\uD2B8\uC758 [\uB418\uB3CC\uB9AC\uAE30]\uB97C \uB204\uB974\uC138\uC694.")
    OpenImport = "ok"
End Function

Private Sub ClearExecRows(ByVal ws As Object)
    Dim lastR As Long, lastC As Long
    UsedLast ws, lastR, lastC
    If lastR < 4 Then lastR = 4
    If lastC < 4 Then lastC = 4
    ws.Range(ws.Cells(4, 1), ws.Cells(lastR, lastC)).ClearContents
End Sub

'
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
            If FirstIn(TrimX(CStr(ls(i))), U("*\uFF0A\u203B")) Then
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
        Info U("\uB418\uB3CC\uB9B4 \uB0B4\uC6A9\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.")
        Exit Sub
    End If
    If CStr(b.Cells(1, 1).Value) = "" Then
        Info U("\uB418\uB3CC\uB9B4 \uB0B4\uC6A9\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.")
        Exit Sub
    End If
    If Not Confirm("'" & b.Cells(1, 1).Value & U("' \uD558\uAE30 \uC804\uC73C\uB85C \uB418\uB3CC\uB9B4\uAE4C\uC694?")) Then Exit Sub
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
    Info U("\uB418\uB3CC\uB838\uC2B5\uB2C8\uB2E4.")
    Exit Sub
fail:
    EndWork
    Info U("\uBB38\uC81C\uAC00 \uC0DD\uACBC\uC2B5\uB2C8\uB2E4: ") & Err.Description & vbLf & "(" & gStep & ")"
End Sub

' ===========================================================================
' New document, template, highlight, seats, presets, password
' ===========================================================================
Public Sub NewDoc()
    On Error GoTo fail
    If Not Confirm(U("8\uBA85 \uBAA8\uB450\uC758 \uC77C\uC815\u00B7\u203B \uCC38\uACE0 \uB0B4\uC6A9\uACFC \uB9E8 \uBC11 \uD45C \uB0B4\uC6A9\uC744 \uBE44\uC6C1\uB2C8\uB2E4.") & vbLf & U("(\uC9C1\uC704\u00B7\uC131\uBA85\u00B7\uD68C\uC0AC\uBA85\u00B7\uD45C \uC774\uB984\u00B7\uC11C\uC2DD\uC740 \uADF8\uB300\uB85C \uB450\uACE0, \uB0A0\uC9DC\uB294 \uC624\uB298\uB85C \uBC14\uB01D\uB2C8\uB2E4)")) Then Exit Sub
    BeginWork
    LoadSettings
    LoadExecs
    SaveBackup U("\uC0C8 \uBB38\uC11C")
    Dim k As Long, tr As Object
    For k = 1 To NX
        ClearExecRows exWs(k)
    Next
    Set tr = NameRange("t_top")
    If Not tr Is Nothing Then tr.Offset(0, 3).Resize(9, 1).ClearContents
    NSet "d_date", Date
    DoBuild
    EndWork
    Info U("\uB0B4\uC6A9\uC744 \uBE44\uC6E0\uC2B5\uB2C8\uB2E4. \uC798\uBABB \uB20C\uB800\uC73C\uBA74 [\uB418\uB3CC\uB9AC\uAE30]\uB97C \uB204\uB974\uC138\uC694.")
    Exit Sub
fail:
    EndWork
    Info U("\uBB38\uC81C\uAC00 \uC0DD\uACBC\uC2B5\uB2C8\uB2E4: ") & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub MakeTemplate()
    Dim p As Variant
    On Error GoTo fail
    On Error Resume Next
    p = Application.GetSaveAsFilename(U("\uC784\uC6D0\uC77C\uC77C\uC608\uC815_\uACF5\uB3D9\uD3B8\uC9D1\uC591\uC2DD.xlsx"), U("\uC5D1\uC140 \uD30C\uC77C (*.xlsx),*.xlsx"))
    On Error GoTo fail
    If VarType(p) = vbBoolean Then Exit Sub
    If CStr(p) = "" Or LCase$(CStr(p)) = "false" Then Exit Sub
    MakeTemplateTo CStr(p)
    Exit Sub
fail:
    EndWork
    Info U("\uBB38\uC81C\uAC00 \uC0DD\uACBC\uC2B5\uB2C8\uB2E4: ") & Err.Description & vbLf & "(" & gStep & ")"
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
    ws.Name = U("\uC791\uC131\uBC29\uBC95")
    guide = Array(U("\uC784\uC6D0 \uC77C\uC77C \uC608\uC815\uC0AC\uD56D \u2014 \uACF5\uB3D9\uD3B8\uC9D1 \uC591\uC2DD"), "", _
        U("1. \uC784\uC6D0\uB9C8\uB2E4 \uC2DC\uD2B8\uAC00 \uD558\uB098\uC529 \uC788\uC2B5\uB2C8\uB2E4. \uC2DC\uD2B8 \uC774\uB984(\uC9C1\uC704 \uC131\uBA85)\uC73C\uB85C \uC784\uC6D0\uC744 \uCC3E\uC73C\uB2C8 \uBC14\uAFB8\uC9C0 \uB9C8\uC138\uC694."), _
        U("2. \uD55C \uC904\uC5D0 \uC77C\uC815 \uD558\uB098: \uC2DC\uAC04 / \uB0B4\uC6A9\uC744 \uAC01\uC790 \uCE78\uC5D0 \uC501\uB2C8\uB2E4. \uC2DC\uAC04\uC740 09:00 \uB610\uB294 14:00~15:30\uCC98\uB7FC \uC501\uB2C8\uB2E4(\uC885\uC77C\u00B7\uC624\uC804\uB3C4 \uB429\uB2C8\uB2E4)."), _
        U("3. \uBE44\uACE0 \uCE78\uC5D0 \uC4F4 \uAC83\uC740 \uADF8 \uC77C\uC815 \uBC11\uC5D0 \uC791\uC740 \uAE00\uC528 \uBD80\uAC00 \uC124\uBA85(*)\uC73C\uB85C \uBD99\uC2B5\uB2C8\uB2E4."), _
        U("4. \uB0B4\uC6A9 \uCE78 \uC548\uC5D0\uC11C \uC904\uC744 \uBC14\uAFD4(Alt+Enter) ""* ""\uB85C \uC2DC\uC791\uD558\uBA74 \uBD80\uAC00 \uC124\uBA85, ""- ""\uB85C \uC2DC\uC791\uD558\uBA74 \uD558\uC704 \uD56D\uBAA9\uC785\uB2C8\uB2E4. \uB0B4\uC6A9\uC774 ""* ""\uB85C \uC2DC\uC791\uD558\uB294 \uC904\uC740 \uBC14\uB85C \uC704 \uC77C\uC815\uC5D0 \uBD99\uC2B5\uB2C8\uB2E4."), _
        U("5. \u203B \uCC38\uACE0 \uCE78\uC5D0 \uC4F4 \uAC83\uC740 \uADF8 \uC784\uC6D0 \uCE78 \uB9E8 \uBC11\uC5D0 \u203B \uD45C\uC2DC\uB85C \uC778\uC1C4\uB429\uB2C8\uB2E4. \uC5B4\uB290 \uC904\uC5D0 \uC368\uB3C4 \uB429\uB2C8\uB2E4."), _
        U("6. \uC5EC\uB7EC \uB0A0\uC744 \uD55C \uC2DC\uD2B8\uC5D0 \uC4F0\uB824\uBA74 \uB9E8 \uC55E\uC5D0 ""\uB0A0\uC9DC"" \uC5F4\uC744 \uB123\uC73C\uC138\uC694. \uBB38\uC11C \uB0A0\uC9DC\uC5D0 \uB9DE\uB294 \uC904\uB9CC \uAC00\uC838\uC635\uB2C8\uB2E4."), _
        U("7. \uB2E4 \uC4F0\uBA74 \uB2F4\uB2F9\uC790\uAC00 \uD30C\uC77C\uC744 \uB0B4\uB824\uBC1B\uC544 \uD55C\uC140 \uD30C\uC77C\uC758 [\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30]\uB85C \uAC00\uC838\uC635\uB2C8\uB2E4."), "", U("\uC608\uC2DC"))
    For i = 0 To UBound(guide)
        ws.Cells(i + 1, 1).Value = guide(i)
    Next
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(UBound(guide) + 1, 1).Font.Bold = True
    TemplateHeader ws, UBound(guide) + 2
    Dim ex As Variant
    ex = Array(Array("08:30", U("\uC784\uC6D0 \uC870\uCC2C \uAC04\uB2F4\uD68C @20\uCE35 \uB77C\uC6B4\uC9C0"), U("\uCC38\uC11D: \uBD80\uC0AC\uC7A5 \uC774\uD558 \uC784\uC6D0 \uC804\uC6D0"), U("11/10(\uD654) \uC2F1\uAC00\uD3EC\uB974 \uCD9C\uC7A5")), _
        Array("10:00", U("\uC8FC\uAC04 \uACBD\uC601\uD68C\uC758 @\uB300\uD68C\uC758\uC2E4") & vbLf & U("- \uC548\uAC74: 4\uBD84\uAE30 \uC2E4\uC801 \uC804\uB9DD"), "", ""), _
        Array("14:00~16:00", U("\uB178\uC0AC\uD611\uC758\uD68C \uC815\uAE30\uD68C\uC758"), U("\uB178\uCE21 6\uBA85, \uC0AC\uCE21 6\uBA85"), ""), _
        Array("", U("* \uD68C\uC758 \uD6C4 \uD604\uC7A5 \uC810\uAC80"), "", ""))
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
        Info U("\uC591\uC2DD\uC744 \uC0C8 \uCC3D\uC73C\uB85C \uB9CC\uB4E4\uC5C8\uC2B5\uB2C8\uB2E4. \uADF8 \uCC3D\uC5D0\uC11C [\uB2E4\uB978 \uC774\uB984\uC73C\uB85C \uC800\uC7A5]\uC73C\uB85C .xlsx \uB85C \uC800\uC7A5\uD558\uC138\uC694.")
        Exit Function
    End If
    wb.Close False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Info U("\uACF5\uB3D9\uD3B8\uC9D1\uC6A9 \uC5D1\uC140 \uC591\uC2DD\uC744 \uC800\uC7A5\uD588\uC2B5\uB2C8\uB2E4.") & vbLf & path & vbLf & U("\uB450\uB808\uC774\uC5D0 \uC62C\uB824 \uD568\uAED8 \uC791\uC131\uD558\uC138\uC694.")
    MakeTemplateTo = "ok"
End Function

Private Sub TemplateHeader(ByVal ws As Object, ByVal r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Value = Array(U("\uC2DC\uAC04"), U("\uB0B4\uC6A9"), U("\uBE44\uACE0 (\uC791\uC740 \uAE00\uC528 *)"), U("\u203B \uCC38\uACE0"))
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
    If n = 0 Then Info U("\uCE60\uD560 \uCE78(\uB0B4\uC6A9\u00B7\uBE44\uACE0\u00B7\u203B \uCC38\uACE0\u00B7\uB9E8 \uBC11 \uD45C \uB0B4\uC6A9)\uC744 \uBA3C\uC800 \uACE0\uB974\uC138\uC694. \uCE60\uD55C \uCE78\uC744 \uB2E4\uC2DC \uACE0\uB974\uACE0 \uB204\uB974\uBA74 \uC9C0\uC6CC\uC9D1\uB2C8\uB2E4.")
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
    ElseIf Left$(t, 1) = U("\u203B") Then
        a = a + (SkipWs(t, 2) - 1)
    ElseIf FirstIn(t, U("*\uFF0A")) Then
        p = 1
        Do While FirstIn(Mid$(t, p), U("*\uFF0A"))
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
    ans = Ask(U("\uC790\uB9AC\uB97C \uBC14\uAFC0 \uB450 \uCE78\uC758 \uBC88\uD638\uB97C \uC4F0\uC138\uC694 (\uC608: 1 4)") & vbLf & seatList, "")
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
    Info U("\uBB38\uC81C\uAC00 \uC0DD\uACBC\uC2B5\uB2C8\uB2E4: ") & Err.Description & vbLf & "(" & gStep & ")"
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
    p1 = Ask(U("\uC5F4 \uB54C \uBB3C\uC5B4\uBCFC \uBB38\uC11C \uC554\uD638\uB97C \uC4F0\uC138\uC694. (\uBE44\uC6CC \uB450\uBA74 \uC554\uD638\uB97C \uC5C6\uC571\uB2C8\uB2E4)"), "")
    If p1 <> "" Then
        p2 = Ask(U("\uD655\uC778\uC744 \uC704\uD574 \uD55C \uBC88 \uB354 \uC4F0\uC138\uC694."), "")
        If p1 <> p2 Then
            Info U("\uB450 \uBC88 \uC4F4 \uC554\uD638\uAC00 \uC11C\uB85C \uB2E4\uB985\uB2C8\uB2E4. \uB2E4\uC2DC \uD574 \uC8FC\uC138\uC694.")
            Exit Sub
        End If
    End If
    On Error Resume Next
    ThisWorkbook.Password = p1
    If Err.Number = 0 Then ThisWorkbook.Save
    If Err.Number <> 0 Then
        Info U("\uB9E4\uD06C\uB85C\uB85C \uBB38\uC11C \uC554\uD638\uB97C \uC815\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.") & vbLf & U("\uD55C\uC140 \uBA54\uB274\uC758 [\uBB38\uC11C \uC554\uD638](\uBCF4\uC548\u00B7\uAC80\uD1A0 \uBA54\uB274\uB098 \uD30C\uC77C \u2192 \uB2E4\uB978 \uC774\uB984\uC73C\uB85C \uC800\uC7A5\uC758 \uB3C4\uAD6C)\uC5D0\uC11C \uC815\uD558\uC138\uC694.")
    ElseIf p1 = "" Then
        Info U("\uBB38\uC11C \uC554\uD638\uB97C \uC5C6\uC574\uC2B5\uB2C8\uB2E4.")
    Else
        Info U("\uBB38\uC11C \uC554\uD638\uB97C \uC815\uD558\uACE0 \uC800\uC7A5\uD588\uC2B5\uB2C8\uB2E4. \uB2E4\uC74C\uC5D0 \uC5F4 \uB54C\uBD80\uD130 \uC554\uD638\uB97C \uBB3B\uC2B5\uB2C8\uB2E4.") & vbLf & U("\uC554\uD638\uB97C \uC78A\uC73C\uBA74 \uC5F4 \uC218 \uC5C6\uC73C\uB2C8 \uB530\uB85C \uC801\uC5B4 \uB450\uC138\uC694.")
    End If
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "^+h", "HighlightSel"
    Application.OnKey "^+H", "HighlightSel"
End Sub
