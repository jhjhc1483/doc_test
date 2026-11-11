Attribute VB_Name = "HcDailyIO"
Option Explicit
' ===========================================================================
' HcDailyIO - 엑셀 불러오기, 붙여넣은 표, 되돌리기, 새 문서, 부서 추가·빼기·순서,
'             형광펜, 공동편집 양식, PDF, 문서 암호
' ===========================================================================

Private Const SEP3 As String = "|~|"
Private Const NCOL As Long = 6

' ---- import entries (one per department and date) -------------------------
Private eN As Long
Private eName() As String, eKnown() As Boolean, eDate() As String, eSource() As String
Private eRows() As Object, eTarget() As Long, eOn() As Boolean, eNote() As String, eTitleKey() As String

Private Sub EntriesReset()
    eN = 0
    ReDim eName(1 To 64): ReDim eKnown(1 To 64): ReDim eDate(1 To 64): ReDim eSource(1 To 64)
    ReDim eRows(1 To 64): ReDim eTarget(1 To 64): ReDim eOn(1 To 64): ReDim eNote(1 To 64): ReDim eTitleKey(1 To 64)
End Sub

' entry for (name, date) from this source, created when missing
Private Function EntryFor(ByVal nm As String, ByVal known As Boolean, ByVal dt As String, ByVal source As String) As Long
    Dim k As String, i As Long
    k = source & "|" & NormName(nm) & "|" & dt
    For i = 1 To eN
        If eTitleKey(i) = k Then
            If known Then eKnown(i) = True
            EntryFor = i
            Exit Function
        End If
    Next
    eN = eN + 1
    If eN > UBound(eName) Then
        ReDim Preserve eName(1 To eN * 2): ReDim Preserve eKnown(1 To eN * 2): ReDim Preserve eDate(1 To eN * 2)
        ReDim Preserve eSource(1 To eN * 2): ReDim Preserve eRows(1 To eN * 2): ReDim Preserve eTarget(1 To eN * 2)
        ReDim Preserve eOn(1 To eN * 2): ReDim Preserve eNote(1 To eN * 2): ReDim Preserve eTitleKey(1 To eN * 2)
    End If
    eName(eN) = nm: eKnown(eN) = known: eDate(eN) = dt: eSource(eN) = source: eTitleKey(eN) = k
    Set eRows(eN) = New Collection
    EntryFor = eN
End Function

Private Sub EntryRow(ByVal i As Long, ByVal joined As String)
    eRows(i).Add joined
End Sub

' fixed-format row: 시간 내용 장소 비고 설명 링크
Private Function FixedRow(ByVal tm As String, ByVal content As String, ByVal place As String, ByVal extra As String, _
        ByVal detail As String, ByVal link As String) As String
    FixedRow = tm & SEP3 & content & SEP3 & place & SEP3 & extra & SEP3 & detail & SEP3 & link
End Function

Private Function MdRow(ByVal ln As String) As String
    MdRow = FixedRow("", ln, "", "", "", "")
End Function

' ---- column roles (port of headerRole / colsFromHeader) ---------------------
Private Function RoleIn(ByVal k As String, ByVal list As String) As Boolean
    RoleIn = (InStr("|" & list & "|", "|" & k & "|") > 0)
End Function

Private Function RoleStarts(ByVal k As String, ByVal list As String) As Boolean
    Dim p As Variant
    For Each p In Split(list, "|")
        If StartsWith(k, CStr(p)) Then
            RoleStarts = True
            Exit Function
        End If
    Next
End Function

Public Function DHeaderRole(ByVal h As String) As String
    Dim k As String, c As Long, baseK As String
    k = LCase$(RemoveChars(h, "()（）[]"))
    If k = "" Or Len(k) > 24 Then Exit Function
    If RoleIn(k, "부서|부서명|소속|과|팀|실|국|과명|팀명|부서/팀|담당부서|주관부서|기관|기관명|dept|department|division|team") Then
        DHeaderRole = "dept"
    ElseIf RoleIn(k, "날짜|일자|일|월일|date|day") Then
        DHeaderRole = "date"
    ElseIf RoleIn(k, "시간|시각|일시|시작|시작시간|시간대|time") Then
        DHeaderRole = "time"
    ElseIf RoleIn(k, "장소|위치|회의실|장소명|place|location|venue|room") Then
        DHeaderRole = "place"
    ElseIf RoleStarts(k, "설명|상세|상세내용|세부내용|세부사항|메모|팝업|desc|description|detail|memo") Then
        DHeaderRole = "detail"
    ElseIf RoleStarts(k, "링크|관련링크|주소|url|link|href|자료링크|게시판링크|바로가기") Then
        DHeaderRole = "link"
    Else
        baseK = k
        If StartsWith(baseK, "주요") Then baseK = Mid$(baseK, 3)
        If RoleIn(baseK, "내용|예정사항|예정내용|일정|일정내용|세부일정|업무|업무내용|행사|행사명|행사내용|사항|제목|건명|항목|회의명|content|title|subject|task|event|schedule|item") Then
            DHeaderRole = "content"
        ElseIf RoleStarts(k, "비고|참석|참석자|참석대상|참석범위|대상|담당|담당자|주관|주재|참고|준비|준비사항|협조|협조사항|연락처|문의|인원|note|notes|remark|remarks|attendee|attendees") Then
            DHeaderRole = "extra"
        End If
    End If
End Function

' first header row (port of findHeader); 0 when none
Private Function FindHeaderRow(v As Variant, ByVal r0 As Long, ByVal nR As Long, ByVal nC As Long) As Long
    Dim r As Long, c As Long, seen As Long, t As String, ok As Boolean, rec As Long, hasMain As Boolean, hasC As Boolean, cells As Long, ro As String
    For r = r0 To nR
        If seen >= 12 Then Exit For
        cells = 0: ok = True: rec = 0: hasMain = False: hasC = False
        For c = 1 To nC
            t = CellText(v(r, c))
            If t <> "" Then
                cells = cells + 1
                If Len(t) > 24 Or InStr(t, vbLf) > 0 Then ok = False
                ro = DHeaderRole(t)
                If ro <> "" Then rec = rec + 1
                If ro = "content" Then hasC = True
                If ro = "content" Or ro = "time" Then hasMain = True
            End If
        Next
        If cells > 0 Then
            seen = seen + 1
            If ok And hasMain And (rec >= 2 Or (seen = 1 And hasC And cells >= 2)) Then
                FindHeaderRow = r
                Exit Function
            End If
        End If
    Next
End Function

Private Function RowAnyV(v As Variant, ByVal r As Long, ByVal nC As Long) As Long
    Dim c As Long, n As Long
    For c = 1 To nC
        If CellText(v(r, c)) <> "" Then n = n + 1
    Next
    RowAnyV = n
End Function

Private Function LooksHead(v As Variant, ByVal r As Long, ByVal nR As Long, ByVal nC As Long) As Boolean
    Dim c As Long, n As Long, t As String
    If r > nR Then Exit Function
    If CellText(v(r, 1)) = "" Then Exit Function
    If DTimeLen(CellText(v(r, 1))) > 0 Then Exit Function
    For c = 1 To nC
        t = CellText(v(r, c))
        If t <> "" Then
            n = n + 1
            If Len(t) > 30 Or InStr(t, vbLf) > 0 Then Exit Function
        End If
    Next
    LooksHead = (n >= 2)
End Function

' a table after a blank row: copied as it is (title row, header, body). returns the row after it, 0 when none
Private Function CopyTableBlock(v As Variant, ByVal r As Long, ByVal nR As Long, ByVal nC As Long, ByVal entry As Long) As Long
    Dim h As Long, e As Long, rr As Long, c As Long, parts() As String
    h = r
    If RowAnyV(v, r, nC) = 1 And CellText(v(r, 1)) <> "" And LooksHead(v, r + 1, nR, nC) Then h = r + 1
    If Not LooksHead(v, h, nR, nC) Then Exit Function
    e = h + 1
    Do While e <= nR
        If RowAnyV(v, e, nC) = 0 Then Exit Do
        e = e + 1
    Loop
    If e - h < 2 Then Exit Function
    EntryRow entry, FixedRow("", "", "", "", "", "")
    For rr = r To e - 1
        ReDim parts(0 To MaxD(nC - 1, NCOL - 1))
        For c = 1 To nC
            parts(c - 1) = CellText(v(rr, c))
        Next
        EntryRow entry, Join(parts, SEP3)
    Next
    CopyTableBlock = e
End Function

' "# 부서 @2026-11-11", "[부서]", or just a department name -> name / date
Private Function SectionHead(ByVal l As String, ByVal allowPlain As Boolean, ByRef nm As String, ByRef dt As String) As Boolean
    Dim t As String, p As Long, d As Long, rest As String, closeCh As String
    t = TrimX(l)
    nm = "": dt = ""
    If Left$(t, 1) = "#" And IsWs(Mid$(t, 2, 1)) Then
        rest = TrimX(Mid$(t, 2))
        Do While Right$(rest, 1) = "#"
            rest = RTrimX(Left$(rest, Len(rest) - 1))
        Loop
        p = InStrRev(rest, "@")
        If p > 0 Then
            If ParseLooseDate(TrimX(Mid$(rest, p + 1)), CLng(Left$(gDate, 4))) <> "" And IsDigitAt(TrimX(Mid$(rest, p + 1)), 1) Then
                dt = ParseLooseDate(TrimX(Mid$(rest, p + 1)), CLng(Left$(gDate, 4)))
                rest = RTrimX(Left$(rest, p - 1))
            End If
        End If
        d = FindDept(rest)
        If d > 0 Then
            nm = dName(d)
        Else
            nm = rest
        End If
        SectionHead = (nm <> "")
        Exit Function
    End If
    If FirstIn(t, "[【<〈《「『") Then
        closeCh = Mid$("]】>〉》」』", InStr("[【<〈《「『", Left$(t, 1)), 1)
        If Right$(t, 1) = closeCh Or FirstIn(Right$(t, 1), "]】>〉》」』") Then
            rest = TrimX(Mid$(t, 2, Len(t) - 2))
            d = FindDept(rest)
            If d > 0 Then
                nm = dName(d)
                SectionHead = True
                Exit Function
            End If
        End If
    End If
    If allowPlain And Len(t) <= 20 Then
        d = FindDept(t)
        If d > 0 Then
            nm = dName(d)
            SectionHead = True
        End If
    End If
End Function

' one sheet -> entries (port of interpretGrid; rows are kept in the department-sheet format)
Private Sub InterpretSheet(ByVal sheetName As String, v As Variant, ByVal r0 As Long, ByVal sheetIsDept As Boolean, ByVal source As String)
    Dim nR As Long, nC As Long, r As Long, c As Long, sheetDept As Long, baseName As String, nameOk As Boolean, baseKnown As Boolean
    Dim hr As Long, colDept As Long, colDate As Long, colTime As Long, colPlace As Long, colLink As Long, colDetail As Long
    Dim cont() As Long, nCont As Long, ex() As Long, exLab() As String, nEx As Long, ro As String, t As String
    Dim usedCols As Long, nonEmpty As Long, filledA As Long, timeA As Long, entry As Long, start As Long, afterBlank As Boolean
    Dim curHead As String, curHeadDate As String, hasHead As Boolean, lastDept As String, lastDate As String, dept As String, dt As String
    Dim tm As String, content As String, place As String, extra As String, detail As String, link As String, e As Long, nm As String
    Dim y0 As Long, lines As New Collection, l As Variant, md As String, k As Long, hd As String, hdate As String, cell As String
    Dim parts As Variant
    If Not IsArray(v) Then Exit Sub
    nR = UBound(v, 1): nC = UBound(v, 2)
    y0 = CLng(Left$(gDate, 4))
    sheetDept = FindDept(sheetName)
    If sheetDept = 0 And Len(NormName(sheetName)) >= 2 Then sheetDept = FindDeptInText(sheetName)
    If sheetDept > 0 Then
        baseName = dName(sheetDept)
    Else
        baseName = sheetName
    End If
    t = LCase$(RemoveChars(sheetName, ""))
    nameOk = sheetIsDept And sheetName <> "" And Not ((StartsWith(t, "sheet") And IsNumeric("0" & Mid$(t, 6))) Or (StartsWith(t, "시트") And IsNumeric("0" & Mid$(t, 3))))
    ReDim cont(1 To nC + 1): ReDim ex(1 To nC + 1): ReDim exLab(1 To nC + 1)
    For r = r0 To nR
        If RowAnyV(v, r, nC) > 0 Then
            nonEmpty = nonEmpty + 1
            If CellText(v(r, 1)) <> "" Then
                filledA = filledA + 1
                If IsTimeLikeD(v(r, 1)) Then timeA = timeA + 1
            End If
        End If
    Next
    If nonEmpty = 0 Then
        EntryFor baseName, (sheetDept > 0 Or nameOk), "", source
        Exit Sub
    End If
    hr = FindHeaderRow(v, r0, nR, nC)
    colDept = 0: colDate = 0: colTime = 0: colPlace = 0: colLink = 0: colDetail = 0
    If hr > 0 Then
        For c = 1 To nC
            t = CellText(v(hr, c))
            If t <> "" Then
                ro = DHeaderRole(t)
                If ro = "content" Then
                    nCont = nCont + 1: cont(nCont) = c
                ElseIf ro = "dept" And colDept = 0 Then
                    colDept = c
                ElseIf ro = "date" And colDate = 0 Then
                    colDate = c
                ElseIf ro = "time" And colTime = 0 Then
                    colTime = c
                ElseIf ro = "place" And colPlace = 0 Then
                    colPlace = c
                ElseIf ro = "link" And colLink = 0 Then
                    colLink = c
                ElseIf ro = "detail" And colDetail = 0 Then
                    colDetail = c
                Else
                    nEx = nEx + 1: ex(nEx) = c: exLab(nEx) = Replace(t, vbLf, " ")
                End If
            End If
        Next
        If nCont = 0 And nEx > 0 Then
            nCont = 1: cont(1) = ex(1)
            For k = 1 To nEx - 1
                ex(k) = ex(k + 1): exLab(k) = exLab(k + 1)
            Next
            nEx = nEx - 1
        End If
        If nCont = 0 And colTime = 0 Then hr = 0
    End If
    baseKnown = (sheetDept > 0) Or (hr > 0 And nameOk)
    start = hr + 1
    If hr = 0 Then
        For c = 1 To nC
            For r = r0 To nR
                If CellText(v(r, c)) <> "" Then
                    usedCols = usedCols + 1
                    Exit For
                End If
            Next
        Next
        If usedCols <= 1 Then
            ' text only, one line per row: markdown (may hold "# 부서" sections)
            entry = 0
            For r = r0 To nR
                cell = ""
                For c = 1 To nC
                    If CellText(v(r, c)) <> "" Then cell = CellTextKeepLead(v(r, c))
                Next
                For Each l In Split(cell, vbLf)
                    If SectionHead(CStr(l), True, hd, hdate) Then
                        entry = EntryFor(hd, True, hdate, source)
                    Else
                        If entry = 0 Then entry = EntryFor(baseName, sheetDept > 0, "", source)
                        EntryRow entry, MdRow(RTrimX(CStr(l)))
                    End If
                Next
            Next
            Exit Sub
        End If
        If timeA >= MaxD(1, filledA * 0.6) Then
            colTime = 1: nCont = 1: cont(1) = 2: colPlace = 3
            nEx = 0
            For c = 4 To MinD(nC, 12)
                nEx = nEx + 1: ex(nEx) = c: exLab(nEx) = ""
            Next
            start = r0
        Else
            ' unknown table: kept as a markdown table
            entry = EntryFor(baseName, sheetDept > 0, "", source)
            k = 0
            For r = r0 To nR
                If RowAnyV(v, r, nC) > 0 Then
                    md = "|"
                    For c = 1 To nC
                        md = md & " " & Replace(Replace(CellText(v(r, c)), "|", "/"), vbLf, " ") & " |"
                    Next
                    EntryRow entry, MdRow(md)
                    k = k + 1
                    If k = 1 Then
                        md = "|"
                        For c = 1 To nC
                            md = md & "---|"
                        Next
                        EntryRow entry, MdRow(md)
                    End If
                End If
            Next
            Exit Sub
        End If
    End If
    r = start
    Do While r <= nR
        If RowAnyV(v, r, nC) = 0 Then
            afterBlank = True
            r = r + 1
            GoTo nextRow
        End If
        If afterBlank Then
            afterBlank = False
            If colDept > 0 And lastDept <> "" Then
                entry = EntryFor(lastDept, True, lastDate, source)
            ElseIf hasHead Then
                entry = EntryFor(curHead, True, curHeadDate, source)
            Else
                entry = EntryFor(baseName, baseKnown, lastDate, source)
            End If
            e = CopyTableBlock(v, r, nR, nC, entry)
            If e > 0 Then
                r = e
                GoTo nextRow
            End If
        End If
        If RowAnyV(v, r, nC) = 1 And colDept = 0 Then
            For c = 1 To nC
                If CellText(v(r, c)) <> "" Then cell = CellText(v(r, c))
            Next
            If SectionHead(cell, True, hd, hdate) Then
                curHead = hd: curHeadDate = hdate: hasHead = True
                r = r + 1
                GoTo nextRow
            End If
        End If
        dept = ""
        If colDept > 0 Then
            parts = Split(CellText(v(r, colDept)) & vbLf, vbLf)
            dept = TrimX(CStr(parts(0)))
            If dept = "" Then dept = lastDept
            lastDept = dept
        End If
        dt = ""
        If colDate > 0 Then
            dt = DateValueIso(v(r, colDate), y0)
            If dt = "" Then dt = lastDate
            lastDate = dt
        ElseIf colTime > 0 And VarType(v(r, colTime)) = vbDate Then
            If CDbl(v(r, colTime)) >= 1 Then dt = IsoOf(CDate(v(r, colTime)))
        ElseIf hasHead Then
            dt = curHeadDate
        End If
        ' the row in the department-sheet format
        tm = ""
        If colTime > 0 Then tm = DNormTime(CellText(v(r, colTime), True))
        content = ""
        For k = 1 To nCont
            If CellText(v(r, cont(k))) <> "" Then content = content & IIf(content = "", "", vbLf) & CellTextKeepLead(v(r, cont(k)))
        Next
        place = ""
        If colPlace > 0 Then place = Replace(CellText(v(r, colPlace)), vbLf, " ")
        extra = ""
        For k = 1 To nEx
            t = Replace(CellText(v(r, ex(k))), vbLf, ", ")
            If t <> "" Then
                If exLab(k) = "" Or exLab(k) = "비고" Or exLab(k) = "참고" Or exLab(k) = "기타" Or StartsWith(t, exLab(k)) Then
                    extra = extra & IIf(extra = "", "", vbLf) & t
                Else
                    extra = extra & IIf(extra = "", "", vbLf) & exLab(k) & ": " & t
                End If
            End If
        Next
        detail = ""
        If colDetail > 0 Then detail = CellText(v(r, colDetail))
        link = ""
        If colLink > 0 Then link = CellText(v(r, colLink))
        If tm <> "" Or content <> "" Or place <> "" Or extra <> "" Or detail <> "" Or link <> "" Then
            If dept <> "" Then
                k = FindDept(dept)
                If k > 0 Then
                    nm = dName(k)
                Else
                    nm = dept
                End If
                entry = EntryFor(nm, True, dt, source)
            ElseIf hasHead Then
                entry = EntryFor(curHead, True, dt, source)
            Else
                entry = EntryFor(baseName, baseKnown, dt, source)
            End If
            EntryRow entry, FixedRow(tm, content, place, extra, detail, link)
        End If
        r = r + 1
nextRow:
    Loop
    If eN = 0 Then EntryFor baseName, baseKnown, "", source
End Sub

Private Function IsTimeLikeD(ByVal x As Variant) As Boolean
    Dim t As String
    If VarType(x) = vbDate Then
        IsTimeLikeD = (CDbl(x) < 1)
        Exit Function
    End If
    t = CellText(x)
    If DTimeLen(t) > 0 Then
        IsTimeLikeD = True
    ElseIf IsNumeric(t) Then
        IsTimeLikeD = (Val(t) >= 0 And Val(t) < 1 And InStr(t, ".") > 0)
    End If
End Function

Private Function EntryLines(ByVal i As Long) As Long
    Dim rw As Variant, n As Long, p As Variant, c As Long
    For Each rw In eRows(i)
        p = Split(rw, SEP3)
        For c = 0 To UBound(p)
            If TrimX(CStr(p(c))) <> "" Then
                n = n + 1
                Exit For
            End If
        Next
    Next
    EntryLines = n
End Function

' ---- buttons ------------------------------------------------------------------
Public Sub ImportExcel()
    Dim f As Variant
    On Error Resume Next
    f = Application.GetOpenFilename("엑셀 파일 (*.xlsx;*.xlsm;*.xls;*.csv;*.xml),*.xlsx;*.xlsm;*.xls;*.csv;*.xml", 1, "두레이에서 내려받은 엑셀 파일을 고르세요", , True)
    If Err.Number <> 0 Then
        Err.Clear
        f = Application.GetOpenFilename("엑셀 파일 (*.xlsx;*.xlsm;*.xls;*.csv;*.xml),*.xlsx;*.xlsm;*.xls;*.csv;*.xml")
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
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Function ImportFiles(files As Variant) As String
    Dim p As Variant, wb As Object, ws As Object, v As Variant, names As String, fn As String, skipped As Long, src As String
    LoadSettings
    LoadDepts False
    EntriesReset
    For Each p In files
        If LCase$(CStr(p)) = LCase$(ThisWorkbook.FullName) Then
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
            Info fn & ": 파일을 열 수 없습니다. (예전 엑셀 .xls는 .xlsx로 저장해서 올려 주세요)"
        Else
            If names <> "" Then names = names & ", "
            names = names & fn
            For Each ws In wb.Worksheets
                gStep = "import: sheet " & ws.Name
                If ws.Visible <> XL_SHEET_VISIBLE Or SkipSheetD(ws.Name) Then
                    skipped = skipped + 1
                Else
                    v = SheetValuesD(ws)
                    If wb.Worksheets.Count > 1 Then
                        src = fn & " › " & ws.Name
                    Else
                        src = fn
                    End If
                    If LCase$(Right$(fn, 4)) = ".csv" Then
                        If FindDeptInText(fn) > 0 Then
                            InterpretSheet dName(FindDeptInText(fn)), v, 1, False, src
                        Else
                            InterpretSheet Left$(fn, Len(fn) - 4), v, 1, False, src
                        End If
                    Else
                        InterpretSheet ws.Name, v, 1, True, src
                    End If
                End If
            Next
            wb.Close False
        End If
    Next
    If names = "" Then Exit Function
    ImportFiles = ReviewImport(names, skipped)
End Function

Private Function SkipSheetD(ByVal nm As String) As Boolean
    Dim k As String, p As Variant
    k = LCase$(RemoveChars(nm, ""))
    For Each p In Array("작성방법", "작성요령", "작성예시", "안내", "도움말", "사용법", "예시", "목록", "설정", "readme", "guide", "help")
        If StartsWith(k, CStr(p)) Then
            SkipSheetD = True
            Exit Function
        End If
    Next
End Function

Private Function SheetValuesD(ByVal ws As Object) As Variant
    Dim ur As Object, v As Variant, a(1 To 1, 1 To 1) As Variant
    Set ur = ws.UsedRange
    Set ur = ws.Range(ws.Cells(1, 1), ws.Cells(ur.Row + ur.Rows.Count - 1, ur.Column + ur.Columns.Count - 1))
    v = ur.Value
    If Not IsArray(v) Then
        a(1, 1) = v
        v = a
    End If
    SheetValuesD = v
End Function

Public Sub ImportPasted()
    Dim ws As Object, v As Variant, lastR As Long, lastC As Long
    On Error GoTo fail
    Set ws = SheetByName(SH_PASTE)
    If ws Is Nothing Then Exit Sub
    UsedLast ws, lastR, lastC
    If lastR < 3 Then
        Info "붙여넣기 시트의 3번째 줄부터 엑셀에서 복사한 내용을 붙여넣은 뒤 다시 누르세요."
        ws.Activate
        Exit Sub
    End If
    LoadSettings
    LoadDepts False
    EntriesReset
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, MaxD(lastC, 2))).Value
    InterpretSheet "", v, 3, False, "붙여넣은 내용"
    If ReviewImport("붙여넣은 내용", 0) = "ok" Then ws.Range(ws.Cells(3, 1), ws.Cells(lastR, MaxD(lastC, 2))).ClearContents
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

' choose date, departments; confirm; apply (port of the import review). returns "ok" when applied
Private Function ReviewImport(ByVal srcLabel As String, ByVal skipped As Long) As String
    Dim i As Long, j As Long, ds() As String, cnt() As Long, nD As Long, k As Long, mx As Long, ref As String, top As String, nTop As Long
    Dim useRef As String, multi As Boolean, changeDate As Boolean, msg As String, ans As String, deptList As String, nOn As Long
    Dim lines As String, offs As String, nNew As Long, ws As Object, nm As String, sc1 As Long, sc2 As Long, w As Long
    If eN = 0 Then
        Info "가져올 내용이 없습니다."
        Exit Function
    End If
    ' dates
    ReDim ds(1 To eN): ReDim cnt(1 To eN)
    For i = 1 To eN
        If eDate(i) <> "" Then
            k = FindKey(ds, nD, eDate(i))
            If k = 0 Then
                nD = nD + 1: ds(nD) = eDate(i): cnt(nD) = 1
            Else
                cnt(k) = cnt(k) + 1
            End If
        End If
    Next
    For k = 1 To nD
        If cnt(k) > mx Then mx = cnt(k)
    Next
    ref = ""
    If nD > 0 Then
        k = FindKey(ds, nD, gDate)
        j = 0
        If k > 0 Then j = cnt(k)
        If j * 2 >= mx And j > 0 Then
            ref = gDate
        Else
            ' most common date; on a tie the first one from today on, else the latest
            top = ""
            For k = 1 To nD
                If cnt(k) = mx Then
                    If ds(k) >= IsoOf(Date) Then
                        If top = "" Or top < IsoOf(Date) Or ds(k) < top Then top = ds(k)
                    ElseIf top = "" Or (top < IsoOf(Date) And ds(k) > top) Then
                        top = ds(k)
                    End If
                End If
            Next
            ref = top
        End If
    End If
    useRef = gDate
    If ref <> "" And ref <> gDate Then
        If Confirm("가져온 내용 대부분이 " & FmtDateK(ref) & " 날짜입니다." & vbLf & "문서 날짜도 " & FmtDateK(ref) & "로 바꿀까요? (지금은 " & FmtDateK(gDate) & ")") Then
            useRef = ref
            changeDate = True
        End If
    End If
    multi = (nD > 1)
    ' targets and defaults (port of prepareRows)
    For i = 1 To eN
        eNote(i) = ""
        eTarget(i) = FindDept(eName(i))
        eOn(i) = True
        If multi And eDate(i) <> "" And eDate(i) <> useRef Then
            eOn(i) = False
            eNote(i) = "다른 날짜"
        ElseIf eTarget(i) = 0 And Not eKnown(i) Then
            eOn(i) = False
            eNote(i) = "부서를 찾지 못함"
        ElseIf eDate(i) <> "" And eDate(i) < useRef Then
            eOn(i) = False
            eNote(i) = "지난 날짜(" & FmtDateK(eDate(i)) & ")"
        ElseIf EntryLines(i) = 0 Then
            eOn(i) = False
            eNote(i) = "내용 없음"
        End If
    Next
    ' the same department twice: keep the one dated for the document, then the first
    For i = 1 To eN
        If eOn(i) Then
            For j = i + 1 To eN
                If eOn(j) Then
                    If (eTarget(i) > 0 And eTarget(i) = eTarget(j)) Or (eTarget(i) = 0 And eTarget(j) = 0 And NormName(eName(i)) = NormName(eName(j))) Then
                        sc1 = IIf(eDate(i) = useRef, 2, IIf(eDate(i) = "", 1, 0))
                        sc2 = IIf(eDate(j) = useRef, 2, IIf(eDate(j) = "", 1, 0))
                        If sc2 > sc1 Then
                            eOn(i) = False
                            eNote(i) = "같은 부서의 더 알맞은 내용이 있음"
                            Exit For
                        Else
                            eOn(j) = False
                            eNote(j) = "같은 부서의 더 알맞은 내용이 있음"
                        End If
                    End If
                End If
            Next
        End If
    Next
    ' departments not found: ask
    For k = 1 To dN
        deptList = deptList & vbLf & k & ". " & dName(k)
    Next
    For i = 1 To eN
        If eNote(i) = "부서를 찾지 못함" And EntryLines(i) > 0 Then
            ans = Ask("'" & eSource(i) & "' (" & EntryLines(i) & "줄) — 부서를 찾지 못했습니다." & vbLf & _
                "넣을 부서 번호를 쓰세요. 0 = 새 부서 '" & eName(i) & "', 비워 두면 넣지 않습니다." & vbLf & deptList, "")
            If TrimX(ans) <> "" And IsNumeric(ans) Then
                w = CLng(ans)
                If w = 0 Then
                    eTarget(i) = 0
                    eOn(i) = True
                    eNote(i) = ""
                ElseIf w >= 1 And w <= dN Then
                    eTarget(i) = w
                    eOn(i) = True
                    eNote(i) = ""
                End If
            End If
        End If
    Next
    ' summary
    For i = 1 To eN
        If eOn(i) Then
            nOn = nOn + 1
            If eTarget(i) > 0 Then
                nm = dName(eTarget(i))
            Else
                nm = eName(i) & " (새 부서)"
                nNew = nNew + 1
            End If
            lines = lines & vbLf & "· " & nm & " ← " & eSource(i) & " (" & EntryLines(i) & "줄)"
        ElseIf eNote(i) <> "" Then
            offs = offs & vbLf & "· " & eSource(i) & IIf(eName(i) <> "", " [" & eName(i) & "]", "") & ": " & eNote(i)
        End If
    Next
    If nOn = 0 Then
        Info "가져올 내용이 없습니다." & IIf(offs <> "", vbLf & offs, "")
        Exit Function
    End If
    msg = srcLabel & vbLf & nOn & "개 부서를 가져옵니다. 고른 부서의 기존 내용은 엑셀 내용으로 바뀝니다." & vbLf & lines
    If offs <> "" Then msg = msg & vbLf & vbLf & "뺀 것:" & offs
    If skipped > 0 Then msg = msg & vbLf & "안내 시트·숨긴 시트 " & skipped & "개는 뺐습니다."
    msg = msg & vbLf & vbLf & "가져올까요?"
    If Len(msg) > 1000 Then msg = Left$(msg, 990) & "…"
    If Not Confirm(msg) Then Exit Function
    ' apply
    gStep = "import: backup"
    SaveBackup "엑셀 불러오기"
    BeginWork
    gStep = "import: write"
    For i = 1 To eN
        If eOn(i) Then
            If eTarget(i) > 0 Then
                Set ws = dWs(eTarget(i))
            Else
                Set ws = NewDeptSheet(eName(i))
                LoadDepts False
                eTarget(i) = FindDept(eName(i))
            End If
            WriteDeptRows ws, eRows(i)
            ws.Cells(1, 6).NumberFormat = "@"
            ws.Cells(1, 6).Value = "엑셀 " & Format$(Now, "hh:mm") & " · " & eSource(i)
        End If
    Next
    If changeDate Then NSet "d_date", IsoToDate(useRef)
    DoBuild
    EndWork
    Info nOn & "개 부서의 내용을 가져왔습니다." & IIf(nNew > 0, " (새 부서 " & nNew & "곳 추가)", "") & vbLf & "잘못 가져왔으면 입력 시트의 [되돌리기]를 누르세요."
    ReviewImport = "ok"
End Function

Private Sub ClearDeptRows(ByVal ws As Object)
    Dim lastR As Long, lastC As Long
    UsedLast ws, lastR, lastC
    If lastR < 3 Then lastR = 3
    If lastC < NCOL Then lastC = NCOL
    ws.Range(ws.Cells(3, 1), ws.Cells(lastR, lastC)).ClearContents
End Sub

Private Sub WriteDeptRows(ByVal ws As Object, ByVal rows As Collection)
    Dim rw As Variant, p As Variant, r As Long, c As Long
    ClearDeptRows ws
    r = 3
    For Each rw In rows
        p = Split(rw, SEP3)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, MaxD(UBound(p) + 1, NCOL))).NumberFormat = "@"
        For c = 0 To UBound(p)
            If CStr(p(c)) <> "" Then ws.Cells(r, c + 1).Value = CStr(p(c))
        Next
        r = r + 1
    Next
End Sub

' a new department sheet after the last one
Public Function NewDeptSheet(ByVal nm As String) As Object
    Dim c As Collection, after As Object, ws As Object
    Set c = DeptSheetList()
    If c.Count > 0 Then
        Set after = c(c.Count)
    Else
        Set after = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If
    Set ws = ThisWorkbook.Worksheets.Add(, after)
    SetupDeptSheet ws, nm, Empty
    ws.Name = UniqueSheetName(nm, ws)
    Set NewDeptSheet = ws
End Function

' ---- backup / undo -----------------------------------------------------------
Private Sub SaveBackup(ByVal reason As String)
    Dim b As Object, k As Long, ws As Object, lastR As Long, lastC As Long, v As Variant, c As Collection, col As Long
    Set b = ScratchSheet(SH_BAK)
    b.Cells.Clear
    b.Cells(1, 1).Value = reason
    b.Cells(1, 2).Value = Now
    b.Cells(2, 1).Value = NVal("d_date")
    Set c = DeptSheetList()
    b.Cells(2, 2).Value = c.Count
    For k = 1 To c.Count
        Set ws = c(k)
        col = (k - 1) * 13 + 1
        UsedLast ws, lastR, lastC
        If lastR < 3 Then lastR = 3
        If lastC > 12 Then lastC = 12
        If lastC < NCOL Then lastC = NCOL
        b.Cells(3, col).NumberFormat = "@"
        b.Cells(3, col).Value = ws.Name
        b.Cells(3, col + 1).Value = lastR
        v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
        b.Range(b.Cells(4, col), b.Cells(4 + lastR - 1, col + lastC - 1)).NumberFormat = "@"
        b.Range(b.Cells(4, col), b.Cells(4 + lastR - 1, col + lastC - 1)).Value = v
    Next
End Sub

Public Sub UndoLast()
    Dim b As Object, n As Long, k As Long, nm As String, ws As Object, lastR As Long, v As Variant, c As Collection, keep As String, col As Long
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
    n = CLng(b.Cells(2, 2).Value)
    For k = 1 To n
        keep = keep & "|" & LCase$(CStr(b.Cells(3, (k - 1) * 13 + 1).Value)) & "|"
    Next
    ' departments added after the backup are removed
    Application.DisplayAlerts = False
    Set c = DeptSheetList()
    For Each ws In c
        If InStr(keep, "|" & LCase$(ws.Name) & "|") = 0 Then ws.Delete
    Next
    Application.DisplayAlerts = True
    For k = 1 To n
        col = (k - 1) * 13 + 1
        nm = CStr(b.Cells(3, col).Value)
        lastR = CLng(b.Cells(3, col + 1).Value)
        Set ws = SheetByName(nm)
        If ws Is Nothing Then
            Set ws = NewDeptSheet(nm)
            ws.Name = UniqueSheetName(nm, ws)
        End If
        ClearDeptRows ws
        v = b.Range(b.Cells(4, col), b.Cells(4 + lastR - 1, col + 11)).Value
        ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 12)).NumberFormat = "@"
        ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 12)).Value = v
    Next
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

' ---- new document / departments ----------------------------------------------
Public Sub NewDoc()
    Dim k As Long
    On Error GoTo fail
    If Not Confirm("모든 부서의 내용을 비웁니다." & vbLf & "(부서 목록·서식은 그대로 두고, 날짜는 오늘로 바뀝니다)") Then Exit Sub
    BeginWork
    LoadSettings
    LoadDepts False
    SaveBackup "새 문서"
    For k = 1 To dN
        ClearDeptRows dWs(k)
        dWs(k).Cells(1, 6).ClearContents
    Next
    NSet "d_date", Date
    DoBuild
    EndWork
    Info "내용을 비웠습니다. 잘못 눌렀으면 [되돌리기]를 누르세요."
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub AddDept()
    Dim nm As String, ws As Object
    On Error GoTo fail
    LoadDepts False
    nm = TrimX(Ask("새 부서 이름을 쓰세요.", ""))
    If nm = "" Then Exit Sub
    If FindDept(nm) > 0 Then
        Info "'" & nm & "' 부서가 이미 있습니다."
        Exit Sub
    End If
    BeginWork
    Set ws = NewDeptSheet(nm)
    DoBuild
    EndWork
    ws.Activate
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Private Function DeptListText() As String
    Dim k As Long, s As String
    For k = 1 To dN
        s = s & vbLf & k & ". " & dName(k)
    Next
    DeptListText = s
End Function

Public Sub RemoveDept()
    Dim ans As String, k As Long
    On Error GoTo fail
    LoadDepts False
    ans = TrimX(Ask("뺄 부서의 번호를 쓰세요. (시트가 지워집니다. 잘못 지웠으면 [되돌리기])" & DeptListText(), ""))
    If ans = "" Or Not IsNumeric(ans) Then Exit Sub
    k = CLng(ans)
    If k < 1 Or k > dN Then Exit Sub
    If Not Confirm("'" & dName(k) & "' 부서를 뺄까요?") Then Exit Sub
    BeginWork
    SaveBackup "부서 빼기"
    Application.DisplayAlerts = False
    dWs(k).Delete
    Application.DisplayAlerts = True
    DoBuild
    EndWork
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub SwapDepts()
    Dim ans As String, parts As Variant, i As Long, j As Long
    On Error GoTo fail
    LoadDepts False
    ans = Ask("순서를 맞바꿀 두 부서의 번호를 쓰세요 (예: 1 4). 부서 시트 탭을 끌어 순서를 바꿔도 됩니다." & DeptListText(), "")
    ans = TrimX(Replace(Replace(ans, ",", " "), "-", " "))
    If ans = "" Then Exit Sub
    parts = Split(CollapseWs(ans), " ")
    If UBound(parts) < 1 Then Exit Sub
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Sub
    i = CLng(parts(0)): j = CLng(parts(1))
    If i < 1 Or i > dN Or j < 1 Or j > dN Or i = j Then Exit Sub
    BeginWork
    SwapDeptsIJ i, j
    DoBuild
    EndWork
    Exit Sub
fail:
    EndWork
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Sub SwapDeptsIJ(ByVal i As Long, ByVal j As Long)
    Dim a As Object, b As Object, ra As Long, rb As Long, ca As Long, cb As Long, n As Long, va As Variant, vb As Variant, na As String, nb As String
    Set a = dWs(i): Set b = dWs(j)
    UsedLast a, ra, ca
    UsedLast b, rb, cb
    n = MaxD(MaxD(ra, rb), 3)
    If ca < cb Then ca = cb
    If ca < NCOL Then ca = NCOL
    va = a.Range(a.Cells(1, 1), a.Cells(n, ca)).Value
    vb = b.Range(b.Cells(1, 1), b.Cells(n, ca)).Value
    a.Range(a.Cells(3, 1), a.Cells(n, ca)).ClearContents
    b.Range(b.Cells(3, 1), b.Cells(n, ca)).ClearContents
    a.Range(a.Cells(1, 1), a.Cells(n, ca)).Value = vb
    b.Range(b.Cells(1, 1), b.Cells(n, ca)).Value = va
    na = a.Name: nb = b.Name
    a.Name = "~swap~"
    b.Name = na
    a.Name = nb
End Sub

' ---- highlight ------------------------------------------------------------------
Public Sub HighlightSel()
    Dim c As Object, s As String, n As Long
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    On Error GoTo 0
    For Each c In Selection.Cells
        If Not IsEmpty(c.Value) And c.Column > 1 And c.Row > 2 Then
            s = CStr(c.Value)
            If InStr(s, "==") > 0 Then
                s = Replace(s, "==", "")
            Else
                s = DMarkLines(s)
            End If
            c.NumberFormat = "@"
            c.Value = s
            n = n + 1
        End If
        If n > 500 Then Exit For
    Next
    If n = 0 Then Info "칠할 칸(부서 시트의 내용·비고 칸)을 먼저 고르세요. 칠한 칸을 다시 고르고 누르면 지워집니다."
End Sub

' wraps the text part of each line in ==...== (markers, time and @place stay outside)
Public Function DMarkLines(ByVal s As String) As String
    Dim ls As Variant, i As Long, ln As String, a As Long, b As Long, t As String, kind As String, mk As String, rank As Double, txt As String, raw As String, p As Long, tl As Long
    ls = SplitLines(s)
    For i = 0 To UBound(ls)
        ln = CStr(ls(i))
        a = SkipWs(ln, 1) - 1
        t = Mid$(ln, a + 1)
        If StartsWith(t, ">>") Then
            a = a + 2
        ElseIf MatchListItem(TrimX(t), kind, mk, rank, txt, raw) Then
            a = a + (Len(RTrimX(t)) - Len(RTrimX(txt)))
        End If
        a = SkipWs(ln, a + 1) - 1
        tl = DTimeLen(Mid$(ln, a + 1))
        If tl > 0 Then a = SkipWs(ln, a + tl + 1) - 1
        b = Len(RTrimX(ln))
        p = InStrRev(Left$(ln, b), "@")
        If p > 1 Then
            If IsWs(Mid$(ln, p - 1, 1)) And TrimX(Mid$(ln, p + 1, b - p)) <> "" Then
                Do While p > 1
                    If Not IsWs(Mid$(ln, p - 1, 1)) Then Exit Do
                    p = p - 1
                Loop
                If p - 1 > a Then b = p - 1
            End If
        End If
        If b > a Then ls(i) = Left$(ln, a) & "==" & Mid$(ln, a + 1, b - a) & "==" & Mid$(ln, b + 1)
    Next
    DMarkLines = Join(ls, vbLf)
End Function

' ---- co-edit template --------------------------------------------------------
Public Sub MakeTemplate()
    Dim p As Variant
    On Error Resume Next
    p = Application.GetSaveAsFilename("일일예정_공동편집양식.xlsx", "엑셀 파일 (*.xlsx),*.xlsx")
    On Error GoTo fail
    If VarType(p) = vbBoolean Then Exit Sub
    If CStr(p) = "" Or LCase$(CStr(p)) = "false" Then Exit Sub
    MakeTemplateTo CStr(p)
    Exit Sub
fail:
    Application.DisplayAlerts = True
    Info "문제가 생겼습니다: " & Err.Description & vbLf & "(" & gStep & ")"
End Sub

Public Function MakeTemplateTo(ByVal path As String) As String
    Dim wb As Object, ws As Object, k As Long, guide As Variant, i As Long, prev As Object, nm As String, ex As Variant, r As Long
    LoadSettings
    LoadDepts False
    If dN = 0 Then
        Info "부서를 먼저 추가하세요."
        Exit Function
    End If
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Set ws = wb.Worksheets(1)
    ws.Name = "작성방법"
    guide = Array("부서별 일일 예정사항 — 공동편집 양식", "", _
        "1. 자기 부서 시트에만 입력하세요. 시트 이름이 부서명이니 바꾸지 마세요.", _
        "2. 한 줄에 일정 하나씩: 시간 / 내용 / 장소를 각자 칸에 씁니다. 시간은 09:00 또는 14:00~15:30처럼 씁니다.", _
        "3. 세부 내용은 내용 칸 안에서 줄을 바꿔(Alt+Enter) 씁니다. ""- ""로 시작하면 하위 항목, ""* ""로 시작하면 작은 글씨입니다.", _
        "4. 내용이 ""* ""로 시작하는 줄은 바로 위 일정의 하위 항목(작은 글씨)이 됩니다.", _
        "5. 비고는 작은 글씨로 붙고, 설명은 화면에서 마우스를 올리면 보이는 메모, 링크는 누르면 열리는 주소입니다(설명·링크 주소는 인쇄되지 않음).", _
        "6. 여러 날을 한 시트에 쓰려면 맨 앞에 ""날짜"" 열을 넣으세요. 문서 날짜에 맞는 줄만 가져옵니다.", _
        "7. 표를 넣으려면 일정 아래 한 줄을 비우고, A열부터 표 제목(한 칸, 선택) → 표 머리글 → 표 내용 순서로 적습니다. 다음 빈 줄까지가 표입니다.", _
        "8. 다 쓰면 담당자가 파일을 내려받아 한셀 파일의 [엑셀 불러오기]로 가져옵니다.", "", "예시")
    For i = 0 To UBound(guide)
        ws.Cells(i + 1, 1).Value = guide(i)
    Next
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(UBound(guide) + 1, 1).Font.Bold = True
    r = UBound(guide) + 2
    TemplateHeaderD ws, r, Array("시간", "내용", "장소", "비고", "설명 (마우스를 올리면 보임)", "링크")
    ex = Array(Array("09:00", "주간 업무보고 회의" & vbLf & "- 안건: 예산안 검토", "본관 3층 대회의실", "참석: 실장, 각 과장", "자료는 회의 전날 배포", "http://intranet/board/123"), _
        Array("14:00~15:30", "예산안 검토 회의", "소회의실", "", "", ""), _
        Array("", "* 실무 협의는 별도 진행", "", "", "", ""), _
        Array("", "상반기 업무계획 실적 취합 (17:00 마감)", "", "", "", ""))
    For i = 0 To UBound(ex)
        TemplateRowD ws, r + 1 + i, ex(i)
    Next
    r = r + UBound(ex) + 3
    ws.Cells(r, 1).Value = "■ 이번 주 교육 일정"
    TemplateHeaderD ws, r + 1, Array("교육명", "일자", "장소", "대상")
    TemplateRowD ws, r + 2, Array("청렴 교육", "11. 12.(목)", "대회의실", "전 직원")
    TemplateRowD ws, r + 3, Array("개인정보보호 교육", "11. 13.(금)", "온라인", "전 직원")
    TemplateWidthsD ws
    Set prev = ws
    For k = 1 To dN
        Set ws = wb.Worksheets.Add(, prev)
        nm = SafeSheetName(dName(k))
        If nm = "" Then nm = "부서" & k
        On Error Resume Next
        ws.Name = nm
        If Err.Number <> 0 Then ws.Name = Left$(nm, 27) & "(" & k & ")"
        On Error GoTo 0
        TemplateHeaderD ws, 1, Array("시간", "내용", "장소", "비고", "설명 (마우스를 올리면 보임)", "링크")
        ws.Range("A2:F300").NumberFormat = "@"
        ws.Range("A2:F300").WrapText = True
        ws.Range("A2:F300").VerticalAlignment = XL_TOP
        TemplateWidthsD ws
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
    Info "공동편집용 엑셀 양식을 저장했습니다 (부서 " & dN & "곳)." & vbLf & path & vbLf & "두레이에 올려 함께 작성하세요."
    MakeTemplateTo = "ok"
End Function

Private Sub TemplateHeaderD(ByVal ws As Object, ByVal r As Long, ByVal heads As Variant)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(heads) + 1))
        .Value = heads
        .Font.Bold = True
        .Interior.Color = COL_HEAD
        .HorizontalAlignment = XL_CENTER
        .Borders.LineStyle = XL_CONT
    End With
End Sub

Private Sub TemplateRowD(ByVal ws As Object, ByVal r As Long, ByVal vals As Variant)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(vals) + 1))
        .NumberFormat = "@"
        .Value = vals
        .Borders.LineStyle = XL_CONT
        .WrapText = True
        .VerticalAlignment = XL_TOP
    End With
End Sub

Private Sub TemplateWidthsD(ByVal ws As Object)
    Dim w As Variant, i As Long
    w = Array(13, 52, 20, 22, 34, 30)
    For i = 0 To 5
        ws.Columns(i + 1).ColumnWidth = w(i)
    Next
End Sub

' ---- PDF / password ------------------------------------------------------------
Public Sub SavePdf()
    Dim p As Variant
    On Error Resume Next
    p = Application.GetSaveAsFilename("일일예정사항_" & NStr("d_date") & ".pdf", "PDF (*.pdf),*.pdf")
    If VarType(p) = vbBoolean Then Exit Sub
    If CStr(p) = "" Or LCase$(CStr(p)) = "false" Then Exit Sub
    ThisWorkbook.Worksheets(SH_A4).ExportAsFixedFormat 0, CStr(p)
    If Err.Number <> 0 Then
        Info "매크로로 PDF를 만들지 못했습니다. A4 시트에서 한셀의 [PDF로 저장](또는 인쇄 → PDF)을 쓰세요."
    Else
        Info "PDF로 저장했습니다. 링크는 PDF에서도 눌립니다." & vbLf & CStr(p)
    End If
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
    ThisWorkbook.Password = p1
    If Err.Number = 0 Then ThisWorkbook.Save
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
