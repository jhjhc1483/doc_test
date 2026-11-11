Attribute VB_Name = "HcDailyMd"
Option Explicit

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
Public Function LayContent(ByVal d As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal useHalves As Boolean, _
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
