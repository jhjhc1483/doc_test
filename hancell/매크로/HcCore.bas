Attribute VB_Name = "HcCore"
Option Explicit
' ===========================================================================
' HcCore - shared helpers for the Hancell (VBA) schedule workbooks
'   * text helpers (no RegExp, so it runs without VBScript)
'   * time parsing (port of splitTime in the HTML version)
'   * rich text builder (**bold** __under__ !!red!! ==mark== ~~strike~~)
'   * layout engine: place rectangles in points, then turn them into a cell grid
'   * text measuring with a hidden scratch sheet (Rows.AutoFit)
' Strings with Korean text are stored as U("\uXXXX") by tools/build.py.
' ===========================================================================

Public Const XL_LEFT As Long = -4131
Public Const XL_CENTER As Long = -4108
Public Const XL_RIGHT As Long = -4152
Public Const XL_TOP As Long = -4160
Public Const XL_BOTTOM As Long = -4107
Public Const XL_CONT As Long = 1
Public Const XL_DOT As Long = -4118
Public Const XL_DASH As Long = -4115
Public Const XL_HAIR As Long = 1
Public Const XL_THIN As Long = 2
Public Const XL_MEDIUM As Long = -4138
Public Const XL_THICK As Long = 4
Public Const XL_EDGE_LEFT As Long = 7
Public Const XL_EDGE_TOP As Long = 8
Public Const XL_EDGE_BOTTOM As Long = 9
Public Const XL_EDGE_RIGHT As Long = 10
Public Const XL_PORTRAIT As Long = 1
Public Const XL_LANDSCAPE As Long = 2
Public Const XL_PAPER_A4 As Long = 9
Public Const XL_SHEET_HIDDEN As Long = 0
Public Const XL_SHEET_VISIBLE As Long = -1
Public Const PT_MM As Double = 2.834645669
Public Const PX As Double = 0.75
Public Const MIN_COL As Double = 4.6     ' narrowest column we create (pt)

Public gQuiet As Boolean          ' True while building/testing: no message boxes
Public gStep As String           ' last step (for error messages)
Public gTrace As String          ' last low-level call (for error messages)
Public gDbg As String
Public gHostLO As Boolean          ' set only by the LibreOffice test tool

' ---------------------------------------------------------------------------
' Unicode strings
' ---------------------------------------------------------------------------
Public Function U(ByVal s As String) As String
    Dim i As Long, p As Long, n As Long, out As String
    n = Len(s)
    i = 1
    Do
        p = InStr(i, s, "\u")
        If p = 0 Or p + 5 > n Then
            out = out & Mid$(s, i)
            Exit Do
        End If
        out = out & Mid$(s, i, p - i) & ChrW(HexVal(Mid$(s, p + 2, 4)))
        i = p + 6
        If i > n Then Exit Do
    Loop
    U = out
End Function

Private Function HexVal(ByVal h As String) As Long
    Dim i As Long, v As Long, d As Long
    For i = 1 To Len(h)
        d = InStr("0123456789ABCDEF", UCase$(Mid$(h, i, 1))) - 1
        If d < 0 Then d = 0
        v = v * 16 + d
    Next
    HexVal = v
End Function

Public Function CodeOf(ByVal ch As String) As Long
    If Len(ch) = 0 Then
        CodeOf = -1
        Exit Function
    End If
    CodeOf = AscW(Left$(ch, 1))
    If CodeOf < 0 Then CodeOf = CodeOf + 65536
End Function

Public Function IsWsCode(ByVal c As Long) As Boolean
    IsWsCode = (c = 32 Or c = 9 Or c = 10 Or c = 13 Or c = 11 Or c = 12 Or c = 160 Or c = 12288 Or c = 65279 Or c = 8203)
End Function

Public Function IsWs(ByVal ch As String) As Boolean
    IsWs = IsWsCode(CodeOf(ch))
End Function

Public Function IsDigitAt(ByVal s As String, ByVal p As Long) As Boolean
    Dim c As Long
    If p < 1 Or p > Len(s) Then Exit Function
    c = CodeOf(Mid$(s, p, 1))
    IsDigitAt = (c >= 48 And c <= 57)
End Function

Public Function IsHangulAt(ByVal s As String, ByVal p As Long) As Boolean
    Dim c As Long
    If p < 1 Or p > Len(s) Then Exit Function
    c = CodeOf(Mid$(s, p, 1))
    IsHangulAt = (c >= 44032 And c <= 55203)
End Function

Public Function IsWideCode(ByVal c As Long) As Boolean
    ' CJK / Hangul / full-width: about 1em wide
    IsWideCode = (c >= 4352 And c <= 4607) Or (c >= 12288 And c <= 12351) Or (c >= 12592 And c <= 12687) _
        Or (c >= 19968 And c <= 40959) Or (c >= 44032 And c <= 55215) Or (c >= 65280 And c <= 65519)
End Function

Public Function TextWidthEm(ByVal s As String) As Double
    Dim i As Long, w As Double
    For i = 1 To Len(s)
        If IsWideCode(CodeOf(Mid$(s, i, 1))) Then
            w = w + 1
        Else
            w = w + 0.58
        End If
    Next
    TextWidthEm = w
End Function

Public Function SkipWs(ByVal s As String, ByVal p As Long) As Long
    Do While p <= Len(s)
        If Not IsWs(Mid$(s, p, 1)) Then Exit Do
        p = p + 1
    Loop
    SkipWs = p
End Function

Public Function TrimX(ByVal s As String) As String
    Dim a As Long, b As Long
    a = 1
    b = Len(s)
    Do While a <= b
        If Not IsWs(Mid$(s, a, 1)) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not IsWs(Mid$(s, b, 1)) Then Exit Do
        b = b - 1
    Loop
    If b >= a Then
        TrimX = Mid$(s, a, b - a + 1)
    Else
        TrimX = ""
    End If
End Function

Public Function RTrimX(ByVal s As String) As String
    Dim b As Long
    b = Len(s)
    Do While b >= 1
        If Not IsWs(Mid$(s, b, 1)) Then Exit Do
        b = b - 1
    Loop
    RTrimX = Left$(s, b)
End Function

Public Function NormNL(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    NormNL = Replace(s, vbCr, vbLf)
End Function

' collapse runs of tabs into one space
Public Function TabsToSpace(ByVal s As String) As String
    Do While InStr(s, vbTab & vbTab) > 0
        s = Replace(s, vbTab & vbTab, vbTab)
    Loop
    TabsToSpace = Replace(s, vbTab, " ")
End Function

' collapse every whitespace run into a single space
Public Function CollapseWs(ByVal s As String) As String
    Dim i As Long, out As String, prevWs As Boolean, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsWs(ch) Then
            If Not prevWs Then out = out & " "
            prevWs = True
        Else
            out = out & ch
            prevWs = False
        End If
    Next
    CollapseWs = out
End Function

Public Function StartsWith(ByVal s As String, ByVal pre As String) As Boolean
    StartsWith = (Left$(s, Len(pre)) = pre)
End Function

' first character is one of the characters in chars
Public Function FirstIn(ByVal s As String, ByVal chars As String) As Boolean
    If Len(s) = 0 Then Exit Function
    FirstIn = (InStr(chars, Left$(s, 1)) > 0)
End Function

' remove leading characters that are in chars, then whitespace
Public Function StripLead(ByVal s As String, ByVal chars As String) As String
    Do While Len(s) > 0
        If InStr(chars, Left$(s, 1)) = 0 Then Exit Do
        s = Mid$(s, 2)
    Loop
    StripLead = TrimX(s)
End Function

' remove "(...)" parts
Public Function RemoveParens(ByVal s As String) As String
    Dim a As Long, b As Long
    Do
        a = InStr(s, "(")
        If a = 0 Then Exit Do
        b = InStr(a + 1, s, ")")
        If b = 0 Then Exit Do
        s = Left$(s, a - 1) & Mid$(s, b + 1)
    Loop
    RemoveParens = s
End Function

' remove every character listed in chars (and all whitespace when stripWs)
Public Function RemoveChars(ByVal s As String, ByVal chars As String, Optional ByVal stripWs As Boolean = True) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(chars, ch) = 0 Then
            If Not (stripWs And IsWs(ch)) Then out = out & ch
        End If
    Next
    RemoveChars = out
End Function

Public Function SplitLines(ByVal s As String) As Variant
    SplitLines = Split(NormNL(s), vbLf)
End Function

Public Function Clamp(ByVal v As Double, ByVal a As Double, ByVal b As Double) As Double
    If v < a Then v = a
    If v > b Then v = b
    Clamp = v
End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        MaxD = a
    Else
        MaxD = b
    End If
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        MinD = a
    Else
        MinD = b
    End If
End Function

' font sizes in 0.5pt steps (what the spreadsheet accepts)
Public Function R5(ByVal v As Double) As Double
    R5 = Int(v * 2 + 0.5) / 2
End Function

Public Function NumStr(ByVal v As Double, Optional ByVal digits As Long = 1) As String
    Dim s As String
    s = CStr(Round(v, digits))
    NumStr = Replace(s, ",", ".")
End Function

' Collection helpers
Public Function ColToArray(ByVal c As Collection) As Variant
    Dim a() As Variant, i As Long
    If c.Count = 0 Then
        ColToArray = Array()
        Exit Function
    End If
    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count
        a(i - 1) = c(i)
    Next
    ColToArray = a
End Function

' position (1-based) of key in keys(1..n), 0 when missing
Public Function FindKey(keys() As String, ByVal n As Long, ByVal key As String) As Long
    Dim i As Long
    For i = 1 To n
        If keys(i) = key Then
            FindKey = i
            Exit Function
        End If
    Next
End Function

Public Function ArrLen(ByVal a As Variant) As Long
    On Error Resume Next
    ArrLen = UBound(a) - LBound(a) + 1
    If Err.Number <> 0 Then ArrLen = 0
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' Time at the start of a line (port of splitTime)
'
' ---------------------------------------------------------------------------
Private cN As Long
Private cE() As Long, cH() As Long, cM() As Long, cPm() As Long

Private Sub AddCand(ByVal e As Long, ByVal h As Long, ByVal mi As Long, ByVal pm As Long)
    cN = cN + 1
    ReDim Preserve cE(1 To cN): ReDim Preserve cH(1 To cN): ReDim Preserve cM(1 To cN): ReDim Preserve cPm(1 To cN)
    cE(cN) = e: cH(cN) = h: cM(cN) = mi: cPm(cN) = pm
End Sub

Private Function IsColonAt(ByVal s As String, ByVal p As Long) As Boolean
    Dim c As Long
    If p < 1 Or p > Len(s) Then Exit Function
    c = CodeOf(Mid$(s, p, 1))
    IsColonAt = (c = 58 Or c = 65306)
End Function

Private Function DigitsVal(ByVal s As String, ByVal p As Long, ByVal n As Long) As Long
    DigitsVal = CLng(Mid$(s, p, n))
End Function

' all possible matches of one time token starting at p, in regex preference order
Private Sub TimeCands(ByVal s As String, ByVal p As Long)
    Dim q As Long, pm As Long, h As Long, q2 As Long, q3 As Long, q4 As Long, q5 As Long, mi As Long, t As String
    cN = 0
    pm = -1
    q = p
    t = Mid$(s, q, 2)
    If t = U("\uC624\uC804") Or t = "AM" Or t = "am" Then
        pm = 0
    ElseIf t = U("\uC624\uD6C4") Or t = "PM" Or t = "pm" Then
        pm = 1
    End If
    If pm >= 0 Then
        q = SkipWs(s, q + 2)
        If Not IsDigitAt(s, q) Then
            pm = -1
            q = p
        End If
    End If
    If Not IsDigitAt(s, q) Then Exit Sub
    h = DigitsVal(s, q, 1)
    q = q + 1
    If IsDigitAt(s, q) Then
        h = h * 10 + DigitsVal(s, q, 1)
        q = q + 1
    End If
    q2 = SkipWs(s, q)
    If IsColonAt(s, q2) Then
        q3 = SkipWs(s, q2 + 1)
        If IsDigitAt(s, q3) And IsDigitAt(s, q3 + 1) Then
            mi = DigitsVal(s, q3, 2)
            q3 = q3 + 2
            q4 = SkipWs(s, q3)
            If IsColonAt(s, q4) Then
                q5 = SkipWs(s, q4 + 1)
                If IsDigitAt(s, q5) And IsDigitAt(s, q5 + 1) Then AddCand q5 + 2, h, mi, pm
            End If
            AddCand q3, h, mi, pm
        End If
    ElseIf Mid$(s, q2, 1) = U("\uC2DC") Then
        q3 = q2 + 1
        q4 = SkipWs(s, q3)
        If IsDigitAt(s, q4) Then
            mi = DigitsVal(s, q4, 1)
            q5 = q4 + 1
            If IsDigitAt(s, q5) Then
                mi = mi * 10 + DigitsVal(s, q5, 1)
                q5 = q5 + 1
            End If
            q5 = SkipWs(s, q5)
            If Mid$(s, q5, 1) = U("\uBD84") Then
                If Not IsHangulAt(s, q5 + 1) Then AddCand q5 + 1, h, mi, pm
            End If
        End If
        q4 = SkipWs(s, q3)
        If Mid$(s, q4, 1) = U("\uBC18") Then
            If Not IsHangulAt(s, q4 + 1) Then AddCand q4 + 1, h, 30, pm
        End If
        If Not IsHangulAt(s, q3) Then AddCand q3, h, 0, pm
    End If
End Sub

Private Function IsSepAt(ByVal s As String, ByVal p As Long) As Boolean
    Dim c As Long
    If p < 1 Or p > Len(s) Then Exit Function
    c = CodeOf(Mid$(s, p, 1))
    IsSepAt = (c = 126 Or c = 8764 Or c = 65374 Or c = 12316 Or c = 45 Or c = 8211 Or c = 8212)
End Function

' normalize: returns "" when invalid. pmIO: in = parsed (-1/0/1), out = final
Private Function NormPart(ByVal h As Long, ByVal mi As Long, ByRef pmIO As Long, ByVal pmDefault As Boolean) As String
    Dim pm As Long
    pm = pmIO
    If pm = -1 And pmDefault And h < 12 Then pm = 1
    If pm = 1 And h < 12 Then h = h + 12
    If h > 24 Or mi > 59 Then
        NormPart = ""
        Exit Function
    End If
    pmIO = pm
    NormPart = Format$(h, "00") & ":" & Format$(mi, "00")
End Function

Public Function SplitTime(ByVal ln As String, ByRef tm As String, ByRef rest As String) As Boolean
    Dim nA As Long, ia As Long, ib As Long, nB As Long, q As Long, endPos As Long, hasB As Boolean, found As Boolean
    Dim aE() As Long, aH() As Long, aM() As Long, aPm() As Long
    Dim bH As Long, bM As Long, bPm As Long, selH As Long, selM As Long, selPm As Long
    Dim ta As String, tb As String, r As String, w As String, wordEnd As Long, ws As Variant
    tm = ""
    rest = ""
    TimeCands ln, 1
    nA = cN
    If nA > 0 Then
        ReDim aE(1 To nA): ReDim aH(1 To nA): ReDim aM(1 To nA): ReDim aPm(1 To nA)
        For ia = 1 To nA
            aE(ia) = cE(ia): aH(ia) = cH(ia): aM(ia) = cM(ia): aPm(ia) = cPm(ia)
        Next
        For ia = 1 To nA
            q = SkipWs(ln, aE(ia))
            If IsSepAt(ln, q) Then
                TimeCands ln, SkipWs(ln, q + 1)
                nB = cN
                For ib = 1 To nB
                    If Not IsDigitAt(ln, cE(ib)) Then
                        hasB = True
                        endPos = cE(ib)
                        bH = cH(ib): bM = cM(ib): bPm = cPm(ib)
                        found = True
                        Exit For
                    End If
                Next
            End If
            If Not found Then
                If Not IsDigitAt(ln, aE(ia)) Then
                    endPos = aE(ia)
                    found = True
                End If
            End If
            If found Then
                selH = aH(ia): selM = aM(ia): selPm = aPm(ia)
                Exit For
            End If
        Next
    End If
    If found Then
        ta = NormPart(selH, selM, selPm, False)
        If ta <> "" And hasB Then tb = NormPart(bH, bM, bPm, (selPm = 1))
        If ta <> "" And (Not hasB Or tb <> "") Then
            tm = ta
            If hasB Then tm = tm & "~" & tb
            r = Mid$(ln, SkipWs(ln, endPos))
            If Len(r) > 0 Then
                If InStr("-" & ChrW(8211) & ChrW(8212) & ":|,", Left$(r, 1)) > 0 Then r = Mid$(r, 2)
            End If
            rest = TrimX(r)
            SplitTime = True
            Exit Function
        End If
    End If
    '
    If StartsWith(ln, U("\uD558\uB8E8")) Then
        q = SkipWs(ln, 3)
        If Mid$(ln, q, 2) = U("\uC885\uC77C") Then
            wordEnd = q + 2
            If q > 3 Then
                w = U("\uD558\uB8E8 \uC885\uC77C")
            Else
                w = U("\uD558\uB8E8\uC885\uC77C")
            End If
        End If
    End If
    If wordEnd = 0 Then
        For Each ws In Array(U("\uC885\uC77C"), U("\uC624\uC804"), U("\uC624\uD6C4"), U("\uBBF8\uC815"), U("\uC218\uC2DC"))
            If StartsWith(ln, CStr(ws)) Then
                w = CStr(ws)
                wordEnd = 3
                Exit For
            End If
        Next
    End If
    If wordEnd > 0 Then
        If wordEnd > Len(ln) Then
            tm = w
            rest = ""
            SplitTime = True
        ElseIf IsWs(Mid$(ln, wordEnd, 1)) Then
            tm = w
            rest = TrimX(Mid$(ln, wordEnd))
            SplitTime = True
        End If
    End If
End Function

Public Function IsTimeCell(ByVal v As String) As Boolean
    Dim tm As String, rest As String
    If SplitTime(TrimX(v), tm, rest) Then IsTimeCell = (rest = "")
End Function

' ---------------------------------------------------------------------------
' Dates
' ---------------------------------------------------------------------------
Public Function IsoOf(ByVal d As Date) As String
    IsoOf = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
End Function

Public Function IsoToDate(ByVal iso As String) As Date
    IsoToDate = DateSerial(CLng(Left$(iso, 4)), CLng(Mid$(iso, 6, 2)), CLng(Mid$(iso, 9, 2)))
End Function

Public Function IsIso(ByVal s As String) As Boolean
    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    IsIso = IsNumeric(Left$(s, 4)) And IsNumeric(Mid$(s, 6, 2)) And IsNumeric(Mid$(s, 9, 2))
End Function

Public Function WeekK(ByVal d As Date) As String
    WeekK = Mid$(U("\uC77C\uC6D4\uD654\uC218\uBAA9\uAE08\uD1A0"), Weekday(d), 1)
End Function

'
Public Function FmtDateK(ByVal iso As String) As String
    If Not IsIso(iso) Then Exit Function
    Dim d As Date
    d = IsoToDate(iso)
    FmtDateK = Year(d) & ". " & Month(d) & ". " & Day(d) & ".(" & WeekK(d) & ")"
End Function

' a cell value -> ISO date ("" when it is not a date)
Public Function DateValueIso(ByVal v As Variant, Optional ByVal y0 As Long = 0) As String
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        DateValueIso = IsoOf(CDate(v))
        Exit Function
    End If
    If IsNumeric(v) And VarType(v) <> vbString Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then
            DateValueIso = IsoOf(CDate(CDbl(v)))
            Exit Function
        End If
    End If
    DateValueIso = ParseLooseDate(CStr(v), y0)
End Function

'
Public Function ParseLooseDate(ByVal v As String, Optional ByVal y0 As Long = 0) As String
    Dim p As Long, a As Long, b As Long, c As Long, n1 As Long, n2 As Long, n3 As Long
    v = TrimX(v)
    If v = "" Then Exit Function
    If y0 = 0 Then y0 = Year(Date)
    p = 1
    n1 = ReadNum(v, p, a)
    If a = 0 Then Exit Function
    If a = 4 Then
        p = SkipWs(v, p)
        If Not SepDate(v, p, U("\uB144")) Then Exit Function
        p = SkipWs(v, p + 1)
        n2 = ReadNum(v, p, b)
        If b = 0 Or b > 2 Then Exit Function
        p = SkipWs(v, p)
        If Not SepDate(v, p, U("\uC6D4")) Then Exit Function
        p = SkipWs(v, p + 1)
        n3 = ReadNum(v, p, c)
        If c = 0 Or c > 2 Then Exit Function
        ParseLooseDate = MkIso(n1, n2, n3)
    ElseIf a <= 2 Then
        p = SkipWs(v, p)
        If Not SepDate(v, p, U("\uC6D4")) Then Exit Function
        p = SkipWs(v, p + 1)
        n2 = ReadNum(v, p, b)
        If b = 0 Or b > 2 Then Exit Function
        ParseLooseDate = MkIso(y0, n1, n2)
    End If
End Function

Private Function ReadNum(ByVal s As String, ByRef p As Long, ByRef cnt As Long) As Long
    Dim v As Long
    cnt = 0
    Do While IsDigitAt(s, p) And cnt < 6
        v = v * 10 + CLng(Mid$(s, p, 1))
        p = p + 1
        cnt = cnt + 1
    Loop
    ReadNum = v
End Function

Private Function SepDate(ByVal s As String, ByVal p As Long, ByVal word As String) As Boolean
    Dim ch As String
    ch = Mid$(s, p, 1)
    SepDate = (ch = "-" Or ch = "." Or ch = "/" Or ch = word)
End Function

Private Function MkIso(ByVal y As Long, ByVal m As Long, ByVal d As Long) As String
    If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then MkIso = Format$(y, "0000") & "-" & Format$(m, "00") & "-" & Format$(d, "00")
End Function

' ---------------------------------------------------------------------------
' Cell value -> text (dates and times as people see them)
' ---------------------------------------------------------------------------
Public Function CellText(ByVal v As Variant, Optional ByVal asTime As Boolean = False) As String
    Dim s As String, d As Double
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = CDbl(v)
        If d < 1 And d >= 0 Then
            s = Format$(v, "hh:mm")
        ElseIf d = Int(d) Then
            s = IsoOf(CDate(v))
        Else
            s = IsoOf(CDate(v)) & " " & Format$(v, "hh:mm")
        End If
    ElseIf VarType(v) = vbString Then
        s = CStr(v)
    ElseIf IsNumeric(v) Then
        If asTime And CDbl(v) >= 0 And CDbl(v) < 1 Then
            s = Format$(CDate(CDbl(v)), "hh:mm")
        Else
            s = CStr(v)
            s = Replace(s, ",", ".")
        End If
    Else
        s = CStr(v)
    End If
    s = Replace(s, ChrW(160), " ")
    CellText = TrimX(NormNL(s))
End Function

' ---------------------------------------------------------------------------
' Rich text builder
' ---------------------------------------------------------------------------
Private rbTxt As String
Private rbRunStr As String
Public rbMark As Boolean

Public Sub RBReset()
    rbTxt = ""
    rbRunStr = ""
    rbMark = False
End Sub

Public Function RBText() As String
    RBText = rbTxt
End Function

Public Function RBRuns() As String
    RBRuns = rbRunStr
End Function

' size = 0: base size. color = -1: base color
Public Sub RBAdd(ByVal t As String, Optional ByVal bold As Boolean = False, Optional ByVal under As Boolean = False, _
        Optional ByVal color As Long = -1, Optional ByVal size As Double = 0, Optional ByVal strike As Boolean = False, _
        Optional ByVal italic As Boolean = False)
    If Len(t) = 0 Then Exit Sub
    Dim st As Long
    st = Len(rbTxt) + 1
    rbTxt = rbTxt & t
    If bold Or under Or color >= 0 Or size > 0 Or strike Or italic Then
        If Len(rbRunStr) > 0 Then rbRunStr = rbRunStr & ";"
        rbRunStr = rbRunStr & st & "," & Len(t) & "," & IIf(bold, 1, 0) & "," & IIf(under, 1, 0) & "," & color & "," & CLng(size * 10) & "," & IIf(strike, 1, 0) & "," & IIf(italic, 1, 0)
    End If
End Sub

' inline marks: **bold** __under__ !!red!! ==mark== (~~strike~~ when allowStrike)
Public Sub RBInline(ByVal s As String, Optional ByVal bold As Boolean = False, Optional ByVal size As Double = 0, _
        Optional ByVal color As Long = -1, Optional ByVal allowStrike As Boolean = False)
    Dim i As Long, n As Long, buf As String, k As Long, mk As String
    Dim f() As Boolean
    ReDim f(1 To 5)
    n = Len(s)
    i = 1
    Do While i <= n
        mk = Mid$(s, i, 2)
        k = 0
        If mk = "**" Then
            k = 1
        ElseIf mk = "__" Then
            k = 2
        ElseIf mk = "!!" Then
            k = 3
        ElseIf mk = "==" Then
            k = 4
        ElseIf mk = "~~" And allowStrike Then
            k = 5
        End If
        If k > 0 Then
            If f(k) Then
                RBFlush buf, f, bold, size, color
                buf = ""
                f(k) = False
                i = i + 2
                GoTo nextCh
            ElseIf InStr(i + 3, s, mk) > 0 Then
                RBFlush buf, f, bold, size, color
                buf = ""
                f(k) = True
                i = i + 2
                GoTo nextCh
            End If
        End If
        buf = buf & Mid$(s, i, 1)
        i = i + 1
nextCh:
    Loop
    RBFlush buf, f, bold, size, color
End Sub

Private Sub RBFlush(ByVal buf As String, f() As Boolean, ByVal bold As Boolean, ByVal size As Double, ByVal color As Long)
    If Len(buf) = 0 Then Exit Sub
    Dim c As Long
    c = color
    If f(3) Then c = RGB(196, 0, 0)
    If f(4) Then rbMark = True
    RBAdd buf, bold Or f(1) Or f(3), f(2), c, size, f(5)
End Sub

' strip inline marks (for plain text uses)
Public Function PlainInline(ByVal s As String) As String
    RBReset
    RBInline s, , , , True
    PlainInline = rbTxt
End Function

Public Sub ApplyRuns(ByVal cell As Object, ByVal runs As String)
    If Len(runs) = 0 Then Exit Sub
    Dim parts As Variant, f As Variant, i As Long
    parts = Split(runs, ";")
    On Error Resume Next
    For i = 0 To UBound(parts)
        f = Split(parts(i), ",")
        With cell.Characters(CLng(f(0)), CLng(f(1))).Font
            If f(2) = "1" Then .Bold = True
            If f(3) = "1" Then .Underline = 2
            If CLng(f(4)) >= 0 Then .Color = CLng(f(4))
            If CLng(f(5)) > 0 Then .Size = CLng(f(5)) / 10
            If f(6) = "1" Then .Strikethrough = True
            If UBound(f) >= 7 Then
                If f(7) = "1" Then .Italic = True
            End If
        End With
    Next
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Layout engine
'   Elements are rectangles in points (x right, y down from the top-left of the
'   printable area). Realize turns every distinct edge into a column/row border.
' ---------------------------------------------------------------------------
Private Const EK As Long = 0      ' 1 text, 2 box
Private Const EX As Long = 1
Private Const EY As Long = 2
Private Const EW As Long = 3
Private Const EH As Long = 4
Private Const ETXT As Long = 5
Private Const ERUN As Long = 6
Private Const ESIZE As Long = 7
Private Const EBOLD As Long = 8
Private Const EHAL As Long = 9
Private Const EVALIGN As Long = 10
Private Const EWRAP As Long = 11
Private Const EFILL As Long = 12
Private Const ECOLOR As Long = 13
Private Const EBT As Long = 14     ' borders top, bottom, left, right
Private Const ELINK As Long = 18
Private Const ENOTE As Long = 19
Private Const EUNDER As Long = 20
Private Const ENF As Long = 21

Private elN As Long
Private elV() As Variant

Public Sub ElReset()
    elN = 0
    ReDim elV(1 To 64)
End Sub

Private Function ElNew(ByVal kind As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double) As Long
    Dim a As Variant
    elN = elN + 1
    If elN > UBound(elV) Then ReDim Preserve elV(1 To UBound(elV) * 2)
    ' order: EK EX EY EW EH ETXT ERUN ESIZE EBOLD EHAL EVALIGN EWRAP EFILL ECOLOR EBT..(4) ELINK ENOTE EUNDER
    a = Array(kind, x, y, w, h, "", "", 0, False, XL_LEFT, XL_TOP, True, -1, 0, Empty, Empty, Empty, Empty, "", "", False)
    elV(elN) = a
    ElNew = elN
End Function

Private Sub ElSet(ByVal i As Long, ByVal k As Long, ByVal v As Variant)
    Dim a As Variant
    a = elV(i)
    a(k) = v
    elV(i) = a
End Sub

Public Function ElText(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal txt As String, _
        ByVal runs As String, ByVal size As Double, Optional ByVal bold As Boolean = False, Optional ByVal hal As Long = XL_LEFT, _
        Optional ByVal vAlign As Long = XL_TOP, Optional ByVal color As Long = 0, Optional ByVal wrap As Boolean = True) As Long
    Dim i As Long, a As Variant
    i = ElNew(1, x, y, w, h)
    a = elV(i)
    a(ETXT) = txt: a(ERUN) = runs: a(ESIZE) = size: a(EBOLD) = bold
    a(EHAL) = hal: a(EVALIGN) = vAlign: a(ECOLOR) = color: a(EWRAP) = wrap
    elV(i) = a
    ElText = i
End Function

Public Function ElBox(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, Optional ByVal fill As Long = -1) As Long
    Dim i As Long
    i = ElNew(2, x, y, w, h)
    ElSet i, EFILL, fill
    ElBox = i
End Function

' edge: 0 top, 1 bottom, 2 left, 3 right
Public Sub ElBorder(ByVal i As Long, ByVal edge As Long, ByVal style As Long, ByVal weight As Long, Optional ByVal color As Long = 0)
    ElSet i, EBT + edge, Array(style, weight, color)
End Sub

Public Sub ElBorderAll(ByVal i As Long, ByVal style As Long, ByVal weight As Long, Optional ByVal color As Long = 0)
    Dim e As Long
    For e = 0 To 3
        ElBorder i, e, style, weight, color
    Next
End Sub

Public Sub ElFill(ByVal i As Long, ByVal color As Long)
    ElSet i, EFILL, color
End Sub

' "#Sheet!A1" = place in this workbook, otherwise a web address
Public Sub ElLink(ByVal i As Long, ByVal url As String)
    ElSet i, ELINK, url
End Sub

Public Sub ElNote(ByVal i As Long, ByVal note As String)
    ElSet i, ENOTE, note
End Sub

Public Function SnapPx(ByVal v As Double) As Double
    SnapPx = Int(v / PX + 0.5) * PX
End Function

Private Sub AddEdge(ByRef arr() As Double, ByRef n As Long, ByVal v As Double)
    Dim i As Long, j As Long, pos As Long
    v = SnapPx(v)
    pos = n + 1
    For i = 1 To n
        If Abs(arr(i) - v) < 0.01 Then Exit Sub
        If arr(i) > v Then
            pos = i
            Exit For
        End If
    Next
    If n + 1 > UBound(arr) Then Exit Sub
    n = n + 1
    For j = n To pos + 1 Step -1
        arr(j) = arr(j - 1)
    Next
    arr(pos) = v
End Sub

' index of the edge nearest to v
Private Function EdgeIdx(ByRef arr() As Double, ByVal n As Long, ByVal v As Double) As Long
    Dim lo As Long, hi As Long, m As Long
    lo = 1
    hi = n
    Do While hi - lo > 1
        m = (lo + hi) \ 2
        If arr(m) <= v Then
            lo = m
        Else
            hi = m
        End If
    Loop
    If Abs(arr(hi) - v) < Abs(arr(lo) - v) Then
        EdgeIdx = hi
    Else
        EdgeIdx = lo
    End If
End Function

' merge edges closer than minGap (spreadsheet columns cannot be very narrow)
Private Sub ClusterEdges(ByRef arr() As Double, ByRef n As Long, ByVal minGap As Double)
    Dim i As Long, k As Long
    If n < 2 Then Exit Sub
    k = 1
    For i = 2 To n
        If arr(i) - arr(k) >= minGap Then
            k = k + 1
            arr(k) = arr(i)
        ElseIf i = n Then
            arr(k) = arr(i)
        End If
    Next
    n = k
End Sub

' turn elements into cells. Returns the last used row/col through lastR/lastC.
Public Sub Realize(ByVal ws As Object, ByVal fontName As String, ByRef lastR As Long, ByRef lastC As Long)
    Dim xs() As Double, ys() As Double, nx As Long, ny As Long, i As Long, e As Variant
    ' sized up front: growing an array passed ByRef loses data in some hosts
    ReDim xs(1 To 2 * elN + 16): ReDim ys(1 To 2 * elN + 80)
    AddEdge xs, nx, 0
    AddEdge ys, ny, 0
    For i = 1 To elN
        e = elV(i)
        If SnapPx(e(EX) + e(EW)) - SnapPx(e(EX)) >= PX And SnapPx(e(EY) + e(EH)) - SnapPx(e(EY)) >= PX Then
            AddEdge xs, nx, e(EX)
            AddEdge xs, nx, e(EX) + e(EW)
            AddEdge ys, ny, e(EY)
            AddEdge ys, ny, e(EY) + e(EH)
        End If
    Next
    If Len(gDbg) > 4000 Then gDbg = Right$(gDbg, 1000)
    gDbg = gDbg & " edges raw nx=" & nx & " ny=" & ny
    ClusterEdges xs, nx, MIN_COL
    gDbg = gDbg & " clustered nx=" & nx & " first=" & xs(1) & "," & xs(2) & "," & xs(3) & " last=" & xs(nx) & " elN=" & elN
    ' rows taller than 400pt are split (row height limit is 409pt)
    i = 1
    Do While i < ny
        If ys(i + 1) - ys(i) > 400 Then AddEdge ys, ny, ys(i) + 300
        i = i + 1
    Loop

    ClearSheet ws
    ' column widths (cumulative, so rounding never adds up)
    Dim cum As Double, target As Double
    cum = 0
    For i = 1 To nx - 1
        target = xs(i + 1) - cum
        If target < PX Then target = PX
        SetColPts ws, i, target
        cum = cum + ws.Columns(i).Width
    Next
    cum = 0
    For i = 1 To ny - 1
        target = ys(i + 1) - cum
        If target < PX Then target = PX
        If target > 409 Then target = 409
        ws.Rows(i).RowHeight = target
        cum = cum + ws.Rows(i).RowHeight
    Next
    lastC = nx - 1
    lastR = ny - 1

    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, rg As Object
    ' pass 1: box fills
    For i = 1 To elN
        e = elV(i)
        If e(EK) = 2 And e(EFILL) >= 0 Then
            If ElRange(ws, e, xs, nx, ys, ny, rg) Then rg.Interior.Color = e(EFILL)
        End If
    Next
    ' pass 2: text cells
    Application.DisplayAlerts = False
    For i = 1 To elN
        e = elV(i)
        If e(EK) = 1 Then
            If ElRange(ws, e, xs, nx, ys, ny, rg) Then
                If rg.Cells.Count > 1 Then rg.Merge
                rg.NumberFormat = "@"
                With rg.Font
                    .Name = fontName
                    .Size = e(ESIZE)
                    .Bold = e(EBOLD)
                    .Color = e(ECOLOR)
                End With
                rg.HorizontalAlignment = e(EHAL)
                rg.VerticalAlignment = e(EVALIGN)
                rg.WrapText = e(EWRAP)
                If e(EFILL) >= 0 Then rg.Interior.Color = e(EFILL)
                rg.Cells(1, 1).Value = e(ETXT)
                ApplyRuns rg.Cells(1, 1), e(ERUN)
                If Len(e(ELINK)) > 0 Then AddLink ws, rg.Cells(1, 1), e(ELINK), e(ECOLOR)
                If Len(e(ENOTE)) > 0 Then
                    On Error Resume Next
                    rg.Cells(1, 1).AddComment e(ENOTE)
                    On Error GoTo 0
                End If
            End If
        End If
    Next
    Application.DisplayAlerts = True
    ' pass 3: borders
    Dim k As Long, b As Variant
    For i = 1 To elN
        e = elV(i)
        If ElRange(ws, e, xs, nx, ys, ny, rg) Then
            For k = 0 To 3
                b = e(EBT + k)
                If IsArray(b) Then
                    With rg.Borders(Choose(k + 1, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_LEFT, XL_EDGE_RIGHT))
                        .LineStyle = b(0)
                        .Weight = b(1)
                        .Color = b(2)
                    End With
                End If
            Next
        End If
    Next
End Sub

Public Sub AddLink(ByVal ws As Object, ByVal cell As Object, ByVal url As String, Optional ByVal keepColor As Long = -1)
    On Error Resume Next
    If Left$(url, 1) = "#" Then
        ws.Hyperlinks.Add cell, "", Mid$(url, 2)
    Else
        ws.Hyperlinks.Add cell, url
    End If
    If keepColor >= 0 Then
        cell.Font.Color = keepColor
        cell.Font.Underline = -4142
    End If
    On Error GoTo 0
End Sub

Private Function ElRange(ByVal ws As Object, e As Variant, xs() As Double, ByVal nx As Long, ys() As Double, ByVal ny As Long, ByRef rg As Object) As Boolean
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    c1 = EdgeIdx(xs, nx, e(EX))
    c2 = EdgeIdx(xs, nx, e(EX) + e(EW)) - 1
    r1 = EdgeIdx(ys, ny, e(EY))
    r2 = EdgeIdx(ys, ny, e(EY) + e(EH)) - 1
    If c2 < c1 Or r2 < r1 Then Exit Function
    Set rg = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    ElRange = True
End Function

Public Sub ClearSheet(ByVal ws As Object)
    On Error Resume Next
    ws.Cells.UnMerge
    ws.Cells.ClearComments
    ws.Hyperlinks.Delete
    ws.Cells.Clear
    ws.Cells.RowHeight = 15
    ws.Cells.ColumnWidth = 8.43
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Column widths in points and text measuring
' ---------------------------------------------------------------------------
Private msWs As Object
Private msK As Double
Private msW1 As Double, msW2 As Double

' small string -> number hash table (keyed Collection lookups behave differently between hosts)
Private Const HB As Long = 2039
Private hHead() As Long, hNext() As Long, hKey() As String, hVal() As Double, hN As Long

Private Sub CacheReset()
    ReDim hHead(0 To HB - 1)
    ReDim hNext(1 To 256): ReDim hKey(1 To 256): ReDim hVal(1 To 256)
    hN = 0
End Sub

Private Function KeyHash(ByVal k As String) As Long
    Dim i As Long, h As Long
    For i = 1 To Len(k)
        h = (h * 31 + CodeOf(Mid$(k, i, 1))) Mod 1000003
    Next
    KeyHash = h Mod HB
End Function

Private Function CacheGet(ByVal k As String, ByRef v As Double) As Boolean
    Dim i As Long
    If hN = 0 Then Exit Function
    i = hHead(KeyHash(k))
    Do While i > 0
        If hKey(i) = k Then
            v = hVal(i)
            CacheGet = True
            Exit Function
        End If
        i = hNext(i)
    Loop
End Function

Private Sub CachePut(ByVal k As String, ByVal v As Double)
    Dim b As Long
    b = KeyHash(k)
    hN = hN + 1
    If hN > UBound(hKey) Then
        ReDim Preserve hKey(1 To hN * 2): ReDim Preserve hVal(1 To hN * 2): ReDim Preserve hNext(1 To hN * 2)
    End If
    hKey(hN) = k
    hVal(hN) = v
    hNext(hN) = hHead(b)
    hHead(b) = hN
End Sub

Public Function ColK(ByVal ws As Object) As Double
    ' points per ColumnWidth unit on this sheet
    Dim a As Double, b As Double, old As Double
    old = ws.Columns(1).ColumnWidth
    ws.Columns(1).ColumnWidth = 10
    a = ws.Columns(1).Width
    ws.Columns(1).ColumnWidth = 20
    b = ws.Columns(1).Width
    ws.Columns(1).ColumnWidth = old
    ColK = (b - a) / 10
    If ColK <= 0 Then ColK = 5.5
End Function

Public Sub SetColPts(ByVal ws As Object, ByVal col As Long, ByVal pts As Double)
    Dim w As Double, a As Double, it As Long, k As Double
    k = msK
    If k <= 0 Then k = 5.5
    w = pts / k
    If w < 0 Then w = 0
    If w > 254 Then w = 254
    gTrace = "SetColPts col=" & col & " pts=" & pts & " w=" & w & " k=" & k
    ws.Columns(col).ColumnWidth = w
    For it = 1 To 4
        a = ws.Columns(col).Width
        If Abs(a - pts) < 0.4 Then Exit For
        w = w + (pts - a) / k
        If w < 0 Then w = 0
        If w > 254 Then w = 254
        ws.Columns(col).ColumnWidth = w
    Next
End Sub

Public Function ScratchSheet(ByVal sheetName As String) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        ws.Visible = XL_SHEET_HIDDEN
    End If
    Set ScratchSheet = ws
End Function

Public Sub MsInit(ByVal sheetName As String)
    Set msWs = ScratchSheet(sheetName)
    msWs.Cells.Clear
    msK = ColK(msWs)
    CacheReset
    msW1 = -1
    msW2 = -1
End Sub

' height (pt) of wrapped text in a cell of the given width
Public Function MeasureH(ByVal txt As String, ByVal runs As String, ByVal widthPt As Double, ByVal fontName As String, _
        ByVal size As Double, Optional ByVal bold As Boolean = False) As Double
    If Len(txt) = 0 Then Exit Function
    Dim key As String, v As Double
    widthPt = Int(widthPt * 4 + 0.5) / 4
    key = "H|" & fontName & "|" & size & "|" & IIf(bold, 1, 0) & "|" & widthPt & "|" & runs & "|" & txt
    If CacheGet(key, v) Then
        MeasureH = v
        Exit Function
    End If
    gTrace = "MeasureH w=" & widthPt & " size=" & size & " font=" & fontName
    If Abs(widthPt - msW1) > 0.01 Then
        SetColPts msWs, 1, MaxD(widthPt - MIN_COL, MIN_COL)
        msW1 = widthPt
    End If
    With msWs.Cells(1, 1)
        .Clear
        .NumberFormat = "@"
        .WrapText = True
        .VerticalAlignment = XL_TOP
        .Font.Name = fontName
        .Font.Size = size
        .Font.Bold = bold
        .Value = txt
    End With
    If gHostLO Then Application.Run "LoShimFont", msWs.Name, "A1", fontName, size, bold
    ApplyRuns msWs.Cells(1, 1), runs
    msWs.Rows(1).AutoFit
    MeasureH = msWs.Rows(1).RowHeight
    CachePut key, MeasureH
End Function

' width (pt) a cell needs to show the text on one line (each line for multi-line text)
Public Function MeasureW(ByVal txt As String, ByVal fontName As String, ByVal size As Double, Optional ByVal bold As Boolean = False) As Double
    If Len(txt) = 0 Then Exit Function
    Dim key As String, v As Double
    key = "W|" & fontName & "|" & size & "|" & IIf(bold, 1, 0) & "|" & txt
    If CacheGet(key, v) Then
        MeasureW = v
        Exit Function
    End If
    gTrace = "MeasureW size=" & size & " font=" & fontName
    With msWs.Cells(1, 2)
        .Clear
        .NumberFormat = "@"
        .WrapText = False
        .Font.Name = fontName
        .Font.Size = size
        .Font.Bold = bold
        .Value = txt
    End With
    If gHostLO Then Application.Run "LoShimFont", msWs.Name, "B1", fontName, size, bold
    msWs.Columns(2).AutoFit
    MeasureW = msWs.Columns(2).Width
    CachePut key, MeasureW
End Function

' ---------------------------------------------------------------------------
' Workbook helpers
' ---------------------------------------------------------------------------
Public Function SheetByName(ByVal nm As String) As Object
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Public Function NameRange(ByVal nm As String) As Object
    On Error Resume Next
    Set NameRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

Public Function NVal(ByVal nm As String) As Variant
    Dim r As Object
    Set r = NameRange(nm)
    If r Is Nothing Then
        NVal = Empty
    Else
        NVal = r.Cells(1, 1).Value
    End If
End Function

Public Function NStr(ByVal nm As String) As String
    Dim v As Variant
    v = NVal(nm)
    NStr = CellText(v)
End Function

Public Function NNum(ByVal nm As String, ByVal def As Double, ByVal lo As Double, ByVal hi As Double) As Double
    Dim v As Variant
    v = NVal(nm)
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) = vbString Then
            If TrimX(CStr(v)) = "" Then
                NNum = def
                Exit Function
            End If
            NNum = Clamp(Val(Replace(CStr(v), ",", ".")), lo, hi)
        Else
            NNum = Clamp(CDbl(v), lo, hi)
        End If
    Else
        NNum = def
    End If
End Function

' "" or a number -> -1 when empty/invalid
Public Function NOptNum(ByVal nm As String) As Double
    Dim v As Variant
    v = NVal(nm)
    NOptNum = -1
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If TrimX(CStr(v)) = "" Then Exit Function
        If Not IsNumeric(Replace(CStr(v), ",", ".")) Then Exit Function
        NOptNum = Val(Replace(CStr(v), ",", "."))
    ElseIf IsNumeric(v) Then
        NOptNum = CDbl(v)
    End If
End Function

Public Function NBool(ByVal nm As String, ByVal def As Boolean) As Boolean
    Dim s As String
    s = NStr(nm)
    If s = "" Then
        NBool = def
    Else
        NBool = Not (s = U("\uC544\uB2C8\uC624") Or s = U("\uC544\uB2C8\uC694") Or s = U("\uB054") Or UCase$(s) = "N" Or UCase$(s) = "NO" Or s = "0" Or UCase$(s) = "FALSE")
    End If
End Function

Public Sub NSet(ByVal nm As String, ByVal v As Variant)
    Dim r As Object
    Set r = NameRange(nm)
    If Not r Is Nothing Then r.Cells(1, 1).Value = v
End Sub

Public Sub DefName(ByVal nm As String, ByVal ws As Object, ByVal addr As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add nm, "='" & ws.Name & "'!" & addr
End Sub

Public Function AddButton(ByVal ws As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
        ByVal caption As String, ByVal macro As String) As Object
    Dim b As Object
    On Error Resume Next
    Set b = ws.Buttons.Add(x, y, w, h)
    If Not b Is Nothing Then
        b.Caption = caption
        b.OnAction = macro
        b.PrintObject = False
        b.Name = "btn_" & macro
    End If
    If b Is Nothing Then
        Err.Clear
        Set b = ws.Shapes.AddShape(5, x, y, w, h)
        b.TextFrame.Characters.Text = caption
        b.OnAction = macro
    End If
    On Error GoTo 0
    Set AddButton = b
End Function

Public Sub Info(ByVal msg As String)
    If gQuiet Then
        gDbg = gDbg & " [Info] " & msg
    Else
        MsgBox msg, vbInformation
    End If
End Sub

Public Function Confirm(ByVal msg As String) As Boolean
    If gQuiet Then
        gDbg = gDbg & " [Confirm] " & msg
        Confirm = True
    Else
        Confirm = (MsgBox(msg, vbYesNo + vbQuestion) = vbYes)
    End If
End Function

Public Function Ask(ByVal msg As String, Optional ByVal def As String = "") As String
    If gQuiet Then
        gDbg = gDbg & " [Ask] " & Left$(msg, 60) & " -> " & def
        Ask = def
    Else
        Ask = InputBox(msg, , def)
    End If
End Function

Public Sub SetScreen(ByVal on_ As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = on_
    On Error GoTo 0
End Sub

Public Sub BeginWork()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo 0
End Sub

Public Sub EndWork()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Public Sub PageSetupA4(ByVal ws As Object, ByVal landscape As Boolean, ByVal marginMm As Double, ByVal lastR As Long, ByVal lastC As Long)
    On Error Resume Next
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Address
    With ws.PageSetup
        .PaperSize = XL_PAPER_A4
        If landscape Then
            .Orientation = XL_LANDSCAPE
        Else
            .Orientation = XL_PORTRAIT
        End If
        .LeftMargin = marginMm * PT_MM
        .RightMargin = marginMm * PT_MM
        .TopMargin = marginMm * PT_MM
        .BottomMargin = marginMm * PT_MM
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHeader = ""
        .CenterFooter = ""
        .LeftHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .RightFooter = ""
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintGridlines = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    On Error GoTo 0
End Sub

Public Sub HideGrid(ByVal ws As Object)
    On Error Resume Next
    Dim cur As Object
    Set cur = ActiveSheet
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    If Not cur Is Nothing Then cur.Activate
    On Error GoTo 0
End Sub

Public Function SafeSheetName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", "?", "*", ":", "[", "]")
        s = Replace(s, CStr(ch), "_")
    Next
    s = TrimX(s)
    If Len(s) > 31 Then s = TrimX(Left$(s, 31))
    SafeSheetName = s
End Function

' debugging aid: first n elements as text
Public Function ElDump(ByVal n As Long) As String
    Dim i As Long, e As Variant, s As String
    s = "elN=" & elN
    For i = 1 To elN
        If i > n Then Exit For
        e = elV(i)
        s = s & " | " & e(EK) & "," & Round(e(EX), 1) & "," & Round(e(EY), 1) & "," & Round(e(EW), 1) & "," & Round(e(EH), 1) & "," & Left$(e(ETXT), 8) & ",sz=" & e(ESIZE) & ",b=" & e(EBOLD) & ",c=" & e(ECOLOR)
    Next
    ElDump = s
End Function
