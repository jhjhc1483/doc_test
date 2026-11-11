Attribute VB_Name = "HancellDashboard"
' 임원 일일 예정사항 - 한셀·엑셀 공통 (셀로만 그림, A4 한 장)
' ★ 표시 줄: 한셀 '라이브러리가 등록되지 않았습니다' 오류를 피하려고 형식을 밝혀 쓴 곳
' 1) Setup 한 번 실행  2) 설정 시트에 직위·성명, 입력 시트에 일정  3) MakeDashboard 실행
Option Explicit

Const DASH As String = "대시보드"
Const CFG As String = "설정"
Const INP As String = "입력"

Dim ln(1 To 8, 1 To 80) As String, kd(1 To 8, 1 To 80) As Long, tl(1 To 8, 1 To 80) As Long
Dim n(1 To 8) As Long, fails As String

Sub Setup()
    Dim ws As Worksheet, i As Long, a As Variant   ' ★
    Set ws = GetSheet(CFG)
    a = Split("제목,날짜,왼쪽 위 문구,용지 방향 (세로/가로),글자크기 최대,글자크기 최소,여백 (mm),줄 높이 보정 (글이 잘리면 크게)", ",")
    For i = 0 To 7
        ws.Range("A" & (i + 1)).Value = a(i)   ' ★
    Next
    ws.Range("B1").Value = "임원 일일 예정사항"
    ws.Range("B2").Value = Date
    ws.Range("B4").Value = "세로"
    ws.Range("B5").Value = 12
    ws.Range("B6").Value = 7
    ws.Range("B7").Value = 10
    ws.Range("B8").Value = 1.7
    a = Split("윗줄 왼쪽,윗줄 가운데,윗줄 오른쪽,가운데 큰 칸,가운데 오른쪽 위,가운데 오른쪽 아래,아랫줄 왼쪽,아랫줄 오른쪽", ",")
    ws.Range("A:A").ColumnWidth = 28   ' ★
    ws.Range("A9").Value = "자리"
    ws.Range("B9").Value = "직위"
    ws.Range("C9").Value = "성명"
    ws.Range("D9").Value = "위치"
    For i = 1 To 8
        ws.Range("A" & (9 + i)).Value = i   ' ★
        ws.Range("D" & (9 + i)).Value = a(i - 1)   ' ★
    Next
    Set ws = GetSheet(INP)
    a = Split("자리 (1~8),시간,내용,비고 (작은 글씨 *),※ 참고", ",")
    For i = 0 To 4
        ws.Range(Chr(65 + i) & "1").Value = a(i)   ' ★
    Next
    GetSheet DASH
    MsgBox "설정 시트에 직위·성명을, 입력 시트에 일정을 쓰고 MakeDashboard 를 실행하세요."
End Sub

Sub MakeDashboard()
    Dim c As Worksheet, wi As Worksheet, d As Worksheet, r As Long, k As Long, i As Long, j As Long, p As Long   ' ★
    Dim a As Variant, s As String, tm As String, first As String, st As Long, rf(1 To 8) As String
    Dim fs As Double, fMax As Double, fMin As Double, mg As Double, pw As Double, ph As Double
    Dim cw As Double, hh As Double, ratio As Double, e(1 To 8) As Double, tot As Double, ex As Double
    Dim r5 As Double, r2 As Double, r8 As Double, r13 As Double, dt As Date, lf As Double
    Dim hdr As Variant, bdy As Variant, nc As Variant

    Set c = GetSheet(CFG)
    Set wi = GetSheet(INP)
    Set d = GetSheet(DASH)
    Erase ln, kd, tl, n

    ' 입력 시트 → 칸마다 줄
    For r = 2 To 500
        k = Val(Txt(wi.Range("A" & r).Value))   ' ★
        If k >= 1 And k <= 8 Then
            tm = TimeText(wi.Range("B" & r).Value)   ' ★
            a = Split(Txt(wi.Range("C" & r).Value), vbLf)   ' ★
            first = ""
            st = 0
            If UBound(a) >= 0 Then
                first = Trim(a(0))
                st = 1
                If Left(first, 1) = "*" Or Left(first, 1) = "※" Then
                    first = ""
                    st = 0
                End If
            End If
            If tm <> "" Or first <> "" Then
                If tm <> "" And first <> "" Then tm = tm & "  "
                AddLn k, tm & Place(first), 0, Len(Trim(tm))
            End If
            For i = st To UBound(a)
                s = Trim(a(i))
                If Left(s, 1) = "※" Then
                    rf(k) = rf(k) & vbLf & "※ " & Trim(Mid(s, 2))
                ElseIf Left(s, 1) = "*" Or Left(s, 1) = "-" Then
                    AddLn k, "   " & Left(s, 1) & " " & Trim(Mid(s, 2)), 1, 0
                ElseIf s <> "" Then
                    AddLn k, "       " & Place(s), 0, 0
                End If
            Next
            a = Split(Txt(wi.Range("D" & r).Value), vbLf)   ' ★
            For i = 0 To UBound(a)
                If Trim(a(i)) <> "" Then AddLn k, "   * " & Trim(a(i)), 1, 0
            Next
            a = Split(Txt(wi.Range("E" & r).Value), vbLf)   ' ★
            For i = 0 To UBound(a)
                s = Trim(a(i))
                If Left(s, 1) = "※" Then s = Trim(Mid(s, 2))
                If s <> "" Then rf(k) = rf(k) & vbLf & "※ " & s
            Next
        End If
    Next
    For k = 1 To 8
        If rf(k) <> "" Then
            If n(k) > 0 Then AddLn k, "", 2, 0
            a = Split(Mid(rf(k), 2), vbLf)
            For i = 0 To UBound(a)
                AddLn k, a(i), 2, 0
            Next
        End If
        If n(k) = 0 Then AddLn k, "일정 없음", 3, 0
    Next

    ' 설정
    fMax = Val(Txt(c.Range("B5").Value))
    If fMax < 6 Then fMax = 12
    fMin = Val(Txt(c.Range("B6").Value))
    If fMin < 5 Or fMin > fMax Then fMin = 7
    mg = Val(Txt(c.Range("B7").Value))
    If mg < 3 Then mg = 10
    lf = Val(Txt(c.Range("B8").Value))
    If lf < 1 Then lf = 1.7
    If InStr(Txt(c.Range("B4").Value), "가로") > 0 Then
        pw = 297: ph = 210
    Else
        pw = 210: ph = 297
    End If
    pw = (pw - 2 * mg) * 2.835 * 0.97
    ph = (ph - 2 * mg) * 2.835 * 0.96
    cw = pw / 12

    ' 글자크기: 최대부터 0.5pt씩 줄이며 A4 한 장에 들어가는 크기 찾기
    nc = Split("4,4,4,8,4,4,6,6", ",")
    fs = fMax
    Do
        hh = fs * 1.8
        For k = 1 To 8
            e(k) = EstH(k, fs, nc(k - 1) * cw - 6, lf)
        Next
        r5 = Mx(Mx(e(1), e(2)), e(3))
        r2 = Mx(e(5) + e(6), e(4) - hh)
        r13 = Mx(e(7), e(8))
        tot = 30 + 18 + 4 + 12 + 4 * hh + r5 + r2 + r13
        If tot <= ph Or fs <= fMin Then Exit Do
        fs = fs - 0.5
    Loop
    ex = ph - tot
    If ex < 0 Then ex = 0
    r5 = r5 + ex / 4
    r13 = r13 + ex / 4
    r2 = r2 + ex / 2
    r8 = e(5) + (r2 - e(5) - e(6)) / 2

    ' 그리기 (한셀에서 안 되는 기능은 건너뛰고 끝에 알려 줌)
    On Error Resume Next
    fails = ""
    d.Range("A1:L30").UnMerge
    d.Range("A1:L30").Clear
    Chk "지우기"
    d.Range("A:A").ColumnWidth = 10   ' ★
    ratio = d.Range("A:A").Width / 10   ' ★
    If Err.Number <> 0 Or ratio <= 0 Then ratio = 5.6
    Err.Clear
    For i = 1 To 12
        d.Range(Chr(64 + i) & ":" & Chr(64 + i)).ColumnWidth = cw / ratio   ' ★
    Next
    Chk "열너비"
    a = Array(30, 18, 4, hh, r5, 6, hh, r8, hh, r2 - r8, 6, hh, r13)
    For i = 0 To 12
        If a(i) > 409 Then a(i) = 409
        d.Range((i + 1) & ":" & (i + 1)).RowHeight = a(i)   ' ★
    Next
    Chk "행높이"

    dt = Date
    If IsDate(c.Range("B2").Value) Then dt = CDate(c.Range("B2").Value)
    s = Year(dt) & ". " & Month(dt) & ". " & Day(dt) & ".(" & Mid("일월화수목금토", Weekday(dt), 1) & ")"
    PutBox d.Range("A1:L1"), Txt(c.Range("B1").Value), 18, True, -4108, -1, False
    PutBox d.Range("A2:F2"), Txt(c.Range("B3").Value), 10, False, -4131, -1, False
    PutBox d.Range("G2:L2"), s, 10, True, -4152, -1, False
    d.Range("A2:L2").Borders(9).LineStyle = 1
    d.Range("A2:L2").Borders(9).Weight = -4138
    Chk "제목"

    hdr = Split("A4:D4,E4:H4,I4:L4,A7:H7,I7:L7,I9:L9,A12:F12,G12:L12", ",")
    bdy = Split("A5:D5,E5:H5,I5:L5,A8:H10,I8:L8,I10:L10,A13:F13,G13:L13", ",")
    For k = 1 To 8
        s = Trim(Txt(c.Range("B" & (9 + k)).Value) & " " & Txt(c.Range("C" & (9 + k)).Value))   ' ★
        PutBox d.Range(hdr(k - 1)), s, fs * 1.05, True, -4108, RGB(231, 233, 236), True
        s = ""
        For j = 1 To n(k)
            If j > 1 Then s = s & vbLf
            s = s & ln(k, j)
        Next
        PutBox d.Range(bdy(k - 1)), s, fs, False, -4131, -1, True
        ' 시간 굵게, * 줄·※ 줄 작게 (안 되면 그냥 넘어감)
        p = 1
        For j = 1 To n(k)
            If tl(k, j) > 0 Then d.Range(bdy(k - 1)).Resize(1, 1).Characters(p + 1, tl(k, j)).Font.Bold = True   ' ★
            If kd(k, j) >= 1 Then d.Range(bdy(k - 1)).Resize(1, 1).Characters(p, Len(ln(k, j))).Font.Size = fs * 0.88   ' ★
            p = p + Len(ln(k, j)) + 1
        Next
        Err.Clear
    Next

    d.Activate
    ActiveWindow.DisplayGridlines = False
    Err.Clear
    d.PageSetup.PrintArea = "$A$1:$L$13"
    d.PageSetup.PaperSize = 9
    If pw > ph Then d.PageSetup.Orientation = 2 Else d.PageSetup.Orientation = 1
    d.PageSetup.LeftMargin = mg * 2.835
    d.PageSetup.RightMargin = mg * 2.835
    d.PageSetup.TopMargin = mg * 2.835
    d.PageSetup.BottomMargin = mg * 2.835
    d.PageSetup.CenterHorizontally = True
    d.PageSetup.CenterHeader = ""
    d.PageSetup.CenterFooter = ""
    Chk "쪽설정"
    d.PageSetup.Zoom = False
    d.PageSetup.FitToPagesWide = 1
    d.PageSetup.FitToPagesTall = 1
    Chk "한쪽맞춤"
    On Error GoTo 0

    s = "대시보드를 그렸습니다. 글자 " & fs & "pt"
    If tot > ph Then s = s & vbLf & "최소 글자크기로도 넘칩니다. 내용을 줄이거나 최소 글자크기를 낮추세요."
    If fails <> "" Then s = s & vbLf & vbLf & "이 프로그램에서 안 된 것: " & fails
    If InStr(fails, "쪽설정") > 0 Or InStr(fails, "한쪽맞춤") > 0 Then
        s = s & vbLf & "인쇄할 때 쪽 설정에서 A4 · 여백 " & mg & "mm · 한 쪽에 맞추기를 직접 고르세요."
    End If
    MsgBox s
End Sub

' 칸 하나 그리기: 서식을 먼저 입히고 병합한 뒤 값을 넣음 (fill < 0 이면 배경 없음)
Sub PutBox(ByVal rg As Range, ByVal v As String, ByVal sz As Double, ByVal bd As Boolean, _
           ByVal ha As Long, ByVal fill As Long, ByVal edge As Boolean)   ' ★
    Dim b As Long
    On Error Resume Next
    rg.Font.Name = "맑은 고딕"
    rg.Font.Size = sz
    rg.Font.Bold = bd
    rg.HorizontalAlignment = ha
    If edge And Not bd Then rg.VerticalAlignment = -4160 Else rg.VerticalAlignment = -4108
    rg.WrapText = True
    If fill >= 0 Then rg.Interior.Color = fill
    Chk "글꼴"
    rg.Merge
    Chk "병합"
    rg.Resize(1, 1).Value = v   ' ★
    Chk "값"
    If edge Then
        For b = 7 To 10
            rg.Borders(b).LineStyle = 1
            rg.Borders(b).Weight = 2
        Next
        Chk "테두리"
    End If
End Sub

Sub AddLn(ByVal k As Long, ByVal s As String, ByVal kind As Long, ByVal tlen As Long)
    If n(k) >= 80 Then Exit Sub
    n(k) = n(k) + 1
    ln(k, n(k)) = " " & s
    kd(k, n(k)) = kind
    tl(k, n(k)) = tlen
End Sub

Sub Chk(ByVal t As String)
    If Err.Number <> 0 Then
        If InStr(fails, t) = 0 Then fails = fails & t & " "
        Err.Clear
    End If
End Sub

Function GetSheet(ByVal nm As String) As Worksheet   ' ★
    Dim ws As Worksheet, w As Worksheet, i As Long   ' ★
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set w = ActiveWorkbook.Worksheets(i)   ' ★
        If w.Name = nm Then Set ws = w   ' ★
    Next
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = nm
    End If
    Set GetSheet = ws
End Function

Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Txt = Trim(Replace(CStr(v), vbCr, ""))
End Function

' 엑셀·한셀 시간 값(9:30)은 09:30 글자로
Function TimeText(ByVal v As Variant) As String
    Dim m As Long
    If VarType(v) = vbDate Or VarType(v) = vbDouble Then
        m = CLng((CDbl(v) - Int(CDbl(v))) * 1440)
        TimeText = Format(m \ 60, "00") & ":" & Format(m Mod 60, "00")
    Else
        TimeText = Txt(v)
    End If
End Function

' "회의 @대회의실" → "회의 (대회의실)"
Function Place(ByVal s As String) As String
    Dim p As Long
    p = InStrRev(s, "@")
    If p = 1 And Trim(Mid(s, 2)) <> "" Then
        s = "(" & Trim(Mid(s, 2)) & ")"
    ElseIf p > 1 And Trim(Mid(s, p + 1)) <> "" Then
        If Mid(s, p - 1, 1) = " " Then s = RTrim(Left(s, p - 1)) & " (" & Trim(Mid(s, p + 1)) & ")"
    End If
    Place = s
End Function

' 칸 하나의 글 높이 어림 (글자크기 fs, 칸 폭 w pt)
Function EstH(ByVal k As Long, ByVal fs As Double, ByVal w As Double, ByVal lf As Double) As Double
    Dim j As Long, q As Double, L As Double
    For j = 1 To n(k)
        q = -Int(-(Units(ln(k, j)) * fs * 1.1) / w)
        If q < 1 Then q = 1
        L = L + q
    Next
    EstH = L * fs * lf + 6
End Function

' 글자 폭 어림 (한글 1, 영문·숫자 0.58, 빈칸 0.32)
Function Units(ByVal s As String) As Double
    Dim i As Long, cd As Long
    For i = 1 To Len(s)
        cd = AscW(Mid(s, i, 1))
        If cd < 0 Then cd = cd + 65536
        Select Case cd
            Case Is >= 4352: Units = Units + 1
            Case 32: Units = Units + 0.32
            Case Else: Units = Units + 0.58
        End Select
    Next
End Function

Function Mx(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Mx = a Else Mx = b
End Function
