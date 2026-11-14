Attribute VB_Name = "HcDailySetup"
Option Explicit
' ===========================================================================
' HcDailySetup - builds every sheet of 부서별 일일 예정사항 (run once)
' ===========================================================================

Public Sub Setup()
    Dim ws As Object
    Set ws = SheetByName(SH_IN)
    If Not ws Is Nothing Then
        If Not Confirm("이미 만들어진 파일입니다. 처음 상태(예시 내용)로 다시 만들까요? 지금 내용은 지워집니다.") Then Exit Sub
    End If
    BeginWork
    SetupAll
    EndWork
    Info "부서별 일일 예정사항 파일을 만들었습니다. [다른 이름으로 저장]에서 '매크로 사용 문서(.xlsm)'로 저장하세요."
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

Private Sub Note9(ByVal ws As Object, ByVal addr As String, ByVal txt As String, Optional ByVal h As Double = 0)
    With ws.Range(addr)
        If .Cells.Count > 1 Then .Merge
        .Cells(1, 1).Value = txt
        .WrapText = True
        .Font.Size = 9
        .Font.Color = COL_GRAY
        .VerticalAlignment = XL_TOP
    End With
    If h > 0 Then ws.Range(addr).Rows(1).RowHeight = h
End Sub

Private Sub SheetBase(ByVal ws As Object)
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 10
    HideGrid ws
End Sub

Public Sub SetupAll()
    Dim wb As Object, ws As Object, k As Long, a4 As Object, prev As Object, samp As Variant
    Set wb = HostBook()
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Styles("Normal").Font.Name = "맑은 고딕"
    wb.Styles("Normal").Font.Size = 10
    On Error GoTo 0
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
    ws.Name = SH_IN
    Set a4 = wb.Worksheets.Add(, ws)
    a4.Name = SH_A4
    Set prev = a4
    For k = 1 To 6
        gStep = "dept sheet " & k
        samp = SampleDept(k)
        Set ws = wb.Worksheets.Add(, prev)
        SetupDeptSheet ws, samp(0), samp(1)
        ws.Name = SafeSheetName(samp(0))
        Set prev = ws
    Next
    gStep = "format"
    Set ws = wb.Worksheets.Add(, prev): ws.Name = SH_FMT: SetupFormat ws: Set prev = ws
    gStep = "help"
    Set ws = wb.Worksheets.Add(, prev): ws.Name = SH_HELP: SetupHelp ws: Set prev = ws
    gStep = "paste"
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

' sample departments: rows "시간|내용|장소|비고|설명|링크" (^ = new line inside a cell)
Private Function SampleDept(ByVal k As Long) As Variant
    Select Case k
        Case 1: SampleDept = Array("기획조정과", Array( _
            "|>> 담당: 기획조정과 ○○○ 주무관 (내선 1234)||||", _
            "09:00|주간 업무보고 회의|본관 3층 대회의실|참석: 실장, 각 과장|안건: 2027년 예산안 검토, 하반기 인사 발령^자료: http://intranet.example/board/123|", _
            "14:00~15:30|2027년도 예산안 검토 회의 (소회의실)||||", _
            "|상반기 업무계획 실적 취합 **(17:00 마감)**||||"))
        Case 2: SampleDept = Array("총무과", Array( _
            "10:00|청사 소방 안전점검 (전 층)||||http://intranet.example/notice/77", _
            "13:30|신규 직원 임용장 수여식|접견실|대상: 신규 임용자 [[3명]]||", _
            "|관용차량 정기점검: 1호차, 3호차||||"))
        Case 3: SampleDept = Array("인사과", Array( _
            "|○ 하반기 승진심사 서류 접수 마감||||", _
            "| - 대상: 6급 이하 승진 후보자||||", _
            "| - 제출처: 인사과 담당자||||", _
            "|○ 15:00 교육훈련 수요조사 설명회 (2층 교육장)||||"))
        Case 4: SampleDept = Array("재무과", Array( _
            "|||||", _
            "시간|내용|장소|비고||", _
            "10:00|계약심의위원회|소회의실|3건 심의||", _
            "16:00|월말 결산 점검|재무과|||"))
        Case 5: SampleDept = Array("정보화담당관", Array( _
            "18:00~20:00|!!내부망 서버 정기 점검!! (업무 시스템 일시 중단)|||점검 대상: 전자결재, 내부 메일, 인트라넷 게시판^문의: 정보화담당관 (내선 5678)|", _
            "|보안 취약점 점검 결과 보고||||", _
            "|> 점검 시간에는 전자결재에 접속할 수 없습니다.||||"))
        Case 6: SampleDept = Array("감사담당관", Empty)
    End Select
End Function

Public Sub SetupDeptSheet(ByVal ws As Object, ByVal nm As String, ByVal rows As Variant)
    Dim i As Long, f As Variant, c As Long, cur As Object
    gStep = "dept sheet clear"
    ClearSheet ws
    SheetBase ws
    Dim w As Variant
    w = Array(13, 52, 20, 22, 34, 30)
    For c = 0 To 5
        ws.Columns(c + 1).ColumnWidth = w(c)
    Next
    gStep = "dept sheet labels"
    ws.Range("A1").Value = "부서"
    ws.Range("C1").Value = "인쇄"
    ws.Range("E1").Value = "상태"
    For Each f In Array("A1", "C1", "E1")
        ws.Range(f).Font.Bold = True
        ws.Range(f).Interior.Color = COL_HEAD
        ws.Range(f).HorizontalAlignment = XL_CENTER
    Next
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value = nm
    ws.Range("B1").Font.Bold = True
    ws.Range("D1").Value = "예"
    ListValid ws, "D1", "예,아니오"
    ws.Range("F1").NumberFormat = "@"
    ws.Range("F1").Font.Size = 9
    ws.Range("F1").Font.Color = COL_GRAY
    ws.Range("A1:F1").Borders.LineStyle = XL_CONT
    ws.Range("A1:F1").Borders.Color = COL_INLINE
    ws.Rows(1).RowHeight = 22
    With ws.Range("A2:F2")
        .Value = Array("시간", "내용", "장소", "비고", "설명 (마우스를 올리면 보임)", "링크")
        .Font.Bold = True
        .Interior.Color = COL_HEAD
        .HorizontalAlignment = XL_CENTER
        .Borders.LineStyle = XL_CONT
        .Borders.Color = COL_INLINE
    End With
    With ws.Range("A3:F300")
        .NumberFormat = "@"
        .WrapText = True
        .VerticalAlignment = XL_TOP
    End With
    ws.Range("A3:A300").HorizontalAlignment = XL_CENTER
    gStep = "dept sheet rows"
    If IsArray(rows) Then
        For i = 0 To UBound(rows)
            f = Split(CStr(rows(i)), "|")
            For c = 0 To UBound(f)
                If c <= 5 And CStr(f(c)) <> "" Then ws.Cells(3 + i, c + 1).Value = Replace(CStr(f(c)), "^", vbLf)
            Next
        Next
    End If
    On Error Resume Next
    Set cur = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
    ws.Range("B3").Select
    If Not cur Is Nothing Then cur.Activate
    On Error GoTo 0
    gStep = "dept sheet buttons"
    AddButton ws, ws.Columns(7).Left + 12, 4, 96, 24, "A4 만들기", "BuildA4"
    AddButton ws, ws.Columns(7).Left + 12, 32, 96, 24, "형광펜", "HighlightSel"
    AddButton ws, ws.Columns(7).Left + 12, 60, 96, 24, "입력으로", "GoInput"
End Sub

Private Sub SetupInput(ByVal ws As Object)
    Dim c As Long, steps As Variant, i As Long, x As Double, y As Double, bw As Double
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 2
    For c = 2 To 13
        ws.Columns(c).ColumnWidth = 7
    Next
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(14).ColumnWidth = 2
    Lbl ws, "B1:I1", "부서별 일일 예정사항", True, 15
    Lbl ws, "J1:M1", "A4 한 장 자동 맞춤", False, 9, COL_GRAY
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 90
    ws.Range("B3:M3").Merge
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").WrapText = True
    ws.Rows(3).RowHeight = 30
    DefName "d_status", ws, "$B$3"
    Lbl ws, "B5:M5", "문서 정보", True, 11
    Lbl ws, "B6", "제목", False, 9, COL_GRAY: Inp ws, "C6:M6", "부서별 일일 예정사항", "d_title"
    Lbl ws, "B7", "날짜", False, 9, COL_GRAY: Inp ws, "C7:D7", Date, "d_date", False
    ws.Range("C7").NumberFormat = "yyyy-mm-dd"
    ws.Range("C7").HorizontalAlignment = XL_LEFT
    Lbl ws, "E7:F7", "작성 부서·부제", False, 9, COL_GRAY: Inp ws, "G7:M7", "○○○○실", "d_sub"
    Lbl ws, "B8", "하단 문구", False, 9, COL_GRAY: Inp ws, "C8:M8", "※ 일정이 바뀌면 기획조정과(내선 1234)로 알려 주시기 바랍니다.", "d_foot"
    ws.Range("C6:M8").Rows.RowHeight = 20
    Lbl ws, "B10:M10", "매일 하는 일", True, 11
    steps = Array("1. 처음 한 번 [공동편집용 엑셀 양식 만들기]로 부서마다 시트가 있는 .xlsx를 만들어 두레이에 올립니다.", _
        "2. 각 부서는 자기 시트에 한 줄에 일정 하나씩 씁니다(시간 / 내용 / 장소 / 비고 / 설명 / 링크).", _
        "3. 파일을 내려받아 [엑셀 불러오기]. 직접 고칠 때는 아래 목록에서 부서를 눌러 그 시트에 씁니다.", _
        "4. [A4 만들기] → [인쇄 미리보기] 또는 [PDF로 저장]. 단추가 안 보이면 Alt+F8에서 BuildA4 등을 실행하세요.")
    For i = 0 To UBound(steps)
        ws.Range(ws.Cells(11 + i, 2), ws.Cells(11 + i, 13)).Merge
        ws.Cells(11 + i, 2).Value = steps(i)
        ws.Cells(11 + i, 2).Font.Size = 9
        ws.Cells(11 + i, 2).WrapText = True
        ws.Rows(11 + i).RowHeight = 24
    Next
    Lbl ws, "B16:M16", "부서별 예정사항 (부서 시트의 탭 순서 = 인쇄 순서)", True, 11
    ws.Range("B17:M17").Merge
    ws.Range("B17").Font.Bold = True
    DefName "d_recv", ws, "$B$17"
    ws.Range("B18:M18").Merge
    ws.Range("B18").WrapText = True
    ws.Range("B18").Font.Size = 9
    ws.Range("B18").Font.Color = COL_BAD
    ws.Rows(18).RowHeight = 28
    DefName "d_miss", ws, "$B$18"
    With ws.Range("B19:F19")
        .Value = Array("번호", "부서 (누르면 그 시트로)", "인쇄", "줄", "상태")
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = COL_HEAD
    End With
    DefName "d_list", ws, "$B$20"
    y = ws.Rows(2).Top + 4
    x = ws.Columns(2).Left
    bw = 102
    AddButton ws, x, y, bw, 24, "A4 만들기", "BuildA4"
    AddButton ws, x + (bw + 6), y, bw, 24, "인쇄 미리보기", "PrintPreviewA4"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, "PDF로 저장", "SavePdf"
    AddButton ws, x + 3 * (bw + 6), y, bw, 24, "엑셀 불러오기", "ImportExcel"
    y = y + 28
    AddButton ws, x, y, bw, 24, "붙여넣은 표 가져오기", "ImportPasted"
    AddButton ws, x + (bw + 6), y, bw, 24, "되돌리기", "UndoLast"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, "새 문서", "NewDoc"
    AddButton ws, x + 3 * (bw + 6), y, bw, 24, "형광펜", "HighlightSel"
    y = y + 28
    AddButton ws, x, y, bw, 24, "부서 추가", "AddDept"
    AddButton ws, x + (bw + 6), y, bw, 24, "부서 빼기", "RemoveDept"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, "순서 바꾸기", "SwapDepts"
    AddButton ws, x + 3 * (bw + 6), y, bw, 24, "공동편집 양식 만들기", "MakeTemplate"
End Sub

Private Sub SetupA4(ByVal ws As Object)
    Dim tb As Object
    ClearSheet ws
    HideGrid ws
    On Error Resume Next
    Set tb = ws.Shapes.AddTextbox(1, 600, 6, 260, 70)
    If Not tb Is Nothing Then
        tb.Name = "a4status"
        tb.TextFrame.Characters.Text = ""
        tb.TextFrame.Characters.Font.Size = 9
        tb.Line.Visible = False
    End If
    On Error GoTo 0
    AddButton ws, 600, 90, 120, 26, "다시 만들기", "BuildA4"
    AddButton ws, 600, 122, 120, 26, "인쇄 미리보기", "PrintPreviewA4"
    AddButton ws, 600, 154, 120, 26, "PDF로 저장", "SavePdf"
    AddButton ws, 600, 186, 120, 26, "입력으로", "GoInput"
End Sub

Private Sub SetupFormat(ByVal ws As Object)
    Dim r As Long
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 24
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 14
    ws.Columns(6).ColumnWidth = 40
    Lbl ws, "B1:F1", "서식", True, 14
    ws.Rows(1).RowHeight = 28
    r = 3
    Lbl ws, "B" & r & ":F" & r, "배치", True, 11: r = r + 1
    Lbl ws, "B" & r, "레이아웃": Inp ws, "C" & r & ":D" & r, "표", "s_layout": ListValid ws, "C" & r, "표,표 2단,제목형,제목형 2단": r = r + 1
    Lbl ws, "B" & r, "예정사항 칸": Inp ws, "C" & r & ":D" & r, "좌우 2칸", "s_split": ListValid ws, "C" & r, "좌우 2칸,한 칸"
    Note9 ws, "E" & r & ":F" & r, "좌우 2칸: 일정을 알아서 나눠 담음 (일정이 하나뿐이거나 열 5개 이상 표가 있는 부서, 2단 레이아웃은 한 칸)", 30: r = r + 1
    Lbl ws, "B" & r, "용지 방향": Inp ws, "C" & r & ":D" & r, "A4 세로", "s_orient": ListValid ws, "C" & r, "A4 세로,A4 가로": r = r + 1
    Lbl ws, "B" & r, "여백 (mm, 3~30)": Inp ws, "C" & r, 12, "s_margin", False: r = r + 1
    Lbl ws, "B" & r, "글꼴": Inp ws, "C" & r & ":D" & r, "맑은 고딕", "s_font"
    ListValid ws, "C" & r, "맑은 고딕,나눔고딕,나눔스퀘어,함초롬돋움,굴림,돋움,Noto Sans KR,함초롬바탕,바탕,나눔명조": r = r + 2
    Lbl ws, "B" & r & ":F" & r, "크기·간격", True, 11: r = r + 1
    Lbl ws, "B" & r, "자동 맞춤": Inp ws, "C" & r, "예", "s_auto": ListValid ws, "C" & r, "예,아니오"
    Note9 ws, "D" & r & ":F" & r, "한 장에 꽉 차도록 글자크기·간격을 함께 조절 (0.5pt 단위)": r = r + 1
    Lbl ws, "B" & r, "글자크기 범위 (pt)": Inp ws, "C" & r, 7, "s_min", False: Inp ws, "D" & r, 13, "s_max", False
    Note9 ws, "E" & r & ":F" & r, "최소 ~ 최대. * 하위 항목은 2pt 작게": r = r + 1
    Lbl ws, "B" & r, "수동: 글자크기 (pt)": Inp ws, "C" & r, 10.5, "s_mfont", False
    Note9 ws, "D" & r & ":F" & r, "자동 맞춤이 '아니오'일 때만": r = r + 1
    Lbl ws, "B" & r, "수동: 항목 간격 (em)": Inp ws, "C" & r, 0.25, "s_mgap", False: r = r + 2
    Lbl ws, "B" & r & ":F" & r, "표시", True, 11: r = r + 1
    Lbl ws, "B" & r, "글머리 기호": Inp ws, "C" & r & ":E" & r, "보고서식 ( ○ → - → · )", "s_bullet"
    ListValid ws, "C" & r, "보고서식 ( ○ → - → · ),보고서식 ( □ → ○ → - → · ),일반 ( • → ◦ → ▪ ),입력한 기호 그대로": r = r + 1
    Lbl ws, "B" & r, "시간 굵게": Inp ws, "C" & r, "예", "s_timebold": ListValid ws, "C" & r, "예,아니오": r = r + 1
    Lbl ws, "B" & r, "표를 페이지 끝까지 늘리기": Inp ws, "C" & r, "예", "s_fill": ListValid ws, "C" & r, "예,아니오": r = r + 1
    Lbl ws, "B" & r, "표 머리글": Inp ws, "C" & r, "예", "s_colhead": ListValid ws, "C" & r, "예,아니오"
    Inp ws, "D" & r, "부서", "s_colh1": Inp ws, "E" & r & ":F" & r, "주요 예정사항", "s_colh2": r = r + 1
    Lbl ws, "B" & r, "내용 없는 부서": Inp ws, "C" & r, "문구 표시", "s_emptymode": ListValid ws, "C" & r, "문구 표시,숨기기"
    Inp ws, "D" & r & ":F" & r, "특이사항 없음", "s_empty": r = r + 2
    Lbl ws, "B" & r & ":F" & r, "문서 암호", True, 11: r = r + 1
    Note9 ws, "B" & r & ":F" & r, "이 파일을 열 때 암호를 묻게 하려면 한셀의 문서 암호를 쓰세요(내용이 암호화됩니다). 아래 단추로 정하거나 한셀 메뉴의 문서 암호에서 정합니다. 암호를 잊으면 열 수 없습니다. 두레이에 올리는 공동편집 파일에는 걸지 마세요.", 40
    r = r + 1
    ws.Rows(r).RowHeight = 34
    AddButton ws, ws.Columns(2).Left, ws.Rows(r).Top + 5, 140, 24, "문서 암호 정하기", "SetDocPassword"
    AddButton ws, ws.Columns(2).Left + 146, ws.Rows(r).Top + 5, 118, 24, "A4 만들기", "BuildA4"
End Sub

Private Sub SetupPaste(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Range("A1:H1").Merge
    ws.Range("A1").Value = "엑셀(두레이)에서 복사한 내용을 3번째 줄 A칸부터 붙여넣고 [붙여넣은 표 가져오기]를 누르세요. 여러 부서면 '부서' 열을 두거나, '# 부서명' 줄로 나누세요."
    ws.Range("A1").WrapText = True
    ws.Range("A1").Font.Size = 9
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 34
    AddButton ws, 4, ws.Rows(2).Top + 5, 140, 24, "붙여넣은 표 가져오기", "ImportPasted"
    AddButton ws, 150, ws.Rows(2).Top + 5, 96, 24, "입력으로", "GoInput"
End Sub

Private Sub SetupHelp(ByVal ws As Object)
    Dim rows As Collection, i As Long, f As Variant, k As Long, ln As String
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 46
    ws.Columns(2).ColumnWidth = 60
    Set rows = New Collection
    rows.Add "#부서 시트 — 한 줄에 일정 하나"
    rows.Add "엑셀 칸|인쇄물"
    rows.Add "시간 · 내용 · 장소|한 줄로: 09:00 (굵게) 주간회의 (대회의실)"
    rows.Add "내용 칸 안에서 Alt+Enter 후 - ...|하위 항목"
    rows.Add "내용 칸 안에서 * ..., 또는 내용이 * 로 시작하는 줄|글씨 2pt 작은 하위 항목 (* 로 시작하는 줄은 바로 위 일정에 붙음)"
    rows.Add "비고|작은 글씨 하위 항목"
    rows.Add "설명|화면에서 칸에 마우스를 올리면 보이는 메모 (인쇄 안 됨)"
    rows.Add "링크|누르면 열리는 링크 (주소는 인쇄 안 됨)"
    rows.Add "한 줄 비우고 A열부터 [표 제목] → 머리글 → 내용|표 그대로"
    rows.Add ""
    rows.Add "#내용 칸에 바로 쓰는 기호 (시간·장소 칸을 비웠을 때)"
    rows.Add "입력|결과"
    rows.Add "- 회의|○ 회의"
    rows.Add "앞에 빈칸 두고 - 참석: 과장|하위 항목(-)"
    rows.Add "* 참석: 실장 (위 항목 아래)|하위 항목, 글씨 2pt 작게"
    rows.Add "- 09:00 회의 @대회의실|○ 09:00 회의 (대회의실)"
    rows.Add "○, □, ※, · 직접 입력|입력한 기호 그대로"
    rows.Add "1. , 가. , ①|번호 그대로"
    rows.Add "**굵게**, __밑줄__, ~~취소~~, [[형광펜]], !!빨간 강조!!|글자 꾸밈 (형광펜은 그 줄의 칸 전체가 노랗게)"
    rows.Add "## 소제목, > 비고, ---|소제목, 비고, 구분선"
    rows.Add ">> 설명 (항목 바로 아래)|마우스를 올리면 보이는 메모 (인쇄 안 됨)"
    rows.Add "[글자](주소) 또는 항목 끝에 주소|누르면 열리는 링크 (주소는 인쇄 안 됨)"
    rows.Add ""
    rows.Add "#단추"
    rows.Add "A4 만들기|A4 시트에 한 장짜리 문서를 새로 만듭니다 (글자크기 자동 맞춤)"
    rows.Add "엑셀 불러오기|두레이 공동편집 파일(.xlsx·.csv·.xml)을 골라 부서별로 넣습니다. 부서는 시트 이름으로 찾고, 처음 보는 부서는 새 부서로 추가"
    rows.Add "붙여넣은 표 가져오기|붙여넣기 시트에 붙인 내용을 부서별로 넣습니다"
    rows.Add "되돌리기|마지막 [엑셀 불러오기]·[붙여넣은 표 가져오기]·[새 문서]·[부서 빼기] 전으로"
    rows.Add "새 문서|모든 부서 내용만 비움 (부서 목록·서식 그대로, 날짜는 오늘)"
    rows.Add "부서 추가 / 부서 빼기 / 순서 바꾸기|부서 시트를 만들고 지우고 맞바꿈 (시트 탭을 끌어 순서를 바꿔도 됨)"
    rows.Add "형광펜|고른 칸의 글자를 [[…]]로 감쌈 (다시 누르면 지움, Ctrl+Shift+H)"
    rows.Add "공동편집 양식 만들기|부서마다 시트가 있는 .xlsx — 두레이에 올려 함께 작성"
    rows.Add "PDF로 저장|A4 시트를 PDF로 (링크는 PDF에서도 눌림)"
    rows.Add ""
    rows.Add "#엑셀 불러오기 규칙"
    rows.Add "부서 찾기|시트 이름. 처음 보는 부서 시트도 머리글(시간·내용…)이 있으면 새 부서로 (Sheet1 같은 기본 이름은 제외)"
    rows.Add "한 시트에 여러 부서|맨 앞에 '부서' 열 (병합된 칸도 됨)"
    rows.Add "날짜 열|문서 날짜에 맞는 줄만. 대부분이 다른 날짜면 문서 날짜를 바꿀지 물어봄"
    rows.Add "빼는 시트|빈 시트, 작성방법 같은 안내 시트, 숨긴 시트"
    rows.Add ""
    rows.Add "#문서 암호"
    rows.Add "열 때 암호 묻기|서식 시트의 [문서 암호 정하기] 또는 한셀 메뉴의 문서 암호"
    rows.Add ""
    rows.Add "#단추가 안 보이거나 눌리지 않을 때"
    rows.Add "Alt+F8 (매크로 목록)|BuildA4, ImportExcel, ImportPasted, UndoLast, NewDoc, AddDept, RemoveDept, SwapDepts, HighlightSel, MakeTemplate, SavePdf, PrintPreviewA4, SetDocPassword"
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
            If f(0) = "입력" Or f(0) = "엑셀 칸" Then ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).Interior.Color = COL_HEAD
        End If
    Next
End Sub

' ---- tests (LibreOffice) -------------------------------------------------
Public Function TestBuild() As String
    gQuiet = True
    gDbg = ""
    On Error GoTo fail
    DoBuild
    gQuiet = False
    TestBuild = "fs=" & gFS & " status=" & NStr("d_status") & " | " & gDbg
    Exit Function
fail:
    gQuiet = False
    TestBuild = "ERROR " & Err.Number & ": " & Err.Description & " [" & gStep & "] [" & gTrace & "]"
End Function

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

' layouts and buttons without dialogs
Public Function TestAll() As String
    Dim r As String, lay As Variant, p As String, ws As Object
    gQuiet = True
    gDbg = ""
    On Error GoTo fail
    For Each lay In Array("표 2단", "제목형", "제목형 2단", "표")
        gStep = "layout " & lay
        NSet "s_layout", lay
        DoBuild
        r = r & lay & ": " & NStr("d_status") & "; "
    Next
    gStep = "landscape"
    NSet "s_orient", "A4 가로"
    DoBuild
    r = r & "가로: " & NStr("d_status") & "; "
    NSet "s_orient", "A4 세로"
    gStep = "manual"
    NSet "s_auto", "아니오"
    DoBuild
    r = r & "수동: " & NStr("d_status") & "; "
    NSet "s_auto", "예"
    gStep = "mark"
    r = r & "mark=" & DMarkLines("- 09:00 회의 @대회의실" & vbLf & "* 참석: 실장" & vbLf & "주간 업무") & "; "
    gStep = "swap"
    LoadDepts False
    SwapDeptsIJ 1, 2
    LoadDepts False
    r = r & "swap=" & dName(1) & "/" & dName(2) & "; "
    SwapDeptsIJ 1, 2
    gStep = "newdoc"
    NewDoc
    LoadDepts True
    r = r & "afterNew lines1=" & dLines(1) & "; "
    UndoLast
    LoadDepts True
    r = r & "afterUndo lines1=" & dLines(1) & "; "
    gStep = "paste"
    Set ws = HostBook().Worksheets(SH_PASTE)
    ws.Range("A3:C3").Value = Array("부서", "시간", "내용")
    ws.Range("A4:C4").Value = Array("감사담당관", "10:00", "붙여넣기 시험")
    ImportPasted
    LoadDepts True
    r = r & "pasted=" & dLines(6) & "; "
    UndoLast
    gStep = "template"
    p = Environ$("HC_TEST_OUT")
    If p <> "" Then r = r & "template=" & MakeTemplateTo(p) & "; "
    gStep = "build"
    DoBuild
    r = r & "final: " & NStr("d_status")
    gQuiet = False
    TestAll = r
    Exit Function
fail:
    gQuiet = False
    TestAll = r & " ERROR " & Err.Number & ": " & Err.Description & " [" & gStep & "] [" & gTrace & "] " & gDbg
End Function

' 이 모듈을 한셀이 읽었는지 확인 (실행했을 때 창이 뜨면 정상)
Public Sub CheckHcDailySetup()
    MsgBox "HcDailySetup OK"
End Sub
