Attribute VB_Name = "HcExecSetup"
Option Explicit

' ===========================================================================
' Setup: builds every sheet (run once in a new workbook)
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
    Info "임원 일일 예정사항 파일을 만들었습니다. [다른 이름으로 저장]에서 '매크로 사용 문서(.xlsm)'로 저장하세요."
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
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 10
    HideGrid ws
End Sub

Private Sub SetupAll()
    Dim wb As Object, ws As Object, k As Long
    Set wb = HostBook()
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Styles("Normal").Font.Name = "맑은 고딕"
    wb.Styles("Normal").Font.Size = 10
    On Error GoTo 0
    ' start from one sheet called 입력
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
        Case 1: SampleExec = Array("부사장", "이정훈", Array( _
            "09:00|인사위원회 @소회의실|하반기 승진 심사|16:00 이후 사내 대기", _
            "10:00|주간 경영회의 @대회의실||", "13:30|감사팀 업무보고||", "15:00|외부 감사인 미팅 @접견실||"))
        Case 2: SampleExec = Array("부사장", "박서연", Array( _
            "08:30|임원 조찬 간담회 @20층 라운지||현장 점검 후 직퇴", "10:00|주간 경영회의 @대회의실||", _
            "14:00~17:00|영업 현장 점검 (부산지사)|동행: 영업본부장, 영남지역팀장|"))
        Case 3: SampleExec = Array("전무", "최영호", Array( _
            "10:00|주간 경영회의 @대회의실||", "11:00|신사업 TF 회의|2027년 신규 진출 시장 검토|", "15:00|협력사 대표 면담 @접견실||"))
        Case 4: SampleExec = Array("대표이사", "김민수", Array( _
            "08:30|임원 조찬 간담회 @20층 라운지|참석: 부사장 이하 임원 전원|11/10(화) 싱가포르 출장 (08:00 인천 출발)", _
            "10:00|주간 경영회의 @대회의실^- 안건: 4분기 실적 전망, 2027년 사업계획 방향||", _
            "11:30|신규 고객사 대표 접견 @접견실||", "12:00|오찬 (고객사 대표)||", _
            "14:00|노사협의회 정기회의 @3층 회의실|노측 6명, 사측 6명|", _
            "16:00|투자심의위원회 @대회의실|안건: 제2공장 설비 증설 (!!승인 필요!!)|", "17:30|결재||", "19:00|업계 CEO 만찬||"))
        Case 5: SampleExec = Array("전무", "정다은", Array( _
            "09:30|재무팀 자금 현황 보고||오후 외부 일정 후 복귀", "10:00|주간 경영회의 @대회의실||", "14:00|주거래 은행 미팅 @외부||"))
        Case 6: SampleExec = Array("상무", "강태준", Array( _
            "종일|해외 바이어 상담 (서울사무소)|베트남 바이어 3개사|10:00 경영회의 불참 (정다은 전무 대리 보고)"))
        Case 7: SampleExec = Array("상무", "윤지훈", Array( _
            "10:00|주간 경영회의 @대회의실||", "13:00|연구소 과제 점검 @기술연구소|하반기 중점과제 5건|", "16:30|품질 회의||"))
        Case 8: SampleExec = Array("상무", "한소영", Array( _
            "09:00|홍보 전략 회의||내일 오전 반차", "11:00|언론사 인터뷰 @홍보실|주제: 친환경 경영|", "15:00|사회공헌 위원회||"))
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
    ws.Range("A1").Value = "직위"
    ws.Range("C1").Value = "성명"
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
    ws.Range("A2").Value = "한 줄에 일정 하나 · 시간은 09:00, 14:00~16:00, 종일 · 내용 끝에 @장소 · 내용이 * 로 시작하면 위 일정의 작은 글씨 · ※ 참고는 칸 맨 밑 · 시트 순서 = 자리 순서"
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = COL_GRAY
    ws.Range("A2").WrapText = True
    ws.Rows(2).RowHeight = 28
    With ws.Range("A3:D3")
        .Value = Array("시간", "내용", "비고 (작은 글씨 *)", "※ 참고")
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
    AddButton ws, ws.Columns(5).Left + 12, 4, 96, 24, "A4 만들기", "BuildA4"
    AddButton ws, ws.Columns(5).Left + 12, 32, 96, 24, "형광펜", "HighlightSel"
    AddButton ws, ws.Columns(5).Left + 12, 60, 96, 24, "입력으로", "GoInput"
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
    Lbl ws, "B1:I1", "임원 일일 예정사항", True, 15
    Lbl ws, "J1:M1", "8명 · A4 한 장 자동 맞춤", False, 9, COL_GRAY
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 60
    ws.Range("B3:M3").Merge
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").WrapText = True
    ws.Rows(3).RowHeight = 30
    DefName "d_status", ws, "$B$3"
    ' 문서 정보
    Lbl ws, "B5:M5", "문서 정보", True, 11
    Lbl ws, "B6:C6", "제목", False, 9, COL_GRAY: Inp ws, "D6:M6", "임원 일일 예정사항", "d_title"
    Lbl ws, "B7:C7", "날짜", False, 9, COL_GRAY: Inp ws, "D7:F7", Date, "d_date", False
    ws.Range("D7").NumberFormat = "yyyy-mm-dd"
    ws.Range("D7").HorizontalAlignment = XL_LEFT
    Lbl ws, "G7:H7", "왼쪽 위 문구", False, 9, COL_GRAY: Inp ws, "I7:M7", "", "d_sub"
    Lbl ws, "B8:C8", "하단 문구", False, 9, COL_GRAY: Inp ws, "D8:M8", "", "d_foot"
    ws.Range("D6:M8").Rows.RowHeight = 20
    ' 배치
    Lbl ws, "B10:M10", "배치 (임원 시트의 탭 순서 = 자리 순서 1~8)", True, 11
    Dim r As Long
    For r = 11 To 14
        ws.Rows(r).RowHeight = 34
    Next
    DefName "d_map", ws, "$B$11"
    ws.Range("B15:M15").Merge
    ws.Rows(15).RowHeight = 20
    ws.Range("B16:M16").Merge
    ws.Range("B16").Value = "칸을 누르면 그 임원 시트로 갑니다. 자리를 바꾸려면 [자리 바꾸기]를 누르거나 임원 시트 탭을 끌어 순서를 바꾸세요. 직위·성명은 각 임원 시트 맨 위에서 고칩니다."
    ws.Range("B16").Font.Size = 9
    ws.Range("B16").Font.Color = COL_GRAY
    ws.Range("B16").WrapText = True
    ws.Rows(16).RowHeight = 28
    ' 순서
    Lbl ws, "B18:M18", "매일 하는 일", True, 11
    Dim steps As Variant, i As Long
    steps = Array("1. 각 임원 시트에 일정을 씁니다(또는 [엑셀 불러오기]로 두레이 공동편집 파일을 가져옵니다).", _
        "2. [A4 만들기]를 누르면 A4 시트에 한 장짜리 문서가 만들어집니다(글자크기 자동 맞춤).", _
        "3. [인쇄 미리보기] 또는 A4 시트에서 Ctrl+P. PDF로 저장도 됩니다.", _
        "※ 단추가 안 보이면 Alt+F8(매크로)에서 BuildA4(A4 만들기), ImportExcel(엑셀 불러오기), UndoLast(되돌리기) 등을 실행하세요.")
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
    AddButton ws, x, y, bw, 24, "A4 만들기", "BuildA4"
    AddButton ws, x + (bw + 6), y, bw, 24, "인쇄 미리보기", "PrintPreviewA4"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, "엑셀 불러오기", "ImportExcel"
    AddButton ws, x + 3 * (bw + 6), y, bw, 24, "붙여넣은 표 가져오기", "ImportPasted"
    AddButton ws, x + 4 * (bw + 6), y, bw, 24, "되돌리기", "UndoLast"
    y = y + 28
    AddButton ws, x, y, bw, 24, "새 문서", "NewDoc"
    AddButton ws, x + (bw + 6), y, bw, 24, "자리 바꾸기", "SwapSeats"
    AddButton ws, x + 2 * (bw + 6), y, bw, 24, "형광펜", "HighlightSel"
    AddButton ws, x + 3 * (bw + 6), y, bw * 2 + 6, 24, "공동편집용 엑셀 양식 만들기", "MakeTemplate"
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
    AddButton ws, 600, 90, 120, 26, "다시 만들기", "BuildA4"
    AddButton ws, 600, 122, 120, 26, "인쇄 미리보기", "PrintPreviewA4"
    AddButton ws, 600, 154, 120, 26, "입력으로", "GoInput"
End Sub

Private Sub SetupTables(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 6
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 70
    Lbl ws, "A1:D1", "맨 밑 표 4개 (인접회사활동 · 휴가 · 출장/방문 · 근무)", True, 12
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = "표 이름·칸 제목은 고쳐 써도 됩니다. 내용 칸에는 여러 줄(Alt+Enter)을 쓸 수 있고, '10:00 회의'처럼 시간으로 시작하면 시간이 굵게, '* '로 시작하면 작은 글씨입니다. 근무의 칸 제목을 비우면 제목 줄 없이 찍힙니다."
    ws.Range("A2").WrapText = True
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = COL_GRAY
    ws.Rows(2).RowHeight = 40
    With ws.Range("A3:D3")
        .Value = Array("표 이름", "칸", "칸 제목", "내용")
        .Font.Bold = True
        .Interior.Color = COL_HEAD
        .HorizontalAlignment = XL_CENTER
    End With
    DefName "t_top", ws, "$A$4"
    Dim k As Long, j As Long, r As Long, heads As Variant, texts As Variant
    For k = 1 To 4
        r = 4 + TblRow0(k)
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = Choose(k, "인접회사활동", "휴가", "출장/방문", "근무")
        ws.Cells(r, 1).Font.Bold = True
        Select Case k
            Case 1
                heads = Array("A사", "B사", "C사", "D사", "E사")
                texts = Array("10:00 사장 주간회의^14:00 노사협의회", "임원 워크숍 (1박 2일)", "특이사항 없음", "15:00 신제품 발표회^* 본사 강당", "11:00 대표 언론 인터뷰")
            Case 2
                heads = Array("")
                texts = Array("상무 한소영 (11/10 오전 반차) · 기획팀장 (11/9~11/11 연차)")
            Case 3
                heads = Array("")
                texts = Array("부사장 박서연: 부산지사 현장 점검 (14:00~17:00) · 상무 강태준: 서울사무소 (종일)")
            Case 4
                heads = Array("당직", "비상대기")
                texts = Array("총무팀 과장 (18:00~익일 09:00)", "비서실장")
        End Select
        For j = 1 To TblN(k)
            ws.Cells(r + j - 1, 2).Value = j
            ws.Cells(r + j - 1, 2).HorizontalAlignment = XL_CENTER
            ws.Range(ws.Cells(r + j - 1, 3), ws.Cells(r + j - 1, 4)).NumberFormat = "@"
            If TblHasHeads(k) Then
                ws.Cells(r + j - 1, 3).Value = heads(j - 1)
            Else
                ws.Cells(r + j - 1, 3).Value = "(제목 없음)"
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
    AddButton ws, ws.Columns(5).Left + 12, 4, 96, 24, "A4 만들기", "BuildA4"
    AddButton ws, ws.Columns(5).Left + 12, 32, 96, 24, "형광펜", "HighlightSel"
    AddButton ws, ws.Columns(5).Left + 12, 60, 96, 24, "입력으로", "GoInput"
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
    Lbl ws, "B1:F1", "서식", True, 14
    ws.Rows(1).RowHeight = 28
    Dim r As Long
    r = 3
    Lbl ws, "B" & r & ":F" & r, "용지·배치", True, 11: r = r + 1
    Lbl ws, "B" & r, "용지 방향": Inp ws, "C" & r & ":D" & r, "A4 세로", "s_orient": ListValid ws, "C" & r, "A4 세로,A4 가로": r = r + 1
    Lbl ws, "B" & r, "여백 (mm, 5~25)": Inp ws, "C" & r, 10, "s_margin", False: r = r + 1
    Lbl ws, "B" & r, "글꼴": Inp ws, "C" & r & ":D" & r, "맑은 고딕", "s_font": ListValid ws, "C" & r, "맑은 고딕,나눔고딕,돋움,굴림,바탕": r = r + 1
    r = r + 1
    Lbl ws, "B" & r & ":F" & r, "칸 크기 (비율)", True, 11: r = r + 1
    Lbl ws, "B" & r, "윗줄 칸 폭 (왼쪽·가운데·오른쪽)": Inp ws, "C" & r, 1, "s_c1a", False: Inp ws, "D" & r, 1, "s_c1b", False: Inp ws, "E" & r, 1, "s_c1c", False: r = r + 1
    Lbl ws, "B" & r, "가운데 줄 폭 (큰 칸·오른쪽)": Inp ws, "C" & r, 2, "s_c2a", False: Inp ws, "D" & r, 1, "s_c2b", False: r = r + 1
    Lbl ws, "B" & r, "아랫줄 칸 폭 (왼쪽·오른쪽)": Inp ws, "C" & r, 1, "s_c3a", False: Inp ws, "D" & r, 1, "s_c3b", False: r = r + 1
    Lbl ws, "B" & r, "줄 높이 (윗줄·가운데·아랫줄)": Inp ws, "C" & r, "", "s_h1", False: Inp ws, "D" & r, "", "s_h2", False: Inp ws, "E" & r, "", "s_h3", False
    Lbl ws, "F" & r, "비우면 자동 (내용에 맞춤)", False, 9, COL_GRAY: r = r + 1
    Lbl ws, "B" & r, "가운데 줄 오른쪽 위 칸 높이": Inp ws, "C" & r, "", "s_split", False
    Lbl ws, "D" & r & ":F" & r, "0.1~0.9 (비우면 자동, 0.5 = 위아래 같게)", False, 9, COL_GRAY: r = r + 1
    Lbl ws, "B" & r, "지금 크기", False, 10, COL_GRAY
    ws.Range("C" & r & ":F" & r).Merge
    ws.Range("C" & r).WrapText = True
    ws.Range("C" & r).Font.Size = 9
    ws.Rows(r).RowHeight = 30
    DefName "s_sizes", ws, "$C$" & r
    r = r + 1
    Dim btnY As Double
    ws.Rows(r).RowHeight = 34
    btnY = ws.Rows(r).Top + 5
    AddButton ws, ws.Columns(2).Left, btnY, 118, 24, "가운데 큰 칸 2/3", "PresetBig23"
    AddButton ws, ws.Columns(2).Left + 124, btnY, 118, 24, "가운데 큰 칸 1/2", "PresetBig12"
    AddButton ws, ws.Columns(2).Left + 248, btnY, 118, 24, "높이만 자동으로", "PresetAutoH"
    AddButton ws, ws.Columns(2).Left + 372, btnY, 118, 24, "모두 처음대로", "PresetReset"
    r = r + 2
    Lbl ws, "B" & r & ":F" & r, "글자크기 (자동 맞춤)", True, 11: r = r + 1
    Lbl ws, "B" & r, "최대 (pt)": Inp ws, "C" & r, 12, "s_max", False: r = r + 1
    Lbl ws, "B" & r, "최소 (pt)": Inp ws, "C" & r, 7, "s_min", False: r = r + 1
    ws.Range("B" & r & ":F" & r).Merge
    ws.Range("B" & r).Value = "8칸 모두 같은 글자크기로, A4 한 장을 꽉 채우는 가장 큰 크기를 이 범위 안에서 고릅니다(0.5pt 단위). 남는 높이는 칸을 늘려 채우고, ※ 참고는 늘 칸의 맨 밑에 붙습니다."
    ws.Range("B" & r).WrapText = True
    ws.Range("B" & r).Font.Size = 9
    ws.Range("B" & r).Font.Color = COL_GRAY
    ws.Rows(r).RowHeight = 30
    r = r + 2
    Lbl ws, "B" & r & ":F" & r, "기타", True, 11: r = r + 1
    Lbl ws, "B" & r, "맨 밑 표 4개 넣기": Inp ws, "C" & r, "예", "s_tables": ListValid ws, "C" & r, "예,아니오": r = r + 1
    Lbl ws, "B" & r, "오른쪽 위 빨간 상자": Inp ws, "C" & r, "예", "s_secret": ListValid ws, "C" & r, "예,아니오"
    Inp ws, "D" & r & ":E" & r, "비공개", "s_secretText": r = r + 1
    Lbl ws, "B" & r, "일정이 없는 칸에 쓸 문구": Inp ws, "C" & r & ":E" & r, "일정 없음", "s_empty"
    Lbl ws, "F" & r, "비워 두면 빈칸", False, 9, COL_GRAY: r = r + 2
    Lbl ws, "B" & r & ":F" & r, "문서 암호", True, 11: r = r + 1
    ws.Range("B" & r & ":F" & r).Merge
    ws.Range("B" & r).Value = "이 파일을 열 때 암호를 묻게 하려면 한셀의 문서 암호를 쓰세요(내용이 암호화됩니다). 아래 단추로 정하거나, 한셀 메뉴의 문서 암호에서 정합니다. 암호를 잊으면 열 수 없습니다."
    ws.Range("B" & r).WrapText = True
    ws.Range("B" & r).Font.Size = 9
    ws.Rows(r).RowHeight = 30
    r = r + 1
    ws.Rows(r).RowHeight = 34
    AddButton ws, ws.Columns(2).Left, ws.Rows(r).Top + 5, 140, 24, "문서 암호 정하기", "SetDocPassword"
    AddButton ws, ws.Columns(2).Left + 146, ws.Rows(r).Top + 5, 118, 24, "A4 만들기", "BuildA4"
End Sub

Private Sub SetupPaste(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Range("A1:H1").Merge
    ws.Range("A1").Value = "엑셀(두레이)에서 복사한 표를 3번째 줄 A칸부터 붙여넣고 [붙여넣은 표 가져오기]를 누르세요. 머리글(시간·내용…)이 있어야 하고, 여러 임원이면 '성명' 열을 두세요."
    ws.Range("A1").WrapText = True
    ws.Range("A1").Font.Size = 9
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 34
    AddButton ws, 4, ws.Rows(2).Top + 5, 140, 24, "붙여넣은 표 가져오기", "ImportPasted"
    AddButton ws, 150, ws.Rows(2).Top + 5, 96, 24, "입력으로", "GoInput"
End Sub

Private Sub SetupHelp(ByVal ws As Object)
    ClearSheet ws
    SheetBase ws
    ws.Columns(1).ColumnWidth = 44
    ws.Columns(2).ColumnWidth = 60
    Dim rows As Collection, i As Long, f As Variant, k As Long, ln As String
    Set rows = New Collection
    rows.Add "#일정 칸 — 한 줄에 일정 하나"
    rows.Add "입력|인쇄물"
    rows.Add "시간 칸 09:00, 내용 칸 주간 업무회의|시간 09:00 (굵게) / 내용"
    rows.Add "14:00~16:00|시간 범위 (~ 뒤에서 줄 바꿈)"
    rows.Add "9:30, 9시 반, 오후 2시|09:30, 09:30, 14:00 처럼 맞춰 줌"
    rows.Add "종일, 오전 (시간 칸에)|시간 칸에 종일 / 오전"
    rows.Add "시간 칸을 비우고 내용에 09:00 회의|시간과 내용을 알아서 나눔"
    rows.Add "내용 끝에 @소회의실|내용 (소회의실)"
    rows.Add "비고 칸, 또는 내용이 * 로 시작하는 줄|바로 위 일정 밑에 작은 글씨 부가 설명"
    rows.Add "내용 칸 안에서 Alt+Enter 후 - 안건|바로 위 일정 밑에 하위 항목"
    rows.Add "※ 참고 칸, 또는 내용이 ※ 로 시작|그 임원 칸 맨 밑 ※ 참고"
    rows.Add "**굵게**, !!빨간 강조!!, __밑줄__, ==형광펜==|글자 꾸밈 (형광펜은 그 줄의 칸 전체가 노랗게)"
    rows.Add ""
    rows.Add "#단추"
    rows.Add "A4 만들기|A4 시트에 한 장짜리 문서를 새로 만듭니다(글자크기 자동 맞춤)"
    rows.Add "엑셀 불러오기|두레이에서 내려받은 공동편집 파일(.xlsx·.csv)을 골라 임원별로 넣습니다"
    rows.Add "붙여넣은 표 가져오기|붙여넣기 시트에 붙인 표를 임원별로 넣습니다"
    rows.Add "되돌리기|마지막 [엑셀 불러오기]·[붙여넣은 표 가져오기]·[새 문서] 전으로"
    rows.Add "새 문서|일정·※ 참고·맨 밑 표 내용만 비움 (직위·성명·서식은 그대로, 날짜는 오늘)"
    rows.Add "자리 바꾸기|두 자리의 임원을 맞바꿈 (임원 시트 탭을 끌어 순서를 바꿔도 됨)"
    rows.Add "형광펜|고른 칸의 글자를 ==…==로 감쌈 (다시 누르면 지움, Ctrl+Shift+H)"
    rows.Add "공동편집용 엑셀 양식 만들기|임원마다 시트가 있는 .xlsx — 두레이에 올려 함께 작성"
    rows.Add ""
    rows.Add "#엑셀 불러오기 규칙"
    rows.Add "임원 찾기|시트 이름에 든 성명으로 (없으면 겹치지 않는 직위로). 못 찾으면 자리 번호를 물어봄"
    rows.Add "한 시트에 여러 임원|'성명' 열을 두세요 (병합된 칸도 됨)"
    rows.Add "날짜 열|문서 날짜의 줄만 가져옴. 문서 날짜에 없으면 일정이 가장 많은 날짜를 권하고, 가져오면 문서 날짜도 바뀜"
    rows.Add "빼는 시트|빈 시트, 작성방법 같은 안내 시트, 숨긴 시트"
    rows.Add ""
    rows.Add "#문서 암호"
    rows.Add "열 때 암호 묻기|서식 시트의 [문서 암호 정하기] 또는 한셀 메뉴의 문서 암호. 파일 내용이 암호화됩니다"
    rows.Add ""
    rows.Add "#단추가 안 보이거나 눌리지 않을 때"
    rows.Add "Alt+F8 (매크로 목록)|BuildA4, ImportExcel, ImportPasted, UndoLast, NewDoc, SwapSeats, HighlightSel, MakeTemplate, PrintPreviewA4, SetDocPassword"
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
            If f(0) = "입력" Then ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 2)).Interior.Color = COL_HEAD
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
    r = r & "mark=" & MarkLines("09:00 회의 @대회의실" & vbLf & "* 참석: 실장" & vbLf & "※ 메모") & "; "
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
    ws.Range("A3:C3").Value = Array("성명", "시간", "내용")
    ws.Range("A4:C4").Value = Array("한소영", "10:00", "붙여넣기 시험")
    ImportPasted
    LoadExecs
    r = r & "pasted8=" & exItems(8).Count & " pasteLeft=" & ws.Range("A4").Value & "; "
    UndoLast
    gStep = "template"
    p = Environ$("HC_TEST_OUT")
    If p <> "" Then r = r & "template=" & MakeTemplateTo(p) & "; "
    gStep = "landscape"
    NSet "s_orient", "A4 가로"
    NSet "s_tables", "아니오"
    NSet "s_secret", "아니오"
    DoBuild
    r = r & "landscape status=" & NStr("d_status") & "; "
    NSet "s_orient", "A4 세로"
    NSet "s_tables", "예"
    NSet "s_secret", "예"
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

' 이 모듈을 한셀이 읽었는지 확인 (실행했을 때 창이 뜨면 정상)
Public Sub CheckHcExecSetup()
    MsgBox "HcExecSetup OK"
End Sub
