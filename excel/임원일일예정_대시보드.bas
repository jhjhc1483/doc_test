Attribute VB_Name = "ExecDashboard"
'=====================================================================
'  임원 일일 예정사항 - 엑셀 대시보드 (A4 한 장 자동 맞춤)
'---------------------------------------------------------------------
'  처음 한 번
'    1) 새 엑셀 파일에서 Alt+F11 → [파일] → [파일 가져오기] → 이 .bas 파일
'    2) Alt+F8 → Setup → [실행]
'    3) [다른 이름으로 저장] → 파일 형식 "Excel 매크로 사용 통합 문서 (*.xlsm)"
'  매일
'    임원별 시트(시트 이름 = 직위 성명)에 일정 입력
'    → [새로 고침] (Ctrl+Shift+R) → [인쇄 미리보기] 또는 [PDF로 저장]
'    두레이 공동편집 파일은 [엑셀 불러오기]로 가져옵니다.
'
'  대시보드 시트는 매번 새로 그립니다. 손으로 고친 도형은 새로 고침하면 사라집니다.
'=====================================================================
Option Explicit

'---------------------------------------------------------------------
' 이름·상수
'---------------------------------------------------------------------
Private Const APP_TITLE As String = "임원 일일 예정사항"
Private Const SH_DASH As String = "대시보드"
Private Const SH_CFG As String = "설정"
Private Const SH_GUIDE As String = "작성방법"
Private Const N_EX As Long = 8
Private Const MAXP As Long = 300
Private Const MAXR As Long = 60
Private Const PFX As String = "ED_"
Private Const PT_MM As Double = 2.83464566929134
Private Const GAP_MM As Double = 2.4

' Office·Excel 상수 (참조 설정 없이 쓰려고 숫자로 둡니다)
Private Const M_TRUE As Long = -1
Private Const M_FALSE As Long = 0
Private Const M_RECT As Long = 1
Private Const M_HORIZ As Long = 1
Private Const M_AUTO_NONE As Long = 0
Private Const M_AUTO_FIT As Long = 1
Private Const M_ANCHOR_TOP As Long = 1
Private Const M_ANCHOR_MIDDLE As Long = 3
Private Const M_ALIGN_LEFT As Long = 1
Private Const M_ALIGN_CENTER As Long = 2
Private Const M_ALIGN_RIGHT As Long = 3
Private Const M_TAB_LEFT As Long = 1
Private Const M_DASH As Long = 4
Private Const M_UNDERLINE As Long = 2
Private Const X_FREE As Long = 3
Private Const X_PORTRAIT As Long = 1
Private Const X_LANDSCAPE As Long = 2
Private Const X_A4 As Long = 9
Private Const X_PDF As Long = 0
Private Const X_VISIBLE As Long = -1
Private Const X_VALID_LIST As Long = 3
Private Const X_CENTER As Long = -4108
Private Const X_TOP As Long = -4160
Private Const X_FORMULAS As Long = -4123
Private Const X_PART As Long = 2
Private Const X_BYROWS As Long = 1
Private Const X_BYCOLS As Long = 2
Private Const X_PREV As Long = 2
Private Const X_WB_SHEET As Long = -4167
Private Const X_XLSX As Long = 51

'---------------------------------------------------------------------
' 임원 칸 내용 (자리 1~8)
'   gKind: M 일정 / C 일정의 다음 줄 / S * 부가 설명 / S2 그 다음 줄
'          D - 하위 항목 / D2 그 다음 줄 / E 일정 없음 문구
'---------------------------------------------------------------------
Private gPos(1 To N_EX) As String
Private gName(1 To N_EX) As String
Private gSheetOf(1 To N_EX) As String
Private gN(1 To N_EX) As Long
Private gKind(1 To N_EX, 1 To MAXP) As String
Private gTime(1 To N_EX, 1 To MAXP) As String
Private gText(1 To N_EX, 1 To MAXP) As String
Private gNR(1 To N_EX) As Long
Private gRef(1 To N_EX, 1 To MAXR) As String

' 설정
Private sTitle As String, sSub As String, sFoot As String, sEmpty As String
Private sDate As Date, sDateISO As String
Private sLandscape As Boolean, sMargin As Double, sFont As String
Private sMaxF As Double, sMinF As Double
Private sC1a As Double, sC1b As Double, sC1c As Double
Private sC2a As Double, sC2b As Double
Private sC3a As Double, sC3b As Double
Private sRowsSet As Boolean, sR1 As Double, sR2 As Double, sR3 As Double
Private sSplit As Double

' 시트에서 읽은 줄 (묶음 = 임원 한 명분)
Private rN As Long
Private rGrp() As Long, rDate() As String, rTime() As String, rCont() As String, rStar() As String, rRef() As String
Private grN As Long
Private grLabel() As String, grSheet() As String, grTitle() As String, grKey() As String, grHasExec() As Boolean

' 대시보드 배치 (단위 pt)
Private dX0 As Double, dY0 As Double, dW As Double, dH As Double, dGap As Double, dPadH As Double
Private dPageW As Double, dPageH As Double, dMarginPt As Double
Private bX(1 To N_EX) As Double, bW(1 To N_EX) As Double
Private bHasTime(1 To N_EX) As Boolean, bTimeU(1 To N_EX) As Double
Private hHead(1 To N_EX) As Double, hBody(1 To N_EX) As Double, hRefBox(1 To N_EX) As Double, hNat(1 To N_EX) As Double
Private hTitle As Double, hMeta As Double, hFoot As Double
Private fReq0 As Double, fReq1 As Double, fReq2 As Double
Private fAvail As Double, fNeed As Double, fInner As Double, fOK As Boolean, fFs As Double
Private fBottom As Double
Private shF(1 To N_EX) As Object, shH(1 To N_EX) As Object, shB(1 To N_EX) As Object, shR(1 To N_EX) As Object
Private shTitle As Object, shSub As Object, shDate As Object, shFoot As Object
Private gStep As String
Private gLastOK As Boolean

'=====================================================================
'  매크로 (Alt+F8 목록·단추에서 부르는 것)
'=====================================================================

' 처음 한 번: 설정·대시보드·임원별 시트를 만들고 예시 일정을 넣습니다.
Public Sub Setup()
    Dim wsD As Object, wsC As Object, ws As Object, i As Long, ans As Long, nm As String
    If SheetExists(SH_CFG) Then
        ans = MsgBox("이미 설치되어 있습니다." & vbLf & vbLf & _
            "[예]  설정 시트를 처음 상태로 다시 만들고 예시 일정을 넣습니다." & vbLf & _
            "       (지금 입력한 설정과 임원별 일정이 지워집니다)" & vbLf & _
            "[아니요]  그대로 두고 대시보드만 다시 그립니다.", _
            vbQuestion + vbYesNoCancel + vbDefaultButton2, APP_TITLE)
        If ans = vbCancel Then Exit Sub
        If ans = vbNo Then
            RefreshDashboard
            Exit Sub
        End If
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False
    gStep = "시트 만들기"
    If SheetExists(SH_DASH) Then
        Set wsD = ThisWorkbook.Worksheets(SH_DASH)
    Else
        Set wsD = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsD.Name = SH_DASH
    End If
    If SheetExists(SH_CFG) Then
        Set wsC = ThisWorkbook.Worksheets(SH_CFG)
    Else
        Set wsC = ThisWorkbook.Worksheets.Add(After:=wsD)
        wsC.Name = SH_CFG
    End If
    If wsD.Index <> 1 Then wsD.Move Before:=ThisWorkbook.Sheets(1)
    If wsC.Index <> 2 Then wsC.Move After:=wsD

    gStep = "설정 시트"
    BuildSettingsSheet wsC
    ReadSettings

    gStep = "임원별 시트"
    For i = 1 To N_EX
        nm = DesiredSheetName(i)
        If SheetExists(nm) Then
            Set ws = ThisWorkbook.Worksheets(nm)
            ws.Cells.Clear
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = nm
        End If
        FormatInputSheet ws
        WriteSample ws, i
    Next

    gStep = "빈 기본 시트 정리"
    DeleteBlankDefaultSheets

    On Error Resume Next
    Application.MacroOptions Macro:="RefreshDashboard", HasShortcutKey:=True, ShortcutKey:="R"
    If Err.Number <> 0 Then
        Err.Clear
        Application.MacroOptions Macro:="'" & ThisWorkbook.Name & "'!RefreshDashboard", HasShortcutKey:=True, ShortcutKey:="R"
    End If
    On Error GoTo Fail

    Application.ScreenUpdating = True
    RefreshDashboard
    If gLastOK Then
        MsgBox "설치를 마쳤습니다. 예시 일정 8명이 들어 있습니다." & vbLf & vbLf & _
            "1. [파일] → [다른 이름으로 저장] → 파일 형식을" & vbLf & _
            "    'Excel 매크로 사용 통합 문서 (*.xlsm)'로 골라 저장하세요." & vbLf & _
            "2. 일정은 임원별 시트(시트 이름 = 직위 성명)에 쓰고" & vbLf & _
            "    [새로 고침] 단추나 Ctrl+Shift+R을 누르세요." & vbLf & _
            "3. 직위·성명·자리·글꼴·용지 방향은 '설정' 시트에서 바꿉니다.", vbInformation, APP_TITLE
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "설치하는 중 문제가 생겼습니다 (" & gStep & ")." & vbLf & Err.Description & " (" & Err.Number & ")", vbExclamation, APP_TITLE
End Sub

' 임원별 시트를 읽어 대시보드를 다시 그리고, 글자크기를 맞춰 A4 한 장에 넣습니다. (Ctrl+Shift+R)
Public Sub RefreshDashboard()
    Dim ws As Object, lastRow As Long
    gLastOK = False
    If Not CheckInstalled() Then Exit Sub
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "임원 대시보드 만드는 중..."

    gStep = "설정 읽기"
    ReadSettings
    gStep = "임원별 시트 확인"
    SyncExecSheets
    gStep = "일정 읽기"
    LoadAllSchedules

    Set ws = DashSheet()
    gStep = "대시보드 준비"
    ClearDashShapes ws
    PrepareGrid ws
    ComputeBoxX
    gStep = "도형 만들기"
    CreateAllShapes ws
    gStep = "글자크기 맞추기"
    FitFont
    gStep = "칸 배치"
    PlaceAll ws
    gStep = "인쇄 설정"
    lastRow = FinishPage(ws)
    AddDashButtons ws

    Application.ScreenUpdating = True
    ' 화면을 끈 채로 잰 높이가 맞는지 한 번 더 확인 (안 맞으면 화면을 켠 채로 다시 맞춤)
    gStep = "확인"
    If Not VerifyBodies() Then
        FitFont
        ClearLinesOnly ws
        PlaceAll ws
        lastRow = FinishPage(ws)
    End If
    WriteStatus ws
    ZoomToPage ws, lastRow
    Application.StatusBar = False
    Application.EnableEvents = True
    gLastOK = True
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "대시보드를 만드는 중 문제가 생겼습니다 (" & gStep & ")." & vbLf & Err.Description & " (" & Err.Number & ")", vbExclamation, APP_TITLE
End Sub

' 새로 고친 뒤 인쇄 미리보기 (여기서 바로 인쇄할 수 있습니다)
Public Sub PreviewDashboard()
    RefreshDashboard
    If Not gLastOK Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_DASH).PrintPreview
End Sub

' 새로 고친 뒤 PDF로 저장
Public Sub SaveDashboardPDF()
    Dim fn As Variant, folder As String
    RefreshDashboard
    If Not gLastOK Then Exit Sub
    folder = ThisWorkbook.Path
    If folder <> "" Then folder = folder & Application.PathSeparator
    fn = Application.GetSaveAsFilename(InitialFileName:=folder & "임원일일예정_" & sDateISO & ".pdf", _
        FileFilter:="PDF 파일 (*.pdf),*.pdf", Title:="PDF로 저장")
    If VarType(fn) = vbBoolean Then Exit Sub
    On Error GoTo Fail
    ThisWorkbook.Worksheets(SH_DASH).ExportAsFixedFormat Type:=X_PDF, Filename:=CStr(fn), Quality:=0, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Exit Sub
Fail:
    MsgBox "PDF로 저장하지 못했습니다." & vbLf & Err.Description, vbExclamation, APP_TITLE
End Sub

' 두레이에서 내려받은 공동편집 엑셀(.xlsx)을 골라 임원별 시트로 가져옵니다.
Public Sub ImportExcel()
    Dim fn As Variant
    If Not CheckInstalled() Then Exit Sub
    ReadSettings
    fn = Application.GetOpenFilename("엑셀 파일 (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", 1, _
        "두레이에서 내려받은 공동편집 엑셀 파일 고르기")
    If VarType(fn) = vbBoolean Then Exit Sub
    If StrComp(CStr(fn), ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "지금 열려 있는 이 파일은 고를 수 없습니다.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If ImportFromFile(CStr(fn), True) <> "" Then RefreshDashboard
End Sub

' 파일 하나를 읽어 임원별 시트에 넣습니다. ask = False 면 묻지 않고 진행(시험용). 넣었으면 요약 글을 돌려줌.
Private Function ImportFromFile(ByVal fn As String, ByVal ask As Boolean) As String
    Dim wb As Object, ws As Object, i As Long, g As Long, k As Long
    Dim dKey() As String, dCnt() As Long, nd As Long, D As String, best As Long, found As Boolean
    Dim tgt() As Long, cnt(1 To N_EX) As Long, srcOf(1 To N_EX) As String, msg As String, miss As String
    Dim nEmpty As Long, nRow As Long, sched As String, refs As String, nOK As Long

    On Error GoTo Fail
    gStep = "파일 열기"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set wb = Workbooks.Open(Filename:=fn, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    gStep = "시트 읽기"
    ResetRead
    For Each ws In wb.Worksheets
        If ws.Visible = X_VISIBLE And Not IsSkipSheet(ws.Name) Then ReadSheet ws, Year(sDate)
    Next
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If grN = 0 Then
        If ask Then MsgBox "가져올 일정을 찾지 못했습니다." & vbLf & "머리글(시간·내용) 줄이 있는지 확인하세요.", vbExclamation, APP_TITLE
        Exit Function
    End If

    ' 날짜 열이 있으면: 문서 날짜의 줄, 없으면 줄이 가장 많은 날짜
    nd = 0
    ReDim dKey(1 To rN + 1): ReDim dCnt(1 To rN + 1)
    For k = 1 To rN
        If rDate(k) <> "" Then
            found = False
            For i = 1 To nd
                If dKey(i) = rDate(k) Then
                    dCnt(i) = dCnt(i) + 1
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                nd = nd + 1
                dKey(nd) = rDate(k)
                dCnt(nd) = 1
            End If
        End If
    Next
    D = ""
    If nd > 0 Then
        best = 1
        For i = 1 To nd
            If dKey(i) = sDateISO Then
                best = i
                Exit For
            End If
            If dCnt(i) > dCnt(best) Then best = i
        Next
        D = dKey(best)
        If D <> sDateISO And ask Then
            If MsgBox("문서 날짜(" & FmtDate(sDateISO) & ")의 일정이 엑셀에 없습니다." & vbLf & _
                "일정이 가장 많은 " & FmtDate(D) & " 일정을 가져오고 문서 날짜도 그 날로 바꿀까요?", _
                vbQuestion + vbYesNo, APP_TITLE) <> vbYes Then Exit Function
        End If
    End If

    ' 묶음 → 임원
    ReDim tgt(1 To grN)
    For g = 1 To grN
        tgt(g) = MatchExec(grLabel(g))
        If tgt(g) = 0 And Not grHasExec(g) And grTitle(g) <> "" Then tgt(g) = MatchExec(grTitle(g))
    Next
    nEmpty = 0
    For g = 1 To grN
        nRow = 0
        sched = "": refs = ""
        For k = 1 To rN
            If rGrp(k) = g And (rDate(k) = "" Or D = "" Or rDate(k) = D) Then
                nRow = nRow + 1
                RowToLines rTime(k), rCont(k), rStar(k), rRef(k), sched, refs
            End If
        Next
        If sched = "" And refs = "" Then
            nEmpty = nEmpty + 1
            tgt(g) = -1
        ElseIf tgt(g) = 0 Then
            miss = miss & vbLf & "   · " & SrcText(g)
        Else
            cnt(tgt(g)) = cnt(tgt(g)) + nRow
            If srcOf(tgt(g)) <> "" Then srcOf(tgt(g)) = srcOf(tgt(g)) & ", "
            srcOf(tgt(g)) = srcOf(tgt(g)) & SrcText(g)
        End If
    Next

    msg = ""
    nOK = 0
    For i = 1 To N_EX
        If cnt(i) > 0 Then
            nOK = nOK + 1
            msg = msg & vbLf & "   " & i & ". " & JTrim(gPos(i) & " " & gName(i)) & "  ←  " & srcOf(i) & " (" & cnt(i) & "줄)"
        End If
    Next
    If nOK = 0 Then
        If ask Then MsgBox "가져올 일정을 찾지 못했습니다." & IIf(miss <> "", vbLf & vbLf & "임원을 찾지 못한 곳:" & miss, "") & vbLf & vbLf & _
            "시트 이름에 임원 성명이 들어 있는지(예: 대표이사 김민수) 확인하세요.", vbExclamation, APP_TITLE
        Exit Function
    End If
    msg = "가져올 내용" & IIf(D <> "", " (" & FmtDate(D) & ")", "") & ":" & msg
    If miss <> "" Then msg = msg & vbLf & vbLf & "임원을 찾지 못해 뺀 곳 (시트 이름에 성명을 넣어 주세요):" & miss
    If nEmpty > 0 Then msg = msg & vbLf & vbLf & "아직 안 쓴 빈 시트 " & nEmpty & "곳은 뺐습니다."
    msg = msg & vbLf & vbLf & "가져올까요? 위 임원의 시트 내용은 새 내용으로 바뀝니다. (엑셀에 없는 임원은 그대로)"
    If ask Then
        If MsgBox(msg, vbQuestion + vbYesNo, APP_TITLE) <> vbYes Then Exit Function
    End If

    gStep = "임원별 시트에 쓰기"
    Application.ScreenUpdating = False
    SyncExecSheets
    For i = 1 To N_EX
        If cnt(i) > 0 Then WriteImported ThisWorkbook.Worksheets(gSheetOf(i)), i, tgt, D
    Next
    If D <> "" And D <> sDateISO Then Cfg("ED_Date").Value = DateSerial(CLng(Left$(D, 4)), CLng(Mid$(D, 6, 2)), CLng(Mid$(D, 9, 2)))
    Application.ScreenUpdating = True
    ImportFromFile = msg
    Exit Function
Fail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If ask Then MsgBox "엑셀을 가져오는 중 문제가 생겼습니다 (" & gStep & ")." & vbLf & Err.Description, vbExclamation, APP_TITLE
    ImportFromFile = ""
End Function

' 두레이에 올릴 공동편집용 양식(.xlsx): 작성방법 + 임원마다 시트 하나
Public Sub MakeCoEditTemplate()
    Dim fn As Variant, wb As Object, ws As Object, i As Long, nm As String, used As String, k As Long
    If Not CheckInstalled() Then Exit Sub
    ReadSettings
    fn = Application.GetSaveAsFilename(InitialFileName:="임원일일예정_공동편집양식.xlsx", _
        FileFilter:="Excel 통합 문서 (*.xlsx),*.xlsx", Title:="공동편집용 엑셀 양식 저장")
    If VarType(fn) = vbBoolean Then Exit Sub
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(X_WB_SHEET)
    Set ws = wb.Worksheets(1)
    ws.Name = SH_GUIDE
    WriteGuideSheet ws
    used = "|" & SH_GUIDE & "|"
    For i = 1 To N_EX
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        nm = DesiredSheetName(i)
        k = 2
        Do While InStr(1, used, "|" & nm & "|", vbTextCompare) > 0
            nm = Left$(DesiredSheetName(i), 27) & "(" & k & ")"
            k = k + 1
        Loop
        used = used & nm & "|"
        ws.Name = nm
        FormatInputSheet ws
    Next
    wb.Worksheets(1).Activate
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(fn), FileFormat:=X_XLSX
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "공동편집용 양식을 저장했습니다." & vbLf & CStr(fn) & vbLf & vbLf & _
        "두레이에 올려 각 비서가 자기 임원 시트에 쓰게 하세요." & vbLf & _
        "다 쓰면 내려받아 [엑셀 불러오기]로 가져옵니다.", vbInformation, APP_TITLE
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "양식을 만들지 못했습니다." & vbLf & Err.Description, vbExclamation, APP_TITLE
End Sub

' 임원별 시트의 일정만 모두 지웁니다 (직위·성명·설정은 그대로).
Public Sub ClearSchedules()
    Dim i As Long, ws As Object, lr As Long
    If Not CheckInstalled() Then Exit Sub
    If MsgBox("임원별 시트의 일정을 모두 지웁니다." & vbLf & "(머리글·직위·성명·설정은 그대로 둡니다)" & vbLf & vbLf & "계속할까요?", _
        vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) <> vbYes Then Exit Sub
    On Error GoTo Fail
    Application.ScreenUpdating = False
    ReadSettings
    SyncExecSheets
    For i = 1 To N_EX
        Set ws = ThisWorkbook.Worksheets(gSheetOf(i))
        lr = LastRowOf(ws)
        If lr >= 2 Then ws.Range(ws.Rows(2), ws.Rows(lr)).ClearContents
    Next
    Application.ScreenUpdating = True
    RefreshDashboard
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "지우는 중 문제가 생겼습니다." & vbLf & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub GoSettings()
    If SheetExists(SH_CFG) Then ThisWorkbook.Worksheets(SH_CFG).Activate
End Sub

Public Sub GoDashboard()
    If SheetExists(SH_DASH) Then ThisWorkbook.Worksheets(SH_DASH).Activate
End Sub

'=====================================================================
'  설정 시트
'=====================================================================
Private Function CheckInstalled() As Boolean
    Dim nm As Object
    If SheetExists(SH_CFG) Then
        On Error Resume Next
        Set nm = ThisWorkbook.Names("ED_Title")
        On Error GoTo 0
        If Not nm Is Nothing Then
            CheckInstalled = True
            Exit Function
        End If
    End If
    MsgBox "아직 설치하지 않았습니다." & vbLf & "Alt+F8 → Setup을 먼저 실행하세요.", vbExclamation, APP_TITLE
End Function

Private Function Cfg(ByVal nm As String) As Object
    Set Cfg = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function CfgText(ByVal nm As String) As String
    CfgText = JTrim(SafeStr(Cfg(nm).Cells(1, 1).Value))
End Function

Private Function NumOf(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = False
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then
        If Not IsNumeric(JTrim(CStr(v))) Then Exit Function
        NumOf = CDbl(JTrim(CStr(v)))
    ElseIf IsNumeric(v) Then
        NumOf = CDbl(v)
    Else
        Exit Function
    End If
    ok = True
End Function

Private Function CfgNum(ByVal nm As String, ByVal def As Double, ByVal lo As Double, ByVal hi As Double) As Double
    Dim ok As Boolean, x As Double
    x = NumOf(Cfg(nm).Cells(1, 1).Value, ok)
    If Not ok Then x = def
    CfgNum = Clamp(x, lo, hi)
End Function

' 비율 칸 n개 → 합이 1인 비율 (잘못 쓰면 기본값)
Private Sub ReadFracs(ByVal nm As String, ByVal n As Long, ByVal d1 As Double, ByVal d2 As Double, ByVal d3 As Double, _
                      ByRef o1 As Double, ByRef o2 As Double, ByRef o3 As Double)
    Dim rg As Object, ok As Boolean, a(1 To 3) As Double, i As Long, s As Double, good As Boolean
    Set rg = Cfg(nm)
    good = True
    For i = 1 To n
        a(i) = NumOf(rg.Cells(1, i).Value, ok)
        If Not ok Or a(i) <= 0 Then good = False
    Next
    If Not good Then
        a(1) = d1: a(2) = d2: a(3) = d3
    End If
    s = 0
    For i = 1 To n
        s = s + a(i)
    Next
    For i = 1 To n
        a(i) = a(i) / s
        If a(i) < 0.12 Then a(i) = 0.12
    Next
    s = 0
    For i = 1 To n
        s = s + a(i)
    Next
    o1 = a(1) / s: o2 = a(2) / s
    If n = 3 Then o3 = a(3) / s Else o3 = 0
End Sub

Private Sub ReadSettings()
    Dim v As Variant, i As Long, rg As Object, ok As Boolean, x As Double, iso As String, t As Double
    Dim r1 As Double, r2 As Double, r3 As Double, okA As Boolean, okB As Boolean, okC As Boolean

    sTitle = CfgText("ED_Title")
    sSub = CfgText("ED_Sub")
    sFoot = CfgText("ED_Foot")
    sEmpty = CfgText("ED_Empty")

    v = Cfg("ED_Date").Cells(1, 1).Value
    sDate = Date
    If VarType(v) = vbDate Then
        sDate = DateSerial(Year(v), Month(v), Day(v))
    Else
        x = NumOf(v, ok)
        If ok And x >= 1 And VarType(v) <> vbString Then
            sDate = CDate(Int(x))
        Else
            iso = ParseLooseDate(SafeStr(v), Year(Date))
            If iso <> "" Then sDate = DateSerial(CLng(Left$(iso, 4)), CLng(Mid$(iso, 6, 2)), CLng(Mid$(iso, 9, 2)))
        End If
    End If
    sDateISO = IsoOf(sDate)

    sLandscape = (InStr(1, CfgText("ED_Orient"), "가로") > 0)
    sMargin = CfgNum("ED_Margin", 10, 5, 25)
    sFont = CfgText("ED_Font")
    If sFont = "" Then sFont = "맑은 고딕"
    sMaxF = CfgNum("ED_FontMax", 12, 6, 20)
    sMinF = CfgNum("ED_FontMin", 7, 5, 16)
    If sMinF > sMaxF Then
        t = sMinF: sMinF = sMaxF: sMaxF = t
    End If

    Set rg = Cfg("ED_Execs")
    For i = 1 To N_EX
        gPos(i) = JTrim(SafeStr(rg.Cells(i, 1).Value))
        gName(i) = JTrim(SafeStr(rg.Cells(i, 2).Value))
    Next

    ReadFracs "ED_ColsTop", 3, 1, 1, 1, sC1a, sC1b, sC1c
    ReadFracs "ED_ColsMid", 2, 2, 1, 0, sC2a, sC2b, x
    ReadFracs "ED_ColsBot", 2, 1, 1, 0, sC3a, sC3b, x

    Set rg = Cfg("ED_Rows")
    r1 = NumOf(rg.Cells(1, 1).Value, okA)
    r2 = NumOf(rg.Cells(1, 2).Value, okB)
    r3 = NumOf(rg.Cells(1, 3).Value, okC)
    sRowsSet = okA And okB And okC
    If sRowsSet Then sRowsSet = (r1 > 0 And r2 > 0 And r3 > 0)
    If sRowsSet Then
        t = r1 + r2 + r3
        sR1 = r1 / t: sR2 = r2 / t: sR3 = r3 / t
    End If

    x = NumOf(Cfg("ED_Split").Cells(1, 1).Value, ok)
    If ok And x >= 0.1 And x <= 0.9 Then sSplit = x Else sSplit = -1
End Sub

Private Sub AddName(ByVal nm As String, ByVal addr As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & SH_CFG & "'!" & addr
End Sub

Private Function SlotLabel(ByVal i As Long) As String
    Select Case i
        Case 1: SlotLabel = "윗줄 왼쪽"
        Case 2: SlotLabel = "윗줄 가운데"
        Case 3: SlotLabel = "윗줄 오른쪽"
        Case 4: SlotLabel = "가운데 줄 왼쪽 (큰 칸)"
        Case 5: SlotLabel = "가운데 줄 오른쪽 위"
        Case 6: SlotLabel = "가운데 줄 오른쪽 아래"
        Case 7: SlotLabel = "아랫줄 왼쪽"
        Case 8: SlotLabel = "아랫줄 오른쪽"
    End Select
End Function

Private Sub Section(ByVal ws As Object, ByVal r As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Interior.Color = RGB(34, 87, 197)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ws.Cells(r, 1).Value = txt
End Sub

Private Sub InputCell(ByVal rg As Object)
    rg.Interior.Color = RGB(255, 249, 219)
    rg.Borders.LineStyle = 1
    rg.Borders.Color = RGB(198, 204, 213)
End Sub

Private Sub BuildSettingsSheet(ByVal ws As Object)
    Dim i As Long, k As Long, r As Long, sp As Variant, guide As Variant, btnY As Double, x As Double
    Dim areas As Variant, a As Variant
    For k = ws.Shapes.Count To 1 Step -1
        ws.Shapes(k).Delete
    Next
    ws.Cells.Clear
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = X_CENTER
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B:D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F:Q").ColumnWidth = 5.5
    ws.Columns("R").ColumnWidth = 3
    ws.Columns("S").ColumnWidth = 30

    ws.Range("A1").Value = "임원 일일 예정사항 - 설정"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "노란 칸을 고친 뒤 [대시보드 새로 고침]을 누르세요. 일정은 임원별 시트(시트 이름 = 직위 성명)에 씁니다."
    ws.Range("A2").Font.Color = RGB(93, 104, 117)

    Section ws, 4, "문서 정보"
    ws.Range("A5").Value = "제목"
    ws.Range("B5").Value = "임원 일일 예정사항"
    ws.Range("A6").Value = "날짜"
    ws.Range("B6").Value = Date
    ws.Range("B6").NumberFormat = "yyyy-mm-dd"
    ws.Range("B6").HorizontalAlignment = -4131
    ws.Range("A7").Value = "왼쪽 위 문구 (작성 부서 등)"
    ws.Range("A8").Value = "하단 문구 (비우면 인쇄 안 함)"
    InputCell ws.Range("B5:D5")
    InputCell ws.Range("B6")
    InputCell ws.Range("B7:D7")
    InputCell ws.Range("B8:D8")
    AddName "ED_Title", "$B$5"
    AddName "ED_Date", "$B$6"
    AddName "ED_Sub", "$B$7"
    AddName "ED_Foot", "$B$8"

    Section ws, 10, "배치 - 자리마다 직위 · 성명 (자리를 바꾸려면 직위·성명을 바꿔 적기)"
    ws.Range("A11:C11").Value = Array("자리", "직위", "성명")
    ws.Range("A11:C11").Font.Bold = True
    ws.Range("A11:C11").Interior.Color = RGB(231, 233, 236)
    sp = Array("부사장", "이정훈", "부사장", "박서연", "전무", "최영호", "대표이사", "김민수", _
               "전무", "정다은", "상무", "강태준", "상무", "윤지훈", "상무", "한소영")
    For i = 1 To N_EX
        r = 11 + i
        ws.Cells(r, 1).Value = i & ".  " & SlotLabel(i)
        ws.Cells(r, 2).Value = sp((i - 1) * 2)
        ws.Cells(r, 3).Value = sp((i - 1) * 2 + 1)
    Next
    InputCell ws.Range("B12:C19")
    AddName "ED_Execs", "$B$12:$C$19"

    ' 배치 그림 (12칸 × 8줄)
    ws.Range("F11").Value = "배치 그림 (번호 = 자리)"
    ws.Range("F11").Font.Bold = True
    areas = Array("F12:I13", "J12:M13", "N12:Q13", "F14:M17", "N14:Q15", "N16:Q17", "F18:K19", "L18:Q19")
    For i = 1 To N_EX
        ws.Range(areas(i - 1)).Cells(1, 1).Formula = "=""" & i & """&CHAR(10)&TRIM(B" & (11 + i) & "&"" ""&C" & (11 + i) & ")"
        With ws.Range(areas(i - 1))
            .Merge
            .HorizontalAlignment = X_CENTER
            .VerticalAlignment = X_CENTER
            .WrapText = True
            .Interior.Color = RGB(233, 240, 253)
            .Borders.LineStyle = 1
            .Borders.Color = RGB(34, 87, 197)
            .Font.Size = 9
        End With
    Next
    ws.Range("A12:Q19").RowHeight = 18

    Section ws, 21, "서식"
    ws.Range("A22").Value = "용지 방향"
    ws.Range("B22").Value = "세로"
    ws.Range("A23").Value = "여백 (mm, 5~25)"
    ws.Range("B23").Value = 10
    ws.Range("A24").Value = "글꼴"
    ws.Range("B24").Value = "맑은 고딕"
    ws.Range("A25").Value = "글자크기 최대 (pt)"
    ws.Range("B25").Value = 12
    ws.Range("A26").Value = "글자크기 최소 (pt)"
    ws.Range("B26").Value = 7
    ws.Range("A27").Value = "일정이 없는 칸에 쓸 문구"
    ws.Range("B27").Value = "일정 없음"
    InputCell ws.Range("B22:B27")
    ws.Range("B22:B27").HorizontalAlignment = -4131
    On Error Resume Next
    With ws.Range("B22").Validation
        .Delete
        .Add Type:=X_VALID_LIST, AlertStyle:=1, Operator:=1, Formula1:="세로,가로"
    End With
    With ws.Range("B24").Validation
        .Delete
        .Add Type:=X_VALID_LIST, AlertStyle:=2, Operator:=1, Formula1:="맑은 고딕,나눔고딕,돋움,굴림,바탕"
    End With
    On Error GoTo 0
    ws.Range("C25").Value = "8칸 모두 같은 글자크기로, 이 범위 안에서 A4 한 장을 꽉 채우는 가장 큰 크기를 고릅니다."
    ws.Range("C25").Font.Color = RGB(93, 104, 117)
    AddName "ED_Orient", "$B$22"
    AddName "ED_Margin", "$B$23"
    AddName "ED_Font", "$B$24"
    AddName "ED_FontMax", "$B$25"
    AddName "ED_FontMin", "$B$26"
    AddName "ED_Empty", "$B$27"

    Section ws, 29, "칸 크기 (비율 - 숫자가 클수록 넓게 · 높게)"
    ws.Range("A30").Value = "윗줄 폭 (1 · 2 · 3)"
    ws.Range("B30:D30").Value = Array(1, 1, 1)
    ws.Range("A31").Value = "가운데 줄 폭 (4 · 5와 6)"
    ws.Range("B31:C31").Value = Array(2, 1)
    ws.Range("A32").Value = "아랫줄 폭 (7 · 8)"
    ws.Range("B32:C32").Value = Array(1, 1)
    ws.Range("A33").Value = "줄 높이 (윗 · 가운데 · 아랫줄) - 비우면 자동"
    ws.Range("A34").Value = "5번 칸 높이 비율 (0.1~0.9) - 비우면 자동"
    InputCell ws.Range("B30:D30")
    InputCell ws.Range("B31:C31")
    InputCell ws.Range("B32:C32")
    InputCell ws.Range("B33:D33")
    InputCell ws.Range("B34")
    AddName "ED_ColsTop", "$B$30:$D$30"
    AddName "ED_ColsMid", "$B$31:$C$31"
    AddName "ED_ColsBot", "$B$32:$C$32"
    AddName "ED_Rows", "$B$33:$D$33"
    AddName "ED_Split", "$B$34"

    Section ws, 36, "결과"
    AddName "ED_Status", "$A$37"
    ws.Range("A37").Font.Bold = True

    Section ws, 39, "입력 방법 (임원별 시트 - 한 줄에 일정 하나)"
    guide = Array( _
        "시간 칸: 09:00 · 14:00~16:00 · 9:30 · 9시 반 · 오후 2시 · 종일 · 오전 → 09:00 꼴로 맞춰 줍니다.", _
        "내용 칸 끝에 @장소 → 내용 (장소)   예) 주간 경영회의 @대회의실", _
        "내용 칸 안에서 줄 바꿈(Alt+Enter) 후 * ... → 작은 글씨 부가 설명,  - ... → 하위 항목", _
        "내용이 * 로 시작하는 줄 → 바로 위 일정의 작은 글씨 부가 설명", _
        "비고 칸 → 그 일정 밑 작은 글씨 부가 설명(*)  /  ※ 참고 칸 → 그 임원 칸 맨 밑 ※ (어느 줄에 써도 됨)", _
        "글자 꾸밈: **굵게**  !!빨간 강조!!  __밑줄__  ==형광펜==", _
        "날짜 열을 넣으면 위 '날짜'와 같은 줄만 씁니다. 시트를 지우거나 이름을 바꾸면 성명으로 찾아 다시 만듭니다.", _
        "두레이 공동편집: [공동편집 양식 만들기]로 만든 파일을 올려 함께 쓰고, 내려받아 [엑셀 불러오기].", _
        "인쇄는 [인쇄 미리보기] 또는 [PDF로 저장]. 항상 A4 한 장에 맞춰 찍힙니다. 단축키 Ctrl+Shift+R = 새로 고침")
    For i = 0 To UBound(guide)
        ws.Cells(40 + i, 1).Value = "· " & guide(i)
    Next

    ' 단추
    x = ws.Columns("S").Left + 4
    btnY = ws.Rows(4).Top
    AddButton ws, "BTN_1", "대시보드 새로 고침 (Ctrl+Shift+R)", "RefreshDashboard", x, btnY, 190
    AddButton ws, "BTN_2", "엑셀 불러오기 (두레이 파일)", "ImportExcel", x, btnY + 30, 190
    AddButton ws, "BTN_3", "공동편집 양식 만들기", "MakeCoEditTemplate", x, btnY + 60, 190
    AddButton ws, "BTN_4", "인쇄 미리보기", "PreviewDashboard", x, btnY + 90, 190
    AddButton ws, "BTN_5", "PDF로 저장", "SaveDashboardPDF", x, btnY + 120, 190
    AddButton ws, "BTN_6", "일정 모두 비우기", "ClearSchedules", x, btnY + 150, 190

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("B5").Select
End Sub

Private Sub AddButton(ByVal ws As Object, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                      ByVal x As Double, ByVal y As Double, ByVal w As Double)
    Dim b As Object
    On Error Resume Next   ' 단추를 못 만들어도 Alt+F8 로 실행할 수 있으니 멈추지 않음
    Set b = ws.Buttons.Add(x, y, w, 24)
    If b Is Nothing Then Exit Sub
    b.Name = nm
    b.Caption = cap
    b.OnAction = mac
    b.Placement = X_FREE
    b.PrintObject = False
End Sub

Private Sub FormatInputSheet(ByVal ws As Object)
    With ws
        .Range("A1:D1").Value = Array("시간", "내용", "비고 (작은 글씨 *)", "※ 참고")
        With .Range("A1:D1")
            .Font.Bold = True
            .Interior.Color = RGB(231, 233, 236)
            .HorizontalAlignment = X_CENTER
            .Borders.LineStyle = 1
            .Borders.Color = RGB(198, 204, 213)
        End With
        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = 54
        .Columns("C").ColumnWidth = 32
        .Columns("D").ColumnWidth = 34
        .Columns("A").NumberFormat = "@"
        .Columns("A:D").VerticalAlignment = X_TOP
        .Columns("B:D").WrapText = True
    End With
    FreezeTop ws
End Sub

Private Sub FreezeTop(ByVal ws As Object)
    Dim cur As Object
    On Error Resume Next
    Set cur = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    If Not cur Is Nothing Then cur.Activate
End Sub

Private Sub WriteGuideSheet(ByVal ws As Object)
    Dim g As Variant, i As Long
    g = Array( _
        "1. 임원마다 시트가 하나씩 있습니다. 시트 이름(직위 성명)으로 임원을 찾으니 바꾸지 마세요.", _
        "2. 한 줄에 일정 하나: 시간 / 내용을 각자 칸에 씁니다. 시간은 09:00 또는 14:00~15:30처럼 씁니다(종일·오전도 됩니다).", _
        "3. 비고 칸에 쓴 것은 그 일정 밑에 작은 글씨 부가 설명(*)으로 붙습니다.", _
        "4. 내용 칸 안에서 줄을 바꿔(Alt+Enter) ""* ""로 시작하면 부가 설명, ""- ""로 시작하면 하위 항목입니다. 내용이 ""* ""로 시작하는 줄은 바로 위 일정에 붙습니다.", _
        "5. ※ 참고 칸에 쓴 것은 그 임원 칸 맨 밑에 ※ 표시로 인쇄됩니다. 어느 줄에 써도 됩니다.", _
        "6. 여러 날을 한 시트에 쓰려면 맨 앞에 ""날짜"" 열을 넣으세요. 문서 날짜에 맞는 줄만 가져옵니다.", _
        "7. 다 쓰면 담당자가 파일을 내려받아 대시보드 파일의 [엑셀 불러오기]로 가져옵니다.")
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 54
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 34
    ws.Range("A1").Value = "임원 일일 예정사항 - 공동편집 양식"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    For i = 0 To UBound(g)
        ws.Cells(3 + i, 1).Value = g(i)
    Next
    ws.Range("A11").Value = "예시"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12:D12").Value = Array("시간", "내용", "비고 (작은 글씨 *)", "※ 참고")
    ws.Range("A12:D12").Font.Bold = True
    ws.Range("A12:D12").Interior.Color = RGB(231, 233, 236)
    ws.Range("A13:A16").NumberFormat = "@"
    ws.Range("A13:D13").Value = Array("08:30", "임원 조찬 간담회 @20층 라운지", "참석: 부사장 이하 임원 전원", "11/10(화) 싱가포르 출장")
    ws.Range("A14:D14").Value = Array("10:00", "주간 경영회의 @대회의실" & vbLf & "- 안건: 4분기 실적 전망", "", "")
    ws.Range("A15:D15").Value = Array("14:00~16:00", "노사협의회 정기회의", "노측 6명, 사측 6명", "")
    ws.Range("A16:D16").Value = Array("", "* 회의 후 현장 점검", "", "")
    ws.Range("A12:D16").Borders.LineStyle = 1
    ws.Range("A13:D16").WrapText = True
    ws.Range("A13:D16").VerticalAlignment = X_TOP
End Sub

' 예시 일정 (시간|내용|비고|※ 참고, 줄은 ^ 로 나눔, 내용 안 줄 바꿈은 \n)
Private Function SampleRows(ByVal i As Long) As String
    Select Case i
        Case 1: SampleRows = "09:00|인사위원회 @소회의실|하반기 승진 심사|16:00 이후 사내 대기^10:00|주간 경영회의 @대회의실||^" & _
                             "13:30|감사팀 업무보고||^15:00|외부 감사인 미팅 @접견실||"
        Case 2: SampleRows = "08:30|임원 조찬 간담회 @20층 라운지||현장 점검 후 직퇴^10:00|주간 경영회의 @대회의실||^" & _
                             "14:00~17:00|영업 현장 점검 (부산지사)|동행: 영업본부장, 영남지역팀장|"
        Case 3: SampleRows = "10:00|주간 경영회의 @대회의실||^11:00|신사업 TF 회의|2027년 신규 진출 시장 검토|^" & _
                             "15:00|협력사 대표 면담 @접견실||"
        Case 4: SampleRows = "08:30|임원 조찬 간담회 @20층 라운지|참석: 부사장 이하 임원 전원|11/10(화) 싱가포르 출장 (08:00 인천 출발)^" & _
                             "10:00|주간 경영회의 @대회의실\n- 안건: 4분기 실적 전망, 2027년 사업계획 방향||^" & _
                             "11:30|신규 고객사 대표 접견 @접견실||^12:00|오찬 (고객사 대표)||^" & _
                             "14:00|노사협의회 정기회의 @3층 회의실|노측 6명, 사측 6명|^" & _
                             "16:00|투자심의위원회 @대회의실|안건: 제2공장 설비 증설 (!!승인 필요!!)|^" & _
                             "17:30|결재||^19:00|업계 CEO 만찬||"
        Case 5: SampleRows = "09:30|재무팀 자금 현황 보고||오후 외부 일정 후 복귀^10:00|주간 경영회의 @대회의실||^" & _
                             "14:00|주거래 은행 미팅 @외부||"
        Case 6: SampleRows = "종일|해외 바이어 상담 (서울사무소)|베트남 바이어 3개사|10:00 경영회의 불참 (정다은 전무 대리 보고)"
        Case 7: SampleRows = "10:00|주간 경영회의 @대회의실||^13:00|연구소 과제 점검 @기술연구소|하반기 중점과제 5건|^" & _
                             "16:30|품질 회의||"
        Case 8: SampleRows = "09:00|홍보 전략 회의||내일 오전 반차^11:00|언론사 인터뷰 @홍보실|주제: 친환경 경영|^" & _
                             "15:00|사회공헌 위원회||"
    End Select
End Function

Private Sub WriteSample(ByVal ws As Object, ByVal i As Long)
    Dim rowsArr As Variant, f As Variant, r As Long, c As Long
    rowsArr = Split(SampleRows(i), "^")
    For r = 0 To UBound(rowsArr)
        f = Split(rowsArr(r), "|")
        For c = 0 To UBound(f)
            If c <= 3 Then ws.Cells(2 + r, 1 + c).Value = CellText(Replace(f(c), "\n", vbLf))
        Next
    Next
End Sub

Private Sub DeleteBlankDefaultSheets()
    Dim k As Long, ws As Object
    Application.DisplayAlerts = False
    On Error Resume Next
    For k = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(k)
        If (ws.Name Like "Sheet#*" Or ws.Name Like "시트#*") And ThisWorkbook.Sheets.Count > 1 Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End If
    Next
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

'=====================================================================
'  임원별 시트 찾기·만들기
'=====================================================================
Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    SheetExists = Not ws Is Nothing
End Function

Private Function IsSpecialSheet(ByVal nm As String) As Boolean
    IsSpecialSheet = (nm = SH_DASH Or nm = SH_CFG Or nm = SH_GUIDE)
End Function

' 불러올 때 빼는 시트 (작성방법·안내 등)
Private Function IsSkipSheet(ByVal nm As String) As Boolean
    Dim k As String, w As Variant, i As Long
    k = LCase$(Replace(Replace(nm, " ", ""), UC(12288), ""))
    If nm = SH_DASH Or nm = SH_CFG Then IsSkipSheet = True: Exit Function
    w = Array("작성방법", "작성요령", "작성예시", "안내", "도움말", "사용법", "예시", "목록", "설정", "readme", "guide", "help")
    For i = 0 To UBound(w)
        If Left$(k, Len(w(i))) = w(i) Then IsSkipSheet = True: Exit Function
    Next
End Function

Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, k As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For k = 0 To UBound(bad)
        s = Replace(s, bad(k), "_")
    Next
    s = JTrim(s)
    Do While Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 31 Then s = JTrim(Left$(s, 31))
    If s = "" Then s = "임원"
    SafeSheetName = s
End Function

Private Function DesiredSheetName(ByVal i As Long) As String
    Dim s As String
    s = JTrim(gPos(i) & " " & gName(i))
    If s = "" Then s = "자리" & i
    DesiredSheetName = SafeSheetName(s)
End Function

' 자리마다 시트를 찾고(시트 이름 → 성명이 든 시트), 없으면 만들고, 이름을 "직위 성명"으로 맞춥니다.
Private Sub SyncExecSheets()
    Dim i As Long, ws As Object, f As Object, want As String, used As String, prev As Object
    Dim curSheet As Object
    On Error Resume Next
    Set curSheet = ActiveSheet
    On Error GoTo 0
    used = "|"
    For i = 1 To N_EX
        Set f = Nothing
        want = DesiredSheetName(i)
        If SheetExists(want) Then
            If InStr(1, used, "|" & want & "|", vbTextCompare) = 0 Then Set f = ThisWorkbook.Worksheets(want)
        End If
        If f Is Nothing And gName(i) <> "" Then
            For Each ws In ThisWorkbook.Worksheets
                If Not IsSpecialSheet(ws.Name) And InStr(1, used, "|" & ws.Name & "|", vbTextCompare) = 0 Then
                    If InStr(1, NormKey(ws.Name), NormKey(gName(i)), vbBinaryCompare) > 0 Then
                        Set f = ws
                        Exit For
                    End If
                End If
            Next
        End If
        If f Is Nothing Then
            Set f = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            If SheetExists(want) Then
                f.Name = SafeSheetName(Left$(want, 27) & "(" & i & ")")
            Else
                f.Name = want
            End If
            FormatInputSheet f
        ElseIf f.Name <> want And Not SheetExists(want) Then
            On Error Resume Next
            f.Name = want
            On Error GoTo 0
        End If
        gSheetOf(i) = f.Name
        used = used & f.Name & "|"
    Next
    ' 시트 순서: 대시보드, 설정, 자리 1~8
    On Error Resume Next
    Set prev = ThisWorkbook.Worksheets(SH_CFG)
    For i = 1 To N_EX
        Set ws = ThisWorkbook.Worksheets(gSheetOf(i))
        If ws.Index <> prev.Index + 1 Then ws.Move After:=prev
        Set prev = ws
    Next
    If Not curSheet Is Nothing Then curSheet.Activate
    On Error GoTo 0
End Sub

'=====================================================================
'  시트 읽기 → 일정 줄
'=====================================================================
Private Sub ResetRead()
    rN = 0: grN = 0
    ReDim rGrp(1 To 64): ReDim rDate(1 To 64): ReDim rTime(1 To 64)
    ReDim rCont(1 To 64): ReDim rStar(1 To 64): ReDim rRef(1 To 64)
    ReDim grLabel(1 To 16): ReDim grSheet(1 To 16): ReDim grTitle(1 To 16)
    ReDim grKey(1 To 16): ReDim grHasExec(1 To 16)
End Sub

Private Function AddGroup(ByVal key As String, ByVal lbl As String, ByVal sheetNm As String, ByVal title As String, ByVal hasExec As Boolean, ByVal fromG As Long) As Long
    Dim g As Long
    For g = fromG To grN
        If grKey(g) = key Then AddGroup = g: Exit Function
    Next
    grN = grN + 1
    If grN > UBound(grLabel) Then
        ReDim Preserve grLabel(1 To grN * 2): ReDim Preserve grSheet(1 To grN * 2): ReDim Preserve grTitle(1 To grN * 2)
        ReDim Preserve grKey(1 To grN * 2): ReDim Preserve grHasExec(1 To grN * 2)
    End If
    grKey(grN) = key: grLabel(grN) = lbl: grSheet(grN) = sheetNm: grTitle(grN) = title: grHasExec(grN) = hasExec
    AddGroup = grN
End Function

Private Sub AddRow(ByVal g As Long, ByVal dt As String, ByVal tm As String, ByVal ct As String, ByVal st As String, ByVal rf As String)
    rN = rN + 1
    If rN > UBound(rGrp) Then
        ReDim Preserve rGrp(1 To rN * 2): ReDim Preserve rDate(1 To rN * 2): ReDim Preserve rTime(1 To rN * 2)
        ReDim Preserve rCont(1 To rN * 2): ReDim Preserve rStar(1 To rN * 2): ReDim Preserve rRef(1 To rN * 2)
    End If
    rGrp(rN) = g: rDate(rN) = dt: rTime(rN) = tm: rCont(rN) = ct: rStar(rN) = st: rRef(rN) = rf
End Sub

Private Function LastRowOf(ByVal ws As Object) As Long
    Dim f As Object
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", LookIn:=X_FORMULAS, LookAt:=X_PART, SearchOrder:=X_BYROWS, SearchDirection:=X_PREV)
    On Error GoTo 0
    If Not f Is Nothing Then LastRowOf = f.Row
End Function

Private Function LastColOf(ByVal ws As Object) As Long
    Dim f As Object
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", LookIn:=X_FORMULAS, LookAt:=X_PART, SearchOrder:=X_BYCOLS, SearchDirection:=X_PREV)
    On Error GoTo 0
    If Not f Is Nothing Then LastColOf = f.Column
End Function

Private Function TimeOfFrac(ByVal x As Double) As String
    Dim mins As Long
    mins = CLng(x * 1440)
    TimeOfFrac = Pad2(mins \ 60) & ":" & Pad2(mins Mod 60)
End Function

' 칸 값 → 글자 (시간 칸은 엑셀 시간 값도 09:30 꼴로, 날짜 칸은 yyyy-mm-dd)
Private Function CellStr(ByVal v As Variant, ByVal role As String) As String
    Dim s As String, x As Double, isNum As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    isNum = (VarType(v) = vbDate)
    If Not isNum And VarType(v) <> vbString And VarType(v) <> vbBoolean Then isNum = IsNumeric(v)
    If isNum Then
        x = CDbl(v)
        If role = "time" Then
            If x >= 0 And x < 1 Then
                CellStr = TimeOfFrac(x)
                Exit Function
            ElseIf x - Int(x) > 0.00001 Then
                CellStr = TimeOfFrac(x - Int(x))
                Exit Function
            End If
        ElseIf role = "date" Then
            If x >= 1 Then
                CellStr = IsoOf(CDate(Int(x)))
                Exit Function
            End If
        End If
    End If
    s = CStr(v)
    s = Replace(s, UC(160), " ")
    s = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    CellStr = JTrim(s)
End Function

' 엑셀 머리글 → 칸의 쓰임 (time content place ref star exec pos date skip)
Private Function HeadRole(ByVal h As String, ByRef lbl As String) As String
    Dim k As String
    lbl = ""
    k = LCase$(RemoveWs(RemoveParens(h)))
    Select Case k
        Case "": HeadRole = "skip"
        Case "시간", "시각", "일시", "time": HeadRole = "time"
        Case "내용", "일정", "일정내용", "주요일정", "주요내용", "예정사항", "주요예정사항", "행사", "행사명", "업무", "content": HeadRole = "content"
        Case "장소", "위치", "place": HeadRole = "place"
        Case UC(8251), "참고", "참고사항", UC(8251) & "참고", UC(8251) & "참고사항": HeadRole = "ref"
        Case "비고", "설명", "부가설명", "메모", "note": HeadRole = "star"
        Case "성명", "이름", "임원", "임원명", "name": HeadRole = "exec"
        Case "직위", "직책", "직급": HeadRole = "pos"
        Case "날짜", "일자", "date": HeadRole = "date"
        Case "순번", "번호", "no", "no.", "연번": HeadRole = "skip"
        Case Else
            HeadRole = "star"
            lbl = JTrim(h)
    End Select
End Function

Private Function RemoveParens(ByVal s As String) As String
    Dim p As Long, q As Long
    Do
        p = InStr(1, s, "(")
        If p = 0 Then Exit Do
        q = InStr(p + 1, s, ")")
        If q = 0 Then Exit Do
        s = Left$(s, p - 1) & Mid$(s, q + 1)
    Loop
    RemoveParens = s
End Function

Private Function RemoveWs(ByVal s As String) As String
    Dim i As Long, c As String, o As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsWs(c) Then o = o & c
    Next
    RemoveWs = o
End Function

Private Function IsHeaderCells(ByRef cv() As String, ByVal n As Long) As Boolean
    Dim c As Long, hasC As Boolean, hasT As Boolean, rl As String, lbl As String
    For c = 1 To n
        rl = HeadRole(cv(c), lbl)
        If rl = "content" Then hasC = True
        If rl = "time" Or rl = "exec" Then hasT = True
    Next
    IsHeaderCells = hasC And hasT
End Function

' 시트 하나 → 묶음·줄. 성명 열이 있으면 임원별로, 없으면 시트 하나가 한 임원
Private Sub ReadSheet(ByVal ws As Object, ByVal y0 As Long)
    Dim lastR As Long, lastC As Long, data As Variant, r As Long, c As Long, hr As Long, lim As Long
    Dim roles() As String, lbls() As String, cv() As String, hasRoles As Boolean
    Dim cE As Long, cP As Long, cD As Long, curExec As String, curPos As String, curDate As String
    Dim g As Long, fromG As Long, title As String, nonEmpty As Boolean, dt As String
    Dim tm As String, ct As String, st As String, rf As String, pl As String, k As Long

    lastR = LastRowOf(ws)
    lastC = LastColOf(ws)
    If lastR = 0 Or lastC = 0 Then Exit Sub
    If lastC > 30 Then lastC = 30
    If lastR > 3000 Then lastR = 3000
    If lastR = 1 And lastC = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = ws.Cells(1, 1).Value
    Else
        data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    End If
    ReDim roles(1 To lastC): ReDim lbls(1 To lastC): ReDim cv(1 To lastC)

    ' 머리글 줄 찾기 (위 20줄 안)
    hr = 0
    lim = lastR
    If lim > 20 Then lim = 20
    For r = 1 To lim
        For c = 1 To lastC
            cv(c) = CellStr(data(r, c), "")
        Next
        If IsHeaderCells(cv, lastC) Then
            hr = r
            Exit For
        End If
    Next
    hasRoles = (hr > 0)
    cE = 0: cP = 0: cD = 0
    If hasRoles Then
        For c = 1 To lastC
            cv(c) = CellStr(data(hr, c), "")
            roles(c) = HeadRole(cv(c), lbls(c))
            If roles(c) = "exec" And cE = 0 Then cE = c
            If roles(c) = "pos" And cP = 0 Then cP = c
            If roles(c) = "date" And cD = 0 Then cD = c
        Next
        title = ""
        For r = 1 To hr - 1
            For c = 1 To lastC
                If CellStr(data(r, c), "") <> "" Then title = JTrim(title & " " & CellStr(data(r, c), ""))
            Next
        Next
    End If

    fromG = grN + 1
    For r = hr + 1 To lastR
        nonEmpty = False
        For c = 1 To lastC
            If hasRoles Then
                cv(c) = CellStr(data(r, c), roles(c))
            ElseIf c = 1 Then
                cv(c) = CellStr(data(r, c), "time")
            Else
                cv(c) = CellStr(data(r, c), "")
            End If
            If cv(c) <> "" Then nonEmpty = True
        Next
        If nonEmpty Then
            If hasRoles Then
                If IsHeaderCells(cv, lastC) Then nonEmpty = False
            End If
        End If
        If nonEmpty Then
            ' 병합된 칸(빈칸)은 위 값을 이어 씀
            If cE > 0 Then
                If cv(cE) <> "" Then
                    curExec = cv(cE)
                    If cP > 0 Then curPos = cv(cP)
                ElseIf cP > 0 Then
                    If cv(cP) <> "" Then curPos = cv(cP)
                End If
            End If
            If cD > 0 Then
                If cv(cD) <> "" Then
                    dt = ParseLooseDate(cv(cD), y0)
                    If dt <> "" Then curDate = dt
                End If
            End If
            If cE > 0 Then
                g = AddGroup("e:" & curPos & "|" & curExec, JTrim(curPos & " " & curExec), ws.Name, title, True, fromG)
            Else
                g = AddGroup("s", ws.Name, ws.Name, title, False, fromG)
            End If
            ' 한 줄 → 시간 / 내용 / 비고 / ※ 참고
            tm = "": ct = "": st = "": rf = "": pl = ""
            If hasRoles Then
                For c = 1 To lastC
                    If cv(c) <> "" Then
                        Select Case roles(c)
                            Case "time": tm = cv(c)
                            Case "content": AddLine ct, cv(c)
                            Case "place": pl = cv(c)
                            Case "ref": AddLine rf, cv(c)
                            Case "star"
                                If lbls(c) <> "" Then
                                    AddLine st, lbls(c) & ": " & cv(c)
                                Else
                                    AddLine st, cv(c)
                                End If
                        End Select
                    End If
                Next
            Else
                k = lastC
                Do While k >= 1
                    If cv(k) <> "" Then Exit Do
                    k = k - 1
                Loop
                c = 1
                If k >= 2 Then
                    If cv(1) = "" Or IsTimeCell(cv(1)) Then
                        tm = cv(1)
                        c = 2
                    End If
                End If
                ct = cv(c)
                For c = c + 1 To k
                    If cv(c) <> "" Then AddLine st, cv(c)
                Next
            End If
            ct = TidyContent(ct, pl)
            If tm <> "" Or ct <> "" Or st <> "" Or rf <> "" Then
                If cD = 0 Then curDate = ""
                AddRow g, curDate, tm, ct, st, rf
            End If
        End If
    Next
End Sub

' 내용 칸: 빈 줄을 빼고, 장소 열이 있으면 첫 줄 끝에 " @장소"
Private Function TidyContent(ByVal ct As String, ByVal pl As String) As String
    Dim arr As Variant, i As Long, o As String, l As String, first As Boolean
    arr = Split(ct, vbLf)
    first = True
    For i = LBound(arr) To UBound(arr)
        l = JTrim(arr(i))
        If l <> "" Then
            If first And pl <> "" Then
                If IsStarC(Left$(l, 1)) Or Left$(l, 1) = UC(8251) Then
                    AddLine o, "@" & pl
                    AddLine o, l
                Else
                    AddLine o, l & " @" & pl
                End If
            Else
                AddLine o, l
            End If
            first = False
        End If
    Next
    If first And pl <> "" Then o = "@" & pl
    TidyContent = o
End Function

Private Sub AddLine(ByRef buf As String, ByVal l As String)
    If Len(buf) > 0 Then buf = buf & vbLf & l Else buf = l
End Sub

' 엑셀 한 줄(시간·내용·비고·※ 참고) → 일정 줄(sched)·※ 참고 줄(refs)
Private Sub RowToLines(ByVal tmRaw As String, ByVal cont As String, ByVal star As String, ByVal refTxt As String, _
                       ByRef sched As String, ByRef refs As String)
    Dim arr As Variant, ln() As String, n As Long, i As Long, l As String, first As String, startI As Long
    Dim tm As String, sp As String, rest As String
    arr = Split(refTxt, vbLf)
    For i = LBound(arr) To UBound(arr)
        l = JTrim(arr(i))
        If l <> "" Then
            If Left$(l, 1) = UC(8251) Then l = LTrimWs(Mid$(l, 2))
            AddLine refs, l
        End If
    Next
    n = 0
    ReDim ln(0 To 0)
    arr = Split(cont, vbLf)
    For i = LBound(arr) To UBound(arr)
        l = JTrim(arr(i))
        If l <> "" Then
            ReDim Preserve ln(0 To n)
            ln(n) = l
            n = n + 1
        End If
    Next
    first = "": startI = 0
    If n > 0 Then
        first = ln(0): startI = 1
        If IsStarC(Left$(first, 1)) Or Left$(first, 1) = UC(8251) Then
            first = "": startI = 0
        End If
    End If
    tm = ""
    If tmRaw <> "" Then
        If SplitTime(tmRaw, sp, rest) Then
            If rest = "" Then tm = sp & " " Else tm = tmRaw & vbTab
        Else
            tm = tmRaw & vbTab
        End If
    End If
    If first <> "" Or tm <> "" Then
        l = tm & first
        If Right$(l, 1) = " " Then l = Left$(l, Len(l) - 1)
        AddLine sched, l
    End If
    For i = startI To n - 1
        l = ln(i)
        If Left$(l, 1) = UC(8251) Then
            AddLine refs, LTrimWs(Mid$(l, 2))
        ElseIf IsStarC(Left$(l, 1)) Or Left$(l, 1) = "-" Or Left$(l, 1) = UC(183) Or Left$(l, 1) = UC(8226) Then
            AddLine sched, l
        Else
            AddLine sched, "  " & l
        End If
    Next
    arr = Split(star, vbLf)
    For i = LBound(arr) To UBound(arr)
        l = JTrim(arr(i))
        If l <> "" Then
            If IsStarC(Left$(l, 1)) Then l = LTrimWs(Mid$(l, 2))
            AddLine sched, "* " & l
        End If
    Next
End Sub

' 자리 1~8: 시트를 읽어 칸 내용을 만듭니다.
Private Sub LoadAllSchedules()
    Dim i As Long, k As Long, ws As Object, sched As String, refs As String, arr As Variant
    For i = 1 To N_EX
        gN(i) = 0: gNR(i) = 0
        sched = "": refs = ""
        Set ws = ThisWorkbook.Worksheets(gSheetOf(i))
        ResetRead
        ReadSheet ws, Year(sDate)
        For k = 1 To rN
            If rDate(k) = "" Or rDate(k) = sDateISO Then
                RowToLines rTime(k), rCont(k), rStar(k), rRef(k), sched, refs
            End If
        Next
        arr = Split(refs, vbLf)
        For k = LBound(arr) To UBound(arr)
            If JTrim(arr(k)) <> "" Then AddRef i, JTrim(arr(k))
        Next
        ParseSched i, sched
        If gN(i) = 0 And sEmpty <> "" Then AddPara i, "E", "", sEmpty
    Next
End Sub

' 가져온 줄을 임원 시트에 씁니다 (시간·내용·비고·※ 참고 네 칸)
Private Sub WriteImported(ByVal ws As Object, ByVal i As Long, ByRef tgt() As Long, ByVal D As String)
    Dim k As Long, n As Long, out() As Variant, lr As Long
    n = 0
    For k = 1 To rN
        If tgt(rGrp(k)) = i And (rDate(k) = "" Or D = "" Or rDate(k) = D) Then n = n + 1
    Next
    lr = LastRowOf(ws)
    If lr >= 2 Then ws.Range(ws.Rows(2), ws.Rows(lr)).ClearContents
    ws.Range("A1:D1").Value = Array("시간", "내용", "비고 (작은 글씨 *)", "※ 참고")
    If n = 0 Then Exit Sub
    ReDim out(1 To n, 1 To 4)
    n = 0
    For k = 1 To rN
        If tgt(rGrp(k)) = i And (rDate(k) = "" Or D = "" Or rDate(k) = D) Then
            n = n + 1
            out(n, 1) = CellText(rTime(k)): out(n, 2) = CellText(rCont(k))
            out(n, 3) = CellText(rStar(k)): out(n, 4) = CellText(rRef(k))
        End If
    Next
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + n, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + n, 4)).Value = out
End Sub

' 칸에 글자로 넣기: = + - @ 로 시작하면 엑셀이 수식으로 읽지 않게 앞에 ' 를 붙임
Private Function CellText(ByVal s As String) As String
    Select Case Left$(s, 1)
        Case "=", "+", "-", "@"
            CellText = "'" & s
        Case Else
            CellText = s
    End Select
End Function

Private Function SrcText(ByVal g As Long) As String
    If grHasExec(g) Then
        SrcText = "'" & grSheet(g) & "' 시트의 " & grLabel(g)
    Else
        SrcText = "'" & grSheet(g) & "' 시트"
    End If
End Function

' 이름표(시트 이름 등)로 임원 찾기: 성명이 들어 있으면 그 사람, 아니면 직위가 겹치지 않을 때만
Private Function MatchExec(ByVal lbl As String) As Long
    Dim L As String, i As Long, n As String, p As String, sc As Long, best As Long, bestSc As Long, tie As Boolean
    L = NormKey(lbl)
    If L = "" Then Exit Function
    For i = 1 To N_EX
        n = NormKey(gName(i))
        p = NormKey(gPos(i))
        sc = 0
        If n <> "" And InStr(1, L, n, vbBinaryCompare) > 0 Then
            sc = 10
            If p <> "" And InStr(1, L, p, vbBinaryCompare) > 0 Then sc = 12
        ElseIf p <> "" And L = p Then
            sc = 5
        ElseIf p <> "" And InStr(1, L, p, vbBinaryCompare) > 0 Then
            sc = 3
        End If
        If sc > bestSc Then
            bestSc = sc
            best = i
            tie = False
        ElseIf sc > 0 And sc = bestSc Then
            tie = True
        End If
    Next
    If tie Then MatchExec = 0 Else MatchExec = best
End Function

Private Function NormKey(ByVal s As String) As String
    Dim i As Long, c As String, o As String
    s = RemoveParens(s)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not IsWs(c) And c <> UC(183) And c <> "." And c <> "," And c <> "_" And c <> "-" Then o = o & c
    Next
    NormKey = LCase$(o)
End Function

'=====================================================================
'  일정 해석 (HTML 판 executive_schedule.html 과 같은 규칙)
'=====================================================================
Private Sub AddPara(ByVal b As Long, ByVal kind As String, ByVal tm As String, ByVal txt As String)
    If gN(b) >= MAXP Then Exit Sub
    gN(b) = gN(b) + 1
    gKind(b, gN(b)) = kind
    gTime(b, gN(b)) = tm
    gText(b, gN(b)) = txt
End Sub

Private Sub AddRef(ByVal b As Long, ByVal txt As String)
    If gNR(b) >= MAXR Then Exit Sub
    gNR(b) = gNR(b) + 1
    gRef(b, gNR(b)) = txt
End Sub

' 일정 줄들 → 칸 내용 (한 줄에 일정 하나)
Private Sub ParseSched(ByVal b As Long, ByVal txt As String)
    Dim arr As Variant, i As Long, raw As String, t As String, k As Long, pre As String, body As String
    Dim after As String, tm As String, rest As String, hasLast As Boolean, lk As String, c1 As String
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    arr = Split(txt, vbLf)
    hasLast = False
    For i = LBound(arr) To UBound(arr)
        raw = arr(i)
        t = JTrim(raw)
        If t = "" Then GoTo NextLine
        c1 = Left$(t, 1)
        ' ※ 참고 → 칸 맨 밑
        If c1 = UC(8251) Then
            AddRef b, LTrimWs(Mid$(t, 2))
            GoTo NextLine
        End If
        ' * 작은 글씨 부가 설명
        If IsStarC(c1) Then
            k = 1
            Do While IsStarC(Mid$(t, k, 1))
                k = k + 1
            Loop
            AddPara b, "S", "", LTrimWs(Mid$(t, k))
            hasLast = True
            GoTo NextLine
        End If
        ' 엑셀에서 온 "시간<Tab>내용": 탭 앞은 시간 칸
        k = InStr(1, t, vbTab)
        If k >= 2 And k <= 17 And Not IsBulletC(c1) Then
            pre = JTrim(Left$(t, k - 1))
            tm = pre
            If SplitTime(pre, after, rest) Then
                If rest = "" Then tm = after
            End If
            body = JTrim(CollapseTabs(Mid$(t, k + 1)))
            AddPara b, "M", tm, body
            hasLast = True
            GoTo NextLine
        End If
        t = CollapseTabs(t)
        ' 글머리 기호
        If IsBulletC(c1) Then
            after = LTrimWs(Mid$(t, 2))
            If SplitTime(after, tm, rest) Then
                AddPara b, "M", tm, rest
                hasLast = True
            ElseIf hasLast And IsDashC(c1) Then
                AddPara b, "D", "", after
            Else
                AddPara b, "M", "", after
                hasLast = True
            End If
            GoTo NextLine
        End If
        ' 시간으로 시작
        If SplitTime(t, tm, rest) Then
            AddPara b, "M", tm, rest
            hasLast = True
            GoTo NextLine
        End If
        ' 앞에 빈칸을 둔 줄: 바로 위 줄에 이어서
        If hasLast And gN(b) > 0 And (Left$(raw, 1) = " " Or Left$(raw, 1) = UC(12288)) Then
            lk = gKind(b, gN(b))
            If lk = "S" Or lk = "S2" Then
                AddPara b, "S2", "", t
            ElseIf lk = "D" Or lk = "D2" Then
                AddPara b, "D2", "", t
            Else
                AddPara b, "C", "", t
            End If
            GoTo NextLine
        End If
        AddPara b, "M", "", t
        hasLast = True
NextLine:
    Next
End Sub

' 줄 앞의 시간을 떼어 냅니다. 시간이면 True (tm = 09:00 또는 09:00~10:30, rest = 나머지)
Private Function SplitTime(ByVal s As String, ByRef tm As String, ByRef rest As String) As Boolean
    Dim p As Long, q As Long, qb As Long, ta As String, pa As Long, tb As String, pb As Long, hasB As Boolean
    Dim r As String, w As String, e As Long
    tm = "": rest = ""
    p = 1
    If TimePart(s, p, ta, pa, False) Then
        hasB = False
        q = SkipWs(s, p)
        If IsSepC(Mid$(s, q, 1)) Then
            qb = SkipWs(s, q + 1)
            If TimePart(s, qb, tb, pb, (ta <> "" And pa = 1)) Then
                If Not IsDigitC(Mid$(s, qb, 1)) Then
                    hasB = True
                    p = qb
                End If
            End If
        End If
        If hasB Or Not IsDigitC(Mid$(s, p, 1)) Then
            If ta <> "" And (Not hasB Or tb <> "") Then
                r = Mid$(s, SkipWs(s, p))
                If Len(r) > 0 Then
                    If InStr(1, "-:|," & UC(8211) & UC(8212), Left$(r, 1), vbBinaryCompare) > 0 Then r = LTrimWs(Mid$(r, 2))
                End If
                tm = ta
                If hasB Then tm = ta & "~" & tb
                rest = JTrim(r)
                SplitTime = True
                Exit Function
            End If
        End If
    End If
    ' 종일·오전 같은 낱말
    w = ""
    If Left$(s, 2) = "하루" Then
        q = SkipWs(s, 3)
        If Mid$(s, q, 2) = "종일" Then
            If q > 3 Then w = "하루 종일" Else w = "하루종일"
            e = q + 2
        End If
    End If
    If w = "" Then
        Select Case Left$(s, 2)
            Case "종일", "오전", "오후", "미정", "수시"
                w = Left$(s, 2)
                e = 3
        End Select
    End If
    If w <> "" Then
        If e > Len(s) Then
            tm = w
            SplitTime = True
        ElseIf IsWs(Mid$(s, e, 1)) Then
            tm = w
            rest = JTrim(Mid$(s, e))
            SplitTime = True
        End If
    End If
End Function

' 시간 조각 하나: [오전|오후|AM|PM] (H:MM[:SS] | H시[ M분| 반]). p는 끝난 다음 자리로 옮김.
' outT = "HH:MM" (시·분 값이 맞지 않으면 ""), outPm = -1 모름 / 0 오전 / 1 오후
Private Function TimePart(ByVal s As String, ByRef p As Long, ByRef outT As String, ByRef outPm As Long, ByVal pmDefault As Boolean) As Boolean
    Dim q As Long, q2 As Long, q3 As Long, q4 As Long, sv As Long, pm As Long, h As Long, mi As Long
    Dim nd As Long, nd2 As Long, c As String, ap As String, got As Boolean
    q = p: pm = -1
    ap = Mid$(s, q, 2)
    If ap = "오전" Or ap = "AM" Or ap = "am" Then
        pm = 0
        q = SkipWs(s, q + 2)
    ElseIf ap = "오후" Or ap = "PM" Or ap = "pm" Then
        pm = 1
        q = SkipWs(s, q + 2)
    End If
    nd = 0
    Do While IsDigitC(Mid$(s, q + nd, 1))
        nd = nd + 1
    Loop
    If nd < 1 Or nd > 2 Then Exit Function
    h = CLng(Mid$(s, q, nd))
    q2 = SkipWs(s, q + nd)
    c = Mid$(s, q2, 1)
    If c = ":" Or c = UC(65306) Then
        q2 = SkipWs(s, q2 + 1)
        If Not IsDigitC(Mid$(s, q2, 1)) Then Exit Function
        If Not IsDigitC(Mid$(s, q2 + 1, 1)) Then Exit Function
        mi = CLng(Mid$(s, q2, 2))
        q = q2 + 2
        q3 = SkipWs(s, q)
        c = Mid$(s, q3, 1)
        If c = ":" Or c = UC(65306) Then
            q3 = SkipWs(s, q3 + 1)
            If IsDigitC(Mid$(s, q3, 1)) And IsDigitC(Mid$(s, q3 + 1, 1)) Then q = q3 + 2
        End If
    ElseIf c = "시" Then
        q = q2 + 1
        sv = q
        mi = 0
        got = False
        q3 = SkipWs(s, q)
        nd2 = 0
        Do While IsDigitC(Mid$(s, q3 + nd2, 1))
            nd2 = nd2 + 1
        Loop
        If nd2 >= 1 And nd2 <= 2 Then
            q4 = SkipWs(s, q3 + nd2)
            If Mid$(s, q4, 1) = "분" Then
                mi = CLng(Mid$(s, q3, nd2))
                q = q4 + 1
                got = True
            End If
        ElseIf Mid$(s, q3, 1) = "반" Then
            mi = 30
            q = q3 + 1
            got = True
        End If
        If IsHangulC(Mid$(s, q, 1)) Then
            If Not got Then Exit Function
            q = sv: mi = 0
            If IsHangulC(Mid$(s, q, 1)) Then Exit Function
        End If
    Else
        Exit Function
    End If
    TimePart = True
    p = q
    If pm = -1 And pmDefault And h < 12 Then pm = 1
    If pm = 1 And h < 12 Then h = h + 12
    outPm = pm
    If h > 24 Or mi > 59 Then
        outT = ""
    Else
        outT = Pad2(h) & ":" & Pad2(mi)
    End If
End Function

Private Function IsTimeCell(ByVal v As String) As Boolean
    Dim tm As String, rest As String
    If SplitTime(JTrim(v), tm, rest) Then IsTimeCell = (rest = "")
End Function

' 날짜 글자 → yyyy-mm-dd (2026-11-10, 2026.11.10, 2026년 11월 10일, 11/10, 11월 10일)
Private Function ParseLooseDate(ByVal v As String, ByVal y0 As Long) As String
    Dim p As Long, n1 As String, n2 As String, n3 As String, c As String
    v = JTrim(v)
    If v = "" Then Exit Function
    p = 1
    n1 = Digits(v, p)
    If Len(n1) = 0 Then Exit Function
    p = SkipWs(v, p)
    c = Mid$(v, p, 1)
    If Len(n1) = 4 Then
        If c <> "-" And c <> "." And c <> "/" And c <> "년" Then Exit Function
        p = SkipWs(v, p + 1)
        n2 = Digits(v, p)
        If Len(n2) < 1 Or Len(n2) > 2 Then Exit Function
        p = SkipWs(v, p)
        c = Mid$(v, p, 1)
        If c <> "-" And c <> "." And c <> "/" And c <> "월" Then Exit Function
        p = SkipWs(v, p + 1)
        n3 = Digits(v, p)
        If Len(n3) < 1 Then Exit Function
        n3 = Left$(n3, 2)
        ParseLooseDate = MkDate(CLng(n1), CLng(n2), CLng(n3))
    ElseIf Len(n1) <= 2 Then
        If c <> "-" And c <> "." And c <> "/" And c <> "월" Then Exit Function
        p = SkipWs(v, p + 1)
        n2 = Digits(v, p)
        If Len(n2) < 1 Then Exit Function
        n2 = Left$(n2, 2)
        ParseLooseDate = MkDate(y0, CLng(n1), CLng(n2))
    End If
End Function

Private Function Digits(ByVal s As String, ByRef p As Long) As String
    Dim q As Long
    q = p
    Do While IsDigitC(Mid$(s, q, 1))
        q = q + 1
    Loop
    Digits = Mid$(s, p, q - p)
    p = q
End Function

Private Function MkDate(ByVal y As Long, ByVal mo As Long, ByVal d As Long) As String
    If mo >= 1 And mo <= 12 And d >= 1 And d <= 31 Then MkDate = y & "-" & Pad2(mo) & "-" & Pad2(d)
End Function

Private Function IsoOf(ByVal d As Date) As String
    IsoOf = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
End Function

' 2026-11-11 → 2026. 11. 11.(수)
Private Function FmtDate(ByVal iso As String) As String
    Dim d As Date
    If Len(iso) <> 10 Then Exit Function
    d = DateSerial(CLng(Left$(iso, 4)), CLng(Mid$(iso, 6, 2)), CLng(Mid$(iso, 9, 2)))
    FmtDate = Year(d) & ". " & Month(d) & ". " & Day(d) & ".(" & Mid$("일월화수목금토", Weekday(d, vbSunday), 1) & ")"
End Function

'---------------------------------------------------------------------
' 글자 꾸밈: **굵게** __밑줄__ !!빨간 강조!! ==형광펜==
'   plain = 표시 글자, runs = "시작,길이,꾸밈;" (꾸밈: 1 굵게 2 밑줄 4 빨강 8 형광펜)
'---------------------------------------------------------------------
Private Sub InlineFmt(ByVal s As String, ByRef plain As String, ByRef runs As String, ByVal offset As Long)
    Dim n As Long, i As Long, j As Long, k As Long, op As Long, cl As Long, cur As Long, st As Long, a As Long
    Dim at() As Long, mk As String, bit As Long
    runs = ""
    plain = s
    n = Len(s)
    If n = 0 Then Exit Sub
    ReDim at(1 To n)
    For k = 1 To 4
        Select Case k
            Case 1: mk = "**": bit = 1
            Case 2: mk = "__": bit = 2
            Case 3: mk = "!!": bit = 4
            Case 4: mk = "==": bit = 8
        End Select
        i = 1
        Do
            If i > Len(s) Then Exit Do
            op = InStr(i, s, mk)
            If op = 0 Then Exit Do
            cl = InStr(op + 3, s, mk)
            If cl = 0 Then Exit Do
            For j = op + 2 To cl - 1
                at(j) = at(j) Or bit
            Next
            DelAt s, at, cl, 2
            DelAt s, at, op, 2
            i = cl - 2
        Loop
    Next
    plain = s
    n = Len(s)
    cur = 0: st = 0
    For i = 1 To n + 1
        If i <= n Then a = at(i) Else a = 0
        If a <> cur Then
            If cur <> 0 Then runs = runs & (offset + st) & "," & (i - st) & "," & cur & ";"
            cur = a
            st = i
        End If
    Next
End Sub

Private Sub DelAt(ByRef s As String, ByRef at() As Long, ByVal p As Long, ByVal cnt As Long)
    Dim n As Long, j As Long
    n = Len(s)
    For j = p To n - cnt
        at(j) = at(j + cnt)
    Next
    s = Left$(s, p - 1) & Mid$(s, p + cnt)
    If n - cnt >= 1 Then
        ReDim Preserve at(1 To n - cnt)
    End If
End Sub

' 일정 내용: 끝의 "@장소" → "(장소)"
Private Sub MainFmt(ByVal s As String, ByRef plain As String, ByRef runs As String, ByVal offset As Long)
    Dim p As Long, place As String, txt As String, a As String, ra As String, b As String, rb As String
    p = InStrRev(s, "@")
    If p > 0 Then
        place = JTrim(Mid$(s, p + 1))
        If place <> "" Then
            If p = 1 Then
                InlineFmt place, b, rb, offset + 1
                plain = "(" & b & ")"
                runs = rb
                Exit Sub
            ElseIf IsWs(Mid$(s, p - 1, 1)) Then
                txt = RTrimWs(Left$(s, p - 1))
                InlineFmt txt, a, ra, offset
                If txt <> "" Then
                    InlineFmt place, b, rb, offset + Len(a) + 2
                    plain = a & " (" & b & ")"
                Else
                    InlineFmt place, b, rb, offset + 1
                    plain = "(" & b & ")"
                End If
                runs = ra & rb
                Exit Sub
            End If
        End If
    End If
    InlineFmt s, plain, runs, offset
End Sub

'---------------------------------------------------------------------
' 글자 도구
'---------------------------------------------------------------------
Private Function UC(ByVal code As Long) As String
    UC = ChrW(code)
End Function

Private Function CodeOf(ByVal c As String) As Long
    If Len(c) = 0 Then
        CodeOf = -1
    Else
        CodeOf = AscW(c)
        If CodeOf < 0 Then CodeOf = CodeOf + 65536
    End If
End Function

Private Function IsWs(ByVal c As String) As Boolean
    Select Case CodeOf(c)
        Case 9 To 13, 32, 160, 5760, 8192 To 8202, 8232, 8233, 8239, 8287, 12288, 65279
            IsWs = True
    End Select
End Function

Private Function IsDigitC(ByVal c As String) As Boolean
    Dim k As Long
    k = CodeOf(c)
    IsDigitC = (k >= 48 And k <= 57)
End Function

Private Function IsHangulC(ByVal c As String) As Boolean
    Dim k As Long
    k = CodeOf(c)
    IsHangulC = (k >= 44032 And k <= 55203)
End Function

Private Function IsStarC(ByVal c As String) As Boolean
    Dim k As Long
    k = CodeOf(c)
    IsStarC = (k = 42 Or k = 65290)
End Function

' 시간 범위 사이: ~ 와 비슷한 물결표들, - 와 대시들
Private Function IsSepC(ByVal c As String) As Boolean
    Select Case CodeOf(c)
        Case 126, 8764, 65374, 12316, 45, 8211, 8212
            IsSepC = True
    End Select
End Function

' 글머리 기호: - (대시들) ○ ● · ▶ ▷ □ ■ (그리고 작은 원·점)
Private Function IsBulletC(ByVal c As String) As Boolean
    Select Case CodeOf(c)
        Case 45, 8211, 9675, 9679, 9702, 183, 8226, 9654, 9655, 9633, 9632
            IsBulletC = True
    End Select
End Function

' 하위 항목 기호: - (대시) · (가운뎃점, 점)
Private Function IsDashC(ByVal c As String) As Boolean
    Select Case CodeOf(c)
        Case 45, 8211, 183, 8226
            IsDashC = True
    End Select
End Function

Private Function SkipWs(ByVal s As String, ByVal p As Long) As Long
    Do While p <= Len(s)
        If Not IsWs(Mid$(s, p, 1)) Then Exit Do
        p = p + 1
    Loop
    SkipWs = p
End Function

Private Function LTrimWs(ByVal s As String) As String
    LTrimWs = Mid$(s, SkipWs(s, 1))
End Function

Private Function RTrimWs(ByVal s As String) As String
    Dim j As Long
    j = Len(s)
    Do While j >= 1
        If Not IsWs(Mid$(s, j, 1)) Then Exit Do
        j = j - 1
    Loop
    RTrimWs = Left$(s, j)
End Function

Private Function JTrim(ByVal s As String) As String
    JTrim = RTrimWs(LTrimWs(s))
End Function

Private Function CollapseTabs(ByVal s As String) As String
    Do While InStr(1, s, vbTab & vbTab) > 0
        s = Replace(s, vbTab & vbTab, vbTab)
    Loop
    CollapseTabs = Replace(s, vbTab, " ")
End Function

Private Function Pad2(ByVal n As Long) As String
    If n < 10 And n >= 0 Then Pad2 = "0" & n Else Pad2 = CStr(n)
End Function

Private Function Clamp(ByVal x As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If x < lo Then x = lo
    If x > hi Then x = hi
    Clamp = x
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function Max3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    Max3 = MaxD(MaxD(a, b), c)
End Function

' 시간 글자의 대략 폭 (글자크기 1 기준, 굵게)
Private Function TextUnits(ByVal s As String) As Double
    Dim i As Long, u As Double
    For i = 1 To Len(s)
        Select Case CodeOf(Mid$(s, i, 1))
            Case 48 To 57: u = u + 0.6
            Case 58: u = u + 0.34
            Case 126: u = u + 0.64
            Case 32: u = u + 0.3
            Case Is >= 4352: u = u + 1
            Case Else: u = u + 0.62
        End Select
    Next
    TextUnits = u * 1.05
End Function

'=====================================================================
'  대시보드 그리기
'=====================================================================
Private Function DashSheet() As Object
    Dim ws As Object
    If SheetExists(SH_DASH) Then
        Set ws = ThisWorkbook.Worksheets(SH_DASH)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_DASH
    End If
    Set DashSheet = ws
End Function

Private Sub ClearDashShapes(ByVal ws As Object)
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(k).Name, Len(PFX)) = PFX Then ws.Shapes(k).Delete
    Next
End Sub

Private Sub ClearLinesOnly(ByVal ws As Object)
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(k).Name, Len(PFX) + 2) = PFX & "L_" Then ws.Shapes(k).Delete
    Next
End Sub

Private Sub SetColWidthPt(ByVal col As Object, ByVal pts As Double)
    Dim k As Long, w As Double, cw As Double
    If pts < 2 Then pts = 2
    cw = pts / 5.25
    If cw > 255 Then cw = 255
    col.ColumnWidth = cw
    For k = 1 To 5
        w = col.Width
        If w <= 0 Then Exit For
        If Abs(w - pts) < 0.7 Then Exit For
        cw = col.ColumnWidth * pts / w
        If cw > 255 Then cw = 255
        If cw < 0.1 Then cw = 0.1
        col.ColumnWidth = cw
    Next
End Sub

' 시트 칸: A 왼쪽 여백 | B 인쇄 영역(본문) | C 오른쪽 여백 | D 틈 | E 단추
Private Sub PrepareGrid(ByVal ws As Object)
    If sLandscape Then
        dPageW = 297 * PT_MM: dPageH = 210 * PT_MM
    Else
        dPageW = 210 * PT_MM: dPageH = 297 * PT_MM
    End If
    dMarginPt = sMargin * PT_MM
    ws.Cells.Clear
    ws.Cells.RowHeight = ws.StandardHeight
    ws.Cells.Interior.Color = RGB(214, 219, 226)
    SetColWidthPt ws.Columns(1), dMarginPt
    SetColWidthPt ws.Columns(2), dPageW - 2 * dMarginPt - 2
    SetColWidthPt ws.Columns(3), dMarginPt
    ws.Columns(4).ColumnWidth = 3
    ws.Columns(5).ColumnWidth = 34
    ws.Rows(1).RowHeight = dMarginPt
    dX0 = ws.Columns(2).Left + 1
    dW = ws.Columns(2).Width - 2
    dY0 = ws.Rows(2).Top + 1
    dH = dPageH - 2 * dMarginPt - 4
    dGap = GAP_MM * PT_MM
End Sub

Private Sub ComputeBoxX()
    Dim inner As Double
    inner = dW - 2 * dGap
    bW(1) = inner * sC1a: bW(2) = inner * sC1b: bW(3) = inner * sC1c
    bX(1) = dX0: bX(2) = bX(1) + bW(1) + dGap: bX(3) = bX(2) + bW(2) + dGap
    inner = dW - dGap
    bW(4) = inner * sC2a: bW(5) = inner * sC2b: bW(6) = bW(5)
    bX(4) = dX0: bX(5) = bX(4) + bW(4) + dGap: bX(6) = bX(5)
    bW(7) = inner * sC3a: bW(8) = inner * sC3b
    bX(7) = dX0: bX(8) = bX(7) + bW(7) + dGap
End Sub

Private Function NewText(ByVal ws As Object, ByVal nm As String, ByVal x As Double, ByVal w As Double) As Object
    Dim s As Object
    Set s = ws.Shapes.AddTextbox(M_HORIZ, x, dY0, w, 20)
    s.Name = PFX & nm
    On Error Resume Next
    s.Placement = X_FREE
    s.Fill.Visible = M_FALSE
    s.Line.Visible = M_FALSE
    With s.TextFrame2
        .WordWrap = M_TRUE
        .AutoSize = M_AUTO_NONE
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = M_ANCHOR_TOP
    End With
    On Error GoTo 0
    Set NewText = s
End Function

' 문단들을 넣고, 글꼴·색·꾸밈을 정합니다.
Private Sub SetParas(ByVal shp As Object, ByRef paras() As String, ByRef runs() As String, ByVal n As Long, ByVal bold As Boolean)
    Dim txt As String, j As Long, tr As Object, parts As Variant, k As Long, f As Variant, cnt As Long
    txt = ""
    For j = 1 To n
        If j > 1 Then txt = txt & vbCr
        If paras(j) = "" Then txt = txt & " " Else txt = txt & paras(j)
    Next
    Set tr = shp.TextFrame2.TextRange
    tr.Text = txt
    cnt = n
    On Error Resume Next
    cnt = tr.Paragraphs.Count
    On Error GoTo 0
    If n > 1 And cnt <> n Then
        tr.Text = Replace(txt, vbCr, vbLf)
        Set tr = shp.TextFrame2.TextRange
    End If
    On Error Resume Next
    With tr.Font
        .Name = sFont
        .NameFarEast = sFont
        .NameComplex = sFont
        .Bold = IIf(bold, M_TRUE, M_FALSE)
        .Italic = M_FALSE
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Visible = M_TRUE
    End With
    tr.ParagraphFormat.Alignment = M_ALIGN_LEFT
    tr.ParagraphFormat.WordWrap = M_TRUE
    For j = 1 To n
        If Len(runs(j)) > 0 Then
            parts = Split(runs(j), ";")
            For k = LBound(parts) To UBound(parts)
                If Len(parts(k)) > 0 Then
                    f = Split(parts(k), ",")
                    StyleRun tr.Paragraphs(j).Characters(CLng(f(0)), CLng(f(1))), CLng(f(2))
                End If
            Next
        End If
    Next
    On Error GoTo 0
End Sub

Private Sub StyleRun(ByVal rg As Object, ByVal a As Long)
    On Error Resume Next
    With rg.Font
        If (a And 1) <> 0 Or (a And 4) <> 0 Then .Bold = M_TRUE
        If (a And 2) <> 0 Then .UnderlineStyle = M_UNDERLINE
        If (a And 4) <> 0 Then .Fill.ForeColor.RGB = RGB(196, 0, 0)
        If (a And 8) <> 0 Then .Highlight.RGB = RGB(255, 230, 128)
    End With
End Sub

Private Sub SetOneText(ByVal shp As Object, ByVal txt As String, ByVal bold As Boolean)
    Dim p(1 To 1) As String, r(1 To 1) As String
    p(1) = txt
    r(1) = ""
    SetParas shp, p, r, 1, bold
End Sub

Private Sub CreateAllShapes(ByVal ws As Object)
    Dim i As Long, j As Long, n As Long, paras() As String, runs() As String, maxU As Double, maxRg As Double, u As Double
    Dim t As String, p As Long, body As String, rb As String, sep As String, hdr As String

    ' 제목·부서·날짜·하단 문구
    Set shTitle = Nothing: Set shFoot = Nothing
    If sTitle <> "" Then
        Set shTitle = NewText(ws, "TITLE", dX0, dW)
        SetOneText shTitle, sTitle, True
        shTitle.TextFrame2.TextRange.ParagraphFormat.Alignment = M_ALIGN_CENTER
        On Error Resume Next
        shTitle.TextFrame2.TextRange.Font.Spacing = 1
        On Error GoTo 0
    End If
    Set shSub = NewText(ws, "SUB", dX0, dW * 0.62)
    SetOneText shSub, IIf(sSub = "", " ", sSub), False
    Set shDate = NewText(ws, "DATE", dX0 + dW * 0.38, dW * 0.62)
    SetOneText shDate, FmtDate(sDateISO), True
    shDate.TextFrame2.TextRange.ParagraphFormat.Alignment = M_ALIGN_RIGHT
    If sFoot <> "" Then
        Set shFoot = NewText(ws, "FOOT", dX0, dW)
        SetOneText shFoot, sFoot, False
    End If

    For i = 1 To N_EX
        ' 머리 (직위 성명)
        Set shH(i) = NewText(ws, "H" & i, bX(i), bW(i))
        hdr = JTrim(gPos(i) & " " & gName(i))
        If hdr = "" Then hdr = " "
        SetOneText shH(i), hdr, True
        With shH(i)
            .Fill.Visible = M_TRUE
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(231, 233, 236)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = M_ALIGN_CENTER
            .TextFrame2.VerticalAnchor = M_ANCHOR_MIDDLE
        End With

        ' 본문 (일정)
        Set shB(i) = NewText(ws, "B" & i, bX(i), bW(i))
        n = gN(i)
        bHasTime(i) = False
        maxU = 0: maxRg = 0
        For j = 1 To n
            If gKind(i, j) = "M" And gTime(i, j) <> "" Then
                bHasTime(i) = True
                t = gTime(i, j)
                p = InStr(1, t, "~")
                If p > 0 Then
                    u = TextUnits(Left$(t, p))
                    If u > maxRg Then maxRg = u
                Else
                    u = TextUnits(t)
                    If u > maxU Then maxU = u
                End If
            End If
        Next
        ' 시간 칸 폭: 시각 하나짜리 기준 (범위만 있으면 "14:00~" 기준)
        If maxU = 0 Then maxU = maxRg
        bTimeU(i) = maxU + 0.6
        If n > 0 Then
            ReDim paras(1 To n): ReDim runs(1 To n)
            For j = 1 To n
                Select Case gKind(i, j)
                    Case "M"
                        sep = ""
                        If bHasTime(i) Then
                            t = gTime(i, j)
                            ' 칸보다 긴 시간(범위)은 탭 대신 빈칸 → 내용은 시간 바로 뒤에서 시작
                            If t <> "" And TextUnits(t) + 0.25 > bTimeU(i) Then sep = " " Else sep = vbTab
                            MainFmt gText(i, j), body, rb, Len(t) + 1
                            paras(j) = t & sep & body
                            runs(j) = rb
                            If t <> "" Then runs(j) = "1," & Len(t) & ",1;" & rb
                        Else
                            MainFmt gText(i, j), body, rb, 0
                            paras(j) = body
                            runs(j) = rb
                        End If
                    Case "C"
                        MainFmt gText(i, j), body, rb, 0
                        paras(j) = body: runs(j) = rb
                    Case "S", "D"
                        InlineFmt gText(i, j), body, rb, 2
                        paras(j) = IIf(gKind(i, j) = "S", "*", "-") & vbTab & body
                        runs(j) = rb
                    Case "S2", "D2"
                        InlineFmt gText(i, j), body, rb, 0
                        paras(j) = body: runs(j) = rb
                    Case Else
                        paras(j) = gText(i, j): runs(j) = ""
                End Select
            Next
            SetParas shB(i), paras, runs, n, False
            If gKind(i, 1) = "E" Then
                With shB(i).TextFrame2.TextRange
                    .ParagraphFormat.Alignment = M_ALIGN_CENTER
                    .Font.Fill.ForeColor.RGB = RGB(138, 138, 138)
                End With
            End If
        End If

        ' ※ 참고 (칸 맨 밑)
        Set shR(i) = Nothing
        If gNR(i) > 0 Then
            Set shR(i) = NewText(ws, "R" & i, bX(i), bW(i))
            ReDim paras(1 To gNR(i)): ReDim runs(1 To gNR(i))
            For j = 1 To gNR(i)
                InlineFmt gRef(i, j), body, rb, 2
                paras(j) = UC(8251) & vbTab & body
                runs(j) = rb
            Next
            SetParas shR(i), paras, runs, gNR(i), False
        End If

        ' 테두리 (맨 위에 둠)
        Set shF(i) = ws.Shapes.AddShape(M_RECT, bX(i), dY0, bW(i), 40)
        With shF(i)
            .Name = PFX & "F" & i
            .Placement = X_FREE
            .Fill.Visible = M_FALSE
            .Line.Visible = M_TRUE
            .Line.ForeColor.RGB = RGB(34, 34, 34)
            .Line.Weight = 0.75
        End With
        On Error Resume Next
        shF(i).Shadow.Visible = M_FALSE
        shF(i).TextFrame2.TextRange.Text = ""
        On Error GoTo 0
    Next
End Sub

' 문단 서식 한 번에 (들여쓰기·탭·간격·크기)
Private Sub FmtPara(ByVal tr As Object, ByVal j As Long, ByVal sz As Double, ByVal leftInd As Double, ByVal firstInd As Double, _
                    ByVal tabPos As Double, ByVal spBefore As Double, ByVal within As Double)
    Dim g As Long, pr As Object
    On Error Resume Next
    Set pr = tr.Paragraphs(j)
    If pr Is Nothing Then Exit Sub
    pr.Font.Size = sz
    With pr.ParagraphFormat
        .FirstLineIndent = 0
        .LeftIndent = leftInd
        .FirstLineIndent = firstInd
        .LineRuleBefore = M_FALSE
        .SpaceBefore = spBefore
        .LineRuleAfter = M_FALSE
        .SpaceAfter = 0
        .LineRuleWithin = M_TRUE
        .SpaceWithin = within
        With .TabStops
            g = 0
            Do While .Count > 0 And g < 20
                .Item(1).Clear
                g = g + 1
            Loop
            If tabPos > 0 Then .Add M_TAB_LEFT, tabPos
        End With
    End With
End Sub

Private Sub SetMargins(ByVal shp As Object, ByVal l As Double, ByVal r As Double, ByVal t As Double, ByVal b As Double)
    On Error Resume Next
    With shp.TextFrame2
        .MarginLeft = l: .MarginRight = r: .MarginTop = t: .MarginBottom = b
    End With
End Sub

Private Sub SizeAll(ByVal shp As Object, ByVal sz As Double, ByVal within As Double)
    On Error Resume Next
    With shp.TextFrame2.TextRange
        .Font.Size = sz
        .ParagraphFormat.LineRuleWithin = M_TRUE
        .ParagraphFormat.SpaceWithin = within
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
    End With
End Sub

' 글자크기 fs에 맞춰 모든 칸의 크기·간격을 정합니다 (HTML 판의 varsFor 와 같은 비례)
Private Sub ApplySizes(ByVal fs As Double)
    Dim uu As Double, lineMul As Double, rp As Double, padv As Double, padh As Double
    Dim starFs As Double, refFs As Double, headFs As Double, metaFs As Double
    Dim i As Long, j As Long, tw As Double, mk As Double, mkS As Double, mkR As Double, inner As Double, tr As Object
    uu = 1
    If sMaxF > sMinF Then uu = Clamp((fs - sMinF) / (sMaxF - sMinF), 0, 1)
    lineMul = 0.94 + 0.16 * uu
    rp = (0.16 + 0.4 * uu) * fs
    padv = (0.22 + 0.3 * uu) * fs
    padh = (0.4 + 0.2 * uu) * fs
    dPadH = padh
    starFs = fs - 2
    If starFs < 0.75 * fs Then starFs = 0.75 * fs
    If starFs < 6 Then starFs = 6
    If starFs > fs Then starFs = fs
    refFs = 0.93 * fs
    headFs = 1.06 * fs
    metaFs = Clamp(fs * 0.95, 8.5, 11)
    mk = 0.85 * fs
    mkS = 0.9 * starFs
    mkR = 1.35 * refFs

    If Not shTitle Is Nothing Then SizeAll shTitle, Clamp(fs * 1.75, 15, 21), 1
    SizeAll shSub, metaFs, 1
    SizeAll shDate, metaFs, 1
    If Not shFoot Is Nothing Then SizeAll shFoot, metaFs, 1

    For i = 1 To N_EX
        SizeAll shH(i), headFs, 1
        SetMargins shH(i), padh, padh, padv * 0.7, padv * 0.7
        SetMargins shB(i), padh, padh, padv * 0.5 + 0.5, padv * 0.5 + 0.5
        tw = 0
        If bHasTime(i) Then
            tw = bTimeU(i) * fs
            inner = bW(i) - 2 * padh
            If tw > inner * 0.45 Then tw = inner * 0.45
        End If
        Set tr = shB(i).TextFrame2.TextRange
        For j = 1 To gN(i)
            Select Case gKind(i, j)
                Case "M": FmtPara tr, j, fs, tw, -tw, tw, IIf(j = 1, 0, rp), lineMul
                Case "C": FmtPara tr, j, fs, tw, 0, 0, 0, lineMul
                Case "S": FmtPara tr, j, starFs, tw + mkS, -mkS, tw + mkS, IIf(j = 1, 0, 0.08 * fs), lineMul * 0.96
                Case "S2": FmtPara tr, j, starFs, tw + mkS, 0, 0, 0, lineMul * 0.96
                Case "D": FmtPara tr, j, fs, tw + mk, -mk, tw + mk, IIf(j = 1, 0, 0.04 * fs), lineMul
                Case "D2": FmtPara tr, j, fs, tw + mk, 0, 0, 0, lineMul
                Case Else: FmtPara tr, j, fs, 0, 0, 0, padv, lineMul
            End Select
        Next
        If gNR(i) > 0 Then
            SetMargins shR(i), padh, padh, padv * 0.8, padv * 0.9
            Set tr = shR(i).TextFrame2.TextRange
            For j = 1 To gNR(i)
                FmtPara tr, j, refFs, mkR, -mkR, mkR, IIf(j = 1, 0, 0.08 * refFs), 0.96
            Next
        End If
    Next
End Sub

' 글 상자에 글이 다 들어가는 높이 (두 가지로 재서 큰 쪽)
Private Function TextNeedH(ByVal shp As Object) As Double
    Dim h1 As Double, h2 As Double
    On Error Resume Next
    With shp.TextFrame2
        h1 = .TextRange.BoundHeight + .MarginTop + .MarginBottom
        shp.Height = 2
        .AutoSize = M_AUTO_FIT
        h2 = shp.Height
        .AutoSize = M_AUTO_NONE
    End With
    On Error GoTo 0
    If h2 > h1 Then TextNeedH = h2 Else TextNeedH = h1
End Function

Private Sub MeasureAll()
    Dim i As Long
    For i = 1 To N_EX
        hHead(i) = TextNeedH(shH(i))
        hBody(i) = TextNeedH(shB(i))
        If gNR(i) > 0 Then hRefBox(i) = TextNeedH(shR(i)) Else hRefBox(i) = 0
        hNat(i) = hHead(i) + hBody(i) + hRefBox(i) + 1
    Next
    If shTitle Is Nothing Then hTitle = 0 Else hTitle = TextNeedH(shTitle)
    hMeta = MaxD(TextNeedH(shDate), TextNeedH(shSub))
    If shFoot Is Nothing Then hFoot = 0 Else hFoot = TextNeedH(shFoot)
End Sub

Private Function HeadBlockH() As Double
    Dim h As Double
    If hTitle > 0 Then h = hTitle + 1.2 * PT_MM
    HeadBlockH = h + hMeta + 1.6 * PT_MM + 1.5 + 3 * PT_MM
End Function

Private Function FootBlockH() As Double
    If hFoot > 0 Then FootBlockH = 2.2 * PT_MM + hFoot
End Function

' 줄마다 꼭 필요한 높이를 재고 A4 한 장에 들어가는지 봅니다.
Private Function CalcFit() As Boolean
    Dim g As Double, t As Double, ok As Boolean
    g = dGap
    fAvail = dH - HeadBlockH() - FootBlockH()
    fInner = fAvail - 2 * g
    fReq0 = Max3(hNat(1), hNat(2), hNat(3))
    If sSplit < 0 Then
        fReq1 = MaxD(hNat(4), hNat(5) + g + hNat(6))
    Else
        fReq1 = Max3(hNat(4), hNat(5) / sSplit + g, hNat(6) / (1 - sSplit) + g)
    End If
    fReq2 = MaxD(hNat(7), hNat(8))
    If Not sRowsSet Then
        fNeed = fReq0 + fReq1 + fReq2 + 2 * g
        ok = (fNeed <= fAvail - 1.5)
    Else
        ok = True
        fNeed = 2 * g
        t = sR1 * (fInner - 1)
        If fReq0 > t Then ok = False
        fNeed = fNeed + MaxD(fReq0, t)
        t = sR2 * (fInner - 1)
        If fReq1 > t Then ok = False
        fNeed = fNeed + MaxD(fReq1, t)
        t = sR3 * (fInner - 1)
        If fReq2 > t Then ok = False
        fNeed = fNeed + MaxD(fReq2, t)
    End If
    fOK = ok
    CalcFit = ok
End Function

Private Function TryFit(ByVal fs As Double) As Boolean
    Application.StatusBar = "임원 대시보드 글자크기 맞추는 중... " & Format$(fs, "0.0") & "pt"
    ApplySizes fs
    MeasureAll
    fFs = fs
    TryFit = CalcFit()
End Function

' 8칸 모두 같은 글자크기로, A4 한 장을 꽉 채우는 가장 큰 크기 (최소~최대 안에서)
Private Sub FitFont()
    Dim lo As Double, hi As Double, a As Double, b As Double, md As Double, f As Double, k As Long
    lo = sMinF: hi = sMaxF
    If TryFit(hi) Then Exit Sub
    If Not TryFit(lo) Then Exit Sub
    a = lo: b = hi
    For k = 1 To 8
        If b - a <= 0.1 Then Exit For
        md = (a + b) / 2
        If TryFit(md) Then a = md Else b = md
    Next
    f = Int(a * 10 + 0.0001) / 10
    If f < lo Then f = lo
    For k = 1 To 20
        If TryFit(f) Then Exit For
        If f <= lo Then Exit For
        f = f - 0.1
        If f < lo Then f = lo
    Next
End Sub

Private Sub PlaceShape(ByVal shp As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    On Error Resume Next
    shp.TextFrame2.AutoSize = M_AUTO_NONE
    On Error GoTo 0
    shp.Left = x
    shp.Top = y
    shp.Width = w
    If h < 1 Then h = 1
    shp.Height = h
End Sub

Private Sub AddHLine(ByVal ws As Object, ByVal nm As String, ByVal x1 As Double, ByVal x2 As Double, ByVal y As Double, _
                     ByVal wt As Double, ByVal clr As Long, ByVal dashed As Boolean)
    Dim s As Object
    Set s = ws.Shapes.AddLine(x1, y, x2, y)
    s.Name = PFX & "L_" & nm
    s.Placement = X_FREE
    With s.Line
        .Visible = M_TRUE
        .ForeColor.RGB = clr
        .Weight = wt
        If dashed Then .DashStyle = M_DASH
    End With
End Sub

Private Sub PlaceBox(ByVal ws As Object, ByVal i As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim rh As Double
    PlaceShape shH(i), x, y, w, hHead(i)
    AddHLine ws, "H" & i, x, x + w, y + hHead(i), 1.1, RGB(34, 34, 34), False
    rh = hRefBox(i)
    PlaceShape shB(i), x, y + hHead(i), w, h - hHead(i) - rh
    If gNR(i) > 0 Then
        PlaceShape shR(i), x, y + h - rh, w, rh
        AddHLine ws, "R" & i, x + dPadH, x + w - dPadH, y + h - rh, 0.75, RGB(102, 102, 102), True
    End If
    PlaceShape shF(i), x, y, w, h
    shF(i).ZOrder 0
End Sub

' 맞춘 글자크기로 줄·칸 높이를 정해 페이지 끝까지 채웁니다.
Private Sub PlaceAll(ByVal ws As Object)
    Dim y As Double, y1 As Double, y2 As Double, g As Double, h0 As Double, h1 As Double, h2 As Double
    Dim in2 As Double, ex As Double, extra As Double, hE As Double, hF As Double, fy As Double
    g = dGap
    y = dY0
    If Not shTitle Is Nothing Then
        PlaceShape shTitle, dX0, y, dW, hTitle
        y = y + hTitle + 1.2 * PT_MM
    End If
    PlaceShape shSub, dX0, y, dW * 0.62, hMeta
    PlaceShape shDate, dX0 + dW * 0.38, y, dW * 0.62, hMeta
    y = y + hMeta + 1.6 * PT_MM
    AddHLine ws, "TOP", dX0, dX0 + dW, y + 0.75, 1.5, RGB(0, 0, 0), False
    y = y + 1.5 + 3 * PT_MM

    If fOK Then
        If Not sRowsSet Then
            extra = fInner - 1 - (fReq0 + fReq1 + fReq2)
            If extra < 0 Then extra = 0
            h0 = fReq0 + extra / 4: h1 = fReq1 + extra / 2: h2 = fReq2 + extra / 4
        Else
            h0 = sR1 * (fInner - 1): h1 = sR2 * (fInner - 1): h2 = sR3 * (fInner - 1)
        End If
    Else
        h0 = fReq0: h1 = fReq1: h2 = fReq2
        If sRowsSet Then
            h0 = MaxD(h0, sR1 * (fInner - 1)): h1 = MaxD(h1, sR2 * (fInner - 1)): h2 = MaxD(h2, sR3 * (fInner - 1))
        End If
    End If
    in2 = h1 - g
    If sSplit < 0 Then
        ex = in2 - hNat(5) - hNat(6)
        If ex < 0 Then ex = 0
        hE = hNat(5) + ex / 2: hF = hNat(6) + ex / 2
    Else
        hE = MaxD(sSplit * in2, hNat(5)): hF = MaxD(in2 - sSplit * in2, hNat(6))
    End If
    h1 = Max3(h1, hE + g + hF, hNat(4))

    PlaceBox ws, 1, bX(1), y, bW(1), h0
    PlaceBox ws, 2, bX(2), y, bW(2), h0
    PlaceBox ws, 3, bX(3), y, bW(3), h0
    y1 = y + h0 + g
    PlaceBox ws, 4, bX(4), y1, bW(4), h1
    PlaceBox ws, 5, bX(5), y1, bW(5), hE
    PlaceBox ws, 6, bX(6), y1 + hE + g, bW(6), hF
    y2 = y1 + h1 + g
    PlaceBox ws, 7, bX(7), y2, bW(7), h2
    PlaceBox ws, 8, bX(8), y2, bW(8), h2
    fBottom = y2 + h2
    If Not shFoot Is Nothing Then
        fy = MaxD(fBottom + 2.2 * PT_MM, dY0 + dH - hFoot)
        PlaceShape shFoot, dX0, fy, dW, hFoot
        fBottom = fy + hFoot
    End If
End Sub

' 화면을 끈 채로 잰 높이가 맞았는지: 본문 글이 칸 안에 다 들어가는지 다시 잼
Private Function VerifyBodies() As Boolean
    Dim i As Long, need As Double, hh As Double
    VerifyBodies = True
    On Error Resume Next
    For i = 1 To N_EX
        If gN(i) > 0 Then
            need = shB(i).TextFrame2.TextRange.BoundHeight + shB(i).TextFrame2.MarginTop + shB(i).TextFrame2.MarginBottom
            hh = shB(i).Height
            If need > hh + 1.5 Then VerifyBodies = False
        End If
    Next
End Function

' 인쇄 영역(B열)·행 높이·용지 설정. 넘치면 인쇄 영역을 늘려 한 장에 줄여 찍히게 합니다.
Private Function FinishPage(ByVal ws As Object) As Long
    Dim totalH As Double, nRows As Long, rh As Double, r As Long, lastRow As Long, pageRows As Long
    totalH = dPageH - 2 * dMarginPt - 2
    If fBottom - (dY0 - 1) + 2 > totalH Then totalH = fBottom - (dY0 - 1) + 2
    nRows = Int(totalH / 400) + 1
    rh = totalH / nRows
    For r = 2 To nRows + 1
        ws.Rows(r).RowHeight = rh
    Next
    lastRow = nRows + 2
    ws.Rows(lastRow).RowHeight = dMarginPt
    pageRows = nRows + 1
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 3))
        .Interior.Color = RGB(255, 255, 255)
        .BorderAround 1, 2, , RGB(160, 167, 177)
    End With

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$B$2:$B$" & pageRows
        .PaperSize = X_A4
        If sLandscape Then .Orientation = X_LANDSCAPE Else .Orientation = X_PORTRAIT
        .LeftMargin = dMarginPt
        .RightMargin = dMarginPt
        .TopMargin = dMarginPt
        .BottomMargin = dMarginPt
        .HeaderMargin = 0
        .FooterMargin = 0
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintGridlines = False
        .PrintHeadings = False
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
    FinishPage = lastRow
End Function

Private Sub AddDashButtons(ByVal ws As Object)
    Dim x As Double, y As Double
    x = ws.Columns(5).Left + 6
    y = dY0
    AddButton ws, PFX & "BTN1", "새로 고침 (Ctrl+Shift+R)", "RefreshDashboard", x, y, 170
    AddButton ws, PFX & "BTN2", "인쇄 미리보기", "PreviewDashboard", x, y + 30, 170
    AddButton ws, PFX & "BTN3", "PDF로 저장", "SaveDashboardPDF", x, y + 60, 170
    AddButton ws, PFX & "BTN4", "엑셀 불러오기 (두레이)", "ImportExcel", x, y + 90, 170
    AddButton ws, PFX & "BTN5", "설정 · 직위 · 성명", "GoSettings", x, y + 120, 170
End Sub

Private Function StatusText() As String
    Dim over As Double
    If fOK Then
        StatusText = "A4 한 장에 맞췄습니다." & vbLf & "글자 " & Format$(fFs, "0.0") & "pt (범위 " & sMinF & "~" & sMaxF & "pt)"
        If fFs >= sMaxF - 0.001 And Not sRowsSet Then
            StatusText = StatusText & vbLf & "최대 글자크기 · 남는 높이는 칸을 늘려 채움"
        End If
    Else
        over = (fNeed - fAvail) / PT_MM
        StatusText = "넘침: 최소 글자크기(" & sMinF & "pt)로도 약 " & Format$(over, "0") & "mm 넘칩니다." & vbLf & _
            "인쇄하면 한 장에 들어가도록 조금 줄여 찍힙니다." & vbLf & _
            "내용을 줄이거나 설정에서 최소 글자크기·여백을 줄여 보세요."
        If sRowsSet Or sSplit >= 0 Then StatusText = StatusText & vbLf & "(줄 높이를 직접 정했다면 비워서 자동으로 해 보세요)"
    End If
End Function

Private Sub WriteStatus(ByVal ws As Object)
    Dim s As Object, txt As String
    txt = StatusText()
    On Error Resume Next
    Set s = ws.Shapes.AddTextbox(M_HORIZ, ws.Columns(5).Left + 6, dY0 + 160, 190, 80)
    s.Name = PFX & "STATUS"
    s.Placement = X_FREE
    s.Line.Visible = M_FALSE
    s.Fill.Visible = M_FALSE
    With s.TextFrame2
        .WordWrap = M_TRUE
        .AutoSize = M_AUTO_FIT
        .TextRange.Text = txt
        .TextRange.Font.Size = 10
        .TextRange.Font.Name = "맑은 고딕"
        .TextRange.Font.NameFarEast = "맑은 고딕"
        If fOK Then
            .TextRange.Font.Fill.ForeColor.RGB = RGB(30, 125, 67)
        Else
            .TextRange.Font.Fill.ForeColor.RGB = RGB(194, 54, 43)
        End If
        .TextRange.Paragraphs(1).Font.Bold = M_TRUE
    End With
    Cfg("ED_Status").Value = Replace(txt, vbLf, "  ")
    If fOK Then
        Cfg("ED_Status").Font.Color = RGB(30, 125, 67)
    Else
        Cfg("ED_Status").Font.Color = RGB(194, 54, 43)
    End If
End Sub

Private Sub ZoomToPage(ByVal ws As Object, ByVal lastRow As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 5)).Select
    ActiveWindow.Zoom = True
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Cells(1, 4).Select
End Sub
