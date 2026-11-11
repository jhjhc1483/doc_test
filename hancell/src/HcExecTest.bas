Attribute VB_Name = "HcExecTest"
Option Explicit
' LibreOffice-only helpers for tools/build.py (not part of the Hancell paste files)

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
    r = r & "afterNew items1=" & LCount(exItems(1)) & "; "
    gStep = "undo"
    UndoLast
    LoadExecs
    r = r & "afterUndo items1=" & LCount(exItems(1)) & "; "
    gStep = "paste"
    Set ws = HostBook().Worksheets(SH_PASTE)
    ws.Range("A3:C3").Value = Array("성명", "시간", "내용")
    ws.Range("A4:C4").Value = Array("한소영", "10:00", "붙여넣기 시험")
    ImportPasted
    LoadExecs
    r = r & "pasted8=" & LCount(exItems(8)) & " pasteLeft=" & ws.Range("A4").Value & "; "
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
