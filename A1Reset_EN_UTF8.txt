Attribute VB_Name = "A1Reset"
Option Explicit

' For Windows desktop Excel. Import this file as a standard VBA module.
' Control!B6 holds the target folder. This macro workbook is excluded.
Private Const CONTROL_SHEET As String = "Control"
Private Const FOLDER_CELL As String = "B6"
Private Const PREVIEW_LIMIT As Long = 30
Private Const PREVIEW_CHARS As Long = 650
Private Const SECURITY_FORCE_DISABLE As Long = 3
Private mRunning As Boolean
Private mCancelRequested As Boolean

Private Type ApplicationState
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    AutomationSecurity As Long
    EnableCancelKey As Long
    StatusBar As Variant
End Type

Public Sub SetupA1ResetButton()
    ' Run once during setup. Add one working Run button over the existing control sheet.
    Dim ws As Worksheet, area As Range, button As Shape
    On Error GoTo Failed
    If mRunning Then Exit Sub
    If ThisWorkbook.FileFormat <> xlOpenXMLWorkbookMacroEnabled And _
        ThisWorkbook.FileFormat <> xlExcel12 Then
        MsgBox "Save this control workbook as an .xlsm file first.", vbExclamation
        Exit Sub
    End If
    Set ws = GetControlSheet()
    If StrComp(ws.Name, CONTROL_SHEET, vbBinaryCompare) <> 0 Then ws.Name = CONTROL_SHEET
    ws.Range("B2").Value2 = "Batch Excel File Processing"
    ws.Range("B3").Value2 = "Choose the folder to process"
    ws.Range("B5").Value2 = "Target folder path"
    ws.Range("H5").Value2 = "Input: B6"
    ws.Range("B8").Value2 = "Paste the target folder path into the white input box."
    ws.Range("D10").Value2 = "Run"
    ws.Range("B14").Value2 = "After setup"
    ws.Range("B15").Value2 = "1  Review the files found in the target folder."
    ws.Range("B16").Value2 = "2  Confirm to process every eligible Excel file."
    ws.Range("B17").Value2 = "3  Save to the original files and close them."
    ws.Range("B19").Value2 = "Back up the target files in another folder before running."
    Set area = ws.Range("D10:H11")
    On Error Resume Next
    Set button = ws.Shapes("btnRunA1Reset")
    On Error GoTo Failed
    If button Is Nothing Then
        Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
        button.Name = "btnRunA1Reset"
    End If
    With button
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!RunA1Reset"
        .TextFrame.Characters.Text = "Run"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Size = 16
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.ForeColor.RGB = RGB(15, 118, 110)
        .Line.Visible = 0
        .Placement = xlMoveAndSize
    End With
    ws.Range("B12").Value2 = "Review the file list and overwrite confirmation before running."
    ws.Range("B22").Value2 = "Files: .xlsx / .xlsm / .xls / .xlsb; no subfolders."
    MsgBox "The Run button is ready. Save this control workbook.", vbInformation
    Exit Sub
Failed:
    MsgBox "Could not set up the Run button." & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub SelectTargetFolder()
    Dim dlg As Object, ws As Worksheet
    On Error GoTo Failed
    If mRunning Then Exit Sub
    Set ws = GetControlSheet()
    Set dlg = Application.FileDialog(4) ' msoFileDialogFolderPicker
    With dlg
        .Title = "Select the target folder (subfolders are excluded)"
        .AllowMultiSelect = False
        If .Show = -1 Then
            ws.Range(FOLDER_CELL).Value2 = .SelectedItems(1)
        End If
    End With
    Exit Sub
Failed:
    MsgBox "Could not set the target folder." & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub CancelA1Reset()
    ' Request a stop at a file boundary. Files already saved are not rolled back.
    If mRunning Then mCancelRequested = True
End Sub

Public Sub RunA1Reset()
    Dim fso As Object, cfg As Worksheet, report As Worksheet
    Dim paths() As String, sizes() As Double, stamps() As Date
    Dim folder As String, n As Long, i As Long, lastProcessed As Long
    Dim state As ApplicationState, captured As Boolean
    Dim status As String, detail As String, stopBatch As Boolean
    Dim savedCount As Long, skippedCount As Long, failedCount As Long
    Dim ending As String, restoreIssues As String, startAt As Date

    If mRunning Then
        MsgBox "A run is already in progress.", vbInformation
        Exit Sub
    End If
    On Error GoTo Failed
    mRunning = True
    mCancelRequested = False
    Set cfg = GetControlSheet()
    If IsError(cfg.Range(FOLDER_CELL).Value) Then
        Err.Raise vbObjectError + 100, , "The folder input in B6 contains an Excel error value."
    End If
    folder = Trim$(CStr(cfg.Range(FOLDER_CELL).Value2))
    If Len(folder) = 0 Then Err.Raise vbObjectError + 101, , "Enter the target folder path in B6."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Err.Raise vbObjectError + 102, , "The folder does not exist."
    folder = fso.GetFolder(folder).Path
    If Len(folder) > 300 Then Err.Raise vbObjectError + 106, , "The folder path is too long. Use a shorter local path."
    n = CollectTargets(fso, folder, paths)
    If n = 0 Then
        MsgBox "No eligible Excel files were found." & vbCrLf & _
            "Files: .xlsx / .xlsm / .xls / .xlsb; no subfolders.", vbInformation
        GoTo Finish
    End If
    If n > cfg.Rows.Count - 4 Then Err.Raise vbObjectError + 103, , "The file count exceeds the results sheet row limit."
    SortPaths paths, 1, n
    ReDim sizes(1 To n)
    ReDim stamps(1 To n)
    For i = 1 To n
        sizes(i) = CDbl(fso.GetFile(paths(i)).Size)
        stamps(i) = fso.GetFile(paths(i)).DateLastModified
    Next i

    ' Before confirmation, do not open target workbooks or modify sheets or files.
    If Not ConfirmTargets(fso, folder, paths, n) Then GoTo Finish
    If AutoSaveIsOn(ThisWorkbook) Then
        Err.Raise vbObjectError + 105, , "Turn off AutoSave for this control workbook and use a local copy."
    End If
    If ThisWorkbook.ProtectStructure Then
        Err.Raise vbObjectError + 104, , "The control workbook structure is protected; a results sheet cannot be created."
    End If

    CaptureApplication state
    captured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = True ' Do not suppress compatibility, save, or security warnings.
    Application.EnableCancelKey = xlErrorHandler
    Application.AutomationSecurity = SECURITY_FORCE_DISABLE
    ' Leave Calculation and CalculateBeforeSave unchanged to preserve calculation settings.
    Set report = NewReportSheet()
    startAt = Now
    report.Range("A1").Value2 = "Batch A1 view reset: " & Format$(startAt, "yyyy/mm/dd hh:nn:ss")
    report.Range("A2").Value2 = "Folder: " & folder
    For i = 1 To n
        report.Cells(i + 4, 1).Value2 = CStr(i)
        report.Cells(i + 4, 2).Value2 = fso.GetFileName(paths(i))
        report.Cells(i + 4, 3).Value2 = "Not processed"
        report.Cells(i + 4, 5).Value2 = paths(i)
    Next i
    ending = "Completed"
    For i = 1 To n
        DoEvents
        If mCancelRequested Then
            ending = "Stopped. Files not yet processed have not been changed."
            Exit For
        End If
        Application.StatusBar = "A1 reset " & i & " / " & n & " : " & fso.GetFileName(paths(i))
        report.Cells(i + 4, 3).Value2 = "Processing"
        detail = vbNullString
        stopBatch = False
        status = ProcessOneFile(fso, paths(i), sizes(i), stamps(i), detail, stopBatch)
        lastProcessed = i
        Select Case status
            Case "Saved": savedCount = savedCount + 1
            Case "Saved; close failed"
                savedCount = savedCount + 1
                failedCount = failedCount + 1
            Case "Skipped": skippedCount = skippedCount + 1
            Case Else: failedCount = failedCount + 1
        End Select
        report.Cells(i + 4, 3).Value2 = status
        report.Cells(i + 4, 4).Value2 = detail
        If stopBatch Then
            ending = "Stopped. Review the details in the results sheet."
            Exit For
        End If
    Next i
    GoTo Finish

Failed:
    If Err.Number = 18 Then
        ending = "Stopped by the Esc key."
    Else
        ending = "Stopped: " & Err.Number & " / " & Err.Description
    End If
Finish:
    If captured Then restoreIssues = RestoreApplication(state)
    mRunning = False
    mCancelRequested = False
    If Not report Is Nothing Then
        On Error Resume Next
        report.Range("A3").Value2 = ending & " / Saved " & savedCount & " / Skipped " & skippedCount & _
            " / Errors/warnings " & failedCount & " / Not processed " & (n - lastProcessed)
        report.Activate
        report.Range("A1").Select
        On Error GoTo 0
    End If
    If Len(ending) > 0 Then
        If report Is Nothing Then
            MsgBox ending, vbExclamation, "A1 View Reset"
        Else
            MsgBox ending & vbCrLf & "Saved: " & savedCount & " file(s)" & vbCrLf & _
                "Skipped: " & skippedCount & " file(s) / Errors/warnings: " & failedCount & " file(s)" & vbCrLf & _
                "Not processed: " & (n - lastProcessed) & " file(s)" & vbCrLf & _
                "See the results sheet: " & report.Name & "" & vbCrLf & _
                "Files already saved are not rolled back." & vbCrLf & _
                "This control workbook has not been saved automatically.", vbInformation, "A1 View Reset"
        End If
    End If
    If Len(restoreIssues) > 0 Then
        MsgBox "Some Excel settings could not be restored:" & vbCrLf & restoreIssues & vbCrLf & _
            "Safely save other open workbooks, then close and restart Excel.", vbExclamation
    End If
End Sub

Private Function GetControlSheet() As Worksheet
    Dim ws As Worksheet, legacyName As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    If ws Is Nothing Then
        ' Recognize the original four-character Japanese tab without non-ASCII source.
        legacyName = ChrW(&H64CD) & ChrW(&H4F5C) & ChrW(&H753B) & ChrW(&H9762)
        Set ws = ThisWorkbook.Worksheets(legacyName)
    End If
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 107, , "Control sheet not found. Rename the input sheet to Control."
    End If
    Set GetControlSheet = ws
End Function

Private Function CollectTargets(ByVal fso As Object, ByVal folder As String, ByRef paths() As String) As Long
    Dim file As Object, ext As String, n As Long
    ReDim paths(1 To 128)
    For Each file In fso.GetFolder(folder).Files
        ext = LCase$(fso.GetExtensionName(file.Name))
        Select Case ext
            Case "xlsx", "xlsm", "xls", "xlsb"
                If Left$(file.Name, 2) <> "~$" Then
                    If StrComp(file.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
                        n = n + 1
                        If n > UBound(paths) Then ReDim Preserve paths(1 To UBound(paths) + 128)
                        paths(n) = file.Path
                    End If
                End If
        End Select
    Next file
    If n > 0 Then ReDim Preserve paths(1 To n)
    CollectTargets = n
End Function

Private Sub SortPaths(ByRef items() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As String, tmp As String
    i = lo: j = hi: pivot = items(lo + (hi - lo) \ 2)
    Do While i <= j
        Do While StrComp(items(i), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(items(j), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            tmp = items(i): items(i) = items(j): items(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortPaths items, lo, j
    If i < hi Then SortPaths items, i, hi
End Sub

Private Function ConfirmTargets(ByVal fso As Object, ByVal folder As String, _
    ByRef paths() As String, ByVal n As Long) As Boolean
    Dim shown As Long, i As Long, page As String, line As String, info As String
    shown = n
    If shown > PREVIEW_LIMIT Then shown = PREVIEW_LIMIT
    info = "Total: " & n & " files. Previewing the first " & shown & " files." & vbCrLf & _
        "OK on a preview page does not start processing." & vbCrLf & _
        "Cancel exits without opening target workbooks." & vbCrLf & vbCrLf
    page = info
    For i = 1 To shown
        line = CStr(i) & ". " & fso.GetFileName(paths(i)) & vbCrLf
        If Len(page & line) > PREVIEW_CHARS And Len(page) > Len(info) Then
            If MsgBox(page & vbCrLf & "OK: Next / Cancel: Stop", _
                vbOKCancel + vbDefaultButton2 + vbInformation, "Files selected for processing") <> vbOK Then Exit Function
            page = info
        End If
        page = page & line
    Next i
    If MsgBox(page & vbCrLf & "OK: Final confirmation / Cancel: Stop", _
        vbOKCancel + vbDefaultButton2 + vbInformation, "Files selected for processing") <> vbOK Then Exit Function

    info = "Folder: " & folder & vbCrLf & _
        "Process ALL " & n & " files in this folder (no subfolders)." & vbCrLf
    If n > shown Then info = info & "Includes " & (n - shown) & " files not shown above." & vbCrLf
    info = info & vbCrLf & "Visit visible worksheets from right to left." & vbCrLf & _
        "Select A1, scroll top-left, OVERWRITE ORIGINALS, then close." & vbCrLf & _
        "Keep hidden/chart sheets unchanged. Do not remove protection." & vbCrLf & _
        "NO automatic backups. Make your own copies first." & vbCrLf & _
        "Stopping after start does NOT undo files already saved." & vbCrLf & _
        "Use trusted local files. Never enable target macros or external content." & vbCrLf & vbCrLf & _
        "Start processing ALL " & n & " files?"
    ConfirmTargets = (MsgBox(info, vbOKCancel + vbDefaultButton2 + vbExclamation, _
        "Confirm overwrite of ALL " & n & " files") = vbOK)
End Function

Private Function ProcessOneFile(ByVal fso As Object, ByVal path As String, ByVal expectedSize As Double, _
    ByVal expectedStamp As Date, ByRef detail As String, ByRef stopBatch As Boolean) As String
    Dim wb As Workbook, win As Window, sh As Object, ws As Worksheet
    Dim visibleWindows As Long, visibleSheets As Long, hiddenSheets As Long, otherSheets As Long
    Dim windowIndex As Long, sheetIndex As Long, saveAttempted As Boolean, saved As Boolean
    Dim result As String, reason As String, errNo As Long, errText As String
    On Error GoTo Failed
    result = "Skipped"
    If Not fso.FileExists(path) Then detail = "The file no longer exists.": GoTo CleanUp
    If Not SnapshotMatches(fso, path, expectedSize, expectedStamp) Then
        detail = "The file changed after the preview was prepared.": GoTo CleanUp
    End If
    If (CLng(fso.GetFile(path).Attributes) And 1) <> 0 Then
        detail = "The file has a read-only attribute.": GoTo CleanUp
    End If
    If WorkbookAlreadyOpen(fso, path) Then
        detail = "This file or a workbook with the same name is already open. It will not be modified.": GoTo CleanUp
    End If
    If fso.FileExists(fso.BuildPath(fso.GetParentFolderName(path), "~$" & fso.GetFileName(path))) Then
        detail = "A lock file exists. Skip the file even if the lock might be stale.": GoTo CleanUp
    End If

    Application.AutomationSecurity = SECURITY_FORCE_DISABLE
    Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False, _
        Password:=vbNullString, WriteResPassword:=vbNullString, IgnoreReadOnlyRecommended:=True, _
        Notify:=False, AddToMru:=False, CorruptLoad:=xlNormalLoad)
    ' Target VBA is disabled. Always cancel any Excel 4.0 macro warning.
    If wb.ReadOnly Then detail = "The workbook opened read-only and will not be saved.": GoTo CleanUp
    If wb.IsAddin Then detail = "Workbooks opened as add-ins are excluded.": GoTo CleanUp
    If wb.Signatures.Count > 0 Then detail = "Digitally signed workbooks are excluded to protect their signatures.": GoTo CleanUp
    If wb.MultiUserEditing Then detail = "Legacy shared workbooks are excluded.": GoTo CleanUp
    If AutoSaveIsOn(wb) Then detail = "Workbooks with AutoSave enabled are excluded.": GoTo CleanUp
    If wb.Excel4MacroSheets.Count > 0 Or wb.Excel4IntlMacroSheets.Count > 0 Then
        detail = "The workbook contains Excel 4.0 macro sheets and is excluded.": GoTo CleanUp
    End If
    If wb.Connections.Count > 0 Then
        detail = "External data connections are present. Excluded; refresh-on-open cannot be fully prevented.": GoTo CleanUp
    End If
    If HasQueryTables(wb) Then detail = "External-data QueryTables are present; the workbook is excluded.": GoTo CleanUp

    ' If any visible worksheet cannot show A1, do not save this workbook.
    For Each sh In wb.Sheets
        If TypeOf sh Is Worksheet Then
            Set ws = sh
            If ws.Type = xlWorksheet And ws.Visible = xlSheetVisible Then
                visibleSheets = visibleSheets + 1
                reason = SheetBlockReason(ws)
                If Len(reason) > 0 Then detail = ws.Name & ": " & reason: GoTo CleanUp
            Else
                hiddenSheets = hiddenSheets + 1
            End If
        Else
            otherSheets = otherSheets + 1
        End If
    Next sh
    If visibleSheets = 0 Then detail = "There are no visible normal worksheets.": GoTo CleanUp
    For Each win In wb.Windows
        If win.Visible Then visibleWindows = visibleWindows + 1
    Next win
    If visibleWindows = 0 Then detail = "There are no visible workbook windows.": GoTo CleanUp

    For windowIndex = wb.Windows.Count To 1 Step -1
        Set win = wb.Windows(windowIndex)
        If win.Visible Then
            win.Activate
            For sheetIndex = wb.Sheets.Count To 1 Step -1
                Set sh = wb.Sheets(sheetIndex)
                If TypeOf sh Is Worksheet Then
                    Set ws = sh
                    If ws.Type = xlWorksheet And ws.Visible = xlSheetVisible Then ResetSheetView ws, win
                End If
            Next sheetIndex
        End If
    Next windowIndex
    If mCancelRequested Then Err.Raise 18, , "Stop requested"
    If Not SnapshotMatches(fso, path, expectedSize, expectedStamp) Then
        Err.Raise vbObjectError + 111, , "The original file changed after opening. Do not save it."
    End If
    saveAttempted = True
    wb.Save
    If Not wb.Saved Then Err.Raise vbObjectError + 112, , "Could not confirm that saving completed."
    saved = True
    result = "Saved"
    detail = "Visible normal worksheets: " & visibleSheets & " / Windows: " & visibleWindows & _
        " / Hidden or other worksheets unchanged: " & hiddenSheets & " / Charts or other sheets unchanged: " & otherSheets & "."
    GoTo CleanUp

Failed:
    errNo = Err.Number
    errText = Err.Description
    If saveAttempted Then
        result = "Save unconfirmed"
        detail = "Save error " & errNo & ": " & errText & ". Check the original file. No automatic retry."
        stopBatch = True
    ElseIf errNo = 18 Then
        result = "Stopped before save"
        detail = "This file will be closed without saving."
        stopBatch = True
    Else
        result = "Failed before save"
        detail = CStr(errNo) & ": " & errText & ". This file will be closed without saving."
    End If
CleanUp:
    ' Close only the workbook opened by this routine and held in wb. Do not save again.
    On Error Resume Next
    If Not wb Is Nothing Then
        Err.Clear
        wb.Close SaveChanges:=False
        If Err.Number <> 0 Then
            If saved Then result = "Saved; close failed"
            detail = detail & " Close failed: " & Err.Description & ". Check the workbook that is still open."
            stopBatch = True
        End If
    End If
    On Error GoTo 0
    ProcessOneFile = result
End Function

Private Sub ResetSheetView(ByVal ws As Worksheet, ByVal win As Window)
    Dim pane As Pane, seen As Boolean, paneIndex As Long
    ws.Select Replace:=True ' Ungroup selected sheets and process one worksheet at a time.
    Application.Goto Reference:=ws.Range("A1"), Scroll:=False
    If win.FreezePanes Then
        ' Keep frozen panes and reset the scrollable area to its top-left position.
        win.ScrollRow = 1
        win.ScrollColumn = 1
    Else
        ' For ordinary split windows, reset every pane to the top-left position.
        For paneIndex = 1 To win.Panes.Count
            Set pane = win.Panes(paneIndex)
            pane.ScrollRow = 1
            pane.ScrollColumn = 1
        Next paneIndex
    End If
    If Not ActiveSheet Is ws Then Err.Raise vbObjectError + 120, , "Could not select the worksheet."
    If ActiveCell.Row <> 1 Or ActiveCell.Column <> 1 Then
        Err.Raise vbObjectError + 121, , "Could not select A1: " & ws.Name
    End If
    For paneIndex = 1 To win.Panes.Count
            Set pane = win.Panes(paneIndex)
        If Not Application.Intersect(pane.VisibleRange, ws.Range("A1")) Is Nothing Then seen = True
    Next paneIndex
    If Not seen Then Err.Raise vbObjectError + 122, , "Could not display A1: " & ws.Name
End Sub

Private Function SheetBlockReason(ByVal ws As Worksheet) As String
    If ws.Rows(1).Hidden Or ws.Columns(1).Hidden Then
        SheetBlockReason = "Row 1 or column A is hidden. Visibility settings will not be changed."
    ElseIf ws.DisplayRightToLeft Then
        SheetBlockReason = "The worksheet uses right-to-left display. Its direction will not be changed."
    ElseIf ws.ProtectContents Then
        If ws.EnableSelection = xlNoSelection Then
            SheetBlockReason = "Protection prevents cell selection."
        ElseIf ws.EnableSelection = xlUnlockedCells Then
            If ws.Range("A1").Locked Then SheetBlockReason = "Protected cell A1 cannot be selected."
        End If
    End If
End Function

Private Function WorkbookAlreadyOpen(ByVal fso As Object, ByVal path As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Or _
            StrComp(wb.Name, fso.GetFileName(path), vbTextCompare) = 0 Then
            WorkbookAlreadyOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function SnapshotMatches(ByVal fso As Object, ByVal path As String, _
    ByVal expectedSize As Double, ByVal expectedStamp As Date) As Boolean
    If Not fso.FileExists(path) Then Exit Function
    SnapshotMatches = (CDbl(fso.GetFile(path).Size) = expectedSize And _
        fso.GetFile(path).DateLastModified = expectedStamp)
End Function

Private Function AutoSaveIsOn(ByVal wb As Workbook) As Boolean
    Dim value As Variant, errNo As Long, errText As String
    On Error Resume Next
    value = CallByName(wb, "AutoSaveOn", VbGet)
    errNo = Err.Number: errText = Err.Description
    On Error GoTo 0
    If errNo = 438 Or errNo = 445 Then Exit Function ' Unsupported in older Excel versions.
    If errNo <> 0 Then Err.Raise errNo, , "Could not read the AutoSave state: " & errText
    AutoSaveIsOn = CBool(value)
End Function

Private Function HasQueryTables(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.QueryTables.Count > 0 Then HasQueryTables = True: Exit Function
    Next ws
End Function

Private Function NewReportSheet() As Worksheet
    Dim ws As Worksheet, base As String, candidate As String, k As Long
    base = "A1Results_" & Format$(Now, "mmdd_hhnnss")
    candidate = base
    Do While SheetNameExists(candidate)
        k = k + 1: candidate = base & "_" & CStr(k)
    Loop
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = candidate
    ws.Columns("A:E").NumberFormat = "@" ' Store file names and paths as text, not formulas.
    ws.Range("A4").Value2 = "No."
    ws.Range("B4").Value2 = "File name"
    ws.Range("C4").Value2 = "Result"
    ws.Range("D4").Value2 = "Details"
    ws.Range("E4").Value2 = "Original file path"
    ws.Columns("A").ColumnWidth = 7
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 23
    ws.Columns("D").ColumnWidth = 75
    ws.Columns("E").ColumnWidth = 65
    ws.Rows(4).Font.Bold = True
    Set NewReportSheet = ws
End Function

Private Function SheetNameExists(ByVal name As String) As Boolean
    Dim sh As Object
    For Each sh In ThisWorkbook.Sheets
        If StrComp(sh.Name, name, vbTextCompare) = 0 Then SheetNameExists = True: Exit Function
    Next sh
End Function

Private Sub CaptureApplication(ByRef state As ApplicationState)
    With Application
        state.EnableEvents = .EnableEvents
        state.ScreenUpdating = .ScreenUpdating
        state.DisplayAlerts = .DisplayAlerts
        state.AutomationSecurity = .AutomationSecurity
        state.EnableCancelKey = .EnableCancelKey
        state.StatusBar = .StatusBar
    End With
End Sub

Private Function RestoreApplication(ByRef state As ApplicationState) As String
    Dim issues As String
    On Error Resume Next
    Err.Clear: Application.AutomationSecurity = state.AutomationSecurity
    If Err.Number <> 0 Then issues = issues & "AutomationSecurity: " & Err.Description & vbCrLf
    Err.Clear: Application.DisplayAlerts = state.DisplayAlerts
    If Err.Number <> 0 Then issues = issues & "DisplayAlerts: " & Err.Description & vbCrLf
    Err.Clear: Application.StatusBar = state.StatusBar
    If Err.Number <> 0 Then issues = issues & "StatusBar: " & Err.Description & vbCrLf
    Err.Clear: Application.EnableCancelKey = state.EnableCancelKey
    If Err.Number <> 0 Then issues = issues & "EnableCancelKey: " & Err.Description & vbCrLf
    Err.Clear: Application.ScreenUpdating = state.ScreenUpdating
    If Err.Number <> 0 Then issues = issues & "ScreenUpdating: " & Err.Description & vbCrLf
    Err.Clear: Application.EnableEvents = state.EnableEvents
    If Err.Number <> 0 Then issues = issues & "EnableEvents: " & Err.Description & vbCrLf
    On Error GoTo 0
    RestoreApplication = issues
End Function
