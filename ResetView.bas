Attribute VB_Name = "ResetView"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const FOLDER_CELL As String = "B2"
Private Const PREVIEW_LIMIT As Long = 30

Public Sub RunA1ViewReset()
    Dim folderPath As String
    Dim files As Collection
    Dim errorMessage As String
    Dim preview As String
    Dim answer As VbMsgBoxResult
    Dim filePath As Variant
    Dim failedFiles As String
    Dim processedCount As Long
    Dim failedCount As Long
    Dim previewCount As Long
    Dim settingsCaptured As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAutomationSecurity As MsoAutomationSecurity

    On Error GoTo Failed

    folderPath = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET).Range(FOLDER_CELL).Value))
    If Len(folderPath) = 0 Then
        MsgBox "Enter the target folder path in " & CONTROL_SHEET & "!" & FOLDER_CELL & ".", _
               vbExclamation, "A1 View Reset"
        Exit Sub
    End If

    Set files = GetTargetFiles(folderPath, errorMessage)
    If files Is Nothing Then
        MsgBox errorMessage, vbExclamation, "A1 View Reset"
        Exit Sub
    End If

    If files.Count = 0 Then
        MsgBox "No supported Excel files were found in the target folder.", _
               vbInformation, "A1 View Reset"
        Exit Sub
    End If

    preview = "Target files: " & files.Count & vbCrLf & vbCrLf
    For Each filePath In files
        If previewCount >= PREVIEW_LIMIT Then Exit For
        preview = preview & FileNameOnly(CStr(filePath)) & vbCrLf
        previewCount = previewCount + 1
    Next filePath
    If files.Count > PREVIEW_LIMIT Then
        preview = preview & "... (" & (files.Count - PREVIEW_LIMIT) & " more files)" & vbCrLf
    End If
    preview = preview & vbCrLf & "Continue? The original files will be overwritten."

    answer = MsgBox(preview, vbYesNo Or vbQuestion Or vbDefaultButton2, "Review Excel files")
    If answer <> vbYes Then Exit Sub

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAutomationSecurity = Application.AutomationSecurity
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    For Each filePath In files
        errorMessage = vbNullString
        If ProcessWorkbook(CStr(filePath), errorMessage) Then
            processedCount = processedCount + 1
        Else
            failedCount = failedCount + 1
            If failedCount <= 10 Then
                failedFiles = failedFiles & vbCrLf & FileNameOnly(CStr(filePath)) & ": " & errorMessage
            End If
        End If
    Next filePath

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AutomationSecurity = oldAutomationSecurity

    preview = "Processing complete." & vbCrLf & _
              "Succeeded: " & processedCount & vbCrLf & _
              "Failed: " & failedCount
    If failedCount > 0 Then
        preview = preview & vbCrLf & failedFiles
        If failedCount > 10 Then preview = preview & vbCrLf & "... and " & (failedCount - 10) & " more."
    End If
    MsgBox preview, IIf(failedCount > 0, vbExclamation, vbInformation), "A1 View Reset"
    Exit Sub

Failed:
    errorMessage = Err.Description
    On Error Resume Next
    If settingsCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.DisplayAlerts = oldDisplayAlerts
        Application.AutomationSecurity = oldAutomationSecurity
    End If
    On Error GoTo 0
    MsgBox "The operation could not continue: " & errorMessage, vbCritical, "A1 View Reset"
End Sub

Private Function GetTargetFiles(ByVal folderPath As String, ByRef errorMessage As String) As Collection
    Dim fso As Object
    Dim targetFolder As Object
    Dim file As Object
    Dim files As Collection
    Dim toolPath As String
    Dim extension As String

    On Error GoTo InvalidFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)
    Set files = New Collection

    If Len(ThisWorkbook.Path) > 0 Then
        toolPath = LCase$(fso.GetAbsolutePathName(ThisWorkbook.FullName))
    End If

    For Each file In targetFolder.Files
        extension = LCase$(fso.GetExtensionName(file.Name))
        If IsExcelExtension(extension) And Left$(file.Name, 2) <> "~$" Then
            If Len(toolPath) = 0 Or LCase$(fso.GetAbsolutePathName(file.Path)) <> toolPath Then
                files.Add file.Path
            End If
        End If
    Next file

    Set GetTargetFiles = files
    Exit Function

InvalidFolder:
    errorMessage = "The target folder could not be read: " & Err.Description
    Set GetTargetFiles = Nothing
End Function

Private Function IsExcelExtension(ByVal extension As String) As Boolean
    Select Case extension
        Case "xls", "xlsx", "xlsm", "xlsb"
            IsExcelExtension = True
    End Select
End Function

Private Function FileNameOnly(ByVal filePath As String) As String
    FileNameOnly = Mid$(filePath, InStrRev(filePath, Application.PathSeparator) + 1)
End Function

Private Function ProcessWorkbook(ByVal filePath As String, ByRef errorMessage As String) As Boolean
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim index As Long
    Dim visibleCount As Long

    On Error GoTo Failed
    Set workbook = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=False, _
                                  AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    If workbook.ReadOnly Then Err.Raise vbObjectError + 1000, "ProcessWorkbook", "The workbook opened as read-only."

    For index = workbook.Worksheets.Count To 1 Step -1
        Set worksheet = workbook.Worksheets(index)
        If worksheet.Visible = xlSheetVisible Then
            visibleCount = visibleCount + 1
            worksheet.Activate
            worksheet.Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next index

    If visibleCount = 0 Then Err.Raise vbObjectError + 1001, "ProcessWorkbook", "No visible worksheets were found."

    workbook.Save
    workbook.Close SaveChanges:=False
    ProcessWorkbook = True
    Exit Function

Failed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not workbook Is Nothing Then workbook.Close SaveChanges:=False
    On Error GoTo 0
End Function
