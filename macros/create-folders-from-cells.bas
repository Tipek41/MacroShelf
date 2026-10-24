Attribute VB_Name = "CreateFoldersFromCells"
Option Explicit

' Create one folder per selected Excel cell.
' Windows Excel | No additional references required.
' The macro never deletes or overwrites existing folders.
Public Sub CreateFoldersFromSelectedCells()
    Dim selectedCells As Range, cell As Range
    Dim parentPath As String, folderName As String, fullPath As String
    Dim fso As Object, picker As FileDialog
    Dim created As Long, existing As Long, skipped As Long, failed As Long
    Dim errors As String, answer As VbMsgBoxResult, rawValue As String

    On Error Resume Next
    Set selectedCells = Application.InputBox( _
        Prompt:="Select cells containing folder names (one folder per cell).", _
        Title:="Step 1 of 2 - Select folder names", Type:=8)
    On Error GoTo FatalError
    If selectedCells Is Nothing Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    With picker
        .Title = "Step 2 of 2 - Choose the destination folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        parentPath = .SelectedItems(1)
    End With
    If Not fso.FolderExists(parentPath) Then
        MsgBox "The selected destination folder does not exist.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox("Create folders from " & selectedCells.Cells.CountLarge & _
        " selected cell(s) inside:" & vbCrLf & parentPath & vbCrLf & vbCrLf & _
        "Existing folders will be left unchanged. Continue?", _
        vbQuestion + vbYesNo, "Confirm folder creation")
    If answer <> vbYes Then Exit Sub

    For Each cell In selectedCells.Cells
        On Error GoTo CellError
        If IsError(cell.Value2) Then GoTo SkipCell
        rawValue = Trim$(CStr(cell.Value2))
        If Len(rawValue) = 0 Then GoTo SkipCell
        folderName = SafeFolderName(rawValue)
        If Len(folderName) = 0 Then GoTo SkipCell
        fullPath = fso.BuildPath(parentPath, folderName)
        If Len(fullPath) > 240 Then GoTo SkipCell
        If fso.FolderExists(fullPath) Then
            existing = existing + 1
        ElseIf fso.FileExists(fullPath) Then
            failed = failed + 1
            If Len(errors) < 600 Then errors = errors & vbCrLf & cell.Address(False, False) & ": a file has this name"
        Else
            fso.CreateFolder fullPath
            created = created + 1
        End If
NextCell:
        On Error GoTo FatalError
    Next cell

    MsgBox "Folder creation finished." & vbCrLf & vbCrLf & _
        "Created: " & created & vbCrLf & _
        "Already exists: " & existing & vbCrLf & _
        "Skipped (empty/invalid/too long): " & skipped & vbCrLf & _
        "Failed: " & failed & IIf(Len(errors) > 0, vbCrLf & errors, ""), _
        vbInformation, "Results"
    Exit Sub

SkipCell:
    skipped = skipped + 1
    GoTo NextCell

CellError:
    failed = failed + 1
    If Len(errors) < 600 Then errors = errors & vbCrLf & cell.Address(False, False) & ": " & Err.Description
    Err.Clear
    Resume NextCell

FatalError:
    MsgBox "Operation stopped: " & Err.Description, vbCritical, "Error"
End Sub

Private Function SafeFolderName(ByVal nameText As String) As String
    Dim invalid As Variant, item As Variant, i As Long, ch As String
    invalid = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each item In invalid
        nameText = Replace$(nameText, CStr(item), "_")
    Next item
    For i = 0 To 31
        nameText = Replace$(nameText, Chr$(i), "_")
    Next i
    nameText = Trim$(nameText)
    Do While Len(nameText) > 0
        ch = Right$(nameText, 1)
        If ch <> "." And ch <> " " Then Exit Do
        nameText = Left$(nameText, Len(nameText) - 1)
    Loop
    If Len(nameText) = 0 Then Exit Function
    If Len(nameText) > 120 Then nameText = Left$(nameText, 120)
    Dim baseName As String
    baseName = UCase$(Split(nameText, ".")(0))
    Select Case baseName
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            nameText = "_" & nameText
    End Select
    SafeFolderName = nameText
End Function
