Attribute VB_Name = "HideTextMatchingFill"
Option Explicit

Private Const BACKUP_MARKER As String = "MACROSHELF_FONT_BACKUP_V1"

' Visual-only concealment: font color is matched to the displayed fill color.
' This does not protect or encrypt cell contents.
Public Sub MetniArkaPlanRengineGoreGizle()
    Dim selectedRange As Range, contentCells As Range, part As Range, cell As Range
    Dim ws As Worksheet, wb As Workbook, backup As Worksheet
    Dim seen As Object, lastBackupRow As Long, nextRow As Long
    Dim key As String, changedCount As Long
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation, settingsCaptured As Boolean
    Dim errText As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select one or more worksheet cells first.", vbExclamation, "MacroShelf"
        Exit Sub
    End If
    Set selectedRange = Selection
    Set ws = selectedRange.Worksheet
    Set wb = ws.Parent

    If ws.ProtectContents Then
        MsgBox "Unprotect the worksheet before running this macro.", vbExclamation, "MacroShelf"
        Exit Sub
    End If

    ' Use SpecialCells so a very large selection does not require looping through every blank cell.
    On Error Resume Next
    Set contentCells = selectedRange.SpecialCells(xlCellTypeFormulas)
    Set part = selectedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If Not part Is Nothing Then
        If contentCells Is Nothing Then
            Set contentCells = part
        Else
            Set contentCells = Application.Union(contentCells, part)
        End If
    End If
    If contentCells Is Nothing Then
        MsgBox "No values or formulas were found in the selection. No changes were made.", _
               vbInformation, "MacroShelf"
        Exit Sub
    End If

    If MsgBox("Match the font color to the displayed fill color for non-empty selected cells?" & _
              vbCrLf & vbCrLf & "This only hides text visually; it does not secure the cell contents." & _
              vbCrLf & "Original font colors will be stored on a very-hidden backup sheet.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Hide Text by Fill Color") <> vbYes Then Exit Sub

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    settingsCaptured = True

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set backup = GetOrCreateBackupSheet(wb)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    lastBackupRow = backup.Cells(backup.Rows.Count, 1).End(xlUp).Row

    For nextRow = 2 To lastBackupRow
        key = CStr(backup.Cells(nextRow, 1).Value2) & "!" & CStr(backup.Cells(nextRow, 2).Value2)
        If Len(key) > 1 Then seen(key) = True
    Next nextRow

    For Each cell In contentCells.Cells
        If cell.MergeCells Then
            If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then GoTo NextSelectedCell
        End If
        key = ws.CodeName & "!" & cell.Address(False, False)
        If Not seen.Exists(key) Then
            lastBackupRow = backup.Cells(backup.Rows.Count, 1).End(xlUp).Row + 1
            backup.Cells(lastBackupRow, 1).Value2 = ws.CodeName
            backup.Cells(lastBackupRow, 2).Value2 = cell.Address(False, False)
            backup.Cells(lastBackupRow, 3).Value2 = cell.Font.Color
                seen(key) = True
        End If
        cell.Font.Color = cell.DisplayFormat.Interior.Color
        cell.Font.TintAndShade = 0
        changedCount = changedCount + 1
NextSelectedCell:
    Next cell

    backup.Visible = xlSheetVeryHidden
    GoTo CleanExit

ErrorHandler:
    errText = Err.Description

CleanExit:
    If settingsCaptured Then
        On Error Resume Next
        Application.Calculation = oldCalculation
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
        On Error GoTo 0
    End If

    If Len(errText) > 0 Then
        MsgBox "The macro stopped: " & errText & vbCrLf & _
               "If any cells changed, run RestoreOriginalFontColors to restore saved colors.", _
               vbCritical, "MacroShelf"
    ElseIf changedCount = 0 Then
        MsgBox "No non-empty cells were found in the selection.", vbInformation, "MacroShelf"
    Else
        MsgBox changedCount & " non-empty cell(s) were changed." & vbCrLf & _
               "This is visual concealment only. Run RestoreOriginalFontColors to undo it.", _
               vbInformation, "MacroShelf"
    End If
End Sub

' Restores the original font color and tint saved by MetniArkaPlanRengineGoreGizle.
Public Sub RestoreOriginalFontColors()
    Dim wb As Workbook, backup As Worksheet, ws As Worksheet
    Dim lastRow As Long, i As Long, restoredCount As Long
    Dim codeName As String, addressText As String
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation, settingsCaptured As Boolean
    Dim errText As String

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set wb = ActiveSheet.Parent
    Set backup = FindBackupSheet(wb)
    If backup Is Nothing Then
        MsgBox "No saved font-color backup was found in this workbook.", vbInformation, "MacroShelf"
        Exit Sub
    End If

    lastRow = backup.Cells(backup.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "There are no saved font colors to restore.", vbInformation, "MacroShelf"
        Exit Sub
    End If
    If MsgBox("Restore saved font colors in this workbook?", vbQuestion + vbYesNo + vbDefaultButton2, _
              "Restore Font Colors") <> vbYes Then Exit Sub

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    settingsCaptured = True
    On Error GoTo RestoreError
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For i = 2 To lastRow
        codeName = CStr(backup.Cells(i, 1).Value2)
        addressText = CStr(backup.Cells(i, 2).Value2)
        Set ws = FindWorksheetByCodeName(wb, codeName)
        If Not ws Is Nothing Then
            With ws.Range(addressText).Font
                .Color = CLng(backup.Cells(i, 3).Value2)
            End With
            restoredCount = restoredCount + 1
        End If
        Set ws = Nothing
    Next i
    backup.Rows("2:" & lastRow).ClearContents
    backup.Visible = xlSheetVeryHidden
    GoTo RestoreCleanExit

RestoreError:
    errText = Err.Description

RestoreCleanExit:
    If settingsCaptured Then
        On Error Resume Next
        Application.Calculation = oldCalculation
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
        On Error GoTo 0
    End If
    If Len(errText) > 0 Then
        MsgBox "Font-color restoration stopped: " & errText & vbCrLf & _
               "The backup remains available; unprotect the affected sheet and try again.", _
               vbCritical, "MacroShelf"
    Else
        MsgBox restoredCount & " cell(s) restored.", vbInformation, "MacroShelf"
    End If
End Sub

Private Function GetOrCreateBackupSheet(ByVal wb As Workbook) As Worksheet
    Dim backup As Worksheet, candidate As String, suffix As Long
    Set backup = FindBackupSheet(wb)
    If Not backup Is Nothing Then
        Set GetOrCreateBackupSheet = backup
        Exit Function
    End If

    candidate = "_MSFontBackup"
    Do While WorksheetNameExists(wb, candidate)
        suffix = suffix + 1
        candidate = "_MSFontBackup" & CStr(suffix)
    Loop
    Set backup = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    backup.Name = candidate
    backup.Cells(1, 1).Value2 = BACKUP_MARKER
    backup.Cells(1, 2).Value2 = "Address"
    backup.Cells(1, 3).Value2 = "FontColor"
    backup.Visible = xlSheetVeryHidden
    Set GetOrCreateBackupSheet = backup
End Function

Private Function FindBackupSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If CStr(ws.Cells(1, 1).Value2) = BACKUP_MARKER Then
            Set FindBackupSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FindWorksheetByCodeName(ByVal wb As Workbook, ByVal targetCodeName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.CodeName, targetCodeName, vbTextCompare) = 0 Then
            Set FindWorksheetByCodeName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function WorksheetNameExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetNameExists = Not ws Is Nothing
    On Error GoTo 0
End Function
