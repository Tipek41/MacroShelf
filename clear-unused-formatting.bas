Attribute VB_Name = "ClearUnusedFormatting"
Option Explicit

' Clears formatting outside the active sheet's formula/value area.
' Cell contents and formulas are preserved. Blank-result formulas count as used cells.
Public Sub VeriDisiTemizleVeSifirla()
    Dim ws As Worksheet
    Dim lastRowCell As Range, lastColCell As Range, refreshedRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim settingsCaptured As Boolean, completed As Boolean
    Dim errorNumber As Long, errorDescription As String
    Dim answer As VbMsgBoxResult

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Select a worksheet and try again.", vbExclamation, "Clear Unused Formatting"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If ws.ProtectContents Then
        MsgBox "This worksheet is protected. Unprotect it before running this macro.", _
               vbExclamation, "Clear Unused Formatting"
        Exit Sub
    End If

    ' Search formulas as well as values, including formulas that currently return "".
    Set lastRowCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False)
    Set lastColCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False)

    If lastRowCell Is Nothing Or lastColCell Is Nothing Then
        MsgBox "No values or formulas were found. No changes were made.", _
               vbInformation, "Clear Unused Formatting"
        Exit Sub
    End If

    lastRow = lastRowCell.Row
    lastCol = lastColCell.Column
    answer = MsgBox("Clear formatting outside the data area on '" & ws.Name & "'?" & vbCrLf & _
        "Cell values and formulas will be kept.", vbQuestion + vbYesNo + vbDefaultButton2, _
        "Clear Unused Formatting")
    If answer <> vbYes Then Exit Sub

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    settingsCaptured = True

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Clear formats only. Do not delete rows or columns, so objects and cell positions stay put.
    If lastRow < ws.Rows.Count Then
        ws.Rows(CStr(lastRow + 1) & ":" & CStr(ws.Rows.Count)).ClearFormats
    End If
    If lastCol < ws.Columns.Count Then
        ws.Range(ws.Cells(1, lastCol + 1), ws.Cells(1, ws.Columns.Count)).EntireColumn.ClearFormats
    End If

    ' Ask Excel to refresh its used-range calculation.
    Set refreshedRange = ws.UsedRange
    completed = True
    GoTo CleanExit

ErrorHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description

CleanExit:
    If settingsCaptured Then
        On Error Resume Next
        Application.Calculation = oldCalculation
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
        On Error GoTo 0
    End If

    If errorNumber <> 0 Then
        MsgBox "Formatting cleanup could not be completed:" & vbCrLf & errorDescription, _
               vbCritical, "Clear Unused Formatting"
    ElseIf completed Then
        MsgBox "Formatting outside the data area was cleared." & vbCrLf & _
               "Last formula/value cell: " & ws.Cells(lastRow, lastCol).Address(False, False) & vbCrLf & _
               "Save and reopen the workbook if Excel still shows an inflated used range.", _
               vbInformation, "Clear Unused Formatting"
    End If
End Sub
