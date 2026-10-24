Attribute VB_Name = "WorksheetInventory"
Option Explicit

' Lists worksheet names and visibility in the active workbook.
' No external references required. Does not change sheet visibility.
Public Sub ListWorksheetsByVisibility()
    Dim wb As Workbook, destination As Range, outputSheet As Worksheet
    Dim ws As Worksheet, rowIndex As Long, groupIndex As Long
    Dim countVisible As Long, countHidden As Long, countVeryHidden As Long
    Dim statusText As String, oldEvents As Boolean, oldScreen As Boolean
    Dim answer As VbMsgBoxResult

    If Application.ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "Worksheet Inventory"
        Exit Sub
    End If
    Set wb = Application.ActiveWorkbook
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then
        MsgBox "Select a worksheet first.", vbExclamation, "Worksheet Inventory"
        Exit Sub
    End If

    On Error Resume Next
    Set destination = Application.InputBox( _
        Prompt:="Select the first output cell (3 columns will be used: Sheet, Visibility, Link).", _
        Title:="Worksheet Inventory - Output Location", Type:=8)
    On Error GoTo Fail
    If destination Is Nothing Then Exit Sub
    Set destination = destination.Cells(1, 1)
    If Not destination.Worksheet.Parent Is wb Then
        MsgBox "Choose a cell in the active workbook.", vbExclamation
        Exit Sub
    End If
    If destination.Column > destination.Worksheet.Columns.Count - 2 Then
        MsgBox "Choose a starting cell with at least 3 available columns.", vbExclamation
        Exit Sub
    End If
    If destination.Row > destination.Worksheet.Rows.Count - wb.Worksheets.Count - 2 Then
        MsgBox "Not enough rows below the selected cell.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox("This will write a header and up to " & wb.Worksheets.Count & _
        " sheet records in three columns starting at " & _
        destination.Address(External:=True) & "." & vbCrLf & _
        "Existing cell contents in this range will be replaced. Continue?", _
        vbYesNo + vbQuestion, "Confirm Output")
    If answer <> vbYes Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set outputSheet = destination.Worksheet

    destination.Value = "Worksheet"
    destination.Offset(0, 1).Value = "Visibility"
    destination.Offset(0, 2).Value = "Open sheet"
    rowIndex = 1

    ' Visible first, then Hidden, then Very Hidden.
    For groupIndex = 0 To 2
        For Each ws In wb.Worksheets
            If (groupIndex = 0 And ws.Visible = xlSheetVisible) Or _
               (groupIndex = 1 And ws.Visible = xlSheetHidden) Or _
               (groupIndex = 2 And ws.Visible = xlSheetVeryHidden) Then

                Select Case ws.Visible
                    Case xlSheetVisible
                        statusText = "Visible": countVisible = countVisible + 1
                    Case xlSheetHidden
                        statusText = "Hidden": countHidden = countHidden + 1
                    Case Else
                        statusText = "Very Hidden": countVeryHidden = countVeryHidden + 1
                End Select

                destination.Offset(rowIndex, 0).Value = ws.Name
                destination.Offset(rowIndex, 1).Value = statusText
                If ws.Visible = xlSheetVisible Then
                    outputSheet.Hyperlinks.Add _
                        Anchor:=destination.Offset(rowIndex, 2), _
                        Address:="", _
                        SubAddress:="'" & Replace(ws.Name, "'", "''") & "'!A1", _
                        TextToDisplay:="Open"
                Else
                    destination.Offset(rowIndex, 2).Value = "Not accessible while hidden"
                End If
                rowIndex = rowIndex + 1
            End If
        Next ws
    Next groupIndex

    With destination.Resize(rowIndex, 3)
        .Columns.AutoFit
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 227, 236)
    End With
    With destination.Resize(1, 3)
        .Font.Bold = True
        .Interior.Color = RGB(26, 47, 77)
        .Font.Color = RGB(255, 255, 255)
    End With

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Worksheet inventory created." & vbCrLf & _
        "Visible: " & countVisible & vbCrLf & _
        "Hidden: " & countHidden & vbCrLf & _
        "Very Hidden: " & countVeryHidden, vbInformation, "Done"
    Exit Sub

Fail:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not create the list: " & Err.Description, vbCritical, "Worksheet Inventory"
End Sub