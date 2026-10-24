Attribute VB_Name = "WeightedDistribution"
Option Explicit

' Weighted distribution with two-decimal rounding.
' Excel desktop, no external references. Existing formulas are preserved.
Public Sub DistributeValueByWeights()
    Dim wb As Workbook, source As Range, weights As Range, firstTarget As Range, targets As Range
    Dim picked As Range, cell As Range, sourceAmount As Double, totalWeight As Double
    Dim weightValues() As Double, shares() As Double
    Dim i As Long, n As Long, lastPositive As Long, distributed As Double
    Dim currentValue As Variant, answer As VbMsgBoxResult
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalc As XlCalculation
    Dim applied As Long, failureText As String, originalFormulas() As Variant
    Dim hasChanges As Boolean

    If Application.ActiveWorkbook Is Nothing Then
        MsgBox "Open an Excel workbook first.", vbExclamation: Exit Sub
    End If
    Set wb = Application.ActiveWorkbook

    Set source = PickRange("1/3 - Select one cell containing the amount to distribute")
    If source Is Nothing Then Exit Sub
    If source.Cells.CountLarge <> 1 Then
        MsgBox "Select exactly one source cell.", vbExclamation: Exit Sub
    End If
    If Not source.Worksheet.Parent Is wb Then
        MsgBox "Select a source cell in the active workbook.", vbExclamation: Exit Sub
    End If
    If IsError(source.Value2) Then
        MsgBox "The source contains an Excel error.", vbExclamation: Exit Sub
    End If
    If Not IsNumeric(source.Value2) Or Len(Trim$(CStr(source.Value2))) = 0 Then
        MsgBox "The source must contain a numeric value.", vbExclamation: Exit Sub
    End If
    sourceAmount = CDbl(source.Value2)

    Set weights = PickRange("2/3 - Select weights in one continuous row")
    If weights Is Nothing Then Exit Sub
    If weights.Areas.Count <> 1 Or weights.Rows.Count <> 1 Then
        MsgBox "Select one continuous row of weights.", vbExclamation: Exit Sub
    End If
    If Not weights.Worksheet.Parent Is wb Then
        MsgBox "Weights must be in the active workbook.", vbExclamation: Exit Sub
    End If
    n = weights.Columns.Count
    ReDim weightValues(1 To n)
    ReDim shares(1 To n)
    For i = 1 To n
        currentValue = weights.Cells(1, i).Value2
        If IsError(currentValue) Then
            MsgBox "A weight contains an Excel error.", vbExclamation: Exit Sub
        End If
        If Len(Trim$(CStr(currentValue))) > 0 Then
            If Not IsNumeric(currentValue) Then
                MsgBox "All weights must be numeric or blank.", vbExclamation: Exit Sub
            End If
            weightValues(i) = CDbl(currentValue)
            If weightValues(i) < 0 Then
                MsgBox "Negative weights are not supported.", vbExclamation: Exit Sub
            End If
            totalWeight = totalWeight + weightValues(i)
            If weightValues(i) > 0 Then lastPositive = i
        End If
    Next i
    If totalWeight <= 0 Then
        MsgBox "The total weight must be greater than zero.", vbExclamation: Exit Sub
    End If

    Set firstTarget = PickRange("3/3 - Select the first destination cell")
    If firstTarget Is Nothing Then Exit Sub
    If firstTarget.Cells.CountLarge <> 1 Then
        MsgBox "Select exactly one destination cell.", vbExclamation: Exit Sub
    End If
    If Not firstTarget.Worksheet.Parent Is wb Then
        MsgBox "Destination must be in the active workbook.", vbExclamation: Exit Sub
    End If
    If firstTarget.Column + n - 1 > firstTarget.Worksheet.Columns.Count Then
        MsgBox "Not enough columns for the output.", vbExclamation: Exit Sub
    End If
    Set targets = firstTarget.Resize(1, n)
    If Not Intersect(targets, source) Is Nothing Then
        MsgBox "Destination cannot overlap the source cell.", vbExclamation: Exit Sub
    End If
    If targets.Worksheet Is weights.Worksheet Then
        If Not Intersect(targets, weights) Is Nothing Then
            MsgBox "Destination cannot overlap the weights.", vbExclamation: Exit Sub
        End If
    End If
    If targets.Worksheet.ProtectContents Then
        MsgBox "The destination worksheet is protected.", vbExclamation: Exit Sub
    End If
    If targets.MergeCells Then
        MsgBox "Merged destination cells are not supported.", vbExclamation: Exit Sub
    End If
    For Each cell In targets.Cells
        If IsError(cell.Value2) Then
            MsgBox "Destination contains an Excel error.", vbExclamation: Exit Sub
        End If
        If Not cell.HasFormula And Len(Trim$(CStr(cell.Value2))) > 0 Then
            If Not IsNumeric(cell.Value2) Then
                MsgBox "Destination cells must contain numbers, formulas or blanks.", vbExclamation: Exit Sub
            End If
        End If
        If cell.HasFormula And cell.HasArray Then
            MsgBox "Array formulas in the destination are not supported.", vbExclamation: Exit Sub
        End If
    Next cell

    ' Round each share to 2 decimals, then assign residual to the last positive weight.
    For i = 1 To n
        If i = lastPositive Then
            shares(i) = Application.WorksheetFunction.Round(sourceAmount - distributed, 2)
        ElseIf weightValues(i) > 0 Then
            shares(i) = Application.WorksheetFunction.Round(sourceAmount * weightValues(i) / totalWeight, 2)
        End If
        distributed = distributed + shares(i)
    Next i

    answer = MsgBox("Amount: " & Format(sourceAmount, "#,##0.00") & vbCrLf & _
        "Weights: " & weights.Address(External:=True) & vbCrLf & _
        "Destination: " & targets.Address(External:=True) & vbCrLf & _
        "Cells: " & n & vbCrLf & vbCrLf & _
        "Nonzero shares will be ADDED to existing values/formulas." & vbCrLf & _
        "Existing formulas will be wrapped in parentheses." & vbCrLf & _
        "Proceed?", vbYesNo + vbQuestion, "Confirm weighted distribution")
    If answer <> vbYes Then Exit Sub

    ReDim originalFormulas(1 To n)
    For i = 1 To n
        originalFormulas(i) = targets.Cells(1, i).Formula
    Next i
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = 1 To n
        If shares(i) <> 0 Then
            Set cell = targets.Cells(1, i)
            If cell.HasFormula Then
                cell.Formula = "=(" & Mid$(cell.Formula, 2) & ")+" & InvariantNumber(shares(i))
            ElseIf Len(Trim$(CStr(cell.Value2))) > 0 Then
                cell.Formula = "=" & InvariantNumber(CDbl(cell.Value2)) & "+" & InvariantNumber(shares(i))
            Else
                cell.Formula = "=" & InvariantNumber(shares(i))
            End If
            applied = applied + 1
            hasChanges = True
        End If
    Next i

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Distribution completed." & vbCrLf & _
        "Total distributed: " & Format(distributed, "#,##0.00") & vbCrLf & _
        "Updated cells: " & applied & vbCrLf & _
        "Destination: " & targets.Address(External:=True), vbInformation, "Done"
    Exit Sub

Failed:
    failureText = Err.Description
    If hasChanges Then
        On Error Resume Next
        For i = 1 To n
            targets.Cells(1, i).Formula = originalFormulas(i)
        Next i
        On Error GoTo 0
    End If
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Distribution failed: " & failureText & vbCrLf & _
        "Attempted to restore original destination cells.", vbCritical, "Error"
End Sub

Private Function PickRange(ByVal message As String) As Range
    On Error Resume Next
    Set PickRange = Application.InputBox(Prompt:=message, _
        Title:="Weighted Distribution", Type:=8)
    On Error GoTo 0
End Function

Private Function InvariantNumber(ByVal value As Double) As String
    ' Str$ always uses a period as decimal separator, unlike CStr.
    InvariantNumber = Trim$(Str$(value))
End Function