Attribute VB_Name = "Module4"
Sub SummarizeResults()
    Dim ws As Worksheet
    Dim capacityRange As Range
    Dim metricRange As Range
    Dim maxMetric As Double
    Dim bestCapacity As Variant
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Outputs")
    Set capacityRange = ws.Range("D19:D25")
    Set metricRange = ws.Range("E19:E25")

    maxMetric = -1E+308
    bestCapacity = "Not Found" '

    For Each cell In metricRange
        If IsNumeric(cell.Value) And cell.Value > maxMetric Then
            maxMetric = cell.Value
            bestCapacity = capacityRange.Cells(cell.Row - metricRange.Row + 1, 1).Value
        End If
    Next cell

    ' Display summary
    If bestCapacity <> "Not Found" Then
        MsgBox "The best capacity selected by the model is: " & bestCapacity & vbCrLf & _
               "This capacity yields the highest Mean NPV of: " & Format(maxMetric, "0.00"), _
               vbInformation, "Summary of Results"
    Else
        MsgBox "No valid results found from the simulation.", vbExclamation, "Summary Error"
    End If
End Sub

