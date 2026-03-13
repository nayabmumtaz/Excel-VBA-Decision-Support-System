Attribute VB_Name = "Module12"
Sub AllMetricVisuals()
    Dim ws As Worksheet
    Dim chart1 As ChartObject
    Dim chart2 As ChartObject
    Dim chart3 As ChartObject
    Dim button As button
    
   
    Set ws = ThisWorkbook.Sheets("Outputs")
    Set chart1 = ws.ChartObjects("Chart 6")
    Set chart2 = ws.ChartObjects("Chart 8")
    Set chart3 = ws.ChartObjects("Chart 11")
    Set button = ws.Buttons("Button 7")
    
    If chart1.Visible = True Then
        chart1.Visible = False
        chart2.Visible = False
        chart3.Visible = False
        button.Caption = "Show All Metric Visuals"
    Else
        chart1.Visible = True
        chart2.Visible = True
        chart3.Visible = True
        button.Caption = "Hide All Metric Visuals"
    End If
End Sub

