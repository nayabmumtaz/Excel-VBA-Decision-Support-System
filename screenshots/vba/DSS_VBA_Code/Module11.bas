Attribute VB_Name = "Module11"
Sub NPVVisuals()
    Dim ws As Worksheet
    Dim chart1 As ChartObject
    Dim chart2 As ChartObject
    Dim button As button
    
    Set ws = ThisWorkbook.Sheets("Outputs")
    Set chart1 = ws.ChartObjects("Chart 3")
    Set chart2 = ws.ChartObjects("Chart 4")
    Set button = ws.Buttons("Button 4")
    
    If chart1.Visible = True Then
        chart1.Visible = False
        chart2.Visible = False
        button.Caption = "Show NPV Visuals"
    Else
        chart1.Visible = True
        chart2.Visible = True
        button.Caption = "Hide NPV Visuals"
    End If
End Sub

