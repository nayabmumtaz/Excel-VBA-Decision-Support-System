Attribute VB_Name = "Module8"
Option Explicit

Sub Export()
    Dim userResponse As VbMsgBoxResult
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Outputs")
    
    'Yes/No options
    userResponse = MsgBox("Would you like to save the report as a PDF?", vbYesNo + vbQuestion, "Export Options")
    
    If userResponse = vbYes Then
        ' Adjust page setup before exporting
        With ws.PageSetup
            .Orientation = xlLandscape
            .FitToPagesWide = 1 '
            .FitToPagesTall = False
            .Zoom = False
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5) '
        End With
        
        ' Export the worksheet as a PDF
        ws.ExportAsFixedFormat xlTypePDF
        MsgBox "Report successfully saved as PDF!", vbInformation, "Export Successful"
    
    ElseIf userResponse = vbNo Then
        ' If user clicks No, display a cancellation message and exit
        MsgBox "Export cancelled. No file was saved.", vbInformation, "Export Cancelled"
        Exit Sub
    End If
End Sub


