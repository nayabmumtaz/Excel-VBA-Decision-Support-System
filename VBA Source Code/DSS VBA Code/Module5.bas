Attribute VB_Name = "Module5"
Sub ExportAndSaveWorkbook()
    Dim filePath As Variant
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim emailSubject As String
    Dim emailBody As String
    Dim recipient As String
    
    ' Display the Save As dialog and set the file path

 On Error Resume Next ' To handle any error in the next line gracefully
    filePath = Application.GetSaveAsFilename( _
    InitialFileName:="VBA Capstone F24 Report_" & Format(Now, "yyyy-mm-dd"), _
    FileFilter:="Excel Files (*.xlsm), *.xlsm", _
    Title:="Save Your Report")
    
    On Error GoTo 0 ' Turn back to normal error handling
        
    ' Check if the user canceled the Save As dialog (i.e., filePath is False)
    If filePath = False Then
        MsgBox "You did not select a file to save. Operation canceled.", vbExclamation, "Save Canceled"
        Exit Sub
    End If
    
    ' Save the file in the chosen location with the correct format
    Application.DisplayAlerts = False ' Turn off alerts to prevent confirmation dialogs during save
    ThisWorkbook.SaveAs filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True ' Turn alerts back on
    
    ' Notify user the file has been saved successfully
    MsgBox "Your workbook has been saved successfully to: " & vbCrLf & filePath, vbInformation, "File Saved"
    
    ' Prepare email details
    emailSubject = "Excel Report from VBA Workbook"
    emailBody = "Please find the attached report. Kindly review the results."
    recipient = "recipient@example.com" ' Replace with actual recipient email address
    
    ' Create Outlook application object
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    
    ' Create and send email with attachment
    With outlookMail
        .Subject = emailSubject
        .Body = emailBody
        .To = recipient
        .Attachments.Add filePath  ' Attach the saved workbook
        .Send ' Or use .Display if you prefer to manually send it
    End With
    
    ' Clear the objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    ' Notify user the email was sent
    MsgBox "The email has been sent successfully with the attached file.", vbInformation, "Email Sent"
End Sub


