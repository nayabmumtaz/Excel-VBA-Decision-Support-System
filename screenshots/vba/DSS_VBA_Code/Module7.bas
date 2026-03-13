Attribute VB_Name = "Module7"
Sub ViewGuide()
    Dim message As String
    
    'Message for the user guide
    message = "Welcome to the Model Input Form" & vbCrLf & vbCrLf & _
              "This is where you customize the simulation parameters for the Prizdol production to choose the best capacity . " & _
              "Each section represents a critical factor in determining profitability." & vbCrLf & vbCrLf & _
              "Here’s how to use the input form:" & vbCrLf & _
              "1. Review the pre-filled values, which serve as default inputs based on the current model sheet." & vbCrLf & _
              "2. Update any parameters as needed to reflect your specific expectations or scenarios." & vbCrLf & _
              "3. Hover over or click the information icons next to each parameter for detailed guidance." & vbCrLf & _
              "4. Once you’ve entered all the inputs, click the Run Simulation button at the bottom to proceed." & vbCrLf & vbCrLf & _
              "Feel free to experiment with different values to explore how changes impact the results!"
    
    ' Display the message in a message box
    MsgBox message, vbInformation, "Guide For Using the Input Form"
End Sub


