Attribute VB_Name = "Module6"
Sub LearnMore()
    Dim message As String
    
    ' Define the updated message with your new text
    message = "This DSS allows you to quickly simulate different production capacity levels and evaluate their impact on profitability. " & _
              "By running multiple simulations, the DSS calculates the expected net present value (NPV) for each capacity, " & _
              "providing you with a range of outcomes (mean, minimum, and maximum NPVs). This enables you to make data-driven decisions " & _
              "while also assessing the associated risks. The system is highly efficient, offering rapid calculations that save you time " & _
              "and help mitigate uncertainties around demand fluctuations and pricing changes. The model optimizes key resources—such as " & _
              "capital for plant construction, operating costs, and pricing strategy—by aligning production capacity with projected demand. " & _
              "This ensures you select the ideal capacity to maximize profitability, improve operational efficiency, and effectively manage " & _
              "risks over the next 10 years."
    
    ' Display the message in a message box
    MsgBox message, vbInformation, "Learn More About the Model"
End Sub


