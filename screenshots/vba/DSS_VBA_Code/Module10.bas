Attribute VB_Name = "Module10"
Sub ZScoreTable()
    
    Dim shp As Shape
    On Error Resume Next
    Set shp = ActiveSheet.Shapes("Picture 2")

    If shp.Visible = msoTrue Then
        shp.Visible = msoFalse
    Else
        shp.Visible = msoTrue
    End If

End Sub


