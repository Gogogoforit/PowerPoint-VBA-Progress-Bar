Sub AutoProgressBar()
    On Error Resume Next ' Suppress errors to avoid runtime interruptions

    With ActivePresentation
        ' Loop through each slide
        For X = 1 To .Slides.Count
            ' Remove existing progress bar
            .Slides(X).Shapes("PB").Delete
            
            ' Create a new progress bar
            Dim s As Shape
            Set s = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
                0, .PageSetup.SlideHeight - 12, _
                X * .PageSetup.SlideWidth / .Slides.Count, 12)
            
            ' Set the fill color of the progress bar
            s.Fill.ForeColor.RGB = RGB(127, 0, 0) ' Red color
            s.Name = "PB" ' Name the shape as "PB"
        Next X
    End With
End Sub
