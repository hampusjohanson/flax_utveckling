Attribute VB_Name = "Module10"
Sub CreateLetieShape()
    Dim slide As slide
    Dim shape As shape
    
    ' Reference the active slide
    Set slide = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)
    
    ' Create the rectangle shape with specified dimensions and position
    Set shape = slide.Shapes.AddShape(msoShapeRectangle, 3.54 * 28.35, 5.14 * 28.35, 1.73 * 28.35, 0.94 * 28.35)
    
    ' Name the shape
    shape.Name = "Leftie"
End Sub

