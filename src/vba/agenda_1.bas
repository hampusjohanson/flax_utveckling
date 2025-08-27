Attribute VB_Name = "agenda_1"
' Declare public variables to store the lists
Public agendaItems As Collection
Public chapterItems As Collection

Sub Agenda_1()
    Dim slide As slide
    Dim shape As shape
    Dim agendaKeywords As Variant
    Dim chapterKeywords As Variant
    Dim key As Variant
    Dim currentAgendaItem As String
    Dim agendaItemList As String
    Dim chapterItemList As String
    Dim agendaItemExists As Boolean
    Dim agendaItem As Variant
    
    ' Initialize collections to store agenda items and chapter items
    Set agendaItems = New Collection
    Set chapterItems = New Collection
    
   agendaKeywords = Array("Agenda", "Table of Contents", "Overview", "Topics", "Sections", _
                       "Dagordning", "InnehŒll", "…versikt", "Punkter", "Kapitel", _
                       "Content of report", "Today's agenda", "Content")

    ' Loop through slides to find agenda items
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            ' Check if shape contains text
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    ' Check for agenda-related keywords in titles (e.g., Agenda, Topics)
                    For Each key In agendaKeywords
                        If InStr(1, shape.TextFrame.textRange.text, key, vbTextCompare) > 0 Then
                            currentAgendaItem = Trim(shape.TextFrame.textRange.text) ' Store the agenda section header
                            ' Now look for the actual agenda items (e.g., bullet points under the header)
                            Dim nextShape As shape
                            For Each nextShape In slide.Shapes
                                If nextShape.HasTextFrame Then
                                    If nextShape.TextFrame.HasText Then
                                        ' Avoid adding the header again
                                        If InStr(1, nextShape.TextFrame.textRange.text, currentAgendaItem, vbTextCompare) = 0 Then
                                            ' Filter out non-relevant items like "2" or empty items
                                            If Len(Trim(nextShape.TextFrame.textRange.text)) > 0 And Not IsNumeric(Trim(nextShape.TextFrame.textRange.text)) Then
                                                ' Check if the item already exists in agendaItems
                                                agendaItemExists = False
                                                For Each agendaItem In agendaItems
                                                    If agendaItem = nextShape.TextFrame.textRange.text Then
                                                        agendaItemExists = True
                                                        Exit For
                                                    End If
                                                Next agendaItem
                                                
                                                ' If not exists, add it to the collection
                                                If Not agendaItemExists Then
                                                    agendaItems.Add nextShape.TextFrame.textRange.text
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next nextShape
                        End If
                    Next key
                End If
            End If
        Next shape
    Next slide
    
    ' Save agenda items as a list string
    For Each Item In agendaItems
        agendaItemList = agendaItemList & Item & vbCrLf
    Next Item
    
    ' Print agenda items in Immediate window
    Debug.Print "Agenda Items:"
    Debug.Print agendaItemList ' Output agenda list in Immediate window
    
    ' Loop through slides to find chapter items based on layout names
    Debug.Print "Chapter Slides:"
    For Each slide In ActivePresentation.Slides
        ' Look for specific layouts that might indicate chapter slides
        If InStr(1, LCase(slide.CustomLayout.Name), "chapter") > 0 Or _
           InStr(1, LCase(slide.CustomLayout.Name), "title slide") > 0 Then
            For Each shape In slide.Shapes
                If shape.HasTextFrame Then
                    If shape.TextFrame.HasText Then
                        chapterItems.Add shape.TextFrame.textRange.text
                    End If
                End If
            Next shape
        End If
    Next slide
    
    ' Save chapter items as a list string
    For Each Item In chapterItems
        chapterItemList = chapterItemList & Item & vbCrLf
    Next Item
    
    ' Print chapter items in Immediate window
    Debug.Print chapterItemList ' Output chapter list in Immediate window
End Sub



