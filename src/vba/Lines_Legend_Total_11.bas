Attribute VB_Name = "Lines_Legend_Total_11"


  Sub Lines_Legend_New_Total_11()

    On Error Resume Next ' skyddar mot enstaka missar utan att bryta hela
    Application.Run "Lines_Legend_Delete_Tables"
  Application.Run "Lines_Legend_New_1"
   Application.Run "Lines_Legend_New_2"
     Application.Run "Lines_Legend_New_3"
        Application.Run "Lines_Legend_New_4"
           Application.Run "Lines_Legend_New_5"
              Application.Run "Lines_Legend_New_6"
                 Application.Run "Lines_Legend_New_7"
                    Application.Run "Lines_Legend_New_8"
                    
                      Application.Run "Lines_Legend_New_B1"
   Application.Run "Lines_Legend_New_B2"
     Application.Run "Lines_Legend_New_B3"
        Application.Run "Lines_Legend_New_B4"
           Application.Run "Lines_Legend_New_B5"
              Application.Run "Lines_Legend_New_B6"
                 Application.Run "Lines_Legend_New_B7"
                    Application.Run "Lines_Legend_New_B8"
     Application.Run "Lines_Legend_Remove_Last_Row"
    On Error GoTo 0
    Debug.Print "? Lines_Legend_New_Total k�rd klart."
End Sub

Function GetVisibleSeriesIndexes(chartObject As chart) As Collection
    Dim i As Integer
    Dim visibleIndexes As New Collection

    For i = 1 To chartObject.SeriesCollection.count
        If chartObject.SeriesCollection(i).Format.line.visible = msoTrue Then
            visibleIndexes.Add i
        End If
    Next i

    Set GetVisibleSeriesIndexes = visibleIndexes
End Function

