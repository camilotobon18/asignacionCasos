Function BuscarHoja(nombreHoja As String) As Boolean
 
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            BuscarHoja = True
            Exit Function
        End If
    Next
     
    BuscarHoja = False
 
End Function
