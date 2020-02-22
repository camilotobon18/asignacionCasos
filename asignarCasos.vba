Sub AsignacionCasos()
Dim largoArray As Integer
Dim anchoArray As Integer
Dim lista() As String
Dim listaAsignacion() As String
Const ticket As Integer = 2
Const fechaApertura As Integer = 4
Const estado As Integer = 8
'Set objArray = CreateObject("System.Array")

'variables para almacenar los responsables
Dim r1 As String
Dim r2 As String
Dim r3 As String
Dim r4 As String

'variables para almacenar los porcentajes para cada responsable
Dim p1 As Double
Dim p2 As Double
Dim p3 As Double
Dim p4 As Double

r1 = Sheets("Maestro").Range("A2").Value
r2 = Sheets("Maestro").Range("A3").Value
r3 = Sheets("Maestro").Range("A4").Value
r4 = Sheets("Maestro").Range("A5").Value

p1 = Sheets("Maestro").Range("B2").Value
p2 = Sheets("Maestro").Range("B3").Value
p3 = Sheets("Maestro").Range("B4").Value
p4 = Sheets("Maestro").Range("B5").Value


Sheets("ListaIncidentes").Select
largoArray = Application.CountA(Sheets("ListaIncidentes").Columns("B")) - 1
anchoArray = Application.CountA(Sheets("ListaIncidentes").Rows("1"))
largoAsignacion = Application.CountA(Sheets("Asignacion").Columns("A"))
anchoAsignacion = Application.CountA(Sheets("Asignacion").Rows("2"))

ReDim lista(anchoArray, largoArray)
ReDim listaAsignacion(anchoAsignacion, largoAsignacion)

For fila = 1 To largoArray
    For columna = 1 To anchoArray
        lista(columna, fila) = ActiveSheet.Cells(fila + 1, columna).Value
    Next
Next

Sheets("Asignacion").Select
For filaAsignacion = 1 To largoAsignacion
    listaAsignacion(1, filaAsignacion) = ActiveSheet.Cells(filaAsignacion + 1, 1).Value
    listaAsignacion(2, filaAsignacion) = ActiveSheet.Cells(filaAsignacion + 1, 2).Value
    listaAsignacion(3, filaAsignacion) = ActiveSheet.Cells(filaAsignacion + 1, 3).Value
    listaAsignacion(4, filaAsignacion) = ActiveSheet.Cells(filaAsignacion + 1, 4).Value
    listaAsignacion(5, filaAsignacion) = ActiveSheet.Cells(filaAsignacion + 1, 5).Value
Next filaAsignacion
'objArray.Sort lista



If BuscarHoja("Asignacion") = True Then
    Application.DisplayAlerts = False
    Sheets("Asignacion").Delete
    Application.DisplayAlerts = True
    Worksheets.Add.Name = "Asignacion"
Else
    Worksheets.Add.Name = "Asignacion"
End If

Worksheets("Asignacion").Range("A1").Value = "Ticket"
Worksheets("Asignacion").Range("B1").Value = "Fecha"
Worksheets("Asignacion").Range("C1").Value = "Estado"
Worksheets("Asignacion").Range("D1").Value = "Tipo"
Worksheets("Asignacion").Range("E1").Value = "Responsable"

Dim contador As Integer

For fila = 1 To largoArray
    'For columna = 1 To 12
        'Worksheets("Asignacion").Cells(fila + 1, columna).Value = lista(columna, fila)
    'Next columna
    Worksheets("Asignacion").Cells(fila + 1, 1).Value = lista(ticket, fila)
    Worksheets("Asignacion").Cells(fila + 1, 2).Value = lista(fechaApertura, fila)
    Worksheets("Asignacion").Cells(fila + 1, 3).Value = lista(estado, fila)
    'Worksheets("Asignacion").Cells(fila + 1, 4).Value = Application.WorksheetFunction.VLookup(lista(ticket, fila), listaAsignacion, 2, False)
    
    contador = 0
    For filaAsignacion = 1 To largoAsignacion
        
        If lista(ticket, fila) = listaAsignacion(1, filaAsignacion) Then
            Worksheets("Asignacion").Cells(fila + 1, 4).Value = "Antiguo"
            Worksheets("Asignacion").Cells(fila + 1, 5).Value = listaAsignacion(5, filaAsignacion)
            contador = contador + 1
        ElseIf lista(ticket, fila) <> listaAsignacion(1, filaAsignacion) And contador = 0 Then
            Worksheets("Asignacion").Cells(fila + 1, 4).Value = "Nuevo"
        End If
    Next filaAsignacion
    
Next fila

Dim largoDataFiltrada As Integer
largoDataFiltrada = Application.CountA(Sheets("Asignacion").Columns("A")) - 1

For fila = 1 To largoDataFiltrada
    'La funcion CStr convierte un numero a texto
    Worksheets("Asignacion").Cells(fila + 1, 2).Value = _
    Application.VLookup(CStr(Worksheets("Asignacion").Cells(fila + 1, 1).Value), Worksheets("ListaIncidentes").Range("B:D"), 3, 0)
Next fila

'Darle formato a una celda
With Worksheets("Asignacion").Range("B:B")
    .NumberFormat = "dd/mm/yyyy h:mm"
End With

'instrucciones para seleccionar el rango dinamico de la hoja Asignacion
Dim rangoOrdenar As Range
Dim nFilas As Integer
nFilas = Worksheets("Asignacion").Cells(1, 1).CurrentRegion.Rows.Count

Set rangoOrdenar = Range(Cells(1, 1), Cells(nFilas, 5))
rangoOrdenar.Select

'instrucciones para ordenar el rango de la lista
ActiveWorkbook.Worksheets("Asignacion").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Asignacion").Sort.SortFields.Add2 Key:=Range( _
        "A2:A63"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Asignacion").Sort
        .SetRange Range("A1:E63")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
