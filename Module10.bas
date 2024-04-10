Attribute VB_Name = "Module10"
'MODULE_10'
'ACTUALIZA LAS NECESIDADES DE PRODUCCIÓN'
Sub WeekProdNeed()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))

    'Elección para actualizar todo o sólo la semana deseada'
    Dim answer As Integer
    answer = MsgBox("¿Desea actualizar todas las semanas?", vbQuestion + vbYesNo, "")
    If answer = vbYes Then
        For i = StartWeek() To CurrentWeekNumber() + FutureWeeks()
            ProdNeed (i)
        Next i
    Else
        Dim WeekSearch As Integer
        WeekSearch = Application.InputBox(prompt:="Indique la semana:", Type:=2, Title:="BÚSQUEDA DE SEMANA")
        ProdNeed (WeekSearch)
    End If
End Sub

Sub WeekProdNeedUpdateAll()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    For i = StartWeek() To CurrentWeekNumber() + FutureWeeks()
        ProdNeed (i)
    Next i
End Sub

Sub WeldingWeekCurrent()
    'Almacena la referencia y semana para llamar a la función de añadir semana
    Dim ws As Worksheet
    Dim Reference As String
    Dim cell As Range
    Dim foundCell As Range
    Dim answer As VbMsgBoxResult
    
    ' Definir la hoja de Excel donde se encuentra la información
    Set ws = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Do
        ' Pedir al usuario que ingrese la parte de la referencia a buscar
        Reference = InputBox("Ingrese la parte de la referencia a buscar:")
        
        ' Inicializar la variable de la primera coincidencia como nula
        Set foundCell = Nothing
        
        ' Recorrer la columna D y buscar la referencia
        For Each cell In ws.Range("D7:D" & ws.Cells(ws.Rows.Count, "D").End(xlUp).Row) 'Step 4
            If InStr(1, cell.Value, Reference, vbTextCompare) > 0 Then
                ' Si se encuentra una coincidencia, almacenarla y salir del bucle
                Set foundCell = cell
                Exit For
            End If
        Next cell
        
        ' Comprobar si se encontró una coincidencia
        If Not foundCell Is Nothing Then
            answer = MsgBox("¿Es esta la referencia deseada?: " & cell.Value & " (Sí/No)", vbQuestion + vbYesNo)
            If answer = vbYes Then
                MsgBox "La referencia se encontró en la fila " & foundCell.Row
                Exit Do ' Salir del bucle si se encuentra la referencia deseada
            End If
        Else
            answer = MsgBox("No se encontraron coincidencias para la referencia proporcionada. ¿Desea intentar de nuevo? (Sí/No)", vbQuestion + vbYesNo)
            If answer = vbNo Then Exit Do ' Salir del bucle si el usuario no quiere intentar de nuevo
        End If
    Loop
End Sub
