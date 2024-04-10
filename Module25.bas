Attribute VB_Name = "Module25"
'MODULE_25'
'COMPROBACIÓN REFERENCIAS COMO CADENAS DE TEXTO EN PESTAÑAS PROCESS Y REFERENCES'
Sub ChkRefStr()
    'Comprobación de Strings en la pestaña REFERENCES'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("References"))
    
    Dim lastRow As Integer

    'Encuentra la última fila con datos en la columna B
    lastRow = ws.Cells(ws.Rows.Count, NumColReference("REFERENCE")).End(xlUp).Row
    
    'Recorre cada celda en la columna A desde la fila 1 hasta la última fila con datos
    For i = 1 To lastRow
        ' Verifica si el valor de la celda es un número y si es una celda en blanco
        If IsNumeric(ws.Cells(i, "A").Value) Then
            If ws.Cells(i, "A").Value = "" Then
                'Comprobación de celda en blanco. (Reconoce la celda en blanco como Zero)
            Else
                ' Agrega un apóstrofo al principio del número para convertirlo en una cadena de caracteres
                ws.Cells(i, "A").Value = "'" & ws.Cells(i, "A").Value
            End If
        End If
    Next i

    'Recorre cada celda en la columna B desde la fila 1 hasta la última fila con datos
    For i = 1 To lastRow
        ' Verifica si el valor de la celda es un número y si es una celda en blanco
        If IsNumeric(ws.Cells(i, NumColReference("REFERENCE")).Value) Then
            If ws.Cells(i, NumColReference("REFERENCE")).Value = "" Then
                'Comprobación de celda en blanco. (Reconoce la celda en blanco como Zero)
            Else
                ' Agrega un apóstrofo al principio del número para convertirlo en una cadena de caracteres
                ws.Cells(i, NumColReference("REFERENCE")).Value = "'" & ws.Cells(i, NumColReference("REFERENCE")).Value
            End If
        End If
    Next i

    'Recorre cada celda en la columna F desde la fila 1 hasta la última fila con datos
    For i = 1 To lastRow
        ' Verifica si el valor de la celda es un número y si es una celda en blanco
        If IsNumeric(ws.Cells(i, NumColReference("FINALREF")).Value) Then
            If ws.Cells(i, NumColReference("FINALREF")).Value = "" Then
                'Comprobación de celda en blanco. (Reconoce la celda en blanco como Zero)
            Else
                ' Agrega un apóstrofo al principio del número para convertirlo en una cadena de caracteres
                ws.Cells(i, NumColReference("FINALREF")).Value = "'" & ws.Cells(i, NumColReference("FINALREF")).Value
            End If
        End If
    Next i

    'Recorre cada celda en la columna G desde la fila 1 hasta la última fila con datos
    For i = 1 To lastRow
        ' Verifica si el valor de la celda es un número y si es una celda en blanco
        If IsNumeric(ws.Cells(i, NumColReference("NEXT_REFERENCE")).Value) Then
            If ws.Cells(i, NumColReference("NEXT_REFERENCE")).Value = "" Then
                'Comprobación de celda en blanco. (Reconoce la celda en blanco como Zero)
            Else
                ' Agrega un apóstrofo al principio del número para convertirlo en una cadena de caracteres
                ws.Cells(i, NumColReference("NEXT_REFERENCE")).Value = "'" & ws.Cells(i, NumColReference("NEXT_REFERENCE")).Value
            End If
        End If
    Next i
End Sub

Sub ChkProcStr()
    'Comprobación de Strings en la pestaña PROCESS'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("Process"))
    
    Dim lastRow As Integer

    'Encuentra la última fila con datos en la columna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Recorre cada celda en la columna A desde la fila 1 hasta la última fila con datos
    For i = 1 To lastRow
        ' Verifica si el valor de la celda es un número
        If IsNumeric(ws.Cells(i, "A").Value) Then
            ' Agrega un apóstrofo al principio del número para convertirlo en una cadena de caracteres
            ws.Cells(i, "A").Value = "'" & ws.Cells(i, "A").Value
        End If
    Next i
End Sub
