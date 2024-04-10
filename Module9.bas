Attribute VB_Name = "Module9"
'MODULE_9'
'Check correct weeks in EDI sheet'
Sub CheckWeekEDI(ValMessage As Boolean)
    'COMPROBACI�N DE SEMANAS EN EDI'
    'Verificaci�n de correcta distribuci�n de semanas en el EDI'
    
    Dim EDISheet As Worksheet
    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
    
    Dim LastColEDI As Long
    LastColEDI = EDISheet.Cells(2, Columns.Count).End(xlToLeft).Column
    
    Dim week As Integer
    Dim ActualDate As Date
    Dim ActualWeekString As String
    'Booleanos para comprobar que no hay errores. Todos se inicializan en False'
    Dim SameWeek As Boolean
    Dim RepWeek As Boolean
    Dim WholeWeeks As Boolean
    SameWeek = False
    RepWeek = False
    WholeWeeks = False
    
    'Bucle para recorrer todas las fechas del EDI y comprobar que coinciden con las semanas'
    For i = 2 To LastColEDI
        ActualDate = StringToDate(EDISheet.Cells(2, i).Value)
        week = DatePart("ww", ActualDate, vbMonday)
        ActualWeekString = "S" & week '- 1
        If ActualWeekString <> EDISheet.Cells(1, i) Then
            Dim answer As Integer
            'MsgBox "ERROR. La semana " & Week & " no coincide con la fecha del EDI"
            answer = MsgBox("ERROR. La semana " & week & " correspondiente a la fecha " & ActualDate & " no coincide con la fecha del EDI. �Desea continuar con la ejecuci�n del procedimiento?", vbQuestion + vbYesNo, "ERROR")
            If answer = vbYes Then
            Else
                Exit For
            End If
        Else
        SameWeek = True
        End If
    Next i
    
    'Comprobaci�n semanas no repetidas'
    For i = 2 To LastColEDI
        If EDISheet.Cells(1, i).Value = EDISheet.Cells(1, i + 1).Value Then
            MsgBox "La semana " & EDISheet.Cells(1, i).Value & " correspondiente al d�a " & EDISheet.Cells(2, i).Value & " se encuentra duplicada"
        Else
        RepWeek = True
        End If
    Next i
    
    'Comprobaci�n todas las semanas' 'Si falta alguna semana del a�o salta error'
    For i = 2 To LastColEDI - 1
        ActualDate = StringToDate(EDISheet.Cells(2, i).Value)
        If ActualDate + 7 <> StringToDate(EDISheet.Cells(2, i + 1).Value) Then
            MsgBox "ERROR en la fecha " & ActualDate
        Else
        WholeWeeks = True
        End If
    Next i
    
    'Mensaje por pantalla si no se encuentrar errores'
    'ValMessage indica si se mostrar� por pantalla un mensaje de verificaci�n de estado correcto.
    'Usado dependiendo de la opci�n elegida en el men� principal'
    If ValMessage = True Then
        If SameWeek = True Then
            If RepWeek = True Then
                If WholeWeeks = True Then
                MsgBox "No se encontraron errores"
                Else
                End If
            Else
            End If
        Else
        End If
    Else
    End If
End Sub

Sub CheckStringEDI()
    'VERIFICA LAS REFERENCIAS DEL EDI PARA ALMACENARLAS COMO CADENAS DE TEXTO.
    'SI NO LO SON, A�ADE UN AP�STROFE AL COMIENZO
    Dim EDISheet As Worksheet
    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
    
    Dim lastRow As Integer
    lastRow = EDISheet.Cells(EDISheet.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
    ' Verifica si el valor de la celda es un n�mero
    If IsNumeric(EDISheet.Cells(i, "A").Value) Then
        ' Agrega un ap�strofo al principio del n�mero para convertirlo en una cadena de caracteres
        EDISheet.Cells(i, "A").Value = "'" & EDISheet.Cells(i, "A").Value
    End If
    Next i
End Sub

