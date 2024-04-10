Attribute VB_Name = "Module20"
'MODULE_20'
'ACTUALIZACI�N SEMANAS BOX'
Sub BoxFirstWeek()
    'A�ADE LA PRIMERA SEMANA DEL A�O A LA PESTA�A BOX'
    Call AddBoxWeekHeaders(1, 5)
    Call BoxWeekBody(1, 5)
    Call BoxWeekFormat(1)

End Sub

Sub BoxWeekUpdate()
    'Actualiza hasta la semana actual + futura'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("Box"))
    
    Dim CurrentWeek As Integer
    Dim BoxWeekLastCol As Integer
    
    CurrentWeek = CurrentWeekNumber()
    BoxWeekLastCol = BoxSheet.Cells(OffsetFilaCabecera() - 2, Columns.Count).End(xlToLeft).Column
    
    Dim CurrentPlusFuture As Integer
    CurrentPlusFuture = CurrentWeek + FutureWeeks()
    'Comprobaci�n semanas desactualizadas'
    If BoxSheet.Cells(OffsetFilaCabecera() - 2, BoxWeekLastCol).Value = "Week " & CurrentPlusFuture Then
        MsgBox "Las semanas se encuentran actualizadas"
    Else
        If NumExtract(BoxSheet.Cells(OffsetFilaCabecera() - 2, BoxWeekLastCol)) < (CurrentWeek + FutureWeeks()) Then
            MsgBox "Semanas desactualizadas. Se van a actualizar hasta la semana: " & CurrentPlusFuture
            Dim NextWeekCol As Integer
            Dim CounterWeek As Integer
            Dim var As Integer 'Variable vac�a para funci�n'
            CounterWeek = NumExtract(BoxSheet.Cells(OffsetFilaCabecera() - 2, BoxWeekLastCol))
            For i = CounterWeek + 1 To CurrentPlusFuture
                CounterWeek = i
                NextWeekCol = BoxWeekLastCol + BoxColDistance()
                Call AddBoxWeekHeaders(CounterWeek, NextWeekCol)
                Call BoxWeekBody(CounterWeek, NextWeekCol - 1)
                Call BoxWeekFormat(CounterWeek)
                ImportWeekEDI (CounterWeek)
                BoxWeekLastCol = BoxWeekLastCol + BoxColDistance() 'Avanzamos a la siguiente celda donde se colocar� la semana'
            Next i
        Else
        End If
    End If
End Sub

Sub BoxWeeksBuilder()
    'Construye todas las semanas para la creaci�n de la pesta�a completa
    BoxFirstWeek 'Primera semana'
    
    'Actualiza hasta la semana actual + futura'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("Box"))
    
    Dim CurrentWeek As Integer
    Dim BoxWeekLastCol As Integer
    
    CurrentWeek = CurrentWeekNumber()
    BoxWeekLastCol = BoxSheet.Cells(OffsetFilaCabecera() - 2, Columns.Count).End(xlToLeft).Column
    
    Dim CurrentPlusFuture As Integer
    CurrentPlusFuture = CurrentWeek + FutureWeeks()
    'Comprobaci�n semanas desactualizadas'
    If BoxSheet.Cells(OffsetFilaCabecera() - 2, BoxWeekLastCol).Value = "Week " & CurrentPlusFuture Then
    Else
        If NumExtract(BoxSheet.Cells(OffsetFilaCabecera() - 2, BoxWeekLastCol)) < (CurrentWeek + FutureWeeks()) Then
            Dim NextWeekCol As Integer
            Dim CounterWeek As Integer
            Dim var As Integer 'Variable vac�a para funci�n'
            CounterWeek = NumExtract(BoxSheet.Cells(OffsetFilaCabecera() - 2, BoxWeekLastCol))
            For i = CounterWeek + 1 To CurrentPlusFuture
                CounterWeek = i
                NextWeekCol = BoxWeekLastCol + BoxColDistance()
                Call AddBoxWeekHeaders(CounterWeek, NextWeekCol)
                Call BoxWeekBody(CounterWeek, NextWeekCol - 1)
                Call BoxWeekFormat(CounterWeek)
                ImportWeekEDI (CounterWeek)
                BoxWeekLastCol = BoxWeekLastCol + BoxColDistance 'Avanzamos a la siguiente celda donde se colocar� la semana'
            Next i
        Else
        End If
    End If
    
End Sub
