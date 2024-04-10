Attribute VB_Name = "Module30"
'MODULE_30'
'ACTUALIZACIÓN SEMANAS BENDING'
Sub BendingFirstWeek()
    Call AddBendingWeekHeaders(1, FirstBendingData())
    Call BendingWeekBody(1, FirstBendingData())
    Call BendingWeekFormat(1)
End Sub

Sub BendingWeekUpdate()
    'Actualiza hasta la semana actual + futura'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    
    Dim CurrentWeek As Integer
    Dim BendingWeekLastCol As Integer
    
    CurrentWeek = CurrentWeekNumber()
    BendingWeekLastCol = BendingSheet.Cells(OffsetFilaCabecera() - 2, Columns.Count).End(xlToLeft).Column
    
    Dim CurrentPlusFuture As Integer
    CurrentPlusFuture = CurrentWeek + FutureWeeks()
    'Comprobación semanas desactualizadas'
    If BendingSheet.Cells(OffsetFilaCabecera() - 2, BendingWeekLastCol).Value = "Week " & CurrentPlusFuture Then
        MsgBox "Las semanas se encuentran actualizadas"
    Else
        If NumExtract(BendingSheet.Cells(OffsetFilaCabecera() - 2, BendingWeekLastCol)) < (CurrentWeek + FutureWeeks()) Then
            MsgBox "Semanas desactualizadas. Se van a actualizar hasta la semana: " & CurrentPlusFuture
            Dim NextWeekCol As Integer
            Dim CounterWeek As Integer
            Dim var As Integer 'Variable vacía para función'
            CounterWeek = NumExtract(BendingSheet.Cells(OffsetFilaCabecera() - 2, BendingWeekLastCol))
            For i = CounterWeek + 1 To CurrentPlusFuture
                CounterWeek = i
                NextWeekCol = BendingWeekLastCol + BoxColDistance()
                Call AddBendingWeekHeaders(CounterWeek, NextWeekCol)
                Call BendingWeekBody(CounterWeek, NextWeekCol)
                Call BendingWeekFormat(CounterWeek)
                ImportWeekEDI (CounterWeek)
                BendingWeekLastCol = BendingWeekLastCol + BendingColDistance() 'Avanzamos a la siguiente celda donde se colocará la semana'
            Next i
        Else
        End If
    End If
End Sub

Sub BendingWeeksBuilder()
    'Construye todas las semanas para la creación de la pestaña completa
    BendingFirstWeek
    
    'Actualiza hasta la semana actual + futura'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    
    Dim CurrentWeek As Integer
    Dim BendingWeekLastCol As Integer
    
    CurrentWeek = CurrentWeekNumber()
    BendingWeekLastCol = BendingSheet.Cells(OffsetFilaCabecera() - 2, Columns.Count).End(xlToLeft).Column
    
    Dim CurrentPlusFuture As Integer
    CurrentPlusFuture = CurrentWeek + FutureWeeks()
    'Comprobación semanas desactualizadas'
    If BendingSheet.Cells(OffsetFilaCabecera() - 2, BendingWeekLastCol).Value = "Week " & CurrentPlusFuture Then
        MsgBox "Las semanas se encuentran actualizadas"
    Else
        If NumExtract(BendingSheet.Cells(OffsetFilaCabecera() - 2, BendingWeekLastCol)) < (CurrentWeek + FutureWeeks()) Then
            MsgBox "Semanas desactulizadas. Se van a actualizar hasta la semana: " & CurrentPlusFuture
            Dim NextWeekCol As Integer
            Dim CounterWeek As Integer
            Dim var As Integer 'Variable vacía para función'
            CounterWeek = NumExtract(BendingSheet.Cells(OffsetFilaCabecera() - 2, BendingWeekLastCol))
            For i = CounterWeek + 1 To CurrentPlusFuture
                CounterWeek = i
                NextWeekCol = BendingWeekLastCol + BoxColDistance()
                Call AddBendingWeekHeaders(CounterWeek, NextWeekCol)
                Call BendingWeekBody(CounterWeek, NextWeekCol)
                Call BendingWeekFormat(CounterWeek)
                ImportWeekEDI (CounterWeek)
                BendingWeekLastCol = BendingWeekLastCol + BendingColDistance() 'Avanzamos a la siguiente celda donde se colocará la semana'
            Next i
        Else
        End If
    End If
End Sub
