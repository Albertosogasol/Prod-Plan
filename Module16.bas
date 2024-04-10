Attribute VB_Name = "Module16"
'MODULE_16'
'ACTUALIZACIÓN SEMANAS WELDING'

Sub ActualizarSemanas()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    
    Dim CurrentWeek As Integer
    Dim WeldingWeekLastCol As Integer
    
    CurrentWeek = CurrentWeekNumber()
    WeldingWeekLastCol = WeldingSheet.Cells(OffsetFilaCabecera() - 2, Columns.Count).End(xlToLeft).Column
    
    Dim CurrentPlusFuture As Integer
    CurrentPlusFuture = CurrentWeek + FutureWeeks()
    'Comprobación semanas desactualizadas'
    If WeldingSheet.Cells(OffsetFilaCabecera() - 2, WeldingWeekLastCol).Value = "Week " & CurrentPlusFuture Then
        MsgBox "Las semanas se encuentran actualizadas"
    Else
        If NumExtract(WeldingSheet.Cells(OffsetFilaCabecera() - 2, WeldingWeekLastCol)) < (CurrentWeek + FutureWeeks()) Then
            MsgBox "Semanas desactulizadas. Se van a actualizar hasta la semana: " & CurrentPlusFuture
            Dim CounterWeek As Integer
            Dim var As Integer 'Variable vacía para función'
            CounterWeek = NumExtract(WeldingSheet.Cells(OffsetFilaCabecera() - 2, WeldingWeekLastCol))
            Dim NextWeekCol As Integer
            NextWeekCol = WeldingWeekLastCol + WeldingColDistance()
            For i = CounterWeek + 1 To CurrentPlusFuture
                CounterWeek = i
                var = AddWeek(CounterWeek, NextWeekCol)
                var = ProdNeed(CounterWeek)
                var = ProdPlan(CounterWeek)
                ImportWeekEDI (CounterWeek)
                var = WeldingAccumulated(CounterWeek)
                'WeldingWeekLastCol = WeldingWeekLastCol + WeldingColDistance 'Avanzamos a la siguiente celda donde se colocará la semana'
                NextWeekCol = NextWeekCol + WeldingColDistance()
            Next i
        Else
        End If
    End If
End Sub

Function NumExtract(cell As Range) As Integer
    Dim regex As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d+"
    
    Set match = regex.Execute(cell.Value)
    
    If match.Count > 0 Then
        NumExtract = CInt(match.Item(0))
    Else
        NumExtract = 0
    End If
End Function

Function AddWeek(week As Integer, WeekCol As Integer) As Integer
    'Añade la semana pasada por argumento a la pestaña WELDING'
    'En el procedimiento del módulo 16 ya se comprueba que las semanas no estén completamente actualizada'
    'Similar al MODULE_4 WeeksHeaders, pero añadiendo unicamente la semana deseada.
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    
    'Actual, Cargas, Necesidad, Plan'
    WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol).Value = "Actual"
    WeldingSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).Value = "Week " & week '+ 1
    WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + 1).Value = "Cargas W" & week '+ 1
    WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + 2).Value = "Necesidad de producción"
    WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + 3).Value = "Plan de producción"
    
    'N,D,T para cada día de la semana'
    For i = 1 To 18 Step 3
        WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 3).Value = "N"
        WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 4).Value = "D"
        WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 5).Value = "T"
    Next i
    
    'Fecha encima de cada día de la semana'
    Dim Counter As Integer
    For i = 1 To 6
        Counter = i
        WeldingSheet.Cells(OffsetFilaCabecera() - 1, WeekCol + (i * 3) + 1).Value = GetDate(week, Counter)
    Next i
    
    Dim FormatSheet As Worksheet
    Set FormatSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    Dim FormatRange As Range
    Set FormatRange = FormatSheet.Range("A14:V16")
    
    'Copia de formatos desde pestaña "Formats"'
    Dim WeldingRange As Range
    Set WeldingRange = WeldingSheet.Range(WeldingSheet.Cells(OffsetFilaCabecera() - 2, WeekCol), WeldingSheet.Cells(OffsetFilaCabecera(), WeekCol + WeldingColDistance() - 1))
    FormatRange.Copy
    WeldingRange.PasteSpecial xlPasteFormats
    CompleteWeekFormat (week)
End Function

