Attribute VB_Name = "Module28"
'MODULE_28'
Sub AddBendingWeekHeaders(week As Integer, WeekCol As Integer)
    'Dimensionado de hojas de trabajo'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    'N,D, T para turno del día'
    For i = 0 To WeekShifts() - 1 Step 3
        BendingSheet.Cells(OffsetFilaCabecera(), WeekCol + i) = "N"
        BendingSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 1) = "D"
        BendingSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 2) = "T"
    Next i
    
    'Fecha encima de cada día de la semana'
    Dim Counter As Integer
    Counter = 1
    For i = 0 To WeekShifts() - 1 Step 3
        BendingSheet.Cells(OffsetFilaCabecera() - 1, WeekCol + i).Value = GetDate(week, Counter)
        Counter = Counter + 1
    Next i
    
    'Número de la semana'
    BendingSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).Value = "Week " & week
    
    'Copiar formatos de celda desde la pestaña FORMATS'
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A66:R68")
    FormatRange.Copy
    BendingSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).PasteSpecial xlPasteFormats
    
End Sub

