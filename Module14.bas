Attribute VB_Name = "Module14"
'MODULE_14'
'FORMATO COMPLETO A CADA SEMANA COPIADO DE PESTAÑA FORMATS'
Function RefWeekFormat(week As Integer, Row As Integer) As Integer
    'Copia el formato unicamente a la referencia deseada en la semana deseada'
    'Ambos parámetros son argumentos de la función'
    Dim WeldingSheet As Worksheet
    Dim formatsSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Set FormatsSheets = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    Dim FormatRange As Range
    Dim WeldingRange As Range
    Set FormatRange = FormatsSheets.Range("A48:V51")
    Dim WeekCol As Integer
    'MsgBox Week
    WeekCol = WeldingWeekSearch(week)
    'MsgBox WeekCol
    'MsgBox row
    FormatRange.Copy
    'Set WeldingRange = WeldingSheet.Range(WeldingSheet.Cells(row, WeekCol)) ', WeldingSheet.Cells(row + 1, WeekCol + WeldingColDistance - 1)
    Set WeldingRange = WeldingSheet.Cells(Row, WeekCol)
    WeldingRange.PasteSpecial xlPasteFormats

End Function

Sub CompleteWeekFormat(week As Integer)
    '    'Actualiza el formato de todas las referencias para la semana deseada'
    '    Dim WeldingSheet As Worksheet
    '    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    ''    Dim Week As Integer
    ''    Week = 2
    '    Dim LastRowWelding As Integer
    '    LastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).row
    '    Dim row As Integer
    '    Dim Valor As Integer
    '    For i = OffsetFilaCabecera() + 1 To LastRowWelding Step WeldingRowDistance()
    '        row = i
    '        Valor = RefWeekFormat(Week, row)
    '    Next i
    
    'Formato mejorado creando un rango de toda la semana' 'Actualización 10/05/2023'
    Dim WeldingSheet As Worksheet
    Dim formatsSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Set FormatsSheets = ThisWorkbook.Worksheets(SheetName("Formats"))
    Dim FormatRange As Range
    Set FormatRange = FormatsSheets.Range("A48:V51")
    Dim lastRowWelding As Integer
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row
    Dim WeldingRange As Range
    Set WeldingRange = WeldingSheet.Range(WeldingSheet.Cells(OffsetFilaCabecera() + 1, WeldingWeekSearch(week)), WeldingSheet.Cells(lastRowWelding + 2, WeldingWeekSearch(week) + 21))
    FormatRange.Copy
    WeldingRange.PasteSpecial xlPasteFormats
End Sub

Sub CompleteFormat()
    For i = StartWeek() To CurrentWeekNumber() + FutureWeeks()
        CompleteWeekFormat (i)
    Next i
End Sub
