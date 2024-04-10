Attribute VB_Name = "Module13"
'MODULE_13'
'APLICA COLORES EN FUNCIÓN DE FORMATOS CONDICIONALES A CELDAS DE PLAN DE PRODUCCIÓN'
Sub PlanFormatWeek(week As Integer)
    'Aplicando formatos plantilla en pestaña "Formats"
    Dim formatsSheet As Worksheet
    Dim WeldingSheet As Worksheet
    
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("B2:B3")
    Dim WeldingRange As Range
    
    Dim lastRowWelding As Long
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row
    
    For i = OffsetFilaCabecera + 1 To lastRowWelding Step 3
        FormatRange.Copy
        Set WeldingRange = WeldingSheet.Range(WeldingSheet.Cells(i, WeldingWeekSearch(week) + 3), WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 3))
        WeldingRange.PasteSpecial xlPasteFormats
    Next i
End Sub

Sub PlanFormatWeekUpdateAll()
    'Aplicado a todas las celdas'
    For i = StartWeek() To CurrentWeekNumber() + FutureWeeks()
        PlanFormatWeek (i)
    Next i
End Sub

