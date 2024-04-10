Attribute VB_Name = "Module17"
'MODULE_17'
'CABECERAS GENERALES BOX'
Sub BoxHeaders()
    'Se crean las cabeceras generales de la pestaña BOX'
    'Línea, CD&V, ID, REFERENCE'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Dim FormatSheet As Worksheet
    Set FormatSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))
    
    BoxSheet.Cells(OffsetFilaCabecera(), NumColWelding("Line")).Value = "Línea"
    BoxSheet.Cells(OffsetFilaCabecera(), NumColWelding("Capacidad")).Value = "CD&V"
    BoxSheet.Cells(OffsetFilaCabecera(), NumColWelding("ID")).Value = "ID"
    BoxSheet.Cells(OffsetFilaCabecera(), NumColWelding("Reference")).Value = "Referencia"
    
    Dim FormatRange As Range
    Set FormatRange = FormatSheet.Range("A19:D19")
    Dim BoxRange As Range
    Set BoxRange = BoxSheet.Cells(OffsetFilaCabecera(), NumColWelding("Line"))
    FormatRange.Copy
    BoxRange.PasteSpecial xlPasteFormats
    
End Sub
