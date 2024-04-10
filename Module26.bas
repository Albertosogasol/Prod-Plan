Attribute VB_Name = "Module26"
'MODULE_26'
'CABECERAS, CUERPO Y DATOS DE PESTAÑA BENDING'
Sub BendingHeaders()
    'Cabeceras generales'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    Dim FormatSheet As Worksheet
    Set FormatSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    
    BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("Linea")).Value = "Línea"
    BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("CD&V")).Value = "CD&V"
    BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("ID")).Value = "ID"
    BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("reference")).Value = "Referencia"
    BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("reference")).Columns.AutoFit
    
    Dim FormatRange As Range
    Set FormatRange = FormatSheet.Range("A55:D55")
    FormatRange.Copy
    
    Dim destRange As Range
    Set destRange = BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("Linea"))
    destRange.PasteSpecial xlPasteFormats
End Sub

