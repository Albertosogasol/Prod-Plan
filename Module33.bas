Attribute VB_Name = "Module33"
'MODULE_33'
'ACTUALIZACI�N COMPLETA PESTA�A BENDING'
Sub BendingSheetClearUpdateAll()
    'Realiza una limpieza completa de la pesta�a Bending. Crea una copia y la recupera de la pesta�a de backups
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    Bending_backup
    BendingSheet.UsedRange.Clear
    BendingHeaders
    BendingReferences
    BendingWeeksBuilder
    BendingBackupToBending
    'MsgBox "Finalizado con �xito"
End Sub
