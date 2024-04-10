Attribute VB_Name = "Module33"
'MODULE_33'
'ACTUALIZACIÓN COMPLETA PESTAÑA BENDING'
Sub BendingSheetClearUpdateAll()
    'Realiza una limpieza completa de la pestaña Bending. Crea una copia y la recupera de la pestaña de backups
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    Bending_backup
    BendingSheet.UsedRange.Clear
    BendingHeaders
    BendingReferences
    BendingWeeksBuilder
    BendingBackupToBending
    'MsgBox "Finalizado con éxito"
End Sub
