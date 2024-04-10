Attribute VB_Name = "Module23"
'MODULE 23' 'COPIAS DE SEGURIDAD'
Sub Welding_backup_sec()
    'CREA COPIA DE SEGURIDAD EN LA PESTAÑA WELDING_BACKUP_SEC'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("WELDING_BACKUP_SEC"))
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    
    WeldingSheet.UsedRange.Copy BackupSheet.Range("A1")
End Sub

Sub Box_backup_sec()
    'CREA COPIA DE SEGURIDAD EN LA PESTAÑA BOX_BACKUP_SEC'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("BOX_BACKUP_SEC"))
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    
    BoxSheet.UsedRange.Copy BackupSheet.Range("A3")
End Sub

Sub Bending_backup_sec()
    'CREA COPIA DE SEGURIDAD EN LA PESTAÑA BENDING_backup_sec'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("BENDING_BACKUP_SEC"))
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    
    BendingSheet.UsedRange.Copy BackupSheet.Range("A1")
End Sub
