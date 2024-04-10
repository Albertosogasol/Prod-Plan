Attribute VB_Name = "Module31"
'MODULE_31'
'BACKUP PESTA�A BENDING'
Sub Bending_backup()
    'Creaci�n de copia de seguridad para recuperaci�n de datos en pesta�a BENDING'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    Dim backup As Worksheet
    Set backup = ThisWorkbook.Worksheets(SheetName("BENDING_BACKUP"))
    
    'Obtenemos el rango de la pesta�a BENDING. Comenzando en columna REFERENCE'
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = BendingSheet.Cells(Rows.Count, NumColBending("reference")).End(xlUp).Row
    lastCol = BendingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Dim BendingRange As Range
    Set BendingRange = BendingSheet.Range(BendingSheet.Cells(OffsetFilaCabecera, NumColBending("reference")), BendingSheet.Cells(lastRow, lastCol))
    
    'Se copia y pega todo el rango. M�todo m�s rapido'
    BendingRange.Copy
    Dim destRow As Integer
    destRow = OffsetFilaCabecera() ' Llamada a tu funci�n para obtener la fila de destino
    Dim destCol As Integer
    destCol = NumColBending("reference") ' Llamada a tu funci�n para obtener la columna de destino
    Dim destCell As Range
    Set destCell = backup.Cells(destRow, destCol)

    ' Pega los valores del rango en la celda de destino
    'destCell.PasteSpecial xlPasteValues
    destCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
    
End Sub

