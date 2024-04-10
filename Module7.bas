Attribute VB_Name = "Module7"
'MODULE_7'
'BACKUP PESTA�A WELDING'
Sub Welding_backup()
    'Realiza una copia de la pesta�a WELDING en la pesta�a WELDING_backup'
    
    Dim Welding As Worksheet
    Dim Welding_backup As Worksheet
    Set Welding = ThisWorkbook.Worksheets(SheetName("Welding"))
    Set Welding_backup = ThisWorkbook.Worksheets(SheetName("Welding_backup"))
    
    'Obtenemos el rango de la pesta�a WELDING. Comenzando en columna REFERENCE'
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = Welding.Cells(Rows.Count, NumColWelding("reference")).End(xlUp).Row
    lastCol = Welding.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Dim WeldingRange As Range
    Set WeldingRange = Welding.Range(Welding.Cells(OffsetFilaCabecera, NumColWelding("reference")), Welding.Cells(lastRow, lastCol))
    
    'Se copia y pega todo el rango. M�todo m�s rapido'
    WeldingRange.Copy
    Dim destRow As Integer
    destRow = OffsetFilaCabecera() ' Llamada a tu funci�n para obtener la fila de destino
    Dim destCol As Integer
    destCol = NumColWelding("reference") ' Llamada a tu funci�n para obtener la columna de destino
    Dim destCell As Range
    Set destCell = Welding_backup.Cells(destRow, destCol)

    ' Pega los valores del rango en la celda de destino
    'destCell.PasteSpecial xlPasteValues
    destCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub

