Attribute VB_Name = "Module8"
'MODULE_8'
'Copy of WELDING_backup to WELDING'

Sub WeldingBackupToWelding() 'MEJORADO
    'Proceso de copia a trav�s de rangos'
    'Se hace uso del procedimiento WeldingRefeferenceRow()
    Dim WeldingSheet As Worksheet
    Dim BackupSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("Welding_backup"))
    
    Dim ReferenceRow As Integer
    Dim BackupRange As Range
    Dim tempRef As String 'Variable donde se almacena la referencia obtenida en el bucle'
    Dim destRange As Range
    Dim initCol As Integer
    Dim initRow As Integer
    Dim finalCol As Integer
    Dim finalRow As Integer
    Dim WeldingBackupLastCol As Integer
    WeldingBackupLastCol = BackupSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Dim WeldingBackupLastRow As Integer
    WeldingBackupLastRow = BackupSheet.Cells(Rows.Count, NumColBox("reference")).End(xlUp).Row
    Dim WeldingLastCol As Integer
    WeldingLastCol = WeldingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Dim WeldingReferenceRowVar As Integer
    For j = (OffsetFilaCabecera() + 1) To WeldingBackupLastRow Step WeldingRowDistance()
        'Para optimizar el proceso, se dimensiona un rango para cada referencia, que contiene todos los valores de las producciones,
        'desde la semana 1 hasta el final de la hoja. Ese rango creado en la pesta�a de backup se pega en la correspondiente fila en la pesta�a
        'WELDING.
        'Previo al pegado especial es necesario localizar la posici�n de la celda de destino en la pesta�a WELDING. Para ello se utiliza
        'la funci�n WeldingReferenceRow() del m�dulo 1
        'On Error Resume Next
        tempRef = BackupSheet.Cells(j, NumColWelding("Reference")).Value
        'Se crea un rango del tama�o 2x(total_columnas)
        initRow = j
        initCol = FirstActualCol()
        finalRow = j + 1
        finalCol = WeldingBackupLastCol
        Set BackupRange = BackupSheet.Range(BackupSheet.Cells(initRow, initCol), BackupSheet.Cells(finalRow, finalCol))
        'Copia del rango completo'
        BackupRange.Copy
        'Se pega el rango copiado en la celda correspondiente'
        WeldingReferenceRowVar = WeldingReferenceRow(tempRef)
        On Error GoTo ErrorHandler
        Set destRange = WeldingSheet.Range(WeldingSheet.Cells(WeldingReferenceRowVar, FirstActualCol()), WeldingSheet.Cells(WeldingReferenceRowVar + 1, WeldingLastCol))
        'destRange.PasteSpecial Paste:=xlPasteValues
        destRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False 'Limpia el portapapeles'
    Next j
    Exit Sub
    
ErrorHandler:
    'Al eliminar una referencia, esta se queda copiada en la pesta�a WELDING_backup, por lo tanto en el momento de la b�squeda,
    'devuelve error al no ser encontrada. Si esta situaci�n se produce, se elimina la referencia de la pesta�a de backup y se vuelve
    'a llamar al procedimiento
    MsgBox "Se ha producido un error en la b�squeda de la referencia " & tempRef & ". Es necesario borrarla de la pesta�a WELDING_backup"
    answer = MsgBox("�Desea borrar la referencia y todo su contenido?", vbQuestion + vbYesNo, "Elegir opci�n")
    If answer = vbYes Then
        Dim Row As Integer
        Row = Application.match(tempRef, BackupSheet.Columns(NumColWelding("Reference")), 0)
        Dim DelRange As Range
        Set DelRange = BackupSheet.Range("A" & Row & ":A" & (Row + 3)) 'NUMERO TOTAL DE FILAS QUE OCUPA UNA PESTA�A'
        DelRange.EntireRow.Delete
        WeldingBackupToWelding
    Else
    End If

    'Sub WeldingBackupToWelding()
    '    'Copia el contenido de la pesta�a WELDING_backup a la pesta�a WELDING'
    '    Dim WeldingSheet As Worksheet
    '    Dim WeldingSheet_backup As Worksheet
    '    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    '    Set WeldingSheet_backup = ThisWorkbook.Worksheets(SheetName("Welding_backup"))
    '
    '    'Dimensionado de rango pesta�a WELDING_backup'
    '    Dim WeldingBackupRange As Range
    '    Dim WeldingBackupLastCol As Long
    '    Dim WeldingBackupLastRow As Long
    '    WeldingBackupLastCol = WeldingSheet_backup.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    '    WeldingBackupLastRow = WeldingSheet_backup.Cells(Rows.Count, NumColWelding("reference")).End(xlUp).row
    '    Set WeldingBackupRange = WeldingSheet_backup.Range(WeldingSheet_backup.Cells(OffsetFilaCabecera, NumColWelding("reference")), WeldingSheet_backup.Cells(WeldingBackupLastRow, WeldingBackupLastCol))
    '
    '    'Dimensionado de rango pesta�a WELDING'
    '    Dim WeldingRange As Range
    '    Dim WeldingRangeContents As Range 'Mismo rango que el anterior pero eliminando cabeceras y refrencias. Unicmente datos en bruto'
    '    Dim WeldingLastCol As Long
    '    Dim WeldingLastRow As Long
    '    WeldingLastCol = WeldingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    '    WeldingLastRow = WeldingSheet.Cells(Rows.Count, NumColWelding("reference")).End(xlUp).row
    '    Set WeldingRange = WeldingSheet.Range(WeldingSheet.Cells(OffsetFilaCabecera, NumColWelding("reference")), WeldingSheet.Cells(WeldingBackupLastRow, WeldingBackupLastCol))
    '    Set WeldingRangeContents = WeldingSheet.Range(WeldingSheet.Cells(OffsetFilaCabecera() + 1, NumColWelding("reference") + 1), WeldingSheet.Cells(WeldingBackupLastRow, WeldingBackupLastCol))
    '
    '    'Copia de datos de WELDING_backup a WELDING mediante VLookup'
    '    Dim RefBackupPosX As Long 'Columna en la que se encuentra la referencia buscada con Vlookup en la pesta�a backup'
    '    Dim RefBackupPosY As Long 'Fila en la que se encuentra la referencia buscada con Vlookup en la pesta�a backup'
    '
    '    'Borra todos los datos copiados del rango para hacer una copia limpia de la actualizaci�n'
    '    WeldingRangeContents.ClearContents
    '
    '    'Bucle para recorrer todo las celdas y copiar los valores'
    '    For i = OffsetFilaCabecera() + 1 To WeldingLastRow Step 3
    '        RefBackupPosX = 2
    '        For j = FirstActualCol() To WeldingLastCol
    '            On Error Resume Next 'Salto a siguiente referencia si Vlookup no encuentra la nueva referencia'
    '            WeldingSheet.Cells(i, j).Value = Application.WorksheetFunction.VLookup(WeldingSheet.Cells(i, NumColWelding("Reference")).Value, WeldingSheet_backup.Range("D:MMM"), RefBackupPosX, False)
    '            On Error GoTo 0 'Manejo de errores no aplicado a otras partes del c�digo'
    '            RefBackupPosX = RefBackupPosX + 1
    '        Next j
    '    Next i
    'End Sub
End Sub
