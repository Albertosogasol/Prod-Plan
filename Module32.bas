Attribute VB_Name = "Module32"
'MODULE_32'
'RECUPERACI�N DE DATOS DESDE PESTA�A BENDING_backup A LA PESTA�A BENDING
Sub BendingBackupToBending()
    'Copia los datos de la pesta�a BENDING_backup a la pesta�a BENDING, en funci�n de la referencia'
    'Se utiliza principalemente en el momento de a�adir referencias nuevas'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("Bending_backup"))
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    
    'Dimensionado de rango pesta�a BENDING_backup'
    Dim BendingBackupRange As Range
    Dim lastCol As Integer
    Dim BendingBackupLastRow As Integer
    BendingBackupLastCol = BackupSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    BendingBackupLastRow = BackupSheet.Cells(Rows.Count, NumColBending("reference")).End(xlUp).Row
    Set BendingBackupRange = BackupSheet.Range(BackupSheet.Cells(6, 4), BackupSheet.Cells(34, 382))

    'Dimensionado pesta�a BOX'
    Dim BendingRange As Range
    Dim BendingRangeContents As Range 'Mismo rango que el anterior pero tomando s�lo los datos, sin cabeceras'
    Dim BendingLastCol As Integer
    Dim BendingLastRow As Integer
    BendingLastCol = BendingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    BendingLastRow = BendingSheet.Cells(Rows.Count, NumColBending("Reference")).End(xlUp).Row
    Set BendingRange = BendingSheet.Range(BendingSheet.Cells(OffsetFilaCabecera(), NumColBending("Reference")), BendingSheet.Cells(BendingBackupLastRow, BendingBackupLastCol))
    Set BendingRangeContents = BendingSheet.Range(BendingSheet.Cells(OffsetFilaCabecera() + 1, NumColBending("Reference") + 1), BendingSheet.Cells(BendingBackupLastRow, BendingBackupLastCol))
    
    'Copia de datos desde pesta�a Box_backup a pesta�a BENDING mediante funci�n VLookup'
    'No se copian todas las celdas. Unicamente las celdas de agregados. Las celdas restantes se calculan a partir de f�rmulas'
    Dim RefBackupPosX As Long 'Columna en la que se encuentra la referencia buscada con VLookup en la pesta�a backup'
    Dim RefBackupPosY As Long 'Fila en la que se encuentra la referencia buscada con VLookup en la pesta�a backup'
    
    'Borrado de todos los datos en la pesta�a Box'
    BendingRangeContents.ClearContents
    
    'Aplicaci�n de f�rmulas a pesta�a BENDING'
    'Llamada a procedimiento BendingWeekBody pas�ndole la semana y posici�n'
    Dim week As Integer
    Dim WeekCol As Integer
    WeekCol = FirstBendingData()
    Dim Contador As Integer
    'Se aplican las f�rmulas
    For i = FirstBendingData() To BendingLastCol Step BendingColDistance()
            week = NumExtract(BendingSheet.Cells(OffsetFilaCabecera() - 2, i))
            Contador = i 'Contador para semanas. No se pasa como argumento la variable "i"'
            Call BendingWeekBody(week, Contador)
    Next i
    
    'Recuperaci�n de datos de agregados desde pesta�a Box_backup'
    'Usar BoxReferenceRow
    Dim ReferenceRow As Integer
    Dim BackupRange As Range
    Dim tempRef As String 'Variable donde se copia la referencia que se lee en el bucle siguiente'
    Dim destRange As Range
    Dim initCol As Integer
    Dim initRow As Integer
    Dim finalCol As Integer
    Dim finalRow As Integer
    Dim BendingReferenceRowVar As Integer
    For j = (OffsetFilaCabecera() + 1) To BendingBackupLastRow Step 4
        'Para optimizar el proceso, se dimensiona un rango para cada referencia, que contiene todos los valores de las producciones,
        'desde la semana 1 hasta el final de la hoja. Ese rango creado en la pesta�a de backup se pega en la correspondiente fila en la pesta�a BOX
        'Previo al pegado especial es necesario localizar la posici�n de la celda de destino en la pesta�a BOX. Para ello se utiliza
        'la funci�n BoxReferenceRow() del m�dulo 1
        'On Error Resume Next
        tempRef = BackupSheet.Cells(j, NumColBending("Reference")).Value
        'Se crea un rango del tama�o 2x(total_columnas)
        initRow = j + 2
        initCol = FirstBendingData()
        finalRow = j + 3
        finalCol = BendingBackupLastCol
        Set BackupRange = BackupSheet.Range(BackupSheet.Cells(initRow, initCol), BackupSheet.Cells(finalRow, finalCol))
        'Se realiza copia del rango completo'
        BackupRange.Copy
        'Se pega el rango copiado en la celda correspondiente'
        BendingReferenceRowVar = BendingReferenceRow(tempRef) + 2
        On Error GoTo ErrorHandler
        Set destRange = BendingSheet.Range(BendingSheet.Cells(BendingReferenceRowVar, FirstBendingData()), BendingSheet.Cells(BendingReferenceRowVar + 1, BendingLastCol))
        destRange.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False 'Limpia el portapapeles'
    Next j
    Exit Sub
    
ErrorHandler:
    'Al eliminar una referencia, esta se queda copiada en la pesta�a WELDING_backup, por lo tanto en el momento de la b�squeda,
    'devuelve error al no ser encontrada. Si esta situaci�n se produce, se elimina la referencia de la pesta�a de backup y se vuelve
    'a llamar al procedimiento
    MsgBox "Se ha producido un error en la b�squeda de la referencia " & tempRef & ". Es necesario borrarla de la pesta�a BOX_backup"
    answer = MsgBox("�Desea borrar la referencia y todo su contenido?", vbQuestion + vbYesNo, "Elegir opci�n")
    If answer = vbYes Then
        Dim Row As Integer
        Row = Application.match(tempRef, BackupSheet.Columns(NumColBending("Reference")), 0)
        Dim DelRange As Range
        Set DelRange = BackupSheet.Range("A" & Row & ":A" & (Row + 3))
        DelRange.EntireRow.Delete
        BendingBackupToBending
    Else
    End If
End Sub
