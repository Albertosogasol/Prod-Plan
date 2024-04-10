Attribute VB_Name = "Module21"
'MODULE_21'
'BOX BACKUP'
Sub Box_backup()
    'Realiza un backup de la pesta�a Box en la pesta�a oculta Box_backup'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("Box_backup"))
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    
    Dim LastColBox As Integer
    Dim LastRowBox As Integer
    
    LastRowBox = BoxSheet.Cells(Rows.Count, NumColBox("capacity")).End(xlUp).Row
    LastColBox = BoxSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    
    'Obtenemos el rango de la pesta�a BOX. Comenzando en columna REFERENCE
    Dim BoxRange As Range
    Set BoxRange = BoxSheet.Range(BoxSheet.Cells(OffsetFilaCabecera(), NumColBox("Reference")), BoxSheet.Cells(LastRowBox, LastColBox))
    
    'Se copia y pega todo el rango'
    BoxRange.Copy
    Dim destRow As Integer
    destRow = OffsetFilaCabecera() ' Llamada para obtener la fila de destino
    Dim destCol As Integer
    destCol = NumColBox("reference") ' Llamada para obtener la columna de destino
    Dim destCell As Range
    Set destCell = BackupSheet.Cells(destRow, destCol)

    ' Pega los valores del rango en la celda de destino
    destCell.PasteSpecial xlPasteValues
    
    ' Limpia el portapapeles
    Application.CutCopyMode = False
End Sub

Sub BoxBackupToBox()
    'Copia los datos de la pesta�a Box_backup a la pesta�a Box, en funci�n de la referencia'
    'Se utiliza principalemente en el momento de a�adir referencias nuevas'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("Box_backup"))
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    
    'Dimensionado de rango pesta�a BOX_backup'
    Dim BoxBackupRange As Range
    Dim lastCol As Integer
    Dim BoxBackupLastRow As Integer
    BoxBackupLastCol = BackupSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    BoxBackupLastRow = BackupSheet.Cells(Rows.Count, NumColBox("reference")).End(xlUp).Row
    'MsgBox OffsetFilaCabecera() & " " & NumColBox("reference") & " a celda " & BoxBackupLastRow & " " & BoxBackupLastCol
    'Set BoxBackupRange = BackupSheet.Range(BackupSheet.Cells(OffsetFilaCabecera(), NumColBox("reference")), BackupSheet.Cells(WeldingBackupLastRow, WeldingBackupLastCol))
    Set BoxBackupRange = BackupSheet.Range(BackupSheet.Cells(6, 4), BackupSheet.Cells(34, 382))
    'Dimensionado pesta�a BOX'
    Dim BoxRange As Range
    Dim BoxRangeContents As Range 'Mismo rango que el anterior pero tomando s�lo los datos, sin cabeceras'
    Dim BoxLastCol As Integer
    Dim BoxLastRow As Integer
    BoxLastCol = BoxSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    BoxLastRow = BoxSheet.Cells(Rows.Count, NumColBox("Reference")).End(xlUp).Row
    Set BoxRange = BoxSheet.Range(BoxSheet.Cells(OffsetFilaCabecera(), NumColBox("Reference")), BoxSheet.Cells(BoxBackupLastRow, BoxBackupLastCol))
    Set BoxRangeContents = BoxSheet.Range(BoxSheet.Cells(OffsetFilaCabecera() + 1, NumColBox("Reference") + 1), BoxSheet.Cells(BoxBackupLastRow, BoxBackupLastCol))
    
    'Copia de datos desde pesta�a Box_backup a pesta�a Box mediante funci�n VLookup'
    'No se copian todas las celdas. Unicamente las celdas de agregados. Las celdas restantes se calculan a partir de f�rmulas'
    Dim RefBackupPosX As Long 'Columna en la que se encuentra la referencia buscada con VLookup en la pesta�a backup'
    Dim RefBackupPosY As Long 'Fila en la que se encuentra la referencia buscada con VLookup en la pesta�a backup'
    
    'Borrado de todos los datos en la pesta�a Box'
    BoxRangeContents.ClearContents
    
    'Aplicaci�n de f�rmulas a pesta�a Box'
    'Llamada a procedimiento BoxWeekBody pas�ndole la semana y posici�n'
    Dim week As Integer
    Dim WeekCol As Integer
    WeekCol = FirstBoxData()
    Dim Contador As Integer
    'Se aplican las f�rmulas
    For i = FirstBoxData() To BoxLastCol Step BoxColDistance()
            week = NumExtract(BoxSheet.Cells(OffsetFilaCabecera() - 2, i))
            Contador = i 'Contador para semanas. No se pasa como argumento la variable "i"'
            Call BoxWeekBody(week, Contador)
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
    Dim BoxReferenceRowVar As Integer
    For j = (OffsetFilaCabecera() + 1) To BoxBackupLastRow Step 4
        'Para optimizar el proceso, se dimensiona un rango para cada referencia, que contiene todos los valores de las producciones,
        'desde la semana 1 hasta el final de la hoja. Ese rango creado en la pesta�a de backup se pega en la correspondiente fila en la pesta�a BOX
        'Previo al pegado especial es necesario localizar la posici�n de la celda de destino en la pesta�a BOX. Para ello se utiliza
        'la funci�n BoxReferenceRow() del m�dulo 1
        'On Error Resume Next
        tempRef = BackupSheet.Cells(j, NumColBox("Reference")).Value
        'Se crea un rango del tama�o 2x(total_columnas)
        initRow = j + 2
        initCol = FirstBoxData()
        finalRow = j + 3
        finalCol = BoxBackupLastCol
        Set BackupRange = BackupSheet.Range(BackupSheet.Cells(initRow, initCol), BackupSheet.Cells(finalRow, finalCol))
        'Se realiza copia del rango completo'
        BackupRange.Copy
        'Se pega el rango copiado en la celda correspondiente'
        BoxReferenceRowVar = BoxReferenceRow(tempRef) + 2
        On Error GoTo ErrorHandler
        Set destRange = BoxSheet.Range(BoxSheet.Cells(BoxReferenceRowVar, FirstBoxData()), BoxSheet.Cells(BoxReferenceRowVar + 1, BoxLastCol))
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
        Row = Application.match(tempRef, BackupSheet.Columns(NumColBox("Reference")), 0)
        Dim DelRange As Range
        Set DelRange = BackupSheet.Range("A" & Row & ":A" & (Row + 3))
        DelRange.EntireRow.Delete
        BoxBackupToBox
    Else
    End If
End Sub
