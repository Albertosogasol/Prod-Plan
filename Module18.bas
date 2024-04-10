Attribute VB_Name = "Module18"
'MODULE_18'
'REFERENCIAS CAJAS'
Sub BoxReferences()
    'Coloca las referencias de soldadura en la pestaña BOX'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Dim ProcessSheet As Worksheet
    Set ProcessSheet = ThisWorkbook.Worksheets(SheetName("Process"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    'Obtener la última fila de la columna de referencias en la pestaña de procesos'
    Dim LastRowProcess As Integer
    LastRowProcess = ProcessSheet.Cells(Rows.Count, ProcessCol("References")).End(xlUp).Row
    
    'Contador para recorrer referencias'
    Dim refCounter As Integer
    refCounter = OffsetFilaCabecera() + 1
    
    'Rango para formatos de celda'
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A22:D25")
    
    'Iterar todas las filas de la columna "A" de la pestaña "Process"'
    For i = 1 To LastRowProcess
        'Comprobar si la celda en la columna C contiene la palabra "BOX"'
        If InStr(1, ProcessSheet.Cells(i, ProcessCol("Process")).Value, "Box", vbTextCompare) > 0 Then
            'Datos desde pestaña de procesos'
            BoxSheet.Cells(refCounter, NumColWelding("Reference")).Value = "'" & ProcessSheet.Cells(i, ProcessCol("Reference")).Value
            BoxSheet.Cells(refCounter, NumColWelding("Linea")).Value = "LC " & ProcessSheet.Cells(i, ProcessCol("Linea")).Value
            BoxSheet.Cells(refCounter + 3, NumColWelding("Capacity")).Value = "Capacidad/turno"
            BoxSheet.Cells(refCounter + 3, NumColWelding("Reference")).Value = ProcessSheet.Cells(i, ProcessCol("Capacity")).Value
            BoxSheet.Cells(refCounter, NumColWelding("ID")).Value = ProcessSheet.Cells(i, ProcessCol("ID")).Value
            BoxSheet.Cells(refCounter, NumColWelding("Capacity")).Value = ProcessSheet.Cells(i, ProcessCol("Project")).Value
            'Formato desde pestaña FORMATS'
            FormatRange.Copy
            BoxSheet.Cells(refCounter, NumColWelding("Linea")).PasteSpecial xlPasteFormats
            refCounter = refCounter + BoxRowDistance()
        End If
    Next i
End Sub
