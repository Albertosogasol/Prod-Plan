Attribute VB_Name = "Module27"
'MODULE_27'
Sub BendingReferences()
    'COLOCA LAS REFERENCIAS DE CURVADO EN LA PESTAÑA BENDING.
    'LAS OBTIENE DE LA PESTAÑA PROCESS
    
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("bending"))
    Dim ProcessSheet As Worksheet
    Set ProcessSheet = ThisWorkbook.Worksheets(SheetName("Process"))
    Dim ReferencesSheet As Worksheet
    Set ReferencesSheet = ThisWorkbook.Worksheets(SheetName("References"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    'Última fila de la columna de rerencias en la pestaña de procesos'
    Dim LastRowProcess As Integer
    LastRowProcess = ProcessSheet.Cells(Rows.Count, ProcessCol("References")).End(xlUp).Row
    
    'Contador para referencias'
    Dim refCounter As Integer 'Posición en la que se colocará la siguiente referencia encontrada en el bucle'
    refCounter = OffsetFilaCabecera() + 1
    
    'Rango para copia de formatos'
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A59:D62")
    FormatRange.Copy
    
    'Bucle para todas las referencias de la pestaña PROCESS'
    For i = 1 To LastRowProcess
        'Comprobar si la celda pertenece a BENDING
        If InStr(1, ProcessSheet.Cells(i, ProcessCol("Process")).Value, "Bending", vbTextCompare) > 0 Then
            'Copia de datos desde la pestaña Process'
            BendingSheet.Cells(refCounter, NumColWelding("Reference")).Value = "'" & ProcessSheet.Cells(i, ProcessCol("Reference")).Value
            BendingSheet.Cells(refCounter, NumColWelding("Line")).Value = "Curv." & ProcessSheet.Cells(i, ProcessCol("Line")).Value
            BendingSheet.Cells(refCounter, NumColWelding("Capacity")).Value = ProcessSheet.Cells(i, ProcessCol("Project")).Value
            BendingSheet.Cells(refCounter, NumColWelding("ID")).Value = ProcessSheet.Cells(i, ProcessCol("ID")).Value
            BendingSheet.Cells(refCounter + 3, NumColWelding("Capacity")).Value = "Capacidad/turno"
            BendingSheet.Cells(refCounter + 3, NumColWelding("Reference")).Value = ProcessSheet.Cells(i, ProcessCol("Capacity")).Value
            
            'Copia de formatos'
            BendingSheet.Cells(refCounter, NumColWelding("Line")).PasteSpecial xlPasteFormats
            
            refCounter = refCounter + BendingRowDistance()
        End If
    Next i
End Sub
