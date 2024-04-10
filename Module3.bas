Attribute VB_Name = "Module3"
'MODULE_3'
'ACTUALIZACIÓN REFERENCIAS SOLDADURA'
Sub WeldingReferences()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim ProcessSheet As Worksheet
    Set ProcessSheet = ThisWorkbook.Worksheets(SheetName("PROCESS"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))
    
    Dim ProcessLastRow As Integer
    ProcessLastRow = ProcessSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Dim iCounter As Integer 'Contador para el interior del bucle'
    iCounter = OffsetFilaCabecera() + 1
    
    'Rangos para copiar formatos'
    Dim formatRange1 As Range
    ' Dim formatRange_WIP1, formatRange_WIP2, formatRange_WIP3 As Range
     Set formatRange1 = formatsSheet.Range("A41:D44")
    ' Set formatRange_WIP1 = FormatsSheet.Range("E41")
    ' Set formatRange_WIP2 = FormatsSheet.Range("F41")
    ' Set formatRange_WIP3 = FormatsSheet.Range("G41")
    Dim TargetRange As Range
    
    For i = 1 To ProcessLastRow
        If InStr(1, ProcessSheet.Cells(i, NumColProcess("PROCESS")).Value, "Welding", vbTextCompare) > 0 Then
        'If ProcessSheet.Cells(i, NumColProcess("PROCESS")).Value = "Welding" Then
            WeldingSheet.Cells(iCounter, NumColWelding("Referencia")).Value = "'" & CStr(ProcessSheet.Cells(i, NumColProcess("Reference")).Value)
            'WeldingSheet.Cells(iCounter, NumColWelding("Referencia")) = "'" & ProcessSheet.Cells(i, NumColProcess("Reference")).Value
            'Agregar valor de la columna F (Capacidad) obtenido con VLOOKUP'
            WeldingSheet.Cells(iCounter + 3, NumColWelding("Referencia")).Value = Application.WorksheetFunction.VLookup(ProcessSheet.Cells(i, 1).Value, ProcessSheet.Range("A:F"), 6, False)
            
            'Agregar valor ID obtenido con VLOOKUP'
            WeldingSheet.Cells(iCounter, NumColWelding("ID")).Value = Application.WorksheetFunction.VLookup(ProcessSheet.Cells(i, 1).Value, ProcessSheet.Range("A:F"), 2, False)
            WeldingSheet.Cells(iCounter + 3, NumColWelding("Capacidad")).Value = "Capacidad/t"
            'WeldingSheet.Cells(iCounter + 3, NumColWelding("Capacidad")).Columns.AutoFit
            
            'Agregar línea'
            WeldingSheet.Cells(iCounter, NumColWelding("Linea")).Value = Application.WorksheetFunction.VLookup(ProcessSheet.Cells(i, 1).Value, ProcessSheet.Range("A:F"), 4, False)
            
            'Formato a celdas'
            Set TargetRange = WeldingSheet.Cells(iCounter, NumColWelding("Line"))
            formatRange1.Copy
            TargetRange.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
            


            iCounter = iCounter + WeldingRowDistance()
            Else
        End If
    Next i
    Application.DisplayAlerts = False
End Sub
