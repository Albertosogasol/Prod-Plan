Attribute VB_Name = "Module70"
'MODULE_70'
'Módulo para la creación de la pestaña process'
'Anteriormente se hacia manualemente, a través de este módulo de construye sola'
Sub createProcTable(Row As Integer, name As String)
    'Creación de tabla en la pestaña Process pasando por argumento la fila y el nombre de la tabla
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("Process"))
    
    Dim table As ListObject
    Set table = ws.ListObjects.Add(xlSrcRange, ws.Range("A" & Row + 1 & ":I" & Row + 1))
    table.name = name

    'Colocación de headers'
    With table.HeaderRowRange
        .Cells(1, 1).Value = "Reference"
        .Cells(1, 2).Value = "ID"
        .Cells(1, 3).Value = "Process"
        .Cells(1, 4).Value = "Line"
        .Cells(1, 5).Value = "Project"
        .Cells(1, 6).Value = "Quantity"
        .Cells(1, 7).Value = "Comments"
        .Cells(1, 8).Value = "Is_next"
        .Cells(1, 9).Value = "Checked"
    End With

End Sub

Sub resetProcSheet()
    'Borra y actualiza toda la pestaña de procesos'
    Dim RefSheet As Worksheet
    Dim procSheet As Worksheet
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))
    Set procSheet = ThisWorkbook.Worksheets(SheetName("PROCESS"))

    'Variables'
    Dim chkWelding As Boolean 'Default: FALSE
    Dim chkBox As Boolean 'Default: FALSE
    Dim chkBending As Boolean 'Default: FALSE
    Dim chkFinal As Boolean 'Default: FALSE
    Dim lastRow As Integer
    Dim table As ListObject
    Dim procNumericTable As Integer 'Relaciona el valor entero de la tabla de la pestaña procesos con su nombre en cadena. Se utiliza para evitar cambiar valores en un futuro
    Dim welding_ID, box_ID, bending_ID, final_ID As Integer 'Contadores para ID de tablas de procesos'
    welding_ID = 1
    box_ID = 1
    bending_ID = 1
    final_ID = 1
    procSheet.UsedRange.Clear

    'Comprobación de 4 tablas de procesos: WELDING, BOX, BENDING, FINAL
    For i = 1 To procSheet.ListObjects.Count
        If (UCase(procSheet.ListObjects(i).name) = "WELDING") Then
            chkWelding = True
        ElseIf (UCase(procSheet.ListObjects(i).name) = "BOX") Then
            chkBox = True
        ElseIf (UCase(procSheet.ListObjects(i).name) = "BENDING") Then
            chkBending = True
        ElseIf (UCase(procSheet.ListObjects(i).name) = "FINAL") Then
            chkFinal = True
        End If
    Next i

    'Si no existe alguna de las tablas, se crea'
    If chkWelding = False Then
        lastRow = procSheet.Cells(Rows.Count, 1).End(xlUp).Row
        Call createProcTable(lastRow, "Welding")
    End If
    If chkBox = False Then
        lastRow = procSheet.Cells(Rows.Count, 1).End(xlUp).Row
        Call createProcTable(lastRow, "Box")
    End If
    If chkBending = False Then
        lastRow = procSheet.Cells(Rows.Count, 1).End(xlUp).Row
        Call createProcTable(lastRow, "Bending")
    End If
    If chkFinal = False Then
        lastRow = procSheet.Cells(Rows.Count, 1).End(xlUp).Row
        Call createProcTable(lastRow, "Final")
    End If

    'Comprobación de referencias desde pestaña REFERENCES'
    'Primero se verifica que las actuales son almacenadas como cadenas'
    Call ChkProcStr   'Module Functions'
    Call ChkRefStr   'Module Functions'

    'Se recorren todas las celdas de cada tabla de la pestaña REFERENCES
    'Para cada valor, se realiza una búsqueda en la tabla correspondiente en la pestaña PROCESS _
    ' y si no está, se añade.

    'MÉTODO BORRAR Y CONSTUIR TODO'

    For i = 1 To RefSheet.ListObjects.Count
        For j = 1 To RefSheet.ListObjects(i).Range.Rows.Count
            If RefSheet.ListObjects(i).Range.Cells(j, 3) = "WELDING" Then
                procNumericTable = 1
                procSheet.ListObjects(procNumericTable).ListRows.Add

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Referencia")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Reference") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("ID")) = welding_ID
                
                welding_ID = welding_ID + 1

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Process")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Process") - 1)
                
                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Line")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Line") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Quantity")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Quantity") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Comments")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Comments") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("CHK")) = True
            
            ElseIf RefSheet.ListObjects(i).Range.Cells(j, 3) = "BOX" Then
                procNumericTable = 2
                procSheet.ListObjects(procNumericTable).ListRows.Add

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Referencia")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Reference") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("ID")) = box_ID
                
                box_ID = box_ID + 1

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Process")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Process") - 1)
                
                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Line")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Line") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Quantity")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Quantity") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Comments")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Comments") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("CHK")) = True
            ElseIf RefSheet.ListObjects(i).Range.Cells(j, 3) = "BENDING" Then
                procNumericTable = 3
                procSheet.ListObjects(procNumericTable).ListRows.Add

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Referencia")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Reference") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("ID")) = bending_ID
                
                bending_ID = bending_ID + 1

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Process")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Process") - 1)
                
                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Line")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Line") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Quantity")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Quantity") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Comments")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Comments") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("CHK")) = True
            ElseIf RefSheet.ListObjects(i).Range.Cells(j, 3) = "FINAL" Then
                procNumericTable = 4
                procSheet.ListObjects(procNumericTable).ListRows.Add

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Referencia")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Reference") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("ID")) = final_ID
                
                final_ID = final_ID + 1

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Process")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Process") - 1)
                
                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Line")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Line") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Quantity")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Quantity") - 1)

                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("Comments")) = RefSheet.ListObjects(i).Range.Cells(j, NumColReference("Comments") - 1)
            
                procSheet.ListObjects(procNumericTable).Range.Cells(procSheet.ListObjects(procNumericTable).Range.Rows.Count - 1, ProcessCol("CHK")) = True
            End If
        Next j
    Next i
    'Set myNewRow = ActiveWorkbook.Worksheets(1).ListObject(1).ListRows.Add
End Sub

Sub updateProcSheet()
    'Actualiza la pestaña de procesos sin borrar contenido'
    Dim RefSheet As Worksheet
    Dim procSheet As Worksheet
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))
    Set procSheet = ThisWorkbook.Worksheets(SheetName("PROCESS"))

    'Variables'
    Dim refMatch As String
    Dim refMatchRow As Integer
    Dim lastRowProc As Integer

    'Check en FASLE. Comprobante de referencia actualizada'
    For i = 1 To procSheet.ListObjects.Count
        procSheet.ListObjects(i).ListColumns(9).DataBodyRange = False
    Next i

    'Comprobación de cada referencia de pestaña REFERENCES en pestaña de procesos'
    For i = 1 To RefSheet.ListObjects.Count 'Se recorren todas las tablas de References'
            For j = 2 To RefSheet.ListObjects(i).Range.Rows.Count
                refMatch = RefSheet.ListObjects(i).Range.Cells(j, 1)
                If Not IsError(Application.match(refMatch, procSheet.Columns(1), 0)) Then
                    refMatchRow = Application.match(refMatch, procSheet.Columns(1), 0)
                    procSheet.Cells(refMatchRow, NumColProcess("CHK")) = True 'Se resta -1 porque las tablas de numColProcess empiezan en la columna 2
                End If
            Next j
    Next i

    lastRowProc = procSheet.Cells(Rows.Count, NumColProcess("REFERENCE")).End(xlUp).Row
    For i = 1 To lastRowProc
        If procSheet.Cells(i, NumColProcess("CHK")) = False Then
            Rows(i).EntireRow.Delete
            'Cells(i, NumColProcess("CHK")).EntireRow.Delete
        End If
    Next i

End Sub

