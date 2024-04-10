Attribute VB_Name = "Module40"
'MODULE_40'
'VERIFICATION SHEET MODULE'

Sub VerificationHeaders()
    'Creación de las cabeceras generales'
    Dim VerSheet As Worksheet
    Dim formatsSheet As Worksheet
    Set VerSheet = ThisWorkbook.Worksheets(SheetName("VERIFICATION"))
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))

    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A83:E83")
    FormatRange.Copy

    VerSheet.Cells(OffsetFilaCabecera(), NumColVer("FINAL_REF")).PasteSpecial
    Application.CutCopyMode = False

End Sub

Sub VerificationRefs()
    'Referencias obtenidas de la pestaña REFERENCES, teniendo en cuenta los niveles'
    Dim VerSheet As Worksheet
    Set VerSheet = ThisWorkbook.Worksheets(SheetName("Verification"))
    
    Dim RefSheet As Worksheet
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("References"))
    
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets("WELDING")

    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))
    
    Dim Reference As String
    Dim lastRowWelding As Integer
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row

    
    'Rango para aislar la tabla con la que se trabaja de la pestaña REFERENCES'
    Dim tableRange As Range
    Dim firstCol As Range
    Dim refColOffset As Integer 'Numero de columnas a la izquierda en la pestaña REFERENCES'
    refColOffset = 1
    
    'Variable última fila en uso en pestaña VERIFICATION'
    Dim lastRowVer As Integer
    lastRowVer = VerSheet.Cells(Rows.Count, NumColVer("FINAL_REF")).End(xlUp).Row
    Dim iCounter As Integer
    iCounter = lastRowVer + 1
    
    For i = OffsetFilaCabecera() + 1 To lastRowWelding Step WeldingRowDistance()
        Reference = WeldingSheet.Cells(i, NumColWelding("Reference")).Value
        On Error Resume Next
        Set tableRange = RefSheet.ListObjects("Table_" & Reference).Range
        
        'Obtención de datos desde tabla aislada'
        For Row = 2 To tableRange.Rows.Count
            If tableRange.Cells(Row, NumColReference("Level") - refColOffset).Value < 0 Then
            Else
                VerSheet.Cells(iCounter, NumColVer("Reference")).Value = tableRange.Cells(Row, NumColReference("Reference") - refColOffset)
                VerSheet.Cells(iCounter, NumColVer("Final_Ref")).Value = tableRange.Cells(Row, NumColReference("Final_ref") - refColOffset)
                VerSheet.Cells(iCounter, NumColVer("Level")).Value = tableRange.Cells(Row, NumColReference("Level") - refColOffset)
                VerSheet.Cells(iCounter, NumColVer("ID")).Value = tableRange.Cells(Row, NumColReference("ID") - refColOffset)
                VerSheet.Cells(iCounter, NumColVer("Process")).Value = tableRange.Cells(Row, NumColReference("Process") - refColOffset)
                iCounter = iCounter + 1
            End If
            
        Next Row
        'Limpiar rango'
        Set tableRange = VerSheet.Range("A1:A1") 'Se hace un rango unitario para no acumular'
        Application.CutCopyMode = False
        
    Next i

    'Formato a celdas'
    Dim FormatRange As Range
    Dim lastRowVerFormats As Integer
    lastRowVerFormats = VerSheet.Cells(Rows.Count, NumColVer("FINAL_REF")).End(xlUp).Row
    For i = OffsetFilaCabecera() + 1 To lastRowVerFormats
        If VerSheet.Cells(i, NumColVer("ID")).Value = 0 Then
            Set FormatRange = formatsSheet.Range("A84:E84")
            FormatRange.Copy
            VerSheet.Cells(i, NumColVer("FINALREF")).PasteSpecial xlPasteFormats
            'MsgBox("Primera condición para " & VerSheet.Cells(i,NumColVer("REFERENCE")) & " en el nivel " & VerSheet.Cells(i,NumColVer("LVL")))
        ElseIf VerSheet.Cells(i + 1, NumColVer("ID")).Value = 0 Then
            Set FormatRange = formatsSheet.Range("A88:E88")
            FormatRange.Copy
            VerSheet.Cells(i, NumColVer("FINALREF")).PasteSpecial xlPasteFormats
            'MsgBox("Segunda condición para " & VerSheet.Cells(i,NumColVer("REFERENCE")) & " en el nivel " & VerSheet.Cells(i,NumColVer("LVL")))
        Else
            Set FormatRange = formatsSheet.Range("A85:E85")
            FormatRange.Copy
            VerSheet.Cells(i, NumColVer("FINALREF")).PasteSpecial xlPasteFormats
            'MsgBox("Tercera condición para " & VerSheet.Cells(i,NumColVer("REFERENCE")) & " en el nivel " & VerSheet.Cells(i,NumColVer("LVL")))
        End If
        FormatRange.ClearContents
        Application.CutCopyMode = False
    Next i
End Sub

Sub AddVerificationWeekHeaders(week As Integer, WeekCol As Integer)
    'Dimensionado de hojas de trabajo'
    Dim VerSheet As Worksheet
    Set VerSheet = ThisWorkbook.Worksheets(SheetName("VERIFICATION"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    'N,D, T para turno del día'
    For i = 0 To WeekShifts() - 1 Step 3
        VerSheet.Cells(OffsetFilaCabecera(), WeekCol + i) = "N"
        VerSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 1) = "D"
        VerSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 2) = "T"
    Next i
    
    'Fecha encima de cada día de la semana'
    Dim Counter As Integer
    Counter = 1
    For i = 0 To WeekShifts() - 1 Step 3
        VerSheet.Cells(OffsetFilaCabecera() - 1, WeekCol + i).Value = GetDate(week, Counter)
        Counter = Counter + 1
    Next i
    
    'Número de la semana'
    VerSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).Value = "Week " & week
    
    'Copiar formatos de celda desde la pestaña FORMATS'
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A66:R68")
    FormatRange.Copy
    VerSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).PasteSpecial xlPasteFormats
    
End Sub

Sub VerificationWeekBody(week As Integer)
    Dim WeldingSheet As Worksheet
    Dim BoxSheet As Worksheet
    Dim BendingSheet As Worksheet
    Dim VerSheet As Worksheet
    Dim RefSheet As Worksheet
    Dim formatsSheet As Worksheet

    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    Set VerSheet = ThisWorkbook.Worksheets(SheetName("Verification"))
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("References"))
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))

    'Variables generales'
    Dim lastRowWelding As Integer 'Última fila en uso en pestaña WELDING'
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row
    Dim lastRowVer As Integer 'Última fila en uso en pestaña VERIFICATION'
    lastRowVer = VerSheet.Cells(Rows.Count, NumColVer("REFERENCE")).End(xlUp).Row
    Dim refTableRange As Range 'Rango para aislar tablas de pestaña REFERENCES'
    Dim verTableRange As Range 'Rango para aislar tablas de pestaña VERIFICATION'
    Dim lastRowTableRange As Integer 'Última fila de la tabla de REFERENCES'
    Dim maxLevel As Integer 'Máximo nivel de la referencia con la que se trabaja'
    Dim refOffset As Integer 'Número de columnas a la izquierda en la pestaña REFERENCES'
    refOffset = ReferencesOffset()
    Dim maxID As Integer 'Número de procesos en una referencia en la tabla correspondiente'
    Dim tempRefRow As Integer 'Fila en la que se encuentra la referencia leída'
    Dim tempRefStatus As Boolean 'Verificación de stock de la referencia con la que se trabaja'
    Dim tempNextRefRow As Integer 'Posición de la siguiente referencia en la tabla de la pestaña VERIFICATION
    Dim iShift As Integer
    Dim tempRefStatusRow As Integer
    Dim tempRefStatusCol As Integer
    Dim currentLevel As Integer
    Dim whileBreak As Boolean
    Dim whileRow As Integer

    'Variables del bucle general'
    Dim weldingReference As String 'Referencia leída en WELDING'
    Dim verFirstMatch As Integer 'Primera fila donde se encuentra la coincidendia'
    Dim verLastMatch As Integer 'Última fila perteneciente al rango de referencias'

    'Se limpia el rango para realizar los cálculos desde 0'
    Dim completeRange As Range
    Set completeRange = VerSheet.Range(VerSheet.Cells(OffsetFilaCabecera() + 1, VerWeekSearch(week)), VerSheet.Cells(lastRowVer, VerWeekSearch(week) + WeekShifts() - 1))
    completeRange.ClearContents
    
    'Bucle principal' 'Se recorren todas las referencias de la pestaña WELDING'
    For weldingRow = OffsetFilaCabecera() + 1 To lastRowWelding Step WeldingRowDistance()
        ' Lectura de referencia en pestaña WELDING
        weldingReference = WeldingSheet.Cells(weldingRow, NumColWelding("Reference")).Value

        'On Error Resume Next
        ' Búsqueda de primera coincidencia de weldingReference en pestaña VERIFICATION'
        'verFirstMatch = Application.match(weldingReference, VerSheet.Columns(NumColVer("FINAL_REF")), 0)
    
        If Not IsError(Application.match(weldingReference, VerSheet.Columns(NumColVer("FINAL_REF")), 0)) Then ' Si se encuentra una coincidencia
            verFirstMatch = Application.match(weldingReference, VerSheet.Columns(NumColVer("FINAL_REF")), 0)
            ' Creación de rango desde la tabla de la pestaña REFERENCES
            'MsgBox (" El error en la tabla: " & "Table_" & weldingReference)
            Set refTable = RefSheet.ListObjects("Table_" & weldingReference).Range
    
            ' Máximo nivel en la referencia con la que se trabaja
            maxLevel = refTable.Cells(refTable.Rows.Count, NumColReference("Level") - refOffset).Value
    
            ' Búsqueda del número de ID's en la tabla correspondiente para dimensionar el rango
            maxID = refTable.Cells(refTable.Rows.Count, NumColReference("ID") - refOffset)
    
            ' Última fila perteneciente a una referencia en la pestaña VERIFICATION
            verLastMatch = verFirstMatch + maxID '- 1 ' Se quita 1 por la cabecera. No se quitan 2 ya que las ID comienzan en 0'
    
            'MsgBox "La referencia " & weldingReference & " empieza en la fila " & verFirstMatch & " y acaba en la fila " & verLastMatch

            'Bucle para recorrer cada referencia de la tabla actual'
            For tempRefRow = verLastMatch To verFirstMatch Step -1
                'Bucle para recorrer cada turno'
                For Shift = 0 To WeekShifts() - 1
                    iShift = Shift
                    If (VerSheet.Cells(tempRefRow, NumColVer("LEVEL"))) = maxLevel Then
                        'Se comprueba el estado del stock de la referencia pasada como argumento'
                        tempRefStatus = checkStock(iShift, week, VerSheet.Cells(tempRefRow, NumColVer("PROCESS")), VerSheet.Cells(tempRefRow, NumColVer("REFERENCE")))
                        If tempRefStatus = True Then
                            VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift).Value = "OK"
                            'MsgBox "OK En la celda: " & tempRefRow & " " & VerWeekSearch(Week) + iShift
                        Else
                            On Error Resume Next
                            VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift).Value = "NOK"
                            nextRef = findNextRef(VerSheet.Cells(tempRefRow, NumColVer("REFERENCE")), weldingReference)
                            'MsgBox "La siguiente referencia es:" & nextRef
                            Set verTableRange = VerSheet.Range(VerSheet.Cells(verFirstMatch, NumColVer("FINAL_REF")), VerSheet.Cells(verLastMatch, NumColVer("REFERENCE")))
                            'MsgBox "Se va a buscar la ref: " & nextRef
                            tempNextRefRow = Application.WorksheetFunction.match(nextRef, verTableRange.Columns(5), 0)
                            'MsgBox "La siguiente referencia (" & nextRef & ") a la referencia " & VerSheet.Cells(tempRefRow, NumColVer("REFERENCE")).Value & " está en la fila " & tempNextRefRow
                            VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift + (VerSheet.Cells(tempRefRow, NumColVer("LEVEL")) - VerSheet.Cells(tempNextRefRow, NumColVer("LEVEL")))).Value = "NOK"
                            VerSheet.Cells(verFirstMatch + tempNextRefRow - 1, VerWeekSearch(week) + iShift).Value = "NOK"
                            'MsgBox "NOK En la celda: " & tempRefRow & " " & VerWeekSearch(Week) + iShift + (VerSheet.Cells(tempRefRow, NumColVer("LEVEL")) - VerSheet.Cells(tempNextRefRow, NumColVer("LEVEL")))
                            On Error GoTo 0
                        End If
                    Else
                        If (VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift).Value = "NOK") Then
                            'En principio no se ejecuta ninguna acción'
                            'REVISAR SI ES NECESARIO AÑADIR'
                        Else
                            'Se comprueba si hay stock'
                            tempRefStatus = checkStock(iShift, week, VerSheet.Cells(tempRefRow, NumColVer("PROCESS")), VerSheet.Cells(tempRefRow, NumColVer("REFERENCE")))
                            If (tempRefStatus = True) Then
                                VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift).Value = "OK" 'Si hay stock se coloca en OK'
                                whileBreak = False
                                whileRow = tempRefRow
                                'SI EXISTE STOCK HAY QUE COMPROBAR SI HAY STOCK DE LAS REFERENCIAS DEL NIVEL N-1
                                'Bucle para comprobar el stock de referencias de nivel n-1
                                While whileBreak = False
                                    If VerSheet.Cells(whileRow + 1, NumColVer("LEVEL")).Value = VerSheet.Cells(tempRefRow, NumColVer("LEVEL")).Value Then 'Cambio de  por tempRefRow
                                        whileRow = whileRow + 1
                                    ElseIf (VerSheet.Cells(whileRow + 1, VerWeekSearch(week) + iShift - 1).Value = "NOK") Then
                                        VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift).Value = "NOK"
                                        whileBreak = True
                                    ElseIf (VerSheet.Cells(whileRow, NumColVer("FINAL_REF")).Value <> VerSheet.Cells(tempRefRow, NumColVer("FINAL_REF")).Value) Then
                                        whileBreak = True
                                    Else
                                        whileRow = whileRow + 1
                                    End If
                                Wend
                            Else
                                VerSheet.Cells(tempRefRow, VerWeekSearch(week) + iShift).Value = "NOK"
                            End If
                        End If
                    End If
                Next Shift
            Set verTableRange = VerSheet.Range("A1:A1")
            Next tempRefRow
        Else
            'MsgBox "La referencia " & weldingReference & " no se encontró en la pestaña VERIFICATION"
        End If

        'On Error GoTo 0
    Next weldingRow
End Sub

Function checkStock(Shift As Integer, week As Integer, Process As String, Reference As String) As Boolean
    'Devuelve en booleano si una referencia tiene falta de stock en la semana y turno pasada por argumento.
    'Necesario pasar el tipo de tecnología para apuntar a una pestaña
    Dim ws As Worksheet
    Dim status As Boolean
    status = True
    Dim Row As Integer
    Dim col As Integer
    Dim stockLimit As Integer 'Variable para almacenar el límite de stock'
    stockLimit = 0 'Se coloca arbitrariamente a 0'
    Process = UCase(Process)
    
    If Process = "WELDING" Then
        Set ws = ThisWorkbook.Worksheets(SheetName(Process))
        Row = WeldingReferenceRow(Reference) + 1
        col = WeldingWeekSearch(week) + 4 + Shift
        If (ws.Cells(Row, col).Value) > stockLimit Then
            status = True
        Else
            status = False
        End If
    ElseIf Process = "BOX" Then
        Set ws = ThisWorkbook.Worksheets(SheetName(Process))
        Row = BoxReferenceRow(Reference) + 1
        col = BoxWeekSearch(week) + Shift
        If (ws.Cells(Row, col).Value) > stockLimit Then
            status = True
        Else
            status = False
        End If
    ElseIf Process = "BENDING" Then
        'MsgBox "Entra en el bucle de BENDING"
        Set ws = ThisWorkbook.Worksheets(SheetName(Process))
        Row = BendingReferenceRow(Reference) + 1
        col = BendingWeekSearch(week) + Shift
        'MsgBox "La referencia " & Reference & " en la semana " & Week & " en el turno " & Shift & " tiene " & ws.Cells(Row, col).Value
        If (ws.Cells(Row, col).Value) > stockLimit Then
            status = True
            'MsgBox "Status TRUE para " & Reference
        Else
            status = False
        End If
    Else
        'MsgBox "ERROR EN EL BUCLE checkStock al pasar la referencia: " & Reference & " con el proceso: " & Process
        'On Error GoTo Error
    End If
    checkStock = status
End Function

Function VerWeekSearch(weekNumber As Integer) As Integer
    'DEVUELVE LA COLUMNA (COMO LONG) EN LA QUE COMIENZA LA SEMANA PASADA COMO ARGUMENTO EN LA PESTAÑA VERIFICATION'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("VERIFICATION"))
    
    Dim week As String
    week = "Week " & weekNumber
    Dim cell As Range
    Set cell = ws.Rows(OffsetFilaCabecera() - 2).Find(What:=week, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        VerWeekSearch = 0
        MsgBox "No se ha encontrado ninguna semana mediante la función VerWeekSearch"
    Else
        VerWeekSearch = cell.Column
    End If
End Function

Function findNextRef(Reference As String, FinalRef As String) As String
    'Busca la referencia siguiente a la pasada por argumento en la pestaña REFERENCES'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("REFERENCES"))

    Dim table As Range
    Dim foundRef As Range

    Set table = ws.ListObjects("Table_" & FinalRef).Range
    Set foundRef = table.Find(What:=Reference, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundRef Is Nothing Then
        'MsgBox "La siguiente referencia es: " & foundRef.offset(0, 5).Value
    Else
        MsgBox "No se ha encontrado ninguna referencia en el rango especificado"
    End If
    findNextRef = foundRef.offset(0, 5).Value
End Function

Sub VerificationWeekHeaderBuilder(week As Integer)
    'Se colocan las cabeceras de la semana pasada por argumento. Principalmente desde el UserForm principal
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("VERIFICATION"))
    'Limpieza de registro anterior'
    Dim lastRow As Integer
    lastRow = ws.Cells(Rows.Count, NumColVer("REFERENCE")).End(xlUp).Row

    Dim completeRange As Range
    Set completeRange = ws.Range(ws.Cells(OffsetFilaCabecera() - 2, NumColVer("REFERENCE") + 1), ws.Cells(lastRow, NumColVer("REFERENCE") + 100))
    completeRange.ClearContents
    
    Call AddVerificationWeekHeaders(week - 1, NumColVer("REFERENCE") + 1)
    Dim lastCol As Integer
    lastCol = ws.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Call AddVerificationWeekHeaders(week, lastCol + 1)
    Application.Wait (Now + TimeValue("0:00:1"))
    lastCol = ws.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Call AddVerificationWeekHeaders(week + 1, lastCol + 1)
    Application.Wait (Now + TimeValue("0:00:1"))
    lastCol = ws.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Call AddVerificationWeekHeaders(week + 2, lastCol + 1)

    MsgBox ("Finalizado con éxito")
End Sub

Sub Verification()
    'Ejecuta la verificación principal
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("VERIFICATION"))

    Dim text As String
    Dim week As Integer
    
    'Se lee el valor de la semana'
    'Se lee la primera semana en aparecer, pero se está trabajando con la week + 1
    text = ws.Cells(OffsetFilaCabecera() - 2, NumColVer("REFERENCE") + 1).Value

    'Se busca la posición del espacio'
    Dim espace As Integer
    espace = InStr(text, " ")

    'Se extrae la parte posterior al espacio'
    text = Right(text, Len(text) - espace)

    'Se convierte el String a Integer
    week = CInt(text)

    'Se ejecuta para la semana central (week + 1) y las dos siguientes
    VerificationWeekBody (week + 1)
    VerificationWeekBody (week + 2)
    'VerificationWeekBody(week + 3)

    MsgBox ("Finalizado con éxito")
End Sub

'SUBRUTINA DE PRUEBA'
Sub pruebaSubrutina()
    'VerificationWeekBody (2)
    'VerificationWeekBody (3)
    VerificationWeekHeaderBuilder (3)
End Sub

'Subrutina de prueba de funciones'
Sub pruebaFunction()
    Dim var As Variant
    var = findNextRef("10082188", "2G6253181AT")
End Sub
