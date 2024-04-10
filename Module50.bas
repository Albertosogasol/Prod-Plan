Attribute VB_Name = "Module50"
'MODULE_50'
'SOLVER SHEET'

Sub SolverHeaders()
    'Cabeceras generales de la pestaña SOLVER'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("SOLVER_SHEET"))
    Dim FormatSheet As Worksheet
    Set FormatSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))

    ws.Cells(OffsetFilaCabecera(), NumColSolver("Process")).Value = "Proceso"
    ws.Cells(OffsetFilaCabecera(), NumColSolver("Linea")).Value = "Línea"
    ws.Cells(OffsetFilaCabecera(), NumColSolver("Referencia")).Value = "Referencia"
    ws.Cells(OffsetFilaCabecera(), NumColSolver("Pers")).Value = "Pers/Turno"
    ws.Cells(OffsetFilaCabecera(), NumColSolver("Pz")).Value = "Pz/Turno"

    'Formatos de celdas
    Dim FormatRange As Range
    Set FormatRange = FormatSheet.Range("A92:E92")
    FormatRange.Copy

    ws.Range(ws.Cells(OffsetFilaCabecera(), NumColSolver("Process")), ws.Cells(OffsetFilaCabecera(), NumColSolver("PZ"))).PasteSpecial xlPasteFormats

End Sub

Sub SolverReferences()
    'Referencias en la pestaña solver'
    Dim SolvSheet As Worksheet
    Set SolvSheet = ThisWorkbook.Worksheets(SheetName("SOLVER"))
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    Dim RefSheet As Worksheet
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))
    Dim procSheet As Worksheet
    Set procSheet = ThisWorkbook.Worksheets(SheetName("PROCESS"))

    'Variables generales'
    Dim lastRowProc As Integer
    lastRowProc = procSheet.Cells(Rows.Count, NumColProcess("References")).End(xlUp).Row
    Dim iCounter As Integer 'Contador para las posiciones de las referencias en la pestaña SOLVER'
    iCounter = OffsetFilaCabecera() + 1
    For i = 1 To lastRowProc
        If procSheet.Cells(i, NumColProcess("REFERENCES")).Value = "Reference" Then
        ElseIf procSheet.Cells(i, NumColProcess("REFERENCES")).Value = "" Then
        Else
            'Si la celda de la referencia no está vacía o no es cabecera de tabla, se continua con el bucle'
            SolvSheet.Cells(iCounter, NumColSolver("Proceso")).Value = procSheet.Cells(i, NumColProcess("Process")).Value
            SolvSheet.Cells(iCounter, NumColSolver("Linea")).Value = procSheet.Cells(i, NumColProcess("Line")).Value
            SolvSheet.Cells(iCounter, NumColSolver("Reference")).Value = "'" & procSheet.Cells(i, NumColProcess("Reference")).Value
            On Error Resume Next
            SolvSheet.Cells(iCounter, NumColSolver("Personas")).Value = RefSheet.Cells(Application.match(procSheet.Cells(i, NumColProcess("Reference")).Value, RefSheet.Columns(NumColReference("Reference")), 0), NumColReference("OP")).Value
            SolvSheet.Cells(iCounter, NumColSolver("Cantidad")).Value = RefSheet.Cells(Application.match(procSheet.Cells(i, NumColProcess("Reference")).Value, RefSheet.Columns(NumColReference("Reference")), 0), NumColReference("Cantidad")).Value
            On Error GoTo 0
            If procSheet.Cells(i, NumColProcess("Line")).Value <> procSheet.Cells(i + 1, NumColProcess("Line")) Then
                iCounter = iCounter + 1
                SolvSheet.Cells(iCounter, NumColSolver("Proceso")).Value = "SUM"
                iCounter = iCounter + 1
            Else
            
            End If
            iCounter = iCounter + 1
        End If
    Next i
End Sub

Sub SolverWeekHeaders(CurrentWeek As Integer)
    'Añade las cabeceras de la semana pasada como argumento en la pestaña SOLVER'
    Dim SolvSheet As Worksheet
    Set SolvSheet = ThisWorkbook.Worksheets(SheetName("SOLVER"))
    Dim FormatSheet As Worksheet
    Set FormatSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))

    Dim firstData As Integer 'Primera columna con información en la pestaña SOLVER'
    firstData = NumColSolver("PZ") + 1

    'Bucle para la construcción de Prod. | N | Prod. | D | Prod. | N |
    'Se repite para cada día de la semana de L - S
    Dim weekDayColStart As Integer
    For wD = 1 To 6
        weekDayColStart = firstData * wD
        SolvSheet.Cells(OffsetFilaCabecera(), weekDayColStart).Value = "Prod."
        SolvSheet.Cells(OffsetFilaCabecera(), weekDayColStart + 1).Value = "N"
        SolvSheet.Cells(OffsetFilaCabecera(), weekDayColStart + 2).Value = "Prod."
        SolvSheet.Cells(OffsetFilaCabecera(), weekDayColStart + 3).Value = "D"
        SolvSheet.Cells(OffsetFilaCabecera(), weekDayColStart + 4).Value = "Prod."
        SolvSheet.Cells(OffsetFilaCabecera(), weekDayColStart + 5).Value = "T"
    Next wD

    'Bucle para fechas de la semana pasada por argumento'
    Dim weekDayCounter As Integer
    weekDayCounter = 1
    For wD = firstData To (firstData * 6) Step 6
        SolvSheet.Cells(OffsetFilaCabecera() - 1, wD).Value = GetDate(CurrentWeek, weekDayCounter)
        weekDayCounter = weekDayCounter + 1
    Next wD
    
    'Número de semana'
    SolvSheet.Cells(OffsetFilaCabecera() - 2, firstData).Value = "Week " & CurrentWeek

    'Formato a celdas'
    Dim FormatRange As Range
    Set FormatRange = FormatSheet.Range("F110:AO112")
    FormatRange.Copy
    SolvSheet.Cells(OffsetFilaCabecera() - 2, firstData).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Sub SolverWeekBody(CurrentWeek As Integer)
    'Estructura del cuerpo de la pestaña Solver'
    Dim SolvSheet As Worksheet
    Set SolvSheet = ThisWorkbook.Worksheets(SheetName("SOLVER"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))

    'Variables para los sumatorios'
    Dim lineBlock As Integer
    lineBlock = 0 'Contador de bloques de grupos de líneas. Se utiliza para la identificación del primer bloque, en el que el sumatorio se hace desde la OffsetFilaCabecera() + 1 hasta el primer SUM. A partir de ahi se hacen diferentes
    Dim startCell, finalCell As Range
    Dim lastSumCell As Integer 'Ultima celda en la que se encontraba un SUM
    Dim languageID As Integer
    languageID = Application.LanguageSettings.languageID(msoLanguageIDUI)
    Dim formulaString As String

    'Variables generales de los bucles'
    Dim lastRowSolver As Integer
    lastRowSolver = SolvSheet.Cells(Rows.Count, NumColSolver("Proceso")).End(xlUp).Row
    Dim lastColSolver As Integer
    lastColSolver = SolvSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    Dim tempRef As String 'Referencia leída en el bucle principal
    Dim tempRefRow As Integer 'Fila de la referencia leía en su pestaña correspondiente
    Dim firstData As Integer 'Primera columna con datos en pestaña SOLVER'
    firstData = NumColSolver("PZ") + 1
    Dim jShift As Integer 'Número de turno en el bucle principal'

    For Row = OffsetFilaCabecera() + 1 To lastRowSolver
        tempRef = SolvSheet.Cells(Row, NumColSolver("REFERENCE")).Value
        'Se comprueba si se está leyendo una referencia o una celda en blanco'
        If tempRef <> "" Then
            'Bucle principal para recorrer todas las columnas de la semana'
            jShift = 1
            For col = firstData + 1 To lastColSolver Step 2
                'Se obtiene la producción de la pestaña correspondiente'
                SolvSheet.Cells(Row, col) = "=" & getProdFormula(tempRef, SolvSheet.Cells(Row, NumColSolver("Process")).Value, CurrentWeek, jShift)
                If SolvSheet.Cells(Row, col) = "=" Then
                    SolvSheet.Cells(Row, col) = ""
                Else
                End If
                jShift = jShift + 1
            Next col
        ElseIf SolvSheet.Cells(Row, NumColSolver("Process")).Value = "SUM" Then
            lineBlock = lineBlock + 1
            If lineBlock = 1 Then 'Caso de primer bloque de referencias
                For col = firstData To lastColSolver
                    Set startCell = SolvSheet.Cells(OffsetFilaCabecera() + 1, col)
                    Set finalCell = SolvSheet.Cells(Row - 1, col)
                    'Se escribe la fórmula en un idioma u otro dependiendo del LanguageID
                    Select Case languageID
                        Case 3082, 1034 'Spanish'
                            formulaString = "=SUM(" & startCell.Address & ":" & finalCell.Address & ")"
                            'SolvSheet.Cells(row, col).Formula = "SUMA(" & startCell.Address & ":" & finalCell.Address & ")"
                            SolvSheet.Cells(Row, col).Formula = formulaString
                        Case Else '1033, 2057 'English'
                        formulaString = "=SUM(" & startCell.Address & ":" & finalCell.Address & ")"
                        SolvSheet.Cells(Row, col).Formula = formulaString
                        
                    End Select
                Next col
                lastSumCell = Row
            Else
                For col = firstData To lastColSolver
                    Set startCell = SolvSheet.Cells(lastSumCell + 2, col)
                    Set finalCell = SolvSheet.Cells(Row - 1, col)
                    'Se escribe la fórmula en un idioma u otro dependiendo del LanguageID
                    Select Case languageID
                        Case 3082, 1034 'Spanish'
                            formulaString = "=SUM(" & startCell.Address & ":" & finalCell.Address & ")"
                            SolvSheet.Cells(Row, col).Formula = formulaString
                        Case Else '1033, 2057 'English'
                            formulaString = "=SUM(" & startCell.Address & ":" & finalCell.Address & ")"
                            SolvSheet.Cells(Row, col).Formula = formulaString
                    End Select
                Next col
                lastSumCell = Row
            End If
        Else
        End If
    Next Row

    'Suma de celdas en cada día de la semana'
    
End Sub

Sub SolverWeekFormat()
    'Aplica el formato a las celdas de la pestaña Solver'
    Dim SolvSheet As Worksheet
    Dim formatsSheet As Worksheet
    Set SolvSheet = ThisWorkbook.Worksheets(SheetName("SOLVER"))
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))

    'Variables generales
    Dim tempRef As String 'Variable para almacenar la variable actual'
    Dim FormatRange As Range
    Dim destRange As Range
    Dim lastRowSolver As Integer
    lastRowSolver = SolvSheet.Cells(Rows.Count, NumColSolver("REFERENCE")).End(xlUp).Row
    Dim lastColSolver As Integer
    lastColSolver = SolvSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column

    'Bucle para recorrer todas las filas de la pestaña SOLVER'
    For Row = OffsetFilaCabecera() + 1 To lastRowSolver + 1
        
        'Condicionales para aplicar formato'
        If SolvSheet.Cells(Row, NumColSolver("PROCESS")) <> "SUM" Then
            If Row Mod 2 = 0 Then 'Fila par
                Set FormatRange = formatsSheet.Range("A131:AO131")
                FormatRange.Copy
                Set destRange = SolvSheet.Range(SolvSheet.Cells(Row, NumColSolver("PROCESS")), SolvSheet.Cells(Row, lastColSolver))
                destRange.PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            Else 'Fila impar
                Set FormatRange = formatsSheet.Range("A133:AO133")
                FormatRange.Copy
                Set destRange = SolvSheet.Range(SolvSheet.Cells(Row, NumColSolver("PROCESS")), SolvSheet.Cells(Row, lastColSolver))
                destRange.PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
        Else
        'Última fila de una referencia'
            Set FormatRange = formatsSheet.Range("A135:AO136")
            FormatRange.Copy
            Set destRange = SolvSheet.Range(SolvSheet.Cells(Row, NumColSolver("PROCESS")), SolvSheet.Cells(Row + 1, lastColSolver))
            destRange.PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            Row = Row + 1
        End If
    Next Row
End Sub

Function getProdFormula(Reference As String, Process As String, week As Integer, Shift As Integer) As String
    'Devuelve la formula de la referencia pasada como argumento.
    'Dependiendo del tipo de proceso, se calculará de una forma u otra.
    'Los turnos van numerados desde el 1 hasta el 18, obtenidos de WeekShifts()
    Dim ws As Worksheet
    Dim Row As Integer
    Dim col As Integer
    Dim cell As Range
    Process = UCase(Process)
    Select Case Process
        Case "WELDING"
            Set ws = ThisWorkbook.Worksheets(SheetName("WELDING"))
            On Error Resume Next
            Row = Application.match(Reference, ws.Columns(NumColWelding("REFERENCE")), 0)
            col = WeldingWeekSearch(week) + 3 + Shift 'Se suma 3 para saltarse las celdas de cargas, necesidades y plan de producción'
            On Error GoTo 0
        Case "BOX", "BOXES"
            Set ws = ThisWorkbook.Worksheets(SheetName("BOX"))
            On Error Resume Next
            Row = Application.match(Reference, ws.Columns(NumColBox("REFERENCE")), 0)
            col = BoxWeekSearch(week) - 1 + Shift 'Se resta 1 porque la función devuelve ya la columna exacta'
            On Error GoTo 0
        Case "BENDING"
            Set ws = ThisWorkbook.Worksheets(SheetName("BENDING"))
            On Error Resume Next
            Row = Application.match(Reference, ws.Columns(NumColBox("REFERENCE")), 0)
            col = BendingWeekSearch(week) - 1 + Shift 'Se resta 1 porque la función devuelve ya la columna exacta
            On Error GoTo 0
        Case Else
            MsgBox "ERROR EN LA FUNCIÓN getProdFormula del MODULE_50. No se está pasando por argumento un proceso correcto"
     End Select
     On Error Resume Next
     Set cell = ws.Cells(Row, col)
     getProdFormula = CStr(Process) & "!" & cell.Address
     On Error GoTo 0
     Exit Function
ErrorLabel:
'MsgBox "ERROR. NO SE ENCUENTRA LA REFERENCIA BUSCADA EN LA FUNCIÓN getProdFormula del MODULE_50"
End Function

Sub BuildWeekSolver()
    'Se construye la pestaña SOLVER completa'
    Dim SolverSheet As Worksheet
    Set SolverSheet = ThisWorkbook.Worksheets(SheetName("Pers_Solver"))
    Dim week As Integer
    answer = MsgBox("¿Desea actualizar la pestaña completa?", vbQuestion + vbYesNo, "Elegir opción")
    If answer = vbYes Then
        week = Application.InputBox("Introduzca la semana: ")
        SolverSheet.UsedRange.Clear
        SolverHeaders
        SolverReferences
        SolverWeekHeaders (week)
        SolverWeekBody (week)
        SolverWeekFormat
    Else
    End If
End Sub

Public Sub pruebaSolver()
    BuildWeekSolver
End Sub
