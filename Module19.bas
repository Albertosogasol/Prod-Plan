Attribute VB_Name = "Module19"
'MODULE_19'
'SEMANAS PESTAÑA BOX'
Sub AddBoxWeekHeaders(week As Integer, WeekCol As Integer)
    'Añade la semana en la pestaña BOX en la columna deseada
    'La posición se pasa como argumento a la columna correspondiente
    
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("Box"))
    
    'N,D,T para cada día de la semana'
    For i = 0 To 17 Step 3
        BoxSheet.Cells(OffsetFilaCabecera(), WeekCol + i).Value = "N"
        BoxSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 1).Value = "D"
        BoxSheet.Cells(OffsetFilaCabecera(), WeekCol + i + 2).Value = "T"
    Next i
    
    'Fecha encima de cada día de la semana'
    Dim Counter As Integer
    Counter = 1
    For i = 0 To (WeekShifts() - 1) Step 3
        BoxSheet.Cells(OffsetFilaCabecera() - 1, WeekCol + i).Value = GetDate(week, Counter)
        Counter = Counter + 1
    Next i
    
    'Número de semana'
    BoxSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).Value = "Week " & week
    
    'Copiar formatos de celda desde pestaña FORMATS'
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A28:R30")
    FormatRange.Copy
    BoxSheet.Cells(OffsetFilaCabecera() - 2, WeekCol).PasteSpecial xlPasteFormats
End Sub

Sub BoxWeekBody(week As Integer, WeekCol As Integer)
    'Añade el cuerpo a cada semana pasada por argumento
    'Columna corresponde a la primera columna de la semana
    Dim BoxSheet As Worksheet
    Dim WeldingSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    Dim ReferencesSheet As Worksheet
    Set ReferencesSheet = ThisWorkbook.Worksheets(SheetName("References"))
    
    Dim Reference As String
    Dim LastRowBox As Integer
    LastRowBox = BoxSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row
    Dim WeldingCell As Range
    Dim Shift As Integer 'Contador de turnos para bucle'
    Dim LastRowReference As Integer
    LastRowReference = ReferencesSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'Variables para almacenar las celdas de los calculos de producción'
    Dim Cell1 As Range
    Dim Cell2 As Range
    Dim Cell3 As Range
    Dim Cell4 As Range
    Dim Cell5 As Range
    
    'Variables creación de fórmula dinámica'
    Dim ArrayForm() As String
    Dim refCounter As Integer
    Dim ArrayPosCounter As Integer
    Dim tempFormula As String
    
    For i = OffsetFilaCabecera() + 1 To LastRowBox Step BoxRowDistance()
        Reference = BoxSheet.Cells(i, NumColWelding("Reference")).Value 'Se recorren todas las referencias'
        refCounter = 0 'Inicializa contador de referencias repetidas a 0'
        
        'Contador de referencias repetidas. Para dimensionar el array que contendrá las referencias finales en las que se usa esa BoxRef'
        For c = 1 To LastRowReference
            If Reference = ReferencesSheet.Cells(c, NumColReference("References")).Value Then
                refCounter = refCounter + 1
                'MsgBox "La Reference " & Reference & " es igual a " & ReferencesSheet.Cells(c, NumColReference("References")).Value & " el contador es " & refCounter
            Else
            End If
        Next c
        
        'Redimensionado del array'
        ReDim ArrayForm(refCounter)
        ArrayPosCounter = 0
        
        'Llenado del array con las referencias a buscar en la pestaña WELDING'
        For k = 1 To LastRowReference
            If Reference = ReferencesSheet.Cells(k, NumColReference("References")).Value Then
                ArrayForm(ArrayPosCounter) = ReferencesSheet.Cells(k, NumColReference("Final_Reference"))
                ArrayPosCounter = ArrayPosCounter + 1
            Else
            End If
        Next k
        
        'Búsqueda de la referencia final en la pestaña References'
        Dim FinalRef As String
        On Error Resume Next
        'FinalRef = Application.WorksheetFunction.VLookup(Reference, ReferencesSheet.Range("B:M"), NumColProcess("Final_Reference"), False)
        'MsgBox "La referencia " & Reference & " pertenece a la referencia final " & FinalRef
        
        Shift = WeldingWeekSearch(week) + StartShiftWeldingCol() 'Antes de comenzar cada semana se inicializa el contador de turnos para que apunte a la celda correcta'
        For j = WeekCol To (WeekCol + 17) 'Bucle para recorrer todos los turnos'
            Dim WeldingFormula As String 'Cadena que almacenará la fórmula'
            Set WeldingCell = WeldingSheet.Cells(WeldingReferenceRow(FinalRef), Shift)
            'BoxSheet.Cells(i, j).Formula = "=" & WeldingSheet.name & "!" & WeldingCell.Address
            tempFormula = BoxFormulaBuilder(ArrayForm(), Shift)
            'MsgBox "tempFormula " & BoxFormulaBuilder(ArrayForm(), Shift)
            BoxSheet.Cells(i, j).Formula = "=" & tempFormula 'CAMBIO DE VALOR (15/01/2024) DE j. ANTES ERA J + 1
            Shift = Shift + 1 'Se avanza un turno'
            
            'Asignamos posiciones de las celdas para cálculo de agregados'
            If week = 1 And j = BoxWeekSearch(1) Then
                BoxSheet.Cells(i + 1, j).Value = 0
            Else
                'Celdas que se utilizan en el cálculo'
                Set Cell1 = BoxSheet.Cells(i, j - 1)
                Set Cell2 = BoxSheet.Cells(i + 1, j - 1)
                Set Cell3 = BoxSheet.Cells(i + 2, j - 1)
                Set Cell4 = BoxSheet.Cells(i + 3, j - 1)
                'Formula para calcular agregados'
                BoxSheet.Cells(i + 1, j).Formula = "=" & Cell2.Address & "-" & Cell1.Address & "+IF(" & Cell4.Address & "=""""," & Cell3.Address & "," & Cell4.Address & ")" 'Esta fórmula no se escribe igual que en la propia celda'
            End If
        Next j
    Next i
End Sub

Sub BoxWeekFormat(week As Integer)
    'APLICA EL FORMATO A CELDAS DE TODA LA SEMANA COMPLETA EN LA PESTAÑA BOX'
    Dim FormatSheet As Worksheet
    Set FormatSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("Box"))
    
    Dim FormatRange As Range
    Set FormatRange = FormatSheet.Range("A34:R37")
    FormatRange.Copy
    
    Dim LastRowBox As Integer
    LastRowBox = BoxSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row
    
    Dim BoxRange As Range
    
    For i = OffsetFilaCabecera() + 1 To LastRowBox Step BoxRowDistance
        Set BoxRange = BoxSheet.Cells(i, BoxWeekSearch(week))
        BoxRange.PasteSpecial xlPasteFormats
    Next i
    
    'Function BoxProdWeekSearch(Week As Integer, Ref As String, Shift As Integer) As Integer
    '    'Busca mediante VLookUp el valor de producción diaria de la semana
    '    'pasada como argumento, de la referencia deseada
    '    'Dicho valor lo busca en la pestaña WELDING
    '    'Shift corresponde al turno de la semana del que se desea conocer el valor. Van numerados del 1 (Lunes Noche), hasta el 18( Sabado Tarde)
    '    Dim WeldingSheet As Worksheet
    '    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    '    Dim ReferencesSheet As Worksheet
    '    Set ReferencesSheet = ThisWorkbook.Worksheets(SheetName("References"))
    '
    '    Dim Reference As String
    '    Reference = CStr(Ref)
    '
    '    'Búsqueda de la referencia final en la pestaña References'
    '    Dim FinalRef As String
    '    On Error Resume Next
    '    'LA SIGUIENTE FUNCIÓN VLOOKUP VIGILAR. NECESARIO INTRODUCIR LAS REFERENCIAS QUE SON SÓLO NÚMEROS COMO TEXTO AÑADIENDO UN APOSTROFE DELANTE
    '    FinalRef = Application.WorksheetFunction.VLookup(Reference, ReferencesSheet.Range("B:M"), NumColProcess("Final_Reference"), False) 'LAS REFERENCIAS QUE SON ÚNICAMENTE NÚMEROS NO LAS BUSCA BIEN!!'
    '
    '    'Búsqueda de la fila en la que se encuentra la fila de la referencia pasada por argumento'
    '    Dim RefRow As Integer
    '    RefRow = Application.match(FinalRef, WeldingSheet.Columns(NumColWelding("Reference")), 0)
    '
    '    'Búsqueda del valor en el turno pasado por argumento'
    '    Dim SearchCol As Integer 'Columna en la que se encuentra el valor buscado
    '    SearchCol = WeldingWeekSearch(Week) + (WeldingColDistance() - (3 * 6) - 1) + Shift
    '    BoxProdWeekSearch = WeldingSheet.Cells(RefRow, SearchCol).Value
    'End Function
End Sub
