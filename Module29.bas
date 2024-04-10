Attribute VB_Name = "Module29"
'MODULE_29'
Sub BendingWeekBody(week As Integer, WeekCol As Integer)
    'Añade el cuerpo a cada semana pasada por ARGUMENTO'
    'Columna corresponde a la primera columna de la semana
    Dim BendingSheet As Worksheet
    Dim formatsSheet As Worksheet
    Dim ReferencesSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("Formats"))
    Set ReferencesSheet = ThisWorkbook.Worksheets(SheetName("References"))
    
    Dim Reference As String
    Dim LastRowBending As Integer
    LastRowBending = BendingSheet.Cells(Rows.Count, NumColWelding("Capacity")).End(xlUp).Row
    Dim WeldingCell As Range
    Dim Shift As Integer 'Contador de turnos para bucle'
    Dim LastRowReference As Integer
    LastRowReference = ReferencesSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'Variables para almacenar las celdas de los cálculos de producción'
    Dim Cell1, Cell2, Cell3, Cell4, Cell5 As Range
    Dim sourceRange, destRange As Range
    
    'Variables de creación de fórmula dinámica'
    Dim ArrayForm() As String
    Dim refCounter As Integer
    Dim ArrayPosCounter As Integer
    Dim tempFormula As String
    
    'Formula para agregados'
    For i = OffsetFilaCabecera() + 1 To LastRowBending Step BendingRowDistance()
        If week = 1 Then 'Semana 1 del año se hace diferente'
            Set Cell1 = BendingSheet.Cells(i, WeekCol)
            Set Cell2 = BendingSheet.Cells(i + 1, WeekCol)
            Set Cell3 = BendingSheet.Cells(i + 2, WeekCol)
            Set Cell4 = BendingSheet.Cells(i + 3, WeekCol)
            Set Cell5 = BendingSheet.Cells(i + 1, WeekCol + 1)
            Cell5.Formula = "=" & Cell2.offset(0, 0).Address(False, False) & "-" & Cell1.offset(0, 0).Address(False, False) & "+IF(" & Cell4.offset(0, 0).Address(False, False) & "=""""," & Cell3.offset(0, 0).Address(False, False) & "," & Cell4.offset(0, 0).Address(False, False) & ")"
            Set sourceRange = BendingSheet.Cells(i + 1, WeekCol + 1)
            Set destRange = BendingSheet.Range(BendingSheet.Cells(i + 1, WeekCol + 1), BendingSheet.Cells(i + 1, WeekCol + WeekShifts() - 1))
            sourceRange.AutoFill Destination:=destRange, Type:=xlFillDefault
        Else
            Set Cell1 = BendingSheet.Cells(i, WeekCol - 1)
            Set Cell2 = BendingSheet.Cells(i + 1, WeekCol - 1)
            Set Cell3 = BendingSheet.Cells(i + 2, WeekCol - 1)
            Set Cell4 = BendingSheet.Cells(i + 3, WeekCol - 1)
            Set Cell5 = BendingSheet.Cells(i + 1, WeekCol)
            Cell5.Formula = "=" & Cell2.offset(0, 0).Address(False, False) & "-" & Cell1.offset(0, 0).Address(False, False) & "+IF(" & Cell4.offset(0, 0).Address(False, False) & "=""""," & Cell3.offset(0, 0).Address(False, False) & "," & Cell4.offset(0, 0).Address(False, False) & ")"
            Set sourceRange = BendingSheet.Cells(i + 1, WeekCol)
            Set destRange = BendingSheet.Range(BendingSheet.Cells(i + 1, WeekCol), BendingSheet.Cells(i + 1, WeekCol + WeekShifts() - 1))
            sourceRange.AutoFill Destination:=destRange, Type:=xlFillDefault
        End If
    Next i
    
    'Obtención de datos para demandas desde pestaña WELDING'
    Dim Ref As String
    For i = OffsetFilaCabecera() + 1 To LastRowBending Step BendingRowDistance()
        'Referencia leída de la columna References'
        Ref = BendingSheet.Cells(i, NumColBending("Reference")).Value
        refCounter = 0 ' Inicializa el contador de referencias repetidas a 0'
        
        'Contador de referencias repetidas para dimensionar el array'
        For c = 1 To LastRowReference
            If Ref = ReferencesSheet.Cells(c, NumColReference("References")).Value Then
            refCounter = refCounter + 1
            Else
            End If
        Next c
        
        'Redimensionado del array'
        ReDim ArrayForm(refCounter)
        ArrayPosCounter = 0
        
        'Llenado del array con las refrencias a buscar en la pestaña WELDING'
        For k = 1 To LastRowReference 'Si encuentra la referencia con la que se está trabajando, guarda en el array la referencia final correspondiente
            If Ref = ReferencesSheet.Cells(k, NumColReference("References")).Value Then
                ArrayForm(ArrayPosCounter) = ReferencesSheet.Cells(k, NumColReference("Final_Reference"))
                ArrayPosCounter = ArrayPosCounter + 1
            Else
            End If
        Next k
        
        'Búsqueda de la referencia final en la pestaña References'
        Dim FinalRef As String
        On Error Resume Next
        Shift = 1  'Se suma 3 para colocarse en la celda del primer turno'
        For j = WeekCol To (WeekCol + WeekShifts() - 1) 'Bucle para recorrer todos los turnos'
            'MsgBox j & "/" & WeekCol + WeekShifts()
            Dim WeldingFormula As String 'Cadena que almacenará la fórmula'
            Set WeldingCell = WeldingSheet.Cells(WeldingReferenceRow(FinalRef), Shift)
            tempFormula = BendingFormulaBuilder(ArrayForm(), week, Shift)
            If tempFormula <> "" Then
                BendingSheet.Cells(i, j).Formula = "=" & tempFormula
            Else
                BendingSheet.Cells(i, j) = ""
            End If
            'MsgBox "El turno total es " & Shift
            Shift = Shift + 1
        Next j
    Next i
End Sub

Sub BendingWeekFormat(week As Integer)
    'Aplica el formato de celda a la semana pasada por argumento en la pestaña BENDING
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("bending"))
    
    Dim LastRowBending As Integer
    LastRowBending = BendingSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row
    
    Dim FormatRange As Range
    Set FormatRange = formatsSheet.Range("A76:R79")
    FormatRange.Copy
    Dim destRange
    Set destRange = BendingSheet.Range(BendingSheet.Cells(OffsetFilaCabecera() + 1, BendingWeekSearch(week)), BendingSheet.Cells(LastRowBending + 3, BendingWeekSearch(week) + WeekShifts() - 1))
    
    destRange.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False 'Limpia el portapapeles'
End Sub
