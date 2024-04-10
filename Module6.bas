Attribute VB_Name = "Module6"
'MODULE_6'
'ACTUALIZACIÓN DEMANDAS EDI'
Sub ImportEDI()
    Dim EDISheet As Worksheet
    Dim WeldingSheet As Worksheet
    Dim RefSheet As Worksheet
    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))

    'Última posición de datos en pestaña WELDING'
    Dim lastColWelding As Integer
    Dim lastRowWelding As Integer
    lastColWelding = WeldingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row

    'Variables para bucle'
    Dim weldingReference As String
    Dim week As Integer
    Dim loadsCounter As Integer
    Dim referenceCounter As Integer
    Dim weldingReferenceEDIRow As Integer 'Fila donde se encuentra la referencia buscada en el bucle
    Dim weldingReferenceLevel As Integer 'Nivel en la tabla de REFERENCES'
    Dim refTable As Range 'Tabla que contiene la referencia en la pestaña REFERENCES
    Dim referencesList() As String
    Dim referencesListDim As Integer
    Dim loopCounter As Integer
    Dim tempWeldingFormula As String

    'Bucle para recorrer todas las filas de referencias en la pestaña WELDING'
    For i = OffsetFilaCabecera() + 1 To lastRowWelding Step WeldingRowDistance()
        weldingReference = WeldingSheet.Cells(i, NumColWelding("Reference")).Value
        week = StartWeek()
        loadsCounter = FirstActualCol() + 1
        referenceCounter = i
        'Comprobamos si la referencia leída es un proceso final o forma parte de otro'
        If (checkLastWelding(weldingReference) = True) Then
            'Referencia final --> Se busca la demanda en el EDI
            'Búsqueda de fila en el EDI para iniciar bucle
            On Error Resume Next
            weldingReferenceEDIRow = Application.match(weldingReference, EDISheet.Columns(1), 0)
            On Error GoTo 0
            'Bucle para recorrer todas las semanas'
            For j = loadsCounter To lastColWelding Step WeldingColDistance()
                On Error Resume Next 'Manejo de errores por si no existe la semana'
                WeldingSheet.Cells(i, j).Value = EDISheet.Cells(weldingReferenceEDIRow, FindWeekColumnEDI(week)).Value
                week = week + 1
                On Error GoTo 0
            Next j
        Else
            If Not IsError(Application.match(weldingReference, EDISheet.Columns(1), 0)) Then
                For j = loadsCounter To lastColWelding Step WeldingColDistance()
                    On Error Resume Next 'Manejo de errores por si no existe la semana'
                    WeldingSheet.Cells(i, j).Value = EDISheet.Cells(weldingReferenceEDIRow, FindWeekColumnEDI(week)).Value
                    week = week + 1
                    On Error GoTo 0
                Next j
            Else
                'Referencia perteneciente a otra superior --> Se busca su referencia final en la tabla.
                'Se trabaja con fórmulas para tener la suma de referencias en caso de que pertenezca a distintas
                'Creación de array de dimensión n para almacenar las referencias finales que intervienen'
                referencesListDim = 0
                For Each cel In RefSheet.Range("B:B")
                    If cel.Value = weldingReference Then
                        referencesListDim = referencesListDim + 1
                    End If
                Next cel
                ReDim referencesList(referencesListDim)
                'Se realiza una búsqueda de las referencias para almacenar en el array las referencias finales
                loopCounter = 0
                For Each cel In RefSheet.Range("B:B") '<---------------------------------------- CAMBIAR LETRAS A FUNCIONES!!! <------------------------
                    If cel.Value = weldingReference Then
                        referencesList(loopCounter) = cel.offset(0, 4).Value
                        loopCounter = loopCounter + 1
                    Else
                    End If
                Next cel
                'Creación de la fórmula
                For j = loadsCounter To lastColWelding Step WeldingColDistance()
                    On Error Resume Next 'Manejo de errores por si no existe la semana'
                    tempWeldingFormula = WeldingFormulaBuilder(referencesList(), week)
                    WeldingSheet.Cells(i, j).Value = "=" & tempWeldingFormula
                    week = week + 1
                    On Error GoTo 0
                Next j
            End If
        End If
    Next i

    ' Sub ImportEDI()
    '     Dim EDISheet As Worksheet
    '     Dim WeldingSheet As Worksheet
    '     Dim RefSheet As Worksheet
    '     Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
    '     Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    '     Set RefSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))

    '     'Última posición de datos en pestaña WELDING'
    '     Dim lastColWelding As Integer
    '     Dim lastRowWelding As Integer
    '     lastColWelding = WeldingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
    '     lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).row

    '     'Variables para bucle'
    '     Dim weldingReference As String
    '     Dim week As Integer
    '     Dim loadsCounter As Integer
    '     Dim referenceCounter As Integer
    '     Dim weldingReferenceEDIRow As Integer 'Fila donde se encuentra la referencia buscada en el bucle
    '     Dim weldingReferenceLevel As Integer 'Nivel en la tabla de REFERENCES'
    '     Dim refTable As Range 'Tabla que contiene la referencia en la pestaña REFERENCES

    '     'Bucle para recorrer todas las filas de referencias en la pestaña WELDING'
    '     For i = OffsetFilaCabecera()+1 To LastRowWelding Step WeldingRowDistance()
    '         WeldingReference = WeldingSheet.Cells(i, NumColWelding("Reference")).Value
    '         week = StartWeek()
    '         loadsCounter = FirstActualCol() + 1
    '         referenceCounter = i
    '         'Comprobamos si la referencia leída es un proceso final o forma parte de otro'
    '         If (checkLastWelding(WeldingReference) = TRUE ) Then
    '             'Referencia final --> Se busca la demanda en el EDI
    '             'Búsqueda de fila en el EDI para iniciar bucle
    '             weldingReferenceEDIRow = Application.Match(WeldingReference, EDISheet.Columns(1), 0)
                
    '             'Bucle para recorrer todas las semanas'
    '             For j = loadsCounter To lastColWelding Step WeldingColDistance()
    '                 On Error Resume Next 'Manejo de errores por si no existe la semana'
    '                 WeldingSheet.Cells(i,j).Value = EDISheet.Cells(weldingReferenceEDIRow,FindWeekColumnEDI(Week)).Value
    '                 Week = Week + 1
    '                 On Error GoTo 0
    '             Next j
    '         Else
    '             'Referencia perteneciente a otra superior --> Se busca su referencia final en la tabla.
    '             'Se trabaja con fórmulas para tener la suma de referencias en caso de que pertenezca a distintas

    '         End If
    '     Next i
    ' End Sub
End Sub

Function FindWeekColumnEDI(weekNumber As Integer) As Integer
    'Devuelve la columna de la pestaña EDI en la que se almacenan las demandas de la semana pasada como argumento'
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim EDISheet As Worksheet
    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))

    Dim currentYear As Integer
    currentYear = Year(Date) ' Obtiene el año actual
    
    Dim CurrentWeek As Integer
    CurrentWeek = DatePart("ww", Date) ' Obtiene la semana actual del año
    
    Dim weekOffset As Integer
    weekOffset = weekNumber - CurrentWeek ' Calcula la diferencia entre la semana actual y la semana deseada
    
    Dim EDILastCol As Integer
    EDILastCol = EDISheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim EDIRange As Range
    Set EDIRange = EDISheet.Range(EDISheet.Cells(1, FindColYearEDI()), EDISheet.Cells(1, EDILastCol))
    
    Dim targetColumn As Integer 'Columna relativa en el EDIRange donde se encuentra la semana buscada'
    On Error GoTo ErrorMsg
    targetColumn = Application.WorksheetFunction.match("S" & weekNumber, EDIRange, 0) ' Busca la columna que contiene la etiqueta de la semana
    
    Dim targetColumnAbs As Integer
    'Para pasar la columna relativa a absoluta, tenemos en cuenta la función FindColYearEDI()'
    targetColumnAbs = FindColYearEDI() + targetColumn - 1 'Se resta 1 porque el rango relativo empieza en la columna 1'
    FindWeekColumnEDI = targetColumnAbs
    Exit Function
    'Si la semana buscada en el EDI no existe, devuelve la columna de referencias. En el plan saldrá como referencias en la columna cargas
ErrorMsg:
    FindWeekColumnEDI = 1
End Function

Sub ImportWeekEDI(week As Integer)
    'Importa los datos del EDI de la semana pasada como argumento'
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim EDISheet As Worksheet
    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
    
    Dim lastRowWelding As Integer
    ReferenceCol = NumColWelding("Reference")
    lastRowWelding = WeldingSheet.Cells(Rows.Count, ReferenceCol).End(xlUp).Row
    
    'Bucle para recorrer toda las filas de referencias en la pestaña WELDING'
    For i = OffsetFilaCabecera() + 1 To lastRowWelding Step WeldingRowDistance()
         weldingReference = WeldingSheet.Cells(i, NumColWelding("Reference")).Value
         referenceCounter = i
            On Error Resume Next
            EDIColSearch = FindWeekColumnEDI(week)
            Reference = WeldingSheet.Cells(referenceCounter, ReferenceCol).Value 'Vigilar. Convierte las referencias de números a Str y las ignora'
            WeldingSheet.Cells(i, WeldingWeekSearch(week) + 1).Value = Application.WorksheetFunction.VLookup(WeldingSheet.Cells(referenceCounter, ReferenceCol).Value, EDISheet.Range("A:ZZ"), EDIColSearch, False)
            'MsgBox "Referencia: " & WeldingSheet.Cells(ReferenceCounter, ReferenceCol).Value & " en la semana " & Week & ". Cargas: " & Application.WorksheetFunction.VLookup(Reference, EDISheet.Range("A:ZZ"), EDIColSearch, False)
            On Error GoTo 0
    Next i
End Sub
