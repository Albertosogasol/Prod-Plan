Attribute VB_Name = "Module6_2"
'MODIFICACIÓN MÓDULO 6 PARA LA ACTUALIZACIÓN DE LAS DEMANDAS DEL EDI'
    'Las referencias que no son finales, deben tener una demanda basada en sus referencia final.
Sub IMPORT_EDI()
    
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
            weldingReferenceEDIRow = Application.match(weldingReference, EDISheet.Columns(1), 0)
            
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
End Sub


'''''''''''''''''MÉTODO DESACTUALIZADO'''''''''''''''''''''''''''''''''''

'MODULE_6'
'ACTUALIZACIÓN DEMANDAS EDI'
'Sub ImportEDI()
'    'Mejora en la obtención de datos del EDI, asignando las demandas por fechas'
'    'Comprobación de semanas correctas en pestaña EDI'
'    'CheckWeekEDI
'
'    Dim EDISheet As Worksheet
'    Dim WeldingSheet As Worksheet
'    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
'    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
'
'    'Columna en la que se encuentas las referencias en la pestaña WELDING'
'    Dim ReferenceCol As Integer
'    ReferenceCol = NumColWelding("Reference")
'
'    'Ultima posición de datos en pestaña WELDING'
'    Dim LastColWelding As Integer
'    Dim LastRowWelding As Integer
'    LastColWelding = WeldingSheet.Cells(OffsetFilaCabecera(), Columns.Count).End(xlToLeft).Column
'    LastRowWelding = WeldingSheet.Cells(Rows.Count, ReferenceCol).End(xlUp).row
'
'    'Variables bucles'
'    Dim WeldingReference As String
'    Dim Week As Integer
'    Dim LoadsCounter As Integer
'    Dim Reference As String
'    Dim ReferenceCounter As Integer
'
'    'Bucle para recorrer toda las filas de referencias en la pestaña WELDING'
'    For i = OffsetFilaCabecera() + 1 To LastRowWelding Step WeldingRowDistance()
'         WeldingReference = WeldingSheet.Cells(i, NumColWelding("Reference")).Value
'         Week = StartWeek()
'         LoadsCounter = FirstActualCol() + 1
'         ReferenceCounter = i
'         'Bucle para recorrer todas las semanas del año en la pestaña WELDING. Celdas "Cargas"'
'         For j = LoadsCounter To LastColWelding Step WeldingColDistance()
'            On Error Resume Next
'            EDIColSearch = FindWeekColumnEDI(Week)
'            Reference = WeldingSheet.Cells(ReferenceCounter, ReferenceCol).Value 'Vigilar. Convierte las referencias de números a Str y las ignora'
'            WeldingSheet.Cells(i, j).Value = Application.WorksheetFunction.VLookup(WeldingSheet.Cells(ReferenceCounter, ReferenceCol).Value, EDISheet.Range("A:ZZ"), EDIColSearch, False)
'            'MsgBox "Referencia: " & WeldingSheet.Cells(ReferenceCounter, ReferenceCol).Value & " en la semana " & Week & ". Cargas: " & Application.WorksheetFunction.VLookup(Reference, EDISheet.Range("A:ZZ"), EDIColSearch, False)
'            Week = Week + 1
'            On Error GoTo 0
'         Next j
'    Next i
'
'End Sub
