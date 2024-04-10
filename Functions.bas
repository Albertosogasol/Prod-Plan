Attribute VB_Name = "Functions"
'MODULE_1'
Function SheetName(name As String) As String
    Select Case UCase(name) 'Lo convertimos previamente a mayúsculas para hacer la comparación'
        Case "EDI"
            SheetName = "EDI"
        Case "WELDING", "SOLDADURA"
            SheetName = "WELDING"
        Case "BENDING", "CURVADO"
            SheetName = "BENDING"
        Case "PROCESS", "PROCESOS", "PROCCESS"
            SheetName = "Process"
        Case "REFERENCES", "REFERENCIAS"
            SheetName = "References"
        Case "BOX", "BOXES", "CAJA", "CAJAS"
            SheetName = "BOX"
        Case "WELDING_BACKUP"
            SheetName = "WELDING_backup"
        Case "Formats", "Format", "formats", "format", "FORMATS", "FORMAT"
            SheetName = "Formats"
        Case "BOX_BACKUP", "BOXBACKUP"
            SheetName = "BOX_backup"
        Case "WELDING_BACKUP_2", "WELDING_BACKUP_SEC"
            SheetName = "WELDING_backup_sec"
        Case "BOX_BACKUP_SEC", "BOX_BACKUP_2"
            SheetName = "BOX_backup_sec"
        Case "BENDING_BACKUP", "BENDINGBACKUP"
            SheetName = "BENDING_backup"
        Case "BENDING_BACKUP_SEC", "BENDING_BACKUP_2"
            SheetName = "BENDING_backup_sec"
        Case "VERIFICATION", "VER", "VERIFICATIONS", "VERIFICACION", "VERIFICACIÓN", "VERIFICACIONES"
            SheetName = "VERIFICATION"
        Case "SOLVER", "SOLVER_SHEET", "SOLVERSHEET", "SHEETSOLVER", "SHEET_SOLVER", "PERS_SOLVER", "PERSSOLVER"
            SheetName = "Pers_Solver"
        Case Else
            MsgBox "NO SE ESTÁ REFERENCIANDO UNA PESTAÑA CORRECTA EN EL MODULO_1 SheetName"
    End Select
End Function

Function FirstActualCol() As Integer
    'DEVUELVE EL VALOR DE LA PRIMERA COLUMNA CON DATOS DEL EDI EN LA PESTAÑA WELDING'
    'Tal cual está es la columna '
    FirstActualCol = NumColWelding("REFERENCE") + 1
End Function

Function NumCol(Valor As String) As Integer
    'DEVUELVE EL VALOR NÚMERO DE LA COLUMNA CORRESPONDIENTE EN LA PESTAÑA REFERENCIAS'
    'Añadimos offset por si se añaden columnas previas en la tablas de referencias'
    Dim offset As Integer
    offset = 0
    Select Case UCase(Valor)
        Case "REFERENCIA", "REFERENCE", "REF"
            NumCol = offset + 1
        Case "NIVEL", "LEVEL", "LVL"
            NumCol = offset + 2
        Case "PROCESO", "PROCESS"
            NumCol = offset + 3
        Case "LÍNEA", "LINE"
            NumCol = offset + 4
        Case "LINEA"
            NumCol = offset + 4
        Case Else
            NumCol = -100
    End Select
End Function

Function OffsetFilaCabecera() As Integer
    'Devuelve el valor en entero de la fila sobre la que se colcan las cabeceras. Tomando como referencia la fila 6'
    OffsetFilaCabecera = 6
End Function

Function GetDate(semana As Integer, diaSemana As Integer) As String
    Dim fecha As Date
    fecha = DateAdd("ww", semana - 1, DateSerial(Year(Date), 1, 1)) ' Calcular la fecha correspondiente al número de semana
    fecha = DateAdd("d", diaSemana - 1, fecha) ' Añadir los días correspondientes al día de la semana
    GetDate = Format(fecha, "dd-mmm") ' Formatear la fecha en el formato "dd-mmm"
End Function

Function CellFormat(Hoja As String, fila As Integer, Columna As Integer, R As Integer, G As Integer, B As Integer, Bold As Boolean, BorderWeight As String)
    'Da el formato a la celdas pasadas por argumento. Cambia al color deseado y pone en negrita. SIEMPRE HACE CENTRADO HORIZONTAL Y VERTICAL.'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Hoja)
    ws.Cells(fila, Columna).Interior.Color = RGB(R, G, B)
    ws.Cells(fila, Columna).Font.Bold = Bold
    
    Select Case BorderWeight
        Case "xlMedium"
            ws.Cells(fila, Columna).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        Case "xlThick"
            ws.Cells(fila, Columna).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        Case "xlThin"
            ws.Cells(fila, Columna).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        Case "xlHairline"
            ws.Cells(fila, Columna).BorderAround LineStyle:=xlContinuous, Weight:=xlHairline
        Case Else
    End Select
    
    ws.Cells(fila, Columna).HorizontalAlignment = xlCenter
    ws.Cells(fila, Columna).VerticalAlignment = xlVAlignCenter
    
End Function

Function StartWeek() As Integer 'CAMBIAR SEGÚN EL AÑO!!!!!!'
    StartWeek = 1
End Function

Function EDIStartColYear() As Integer
    Dim EDISheet As Worksheet
    Set EDISheet = ThisWorkbook.Worksheets(SheetName("EDI"))
    Dim SearchDate As Date
    SearchDate = "01/01/2024" 'Establecer la fecha a buscar. CAMBIARÁ CADA AÑO!!!!!!!!!!'
    Dim FoundDate As Range
    Set FoundDate = EDISheet.Rows(2).Find(What:="01/01/2024", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not FoundDate Is Nothing Then
        Dim FoundCol As Integer
        FoundCol = FoundDate.Column
    Else
        MsgBox "No se encontró ninguna fecha con el año 2024 en la fila 2 del EDI!"
    End If
    EDIStartColYear = FoundCol
End Function

Function currentYear() As Integer
    'Devuelve entero con el año con el que se está trabajando'
    currentYear = 2024
End Function

Function FindColYearEDI() As Integer
    'Devuelve un entero con la columna en la que se encuentra la primera fecha del año actual'
    Dim HojaEDI As Worksheet
    Set HojaEDI = ThisWorkbook.Worksheets(SheetName("EDI"))

    Dim col As Integer
    Dim currentYear As Integer
    'CurrentYear = CurrentYear() 'CADA AÑO CAMBIAR ESTO!'
    currentYear = 2024
    ' Buscar la primera fecha en la fila 2 que coincide con el año 2024
    'For col = 1 To HojaEDI.Columns.Count
    '    If IsDate(HojaEDI.Cells(2, col).Value) Then ' comprobamos si la celda contiene una fecha
    '        If Year(CDate(HojaEDI.Cells(2, col).Value)) = currentYear Then ' convertimos a fecha y comprobamos si es del año buscado
     '           Exit For
      '      End If
       ' End If
    'Next col
    
    'Devuelve el número de columna ABSOLUTO de la hoja EDI'
    'FindColYearEDI = col
    FindColYearEDI = 2
End Function

Function WeldingColDistance() As Integer
    'Devuelve el número de columnas entre dos celdas equivalentes a cada semana en la pestaña WELDING'
    WeldingColDistance = 22
End Function

Function WeldingRowDistance() As Integer
    'Devuelve el número de filas entre dos celdas equivalentes a cada referencia en la pestaña WELDING'
    WeldingRowDistance = 4
End Function

Function FutureWeeks() As Integer
    'Devuelve el número de semanas a futuro (sobre la actual) para los cálculos'
    FutureWeeks = 3
End Function

Function NumColWelding(col As String) As Integer
    'DEVUELVE EL VALOR DE LA COLUMNA DESEADA EN LA PESTAÑA WELDING'
    Select Case UCase(col)
        Case "LÍNEA", "LINEA", "LINE", "LINES", "LINEAS", "LÍNEAS"
            NumColWelding = 1
        Case "CAPACIDAD", "CAPACITY", "CAPACIDADES"
            NumColWelding = 2
        Case "ID", "IDENTIFICACIÓN", "IDENTIFICACION", "IDENTIFICATION"
            NumColWelding = 3
        Case "REFERENCIA", "REF", "REFERENCE"
            NumColWelding = 4
        ' Case "WIP_1", "WIP1"
        '     NumColWelding = 5
        ' Case "WIP_2", "WIP2"
        '     NumColWelding = 6
        ' Case "WIP_3", "WIP3", "OTHER", "OTHERS"
            'NumColWelding = 7
        Case Else
            MsgBox "ERROR. Argumento inválido en función: NumColWelding(). MODULE1"
    End Select
End Function

Function StringToDate(InitValue As String) As Date
    'CONVIERTE UNA FECHA DE TIPO STRING A UNA DE TIPO DATE'
    On Error GoTo ErrorHandler
    StringToDate = CDate(InitValue)
    Exit Function
ErrorHandler:
    MsgBox "Se ha producido un error en la ejecución de StringToDate debido a una fecha con formato incorrecto en la pestaña EDI"
End Function

Function ProdNeed(week As Integer) As Integer
    'Calcula las necesidades de producción de la semana pasada como argumento'
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))

    Dim lastRowWelding As Long
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row + 3
    
    'Columna respecto a la cabecera "Week" donde se coloca la celda Necesidades de producción'
    Dim NeedCol As Integer
    NeedCol = WeldingWeekSearch(week) + 2
    
    'ESTE BUCLE APLICA DIRECTAMENTE LOS VALORES. EL 12/04/2023 SE CAMBIA PARA QUE APLIQUE FÓRMULAS'
    'Bucle recorrer todas las referencias'


    For i = OffsetFilaCabecera() + 1 To lastRowWelding Step WeldingRowDistance()
        'Obtener las referencias de las celdas a restar
        Dim Cell1 As Range
        Dim Cell2 As Range
        
        Set Cell1 = WeldingSheet.Cells(i, NeedCol - 1)
        Set Cell2 = WeldingSheet.Cells(i, NeedCol - 2)
        
        'Aplicar la fórmula en la celda correspondiente
        WeldingSheet.Cells(i, NeedCol).Formula = "=" & Cell1.Address & "-" & Cell2.Address
        
        'Obtener las referencias de las celdas a restar
        Set Cell1 = WeldingSheet.Cells(i + 1, NeedCol - 1)
        Set Cell2 = WeldingSheet.Cells(i + 1, NeedCol - 2)
        
        'Aplicar la fórmula en la celda correspondiente
        WeldingSheet.Cells(i + 1, NeedCol).Formula = "=" & Cell1.Address & "-" & Cell2.Address
        
        'Borrar la celda en blanco
        'WeldingSheet.Cells(i + 2, NeedCol).ClearContents
    Next i
End Function

Function WeldingWeekSearch(weekNumber As Integer) As Integer
    'DEVUELVE LA COLUMNA (COMO LONG) EN LA QUE COMIENZA LA SEMANA PASADA COMO ARGUMENTO EN LA PESTAÑA WELDING'
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Dim week As String
    week = "Week " & weekNumber
    Dim cell As Range
    Set cell = WeldingSheet.Rows(OffsetFilaCabecera() - 2).Find(What:=week, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        WeldingWeekSearch = 0
        MsgBox "No se ha encontrado ninguna semana mediante la función WeldingWeekSearch"
    Else
        WeldingWeekSearch = cell.Column
    End If
End Function

Function CurrentWeekNumber() As Integer
    'DEVUELVE LA SEMANA ACTUAL DEL AÑO'
    CurrentWeekNumber = Application.WorksheetFunction.IsoWeekNum(Date)
End Function

Function ProdPlan(week As Integer) As Integer

    'CALCULA LOS PLANES DE PRODUCCIÓN DE LA SEMANA PASADA COMO ARGUMENTO'
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Dim lastRowWelding As Long
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row + 2
    
    'Columna respecto a la cabecera "Week" donde se coloca la celda Plan de producción'
    Dim PlanProdCol As Integer
    PlanProd = WeldingWeekSearch(week) + 3
    
    'Bucle recorrer todas las referencias'
    Dim RealPlan As Long
    Dim TeoricPlan As Long
    
    'ESTE BUCLE APLICA DIRECTAMENTE LOS VALORES. EL 12/04/2023 SE CAMBIA PARA QUE APLIQUE FÓRMULAS'
    'Dim RealPlanRange As Range' 'Comentadas estas dos variables porque ya están ene el bucle'
    'Dim TeoricPlanRange As Range' 'Comentadas estas dos variables porque ya están ene el bucle'



    For i = OffsetFilaCabecera() + 1 To lastRowWelding Step 4
        'Obtener el rango de celdas para sumar
        Dim RealPlanRange As Range
        Set RealPlanRange = WeldingSheet.Range(WeldingSheet.Cells(i, PlanProd + 1), WeldingSheet.Cells(i, PlanProd + 18))
        
        'Aplicar la fórmula en la celda correspondiente
        WeldingSheet.Cells(i, PlanProd).Formula = "=SUM(" & RealPlanRange.Address & ")"
        
        'Obtener el rango de celdas para sumar
        Dim TeoricPlanRange As Range
        Set TeoricPlanRange = WeldingSheet.Range(WeldingSheet.Cells(i + 2, PlanProd + 1), WeldingSheet.Cells(i + 2, PlanProd + 18))
        
        'Aplicar la fórmula en la celda correspondiente
        WeldingSheet.Cells(i + 2, PlanProd).Formula = "=SUM(" & TeoricPlanRange.Address & ")"
        
        'Borrar la celda en blanco
        'WeldingSheet.Cells(i + 2, PlanProd).ClearContents
    Next i
End Function

Function ProcessCol(Column As String) As Integer
    'Devuelve el número de columna donde se encuentra el string buscado en la pestaña Process'
    '---> Vigilar la referenciación del código. Hay partes del código que llaman a la subrutina NumColProces()
    Dim offset As Integer
    offset = 0
    Select Case UCase(Column)
        Case "REFERENCIA", "REFERENCIAS", "REFERENCES", "REFERENCE", "REF"
            ProcessCol = offset + 1
        Case "ID", "IDENTIFICACION", "IDENTIFICACIÓN", "IDENTIFICACIONES", "IDENTIFICATION", "IDENTIFICATIONS"
            ProcessCol = offset + 2
        Case "PROCESO", "PROCESS", "PROCESOS"
            ProcessCol = offset + 3
        Case "LÍNEA", "LINEA", "LINE", "LÍNEAS", "LINES", "LINEAS"
            ProcessCol = offset + 4
        Case "PROYECTO", "PROJECT", "PROYECTOS", "PROJECTS"
            ProcessCol = offset + 5
        Case "CAPACIDAD", "CAPACITY", "CAP", "CAPACIDADES", "QUANTITY"
            ProcessCol = offset + 6
        Case "COMENTARIO", "COMENTARIOS", "COMMENT", "COMMENTS"
            ProcessCol = offset + 7
        Case "NEXT", "ISNEXT"
            ProcessCol = offset + 8
        Case "CHECK", "CHK", "CHECKED"
            ProcessCol = offset + 9
        Case Else
            ProcessCol = -100
    End Select
End Function

Function BoxRowDistance() As Integer
    'Devuelve el número de filas entre dos celdas equivalentes a cada referencia en la pestaña BOX'
    BoxRowDistance = 4
End Function

Function NumColProcess(col As String) As Integer
    'Devuelve la columna correspondiente en la pestaña PROCESS'
    'SE IGNORA LA PRIMERA COLUMNA PARA TRABAJAR UNICAMENTE CON LAS TABLAS!'
    '--->Vigilar la referenciación del módulo. Hay partes del código que llaman a la subrutina ProcessCol()
    Dim ProcessSheet As Worksheet
    Set ProcessSheet = ThisWorkbook.Worksheets(SheetName("Process"))
    
    Dim offset As Integer 'Añadimos offset por si se necesita en el futuro'
    offset = 0
    Select Case UCase(col)
        Case "REFERENCIA", "REFERENCIAS", "REFERENCES", "REFERENCE", "REF"
            NumColProcess = offset + 1
        Case "ID", "IDENTIFICACION", "IDENTIFICACIÓN", "IDENTIFICACIONES", "IDENTIFICATION", "IDENTIFICATIONS"
            NumColProcess = offset + 2
        Case "PROCESO", "PROCESS", "PROCESOS"
            NumColProcess = offset + 3
        Case "LÍNEA", "LINEA", "LINE", "LÍNEAS", "LINES", "LINEAS"
            NumColProcess = offset + 4
        Case "PROYECTO", "PROJECT", "PROYECTOS", "PROJECTS"
            NumColProcess = offset + 5
        Case "CAPACIDAD", "CAPACITY", "CAP", "CAPACIDADES", "QUANTITY"
            NumColProcess = offset + 6
        Case "COMENTARIO", "COMENTARIOS", "COMMENT", "COMMENTS"
            NumColProcess = offset + 7
        Case "NEXT", "ISNEXT"
            NumColProcess = offset + 8
        Case "CHECK", "CHK", "CHECKED"
            NumColProcess = offset + 9
        Case Else
            NumColProcess = -1000
            MsgBox "ERROR en la llamada de la función NumColProcess. Se está referenciando una columna errónea!"
    End Select
End Function

Function WeldingReferenceRow(Reference As String) As Integer
    'DEVUELVE LA FILA EN LA QUE SE ENCUENTRA LA REFERENCIA PASADA POR ARGUMENTO EN LA PESTAÑA WELDING'
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    On Error Resume Next
    'On Error GoTo 0
    WeldingReferenceRow = Application.match(Reference, WeldingSheet.Columns(NumColWelding("Reference")), 0)
    
End Function

Function BoxWeekSearch(weekNumber As Integer) As Integer
    'DEVUELVE LA COLUMNA (COMO LONG) EN LA QUE COMIENZA LA SEMANA PASADA COMO ARGUMENTO EN LA PESTAÑA BOX'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    
    Dim week As String
    week = "Week " & weekNumber
    Dim cell As Range
    Set cell = BoxSheet.Rows(OffsetFilaCabecera() - 2).Find(What:=week, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        BoxWeekSearch = 0
        MsgBox "No se ha encontrado ninguna semana mediante la función WeldingWeekSearch"
    Else
        BoxWeekSearch = cell.Column
    End If
End Function

Function BoxColDistance() As Integer
    'Devuelve la distancia entre dos columnas equivalentes en semanas distintas en la pestaña BOX'
    BoxColDistance = 18
End Function

Function FirstBoxData() As Integer
    'DEVUELVE EL VALOR DE LA PRIMERA COLUMNA CON DATOS EN LA PESTAÑA BOX'
    'Tal cual está es la columna E'
    FirstBoxData = 5
End Function

Function NumColBox(col As String) As Integer
    'DEVUELVE EL VALOR DE LA COLUMNA DESEADA EN LA PESTAÑA WELDING'
    Select Case UCase(col)
        Case "LÍNEA", "LINEA", "LINE", "LINES", "LINEAS", "LÍNEAS"
            NumColBox = FirstBoxData() - 4
        Case "CAPACIDAD", "CAPACITY", "CAPACIDADES"
            NumColBox = FirstBoxData() - 3
        Case "ID", "IDENTIFICACIÓN", "IDENTIFICACION", "IDENTIFICATION"
            NumColBox = FirstBoxData() - 2
        Case "REFERENCIA", "REF", "REFERENCE"
            NumColBox = FirstBoxData() - 1
        Case Else
            MsgBox "ERROR. Argumento inválido en función: NumColBox(). MODULE1"
    End Select
End Function

Function BoxBackupReferenceRow(Reference As String) As Integer
    'DEVUELVE LA FILA EN LA QUE SE ENCUENTRA LA REFERENCIA PASADA POR ARGUMENTO EN LA PESTAÑA BOX_Backup'
    Dim BackupSheet As Worksheet
    Set BackupSheet = ThisWorkbook.Worksheets(SheetName("Box_backup"))
  
    BoxBackupReferenceRow = Application.match(Reference, BackupSheet.Columns(NumColBox("Reference")), 0)
    
End Function

Function BoxReferenceRow(Reference As String) As Integer
    'DEVUELVE LA FILA EN LA QUE SE ENCUENTRA LA REFERENCIA PASADA POR ARGUMENTO EN LA PESTAÑA BOX'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("Box"))
    BoxReferenceRow = Application.match(Reference, BoxSheet.Columns(NumColBox("Reference")), 0)
End Function

Function BendingReferenceRow(Reference As String) As Integer
    'DEVUELVE LA FILA EN LA QUE SE ENCUENTRA LA REFERENCIA PASADA POR ARGUMENTO EN LA PESTAÑA BENDING'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    If IsError(Application.match(Reference, BendingSheet.Columns(NumColBending("Reference")), 0)) Then
        'MsgBox ("Se ha producido un error en la función BendingReferenceRow, durante la búsqueda de: " & Reference)
    Else
        BendingReferenceRow = Application.match(Reference, BendingSheet.Columns(NumColBending("Reference")), 0)
    End If
End Function

Function WeldingCenterViewWeek(week As Integer) As Integer
    'Centra la vista en la semana introducida por pantalla. Principalmente usado para buscar a través de userform1
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    Dim col As Integer
    col = WeldingWeekSearch(week)
    Dim WeekRange As Range
    On Error GoTo ErrorHandler
    Set WeekRange = WeldingSheet.Cells(OffsetFilaCabecera() - 2, col)
    WeekRange.Select
    Exit Function
ErrorHandler:
    'No hace falta mostrar mensaje por pantalla. La función WeldingWeekSearch ya lo muestra
    'MsgBox "No se encontraron semanas"
End Function

Function BoxCenterViewWeek(week As Integer) As Integer
    'Centra la vista en la semana introducida por pantalla. Principalmente usado para buscar a través de Userform1
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Dim col As Integer
    col = BoxWeekSearch(week)
    Dim WeekRange As Range
    On Error GoTo ErrorHandler
    Set WeekRange = BoxSheet.Cells(OffsetFilaCabecera() - 2, col)
    WeekRange.Select
    Exit Function
ErrorHandler:
    'No hace falta mostrar mensaje por pantalla. La función WeldingWeekSearch ya lo muestra
    'MsgBox "No se encontraron semanas"
End Function

Function BendingCenterViewWeek(week As Integer) As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("Bending"))
    Dim col As Integer
    col = BendingWeekSearch(week)
    Dim WeekRange As Range
    Set WeekRange = ws.Cells(OffsetFilaCabecera() - 2, col)
    WeekRange.Select
End Function

Function NumColReference(col As String) As Integer
    'Devuelve la columna correspondiente en la pestaña REFERENCES'
    'SE IGNORA LA PRIMERA COLUMNA PARA TRABAJAR UNICAMENTE CON LAS TABLAS!'
    Dim ProcessSheet As Worksheet
    Set ProcessSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))
    
    Dim offset As Integer 'Añadimos offset por si se necesita en el futuro'
    offset = 0
    Select Case UCase(col)
        Case "REFERENCIA", "REFERENCE", "REFERENCIAS", "REFERENCES", "REF"
            NumColReference = offset + 2
        Case "LEVEL", "NIVEL", "NIVELES", "LEVELS", "LVL", "NVL"
            NumColReference = offset + 3
        Case "PROCESS", "PROCCESS", "PROCESSES", "PROCESO", "PROCESOS", "PROC"
            NumColReference = offset + 4
        Case "LINE", "LIN", "LINEA", "LÍNEA", "LINES", "LÍNEAS", "LINEAS"
            NumColReference = offset + 5
        Case "FINALREF", "FINAL_REF", "REF_FINAL", "FIN_REF", "REFERENCIA_FINAL", "REFERENCIAFINAL", "FINAL_REFERENCE", "FINAL"
            NumColReference = offset + 6
        Case "NEXT_REFERENCE", "NEXTREFERENCE", "SIGUIENTE_REFERENCIA", "NEXT", "REFERENCIA_SIGUIENTE", "SIGUIENTEREFERENCIA", "NEXT_REF", "REF_NEXT"
            NumColReference = offset + 7
        Case "PREVIOUS_REFERENCE", "PREV_REFERENCE", "PREVIOUS", "PREV_REF", "PREVREF", "PREVREFERENCE", "PREVIOUSREF", "PREVIOUS_REF"
            NumColReference = offset + 8
        Case "CANTIDAD", "CUANTITY", "QY", "CANT", "QUANTITY", "CANTIDADES", "QUANTITIES", "CUANTITIES"
            NumColReference = offset + 9
        Case "OPERARIOS", "OP", "WORKERS", "WORK", "OPERARIO"
            NumColReference = offset + 10
        Case "IS_NEXT"
            NumColReference = offset + 11
        Case "CHK", "CHECK"
            NumColReference = offset + 12
        Case "ID"
            NumColReference = offset + 13
        Case "COMMENTS", "COMMENT"
            NumColReference = offset + 14
        Case "SAGE", "IS_SAGE"
            NumColReference = offset + 15
        Case Else
            NumColReference = -1000
            MsgBox "ERROR en la llamada de la función NumColReference. Se está referenciando una columna errónea!"
    End Select
End Function

Function BoxFormulaBuilder(References() As String, Shift As Integer) As String
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim WeldingCell As Range
    Dim BoxFormulaBuilderTemp As String: BoxFormulaBuilderTemp = ""
    
    'BoxFormulaBuilder = "="

    Dim Size As Integer
    If IsEmpty(References) Then
      Size = 0
    Else
       Size = UBound(References) - LBound(References) - 1
       'MsgBox "El tamaño size es: " & Size
    End If
    For m = 0 To Size Step 1
        Set WeldingCell = WeldingSheet.Cells(WeldingReferenceRow(References(m)), Shift)
        BoxFormulaBuilderTemp = BoxFormulaBuilderTemp & WeldingSheet.name & "!" & WeldingCell.Address & "+"
        BoxFormulaBuilder = Left(BoxFormulaBuilderTemp, Len(BoxFormulaBuilderTemp) - 1) 'Se elimina el último "+"'
    Next m
    'BoxFormulaBuilder = Left(BoxFormulaBuilderTemp, Len(BoxFormulaBuilderTemp) - 1)
    
End Function

Function WeldingDateSearch(myDate As Date) As Integer
    'DEVUELVE LA COLUMNA EN LA QUE SE ENCUENTRA LA FECHA PASADA POR ARG'
    Dim Weld As Worksheet
    Set Weld = ThisWorkbook.Worksheets(SheetName("WELDING"))
End Function

Function WeldingAccumulated(week As Integer) As Integer
    'Aplica la fórmula del acumulado en la pestaña WELDING para la semana pasada como argument
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    
    Dim lastRowWelding As Integer
    lastRowWelding = WeldingSheet.Cells(Rows.Count, NumColWelding("Capacity")).End(xlUp).Row
    
    Dim Cell1, Cell2, Cell3, Cell4, Cell5 As Range
    Dim sourceRange As Range
    Dim destRange As Range
    
    For i = OffsetFilaCabecera() + 1 To lastRowWelding Step WeldingRowDistance() 'Bucle para recorrer todas las referencias'
        If week = 1 Then
            WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 4).Value = 0
            Set Cell1 = WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 4)
            Set Cell2 = WeldingSheet.Cells(i, WeldingWeekSearch(week) + 4)
            Set Cell3 = WeldingSheet.Cells(i + 2, WeldingWeekSearch(week) + 4)
            Set Cell4 = WeldingSheet.Cells(i + 3, WeldingWeekSearch(week) + 4)
            Set Cell5 = WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 5)
            'Cell5.Formula = "=" & Cell1.Address & "-" & Cell2.Address & "+IF(" & Cell4.Address & "=""""," & Cell3.Address & "," & Cell4.Address & ")"
            Cell5.Formula = "=" & Cell1.offset(0, 0).Address(False, False) & "-" & Cell2.offset(0, 0).Address(False, False) & "+IF(" & Cell4.offset(0, 0).Address(False, False) & "=""""," & Cell3.offset(0, 0).Address(False, False) & "," & Cell4.offset(0, 0).Address(False, False) & ")"
            Set WeekRange = WeldingSheet.Range(WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 5), WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 21))
            WeekRange.FillRight
        Else
            WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 4).Formula = "=" & SheetName("Welding") & "!" & WeldingSheet.Cells(i, WeldingWeekSearch(week)).offset(0, 0).Address(False, False)
            'WeldingSheet.Cells(i + 1, WeldingWeekSearch(Week) + 4).Formula = "=" & SheetName("Welding") & "!" & WeldingSheet.Cells(i, WeldingWeekSearch(Week)).Address
            Set Cell1 = WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 4)
            Set Cell2 = WeldingSheet.Cells(i, WeldingWeekSearch(week) + 4)
            Set Cell3 = WeldingSheet.Cells(i + 2, WeldingWeekSearch(week) + 4)
            Set Cell4 = WeldingSheet.Cells(i + 3, WeldingWeekSearch(week) + 4)
            Set Cell5 = WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 5)
            'Cell5.Formula = "=" & Cell1.Address & "-" & Cell2.Address & "+IF(" & Cell4.Address & "=""""," & Cell3.Address & "," & Cell4.Address & ")"
            Cell5.Formula = "=" & Cell1.offset(0, 0).Address(False, False) & "-" & Cell2.offset(0, 0).Address(False, False) & "+IF(" & Cell4.offset(0, 0).Address(False, False) & "=""""," & Cell3.offset(0, 0).Address(False, False) & "," & Cell4.offset(0, 0).Address(False, False) & ")"
            Set sourceRange = WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 5)
            Set destRange = WeldingSheet.Range(WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 5), WeldingSheet.Cells(i + 1, WeldingWeekSearch(week) + 21))
            sourceRange.AutoFill Destination:=destRange, Type:=xlFillDefault
        End If
    Next i
End Function

Function NumColBending(col As String) As Integer
    'Devuelve el número de columna pasada por argumento como String en la pestaña BENDING'
    Dim offset As Integer 'Offset por si se añaden columnas'
    offset = 0 '<----------- EDITAR SI ES NECESARIO
    Select Case UCase(col)
    Case "LÍNEA", "LÍNEAS", "LINE", "LINES", "LINEA", "LINEAS"
        NumColBending = 1
    Case "CD&V", "CAPACITY", "CAP", "CAPACIDAD", "CAPACIDADES", "CD"
        NumColBending = 2
    Case "ID", "IDENTITY", "IDNUM", "IDENTIFICACION"
        NumColBending = 3
    Case "REFERENCE", "REFERENCIA", "REFERENCIAS", "REFERENCES", "REF", "REFS"
        NumColBending = 4
    Case Else
        MsgBox "ERROR. No se está referenciando una columna correcta en la función NumColBending del MODULE_1"
    End Select
    NumColBending = NumColBending + offset
End Function

Function WeekShifts() As Integer
    'Número de turnos de trabajo por semana'
    WeekShifts = 3 * 6
End Function

Function StartShiftWeldingCol() As Integer
    'NÚMERO DE COLUMNAS QUE HAY ENTRE LA LA PRIMERA CELDA DE LA SEMANA Y
    ' LA PRIMERA CELDA DE TURNOS (N).
    StartShiftWeldingCol = WeldingColDistance() - WeekShifts()
End Function

Function BendingRowDistance() As Integer
    BendingRowDistance = 4
End Function

Function BendingWeekSearch(weekNumber As Integer) As Integer
    'DEVUELVE LA COLUMNA (COMO LONG) EN LA QUE COMIENZA LA SEMANA PASADA COMO ARGUMENTO EN LA PESTAÑA BENDING'
    Dim BendingSheet As Worksheet
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("Bending"))
    
    Dim week As String
    week = "Week " & weekNumber
    Dim cell As Range
    Set cell = BendingSheet.Rows(OffsetFilaCabecera() - 2).Find(What:=week, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        BendingWeekSearch = 0
        MsgBox "No se ha encontrado ninguna semana mediante la función BendingWeekSearch"
    Else
        BendingWeekSearch = cell.Column
    End If
End Function

Function BendingFormulaBuilder(References() As String, week As Integer, Shift As Integer) As String
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim WeldingCell As Range
    Dim BendingFormulaBuilderTemp As String: BendingFormulaBuilderTemp = ""
    
    'BoxFormulaBuilder = "="

    Dim Size As Integer
    If IsEmpty(References) Then
      Size = 0
    Else
       Size = UBound(References) - LBound(References) - 1
       'MsgBox "El tamaño size es: " & Size
    End If
    Dim partialShift
    partialShift = WeldingWeekSearch(week) + 3 + Shift
    For m = 0 To Size Step 1
        Set WeldingCell = WeldingSheet.Cells(WeldingReferenceRow(References(m)), partialShift)
        BendingFormulaBuilderTemp = BendingFormulaBuilderTemp & WeldingSheet.name & "!" & WeldingCell.Address & "+"
        BendingFormulaBuilder = Left(BendingFormulaBuilderTemp, Len(BendingFormulaBuilderTemp) - 1) 'Se elimina el último "+"'
    Next m
    'BoxFormulaBuilder = Left(BoxFormulaBuilderTemp, Len(BoxFormulaBuilderTemp) - 1)
    
End Function

Function FirstBendingData() As Integer
    'Devuelve la primera columna con datos de producción en la pestaña BENDING'
    FirstBendingData = NumColBending("Reference") + 1
End Function

Function BendingColDistance() As Integer
    'Devuelve la distancia entre dos columnas equivalentes en semanas distintas en la pestaña BENDING'
    BendingColDistance = 18
End Function

Function NumColVer(col As String) As Integer
    'Devuelve el número de columna pasada por argumento como String en la pestaña VERIFICATION'
    Dim offset As Integer 'Offset por si se añaden columnas'
    offset = 0 '<----------- EDITAR SI ES NECESARIO
    
    Select Case UCase(col)
    Case "FINAL_REF", "FINALREF", "FINAL_REFERENCE", "FINALREFERENCE"
        NumColVer = 2
    Case "ID"
        NumColVer = 3
    Case "PROCESS", "PROCC", "PROCESO", "PROCESOS"
        NumColVer = 4
    Case "LVL", "LEVEL", "NIVEL", "NVL", "NIVELES", "LEVELS"
        NumColVer = 5
    Case "REFERENCE", "REFERENCIA", "REFERENCIAS", "REFERENCES", "REF", "REFS"
        NumColVer = 6
    Case Else
        MsgBox "ERROR. No se está referenciando una columna correcta en la función NumColVer del MODULE_1"
    End Select
    NumColVer = NumColVer + offset
    
End Function
Function WIPSet(Reference As String) As Variant  'TRABAJANDO COMO ARRAY PERO SE PUEDE MODIFICAR A COLLECTION
    'Devuelve los wips anteriores y posteriores de las referencias pasadas por argumento
    'Dependiendo de la referencia pasada por argumento, se mostrarán unicamente los wips posteriores o también los anteriores.
    'Esto se realiza mediante la lectura del nivel en la pestaña REFERENCES
    Dim WeldingSheet As Worksheet
    Dim BoxSheet As Worksheet
    Dim BendingSheet As Worksheet
    Dim RefSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Set BendingSheet = ThisWorkbook.Worksheets(SheetName("BENDING"))
    Set RefSheet = ThisWorkbook.Worksheets(SheetName("REFERENCES"))
    
    Reference = CStr(Reference)

    Dim tempArray(3) As String
    Dim tempString As String


    Dim lastRowRef As Integer
    lastRowRef = RefSheet.Cells(RefSheet.Rows.Count, NumColReference("Reference")).End(xlUp).Row

    Dim findRange As Range
    Set findRange = RefSheet.Range(RefSheet.Cells(1, 2), RefSheet.Cells(lastRowRef, NumColReference("Reference")))

    Dim foundRow As Integer
    Dim lvlTemp As Integer
    Dim srchoffset As Integer 'Número de filas en las que realiza búsqueda de los wips siguientes. Se evita que busque en toda la hoja' 'Calculado a ojo'
    srcoffset = 5

    'Búsqueda de referencia pasada por argumento'
    'REVISAR CORREO'
    'FOR EACH CELDA IN RANGE
    foundRow = Application.WorksheetFunction.match(Reference, findRange, 0)
    If RefSheet.Cells(foundRow, NumColReference("Level")).Value = 0 Then
        
    Else
    End If
End Function

Function VerificationWeekSearch(weekNumber As Integer) As Integer
    'DEVUELVE LA COLUMNA (COMO LONG) EN LA QUE COMIENZA LA SEMANA PASADA COMO ARGUMENTO EN LA PESTAÑA VERIFICATION'
    Dim VerSheet As Worksheet
    Set VerSheet = ThisWorkbook.Worksheets(SheetName("VERIFICATION"))
    
    Dim week As String
    week = "Week " & weekNumber
    Dim cell As Range
    Set cell = VerSheet.Rows(OffsetFilaCabecera() - 2).Find(What:=week, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        VerificationWeekSearch = 0
        MsgBox "No se ha encontrado ninguna semana mediante la función BendingWeekSearch"
    Else
        VerificationWeekSearch = cell.Column
    End If
End Function

Function ReferencesOffset() As Integer
    'Devuelve el número de columnas a la izquierda en la pestaña REFERENCES
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("REFERENCES"))
    ReferencesOffset = 1
End Function

Function WeldingFormulaBuilder(References() As String, week As Integer) As String
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("WELDING"))
    Dim WeldingCell As Range
    Dim WeldingFormulaBuilderTemp As String: WeldingFormulaBuilderTemp = ""

    Dim Size As Integer
    If IsEmpty(References) Then
        Size = 0
    Else
        Size = UBound(References) - LBound(References) - 1
        'MsgBox "El tamaño size es: " & Size
    End If
    For m = 0 To Size Step 1
        Set WeldingCell = WeldingSheet.Cells(WeldingReferenceRow(References(m)), WeldingWeekSearch(week) + 1)
        WeldingFormulaBuilderTemp = WeldingFormulaBuilderTemp & WeldingSheet.name & "!" & WeldingCell.Address & "+"
        WeldingFormulaBuilder = Left(WeldingFormulaBuilderTemp, Len(WeldingFormulaBuilderTemp) - 1) 'Se elimina el último "+"'
    Next m
    'BoxFormulaBuilder = Left(BoxFormulaBuilderTemp, Len(BoxFormulaBuilderTemp) - 1)

End Function

Function checkLastWelding(Reference As String) As Boolean
    'Comprueba si la referencia de soldadura pasada por argumento es una referencia final'
    'Para ello comprueba si existe una tabla con su nombre en la pestaña REFERENCES'
    Dim table As ListObject
    On Error Resume Next
    Set table = ThisWorkbook.Worksheets(SheetName("REFERENCES")).ListObjects("Table_" & Reference)
    On Error GoTo 0
    
    If Not table Is Nothing Then
        checkLastWelding = True
    Else
        checkLastWelding = False
    End If
End Function

Function NumColSolver(col As String) As Integer
    'Devuelve como entero la columna pasada por argumento en la pestaña SOLVER
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("SOLVER_SHEET"))

    Dim offset As Integer
    offset = 0

    Select Case UCase(col)
    Case "PROCESO", "PROCESS", "PROCESOS"
        NumColSolver = 1
    Case "LÍNEA", "LINEA", "LINE", "LÍNEAS", "LINEAS", "LINES"
        NumColSolver = 2
    Case "REF", "REFERENCIA", "REFERENCE", "REFERENCIAS", "REFERENCES"
        NumColSolver = 3
    Case "PERSONAL", "PERSONAS", "PERS", "PERSONS", "PERSON"
        NumColSolver = 4
    Case "PIEZAS", "PZ", "PIECES", "PIEZA", "CANTIDAD", "QUANTITY", "CANT"
        NumColSolver = 5
    Case Else
        MsgBox "ERROR en la función NumColSolver. No se está pasando un argumento válido", , "NumColSolver_ERROR"
        End Select
End Function

Function CheckString(sheet As String) As Integer
    'Comprueba si todas las referencias de la pestaña pasada por argumento se encuentran almacenadas como cadenas
    Dim ws As Worksheet
    Dim lastRow As Integer
    sheet = UCase(sheet)
    Select Case sheet
        Case "WELDING"
            Set ws = ThisWorkbook.Worksheets(SheetName("WELDING"))
            lastRow = ws.Cells(Rows.Count, NumColWelding("Reference")).End(xlUp).Row
            For Row = OffsetFilaCabecera() + 1 To lastRow Step WeldingRowDistance
                If IsNumeric(ws.Cells(Row, NumColWelding("Reference"))) Then
                    ws.Cells(Row, NumColWelding("Reference")).Value = "'" & ws.Cells(Row, NumColWelding("Reference")).Value
                End If
            Next Row
        Case Else
            'statments
    End Select
End Function

Sub PruebaCheckString()
    CheckString ("WELDING")
End Sub

