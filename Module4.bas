Attribute VB_Name = "Module4"
'MODULE_4'
'ACTUALIZACIÓN CABECERAS SEMANAS'
Sub WeeksHeaders()
    'Actualiza las cabeceras de las semanas en la pestaña WELDING hasta la semana actual'
    'Este modulo es el encargado de asignar el número de semanas en cada año
    'Debido a la diferencia al comienzo de contabilizar la primera semana del año, entre los ordenadores y los calendarios, es importante revisar
    'este módulo cada año para ajustar la primera semana del año
    'Se marca una de las líneas clave con asteriscos
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    Dim formatsSheet As Worksheet
    Set formatsSheet = ThisWorkbook.Worksheets(SheetName("FORMATS"))
    
    Dim CurrentWeek As Integer
    CurrentWeek = Application.WorksheetFunction.IsoWeekNum(Date)
    
    Dim CurrentRow As Integer
    CurrentRow = OffsetFilaCabecera()
    
    Dim CurrentCol As Integer
    CurrentCol = FirstActualCol() 'Columna E'

    Dim week As Integer
    Dim inicioAno As Integer 'INDICA EN QUE SEMANA DEL AÑO SE COMIENZA A TRABAJAR. EJ: EN 2023 SE COMIENZA EN LA SEMANA 2'
    inicioAno = StartWeek() '- 1
    Dim Counter As Integer
    
    'Rango para la copia de formatos en cabeceras'
    Dim FormatRange As Range

    'Cabeceras principales: LINEA/CD&V/REFERENCE'
    Set FormatRange = formatsSheet.Range("A72:D72")
    WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("LINE")).Value = "LÍNEA"
    WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("CAPACITY")).Value = "CD&V"
    WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("ID")).Value = "ID"
    WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("REFERENCE")).Value = "REFERENCE"
    'WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("WIP_1")).Value = "BOX"
    'WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("WIP_2")).Value = "BENDING"
    'WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("WIP_3")).Value = "OTHER"
    
    FormatRange.Copy
    WeldingSheet.Cells(OffsetFilaCabecera(), NumColWelding("LINE")).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False 'Limpia el portapapeles'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Desactualizado'
    '26/05/2023 SE MODIFICA PARA TRABAJAR CON RANGOS'
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 1) = CellFormat("WELDING", CurrentRow, CurrentCol - 1, 208, 206, 206, True, "xlMedium")
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 1) = "REFERENCE"
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 1).Columns.AutoFit
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 2) = CellFormat("WELDING", CurrentRow, CurrentCol - 2, 208, 206, 206, True, "xlMedium")
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 2) = "ID"
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 2).Columns.AutoFit
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 3) = CellFormat("WELDING", CurrentRow, CurrentCol - 3, 208, 206, 206, True, "xlMedium")
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 3) = "CD&V"
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 4) = CellFormat("WELDING", CurrentRow, CurrentCol - 4, 208, 206, 206, True, "xlMedium")
    '    WeldingSheet.Cells(CurrentRow, CurrentCol - 4) = "LÍNEA"
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Bucle desde semana 1 del año 2023 hasta la semana actual, sumando dos semanas adicionales'
    For week = inicioAno To CurrentWeek + FutureWeeks()
        
        'Escribir "Actual" en la celda correspondiente. Añadiendo formato de celda'
        WeldingSheet.Cells(CurrentRow, CurrentCol) = CellFormat("WELDING", CurrentRow, CurrentCol, 255, 192, 0, True, "xlMedium")
        WeldingSheet.Cells(CurrentRow, CurrentCol) = "Actual"
        WeldingSheet.Cells(CurrentRow, CurrentCol).Columns.AutoFit
        WeldingSheet.Cells(CurrentRow - 2, CurrentCol).Interior.Color = RGB(191, 191, 191)
        WeldingSheet.Cells(CurrentRow - 2, CurrentCol) = "Week " & week '+ 1' '********************************************'
        WeldingSheet.Cells(CurrentRow - 2, CurrentCol).Font.Bold = True
        WeldingSheet.Cells(CurrentRow - 2, CurrentCol).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        WeldingSheet.Cells(CurrentRow - 2, CurrentCol).HorizontalAlignment = xlCenter
        WeldingSheet.Cells(CurrentRow - 2, CurrentCol).VerticalAlignment = xlVAlignCenter
        CurrentCol = CurrentCol + 1
        
        'Escribir "Cargas" en la celda correspondiente'
        WeldingSheet.Cells(CurrentRow, CurrentCol) = CellFormat("WELDING", CurrentRow, CurrentCol, 255, 230, 153, True, "xlMedium")
        WeldingSheet.Cells(CurrentRow, CurrentCol) = "Cargas W" & week
        WeldingSheet.Cells(CurrentRow, CurrentCol).ColumnWidth = 9
        WeldingSheet.Cells(CurrentRow, CurrentCol).WrapText = True
        CurrentCol = CurrentCol + 1
        
        'Escribir "Necesidad" en la celda correspondiente'
        WeldingSheet.Cells(CurrentRow, CurrentCol) = CellFormat("WELDING", CurrentRow, CurrentCol, 226, 239, 218, True, "xlMedium")
        WeldingSheet.Cells(CurrentRow, CurrentCol) = "Necesidad de producción"
        WeldingSheet.Cells(CurrentRow, CurrentCol).ColumnWidth = 15
        WeldingSheet.Cells(CurrentRow, CurrentCol).WrapText = True
        CurrentCol = CurrentCol + 1
        
        'Escribir "Plan" en la celda correspondiente'
        WeldingSheet.Cells(CurrentRow, CurrentCol) = CellFormat("WELDING", CurrentRow, CurrentCol, 226, 239, 218, True, "xlMedium")
        WeldingSheet.Cells(CurrentRow, CurrentCol) = "Plan de producción"
        WeldingSheet.Cells(CurrentRow, CurrentCol).ColumnWidth = 15
        WeldingSheet.Cells(CurrentRow, CurrentCol).WrapText = True
        CurrentCol = CurrentCol + 1
        
        For i = 1 To 6
            Counter = i
            'Escribir "D" en la celda correspondiente'
            WeldingSheet.Cells(CurrentRow, CurrentCol) = "N"
            WeldingSheet.Cells(CurrentRow, CurrentCol).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            WeldingSheet.Cells(CurrentRow - 1, CurrentCol).Value = GetDate(week, Counter)
            WeldingSheet.Cells(CurrentRow - 1, CurrentCol).Font.Bold = True
            WeldingSheet.Cells(CurrentRow - 1, CurrentCol).HorizontalAlignment = xlCenter
            WeldingSheet.Cells(CurrentRow - 1, CurrentCol).VerticalAlignment = xlVAlignCenter
            
            'Sábados en gris claro'
            If Counter = 6 Then
                WeldingSheet.Cells(CurrentRow - 1, CurrentCol).Interior.Color = RGB(232, 232, 232)
            End If
            CurrentCol = CurrentCol + 1
            
            'Escribir "T" en la celda correspondiente'
            WeldingSheet.Cells(CurrentRow, CurrentCol) = "D"
            WeldingSheet.Cells(CurrentRow, CurrentCol).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            CurrentCol = CurrentCol + 1
            
            'Escribir "N" en la celda correspondiente'
            WeldingSheet.Cells(CurrentRow, CurrentCol) = "T"
            WeldingSheet.Cells(CurrentRow, CurrentCol).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            
            'Unir celdas superiores a "N"'
            If WeldingSheet.Cells(CurrentRow - 1, CurrentCol - 2).MergeCells Then '--> Las une si no estaban ya unidas'
            Else
                WeldingSheet.Range(WeldingSheet.Cells(CurrentRow - 1, CurrentCol - 2), WeldingSheet.Cells(CurrentRow - 1, CurrentCol)).Merge
            End If
            CurrentCol = CurrentCol + 1
            
            'Formato a las celdas'
            WeldingSheet.Range(WeldingSheet.Cells(CurrentRow, CurrentCol - 3), WeldingSheet.Cells(CurrentRow, CurrentCol - 1)).Interior.Color = RGB(217, 225, 242)
            WeldingSheet.Range(WeldingSheet.Cells(CurrentRow, CurrentCol - 3), WeldingSheet.Cells(CurrentRow, CurrentCol - 1)).Font.Bold = True
            WeldingSheet.Range(WeldingSheet.Cells(CurrentRow, CurrentCol - 3), WeldingSheet.Cells(CurrentRow, CurrentCol - 1)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
            WeldingSheet.Range(WeldingSheet.Cells(CurrentRow, CurrentCol - 3), WeldingSheet.Cells(CurrentRow, CurrentCol - 1)).HorizontalAlignment = xlCenter
            WeldingSheet.Range(WeldingSheet.Cells(CurrentRow, CurrentCol - 3), WeldingSheet.Cells(CurrentRow, CurrentCol - 1)).VerticalAlignment = xlVAlignCenter
            
        Next i
    
    'Merge celdas correspondientes a Week'
    WeldingSheet.Range(WeldingSheet.Cells(CurrentRow - 2, CurrentCol - WeldingColDistance()), WeldingSheet.Cells(CurrentRow - 2, CurrentCol - 1)).Merge
    
    'Bordes gruesos para separar Weeks'
    WeldingSheet.Range(WeldingSheet.Cells(CurrentRow - 2, CurrentCol - WeldingColDistance()), WeldingSheet.Cells(CurrentRow, CurrentCol - 1)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    
    'Ancho de la fila de cabeceras'
    WeldingSheet.Range("A6").RowHeight = 52

    Next week
    
End Sub

