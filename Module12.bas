Attribute VB_Name = "Module12"
'MODULE_12'
'APLICA ESTILO A CELDAS DE LA SEMANA PASADA COMO ARGUMENTO'

'ACTUAL'
Function ActualCellFormatWeek(week As Integer) As Integer
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Dim lastRow As Integer
    lastRow = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row + 2
    
    Dim ActualCellCol As Integer
    Dim ActualCellRow As Integer
    
    For i = OffsetFilaCabecera() + 1 To lastRow Step 3
        ActualCellCol = WeldingWeekSearch(week)
        ActualCellRow = i
        WeldingSheet.Cells(ActualCellRow, ActualCellCol) = CellFormat("WELDING", ActualCellRow, ActualCellCol, 255, 255, 0, True, "xlMedium")
        WeldingSheet.Cells(ActualCellRow + 1, ActualCellCol) = CellFormat("WELDING", ActualCellRow + 1, ActualCellCol, 255, 255, 0, True, "xlMedium")
        WeldingSheet.Range(WeldingSheet.Cells(ActualCellRow, ActualCellCol), WeldingSheet.Cells(ActualCellRow + 1, ActualCellCol)).Merge
    Next i
End Function

Sub ActualCellFormatUpdateAll()
    For i = StartWeek() To CurrentWeekNumber + FutureWeeks
        ActualCellFormatWeek (i)
    Next i
End Sub

'CARGAS'
Function LoadsCellFormatWeek(week As Integer) As Integer
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Dim lastRow As Integer
    lastRow = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row + 2
    
    Dim LoadCellCol As Integer
    Dim LoadCellRow As Integer
    
    For i = OffsetFilaCabecera() + 1 To lastRow Step 3
        LoadCellCol = WeldingWeekSearch(week) + 1
        LoadCellRow = i
        WeldingSheet.Cells(LoadCellRow, LoadCellCol) = CellFormat("WELDING", LoadCellRow, LoadCellCol, 255, 230, 153, True, "xlMedium")
        WeldingSheet.Cells(LoadCellRow + 1, LoadCellCol) = CellFormat("WELDING", LoadCellRow + 1, LoadCellCol, 255, 230, 153, True, "xlMedium")
        WeldingSheet.Range(WeldingSheet.Cells(LoadCellRow, LoadCellCol), WeldingSheet.Cells(LoadCellRow + 1, LoadCellCol)).Merge
    Next i
End Function

Sub LoadsCellFormatUpdateAll()
    For i = StartWeek() To CurrentWeekNumber + FutureWeeks
        LoadsCellFormatWeek (i)
    Next i
End Sub

'NECESIDADES DE PRODUCCIÓN'
Function NeedsCellFormatWeek(week As Integer) As Integer
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Dim lastRow As Integer
    lastRow = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row + 2
    
    Dim NeedsCellCol As Integer
    Dim NeedsCellRow As Integer
    
    For i = OffsetFilaCabecera() + 1 To lastRow Step 3
        NeedsCellCol = WeldingWeekSearch(week) + 2
        NeedsCellRow = i
        WeldingSheet.Cells(NeedsCellRow, NeedsCellCol) = CellFormat("WELDING", NeedsCellRow, NeedsCellCol, 255, 242, 204, True, "xlMedium")
        WeldingSheet.Cells(NeedsCellRow + 1, NeedsCellCol) = CellFormat("WELDING", NeedsCellRow + 1, NeedsCellCol, 255, 242, 204, True, "xlMedium")
        WeldingSheet.Range(WeldingSheet.Cells(NeedsCellRow, NeedsCellCol), WeldingSheet.Cells(NeedsCellRow + 1, NeedsCellCol)).Merge
    Next i
End Function

Sub NeedsCellFormatUpdateAll()
    For i = StartWeek() To CurrentWeekNumber + FutureWeeks
        NeedsCellFormatWeek (i)
    Next i
End Sub

'PLAN DE PRODUCCIÓN'
Sub PlanCellFormatWeek(week As Integer)
    
End Sub

Sub prueba()
    Call PlanCellFormatWeek(2)
End Sub

'ACTUALIZA TODAS DE GOLPE PARA EL PROCEDIMIENTO UPDATE&CLEAR DEL MODULO 5'
Sub ActualLoadsNeedsFormatUpdateAll()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    
    Dim lastRow As Integer
    lastRow = WeldingSheet.Cells(Rows.Count, NumColWelding("Line")).End(xlUp).Row + 2
    
    Dim ActualCellCol As Integer
    Dim ActualCellRow As Integer
    Dim LoadCellCol As Integer
    Dim LoadCellRow As Integer
    Dim NeedsCellCol As Integer
    Dim NeedsCellRow As Integer
    Dim week As Integer
    
    For j = StartWeek() To CurrentWeekNumber() + FutureWeeks
        week = j
        For i = OffsetFilaCabecera() + 1 To lastRow Step 3
            ActualCellCol = WeldingWeekSearch(week)
            ActualCellRow = i
            WeldingSheet.Cells(ActualCellRow, ActualCellCol) = CellFormat("WELDING", ActualCellRow, ActualCellCol, 255, 255, 0, True, "xlMedium")
            WeldingSheet.Cells(ActualCellRow + 1, ActualCellCol) = CellFormat("WELDING", ActualCellRow + 1, ActualCellCol, 255, 255, 0, True, "xlMedium")
            WeldingSheet.Range(WeldingSheet.Cells(ActualCellRow, ActualCellCol), WeldingSheet.Cells(ActualCellRow + 1, ActualCellCol)).Merge
            LoadCellCol = WeldingWeekSearch(week) + 1
            LoadCellRow = i
            WeldingSheet.Cells(LoadCellRow, LoadCellCol) = CellFormat("WELDING", LoadCellRow, LoadCellCol, 255, 230, 153, True, "xlMedium")
            WeldingSheet.Cells(LoadCellRow + 1, LoadCellCol) = CellFormat("WELDING", LoadCellRow + 1, LoadCellCol, 255, 230, 153, True, "xlMedium")
            WeldingSheet.Range(WeldingSheet.Cells(LoadCellRow, LoadCellCol), WeldingSheet.Cells(LoadCellRow + 1, LoadCellCol)).Merge
            NeedsCellCol = WeldingWeekSearch(week) + 2
            NeedsCellRow = i
            WeldingSheet.Cells(NeedsCellRow, NeedsCellCol) = CellFormat("WELDING", NeedsCellRow, NeedsCellCol, 255, 242, 204, True, "xlMedium")
            WeldingSheet.Cells(NeedsCellRow + 1, NeedsCellCol) = CellFormat("WELDING", NeedsCellRow + 1, NeedsCellCol, 255, 242, 204, True, "xlMedium")
            WeldingSheet.Range(WeldingSheet.Cells(NeedsCellRow, NeedsCellCol), WeldingSheet.Cells(NeedsCellRow + 1, NeedsCellCol)).Merge
        Next i
    Next j
End Sub
