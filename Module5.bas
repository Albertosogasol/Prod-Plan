Attribute VB_Name = "Module5"
'MODULE_5'
Sub WeldingSheetClearUpdateAll()
    'ACTULIZACI�N COMPLETA DE PESTA�A WELDING BORRANDO EL CONTENIDO PREVIAMENTE. REFERENCIAS, EDI Y CABECERAS'
    Dim answer As Integer
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    answer = MsgBox("�Desea borrar todo y actualizar la pesta�a?", vbQuestion + vbYesNo, "Elegir opci�n")
    If answer = vbYes Then
        Answer2 = MsgBox("Este proceso borrar� todos los datos y los volver� a cargar en la hoja. Esto podr�a demorarse. �Desea continuar?", vbQuestion + vbYesNo, "Elegir opci�n")
            If Answer2 = vbYes Then
                Welding_backup
                WeldingSheet.UsedRange.Clear
                WeeksHeaders
                WeldingReferences
                WeldingBackupToWelding
                ImportEDI
                WeekProdPlanUpdateAll
                WeekProdNeedUpdateAll
                CompleteFormat
                MsgBox "Finalizado con �xito"
            Else
            End If
    Else
    End If
    'Desactualizados'
'--------------------------------------------------------------------------------------'
'Sub WeldingHeadersUpdate()
'    WeeksHeaders
'    WeldingReferences
'End Sub
'
'Sub WeldingSheetUpdate()
'    'ACTUALIZACI�N CABECERAS, REFERENCIAS Y DEMANDAS EDI'
'    Dim Answer As Integer
'    Answer = MsgBox("�Desea actualizar los valores de las demandas semanales?", vbQuestion + vbYesNo, "Elegir opci�n")
'    If Answer = vbYes Then
'        WeldingHeadersUpdate
'        ImportEDI
'    Else
'        WeldingHeadersUpdate
'        'ImportEDI
'    End If
'    MsgBox "Finalizado con �xito"
'End Sub
'
'Sub WeldingSheetUpdateAll()
'    'ACTULIZACI�N COMPLETA DE PESTA�A WELDING. REFERENCIAS, EDI Y CABECERAS'
'    Dim Answer As Integer
'    Answer = MsgBox("�Desea actualizar todo?", vbQuestion + vbYesNo, "Elegir opci�n")
'    If Answer = vbYes Then
'        WeeksHeaders
'        WeldingReferences
'        ImportEDI
'        WeekProdNeedUpdateAll
'        Welding_backup
'        WeldingBackupToWelding
'        MsgBox "Finalizado con �xito"
'    Else
'    End If
'End Sub

'Sub WeldingSheetClearUpdateAll()
'    'ACTULIZACI�N COMPLETA DE PESTA�A WELDING BORRANDO EL CONTENIDO PREVIAMENTE. REFERENCIAS, EDI Y CABECERAS'
'    Dim Answer As Integer
'    Dim WeldingSheet As Worksheet
'    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
'    Answer = MsgBox("�Desea borrar todo y actualizar la pesta�a?", vbQuestion + vbYesNo, "Elegir opci�n")
'    If Answer = vbYes Then
'        Answer2 = MsgBox("Este proceso borrar� todos los datos y los volver� a cargar en la hoja. Esto podr�a demorarse. �Desea continuar?", vbQuestion + vbYesNo, "Elegir opci�n")
'            If Answer2 = vbYes Then
'                WeekProdNeedUpdateAll
'                WeekProdPlanUpdateAll
'                Welding_backup
'                WeldingSheet.UsedRange.Clear
'                WeeksHeaders
'                WeldingReferences
'                ActualCellFormatUpdateAll
'                LoadsCellFormatUpdateAll
'                NeedsCellFormatUpdateAll
'                ImportEDI
'                WeldingBackupToWelding
'                WeekProdNeedUpdateAll
'                WeekProdPlanUpdateAll
'                MsgBox "Finalizado con �xito"
'            Else
'            End If
'    Else
'    End If
'End Sub
End Sub

