Attribute VB_Name = "Module5"
'MODULE_5'
Sub WeldingSheetClearUpdateAll()
    'ACTULIZACIÓN COMPLETA DE PESTAÑA WELDING BORRANDO EL CONTENIDO PREVIAMENTE. REFERENCIAS, EDI Y CABECERAS'
    Dim answer As Integer
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    answer = MsgBox("¿Desea borrar todo y actualizar la pestaña?", vbQuestion + vbYesNo, "Elegir opción")
    If answer = vbYes Then
        Answer2 = MsgBox("Este proceso borrará todos los datos y los volverá a cargar en la hoja. Esto podría demorarse. ¿Desea continuar?", vbQuestion + vbYesNo, "Elegir opción")
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
                MsgBox "Finalizado con éxito"
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
'    'ACTUALIZACIÓN CABECERAS, REFERENCIAS Y DEMANDAS EDI'
'    Dim Answer As Integer
'    Answer = MsgBox("¿Desea actualizar los valores de las demandas semanales?", vbQuestion + vbYesNo, "Elegir opción")
'    If Answer = vbYes Then
'        WeldingHeadersUpdate
'        ImportEDI
'    Else
'        WeldingHeadersUpdate
'        'ImportEDI
'    End If
'    MsgBox "Finalizado con éxito"
'End Sub
'
'Sub WeldingSheetUpdateAll()
'    'ACTULIZACIÓN COMPLETA DE PESTAÑA WELDING. REFERENCIAS, EDI Y CABECERAS'
'    Dim Answer As Integer
'    Answer = MsgBox("¿Desea actualizar todo?", vbQuestion + vbYesNo, "Elegir opción")
'    If Answer = vbYes Then
'        WeeksHeaders
'        WeldingReferences
'        ImportEDI
'        WeekProdNeedUpdateAll
'        Welding_backup
'        WeldingBackupToWelding
'        MsgBox "Finalizado con éxito"
'    Else
'    End If
'End Sub

'Sub WeldingSheetClearUpdateAll()
'    'ACTULIZACIÓN COMPLETA DE PESTAÑA WELDING BORRANDO EL CONTENIDO PREVIAMENTE. REFERENCIAS, EDI Y CABECERAS'
'    Dim Answer As Integer
'    Dim WeldingSheet As Worksheet
'    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
'    Answer = MsgBox("¿Desea borrar todo y actualizar la pestaña?", vbQuestion + vbYesNo, "Elegir opción")
'    If Answer = vbYes Then
'        Answer2 = MsgBox("Este proceso borrará todos los datos y los volverá a cargar en la hoja. Esto podría demorarse. ¿Desea continuar?", vbQuestion + vbYesNo, "Elegir opción")
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
'                MsgBox "Finalizado con éxito"
'            Else
'            End If
'    Else
'    End If
'End Sub
End Sub

