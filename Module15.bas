Attribute VB_Name = "Module15"
'MODULE_15'
'CREACIÓN DE LA PESTAÑA WELDING COMPLETA'
Sub WeldingSheetCreateComplete()
    'Crea la pestaña de WELDING completa en blanco'
    'NO CREA BACKUP PREVIO'
    Dim answer As Integer
    answer = MsgBox("¿Desea rellenar la pestaña Welding completa?", vbQuestion + vbYesNo, "Elegir opción")
        If answer = vbYes Then
        Dim Answer2
        Answer2 = MsgBox("Este proceso puede tardar unos minutos. ¿Desea continuar?", vbQuestion + vbYesNo, "Elegir opción")
        If Answer2 = vbYes Then
            Application.ScreenUpdating = False
            WeeksHeaders
            WeldingReferences
            ImportEDI
            WeekProdNeedUpdateAll
            WeekProdPlanUpdateAll
            CompleteFormat
            Application.ScreenUpdating = True
            Else
            End If
        Else
        End If
        MsgBox "Finalizado con éxito"
End Sub
