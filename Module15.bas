Attribute VB_Name = "Module15"
'MODULE_15'
'CREACI�N DE LA PESTA�A WELDING COMPLETA'
Sub WeldingSheetCreateComplete()
    'Crea la pesta�a de WELDING completa en blanco'
    'NO CREA BACKUP PREVIO'
    Dim answer As Integer
    answer = MsgBox("�Desea rellenar la pesta�a Welding completa?", vbQuestion + vbYesNo, "Elegir opci�n")
        If answer = vbYes Then
        Dim Answer2
        Answer2 = MsgBox("Este proceso puede tardar unos minutos. �Desea continuar?", vbQuestion + vbYesNo, "Elegir opci�n")
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
        MsgBox "Finalizado con �xito"
End Sub
