Attribute VB_Name = "Module22"
'MODULE_22
Sub BoxSheetClearUpdateAll()
    'ACTUALIZACI�N COMPLETA PESTA�A BOX'
    'PRINCIPALMENTE PARA CAMBIOS DE REFERENCIA'
    'GUARDA TODO PREVIAMENTE EN BACKUP'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Dim answer As Integer
    answer = MsgBox("�Desea borrar todo y actualizar la pesta�a?", vbQuestion + vbYesNo, "Elegir opci�n")
    If answer = vbYes Then
        Answer2 = MsgBox("Este proceso borrar� todos los datos y los volver� a cargar en la hoja. Esto podr�a demorarse. �Desea continuar?", vbQuestion + vbYesNo, "Elegir opci�n")
            If Answer2 = vbYes Then
                Box_backup
                BoxSheet.UsedRange.Clear
                BoxHeaders
                BoxReferences
                BoxWeeksBuilder
                BoxBackupToBox
                MsgBox "Finalizado con �xito"
            Else
            End If
    Else
    End If
End Sub
