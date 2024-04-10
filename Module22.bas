Attribute VB_Name = "Module22"
'MODULE_22
Sub BoxSheetClearUpdateAll()
    'ACTUALIZACIÓN COMPLETA PESTAÑA BOX'
    'PRINCIPALMENTE PARA CAMBIOS DE REFERENCIA'
    'GUARDA TODO PREVIAMENTE EN BACKUP'
    Dim BoxSheet As Worksheet
    Set BoxSheet = ThisWorkbook.Worksheets(SheetName("BOX"))
    Dim answer As Integer
    answer = MsgBox("¿Desea borrar todo y actualizar la pestaña?", vbQuestion + vbYesNo, "Elegir opción")
    If answer = vbYes Then
        Answer2 = MsgBox("Este proceso borrará todos los datos y los volverá a cargar en la hoja. Esto podría demorarse. ¿Desea continuar?", vbQuestion + vbYesNo, "Elegir opción")
            If Answer2 = vbYes Then
                Box_backup
                BoxSheet.UsedRange.Clear
                BoxHeaders
                BoxReferences
                BoxWeeksBuilder
                BoxBackupToBox
                MsgBox "Finalizado con éxito"
            Else
            End If
    Else
    End If
End Sub
