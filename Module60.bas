Attribute VB_Name = "mODULE60"
Sub hideCol(initCol As Integer, lastCol As Integer, sheet As String)
    'Oculta las pestañas en el intervalo pasado por argumento'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName(sheet))
    Dim initColLetter As String
    Dim lastColLetter As String

    initColLetter = Split(Cells(1, initCol).Address, "$")(1)
    lastColLetter = Split(Cells(1, lastCol).Address, "$")(1)

    ws.Range(initColLetter & ":" & lastColLetter).EntireColumn.Hidden = True
End Sub

Sub UnhideCol(sheet As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName(sheet))

    ws.Cells.EntireColumn.Hidden = False
End Sub

Sub ChkRefTablesNames()
    'Verificación de nombres de tablas en pestaña REFERENCES'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("References"))

    Dim oldName As String
    Dim newName As String
    
    Dim tableRange As Range

    For i = 1 To ws.ListObjects.Count
        oldName = ws.ListObjects(i).name
        Set tableRange = ws.ListObjects(i).DataBodyRange
        newName = "Table_" & UCase(tableRange.Cells(1, 1).Value)
        If oldName <> newName Then
            ws.ListObjects(i).name = newName
        End If
    Next i

End Sub

Sub hideColWelding(initWeek As Integer, finalWeek As Integer)
    'Oculta las columnas deseadas en la pestaña WELDING'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("WELDING"))

    Dim initWeekCol As Integer
    Dim finalWeekCol As Integer
    
    Call UnhideCol("WELDING")
    initWeekCol = WeldingWeekSearch(initWeek)
    finalWeekCol = WeldingWeekSearch(finalWeek) + WeldingColDistance - 1

    Call hideCol(initWeekCol, finalWeekCol, "WELDING")
End Sub
    
Sub hideColWeldingFORM()
    'Manejo del formulario para ocultar columnas en la pestaña WELDING'
    Dim initWeek As Integer
    Dim finalWeek As Integer
    

    initWeek = InputBox("Introduzca la primera semana que desea ocultar: ")
    finalWeek = InputBox("Introduzca la última semana que desea ocultar: ")
    If (initWeek < finalWeek) Then
        Call hideColWelding(initWeek, finalWeek)
        MsgBox "Finalizado con éxito", , "Finalizado"
    Else
        MsgBox "La semana final no puede ser anterior a la semana inicial", , "ERROR"
        Dim answer As Integer
        answer = MsgBox("¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "ERROR")
        If (answer = vbYes) Then
            Call hideColWeldingFORM
        Else
            Exit Sub
        End If
    End If
End Sub

Sub hideColBox(initWeek As Integer, finalWeek As Integer)
    'Oculta las columnas deseadas en la pestaña WELDING'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("BOX"))

    Dim initWeekCol As Integer
    Dim finalWeekCol As Integer
    
    Call UnhideCol("BOX")
    initWeekCol = BoxWeekSearch(initWeek)
    finalWeekCol = BoxWeekSearch(finalWeek) + BoxColDistance - 1

    Call hideCol(initWeekCol, finalWeekCol, "BOX")
End Sub

Sub hideColBoxFORM()
    'Manejo del formulario para ocultar columnas en la pestaña WELDING'
    Dim initWeek As Integer
    Dim finalWeek As Integer
    

    initWeek = InputBox("Introduzca la primera semana que desea ocultar: ")
    finalWeek = InputBox("Introduzca la última semana que desea ocultar: ")
    If (initWeek < finalWeek) Then
        Call hideColBox(initWeek, finalWeek)
        MsgBox "Finalizado con éxito", , "Finalizado"
    Else
        MsgBox "La semana final no puede ser anterior a la semana inicial", , "ERROR"
        Dim answer As Integer
        answer = MsgBox("¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "ERROR")
        If (answer = vbYes) Then
            Call hideColBoxFORM
        Else
            Exit Sub
        End If
    End If
End Sub

Sub hideColBending(initWeek As Integer, finalWeek As Integer)
    'Oculta las columnas deseadas en la pestaña WELDING'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName("BENDING"))

    Dim initWeekCol As Integer
    Dim finalWeekCol As Integer
    
    Call UnhideCol("BENDING")
    initWeekCol = BendingWeekSearch(initWeek)
    finalWeekCol = BendingWeekSearch(finalWeek) + BendingColDistance - 1

    Call hideCol(initWeekCol, finalWeekCol, "BENDING")
End Sub

Sub hideColBendingFORM()
    'Manejo del formulario para ocultar columnas en la pestaña WELDING'
    Dim initWeek As Integer
    Dim finalWeek As Integer
    

    initWeek = InputBox("Introduzca la primera semana que desea ocultar: ")
    finalWeek = InputBox("Introduzca la última semana que desea ocultar: ")
    If (initWeek < finalWeek) Then
        Call hideColBending(initWeek, finalWeek)
        MsgBox "Finalizado con éxito", , "Finalizado"
    Else
        MsgBox "La semana final no puede ser anterior a la semana inicial", , "ERROR"
        Dim answer As Integer
        answer = MsgBox("¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "ERROR")
        If (answer = vbYes) Then
            Call hideColBendingFORM
        Else
            Exit Sub
        End If
    End If
End Sub
