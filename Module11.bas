Attribute VB_Name = "Module11"
'MODULE_11'
'ACTUALIZA EL PLAN DE PRODUCCIÓN'
Sub WeekProdPlan()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))

    'Elección para actualizar todo o sólo la semana deseada'
    Dim answer As Integer
    answer = MsgBox("¿Desea actualizar todas las semanas?", vbQuestion + vbYesNo, "")
    If answer = vbYes Then
        For i = StartWeek() To CurrentWeekNumber() + FutureWeeks()
            ProdPlan (i)
        Next i
    Else
        Dim WeekSearch As Integer
        WeekSearch = Application.InputBox(prompt:="Indique la semana:", Type:=2, Title:="BÚSQUEDA DE SEMANA")
        ProdPlan (WeekSearch)
    End If
End Sub

Sub WeekProdPlanUpdateAll()
    Dim WeldingSheet As Worksheet
    Set WeldingSheet = ThisWorkbook.Worksheets(SheetName("Welding"))
    For i = StartWeek() To CurrentWeekNumber() + FutureWeeks()
        ProdPlan (i)
        WeldingAccumulated (i)
    Next i
End Sub
