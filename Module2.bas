Attribute VB_Name = "Module2"
'MODULE_2'
'LISTADO DE WIPS'
Sub WIPSList()
    Dim Reference As String
    Dim SearchReference As String
    Dim tableRange As Range
    Dim ReferencesSheet As Worksheet
    Set ReferencesSheet = ThisWorkbook.Worksheets(SheetName("References"))
    'Contadores para bucles'
    Dim n As Integer, m As Integer, k As Integer
    k = 1
   
    'Rango de hoja usado. Contiene todas las celdas con valores'
    Dim FullRange As Range
   
    'Calculamos el rango que ocupa la hoja completa con todas las celdad en uso'
    Set FullRange = ReferencesSheet.UsedRange
   
    'Seleccionamos la referencia con la que trabajar'
    SearchReference = Application.InputBox(prompt:="Indique la referencia:", Type:=2, Title:="BÚSQUEDA DE REFERENCIA")
    Set FoundRange = FullRange.Find(SearchReference, LookIn:=xlValues, LookAt:=xlPart)
    
    If FoundRange Is Nothing Then
        MsgBox "No se encontró ninguna referencia que contenga la cadena '" & SearchReference & "'."
        Exit Sub
    Else
        Reference = "Table_" & FoundRange.Value
    End If
    
    'Trabajamos con la referencia deseada'
    Set tableRange = ReferencesSheet.ListObjects(Reference).Range
   
    'Agrupamos todos los procesos por nivel'
    Dim LevelProcess As Object
    Set LevelProcess = CreateObject("Scripting.Dictionary")
   
    'Agregamos los procesos al diccionario'
    n = tableRange.Rows.Count
    While k < n
        Level = tableRange.Cells(k + 1, NumCol("Level")).Value
        If Not LevelProcess.Exists(Level) Then
            LevelProcess.Add Level, ""
        End If
        LevelProcess(Level) = LevelProcess(Level) & "Reference " & tableRange.Cells(k + 1, NumCol("Reference")).Value & " Proceso: " & tableRange.Cells(k + 1, NumCol("Proceso")) & " Línea/s: " & tableRange.Cells(k + 1, NumCol("Linea")).Value & vbCrLf
        k = k + 1
    Wend
 
    'Mostramos los procesos agrupados por Level'
    Dim message As String
    For Each Level In LevelProcess.Keys
        message = message & "Level " & Level & ":" & vbCrLf & LevelProcess(Level) & vbCrLf
    Next Level
    MsgBox message, Title:="Reference: " & FoundRange.Value
 
 
End Sub
