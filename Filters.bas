Attribute VB_Name = "Filters"
Sub ClearFilters()       'Borra todos los filtros.

    If Sheets(SheetName).FilterMode Then Sheets(SheetName).ShowAllData

End Sub

'SheetName = ActiveSheet.Name
'Sheets(SheetName).Cells(i + 1, j).Select                   'Selecciona una celda dentro de la tabla donde aplicar el filtro.
'TableName = ActiveSheet.ListObjects(1).Name                'Selecciona el nombre de la primera tabla en la hoja activa.
'FilterSet = Sheets(SheetName).Cells(i, j).Value            'Posición del encabezado donde aplicar el filtro.

Sub AlfabeticOrder(SheetName As String, TableName As String, FilterSet As String)       'Filtro: Ordenar en orden alfabético.
    
    On Error GoTo ErrorHandler:
    
        ActiveWorkbook.Worksheets(SheetName).ListObjects(TableName).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(SheetName).ListObjects(TableName).Sort.SortFields.Add2 _
            Key:=Range(TableName & "[[#All],[" & FilterSet & "]]"), SortOn:= _
            xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets(SheetName).ListObjects(TableName).Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
ErrorHandler:
    
        On Error GoTo 0

End Sub



