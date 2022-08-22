Attribute VB_Name = "ActualizarFormato"
Sub ActualizarFormato()
    
    Dim palette(9) As String
    Dim i As Integer
    Dim Conti As Integer
    Dim TableName As String
    Dim FilterSet As String
    
    Application.ScreenUpdating = False
    
    'CÓDIGO DE COLORES USADO
    palette(1) = 35
    palette(3) = 36
    palette(5) = 20
    palette(7) = 39
    palette(9) = 40
    i = 0
    
    Call Locate_Positions_CP
    
    Call ClearFilters
    
    Sheets(SheetName).Cells(CPAuxi, CPsupplierj).Select                     'Selecciona una celda dentro de la tabla donde aplicar el filtro.
    TableName = ActiveSheet.ListObjects(1).Name                             'Selecciona el nombre de la primera tabla en la hoja activa.
    FilterSet = Sheets(SheetName).Cells(CPsupplieri, CPsupplierj).Value     'Posición del encabezado donde aplicar el filtro.
    Call AlfabeticOrder(SheetName, TableName, FilterSet)
         
    Call Mayus_Clean(1, CPsupplierj)
    Call Mayus_Clean(2, CPlanguagej)

    For Conti = CPAuxi To CPendi    'Rellenar celdas con colores
        
        Application.StatusBar = "Format Progress (3/4): " & Conti - CPAuxi & " of " & CPendi - CPAuxi & ": " & Format((Conti - CPAuxi) / (CPendi - CPAuxi), "0%")
        
        If i Mod 2 = 0 Then
        
            palette(i) = 2
            
        End If
        
        ws_contact.Range(NC_CPvendorcodej & Conti & ":" & NC_CPOKj & Conti).Interior.ColorIndex = palette(i)
        
        If ws_contact.Cells(Conti + 1, CPsupplierj).Value = ws_contact.Cells(Conti, CPsupplierj).Value And Conti <= CPendi Then
            
            ws_contact.Range(NC_CPvendorcodej & Conti + 1 & ":" & NC_CPOKj & Conti + 1).Interior.ColorIndex = palette(i) 'Se agrupan las líneas por colores y por proveedor
            
        Else
        
            i = i + 1
            
        End If
        
        If i > 9 Then
        
            i = 0
            
        End If
        
         If ws_contact.Cells(Conti, CPmailj).Value <> "" And ws_contact.Cells(Conti, CPOKj).Value <> "" And Conti <= CPendi Then
        
            ws_contact.Cells(Conti, CPOKj).Value = ""  'Si se había marcado como que no tenía contacto se limpia el estado al añadir el correo
            
        End If
        
    Next
    
    For Conti = CPAuxi To CPendi    'Comprobar si hay información de contacto
        
        Application.StatusBar = "Format Progress (4/4): " & Conti - CPAuxi & " of " & CPendi - CPAuxi & ": " & Format((Conti - CPAuxi) / (CPendi - CPAuxi), "0%")
        
        If (ws_contact.Cells(Conti, CPmailj).Value = "" And Conti <= CPendi) Then
            
            If ws_contact.Cells(Conti, CPOKj).Value = "" Then
                ws_contact.Cells(Conti, CPOKj).Value = "Falta información del proveedor"
            End If
            
            ws_contact.Range(NC_CPvendorcodej & Conti & ":" & NC_CPOKj & Conti).Interior.ColorIndex = 3
            
        End If
    
    Next
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
End Sub

Function Mayus_Clean(Process As Integer, Field As Integer)
'Poner en mayúscula y eliminar espacios innecesarios la columna elegida.
    
    Dim Conti As Integer

    For Conti = CPAuxi To CPendi

        Application.StatusBar = "Format Progress (" & Process & "/4): " & Conti - CPAuxi & " of " & CPendi - CPAuxi & ": " & Format((Conti - CPAuxi) / (CPendi - CPAuxi), "0%")
        
        ws_contact.Cells(Conti, Field).Value = UCase(ws_contact.Cells(Conti, Field).Value)
        ws_contact.Cells(Conti, Field).Value = Trim(ws_contact.Cells(Conti, Field).Value)
    
    Next
    
End Function

