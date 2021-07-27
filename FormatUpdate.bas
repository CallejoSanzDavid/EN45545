Attribute VB_Name = "FormatUpdate"
Sub ActualizarFormato()
    
    Dim paleta(9) As String
    Dim i As Integer
    Dim Conti As Integer
    Dim CPmailj As Integer
    
    Application.ScreenUpdating = False
    
    Call Locate_Positions_OG
    
    'CÓDIGO DE COLORES USADO
    paleta(1) = 35
    paleta(3) = 36
    paleta(5) = 20
    paleta(7) = 39
    paleta(9) = 40
    i = 0
    
    With ActiveSheet.ListObjects("Tabla1").Sort         'Ordenar alfabeticamente una tabla
        
        .SortFields.Clear                               'Elimina los filtros activos
        .SortFields.Add Key:=Range("Tabla1[Supplier]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal        'Selecciona el filtro y donde aplicarlo
        .Apply                                          'Aplicar
    
    End With
    
    CPmailj = ws_contact.Range("A1:Z1").Find("Mail").Column
    
    For Conti = CPAuxi To CPendi  'Poner en mayúscula y eliminar espacios innecesarios la columna de supplier
    
        Application.StatusBar = "Format Progress (1/4): " & Conti - CPAuxi & " of " & CPendi - CPAuxi & ": " & Format((Conti - CPAuxi) / (CPendi - CPAuxi), "0%")
        
        ws_contact.Cells(Conti, CPsupplierj).Value = UCase(ws_contact.Cells(Conti, CPsupplierj).Value)
        ws_contact.Cells(Conti, CPsupplierj).Value = Trim(ws_contact.Cells(Conti, CPsupplierj).Value)
    
    Next
    
    For Conti = CPAuxi To CPendi  'Poner en mayúscula y eliminar espacios innecesarios la columna de language
    
        Application.StatusBar = "Format Progress (2/4): " & Conti - CPAuxi & " of " & CPendi - CPAuxi & ": " & Format((Conti - CPAuxi) / (CPendi - CPAuxi), "0%")
        
        ws_contact.Cells(Conti, CPsupplierj).Value = UCase(ws_contact.Cells(Conti, CPsupplierj).Value)
        ws_contact.Cells(Conti, CPsupplierj).Value = Trim(ws_contact.Cells(Conti, CPsupplierj).Value)
    
    Next

    For Conti = CPAuxi To CPendi    'Rellenar celdas con colores
        
        Application.StatusBar = "Format Progress (3/4): " & Conti - CPAuxi & " of " & CPendi - CPAuxi & ": " & Format((Conti - CPAuxi) / (CPendi - CPAuxi), "0%")
        
        If i Mod 2 = 0 Then
        
            paleta(i) = 2
            
        End If
        
        If ws_contact.Cells(Conti, CPmailj).Value <> "" And ws_contact.Cells(Conti, CPOKj).Value = "NO HAY CONTACTO" And Conti <= CPendi Then
        
            ws_contact.Cells(Conti, CPOKj).Value = ""  'Si se había marcado como que no tenía contacto se limpia el estado al añadir el correo
            
        End If
        
        ws_contact.Range(NC_CPvendorcodej & Conti & ":" & NC_CPOKj & Conti).Interior.ColorIndex = paleta(i)
        
        If ws_contact.Cells(Conti + 1, CPsupplierj).Value = ws_contact.Cells(Conti, CPsupplierj).Value And Conti <= CPendi Then
            
            ws_contact.Range(NC_CPvendorcodej & Conti + 1 & ":" & NC_CPOKj & Conti + 1).Interior.ColorIndex = paleta(i) 'Se agrupan las líneas por colores y por proveedor
            
        Else
        
            i = i + 1
            
        End If
        
        If i > 9 Then
        
            i = 0
            
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


