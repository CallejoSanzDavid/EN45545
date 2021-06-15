Attribute VB_Name = "InfoProveedores_SAP"
Sub BaseProveedores()      'Busca y registra la información de contacto.

    Dim supplj As Integer
    Dim supplier As String
    Dim m As Integer
    Dim N As Integer
    Dim ContarDBi As Integer
    Dim InfoContj As Integer
    Dim Inicioi As Integer
    Dim InicioConti As Integer
    Dim mailj As Integer
    Dim c As Range
    Dim a As Integer
    
    a = MsgBox("Para el correcto funcionamiento de la función, asegúrese de estar registrado en SAP y tener la ventana inicial abierta (SAP Easy Access)." + vbCrLf + vbCrLf + "Para evitar interrupciones en el programa, pulse el último icono de la parte superior y seleccione 'Options...'. Dentro de 'Accesibility & Scripting' > 'Scripting' > 'User Settings' desleccione las notificaciones. Deje activa la opción 'Enable scripting'", vbOKCancel)
    
    If a = 2 Then
    
        Exit Sub
    
    End If
       
    Application.ScreenUpdating = False
    
    SheetName = ActiveSheet.name
    
    'Localizamos las posiciones en Articles
    codej = Sheets(SheetName).Range("A1:Z4").Find("KB article number").Column
    supplj = Sheets(SheetName).Range("A1:Z4").Find("Supplier").Column
    InfoContj = Sheets(SheetName).Range("A1:Z4").Find("Contact Info").Column
    Inicioi = Sheets(SheetName).Range(Cells(1, supplj), Cells(10, supplj)).Find("Supplier").Row + 1
    N = Sheets(SheetName).Cells(Rows.Count, supplj).End(xlUp).Row
    
    'Localizamos las posiciones en Información de Contacto
    Sheets("Información de Contacto").Activate
    
    InicioConti = Sheets("Información de Contacto").Range("A1:Z4").Find("Supplier").Row + 1
    merakj = Sheets("Información de Contacto").Range(Cells(InicioConti - 1, 1), Cells(InicioConti - 1, 10)).Find("Vendor Code").Column
    supplierj = Sheets("Información de Contacto").Range(Cells(InicioConti - 1, 1), Cells(InicioConti - 1, 10)).Find("Supplier").Column
    mailj = Sheets("Información de Contacto").Range(Cells(InicioConti - 1, 1), Cells(InicioConti - 1, 10)).Find("Mail").Column
    tlfnoj = Sheets("Información de Contacto").Range(Cells(InicioConti - 1, 1), Cells(InicioConti - 1, 10)).Find("Telephone").Column
    countryj = Sheets("Información de Contacto").Range(Cells(InicioConti - 1, 1), Cells(InicioConti - 1, 10)).Find("Country").Column
    languagej = Sheets("Información de Contacto").Range(Cells(InicioConti - 1, 1), Cells(InicioConti - 1, 10)).Find("Language").Column
    ContarDBi = Sheets("Información de Contacto").Cells(Rows.Count, supplierj).End(xlUp).Row
    
    Sheets(SheetName).Activate
    
    For m = Inicioi To N
        
        Application.StatusBar = "Updating Supplier's Contact Information: " & m - Inicioi - 1 & " of " & N - Inicioi - 1 & ": " & Format((m - Inicioi - 1) / (N - Inicioi - 1), "0%")
        
        supplier = Sheets(SheetName).Cells(m, supplj).Value
        
        If supplier <> "" Then
            
            m = LocateSupplier(SheetName, codej, InicioConti, merakj, supplierj, mailj, tlfnoj, countryj, languagej, ContarDBi, supplier, m, InfoContj)
            
        Else
            
            InfoUpdated = ME2M_SAP_SUPPLIER_CONTACT(SheetName, codej, m, InfoContj, InicioConti, merakj, supplierj, mailj, tlfnoj, countryj, languagej, ContarDBi)
            
            Sheets(SheetName).Cells(m, InfoContj) = "Does NOT Exist"
            Sheets(SheetName).Cells(m, InfoContj).Interior.ColorIndex = 3
            
            'Copy Paste Module
            finalj = 17 'Columna Q
            Sheets(SheetName).Range(Cells(m, codej), Cells(m, finalj)).Copy
            
            Sheets("No SAP Info").Activate
            
            auxfinali = Sheets("No SAP Info").Cells(Rows.Count, codej).End(xlUp).Row + 1
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            'ActiveSheet.ListObjects(1).Resize Range(Cells(Inicioi, codej), Cells(auxfinali, finalj)) 'Ampliamos el rango de la tabla para que añada la nueva línea
            
            Sheets(SheetName).Activate
        
        End If
    
    Next
    
    Application.StatusBar = ""
    On Error GoTo 0
    Application.ScreenUpdating = True
    
End Sub

Function LocateSupplier(SheetName, codej, InicioConti, merakj, supplierj, mailj, tlfnoj, countryj, languagej, ContarDBi, supplier, m, InfoContj)   'Localizamos la posición del proveedor en la hoja de Información de Contacto
    
    Dim linea As Integer
    Dim InfoUpdated As Integer
    
    InfoUpdated = 0
    
    Set c = Range(Sheets("Información de Contacto").Cells(InicioConti, supplierj), Sheets("Información de Contacto").Cells(ContarDBi, supplierj)).Find(supplier)
    
    If c Is Nothing Then 'No existe el proveedor en la BBDD de Contactos
    
        linea = 0
        
        InfoUpdated = ME2M_SAP_SUPPLIER_CONTACT(SheetName, codej, m, InfoContj, InicioConti, merakj, supplierj, mailj, tlfnoj, countryj, languagej, ContarDBi)
        
        If InfoUpdated = 0 Then
        
            Sheets(SheetName).Cells(m, InfoContj) = "Does NOT Exist"
            Sheets(SheetName).Cells(m, InfoContj).Interior.ColorIndex = 3
            
            'Copy Paste Module
            finalj = 17 'Columna Q
            Sheets(SheetName).Range(Cells(m, codej), Cells(m, finalj)).Copy
            
            Sheets("No SAP Info").Activate
            
            auxfinali = Sheets("No SAP Info").Cells(Rows.Count, codej).End(xlUp).Row + 1
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            'ActiveSheet.ListObjects(1).Resize Range(Cells(Inicioi, codej), Cells(auxfinali, finalj)) 'Ampliamos el rango de la tabla para que añada la nueva línea
            
            Sheets(SheetName).Activate
            
        End If
        
    Else
    
        linea = c.Row
        
        If Sheets("Información de Contacto").Cells(linea, mailj) = "" Then  'Existe el proveedor en la BBDD de Contactos pero no tiene info de contacto
            
            InfoUpdated = ME2M_SAP_SUPPLIER_CONTACT(SheetName, codej, m, InfoContj, InicioConti, merakj, supplierj, mailj, tlfnoj, countryj, languagej, ContarDBi)
            
            If InfoUpdated = 0 Then
                
                Sheets(SheetName).Cells(m, InfoContj) = "Does NOT Exist"
                Sheets(SheetName).Cells(m, InfoContj).Interior.ColorIndex = 3
                
                linea = 0
                
                'Copy Paste Module
                finalj = 17 'Columna Q
                Sheets("No SAP Info").Activate
                
                auxfinali = Sheets("No SAP Info").Cells(Rows.Count, codej).End(xlUp).Row + 1
                Cells(auxfinali, "A").Select
                ActiveSheet.Paste
                
                'ActiveSheet.ListObjects(1).Resize Range(Cells(Inicioi, codej), Cells(auxfinali, finalj)) 'Ampliamos el rango de la tabla para que añada la nueva línea
                
                Sheets(SheetName).Activate
                
            End If
            
        End If
        
    End If
    
    If linea <> 0 Then
            
        Sheets(SheetName).Cells(m, InfoContj) = Sheets("Información de Contacto").Cells(linea, mailj)
        Sheets(SheetName).Cells(m, InfoContj).Interior.ColorIndex = 43
         
    End If
    
    If InfoUpdated = 1 Then
        
        'Actualizamos la información de contacto vinculada
        Sheets("Información de Contacto").Activate
        ActiveWorkbook.RefreshAll
        'Relocalizamos el final de la BB.DD. de contactos
        ContarDBi = Sheets("Información de Contacto").Cells(Rows.Count, supplierj).End(xlUp).Row
        Sheets(SheetName).Activate
        LocateSupplier = m - 1       'Si se actualiza la info volvemos a hacer la consulta.
    
    Else
    
        LocateSupplier = m
    
    End If

End Function

Function ME2M_SAP_SUPPLIER_CONTACT(SheetName, codej, m, InfoContj, InicioConti, merakj, supplierj, mailj, tlfnoj, countryj, languagej, ContarDBi)        'Obtiene información de la función ME2M de SAP: CONTACTO Y NOMBRE DEL ÚLTIMO PROVEEDOR (Last Supplier's Contact And Name).

    Dim nproductoj As Integer
    Dim nproducto As String
    Dim mark As Integer
    Dim InfoUpdated As Integer
    Dim meraki As Integer
    Dim EndDBi As Integer
    Dim contacto As String
    Dim language As String
    Dim saveflag As Integer
    Dim obj As Object
       
    'Cabecera para tener ya abierto SAP
    Set SapGuiAuto = GetObject("SAPGUI")        'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)     'Get the first system that is currently connected
    Set Session = SAPCon.Children(0)    'Get the first session (window) on that connection
    'Fin de la cabecera
    
    nproductoj = codej
    nproducto = Sheets(SheetName).Cells(m, nproductoj).Value
    saveflag = 0
    
    'Ejecuto transacción ME2M
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "ME2M"
    Session.findById("wnd[0]").sendVKey 0
    
    'Busco PartNumber
    Session.findById("wnd[0]/usr/ctxtEM_MATNR-LOW").Text = nproducto
    'Busco en planta ES20
    Session.findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = "ES20"
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    On Error GoTo ErrorHandler1:
    'Seleccionamos la celda del último proveedor al que hemos comprado el material
    Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "SUPERFIELD"
    Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 1
    Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
    
    On Error GoTo ErrorHandler2:
    'En caso de que se soluciones el error se vuelve a ejecutar desde este punto
ErrorFixed:
    'Obtenemos el nombre del último proveedor
    Name1 = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text
    Name2 = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME2").Text
    
    If Name2 <> "" Then
            
        supplier = Name1 & " " & Name2
        
    Else
    
        supplier = Name1
        
    End If
       
    nombre_RecordSheet = ActiveWorkbook.name
    'Abrir Excel BBDD de contactos
    Workbooks.Open (Sheets("AUX1").Range("H3").Value)
    nombre_bbdd = ActiveWorkbook.name
    Workbooks(nombre_bbdd).Sheets("OG PROVEEDORES").Activate          'Activamos el libro en la hoja de registro
    
    meraki = Sheets("OG PROVEEDORES").Range(Cells(1, supplierj), Cells(4, supplierj)).Find("Supplier").Row
    
    EndDBi = Sheets("OG PROVEEDORES").Cells(Rows.Count, supplierj).End(xlUp).Row
    
    Set c = Range(Sheets("OG PROVEEDORES").Cells(meraki + 1, supplierj), Sheets("OG PROVEEDORES").Cells(EndDBi, supplierj)).Find(supplier)
    
    If c Is Nothing Then 'No existe el proveedor en la BBDD de Contactos
        
        saveflag = 1
        
        i = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row + 1            'Localizamos la última fila con info en una columna sin celdas combinadas
        
        'Código de proveedor
        Sheets("OG PROVEEDORES").Cells(i, merakj).Value = Session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").Text
        
        'Nombre del proveedor
        Sheets("OG PROVEEDORES").Cells(i, supplierj).Value = supplier
        
        'Teléfono de contacto
        Sheets("OG PROVEEDORES").Cells(i, tlfnoj).Value = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-TEL_NUMBER").Text
        
        'Pais del proveedor
        Sheets("OG PROVEEDORES").Cells(i, countryj).Value = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").Text
        
        'Idioma de contacto preferido
        Sheets("OG PROVEEDORES").Cells(i, languagej).Value = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbADDR1_DATA-LANGU").Text
        
        'Correo de contacto
        contacto = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").Text
        
        If contacto <> "" Then
            
            Sheets("OG PROVEEDORES").Cells(i, mailj).Value = contacto
            ME2M_SAP_SUPPLIER_CONTACT = 1
            
        Else
        
            ME2M_SAP_SUPPLIER_CONTACT = 0
            
        End If
        
    Else            'Existe el proveedor en la BBDD de Contactos
        
        'Código de proveedor
        If Sheets("OG PROVEEDORES").Cells(c.Row, merakj).Value = "" Then
            Sheets("OG PROVEEDORES").Cells(c.Row, merakj).Value = Session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").Text
            saveflag = 1
        End If
        
        'Teléfono de contacto
        If Sheets("OG PROVEEDORES").Cells(c.Row, tlfnoj).Value = "" Then
            Sheets("OG PROVEEDORES").Cells(c.Row, tlfnoj).Value = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-TEL_NUMBER").Text
            saveflag = 1
        End If
        
        'Pais del proveedor e Idioma de contacto preferido
        If Sheets("OG PROVEEDORES").Cells(c.Row, countryj).Value = "" Then
            Sheets("OG PROVEEDORES").Cells(c.Row, countryj).Value = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").Text
            saveflag = 1
        End If
        
        'Idioma de contacto preferido
        If Sheets("OG PROVEEDORES").Cells(c.Row, languagej).Value = "" Then
            Sheets("OG PROVEEDORES").Cells(c.Row, languagej).Value = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbADDR1_DATA-LANGU").Text
            saveflag = 1
        End If
        
        'Correo de contacto
        contacto = Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").Text
        
        If contacto <> "" And Sheets("OG PROVEEDORES").Cells(c.Row, mailj).Value = "" Then
            
            Sheets("OG PROVEEDORES").Cells(c.Row, mailj).Value = contacto
            ME2M_SAP_SUPPLIER_CONTACT = 1     'Devuelve 1 si se ha actualizado la info
            saveflag = 1
            
        Else
        
            ME2M_SAP_SUPPLIER_CONTACT = 0     'Devuelve 0 si no se ha actualizado el correo de contacto
            
        End If
        
    End If
    
    If saveflag = 1 Then
        
        'Ejecutar Macro en otro libro
        Application.Workbooks(nombre_bbdd).Sheets("OG PROVEEDORES").Activate
        Application.Run "'" & nombre_bbdd & "'!" & "ActualizarFormato.ActualizarFormato"        '   "Nombre del módulo.Nombre de la macro"
        
        ActiveWorkbook.Save
        
    End If
    
    ActiveWorkbook.Close
    
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
    Session.findById("wnd[0]/tbar[0]/btn[15]").press
            
ErrorHandler1:
    If Err.Number <> 0 Then
            
        'Cierro transacción de SAP en caso de error
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        Resume ErrorOK
        
    End If
    
ErrorHandler2:
    'En ocasiones hay dos lineas en blanco al principio. Con esto evitamos que no obtenga info cuando la hay.
    'PROVEEDOR
    If Err.Number <> 0 Then
        
        Resume Here:
Here:
        On Error GoTo ErrorHandler4:
        
        Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "SUPERFIELD"
        Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 2
        Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
        
        If Session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text <> "" Then
            
            On Error GoTo 0
            Err.Clear
            Resume ErrorFixed:        'Si se soluciona el error vuelve al código para actualizar la info
                
        End If
        
        If Err.Number = 619 Then        'Evita error en el que el proveedor no tiene info de contacto
ErrorHandler4:
            Session.findById("wnd[0]/tbar[0]/btn[15]").press
        
        End If
    
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        Session.findById("wnd[0]/tbar[0]/btn[15]").press
        
        Resume ErrorOK
    
    End If
    
ErrorOK:

    Err.Clear
    
End Function


