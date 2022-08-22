Attribute VB_Name = "SAP_InfoProveedores"
Sub SAP_InfoProveedores()
'Busca en SAP y registra la información de contacto en la BB.DD. de contactos.
    
    Dim VendorCode As String
    Dim CPsupplier As String
    Dim Mail As String
    Dim Telephone As String
    Dim Country As String
    Dim Language As String
    Dim m As Integer
    Dim a As Integer
        
    a = MsgBox("Para el correcto funcionamiento de la función, asegúrese de estar registrado en SAP y tener la ventana inicial abierta (SAP Easy Access)." + vbCrLf + vbCrLf + "Para evitar interrupciones en el programa, pulse el último icono de la parte superior y seleccione 'Options...'. Dentro de 'Accesibility & Scripting' > 'Scripting' > 'User Settings' deseleccione las notificaciones. Deje activa la opción 'Enable scripting'.", vbOKCancel)
    
    If a = 2 Then
    
        Exit Sub
    
    End If
           
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    Call Locate_Positions_CP
    
    'Borrar todos los filtros que haya aplicados
    Call ClearFilters
    
    Call Locate_Positions_DDBB
    
    For m = CPsupplieri + 1 To CPendi
        
        Application.StatusBar = "Updating Supplier's Contact Information: " & m - CPsupplieri & " of " & CPendi - CPsupplieri & ": " & Format((m - CPsupplieri) / (CPendi - CPsupplieri), "0%")
        
        VendorCode = ws_contact.Cells(m, CPvendorcodej).Value
        CPsupplier = ws_contact.Cells(m, CPsupplierj).Value
        Mail = ws_contact.Cells(m, CPmailj).Value
        Telephone = ws_contact.Cells(m, CPtlfnoj).Value
        Country = ws_contact.Cells(m, CPcountryj).Value
        Language = ws_contact.Cells(m, CPlanguagej).Value

        If VendorCode = "" Or Mail = "" Or Telephone = "" Or Country = "" Or Language = "" Then
            
            Call LocateSupplier(m, CPsupplier)
            ws_contact.Activate
            
        End If
        
    Next
    
    Application.StatusBar = ""
    On Error GoTo 0
    Application.ScreenUpdating = True
    
End Sub

Function LocateSupplier(m As Integer, CPsupplier As String)
'Localizamos la posición del proveedor en la hoja de Información de Contacto
    
    Dim InfoUpdated As Integer
    Dim linea As Integer
    Dim PartNumber As String
    
    InfoUpdated = 0

    ws_OG.Activate
    Set c = Range(ws_OG.Cells(Auxi + 1, manufj), ws_OG.Cells(N, manufj)).Find(CPsupplier)
    
    If c Is Nothing Then 'No existe el proveedor en la BBDD de Contactos

        ws_contact.Cells(m, CPOKj).Value = "NO HAY PART NUMBER EN LA BBDD DE F&H"
        
    Else
    
        linea = c.Row
        PartNumber = ws_OG.Cells(linea, nprodj).Value
        
        If PartNumber <> "" Then        'Si no hay Part Number SAP se bloquea.
        
            InfoUpdated = ME2M_SAP_SUPPLIER_CONTACT(PartNumber, CPsupplier, m)
            
        End If
            
        Do While (linea <= N And InfoUpdated <> 1)
            
            'Si no encuentra la info con el primer Part Number encontrado se mete en el Loop para
            'comprobar la información con otros Part Numbers del mismo Proveedor.
            Do While ws_OG.Cells(linea, nprodj).Value = PartNumber
            
                linea = linea + 1
                
                If PartNumber = "" Then
                    Exit Do
                End If
                
            Loop
            
            Set c = Range(ws_OG.Cells(linea, manufj), ws_OG.Cells(N, manufj)).Find(CPsupplier)
            
            If Not c Is Nothing Then
                    
                linea = c.Row
                PartNumber = ws_OG.Cells(linea, nprodj).Value
                
                If PartNumber <> "" Then        'Si no hay Part Number SAP se bloquea.
                    
                    InfoUpdated = ME2M_SAP_SUPPLIER_CONTACT(PartNumber, CPsupplier, m)
                    
                End If
            
            Else
                
                Exit Do
                         
            End If
            
        Loop
        
        If InfoUpdated = 0 Then

            ws_contact.Cells(m, CPOKj).Value = "NO HAY INFO EN SAP"
        
        End If
        
    End If

End Function

Function ME2M_SAP_SUPPLIER_CONTACT(PartNumber As String, CPsupplier As String, m As Integer) As Integer
'Busca la información del proveedor con la función ME2M de SAP.

    'Cabecera para tener ya abierto SAP
    Set SapGuiAuto = GetObject("SAPGUI")        'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)             'Get the first system that is currently connected
    Set session = SAPCon.Children(0)            'Get the first session (window) on that connection
    'Fin de la cabecera
    
    Dim saveflag As Boolean
    Dim Name1 As String
    Dim Name2 As String
    Dim supplier As String
    
    saveflag = False
    ME2M_SAP_SUPPLIER_CONTACT = 0
    
    'Ejecuto transacción ME2M
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ME2M"
    session.findById("wnd[0]").sendVKey 0
    
    'Busco PartNumber
    session.findById("wnd[0]/usr/ctxtEM_MATNR-LOW").Text = PartNumber
    'Busco en planta ES20
    session.findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = "ES20"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    On Error GoTo ErrorHandler1:
    'Seleccionamos la celda del último proveedor al que hemos comprado el material
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "SUPERFIELD"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 1
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
    
    On Error GoTo ErrorHandler2:
    'En caso de que se soluciones el error se vuelve a ejecutar desde este punto
ErrorFixed:
    'Obtenemos el nombre del último proveedor
    Name1 = session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text
    Name2 = session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME2").Text
    
    If Name2 <> "" Then

        supplier = UCase(Name1 & " " & Name2)
        
    Else

        supplier = UCase(Name1)
        
    End If
    
    If supplier <> CPsupplier Then
    
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press

        ws_contact.Cells(m, CPOKj).Value = "El nombre del proveedor encontrado en SAP es diferente"
        
        ME2M_SAP_SUPPLIER_CONTACT = 2     'Devuelve 2 si no coinciden los nombres de los proveedores.
        Exit Function
    
    End If
    
    ws_contact.Activate          'Activamos el libro en la hoja de registro

    saveflag = RegisterInfo(m)
    
    Do While ws_contact.Cells(m, CPsupplierj) = ws_contact.Cells(m + 1, CPsupplierj).Value
    'Loop para rellenar la info en caso de haber varias líneas de contacto.
        
        m = m + 1
        
        If saveflag = True Then
            
            Call FillInfo(m)
            
        End If
        
    Loop
    
    If saveflag = True Then
        
        ME2M_SAP_SUPPLIER_CONTACT = 1
    
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press
            
ErrorHandler1:
    If Err.Number <> 0 Then
        
        'Cierro transacción de SAP en caso de error
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        Resume ErrorOK
        
    End If
    
ErrorHandler2:
    'En ocasiones hay dos líneas en blanco al principio. Con esto evitamos que no obtenga info cuando la hay.
    If Err.Number <> 0 Then
        
        Err.Clear
        Resume Here:
Here:
        On Error GoTo ErrorHandler3:

        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "SUPERFIELD"
        'Si no se escribe así se bloquea al encontrar otro error.
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 2
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
        
        If session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").Text <> "" Then
            
            Err.Clear
            Resume ErrorFixed:        'Si se soluciona el error vuelve al código para actualizar la info
                
        End If
        
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        
        Resume ErrorOK
        
    End If

ErrorHandler3:
    'En ocasiones encuentra información sobre el part number pero no hay información del proveedor.
    If Err.Number <> 0 Then
        
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        Resume ErrorOK
    
    End If

ErrorOK:

    Err.Clear
    
End Function

Function RegisterInfo(m As Integer) As Boolean
'Registra la información del proveedor encontrada con la función ME2M de SAP: VENDOR CODE, INFO DE CONTACTO, PAIS E IDIOMA DE PREFERENCIA (Last Supplier's Contact And Name).
    
    'Cabecera para tener ya abierto SAP
    Set SapGuiAuto = GetObject("SAPGUI")        'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine  'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)     'Get the first system that is currently connected
    Set session = SAPCon.Children(0)    'Get the first session (window) on that connection
    'Fin de la cabecera
    
    Dim auxvendorcode As String
    Dim auxcontact As String
    Dim auxtelephone As String
    Dim auxcountry As String
    Dim auxlanguage As String

    RegisterInfo = False
    
    'Código de proveedor
    auxvendorcode = Trim(session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").Text)
        
    If auxvendorcode <> "" And ws_contact.Cells(m, CPvendorcodej).Value = "" Then
        
        ws_contact.Cells(m, CPvendorcodej).Value = auxvendorcode
        RegisterInfo = True     'Devuelve True si se ha actualizado la info
    
    End If
    
    'Correo de contacto
    auxcontact = Trim(session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").Text)
    
    If auxcontact <> "" And ws_contact.Cells(m, CPmailj).Value = "" Then
        
        ws_contact.Cells(m, CPmailj).Value = auxcontact
        RegisterInfo = True     'Devuelve True si se ha actualizado la info
         
    End If
    
    'Teléfono de contacto
    auxtelephone = Trim(session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-TEL_NUMBER").Text)
    
    If auxtelephone <> "" And ws_contact.Cells(m, CPtlfnoj).Value = "" Then
        
        ws_contact.Cells(m, CPtlfnoj).Value = auxtelephone
        RegisterInfo = True     'Devuelve True si se ha actualizado la info
        
    End If
    
    'Pais del proveedor
    auxcountry = Trim(session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").Text)
    
    If auxcountry <> "" And ws_contact.Cells(m, CPcountryj).Value = "" Then
        
        ws_contact.Cells(m, CPcountryj).Value = UCase(auxcountry)
        RegisterInfo = True     'Devuelve True si se ha actualizado la info
        
    End If
    
    'Idioma de contacto preferido
    auxlanguage = Trim(session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbADDR1_DATA-LANGU").Text)
    
    If auxlanguage <> "" And ws_contact.Cells(m, CPlanguagej).Value = "" Then
        
        ws_contact.Cells(m, CPlanguagej).Value = UCase(auxlanguage)
        RegisterInfo = True     'Devuelve True si se ha actualizado la info
        
    End If
    
    If RegisterInfo = True Then
        
        ws_contact.Cells(m, CPOKj).Value = ""
        
    End If
        
End Function

Function FillInfo(m As Integer)
'Si hay varia líneas para un proveedor se rellanan con esta función

    Dim auxvendorcode As String
    Dim auxcontact As String
    Dim auxtelephone As String
    Dim auxcountry As String
    Dim auxlanguage As String
    
    'Código de proveedor
    auxvendorcode = ws_contact.Cells(m - 1, CPvendorcodej).Value
        
    If auxvendorcode <> "" And ws_contact.Cells(m, CPvendorcodej).Value = "" Then
        
        ws_contact.Cells(m, CPvendorcodej).Value = auxvendorcode
        
    End If
    
    'Correo de contacto
    auxcontact = ws_contact.Cells(m - 1, CPmailj).Value
    
    If auxcontact <> "" And ws_contact.Cells(m, CPmailj).Value = "" Then
        
        ws_contact.Cells(m, CPmailj).Value = auxcontact
         
    End If
    
    'Teléfono de contacto
    auxtelephone = ws_contact.Cells(m - 1, CPtlfnoj).Value
    
    If auxtelephone <> "" And ws_contact.Cells(m, CPtlfnoj).Value = "" Then
        
        ws_contact.Cells(m, CPtlfnoj).Value = auxtelephone
        
    End If
    
    'Pais del proveedor
    auxcountry = ws_contact.Cells(m - 1, CPcountryj).Value
    
    If auxcountry <> "" And ws_contact.Cells(m, CPcountryj).Value = "" Then
        
        ws_contact.Cells(m, CPcountryj).Value = auxcountry
        
    End If
    
    'Idioma de contacto preferido
    auxlanguage = ws_contact.Cells(m - 1, CPlanguagej).Value
    
    If auxlanguage <> "" And ws_contact.Cells(m, CPlanguagej).Value = "" Then
        
        ws_contact.Cells(m, CPlanguagej).Value = auxlanguage
        
    End If
    
    ws_contact.Cells(m, CPOKj).Value = ws_contact.Cells(m - 1, CPOKj).Value
        
End Function

