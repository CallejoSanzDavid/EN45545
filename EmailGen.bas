Attribute VB_Name = "EmailGen"
Sub Email_Gen()         'Genera correos de las líneas que contengan certificados caducados.
      
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    SheetName = ActiveSheet.Name
    ContactSheetName = "Contacto de proveedores"
    
    'Localizamos las posiciones en el FCIL
    Aux = Sheets(SheetName).Range("A1:DA20").Find("Assembly Name").Row
    G = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Certificate global status*").Column
    nprodj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier part number").Column
    N = Sheets(SheetName).Cells(Rows.Count, nprodj).End(xlUp).Row
    
    nombj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Part name").Column
    matj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Raw material or product name*").Column
    manufj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer name*").Column
    contactj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier's Contact").Column
    
    Call MAYUSCULAS
    
    'Localizamos posiciones en la hoja de Contacto de proveedores
    CPsupplierj = Sheets(ContactSheetName).Range("A1:J1").Find("Supplier").Column
    CPendi = Sheets(ContactSheetName).Cells(Rows.Count, CPsupplierj).End(xlUp).Row
    CPmailj = Sheets(ContactSheetName).Range("A1:J1").Find("Mail").Column
    
    Sheets(SheetName).Cells(Aux + 1, G).Select
       
    TableName = ActiveSheet.ListObjects(1).Name
    
    'Borrar todos los filtros que haya aplicados
    Call ClearFilters
    'Ordenar los Part Names por orden alfabético
    FilterSet = Sheets(SheetName).Cells(Aux, nombj).Value
    Call AlfabeticOrder
    'Ordenar los Part Numbers por orden alfabético
    FilterSet = Sheets(SheetName).Cells(Aux, nprodj).Value
    Call AlfabeticOrder
    'Ordenar los Proveedores por orden alfabético
    FilterSet = Sheets(SheetName).Cells(Aux, manufj).Value
    Call AlfabeticOrder
       
    'Información fija del email
    EncabezadoEN = "Dear Supplier," + vbCrLf + vbCrLf + "With this email we inform you that the Fire & Smoke declaration under the standard EN45545-2 related to the listed MERAK part number/s supplied by you are expired or will expire shortly. We kindly ask you to provide the extension declaration dossier as soon as possible." + vbCrLf + vbCrLf + "Product information: " + vbCrLf + vbCrLf
    DespedidaEN = "We remain waiting for your answer." + vbCrLf + vbCrLf + "Thank you very much in advance." + vbCrLf + vbCrLf
    Separacion = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" + vbCrLf + vbCrLf
    EncabezadoES = "Estimado Proveedor," + vbCrLf + vbCrLf + "Con este correo electrónico le informamos de que su declaración de Fuegos y Humos bajo el estándar EN45545-2 en relación al listado de número/s MERAK distribuido por ustedes ha expirado o expirará pronto. Les pedimos que nos faciliten la declaración de conformidad lo antes posible." + vbCrLf + vbCrLf + "Información del producto: " + vbCrLf + vbCrLf
    DespedidaES = "Esperamos su respuesta." + vbCrLf + vbCrLf + "Gracias de antemano." + vbCrLf + vbCrLf
    Firma = "MERAK Spain, S.A." + vbCrLf + "Miguel Faraday, 1" + vbCrLf + "Parque Empresarial 'La Carpetania'" + vbCrLf + "28906 Getafe (Madrid)" + vbCrLf + "mailto: f&s@merak-hvac.com"
        
    'Inicialización de marcadores
    ncorreos = 0
    nsincontacto = 0
    nexport = 0
    Export = 0
    
    For i = Aux + 1 To N
        
        NoContact = ManufacturerContact(SheetName, i, manufj)
        
        If NoContact = 0 Then
        
            GoTo NoContact:     'Si no existe el contacto se pasa a la siguiente iteración.
        
        End If
    
NextPartNumber:
        
        Application.StatusBar = "Checking expired certificates and generating emails: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        Valj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Email Sended").Column
        validacion = Alarmas()
        
        Valj = G
        statusmin = Alarmas()
        
        'Inicialización de marcadores
        stat = 3
        auxstatus = 30
        marc1 = 0               'Marcador para identificar en qué bucle ha entrado:
                                '0 - Estado inicial: Part Number con un material.
                                '1 - Diversos materiales para un Part Number.
        lasterror = 0           'Marcador para evitar el error en el que no se registran las últimas líneas del Part Number si no están OK y son del mismo material.
        
        
        If statusmin <= 21 And validacion > statusmin Then
            
            status = Sheets(SheetName).Cells(i, G)
            
            nproducto = Sheets(SheetName).Cells(i, nprodj).Value
            
            Auxsplit = 0            'Esta variable se inicializa en 0 y se usa para detectar si el Part Number está compuesto por varios materiales
            auxname = Split(Cells(i, nombj).Value, " - MATERIAL")
            nombre = auxname(0)
            
            On Error GoTo ErrorHandler:
            
            Auxsplit = auxname(1)
            
ErrorHandler:
            
            If Err.Number = 9 Then      'Solución error 9. Subíndice fuera del intervalo.

                Auxsplit = 0
                Err.Clear               'Marca que el error se ha solucionado.
                Resume ErrorHandler:
                
            End If
            
            On Error GoTo 0
            
            material = Sheets(SheetName).Cells(i, matj).Value
                         
            Do While manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value   'Si el Proveedor en la siguiente línea es igual entra en el bucle
                
                status = Sheets(SheetName).Cells(i, G)
                '-------------------------------Diversos materiales para un Part Number ---------------------------------
                If Auxsplit <> 0 And status <> "OK" Then
                    
                    marc1 = ComplexPartNumber
                        
                End If
                    
                '-------------------------------Si el part number solo tiene un material---------------------------------
                If Auxsplit = 0 And marc1 = 0 And status <> "OK" Then
                                            
                    Call SimplePartNumber
                    
                End If
                
                Select Case stat
                    
                    Case 2        'Si faltan día/s para que el material más restrictivo caduque.
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & " (" & auxstatus & " day/s to expire)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & " (" & auxstatus & " día/s para expirar)." + vbCrLf
                        expstatus = auxstatus & " día/s para expirar"
                                      
                    Case 1        'Si faltan mes/es para que el material más restrictivo caduque.
                        
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & " (" & auxstatus & " month/s to expire)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & " (" & auxstatus & " mes/es para expirar)." + vbCrLf
                        expstatus = auxstatus & " mes/es para expirar"
                        
                    Case 0        'Si algunos de los materiales ha expirado.
                        
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & " (EXPIRED)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & " (EXPIRADO)." + vbCrLf
                        expstatus = "EXPIRADO"
                        
                End Select
                
                Export = 1
                
                FinalInfoEN = FinalInfoEN & InfoEN & InfoENRW + vbCrLf
                FinalInfoES = FinalInfoES & InfoES & InfoESRW + vbCrLf
                
                InfoENRW = ""
                InfoESRW = ""
                
                If Export = 1 Then         'Se genera una línea en la BB.DD. de Pedidos por Part Number
        
                    Call EXPORT_DATA
                    nexport = nexport + 1
                    
                End If
                
                If Export = 1 And manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value Then
                    
                    i = i + 1               'Así evitamos que vuelva a analizar el último part number
                    GoTo NextPartNumber:    'Vuelve al bucle saltándose las funciones que identifican el contacto
                
                End If
                
            Loop                    'Bucle Proveedor.
                
        End If
        
        If Export = 1 And manufacturer <> Sheets(SheetName).Cells(i + 1, manufj).Value Then
            
            Call EmailDisplay
            ncorreos = ncorreos + 1
            
            Export = 0
            
            FinalInfoEN = ""
            FinalInfoES = ""
            
        End If
        
NoContact:

    Next
    
    MsgBox (nsincontacto & " elemento/s expirado/s no tiene/n información de contacto." + vbCrLf + vbCrLf + "Se han generado " & ncorreos & " correo/s para " & nexport & " part numbers.")
    
    'Borra los filtros y ordena los Part Numbers por orden alfabético.
    FilterSet = Sheets(SheetName).Cells(Aux, nprodj).Value
    Call ClearFilters
    Call AlfabeticOrder
    
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Sub

Function MAYUSCULAS()           'Corrige el formato de los campos seleccionados.
    
    Dim Inicioi As Integer
    
    'Encuentra la posición de la columna con la palabra clave
    'nombj = Sheets(SheetName).Range("A10:Z10").Find("Part name").Column
    
    For Inicioi = Aux + 1 To N
        
        Application.StatusBar = "Format Progress (1/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        Sheets(SheetName).Cells(Inicioi, nombj).Value = UCase(Sheets(SheetName).Cells(Inicioi, nombj).Value)
    
    Next
    
    For Inicioi = Aux + 1 To N
        
        Application.StatusBar = "Format Progress (2/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        Sheets(SheetName).Cells(Inicioi, matj).Value = UCase(Sheets(SheetName).Cells(Inicioi, matj).Value)
    
    Next
    
    For Inicioi = Aux + 1 To N
        
        Application.StatusBar = "Format Progress (3/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        Sheets(SheetName).Cells(Inicioi, manufj).Value = UCase(Sheets(SheetName).Cells(Inicioi, manufj).Value)
    
    Next
    
    Application.StatusBar = ""

End Function

Function ManufacturerContact(SheetName, i, manufj) As Integer
    
    ManufacturerContact = 1
    
    manufacturer = Sheets(SheetName).Cells(i, manufj).Value
    
    Destinatario = Sheets(SheetName).Cells(i, contactj).Value
    
    If Destinatario = "Does NOT Exist" Then
    
        nsincontacto = nsincontacto + 1
        ManufacturerContact = 0
        Exit Function
        
    End If
    
    Sheets(ContactSheetName).Activate   'Para evitar que la siguiente línea de un error activamos la hoja donde tiene que buscar.
    CPmaili = Sheets(ContactSheetName).Range(Cells(1, CPmailj), Cells(CPendi, CPmailj)).Find(Destinatario).Row
    
    Do While Destinatario <> "Does NOT Exist" And Sheets(ContactSheetName).Cells(CPmaili, CPsupplierj).Value = Sheets(ContactSheetName).Cells(CPmaili + 1, CPsupplierj).Value     'Bucle para enviar email a todos los correos de contacto.
        
        Destinatario = Destinatario & "; " & Sheets(ContactSheetName).Cells(CPmaili + 1, CPmailj).Value
        CPmaili = CPmaili + 1
        
    Loop
    
    Sheets(SheetName).Activate
    
End Function

Function EXPORT_DATA()       'Registra la información de los correos generados. (nproducto, nombre, material, manufacturer, Destinatario, status)

    Dim expi As Integer
    
    nombre_RecordSheet = ActiveWorkbook.Name
    
    Workbooks.Open (Sheets("Listas de Validación").Range("G2").Value)
    
    nombre_bbdd = ActiveWorkbook.Name
    
    Workbooks(nombre_bbdd).Sheets("TEMP").Activate          'Activamos el libro en la hoja de registro
    
    expi = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row + 1            'Localizamos la última fila con info en una columna sin celdas combinadas
    
    Workbooks(nombre_RecordSheet).Activate          'Activamos la BBDD F&H para extraer la info de esta
    
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 1).Value = nproducto             'Part Number
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 2).Value = nombre                'Part Name
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 3).Value = material              'Raw Material
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 4).Value = manufacturer          'Supplier
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 5).Value = "---"                 'TR number
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 6).Value = Destinatario          'Contact e-mails
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 7).Value = "BB.DD."              'Quien lo pide
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 8).Value = Date                  'Cuando se ha pedido
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 9).Value = Date                  'Fecha del último coreo enviado
    
    Workbooks(nombre_bbdd).Sheets("TEMP").Activate                                     'Activamos la BBDD de pedidos para que guarde la info archivada
    Workbooks(nombre_bbdd).Sheets("AUX2").Range("A1").Copy Range("J" & expi)           'Lista de validación
    
    '<------------- Se registra el status de la última línea, no la de el material más restrictivo.
    Workbooks(nombre_bbdd).Sheets("TEMP").Cells(expi, 11).Value = expstatus            'Estatus de los TR
       
    ActiveWorkbook.Save
    ActiveWorkbook.Close

End Function

Function Alarmas() As Integer          'Módulo de comparación de alarmas (SheetName, i, Valj, G)
    
    daynum = 0
    auxday = Split(Cells(i, G).Value, " day/s")
    daynum = auxday(0)
    
    Case_Option = Sheets(SheetName).Cells(i, Valj)
    
    Select Case Case_Option
           
        Case ""
            Alarmas = 24
        
        Case "No date"
            Alarmas = 23
           
        Case "OK"
            Alarmas = 22
        
        Case "6 month/s"
            Alarmas = 21
        
        Case "5 month/s"
            Alarmas = 21
            
        Case "4 month/s"
            Alarmas = 21
        
        Case "3 month/s"
            Alarmas = 18
        
        Case "2 month/s"
            Alarmas = 17
        
        Case "1 month/s"
            Alarmas = 16
        
        Case "15 day/s"
            Alarmas = 15
        
        Case "14 day/s"
            Alarmas = 14
        
        Case "13 day/s"
            Alarmas = 13
        
        Case "12 day/s"
            Alarmas = 12
        
        Case "11 day/s"
            Alarmas = 11
        
        Case "10 day/s"
            Alarmas = 10
        
        Case "9 day/s"
            Alarmas = 9
        
        Case "8 day/s"
            Alarmas = 8
        
        Case "7 day/s"
            Alarmas = 7
        
        Case "6 day/s"
            Alarmas = 6
        
        Case "5 day/s"
            Alarmas = 5
        
        Case "4 day/s"
            Alarmas = 4
        
        Case "3 day/s"
            Alarmas = 3
        
        Case "2 day/s"
            Alarmas = 2
                
        Case "1 day/s"
            Alarmas = 1
        
        Case "EXPIRED"
            Alarmas = 0
    
        Case Else        'Opciones: "---"; "PRIORITY"; Falta menos de 1 mes pero más de 15 días para expirar u otro.
            
            If Valj = G And daynum < 31 And daynum > 15 Then
            
                Alarmas = 15
            
            Else
            
                Alarmas = 24
            
            End If
            
    End Select
    
End Function


Function ComplexPartNumber()
'-----------------------------------------------Diversos materiales para un Part Number ------------------------------------------------
        
    ComplexPartNumber = 1       'Marca tipo de material
    
    nombi = Sheets(SheetName).Range(Cells(Aux, nprodj), Cells(N, nprodj)).Find(nproducto).Row
    
    Do While nproducto = Sheets(SheetName).Cells(nombi + 1, nprodj).Value              'Bucle para registrar todos los materiales de dicho Part Number.
        
        material = Sheets(SheetName).Cells(nombi, matj).Value
        material1 = Sheets(SheetName).Cells(nombi + 1, matj).Value
        
        status = Sheets(SheetName).Cells(nombi, G)
        
        If material <> material1 And status <> "OK" Then                           'Condición para evitar la repetición de un material.
                                                       
            Call StatusCase
            
        End If
       
        nombi = nombi + 1
        
    Loop
          
    material = Sheets(SheetName).Cells(nombi, matj).Value
    material1 = Sheets(SheetName).Cells(nombi - 1, matj).Value
    
    status = Sheets(SheetName).Cells(nombi, G)
    
    'Condición para que se añada el último material del grupo.
    If (material <> material1 And nproducto = Sheets(SheetName).Cells(nombi - 1, nprodj).Value And status <> "OK") Or (material = material1 And status <> "OK") Then
                                    
        lasterror = 1       'Evita que se registre infinitamente el Part Number
        
        status = Sheets(SheetName).Cells(nombi, G)
        
        Call StatusCase
                                    
    End If
            
    If i <> nombi Or lasterror = 1 Then
        
        i = nombi
        
    End If
    
End Function

Function StatusCase()
'Bloque para generar la información del correo según su status

    Select Case status
            
        Case "EXPIRED"
            
            AuxENRW = "- Raw material or product name: " & material & " (" & status & ")." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW
        
            AuxESRW = "- Materia prima o nombre del producto: " & material & " (EXPIRADO)." + vbCrLf
            InfoESRW = InfoESRW & AuxESRW
            
            auxstatus = 0
            stat = 0        'Estado global del Part Number. EXPIRADO.
            
        Case "No date"
            AuxENRW = "- Raw material or product name: " & material & " (" & status & ")." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW

            AuxESRW = "- Materia prima o nombre del producto: " & material & " (Sin fecha)." + vbCrLf
            InfoESRW = InfoESRW & AuxESRW
            
        Case Else           'Condición para cuando faltan día/s o mes/es para expirar.  If status <> "OK" And status <> "EXPIRED" And status <> "No date" Then
            AuxENRW = "- Raw material or product name: " & material & " (" & status & " to expire)." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW
            
            Call SpanishModule
        
    End Select

End Function

Function SpanishModule()
'Bloque para poner el mensaje en Español.
    statusES = Split(status, " ")
                
    If statusES(1) = "day/s" Then
    
        'Bloque de estado global del Part Number
        If stat <> 0 Then

            If stat = 1 Then    'Si el estado anterior era de meses y la nueva línea tiene el estado de días.
                   
                auxstatus = statusES(0)
                                          
            End If
            
            stat = 2
            
            If statusES(0) < auxstatus Then    'Actualizamos el estado de meses a días.

                auxstatus = statusES(0)
                                              
            End If
                                                    
        End If
        
        AuxESRW = "- Materia prima o nombre del producto: " & material & " (" & statusES(0) & " día/s para expirar)." + vbCrLf
        InfoESRW = InfoESRW & AuxESRW
        
    End If
        
    If statusES(1) = "month/s" Then
        
        AuxESRW = "- Materia prima o nombre del producto: " & material & " (" & statusES(0) & " mes/es para expirar)." + vbCrLf
        InfoESRW = InfoESRW & AuxESRW
        
        If statusES(0) < auxstatus And stat <> 0 And stat <> 2 Then     'Actualizamos el estado global del Part Number a meses.

            auxstatus = statusES(0)
            stat = 1
            
        End If
        
    End If

End Function


Function SimplePartNumber()
'-------------------------------Si el part number solo tiene un material---------------------------------
    If nproducto <> Sheets(SheetName).Cells(i + 1, nprodj).Value Then
                        
        status = Sheets(SheetName).Cells(i, G)
        
        Call StatusCase
    
    End If
    
    
    Do While nproducto = Sheets(SheetName).Cells(i + 1, nprodj).Value        'En caso de que haya varias líneas para el mismo part number estas se saltan generándose un correo con la fecha más restrictiva.

        status = Sheets(SheetName).Cells(i, G)
        
        If status = "OK" Or (status = Sheets(SheetName).Cells(i + 1, G) And nproducto = Sheets(SheetName).Cells(i - 1, nprodj).Value) Then
                
            statusES(1) = 0
            GoTo NextIterarion:
                                            
        End If
        
        Call StatusCase
        
NextIterarion:

        i = i + 1
        
        nproducto = Sheets(SheetName).Cells(i, nprodj).Value
        
    Loop
      
    status = Sheets(SheetName).Cells(i, G)
    
    'Condición para que se analice el último material del grupo.
    If nproducto = Sheets(SheetName).Cells(nombi - 1, nprodj).Value And status <> "OK" And status <> Sheets(SheetName).Cells(i - 1, G) Then
        'STOP
        Call StatusCase
                                    
    End If

End Function

Function EmailDisplay()

    'Encabezado correo.
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
    
        'Generación del correo.
        .To = Destinatario
        .CC = "f&s@merak-hvac.com"
        .Attachments.Add "T:\Compartir\F&S Certificates\20150223_Manufacturer_Declaration.doc"
        .Subject = "EN45545 Certificate update - " & manufacturer
        .Body = EncabezadoEN & FinalInfoEN & DespedidaEN & Separacion & EncabezadoES & FinalInfoES & DespedidaES & Firma
        .Display
        'TEST STOP
    
    End With

End Function

