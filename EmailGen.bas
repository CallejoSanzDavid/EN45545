Attribute VB_Name = "Módulo2"
Sub Email_Gen()
    
    Dim OutApp As Object
    Dim OutMail As Object
    Dim status As String
    Dim statusES() As String
    Dim nproducto As String
    Dim nprodj As Integer
    Dim auxname() As String
    Dim nombre As String
    Dim nombj As Integer
    Dim material As String
    Dim matj As Integer
    Dim manufacturer As String
    Dim manufj As Integer
    Dim EncabezadoEN As String
    Dim EncabezadoES As String
    Dim InfoEN As String
    Dim InfoES As String
    Dim DespedidaEN As String
    Dim DespedidaES As String
    Dim Firma As String
    Dim Separacion As String
    Dim Destinatario As String
    Dim contactj As Integer
    Dim G As Integer
    Dim validacion As Integer
    Dim ncorreos As Integer
    Dim nsincontacto As Integer
    Dim Aux As Integer
    
    If Sheets("FCIL").FilterMode Then Sheets("FCIL").ShowAllData
    
    Call MAYUSCULAS
    
    G = Sheets("FCIL").Range("A10:DA10").Find("Certificate global status*").Column
    N = Contar_Elem
    
    ncorreos = 0
    nsincontacto = 0
    
    Aux = Sheets("FCIL").Range("A10:DA10").Find("Assembly Name").Row
    
    For i = Sheets("FCIL").Range("A10:DA10").Find("Assembly Name").Row + 1 To N
        
        validacion = Alarmas(i)
        statusmin = AlarmasX(i)
    
        If Sheets("FCIL").Cells(i, G).Value <> "OK" And Sheets("FCIL").Cells(i, G).Value <> "No date" And validacion > statusmin Then
        
            status = Sheets("FCIL").Cells(i, G)
            statusES = Split(status, " ")
            
            nprodj = Sheets("FCIL").Range("A10:DA10").Find("Supplier part number").Column
            nproducto = Sheets("FCIL").Cells(i, nprodj).Value
            
            nombj = Sheets("FCIL").Range("A10:DA10").Find("Part name").Column
            auxname = Split(Cells(i, nombj).Value, " - MATERIAL")
            nombre = auxname(0)
            
            matj = Sheets("FCIL").Range("A10:DA10").Find("Raw material or product name*").Column
            material = Sheets("FCIL").Cells(i, matj).Value
            
            manufj = Sheets("FCIL").Range("A10:DA10").Find("Manufacturer name*").Column
            manufacturer = Sheets("FCIL").Cells(i, manufj).Value
            
            Firma = "MERAK Spain, S.A." + vbCrLf + "Miguel Faraday, 1" + vbCrLf + "Parque Empresarial 'La Carpetania'" + vbCrLf + "28906 Getafe (Madrid)" + vbCrLf + "mailto: f&s@merak-hvac.com"
            Separacion = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" + vbCrLf + vbCrLf
            
            contactj = Sheets("FCIL").Range("A10:DA10").Find("Supplier's Contact").Column
            Destinatario = Sheets("FCIL").Cells(i, contactj).Value
            
            If Destinatario = "Does NOT Exist" Then
            
                nsincontacto = nsincontacto + 1
               
            End If
            
            'Localizamos posiciones en la hoja de Contacto de proveedores
            
            CPsupplierj = Sheets("Contacto de proveedores").Range("A1:J1").Find("Supplier").Column
            CPendi = Sheets("Contacto de proveedores").Cells(Rows.Count, CPsupplierj).End(xlUp).Row
            
            CPmailj = Sheets("Contacto de proveedores").Range("A1:J1").Find("Mail").Column
            
            Sheets("Contacto de proveedores").Activate   'Para evitar que la siguiente línea de un error activamos la hoja donde tiene que buscar.
            CPmaili = Sheets("Contacto de proveedores").Range(Cells(1, CPmailj), Cells(CPendi, CPmailj)).Find(Destinatario).Row
            
            Do While Destinatario <> "Does NOT Exist" And Sheets("Contacto de proveedores").Cells(CPmaili, CPsupplierj).Value = Sheets("Contacto de proveedores").Cells(CPmaili + 1, CPsupplierj).Value     'Bucle para enviar email a todos los correos de contacto.
                
                Destinatario = Destinatario & "; " & Sheets("Contacto de proveedores").Cells(CPmaili + 1, CPmailj).Value
                CPmaili = CPmaili + 1
                
            Loop
            
            Sheets("FCIL").Activate
            
            If Destinatario <> "Does NOT Exist" And status = "EXPIRED" Then
                
                Set OutApp = CreateObject("Outlook.Application")
                OutApp.Session.Logon
                
                Set OutMail = OutApp.CreateItem(0)
                
                On Error Resume Next
                
                With OutMail
                
                    .To = Destinatario
                    .CC = "f&s@merak-hvac.com"
                    .Attachments.Add "T:\Compartir\F&S Certificates\20150223_Manufacturer_Declaration.doc"
                    .Subject = "Certificate update - " & nombre
                    
                    EncabezadoEN = "Dear Supplier," + vbCrLf + vbCrLf + "With this email we inform you that the Fire & Smoke declaration under the standard EN45545-2 related to the MERAK part number " & nproducto & " supplied by you has " & status & ". We kindly ask you to provide the extension declaration dossier as soon as possible." + vbCrLf + vbCrLf
                    InfoEN = "Product information: " + vbCrLf + "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & "." + vbCrLf + "- Raw material or product name: " & material & "." + vbCrLf + "- Manufacturer name: " & manufacturer & "." + vbCrLf + vbCrLf
                    DespedidaEN = "We remain waiting for your answer." + vbCrLf + vbCrLf + "Thank you very much in advance." + vbCrLf + vbCrLf
                    
                    EncabezadoES = "Estimado Proveedor," + vbCrLf + vbCrLf + "Con este correo electrónico le informamos de que su declaración de Fuegos y Humos bajo el estándar EN45545-2 en relación al número del elemento de MERAK " & nproducto & " distribuido por ustedes ha EXPIRADO. Les pedimos que nos faciliten la declaración de conformidad lo antes posible." + vbCrLf + vbCrLf
                    InfoES = "Información del producto: " + vbCrLf + "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & "." + vbCrLf + "- Materia prima o nombre del producto: " & material & "." + vbCrLf + "- Nombre del fabricante: " & manufacturer & "." + vbCrLf + vbCrLf
                    DespedidaES = "Esperamos su respuesta ." + vbCrLf + vbCrLf + "Gracias de antemano." + vbCrLf + vbCrLf
                
                    .Body = EncabezadoEN & InfoEN & DespedidaEN & Separacion & EncabezadoES & InfoES & DespedidaES & Firma
        
                    .Display
                    
                    ncorreos = ncorreos + 1
                    
                End With
            
            Else
            
                Set OutApp = CreateObject("Outlook.Application")
                OutApp.Session.Logon
                
                Set OutMail = OutApp.CreateItem(0)
                
                On Error Resume Next
                
                With OutMail
                    
                    .To = Destinatario
                    .CC = "f&s@merak-hvac.com"
                    .Attachments.Add "T:\Compartir\F&S Certificates\20150223_Manufacturer_Declaration.doc"
                    .Subject = "Certificate update - " & nombre
                    
                    EncabezadoEN = "Dear Supplier," + vbCrLf + vbCrLf + "With this email we inform you that the Fire & Smoke declaration under the standard EN45545-2 related to the MERAK part number " & nproducto & " supplied by you will expire in " & status & ". We kindly ask you to provide the extension declaration dossier as soon as possible." + vbCrLf + vbCrLf
                    InfoEN = "Product information: " + vbCrLf + "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK Part name: " & nombre & "." + vbCrLf + "- Raw material or product name: " & material & "." + vbCrLf + "- Manufacturer name: " & manufacturer & "." + vbCrLf + vbCrLf
                    DespedidaEN = "We remain waiting for your answer." + vbCrLf + vbCrLf + "Thank you very much in advance." + vbCrLf + vbCrLf
                    
                    If statusES(1) = "month/s" Then
                    
                        EncabezadoES = "Estimado Proveedor," + vbCrLf + vbCrLf + "Con este correo electrónico le informamos de que su declaración de Fuegos y Humos bajo el estándar EN45545-2 en relación al número del elemento de MERAK " & nproducto & " distribuido por ustedes expirará en " & statusES(0) & " mes/es. Les pedimos que nos faciliten la declaración de conformidad lo antes posible." + vbCrLf + vbCrLf
                        InfoES = "Información del producto: " + vbCrLf + "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & "." + vbCrLf + "- Materia prima o nombre del producto: " & material & "." + vbCrLf + "- Nombre del fabricante: " & manufacturer & "." + vbCrLf + vbCrLf
                        DespedidaES = "Esperamos su respuesta ." + vbCrLf + vbCrLf + "Gracias de antemano." + vbCrLf + vbCrLf
                    
                    End If
                    
                    If statusES(1) = "day/s" Then
                    
                        EncabezadoES = "Estimado Proveedor," + vbCrLf + vbCrLf + "Con este correo electrónico le informamos de que su declaración de Fuegos y Humos bajo el estándar EN45545-2 en relación al número del elemento de MERAK " & nproducto & " distribuido por ustedes expirará en " & statusES(0) & " día/s. Les pedimos que nos faciliten la declaración de conformidad lo antes posible." + vbCrLf + vbCrLf
                        InfoES = "Información del producto: " + vbCrLf + "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & "." + vbCrLf + "- Materia prima o nombre del producto: " & material & "." + vbCrLf + "- Nombre del fabricante: " & manufacturer & "." + vbCrLf + vbCrLf
                        DespedidaES = "Esperamos su respuesta ." + vbCrLf + vbCrLf + "Gracias de antemano." + vbCrLf + vbCrLf
                    
                    End If
                    
                    .Body = EncabezadoEN & InfoEN & DespedidaEN & Separacion & EncabezadoES & InfoES & DespedidaES & Firma
        
                    .Display
            
                    ncorreos = ncorreos + 1
            
                End With
                
            End If
    
        End If
    
    Application.StatusBar = "Checking expired certificates and generating emails: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")

    Next
    
    MsgBox (nsincontacto & " elemento/s expirado/s no tiene/n información de contacto." + vbCrLf + vbCrLf + "Se han generado " & ncorreos & " correo/s.")
    
    Application.StatusBar = ""
    
End Sub

Function Alarmas(i) As Integer
    
    Dim Val As Integer
        
    Val = Sheets("FCIL").Range("A10:DA10").Find("Email Sended").Column
    
    If Sheets("FCIL").Cells(i, Val) = "---" Then
        Alarmas = 24
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "6 month/s" Then
        Alarmas = 21
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "3 month/s" Then
        Alarmas = 18
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "2 month/s" Then
        Alarmas = 17
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "1 month/s" Then
        Alarmas = 16
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "15 day/s" Then
        Alarmas = 15
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "14 day/s" Then
        Alarmas = 14
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "13 day/s" Then
        Alarmas = 13
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "12 day/s" Then
        Alarmas = 12
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "11 day/s" Then
        Alarmas = 11
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "10 day/s" Then
        Alarmas = 10
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "9 day/s" Then
        Alarmas = 9
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "8 day/s" Then
        Alarmas = 8
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "7 day/s" Then
        Alarmas = 7
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "6 day/s" Then
        Alarmas = 6
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "5 day/s" Then
        Alarmas = 5
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "4 day/s" Then
        Alarmas = 4
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "3 day/s" Then
        Alarmas = 3
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "2 day/s" Then
        Alarmas = 2
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "1 day/s" Then
        Alarmas = 1
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "PRIORITY" Then
        Alarmas = 24
    End If

End Function

Function AlarmasX(i) As Integer
    
    Dim Val As Integer
        
    Val = Sheets("FCIL").Range("A10:DA10").Find("Certificate global status").Column
    
    If Sheets("FCIL").Cells(i, Val) = "OK" Then
        AlarmasX = 24
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "6 month/s" Or Sheets("FCIL").Cells(i, Val) = "5 month/s" Or Sheets("FCIL").Cells(i, Val) = "4 month/s" Then
        AlarmasX = 21
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "3 month/s" Then
        AlarmasX = 18
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "2 month/s" Then
        AlarmasX = 17
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "1 month/s" Then
        AlarmasX = 16
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "15 day/s" Then
        AlarmasX = 15
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "14 day/s" Then
        AlarmasX = 14
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "13 day/s" Then
        AlarmasX = 13
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "12 day/s" Then
        AlarmasX = 12
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "11 day/s" Then
        AlarmasX = 11
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "10 day/s" Then
        AlarmasX = 10
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "9 day/s" Then
        AlarmasX = 9
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "8 day/s" Then
        AlarmasX = 8
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "7 day/s" Then
        AlarmasX = 7
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "6 day/s" Then
        AlarmasX = 6
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "5 day/s" Then
        AlarmasX = 5
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "4 day/s" Then
        AlarmasX = 4
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "3 day/s" Then
        AlarmasX = 3
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "2 day/s" Then
        AlarmasX = 2
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "1 day/s" Then
        AlarmasX = 1
    End If
    
    If Sheets("FCIL").Cells(i, Val) = "EXPIRED" Then
        AlarmasX = 0
    End If

End Function

Function MAYUSCULAS()
    
    Dim Inicioi As Integer
    Dim partnamej As Integer
    Dim rawmaterialj As Integer
    Dim manufj As Integer
    Dim Aux As Integer
    
    partnamej = Sheets("FCIL").Range("A10:Z10").Find("Part name").Column            'Encuentra la posición de la columna con la palabra clave
    rawmaterialj = Sheets("FCIL").Range("A10:Z10").Find("Raw material or*").Column
    manufj = Sheets("FCIL").Range("A10:Z10").Find("Manufacturer name*").Column
    
    Aux = Sheets("FCIL").Range("A10:DA10").Find("Assembly Name").Row
    
    N = Contar_Elem
    
    For Inicioi = Sheets("FCIL").Range("A10:Z10").Find("Supplier").Row + 1 To N
        
        Application.StatusBar = "Format Progress (1/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        Sheets("FCIL").Cells(Inicioi, partnamej).Value = UCase(Sheets("FCIL").Cells(Inicioi, partnamej).Value)
    
    Next
    
    For Inicioi = Sheets("FCIL").Range("A10:Z10").Find("Supplier").Row + 1 To N
        
        Application.StatusBar = "Format Progress (2/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        Sheets("FCIL").Cells(Inicioi, rawmaterialj).Value = UCase(Sheets("FCIL").Cells(Inicioi, rawmaterialj).Value)
    
    Next
    
    For Inicioi = Sheets("FCIL").Range("A10:Z10").Find("Supplier").Row + 1 To N
        
        Application.StatusBar = "Format Progress (3/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        Sheets("FCIL").Cells(Inicioi, manufj).Value = UCase(Sheets("FCIL").Cells(Inicioi, manufj).Value)
    
    Next
    
    Application.StatusBar = ""
    
End Function


