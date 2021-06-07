Attribute VB_Name = "EmailGen"
Sub Email_Gen()
'Creates emails with the information of expired or about to expire to its pertinent supplier.
    
    Dim Export As Integer
    
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    Call Locate_Positions_OG
    
    '<-----------------------------------
    'Call Format_Capitalization
    
    Sheets(SheetName).Cells(Aux + 1, GlobalStatusj).Select
    
    '<-----------------------------------
    'TableName = ActiveSheet.ListObjects(1).Name
    'Call ClearFilters
    'Sorts Part Names in Alfabetic Order.
    'FilterSet = Sheets(SheetName).Cells(Aux, nombj).Value
    'Call AlfabeticOrder
    'Sorts Part Numbers in Alfabetic Order.
    'FilterSet = Sheets(SheetName).Cells(Aux, nprodj).Value
    'Call AlfabeticOrder
    'Sorts Suppliers in Alfabetic Order.
    'FilterSet = Sheets(SheetName).Cells(Aux, manufj).Value
    'Call AlfabeticOrder
    
    Call Locate_Positions_Contacts
    
    Call Locate_Positions_Email_Body
    
    Call Email_Body
    '<----------------------------------- DELETOS?
    'Información fija del email
    'EBHeadingEN = "Dear Supplier," + vbCrLf + vbCrLf + "With this email we inform you that the Fire & Smoke declaration under the standard EN45545-2 related to the listed MERAK part number/s supplied by you are expired or will expire shortly. We kindly ask you to provide the extension declaration dossier as soon as possible." + vbCrLf + vbCrLf + "Product information: " + vbCrLf + vbCrLf
    'EBFarewellEN = "We remain waiting for your answer." + vbCrLf + vbCrLf + "Thank you very much in advance." + vbCrLf + vbCrLf
    'EBSeparation = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" + vbCrLf + vbCrLf
    'EBHeadingES = "Estimado Proveedor," + vbCrLf + vbCrLf + "Con este correo electrónico le informamos de que su declaración de Fuegos y Humos bajo el estándar EN45545-2 en relación al listado de número/s MERAK distribuido por ustedes ha expirado o expirará pronto. Les pedimos que nos faciliten la declaración de conformidad lo antes posible." + vbCrLf + vbCrLf + "Información del producto: " + vbCrLf + vbCrLf
    'EBFarewellES = "Esperamos su respuesta." + vbCrLf + vbCrLf + "Gracias de antemano." + vbCrLf + vbCrLf
    'EBSignature = "MERAK Spain, S.A." + vbCrLf + "Miguel Faraday, 1" + vbCrLf + "Parque Empresarial 'La Carpetania'" + vbCrLf + "28906 Getafe (Madrid)" + vbCrLf + "mailto: f&s@merak-hvac.com"
    '<----------------------------------- DELETOS?
    
    Call Locate_Positions_RankingStatus
    
    'flags initial values
    ncorreos = 0
    nsincontacto = 0
    nexport = 0
    Export = 0
    
    For i = Aux + 1 To N
        
        NoContact = Manufacturer_Contact(SheetName, i, manufj)
        
        If NoContact = 0 Then
        
            GoTo NoContact:     'If there is no contact goes to the next line.
        
        End If
    
NextPartNumber:
        
        Application.StatusBar = "Checking expired certificates and generating emails: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        ColumnPosition = GlobalStatusj
        statusmin = Alarms_Check(i, ColumnPosition)
        
        ColumnPosition = EmailSendedj
        Validation = Alarms_Check(i, ColumnPosition)
        
        'flags initial values
        stat = 3
        auxstatus = 30
        marc1 = 0               'flags to identify the correct loop:
                                '0 - Initial value: Part Number with a simple material.
                                '1 - Part Number with several materials.
        lasterror = 0           'flag to prevent the error in which the last lines are not loged if it's not OK or are the same material.
        
        If statusmin <= 21 And Validation > statusmin Then
            
            'status = Sheets(SheetName).Cells(i, GlobalStatusj)     DELETOS?
            
            nproducto = Sheets(SheetName).Cells(i, nprodj).Value
            
            Auxsplit = 0            'flag initialized as 0 to detect if the Part Number has several materials.
            auxname = Split(Cells(i, nombj).Value, " - MATERIAL")
            nombre = auxname(0)
            
            On Error GoTo ErrorHandler:
            
            Auxsplit = auxname(1)
            
ErrorHandler:
            
            If Err.Number = 9 Then      'Solution for error 9: Subindex out of interval.

                Auxsplit = 0
                Err.Clear               'Error solved.
                Resume ErrorHandler:
                
            End If
            
            On Error GoTo 0
            
            material = Sheets(SheetName).Cells(i, matj).Value
                         
            Do While manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value   'If the next supplier is the same gets in this loop.
                
                status = Sheets(SheetName).Cells(i, GlobalStatusj)
                '-------------------------------Part Numbers with several materials---------------------------------
                If Auxsplit <> 0 And status <> "OK" Then
                    
                    marc1 = Complex_Part_Number
                        
                End If
                    
                '-------------------------------Part Numbers with one material---------------------------------
                If Auxsplit = 0 And marc1 = 0 And status <> "OK" Then
                                            
                    Call Simple_Part_Number
                    
                End If
                
                Select Case stat
                    
                    Case 2        'Days left for the certificate to expire.
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & " (" & auxstatus & " day/s to expire)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & " (" & auxstatus & " día/s para expirar)." + vbCrLf
                        expstatus = auxstatus & " día/s para expirar"
                                      
                    Case 1        'Months left for the certificate to expire.
                        
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & " (" & auxstatus & " month/s to expire)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & " (" & auxstatus & " mes/es para expirar)." + vbCrLf
                        expstatus = auxstatus & " mes/es para expirar"
                        
                    Case 0        'Any of the materials has expired.
                        
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & nombre & " (EXPIRED)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & nombre & " (EXPIRADO)." + vbCrLf
                        expstatus = "EXPIRADO"
                        
                End Select
                
                Export = 1
                
                FinalInfoEN = FinalInfoEN & InfoEN & InfoENRW + vbCrLf
                FinalInfoES = FinalInfoES & InfoES & InfoESRW + vbCrLf
                
                InfoENRW = ""
                InfoESRW = ""
                
                If Export = 1 Then         'A new line is loged for each Part Number in the Data Base "Pedidos".
        
                    Call Export_Data
                    nexport = nexport + 1
                    
                    If manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value Then
                    
                        i = i + 1               'With this line we prevent the code to analize the last line again.
                        GoTo NextPartNumber:    'Starts the loop again skipping the "Manufacturer_Contact" function.
                    
                    End If
                    
                End If
                '<--------------------------
                'If Export = 1 And manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value Then
                    
                    'i = i + 1               'Así evitamos que vuelva a analizar el último part number
                    'GoTo NextPartNumber:    'Vuelve al bucle saltándose las funciones que identifican el contacto
                
                'End If
                '<--------------------------
            Loop
                
        End If
        
        If Export = 1 And manufacturer <> Sheets(SheetName).Cells(i + 1, manufj).Value Then
            
            Call Email_Display
            ncorreos = ncorreos + 1
            
            Export = 0
            
            FinalInfoEN = ""
            FinalInfoES = ""
            
        End If
        
NoContact:

    Next
    
    MsgBox (nsincontacto & " elemento/s expirado/s no tiene/n información de contacto." + vbCrLf + vbCrLf + "Se han generado " & ncorreos & " correo/s para " & nexport & " part numbers.")
    
    'Clears the filters and sorts the Part Numbers by alfabetic order.
    FilterSet = Sheets(SheetName).Cells(Aux, nprodj).Value
    Call ClearFilters
    Call AlfabeticOrder
    
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Sub

Function Format_Capitalization()
'Corrects the format of the selected areas.
    
    Dim Starti As Integer
    
    For Starti = Aux + 1 To N
        
        Application.StatusBar = "Format Progress (1/3): " & Starti - Aux & " of " & N - Aux & ": " & Format((Starti - Aux) / (N - Aux), "0%")
        Sheets(SheetName).Cells(Starti, nombj).Value = UCase(Sheets(SheetName).Cells(Starti, nombj).Value)
    
    Next
    
    For Starti = Aux + 1 To N
        
        Application.StatusBar = "Format Progress (2/3): " & Starti - Aux & " of " & N - Aux & ": " & Format((Starti - Aux) / (N - Aux), "0%")
        Sheets(SheetName).Cells(Starti, matj).Value = UCase(Sheets(SheetName).Cells(Starti, matj).Value)
    
    Next
    
    For Starti = Aux + 1 To N
        
        Application.StatusBar = "Format Progress (3/3): " & Starti - Aux & " of " & N - Aux & ": " & Format((Starti - Aux) / (N - Aux), "0%")
        Sheets(SheetName).Cells(Starti, manufj).Value = UCase(Sheets(SheetName).Cells(Starti, manufj).Value)
    
    Next
    
    Application.StatusBar = ""

End Function

Function Email_Body()
'Gets the body information from the "Email Body" page.
    
    'Public EBSubject As String
    'Public EBHeadingEN As String
    'Public EBFarewellEN As String
    'Public EBSeparation As String
    'Public EBHeadingES As String
    'Public EBFarewellES As String
    'Public EBSignature As String
    
    EBSubject = Sheets(EmailBodySheetName).Cells(EBSubjecti, EBInfoj + 1).Value
    EBHeadingEN = Sheets(EmailBodySheetName).Cells(EBSubjecti + 1, EBInfoj + 1).Value
    
    EBFarewellEN = Sheets(EmailBodySheetName).Cells(EBSubjecti + 3, EBInfoj + 1).Value
    EBSeparation = Sheets(EmailBodySheetName).Cells(EBSubjecti + 4, EBInfoj + 1).Value
    EBHeadingES = Sheets(EmailBodySheetName).Cells(EBSubjecti + 5, EBInfoj + 1).Value
    
    EBFarewellES = Sheets(EmailBodySheetName).Cells(EBSubjecti + 7, EBInfoj + 1).Value
    EBSignature = Sheets(EmailBodySheetName).Cells(EBSubjecti + 8, EBInfoj + 1).Value
    
    'EBHeadingEN = "Dear Supplier," + vbCrLf + vbCrLf + "With this email we inform you that the Fire & Smoke declaration under the standard EN45545-2 related to the listed MERAK part number/s supplied by you are expired or will expire shortly. We kindly ask you to provide the extension declaration dossier as soon as possible." + vbCrLf + vbCrLf + "Product information: " + vbCrLf + vbCrLf
    'EBFarewellEN = "We remain waiting for your answer." + vbCrLf + vbCrLf + "Thank you very much in advance." + vbCrLf + vbCrLf
    'EBSeparation = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" + vbCrLf + vbCrLf
    'EBHeadingES = "Estimado Proveedor," + vbCrLf + vbCrLf + "Con este correo electrónico le informamos de que su declaración de Fuegos y Humos bajo el estándar EN45545-2 en relación al listado de número/s MERAK distribuido por ustedes ha expirado o expirará pronto. Les pedimos que nos faciliten la declaración de conformidad lo antes posible." + vbCrLf + vbCrLf + "Información del producto: " + vbCrLf + vbCrLf
    'EBFarewellES = "Esperamos su respuesta." + vbCrLf + vbCrLf + "Gracias de antemano." + vbCrLf + vbCrLf
    'EBSignature = "MERAK Spain, S.A." + vbCrLf + "Miguel Faraday, 1" + vbCrLf + "Parque Empresarial 'La Carpetania'" + vbCrLf + "28906 Getafe (Madrid)" + vbCrLf + "mailto: f&s@merak-hvac.com"
    
End Function

Function Manufacturer_Contact(SheetName, i, manufj) As Integer
    
    Manufacturer_Contact = 1
    
    manufacturer = Sheets(SheetName).Cells(i, manufj).Value
    
    Destinatario = Sheets(SheetName).Cells(i, ContactDBj).Value
    
    If Destinatario = "Does NOT Exist" Then
    
        nsincontacto = nsincontacto + 1
        Manufacturer_Contact = 0
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

'Needs to have this line before the Call:
'ColumnPosition = Column_Position_j
Function Alarms_Check(i, ColumnPosition) As Integer
'Logs the Global Status of each Part Number.
    
    Dim findstatus As String
    
    findstatus = Sheets(SheetName).Cells(i, ColumnPosition).Value

    Set Alarms_Check_i = Range(Sheets(RankingStatusSheet).Cells(RSRankingi, RSStatusENj), Sheets(RankingStatusSheet).Cells(RSEndi, RSStatusENj)).Find(findstatus)
    
    If Alarms_Check_i Is Nothing Then
        
        Alarms_Check = 24
        
    Else
        
        Alarms_Check = Sheets(RankingStatusSheet).Cells(Alarms_Check_i.Row, RSRankingj).Value
        
    End If
    
End Function

Function Export_Data()       'Registra la información de los correos generados. (nproducto, nombre, material, manufacturer, Destinatario, status)
'Exporta la info de los Part numbers en los que se ha generado una notificación
    Dim expi As Integer
    
    nombre_RecordSheet = ActiveWorkbook.Name
    
    Workbooks.Open (Sheets("validation Lists and Routes").Range("G2").Value)
    
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

Function Complex_Part_Number()
'-----------------------------------------------Diversos materiales para un Part Number ------------------------------------------------
        
    Complex_Part_Number = 1       'Marca tipo de material
    
    nombi = Sheets(SheetName).Range(Cells(Aux, nprodj), Cells(N, nprodj)).Find(nproducto).Row
    
    Do While nproducto = Sheets(SheetName).Cells(nombi + 1, nprodj).Value              'Bucle para registrar todos los materiales de dicho Part Number.
        
        material = Sheets(SheetName).Cells(nombi, matj).Value
        material1 = Sheets(SheetName).Cells(nombi + 1, matj).Value
        
        status = Sheets(SheetName).Cells(nombi, GlobalStatusj)
        
        If material <> material1 And status <> "OK" Then                           'Condición para evitar la repetición de un material.
                                                       
            Call Status_Case
            
        End If
       
        nombi = nombi + 1
        
    Loop
          
    material = Sheets(SheetName).Cells(nombi, matj).Value
    material1 = Sheets(SheetName).Cells(nombi - 1, matj).Value
    
    status = Sheets(SheetName).Cells(nombi, GlobalStatusj)
    
    'Condición para que se añada el último material del grupo.
    If (material <> material1 And nproducto = Sheets(SheetName).Cells(nombi - 1, nprodj).Value And status <> "OK") Or (material = material1 And status <> "OK") Then
                                    
        lasterror = 1       'Evita que se registre infinitamente el Part Number
        
        status = Sheets(SheetName).Cells(nombi, GlobalStatusj)
        
        Call Status_Case
                                    
    End If
            
    If i <> nombi Or lasterror = 1 Then
        
        i = nombi
        
    End If
    
End Function

Function Status_Case()
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
            
            Call Spanish_Module
        
    End Select

End Function

Function Spanish_Module()
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


Function Simple_Part_Number()
'-------------------------------Si el part number solo tiene un material---------------------------------
    If nproducto <> Sheets(SheetName).Cells(i + 1, nprodj).Value Then
                        
        status = Sheets(SheetName).Cells(i, GlobalStatusj)
        
        Call Status_Case
    
    End If
    
    
    Do While nproducto = Sheets(SheetName).Cells(i + 1, nprodj).Value        'En caso de que haya varias líneas para el mismo part number estas se saltan generándose un correo con la fecha más restrictiva.

        status = Sheets(SheetName).Cells(i, GlobalStatusj)
        
        If status = "OK" Or (status = Sheets(SheetName).Cells(i + 1, GlobalStatusj) And nproducto = Sheets(SheetName).Cells(i - 1, nprodj).Value) Then
                
            statusES(1) = 0
            GoTo NextIterarion:
                                            
        End If
        
        Call Status_Case
        
NextIterarion:

        i = i + 1
        
        nproducto = Sheets(SheetName).Cells(i, nprodj).Value
        
    Loop
      
    status = Sheets(SheetName).Cells(i, GlobalStatusj)
    
    'Condición para que se analice el último material del grupo.
    If nproducto = Sheets(SheetName).Cells(nombi - 1, nprodj).Value And status <> "OK" And status <> Sheets(SheetName).Cells(i - 1, GlobalStatusj) Then
        'STOP
        Call Status_Case
                                    
    End If

End Function

Function Email_Display()

    'Encabezado correo.
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.session.Logon
    
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
    
        'Generación del correo.
        .To = Destinatario
        .CC = "f&s@merak-hvac.com"
        .Attachments.Add "T:\Compartir\F&S Certificates\20150223_Manufacturer_Declaration.doc"
        .Subject = EBSubject & manufacturer
        .Body = EBHeadingEN & FinalInfoEN & EBFarewellEN & EBSeparation & EBHeadingES & FinalInfoES & EBFarewellES & EBSignature
        .Display
        'TEST STOP
    
    End With

End Function
