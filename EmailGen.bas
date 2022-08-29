Attribute VB_Name = "EmailGen"
Sub Email_Gen()
'Genera correos con la información de los materiales con la certificación expirada o a punto de expirar a los proveedores correspondientes.
    
    'Contadores.
    Dim i As Integer
    Dim nmails As Integer
    Dim nnocontact As Integer
    Dim nexport As Integer
    'Marcadores.
    Dim Export As Boolean
    Dim NoContact As Boolean
    Dim validation As Integer
    Dim OpenDataBase As Boolean
    'Separadores.
    Dim Auxsplit As String
    Dim auxname() As String
    Dim partname As String
    'Buscadores e identificadores.
    Dim pnamei As Integer
    Dim ColumnPosition As Integer
    Dim status As String
    Dim statusmin As Integer
    'Variables.
    Dim expstatus As String
    'Nombres de libros.
    Dim partname_RecordSheet As String
    Dim partname_bbdd As String
    'Información del Email.
    Dim InfoEN As String
    Dim InfoES As String
    Dim FinalInfoEN As String
    Dim FinalInfoES As String
    
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    t1 = Time
    
    Call Locate_Positions_OG
    
    Call Mayus_Clean(1, nombj)
    Call Mayus_Clean(2, matj)
    
    ws_OG.Cells(Aux + 1, GlobalStatusj).Select
    
    TableName = ActiveSheet.ListObjects(1).Name
    Call ClearFilters
    'Ordena los Part Names en orden alfabético.
    FilterSet = ws_OG.Cells(Aux, nombj).Value
    Call AlfabeticOrder
    'Ordena los Part Numbers en orden alfabético.
    FilterSet = ws_OG.Cells(Aux, nprodj).Value
    Call AlfabeticOrder
    'Ordena los Suppliers en orden alfabético.
    FilterSet = ws_OG.Cells(Aux, manufj).Value
    Call AlfabeticOrder
    
    Call Locate_Positions_Contacts
    
    Call Locate_Positions_Email_Body
    
    Call Email_Body
    
    Call Locate_Positions_RankingStatus
    
    'Inicialización de marcadores.
    nmails = 0
    nnocontact = 0
    nexport = False
    Export = False
    OpenDataBase = False
    language = ""
    NoContact = False
    
    For i = Aux + 1 To N
    
NextPartNumber:
        
        Application.StatusBar = "Checking expired certificates and generating emails: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        ColumnPosition = GlobalStatusj
        statusmin = Alarms_Check(i, ColumnPosition)
        
        ColumnPosition = EmailSendedj
        validation = Alarms_Check(i, ColumnPosition)
        
        'Inicialización de marcadores.
        stat = 3                'Bandera que identifica el estado mínimo:
                                '3 - Valor auxiliar.
                                '2 - day/s
                                '1 - month/s
                                '0 - EXPIRED
        auxstatus = 30          'Valor auxiliar para identificar el estado mínimo de la lista de materiales de cada proveedor.
        
        If statusmin <= 21 And validation > statusmin Then
            
            If NoContact = False Then
            
                NoContact = Manufacturer_Contact(i, nnocontact)
                
                If NoContact = False Then
            
                    GoTo NoContact:     'Si después de la búsqueda no hay contacto continuamos con la siguiente iteración.
                
                End If
                
            End If
            
            nproducto = ws_OG.Cells(i, nprodj).Value
            
            Auxsplit = "0"            'Marcador inicializado en "0" para detectar si el paterial es Simple o Compuesto.
            auxname = Split(ws_OG.Cells(i, nombj).Value, " - MATERIAL")
            partname = auxname(0)
            
            On Error GoTo ErrorHandler:
            
            Auxsplit = auxname(1)
            
ErrorHandler:
            
            If Err.Number = 9 Then      'Solution for error 9: Subindex out of interval.

                Auxsplit = "0"
                Err.Clear               'Error solved.
                Resume ErrorHandler:
                
            End If
            
            On Error GoTo 0
            
            material = ws_OG.Cells(i, matj).Value
            
            status = ws_OG.Cells(i, GlobalStatusj)
            pnamei = ws_OG.Range(Cells(Aux, nprodj), Cells(N, nprodj)).Find(nproducto).Row
            
            '-------------------------------Part Numbers con varios materiales (Compuestos)---------------------------------
            If Auxsplit <> "0" And status <> "OK" Then
                
                i = Complex_Part_Number(pnamei, i)
              
            ElseIf status <> "OK" Then
                                    
                i = Simple_Part_Number(i)
                
            End If
            
            Select Case stat
                
                Case 2        'Faltan días para que expire alguno de los certificados.
                    InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & partname & " (" & auxstatus & " day/s to expire)." + vbCrLf
                    InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & partname & " (" & auxstatus & " día/s para expirar)." + vbCrLf
                    expstatus = auxstatus & " día/s para expirar"
                                  
                Case 1        'Faltan meses para que expire alguno de los certificados.
                    
                    InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & partname & " (" & auxstatus & " month/s to expire)." + vbCrLf
                    InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & partname & " (" & auxstatus & " mes/es para expirar)." + vbCrLf
                    expstatus = auxstatus & " mes/es para expirar"
                    
                Case 0        'Alguno de los certificados ha expirado.
                    
                    InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & partname & " (EXPIRED)." + vbCrLf
                    InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & partname & " (EXPIRADO)." + vbCrLf
                    expstatus = "EXPIRADO"
                    
            End Select
            
            FinalInfoEN = FinalInfoEN & InfoEN & InfoENRW + vbCrLf
            FinalInfoES = FinalInfoES & InfoES & InfoESRW + vbCrLf
            
            InfoENRW = ""
            InfoESRW = ""
            
            If OpenDataBase = False Then
            
                partname_RecordSheet = ActiveWorkbook.Name
                
                Workbooks.Open (Sheets("Validation Lists and Routes").Range("H2").Value)
                
                partname_bbdd = ActiveWorkbook.Name
                
                OpenDataBase = True
                
            End If
            
            Export = Export_Data(partname_RecordSheet, partname_bbdd, partname, expstatus)
            ws_OG.Activate        'Workbooks(partname_RecordSheet).Activate
            
            nexport = nexport + 1
            
            If manufacturer = ws_OG.Cells(i + 1, manufj).Value Then
            
                i = i + 1               'Con esta línea evitamos que el código vuelva a analizar otra vez la misma línea.
                GoTo NextPartNumber:    'Vuelve al comienzo del bucle For.
            
            End If
                
        End If
NoContact:
        If Export = True And manufacturer <> ws_OG.Cells(i + 1, manufj).Value Then
            
            Call Email_Display(FinalInfoEN, FinalInfoES)
            nmails = nmails + 1
            
            NoContact = False
            Export = False
            
            FinalInfoEN = ""
            FinalInfoES = ""
            
        End If

    Next
    
    Workbooks(partname_bbdd).Activate
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    Workbooks(partname_RecordSheet).Activate
    
    'Elimina los filtros y ordena los Part Numbers por orden alfabético.
    FilterSet = ws_OG.Cells(Aux, nprodj).Value
    Call ClearFilters
    Call AlfabeticOrder
    
    t2 = Time
    crono = Format(t2 - t1, "hh:mm:ss")
    
    MsgBox (nnocontact & " elemento/s expirado/s no tiene/n información de contacto." + vbCrLf + vbCrLf + "Se han generado " & nmails & " correo/s para " & nexport & " part numbers." + vbCrLf + vbCrLf + "Tiempo de operación: " & crono & ".")
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Sub

Function Mayus_Clean(Process As Integer, Field As Integer)
'Poner en mayúscula y eliminar espacios innecesarios de la columna elegida.
    
    Dim Starti As Integer

    For Starti = Aux + 1 To N

        Application.StatusBar = "Format Progress (" & Process & "/2): " & Starti - Aux & " of " & N - Aux & ": " & Format((Starti - Aux) / (N - Aux), "0%")
        
        ws_OG.Cells(Starti, Field).Value = UCase(ws_OG.Cells(Starti, Field).Value)
        ws_OG.Cells(Starti, Field).Value = Trim(ws_OG.Cells(Starti, Field).Value)
    
    Next
    
    Application.StatusBar = ""
    
End Function

Function Email_Body()
'obtiene la información de la hoja "Email Body".
    
    EBcc = ws_emailb.Cells(EBcci, EBInfoj).Value
    
    EBSubjectEN = ws_emailb.Cells(EBSubjectENi, EBInfoj).Value
    EBSubjectES = ws_emailb.Cells(EBSubjectESi, EBInfoj).Value
    
    EBAttachment = ws_emailb.Cells(EBAttachmenti, EBInfoj).Value
    
    EBHeadingEN = ws_emailb.Cells(EBHeadingENi, EBInfoj).Value
    EBFarewellEN = ws_emailb.Cells(EBFarewellENi, EBInfoj).Value
    
    EBSeparation = ws_emailb.Cells(EBSeparationi, EBInfoj).Value
    
    EBHeadingES = ws_emailb.Cells(EBHeadingESi, EBInfoj).Value
    EBFarewellES = ws_emailb.Cells(EBFarewellESi, EBInfoj).Value
    
    EBSignature = ws_emailb.Cells(EBSignaturei, EBInfoj).Value

End Function

Function Manufacturer_Contact(i As Integer, nnocontact As Integer) As Boolean
    
    Dim CPmaili As Integer
    
    manufacturer = ws_OG.Cells(i, manufj).Value
    
    Recipient = ws_OG.Cells(i, ContactDBj).Value
    
    If Recipient = "Does NOT Exist" Then
    
        nnocontact = nnocontact + 1
        Manufacturer_Contact = False
        Exit Function
        
    End If
    
    Manufacturer_Contact = True
    
    ws_contact.Activate   'Para evitar errores en las siguientes líneas se activa la hoja de contactos.
    
    On Error GoTo ErrorHandler1:
    
    CPmaili = ws_contact.Range(Cells(1, CPmailj), Cells(CPendi, CPmailj)).Find(Recipient).Row
            
    If Err.Number = 9 Then      'Solución del error 9: Subindex out of interval.

        Err.Clear
        Call ClearFilters
        CPmaili = ws_contact.Range(Cells(1, CPmailj), Cells(CPendi, CPmailj)).Find(Recipient).Row
        If Err.Number = 9 Then
            Manufacturer_Contact = False
            Exit Function
        End If
        Resume ErrorHandler1:
ErrorHandler1:
    End If
    
    On Error GoTo 0

    language = Trim(ws_contact.Cells(CPmaili, CPlanguagej).Value)
    
    Do While Recipient <> "Does NOT Exist" And ws_contact.Cells(CPmaili, CPsupplierj).Value = ws_contact.Cells(CPmaili + 1, CPsupplierj).Value
    'Registra todos los contactos del proveedor si tuviera varios correos registrados.
        
        Recipient = Recipient & "; " & ws_contact.Cells(CPmaili + 1, CPmailj).Value
        CPmaili = CPmaili + 1
        
    Loop
    
    ws_OG.Activate
    
End Function

'Necesita esta línea antes de la llamada.
'ColumnPosition = Column_Position_j
Function Alarms_Check(i As Integer, ColumnPosition As Integer) As Integer
'Obtiene el ranking del estado de cada Part Number.
    
    Dim findstatus As String
    
    findstatus = ws_OG.Cells(i, ColumnPosition).Value

    Set Alarms_Check_i = Range(ws_ranking.Cells(RSRankingi, RSStatusENj), ws_ranking.Cells(RSEndi, RSStatusENj)).Find(findstatus)
    
    If Alarms_Check_i Is Nothing Then
        
        Alarms_Check = 24
        
    Else
        
        Alarms_Check = ws_ranking.Cells(Alarms_Check_i.Row, RSRankingj).Value
        
    End If
    
End Function

Function Complex_Part_Number(pnamei As Integer, i As Integer) As Integer
'-----------------------------------------------Part Numbers con varios materiales (Compuestos)------------------------------------------------
    Dim lasterror As Integer
    Dim material1 As String
    Dim status As String
    Dim status1 As String
    
    lasterror = 0           'Marcador para evitar que las últimas líneas de un Part Number con estado diferente a "OK" no se registren.
    
    Do While nproducto = ws_OG.Cells(pnamei + 1, nprodj).Value              'Con este loop se registran todas las líneas del material, salvo la última.
        
        material = ws_OG.Cells(pnamei, matj).Value
        material1 = ws_OG.Cells(pnamei + 1, matj).Value
        
        status = ws_OG.Cells(pnamei, GlobalStatusj)
        status1 = ws_OG.Cells(pnamei + 1, GlobalStatusj)
        
        If ((material <> material1) Or (material = material1 And status <> status1)) And status <> "OK" Then
        'Condición para evitar la repetición del material.
                                                       
            Call Status_Case(status)
            
        End If
       
        pnamei = pnamei + 1
        
    Loop
          
    material = ws_OG.Cells(pnamei, matj).Value
    material1 = ws_OG.Cells(pnamei - 1, matj).Value
    
    status = ws_OG.Cells(pnamei, GlobalStatusj)
    
    'Condición para añadir el último material del Part Number.
    If nproducto = ws_OG.Cells(pnamei - 1, nprodj).Value And status <> "OK" Then
                                    
        lasterror = 1       'Evita que el bucle del cuerpo se repita eternamente.
        
        status = ws_OG.Cells(pnamei, GlobalStatusj)
        
        Call Status_Case(status)
                                    
    End If
            
    If i <> pnamei Or lasterror = 1 Then
        
        Complex_Part_Number = pnamei
        
    End If
    
End Function

Function Simple_Part_Number(i As Integer) As Integer
'-------------------------------Part Numbers con un material (Simple)---------------------------------
    Dim status As String
    
    If nproducto <> ws_OG.Cells(i + 1, nprodj).Value Then
                 
        status = ws_OG.Cells(i, GlobalStatusj)
        
        Call Status_Case(status)
    
    Else
    
        Do While nproducto = ws_OG.Cells(i + 1, nprodj).Value
        'En caso de que haya varias líneas para el mismo material solo se registra una vez.
            status = ws_OG.Cells(i, GlobalStatusj)
            
            If status = "OK" Or (status = ws_OG.Cells(i - 1, GlobalStatusj) And nproducto = ws_OG.Cells(i - 1, nprodj).Value) Then
                    
                GoTo NextIterarion:
                                                
            End If
            
            Call Status_Case(status)
            
NextIterarion:
    
            i = i + 1
            
            nproducto = ws_OG.Cells(i, nprodj).Value
            
        Loop
          
        status = ws_OG.Cells(i, GlobalStatusj)
        
        'Condición para añadir el último material del Part Number.
        If nproducto = ws_OG.Cells(i - 1, nprodj).Value And status <> "OK" And status <> ws_OG.Cells(i - 1, GlobalStatusj) Then
    
            Call Status_Case(status)
                                        
        End If
    
    End If
    
    Simple_Part_Number = i
    
End Function

Function Status_Case(status As String)
'Genera la información de contacto dependiendo del estado de la línea.

    Dim AuxENRW As String
    Dim AuxESRW As String

    Select Case status
            
        Case "EXPIRED"
            
            AuxENRW = "- Raw material or product name: " & material & " (" & status & ")." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW
        
            AuxESRW = "- Materia prima o part name del producto: " & material & " (EXPIRADO)." + vbCrLf
            InfoESRW = InfoESRW & AuxESRW
            
            auxstatus = 0
            stat = 0        'Estado global del Part Number: EXPIRED.
            
        Case "No date"
            AuxENRW = "- Raw material or product name: " & material & " (" & status & ")." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW

            AuxESRW = "- Materia prima o part name del producto: " & material & " (Sin fecha)." + vbCrLf
            InfoESRW = InfoESRW & AuxESRW
            
        Case Else           'Cuando al certificado le quedan meses o días para expirar.
            AuxENRW = "- Raw material or product name: " & material & " (" & status & " to expire)." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW
            
            Call Spanish_Module(status)
        
    End Select
    
    If exp_material = "" Then
    
        exp_material = material
    
    Else
        
        exp_material = exp_material + vbCrLf + material
    
    End If
    
End Function

Function Spanish_Module(status As String)
'Función para generar la información del correo en español.

    Dim AuxESRW As String
    Dim statusES() As String
    
    statusES = Split(status, " ")
                
    If statusES(1) = "day/s" Then
    
        If stat <> 0 Then

            If stat = 1 Then    'Si la última línea expirará meses y la nueva línea expirará en días.
                   
                auxstatus = statusES(0)
                                          
            End If
            
            stat = 2            'Actualiza el estado global de meses a días.
            
            If statusES(0) < auxstatus Then

                auxstatus = statusES(0)
                                              
            End If
                                                    
        End If
        
        AuxESRW = "- Materia prima o part name del producto: " & material & " (" & statusES(0) & " día/s para expirar)." + vbCrLf
        InfoESRW = InfoESRW & AuxESRW
        
    End If
        
    If statusES(1) = "month/s" Then
               
        If statusES(0) < auxstatus And stat <> 0 And stat <> 2 Then

            auxstatus = statusES(0)
            stat = 1        'Actualiza el estado global a meses.
            
        End If
        
        AuxESRW = "- Materia prima o part name del producto: " & material & " (" & statusES(0) & " mes/es para expirar)." + vbCrLf
        InfoESRW = InfoESRW & AuxESRW
        
    End If

End Function

Function Export_Data(partname_RecordSheet As String, partname_bbdd As String, partname As String, expstatus As String) As Boolean
'Exporta la información de las notificaciones generadas a la Base de Datos "PEDIDOS".
    Dim expi As Integer
    
    Workbooks(partname_bbdd).Sheets("TEMP").Activate                  'Activates the logging Sheet in "PEDIDOS" Data Base.
    
    expi = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row + 1         'Localiza la última línea con información.
    
    Workbooks(partname_RecordSheet).Activate                            'Activa la Base de Datos F&S para extraer la información de ella.
    
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 1).Value = nproducto              'Part Number.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 2).Value = partname               'Part Name.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 3).Value = exp_material           'Materia prima.
    exp_material = ""
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 4).Value = manufacturer           'Proveedor.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 5).Value = "---"                  'Número de Test Report.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 6).Value = Recipient              'E-mail de contacto.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 7).Value = "BB.DD."               'Quién pide la actualización.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 8).Value = Date                   'Fecha de la primera notificación.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 8).NumberFormat = "dd/mm/yyy"
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 9).Value = Date                   'Fecha del último correo enviado.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 9).NumberFormat = "dd/mm/yyy"
    
    Workbooks(partname_bbdd).Sheets("TEMP").Activate                                      'Activate the Data Base "PEDIDOS" to save the loged info.
    Workbooks(partname_bbdd).Sheets("AUX2").Range("A1").Copy Range("J" & expi)            'Validation list.
    
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 11).Value = expstatus             'Test Reports status.
    
    Export_Data = True
    
End Function

Function Email_Display(FinalInfoEN As String, FinalInfoES As String)
'Genera el correo al proveedor.

    'Cabecera Email.
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.session.Logon
    
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
        
        Select Case language
                    
            Case "SPANISH"       'Si el idioma de preferencia es español.
                              
                .Subject = EBSubjectES & manufacturer
                .Body = EBHeadingES & FinalInfoES & EBFarewellES & EBSignature
                              
            Case "ENGLISH"       'Si el idioma de preferencia es inglés.
                
                .Subject = EBSubjectEN & manufacturer
                .Body = EBHeadingEN & FinalInfoEN & EBFarewellEN & EBSignature
                
            Case Else           'Si el idioma de preferencia no es ni inglés ni español.
                
                .Subject = EBSubjectEN & manufacturer
                .Body = EBHeadingEN & FinalInfoEN & EBFarewellEN & EBSeparation & EBHeadingES & FinalInfoES & EBFarewellES & EBSignature
                
        End Select
    
        'Creación del Email
        .To = Recipient
        .CC = EBcc
        .Attachments.Add EBAttachment           'Adjunta el formato de declaración de conformidad vacío.
        
        .Display
            
    End With

End Function
 ­º­º                                                                                                                                                            