Attribute VB_Name = "EmailGen"
Sub Email_Gen()
'Creates emails with the information of expired or about to expire to its pertinent supplier.
    
    'counters
    Dim nmails As Integer
    Dim nnocontact As Integer
    Dim nexport As Integer
    'flags
    Dim Export As Integer
    Dim NoContact As Integer
    Dim validation As Integer
    'splitters
    Dim Auxsplit As String
    Dim auxname() As String
    'finders and identifiers
    Dim pnamei As Integer
    Dim ColumnPosition As Integer
    Dim status As String
    Dim statusmin As Integer
    'workbooks names
    Dim partname_RecordSheet As String
    Dim partname_bbdd As String
    'email info
    Dim FinalInfoEN As String
    Dim FinalInfoES As String
                    
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    Call Locate_Positions_OG
    
    '<-----------------------------------
    Call Format_Capitalization
    
    Sheets(SheetName).Cells(Aux + 1, GlobalStatusj).Select
    
    '<-----------------------------------
    TableName = ActiveSheet.ListObjects(1).Name
    Call ClearFilters
    'Sorts Part Names in Alfabetic Order.
    FilterSet = Sheets(SheetName).Cells(Aux, nombj).Value
    Call AlfabeticOrder
    'Sorts Part Numbers in Alfabetic Order.
    FilterSet = Sheets(SheetName).Cells(Aux, nprodj).Value
    Call AlfabeticOrder
    'Sorts Suppliers in Alfabetic Order.
    FilterSet = Sheets(SheetName).Cells(Aux, manufj).Value
    Call AlfabeticOrder
    
    Call Locate_Positions_Contacts
    
    Call Locate_Positions_Email_Body
    
    Call Email_Body
    
    Call Locate_Positions_RankingStatus
    
    'flags initial values
    nmails = 0
    nnocontact = 0
    nexport = 0
    Export = 0
    OpenDatabase = 0
    
    For i = Aux + 1 To N

        NoContact = Manufacturer_Contact(i, nnocontact)
        
        If NoContact = 0 Then
        
            GoTo NoContact:     'If there is no contact goes to the next line.
        
        End If
    
NextPartNumber:
        
        Application.StatusBar = "Checking expired certificates and generating emails: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        ColumnPosition = GlobalStatusj
        statusmin = Alarms_Check(i, ColumnPosition)
        
        ColumnPosition = EmailSendedj
        validation = Alarms_Check(i, ColumnPosition)
        
        'flags initial values
        stat = 3                'flag to identify the minimum status.
                                '3 - auxiliar value
                                '2 - day/s
                                '1 - month/s
                                '0 - EXPIRED
        auxstatus = 30          'auxiliar value to identify the minimum status.
        marc1 = 0               'flags to identify the correct loop:
                                '0 - Initial value: Part Number with a single material.
                                '1 - Part Number with several materials.
        
        If statusmin <= 21 And validation > statusmin Then
            
            nproducto = Sheets(SheetName).Cells(i, nprodj).Value
            
            Auxsplit = "0"            'flag initialized as "0" to detect if the Part Number has several materials.
            auxname = Split(Cells(i, nombj).Value, " - MATERIAL")
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
            
            material = Sheets(SheetName).Cells(i, matj).Value
                         
            If manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value Then   'If the next supplier is the same.
                
                status = Sheets(SheetName).Cells(i, GlobalStatusj)
                pnamei = Sheets(SheetName).Range(Cells(Aux, nprodj), Cells(N, nprodj)).Find(nproducto).Row
                
                '-------------------------------Part Numbers with several materials---------------------------------
                If Auxsplit <> "0" And status <> "OK" Then
                    
                    i = Complex_Part_Number(pnamei, i)
                    marc1 = 1       'flag value for Part Number type
                        
                End If
                    
                '-------------------------------Part Numbers with one material---------------------------------
                If Auxsplit = "0" And marc1 = 0 And status <> "OK" Then
                                            
                    i = Simple_Part_Number(pnamei, i)
                    
                End If
                
                Select Case stat
                    
                    Case 2        'Days left for the certificate to expire.
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & partname & " (" & auxstatus & " day/s to expire)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & partname & " (" & auxstatus & " día/s para expirar)." + vbCrLf
                        expstatus = auxstatus & " día/s para expirar"
                                      
                    Case 1        'Months left for the certificate to expire.
                        
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & partname & " (" & auxstatus & " month/s to expire)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & partname & " (" & auxstatus & " mes/es para expirar)." + vbCrLf
                        expstatus = auxstatus & " mes/es para expirar"
                        
                    Case 0        'Any of the materials has expired.
                        
                        InfoEN = "- MERAK part number: " & nproducto & "." + vbCrLf + "- MERAK part name: " & partname & " (EXPIRED)." + vbCrLf
                        InfoES = "- Número del elemento de MERAK: " & nproducto & "." + vbCrLf + "- Nombre del elemento MERAK: " & partname & " (EXPIRADO)." + vbCrLf
                        expstatus = "EXPIRADO"
                        
                End Select
                
                Export = 1
                
                FinalInfoEN = FinalInfoEN & InfoEN & InfoENRW + vbCrLf
                FinalInfoES = FinalInfoES & InfoES & InfoESRW + vbCrLf
                
                InfoENRW = ""
                InfoESRW = ""
                
                If Export = 1 Then         'A new line is loged for each Part Number in the Data Base "Pedidos".
                    
                    If OpenDatabase = 0 Then
                    
                        partname_RecordSheet = ActiveWorkbook.Name
                        
                        Workbooks.Open (Sheets("Validation Lists and Routes").Range("G2").Value)
                        
                        partname_bbdd = ActiveWorkbook.Name
                        
                        OpenDatabase = 1
                        
                    End If
                    
                    Call Export_Data(partname_RecordSheet, partname_bbdd)
                    Workbooks(partname_RecordSheet).Activate
                    
                    nexport = nexport + 1
                    
                    If manufacturer = Sheets(SheetName).Cells(i + 1, manufj).Value Then
                    
                        i = i + 1               'With this line we prevent the code to analize the last line again.
                        GoTo NextPartNumber:    'Starts the loop again skipping the "Manufacturer_Contact" function.
                    
                    End If
                    
                End If
                
            End If
                
        End If
        
        If Export = 1 And manufacturer <> Sheets(SheetName).Cells(i + 1, manufj).Value Then
            
            Call Email_Display(FinalInfoEN, FinalInfoES)
            nmails = nmails + 1
            
            Export = 0
            
            FinalInfoEN = ""
            FinalInfoES = ""
            
        End If
        
NoContact:

    Next
    
    Workbooks(partname_bbdd).Activate
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    Workbooks(partname_RecordSheet).Activate
    
    MsgBox (nnocontact & " elemento/s expirado/s no tiene/n información de contacto." + vbCrLf + vbCrLf + "Se han generado " & nmails & " correo/s para " & nexport & " part numbers.")
    
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
    
    EBcc = Sheets(EmailBodySheetName).Cells(EBcci, EBInfoj).Value
    
    EBSubjectEN = Sheets(EmailBodySheetName).Cells(EBSubjectENi, EBInfoj).Value
    EBSubjectES = Sheets(EmailBodySheetName).Cells(EBSubjectESi, EBInfoj).Value
    
    EBAttachment = Sheets(EmailBodySheetName).Cells(EBAttachmenti, EBInfoj).Value
    
    EBHeadingEN = Sheets(EmailBodySheetName).Cells(EBHeadingENi, EBInfoj).Value
    EBFarewellEN = Sheets(EmailBodySheetName).Cells(EBFarewellENi, EBInfoj).Value
    
    EBSeparation = Sheets(EmailBodySheetName).Cells(EBSeparationi, EBInfoj).Value
    
    EBHeadingES = Sheets(EmailBodySheetName).Cells(EBHeadingESi, EBInfoj).Value
    EBFarewellES = Sheets(EmailBodySheetName).Cells(EBFarewellESi, EBInfoj).Value
    
    EBSignature = Sheets(EmailBodySheetName).Cells(EBSignaturei, EBInfoj).Value

End Function

Function Manufacturer_Contact(i, nnocontact) As Integer
    
    Dim CPmaili As Integer
    
    Manufacturer_Contact = 1
    
    manufacturer = Sheets(SheetName).Cells(i, manufj).Value
    
    Recipient = Sheets(SheetName).Cells(i, ContactDBj).Value
    
    If Recipient = "Does NOT Exist" Then
    
        nnocontact = nnocontact + 1
        Manufacturer_Contact = 0
        Exit Function
        
    End If
    
    Sheets(ContactSheetName).Activate   'To prevent an error in the next code line, we activate the Sheet.
    CPmaili = Sheets(ContactSheetName).Range(Cells(1, CPmailj), Cells(CPendi, CPmailj)).Find(Recipient).Row
    
    Do While Recipient <> "Does NOT Exist" And Sheets(ContactSheetName).Cells(CPmaili, CPsupplierj).Value = Sheets(ContactSheetName).Cells(CPmaili + 1, CPsupplierj).Value
    'Loop to send the email to all the contacts.
        
        Recipient = Recipient & "; " & Sheets(ContactSheetName).Cells(CPmaili + 1, CPmailj).Value
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

Function Complex_Part_Number(pnamei, i) As Integer

'-----------------------------------------------Part Numbers with several materials------------------------------------------------
    Dim lasterror As Integer
    Dim material1 As String
    
    lasterror = 0           'flag to prevent the error in which the last lines are not loged if it's not OK or are the same material.
    
    Do While nproducto = Sheets(SheetName).Cells(pnamei + 1, nprodj).Value              'Loop to log all the Part Number materials.
        
        material = Sheets(SheetName).Cells(pnamei, matj).Value
        material1 = Sheets(SheetName).Cells(pnamei + 1, matj).Value
        
        status = Sheets(SheetName).Cells(pnamei, GlobalStatusj)
        
        If material <> material1 And status <> "OK" Then                                'Condition to prevent the repetition of a material.
                                                       
            Call Status_Case(status)
            
        End If
       
        pnamei = pnamei + 1
        
    Loop
          
    material = Sheets(SheetName).Cells(pnamei, matj).Value
    material1 = Sheets(SheetName).Cells(pnamei - 1, matj).Value
    
    status = Sheets(SheetName).Cells(pnamei, GlobalStatusj)
    
    'Condition to add the last material of the Part number.
    If (material <> material1 And nproducto = Sheets(SheetName).Cells(pnamei - 1, nprodj).Value And status <> "OK") Or (material = material1 And status <> "OK") Then
                                    
        lasterror = 1       'Prevents the part number to be loged infinetly.
        
        status = Sheets(SheetName).Cells(pnamei, GlobalStatusj)
        
        Call Status_Case(status)
                                    
    End If
            
    If i <> pnamei Or lasterror = 1 Then
        
        Complex_Part_Number = pnamei
        
    End If
    
End Function

Function Simple_Part_Number(pnamei, i) As Integer

'-------------------------------Part Numbers with one material---------------------------------
    If nproducto <> Sheets(SheetName).Cells(i + 1, nprodj).Value Then
                        
        status = Sheets(SheetName).Cells(i, GlobalStatusj)
        
        Call Status_Case(status)
    
    End If
    
    
    Do While nproducto = Sheets(SheetName).Cells(i + 1, nprodj).Value        'In case there are several lines for the same Part Number, only logs the most restrictive.

        status = Sheets(SheetName).Cells(i, GlobalStatusj)
        
        If status = "OK" Or (status = Sheets(SheetName).Cells(i + 1, GlobalStatusj) And nproducto = Sheets(SheetName).Cells(i - 1, nprodj).Value) Then
                
            statusES(1) = 0
            GoTo NextIterarion:
                                            
        End If
        
        Call Status_Case(status)
        
NextIterarion:

        i = i + 1
        
        nproducto = Sheets(SheetName).Cells(i, nprodj).Value
        
    Loop
      
    status = Sheets(SheetName).Cells(i, GlobalStatusj)
    
    'Condition to log last material.
    If nproducto = Sheets(SheetName).Cells(pnamei - 1, nprodj).Value And status <> "OK" And status <> Sheets(SheetName).Cells(i - 1, GlobalStatusj) Then
        'STOP
        Call Status_Case(status)
                                    
    End If
    
    Simple_Part_Number = i
    
End Function

Function Status_Case(status)
'Funtion to generate the information of the expired or about to expire material according to its status.

    Select Case status
            
        Case "EXPIRED"
            
            AuxENRW = "- Raw material or product name: " & material & " (" & status & ")." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW
        
            AuxESRW = "- Materia prima o partname del producto: " & material & " (EXPIRADO)." + vbCrLf
            InfoESRW = InfoESRW & AuxESRW
            
            auxstatus = 0
            stat = 0        'Global status of the Part Number. EXPIRED.
            
        Case "No date"
            AuxENRW = "- Raw material or product name: " & material & " (" & status & ")." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW

            AuxESRW = "- Materia prima o partname del producto: " & material & " (Sin fecha)." + vbCrLf
            InfoESRW = InfoESRW & AuxESRW
            
        Case Else           'When the certificates have months or days to expire left.
            AuxENRW = "- Raw material or product name: " & material & " (" & status & " to expire)." + vbCrLf
            InfoENRW = InfoENRW & AuxENRW
            
            Call Spanish_Module(status)
        
    End Select

End Function

Function Spanish_Module(status)
'Function to log the info in spanish.
    statusES = Split(status, " ")
                
    If statusES(1) = "day/s" Then
    
        If stat <> 0 Then

            If stat = 1 Then    'If the last status was expiring in months and the new line will expire in days.
                   
                auxstatus = statusES(0)
                                          
            End If
            
            stat = 2
            
            If statusES(0) < auxstatus Then    'Updates the global status from months to days.

                auxstatus = statusES(0)
                                              
            End If
                                                    
        End If
        
        AuxESRW = "- Materia prima o partname del producto: " & material & " (" & statusES(0) & " día/s para expirar)." + vbCrLf
        InfoESRW = InfoESRW & AuxESRW
        
    End If
        
    If statusES(1) = "month/s" Then
        
        AuxESRW = "- Materia prima o partname del producto: " & material & " (" & statusES(0) & " mes/es para expirar)." + vbCrLf
        InfoESRW = InfoESRW & AuxESRW
        
        If statusES(0) < auxstatus And stat <> 0 And stat <> 2 Then     'Updates the global status to months.

            auxstatus = statusES(0)
            stat = 1
            
        End If
        
    End If

End Function

Function Export_Data(partname_RecordSheet, partname_bbdd)
'Exports to the "PEDIDOS" Data Base the info for the notified Part Numbers.
    Dim expi As Integer
    
    Workbooks(partname_bbdd).Sheets("TEMP").Activate                  'Activates the logging Sheet in "PEDIDOS" Data Base.
    
    expi = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row + 1     'Locates the last line with info.
    
    Workbooks(partname_RecordSheet).Activate                          'Activate the F&S Data Base to extract the info from it.
    
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 1).Value = nproducto              'Part Number.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 2).Value = partname                 'Part Name.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 3).Value = material               'Raw Material.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 4).Value = manufacturer           'Supplier.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 5).Value = "---"                  'TR number.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 6).Value = Recipient              'E-mail Contact.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 7).Value = "BB.DD."               'Who demands the update.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 8).Value = Date                   'Date of the first notification.
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 9).Value = Date                   'Date of the last email sended.
    
    Workbooks(partname_bbdd).Sheets("TEMP").Activate                                      'Activate the Data Base "PEDIDOS" to save the loged info.
    Workbooks(partname_bbdd).Sheets("AUX2").Range("A1").Copy Range("J" & expi)            'Validation list.
    
    Workbooks(partname_bbdd).Sheets("TEMP").Cells(expi, 11).Value = expstatus             'Test Reports status.

End Function

Function Email_Display(FinalInfoEN, FinalInfoES)

    'Email Header.
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.session.Logon
    
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
    
        'Email generation
        .To = Recipient
        .CC = "f&s@merak-hvac.com"
        '.SentOnBehalfOfName = "f&s@merak-hvac.com"
        .Attachments.Add EBAttachment           '"T:\Compartir\F&S Certificates\20150223_Manufacturer_Declaration.doc"
        .Subject = EBSubjectEN & manufacturer
        .Body = EBHeadingEN & FinalInfoEN & EBFarewellEN & EBSeparation & EBHeadingES & FinalInfoES & EBFarewellES & EBSignature
        .Display
        '.Send
    
    End With

End Function
