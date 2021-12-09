Attribute VB_Name = "CheckStatus"
Sub Check_Status()
'Check if the certificates are OK, EXPIRED or about to EXPIRE.
'Efficiency: 909 lines in 3:12 minutes.

'<------------------------Si da error statusmin iniciarlo en GlobalEntities
    Dim i As Integer
    Dim DateT1j As Integer
    Dim Current_Date As Date
    Dim status0 As Integer
    Dim status1 As String
    Dim statusmin As Integer
    
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    t1 = Time
    
    Call Locate_Positions_OG
    Call Locate_Positions_RankingStatus
    
    Current_Date = Date
      
    ws_OG.Cells(Aux + 1, TMexpirej).Select
    
    'Clears the filters and sorts the Part Numbers by alfabetic order.
    TableName = ActiveSheet.ListObjects(1).Name
    FilterSet = ws_OG.Cells(Aux, nprodj).Value
    Call ClearFilters
    Call AlfabeticOrder
    
    Call Check_Contacts
    
    For i = Aux + 1 To N
        
        statusmin = 24              'Auxiliar value to prevent bugs in the comparisons
        
        Application.StatusBar = "Checking Certificates Status: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        For DateT1j = ws_OG.Range("A10:DA10").Find("Date * T1").Column To DateT6j Step 6
            
            status0 = 24            'Auxiliar value to prevent bugs in the comparisons
            
            Call Identify_Status(i, DateT1j, Current_Date, status0, statusmin)

            Call Counters_Check
            
        Next
        
    Next
    
    t2 = Time
    crono = Format(t2 - t1, "hh:mm:ss")
    
    MsgBox ("Operación ejecutada con éxito." + vbCrLf + vbCrLf + "Tiempo de operación: " & crono & ".")
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Sub

Function Check_Contacts()
'Finds and logs the contact information in the Supplier's Contact column.
'Efficiency: 909 lines in 29.10 seconds.
    
    Dim manufacturer As String
    Dim auxmanufacturer As String
    Dim m As Integer
    Dim linea As Integer
    Dim c As Range
    
    Call Locate_Positions_Contacts
    
    auxmanufacturer = ""
    manufacturer = ws_OG.Cells(Aux + 1, manufj).Value
    
    For m = Aux + 1 To N
        
        Application.StatusBar = "Updating Supplier's Contact Information: " & m - Aux & " of " & N - Aux & ": " & Format((m - Aux) / (N - Aux), "0%")
        
        If auxmanufacturer <> manufacturer Then     'Only finds the contact info once.
            
            auxmanufacturer = manufacturer
            Set c = Range(ws_contact.Cells(CPsupplieri + 1, CPsupplierj), ws_contact.Cells(CPendi, CPsupplierj)).Find(manufacturer)
        
        End If
        
        If c Is Nothing Then                        'No contact info.
        
            ws_OG.Cells(m, ContactDBj) = "Does NOT Exist"
            ws_OG.Cells(m, ContactDBj).Interior.ColorIndex = 3
        
            linea = 0
            
        Else
        
            linea = c.Row
            
            If ws_contact.Cells(linea, CPmailj) = "" Then     'Exists the Provider in the list but has no contact info.
            
                ws_OG.Cells(m, ContactDBj) = "Does NOT Exist"
                ws_OG.Cells(m, ContactDBj).Interior.ColorIndex = 3
                
                linea = 0
                
            End If
            
        End If
        
        If linea <> 0 Then      'Exists contact info.
                
            ws_OG.Cells(m, ContactDBj) = ws_contact.Cells(linea, CPmailj)
            ws_OG.Cells(m, ContactDBj).Interior.ColorIndex = 43
             
        End If
        
        manufacturer = ws_OG.Cells(m + 1, manufj).Value
        
    Next
    
    Application.StatusBar = ""
    
End Function

Function Identify_Status(i, DateT1j, Current_Date, status0, statusmin)
'Calls the funtions in order to correctly identify and log the certificates status.

    Dim ColumnPosition As Integer

    ColumnPosition = DateT1j
    status1 = Check_Dates(i, ColumnPosition, Current_Date, status0, statusmin)
    
    If status0 <> 23 And ws_OG.Cells(i, ManufDeclarationj) <> "" And IsDate(ws_OG.Cells(i, ManufDeclarationj)) Then
        
        ColumnPosition = ManufDeclarationj
        status1 = Check_Dates(i, ColumnPosition, Current_Date, status0, statusmin)
        
    End If
        
    ColumnPosition = TMexpirej
    Call Log_Status(i, ColumnPosition, statusmin, status0, status1)
    
    If status0 < statusmin Then
        
        statusmin = status0        'Logs the new minimum status.
        ColumnPosition = GlobalStatusj
        Call Log_Status(i, ColumnPosition, statusmin, status0, status1)
    
    End If
    
End Function

Function Check_Dates(i, ColumnPosition, Current_Date, status0, statusmin) As String
'Compares the Certificates and Manufacturers' declarations dates and logs the Part Number status.
    
    Dim Dif_Months As Integer
    Dim Dif_Days As Integer
    
    If ws_OG.Cells(i, ColumnPosition) <> "" And IsDate(ws_OG.Cells(i, ColumnPosition)) Then
    
        Dif_Months = 60 - DateDiff("m", ws_OG.Cells(i, ColumnPosition), Current_Date)
        Dif_Days = 1827 - DateDiff("d", ws_OG.Cells(i, ColumnPosition), Current_Date)
    
    Else

        status0 = 23
        Check_Dates = "No date"
        Exit Function
        
    End If
    
    Select Case Dif_Months

        Case Is > 6
            status0 = 22
            Check_Dates = "OK"
            Exit Function
    
        Case 2 To 6
            status0 = 15 + Dif_Months
            Check_Dates = Dif_Months & " month/s"
            Exit Function
            
        Case Is <= 1
            
            Select Case Dif_Days
            
            Case Is > 16
                status0 = 16
                Check_Dates = "1 month/s"
                Exit Function
            
            Case 1 To 15
                status0 = Dif_Days
                Check_Dates = Dif_Days & " day/s"
                Exit Function
                
            Case Is < 1
                status0 = 0
                Check_Dates = "EXPIRED"
                Exit Function
        
        End Select
    
    End Select
    
End Function

'Needs to have this line before the Call:
'ColumnPosition = Column_Position_j
Function Log_Status(i, ColumnPosition, statusmin, status0, status1)
'Logs the Global Status of each Part Number.
    
    ws_OG.Cells(i, ColumnPosition).Value = status1

    Set findstatus = Range(ws_ranking.Cells(RSRankingi, RSStatusENj), ws_ranking.Cells(RSEndi, RSStatusENj)).Find(status1)
    
    ws_OG.Cells(i, ColumnPosition).Interior.ColorIndex = ws_ranking.Cells(findstatus.Row, RSColorCodej).Value
    
End Function

Function Counters_Check()
'Adds or resets counters
    If TMexpirej = ws_OG.Range("A10:DA10").Find("Test Method 1 time to expire*").Column + 5 Then

        TMexpirej = ws_OG.Range("A10:DA10").Find("Test Method 1 time to expire*").Column
        
    Else
    
        TMexpirej = TMexpirej + 1
        
    End If
            
End Function
