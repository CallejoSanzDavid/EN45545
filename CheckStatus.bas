Attribute VB_Name = "CheckStatus"
Sub Check_Status()
'Check if the certificates are OK, EXPIRED or about to EXPIRE.
'Efficiency: 909 lines in 3:12 minutes.

'<------------------------Si da error statusmin iniciarlo en GlobalEntities
    Dim i As Integer
    Dim DateT1j As Integer
    Dim Current_Date As Date
    Dim Dif_Months As Integer
    Dim Dif_Days As Integer
    Dim status0 As Integer
    Dim status1 As String
    Dim statusmin As Integer
    
    Call Locate_Positions_OG
    Call Locate_Positions_RankingStatus
    
    Current_Date = Date
      
    Sheets(SheetName).Cells(Aux + 1, TMexpirej).Select
    
    Call ClearFilters
    
    Call Check_Contacts
    
    For i = Aux + 1 To N
        
        statusmin = 24              'Auxiliar value to prevent bugs in the comparisons
        
        Application.StatusBar = "Checking Certificates Status: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        For DateT1j = Sheets(SheetName).Range("A10:DA10").Find("Date * T1").Column To DateT6j Step 6
            
            status0 = 24            'Auxiliar value to prevent bugs in the comparisons
            
            Call Identify_Status(i, DateT1j, Current_Date, status0, statusmin)

            Call Counters_Check
            
        Next
        
    Next
    
    Application.StatusBar = ""
    
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
    manufacturer = Sheets(SheetName).Cells(Aux + 1, manufj).Value
    
    For m = Aux + 1 To N
        
        Application.StatusBar = "Updating Supplier's Contact Information: " & m - Aux & " of " & N - Aux & ": " & Format((m - Aux) / (N - Aux), "0%")
        
        If auxmanufacturer <> manufacturer Then     'Only finds the contact info once.
            
            auxmanufacturer = manufacturer
            Set c = Range(Sheets(ContactSheetName).Cells(CPsupplieri + 1, CPsupplierj), Sheets(ContactSheetName).Cells(CPendi, CPsupplierj)).Find(manufacturer)
        
        End If
        
        If c Is Nothing Then                        'No contact info.
        
            Sheets(SheetName).Cells(m, ContactDBj) = "Does NOT Exist"
            Sheets(SheetName).Cells(m, ContactDBj).Interior.ColorIndex = 3
        
            linea = 0
            
        Else
        
            linea = c.Row
            
            If Sheets(ContactSheetName).Cells(linea, CPmailj) = "" Then     'Exists the Provider in the list but has no contact info.
            
                Sheets(SheetName).Cells(m, ContactDBj) = "Does NOT Exist"
                Sheets(SheetName).Cells(m, ContactDBj).Interior.ColorIndex = 3
                
                linea = 0
                
            End If
            
        End If
        
        If linea <> 0 Then      'Exists contact info.
                
            Sheets(SheetName).Cells(m, ContactDBj) = Sheets(ContactSheetName).Cells(linea, CPmailj)
            Sheets(SheetName).Cells(m, ContactDBj).Interior.ColorIndex = 43
             
        End If
        
        manufacturer = Sheets(SheetName).Cells(m + 1, manufj).Value
        
    Next
    
    Application.StatusBar = ""
    
End Function

Function Identify_Status(i, DateT1j, Current_Date, status0, statusmin)
'Calls the funtions in order to correctly identify and log the certificates status.

    Dim ColumnPosition As Integer

    ColumnPosition = DateT1j
    status1 = Check_Dates(i, ColumnPosition, Current_Date, status0, statusmin)
    
    If status0 <> 23 And Sheets(SheetName).Cells(i, ManufDeclarationj) <> "" And IsDate(Sheets(SheetName).Cells(i, ManufDeclarationj)) Then
        
        ColumnPosition = ManufDeclarationj
        status1 = Check_Dates(i, ColumnPosition, Current_Date, status0, statusmin)
        
    End If
        
    ColumnPosition = TMexpirej
    Call Global_Status(i, ColumnPosition, statusmin, status0, status1)
    
    If status0 < statusmin Then
        
        statusmin = status0        'Logs the new minimum status.
        ColumnPosition = GlobalStatusj
        Call Global_Status(i, ColumnPosition, statusmin, status0, status1)
    
    End If
    
End Function

Function Check_Dates(i, ColumnPosition, Current_Date, status0, statusmin) As String
'Compares the Certificates and Manufacturers' declarations dates and logs the Part Number status.

    If Sheets(SheetName).Cells(i, ColumnPosition) <> "" And IsDate(Sheets(SheetName).Cells(i, ColumnPosition)) Then
    
        Dif_Months = 60 - DateDiff("m", Sheets(SheetName).Cells(i, ColumnPosition), Current_Date)
        Dif_Days = 1827 - DateDiff("d", Sheets(SheetName).Cells(i, ColumnPosition), Current_Date)
    
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
Function Global_Status(i, ColumnPosition, statusmin, status0, status1)
'Logs the Global Status of each Part Number.
    
    Sheets(SheetName).Cells(i, ColumnPosition).Value = status1

    Set findstatus = Range(Sheets(RankingStatusSheet).Cells(RSRankingi, RSStatusENj), Sheets(RankingStatusSheet).Cells(RSEndi, RSStatusENj)).Find(status1)
    
    Sheets(SheetName).Cells(i, ColumnPosition).Interior.ColorIndex = Sheets(RankingStatusSheet).Cells(findstatus.Row, RSColorCodej).Value
    
End Function

Function Counters_Check()
'Adds or resets counters
    If TMexpirej = Sheets(SheetName).Range("A10:DA10").Find("Test Method 1 time to expire*").Column + 5 Then

        TMexpirej = Sheets(SheetName).Range("A10:DA10").Find("Test Method 1 time to expire*").Column
        
    Else
    
        TMexpirej = TMexpirej + 1
        
    End If
            
End Function
