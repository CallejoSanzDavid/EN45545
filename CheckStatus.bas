Attribute VB_Name = "CheckStatus"
Sub Check_Status()           'Comprueba el estado de los certificados.
    
    Dim i As Integer
    Dim DateT1j As Integer
    Dim Current_Date As Date
    Dim Dif_Months As Integer
    Dim Dif_Days As Integer
    Dim Dif_MonthsDC As Integer
    Dim Dif_DaysDC As Integer
    Dim status0 As Integer
    Dim status1 As String
    Dim auxstatus1(1) As String
    Dim statusmin As Integer
    Dim ManufDeclaration As Integer
    Dim No_Date_flag As Integer
    
    Call Locate_Positions_OG
    
    Current_Date = Date
      
    Sheets(SheetName).Cells(Aux + 1, TMexpirej).Select
    '<----------------------
    'Call ClearFilters
    '<----------------------
    'Call Check_Contacts
    
    For i = Aux + 1 To N
            
        'STOP<----------------------------------
        'Error: Test Method 5 se registra incorrectamente. "0 month/s", pero no tiene fecha.
        i = 156
            
        statusmin = 24              'Auxiliar value to prevent bugs in the comparisons
        
        Application.StatusBar = "Checking Certificates Status: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        For DateT1j = Sheets(SheetName).Range("A10:DA10").Find("Date * T1").Column To DateT6j Step 6
            
            No_Date_flag = 0
            status0 = 24            'Auxiliar value to prevent bugs in the comparisons
            
            If IsDate(Sheets(SheetName).Cells(i, DateT1j)) = False Then                   'Error: Cell with no date.

                No_Date_flag = No_Date(i, DateT1j, DateT6j, TMexpirej, status0, statusmin)
                Exit For            'Salir del for al terminar la función ¿?
            
            End If
                            
            If No_Date_flag = 0 Then

                Call Check_Dates(i, DateT1j, Current_Date, ManufDeclarationj, TMexpirej, status0, statusmin)
            
            End If
            
            Call Counters_Check(TMexpirej)
            
        Next
        
    Next
    
    Application.StatusBar = ""
    
End Sub

Function No_Date(i, DateT1j, DateT6j, TMexpirej, status0, statusmin) As Integer
'Logs the Certificates with No Date.

    Do While DateT1j <= DateT6j
        
        status0 = 23
        status1 = "No date"
        
        Sheets(SheetName).Cells(i, TMexpirej).Value = status1
        Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 2      'No Date: White.
        
        No_Date = 1
                      
        If Sheets(SheetName).Cells(i, ManufDeclarationj) <> "" And IsDate(Sheets(SheetName).Cells(i, ManufDeclarationj)) Then
'STOP
            Call Check_Dates(i, DateT1j, Current_Date, ManufDeclarationj, TMexpirej, status0, statusmin)
        
        End If
        
        If status0 < statusmin Then

            Call Global_Status(i, statusmin, status0, status1)
        
        End If

        Call Counters_Check(TMexpirej)
        
        DateT1j = DateT1j + 6
        
    Loop

End Function

Function Counters_Check(TMexpirej)
'Adds or resets counters
    If TMexpirej = Sheets(SheetName).Range("A10:DA10").Find("Test Method 1 time to expire*").Column + 5 Then

        TMexpirej = Sheets(SheetName).Range("A10:DA10").Find("Test Method 1 time to expire*").Column
        
    Else
    
        TMexpirej = TMexpirej + 1
        
    End If
            
End Function

Function Check_Dates(i, DateT1j, Current_Date, ManufDeclarationj, TMexpirej, status0, statusmin)
'Compares the Certificates and Manufacturers' declarations dates and logs the Part Number status.
    'Función repetida.
    If Sheets(SheetName).Cells(i, DateT1j) <> "" And IsDate(Sheets(SheetName).Cells(i, DateT1j)) Then
    
        Dif_Months = 60 - DateDiff("m", Sheets(SheetName).Cells(i, DateT1j), Current_Date)
        Dif_Days = 1827 - DateDiff("d", Sheets(SheetName).Cells(i, DateT1j), Current_Date)
    
    Else

        Dif_Months = 0
        Dif_Days = 0
        
    End If
    'Función repetida.
    If Sheets(SheetName).Cells(i, ManufDeclarationj) <> "" And IsDate(Sheets(SheetName).Cells(i, ManufDeclarationj)) Then

        Dif_MonthDC = 60 - DateDiff("m", Sheets(SheetName).Cells(i, ManufDeclarationj), Current_Date)
        Dif_DaysDC = 1827 - DateDiff("d", Sheets(SheetName).Cells(i, ManufDeclarationj), Current_Date)
        
    Else
    
        Dif_MonthsDC = 0
        Dif_DaysDC = 0
        
    End If
    
    If status0 <> 23 And (Dif_Months > 6 Or Dif_MonthsDC > 6) Then                    'Si faltan más de 6 meses para que expire: OK
'STOP: AQUÍ TENGO QUE MIRAR DE RESOLVER EL TEMA DEL ESTADO GLOBAL.
        status0 = 22
        status1 = "OK"
        
        Sheets(SheetName).Cells(i, TMexpirej) = status1
        Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 4  'Verde si es OK
        
    End If
    'Reescribir función.
    If Dif_Months <= 6 And Dif_MonthsDC <= 6 Then                     'Si faltan menos de 6 meses para que expire
'STOP
        Sheets(SheetName).Cells(i, TMexpirej).Value = Dif_Months & " month/s"
        Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 6      'Amarillo si falta entre 6 y 3 meses
        status0 = 15 + Dif_Months
        status1 = Dif_Months & " month/s"
        
        If Dif_Months <= 3 And Dif_MonthsDC <= 3 Then
        
            Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 44 'Amarillo oscuro si está entre 3 y 2 meses
            
            If Dif_Months <= 2 And Dif_MonthsDC <= 2 Then
        
                Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 45 'Naranja claro si está entre 2 y 1 mes/es.
            
                If Dif_Months <= 1 And Dif_MonthsDC <= 1 And Dif_Days <= 30 And Dif_DaysDC <= 30 Then   'Si faltan días para que expire
            
                    Sheets(SheetName).Cells(i, TMexpirej).Value = Dif_Days & " day/s"

                    status1 = Dif_Days & " day/s"
                                             
                    Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 46 'Naranja oscuro faltan entre 30 y 1 días
                    
                    If Dif_Days <= 15 And Dif_DaysDC <= 15 Then
                        
                        status0 = Dif_Days
                    
                        If Dif_Days <= 0 And Dif_DaysDC <= 0 Then
                
                            status0 = 0
                            status1 = "EXPIRED"
                            Sheets(SheetName).Cells(i, TMexpirej).Value = status1
                            Sheets(SheetName).Cells(i, TMexpirej).Interior.ColorIndex = 3  'Rojo si está caducado
                            
                        End If
                    
                    End If
            
                End If
                
            End If
            
        End If
        
    End If
    
    If status0 < statusmin Then

        Call Global_Status(i, statusmin, status0, status1)
        
    End If
    
End Function

Function Global_Status(i, statusmin, status0, status1)    'Marca el estado global del Part number
'Logs the Global Status of each Part Number.
              
    Sheets(SheetName).Cells(i, GlobalStatusj).Value = status1
        
    statusmin = status0        'Logs the new minimum status.
    
    Select Case statusmin
           
        Case 23
            Sheets(SheetName).Cells(i, GlobalStatusj).Interior.ColorIndex = 2       'White: No Date.
        
        Case 22
            Sheets(SheetName).Cells(i, GlobalStatusj).Interior.ColorIndex = 4       'Green: OK.
        
        Case 19 To 21
            Sheets(SheetName).Cells(i, GlobalStatusj).Interior.ColorIndex = 6       'Yellow: Between 6 to 3 months.
'STOP
        Case 17, 18
            Sheets(SheetName).Cells(i, GlobalStatusj).Interior.ColorIndex = 44      'Dark Yellow: Between 3 to 2 months.
        
        Case 16
            Sheets(SheetName).Cells(i, GlobalStatusj).Interior.ColorIndex = 45      'Orange: Between 2 to 1 month/s.
        
        Case 0
            Sheets(SheetName).Cells(i, GlobalStatusj).Interior.ColorIndex = 3       'Red: EXPIRED.
    
        Case Else        'Less than 1 month to expire or other.
'STOP
            daynum = 0
            auxday = Split(Cells(i, G).Value, " day/s")
            daynum = auxday(0)
            
            If statusmin <= 16 Then         'And m_d = "day/s"
                Sheets(SheetName).Cells(i, G).Interior.ColorIndex = 46              'Dark Orange: Between 30 to 1 day
            End If
            
    End Select
    
End Function


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
            Set c = Range(Sheets(ContactSheetName).Cells(CPsupplieri, CPsupplierj), Sheets(ContactSheetName).Cells(CPendi, CPsupplierj)).Find(manufacturer)
        
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
