Attribute VB_Name = "CheckStatus"
Sub Check_Status()
'Comprueba el estado de los certificados: OK, EXPIRED o about to EXPIRE (month/s, day/s).

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
    
    'Elimina todos los filtros y ordena los Part Numbers por orden alfab�tico.
    TableName = ActiveSheet.ListObjects(1).Name
    FilterSet = ws_OG.Cells(Aux, nprodj).Value
    Call ClearFilters
    Call AlfabeticOrder
    
    Call Check_Contacts
    
    For i = Aux + 1 To N
        
        statusmin = 24              'Valor auxiliara para evitar errores al comparar.
        
        Application.StatusBar = "Checking Certificates Status: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        For DateT1j = ws_OG.Range("A10:DA10").Find("Date * T1").Column To DateT6j Step 6
            
            status0 = 24            'Valor auxiliara para evitar errores al comparar.
            
            Call Identify_Status(i, DateT1j, Current_Date, status0, statusmin)

            Call Counters_Check
            
        Next
        
    Next
    
    t2 = Time
    crono = Format(t2 - t1, "hh:mm:ss")
    
    MsgBox ("Operaci�n ejecutada con �xito." + vbCrLf + vbCrLf + "Tiempo de operaci�n: " & crono & ".")
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Sub

Function Check_Contacts()
'Busca y registra la informaci�n de contacto del proveedor en la columna "Supplier's Contact".
    
    Dim manufacturer As String
    Dim auxmanufacturer As String
    Dim m As Integer
    Dim c As Range
    
    Call Locate_Positions_Contacts
    
    auxmanufacturer = ""
    
    
    For m = Aux + 1 To N
        
        Application.StatusBar = "Updating Supplier's Contact Information: " & m - Aux & " of " & N - Aux & ": " & Format((m - Aux) / (N - Aux), "0%")
        
        manufacturer = ws_OG.Cells(m, manufj).Value
        
        If auxmanufacturer <> manufacturer Then     'Con este condicional �nicamente busca la informaci�n de contacto una vez.
            
            auxmanufacturer = manufacturer
            Set c = Range(ws_contact.Cells(CPsupplieri + 1, CPsupplierj), ws_contact.Cells(CPendi, CPsupplierj)).Find(manufacturer)
        
        End If
        
        If c Is Nothing Then                        'No hay informaci�n de contacto.
        
            ws_OG.Cells(m, ContactDBj) = "Does NOT Exist"
            ws_OG.Cells(m, ContactDBj).Interior.ColorIndex = 3
            
        Else
        
            If ws_contact.Cells(c.Row, CPmailj) = "" Then     'Existe el proveedor en la lista pero no hay informaci�n de contacto.
            
                ws_OG.Cells(m, ContactDBj) = "Does NOT Exist"
                ws_OG.Cells(m, ContactDBj).Interior.ColorIndex = 3
                
            Else    'Existe informaci�n de contacto.
            
                ws_OG.Cells(m, ContactDBj) = ws_contact.Cells(c.Row, CPmailj)
                ws_OG.Cells(m, ContactDBj).Interior.ColorIndex = 43
                
            End If
            
        End If
        
    Next
    
    Application.StatusBar = ""
    
End Function

Function Identify_Status(i As Integer, DateT1j As Integer, Current_Date As Date, status0 As Integer, statusmin As Integer)
'Llama a las funciones en orden para identificar correctamente el estado de los certificados.

    Dim ColumnPosition As Integer

    ColumnPosition = DateT1j
    status1 = Check_Dates(i, ColumnPosition, Current_Date, status0)
    
    If status0 <> 23 And ws_OG.Cells(i, ManufDeclarationj) <> "" And IsDate(ws_OG.Cells(i, ManufDeclarationj)) Then
        
        ColumnPosition = ManufDeclarationj
        status1 = Check_Dates(i, ColumnPosition, Current_Date, status0)
        
    End If
        
    ColumnPosition = TMexpirej
    Call Log_Status(i, ColumnPosition, status0, status1)
    
    If status0 < statusmin Then
        
        statusmin = status0        'Registra el nuevo estado m�nimo de la l�nea.
        ColumnPosition = GlobalStatusj
        Call Log_Status(i, ColumnPosition, status0, status1)
    
    End If
    
End Function

Function Check_Dates(i As Integer, ColumnPosition As Integer, Current_Date As Date, status0 As Integer) As String
'Compara las fechas de los certificados y de la declaraci�n de conformidad (si existiera) y registra el estado de los ensayos.
       
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

Function Log_Status(i, ColumnPosition, status0, status1)
'Registra el estado de cada certificado (texto + c�digo de color).
    
    ws_OG.Cells(i, ColumnPosition).Value = status1

    Set findstatus = Range(ws_ranking.Cells(RSRankingi, RSStatusENj), ws_ranking.Cells(RSEndi, RSStatusENj)).Find(status1)
    
    ws_OG.Cells(i, ColumnPosition).Interior.ColorIndex = ws_ranking.Cells(findstatus.Row, RSColorCodej).Value
    
End Function

Function Counters_Check()
'Aumenta o resetea el contador.
    If TMexpirej = ws_OG.Range("A10:DA10").Find("Test Method 1 time to expire*").Column + 5 Then

        TMexpirej = ws_OG.Range("A10:DA10").Find("Test Method 1 time to expire*").Column
        
    Else
    
        TMexpirej = TMexpirej + 1
        
    End If
            
End Function
                                                                       ����                                                                                                                                                                                                                                                                                                                                    ����                                                                                                             