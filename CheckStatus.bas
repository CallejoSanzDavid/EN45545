Attribute VB_Name = "Módulo1"
Sub Comprobar_Caducidad()
    
    Dim i As Integer
    Dim j As Integer
    Dim nj As Integer
    Dim x As Integer
    Dim N As Integer
    Dim k As Integer
    Dim G As Integer
    Dim fechaActual As Date
    Dim Dif_Mes As Integer
    Dim Dif_Dia As Integer
    Dim Dif_MesDC As Integer
    Dim Dif_DiaDC As Integer
    Dim status(6, 1) As String
    Dim Aux As Integer
    Dim statusmin As Integer
    Dim DeclaracionConformidad As Integer
    Dim DeclaracionConformidadj As Integer
    
    nj = Sheets("FCIL").Range("A10:DA10").Find("Date * T6").Column
    x = Sheets("FCIL").Range("A10:DA10").Find("Test Method 1 time to expire*").Column
    N = Contar_Elem
    k = 1
    G = Sheets("FCIL").Range("A10:DA10").Find("Certificate global status*").Column
    fechaActual = Date
    Aux = Sheets("FCIL").Range("A10:DA10").Find("Assembly Name").Row                'Fila inicial
    DeclaracionConformidadj = Sheets("FCIL").Range("A10:DA10").Find("Manufacturer Declaration Date").Column
    
    Call BaseProveedores
    
    For i = Sheets("FCIL").Range("A10:DA10").Find("Assembly Name").Row + 1 To N
        
        statusmin = 24                                                             'Valores de cadena auxiliares para evitar errores al comparar
        Application.StatusBar = "Checking Certificates Status: " & i - Aux & " of " & N - Aux & ": " & Format((i - Aux) / (N - Aux), "0%")
        
        For j = Sheets("FCIL").Range("A10:DA10").Find("Date * T1").Column To nj Step 6
        
            Do While IsDate(Sheets("FCIL").Cells(i, j)) = False And j <= nj                   'Error: en la celda no hay una fecha
                
                Sheets("FCIL").Cells(i, x).Value = "No date"
                
                status(k, 0) = 23
                status(k, 1) = "No date"
                
                If status(k, 0) < statusmin Then
                    
                    Sheets("FCIL").Cells(i, G).Value = status(k, 1)
                    statusmin = status(k, 0)
                
                End If
                
                If k = 6 Then
                    
                    k = 1
                    
                Else
                    
                    k = k + 1
                
                End If
                
                    x = x + 1
                
                If j = nj Or x = Sheets("FCIL").Range("A10:DA10").Find("Test Method 1 time to expire*").Column + 6 Then
                    
                    x = Sheets("FCIL").Range("A10:DA10").Find("Test Method 1 time to expire*").Column
                    Exit For
                
                End If
                
                j = j + 6
                
            Loop
                            
            Dif_Mes = 60 - DateDiff("m", Sheets("FCIL").Cells(i, j), fechaActual)
            Dif_Dia = 1827 - DateDiff("d", Sheets("FCIL").Cells(i, j), fechaActual)
            
            If Sheets("FCIL").Cells(i, DeclaracionConformidadj) <> "" And IsDate(Sheets("FCIL").Cells(i, DeclaracionConformidadj)) Then
                Dif_MesDC = 60 - DateDiff("m", Sheets("FCIL").Cells(i, DeclaracionConformidadj), fechaActual)
                Dif_DiaDC = 1827 - DateDiff("d", Sheets("FCIL").Cells(i, DeclaracionConformidadj), fechaActual)
            Else
                Dif_MesDC = 0
                Dif_DiaDC = 0
            End If
                  
            If Dif_Mes > 6 Or Dif_MesDC > 6 Then                    'Si faltan más de 6 meses para que caduque: OK
            
                Sheets("FCIL").Cells(i, x) = "OK"
                Sheets("FCIL").Cells(i, x).Interior.ColorIndex = 4  'Verde si es OK
                
                status(k, 0) = 22
                status(k, 1) = "OK"
                
            End If
                  
            If Dif_Mes <= 6 And Dif_MesDC <= 6 Then                     'Si faltan menos de 6 meses para que caduque
            
                Sheets("FCIL").Cells(i, x).Value = Dif_Mes & " month/s"
                Sheets("FCIL").Cells(i, x).Interior.ColorIndex = 6      'Amarillo si falta entre 6 y 3 meses
                status(k, 0) = 15 + Dif_Mes
                status(k, 1) = Dif_Mes & " month/s"
                
                If Dif_Mes <= 3 And Dif_MesDC <= 3 Then
                
                    Sheets("FCIL").Cells(i, x).Interior.ColorIndex = 44 'Amarillo oscuro si está entre 3 y 2 meses
                    
                End If
                
                If Dif_Mes <= 2 And Dif_MesDC <= 2 Then
                
                    Sheets("FCIL").Cells(i, x).Interior.ColorIndex = 45 'Naranja claro si está entre 2 y 1 mes/es.
                    
                    
                End If
                
                If Dif_Mes <= 1 And Dif_MesDC <= 1 And Dif_Dia <= 30 And Dif_DiaDC <= 30 Then   'Si faltan días para que caduque
                    
                    Sheets("FCIL").Cells(i, x).Value = Dif_Dia & " day/s"

                    status(k, 0) = Dif_Dia
                    status(k, 1) = Dif_Dia & " day/s"
                    
                    If Dif_Dia <= 30 And Dif_DiaDC <= 30 Then
                    
                        Sheets("FCIL").Cells(i, x).Interior.ColorIndex = 46 'Naranja oscuro faltan entre 30 y 1 días
                        
                        If Dif_Dia <= 15 And Dif_DiaDC <= 15 Then
                            
                            status(k, 0) = Dif_Dia
                        
                        End If
                        
                    End If
                    
                    If Dif_Dia <= 0 And Dif_DiaDC <= 0 Then
                    
                        Sheets("FCIL").Cells(i, x).Value = "EXPIRED"
                        Sheets("FCIL").Cells(i, x).Interior.ColorIndex = 3  'Rojo si está caducado
                        status(k, 0) = 0
                        status(k, 1) = "EXPIRED"
                        
                    End If
                    
                End If
                
            End If
            
            If status(k, 0) < statusmin Then
                
                Sheets("FCIL").Cells(i, G).Value = status(k, 1)
                
                If status(k, 0) <= 21 And status(k, 0) >= 19 Then
                
                    status(k, 0) = 21
                    
                End If
                
                statusmin = status(k, 0)
                
                If status(k, 0) = 22 Then       'Si faltan más de 6 meses para que caduque: OK
                    
                    Sheets("FCIL").Cells(i, G).Interior.ColorIndex = 4  'Verde si es OK
                    
                End If
                
                If status(k, 0) <= 21 Then      'Si faltan menos de 6 meses para que caduque
                    
                    Sheets("FCIL").Cells(i, G).Interior.ColorIndex = 6  'Amarillo si falta entre 6 y 3 meses
                    
                    If status(k, 0) <= 18 Then
                        
                        Sheets("FCIL").Cells(i, G).Interior.ColorIndex = 44 'Amarillo oscuro si está entre 3 y 2 meses
                        
                        If status(k, 0) <= 17 Then
                            
                            Sheets("FCIL").Cells(i, G).Interior.ColorIndex = 45 'Naranja claro si está entre 2 y 1 mes/es.
    
                            If status(k, 0) <= 15 Then
                            
                                Sheets("FCIL").Cells(i, G).Interior.ColorIndex = 46 'Naranja oscuro faltan entre 30 y 1 días
                            
                            End If
                            
                            If status(k, 0) <= 0 Then
                            
                                Sheets("FCIL").Cells(i, G).Interior.ColorIndex = 3  'Rojo si está caducado
                                
                            End If
                        
                        End If
                        
                    End If
                        
                End If
                
            End If
            
            If k = 6 Then
                k = 1
                status(0, 0) = 24                                                             'Valores de cadena auxiliares para evitar errores al comparar
            Else
                k = k + 1
            End If
            
            If x = Sheets("FCIL").Range("A10:DA10").Find("Test Method 1 time to expire*").Column + 5 Then
                x = Sheets("FCIL").Range("A10:DA10").Find("Test Method 1 time to expire*").Column
            Else
                x = x + 1
            End If
        
        Next
        
    Next
    
    Application.StatusBar = ""
    
End Sub


Function Contar_Elem() As Integer

    Dim ContadorN As Integer
    Dim ContadorP As Integer
    Dim nprodj As Integer
    Dim nombj As Integer
    
    nprodj = Sheets("FCIL").Range("A10:DA10").Find("Supplier part number").Column
    nombj = Sheets("FCIL").Range("A10:DA10").Find("Part name").Column
        
    ContadorN = Sheets("FCIL").Cells(Rows.Count, nprodj).End(xlUp).Row
    ContadorP = Sheets("FCIL").Cells(Rows.Count, nombj).End(xlUp).Row
    

    If (ContadorN = ContadorP) Then
        Contar_Elem = ContadorN
    End If
    
    If (ContadorN > ContadorP) Then
        MsgBox ("Hay campos sin rellenar en la columna P.")
        Contar_Elem = ContadorN
    End If
    
    If (ContadorN < ContadorP) Then
        MsgBox ("Hay campos sin rellenar en la columna N.")
        Contar_Elem = ContadorP
    End If

End Function

Function BaseProveedores()

    Dim manufj As Integer
    Dim manufacturer As String
    Dim m As Integer
    Dim N As Integer
    Dim ContarDB As Integer
    Dim ContactDB As Integer
    Dim ContactoDBi As Integer
    Dim supplieri As Integer
    Dim supplierj As Integer
    Dim mailj As Integer
    Dim linea As Integer
    Dim c As Range
    
    manufj = Sheets("FCIL").Range("A10:DA10").Find("Manufacturer name*").Column
    
    ContactDB = Sheets("FCIL").Range("A10:DA10").Find("Supplier's Contact").Column
    ContactoDBi = Sheets("FCIL").Range("A10:DA10").Find("Supplier's Contact").Row + 1
    supplieri = Sheets("Contacto de proveedores").Range("A1:Z1").Find("Supplier").Row + 1
    supplierj = Sheets("Contacto de proveedores").Range("A1:Z1").Find("Supplier").Column
    mailj = Sheets("Contacto de proveedores").Range("A1:Z1").Find("Mail").Column
    
    
    ContarDB = Contar_Elem_DB
    N = Contar_Elem


    For m = ContactoDBi To N
        
        manufacturer = Sheets("FCIL").Cells(m, manufj).Value
        
        Application.StatusBar = "Updating Supplier's Contact Information: " & m - ContactoDBi + 1 & " of " & N - ContactoDBi + 1 & ": " & Format((m - ContactoDBi + 1) / (N - ContactoDBi + 1), "0%")
        
        Set c = Range(Sheets("Contacto de proveedores").Cells(supplieri, supplierj), Sheets("Contacto de proveedores").Cells(ContarDB, supplierj)).Find(manufacturer)
        
        If c Is Nothing Then
                
            Sheets("FCIL").Cells(m, ContactDB) = "Does NOT Exist"
            Sheets("FCIL").Cells(m, ContactDB).Interior.ColorIndex = 3
              
        Else
        
            linea = c.Row
                     
            Sheets("FCIL").Cells(m, ContactDB) = Sheets("Contacto de proveedores").Cells(linea, mailj)
            Sheets("FCIL").Cells(m, ContactDB).Interior.ColorIndex = 43
        
        End If
        
    Next
    
    Application.StatusBar = ""
    
End Function

Function Contar_Elem_DB() As Integer

    Dim j As Integer
    
    j = Sheets("Contacto de proveedores").Range("A1:Z1").Find("Supplier").Column
    
    Contar_Elem_DB = Sheets("Contacto de proveedores").Cells(Rows.Count, j).End(xlUp).Row

End Function

