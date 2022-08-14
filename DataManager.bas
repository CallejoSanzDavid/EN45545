Attribute VB_Name = "DataManager"
Sub Data_Manager()
'Cuerpo principal de la herramienta. Selecciona las funciones dependiendo de la hoja en la que se ejecutan.
    
    Dim SheetName As String
    Dim Status As String
    Dim i As Integer
    Dim Today As Date
    Dim Day_Dif As Integer
    Dim PA_Status As String
                
    Application.StatusBar = ""
    Application.ScreenUpdating = False
    
    SheetName = ActiveSheet.Name
    
    If SheetName <> "POR ARCHIVAR" Then
    
        Call Locate_Positions_OG(SheetName)
    
    Else
    
        Call Locate_Positions_PA(SheetName)
    
    End If
    
    Select Case SheetName
                
        Case "EN CURSO"
            
            For i = Starti + 1 To Endi
                
                Status = Sheets(SheetName).Cells(i, Statusj).Value
                
                Select Case Status
                
                    Case ""
                    
                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                    
                    Case "NOK"
                    
                        Call Delete_Row(i)
                                        
                    Case "OK"
                        
                        If IsDate(Sheets(SheetName).Cells(i, LastMsgj)) = True Then                   'Error: en la celda no hay una fecha
                
                            Today = Date
                            Day_Dif = DateDiff("d", Sheets(SheetName).Cells(i, LastMsgj), Today)
                            
                            If Day_Dif >= 7 Then           'Cortar y pegar si cumple.
                        
                                Call Move_Line(SheetName, Status, i, AdActj)
                                
                            End If
                            
                        End If
                    
                    Case "POR ARCHIVAR"
                        
                        If Sheets(SheetName).Cells(i, AdActj + 2).Value <> 1 Then                    'Copiar y pegar si cumple en POR ARCHIVAR.
                            
                            Call Locate_Positions_PA(SheetName)
                            Call Move_Line(SheetName, Status, i, Supplierj)
                        
                        End If
                    
                    Case "NO EN45545"
                        
                        Call Move_Line(SheetName, Status, i, AdActj)
                        
                End Select
                        
            Next

        Case "POR ARCHIVAR"
            
            For i = PA_Starti + 1 To PA_Endi
    
                PA_Status = Sheets(SheetName).Cells(i, PA_Statusj).Value
                                
                Select Case PA_Status
                
                    Case ""

                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                    
                    Case "NOK"
                    
                        Call Delete_Row(i)
                    
                    Case "OK"

                        Call Update_Status(i, SheetName)
                        Call Move_Line(SheetName, "ARCHIVADOS", i, PA_Statusj)
                                        
                End Select
                
            Next

        Case "OK", "NO EN45545", "TEMP"
            
            For i = Starti + 1 To Endi
                
                Status = Sheets(SheetName).Cells(i, Statusj).Value
                
                Select Case Status
                
                    Case ""
                        
                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                    
                    Case "NOK"
                
                        Call Delete_Row(i)

                    Case SheetName, "---"
                    'Si coincide el estado con el nombre de la hoja no se hace nada
                    
                    Case "OK", "NO EN45545"
                    'Si estamos en la hoja OK y el estado es NO EN45545 se mueve directamente a la hoja NO EN45545.
                    'Si estamos en la hoja NO EN45545 y el estado es OK se mueve directamente a la hoja OK.
                        Call Move_Line(SheetName, Status, i, AdActj)
                    
                    Case Else
                        
                        Call Move_Line(SheetName, "EN CURSO", i, AdActj)
                    
                End Select

            Next

    End Select

    Application.ScreenUpdating = True
                        
End Sub
                       
Function Move_Line(SheetName As String, Status As String, i As Integer, Lastj As Integer)
'Mueve de una hoja a otra cualquier línea de información.
    
    Dim AuxEndi As Integer
    
    If Status = "POR ARCHIVAR" Then
        Sheets(SheetName).Range(Cells(i, PartNumj), Cells(i, Lastj)).Copy
    Else
        Sheets(SheetName).Range(Cells(i, PartNumj), Cells(i, Lastj)).Cut
    End If
    
    Sheets(Status).Activate
    AuxEndi = Sheets(Status).Cells(Rows.Count, PartNumj).End(xlUp).Row + 1
    Cells(AuxEndi, "A").Select
    ActiveSheet.Paste
    
    'Ampliamos el rango de la tabla para que añada la nueva línea.
    If Status = "POR ARCHIVAR" Then
        ActiveSheet.ListObjects(1).Resize Range(Cells(PA_Starti, PA_PartNumj), Cells(AuxEndi, PA_Statusj))
        Sheets("AUX2").Range("C1").Copy Destination:=Sheets("POR ARCHIVAR").Range(PA_StatusLetterj & AuxEndi)    'Lista de validación: "PENDIENTE".
    Else
        ActiveSheet.ListObjects(1).Resize Range(Cells(Starti, PartNumj), Cells(AuxEndi, Lastj))
    End If
    
    Sheets(SheetName).Activate
    
    If Status = "POR ARCHIVAR" Then
        'Marca que indica que la línea ya se ha movido a la hoja "POR ARCHIVAR" anteriormente.
        Sheets(SheetName).Cells(i, AdActj + 2).Value = 1
    Else
        Call Delete_Row(i)
    End If
    
End Function

Function Delete_Row(i As Integer)

    ActiveSheet.Cells(i, 1).EntireRow.Delete
    i = i - 1

End Function

Function Update_Status(i As Integer, SheetName As String)
    
    Dim PartNum As String

    PA_PartNum = ws_PA.Cells(i, PA_PartNumj)
    
    Sheets("EN CURSO").Activate
    
    Call Locate_Positions_OG("EN CURSO")
    
    Set c = Range(Sheets("EN CURSO").Cells(Starti, PartNumj), Sheets("EN CURSO").Cells(Endi, PartNumj)).Find(PA_PartNum)

    If Not c Is Nothing Then
        
        Sheets("AUX2").Range("B1").Copy Sheets("EN CURSO").Cells(c.Row, Statusj)      'Lista de validación "OK"
                                                                                        
    End If
    
    Sheets(SheetName).Activate
    
End Function

                          
        

    
