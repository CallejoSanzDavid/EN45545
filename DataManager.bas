Attribute VB_Name = "DataManager"
Sub Data_Manager()
'Cuerpo principal de la herramienta. Selecciona las funciones dependiendo de la hoja en la que se ejecutan.
    
    Dim SheetName As String
    Dim Status As String
    Dim i As Integer
    Dim Today As Date
    Dim Day_Dif As Integer
    Dim PA_Status As String
                
    Application.ScreenUpdating = False
    
    SheetName = ActiveSheet.Name
    
    If SheetName <> "POR ARCHIVAR" Then
    
        Call Locate_Positions_OG(SheetName)
    
    Else
    
        Call Locate_Positions_PA(SheetName)
    
    End If
    
    Select Case SheetName
                
        Case "EN CURSO"     '---CORRECTO---
            
            For i = Starti + 1 To Endi
                
                Status = Sheets(SheetName).Cells(i, Statusj).Value
                
                Select Case Status
                
                    Case "OK"
                        
                        If IsDate(Sheets(SheetName).Cells(i, LastMsgj)) = True Then                   'Error: en la celda no hay una fecha
                
                            Today = Date
                            Day_Dif = DateDiff("d", Sheets(SheetName).Cells(i, LastMsgj), Today)
                            
                            If Day_Dif >= 7 Then           'Cortar y pegar si cumple.
                        
                                Call Cut_Paste(SheetName, Status, i, AdActj)
                                
                            End If
                            
                        End If
                    
                    Case "POR ARCHIVAR"
                        
                        If Sheets(SheetName).Cells(i, AdActj + 2).Value <> 1 Then                    'Copiar y pegar si cumple en POR ARCHIVAR.
                            
                            Call Locate_Positions_PA(SheetName)
                            Call PA_Copy_Paste(SheetName, Status, i)
                        
                        End If
                    
                    Case "NO EN45545"
                        
                        Call Cut_Paste(SheetName, Status, i, AdActj)
                    
                    Case "NOK"
                    
                        Call Delete_Row(i)
                        
                    Case ""
                    
                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                        
                End Select
                        
            Next

        Case "POR ARCHIVAR"                 '---CORRECTO---
            
            For i = PA_Starti + 1 To PA_Endi
    
                PA_Status = Sheets(SheetName).Cells(i, PA_Statusj).Value
                                
                Select Case PA_Status
                
                    Case "OK"

                        Call Update_Status(i, "EN CURSO", SheetName)
                        Call Cut_Paste(SheetName, "ARCHIVADOS", i, PA_Statusj)
                    
                    Case "NOK"
                    
                        Call Delete_Row(i)
                    
                    Case ""

                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                
                End Select
                
            Next

        Case "OK", "NO EN45545"         '---CORRECTO---
            
            For i = Starti + 1 To Endi
                
                Status = Sheets(SheetName).Cells(i, Statusj).Value
                
                Select Case Status
                
                    Case ""
                        
                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                    
                    Case "NOK"
                
                        Call Delete_Row(i)

                    Case SheetName
                    'Si coincide el estado con el nombre de la hoja no se hace nada
                    
                    Case "OK", "NO EN45545"
                    'Si estamos en la hoja OK y el estado es NO EN45545 se mueve directamente a la hoja NO EN45545.
                    'Si estamos en la hoja NO EN45545 y el estado es OK se mueve directamente a la hoja OK.
                        Call Cut_Paste(SheetName, Status, i, AdActj)
                    
                    Case Else
                        
                        Call Cut_Paste(SheetName, "EN CURSO", i, AdActj)
                    
                End Select

            Next
'<------------------STOP
        Case "TEMP"
            
            For i = Starti + 1 To Endi
                
                Status = Sheets(SheetName).Cells(i, Statusj).Value
                
                Select Case Status
                    
                    Case ""
                        
                        Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
                    
                    Case "---"
                    
                    Case "NOK"
                
                        Call Delete_Row(i)
                    
                    Case "OK", "NO EN45545"
                        
                        Call Cut_Paste(SheetName, Status, i, AdActj)
                    
                    Case Else
                    
                        Call Cut_Paste(SheetName, "EN CURSO", i, AdActj)
                    
                End Select
                
            Next

    End Select
'<------------------STOP
    Application.ScreenUpdating = True
                        
End Sub
                       
Function Cut_Paste(SheetName As String, Status As String, i As Integer, Lastj As Integer)        '---CORRECTO---
'Corta y pega cualquier línea de información
    
    Dim AuxEndi As Integer
    
    Sheets(SheetName).Range(Cells(i, PartNumj), Cells(i, Lastj)).Cut
    
    Sheets(Status).Activate
    AuxEndi = Sheets(Status).Cells(Rows.Count, PartNumj).End(xlUp).Row + 1
    Cells(AuxEndi, "A").Select
    ActiveSheet.Paste
    
    'Ampliamos el rango de la tabla para que añada la nueva línea
    ActiveSheet.ListObjects(1).Resize Range(Cells(Starti, PartNumj), Cells(AuxEndi, Lastj))
    
    Sheets(SheetName).Activate
    
    Call Delete_Row(i)
    
End Function

Function Delete_Row(i As Integer)           '---CORRECTO---

    ActiveSheet.Cells(i, 1).EntireRow.Delete
    i = i - 1

End Function

Function PA_Copy_Paste(SheetName As String, Status As String, i As Integer)
'Copia y pega información a la hoja "POR ARCHIVAR".
'Esta hoja no tiene tantos campos como el resto de hojas, por eso requiere una función especial
    
    Dim AuxEndi As Integer
    
    Sheets(SheetName).Range(Cells(i, PartNumj), Cells(i, Supplierj)).Copy
    
    Sheets(Status).Activate
    
    AuxEndi = Sheets(Status).Cells(Rows.Count, PartNumj).End(xlUp).Row + 1
    Range("A" & AuxEndi).Select
    ActiveSheet.Paste
    
    'Ampliamos el rango de la tabla para que añada la nueva línea.
    ActiveSheet.ListObjects(1).Resize Range(Cells(PA_Starti, PA_PartNumj), Cells(AuxEndi, PA_Statusj))
    
    Sheets("AUX2").Range("C1").Copy Sheets("POR ARCHIVAR").Range(PA_StatusLetterj & AuxEndi)            'Lista de validación: "PENDIENTE".
    
    Sheets(SheetName).Activate
    
    'Marca que indica que la línea ya se ha movido a la hoja "POR ARCHIVAR" anteriormente.
    Sheets(SheetName).Cells(i, AdActj + 2).Value = 1

End Function

Function Update_Status(i As Integer, Update_Sheet As String, SheetName As String)       '---CORRECTO---
    
    Dim PartNum As String

    PA_PartNum = ws_PA.Cells(i, PA_PartNumj)
    
    Sheets(Update_Sheet).Activate
    
    Call Locate_Positions_OG(Update_Sheet)
    
    Set c = Range(Sheets(Update_Sheet).Cells(Starti, PartNumj), Sheets(Update_Sheet).Cells(Endi, PartNumj)).Find(PA_PartNum)

    If Not c Is Nothing Then
        
        Sheets("AUX2").Range("B1").Copy Sheets(Update_Sheet).Cells(c.Row, Statusj)      'Lista de validación "OK"
                                                                                        
    End If
    
    Sheets(SheetName).Activate
    
End Function

                          
        

    
