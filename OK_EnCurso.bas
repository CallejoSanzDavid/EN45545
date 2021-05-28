Attribute VB_Name = "Módulo4"
Sub DESHACER()  'Devuelve lineas que estaban en "OK" a "EN CURSO".

    Dim inicioi As Integer
    Dim inicioj As Integer
    Dim finali As Integer
    Dim finalj As Integer
    Dim estadoj As Integer
    Dim fechaj As Integer
    Dim i As Integer
    Dim estado As String
    Dim auxfinali As Integer
    
    Application.ScreenUpdating = False
    
    inicioi = Sheets("OK").Range("A1:A10").Find("PART NUMBER").Row            'Posiciones
    inicioj = Sheets("OK").Range("A1:A10").Find("PART NUMBER").Column
    
    finali = Sheets("OK").Cells(Rows.Count, inicioj).End(xlUp).Row
    finalj = Sheets("OK").Cells(inicioi, Columns.Count).End(xlToLeft).Column
    
    estadoj = Sheets("OK").Range(Cells(inicioi, inicioj), Cells(inicioi, finalj)).Find("ESTADO").Column
    
    For i = inicioi + 1 To finali
    
        estado = Sheets("OK").Cells(i, estadoj).Value
        
        If estado = "" Then
            
            Exit For            'Con esto evitamos que se quede atascado en el bucle añadiendo líneas vacías
            
        End If
        
        If estado <> "OK" Then          'Cortar y pegar si cumple
                
            Sheets("OK").Range(Cells(i, inicioj), Cells(i, finalj)).Cut
            
            Sheets("EN CURSO").Activate
            
            auxfinali = Sheets("EN CURSO").Cells(Rows.Count, estadoj).End(xlUp).Row + 1
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            ActiveSheet.ListObjects(1).Resize Range(Cells(inicioi, inicioj), Cells(auxfinali, finalj)) 'Ampliamos el rango de la tabla para que añada la nueva línea
            
            Sheets("OK").Activate
            
            Sheets("OK").Cells(i, inicioj).EntireRow.Delete
            i = i - 1
            
        End If
        
    Next

    Application.ScreenUpdating = True
    
End Sub

