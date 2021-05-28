Attribute VB_Name = "Módulo3"
Sub Nuevos_Pedidos()        'Archiva los nuevos pedidos que están PEDIDOS en "EN CURSO"
    
    Dim inicioi As Integer
    Dim inicioj As Integer
    Dim finali As Integer
    Dim finalj As Integer
    Dim estadoj As Integer
    Dim i As Integer
    Dim estado As String
    Dim auxfinali As Integer
    
    inicioi = Sheets("TEMP").Range("A1:A10").Find("PART NUMBER").Row            'Posiciones
    inicioj = Sheets("TEMP").Range("A1:A10").Find("PART NUMBER").Column
    
    finali = Sheets("TEMP").Cells(Rows.Count, inicioj).End(xlUp).Row
    finalj = Sheets("TEMP").Cells(inicioi, Columns.Count).End(xlToLeft).Column
    
    estadoj = Sheets("TEMP").Range(Cells(inicioi, inicioj), Cells(inicioi, finalj)).Find("ESTADO").Column
    
    For i = inicioi + 1 To finali
    
        estado = Sheets("TEMP").Cells(i, estadoj).Value
        
        If estado = "PEDIDO" Then          'Cortar y pegar si cumple
                
            Sheets("TEMP").Range(Cells(i, inicioj), Cells(i, finalj)).Cut
            
            Sheets("EN CURSO").Activate
            auxfinali = Sheets("EN CURSO").Cells(Rows.Count, estadoj).End(xlUp).Row + 1
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            ActiveSheet.ListObjects(1).Resize Range(Cells(inicioi, inicioj), Cells(auxfinali, finalj)) 'Ampliamos el rango de la tabla para que añada la nueva línea
            
            Sheets("TEMP").Activate
            
            Sheets("TEMP").Cells(i, inicioj).EntireRow.Delete
            i = i - 1
            
        End If
    
    Next
    
End Sub



