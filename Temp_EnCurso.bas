Attribute VB_Name = "Módulo3"
Sub Nuevos_Pedidos()
    
    Dim inicioi As Integer
    Dim inicioj As Integer
    Dim finali As Integer
    Dim finalj As Integer
    Dim estadoj As Integer
    Dim fechaj As Integer
    Dim i As Integer
    Dim estado As String
    Dim fechaActual As Date
    Dim Dif_Dia As Integer
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
            auxfinali = Sheets("EN CURSO").Cells(Rows.Count, estadoj).End(xlUp).Row + 1
            Sheets("EN CURSO").Activate
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            Sheets("TEMP").Cells(i, inicioj).EntireRow.Delete
            i = i - 1
            
            Sheets("TEMP").Activate
            
        End If
    
    Next
    
End Sub
