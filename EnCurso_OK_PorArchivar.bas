Attribute VB_Name = "Módulo1"
Sub limpieza_bbdd()
    
    Dim inicioi As Integer
    Dim inicioj As Integer
    Dim finali As Integer
    Dim finalj As Integer
    Dim supplierj As Integer
    Dim estadoj As Integer
    Dim fechaj As Integer
    Dim i As Integer
    Dim estado As String
    Dim fechaActual As Date
    Dim Dif_Dia As Integer
    Dim auxfinali As Integer
    
    inicioi = Sheets("EN CURSO").Range("A1:A10").Find("PART NUMBER").Row            'Posiciones
    inicioj = Sheets("EN CURSO").Range("A1:Z1").Find("PART NUMBER").Column
    
    finali = Sheets("EN CURSO").Cells(Rows.Count, inicioj).End(xlUp).Row
    finalj = Sheets("EN CURSO").Cells(inicioi, Columns.Count).End(xlToLeft).Column
    
    supplierj = Sheets("EN CURSO").Range("A1:Z1").Find("SUPPLIER").Column
    
    estadoj = Sheets("EN CURSO").Range(Cells(inicioi, inicioj), Cells(inicioi, finalj)).Find("ESTADO").Column
    fechaj = Sheets("EN CURSO").Range(Cells(inicioi, inicioj), Cells(inicioi, finalj)).Find("FECHA DE ÚLTIMO CORREO ENVIADO").Column
    
    For i = inicioi + 1 To finali
    
        estado = Sheets("EN CURSO").Cells(i, estadoj).Value
        
        If IsDate(Sheets("EN CURSO").Cells(i, fechaj)) = True Then                   'Error: en la celda no hay una fecha
        
        fechaActual = Date
        Dif_Dia = DateDiff("d", Sheets("EN CURSO").Cells(i, fechaj), fechaActual)
        
            If estado = "OK" And Dif_Dia >= 7 Then           'Cortar y pegar si cumple. En OK.
                
                Sheets("EN CURSO").Range(Cells(i, inicioj), Cells(i, finalj)).Cut
                auxfinali = Sheets("OK").Cells(Rows.Count, inicioj).End(xlUp).Row + 1
                Sheets("OK").Activate
                Cells(auxfinali, "A").Select
                ActiveSheet.Paste
                
                Sheets("EN CURSO").Cells(i, inicioj).EntireRow.Delete
                i = i - 1
                
                Sheets("EN CURSO").Activate
                
            End If
            
            If estado = "POR ARCHIVAR" And Sheets("EN CURSO").Cells(i, finalj + 2) <> 1 Then                       'Cortar y pegar si cumple. En POR ARCHIVAR.
            
                Sheets("EN CURSO").Range(Cells(i, inicioj), Cells(i, supplierj)).Copy
                auxfinali = Sheets("POR ARCHIVAR").Cells(Rows.Count, inicioj).End(xlUp).Row + 1
                Sheets("POR ARCHIVAR").Activate
                Range("A" & auxfinali).Select
                ActiveSheet.Paste
                
                Sheets("OK").Activate
                Sheets("OK").Range("Z1").Copy Sheets("POR ARCHIVAR").Range("F" & auxfinali)           'Lista de validación
                
                Sheets("EN CURSO").Activate
                
                Sheets("EN CURSO").Cells(i, finalj + 2).Value = 1
            
            End If
    
        End If
    
    Next
    
End Sub


