Attribute VB_Name = "Módulo2"
Sub ARCHIVADOS()            'Archiva las lineas que están OK de "POR ARCHIVAR" a "ARCHIVADOS". Actualiza el estado de la linea archivada en "EN CURSO" -> OK
    
    Dim inicioi As Integer
    Dim inicioj As Integer
    Dim finali As Integer
    Dim finalj As Integer
    Dim estadoj As Integer
    Dim inicioECi As Integer
    Dim inicioECj As Integer
    Dim finalECi As Integer
    Dim finalECj As Integer
    Dim fechaj As Integer
    Dim i As Integer
    Dim estado As String
    Dim fechaActual As Date
    Dim Dif_Dia As Integer
    Dim auxfinali As Integer
    Dim auxpartnumberi As Integer
    Dim partnumber As String
    
    inicioi = Sheets("POR ARCHIVAR").Range("A1:A10").Find("PART NUMBER").Row            'Posiciones iniciales "POR ARCHIVAR"
    inicioj = Sheets("POR ARCHIVAR").Range("A1:Z1").Find("PART NUMBER").Column
    
    finali = Sheets("POR ARCHIVAR").Cells(Rows.Count, inicioj).End(xlUp).Row            'Posiciones finales "POR ARCHIVAR"
    finalj = Sheets("POR ARCHIVAR").Cells(inicioi, Columns.Count).End(xlToLeft).Column
    
    estadoj = Sheets("POR ARCHIVAR").Range(Cells(inicioi, inicioj), Cells(inicioi, finalj)).Find("ESTADO").Column
    
    Sheets("EN CURSO").Activate
    
    inicioECi = Sheets("EN CURSO").Range("A1:A10").Find("PART NUMBER").Row              'Posiciones iniciales "EN CURSO"
    inicioECj = Sheets("EN CURSO").Range("A1:Z1").Find("PART NUMBER").Column
    
    finalECi = Sheets("EN CURSO").Cells(Rows.Count, inicioj).End(xlUp).Row              'Posiciones finales "EN CURSO"
    finalECj = Sheets("EN CURSO").Cells(inicioi, Columns.Count).End(xlToLeft).Column
    
    For i = inicioi + 1 To finali
    
        estado = Sheets("POR ARCHIVAR").Cells(i, estadoj).Value
        partnumber = Sheets("POR ARCHIVAR").Cells(i, inicioj).Value
        
        If estado = "OK" Then          'Cortar y pegar si cumple a "ARCHIVADO"
                     
            Sheets("EN CURSO").Activate
            
            auxpartnumberi = Sheets("EN CURSO").Range(Cells(inicioECi, inicioECj), Cells(finalECi, inicioECj)).Find(partnumber).Row
            Sheets("AUX2").Range("B1").Copy Sheets("EN CURSO").Range("J" & auxpartnumberi)           'Lista de validación "OK"
            
            Sheets("POR ARCHIVAR").Activate
            
            Sheets("POR ARCHIVAR").Range(Cells(i, inicioj), Cells(i, finalj)).Cut
            auxfinali = Sheets("ARCHIVADOS").Cells(Rows.Count, estadoj).End(xlUp).Row + 1
            Sheets("ARCHIVADOS").Activate
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            ActiveSheet.ListObjects(1).Resize Range(Cells(inicioi, inicioj), Cells(auxfinali, finalj)) 'Ampliamos el rango de la tabla para que añada la nueva línea
            
            Sheets("POR ARCHIVAR").Activate
            
            Sheets("POR ARCHIVAR").Cells(i, inicioj).EntireRow.Delete
            i = i - 1
            
        End If
    
    Next
    
End Sub

