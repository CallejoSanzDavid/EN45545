Attribute VB_Name = "Módulo2"
Sub ARCHIVADOS()
    
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
    
    inicioi = Sheets("POR ARCHIVAR").Range("A1:A10").Find("PART NUMBER").Row            'Posiciones
    inicioj = Sheets("POR ARCHIVAR").Range("A1:Z1").Find("PART NUMBER").Column
    
    finali = Sheets("POR ARCHIVAR").Cells(Rows.Count, inicioj).End(xlUp).Row
    finalj = Sheets("POR ARCHIVAR").Cells(inicioi, Columns.Count).End(xlToLeft).Column
    
    estadoj = Sheets("POR ARCHIVAR").Range(Cells(inicioi, inicioj), Cells(inicioi, finalj)).Find("ESTADO").Column
    
    For i = inicioi + 1 To finali
    
        estado = Sheets("POR ARCHIVAR").Cells(i, estadoj).Value
        
        If estado = "OK" Then          'Cortar y pegar si cumple
            
            Sheets("POR ARCHIVAR").Range(Cells(i, inicioj), Cells(i, finalj)).Cut
            auxfinali = Sheets("ARCHIVADOS").Cells(Rows.Count, estadoj).End(xlUp).Row + 1
            Sheets("ARCHIVADOS").Activate
            Cells(auxfinali, "A").Select
            ActiveSheet.Paste
            
            Sheets("POR ARCHIVAR").Cells(i, inicioj).EntireRow.Delete
            i = i - 1
            
            Sheets("POR ARCHIVAR").Activate
            
        End If
    
    Next
    
End Sub

