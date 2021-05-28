Attribute VB_Name = "Módulo3"
Sub ODD1OUT()               'Este código encuentra inconsistencias en los Part Names

    nprodj = Sheets("FCIL").Range("A10:DA10").Find("Supplier part number").Column
    N = Sheets("FCIL").Cells(Rows.Count, nprodj).End(xlUp).Row
    
    For i = Sheets("FCIL").Range("M1:M15").Find("Assembly Name").Row + 1 To N
              
        nprodj = Sheets("FCIL").Range("A10:DA10").Find("Supplier part number").Column
        nproducto = Sheets("FCIL").Cells(i, nprodj).Value
        
        nombj = Sheets("FCIL").Range("A10:DA10").Find("Part name").Column
        auxname = Split(Cells(i, nombj).Value, " - MATERIAL")
        nombre = auxname(0)
        
        auxname1 = Split(Cells(i + 1, nombj).Value, " - MATERIAL")
        nombre1 = auxname1(0)
        
        If nproducto = Sheets("FCIL").Cells(i + 1, nprodj).Value And nombre <> nombre1 Then
        
            ActiveSheet.Cells(i, nombj).Select              'Poner debugger aquí y correr programa
        
        End If

    Next

End Sub

