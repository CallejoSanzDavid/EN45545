Attribute VB_Name = "Test_PartNameCheck"
Sub ODD1OUT()               'Este c�digo encuentra inconsistencias en los Part Names

    nprodj = Sheets("FCIL").Range("A10:DA10").Find("Supplier part number").Column
    nombj = Sheets("FCIL").Range("A10:DA10").Find("Part name").Column
    
    N = Sheets("FCIL").Cells(Rows.Count, nprodj).End(xlUp).Row
    
    For i = Sheets("FCIL").Range("M1:M15").Find("Assembly Name").Row + 1 To N
              
        nproducto = Sheets("FCIL").Cells(i, nprodj).Value
        
        auxname = Split(Cells(i, nombj).Value, " - MATERIAL")
        nombre = auxname(0)
        
        auxname1 = Split(Cells(i + 1, nombj).Value, " - MATERIAL")
        nombre1 = auxname1(0)
        
        If nproducto = Sheets("FCIL").Cells(i + 1, nprodj).Value And nombre <> nombre1 Then
        
            ActiveSheet.Cells(i, nombj).Select              'Poner debugger aqu� y correr programa
        
        End If

    Next

End Sub

 y . R e s u l t . T a g " :   2 3 7 0 0 1 2 9 ,   " D a t a . N u l l O L D o c " :   f a l s e ,   " D a t a . W o r k b o o k I d " :   " { 1 E 7 C D D 3 9 - 5 9 6 5 - 4 3 1 B - 9 F F 7 - 5 2 5 5 5 0 2 0 0 D D 3 } " ,   " D a t a . C a c h e F i l e S i z e " :   " 3 3 4 9 " ,   " D a t a . P a r s e U s e r " :   4 0 }         S  ��ɫ  ��,ʫ                                                                                            S  ��ɫ  u���                                               