Attribute VB_Name = "Test_ClearFormat"
Sub Formato_Prueba()                'Limpia los rangos donde actua la macro para comprobar el funcionamiento

    Auxi = Sheets("FCIL").Range("A10:DA10").Find("Assembly Name").Row + 1
        
    G = Sheets("FCIL").Range("A10:DA10").Find("Certificate global status*").Column
    x1 = Sheets("FCIL").Range("A10:DA10").Find("Test Method 1 time to expire*").Column
    x2 = Sheets("FCIL").Range("A10:DA10").Find("Test Method 6 time to expire*").Column + 1
    
    N = Sheets("FCIL").Cells(Rows.Count, G).End(xlUp).Row
    
    Sheets("FCIL").Range(Cells(Auxi, G), Cells(N, G)).ClearContents
    Sheets("FCIL").Range(Cells(Auxi, G), Cells(N, G)).Interior.ColorIndex = 41
    
    Sheets("FCIL").Range(Cells(Auxi, x1), Cells(N, x2)).ClearContents
    Sheets("FCIL").Range(Cells(Auxi, x1), Cells(N, x2)).Interior.ColorIndex = 41
    
End Sub
 l t . C o d e " :   2 6 ,   " A c t i v i t y . R e s u l t . T y p e " :   " O S M R e s u l t " ,   " A c t i v i t y . A g g I n t e r v a l " :   1 }                                                                             (P┌ └Юог  PKог                                                                                                                                                                                                                                                                 