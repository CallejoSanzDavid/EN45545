Attribute VB_Name = "Módulo4"
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
