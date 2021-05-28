VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    
    Dim paleta(9) As String
    Dim i As Integer
    Dim Inicioi As Integer
    Dim supplierj As Integer
    Dim ContactNamej As Integer
    Dim Countryj As Integer
    Dim mailj As Integer
    Dim OKj As Integer
    Dim N As Integer
    Dim Aux As Integer
    
    Aux = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Merak").Row + 1
    
    paleta(1) = 35
    paleta(3) = 36
    paleta(5) = 38
    paleta(7) = 45
    paleta(9) = 39

    i = 0
    
    supplierj = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Supplier").Column
    
    N = Sheets("OG PROVEEDORES").Cells(Rows.Count, supplierj).End(xlUp).Row
    
    With ActiveSheet.ListObjects("Tabla1").Sort         'Ordenar alfabeticamente una tabla
        
        .SortFields.Clear                               'Elimina los filtros activos
        .SortFields.Add Key:=Range("Tabla1[Supplier]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal        'Selecciona el filtro y donde aplicarlo
        .Apply                                          'Aplicar
    
    End With
    
    ContactNamej = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Contact Name").Column
    Countryj = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Country").Column
    mailj = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Mail").Column
    OKj = Sheets("OG PROVEEDORES").Range("A2:H2").Find("OK/NOK").Column
    
    For Inicioi = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Merak").Row + 1 To N  'Poner en mayúscula la columna de supplier
    
        Application.StatusBar = "Format Progress (1/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        
        Sheets("OG PROVEEDORES").Cells(Inicioi, supplierj).Value = UCase(Sheets("OG PROVEEDORES").Cells(Inicioi, supplierj).Value)
    
    Next


    For Inicioi = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Merak").Row + 1 To N    'Rellenar celdas con colores
        
        Application.StatusBar = "Format Progress (2/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        
        If i Mod 2 = 0 Then
            paleta(i) = 2
        End If
        
        If Sheets("OG PROVEEDORES").Cells(Inicioi, mailj).Value <> "" And Sheets("OG PROVEEDORES").Cells(Inicioi, OKj).Value = "NO HAY CONTACTO" And Inicioi <= N Then
        
            Sheets("OG PROVEEDORES").Cells(Inicioi, OKj).Value = ""
            
        End If
        
        Sheets("OG PROVEEDORES").Range("A" & Inicioi & ":H" & Inicioi).Interior.ColorIndex = paleta(i)
        
        If Sheets("OG PROVEEDORES").Cells(Inicioi + 1, supplierj).Value = Sheets("OG PROVEEDORES").Cells(Inicioi, supplierj).Value And Inicioi <= N Then
            
            Sheets("OG PROVEEDORES").Range("A" & Inicioi + 1 & ":H" & Inicioi + 1).Interior.ColorIndex = paleta(i)
            
        Else
        
            i = i + 1
            
        End If
        
        If i > 9 Then
        
            i = 0
            
        End If

    Next
    
    For Inicioi = Sheets("OG PROVEEDORES").Range("A2:H2").Find("Merak").Row + 1 To N    'Comprobar si hay información de contacto
        
        Application.StatusBar = "Format Progress (3/3): " & Inicioi - Aux & " of " & N - Aux & ": " & Format((Inicioi - Aux) / (N - Aux), "0%")
        
        If Sheets("OG PROVEEDORES").Cells(Inicioi, mailj).Value = "" And Inicioi <= N Then
            
            Sheets("OG PROVEEDORES").Cells(Inicioi, OKj).Value = "NO HAY CONTACTO"
            Sheets("OG PROVEEDORES").Range("A" & Inicioi & ":H" & Inicioi).Interior.ColorIndex = 3
            
        End If
    
    Next
    
    
    Application.StatusBar = ""
    
End Sub

