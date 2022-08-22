Attribute VB_Name = "LocatePositions"
Option Explicit
    
    Public wb As Workbook
    
'----------Locate_Positions_OG----------
    Public ws_OG As Object
    'Rows
    Public Starti As Integer
    Public Endi As Integer
    'Columns
    Public PartNumj As Integer
    Public PartNamej As Integer
    Public RawMatj As Integer
    Public Supplierj As Integer
    Public TRj As Integer
    Public Contactj As Integer
    Public Whoj As Integer
    Public Whenj As Integer
    Public LastMsgj As Integer
    Public Statusj As Integer
    Public Commentsj As Integer
    Public AdActj As Integer
    
'----------Locate_Positions_PA----------
    Public PA_SheetName As String
    Public ws_PA As Object
    'Rows
    Public PA_Starti As Integer
    Public PA_Endi As Integer
    'Columns
    Public PA_PartNumj As Integer
    Public PA_PartNamej As Integer
    Public PA_RawMatj As Integer
    Public PA_Supplierj As Integer
    Public PA_Statusj As Integer
    Public PA_StatusLetterj As String
   
Sub Locate_Positions_OG(SheetName As String)
'Localiza las posiciones necesarias enla hoja activa para el correcto funcionamiento del código.
'OG: OriGinal -> Página desde la que se acciona la macro.

    Set wb = ThisWorkbook
    Set ws_OG = wb.Sheets(SheetName)
    
    Starti = Find_Row(0, "PART NUMBER", SheetName)
    
    PartNumj = Find_Column(Starti, "PART NUMBER", SheetName)
    PartNamej = Find_Column(Starti, "PART NAME", SheetName)
    RawMatj = Find_Column(Starti, "RAW MATERIAL", SheetName)
    Supplierj = Find_Column(Starti, "SUPPLIER", SheetName)
    TRj = Find_Column(Starti, "TR NUMBER*", SheetName)
    Contactj = Find_Column(Starti, "CONTACT EMAIL", SheetName)
    Whoj = Find_Column(Starti, "QUIÉN LO PIDE", SheetName)
    Whenj = Find_Column(Starti, "CUANDO SE HA PEDIDO", SheetName)
    LastMsgj = Find_Column(Starti, "FECHA DE ÚLTIMO CORREO ENVIADO", SheetName)
    Statusj = Find_Column(Starti, "ESTADO", SheetName)
    Commentsj = Find_Column(Starti, "COMENTARIOS", SheetName)
    AdActj = Find_Column(Starti, "ACCIONES ADICIONALES", SheetName)
    
    Endi = ws_OG.Cells(Rows.Count, PartNumj).End(xlUp).Row
    
End Sub

Sub Locate_Positions_PA(SheetName As String)
'Localiza las posiciones necesarias enla hoja activa para el correcto funcionamiento del código.
'PA: POR ARCHIVAR

    Dim AuxSheet As String
    
    'Para obtener las posiciones en la hoja "POR ARCHIVAR" necesitamos activar dicha hoja.
    'Si la hoja activa es otra, usaremos esta variable para volver a activar la hoja desde la que se inició la macro.
    AuxSheet = ActiveSheet.Name
    
    PA_SheetName = "POR ARCHIVAR"
    Set wb = ThisWorkbook
    Set ws_PA = wb.Sheets(PA_SheetName)
    
    ws_PA.Activate
    
    PA_Starti = Find_Row(0, "PART NUMBER", PA_SheetName)
    
    PA_PartNumj = Find_Column(PA_Starti, "PART NUMBER", PA_SheetName)
    PA_PartNamej = Find_Column(PA_Starti, "PART NAME", PA_SheetName)
    PA_RawMatj = Find_Column(PA_Starti, "RAW MATERIAL", PA_SheetName)
    PA_Supplierj = Find_Column(PA_Starti, "SUPPLIER", PA_SheetName)
    
    PA_Statusj = Find_Column(PA_Starti, "ESTADO", PA_SheetName)
    'Obtiene la letra de la columna.
    PA_StatusLetterj = Mid(ws_PA.Cells(PA_Starti, PA_Statusj).Address, 2, InStr(2, ws_PA.Cells(PA_Starti, PA_Statusj).Address, "$") - 2)         '= F
    
    PA_Endi = ws_PA.Cells(Rows.Count, PA_Supplierj).End(xlUp).Row
    
    'Activa la hoja desde la que se inició la macro.
    Sheets(AuxSheet).Activate
    
End Sub

Function Find_Row(Column_Lim As Integer, Find_Me As String, Sheet As String) As Integer
    
    'Si Column_Lim = 0 quiere decir que aún no se ha localizado una línea de referencia para buscar los valores.
    'Por lo que se usa un rango lo suficientemente grande como para encontrar la palabra clave.
        
    If Column_Lim = 0 Then
    
        Find_Row = Sheets(Sheet).Range("A1:DA20").Find(Find_Me).Row
    
    Else
    
        Find_Row = Sheets(Sheet).Range(Cells(1, Column_Lim), Cells(100, Column_Lim)).Find(Find_Me).Row
        
    End If
    
End Function

Function Find_Column(Row_Lim As Integer, Find_Me As String, Sheet As String) As Integer
    
    'Si Row_Lim = 0 quiere decir que aún no se ha localizado una línea de referencia para buscar los valores.
    'Por lo que se usa un rango lo suficientemente grande como para encontrar la palabra clave.
        
    If Row_Lim = 0 Then
    
        Find_Column = Sheets(Sheet).Range("A1:DA20").Find(Find_Me).Column
    
    Else
    
        Find_Column = Sheets(Sheet).Range(Cells(Row_Lim, 1), Cells(Row_Lim, 100)).Find(Find_Me).Column
        
    End If
    
End Function
