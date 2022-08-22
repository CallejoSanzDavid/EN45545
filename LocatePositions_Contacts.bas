Attribute VB_Name = "LocatePositions"
Option Explicit
    
    Public wb As Workbook
    
'----------Locate_Positions_OG----------
    Public SheetName As String
    Public ws_contact As Object
    'Rows
    Public CPsupplieri As Integer
    Public CPAuxi As Integer
    Public CPendi As Integer
    'Columns
    Public CPvendorcodej As Integer
    Public NC_CPvendorcodej As String
    Public CPsupplierj As Integer
    Public CPmailj As Integer
    Public CPtlfnoj As Integer
    Public CPcountryj As Integer
    Public CPlanguagej As Integer
    Public CPOKj As Integer
    Public NC_CPOKj As String
    
'----------Locate_Positions_DDBB----------
    Public ws_OG As Object
    Public wb_OG As String
    'Rows
    Public Auxi As Integer   'Initial
    Public N As Integer     'Final
    'Columns
    Public nprodj As Integer
    Public manufj As Integer
  
Sub Locate_Positions_CP()
'Localiza las posiciones necesarias enla hoja activa para el correcto funcionamiento del código.
'CP: Contact Page.

    SheetName = ActiveSheet.Name
    
    Set wb = ThisWorkbook
    Set ws_contact = wb.Sheets(SheetName)
        
    CPsupplieri = Find_Row(0, "Supplier", SheetName)
    CPAuxi = CPsupplieri + 1
    
    CPvendorcodej = Find_Column(CPsupplieri, "Vendor Code", SheetName)
    'Obtiene la letra de la columna.
    NC_CPvendorcodej = Mid(ws_contact.Cells(CPsupplieri, CPvendorcodej).Address, 2, InStr(2, ws_contact.Cells(CPsupplieri, CPvendorcodej).Address, "$") - 2) '= A
    
    CPsupplierj = Find_Column(CPsupplieri, "Supplier", SheetName)
    CPmailj = Find_Column(CPsupplieri, "Mail", SheetName)
    CPtlfnoj = Find_Column(CPsupplieri, "Telephone", SheetName)
    CPcountryj = Find_Column(CPsupplieri, "Country", SheetName)
    CPlanguagej = Find_Column(CPsupplieri, "Language", SheetName)
    
    CPOKj = Find_Column(CPsupplieri, "OK/NOK", SheetName)
    'Obtiene la letra de la columna.
    NC_CPOKj = Mid(ws_contact.Cells(CPsupplieri, CPOKj).Address, 2, InStr(2, ws_contact.Cells(CPsupplieri, CPOKj).Address, "$") - 2) '= H
    
    CPendi = ws_contact.Cells(Rows.Count, CPsupplierj).End(xlUp).Row
    
End Sub

Sub Locate_Positions_DDBB()
'Localiza las posiciones necesarias enla hoja activa para el correcto funcionamiento del código.
'OG: OriGinal -> Página FCIL de BBDD de F&H. Relacionada a la hoja AUX FCIL.

    Dim FCILSheetName As String
    
    FCILSheetName = "AUX FCIL"
    Set ws_OG = wb.Sheets(FCILSheetName)
    
    ws_OG.Activate
    
    Auxi = Find_Row(0, "Supplier part number", FCILSheetName)
    
    nprodj = Find_Column(Auxi, "Supplier part number", FCILSheetName)
    manufj = Find_Column(Auxi, "Manufacturer name*", FCILSheetName)
    
    N = ws_OG.Cells(Rows.Count, nprodj).End(xlUp).Row
    
    ws_contact.Activate
    
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
