Attribute VB_Name = "LocatePositions"
Option Explicit
'----------Locate_Positions_OG----------
    Public SheetName As String
    'Rows
    Public Aux As Integer   'Initial
    Public N As Integer     'Final
    'Columns
    Public nprodj As Integer
    Public nombj As Integer
    Public matj As Integer
    Public manufj As Integer
    Public DateT6j As Integer
    Public ManufDeclarationj As Integer
    Public GlobalStatusj As Integer
    Public TMexpirej As Integer
    Public ContactDBj As Integer

'----------Locate_Positions_Contacts----------
    Public ContactSheetName As String
    'Rows
    Public CPendi As Integer
    Public CPsupplieri As Integer
    'Columns
    Public CPmailj As Integer
    Public CPsupplierj As Integer

'----------Locate_Positions_RankingStatus----------
    Public RankingStatusSheet As String
    'Rows
    Public RSRankingi As Integer
    Public RSEndi As Integer
    'Columns
    Public RSRankingj As Integer
    Public RSStatusENj As Integer
    Public RSStatusESj As Integer
    Public RSColorCodej As Integer
    
Sub Locate_Positions_OG()
'Locates the necesary positions in the current activated sheet for the correct function of the code.
    
    SheetName = ActiveSheet.Name
    
    Aux = Sheets(SheetName).Range("A1:DA20").Find("Assembly Name").Row
    
    nprodj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier part number").Column
    nombj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Part name").Column
    matj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Raw material or product name*").Column
    manufj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer name*").Column
    
    DateT6j = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Date * T6").Column
    ManufDeclarationj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer Declaration Date").Column
    GlobalStatusj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Certificate global status*").Column
    TMexpirej = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Test Method 1 time to expire*").Column
    ContactDBj = Sheets(SheetName).Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier's Contact").Column                      'EmailGen: contactj
                                                                                                                                'Public contactj As Integer
    N = Sheets(SheetName).Cells(Rows.Count, nprodj).End(xlUp).Row
    
End Sub

Sub Locate_Positions_Contacts()
'Locates the necesary positions for the correct function of the code.
    
    Dim AuxSheet As String
    
    AuxSheet = ActiveSheet.Name
    
    ContactSheetName = "Contacto de proveedores"
    Sheets(ContactSheetName).Activate
    
    CPsupplieri = Sheets(ContactSheetName).Range("A1:Z10").Find("Supplier").Row + 1
    
    CPsupplierj = Sheets(ContactSheetName).Range(Cells(CPsupplieri - 1, 1), Cells(CPsupplieri - 1, 100)).Find("Supplier").Column
    CPmailj = Sheets(ContactSheetName).Range(Cells(CPsupplieri - 1, 1), Cells(CPsupplieri - 1, 100)).Find("Mail").Column
    
    CPendi = Sheets(ContactSheetName).Cells(Rows.Count, CPsupplierj).End(xlUp).Row
    
    Sheets(AuxSheet).Activate
    
End Sub

Sub Locate_Positions_RankingStatus()
    
    Dim AuxSheet As String
    
    AuxSheet = ActiveSheet.Name
    
    RankingStatusSheet = "Ranking Status"
    Sheets(RankingStatusSheet).Activate
        
    RSRankingi = Sheets(RankingStatusSheet).Range("A1:Z10").Find("Ranking").Row

    RSRankingj = Sheets(RankingStatusSheet).Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Ranking").Column
    RSStatusENj = Sheets(RankingStatusSheet).Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Status (EN)").Column
    RSStatusESj = Sheets(RankingStatusSheet).Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Status (ES)").Column
    RSColorCodej = Sheets(RankingStatusSheet).Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Color Code").Column
    
    RSEndi = Sheets(RankingStatusSheet).Cells(Rows.Count, RSRankingj).End(xlUp).Row
    
    Sheets(AuxSheet).Activate
    
End Sub















