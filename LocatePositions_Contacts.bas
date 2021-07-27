Attribute VB_Name = "LocatePositions_Contacts"
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
    Public EmailSendedj As Integer
  
Sub Locate_Positions_OG()
'Locates the necesary positions in the current activated sheet for the correct function of the code.
'OG: OriGinal -> Page from where the macros are initiated.

    SheetName = ActiveSheet.Name
    
    Set wb = ThisWorkbook
    Set ws_contact = wb.Sheets(SheetName)
    
    CPsupplieri = ws_contact.Range("A1:Z10").Find("Supplier").Row   '= 1
    CPAuxi = CPsupplieri + 1                                        '= 2
    
    CPvendorcodej = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Vendor Code").Column  'A = 1
    'Obtiene la letra de la columna.
    NC_CPvendorcodej = Mid(ws_contact.Cells(CPsupplieri, CPvendorcodej).Address, 2, InStr(2, ws_contact.Cells(CPsupplieri, CPvendorcodej).Address, "$") - 2) '= A
    
    CPsupplierj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 100)).Find("Supplier").Column      'B = 2
    CPmailj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 100)).Find("Mail").Column              'D = 4
    CPtlfnoj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Telephone").Column         'E = 5
    CPcountryj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Country").Column         'F = 6
    CPlanguagej = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Language").Column       'G = 7
    
    CPOKj = ws_contact.Range("A1:Z1").Find("OK/NOK").Column                                                     'H = 8
    'Obtiene la letra de la columna.
    NC_CPOKj = Mid(ws_contact.Cells(CPsupplieri, CPOKj).Address, 2, InStr(2, ws_contact.Cells(CPsupplieri, CPOKj).Address, "$") - 2) '= H
    
    CPendi = ws_contact.Cells(Rows.Count, CPsupplierj).End(xlUp).Row    '= 307
    
End Sub

Sub Locate_Positions_DDBB()
'Locates the necesary positions for the correct function of the code.
'CP: Contact Page.

    Dim auxwb As Workbook
    Dim AuxSheet As String
    Dim RoutesSheetName As String
    Dim Routei As Integer
    Dim Routej As Integer
    Dim FCILSheetName As String
    
    AuxSheet = ActiveSheet.Name
    Set auxwb = ThisWorkbook
    
    RoutesSheetName = "Routes"
    auxwb.Sheets(RoutesSheetName).Activate
    
    Routei = Sheets(RoutesSheetName).Range("A1:Z20").Find("EN45545 DDBB").Row                                    '= 2
    Routej = Sheets(RoutesSheetName).Range("A1:Z20").Find("FULL ROUTE OF THE CONF. SHEET DOCUMENT").Column       'B = 2
    
    Workbooks.Open (Sheets(RoutesSheetName).Cells(Routei, Routej).Value)
    
    wb_OG = ActiveWorkbook.Name
    FCILSheetName = "FCIL"
    Set ws_OG = Workbooks(wb_OG).Sheets(FCILSheetName)
    
    ws_OG.Activate
    
    Aux = ws_OG.Range("A1:DA20").Find("Assembly Name").Row          '= 10
    
    nprodj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier part number").Column                            'N = 14
    nombj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Part name").Column                                        'P = 16
    matj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Raw material or product name*").Column                     'Q = 17
    manufj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer name*").Column                              'R = 18
    
    DateT6j = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Date * T6").Column                                      'BW = 75
    ManufDeclarationj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer Declaration Date").Column        'CB = 80
    GlobalStatusj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Certificate global status*").Column               'CD = 82
    EmailSendedj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Email Sended").Column                              'CE = 83
    TMexpirej = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Test Method 1 time to expire*").Column                'CF = 84
    ContactDBj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier's Contact").Column                          'CL = 90
                                                                                                                        'EmailGen: contactj
                                                                                                                        'Public contactj As Integer
    N = ws_OG.Cells(Rows.Count, nprodj).End(xlUp).Row               '= 919
    
    auxwb.Sheets(AuxSheet).Activate
    
End Sub
