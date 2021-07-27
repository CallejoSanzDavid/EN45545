Attribute VB_Name = "LocatePositions"
Option Explicit
    
    Public wb As Workbook
    
'----------Locate_Positions_OG----------
    Public SheetName As String
    Public ws_OG As Object
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
    
'----------Locate_Positions_Contacts----------
    Public ContactSheetName As String
    Public ws_contact As Object
    'Rows
    Public CPsupplieri As Integer
    Public CPendi As Integer
    'Columns
    Public CPvendorcodej As Integer
    Public CPsupplierj As Integer
    Public CPmailj As Integer
    Public CPtlfnoj As Integer
    Public CPcountryj As Integer
    Public CPlanguagej As Integer

'----------Locate_Positions_RankingStatus----------
    Public RankingStatusSheet As String
    Public ws_ranking As Object
    'Rows
    Public RSRankingi As Integer
    Public RSEndi As Integer
    'Columns
    Public RSRankingj As Integer
    Public RSStatusENj As Integer
    Public RSStatusESj As Integer
    Public RSColorCodej As Integer
    
'----------Locate_Positions_Email_Body----------
    Public EmailBodySheetName As String
    Public ws_emailb As Object
    'Rows
    Public EBcci As Integer
    Public EBEndi As Integer
    Public EBSubjectENi As Integer
    Public EBSubjectESi As Integer
    Public EBAttachmenti As Integer
    Public EBHeadingENi As Integer
    Public EBFarewellENi As Integer
    Public EBSeparationi As Integer
    Public EBHeadingESi As Integer
    Public EBFarewellESi As Integer
    Public EBSignaturei As Integer
    'Columns
    Public EBccj As Integer
    Public EBInfoj As Integer
    
Sub Locate_Positions_OG()
'Locates the necesary positions in the current activated sheet for the correct function of the code.
'OG: OriGinal -> Page from where the macros are initiated.

    SheetName = ActiveSheet.Name
    
    Set wb = ThisWorkbook
    Set ws_OG = wb.Sheets(SheetName)
    
    Aux = ws_OG.Range("A1:DA20").Find("Assembly Name").Row
    
    nprodj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier part number").Column
    nombj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Part name").Column
    matj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Raw material or product name*").Column
    manufj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer name*").Column
    
    DateT6j = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Date * T6").Column
    ManufDeclarationj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Manufacturer Declaration Date").Column
    GlobalStatusj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Certificate global status*").Column
    EmailSendedj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Email Sended").Column
    TMexpirej = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Test Method 1 time to expire*").Column
    ContactDBj = ws_OG.Range(Cells(Aux, 1), Cells(Aux, 100)).Find("Supplier's Contact").Column                      'EmailGen: contactj
                                                                                                                                'Public contactj As Integer
    N = ws_OG.Cells(Rows.Count, nprodj).End(xlUp).Row
    
End Sub

Sub Locate_Positions_Contacts()
'Locates the necesary positions for the correct function of the code.
'CP: Contact Page.

    Dim AuxSheet As String
    
    AuxSheet = ActiveSheet.Name
    
    ContactSheetName = "Suppliers Contact Info"
    Set ws_contact = wb.Sheets(ContactSheetName)
    
    ws_contact.Activate
    
    CPsupplieri = ws_contact.Range("A1:Z10").Find("Supplier").Row       '= 1
    
    CPvendorcodej = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Vendor Code").Column      'A = 1
    CPsupplierj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 100)).Find("Supplier").Column          'B = 2
    CPmailj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 100)).Find("Mail").Column                  'D = 4
    CPtlfnoj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Telephone").Column             'E = 5
    CPcountryj = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Country").Column             'F = 6
    CPlanguagej = ws_contact.Range(Cells(CPsupplieri, 1), Cells(CPsupplieri, 10)).Find("Language").Column           'G = 7
    
    CPendi = ws_contact.Cells(Rows.Count, CPsupplierj).End(xlUp).Row    '= 307
    
    Sheets(AuxSheet).Activate
    
End Sub

Sub Locate_Positions_RankingStatus()
'Locates the necesary positions for the correct function of the code.
'RS: Ranking Sheet

    Dim AuxSheet As String
    
    AuxSheet = ActiveSheet.Name
    
    RankingStatusSheet = "Ranking Status"
    Set ws_ranking = wb.Sheets(RankingStatusSheet)
    
    ws_ranking.Activate
        
    RSRankingi = ws_ranking.Range("A1:Z10").Find("Ranking").Row

    RSRankingj = ws_ranking.Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Ranking").Column
    RSStatusENj = ws_ranking.Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Status (EN)").Column
    RSStatusESj = ws_ranking.Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Status (ES)").Column
    RSColorCodej = ws_ranking.Range(Cells(RSRankingi, 1), Cells(RSRankingi, 20)).Find("Color Code").Column
    
    RSEndi = ws_ranking.Cells(Rows.Count, RSRankingj).End(xlUp).Row
    
    Sheets(AuxSheet).Activate
    
End Sub

Sub Locate_Positions_Email_Body()
'Locates the necesary positions for the correct function of the code.
'EB: Email Body

    Dim AuxSheet As String
    
    AuxSheet = ActiveSheet.Name
    
    EmailBodySheetName = "Email Body"
    Set ws_emailb = wb.Sheets(EmailBodySheetName)
    
    ws_emailb.Activate
    
    EBcci = ws_emailb.Range("A1:Z10").Find("CC").Row           '= 1
    EBccj = ws_emailb.Range(Cells(EBcci, 1), Cells(EBcci, 10)).Find("CC").Column     'A = 1
    
    EBInfoj = EBccj + 1     'B = 2
    
    EBEndi = ws_emailb.Cells(Rows.Count, EBccj).End(xlUp).Row
    
    EBSubjectENi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("SubjectEN").Row         '= 2
    EBSubjectESi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("SubjectES").Row         '= 3
    
    EBAttachmenti = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("Attachment").Row       '= 4
    
    EBHeadingENi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("HeadingEN").Row         '= 5
    EBFarewellENi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("FarewellEN").Row       '= 7
    
    EBSeparationi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("Separation").Row       '= 8
    
    EBHeadingESi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("HeadingES").Row         '= 9
    EBFarewellESi = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("FarewellES").Row       '= 11
    
    EBSignaturei = ws_emailb.Range(Cells(EBcci, EBccj), Cells(EBEndi, EBccj)).Find("Signature").Row         '= 12
    
    Sheets(AuxSheet).Activate

End Sub













