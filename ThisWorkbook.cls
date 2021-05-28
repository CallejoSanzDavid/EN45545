VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    Dim Ans As String
    
    If Sheets("FCIL").FilterMode Then Sheets("FCIL").ShowAllData
    
    Ans = MsgBox("¿Desea actualizar las alarmas del archivo?", vbYesNo)
    
    If Ans = vbYes Then
        
        Call Comprobar_Caducidad
    
    End If
    
End Sub


