VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_tela_inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub btn_encerrar_Click()
resp = MsgBox("Deseja realmente sair do sistema?", vbQuestion + vbYesNo, "ATENÇÃO")
If resp = vbYes Then
   DoCmd.Quit acQuitSaveAll
End If
End Sub
Private Sub btn_login_adm_Click()
DoCmd.Close
DoCmd.OpenForm "login", acNormal
End Sub

Private Sub btn_login_professor_Click()
DoCmd.Close
DoCmd.OpenForm "login_prof", acNormal
End Sub

Private Sub Form_Load()
DoCmd.Maximize
End Sub
