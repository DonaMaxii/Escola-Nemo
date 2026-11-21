VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_inicio_prof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub btn_alunos_Click()
DoCmd.Close
DoCmd.OpenForm "cadastro_alunos", acNormal
End Sub

Private Sub btn_cad_prof_Click()
DoCmd.Close
DoCmd.OpenForm "cadastro_prof", acNormal
End Sub

Private Sub btn_aluno_turma1_Click()
Dim stDocName As String
    stDocName = "relat_todas_turmas"
    DoCmd.OpenReport stDocName, acViewReport
    DoCmd.Maximize
End Sub

Private Sub btn_encerrar_Click()
resp = MsgBox("Deseja realmente sair?", vbQuestion + vbYesNo, "ATENÇÃO")
If resp = vbYes Then
   DoCmd.Save
   
   DoCmd.Close
   
   DoCmd.OpenForm "tela_inicial", acNormal
End If
End Sub

Private Sub btn_logoff_Click()
resp = MsgBox("Deseja realizar logoff?", vbInformation + vbYesNo, "Logoff")
If resp = vbYes Then
    DoCmd.Close
    DoCmd.OpenForm "tela_inicial", acNormal
End If
End Sub

Private Sub btn_niver_Click()
Dim stDocName As String
    stDocName = "relat_niver_mes"
    DoCmd.OpenReport stDocName, acViewReport
    DoCmd.Maximize
End Sub


Private Sub btn_relat_alunos_Click()
Dim stDocName As String
    stDocName = "relat_matriculados"
    DoCmd.OpenReport stDocName, acViewReport
    DoCmd.Maximize
End Sub

Private Sub btn_relat_professor_Click()
Dim stDocName As String
    stDocName = "relat_professores"
    DoCmd.OpenReport stDocName, acViewReport
    DoCmd.Maximize
End Sub

Private Sub btn_relat_turma_Click()
Dim stDocName As String
    stDocName = "relat_turmas"
    DoCmd.OpenReport stDocName, acViewReport
    DoCmd.Maximize
End Sub

Private Sub Form_Load()
DoCmd.Maximize
End Sub
