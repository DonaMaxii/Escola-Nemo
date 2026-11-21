VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_login_prof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_relatorios_Click()
DoCmd.Close 'Fecha o formulário atual
DoCmd.OpenForm ("relatorios"), acNormal 'Abre o form destino

End Sub

Private Sub btn_voltar_Click()
DoCmd.Close
DoCmd.OpenForm "tela_inicial", acNormal
End Sub

Private Sub Form_Load()
conectar_banco
DoCmd.Maximize
End Sub
Private Sub btn_entrar_Click()
If IsNull(txt_usuario) Or IsNull(txt_senha) Then
    MsgBox ("Por favor, preencha os campos para continuar."), vbExclamation + vbOKOnly, "Login"
Else
    sql = "select * from tb_prof where (usuario ='" & txt_usuario & "' or email = '" & txt_usuario & "') and senha = '" & txt_senha & "'"
    validar_leitura
    If rs.EOF = False Then 'logar
        txt_email = rs.Fields(3)
        txt_usuario = rs.Fields(4)
        txt_senha = rs.Fields(5)
        
        DoCmd.Close
        DoCmd.OpenForm "inicio_prof", acNormal
   Else
       MsgBox ("Dados incorretos ou usuário não localizado." & vbNewLine & "Revise os dados informados e tente novamente."), vbExclamation + vbOKOnly, "AVISO"
  End If
End If
End Sub
