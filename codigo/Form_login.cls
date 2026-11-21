VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_voltar_Click()
DoCmd.Close
DoCmd.OpenForm "tela_inicial", acNormal

End Sub

Private Sub Form_Load()
conectar_banco

DoCmd.Maximize

End Sub
Private Sub btn_entrar_Click()
    sql = "select * from tb_login where (usuario ='" & txt_usuario & "' or email = '" & txt_usuario & "') and senha = '" & txt_senha & "'"
    validar_leitura
    If rs.EOF = False Then
        txt_email = rs.Fields(3)
        txt_usuario = rs.Fields(4)
        txt_senha = rs.Fields(5)
        
       DoCmd.Close

        DoCmd.OpenForm "inicio", acNormal
    Else
        MsgBox ("Usuário não localizado!"), vbInformation + vbOKOnly, "AVISO"
    End If
End Sub

