VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_cadastro_prof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub btn_buscar_prof_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strCPF As String
    Dim strSQL As String

    Set db = CurrentDb
    strCPF = txt_cpf.Value

    If Len(strCPF) = 0 Then
        MsgBox "Por favor, digite o CPF do professor para buscar.", vbInformation
        txt_cpf.SetFocus
        Exit Sub
    End If

    strSQL = "SELECT CPF, NOME, DATA_NASC, EMAIL, CREF, USUARIO, SENHA, CSENHA FROM tb_prof WHERE CPF = '" & strCPF & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)

    If Not rs.EOF Then
        ' Professor encontrado, preenche os campos
        txt_cpf.Value = Nz(rs!CPF, "")
        txt_nome.Value = Nz(rs!nome, "")
        txt_email.Value = Nz(rs!EMAIL, "")
        txt_data_nasc.Value = Nz(rs!DATA_NASC, Null)
        txt_cref.Value = Nz(rs!CREF, Null)
        txt_usuario.Value = Nz(rs!USUARIO, Null)
        txt_senha.Value = Nz(rs!SENHA, Null)
        txt_csenha.Value = Nz(rs!CSENHA, Null)
        
        MsgBox "Professor encontrado e dados preenchidos.", vbInformation '
    Else
        ' Professor não encontrado
        MsgBox "Professor não encontrado. Por favor, preencha os dados.", vbInformation
        txt_nome.Value = ""
        txt_email.Value = ""
        txt_data_nasc.Value = ""
        txt_cref.Value = ""
        txt_usuario.Value = ""
        txt_senha.Value = ""
        txt_csenha.Value = ""
        txt_cpf.SetFocus
    End If
End Sub
Private Sub btn_excluir1_Click()
Set db = CurrentDb
resp = MsgBox("Deseja realmente excluir o Professor?", vbQuestion + vbYesNo, "ATENÇÃO")
    If resp = vbYes Then
        sql = " delete * from tb_prof where CPF = '" & txt_cpf & "' "
        db.Execute (sql)
        MsgBox ("Professor excluído com sucesso!!!"), vbInformation + vbOKOnly, "AVISO"
    
        txt_cpf = ""
        txt_nome = ""
        txt_data_nasc = ""
        txt_email = ""
        txt_cref = ""
        txt_usuario = ""
        txt_senha = ""
        txt_csenha = ""
    Else
        MsgBox ("Registro de professor mantido"), vbInformation + vbOKOnly, "AVISO"
    End If
End Sub
Private Sub btn_salvar_Click()

If IsNull(txt_cpf) Or _
   IsNull(txt_nome) Or _
   IsNull(txt_data_nasc) Or _
   IsNull(txt_cref) Or _
   IsNull(txt_email) Or _
   IsNull(txt_usuario) Or _
   IsNull(txt_senha) Or _
   IsNull(txt_csenha) = True Then
   MsgBox ("Preencha todos os campos!"), vbExclamation + vbOKOnly, "ATENÇÃO"
   Exit Sub
End If

    Set db = CurrentDb
    sql = "select * from tb_prof where cpf='" & txt_cpf & "'"
    Set rs = db.OpenRecordset(sql)
    If rs.EOF = True Then
    validar_leitura
        sql = "insert into tb_prof values ('" & txt_cpf & "', " & _
                                            "'" & txt_nome & "', " & _
                                            "'" & txt_data_nasc & "', " & _
                                            "'" & txt_email & "', " & _
                                            "'" & txt_cref & "', " & _
                                            "'" & txt_usuario & "', " & _
                                            "'" & txt_senha & "', " & _
                                            "'" & txt_csenha & "')"
        db.Execute (UCase(sql))
        MsgBox ("Registro gravado com sucesso!"), vbInformation + vbOKOnly, "AVISO"
        

        txt_cpf = ""
        txt_nome = ""
        txt_data_nasc = ""
        txt_email = ""
        txt_cref = ""
        txt_usuario = ""
        txt_senha = ""
        txt_csenha = ""
      End If

End Sub
Private Sub btn_volta_Click()
DoCmd.Close
DoCmd.OpenForm " inicio ", acNormal
End Sub
Private Sub Form_Load()
DoCmd.Maximize
End Sub

