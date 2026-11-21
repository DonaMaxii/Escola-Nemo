VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_cadastro_alunos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub btn_excluir_Click()
resp = MsgBox("Deseja realmente excluir o Aluno?", vbQuestion + vbYesNo, "ATENÇÃO")
    If resp = vbYes Then
        sql = " delete * from tb_alunos where CPF= '" & txt_cpf & "' "
        db.Execute (sql)
        MsgBox ("Cliente excluído com sucesso!!!"), vbInformation + vbOKOnly, "AVISO"
    End If
    limpar_cadastro
End Sub
Private Sub btn_gravar_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dataAtual As Date
    Dim idade As Integer
    Dim turmaCalculada As String
    Dim diaAulaCalculado As String
    Dim horarioAulaCalculado As String
    Dim matriculaAluno As Long

    ' Verificação de campos obrigatórios
    If IsNull(txt_cpf) Or _
       IsNull(txt_nome) Or _
       IsNull(txt_data_nasc) Or _
       IsNull(txt_mae) Or _
       IsNull(txt_cep) Or _
       IsNull(txt_num_comp) Or _
       IsNull(txt_fone) Or _
       IsNull(txt_email) = True Then

        MsgBox ("Preencha todos os campos!"), vbExclamation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

    Set db = CurrentDb
    sql = "SELECT * FROM tb_alunos WHERE cpf='" & txt_cpf & "'"
    Set rs = db.OpenRecordset(sql)

    If rs.EOF = True Then
        ' Inserir novo aluno
        sql = "INSERT INTO tb_alunos (cpf, nome, data_nasc, nome_mae, cep, endereco, comp, bairro, cidade, uf, celular, email) VALUES ('" & _
              txt_cpf & "', '" & _
              txt_nome & "', " & _
              "#" & Format(CDate(txt_data_nasc), "yyyy\/mm\/dd") & "#, '" & _
              txt_mae & "', '" & _
              txt_cep & "', '" & _
              txt_endereco & "', '" & _
              txt_num_comp & "', '" & _
              txt_bairro & "', '" & _
              txt_cidade & "', '" & _
              txt_uf & "', '" & _
              txt_fone & "', '" & _
              txt_email & "')"

        db.Execute (UCase(sql))

        rs.Close
        Set rs = Nothing

        Call ProcessarNovoAlunoSemDuplicar

        MsgBox ("Registro gravado com sucesso!"), vbInformation + vbOKOnly, "AVISO"

    Else
        sql = "UPDATE tb_alunos SET nome='" & txt_nome & "', " & _
              "data_nasc= #" & Format(CDate(txt_data_nasc), "yyyy\/mm\/dd") & "#, " & _
              "nome_mae='" & txt_mae & "', " & _
              "cep='" & txt_cep & "', " & _
              "endereco='" & txt_endereco & "', " & _
              "comp='" & txt_num_comp & "', " & _
              "bairro='" & txt_bairro & "', " & _
              "cidade='" & txt_cidade & "', " & _
              "uf='" & txt_uf & "', " & _
              "celular='" & txt_fone & "', " & _
              "email='" & txt_email & "' WHERE cpf='" & txt_cpf & "'"
        db.Execute (UCase(sql))

        dataAtual = Date
        
        matriculaAluno = DLookup("matricula", "tb_alunos", "cpf = '" & txt_cpf & "'")

        idade = DateDiff("yyyy", CDate(txt_data_nasc), dataAtual)
        If Month(CDate(txt_data_nasc)) > Month(dataAtual) Or _
           (Month(CDate(txt_data_nasc)) = Month(dataAtual) And _
            Day(CDate(txt_data_nasc)) > Day(dataAtual)) Then
            idade = idade - 1
        End If

        Select Case idade
            Case Is < 4
                turmaCalculada = "A"
                diaAulaCalculado = "Segunda e Quarta"
                horarioAulaCalculado = "09:00 - 10:00"
            Case 4 To 7
                turmaCalculada = "B"
                diaAulaCalculado = "Terça e Quinta"
                horarioAulaCalculado = "10:30 - 11:30"
            Case 7 To 10
                turmaCalculada = "C"
                diaAulaCalculado = "Segunda e Quarta"
                horarioAulaCalculado = "14:00 - 15:30"
            Case 11 To 17
                turmaCalculada = "D"
                diaAulaCalculado = "Terça e Quinta"
                horarioAulaCalculado = "16:00 - 17:30"
            Case 18 To 59
                turmaCalculada = "E"
                diaAulaCalculado = "Segunda e Sexta"
                horarioAulaCalculado = "19:00 - 20:30"
            Case Else
                turmaCalculada = "F"
                diaAulaCalculado = "Quarta e Sábado"
                horarioAulaCalculado = "10:00 - 11:30"
        End Select

        db.Execute "DELETE FROM tb_turmas WHERE matricula = " & matriculaAluno
        
        sql = "INSERT INTO tb_turmas (matricula, nome, idade, turma, dia_aula, horario_aula) VALUES (" & _
              matriculaAluno & ", '" & _
              Replace(txt_nome, "'", "''") & "', " & _
              idade & ", '" & turmaCalculada & "', '" & _
              diaAulaCalculado & "', '" & _
              horarioAulaCalculado & "')"
        db.Execute (UCase(sql))

        rs.Close
        Set rs = Nothing

        MsgBox ("Dados alterados com sucesso!"), vbInformation + vbOKOnly, "AVISO"
    End If

    txt_cpf = ""
    txt_nome = ""
    txt_data_nasc = ""
    txt_mae = ""
    txt_cep = ""
    txt_endereco = ""
    txt_num_comp = ""
    txt_bairro = ""
    txt_cidade = ""
    txt_uf = ""
    txt_fone = ""
    txt_email = ""

    Set db = Nothing
    Exit Sub

End Sub
Private Sub btn_relatorios_Click()
DoCmd.Close
DoCmd.OpenForm ("relatorios"), acNormal
End Sub
Private Sub btn_retorne1_Click()
DoCmd.Close
DoCmd.OpenForm " inicio ", acNormal
End Sub
Private Sub btn_turmas_Click()
DoCmd.OpenForm "aviso", acNormal
End Sub
Private Sub Form_Load()
conectar_banco
DoCmd.Maximize
End Sub
Private Sub txt_cep_LostFocus()
sql = "select * from tb_cep where cep = '" & txt_cep & "' "
validar_leitura
If rs.EOF = False Then 'Se existir o cep na tabela
    txt_endereco = rs.Fields(1).Value
    txt_cidade = rs.Fields(2).Value
    txt_bairro = rs.Fields(3).Value
    txt_uf = rs.Fields(4).Value
    txt_num_comp.SetFocus
    txt_endereco.Enabled = False
    txt_cidade.Enabled = False
    txt_bairro.Enabled = False
    txt_uf.Enabled = False
Else
    MsgBox ("CEP não localizado!"), vbInformation + vbOKOnly, "AVISO"
End If
End Sub
Private Sub txt_cpf_DblClick(Cancel As Integer)
limpar_cadastro
End Sub
Private Sub txt_cpf_LostFocus()
sql = "select * from tb_alunos where cpf='" & txt_cpf & "'"
validar_leitura
If rs.EOF = False Then
        txt_nome = rs.Fields(2).Value
        txt_data_nasc = rs.Fields(3).Value
        txt_mae = rs.Fields(4).Value
        txt_cep = rs.Fields(5).Value
        txt_endereco = rs.Fields(6).Value
        txt_num_comp = rs.Fields(7).Value
        txt_bairro = rs.Fields(8).Value
        txt_cidade = rs.Fields(9).Value
        txt_uf = rs.Fields(10).Value
        txt_fone = rs.Fields(11).Value
        txt_email = rs.Fields(12).Value
        txt_nome.Enabled = False
        txt_data_nasc.Enabled = False
        txt_mae.Enabled = False
        txt_cep.Enabled = False
        txt_endereco.Enabled = False
        txt_num_comp.Enabled = False
        txt_bairro.Enabled = False
        txt_cidade.Enabled = False
        txt_uf.Enabled = False
        txt_fone.Enabled = False
        txt_email.Enabled = False
        btn_excluir.Enabled = True
        resp = MsgBox("Deseja alterar os dados?", vbQuestion + vbYesNo, "ATENÇÃO")
        If resp = vbYes Then
           txt_nome.Enabled = True
           txt_data_nasc.Enabled = True
           txt_mae.Enabled = True
           txt_cep.Enabled = True
           txt_num_comp.Enabled = True
           txt_fone.Enabled = True
           txt_email.Enabled = True
        Else
            txt_cpf = ""
            txt_nome = ""
            txt_data_nasc = ""
            txt_mae = ""
            txt_cep = ""
            txt_endereco = ""
            txt_num_comp = ""
            txt_bairro = ""
            txt_cidade = ""
            txt_uf = ""
            txt_fone = ""
            txt_email = ""
        End If
End If
End Sub
