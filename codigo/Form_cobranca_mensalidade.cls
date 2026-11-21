VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_cobranca_mensalidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private m_AlunoExistenteCPF As String
Private m_UltimoPagamentoID As Long
Private m_CPFRelatorio As String
Private Sub Form_Load()
    Call LimparCampos
    m_AlunoExistenteCPF = ""
    
    m_UltimoPagamentoID = 0
    m_CPFRelatorio = ""
    DoCmd.Maximize

    cmb_turma.RowSourceType = "Table/Query"
    cmb_turma.RowSource = "SELECT DISTINCT turma FROM tb_turmas ORDER BY turma;"
    cmb_turma.BoundColumn = 1
    cmb_turma.ColumnCount = 1
    cmb_turma.ColumnWidths = "5cm"
End Sub

Private Sub LimparCampos()

    txt_cpf.Value = ""
    txt_nome.Value = ""
    txt_mae.Value = ""
    txt_data_nasc.Value = Null
    txt_idade.Value = ""
    cmb_turma.Value = Null
    txt_mes_referencia.Value = Format(Date, "dd/mm/yyyy")
    txt_valor.Value = ""
    txt_data_pgto.Value = Date
    cmb_forma_pagamento.Value = Null
    txt_observacoes.Value = ""
    txt_cpf.SetFocus
    m_AlunoExistenteCPF = ""
   
End Sub
Private Sub btn_retorne3_Click()
    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm "inicio", acNormal
End Sub
Private Sub btn_buscar_aluno_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strCPF As String
    Dim strSQL As String
    
    Set db = CurrentDb
    strCPF = txt_cpf.Value

    If Len(strCPF) = 0 Then
        MsgBox "Por favor, digite o CPF do aluno para buscar.", vbInformation
        txt_cpf.SetFocus
        Exit Sub
    End If

    strSQL = "SELECT CPF, MATRICULA, NOME, DATA_NASC, NOME_MAE, TURMA FROM tb_alunos WHERE CPF = '" & strCPF & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)

    If Not rs.EOF Then
        m_AlunoExistenteCPF = Nz(rs!CPF, "")
        txt_nome.Value = Nz(rs!nome, "")
        txt_mae.Value = Nz(rs!NOME_MAE, "")

        txt_data_nasc.Value = Nz(rs!DATA_NASC, Null)

        If Not IsNull(rs!turma) Then
            cmb_turma.Value = rs!turma
        Else
            cmb_turma.Value = Null
        End If

        If Not IsNull(txt_data_nasc.Value) Then
            txt_idade.Value = CalcularIdade(txt_data_nasc.Value)
        Else
            txt_idade.Value = ""
        End If

        m_CPFRelatorio = strCPF

        MsgBox "Aluno encontrado e dados preenchidos.", vbInformation
    Else
        MsgBox "Aluno não encontrado. Por favor, preencha os dados para um novo cadastro.", vbInformation
        m_AlunoExistenteCPF = ""
        Call LimparCampos
        txt_nome.SetFocus
        m_CPFRelatorio = ""
    End If
End Sub
Private Sub btn_gravar_Click()
    Dim db As DAO.Database
    Dim rsAlunos As DAO.Recordset
    Dim strAlunoCPF As String
    Dim strAlunoNome As String
    Dim boolAlunoRecordsetOpened As Boolean
    Dim strSQLInsertPagamento As String
    Dim varPagamentoID As Variant

    Set db = CurrentDb
    boolAlunoRecordsetOpened = False
    If IsNull(txt_cpf) Or _
       IsNull(txt_nome) Or _
       IsNull(txt_idade) Or _
       IsNull(txt_mae) Or _
       IsNull(txt_mes_referencia) Or _
       IsNull(txt_valor) Or _
       IsNull(txt_data_pgto) Or _
       IsNull(cmb_turma) Or _
       IsNull(cmb_forma_pagamento) Then
    
        MsgBox ("Por favor, preencha todos os campos obrigatórios"), vbInformation, " AVISO"

    End If

    Set rsAlunos = db.OpenRecordset("SELECT * FROM tb_alunos WHERE CPF = '" & txt_cpf.Value & "'", dbOpenDynaset)
    boolAlunoRecordsetOpened = True

    If rsAlunos.EOF Then
        MsgBox ("Aluno não existe"), vbInformation + vbOKOnly, "AVISO"
    Else
        rsAlunos.Edit
        rsAlunos!nome = txt_nome.Value
        rsAlunos!NOME_MAE = txt_mae.Value
        
        If IsDate(txt_data_nasc.Value) Then
            rsAlunos!DATA_NASC = CDate(txt_data_nasc.Value)
        Else
            rsAlunos!DATA_NASC = Null
        End If

        If Not IsNull(cmb_turma.Value) Then
            rsAlunos!turma = cmb_turma.Value
        Else
            rsAlunos!turma = Null
        End If
        rsAlunos.Update
        strAlunoCPF = txt_cpf.Value
        strAlunoNome = txt_nome.Value
    End If
    
    strAlunoCPF = txt_cpf.Value
    strAlunoNome = txt_nome.Value
    m_AlunoExistenteCPF = strAlunoCPF


    Dim sqlCPF As String: sqlCPF = "'" & Replace(strAlunoCPF, "'", "''") & "'"
    Dim sqlNome As String: sqlNome = "'" & Replace(strAlunoNome, "'", "''") & "'"
    Dim sqlMesReferencia As String
    If IsDate(txt_mes_referencia.Value) Then
        sqlMesReferencia = "#" & Format(CDate(txt_mes_referencia.Value), "yyyy\/mm\/dd") & "#"
    Else
        MsgBox "O 'Mês de Referência' digitado não é uma data válida. Por favor, corrija.", vbExclamation
    End If

    Dim sqlValorPago As String: sqlValorPago = Replace(txt_valor.Value, ",", ".")
    Dim sqlDataPagamento As String: sqlDataPagamento = "#" & Format(IIf(IsNull(txt_data_pgto.Value), Date, txt_data_pgto.Value), "yyyy\/mm\/dd") & "#"
    Dim sqlFormaPagamento As String: sqlFormaPagamento = "Null"
    If Not IsNull(cmb_forma_pagamento.Value) Then
        sqlFormaPagamento = "'" & Replace(cmb_forma_pagamento.Value, "'", "''") & "'"
    End If
    Dim sqlObservacoes As String: sqlObservacoes = "Null"
    If Len(txt_observacoes.Value & "") > 0 Then
        sqlObservacoes = "'" & Replace(txt_observacoes.Value, "'", "''") & "'"
    End If
    
    strSQLInsertPagamento = "INSERT INTO tb_Pagamentos (CPF, Nome, Mes_referencia, Valor_pago, Data_pagamento, Forma_pagamento, Observacoes) VALUES (" & _
                            sqlCPF & ", " & _
                            sqlNome & ", " & _
                            sqlMesReferencia & ", " & _
                            sqlValorPago & ", " & _
                            sqlDataPagamento & ", " & _
                            sqlFormaPagamento & ", " & _
                            sqlObservacoes & ");"
    
    db.Execute strSQLInsertPagamento
  
    m_CPFRelatorio = strAlunoCPF

    MsgBox "Dados salvos com sucesso!", vbInformation

    txt_cpf = ""
    txt_nome = ""
    txt_idade = ""
    txt_mae = ""
    txt_valor = ""
    cmb_forma_pagamento = Null
    txt_observacoes = ""
    cmb_turma = Null
    txt_cpf.SetFocus

    Set rsAlunos = Nothing
    Set db = Nothing
    Exit Sub
End Sub
Private Sub txt_data_nasc_AfterUpdate()
    If IsDate(txt_data_nasc.Value) Then
        txt_idade.Value = CalcularIdade(txt_data_nasc.Value)
    Else
        txt_idade.Value = ""
    End If
End Sub

Private Function CalcularIdade(DataNascimento As Date) As Integer
    If IsDate(DataNascimento) Then
        CalcularIdade = DateDiff("yyyy", DataNascimento, Date)
        If Date < DateSerial(Year(Date), Month(DataNascimento), Day(DataNascimento)) Then
            CalcularIdade = CalcularIdade - 1
        End If
    Else
        CalcularIdade = 0
    End If
End Function

Private Sub btn_relatorio_Click()

    If Len(m_CPFRelatorio) > 0 Then
        DoCmd.OpenReport "relat_nf", acViewReport, , "[CPF] = '" & m_CPFRelatorio & "'"
    Else
        MsgBox "Nenhum CPF de aluno disponível para gerar o relatório. Por favor, busque ou grave um aluno primeiro.", vbInformation
    End If

End Sub

Private Sub btn_limpar_Click()
    Call LimparCampos
   
End Sub
