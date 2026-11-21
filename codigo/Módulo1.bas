Attribute VB_Name = "Módulo1"
Option Compare Database
Public nome_relat As String
Public db As Database 'Variável destinada ao banco de dados
Public rs As Recordset 'Variável destinada as tabelas
Public sql, resp As String 'Variável destinada a query's crud
Function conectar_banco()
Set db = CurrentDb
'MsgBox ("Conexão OK"), vbInformation + vbOKOnly, "AVISO"
End Function
Function validar_leitura()
Set rs = db.OpenRecordset(sql, dbOpenDynaset)
End Function
Function limpar_cadastro()
With Form_cadastro
     .txt_nome.Enabled = True
     .txt_data_nasc.Enabled = True
     .txt_cep.Enabled = True
     .txt_endereco.Enabled = True
     .txt_num_comp.Enabled = True
     .txt_cidade.Enabled = True
     .txt_bairro.Enabled = True
     .txt_uf.Enabled = True
     .txt_fone.Enabled = True
     .txt_email.Enabled = True
     .txt_cpf = ""
     .txt_nome = ""
     .txt_data_nasc = ""
     .txt_endereco = ""
     .txt_cep = ""
     .txt_num_comp = ""
     .txt_bairro = ""
     .txt_cidade = ""
     .txt_uf = ""
     .txt_fone = ""
     .txt_email = ""
     .btn_excluir.Enabled = False
     .txt_cpf.SetFocus
End With
End Function

Function ListaProfessores()
Dim resultado As String
Set db = CurrentDb 'este é o nosso abnco de dados atual
Set rs = db.OpenRecordset("query_profs") 'trazemos nossa query de professores aqui

'variável para armazenar os nomes em sequencia
result = "PROFESSORES REGISTRADOS:" & vbNewLine & vbNewLine & ""

If Not rs.EOF Then 'loop para puxar cada nome
    rs.MoveFirst
    Do Until rs.EOF '"faça enquanto" não chegar ao fim da lista
        result = result & rs!nome & ", " 'concatenando os nomes
        rs.MoveNext
    Loop
End If

If Len(result) > 2 Then
    result = Left(result, Len(result) - 2)
End If

ListaProfessores = result

'as funções abaixo ajudam a liberar memória. Obrigada, chatGPT.
rs.Close
Set rs = Nothing
Set db = Nothing

End Function
