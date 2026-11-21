Attribute VB_Name = "modulo2"
Public Sub SepararTurmasPorIdadeComHorario()
    Dim db As DAO.Database
    Dim rsAlunos As DAO.Recordset
    Dim dataAtual As Date
    Dim idade As Integer
    Dim turma As String
    Dim diaAula As String
    Dim horarioAula As String     '
    Dim sql As String
    Dim sqlUpdate As String
    Dim contadorTurmaA As Integer
    Dim contadorTurmaB As Integer
    Dim contadorTurmaC As Integer
    Dim contadorTurmaD As Integer
    Dim contadorTurmaE As Integer
    Dim contadorTurmaF As Integer

    contadorTurmaA = 0
    contadorTurmaB = 0
    contadorTurmaC = 0
    contadorTurmaD = 0
    contadorTurmaE = 0
    contadorTurmaF = 0

    dataAtual = Date
    
    Set db = CurrentDb
    
    sql = "SELECT matricula, nome, data_nasc FROM tb_alunos ORDER BY data_nasc"
    Set rsAlunos = db.OpenRecordset(sql)
    
    If rsAlunos.EOF Then
        MsgBox "Não foram encontrados alunos no banco de dados.", vbInformation, "Aviso"
        rsAlunos.Close
        Set rsAlunos = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    rsAlunos.MoveFirst
    Do Until rsAlunos.EOF
        ' Calcula a idade
        idade = DateDiff("yyyy", rsAlunos("data_nasc"), dataAtual)
        
        ' Ajusta a idade se o aniversário deste ano ainda não ocorreu
        If Month(rsAlunos("data_nasc")) > Month(dataAtual) Or _
           (Month(rsAlunos("data_nasc")) = Month(dataAtual) And _
            Day(rsAlunos("data_nasc")) > Day(dataAtual)) Then
            idade = idade - 1
        End If
        
        ' Define a turma, dia e horário com base na idade
        Select Case idade
            Case Is < 4
                turma = "A"
                diaAula = "Segunda e Quarta"
                horarioAula = "09:00 - 10:00"
                contadorTurmaA = contadorTurmaA + 1
            Case 4 To 7
                turma = "B"
                diaAula = "Terça e Quinta"
                horarioAula = "10:30 - 11:30"
                contadorTurmaB = contadorTurmaB + 1
            Case 7 To 10
                turma = "C"
                diaAula = "Segunda e Quarta"
                horarioAula = "14:00 - 15:30"
                contadorTurmaC = contadorTurmaC + 1
            Case 11 To 17
                turma = "D"
                diaAula = "Terça e Quinta"
                horarioAula = "16:00 - 17:30"
                contadorTurmaD = contadorTurmaD + 1
            Case 18 To 59
                turma = "E"
                diaAula = "Segunda e Sexta"
                horarioAula = "19:00 - 20:30"
                contadorTurmaE = contadorTurmaE + 1
            Case Else ' 60 anos ou mais
                turma = "F"
                diaAula = "Quarta e Sábado"
                horarioAula = "10:00 - 11:30"
                contadorTurmaF = contadorTurmaF + 1
        End Select
        
        sqlUpdate = "INSERT INTO tb_turmas (matricula, nome, idade, turma, dia_aula, horario_aula) VALUES (" & _
                    rsAlunos("matricula") & ", '" & _
                    Replace(rsAlunos("nome"), "'", "''") & "', " & _
                    idade & ", '" & turma & "', '" & _
                    diaAula & "', '" & _
                    horarioAula & "')"
        db.Execute sqlUpdate
        
        rsAlunos.MoveNext
    Loop

    rsAlunos.Close
    

    MsgBox "Distribuição de turmas concluída:" & vbCrLf & _
           "Turma A (até 4 anos): " & contadorTurmaA & " alunos" & vbCrLf & _
           "Turma B (4-7 anos): " & contadorTurmaB & " alunos" & vbCrLf & _
           "Turma C (8-10 anos): " & contadorTurmaC & " alunos" & vbCrLf & _
           "Turma D (11-17 anos): " & contadorTurmaD & " alunos" & vbCrLf & _
           "Turma E (18-59 anos): " & contadorTurmaE & " alunos" & vbCrLf & _
           "Turma F (60+ anos): " & contadorTurmaF & " alunos", _
           vbInformation, "Separação de Turmas"
    
    Set rsAlunos = Nothing
    Set db = Nothing
End Sub
Public Sub ProcessarNovoAlunoSemDuplicar()
    Dim db As DAO.Database
    Dim rsAluno As DAO.Recordset
    Dim dataAtual As Date
    Dim idade As Integer
    Dim turma As String
    Dim diaAula As String
    Dim horarioAula As String
    Dim sql As String
    Dim sqlInsert As String
    Dim ultimaMatricula As Long

    
    Set db = CurrentDb
    dataAtual = Date
    
    ultimaMatricula = DMax("matricula", "tb_alunos")
    
    db.Execute "DELETE FROM tb_turmas WHERE matricula = " & ultimaMatricula
    
    sql = "SELECT matricula, nome, data_nasc FROM tb_alunos WHERE matricula = " & ultimaMatricula
    Set rsAluno = db.OpenRecordset(sql)
    
    If Not rsAluno.EOF Then
        idade = DateDiff("yyyy", rsAluno("data_nasc"), dataAtual)
        If Month(rsAluno("data_nasc")) > Month(dataAtual) Or _
           (Month(rsAluno("data_nasc")) = Month(dataAtual) And _
            Day(rsAluno("data_nasc")) > Day(dataAtual)) Then
            idade = idade - 1
        End If
        
        Select Case idade
            Case Is < 4
                turma = "A"
                diaAula = "Segunda e Quarta"
                horarioAula = "09:00 - 10:00"
            Case 4 To 7
                turma = "B"
                diaAula = "Terça e Quinta"
                horarioAula = "10:30 - 11:30"
            Case 7 To 10
                turma = "C"
                diaAula = "Segunda e Quarta"
                horarioAula = "14:00 - 15:30"
            Case 11 To 17
                turma = "D"
                diaAula = "Terça e Quinta"
                horarioAula = "16:00 - 17:30"
            Case 18 To 59
                turma = "E"
                diaAula = "Segunda e Sexta"
                horarioAula = "19:00 - 20:30"
            Case Else
                turma = "F"
                diaAula = "Quarta e Sábado"
                horarioAula = "10:00 - 11:30"
        End Select
        
        sqlInsert = "INSERT INTO tb_turmas (matricula, nome, idade, turma, dia_aula, horario_aula) VALUES (" & _
                    rsAluno("matricula") & ", '" & _
                    Replace(rsAluno("nome"), "'", "''") & "', " & _
                    idade & ", '" & turma & "', '" & _
                    diaAula & "', '" & _
                    horarioAula & "')"
        db.Execute sqlInsert
        
    End If
    
    rsAluno.Close
    Set rsAluno = Nothing
    Set db = Nothing
    Exit Sub
    
    If Not rsAluno Is Nothing Then
        If Not rsAluno.EOF Or Not rsAluno.BOF Then
            rsAluno.Close
        End If
        Set rsAluno = Nothing
    End If
    Set db = Nothing
End Sub
Public Sub AtualizarTurmaNovoAlunoSimples()
    DoEvents
    
    Call ProcessarNovoAlunoSemDuplicar
    
    Exit Sub
End Sub
Public Sub ReconstruirTabelaTurmas()
    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' Esta linha irá deletar TODOS os registros da tb_turmas. É para ser usado quando você quer recriar a tabela do zero.
    db.Execute "DELETE FROM tb_turmas"
    
    ' Reprocessa todos os alunos para popular a tabela de turmas novamente
    Call SepararTurmasPorIdadeComHorario
    
    Set db = Nothing
    MsgBox "Tabela de Turmas reconstruída e atualizada com sucesso!", vbInformation
End Sub
