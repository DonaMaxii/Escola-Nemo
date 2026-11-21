SELECT *
FROM tb_alunos
WHERE year(data_nasc) between [Digite o ano inicial de nascimento] and [digite o ano final de nascimento]
ORDER BY data_nasc;
