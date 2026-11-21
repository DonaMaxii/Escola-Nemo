SELECT *
FROM tb_cep
WHERE cidade=[Digite a cidade desejada] and
endereco like 'AV*'
ORDER BY endereco;
