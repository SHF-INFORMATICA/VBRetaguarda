UPDATE LANCAMENTO SET NOME_PESSOA = left(PESSOA.DESCRICAO,30)
FROM LANCAMENTO INNER JOIN PESSOA
ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID 

SELECT top 10 * FROM LANCAMENTO
select * from pessoa where pessoa_id = 4991

--UPDATE PEDIDO SET PEDIDO.TIPOVENDA_ID = PEDIDOVACA.TIPOVENDA_ID
--FROM PEDIDOVACA 
--WHERE PEDIDO.PEDIDO_ID = PEDIDOVACA.PEDIDO_ID