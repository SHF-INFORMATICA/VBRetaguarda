select produto_id, codg_produto, qtde, qtde_retido from PRODUTO where codg_produto = '12101'
SELECT sum(QTD_PEDIDA) as SAIDA FROM PEDIDO INNER JOIN PEDIDOITEM ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID where produto_id = 95 AND PEDIDO.STATUS IN (3,5)
SELECT sum(QTD_entrada) AS ENTRADA_NOTA FROM NOTAENTRADA INNER JOIN NOTAENTRADAITEM ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID where produto_id = 95 AND NOTAENTRADA.STATUS =  'E'
select SUM(QTD_PRIMEIRA) AS ENTRA_INVENT from inventario where codg_prod = '12101' AND STATUS = 'F'