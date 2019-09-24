update pedido set cliente_id = cliente.CLIENTE_ID

--SELECT pedido.cliente_id,cliente.CLIENTE_ID , pedido.CGCCPF,cliente.CGCCPF ,pedido.NOME_CLIENTE ,cliente.NOME 
FROM         PEDIDO INNER JOIN
                      CLIENTE ON PEDIDO.CGCCPF = CLIENTE.CGCCPF
where pedido.cliente_id <> cliente.cliente_id
and pedido.CGCCPF = cliente.CGCCPF 
--order by nome