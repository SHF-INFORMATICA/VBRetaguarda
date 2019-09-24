ALTER procedure spPedidoComanda  @Acao int, @PEDIDO_ID bigint, @CARTAOBARRA_ID BIGINT, @SEQ_COMANDA_ID BIGINT, @SEQ_PEDIDO_ID BIGINT
as  begin     
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			INSERT INTO PEDIDOCOMANDA (PEDIDO_ID, CARTAOBARRA_ID, SEQ_COMANDA_ID, SEQ_PEDIDO_ID)
			Values (@PEDIDO_ID,@CARTAOBARRA_ID,@SEQ_COMANDA_ID,@SEQ_PEDIDO_ID)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update PEDIDOCOMANDA SET 
				PEDIDO_ID = @PEDIDO_ID, CARTAOBARRA_ID=@CARTAOBARRA_ID, SEQ_COMANDA_ID=@SEQ_COMANDA_ID,SEQ_PEDIDO_ID=@SEQ_PEDIDO_ID
			where PEDIDO_ID = @PEDIDO_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from PEDIDOCOMANDA where PEDIDO_ID = @PEDIDO_ID
			delete from PEDIDOCOMANDA where CARTAOBARRA_ID = @CARTAOBARRA_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spPEDIDOCOMANDA',14,1)     
	End  
End 