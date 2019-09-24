ALTER procedure spPedidoItem @Acao int, @PEDIDO_ID bigint, @SEQ_ID bigint, @PRODUTO_ID bigint, @QTD_PEDIDA float, @VALOR_ITEM float, @PERC_DESC float,
										@CFOP_ID nvarchar(10), @STRIBUTARIA nvarchar(4), @VLRBASEICMS float, @PERCICMS real, @VLRICMS float, 
										@VLRBASEICMSSUBST float, @PERCICMSSUBST real, @VLRICMSSUBST float, @PERCREDUCAOICMS real, @PERCIVA real, 
										@PERC_IPI real, @VLR_IPI float, @VALOR_DESCONTO float, @STATUS char(4), @PRECO_CUSTO float, @TIPO_REG char(2),
										@PESO_ITEM float, @QTDE_BALANCA float, @USU_ATENDE int, @ALTURA float, @LARGURA float
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION

			INSERT INTO PedidoItem (PEDIDO_ID,SEQ_ID,PRODUTO_ID,QTD_PEDIDA,VALOR_ITEM,PERC_DESC,
									CFOP_ID,STRIBUTARIA,VLRBASEICMS,PERCICMS,VLRICMS,VLRBASEICMSSUBST,
									PERCICMSSUBST,VLRICMSSUBST,PERCREDUCAOICMS,PERCIVA,PERC_IPI,
									VLR_IPI,VALOR_DESCONTO,STATUS,PRECO_CUSTO,TIPO_REG,PESO_ITEM,
									QTDE_BALANCA,USU_ATENDE,ALTURA,LARGURA)
			Values (@PEDIDO_ID,@SEQ_ID,@PRODUTO_ID,@QTD_PEDIDA,@VALOR_ITEM,@PERC_DESC,@CFOP_ID,
					@STRIBUTARIA,@VLRBASEICMS,@PERCICMS,@VLRICMS,@VLRBASEICMSSUBST,@PERCICMSSUBST,
					@VLRICMSSUBST,@PERCREDUCAOICMS,@PERCIVA,@PERC_IPI,@VLR_IPI,@VALOR_DESCONTO,
					@STATUS,@PRECO_CUSTO,@TIPO_REG,@PESO_ITEM,@QTDE_BALANCA,@USU_ATENDE,@ALTURA,@LARGURA)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update PedidoItem SET QTD_PEDIDA=@QTD_PEDIDA,VALOR_ITEM=@VALOR_ITEM,PERC_DESC=@PERC_DESC,CFOP_ID=@CFOP_ID,
					STRIBUTARIA=@STRIBUTARIA,VLRBASEICMS=@VLRBASEICMS,PERCICMS=@PERCICMS,VLRICMS=@VLRICMS,
					VLRBASEICMSSUBST=@VLRBASEICMSSUBST,PERCICMSSUBST=@PERCICMSSUBST,VLRICMSSUBST=@VLRICMSSUBST,
					PERCREDUCAOICMS=@PERCREDUCAOICMS,PERCIVA=@PERCIVA,PERC_IPI=@PERC_IPI,VLR_IPI=@VLR_IPI,
					VALOR_DESCONTO=@VALOR_DESCONTO,STATUS=@STATUS,PRECO_CUSTO=@PRECO_CUSTO,TIPO_REG=@TIPO_REG,
					PESO_ITEM=@PESO_ITEM,QTDE_BALANCA=@QTDE_BALANCA,USU_ATENDE=@USU_ATENDE,ALTURA=@ALTURA,LARGURA=@LARGURA
			where PEDIDO_ID = @PEDIDO_ID
			AND   SEQ_ID = @SEQ_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			if @SEQ_ID > 0 
			delete from PedidoItem where PEDIDO_ID = @PEDIDO_ID
								   AND   SEQ_ID = @SEQ_ID
			else 
				delete from PedidoItem where PEDIDO_ID = @PEDIDO_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, n�o executado, spPedidoItem',14,1)     
	End  
End 