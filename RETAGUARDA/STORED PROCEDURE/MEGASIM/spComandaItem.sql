ALTER procedure spComandaItem @Acao int, @COMANDA_ID bigint, @SEQ_ID bigint, @PRODUTO_ID bigint, @QTDE float, @VALOR_ITEM float, @SITUACAO nvarchar(10), @USUARIO_ID int
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION

			INSERT INTO COMANDAITEM (COMANDA_ID,SEQ_ID,PRODUTO_ID,QTDE,VALOR_ITEM,SITUACAO,USUARIO_ID)
			Values (@COMANDA_ID,@SEQ_ID,@PRODUTO_ID,@QTDE,@VALOR_ITEM,@SITUACAO,@USUARIO_ID)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update COMANDAITEM SET	SITUACAO=@SITUACAO,
									QTDE=@QTDE,
									PRODUTO_ID=@PRODUTO_ID,
									VALOR_ITEM=@VALOR_ITEM
			where COMANDA_ID = @COMANDA_ID
			AND   SEQ_ID = @SEQ_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			if @SEQ_ID > 0 
			delete from COMANDAITEM where COMANDA_ID = @COMANDA_ID
									AND   SEQ_ID = @SEQ_ID
			else 
				delete from COMANDAITEM where COMANDA_ID = @COMANDA_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spComandaITEM',14,1)     
	End  
End 