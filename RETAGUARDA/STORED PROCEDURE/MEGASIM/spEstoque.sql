ALTER procedure spMovimentaEstoque @Acao int,@Estoque_ID bigint,@Estab_ID int,@Produto_ID float,@QTDE FLOAT
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			SET @Estoque_ID = (select max(estoque_id) from ESTOQUE) + 1 

			IF @Estoque_ID IS NULL 
				SET @Estoque_ID = 1
			IF @Estoque_ID <= 0 
				SET @Estoque_ID = 1
			
			INSERT INTO ESTOQUE (estoque_id, ESTABELECIMENTO_ID, PRODUTO_ID, QTDE_ESTOQUE)
			Values (@Estoque_ID,@Estab_ID,@Produto_ID,@QTDE)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update ESTOQUE SET @QTDE = @QTDE
			where estoque_id = @Estoque_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin
		BEGIN TRANSACTION
			delete from ESTOQUE where estoque_id = @Estoque_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spESTOQUE',14,1)     
	End  
End 