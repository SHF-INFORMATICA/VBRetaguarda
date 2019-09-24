ALTER procedure spCOMANDA  @Acao int, @COMANDA_ID bigint, @CARTAOBARRA_ID bigint, @USUARIO_ID int, @DT_REGISTRO datetime, @SITUACAO nvarchar(10)
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			SET @COMANDA_ID = (select max(COMANDA_ID) from COMANDA) + 1

			IF @COMANDA_ID IS NULL 
				SET @COMANDA_ID = 1
			IF @COMANDA_ID <= 0 
				SET @COMANDA_ID = 1
			IF @SITUACAO = ''
				SET @SITUACAO = NULL

			--INSERT INTO COMANDA (COMANDA_ID,PEDIDO_ID,CARTAOBARRA_ID,USUARIO_ID,DT_REGISTRO,SITUACAO)
			--Values (@COMANDA_ID,@PEDIDO_ID,@CARTAOBARRA_ID,@USUARIO_ID,@DT_REGISTRO,@SITUACAO)

			INSERT INTO COMANDA (COMANDA_ID,CARTAOBARRA_ID,USUARIO_ID,DT_REGISTRO,SITUACAO)
			Values (@COMANDA_ID,@CARTAOBARRA_ID,@USUARIO_ID,@DT_REGISTRO,@SITUACAO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update COMANDA SET SITUACAO=@SITUACAO,DT_REGISTRO=@DT_REGISTRO
			where COMANDA_ID = @COMANDA_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from COMANDA where COMANDA_ID = @COMANDA_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, n�o executado, spCOMANDA',14,1)     
	End  
End 