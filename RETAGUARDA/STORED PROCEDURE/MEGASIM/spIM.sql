ALTER procedure spIM @Acao int, @IM_ID bigint, @PESSOA_ID BIGINT, @NUMR_IM nvarchar(MAX), @ENDERECO_ID BIGINT
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			SET @IM_ID = (select max(IM_id) from IM) + 1 

			IF @IM_ID IS NULL 
				SET @IM_ID = 1
			IF @IM_ID <= 0 
				SET @IM_ID = 1

			INSERT INTO IM (IM_ID, PESSOA_ID, NUMR_IM, ENDERECO_ID)
			Values (@IM_ID, @PESSOA_ID, @NUMR_IM, @ENDERECO_ID)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update IM SET NUMR_IM = @NUMR_IM, ENDERECO_ID = @ENDERECO_ID
			where IM_id = @IM_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION
			
			delete from IM where IM_id = @IM_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spIM',14,1)     
	End  
End 