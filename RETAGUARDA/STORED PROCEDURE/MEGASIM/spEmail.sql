alter procedure [dbo].[spEMAIL]  @Acao int,@EMAIL_ID bigint,@EMAIL nvarchar(MAX),@PESSOA_ID BIGINT
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			SET @EMAIL_ID = (select max(EMAIL_id) from EMAIL) + 1 

			IF @EMAIL_ID IS NULL 
				SET @EMAIL_ID = 1
			IF @EMAIL_ID <= 0 
				SET @EMAIL_ID = 1

			INSERT INTO EMAIL (EMAIL_ID, EMAIL, PESSOA_ID)          
			Values (@EMAIL_ID,@EMAIL,@PESSOA_ID)     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update EMAIL SET EMAIL = @EMAIL
			where EMAIL_id = @EMAIL_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from EMAIL where EMAIL_id = @EMAIL_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spEMAIL',14,1)     
	End  
End 