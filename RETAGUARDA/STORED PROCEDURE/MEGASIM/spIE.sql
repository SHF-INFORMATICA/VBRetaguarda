ALTER procedure spIE @Acao int, @IE_ID bigint, @PESSOA_ID BIGINT, @NUMR_IE nvarchar(MAX), @ENDERECO_ID BIGINT
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			SET @IE_ID = (select max(IE_id) from IE) + 1 

			IF @IE_ID IS NULL 
				SET @IE_ID = 1
			IF @IE_ID <= 0 
				SET @IE_ID = 1

			INSERT INTO IE (IE_ID, PESSOA_ID, NUMR_IE, ENDERECO_ID)
			Values (@IE_ID, @PESSOA_ID, @NUMR_IE, @ENDERECO_ID)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update IE SET NUMR_IE = @NUMR_IE, ENDERECO_ID = @ENDERECO_ID
			where IE_id = @IE_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from IE where IE_id = @IE_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spIE',14,1)     
	End  
End 