alter procedure spOSTERMO @Acao int, @OSTERMO_ID bigint, @OS_ID BIGINT, @OSTERMOOBS nvarchar(MAX)
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			SET @OSTERMO_ID = (select max(OSTERMO_ID) from OSTERMO) + 1 

			IF @OSTERMO_ID IS NULL 
				SET @OSTERMO_ID = 1
			IF @OSTERMO_ID <= 0 
				SET @OSTERMO_ID = 1

			INSERT INTO OSTERMO (OSTERMO_ID, OS_ID, OSTERMOOBS)
			Values (@OSTERMO_ID, @OS_ID, @OSTERMOOBS)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update OSTERMO SET OSTERMOOBS = @OSTERMOOBS
			where OSTERMO_ID = @OSTERMO_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from OSTERMO where OSTERMO_ID = @OSTERMO_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spOSTERMO',14,1)     
	End  
End 