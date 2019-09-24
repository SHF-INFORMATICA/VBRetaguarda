ALTER procedure spOSOBS @Acao int, @OSOBS_ID bigint, @OS_ID BIGINT, @OBS nvarchar(MAX)
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			SET @OSOBS_ID = (select max(OSOBS_ID) from OSOBS) + 1 

			IF @OSOBS_ID IS NULL 
				SET @OSOBS_ID = 1
			IF @OSOBS_ID <= 0 
				SET @OSOBS_ID = 1

			INSERT INTO OSOBS (OSOBS_ID, OS_ID, OBS, DT_CAD)
			Values (@OSOBS_ID, @OS_ID, @OBS, getdate())

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update OSOBS SET OBS = @OBS
			where OSOBS_ID = @OSOBS_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from OSOBS where OSOBS_ID = @OSOBS_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spOSOBS',14,1)     
	End  
End 