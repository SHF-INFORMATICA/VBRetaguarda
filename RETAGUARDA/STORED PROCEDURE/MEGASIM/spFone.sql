--ALTER procedure [dbo].[spFONE]  @Acao int,@FONE_ID_N bigint,@NUMR_FONE nvarchar(MAX),@PESSOA_ID BIGINT,@DDD INT,@LOCAL nvarchar(MAX)
alter procedure [dbo].[spFONE]  @Acao int, @FONE_ID_N bigint, @NUMR_FONE nvarchar(MAX), @PESSOA_ID BIGINT, @DDD INT, @LOCAL nvarchar(MAX)
as  begin     
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION

			SET @FONE_ID_N = (select max(FONE_id) from FONE) + 1 

			IF @FONE_ID_N IS NULL 
				SET @FONE_ID_N = 1
			IF @FONE_ID_N <= 0 
				SET @FONE_ID_N = 1
		
			INSERT INTO FONE (FONE_ID, NUMERO, PESSOA_ID, DDD, LOCAL)          
			Values (@FONE_ID_N,@NUMR_FONE,@PESSOA_ID,@DDD,@LOCAL)     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update FONE SET NUMERO = @NUMR_FONE, DDD = @DDD
			where FONE_id = @FONE_ID_N     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from FONE where FONE_id = @FONE_ID_N     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spFONE',14,1)     
	End  
End 