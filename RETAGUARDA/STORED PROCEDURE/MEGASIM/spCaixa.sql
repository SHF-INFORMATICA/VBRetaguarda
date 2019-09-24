alter procedure spCAIXA  @Acao int,@CAIXADIA_ID bigint,@USUARIO_ID bigint,@ESTABELECIMENTO_ID bigint,@NUMERO_CAIXA_CPU bigint,@DT_ABERTURA datetime,@DT_FECHAMENTO datetime
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			SET @CAIXADIA_ID = (select max(CAIXADIA_ID) from CAIXADIA) + 1

			IF @CAIXADIA_ID IS NULL 
				SET @CAIXADIA_ID = 1
			IF @CAIXADIA_ID <= 0 
				SET @CAIXADIA_ID = 1
			IF @DT_FECHAMENTO = ''
				SET @DT_FECHAMENTO = NULL

			INSERT INTO CAIXADIA (CAIXADIA_ID,USUARIO_ID,ESTABELECIMENTO_ID,NUMERO_CAIXA_CPU,DT_ABERTURA,DT_FECHAMENTO)
			Values (@CAIXADIA_ID,@USUARIO_ID,@ESTABELECIMENTO_ID,@NUMERO_CAIXA_CPU,@DT_ABERTURA,@DT_FECHAMENTO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update CAIXADIA SET DT_FECHAMENTO=@DT_FECHAMENTO
			where CAIXADIA_ID = @CAIXADIA_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from CAIXADIA where CAIXADIA_ID = @CAIXADIA_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spCAIXA',14,1)     
	End  
End 