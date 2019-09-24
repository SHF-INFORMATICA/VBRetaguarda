alter procedure spCaixaItem @Acao int,@CAIXADIA_ID bigint,@CAIXADIAITEM_ID bigint,@FORMAPAGTO_ID bigint,@VALOR float,@TIPO char(1),@HISTORICO nvarchar(max)
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION

			INSERT INTO CAIXADIAITEM (CAIXADIA_ID,CAIXADIAITEM_ID,FORMAPAGTO_ID,VALOR,TIPO,HISTORICO)
			Values (@CAIXADIA_ID,@CAIXADIAITEM_ID,@FORMAPAGTO_ID,@VALOR,@TIPO,@HISTORICO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update CAIXADIAITEM SET HISTORICO=@HISTORICO,
									VALOR=@VALOR,
									FORMAPAGTO_ID=@FORMAPAGTO_ID,
									TIPO=@TIPO
			where CAIXADIA_ID = @CAIXADIA_ID
			AND   CAIXADIAITEM_ID = @CAIXADIAITEM_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from CAIXADIAITEM where CAIXADIA_ID = @CAIXADIA_ID
								 AND   CAIXADIAITEM_ID = @CAIXADIAITEM_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spCAIXAITEM',14,1)     
	End  
End 