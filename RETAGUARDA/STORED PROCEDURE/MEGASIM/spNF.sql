ALTER procedure spNF @Acao int, @NF_ID int, @PESSOA_ID bigint, @TRANSP_ID bigint, @NF_TIPO nvarchar(2),
					  @NUMR_NOTA int, @SERIE_NOTA nvarchar(6), @DT_EMISSAO datetime, @DT_ENTRASAI datetime, @STATUS nvarchar(50),
					  @DT_CANCELA datetime, @QTD_VOLUME float, @PESO_BRUTO float, @PESO_LIQUIDO float,
					  @NUMR_REQ_DEV int, @ESTABELECIMENTO_ID int, @indPres int, @idDest int, @MODELO_DOC nvarchar(3)

as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			SET @NF_ID = (select max(NF_ID) from NF) + 1

			IF @NF_ID IS NULL 
				SET @NF_ID = 1
			IF @NF_ID <= 0 
				SET @NF_ID = 1
			IF @NF_TIPO = ''
				SET @NF_TIPO = NULL

			INSERT INTO NF (NF_ID,PESSOA_ID,TRANSP_ID,NF_TIPO,
							NUMR_NOTA,SERIE_NOTA,DT_EMISSAO,DT_ENTRASAI,STATUS,
							DT_CANCELA,QTD_VOLUME,PESO_BRUTO,PESO_LIQUIDO,
							NUMR_REQ_DEV,ESTABELECIMENTO_ID,indPres,idDest,MODELO_DOC)
			Values (@NF_ID,@PESSOA_ID,@TRANSP_ID,@NF_TIPO,
					@NUMR_NOTA,@SERIE_NOTA,@DT_EMISSAO,@DT_ENTRASAI,@STATUS,
					@DT_CANCELA,@QTD_VOLUME,@PESO_BRUTO,@PESO_LIQUIDO,
					@NUMR_REQ_DEV,@ESTABELECIMENTO_ID,@indPres,@idDest,@MODELO_DOC)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update NF SET 
					PESSOA_ID=@PESSOA_ID,TRANSP_ID=@TRANSP_ID,NF_TIPO=@NF_TIPO,
					NUMR_NOTA=@NUMR_NOTA,SERIE_NOTA=@SERIE_NOTA,DT_EMISSAO=@DT_EMISSAO,DT_ENTRASAI=@DT_ENTRASAI,
					STATUS=@STATUS,DT_CANCELA=@DT_CANCELA,QTD_VOLUME=@QTD_VOLUME,
					PESO_BRUTO=@PESO_BRUTO,PESO_LIQUIDO=@PESO_LIQUIDO,NUMR_REQ_DEV=@NUMR_REQ_DEV,
					ESTABELECIMENTO_ID=@ESTABELECIMENTO_ID,indPres=@indPres,idDest=@idDest,MODELO_DOC=@MODELO_DOC
			where NF_ID = @NF_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from NF where NF_ID = @NF_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spNF',14,1)     
	End  
End 