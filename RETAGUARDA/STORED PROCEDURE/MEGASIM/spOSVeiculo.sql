alter procedure spOSVeiculo @Acao int,
							@VEICULO_ID bigint,
							@PESSOA_ID bigint,
							@COMBUSTIVEL_ID bigint,
							@COR_ID bigint,
							@TIPO_VEICULO_ID bigint,
							@MARCA_ID bigint,
							@PLACA nvarchar(10),
							@DESCRICAO nvarchar(100),
							@MOTOR nvarchar(100),
							@CHASSI nvarchar(100),
							@NUMR_FROTA nvarchar(10),
							@ANO int,
							@MODELO int

as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
		    SET @VEICULO_ID = (select max(VEICULO_ID) from OSVEICULO) + 1

			IF @VEICULO_ID IS NULL 
				SET @VEICULO_ID = 1
			IF @VEICULO_ID <= 0 
				SET @VEICULO_ID = 1

			INSERT INTO OSVEICULO (VEICULO_ID,PESSOA_ID,COMBUSTIVEL_ID,COR_ID,TIPO_VEICULO_ID,
								   MARCA_ID,PLACA,DESCRICAO,MOTOR,CHASSI,NUMR_FROTA,ANO,MODELO)
			Values (@VEICULO_ID,@PESSOA_ID,@COMBUSTIVEL_ID,@COR_ID,@TIPO_VEICULO_ID,
					@MARCA_ID,@PLACA,@DESCRICAO,@MOTOR,@CHASSI,@NUMR_FROTA,@ANO,@MODELO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update OSVEICULO SET PESSOA_ID=@PESSOA_ID,COMBUSTIVEL_ID=@COMBUSTIVEL_ID,COR_ID=@COR_ID,TIPO_VEICULO_ID=@TIPO_VEICULO_ID,MARCA_ID=@MARCA_ID,
								 PLACA=@PLACA,DESCRICAO=@DESCRICAO,MOTOR=@MOTOR,CHASSI=@CHASSI,NUMR_FROTA=@NUMR_FROTA,ANO=@ANO,MODELO=@MODELO
			where VEICULO_ID = @VEICULO_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 3) begin
		BEGIN TRANSACTION
			delete from OSVEICULO where VEICULO_ID = @VEICULO_ID

		IF @@ERROR <> 0
		   ROLLBACK
		   ELSE
		      COMMIT
	End
	begin raiserror('Erro, não executado, spOSVeiculo',14,1)
	End
End