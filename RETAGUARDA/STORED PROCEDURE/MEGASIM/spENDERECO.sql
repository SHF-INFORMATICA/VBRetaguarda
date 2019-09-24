ALTER procedure spENDERECO  @Acao int, @ENDERECO_ID bigint, @PESSOA_ID bigint, @CEP_ID nvarchar(8), @RUA nvarchar(50), @BAIRRO nvarchar(50), @COMPLEMENTO nvarchar(50), @TIPO nvarchar(1), @NUMERO nvarchar(50)
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION
			SET @ENDERECO_ID = (select max(ENDERECO_ID) from ENDERECO) + 1

			IF @ENDERECO_ID IS NULL 
				SET @ENDERECO_ID = 1
			IF @ENDERECO_ID <= 0 
				SET @ENDERECO_ID = 1

			INSERT INTO ENDERECO (ENDERECO_ID, PESSOA_ID, CEP_ID, RUA, BAIRRO, COMPLEMENTO, TIPO, NUMERO)
			Values (@ENDERECO_ID, @PESSOA_ID, @CEP_ID, @RUA, @BAIRRO, @COMPLEMENTO, @TIPO, @NUMERO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update ENDERECO SET CEP_ID=@CEP_ID, RUA=@RUA, BAIRRO=@BAIRRO, COMPLEMENTO=@COMPLEMENTO, TIPO=@TIPO, NUMERO=@NUMERO
			where ENDERECO_ID = @ENDERECO_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from ENDERECO where ENDERECO_ID = @ENDERECO_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	begin raiserror('Erro, não executado, spENDERECO',14,1)     
	End
End