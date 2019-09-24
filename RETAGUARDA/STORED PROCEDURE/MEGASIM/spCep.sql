ALTER procedure [dbo].[spCEP]  @Acao int,@CEP_ID nvarchar(8),@CIDADE nvarchar(50),@UF nvarchar(2),@IBGE_ID int
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION

			INSERT INTO CEP (CEP_ID, CIDADE, UF, IBGE_ID)
			Values (@CEP_ID, @CIDADE, @UF, @IBGE_ID)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update CEP SET CIDADE = @CIDADE, UF = @UF, IBGE_ID = @IBGE_ID
			where CEP_id = @CEP_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from CEP where CEP_id = @CEP_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	begin raiserror('Erro, não executado, spCEP',14,1)     
	End
End