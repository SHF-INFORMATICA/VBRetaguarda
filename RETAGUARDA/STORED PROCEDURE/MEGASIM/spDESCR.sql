ALTER procedure spDESCR @Acao int,@CODIGO bigint,@TIPO nvarchar(2),@DESCRICAO nvarchar(100)
as  begin
	if (@Acao = 1) begin			--incluir registro
		BEGIN TRANSACTION

			INSERT INTO DESCR (CODIGO,TIPO,DESCRICAO)
			Values (@CODIGO, @TIPO, @DESCRICAO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 2) begin       --alterar registro
		BEGIN TRANSACTION

			update DESCR SET DESCRICAO = @DESCRICAO
			where CODIGO = @CODIGO

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 3) begin       --excluir registro
		BEGIN TRANSACTION

			delete from DESCR where CODIGO = @CODIGO and TIPO = @TIPO

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	begin raiserror('Erro, não executado, spDESCR',14,1)     
	End
End