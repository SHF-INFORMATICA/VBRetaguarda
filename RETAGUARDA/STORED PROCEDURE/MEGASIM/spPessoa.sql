--create procedure spPESSOA  @Acao int,@PESSOA_ID bigint,@CNPJCPF nvarchar(14),@DESCRICAO nvarchar(MAX),@RAZAO nvarchar(MAX),@SITUACAO nvarchar(1)  
alter procedure spPESSOA  @Acao int,@PESSOA_ID bigint,@CNPJCPF nvarchar(14),@DESCRICAO nvarchar(MAX),@RAZAO nvarchar(MAX),@SITUACAO nvarchar(1)
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			SET @PESSOA_ID = (select max(pessoa_id) from pessoa) + 1

			IF @PESSOA_ID IS NULL 
				SET @PESSOA_ID = 1
			IF @PESSOA_ID <= 0 
				SET @PESSOA_ID = 1

			INSERT INTO PESSOA (PESSOA_ID, CnpjCpf, Descricao, RAZAO, DATA_CAD, SITUACAO)          
			Values (@PESSOA_ID,@CNPJCPF,@DESCRICAO,@RAZAO,getdate(),@SITUACAO)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update PESSOA SET Descricao = @Descricao, RAZAO = @RAZAO, SITUACAO = @SITUACAO 
			where pessoa_id = @PESSOA_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from PESSOA where pessoa_id = @PESSOA_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spPessoa',14,1)     
	End  
End 