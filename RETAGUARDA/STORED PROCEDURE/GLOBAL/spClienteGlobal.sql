ALTER procedure spClienteGlobal @Acao int, 
								@A1_FILIAL varchar(2), @A1_COD varchar(6), @A1_LOJA varchar(2), @A1_NOME varchar(70), 
								@A1_PESSOA varchar(1), @A1_NREDUZ varchar(45), @A1_TIPO varchar(1), @A1_END varchar(60),
								@A1_NUMERO char(60), @A1_MUN varchar(60), @A1_EST varchar(2), @A1_BAIRRO varchar(60), 
								@A1_ESTADO varchar(20), @A1_CEP varchar(8), @A1_DDI varchar(6), @A1_DDD varchar(3),
								@A1_TEL varchar(15), @A1_TELEX varchar(10), @A1_FAX varchar(15), @A1_ENDENTNR char(60),
								@A1_ENDENT varchar(60), @A1_ENENTCGC varchar(14), @A1_CGC varchar(14), @A1_CONTATO varchar(15),
								@A1_INSCR varchar(18), @A1_SUFRAMA varchar(12), @A1_BAIRROC varchar(20), @A1_CEPC varchar(8),
								@A1_MUNC varchar(15), @A1_ESTC varchar(2), @A1_BAIRROE varchar(60), @A1_CEPE varchar(8),
								@A1_MUNE varchar(60), @A1_ESTE varchar(2), @A1_EMAIL varchar(60), @A1_CODMUN int, 
								@A1_INSCRUR varchar(18), @A1_CODLOJA varchar(8), 
								@A1_CODCIDADE int, @A1_CODCIDENT int, @A1_UFENTREGA char(2)
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			DECLARE @R_E_C_N_O_ int

			SET @R_E_C_N_O_ = (select max(R_E_C_N_O_) from SA1010) + 1

			IF @R_E_C_N_O_ IS NULL 
				SET @R_E_C_N_O_ = 1
			IF @R_E_C_N_O_ <= 0 
				SET @R_E_C_N_O_ = 1

			INSERT INTO SA1010 (A1_FILIAL,A1_COD,A1_LOJA,A1_NOME,A1_PESSOA,A1_NREDUZ,A1_TIPO,A1_END,A1_NUMERO,A1_MUN,A1_EST,
								A1_BAIRRO,A1_ESTADO,A1_CEP,A1_DDI,A1_DDD,A1_TEL,A1_TELEX,A1_FAX,A1_ENDENTNR,A1_ENDENT,
								A1_ENENTCGC,A1_CGC,A1_CONTATO,A1_INSCR,A1_SUFRAMA,A1_BAIRROC,A1_CEPC,A1_MUNC,A1_ESTC,
								A1_BAIRROE,A1_CEPE,A1_MUNE,A1_ESTE,A1_EMAIL,A1_CODMUN,R_E_C_N_O_,A1_INSCRUR,
								A1_CODLOJA,A1_CODCIDADE,A1_CODCIDENT,A1_UFENTREGA)
			Values (@A1_FILIAL,@A1_COD,@A1_LOJA,@A1_NOME,@A1_PESSOA,@A1_NREDUZ,@A1_TIPO,@A1_END,@A1_NUMERO,@A1_MUN,@A1_EST,
								@A1_BAIRRO,@A1_ESTADO,@A1_CEP,@A1_DDI,@A1_DDD,@A1_TEL,@A1_TELEX,@A1_FAX,@A1_ENDENTNR,@A1_ENDENT,
								@A1_ENENTCGC,@A1_CGC,@A1_CONTATO,@A1_INSCR,@A1_SUFRAMA,@A1_BAIRROC,@A1_CEPC,@A1_MUNC,@A1_ESTC,
								@A1_BAIRROE,@A1_CEPE,@A1_MUNE,@A1_ESTE,@A1_EMAIL,@A1_CODMUN,@R_E_C_N_O_,@A1_INSCRUR,
								@A1_CODLOJA,@A1_CODCIDADE,@A1_CODCIDENT,@A1_UFENTREGA)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update SA1010 SET A1_NOME=@A1_NOME, A1_PESSOA=@A1_PESSOA, A1_NREDUZ=@A1_NREDUZ, A1_TIPO=@A1_TIPO, A1_END=@A1_END, 
							  A1_NUMERO=@A1_NUMERO, A1_MUN=@A1_MUN, A1_EST=@A1_EST, A1_BAIRRO=@A1_BAIRRO, A1_ESTADO=@A1_ESTADO, 
							  A1_CEP=@A1_CEP, A1_DDI=@A1_DDI, A1_DDD=@A1_DDD, A1_TEL=@A1_TEL, A1_TELEX=@A1_TELEX, A1_FAX=@A1_FAX, 
							  A1_ENDENTNR=@A1_ENDENTNR, A1_ENDENT=@A1_ENDENT, A1_ENENTCGC=@A1_ENENTCGC,  
							  A1_CONTATO=@A1_CONTATO, A1_INSCR=@A1_INSCR, A1_SUFRAMA=@A1_SUFRAMA, A1_BAIRROC=@A1_BAIRROC, 
							  A1_CEPC=@A1_CEPC, A1_MUNC=@A1_MUNC, A1_ESTC=@A1_ESTC, A1_BAIRROE=@A1_BAIRROE, A1_CEPE=@A1_CEPE, 
							  A1_MUNE=@A1_MUNE, A1_ESTE=@A1_ESTE, A1_EMAIL=@A1_EMAIL, A1_CODMUN=@A1_CODMUN, A1_INSCRUR=@A1_INSCRUR, 
							  A1_CODLOJA=@A1_CODLOJA, A1_CODCIDADE=@A1_CODCIDADE, A1_CODCIDENT=@A1_CODCIDENT, A1_UFENTREGA=@A1_UFENTREGA
			where A1_CGC = @A1_CGC

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from SA1010 where A1_CGC = @A1_CGC

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spClienteGlobal',14,1)     
	End  
End 