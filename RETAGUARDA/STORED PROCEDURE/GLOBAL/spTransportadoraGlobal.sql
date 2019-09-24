
ALTER procedure spTransportadoraGlobal @Acao int,
 										@MFTFILIAL varchar(2),@MFTCOD nvarchar(6),@MFTNOME nvarchar(60),@MFTNREDUZ nvarchar(60),
 										@MFTEND nvarchar(60),@MFTBAIRRO nvarchar(60),@MFTMUN nvarchar(60),@MFTEST varchar(2),
 										@MFTCEP varchar(8),@MFTDDD varchar(3),@MFTTEL varchar(15),@MFTCGC varchar(14),
 										@MFTINSEST varchar(15),@MFTEMAIL nvarchar(30),@MFTREGISTRO int,@MFTCODCID int,
 										@MFTNATUREZA char(1),@MFTRUA nvarchar(12)
AS BEGIN
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			INSERT INTO MFT010 (MFTFILIAL,MFTCOD,MFTNOME,MFTNREDUZ,MFTEND,MFTBAIRRO,MFTMUN,MFTEST,
								MFTCEP,MFTDDD,MFTTEL,MFTCGC,MFTINSEST,MFTEMAIL,MFTREGISTRO,MFTCODCID,
								MFTNATUREZA,MFTRUA)
			Values (@MFTFILIAL,@MFTCOD,@MFTNOME,@MFTNREDUZ,@MFTEND,@MFTBAIRRO,@MFTMUN,@MFTEST,
					@MFTCEP,@MFTDDD,@MFTTEL,@MFTCGC,@MFTINSEST,@MFTEMAIL,@MFTREGISTRO,@MFTCODCID,
					@MFTNATUREZA,@MFTRUA)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update MFT010 SET MFTFILIAL=@MFTFILIAL,MFTNOME=@MFTNOME,MFTNREDUZ=@MFTNREDUZ,MFTEND=@MFTEND,MFTBAIRRO=@MFTBAIRRO,
							  MFTMUN=@MFTMUN,MFTEST=@MFTEST,MFTCEP=@MFTCEP,MFTDDD=@MFTDDD,MFTTEL=@MFTTEL,MFTINSEST=@MFTINSEST,
							  MFTEMAIL=@MFTEMAIL,MFTCODCID=@MFTCODCID,MFTNATUREZA=@MFTNATUREZA,MFTRUA=@MFTRUA
			where MFTCGC = @MFTCGC

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from MFT010 where MFTCGC = @MFTCGC

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spClienteGlobal',14,1)     
	End  
End 