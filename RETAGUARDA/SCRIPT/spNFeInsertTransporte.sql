USE [GLOBAL]
GO
/****** Object:  StoredProcedure [dbo].[spNFeInsertTransporte]    Script Date: 24/09/2017 10:24:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Yuri Grandinetti Lemes
-- Create date: 06/06/2010
-- Description:	Inserir Cadastro de Cliente

-- ALTERAÇÕES date: 24/09/2017	Horácio
-- CORREÇÕES: Description:Inserir Cadastro de TRANSPORTADORA
-- O campo MFTCOD será preenchido pela aplicação RETAGUARDA com dados da 
-- tabela TRANSPORTADORA.TRANSP_ID = @MFTCOD daqui desta sp
-- A formatação deste campo é feito com 0000 antes e isso está incorreto.
-- =============================================
ALTER PROCEDURE [dbo].[spNFeInsertTransporte] 
	-- Add the parameters for the stored procedure here
	       @REGISTRO	int	OUT,	
	       @MFTFILIAL	varchar(2)= null,
           @MFTCOD		varchar(6)= null,
           @MFTNOME		varchar(60)= null,
           @MFTNREDUZ	varchar(60)= null,
           @MFTVIA		varchar(15)= null,
           @MFTEND		varchar(60)= null,
           @MFTBAIRRO	varchar(30)= null,
           @MFTMUN		varchar(60)= null,
           @MFTEST		varchar(2)= null,
           @MFTCEP		varchar(8)= null,
           @MFTDDI		varchar(6)= null,
           @MFTDDD		varchar(3)= null,
           @MFTTEL		varchar(15)= null,
           @MFTCGC		varchar(14)= null,
           @MFTTELEX	varchar(10)= null,
           @MFTINSEST	varchar(15)= null,
           @MFTEMAIL	varchar(30)= null,
           @MFTHPAGE	varchar(30)= null,
           @MFTENDPAD	varchar(15)= null,
           @MFTCONTATO	varchar(15)= null,          
           @MFTCODCID	int= null,
           @MFTNATUREZA char(1)= null,
           @MFTRUA		varchar(12)= null	
AS

BEGIN TRANSACTION

Declare @Utimoregistro int
DECLARE @Ultimo VARCHAR(6)
DECLARE @WULTIMO VARCHAR(6)
DECLARE @WZERO VARCHAR(6)

    SET @Utimoregistro = (SELECT max(MFTREGISTRO) as Ultimore FROM dbo.MFT010)+1	

    SET	@Ultimo = (SELECT max(MFTCOD) as Ultimo FROM dbo.MFT010)+1

	IF(ISNULL(@Ultimo,0) = 0)
	begin
	   set @Ultimo=1
	end

	IF(ISNULL(@Utimoregistro,0) = 0)
	begin
	   set @Utimoregistro=1
	end

    IF(LEN(@Ultimo)=1)
	BEGIN
	   set @WZERO='00000'
    END  

	IF(LEN(@Ultimo)=2)
	BEGIN
	   set @WZERO='0000'
	END

	IF(LEN(@Ultimo)=3)
	BEGIN
	   set @WZERO='000'
	END

	IF(LEN(@Ultimo)=4)
	BEGIN
	   set @WZERO='00'
	END

    IF(LEN(@Ultimo)=5)
	BEGIN
	   set @WZERO='0'
	END  
	
    SET @WULTIMO = @WZERO+CONVERT(VARCHAR(3),@Ultimo)

	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO dbo.MFT010
           ([MFTFILIAL]
		   ,[MFTCOD]
           ,[MFTNOME]
           ,[MFTNREDUZ]
           ,[MFTVIA]
           ,[MFTEND]
           ,[MFTBAIRRO]
           ,[MFTMUN]
           ,[MFTEST]
           ,[MFTCEP]
           ,[MFTDDI]
           ,[MFTDDD]
           ,[MFTTEL]
           ,[MFTCGC]
           ,[MFTTELEX]
           ,[MFTINSEST]
           ,[MFTEMAIL]
           ,[MFTHPAGE]
           ,[MFTENDPAD]
           ,[MFTCONTATO]                     
           ,[MFTCODCID]
           ,[MFTNATUREZA]
           ,[MFTRUA]
           ,[MFTREGISTRO])
           
     VALUES
           (@MFTFILIAL,
		   @WULTIMO,
		   @MFTNOME,
           @MFTNREDUZ,
           @MFTVIA,
           @MFTEND,
           @MFTBAIRRO,
           @MFTMUN,
           @MFTEST,
           @MFTCEP,
           @MFTDDI,
           @MFTDDD,
           @MFTTEL,
           @MFTCGC,
           @MFTTELEX,
           @MFTINSEST,
           @MFTEMAIL,
           @MFTHPAGE,
           @MFTENDPAD,
           @MFTCONTATO,          
           @MFTCODCID,
           @MFTNATUREZA,
           @MFTRUA,
		   @Utimoregistro   )

commit
SET	 @REGISTRO	= @Utimoregistro
