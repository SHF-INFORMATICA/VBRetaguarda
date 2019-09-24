USE [GLOBAL]
GO
/****** Object:  StoredProcedure [dbo].[spNFEAtualizacao]    Script Date: 01/10/2017 15:40:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Yuri G. Lemes
-- Create date: 21/03/2010
-- Description:	Atualização de alguns Campos referentes à 
--              Nota Fiscal Eletrônica
-- =============================================
create PROCEDURE [dbo].[spNFEAtualizacao]
  @MFAMOTRESU varchar(200)
 ,@MFADOC varchar(06)
 ,@MFACHAVENFE varchar(100)
 ,@MFACODSTAT int 
 ,@MFASERIE varchar(03)
 ,@MFASEQUENCIA int
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	--SELECT <@Param1, sysname, @p1>, <@Param2, sysname, @p2>
	UPDATE [dbo].[MFA010]
   SET [MFAMOTRESU] = @MFAMOTRESU
    --  ,[MFADOC] = @MFADOC
      ,[MFACHAVENFE] = @MFACHAVENFE
      ,[MFACODSTAT] = @MFACODSTAT
	  ,MFACODMORE = 1
    --  ,[MFASERIE] = @MFASERIE
    WHERE [MFASEQUENCIA] = @MFASEQUENCIA

END
