USE [GLOBAL]
GO
/****** Object:  StoredProcedure [dbo].[spNFeFinanceiro]    Script Date: 01/10/2017 08:49:57 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create PROCEDURE [dbo].[spNFeFinanceiro](@CodigoFilial varchar(02),@CodigoCliente varchar(06), @NumeroNota varchar(09),@Serie 
varchar(03))
   -- Add the parameters for the stored procedure here
   AS
BEGIN
   -- SET NOCOUNT ON added to prevent extra result sets from
   -- interfering with SELECT statements.
   SET NOCOUNT ON;

    -- Insert statements for procedure here
   SELECT [E1_FILIAL]
      ,[E1_PREFIXO]
      ,[E1_NUM] AS nDup
      ,[E1_PARCELA] as parcela
      ,[E1_TIPO]
      ,[E1_NATUREZ]
      ,[E1_CLIENTE]
      ,[E1_LOJA]
      ,[E1_NOMCLI]
      ,[E1_EMISSAO]
      ,[E1_VENCTO] as dVenc
      ,[E1_VENCREA]
      ,[E1_VALOR] as vDup
      ,[E1_SITUACA]
      ,[E1_SALDO]
      ,[E1_VENCORI]
      ,[E1_MOEDA]
      ,[E1_DTFATUR]
      ,[E1_NUMNOTA]
      ,[E1_SERIE]
      ,[E1_STATUS]
      ,[E1_ORIGEM]
      ,[E1_FLUXO]
      ,[E1_VLRREAL]
      ,[E1_MESBASE]
      ,[E1_ANOBASE]
      ,[E1_CODEMP]
      ,[E1_TXMOEDA]
      ,[E1_DESDOBR]
      ,[E1_NRDOC]
     ,E1_CARTAO AS CNPJCARTAO
     ,E1_ADM AS tBand
     ,E1_CARTAUT AS cAut
      ,[D_E_L_E_T_]
      ,[R_E_C_N_O_]
      ,[E1_FLAG]
  From [dbo].[SE1010]
  where E1_FILIAL=@CodigoFilial and E1_CLIENTE=@CodigoCliente and E1_NUMNOTA=@NumeroNota --and E1_SERIE=@Serie
End
