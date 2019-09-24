USE [MEGASIM]
GO

/****** Object:  Table [dbo].[BOLETO]    Script Date: 02/22/2012 14:19:12 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BOLETO](
	[EMPRESA_ID] [int] NOT NULL,
	[NUMR_DOC] [int] NOT NULL,
	[SEQ] [int] NOT NULL,
	[NUMR_DP] [nvarchar](12) NULL,
	[CGCCPF] [nvarchar](25) NOT NULL,
	[LOCAL_PAGTO] [nvarchar](60) NULL,
	[DT_VENC] [datetime] NULL,
	[DT_DOC] [datetime] NULL,
	[VALOR_DOC] [float] NULL,
	[VALOR_DESCONTO] [float] NULL,
	[VALOR_COBRADO] [float] NULL,
	[INSTRUCAO] [nvarchar](255) NULL,
	[TIPO_DOC] [nvarchar](2) NULL,
	[CLIENTE] [nvarchar](70) NULL,
	[ENDERECO] [nvarchar](40) NULL,
	[UF] [nvarchar](2) NULL,
	[CIDADE] [nvarchar](40) NULL,
	[CEP] [nvarchar](10) NULL
) ON [PRIMARY]

GO


