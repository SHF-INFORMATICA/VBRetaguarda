--USE [MEGASIM]
GO

/****** Object:  Table [dbo].[OSVEICULO]    Script Date: 29/07/2019 11:39:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[OSVEICULO](
	[VEICULO_ID] [bigint] NOT NULL,
	[PESSOA_ID] [bigint] NOT NULL,
	[PLACA] [nvarchar](10) NOT NULL,
	[DESCRICAO] [nvarchar](100) NOT NULL,
	[MOTOR] [nvarchar](100) NULL,
	[CHASSI] [nvarchar](100) NULL,
	[NUMR_FROTA] [nvarchar](10) NOT NULL,
	[ANO] [int] NULL,
	[MODELO] [int] NULL,
	[COMBUSTIVEL_ID] [bigint] NOT NULL,
	[COR_ID] [bigint] NULL,
	[TIPO_VEICULO_ID] [bigint] NULL,
	[MARCA_ID] [bigint] NULL,
 CONSTRAINT [PK_OSVEICULO] PRIMARY KEY CLUSTERED 
(
	[VEICULO_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


