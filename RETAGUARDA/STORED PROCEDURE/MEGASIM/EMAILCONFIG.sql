--USE [SHFINFO]
GO

/****** Object:  Table [dbo].[EMAILCONFIG]    Script Date: 27/08/2019 16:40:17 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[EMAILCONFIG](
	[EMAILCONFIG_ID] [bigint] NOT NULL,
	[ESTABELECIMENTO_ID] [bigint] NOT NULL,
	[REMETENTE_A] [nvarchar](50) NOT NULL,
	[DESTINATARIO_A] [nvarchar](50) NOT NULL,
	[SENHA_EMAIL] [nvarchar](50) NOT NULL,
	[SMTP] [nvarchar](50) NOT NULL,
	[PORTA] [nvarchar](50) NOT NULL,
	[SSL] [nchar](10) NULL,
	[TLS] [nchar](10) NULL,
 CONSTRAINT [PK_EMAILCONFIG] PRIMARY KEY CLUSTERED 
(
	[EMAILCONFIG_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


