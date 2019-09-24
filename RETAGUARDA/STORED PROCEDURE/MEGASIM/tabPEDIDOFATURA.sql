--USE [MEGASIM]
GO

/****** Object:  Table [dbo].[PEDIDOFATURA]    Script Date: 05/07/2019 11:51:31 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[PEDIDOFATURA](
	[PEDIDOFATURA_ID] [bigint] NOT NULL,
	[PEDIDO_ID] [bigint] NOT NULL,
	[TABELAPRECO_ID] [int] NOT NULL,
	[FORMAPAGTO_ID] [int] NOT NULL,
	[TIPOVENDA_ID] [int] NOT NULL,
 CONSTRAINT [PK_PEDIDOFATURA] PRIMARY KEY CLUSTERED 
(
	[PEDIDOFATURA_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[PEDIDOFATURA]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOFATURA_FORMAPAGTO] FOREIGN KEY([FORMAPAGTO_ID])
REFERENCES [dbo].[FORMAPAGTO] ([FORMAPAGTO_ID])
GO

ALTER TABLE [dbo].[PEDIDOFATURA] CHECK CONSTRAINT [FK_PEDIDOFATURA_FORMAPAGTO]
GO

ALTER TABLE [dbo].[PEDIDOFATURA]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOFATURA_PEDIDO] FOREIGN KEY([PEDIDO_ID])
REFERENCES [dbo].[PEDIDO] ([PEDIDO_ID])
GO

ALTER TABLE [dbo].[PEDIDOFATURA] CHECK CONSTRAINT [FK_PEDIDOFATURA_PEDIDO]
GO

ALTER TABLE [dbo].[PEDIDOFATURA]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOFATURA_TABELAPRECO] FOREIGN KEY([TABELAPRECO_ID])
REFERENCES [dbo].[TABELAPRECO] ([TABELAPRECO_ID])
GO

ALTER TABLE [dbo].[PEDIDOFATURA] CHECK CONSTRAINT [FK_PEDIDOFATURA_TABELAPRECO]
GO

