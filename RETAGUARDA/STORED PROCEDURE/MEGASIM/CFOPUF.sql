/****** Object:  Table [dbo].[CFOPUF]    Script Date: 26/05/2019 10:09:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CFOPUF](
	[CFOPUF_ID] [bigint] NOT NULL,
	[CFOP_ID] [nvarchar](10) NOT NULL,
	[UF_ORIGEM] [nvarchar](2) NOT NULL,
	[UF_DESTINO] [nvarchar](2) NOT NULL,
 CONSTRAINT [PK_CFOPUF] PRIMARY KEY CLUSTERED 
(
	[CFOPUF_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[CFOPUF]  WITH CHECK ADD  CONSTRAINT [FK_CFOPUF_CFOP] FOREIGN KEY([CFOP_ID])
REFERENCES [dbo].[CFOP] ([CFOP_ID])
GO

ALTER TABLE [dbo].[CFOPUF] CHECK CONSTRAINT [FK_CFOPUF_CFOP]
GO


