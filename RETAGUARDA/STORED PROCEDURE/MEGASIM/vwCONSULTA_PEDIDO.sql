USE [MEGASIM]
GO

/****** Object:  View [dbo].[vwCONSULTA_PEDIDO]    Script Date: 13/09/2019 11:14:03 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

DROP VIEW vwCONSULTA_PEDIDO

CREATE VIEW vwCONSULTA_PEDIDO
AS
SELECT        dbo.PEDIDO.CLIENTE_ID, dbo.PEDIDO.EMPRESA_ID, dbo.PEDIDO.DT_REQ, dbo.PEDIDO.NOME_CLIENTE, dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.STATUS AS SIT_PEDIDO, dbo.PEDIDO.VALOR_DESCONTO AS DescCabeca, 
                         dbo.PEDIDOITEM.PEDIDO_ID, dbo.PEDIDOITEM.SEQ_ID, dbo.PEDIDOITEM.PRODUTO_ID, dbo.PEDIDOITEM.QTD_PEDIDA, dbo.PEDIDOITEM.VALOR_ITEM, dbo.PEDIDOITEM.CFOP_ID, dbo.PEDIDOITEM.STRIBUTARIA, 
                         dbo.PEDIDOITEM.VALOR_DESCONTO, dbo.PEDIDOITEM.STATUS, dbo.PEDIDOITEM.PRECO_CUSTO, dbo.PRODUTO.DESCRICAO, dbo.PRODUTO.FAMILIAPRODUTO_ID, dbo.PRODUTO.SITUACAO, 
                         dbo.PRODUTO.SITUACAO_TRIBUTARIA, dbo.PRODUTO.CODG_NCM, dbo.PRODUTO.PRECO_CUSTO AS Preço_Custo_Produto, dbo.PRODUTO.PRECO_ATACADO, dbo.PRODUTO.PRECO_Venda, dbo.NF.NUMR_NOTA, 
                         dbo.NF.SERIE_NOTA, dbo.NF.DT_EMISSAO, dbo.NF.QTD_VOLUME, dbo.NF.DT_CANCELA, dbo.NF.PESO_BRUTO, dbo.NF.PESO_LIQUIDO, dbo.CUPOM.NUMR_CUPOM, dbo.CUPOM.MODELO_DOC, 
                         dbo.PEDIDO.NUMERO_CAIXA_CPU, dbo.PEDIDO.TIPO_REGISTRO, dbo.PEDIDO.CARTAOBARRA_ID, dbo.PEDIDO.VALOR_TOTAL, dbo.PEDIDO.ESTABELECIMENTO_ID, dbo.PEDIDO.USUARIO_LIBERA_VENDA, 
                         dbo.PRODUTO.UNIDADE_MEDIDA, dbo.PEDIDOITEM.TIPO_REG, dbo.CLIENTE.NOME, dbo.PESSOA.CNPJCPF, dbo.PESSOA.DESCRICAO AS Expr1, dbo.PESSOA.RAZAO, dbo.PEDIDOFATURA.PEDIDOFATURA_ID, 
                         dbo.PEDIDOFATURA.TABELAPRECO_ID, dbo.PEDIDOFATURA.FORMAPAGTO_ID, dbo.PEDIDOFATURA.TIPOVENDA_ID
FROM            dbo.PEDIDO WITH (NOLOCK) INNER JOIN
                         dbo.CLIENTE WITH (NOLOCK) ON dbo.PEDIDO.CLIENTE_ID = dbo.CLIENTE.CLIENTE_ID INNER JOIN
                         dbo.PESSOA WITH (NOLOCK) ON dbo.CLIENTE.PESSOA_ID = dbo.PESSOA.PESSOA_ID INNER JOIN
                         dbo.PEDIDOFATURA ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOFATURA.PEDIDO_ID INNER JOIN
                         dbo.PEDIDONF ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDONF.PEDIDO_ID LEFT OUTER JOIN
                         dbo.NF WITH (NOLOCK) ON dbo.PEDIDONF.NF_ID = dbo.NF.NF_ID LEFT OUTER JOIN
                         dbo.PEDIDOITEM WITH (NOLOCK) ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOITEM.PEDIDO_ID LEFT OUTER JOIN
                         dbo.PRODUTO WITH (NOLOCK) ON dbo.PEDIDOITEM.PRODUTO_ID = dbo.PRODUTO.PRODUTO_ID LEFT OUTER JOIN
                         dbo.CUPOM WITH (NOLOCK) ON dbo.PEDIDO.PEDIDO_ID = dbo.CUPOM.PEDIDO_ID

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[44] 4[6] 2[32] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "PEDIDO"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 136
               Right = 261
            End
            DisplayFlags = 280
            TopColumn = 16
         End
         Begin Table = "CLIENTE"
            Begin Extent = 
               Top = 138
               Left = 38
               Bottom = 268
               Right = 253
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PESSOA"
            Begin Extent = 
               Top = 270
               Left = 38
               Bottom = 400
               Right = 235
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDOFATURA"
            Begin Extent = 
               Top = 0
               Left = 353
               Bottom = 148
               Right = 550
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "NF"
            Begin Extent = 
               Top = 39
               Left = 828
               Bottom = 169
               Right = 1038
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDOITEM"
            Begin Extent = 
               Top = 270
               Left = 273
               Bottom = 400
               Right = 472
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PRODUTO"
            Begin Extent = 
               Top = 534
               Left = 38
               Bottom = 664
               Right = 286
            End
            DisplayFlags = 280
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCONSULTA_PEDIDO'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'
            TopColumn = 0
         End
         Begin Table = "CUPOM"
            Begin Extent = 
               Top = 666
               Left = 38
               Bottom = 796
               Right = 235
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDONF"
            Begin Extent = 
               Top = 6
               Left = 588
               Bottom = 119
               Right = 785
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCONSULTA_PEDIDO'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCONSULTA_PEDIDO'
GO


