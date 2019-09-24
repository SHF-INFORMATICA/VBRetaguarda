--USE [MEGASIM]
GO

drop view [vwRelVenda]

/****** Object:  View [dbo].[vwRelVenda]    Script Date: 05/07/2019 08:19:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwRelVenda]
AS
SELECT   dbo.PEDIDO.NOME_CLIENTE, dbo.TIPOVENDA.DESCRICAO AS DescTipoVenda, dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.CLIENTE_ID, dbo.PEDIDO.EMPRESA_ID, 
                         dbo.PEDIDO.CGCCPF, dbo.PEDIDO.DT_REQ, dbo.PEDIDO.STATUS, dbo.PEDIDO.VALOR_DESCONTO, dbo.PEDIDO.VALOR_TOTAL, dbo.PEDIDO.PEDIDO_ID, 
                         dbo.PEDIDOITEM.SEQ_ID, dbo.PEDIDO.ESTABELECIMENTO_ID, dbo.PEDIDOITEM.PRODUTO_ID, dbo.PRODUTO.CODG_PRODUTO, dbo.PEDIDOITEM.QTD_PEDIDA, 
                         dbo.PEDIDOITEM.VALOR_ITEM, dbo.PEDIDOITEM.VALOR_DESCONTO AS DescontoItem, dbo.PEDIDOITEM.STATUS AS StatusItem, 
                         dbo.PEDIDOITEM.PRECO_CUSTO AS CustoItem, dbo.PRODUTO.DESCRICAO AS DescProduto, dbo.PRODUTO.FAMILIAPRODUTO_ID, dbo.PRODUTO.CODG_BARRA, 
                         dbo.PRODUTO.SITUACAO, dbo.PRODUTO.SITUACAO_TRIBUTARIA, dbo.PRODUTO.REFERENCIA, dbo.PRODUTO.CODG_NCM, 
                         dbo.PRODUTO.PRECO_CUSTO AS CustoProduto, dbo.PRODUTO.PRECO_ATACADO, dbo.PRODUTO.PRECO_Venda, dbo.PEDIDO.CARTAOBARRA_ID, 
                         dbo.PEDIDO.NUMERO_CAIXA_CPU, dbo.PEDIDO.VALOR_RECEBIDO, dbo.PEDIDO.USUARIO_LIBERA_VENDA, dbo.PEDIDO.TIPO_REGISTRO, dbo.PEDIDO.USUARIO_ID, 
                         dbo.ESTABELECIMENTO.DESCRICAO AS NomeEstab, dbo.ESTABELECIMENTO.CNPJCPF, dbo.PRODUTO.UNIDADE_MEDIDA, dbo.PEDIDOFATURA.PEDIDOFATURA_ID, 
                         dbo.PEDIDOFATURA.TABELAPRECO_ID, dbo.PEDIDOFATURA.FORMAPAGTO_ID, dbo.PEDIDOFATURA.TIPOVENDA_ID AS Expr1
FROM         dbo.ESTABELECIMENTO WITH (NOLOCK) INNER JOIN
                         dbo.PEDIDO WITH (NOLOCK) ON dbo.ESTABELECIMENTO.ESTABELECIMENTO_ID = dbo.PEDIDO.ESTABELECIMENTO_ID INNER JOIN
                         dbo.PEDIDOFATURA ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOFATURA.PEDIDO_ID INNER JOIN
                         dbo.TIPOVENDA ON dbo.PEDIDOFATURA.TIPOVENDA_ID = dbo.TIPOVENDA.TIPOVENDA_ID LEFT OUTER JOIN
                         dbo.PEDIDOITEM ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOITEM.PEDIDO_ID FULL OUTER JOIN
                         dbo.PRODUTO ON dbo.PEDIDOITEM.PRODUTO_ID = dbo.PRODUTO.PRODUTO_ID

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[57] 4[4] 2[20] 3) )"
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
         Begin Table = "ESTABELECIMENTO"
            Begin Extent = 
               Top = 57
               Left = 28
               Bottom = 187
               Right = 263
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDO"
            Begin Extent = 
               Top = 3
               Left = 322
               Bottom = 133
               Right = 545
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "TIPOVENDA"
            Begin Extent = 
               Top = 0
               Left = 881
               Bottom = 130
               Right = 1103
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDOITEM"
            Begin Extent = 
               Top = 158
               Left = 765
               Bottom = 288
               Right = 965
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PRODUTO"
            Begin Extent = 
               Top = 181
               Left = 1025
               Bottom = 311
               Right = 1273
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDOFATURA"
            Begin Extent = 
               Top = 6
               Left = 583
               Bottom = 163
               Right = 780
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
         Wi' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelVenda'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'dth = 1500
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelVenda'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelVenda'
GO


