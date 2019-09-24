--USE [MEGASIM]
GO

drop view [vwPOSICAOESTOQUE]

/****** Object:  View [dbo].[vwPOSICAOESTOQUE]    Script Date: 05/07/2019 08:10:13 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwPOSICAOESTOQUE]
AS
SELECT   dbo.ESTABELECIMENTO.EMPRESA_ID, dbo.ESTABELECIMENTO.DESCRICAO AS DescEstab, dbo.ESTABELECIMENTO.INDR_PANIFIC, 
                         dbo.ESTABELECIMENTO.CNPJCPF AS CNPJ_Estab, dbo.PEDIDO.PEDIDO_ID, dbo.PEDIDO.CLIENTE_ID, dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.USUARIO_ID, 
                         dbo.PEDIDO.CARTAOBARRA_ID, dbo.PEDIDO.DT_REQ, dbo.PEDIDO.STATUS AS StatusPedido, dbo.PEDIDO.TIPO_REGISTRO, dbo.PEDIDO.NOME_CLIENTE, 
                         dbo.PEDIDO.NUMERO_CAIXA_CPU, dbo.PEDIDOITEM.SEQ_ID, dbo.PEDIDOITEM.PRODUTO_ID, dbo.PEDIDOITEM.QTD_PEDIDA, dbo.PEDIDOITEM.VALOR_ITEM, 
                         dbo.PEDIDOITEM.PERC_DESC, dbo.PEDIDOITEM.CFOP_ID, dbo.PEDIDOITEM.STRIBUTARIA, dbo.PEDIDOITEM.VLRBASEICMS, dbo.PEDIDOITEM.PERCICMS, 
                         dbo.PEDIDOITEM.VLRICMS, dbo.PEDIDOITEM.VLRBASEICMSSUBST, dbo.PEDIDOITEM.PERCICMSSUBST, dbo.PEDIDOITEM.VLRICMSSUBST, 
                         dbo.PEDIDOITEM.PERCREDUCAOICMS, dbo.PEDIDOITEM.PERCIVA, dbo.PEDIDOITEM.PERC_IPI, dbo.PEDIDOITEM.VLR_IPI, dbo.PEDIDOITEM.VALOR_DESCONTO, 
                         dbo.PEDIDOITEM.STATUS AS StatusItem, dbo.PEDIDOITEM.PRECO_CUSTO AS PrCustoItem, dbo.PEDIDOITEM.TIPO_REG, dbo.PEDIDOITEM.PESO_ITEM, 
                         dbo.PEDIDOITEM.USU_ATENDE, dbo.PRODUTO.CODG_PRODUTO, dbo.PRODUTO.DESCRICAO AS DescProduto, dbo.PRODUTO.FAMILIAPRODUTO_ID, 
                         dbo.PRODUTO.UNIDADE_MEDIDA, dbo.PRODUTO.CODG_BARRA, dbo.PRODUTO.SITUACAO AS StatusProduto, dbo.PRODUTO.SITUACAO_TRIBUTARIA, 
                         dbo.PRODUTO.ALIQUOTA_ICMS, dbo.PRODUTO.TIPO_PROD, dbo.PRODUTO.REFERENCIA, dbo.PRODUTO.CODG_NCM, dbo.PRODUTO.FORNECEDOR_ID, 
                         dbo.PRODUTO.PRECO_CUSTO_ANTERIOR, dbo.PRODUTO.PRECO_CUSTO AS PrCustoProduto, dbo.PRODUTO.PRECO_ATACADO, dbo.PRODUTO.PRECO_Venda, 
                         dbo.PRODUTO.DT_ULT_VENDA, dbo.PRODUTO.DT_ULT_COMPRA, dbo.PRODUTO.PESO_LIQUIDO, dbo.PRODUTO.PESO_BRUTO, dbo.PRODUTO.TAMANHO, 
                         dbo.PRODUTO.MARCA_ID, dbo.PRODUTO.PRODUTO_BALANCA, dbo.PRODUTO.CONCEDER_PRODUCAO, dbo.PEDIDO.ESTABELECIMENTO_ID
FROM         dbo.ESTABELECIMENTO WITH (NOLOCK) INNER JOIN
                         dbo.PEDIDO WITH (NOLOCK) ON dbo.ESTABELECIMENTO.ESTABELECIMENTO_ID = dbo.PEDIDO.ESTABELECIMENTO_ID INNER JOIN
                         dbo.PEDIDOITEM WITH (NOLOCK) ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOITEM.PEDIDO_ID INNER JOIN
                         dbo.PRODUTO WITH (NOLOCK) ON dbo.PEDIDOITEM.PRODUTO_ID = dbo.PRODUTO.PRODUTO_ID

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
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
               Top = 6
               Left = 38
               Bottom = 136
               Right = 273
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDO"
            Begin Extent = 
               Top = 4
               Left = 382
               Bottom = 134
               Right = 605
            End
            DisplayFlags = 280
            TopColumn = 16
         End
         Begin Table = "PEDIDOITEM"
            Begin Extent = 
               Top = 95
               Left = 756
               Bottom = 225
               Right = 955
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PRODUTO"
            Begin Extent = 
               Top = 173
               Left = 260
               Bottom = 303
               Right = 508
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
  ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPOSICAOESTOQUE'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'    End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPOSICAOESTOQUE'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPOSICAOESTOQUE'
GO


