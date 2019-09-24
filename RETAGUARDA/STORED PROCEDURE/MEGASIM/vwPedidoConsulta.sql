--USE [MEGASIM]
GO

drop view vwPedidoConsulta

/****** Object:  View [dbo].[vwPedidoConsulta]    Script Date: 04/07/2019 17:48:41 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwPedidoConsulta]
AS
SELECT   dbo.PEDIDO.PEDIDO_ID, dbo.CUPOM.NUMR_CUPOM, dbo.NF.NUMR_NOTA, dbo.PEDIDO.NOME_CLIENTE, dbo.PEDIDO.VALOR_TOTAL, dbo.PEDIDO.DT_REQ, 
                         dbo.TIPOVENDA.DESCRICAO AS Faturamento, Pessoa_Vendedor.DESCRICAO AS Vendedor, dbo.PEDIDO.STATUS, dbo.PEDIDO.CLIENTE_ID, dbo.PEDIDO.EMPRESA_ID, 
                         dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.VALOR_DESCONTO AS DescCabeca, dbo.PEDIDO.NUMERO_CAIXA_CPU, dbo.PEDIDO.TIPO_REGISTRO, 
                         dbo.PEDIDO.CARTAOBARRA_ID, dbo.PEDIDO.ESTABELECIMENTO_ID, dbo.PEDIDOFATURA.PEDIDOFATURA_ID, dbo.PEDIDOFATURA.FORMAPAGTO_ID, 
                         dbo.PEDIDOFATURA.TIPOVENDA_ID, dbo.PEDIDOFATURA.TABELAPRECO_ID
FROM         dbo.PEDIDO WITH (NOLOCK) INNER JOIN
                         dbo.CLIENTE WITH (NOLOCK) ON dbo.PEDIDO.CLIENTE_ID = dbo.CLIENTE.CLIENTE_ID INNER JOIN
                         dbo.VENDEDOR ON dbo.PEDIDO.VENDEDOR_ID = dbo.VENDEDOR.VENDEDOR_ID INNER JOIN
                         dbo.PESSOA AS Pessoa_Vendedor ON dbo.VENDEDOR.PESSOA_ID = Pessoa_Vendedor.PESSOA_ID INNER JOIN
                         dbo.PEDIDOFATURA ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOFATURA.PEDIDO_ID INNER JOIN
                         dbo.TIPOVENDA ON dbo.PEDIDOFATURA.TIPOVENDA_ID = dbo.TIPOVENDA.TIPOVENDA_ID LEFT OUTER JOIN
                         dbo.NF WITH (NOLOCK) ON dbo.PEDIDO.PEDIDO_ID = dbo.NF.PEDIDO_ID LEFT OUTER JOIN
                         dbo.CUPOM WITH (NOLOCK) ON dbo.PEDIDO.PEDIDO_ID = dbo.CUPOM.PEDIDO_ID

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[62] 4[5] 2[15] 3) )"
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
            TopColumn = 4
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
         Begin Table = "TIPOVENDA"
            Begin Extent = 
               Top = 8
               Left = 852
               Bottom = 138
               Right = 1074
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "VENDEDOR"
            Begin Extent = 
               Top = 402
               Left = 38
               Bottom = 532
               Right = 235
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Pessoa_Vendedor"
            Begin Extent = 
               Top = 402
               Left = 273
               Bottom = 532
               Right = 470
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "NF"
            Begin Extent = 
               Top = 534
               Left = 38
               Bottom = 664
               Right = 248
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "CUPOM"
            Begin Extent = 
               Top = 666
               Left = 38
               Bottom = 796
               Right = 235
            End
            DisplayFlags = 280' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPedidoConsulta'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'
            TopColumn = 0
         End
         Begin Table = "PEDIDOFATURA"
            Begin Extent = 
               Top = 20
               Left = 469
               Bottom = 179
               Right = 666
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPedidoConsulta'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPedidoConsulta'
GO


