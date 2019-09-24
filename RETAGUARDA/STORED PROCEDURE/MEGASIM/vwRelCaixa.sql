--USE [MEGASIM]
GO

drop view [vwRelCaixa]

/****** Object:  View [dbo].[vwRelCaixa]    Script Date: 05/07/2019 08:14:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwRelCaixa]
AS
SELECT   dbo.ITEMLANCAMENTO.SEQ, dbo.ITEMLANCAMENTO.FORMAPAGTO_ID, dbo.ITEMLANCAMENTO.VALOR_ITEM, dbo.ITEMLANCAMENTO.STATUS, 
                         dbo.ITEMLANCAMENTO.DT_VENCIMENTO, dbo.ITEMLANCAMENTO.DT_BAIXA, dbo.ITEMLANCAMENTO.DT_CANCELA, dbo.ITEMLANCAMENTO.VALOR_DESCONTO, 
                         dbo.FORMAPAGTO.DESCRICAO, dbo.PEDIDO.CLIENTE_ID, dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.DT_REQ, dbo.PEDIDO.STATUS AS Sit_Pedido, 
                         dbo.PEDIDO.NOME_CLIENTE, dbo.PEDIDO.PEDIDO_ID, dbo.PEDIDO.ESTABELECIMENTO_ID, dbo.PEDIDO.VALOR_TOTAL, dbo.PEDIDO.VALOR_DESCONTO AS DescCabeca,
                          dbo.PEDIDO.CARTAOBARRA_ID, dbo.PEDIDO.NUMERO_CAIXA_CPU, dbo.LANCAMENTO.TIPO_LANCAMENTO, 
                         dbo.ESTABELECIMENTO.DESCRICAO AS Estabelecimento
FROM         dbo.PEDIDO WITH (NOLOCK) INNER JOIN
                         dbo.ESTABELECIMENTO WITH (NOLOCK) ON dbo.PEDIDO.ESTABELECIMENTO_ID = dbo.ESTABELECIMENTO.ESTABELECIMENTO_ID LEFT OUTER JOIN
                         dbo.ITEMLANCAMENTO WITH (NOLOCK) RIGHT OUTER JOIN
                         dbo.FORMAPAGTO WITH (NOLOCK) ON dbo.ITEMLANCAMENTO.FORMAPAGTO_ID = dbo.FORMAPAGTO.FORMAPAGTO_ID RIGHT OUTER JOIN
                         dbo.LANCAMENTO WITH (NOLOCK) ON dbo.ITEMLANCAMENTO.LANCAMENTO_ID = dbo.LANCAMENTO.LANCAMENTO_ID ON 
                         dbo.PEDIDO.PEDIDO_ID = dbo.LANCAMENTO.NUMR_DOC

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[55] 4[8] 2[27] 3) )"
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
               Top = 10
               Left = 293
               Bottom = 227
               Right = 516
            End
            DisplayFlags = 280
            TopColumn = 5
         End
         Begin Table = "ESTABELECIMENTO"
            Begin Extent = 
               Top = 0
               Left = 12
               Bottom = 130
               Right = 223
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "ITEMLANCAMENTO"
            Begin Extent = 
               Top = 55
               Left = 852
               Bottom = 185
               Right = 1049
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "FORMAPAGTO"
            Begin Extent = 
               Top = 18
               Left = 1123
               Bottom = 148
               Right = 1320
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "LANCAMENTO"
            Begin Extent = 
               Top = 31
               Left = 573
               Bottom = 161
               Right = 783
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
         ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelCaixa'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'Alias = 900
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelCaixa'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelCaixa'
GO


