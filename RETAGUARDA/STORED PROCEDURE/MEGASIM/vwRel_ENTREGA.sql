--USE [MEGASIM]
GO

drop view [vwRel_ENTREGA]

/****** Object:  View [dbo].[vwRel_ENTREGA]    Script Date: 04/07/2019 18:01:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwRel_ENTREGA]
AS
SELECT   dbo.ENTREGA.ENTREGA_ID, dbo.ENTREGA.PESSOA_ID, dbo.ENTREGA.PEDIDO_ID, dbo.ENTREGA.EMPRESA_ID, dbo.ENTREGA.DT_CAD, dbo.ENTREGA.DT_AGENDA, 
                         dbo.ENTREGA.DT_ENTREGA, dbo.ENTREGA.CEP_ID, dbo.ENTREGA.RUA, dbo.ENTREGA.COMPLEMENTO, dbo.ENTREGA.BAIRRO, dbo.ENTREGA.CIDADE, 
                         dbo.ENTREGA.UF, dbo.ENTREGA.ENTREGADOR_ID, dbo.ENTREGA.ENTREGADOR, dbo.ENTREGA.ATENDENTE_ID, dbo.ENTREGA.ATENDENTE, 
                         dbo.ENTREGA.ENTREGADOR_ID AS Expr1, dbo.ENTREGA.ENTREGADOR AS Expr2, dbo.PESSOA.CNPJCPF, dbo.PESSOA.DESCRICAO, dbo.PEDIDO.ESTABELECIMENTO_ID,
                          dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.USUARIO_ID, dbo.PEDIDO.CGCCPF, dbo.PEDIDO.NOME_CLIENTE, dbo.ESTABELECIMENTO.DESCRICAO AS DescEstab
FROM         dbo.ENTREGA INNER JOIN
                         dbo.PEDIDO ON dbo.ENTREGA.PEDIDO_ID = dbo.PEDIDO.PEDIDO_ID INNER JOIN
                         dbo.PESSOA ON dbo.ENTREGA.PESSOA_ID = dbo.PESSOA.PESSOA_ID INNER JOIN
                         dbo.ESTABELECIMENTO ON dbo.PEDIDO.ESTABELECIMENTO_ID = dbo.ESTABELECIMENTO.ESTABELECIMENTO_ID

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
         Begin Table = "ENTREGA"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 136
               Right = 235
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDO"
            Begin Extent = 
               Top = 0
               Left = 456
               Bottom = 130
               Right = 679
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PESSOA"
            Begin Extent = 
               Top = 6
               Left = 273
               Bottom = 136
               Right = 470
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "ESTABELECIMENTO"
            Begin Extent = 
               Top = 270
               Left = 38
               Bottom = 400
               Right = 273
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
      End' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRel_ENTREGA'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRel_ENTREGA'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRel_ENTREGA'
GO


