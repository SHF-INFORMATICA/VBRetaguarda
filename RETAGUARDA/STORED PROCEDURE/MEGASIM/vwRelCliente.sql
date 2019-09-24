--USE [MEGASIM]
GO

drop view [vwRelCliente]

/****** Object:  View [dbo].[vwRelCliente]    Script Date: 04/07/2019 17:52:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwRelCliente]
AS
SELECT   dbo.PEDIDO.PEDIDO_ID, dbo.PEDIDO.CLIENTE_ID, dbo.PEDIDO.EMPRESA_ID, dbo.PEDIDO.VENDEDOR_ID, dbo.PEDIDO.ESTABELECIMENTO_ID, dbo.PEDIDO.DT_REQ, 
                         dbo.PEDIDO.STATUS AS status_pedido, dbo.CLIENTE.STATUS AS status_cli, dbo.CLIENTE.SEXO, dbo.CLIENTE.CONTATO, dbo.CLIENTE.LIMITE_CREDITO, 
                         dbo.ENDERECO.ENDERECO_ID, dbo.ENDERECO.PESSOA_ID, dbo.ENDERECO.CEP_ID, dbo.ENDERECO.RUA, dbo.ENDERECO.BAIRRO, dbo.ENDERECO.COMPLEMENTO, 
                         dbo.ENDERECO.TIPO, dbo.ENDERECO.NUMERO, dbo.PESSOA.CNPJCPF, dbo.PESSOA.DESCRICAO, dbo.PESSOA.RAZAO, dbo.PESSOA.DATA_CAD, 
                         dbo.FONE.NUMERO AS NumeroFone, dbo.FONE.DDD, dbo.FONE.LOCAL, dbo.CEP.CIDADE, dbo.CEP.UF, dbo.CEP.IBGE_ID, dbo.PEDIDOFATURA.PEDIDOFATURA_ID, 
                         dbo.PEDIDOFATURA.TABELAPRECO_ID, dbo.PEDIDOFATURA.FORMAPAGTO_ID, dbo.PEDIDOFATURA.TIPOVENDA_ID
FROM         dbo.PEDIDO WITH (NOLOCK) INNER JOIN
                         dbo.CLIENTE WITH (NOLOCK) ON dbo.PEDIDO.CLIENTE_ID = dbo.CLIENTE.CLIENTE_ID INNER JOIN
                         dbo.PESSOA WITH (NOLOCK) ON dbo.CLIENTE.PESSOA_ID = dbo.PESSOA.PESSOA_ID INNER JOIN
                         dbo.ENDERECO WITH (NOLOCK) ON dbo.PESSOA.PESSOA_ID = dbo.ENDERECO.PESSOA_ID INNER JOIN
                         dbo.FONE WITH (NOLOCK) ON dbo.PESSOA.PESSOA_ID = dbo.FONE.PESSOA_ID INNER JOIN
                         dbo.PEDIDOFATURA ON dbo.PEDIDO.PEDIDO_ID = dbo.PEDIDOFATURA.PEDIDO_ID LEFT OUTER JOIN
                         dbo.CEP ON dbo.ENDERECO.CEP_ID = dbo.CEP.CEP_ID

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[20] 2[13] 3) )"
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
               Bottom = 194
               Right = 261
            End
            DisplayFlags = 280
            TopColumn = 7
         End
         Begin Table = "CLIENTE"
            Begin Extent = 
               Top = 63
               Left = 301
               Bottom = 193
               Right = 516
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
         Begin Table = "ENDERECO"
            Begin Extent = 
               Top = 270
               Left = 273
               Bottom = 400
               Right = 470
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "FONE"
            Begin Extent = 
               Top = 402
               Left = 38
               Bottom = 532
               Right = 235
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "CEP"
            Begin Extent = 
               Top = 402
               Left = 273
               Bottom = 532
               Right = 470
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PEDIDOFATURA"
            Begin Extent = 
               Top = 5
               Left = 804
               Bottom = 163
               Right = 1001
            End
            DisplayFlags = 280
  ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelCliente'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'          TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 30
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelCliente'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwRelCliente'
GO


