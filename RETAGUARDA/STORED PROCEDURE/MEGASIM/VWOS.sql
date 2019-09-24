--USE [MEGASIM]
GO

/****** Object:  View [dbo].[vwOS]    Script Date: 08/04/2019 10:02:28 ******/
SET ANSI_NULLS ON
GO

drop view vwos

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwOS]
AS
SELECT        dbo.OSEQUIPAMENTO.EQUIPAMENTO_ID, dbo.OSEQUIPAMENTO.MARCA_ID, dbo.OSEQUIPAMENTO.COR_ID, dbo.OSEQUIPAMENTO.TIPO_EQP, dbo.OSEQUIPAMENTO.DT_CAD, 
                         dbo.OSEQUIPAMENTO.DESCRICAO AS Nome_Eqp, dbo.OSEQUIPAMENTO.IDENTIFICACAO, dbo.OSEQUIPAMENTO.ANO, dbo.OSEQUIPAMENTO.MODELO, dbo.OSVEICULO.VEICULO_ID, 
                         dbo.OSVEICULO.COMBUSTIVEL_ID, dbo.OSVEICULO.PLACA, dbo.OSVEICULO.DESCRICAO AS DESCRICAOVEICULO, dbo.OSVEICULO.MOTOR, dbo.OSVEICULO.CHASSI, dbo.OS.OS_ID, 
                         dbo.OS.ESTABELECIMENTO_ID, dbo.OS.PESSOA_ID, dbo.OS.CT_ID, dbo.OS.DT_OS, dbo.OS.DT_FECHA, dbo.OS.TIPO_OS, dbo.OS.SITUACAO_OS, dbo.OS.KM, dbo.OS.CLIENTE, dbo.PESSOA.CNPJCPF, 
                         dbo.PESSOA.DESCRICAO AS DESCRICAOPESSOA, dbo.PESSOA.RAZAO, dbo.PESSOA.SITUACAO AS SITUACAOPESSOA, dbo.OSEQUIPAMENTO.NOME_CLIENTE, dbo.OSVEICULO.NUMR_FROTA, 
                         dbo.OSVEICULO.ANO AS AnoVeiculo, dbo.OSVEICULO.MODELO AS ModeloVeiculo, dbo.OSVEICULO.COR_ID AS CorVeiculo, dbo.OSVEICULO.TIPO_VEICULO_ID, dbo.OSVEICULO.MARCA_ID AS MarcaVeiculo
FROM            dbo.OSEQUIPAMENTO WITH (NOLOCK) RIGHT OUTER JOIN
                         dbo.OS WITH (NOLOCK) LEFT OUTER JOIN
                         dbo.OSVEICEQP WITH (NOLOCK) ON dbo.OS.OS_ID = dbo.OSVEICEQP.OS_ID LEFT OUTER JOIN
                         dbo.PESSOA WITH (NOLOCK) ON dbo.OS.PESSOA_ID = dbo.PESSOA.PESSOA_ID LEFT OUTER JOIN
                         dbo.OSVEICULO WITH (NOLOCK) ON dbo.OSVEICEQP.VEICULO_ID = dbo.OSVEICULO.VEICULO_ID ON dbo.OSEQUIPAMENTO.EQUIPAMENTO_ID = dbo.OSVEICEQP.EQUIPAMENTO_ID

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[52] 4[3] 2[29] 3) )"
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
         Begin Table = "OSEQUIPAMENTO"
            Begin Extent = 
               Top = 140
               Left = 964
               Bottom = 270
               Right = 1242
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "OS"
            Begin Extent = 
               Top = 0
               Left = 368
               Bottom = 130
               Right = 578
            End
            DisplayFlags = 280
            TopColumn = 1
         End
         Begin Table = "OSVEICEQP"
            Begin Extent = 
               Top = 2
               Left = 697
               Bottom = 115
               Right = 894
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "PESSOA"
            Begin Extent = 
               Top = 8
               Left = 20
               Bottom = 138
               Right = 217
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "OSVEICULO"
            Begin Extent = 
               Top = 0
               Left = 1054
               Bottom = 130
               Right = 1251
            End
            DisplayFlags = 280
            TopColumn = 9
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
     ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwOS'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'    Table = 1170
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwOS'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwOS'
GO


