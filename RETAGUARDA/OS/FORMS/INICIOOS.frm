VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{EDF439C0-99E5-11CF-AFF3-004005100200}#8.0#0"; "PVMarq.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmINICIO 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administração e controle de ordem de Serviço"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INICIOOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8550
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmdOSAbre 
      Height          =   615
      Left            =   45
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":5C12
      Caption         =   "&Ordem de Serviço"
      PictureAlignment=   1
   End
   Begin PVMarqueeLib.PVMarquee PVMarquee1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _Version        =   524288
      _ExtentX        =   15055
      _ExtentY        =   873
      _StockProps     =   29
      Text            =   "Adminitração de Ordem de Serviço"
      ForeColor       =   16711680
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TickIncrement   =   2
      Frame           =   3
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Text            =   "Adminitração de Ordem de Serviço"
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   8415
      _Version        =   524288
      _ExtentX        =   14843
      _ExtentY        =   53
      _StockProps     =   8
      ForeColor       =   -2147483647
      LineWidth       =   3
      ShadowHorizontal=   0
      ShadowVertical  =   0
      Transparent     =   -1  'True
   End
   Begin MSComctlLib.StatusBar BARI 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   5640
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "INICIOOS.frx":6064
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgOficina 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":64B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":7A85
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":8C37
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":9C85
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":B35A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":C782
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":CBD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":DCB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":FE59
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":11E3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":136D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":14AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":14DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":16600
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":17CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIOOS.frx":1DF50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   8550
      DesignHeight    =   6015
   End
   Begin Threed.SSCommand cmdVeiculo 
      Height          =   615
      Left            =   45
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":1F2B0
      Caption         =   "&Cadastro de Veículo"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdServico 
      Height          =   615
      Left            =   45
      TabIndex        =   3
      Top             =   2040
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":1F92B
      Caption         =   "Cadastra Servi&ço"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdProduto 
      Height          =   615
      Left            =   45
      TabIndex        =   4
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":20EFA
      Caption         =   "Cadastra &Produto"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdCliente 
      Height          =   615
      Left            =   45
      TabIndex        =   5
      Top             =   3480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":2134C
      Caption         =   "Cadastra &Cliente"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdParam 
      Height          =   615
      Left            =   45
      TabIndex        =   6
      Top             =   4200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":226AC
      Caption         =   "Cadastra Paramet&ros"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdUsuario 
      Height          =   615
      Left            =   45
      TabIndex        =   7
      Top             =   4920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":23EBE
      Caption         =   "Cadastra &Usuário"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   615
      Left            =   4320
      TabIndex        =   13
      Top             =   4920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":24534
      Caption         =   "&Sair"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdConsultaOS 
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":2595C
      Caption         =   "Consulta &Ordem Serviço"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdPagar 
      Height          =   615
      Left            =   4320
      TabIndex        =   12
      Top             =   3480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":26A67
      Caption         =   "Contas a P&agar"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdReceber 
      Height          =   615
      Left            =   4320
      TabIndex        =   11
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":27023
      Caption         =   "Contas a &Receber"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdCaixa 
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":28333
      Caption         =   "&Faturamento"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdOSFecha 
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":28886
      Caption         =   "&Fechar Ordem de Serviço"
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdCadVendedor 
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   4200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "INICIOOS.frx":29A11
      Caption         =   "Cadastra &Vendedor"
      PictureAlignment=   1
   End
End
Attribute VB_Name = "frmINICIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   INICIALIZA_SISTEMA

   If INDR_OS_VEICULO = False Then
      cmdVeiculo.Caption = "Cadastro de Equipamento"
      cmdOSFecha.Visible = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         frmCADASTROCLIENTE.Show 1
      Case vbKeyF3
         frmCADASTROFORNECEDOR.Show 1
      Case vbKeyF4
         frmCADASTROPRODUTO.Show 1
      Case vbKeyF6
         frmDISPLAYEMISSOR.Show 1
      Case vbKeyF8
         CRITERIO = InputBox("Entre com a senha", "Atualizações Banco de dados e Tabelas")
         If UCase(CRITERIO) = UCase("vacaveia") Then _
            frmATUALIZACAO.Show 1
         If UCase(CRITERIO) = UCase("PROTEC") Then _
            frmPROTECAO.Show 1
      Case vbKeyF12
         End
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub cmdConsultaOS_Click()
   frmOSCONSULTA.Show 1
End Sub

Private Sub cmdOSAbre_Click()
   SINAL_INDICADOR_N = 1
   If INDR_OS_VEICULO = False Then
      frmOSSERVIÇO.Show 1
      Else: frmOSVEICULO.Show 1
   End If
   SINAL_INDICADOR_N = 0
   MOSTRA_RODAPE "", "", "", "", ""
End Sub

Private Sub cmdOSFecha_Click()
   SINAL_INDICADOR_N = 2
   frmOSVEICULO.Show 1
   SINAL_INDICADOR_N = 0
End Sub

Private Sub cmdVeiculo_Click()
   If INDR_OS_VEICULO = False Then
      frmOSEQPCADASTRO.Show 1
      Else: frmOSVEICULOCADASTRO.Show 1
   End If
End Sub

Private Sub cmdSERVIcO_Click()
   frmOSSERVICOCADASTRO.Show 1
End Sub

Private Sub cmdProduto_Click()
   frmCADASTROPRODUTO.Show 1
End Sub

Private Sub cmdCLIENTE_Click()
   frmCADASTROCLIENTE.Show 1
End Sub

Private Sub cmdParam_Click()
   frmCADASTROPARAMETRO.Show 1
End Sub

Private Sub cmdUsuario_Click()
   frmCADASTROUSUARIO.Show 1
End Sub

Private Sub cmdCadVendedor_Click()
   frmCADASTROVENDEDOR.Show 1
End Sub

Private Sub cmdReceber_Click()
   SINAL_INDICADOR_N = 1
   frmFINGERALANC.Show 1
End Sub

Private Sub cmdPagar_Click()
   SINAL_INDICADOR_N = 2
   frmFINGERALANC.Show 1
End Sub

Private Sub cmdSair_Click()
   End
End Sub
