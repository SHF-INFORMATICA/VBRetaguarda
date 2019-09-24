VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedidoConsultaSimples 
   Caption         =   "Consulta Pedido Venda"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoConsultaSimples.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbAuxForma 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   48
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbCPUaux 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbVend 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   22
      Top             =   2280
      Width           =   3585
   End
   Begin VB.ComboBox cmbForma 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   21
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtTotalVenda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10170
      TabIndex        =   20
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   19
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtCli 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      MaxLength       =   100
      TabIndex        =   18
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbSITUACAO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cmbVendAux 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNOTA 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtCupom 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox cmbSituacaoAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   8760
      TabIndex        =   12
      Top             =   2280
      Width           =   3015
   End
   Begin VB.ComboBox cmbFamiliaAUX 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtQtdeProd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1440
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.ComboBox cmbCPU 
      Height          =   360
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdConsCli 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5640
      Picture         =   "PedidoConsultaSimples.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox txtTotDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   5865
      TabIndex        =   7
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtTotVendas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7920
      TabIndex        =   6
      Top             =   6960
      Width           =   1575
   End
   Begin VB.ComboBox cmbCC 
      Height          =   360
      Left            =   7080
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cmbCCAux 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10320
      TabIndex        =   3
      Top             =   2040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1270
      ButtonWidth     =   3519
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir Tela"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Pedido"
            Key             =   "pedido"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7680
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaSimples.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaSimples.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaSimples.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaSimples.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaSimples.frx":A8FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaSimples.frx":C8DF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   360
      Left            =   3600
      TabIndex        =   24
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11865
      DesignHeight    =   7695
   End
   Begin MSComctlLib.ListView lstPedidoItem 
      Height          =   1905
      Left            =   45
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   3360
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648384
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2252
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   7508
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   1668
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valr.Unitário"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Desconto"
         Object.Width           =   1877
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valr.Total"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "NCM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Pedido"
         Object.Width           =   1502
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3015
      Left            =   0
      TabIndex        =   26
      Top             =   3240
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   5318
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      Appearance      =   1
      MousePointer    =   99
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cupom"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NFe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr.Venda"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Vlr.Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Dt.Emisão"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Faturamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ID"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "NºCaixa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "TpRegistro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Comanda"
         Object.Width           =   2540
      EndProperty
   End
   Begin Threed.SSOption optSintetico 
      Height          =   255
      Left            =   10680
      TabIndex        =   27
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sintético"
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin Threed.SSOption optAnalitico 
      Height          =   255
      Left            =   10680
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Analítico"
      Value           =   -1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   7440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor(a):"
      Height          =   240
      Left            =   105
      TabIndex        =   46
      Top             =   2280
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      Height          =   240
      Left            =   60
      TabIndex        =   45
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.TotalVendas"
      Height          =   240
      Index           =   0
      Left            =   7905
      TabIndex        =   44
      Top             =   6600
      Width           =   1515
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos"
      Height          =   240
      Left            =   150
      TabIndex        =   43
      Top             =   6600
      Width           =   765
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   240
      Left            =   540
      TabIndex        =   42
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   240
      Left            =   2820
      TabIndex        =   41
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   9000
      TabIndex        =   40
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N.F.:"
      Height          =   240
      Left            =   5640
      TabIndex        =   39
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ECF:"
      Height          =   240
      Left            =   5610
      TabIndex        =   38
      Top             =   2280
      Width           =   435
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   240
      Left            =   120
      TabIndex        =   37
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   240
      Left            =   3360
      TabIndex        =   36
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Família:"
      Height          =   240
      Left            =   7920
      TabIndex        =   35
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos.Vendidos"
      Height          =   240
      Left            =   1470
      TabIndex        =   34
      Top             =   6600
      Width           =   1785
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Estação:"
      Height          =   240
      Left            =   600
      TabIndex        =   33
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "TotalDesconto"
      Height          =   240
      Left            =   5820
      TabIndex        =   32
      Top             =   6600
      Width           =   1350
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.Faturado"
      Height          =   240
      Index           =   1
      Left            =   10200
      TabIndex        =   31
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CentroCusto:"
      Height          =   240
      Left            =   5805
      TabIndex        =   30
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmPedidoConsultaSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CARREGA_COMBOS
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lstPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  CRITERIO = ""
                  CRITERIO = Trim(InputBox("Informe CPF/CNPJ do cliente", "Atualização de Dados Pedido Venda", CRITERIO))

                  If Trim(CRITERIO) <> "" Then
                     SQL = "update PEDIDO set "
                     SQL = SQL & " cgccpf = '" & Trim(CRITERIO) & "'"
                     SQL = SQL & " where pedido_id = " & lstPedido.SelectedItem.Text
                     CONECTA_RETAGUARDA.Execute SQL
                  End If

                  SQL = ""
                  CRITERIO = ""
                  MONTA_CONSULTA_SQL True
                  SETA_GRID
               End If
            End If
         End If
      Case vbKeyF6
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
                     CRITERIO = ""
                     frmPedidoCancela.txtPedido.Text = 0 & lstPedido.SelectedItem.Text
                     frmPedidoCancela.Show 1
                     SQL = ""
                     CRITERIO = ""
                     MONTA_CONSULTA_SQL True
                     SETA_GRID
                     Else: MsgBox "Não permitido."
                  End If
               End If
            End If
         End If
      Case vbKeyF7
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.ListItems.Clear

            SQL = "SELECT PRODUTO.CODG_PRODUTO, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "
            SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.PRECO_CUSTO, pedidoitem.seq_id,"
            SQL = SQL & " PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.CFOP_id, pedidoitem.status, "
            SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, Produto.FORNECEDOR_ID"
            SQL = SQL & " FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
            SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

            SQL = SQL & " where pedidoitem.produto_id = produto.produto_id "
            SQL = SQL & " and PEDIDO_ID = " & lstPedido.SelectedItem.ListSubItems.Item(11).Text
            SQL = SQL & " and tipo_reg = 'PC' "
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               MOSTRA_TOP "Duplo Click no grid ocultar", " ", " ", " ", ""
               lstPedidoItem.Visible = True
            End If
            While Not TabTemp.EOF
               VALOR_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
               VALOR_ITEM_N = TabTemp.Fields("qtd_pedida").Value * (TabTemp.Fields("valor_item").Value - VALOR_DESCONTO_N)

               Set Item = lstPedidoItem.ListItems.Add(, "seq." & TabTemp.Fields("seq_id").Value, Trim(TabTemp.Fields("codg_produto").Value))
               Item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
               Item.SubItems(2) = "" & Format(Trim(TabTemp.Fields("qtd_pedida").Value), strFormatacao3Digitos)
               Item.SubItems(3) = "" & Format(Trim(TabTemp.Fields("valor_item").Value), strFormatacao2Digitos)
               Item.SubItems(4) = "" & Format(Trim(TabTemp.Fields("valor_desconto").Value), strFormatacao2Digitos)
               Item.SubItems(5) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
               Item.SubItems(6) = "" & Trim(TabTemp.Fields("CODG_ncm").Value)
               Item.SubItems(7) = "" & lstPedido.SelectedItem.Text

               If Trim(TabTemp.Fields("status").Value) = "A" Then
                  Item.ForeColor = vbBlue
                  Item.ListSubItems(1).ForeColor = vbBlue
                  Item.ListSubItems(2).ForeColor = vbBlue
                  Item.ListSubItems(3).ForeColor = vbBlue
                  Item.ListSubItems(4).ForeColor = vbBlue
                  Item.ListSubItems(5).ForeColor = vbBlue
                  Item.ListSubItems(6).ForeColor = vbBlue
               End If
               If Trim(TabTemp.Fields("status").Value) = "P" Then
                  Item.ForeColor = vbBlack
                  Item.ListSubItems(1).ForeColor = vbBlack
                  Item.ListSubItems(2).ForeColor = vbBlack
                  Item.ListSubItems(3).ForeColor = vbBlack
                  Item.ListSubItems(4).ForeColor = vbBlack
                  Item.ListSubItems(5).ForeColor = vbBlack
                  Item.ListSubItems(6).ForeColor = vbBlack
               End If
               If Trim(TabTemp.Fields("status").Value) = "C" Then
                  Item.ForeColor = vbRed
                  Item.ListSubItems(1).ForeColor = vbRed
                  Item.ListSubItems(2).ForeColor = vbRed
                  Item.ListSubItems(3).ForeColor = vbRed
                  Item.ListSubItems(4).ForeColor = vbRed
                  Item.ListSubItems(5).ForeColor = vbRed
                  Item.ListSubItems(6).ForeColor = vbRed
               End If
               TabTemp.MoveNext
               CRITERIO = ""
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.Refresh
         End If
      Case vbKeyF11
         'frmSenha.Show 1

         'If UCase(CRITERIO) = UCase("acerto") Then
            PEDIDO_ID_N = 0
            If Not IsNull(lstPedido.SelectedItem.Text) Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  PEDIDO_ID_N = lstPedido.SelectedItem.Text

                  frmPedidoClienteAcerto.Show 1
                  MONTA_CONSULTA_SQL True
                  SETA_GRID
                  PEDIDO_ID_N = 0
               End If
            End If
         'End If
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoVENDA_KeyDown"
End Sub

Private Sub cmdConsCli_Click()
   frmDISPLAYCLIENTE.Show 1
   If Trim(CNPJCPF_A) <> "" Then _
      txtCGCCPF.Text = CNPJCPF_A
   CNPJCPF_A = ""
   txtCGCCPF.SetFocus
End Sub

Private Sub txtCGCCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   If txtCGCCPF.Text = "" Then _
      txtCGCCPF.Text = "99999999999"

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome,cliente_id from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where CGCCPF = '" & Trim(txtCGCCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCliente.EOF Then
      Beep
      MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
      txtCGCCPF.SetFocus
      Exit Sub
      Else:
         CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
         If TabCliente!NOME <> "" Then _
            txtCli.Text = TabCliente!NOME
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_LostFocus"
End Sub

Private Sub txtDtFim_LostFocus()
   CHECA_ULTIMO_DIA_MES
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub cmbcpu_Click()
'On Error GoTo ERRO_TRATA

   cmbCPUaux.ListIndex = cmbCPU.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbcpu_Click"
End Sub

Private Sub txtReg_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtQtdeProd_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtTotalVenda_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtTotDesconto_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtTotVendas_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      MONTA_CONSULTA_SQL True
      SETA_GRID
      'SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtpedido_KeyPress"
End Sub

Private Sub lstPedido_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedido.SelectedItem.Text) Then
      CRITERIO = lstPedido.SelectedItem.Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_DblClick"
End Sub

Private Sub lstpedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedido, ColumnHeader
End Sub

Private Sub lstPedido_Click()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 0
   If Not IsNull(lstPedido.SelectedItem.Text) Then _
      If IsNumeric(lstPedido.SelectedItem.Text) Then _
         PEDIDO_ID_N = lstPedido.SelectedItem.Text

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_Click"
End Sub

Private Sub lstPedidoitem_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.Visible = False

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoitem_DblClick"
End Sub

Private Sub lstpedidoitem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedidoItem, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         MONTA_CONSULTA_SQL True
         SETA_GRID
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         CRITERIO = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
      Case "pedido"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False
         FORMULA_REL = ""
         If Not IsNull(lstPedido.SelectedItem.ListSubItems.Item(11).Text) Then
            FORMULA_REL = lstPedido.SelectedItem.ListSubItems.Item(11).Text

            If Not IsNumeric(FORMULA_REL) Then _
               Exit Sub

            PEDIDO_ID_N = FORMULA_REL

            FORMULA_REL = "{vwRelVenda.empresa_id} = " & EMPRESA_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & PEDIDO_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.status_item} <> 'C' "

            If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_pedido_venda.rpt"
            If CNPJ_GERAL = "15333554000188" Then _
               Nome_Relatorio = "pedido_shf.rpt"
            frmRELATORIO10.Show 1
         End If
      Case "print"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False

         MONTA_CONSULTA_SQL False
         GERA_REL

   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxForma.ListIndex = cmbForma.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFORMA_Click"
End Sub

Private Sub cmbSituacao_Click()
'On Error GoTo ERRO_TRATA

   cmbSituacaoAUX.ListIndex = cmbSITUACAO.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbsituacao_Click"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_Click"
End Sub

Private Sub TXTCGCCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""

   If Trim(CNPJCPF_A) <> "" Then
      txtCGCCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_GotFocus"
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYCLIENTE.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCGCCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_KeyDown"
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_KeyPress"
End Sub

Private Sub txtDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDTfim_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub cmbCC_Click()
On Error Resume Next

   cmbCCAux.ListIndex = cmbCC.ListIndex
   Call cmbCC_LostFocus
End Sub

Private Sub cmbCC_LostFocus()

   If Trim(cmbCC.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from DESCR d WITH (NOLOCK)"
      SQL = SQL & " where TIPO = 'O' "
      SQL = SQL & " and codigo = '" & Trim(cmbCCAux.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         cmbCC.Text = "" & TabTemp.Fields("descricao").Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   cmbCC.Text = ""
   cmbCCAux.Text = ""
   cmbCPU.Text = ""
   cmbCPUaux.Text = ""
   cmbSITUACAO.Text = ""
   cmbSituacaoAUX.Text = ""

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False
   PRODUTO_ID_N = 0
   txtNOTA.Text = ""
   txtCupom.Text = ""
   lstPedido.ListItems.Clear
   txtPedido.Text = ""
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   txtCli.Text = ""

   If cmbVend.Enabled = True Then _
      cmbVend.Text = ""

   cmbForma.Text = ""
   cmbAuxForma.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtTotalVenda.Text = ""
   txtTotVendas.Text = ""
   txtTotDesconto.Text = ""
   txtReg.Text = ""
   txtQtdeProd.Text = ""

   lstPedido.Visible = True
   optSintetico.Value = True
   txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = lstPedido.SelectedItem.Text
   CNPJCPF_A = ""

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status, cgccpf from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      CNPJCPF_A = "" & Trim(TabCABECA.Fields("cgccpf").Value)
      If Not IsNull(TabCABECA!Status) Then
         If TabCABECA!Status <> "9" Then
            If Trim(CNPJCPF_A) = "99999999999" Then
               Msg = "Para geração de nota fiscal eletrônica, os dados do cliente devem ser cadastrados, deseja continuar essa operação ?"
               PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CNPJCPF_A = ""
                  frmDISPLAYCLIENTE.Show 1
                  If Trim(CNPJCPF_A) <> "" Then
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

SQL = "select nome,cgccpf from CLIENTE WITH (NOLOCK)"
SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_A) & "'"
TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
If TabConsulta.EOF Then
   Msg = "CNPF/CPF informado não cadastrado, deseja realizar cadastro de cliente agora ?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      frmCADASTROCLIENTE.Show 1
      MsgBox "Repetir operação."
      Else
         If TabCABECA.State = 1 Then _
            TabCABECA.Close
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Exit Sub
   End If
   Else
      If TabCABECA.State = 1 Then _
         TabCABECA.Close
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
End If
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

                     SQL = "update PEDIDO set "
                     SQL = SQL & " cgccpf = '" & Trim(CNPJCPF_A) & "'"
                     SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                     SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                     SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
                     Else
                        If TabCABECA.State = 1 Then _
                           TabCABECA.Close
                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close
                        Exit Sub
                  End If
                  Else
                     If TabCABECA.State = 1 Then _
                        TabCABECA.Close
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
                     Exit Sub
               End If
            End If

            CRITERIO = PEDIDO_ID_N
            'TIPO_NFe_GERAR = "S"
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
            frmNOTAGERA.Show 1
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Sub CHECA_ULTIMO_DIA_MES()
   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
      txtDtFim.Text = ""

      txtDtIni.PromptInclude = True
      If IsDate(txtDtIni.Text) Then
         CRITERIO = FimDoMes(DMA(txtDtIni.Text), False)
         CRITERIO = Right(CRITERIO, 2) & "/" & Mid(CRITERIO, 5, 2) & "/" & Left(CRITERIO, 4)
         txtDtFim.Text = CRITERIO
         txtDtFim.PromptInclude = True
      End If
   End If
End Sub

Sub GERA_REL()
'On Error GoTo ERRO_TRATA

   Dim VALOR_ITEM_N     As Double
   Dim Qtde_N           As Double
   Dim DESCONTO_ITEM    As Double
   Dim DESCONTO_CABEÇA  As Double
   Dim VALOR_CUSTO_N    As Double
   Dim NOME_CLIENTE_A   As String
   Dim CARTAO_ID        As Long

   VALOR_ITEM_N = 0
   Qtde_N = 0
   DESCONTO_ITEM = 0
   DESCONTO_CABEÇA = 0
   VALOR_CUSTO_N = 0
   CARTAO_ID = 0

   Me.Enabled = False
   CONT_N = 0
   CRITERIO = SQL

   If ExisteTabela("RETAGUARDA", "RELVENDA", "U") = True Then
      strsql = "drop table RELVENDA"
      CONECTA_RETAGUARDA.Execute strsql
   End If

   If ExisteTabela("RETAGUARDA", "RELVENDA", "U") = False Then
      strsql = "create table RELVENDA"
      strsql = strsql & " ("
         strsql = strsql & " RELVENDA_ID        bigint      not null,"
         strsql = strsql & " EMPRESA_ID         bigint      not null,"
         strsql = strsql & " PEDIDO_ID          bigint      not null,"
         strsql = strsql & " CLIENTE_ID         bigint      not null,"
         strsql = strsql & " VENDEDOR_ID        bigint      not null,"
         strsql = strsql & " tipovenda_ID       bigint      not null,"
         strsql = strsql & " DT_VENDA           datetime    not null,"

         strsql = strsql & " VALOR_VENDA        float       not null,"
         strsql = strsql & " VLR_TOT_CUSTO      float       null    ,"
         strsql = strsql & " VLR_TOT_DESCONTO   float       null    ,"
         strsql = strsql & " CLIENTE            varchar(50) null    ,"

         strsql = strsql & " QTDE_VENDIDA       float       not null,"
         strsql = strsql & " CARTAOBARRA_ID BIGINT,"

         strsql = strsql & " constraint PK_RELVENDA primary key (RELVENDA_ID)"
      strsql = strsql & " )"
      CONECTA_RETAGUARDA.Execute strsql
   End If

   strsql = "delete from RELVENDA"
   CONECTA_RETAGUARDA.Execute strsql

   Qtde_N = 0
   VALOR_ITEM_N = 0
   VALOR_CUSTO_N = 0
   VALOR_DESCONTO_N = 0
   NUMR_ID_N = 0
   PEDIDO_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

'============================

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL3, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then _
      CONTA_REG_PROGRESSO = TabTemp.Fields(0).Value
'============================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   ProgressBar1.Min = 0                   'Indica o valor inicial
   ProgressBar1.Max = CONTA_REG_PROGRESSO 'Indica o valor final
   CONT_N = 0

   TabTemp.Open CRITERIO, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Registro não encontrado."
      Exit Sub
   End If
   While Not TabTemp.EOF
      If CONT_N < CONTA_REG_PROGRESSO Then
         CONT_N = CONT_N + 1
         ProgressBar1.Value = CONT_N
      End If

      CARTAO_ID = 0 & TabTemp.Fields("cartaobarra_id").Value
      NOME_CLIENTE_A = Trim(TabTemp.Fields("NOME_CLIENTE").Value)
      If Trim(TabTemp.Fields("NOME_CLIENTE").Value) = "" Then
         If TabCliente.State = 1 Then _
            TabCliente.Close
      
            SQL = "select nome from CLIENTE WITH (NOLOCK)"
            SQL = SQL & " where cliente_id = " & TabTemp.Fields("cliente_id").Value
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               NOME_CLIENTE_A = Trim(TabCliente.Fields(0).Value)
      
         If TabCliente.State = 1 Then _
            TabCliente.Close
      End If

      VALOR_ITEM_N = 0 & TabTemp.Fields("valor_total").Value
      'VALOR_ITEM_N = 0 & (TabTemp.Fields("QTD_PEDIDA").Value * TabTemp.Fields("VALOR_ITEM").Value)
      Qtde_N = 0 & TabTemp.Fields("QTD_PEDIDA").Value
      DESCONTO_ITEM = 0 & TabTemp.Fields("VALOR_DESCONTO").Value
      DESCONTO_CABEÇA = 0 & TabTemp.Fields("DESCCABECA").Value
      VALOR_CUSTO_N = 0 & TabTemp.Fields("preco_custo").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from RELVENDA WITH (NOLOCK)"
      SQL = SQL & " where PEDIDO_ID = " & TabTemp.Fields("pedido_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         NUMR_ID_N = MAX_ID("RELVENDA_id", "RELVENDA", "", "", "", "")

         SQL = "insert into RELVENDA "
         SQL = SQL & "("
            SQL = SQL & " RELVENDA_ID,EMPRESA_ID,PEDIDO_ID,CLIENTE_ID,VENDEDOR_ID,DT_VENDA,"
            SQL = SQL & " VALOR_VENDA,VLR_TOT_CUSTO,VLR_TOT_DESCONTO,CLIENTE,QTDE_VENDIDA,"
            SQL = SQL & " cartaobarra_id,tipovenda_id "
         SQL = SQL & ")"
         SQL = SQL & " values("
            SQL = SQL & NUMR_ID_N                                                'RELVENDA_ID
            SQL = SQL & "," & TabTemp.Fields("EMPRESA_ID").Value                 'EMPRESA_ID
            SQL = SQL & "," & TabTemp.Fields("PEDIDO_ID").Value                  'PEDIDO_ID
            SQL = SQL & "," & TabTemp.Fields("CLIENTE_ID").Value                 'CLIENTE_ID
            SQL = SQL & "," & TabTemp.Fields("vendedor_ID").Value                'VENDEDOR_ID
            SQL = SQL & ",'" & Trim(TabTemp.Fields("dt_req").Value) & "'"         'DT_VENDA

            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                              'VALOR_VENDA
            SQL = SQL & "," & tpMOEDA(VALOR_CUSTO_N)                             'VLR_TOT_CUSTO
            SQL = SQL & "," & tpMOEDA(DESCONTO_CABEÇA)                           'VLR_TOT_DESCONTO
            SQL = SQL & ",'" & Trim(TabTemp.Fields("NOME_CLIENTE").Value) & "'"  'CLIENTE
            
            SQL = SQL & "," & tpMOEDA(Qtde_N)                                    'QTDE_VENDIDA
            SQL = SQL & "," & CARTAO_ID             'comanda
            SQL = SQL & "," & TabTemp.Fields("tipovenda_id").Value               'tipovenda_id
         SQL = SQL & ")"
         Else
            SQL = "update RELVENDA set "

            'SQL = SQL & "VALOR_VENDA = valor_venda + " & tpMOEDA(VALOR_ITEM_N)                  'VALOR_VENDA
            SQL = SQL & " VLR_TOT_CUSTO = VLR_TOT_CUSTO + " & tpMOEDA(VALOR_CUSTO_N * Qtde_N)   'VLR_TOT_CUSTO
            SQL = SQL & ",QTDE_VENDIDA = QTDE_VENDIDA + " & tpMOEDA(Qtde_N)                     'QTDE_VENDIDA

            SQL = SQL & " where PEDIDO_ID = " & TabTemp.Fields("pedido_id").Value
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      DoEvents

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   Me.Enabled = True

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If optAnalitico.Value = True Then
      Nome_Relatorio = "venda_totais_analitico.rpt"
      Else: Nome_Relatorio = "venda_totais.rpt"
   End If

   frmRELATORIO10.Show 1

   Me.Enabled = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "GERA_REL"
End Sub

Sub CARREGA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbVend.Enabled = True

      Toolbar1.Buttons(6).Visible = False
      If USA_NFe = True Then _
         Toolbar1.Buttons(6).Visible = True

      Else
         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select logon from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where usuario_id = " & USUARIO_ID_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            CRITERIO = Chr$(39) & Trim(TabUSU.Fields("logon").Value) & "%" & Chr(39)

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            SQL = "select nome_vend, vendedor_id from VENDEDOR WITH (NOLOCK)"
            SQL = SQL & " where nome_vend like " & CRITERIO
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVend.Text = TabVENDEDOR!NOME_VEND
               cmbVendAux.Text = TabVENDEDOR!vendedor_id
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_VENDEDOR"
End Sub

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

   If TRAZ_TIPO_USUARIO = 7 Then
      txtTotalVenda.Visible = False
      txtTotVendas.Visible = False
      Label7(0).Visible = False
      Label7(1).Visible = False
      Label18.Visible = False
      txtTotDesconto.Visible = False
   End If

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   cmbSITUACAO.AddItem "Todos"
   cmbSituacaoAUX.AddItem ""

   cmbSITUACAO.AddItem "Cupom Fiscal"
   cmbSituacaoAUX.AddItem "'7'"

   cmbSITUACAO.AddItem "Nota Eletrônica"
   cmbSituacaoAUX.AddItem "'7','5','3'"

   cmbSITUACAO.AddItem "Pendente"
   cmbSituacaoAUX.AddItem "'1','2','4'"

   cmbSITUACAO.AddItem "Faturado"
   cmbSituacaoAUX.AddItem "'3','5','7'"

   cmbSITUACAO.AddItem "Cancelado"
   cmbSituacaoAUX.AddItem "'9'"

   cmbSITUACAO.Text = "Faturado"
   cmbSituacaoAUX.Text = "'3','5','7'"

   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False

   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFamilia.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("familiaproduto_id").Value)
      cmbFamiliaAUX.AddItem Trim(TabDESCR.Fields("familiaproduto_id").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbCPU.Clear
   cmbCPUaux.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select distinct(numero_caixa_cpu) from PEDIDO WITH (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      If Not IsNull(TabDESCR.Fields(0).Value) Then
         cmbCPU.AddItem Trim("CAIXA") & "-" & Trim(TabDESCR.Fields(0).Value)
         cmbCPUaux.AddItem Trim(TabDESCR.Fields(0).Value)
      End If
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   If USA_ECF = False Then _
      lstPedido.ColumnHeaders(2).Width = 1
   If USA_NFe = False Then _
      lstPedido.ColumnHeaders(3).Width = 1

   cmbForma.Clear
   cmbAuxForma.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbForma.AddItem TabTemp!DESCRICAO & " - " & TabTemp!TipoVenda_ID
      cmbAuxForma.AddItem TabTemp!TipoVenda_ID
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbVend.Clear
   SQL = "select vendedor_id,nome_vend from VENDEDOR WITH (NOLOCK)"
   SQL = SQL & " order by nome_vend "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbVend.AddItem Trim(TabTemp!NOME_VEND) & " - " & Trim(TabTemp!vendedor_id)
      cmbVendAux.AddItem Trim(TabTemp!vendedor_id)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbVend.Enabled = False
   Toolbar1.Buttons(6).Visible = False

   CARREGA_VENDEDOR

   If INDR_PANIFIC = False Then
      If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
         cmbVend.Enabled = True
         cmbVend.Text = ""
      End If

      MONTA_CONSULTA_SQL True
      SETA_GRID
   End If

   Me.Enabled = True
   Me.KeyPreview = True
   VALOR_TOTAL_N = 0

   cmbCCAux.Clear
   cmbCC.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'O'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCC.AddItem Trim(TabTemp!DESCRICAO)
      cmbCCAux.AddItem TabTemp!codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   CHECA_ULTIMO_DIA_MES

   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False

   VALOR_TOTAL_N = 0
   If Indr_Consulta = True Then
      txtTotalVenda.Text = ""
      txtReg.Text = ""
      txtQtdeProd.Text = ""
   End If
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SqL2 = "SELECT  * "
   SQL3 = "SELECT count(vwCONSULTA_PEDIDO.cliente_ID) "

   SQL = " FROM vwCONSULTA_PEDIDO WITH (NOLOCK) "

   SQL = SQL & " where pedido_id Is Not Null"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(cmbCPUaux.Text) <> "" Then _
      If IsNumeric(cmbCPUaux.Text) Then _
         SQL = SQL & " and numero_caixa_cpu = " & cmbCPUaux.Text

   If Trim(txtCupom.Text) <> "" Then _
      SQL = SQL & " and numr_cupom = " & txtCupom.Text

   If Trim(txtNOTA.Text) <> "" Then _
      SQL = SQL & " and numr_nota = " & txtNOTA.Text

   If Trim(txtPedido.Text) <> "" Then _
      SQL = SQL & " and pedido_id = " & txtPedido.Text

   txtCGCCPF.PromptInclude = False
   If Trim(txtCGCCPF.Text) <> "" Then _
      If CLIENTE_ID_N > 0 Then _
         SQL = SQL & " and cliente_id = " & CLIENTE_ID_N
         'SQL = SQL & " and cgccpf = '" & Trim(txtCGCCPF.Text) & "'"
   txtCGCCPF.PromptInclude = True

   If Trim(cmbVend.Text) <> "" Then _
      SQL = SQL & " and vendedor_id = " & cmbVendAux.Text

   If Trim(cmbSituacaoAUX.Text) <> "" Then
      If Trim(cmbSituacaoAUX.Text) = "'7','5','3'" Then _
         SQL = SQL & " and numr_nota > 0 "

      SQL = SQL & " and SIT_PEDIDO in (" & Trim(cmbSituacaoAUX.Text) & ")"
   End If

   If Trim(cmbAuxForma.Text) <> "" Then _
      SQL = SQL & " and tipovenda_id = " & cmbAuxForma.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_req <= '" & txtDtFim.Text & "'"
   End If

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   SQL3 = SQL3 & " " & SQL

   SQL = SQL & " order by PEDIDO_ID desc"

   SQL = SqL2 & " " & SQL

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   HORA_INI = Time

   Dim TabTemp          As New ADODB.Recordset
   Dim Conta_Produto_N  As Long

   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0

   Me.Enabled = False
   Me.KeyPreview = False

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
'MsgBox SQL3
   TabTemp.Open SQL3, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then _
      CONTA_REG_PROGRESSO = TabTemp.Fields(0).Value
'============================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If CONTA_REG_PROGRESSO > 0 Then
      ProgressBar1.Min = 0                   'Indica o valor inicial
      ProgressBar1.Max = CONTA_REG_PROGRESSO 'Indica o valor final
      'frmProgresso.Show 1
   End If
   CONT_N = 0

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then
      While Not TabTemp.EOF
         DoEvents

         If CONT_N < CONTA_REG_PROGRESSO Then
            CONT_N = CONT_N + 1
            ProgressBar1.Value = CONT_N
         End If

         If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
            CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
            txtReg.Text = CONTA_REGISTRO_N
            txtReg.Refresh

            PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value

            NUMR_SEQ_N = NUMR_SEQ_N + 1
            Set Item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp.Fields("PEDIDO_ID").Value)

            Item.SubItems(11) = "" & TabTemp.Fields("PEDIDO_ID").Value
            Item.SubItems(1) = "" & TabTemp.Fields("numr_cupom").Value
            Item.SubItems(2) = "" & TabTemp.Fields("numr_nota").Value
            Item.SubItems(3) = "" & Trim(TabTemp!NOME_CLIENTE) & " - " & Trim(TabTemp.Fields("CNPJCPF").Value)

            If IsNull(TabTemp!NOME_CLIENTE) Or Trim(TabTemp!NOME_CLIENTE) = "" Then
               If TabCliente.State = 1 Then _
                  TabCliente.Close

               SQL = "select nome from CLIENTE WITH (NOLOCK)"
               SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"
               TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCliente.EOF Then _
                  Item.SubItems(3) = "" & TabCliente!NOME

               If TabCliente.State = 1 Then _
                  TabCliente.Close
            End If

            Item.SubItems(7) = TabTemp!DT_REQ
            Item.SubItems(8) = ""

            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            SQL = "select * from TIPOVENDA WITH (NOLOCK)"
            SQL = SQL & " where tipovenda_id = " & TabTemp!TipoVenda_ID
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then _
               Item.SubItems(8) = TabDESCR!DESCRICAO
            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            If TabUSU.State = 1 Then _
               TabUSU.Close

            Item.SubItems(9) = ""
   
            SQL = "select * from VENDEDOR WITH (NOLOCK)"
            SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id").Value
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               Item.SubItems(9) = TabUSU!NOME_VEND
            If TabUSU.State = 1 Then _
               TabUSU.Close

            Item.SubItems(10) = ""

            If Not IsNull(TabTemp.Fields("Status")) Then
               If TabTemp.Fields("Status") = 2 Then
                  If TabTemp.Fields("tipo_registro") = "O" Then
                     Item.SubItems(10) = "Orcamento"
                     Else: Item.SubItems(10) = "Pedido"
                  End If
               End If
               If TabTemp.Fields("Status").Value = 3 Then _
                  Item.SubItems(10) = "3-Faturado"
               If TabTemp.Fields("Status").Value = 4 Then _
                  Item.SubItems(10) = "4-Cupom"
               If TabTemp.Fields("Status").Value = 5 Then _
                  Item.SubItems(10) = "5-Faturado"
               If TabTemp.Fields("Status").Value = 7 Then _
                  Item.SubItems(10) = "7-Cupom Fiscal"
               If TabTemp.Fields("Status").Value = 9 Then _
                  Item.SubItems(10) = "9-Cancelado"
            End If

            If Not IsNull(TabTemp.Fields("numero_caixa_cpu").Value) Then _
               Item.SubItems(12) = TabTemp.Fields("numero_caixa_cpu").Value

            Item.SubItems(13) = TabTemp.Fields("tipo_registro").Value

            VALOR_DESCONTO_N = 0
            'VALOR_TOTAL_DESCONTO_N = 0

            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPedidoItem.EOF Then _
               If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
                  VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

           If Not IsNull(TabTemp.Fields("desccabeca").Value) Then _
              VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabTemp.Fields("desccabeca").Value

            VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

            'BUSCA VALOR TOTAL VENDA
            VALOR_ITEM_N = 0

            SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            'SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            SQL = "select sum(qtd_pedida) FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               Conta_Produto_N = Conta_Produto_N + TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            txtQtdeProd.Text = Conta_Produto_N
            txtQtdeProd.Refresh
            'VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N
            VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

            txtTotDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
            txtTotVendas.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
            txtTotalVenda.Text = Format(VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

            Item.SubItems(4) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
            Item.SubItems(5) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
            Item.SubItems(6) = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

            If TabTemp.Fields("SIT_PEDIDO").Value = 1 Then
               Item.ForeColor = vbRed
               Item.ListSubItems(1).ForeColor = vbRed
               Item.ListSubItems(2).ForeColor = vbRed
               Item.ListSubItems(3).ForeColor = vbRed
               Item.ListSubItems(4).ForeColor = vbRed
               Item.ListSubItems(5).ForeColor = vbRed
               Item.ListSubItems(6).ForeColor = vbRed
               Item.ListSubItems(7).ForeColor = vbRed
               Item.ListSubItems(8).ForeColor = vbRed
               Item.SubItems(10) = "" & "Em Aberto - 1"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 2 Then
               Item.ForeColor = vbBlue
               Item.ListSubItems(1).ForeColor = vbBlue
               Item.ListSubItems(2).ForeColor = vbBlue
               Item.ListSubItems(3).ForeColor = vbBlue
               Item.ListSubItems(4).ForeColor = vbBlue
               Item.ListSubItems(5).ForeColor = vbBlue
               Item.ListSubItems(6).ForeColor = vbBlue
               Item.ListSubItems(7).ForeColor = vbBlue
               Item.ListSubItems(8).ForeColor = vbBlue
               Item.ListSubItems(9).ForeColor = vbBlue
               Item.SubItems(10) = "" & "A Faturar - 2"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 3 Then
               Item.ForeColor = vbBlack
               Item.ListSubItems(1).ForeColor = vbBlack
               Item.ListSubItems(2).ForeColor = vbBlack
               Item.ListSubItems(3).ForeColor = vbBlack
               Item.ListSubItems(4).ForeColor = vbBlack
               Item.ListSubItems(5).ForeColor = vbBlack
               Item.ListSubItems(6).ForeColor = vbBlack
               Item.ListSubItems(7).ForeColor = vbBlack
               Item.ListSubItems(8).ForeColor = vbBlack
               Item.ListSubItems(9).ForeColor = vbBlack
               Item.SubItems(10) = "" & "Faturado - 3"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 5 Then
               Item.ForeColor = vbBlack
               Item.ListSubItems(1).ForeColor = vbBlack
               Item.ListSubItems(2).ForeColor = vbBlack
               Item.ListSubItems(3).ForeColor = vbBlack
               Item.ListSubItems(4).ForeColor = vbBlack
               Item.ListSubItems(5).ForeColor = vbBlack
               Item.ListSubItems(6).ForeColor = vbBlack
               Item.ListSubItems(7).ForeColor = vbBlack
               Item.ListSubItems(8).ForeColor = vbBlack
               Item.ListSubItems(9).ForeColor = vbBlack
               Item.SubItems(10) = "" & "Faturado - 5"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 6 Then
               Item.ForeColor = vbBlack
               'Item.ListSubItems(1).ForeColor = vbYellow
               'Item.ListSubItems(2).ForeColor = vbYellow
               'Item.ListSubItems(3).ForeColor = vbYellow
               'Item.ListSubItems(4).ForeColor = vbYellow
               'Item.ListSubItems(5).ForeColor = vbYellow
               'Item.ListSubItems(6).ForeColor = vbYellow
               'Item.ListSubItems(7).ForeColor = vbYellow
               'Item.ListSubItems(8).ForeColor = vbYellow
               Item.ListSubItems(10).ForeColor = vbYellow
               Item.SubItems(10) = "" & "Não Contabilizado"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 7 Then
               Item.ForeColor = vbMagenta
               Item.ListSubItems(1).ForeColor = vbMagenta
               Item.ListSubItems(2).ForeColor = vbMagenta
               Item.ListSubItems(3).ForeColor = vbMagenta
               Item.ListSubItems(4).ForeColor = vbMagenta
               Item.ListSubItems(5).ForeColor = vbMagenta
               Item.ListSubItems(6).ForeColor = vbMagenta
               Item.ListSubItems(7).ForeColor = vbMagenta
               Item.ListSubItems(8).ForeColor = vbMagenta
               Item.ListSubItems(9).ForeColor = vbMagenta
               Item.SubItems(10) = "" & "Cupom Fiscal - 7"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 9 Then
               'Item.ForeColor = &HC0E0FF '&HC0C0C0
               'Item.ListSubItems(1).ForeColor = &HC0E0FF
               'Item.ListSubItems(2).ForeColor = &HC0E0FF
               'Item.ListSubItems(3).ForeColor = &HC0E0FF
               'Item.ListSubItems(4).ForeColor = &HC0E0FF
               'Item.ListSubItems(5).ForeColor = &HC0E0FF
               'Item.ListSubItems(6).ForeColor = &HC0E0FF
               'Item.ListSubItems(7).ForeColor = &HC0E0FF
               Item.ListSubItems(8).ForeColor = &HC0E0FF
               Item.ListSubItems(9).ForeColor = &HC0E0FF
               Item.ListSubItems(10).ForeColor = &HC0E0FF
               Item.SubItems(10) = "" & "Cancelado - 9"
            End If
         End If

         Item.ListSubItems(2).ForeColor = vbRed
         Item.ListSubItems(1).ForeColor = vbRed

'verificando se é venda com comanda eletronica
         Item.SubItems(14) = "" & TabTemp.Fields("cartaobarra_id").Value

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select cartaobarra_id from PEDIDOTEMP WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("PEDIDO_ID").Value
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            'Item.ForeColor = vbBlack
            'Item.ListSubItems(1).ForeColor = vbBlack
            'Item.ListSubItems(2).ForeColor = vbBlack
            'Item.ListSubItems(3).ForeColor = vbBlack
            'Item.ListSubItems(4).ForeColor = vbBlack
            'Item.ListSubItems(5).ForeColor = vbBlack
            'Item.ListSubItems(6).ForeColor = vbBlack
            'Item.ListSubItems(7).ForeColor = vbBlack
            Item.ListSubItems(8).ForeColor = vbBlack
            Item.ListSubItems(9).ForeColor = vbBlack
            Item.SubItems(14) = "" & TabDESCR.Fields(0).Value
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         TabTemp.MoveNext
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

'vbBlack  vbRed  vbGreen  vbYellow  vbBlue  vbMagenta  vbCyan  vbWhite
Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub


