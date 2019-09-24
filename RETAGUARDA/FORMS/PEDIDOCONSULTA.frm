VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoConsulta 
   Caption         =   "Consulta Pedido Venda"
   ClientHeight    =   7290
   ClientLeft      =   1860
   ClientTop       =   2580
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
   Icon            =   "PEDIDOCONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkImp 
      Caption         =   "Imp"
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
      Left            =   5160
      TabIndex        =   58
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtPeso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3600
      TabIndex        =   56
      Top             =   6840
      Width           =   1335
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
      TabIndex        =   55
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbCC 
      Height          =   360
      Left            =   7080
      TabIndex        =   53
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtTotVendas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7920
      TabIndex        =   52
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtTotDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   5865
      TabIndex        =   49
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtComanda 
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
      Left            =   5160
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdConsCli 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5640
      Picture         =   "PEDIDOCONSULTA.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   840
      Width           =   405
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
      Left            =   4200
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbCPU 
      Height          =   360
      Left            =   4200
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmbEstabAUX 
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
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbEstab 
      Height          =   360
      Left            =   1440
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtQtdeProd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1440
      TabIndex        =   41
      Top             =   6840
      Width           =   1335
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
      Left            =   9000
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   9000
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3315
      Picture         =   "PEDIDOCONSULTA.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2280
      Width           =   405
   End
   Begin VB.TextBox txtProduto 
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
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtDescProd 
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
      Left            =   3780
      MaxLength       =   100
      TabIndex        =   30
      Top             =   2280
      Width           =   3945
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
      TabIndex        =   29
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
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
      Left            =   6600
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
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
      Left            =   6600
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
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
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
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
      TabIndex        =   0
      Top             =   840
      Width           =   1215
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
      TabIndex        =   22
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   20
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox txtTotalVenda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10170
      TabIndex        =   18
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ComboBox cmbFormaAUX 
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
      Top             =   1380
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
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
      TabIndex        =   5
      Top             =   1800
      Width           =   3585
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   14
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
         NumButtons      =   7
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
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NFe"
            Key             =   "caixa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "NFCE"
            Key             =   "nfce"
            ImageIndex      =   6
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
               Picture         =   "PEDIDOCONSULTA.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCONSULTA.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCONSULTA.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCONSULTA.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCONSULTA.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCONSULTA.frx":D2E1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   360
      Left            =   3600
      TabIndex        =   1
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
      DesignHeight    =   7290
   End
   Begin MSComctlLib.ListView lstPedidoItem 
      Height          =   1905
      Left            =   45
      TabIndex        =   32
      Top             =   4440
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
      TabIndex        =   33
      Top             =   3360
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
         Text            =   "Doc.Fiscal"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vlr.Venda"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr.Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Dt.Emisão"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Faturamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ID"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "NºCaixa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "TpRegistro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Comanda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ipPedidoStatus"
         Object.Width           =   18
      EndProperty
   End
   Begin Threed.SSOption optSintetico 
      Height          =   255
      Left            =   10680
      TabIndex        =   35
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
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      TabIndex        =   36
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
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      Height          =   240
      Left            =   3600
      TabIndex        =   57
      Top             =   6555
      Width           =   1290
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CentroCusto:"
      Height          =   240
      Left            =   5805
      TabIndex        =   54
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.Faturado"
      Height          =   240
      Index           =   1
      Left            =   10200
      TabIndex        =   51
      Top             =   6555
      Width           =   1185
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "TotalDesconto"
      Height          =   240
      Left            =   5820
      TabIndex        =   50
      Top             =   6555
      Width           =   1350
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CMD:"
      Height          =   240
      Left            =   5160
      TabIndex        =   48
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Estação:"
      Height          =   240
      Left            =   3360
      TabIndex        =   45
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Estab.:"
      Height          =   240
      Left            =   720
      TabIndex        =   43
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "QtdeProdutos"
      Height          =   240
      Left            =   1440
      TabIndex        =   42
      Top             =   6555
      Width           =   1290
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Família:"
      Height          =   240
      Left            =   8160
      TabIndex        =   39
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   255
      Left            =   7920
      TabIndex        =   38
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   7800
      TabIndex        =   37
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   240
      Left            =   480
      TabIndex        =   31
      Top             =   2325
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NFCe:"
      Height          =   240
      Left            =   5970
      TabIndex        =   28
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NFe:"
      Height          =   240
      Left            =   6120
      TabIndex        =   27
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   9000
      TabIndex        =   25
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   240
      Left            =   2820
      TabIndex        =   24
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   240
      Left            =   540
      TabIndex        =   23
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos"
      Height          =   240
      Left            =   150
      TabIndex        =   21
      Top             =   6555
      Width           =   765
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.TotalVendas"
      Height          =   240
      Index           =   0
      Left            =   7905
      TabIndex        =   19
      Top             =   6555
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      Height          =   240
      Left            =   60
      TabIndex        =   17
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor(a):"
      Height          =   240
      Left            =   105
      TabIndex        =   15
      Top             =   1800
      Width           =   1230
   End
End
Attribute VB_Name = "frmPedidoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim SQL_CONSULTA        As String
   Dim SQL_CONSULTA_CORPO  As String
   Dim SQL_CONSULTA3       As String
   Dim CANCELA_LOOP        As Boolean

Private Sub Form_Load()
   CANCELA_LOOP = False
   CARREGA_COMBOS
   If MULT_EMPRESA_B = True Then
      txtPeso.Visible = False
      Label20.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         CANCELA_LOOP = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "nfce"
      Case "caixa"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False

         CRITERIO_A = ""
         TIPO_NFe_GERAR = "R"          'Tipo Saida

         If Not IsNull(lstPedido.SelectedItem.ListSubItems.item(13).Text) Then
            If Trim(lstPedido.SelectedItem.ListSubItems.item(13).Text) = "D" Then
               TIPO_NFe_GERAR = Trim(lstPedido.SelectedItem.ListSubItems.item(13).Text)
            End If
         End If

         Indr_Consulta = True
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then _
            If Not IsNull(lstPedido.SelectedItem.Text) Then _
               If IsNumeric(lstPedido.SelectedItem.Text) Then _
                  GERA_NOTA

         CRITERIO_A = ""
         Indr_Consulta = False
      Case "consultar"
         MONTA_CONSULTA_SQL True
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         PEDIDO_ID_N = 0
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
      Case "pedido"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False
         FORMULA_REL = ""
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            FORMULA_REL = lstPedido.SelectedItem.Text

            If Not IsNumeric(FORMULA_REL) Then _
               Exit Sub

            PEDIDO_ID_N = FORMULA_REL

            FORMULA_REL = "{vwRelVenda.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & PEDIDO_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.statusitem} <> 'C' "

            If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_pedido_venda.RPT"
            If CNPJ_EMPRESA_N = "15333554000188" Then _
               Nome_Relatorio = "pedido_shf.rpt"
            frmRELATORIO10.Show 1
         End If
      Case "print"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False

         MONTA_CONSULTA_SQL False
         GERA_REL
   End Select
   PEDIDO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  CRITERIO_A = ""
                  CRITERIO_A = Trim(InputBox("Informe CPF/CNPJ do cliente", "Atualização de Dados Pedido Venda", CRITERIO_A))

                  If Trim(CRITERIO_A) <> "" Then
                     SQL = "update PEDIDO set "
                     SQL = SQL & " cgccpf = '" & Trim(CRITERIO_A) & "'"
                     SQL = SQL & " where pedido_id = " & lstPedido.SelectedItem.Text
                     CONECTA_RETAGUARDA.Execute SQL
                  End If

                  SQL = ""
                  CRITERIO_A = ""
                  MONTA_CONSULTA_SQL True
               End If
            End If
         End If
      Case vbKeyF6
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then

                  NUMR_SEQ_N = 0 & Trim(lstPedido.SelectedItem.ListSubItems.item(14).Text)
                  If NUMR_SEQ_N < 3 Then
                     CRITERIO_A = ""
                     frmPedidoCancela.txtPedido.Text = 0 & lstPedido.SelectedItem.Text
                     frmPedidoCancela.Show 1
                     SQL = ""
                     CRITERIO_A = ""
                     MONTA_CONSULTA_SQL True
                     Else
                        If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
                           CRITERIO_A = ""
                           frmPedidoCancela.txtPedido.Text = 0 & lstPedido.SelectedItem.Text
                           frmPedidoCancela.Show 1
                           SQL = ""
                           CRITERIO_A = ""
                           MONTA_CONSULTA_SQL True
                           Else: MsgBox "Não permitido."
                        End If
                  End If

               End If
            End If
         End If
      Case vbKeyF7
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.ListItems.Clear

            SQL = "select PRODUTO.CODG_PRODUTO, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "
            SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.PRECO_CUSTO, pedidoitem.seq_id,"
            SQL = SQL & " PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.CFOP_id, pedidoitem.status, "
            SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, Produto.FORNECEDOR_ID"
            SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
            SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

            SQL = SQL & " where pedidoitem.produto_id = produto.produto_id "
            SQL = SQL & " and PEDIDO_ID = " & lstPedido.SelectedItem.Text

            'SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               MOSTRA_TOP "Duplo Click no grid ocultar", " ", " ", " ", ""
               lstPedidoItem.Visible = True
            End If
            While Not TabTemp.EOF
               VALOR_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
               VALOR_ITEM_N = TabTemp.Fields("qtd_pedida").Value * (TabTemp.Fields("valor_item").Value - VALOR_DESCONTO_N)

               Set item = lstPedidoItem.ListItems.Add(, "seq." & TabTemp.Fields("seq_id").Value, Trim(TabTemp.Fields("codg_produto").Value))
               item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
               item.SubItems(2) = "" & Format(Trim(TabTemp.Fields("qtd_pedida").Value), strFormatacao3Digitos)
               item.SubItems(3) = "" & Format(Trim(TabTemp.Fields("valor_item").Value), strFormatacao2Digitos)
               item.SubItems(4) = "" & Format(Trim(TabTemp.Fields("valor_desconto").Value), strFormatacao2Digitos)
               item.SubItems(5) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
               item.SubItems(6) = "" & Trim(TabTemp.Fields("CODG_ncm").Value)
               item.SubItems(7) = "" & lstPedido.SelectedItem.Text

               If Trim(TabTemp.Fields("status").Value) = "A" Then
                  item.ForeColor = vbBlue
                  item.ListSubItems(1).ForeColor = vbBlue
                  item.ListSubItems(2).ForeColor = vbBlue
                  item.ListSubItems(3).ForeColor = vbBlue
                  item.ListSubItems(4).ForeColor = vbBlue
                  item.ListSubItems(5).ForeColor = vbBlue
                  item.ListSubItems(6).ForeColor = vbBlue
               End If
               If Trim(TabTemp.Fields("status").Value) = "P" Then
                  item.ForeColor = vbBlack
                  item.ListSubItems(1).ForeColor = vbBlack
                  item.ListSubItems(2).ForeColor = vbBlack
                  item.ListSubItems(3).ForeColor = vbBlack
                  item.ListSubItems(4).ForeColor = vbBlack
                  item.ListSubItems(5).ForeColor = vbBlack
                  item.ListSubItems(6).ForeColor = vbBlack
               End If
               If Trim(TabTemp.Fields("status").Value) = "C" Then
                  item.ForeColor = vbRed
                  item.ListSubItems(1).ForeColor = vbRed
                  item.ListSubItems(2).ForeColor = vbRed
                  item.ListSubItems(3).ForeColor = vbRed
                  item.ListSubItems(4).ForeColor = vbRed
                  item.ListSubItems(5).ForeColor = vbRed
                  item.ListSubItems(6).ForeColor = vbRed
               End If
               TabTemp.MoveNext
               CRITERIO_A = ""
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.Refresh
         End If
      Case vbKeyF11
         'frmSenha.Show 1

         'If UCase(CRITERIO_A) = UCase("acerto") Then
            'PEDIDO_ID_N = 0
            'If Not IsNull(lstPedido.selectedItem.Text) Then
            '   If IsNumeric(lstPedido.selectedItem.Text) Then
            '      PEDIDO_ID_N = lstPedido.selectedItem.Text

            '      frmPedidoClienteAcerto.Show 1
            '      MONTA_CONSULTA_SQL True
            '      PEDIDO_ID_N = 0
            '   End If
            'End If
         'End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoVENDA_KeyDown"
End Sub

Private Sub cmdConsCli_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then _
      txtCNPJCPF.Text = CNPJCPF_A
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmdConsProd_Click()
   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Text = "99999999999"

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome,cliente_id from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCliente.EOF Then
      Beep
      MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
      txtCNPJCPF.SetFocus
      Exit Sub
      Else:
         CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
         If TabCliente!NOME <> "" Then _
            txtCli.Text = TabCliente!NOME
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_LostFocus"
End Sub

Private Sub txtDtFim_LostFocus()
   CHECA_ULTIMO_DIA_MES
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub cmbcpu_Click()
'On Error GoTo ERRO_TRATA

   cmbCPUaux.ListIndex = cmbCPU.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcpu_Click"
End Sub

Private Sub cmbestab_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbestab_Click"
End Sub

Private Sub txtReg_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtQtdeProd_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtPeso_GotFocus()
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
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpedido_KeyPress"
End Sub

Private Sub lstPedido_DblClick()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""
   If Not IsNull(lstPedido.SelectedItem.Text) Then
      NUMR_SEQ_N = 0 & Trim(lstPedido.SelectedItem.ListSubItems.item(14).Text)
      If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
         CRITERIO_A = lstPedido.SelectedItem.Text
         Unload Me
         Else
            If UCase(Trim(cmbForma.Text)) = "ENCOMENDA" Then
               CRITERIO_A = lstPedido.SelectedItem.Text
               Unload Me
               Else
               If NUMR_SEQ_N = 3 Or NUMR_SEQ_N = 5 Or NUMR_SEQ_N = 7 Or NUMR_SEQ_N = 9 Then
                  MsgBox "Permitido somente consulta."
                  PEDIDO_ID_N = 0
                  Exit Sub
                  Else
                     If NUMR_SEQ_N = 1 Then
                        CRITERIO_A = lstPedido.SelectedItem.Text
                        Unload Me
                     End If
               End If
            End If
      End If
   End If

Exit Sub
ERRO_TRATA:
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
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_Click"
End Sub

Private Sub lstPedidoitem_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoitem_DblClick"
End Sub

Private Sub lstpedidoitem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedidoItem, ColumnHeader
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbFormaAUX.ListIndex = cmbForma.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFORMA_Click"
End Sub

Private Sub cmbSituacao_Click()
'On Error GoTo ERRO_TRATA

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbsituacao_Click"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_Click"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
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
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub cmbCC_Click()
On Error Resume Next

   cmbCCAux.ListIndex = cmbCC.ListIndex
   Call cmbcc_LostFocus
End Sub

Private Sub cmbcc_LostFocus()

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

   CANCELA_LOOP = False
   CLIENTE_ID_N = 0
   cmbCC.Text = ""
   cmbCCAux.Text = ""
   txtComanda.Text = ""
   cmbCPU.Text = ""
   cmbCPUaux.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   txtNOTA.Text = ""
   txtCupom.Text = ""
   lstPedido.ListItems.Clear
   txtPedido.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""

   If cmbVend.Enabled = True Then _
      cmbVend.Text = ""

   cmbForma.Text = ""
   cmbFormaAUX.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtTotalVenda.Text = ""
   txtTotVendas.Text = ""
   txtTotDesconto.Text = ""
   txtReg.Text = ""
   txtQtdeProd.Text = ""
   txtPeso.Text = ""

   lstPedido.Visible = True
   optSintetico.Value = True
   txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = lstPedido.SelectedItem.Text
   CNPJCPF_A = ""

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select status, cgccpf from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      CNPJCPF_A = "" & Trim(TabCabeca.Fields("cgccpf").Value)
      If Not IsNull(TabCabeca!STATUS) Then
         If TabCabeca!STATUS <> "9" Then
            If Trim(CNPJCPF_A) = "99999999999" Then
               Msg = "Para geração de nota fiscal eletrônica, os dados do cliente devem ser cadastrados, deseja continuar essa operação ?"
               PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CNPJCPF_A = ""
                  TIPO_PESSOA_CADASTRO = "CLIENTE"
                  frmPessoaConsulta.Show 1
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
      TIPO_PESSOA_CADASTRO = "CLIENTE"
      frmPessoaCadastro.Show 1

      MsgBox "Repetir operação."
      Else
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Exit Sub
   End If
   Else
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
End If
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

                     SQL = "update PEDIDO set "
                     SQL = SQL & " cgccpf = '" & Trim(CNPJCPF_A) & "'"
                     SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                     SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
                     Else
                        If TabCabeca.State = 1 Then _
                           TabCabeca.Close
                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close
                        Exit Sub
                  End If
                  Else
                     If TabCabeca.State = 1 Then _
                        TabCabeca.Close
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
                     Exit Sub
               End If
            End If

            CRITERIO_A = PEDIDO_ID_N
            'TIPO_NFe_GERAR = "R"
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            If USA_DOC_FISCAL = True Then _
               If USA_NFe = True Then _
                  frmNOTAGERA.Show 1
         End If
      End If
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
End Sub

Sub PROCURA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select produto_id,descricao from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDescProd.Text = TabConsulta.Fields("descricao").Value
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
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
         CRITERIO_A = FimDoMes(DMA(txtDtIni.Text), False)
         CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
         txtDtFim.Text = CRITERIO_A
         txtDtFim.PromptInclude = True
      End If
   End If
End Sub

Sub CARREGA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(6).Visible = False 'NFe
   If USA_NFe = True Then _
      Toolbar1.Buttons(6).Visible = True

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbVend.Enabled = True

      Else
         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select logon from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            CRITERIO_A = Chr$(39) & Trim(TabUSU.Fields("logon").Value) & "%" & Chr(39)

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
            SQL = SQL & " where descricao like " & CRITERIO_A
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVend.Text = TabVENDEDOR!DESCRICAO
               cmbVendAux.Text = TabVENDEDOR!VENDEDOR_ID
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
   End If

Exit Sub
ERRO_TRATA:
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

   cmbSituacao.AddItem "Todos"
   cmbSituacaoAUX.AddItem ""

   cmbSituacao.AddItem "Faturado"
   cmbSituacaoAUX.AddItem "'3','5','7'"

   cmbSituacao.AddItem "Pendente"
   cmbSituacaoAUX.AddItem "'1','2','4'"

   cmbSituacao.AddItem "Cupom Fiscal"
   cmbSituacaoAUX.AddItem "'7'"

   cmbSituacao.AddItem "Nota Eletrônica"
   cmbSituacaoAUX.AddItem "'7','5','3'"

   cmbSituacao.AddItem "Cancelado"
   cmbSituacaoAUX.AddItem "'9'"

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

   cmbEstabAUX.Clear
   cmbEstab.Clear
   cmbEstab.AddItem "Todos"
   cmbEstabAUX.AddItem ""

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbEstab.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      cmbEstabAUX.AddItem Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N

   Toolbar1.Buttons(6).Visible = False 'NFe
   If USA_NFe = True Then _
      Toolbar1.Buttons(6).Visible = True

   cmbEstab.Visible = False
   Label15.Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbEstab.Visible = True
      Label15.Visible = True
   End If

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

   cmbForma.Clear
   cmbFormaAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where receber = 'true' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbForma.AddItem Trim(TabTemp!DESCRICAO)
      cmbFormaAUX.AddItem TabTemp!TIPOVENDA_ID
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbVend.Clear
   SQL = "select vendedor_id,descricao from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by descricao "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbVend.AddItem Trim(TabTemp!DESCRICAO) & " - " & Trim(TabTemp!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabTemp!VENDEDOR_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbVend.Enabled = False
   Toolbar1.Buttons(6).Visible = False 'NFe

   CARREGA_VENDEDOR

   If MULT_EMPRESA_B = False Then
      If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
         cmbVend.Enabled = True
         cmbVend.Text = ""
      End If

      MONTA_CONSULTA_SQL True
   End If

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
      cmbCCAux.AddItem TabTemp!Codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   HORA_INI = Time
   Me.Caption = "Aguarde, Pesquisando ..."
   SQL_CONSULTA = ""
   SQL_CONSULTA_CORPO = ""
   SQL_CONSULTA3 = ""
   CRITERIO_A = ""

   CHECA_ULTIMO_DIA_MES

   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False

   VALOR_TOTAL_N = 0
   If Indr_Consulta = True Then
      txtTotalVenda.Text = ""
      txtReg.Text = ""
      txtQtdeProd.Text = ""
      txtPeso.Text = ""
   End If
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SQL_CONSULTA = ""
   SQL_CONSULTA_CORPO = ""
   SQL_CONSULTA3 = ""
   CRITERIO_A = ""

   Dim SQL_COUNT  As String
   SQL_COUNT = ""

   SQL_CONSULTA = "select * from vwCONSULTA_PEDIDO WITH (NOLOCK) "
SQL_COUNT = "select count(pedido_id) from PEDIDO "

   SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " where pedido_id Is Not Null"

   If Trim(cmbEstabAUX.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and estabelecimento_id = " & cmbEstabAUX.Text

'itens pedido
SQL = SQL & " and status <> 'C'"

   If Trim(cmbCPUaux.Text) <> "" Then _
      If IsNumeric(cmbCPUaux.Text) Then _
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and numero_caixa_cpu = " & cmbCPUaux.Text

   If Trim(txtPedido.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and pedido_id = " & txtPedido.Text

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then _
      If CLIENTE_ID_N > 0 Then _
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and cliente_id = " & CLIENTE_ID_N
   txtCNPJCPF.PromptInclude = True

   If Trim(cmbFormaAUX.Text) <> "" Then
      If UCase(Trim(cmbForma.Text)) = "ENCOMENDA" Then
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and tipovenda_id = " & cmbFormaAUX.Text
         Else
            SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and tipovenda_id = " & cmbFormaAUX.Text

            If Trim(cmbVend.Text) <> "" Then _
               SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and vendedor_id = " & cmbVendAux.Text
      End If
      Else
         If Trim(cmbVend.Text) <> "" Then _
            SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and vendedor_id = " & cmbVendAux.Text
   End If


   If Trim(txtComanda.Text) <> "" Then _
      If IsNumeric(txtComanda.Text) Then _
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and cartaobarra_id = " & txtComanda.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and dt_req <= '" & txtDtFim.Text & "'"
   End If

'===============
SQL_COUNT = SQL_COUNT & SQL_CONSULTA_CORPO

   If Trim(cmbSituacaoAUX.Text) <> "" Then
      If Trim(cmbSituacaoAUX.Text) = "'7','5','3'" Then _
         SQL_COUNT = SQL_COUNT & " and numr_nota > 0 "

      SQL_COUNT = SQL_COUNT & " and status in (" & Trim(cmbSituacaoAUX.Text) & ")"
   End If
'===============
   
   If Trim(cmbSituacaoAUX.Text) <> "" Then
      If Trim(cmbSituacaoAUX.Text) = "'7','5','3'" Then _
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and numr_nota > 0 "

      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and SIT_PEDIDO in (" & Trim(cmbSituacaoAUX.Text) & ")"
   End If
   If Trim(txtProduto.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and produto_id = " & PRODUTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   If Trim(txtCupom.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and numr_cupom = " & txtCupom.Text

   If Trim(txtNOTA.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and numr_nota = " & txtNOTA.Text

'===============

   Dim SQL_ITENS        As String
   Dim TabTemp          As New ADODB.Recordset
   Dim Conta_Produto_N  As Long
   Dim Peso_N           As Double

   Peso_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   SQL_ITENS = ""

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

   CONT_N = 0
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL_CONSULTA = SQL_CONSULTA & SQL_CONSULTA_CORPO & " order by PEDIDO_ID desc"

'SQL3 = "EXECUTE sp_PedidoConsulta 1," & txtDtIni.Text & "," & txtDtFim.Text & "," & Trim(cmbSituacaoAUX.Text)


'Debug.Print SQL_CONSULTA

'Set TabTemp = CONECTA_RETAGUARDA.Execute(SQL3)

   TabTemp.Open SQL_CONSULTA, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Exit Sub
   End If
   If Not TabTemp.EOF Then
      While Not TabTemp.EOF
         DoEvents

         If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
            CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
            txtReg.Text = CONTA_REGISTRO_N
            txtReg.Refresh

            PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value

            NUMR_SEQ_N = NUMR_SEQ_N + 1
            Set item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp.Fields("PEDIDO_ID").Value)

            item.SubItems(11) = "" & TabTemp.Fields("PEDIDO_ID").Value
            item.SubItems(1) = "" & TabTemp.Fields("numr_cupom").Value

If Not IsNull(TabTemp.Fields("numr_nota").Value) Then _
   If Trim(TabTemp.Fields("numr_nota").Value) Then _
      item.SubItems(1) = "" & TabTemp.Fields("numr_nota").Value

            item.SubItems(2) = "" & Trim(TabTemp!NOME_CLIENTE) & " - " & Trim(TabTemp.Fields("CNPJCPF").Value)

            If IsNull(TabTemp!NOME_CLIENTE) Or Trim(TabTemp!NOME_CLIENTE) = "" Then
               If TabCliente.State = 1 Then _
                  TabCliente.Close

               SQL = "select nome from CLIENTE WITH (NOLOCK)"
               SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"
               TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCliente.EOF Then _
                  item.SubItems(2) = "" & TabCliente!NOME

               If TabCliente.State = 1 Then _
                  TabCliente.Close
            End If

            item.SubItems(6) = TabTemp!DT_REQ
            item.SubItems(7) = ""

            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            SQL = "select * from TIPOVENDA WITH (NOLOCK)"
            SQL = SQL & " where tipovenda_id = " & TabTemp!TIPOVENDA_ID
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then
               item.SubItems(7) = Trim(TabDESCR!DESCRICAO)

               If Trim(txtPedido.Text) <> "" Then _
                  If UCase(Trim(TabDESCR!DESCRICAO)) = "ENCOMENDA" Then _
                     cmbForma.Text = "ENCOMENDA"
            End If
            If TabDESCR.State = 1 Then _
               TabDESCR.Close
            If TabUSU.State = 1 Then _
               TabUSU.Close

            item.SubItems(8) = ""
   
            SQL = "select * from vwVendedor WITH (NOLOCK)"
            SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id").Value
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               item.SubItems(8) = TabUSU!DESCRICAO
            If TabUSU.State = 1 Then _
               TabUSU.Close

            item.SubItems(9) = ""

            If Not IsNull(TabTemp.Fields("Status")) Then
               If TabTemp.Fields("Status") = 2 Then
                  If TabTemp.Fields("tipo_registro") = "O" Then
                     item.SubItems(9) = "Orcamento"
                     Else: item.SubItems(9) = "Pedido"
                  End If
               End If
               If TabTemp.Fields("Status").Value = 3 Then _
                  item.SubItems(9) = "3-Faturado"
               If TabTemp.Fields("Status").Value = 4 Then _
                  item.SubItems(9) = "4-Cupom"
               If TabTemp.Fields("Status").Value = 5 Then _
                  item.SubItems(9) = "5-Faturado"
               If TabTemp.Fields("Status").Value = 7 Then _
                  item.SubItems(9) = "7-Cupom Fiscal"
               If TabTemp.Fields("Status").Value = 9 Then _
                  item.SubItems(9) = "9-Cancelado"
            End If

            If Not IsNull(TabTemp.Fields("numero_caixa_cpu").Value) Then _
               item.SubItems(11) = TabTemp.Fields("numero_caixa_cpu").Value

            item.SubItems(12) = TabTemp.Fields("tipo_registro").Value

            VALOR_DESCONTO_N = 0
            'VALOR_TOTAL_DESCONTO_N = 0

            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
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

            SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            'SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            If Not IsNull(TabTemp.Fields("unidade_medida").Value) Then
               If UCase(TabTemp.Fields("unidade_medida").Value) = UCase("kg") Then
                  SQL = "select sum(qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
                  SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
                  SQL = SQL & " and tipo_reg = 'PC' "
                  SQL = SQL & " and pedidoitem.status <> 'C' "
                  TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
                     Peso_N = Peso_N + TabPedidoItem.Fields(0).Value
                  If TabPedidoItem.State = 1 Then _
                     TabPedidoItem.Close

                  txtPeso.Text = Peso_N
                  txtPeso.Refresh
               End If
            End If
            'VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N
            VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

            txtTotDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
            txtTotVendas.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
            txtTotalVenda.Text = Format(VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

            item.SubItems(3) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
            item.SubItems(4) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
            item.SubItems(5) = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
            item.SubItems(14) = "" & TabTemp.Fields("sit_pedido")
            If TabTemp.Fields("SIT_PEDIDO").Value = 1 Then
               item.ForeColor = vbRed
               item.ListSubItems(1).ForeColor = vbRed
               item.ListSubItems(2).ForeColor = vbRed
               item.ListSubItems(3).ForeColor = vbRed
               item.ListSubItems(4).ForeColor = vbRed
               item.ListSubItems(5).ForeColor = vbRed
               item.ListSubItems(6).ForeColor = vbRed
               item.ListSubItems(7).ForeColor = vbRed
               item.ListSubItems(8).ForeColor = vbRed
               item.SubItems(9) = "" & "Em Aberto - 1"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 2 Then
               item.ForeColor = vbBlue
               item.ListSubItems(1).ForeColor = vbBlue
               item.ListSubItems(2).ForeColor = vbBlue
               item.ListSubItems(3).ForeColor = vbBlue
               item.ListSubItems(4).ForeColor = vbBlue
               item.ListSubItems(5).ForeColor = vbBlue
               item.ListSubItems(6).ForeColor = vbBlue
               item.ListSubItems(7).ForeColor = vbBlue
               item.ListSubItems(8).ForeColor = vbBlue
               item.ListSubItems(9).ForeColor = vbBlue
               item.SubItems(9) = "" & "A Faturar - 2"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 3 Then
               item.ForeColor = vbBlack
               item.ListSubItems(1).ForeColor = vbBlack
               item.ListSubItems(2).ForeColor = vbBlack
               item.ListSubItems(3).ForeColor = vbBlack
               item.ListSubItems(4).ForeColor = vbBlack
               item.ListSubItems(5).ForeColor = vbBlack
               item.ListSubItems(6).ForeColor = vbBlack
               item.ListSubItems(7).ForeColor = vbBlack
               item.ListSubItems(8).ForeColor = vbBlack
               item.ListSubItems(9).ForeColor = vbBlack
               item.SubItems(9) = "" & "Faturado - 3"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 5 Then
               item.ForeColor = vbBlack
               item.ListSubItems(1).ForeColor = vbBlack
               item.ListSubItems(2).ForeColor = vbBlack
               item.ListSubItems(3).ForeColor = vbBlack
               item.ListSubItems(4).ForeColor = vbBlack
               item.ListSubItems(5).ForeColor = vbBlack
               item.ListSubItems(6).ForeColor = vbBlack
               item.ListSubItems(7).ForeColor = vbBlack
               item.ListSubItems(8).ForeColor = vbBlack
               item.ListSubItems(9).ForeColor = vbBlack
               item.SubItems(9) = "" & "Faturado - 5"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 6 Then
               item.ForeColor = vbBlack
               'Item.ListSubItems(1).ForeColor = vbYellow
               'Item.ListSubItems(2).ForeColor = vbYellow
               'Item.ListSubItems(3).ForeColor = vbYellow
               'Item.ListSubItems(4).ForeColor = vbYellow
               'Item.ListSubItems(5).ForeColor = vbYellow
               'Item.ListSubItems(6).ForeColor = vbYellow
               'Item.ListSubItems(7).ForeColor = vbYellow
               'Item.ListSubItems(8).ForeColor = vbYellow
               item.ListSubItems(9).ForeColor = vbYellow
               item.SubItems(9) = "" & "Não Contabilizado"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 7 Then
               item.ForeColor = vbMagenta
               item.ListSubItems(1).ForeColor = vbMagenta
               item.ListSubItems(2).ForeColor = vbMagenta
               item.ListSubItems(3).ForeColor = vbMagenta
               item.ListSubItems(4).ForeColor = vbMagenta
               item.ListSubItems(5).ForeColor = vbMagenta
               item.ListSubItems(6).ForeColor = vbMagenta
               item.ListSubItems(7).ForeColor = vbMagenta
               item.ListSubItems(8).ForeColor = vbMagenta
               item.ListSubItems(9).ForeColor = vbMagenta
               item.SubItems(9) = "" & "Cupom Fiscal - 7"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 9 Then
               item.ListSubItems(7).ForeColor = &HC0E0FF
               item.ListSubItems(8).ForeColor = &HC0E0FF
               item.ListSubItems(9).ForeColor = &HC0E0FF
               item.SubItems(9) = "" & "Cancelado - 9"
            End If
         End If

         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed

         'verificando se é venda com comanda eletronica
         item.SubItems(13) = "" & TabTemp.Fields("cartaobarra_id").Value

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select cartaobarra_id from PEDIDOTEMP WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("PEDIDO_ID").Value
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            item.ListSubItems(7).ForeColor = vbBlack
            item.ListSubItems(8).ForeColor = vbBlack
            item.SubItems(13) = "" & TabDESCR.Fields(0).Value
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         TabTemp.MoveNext

         If CANCELA_LOOP = True Then _
            Exit Sub
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
''''''''''''''''''''''''''''''''
   SQL_ITENS = "select count(qtd_pedida) from vwCONSULTA_PEDIDO WITH (NOLOCK) "
   SQL_ITENS = SQL_ITENS & SQL_CONSULTA_CORPO
   SQL_ITENS = SQL_ITENS & " and vwCONSULTA_PEDIDO.tipo_reg = 'PC' "
   SQL_ITENS = SQL_ITENS & " and vwCONSULTA_PEDIDO.status <> 'C' "
   SQL = SQL & " and status <> 'C'"
   TabTemp.Open SQL_ITENS, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabTemp.Fields(0).Value) Then _
      Conta_Produto_N = Conta_Produto_N + TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtQtdeProd.Text = Conta_Produto_N
   txtQtdeProd.Refresh
''''''''''''''''''''''''''''''''
   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Sub GERA_REL()
'On Error GoTo ERRO_TRATA

   Dim VALOR_ITEM_N     As Double
   Dim DESCONTO_ITEM    As Double
   Dim DESCONTO_CABEÇA  As Double
   Dim VALOR_CUSTO_N    As Double
   Dim CARTAO_ID        As Long
   Dim strSQL           As String
   Dim TabTempRel       As New ADODB.Recordset

   VALOR_ITEM_N = 0
   QTDE_N = 0
   DESCONTO_ITEM = 0
   DESCONTO_CABEÇA = 0
   VALOR_CUSTO_N = 0
   CARTAO_ID = 0
   CONT_N = 0
   CRITERIO_A = SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "RELVENDA", "U") = True Then
      strSQL = "drop table RELVENDA"
      CONECTA_RETAGUARDA.Execute strSQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "RELVENDA", "U") = False Then
      strSQL = "create table RELVENDA"
      strSQL = strSQL & " ("
         strSQL = strSQL & " RELVENDA_ID        bigint      not null,"
         strSQL = strSQL & " ESTABELECIMENTO_ID         bigint      not null,"
         strSQL = strSQL & " PEDIDO_ID          bigint      not null,"
         strSQL = strSQL & " CLIENTE_ID         bigint      not null,"
         strSQL = strSQL & " VENDEDOR_ID        bigint      not null,"
         strSQL = strSQL & " tipovenda_ID       bigint      not null,"
         strSQL = strSQL & " DT_VENDA           datetime    not null,"

         strSQL = strSQL & " VALOR_VENDA        float       not null,"
         strSQL = strSQL & " VLR_TOT_CUSTO      float       null    ,"
         strSQL = strSQL & " VLR_TOT_DESCONTO   float       null    ,"
         strSQL = strSQL & " CLIENTE            varchar(50) null    ,"

         strSQL = strSQL & " QTDE_VENDIDA       float       not null,"
         strSQL = strSQL & " CARTAOBARRA_ID BIGINT,"

         strSQL = strSQL & " constraint PK_RELVENDA primary key (RELVENDA_ID)"
      strSQL = strSQL & " )"
      CONECTA_RETAGUARDA.Execute strSQL
   End If

   strSQL = "delete from RELVENDA"
   CONECTA_RETAGUARDA.Execute strSQL

   QTDE_N = 0
   VALOR_ITEM_N = 0
   VALOR_CUSTO_N = 0
   VALOR_DESCONTO_N = 0
   NUMR_ID_N = 0
   PEDIDO_ID_N = 0

   If TabTempRel.State = 1 Then _
      TabTempRel.Close

'SQL_CONSULTA = SQL_CONSULTA & SQL_CONSULTA_CORPO & " order by PEDIDO_ID desc"

   TabTempRel.Open SQL_CONSULTA, CONECTA_RETAGUARDA, , , adCmdText
   If TabTempRel.EOF Then
      If TabTempRel.State = 1 Then _
         TabTempRel.Close
      MsgBox "Registro não encontrado."
      Exit Sub
   End If
   While Not TabTempRel.EOF
      CARTAO_ID = 0 & TabTempRel.Fields("cartaobarra_id").Value
      NOME_CLIENTE_A = "" & Trim(TabTempRel.Fields("NOME_CLIENTE").Value)
      PEDIDO_ID_N = 0 & Trim(TabTempRel.Fields("pedido_id").Value)
      If Trim(TabTempRel.Fields("NOME_CLIENTE").Value) = "" Then
         If TabCliente.State = 1 Then _
            TabCliente.Close
      
            SQL = "select nome from CLIENTE WITH (NOLOCK)"
            SQL = SQL & " where cliente_id = " & TabTempRel.Fields("cliente_id").Value
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               NOME_CLIENTE_A = Trim(TabCliente.Fields(0).Value)
      
         If TabCliente.State = 1 Then _
            TabCliente.Close
      End If

      VALOR_ITEM_N = 0 & TabTempRel.Fields("valor_total").Value
      'VALOR_ITEM_N = 0 & (TabTempRel.Fields("QTD_PEDIDA").Value * TabTempRel.Fields("VALOR_ITEM").Value)
      QTDE_N = 0 & TabTempRel.Fields("QTD_PEDIDA").Value
      DESCONTO_ITEM = 0 & TabTempRel.Fields("VALOR_DESCONTO").Value
      DESCONTO_CABEÇA = 0 & TabTempRel.Fields("DESCCABECA").Value
      VALOR_CUSTO_N = 0 & TabTempRel.Fields("preco_custo").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from RELVENDA WITH (NOLOCK)"
      SQL = SQL & " where PEDIDO_ID = " & PEDIDO_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         NUMR_ID_N = MAX_ID("RELVENDA_id", "RELVENDA", "", "", "", "")

         SQL = "insert into RELVENDA "
         SQL = SQL & "("
            SQL = SQL & " RELVENDA_ID,ESTABELECIMENTO_ID,PEDIDO_ID,CLIENTE_ID,VENDEDOR_ID,DT_VENDA,"
            SQL = SQL & " VALOR_VENDA,VLR_TOT_CUSTO,VLR_TOT_DESCONTO,CLIENTE,QTDE_VENDIDA,"
            SQL = SQL & " cartaobarra_id,tipovenda_id "
         SQL = SQL & ")"
         SQL = SQL & " values("
            SQL = SQL & NUMR_ID_N                                                            'RELVENDA_ID
            SQL = SQL & "," & TabTempRel.Fields("ESTABELECIMENTO_ID").Value                  'ESTABELECIMENTO_ID
            SQL = SQL & "," & PEDIDO_ID_N                                                    'PEDIDO_ID
            SQL = SQL & "," & TabTempRel.Fields("CLIENTE_ID").Value                          'CLIENTE_ID
            SQL = SQL & "," & TabTempRel.Fields("vendedor_ID").Value                         'VENDEDOR_ID
            SQL = SQL & ",'" & Trim(TabTempRel.Fields("dt_req").Value) & "'"                 'DT_VENDA

            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                                          'VALOR_VENDA
            SQL = SQL & "," & tpMOEDA(VALOR_CUSTO_N)                                         'VLR_TOT_CUSTO
            SQL = SQL & "," & tpMOEDA(DESCONTO_CABEÇA)                                       'VLR_TOT_DESCONTO
            SQL = SQL & ",'" & Trim(Left(TabTempRel.Fields("NOME_CLIENTE").Value, 50)) & "'" 'CLIENTE
            
            SQL = SQL & "," & tpMOEDA(QTDE_N)                                                'QTDE_VENDIDA
            SQL = SQL & "," & CARTAO_ID                                                      'comanda
            SQL = SQL & "," & TabTempRel.Fields("tipovenda_id").Value                        'tipovenda_id
         SQL = SQL & ")"
         Else
            SQL = "update RELVENDA set "

            'SQL = SQL & "VALOR_VENDA = valor_venda + " & tpMOEDA(VALOR_ITEM_N)                  'VALOR_VENDA
            SQL = SQL & " VLR_TOT_CUSTO = VLR_TOT_CUSTO + " & tpMOEDA(VALOR_CUSTO_N * QTDE_N)   'VLR_TOT_CUSTO
            SQL = SQL & ",QTDE_VENDIDA = QTDE_VENDIDA + " & tpMOEDA(QTDE_N)                     'QTDE_VENDIDA

            SQL = SQL & " where PEDIDO_ID = " & PEDIDO_ID_N
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      DoEvents

      TabTempRel.MoveNext
   Wend
   If TabTempRel.State = 1 Then _
      TabTempRel.Close

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If optAnalitico.Value = True Then
      Nome_Relatorio = "venda_totais_analitico.rpt"
      Else: Nome_Relatorio = "venda_totais.rpt"
   End If

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_REL"
End Sub
