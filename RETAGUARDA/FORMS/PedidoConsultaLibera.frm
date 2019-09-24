VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoConsultaLibera 
   Caption         =   "Consulta Liberação de Pedidos"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoConsultaLibera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkTabela 
      Caption         =   "Abaixo Pr.Tabela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   210
      Left            =   10080
      TabIndex        =   46
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkCusto 
      Caption         =   "Abaixo Pr.Custo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   10080
      TabIndex        =   45
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   7560
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbFamiliaAUX 
      Appearance      =   0  'Flat
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
      TabIndex        =   43
      Top             =   2760
      Visible         =   0   'False
      Width           =   870
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
      Left            =   7800
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbVendAux 
      Appearance      =   0  'Flat
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
      Left            =   6480
      TabIndex        =   41
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
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
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbVend 
      Height          =   360
      Left            =   6480
      TabIndex        =   5
      Top             =   1320
      Width           =   3105
   End
   Begin VB.ComboBox cmbForma 
      Height          =   360
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtTotalVenda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10170
      TabIndex        =   19
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   18
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox txtCli 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   17
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   10560
      TabIndex        =   3
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
      Left            =   7800
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtDescProd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   16
      Top             =   2280
      Width           =   5655
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Picture         =   "PedidoConsultaLibera.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtQtdeProd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1440
      TabIndex        =   14
      Top             =   6840
      Width           =   1815
   End
   Begin VB.ComboBox cmbEstab 
      Height          =   360
      Left            =   7560
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdConsCli 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Picture         =   "PedidoConsultaLibera.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtTotDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   5865
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtTotVendas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7920
      TabIndex        =   11
      Top             =   6840
      Width           =   1575
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
      TabIndex        =   10
      Top             =   1920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1270
      ButtonWidth     =   2858
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
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
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "PedidoConsultaLibera.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaLibera.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaLibera.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaLibera.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaLibera.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaLibera.frx":D2E1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      DesignWidth     =   11880
      DesignHeight    =   7245
   End
   Begin MSComctlLib.ListView lstPedidoItem 
      Height          =   1905
      Left            =   45
      TabIndex        =   21
      Top             =   4320
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2252
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valr.Unitário"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Vlr.Venda"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Desconto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valr.Total"
         Object.Width           =   3003
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3015
      Left            =   0
      TabIndex        =   22
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Vlr.Venda"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vlr.Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dt.Emisão"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Faturamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
   End
   Begin Threed.SSOption optSintetico 
      Height          =   270
      Left            =   10320
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   476
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sintético"
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin Threed.SSOption optAnalitico 
      Height          =   255
      Left            =   10320
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Analítico"
      Value           =   -1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor(a):"
      Height          =   240
      Left            =   5145
      TabIndex        =   39
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.TotalVendas"
      Height          =   240
      Index           =   0
      Left            =   7905
      TabIndex        =   37
      Top             =   6480
      Width           =   1515
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos"
      Height          =   240
      Left            =   150
      TabIndex        =   36
      Top             =   6480
      Width           =   765
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   240
      Left            =   9660
      TabIndex        =   35
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   6840
      TabIndex        =   33
      Top             =   840
      Width           =   900
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
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   240
      Left            =   3480
      TabIndex        =   30
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Família:"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos.Vendidos"
      Height          =   240
      Left            =   1470
      TabIndex        =   28
      Top             =   6480
      Width           =   1785
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Estab.:"
      Height          =   240
      Left            =   6840
      TabIndex        =   27
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "TotalDesconto"
      Height          =   240
      Left            =   5820
      TabIndex        =   26
      Top             =   6480
      Width           =   1350
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.Faturado"
      Height          =   240
      Index           =   1
      Left            =   10200
      TabIndex        =   25
      Top             =   6480
      Width           =   1185
   End
End
Attribute VB_Name = "frmPedidoConsultaLibera"
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
      Case vbKeyF7
         MOSTRA_ITENS_PEDIDO
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_LostFocus"
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
   Me.Enabled = True
   Me.KeyPreview = True
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
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

Private Sub cmbestab_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbestab_Click"
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
      CRITERIO_A = lstPedido.SelectedItem.Text
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
      Case "imprimir"
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         FORMULA_REL = ""

         Nome_Relatorio = "pedido_Libera.rpt"
         frmRELATORIO10.Show 1
      Case "consultar"
         MONTA_CONSULTA_SQL True
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
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

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex

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

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
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

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
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

Private Sub TXTDTFIM_GotFocus()
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

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   chkTabela.Value = 1
   chkCusto.Value = 1

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   lstPedido.ListItems.Clear
   txtPedido.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
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
   Me.Enabled = True
   Me.KeyPreview = True
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
   Me.Enabled = True
   Me.KeyPreview = True
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

   cmbSituacao.AddItem "Todos"
   cmbSituacaoAUX.AddItem ""

   cmbSituacao.AddItem "Cupom Fiscal"
   cmbSituacaoAUX.AddItem "'7'"

   cmbSituacao.AddItem "Nota Eletrônica"
   cmbSituacaoAUX.AddItem "'7','5','3'"

   cmbSituacao.AddItem "Pendente"
   cmbSituacaoAUX.AddItem "'1','2','4'"

   cmbSituacao.AddItem "Faturado"
   cmbSituacaoAUX.AddItem "'3','5','7'"

   cmbSituacao.AddItem "Cancelado"
   cmbSituacaoAUX.AddItem "'9'"

   cmbSituacao.Text = "Todos"
   cmbSituacaoAUX.Text = ""

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

   cmbEstab.Visible = False
   Label15.Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbEstab.Visible = True
      Label15.Visible = True
   End If

   cmbForma.Clear
   cmbAuxForma.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where receber = 'true' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbForma.AddItem TabTemp!DESCRICAO & " - " & TabTemp!TIPOVENDA_ID
      cmbAuxForma.AddItem TabTemp!TIPOVENDA_ID
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

   CARREGA_VENDEDOR

   If MULT_EMPRESA_B = False Then
      If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
         cmbVend.Enabled = True
         cmbVend.Text = ""
      End If
   End If

   Me.Enabled = True
   Me.KeyPreview = True
   VALOR_TOTAL_N = 0

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Sub MOSTRA_ITENS_PEDIDO()
   If Not IsNull(lstPedido.SelectedItem.Text) Then
      Dim ValorVenda_N     As Double
      Dim ValorTabela_N    As Double

      lstPedidoItem.ListItems.Clear
      CONT_N = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select PEDIDO.pedido_id,PRODUTO.CODG_PRODUTO, PRODUTO.PRECO_Venda, PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.PRECO_CUSTO, "
      SQL = SQL & " PEDIDOITEM.SEQ_ID, PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.CFOP_ID, PEDIDOITEM.STATUS, PEDIDOITEM.QTD_PEDIDA, PRODUTO.DESCRICAO,"
      SQL = SQL & " PEDIDOITEM.PRODUTO_ID , Produto.Tipo_Prod, Produto.CODG_NCM, Produto.FORNECEDOR_ID, PEDIDO.TABELAPRECO_ID"
      SQL = SQL & " from PEDIDOITEM WITH (NOLOCK) "
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " INNER JOIN PEDIDO "
      SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = PEDIDO.PEDIDO_ID"

      SQL = SQL & " where pedidoitem.PEDIDO_ID = " & lstPedido.SelectedItem.Text

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         MOSTRA_TOP "Duplo Click no grid ocultar", " ", " ", " ", ""
         lstPedidoItem.Visible = True
      End If
      While Not TabTemp.EOF
         CONT_N = CONT_N + 1
         VALOR_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
         VALOR_ITEM_N = TabTemp.Fields("qtd_pedida").Value * (TabTemp.Fields("valor_item").Value - VALOR_DESCONTO_N)

         Set item = lstPedidoItem.ListItems.Add(, "seqitens." & CONT_N, Trim(TabTemp.Fields("codg_produto").Value))
         item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
         item.SubItems(2) = "" & Format(Trim(TabTemp.Fields("qtd_pedida").Value), strFormatacao3Digitos)

         ValorVenda_N = 0 & TabTemp.Fields("preco_venda").Value
         FORMAPAGTO_ID_N = 1

'=================== buscando a forma que o pedido foi faturado
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select ITEMLANCAMENTO.FORMAPAGTO_ID from LANCAMENTO WITH (NOLOCK) "
         SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
         SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
         SQL = SQL & " AND LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

         SQL = SQL & " where lancamento.numr_doc = " & TabTemp.Fields("pedido_id").Value

         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               FORMAPAGTO_ID_N = 0 & TabConsulta.Fields(0).Value
'===================

         ValorTabela_N = 0 & (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(TabTemp.Fields("produto_id").Value, TabTemp.Fields("tabelapreco_id").Value, FORMAPAGTO_ID_N))
         If ValorTabela_N > 0 Then _
            ValorVenda_N = 0 & ValorTabela_N

         item.SubItems(3) = "" & Format(ValorVenda_N, strFormatacao2Digitos)
         item.SubItems(4) = "" & Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
         item.SubItems(5) = "" & Format(Trim(TabTemp.Fields("valor_desconto").Value), strFormatacao2Digitos)
         item.SubItems(6) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
         'Item.SubItems(7) = "" '& Trim(TabTemp.Fields("CODG_ncm").Value)
         'Item.SubItems(8) = "" & lstPedido.selectedItem.Text

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
End Sub

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   HORA_INI = Time

   If EXISTE_OBJ_BANCO("RETAGUARDA", "RELLIBERA", "U") = True Then
      strSQL = "drop table RELLIBERA"
      CONECTA_RETAGUARDA.Execute strSQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "RELLIBERA", "U") = False Then
      strSQL = "create table RELLIBERA"
      strSQL = strSQL & " ("
         strSQL = strSQL & " RELLIBERA_ID        bigint        not null,"
         strSQL = strSQL & " PEDIDO_ID          bigint         not null,"
         strSQL = strSQL & " CLIENTE_ID         bigint         not null,"
         strSQL = strSQL & " VENDEDOR_ID        bigint         not null,"
         strSQL = strSQL & " tipovenda_ID       bigint         not null,"
         strSQL = strSQL & " DT_VENDA           datetime       not null,"

         strSQL = strSQL & " PRECO_VENDA        float          not null,"
         strSQL = strSQL & " VALOR_VENDIDO      float          NOT null    ,"
         strSQL = strSQL & " CLIENTE            nvarchar(50)   NOT null    ,"

         strSQL = strSQL & " QTDE_VENDIDA       float          not null,"
         strSQL = strSQL & " PRODUTO_ID         BIGINT         NOT NULL,"
         strSQL = strSQL & " CodgProduto        nvarchar(100)  NOT NULL,"
         strSQL = strSQL & " DescProduto        nvarchar(100)  NOT NULL,"
         strSQL = strSQL & " DescTipoVenda      nvarchar(30)   "

         strSQL = strSQL & " constraint PK_RELLIBERA primary key (RELLIBERA_ID)"
      strSQL = strSQL & " )"
      CONECTA_RETAGUARDA.Execute strSQL
   End If

   strSQL = "delete from RELLIBERA"
   CONECTA_RETAGUARDA.Execute strSQL

   Dim TabVaca      As New ADODB.Recordset

   Dim Conta_Produto_N  As Long
   Dim ValorTabela_N    As Double
   Dim VALOR_ITEM_N     As Double
   Dim DESCONTO_ITEM    As Double
   Dim DESCONTO_CABEÇA  As Double
   Dim VALOR_CUSTO_N    As Double
   Dim CARTAO_ID        As Long

   VALOR_ITEM_N = 0
   QTDE_N = 0
   DESCONTO_ITEM = 0
   DESCONTO_CABEÇA = 0
   VALOR_CUSTO_N = 0
   CARTAO_ID = 0

   Me.Enabled = False
   CONT_N = 0

   QTDE_N = 0
   VALOR_ITEM_N = 0
   VALOR_CUSTO_N = 0
   VALOR_DESCONTO_N = 0
   NUMR_ID_N = 0
   PEDIDO_ID_N = 0
   CONT_N = 0
   ValorTabela_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   CONT_N = 0

   Me.Enabled = False
   Me.KeyPreview = False

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

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

'set pegar pelo campo de quem liberou a venda com valor diferente]

   SQL = "select PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.VENDEDOR_ID, PEDIDO.USUARIO_ID, PEDIDO.ESTABELECIMENTO_ID,"
   SQL = SQL & " PEDIDO.TABELAPRECO_ID, PEDIDO.DT_REQ, PEDIDO.STATUS, PEDIDO.USUARIO_LIBERA_VENDA, PEDIDO.VALOR_DESCONTO as DescontoCabeca,"
   SQL = SQL & " PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_TOTAL, PEDIDO.STATUS, "
   SQL = SQL & " PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM as ValorVendido, PEDIDOITEM.VALOR_DESCONTO AS DescontoItem,"
   SQL = SQL & " PEDIDOITEM.PRECO_CUSTO as PrCustoItem,"
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO as DescricaoProduto, PRODUTO.FAMILIAPRODUTO_ID,"
   SQL = SQL & " PRODUTO.PRECO_CUSTO as PrCustoProduto, PRODUTO.PRECO_Venda, ITEMLANCAMENTO.FORMAPAGTO_ID"

      SQL = SQL & " from PEDIDOITEM WITH (NOLOCK) "

      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " INNER JOIN PEDIDO WITH (NOLOCK) "
      SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = PEDIDO.PEDIDO_ID "
      SQL = SQL & " INNER JOIN LANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = LANCAMENTO.NUMR_DOC "
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
      SQL = SQL & " AND LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

   SQL = SQL & " where PEDIDO.PEDIDO_ID > 0 "

'SQL = SQL & " and USUARIO_LIBERA_VENDA > 0"

   SQL = SQL & " and PEDIDO.estabelecimento_id = " & cmbEstabAUX.Text

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and PEDIDOITEM.produto_id = " & PRODUTO_ID_N

   If Trim(txtPedido.Text) <> "" Then _
      SQL = SQL & " and PEDIDO.pedido_id = " & txtPedido.Text

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then _
      If CLIENTE_ID_N > 0 Then _
         SQL = SQL & " and cliente_id = " & CLIENTE_ID_N
   txtCNPJCPF.PromptInclude = True

   If Trim(cmbVend.Text) <> "" Then _
      SQL = SQL & " and vendedor_id = " & cmbVendAux.Text

   If Trim(cmbSituacaoAUX.Text) <> "" Then
      If Trim(cmbSituacaoAUX.Text) = "'7','5','3'" Then _
         SQL = SQL & " and numr_nota > 0 "

      SQL = SQL & " and PEDIDO.status in (" & Trim(cmbSituacaoAUX.Text) & ")"
   End If

   'If Trim(cmbAuxForma.Text) <> "" Then _
      SQL = SQL & " and PEDIDO.tipovenda_id = " & cmbAuxForma.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_req <= '" & txtDtFim.Text & "'"
   End If

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   SQL = SQL & " order by PEDIDO.PEDIDO_ID desc"

'============================
   If TabVaca.State = 1 Then _
      TabVaca.Close

   TabVaca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabVaca.EOF Then
      If TabVaca.State = 1 Then _
         TabVaca.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   INDR_PRI = True

   While Not TabVaca.EOF
      CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
      txtReg.Text = CONTA_REGISTRO_N
      txtReg.Refresh
      DoEvents

      If chkTabela.Value = 1 Then
         INDR_PRI = False
         If Not IsNull(TabVaca.Fields("TABELAPRECO_ID").Value) Then
            INDR_PRI = False

            VALOR_ITEM_N = 0 & TabVaca.Fields("preco_venda").Value
            FORMAPAGTO_ID_N = 1
            If Not IsNull(TabVaca.Fields("formapagto_id").Value) Then
               FORMAPAGTO_ID_N = TabVaca.Fields("formapagto_id").Value
               ValorTabela_N = 0 & (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(TabVaca.Fields("produto_id").Value, TabVaca.Fields("tabelapreco_id").Value, FORMAPAGTO_ID_N))
               If ValorTabela_N > 0 Then _
                  VALOR_ITEM_N = ValorTabela_N
            End If
      
            If TabVaca.Fields("valorvendido").Value < VALOR_ITEM_N Then _
               INDR_PRI = True
      
            Else  'vai pegar do cadastro de produto
               If TabProduto.State = 1 Then _
                  TabProduto.Close
      
               SQL = "select preco_custo,preco_atacado,preco_venda from PRODUTO "
               SQL = SQL & " where produto_id = " & TabVaca.Fields("produto_id").Value
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then
                  INDR_PRI = False
                  VALOR_ITEM_N = 0 & TabProduto.Fields("preco_venda").Value
                  If TabVaca.Fields("valorvendido").Value < VALOR_ITEM_N Then _
                     INDR_PRI = True
               End If
               If TabProduto.State = 1 Then _
                  TabProduto.Close
         End If
      End If

      If PEDIDO_ID_N <> TabVaca.Fields("pedido_id").Value And INDR_PRI = True Then
         PEDIDO_ID_N = TabVaca.Fields("pedido_id").Value

         NUMR_SEQ_N = NUMR_SEQ_N + 1
         Set item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, TabVaca.Fields("PEDIDO_ID").Value)
         item.SubItems(1) = "" & Trim(TabVaca!NOME_CLIENTE)

         If IsNull(TabVaca!NOME_CLIENTE) Or Trim(TabVaca!NOME_CLIENTE) = "" Then
            If TabCliente.State = 1 Then _
               TabCliente.Close
            SQL = "select nome from CLIENTE WITH (NOLOCK)"
            SQL = SQL & " where cgccpf = '" & Trim(TabVaca.Fields("CNPJCPF").Value) & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               item.SubItems(1) = "" & TabCliente!NOME
            If TabCliente.State = 1 Then _
               TabCliente.Close
         End If

'================

         VALOR_DESCONTO_N = 0

         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabVaca.Fields("pedido_id").Value
         'SQL = SQL & " and tipo_reg = 'PC' "
         SQL = SQL & " and status <> 'C' "
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedidoItem.EOF Then _
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

        If Not IsNull(TabVaca.Fields("DescontoCabeca").Value) Then _
           VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabVaca.Fields("DescontoCabeca").Value

         VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

         'BUSCA VALOR TOTAL VENDA
         VALOR_ITEM_N = 0

         SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabVaca.Fields("pedido_id").Value
         SQL = SQL & " and status <> 'C' "
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
            VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

         item.SubItems(2) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
         item.SubItems(3) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
         item.SubItems(4) = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
'=====================

         item.SubItems(5) = TabVaca!DT_REQ
         item.SubItems(6) = ""

         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         SQL = "select descricao from TIPOVENDA WITH (NOLOCK)"
         SQL = SQL & " where tipovenda_id = " & TabVaca!TIPOVENDA_ID
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            item.SubItems(6) = TabDESCR!DESCRICAO
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         item.SubItems(7) = ""

         If TabUSU.State = 1 Then _
            TabUSU.Close
         SQL = "select * from vwVendedor WITH (NOLOCK)"
         SQL = SQL & " where vendedor_id = " & TabVaca.Fields("vendedor_id").Value
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then _
            item.SubItems(7) = TabUSU!DESCRICAO
         If TabUSU.State = 1 Then _
            TabUSU.Close

         item.SubItems(8) = ""

         If Not IsNull(TabVaca.Fields("Status")) Then
            If TabVaca.Fields("Status") = 2 Then
               If TabVaca.Fields("tipo_registro") = "O" Then
                  item.SubItems(8) = "Orcamento"
                  Else: item.SubItems(8) = "Pedido"
               End If
            End If
            If TabVaca.Fields("Status").Value = 3 Then _
               item.SubItems(8) = "3-Faturado"
            If TabVaca.Fields("Status").Value = 4 Then _
               item.SubItems(8) = "4-Cupom"
            If TabVaca.Fields("Status").Value = 5 Then _
               item.SubItems(8) = "5-Faturado"
            If TabVaca.Fields("Status").Value = 7 Then _
               item.SubItems(8) = "7-Cupom Fiscal"
            If TabVaca.Fields("Status").Value = 9 Then _
               item.SubItems(8) = "9-Cancelado"
         End If

         SQL = "select sum(qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabVaca.Fields("pedido_id").Value
         SQL = SQL & " and tipo_reg = 'PC' "
         SQL = SQL & " and status <> 'C' "
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
            Conta_Produto_N = Conta_Produto_N + TabPedidoItem.Fields(0).Value
         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         txtQtdeProd.Text = Conta_Produto_N
         txtQtdeProd.Refresh

         txtTotDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
         txtTotVendas.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
         txtTotalVenda.Text = Format(VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

'==============
         If TabVaca.Fields("status").Value = 1 Then
            item.ForeColor = vbRed
            item.ListSubItems(1).ForeColor = vbRed
            item.ListSubItems(2).ForeColor = vbRed
            item.ListSubItems(3).ForeColor = vbRed
            item.ListSubItems(4).ForeColor = vbRed
            item.ListSubItems(5).ForeColor = vbRed
            item.ListSubItems(6).ForeColor = vbRed
            item.ListSubItems(7).ForeColor = vbRed
            item.ListSubItems(8).ForeColor = vbRed
            item.SubItems(8) = "" & "Em Aberto - 1"
         End If
         If TabVaca.Fields("status").Value = 2 Then
            item.ForeColor = vbBlue
            item.ListSubItems(1).ForeColor = vbBlue
            item.ListSubItems(2).ForeColor = vbBlue
            item.ListSubItems(3).ForeColor = vbBlue
            item.ListSubItems(4).ForeColor = vbBlue
            item.ListSubItems(5).ForeColor = vbBlue
            item.ListSubItems(6).ForeColor = vbBlue
            item.ListSubItems(7).ForeColor = vbBlue
            item.ListSubItems(8).ForeColor = vbBlue
            item.SubItems(8) = "" & "A Faturar - 2"
         End If
         If TabVaca.Fields("status").Value = 3 Then
            item.ForeColor = vbBlack
            item.ListSubItems(1).ForeColor = vbBlack
            item.ListSubItems(2).ForeColor = vbBlack
            item.ListSubItems(3).ForeColor = vbBlack
            item.ListSubItems(4).ForeColor = vbBlack
            item.ListSubItems(5).ForeColor = vbBlack
            item.ListSubItems(6).ForeColor = vbBlack
            item.ListSubItems(7).ForeColor = vbBlack
            item.ListSubItems(8).ForeColor = vbBlack
            item.SubItems(8) = "" & "Faturado - 3"
         End If
         If TabVaca.Fields("status").Value = 5 Then
            item.ForeColor = vbBlack
            item.ListSubItems(1).ForeColor = vbBlack
            item.ListSubItems(2).ForeColor = vbBlack
            item.ListSubItems(3).ForeColor = vbBlack
            item.ListSubItems(4).ForeColor = vbBlack
            item.ListSubItems(5).ForeColor = vbBlack
            item.ListSubItems(6).ForeColor = vbBlack
            item.ListSubItems(7).ForeColor = vbBlack
            item.ListSubItems(8).ForeColor = vbBlack
            item.SubItems(8) = "" & "Faturado - 5"
         End If
         If TabVaca.Fields("status").Value = 6 Then
            item.ForeColor = vbBlack
            item.ListSubItems(10).ForeColor = vbYellow
            item.SubItems(8) = "" & "Não Contabilizado"
         End If
         If TabVaca.Fields("status").Value = 7 Then
            item.ForeColor = vbMagenta
            item.ListSubItems(1).ForeColor = vbMagenta
            item.ListSubItems(2).ForeColor = vbMagenta
            item.ListSubItems(3).ForeColor = vbMagenta
            item.ListSubItems(4).ForeColor = vbMagenta
            item.ListSubItems(5).ForeColor = vbMagenta
            item.ListSubItems(6).ForeColor = vbMagenta
            item.ListSubItems(7).ForeColor = vbMagenta
            item.ListSubItems(8).ForeColor = vbMagenta
            item.SubItems(8) = "" & "Cupom Fiscal - 7"
         End If
         If TabVaca.Fields("status").Value = 9 Then
            item.ListSubItems(8).ForeColor = &HC0E0FF
            item.SubItems(8) = "" & "Cancelado - 9"
         End If
      End If   'If PEDIDO_ID_N <> TabVaca.Fields("pedido_id").Value And INDR_PRI = True Then

      TabVaca.MoveNext
   Wend
   If TabVaca.State = 1 Then _
      TabVaca.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time


   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub
