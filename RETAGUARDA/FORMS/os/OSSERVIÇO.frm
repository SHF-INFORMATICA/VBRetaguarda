VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSServico 
   Caption         =   "O.S. (frmOSSERVIÇO)"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "OSSERVIÇO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotOS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   10620
      TabIndex        =   70
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtTotDescontoProduto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2310
      TabIndex        =   69
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox txtTotDescontoServico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2310
      TabIndex        =   68
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox txtTotServico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5340
      TabIndex        =   67
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox txtTotProduto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   5340
      TabIndex        =   66
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox txtTotGeralProduto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   8175
      TabIndex        =   65
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox txtTotGeralServico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8175
      TabIndex        =   64
      Top             =   8160
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Produtos Ordem de Serviço"
      ForeColor       =   &H00008000&
      Height          =   2415
      Left            =   50
      TabIndex        =   52
      Top             =   5760
      Width           =   11775
      Begin VB.CommandButton cmdCadProduto 
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
         Height          =   350
         Left            =   2880
         Picture         =   "OSSERVIÇO.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Consulta Cadastro Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtDescProduto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   56
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   20
         ToolTipText     =   "Informe Código Produto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtValorProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7800
         TabIndex        =   25
         ToolTipText     =   "Valor Venda Produto"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDescontoProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         ToolTipText     =   "Desconto Produto"
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbVendedor 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10080
         TabIndex        =   21
         ToolTipText     =   "Responsável Venda"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   23
         ToolTipText     =   "Digite a Quantidade"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtTotalProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   10440
         TabIndex        =   55
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbVendedorAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   10080
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdProduto 
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
         Height          =   350
         Left            =   2444
         Picture         =   "OSSERVIÇO.frx":7D3C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin MSComctlLib.ListView lstProduto 
         Height          =   1005
         Left            =   45
         TabIndex        =   57
         Top             =   1320
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   1773
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   32768
         BackColor       =   16777215
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qtde"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Desconto"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Total"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Vendedor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "OSPECA_ID"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Garantia"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "produto_ID"
            Object.Width           =   2
         EndProperty
      End
      Begin MSMask.MaskEdBox txtGarantia 
         Height          =   375
         Left            =   1080
         TabIndex        =   22
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Garantia:"
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   85
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   90
         TabIndex        =   63
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Item = "
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   6600
         TabIndex        =   62
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto ="
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   4260
         TabIndex        =   61
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável:"
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   8730
         TabIndex        =   60
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtde ="
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   0
         Left            =   2520
         TabIndex        =   59
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Item ="
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   9240
         TabIndex        =   58
         Top             =   840
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serviços Ordem de Serviço"
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   50
      TabIndex        =   43
      Top             =   2880
      Width           =   11775
      Begin VB.ComboBox cmbMecanicoAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   9600
         TabIndex        =   87
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbSitServ 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Selecione Mecanico para tarefa"
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdCadServiço 
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
         Height          =   350
         Left            =   2400
         Picture         =   "OSSERVIÇO.frx":873E
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Consulta Cadastro Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtTotalTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   10320
         TabIndex        =   45
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cmbMecanico 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9600
         TabIndex        =   14
         ToolTipText     =   "Selecione Mecanico para tarefa"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtDescontoTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         ToolTipText     =   "Desconto Serviço"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtValorTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   19
         ToolTipText     =   "Informe Valor Serviço"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtDescTarefa 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         ToolTipText     =   "Descrição Serviço"
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox txtServico 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   12
         ToolTipText     =   "Digite Código Tarefa ou 0 para Diversar"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdServico 
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
         Height          =   350
         Left            =   1920
         Picture         =   "OSSERVIÇO.frx":A868
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin MSComctlLib.ListView lstServiço 
         Height          =   1005
         Left            =   45
         TabIndex        =   46
         Top             =   1800
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   1773
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   12582912
         BackColor       =   16777215
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Serviço"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Desconto"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Mecanico"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DtInicio"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DtFim"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   375
         Left            =   6480
         TabIndex        =   17
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox txtDtInicio 
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Inicio:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1995
         TabIndex        =   86
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Fim:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5760
         TabIndex        =   82
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Serviço ="
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8160
         TabIndex        =   51
         Top             =   1320
         Width           =   1965
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8520
         TabIndex        =   50
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto ="
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1680
         TabIndex        =   49
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr.Serviço ="
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5115
         TabIndex        =   48
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label lblTarefa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serviço:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados Equipamento"
      ForeColor       =   &H00800080&
      Height          =   1335
      Left            =   50
      TabIndex        =   34
      Top             =   1560
      Width           =   11775
      Begin PVMaskEditLib.PVMaskEdit txtPlaca 
         Height          =   375
         Left            =   1560
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _Version        =   524288
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "@@@-####"
      End
      Begin VB.TextBox txtKM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdCadCli 
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
         Height          =   350
         Left            =   6960
         Picture         =   "OSSERVIÇO.frx":B26A
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Consulta Cadastro Veículo"
         Top             =   360
         Width           =   405
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
         Height          =   350
         Left            =   6530
         Picture         =   "OSSERVIÇO.frx":D394
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtEqp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Informe Kilometragem atual"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   10080
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtMODELO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtANO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdCadPlaca 
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
         Height          =   350
         Left            =   3360
         Picture         =   "OSSERVIÇO.frx":DD96
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Consulta Cadastro Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdConsultaPlaca 
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
         Height          =   350
         Left            =   2985
         Picture         =   "OSSERVIÇO.frx":FEC0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KM:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   7320
         TabIndex        =   91
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   405
         TabIndex        =   42
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   8880
         TabIndex        =   41
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ano/Mod.:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   4575
         TabIndex        =   40
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lblEqp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identificação:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   165
         TabIndex        =   39
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   3780
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem de Serviço"
      Height          =   975
      Left            =   50
      TabIndex        =   27
      Top             =   600
      Width           =   11775
      Begin VB.OptionButton optVeiculo 
         Caption         =   "&Veículo"
         Height          =   255
         Left            =   10080
         TabIndex        =   90
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optEqp 
         Caption         =   "&Equipamento"
         Height          =   255
         Left            =   10080
         TabIndex        =   89
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox cmbTipoOSAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   5880
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbSituacaoAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   83
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbSituacao 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8280
         TabIndex        =   4
         ToolTipText     =   "Selecione situação Ordem de Serviço"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbTipoOS 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5880
         TabIndex        =   3
         ToolTipText     =   "Selecione Tipo Ordem de Serviço"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtOS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "<<Enter>> gerar nova O.S."
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbConsultor 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3600
         TabIndex        =   2
         ToolTipText     =   "Selecione Consultor Tecnico"
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cmbConsultorAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4200
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtDtOS 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação"
         Height          =   240
         Left            =   8280
         TabIndex        =   33
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo O.S."
         Height          =   240
         Left            =   5880
         TabIndex        =   32
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblOs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O.S."
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblCt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consultor/Vendedor"
         Height          =   240
         Left            =   3600
         TabIndex        =   30
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   240
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   435
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":108C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":10D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":11032
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":11486
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":118DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":11BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":1204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":1236E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":12F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVIÇO.frx":18552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1958
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "cons"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cliente"
            Key             =   "cli"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Termo"
            Key             =   "termo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Observações"
            Key             =   "obs"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   10200
         TabIndex        =   92
         Top             =   240
         Width           =   1455
      End
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
      DesignHeight    =   8865
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total O.S.="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9240
      TabIndex        =   77
      Top             =   8280
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desc. Produto = "
      Height          =   240
      Left            =   120
      TabIndex        =   76
      Top             =   8520
      Width           =   2100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desc. Serviço = "
      Height          =   240
      Left            =   150
      TabIndex        =   75
      Top             =   8160
      Width           =   2070
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal Serviço = "
      Height          =   240
      Left            =   3375
      TabIndex        =   74
      Top             =   8160
      Width           =   1875
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal Produto = "
      Height          =   240
      Left            =   3345
      TabIndex        =   73
      Top             =   8520
      Width           =   1905
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Produto = "
      Height          =   240
      Left            =   6555
      TabIndex        =   72
      Top             =   8520
      Width           =   1530
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Serviço = "
      Height          =   240
      Left            =   6585
      TabIndex        =   71
      Top             =   8160
      Width           =   1500
   End
End
Attribute VB_Name = "frmOSServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim EQUIPAMENTO_ID_N       As Long
   Dim VEICULO_ID_N           As Long
   Dim OSPECA_ID_N            As Long
   Dim TAREFA_ID_N            As Long
   Dim Situação_Ordem_Serviço As String
   Dim DT_FECHAMENTO_OS       As Date

Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   LIMPA_OS
   CARREGA_COMBOS

   EQP_VEICULO

End Sub

Private Sub Form_Unload(Cancel As Integer)
   MOSTRA_RODAPE "", "", "", "", ""
End Sub

Private Sub optEqp_Click()
   EQP_VEICULO
End Sub

Private Sub optVeiculo_Click()
   EQP_VEICULO
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "obs"
         If Trim(txtOs.Text) <> "" Then
            OS_ID_N = 0 & txtOs.Text

            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            SQL = "select * from OS "
            SQL = SQL & " where os_id = " & OS_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCabeca.EOF Then
               CHAMADA_A = "OBS"

               frmOBS.Show 1
               Else
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close
                  MsgBox "O.S. não informada !!!"
            End If
            If TabCabeca.State = 1 Then _
               TabCabeca.Close
         End If
         CHAMADA_A = ""
      Case "termo"
         If Trim(txtOs.Text) <> "" Then
            OS_ID_N = 0 & txtOs.Text

            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            SQL = "select * from OS "
            SQL = SQL & " where os_id = " & OS_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCabeca.EOF Then
               CHAMADA_A = "TERMOGARANTIA"

               frmOBS.Show 1
'Garantia de Serviços 30 Dias
'Garantia de Peças 90 dias sem impurezas no óleo
               Else
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close
                  MsgBox "O.S. não informada !!!"
            End If
            If TabCabeca.State = 1 Then _
               TabCabeca.Close
         End If
         CHAMADA_A = ""
      Case "matar"
         EXCLUIR_OS
      Case "cons"
         SQL3 = ""
         frmOSConsulta.Show 1
         If SQL3 <> "" Then _
            If IsNumeric(SQL3) Then _
               txtOs.Text = SQL3
         SQL3 = ""
         txtOs.SetFocus
      Case "sair"
        Unload Me
      Case "limpar"
         LIMPA_OS
         txtOs.SetFocus
      Case "gravar"
         If CHECA_DADOS_OS = True Then
            GRAVA_OS
            LIMPA_OS
            txtOs.SetFocus
         End If
      Case "imprimir"
         If Trim(txtOs.Text) <> "" Then _
            If IsNumeric(txtOs.Text) Then _
               IMPRIMIR_ORDEM_SERVICO txtOs.Text, "SERVIÇO", Trim(txtCliente.Text)
      Case "excluir"
      Case "cli"
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstProduto_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtOs.Text) <> "" Then
      If IsNumeric(txtOs.Text) Then
         If Trim(lstProduto.SelectedItem.Text) <> "" Then
            OSPECA_ID_N = 0 & Trim(lstProduto.SelectedItem.ListSubItems(6).Text)
            txtProduto.Text = "" & Trim(lstProduto.SelectedItem.Text)
            'txtProduto.SetFocus
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_Click"
End Sub

Private Sub lstProduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Trim(txtOs.Text) <> "" Then
            If IsNumeric(txtOs.Text) Then
               If Trim(lstProduto.SelectedItem.Text) <> "" Then
                  'If IsNumeric(lstProduto.SelectedItem.Text) Then
                     EXCLUIR_PRODUTO_ITEM
                  'End If
               End If
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_KeyDown"
End Sub

Private Sub lstServiço_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Trim(txtOs) <> "" Then
            If IsNumeric(txtOs) Then
               If Trim(lstServiço.SelectedItem.Text) <> "" Then
                  If IsNumeric(lstServiço.SelectedItem.Text) Then
                     EXCLUIR_SERVIÇO_ITEM
                  End If
               End If
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstServiço_KeyDown"
End Sub

Private Sub cmbConsultor_GotFocus()
   cmbConsultor.SelStart = 0
   cmbConsultor.SelLength = Len(cmbConsultor.Text)
   cmbConsultor.BackColor = &HC0FFFF
End Sub

Private Sub cmbMecanico_GotFocus()
   cmbMecanico.SelStart = 0
   cmbMecanico.SelLength = Len(cmbMecanico.Text)
   cmbMecanico.BackColor = &HC0FFFF
End Sub

Private Sub cmbSituacao_GotFocus()
   cmbSituacao.SelStart = 0
   cmbSituacao.SelLength = Len(cmbSituacao.Text)
   cmbSituacao.BackColor = &HC0FFFF
End Sub

Private Sub cmbSituacao_LostFocus()
   cmbSituacao.BackColor = &HFFFFFF
   If Trim(cmbSituacao.Text) = "" Then _
      cmbSituacao.ListIndex = 0
End Sub

Private Sub cmbTipoOS_GotFocus()
   cmbTipoOS.SelStart = 0
   cmbTipoOS.SelLength = Len(cmbTipoOS.Text)
   cmbTipoOS.BackColor = &HC0FFFF
End Sub

Private Sub cmbVENDEDOR_GotFocus()
   cmbVendedor.SelStart = 0
   cmbVendedor.SelLength = Len(cmbVendedor.Text)
   cmbVendedor.BackColor = &HC0FFFF
End Sub

Private Sub cmdCadProduto_Click()
   frmCADASTROPRODUTO.Show 1
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Text = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = ""
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      'Call TXTCNPJCPF_LostFocus

      'Exit Sub
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsCli_Click"
End Sub

Private Sub cmdCadCli_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaCadastro.Show 1
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmdCadServiço_Click()
   frmOSServicoCadastro.Show 1
End Sub

Private Sub cmbConsultor_LostFocus()
   cmbConsultor.BackColor = &HFFFFFF
   If Trim(cmbConsultor.Text) = "" Then _
      cmbConsultor.ListIndex = 0
End Sub

Private Sub cmbTipoOS_LostFocus()
   cmbTipoOS.BackColor = &HFFFFFF
   If Trim(cmbTipoOS.Text) = "" Then _
      cmbTipoOS.ListIndex = 0
End Sub

Private Sub cmbMecanico_LostFocus()
   cmbMecanico.BackColor = &HFFFFFF
   'If Trim(cmbMecanico.Text) = "" Then _
      cmbMecanico.ListIndex = 0
End Sub

Private Sub cmbVendedor_LostFocus()
   cmbVendedor.BackColor = &HFFFFFF
   If Trim(cmbVendedor.Text) = "" Then _
      cmbVendedor.ListIndex = 0
End Sub

Private Sub cmbSituacao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If EQP_VEICULO = False Then
         txtPlaca.SetFocus
         Else: txtEqp.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbSituacao_KeyPress"
End Sub

Private Sub cmbConsultor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTipoOS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbConsultor_KeyPress"
End Sub

Private Sub cmbConsultor_Click()
On Error Resume Next

   cmbConsultorAUX.ListIndex = cmbConsultor.ListIndex

Err.Clear
End Sub

Private Sub cmbtipoos_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSituacao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtipoos_KeyPress"
End Sub

Private Sub cmbtipoos_Click()
On Error Resume Next

   cmbTipoOSAUX.ListIndex = cmbTipoOS.ListIndex

Err.Clear
End Sub

Private Sub cmbSituacao_Click()
On Error Resume Next

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex
   INDR_RECEITA = cmbSituacaoAUX.Text

Err.Clear
End Sub

Private Sub txtAno_GotFocus()
   txtANO.SelStart = 0
   txtANO.SelLength = Len(txtANO.Text)
   txtANO.BackColor = &HC0FFFF
End Sub

Private Sub txtANO_LostFocus()
   txtANO.BackColor = &HFFFFFF
End Sub

Private Sub txtKM_GotFocus()
   txtKM.SelStart = 0
   txtKM.SelLength = Len(txtKM.Text)
   txtKM.BackColor = &HC0FFFF
End Sub

Private Sub txtKM_LostFocus()
   txtKM.BackColor = &HFFFFFF
End Sub

Private Sub txtKM_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtServico.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtKM_KeyPress"
End Sub

Private Sub txtCliente_GotFocus()
   txtCliente.SelStart = 0
   txtCliente.SelLength = Len(txtCliente.Text)
   txtCliente.BackColor = &HC0FFFF
End Sub

Private Sub txtCliente_LostFocus()
   txtCliente.BackColor = &HFFFFFF
End Sub

'==================cgccpf
Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC-SAIR", "F7-Consulta Clientes", "Inform CNPJ/CPF Cliente e Tecle <<Enter>>", "", ""
   txtCliente.Enabled = True
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then _
      txtCNPJCPF.Mask = "##############"

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Text)
   txtCNPJCPF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.Text = ""
            txtCNPJCPF.Mask = "##############"

            txtCNPJCPF.Text = CNPJCPF_A

            Exit Sub
         End If
         CNPJCPF_A = ""
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         txtCliente.Enabled = True
         txtCliente.SetFocus
         Else
            'txtServico.SetFocus
            txtKM.SetFocus
            txtCliente.Enabled = False
      End If
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      UCase (txtProduto.Text)
      txtServico.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCLIENTE_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Text = "99999999999"
      Else
         If Trim(txtCNPJCPF.Text) <> "99999999999" Then _
            TRATA_PESSOA txtCNPJCPF.Text
   End If

   txtCNPJCPF.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtDesc_GotFocus()
   txtDesc.SelStart = 0
   txtDesc.SelLength = Len(txtDesc.Text)
   txtDesc.BackColor = &HC0FFFF
End Sub

Private Sub txtDesc_LostFocus()
   txtDesc.BackColor = &HFFFFFF
End Sub

Private Sub txtDescontoProduto_GotFocus()
   txtDESCONTOPRODUTO.SelStart = 0
   txtDESCONTOPRODUTO.SelLength = Len(txtDESCONTOPRODUTO.Text)
   txtDESCONTOPRODUTO.BackColor = &HC0FFFF
End Sub

Private Sub txtDescontoTarefa_GotFocus()
   txtDescontoTarefa.SelStart = 0
   txtDescontoTarefa.SelLength = Len(txtDescontoTarefa.Text)
   txtDescontoTarefa.BackColor = &HC0FFFF
End Sub

Private Sub txtDescProduto_GotFocus()
   txtDESCPRODUTO.SelStart = 0
   txtDESCPRODUTO.SelLength = Len(txtDESCPRODUTO.Text)
   txtDESCPRODUTO.BackColor = &HC0FFFF
End Sub

Private Sub txtDescProduto_LostFocus()
   txtDESCPRODUTO.BackColor = &HFFFFFF
End Sub

Private Sub txtDescTarefa_GotFocus()
   txtDescTarefa.SelStart = 0
   txtDescTarefa.SelLength = Len(txtDescTarefa.Text)
   txtDescTarefa.BackColor = &HC0FFFF
End Sub

Private Sub txtDescTarefa_LostFocus()
   txtDescTarefa.BackColor = &HFFFFFF
End Sub

Private Sub TXTDTFIM_GotFocus()
   txtDtFim.PromptInclude = True
   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim.Text)
   txtDtFim.BackColor = &HC0FFFF
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescontoTarefa.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

End Sub

Private Sub txtDtFim_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
         txtDtFim.Text = ""
      txtDtFim.PromptInclude = True
   End If

   txtDtFim.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTDTFIM_LostFocus"
End Sub

Private Sub txtDTINICIO_GotFocus()
   txtDtInicio.PromptInclude = False
'   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")

   txtDtInicio.PromptInclude = True
   txtDtInicio.SelStart = 0
   txtDtInicio.SelLength = Len(txtDtInicio.Text)
   txtDtInicio.BackColor = &HC0FFFF
End Sub

Private Sub txtDTINICIO_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtDtFim.Enabled = True Then
         txtDtFim.SetFocus
         Else: txtDescontoTarefa.SetFocus
      End If
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

End Sub

Private Sub txtDTINICIO_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtInicio.PromptInclude = True
   If Not IsDate(txtDtInicio.Text) Then
      txtDtInicio.PromptInclude = False
         txtDtInicio.Text = ""
      txtDtInicio.PromptInclude = True
   End If

   txtDtInicio.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINICIO_LostFocus"
End Sub

Private Sub txtDtOS_GotFocus()
   txtDtOS.SelStart = 0
   txtDtOS.SelLength = Len(txtDtOS.Text)
   txtDtOS.BackColor = &HC0FFFF
   
   cmbConsultor.SetFocus
End Sub

Private Sub txtDtOS_LostFocus()
   txtDtOS.BackColor = &HFFFFFF
End Sub

Private Sub txtPlaca_GotFocus()
   txtPlaca.BackColor = &HC0FFFF
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtPlaca.Text = "" & CONSULTA_EQP_VEICULO
         MOSTRA_VEICULO
         txtPlaca.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPLACA_KeyDown"
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtplaca_KeyPress"
End Sub

Private Sub txtPLACA_LostFocus()
   If Trim(txtPlaca.Text) <> "" Then _
      MOSTRA_VEICULO
   txtPlaca.BackColor = &HFFFFFF
End Sub

Private Sub txtEqp_GotFocus()
   txtEqp.SelStart = 0
   txtEqp.SelLength = Len(txtEqp.Text)
   txtEqp.BackColor = &HC0FFFF
End Sub

Private Sub txtMarca_GotFocus()
   txtMarca.SelStart = 0
   txtMarca.SelLength = Len(txtMarca.Text)
   txtMarca.BackColor = &HC0FFFF
End Sub

Private Sub txtMarca_LostFocus()
   txtMarca.BackColor = &HFFFFFF
End Sub

Private Sub txtmodelo_GotFocus()
   txtMODELO.SelStart = 0
   txtMODELO.SelLength = Len(txtMODELO.Text)
   txtMODELO.BackColor = &HC0FFFF
End Sub

Private Sub txtMODELO_LostFocus()
   txtMODELO.BackColor = &HFFFFFF
End Sub

Private Sub txtOs_GotFocus()
   txtOs.SelStart = 0
   txtOs.SelLength = Len(txtOs.Text)
   txtOs.BackColor = &HC0FFFF
End Sub

Private Sub txtOS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbConsultor.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOS_KeyPress"
End Sub

Private Sub txtOs_LostFocus()
'On Error GoTo ERRO_TRATA

   INDR_PRI = False
   If Trim(txtOs.Text) <> "" Then
      If IsNumeric(txtOs.Text) Then
         OS_ID_N = txtOs.Text

         MOSTRA_OS

         Exit Sub
      End If
   End If

   If INDR_PRI = False Then
      GERA_PEDIDO_ID
      OS_ID_N = 0 & PEDIDO_ID_N
      txtOs.Text = OS_ID_N
   End If

   txtOs.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOS_KeyPress"
End Sub

Private Sub cmdConsultaPlaca_Click()
'On Error GoTo ERRO_TRATA

   If EQP_VEICULO = False Then
      INDR_OS_VEICULO = True
      txtPlaca.Text = "" & CONSULTA_EQP_VEICULO
      txtPlaca.SetFocus
      Else
         INDR_OS_VEICULO = False
         txtEqp.Text = "" & CONSULTA_EQP_VEICULO
         txtEqp.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsultaPlaca_Click"
End Sub

Private Sub cmdCadPlaca_Click()
   If EQP_VEICULO = False Then
      frmOSVeiculoCadastro.Show 1
      Else: frmOSEqpCadastro.Show 1
   End If
End Sub

Private Sub txtEqp_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtEqp.Text = "" & CONSULTA_EQP_VEICULO
         MOSTRA_EQP
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEqp_KeyDown"
End Sub

Private Sub txtEqp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEqp_KeyPress"
End Sub

Private Sub txtEqp_LostFocus()
   txtEqp.BackColor = &HFFFFFF
   If Trim(txtEqp.Text) = "" Then _
      txtEqp.Text = 1

   MOSTRA_EQP
End Sub

Private Sub cmdProduto_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then _
      txtProduto.Text = SQL3
   SQL3 = ""
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdPRODUTO_Click"
End Sub

Private Sub txtProduto_GotFocus()
   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto.Text)
   txtProduto.BackColor = &HC0FFFF
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then _
            txtProduto.Text = SQL3
         SQL3 = ""
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSERVIcO_Click"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtGARANTIA_GotFocus()
   txtGarantia.PromptInclude = False
   If Trim(txtGarantia.Text) = "" Then _
      txtGarantia.Text = DMA(Date, "I")
   txtGarantia.PromptInclude = True
  
   txtGarantia.SelStart = 0
   txtGarantia.SelLength = Len(txtGarantia.Text)
   txtGarantia.BackColor = &HC0FFFF
End Sub

Private Sub txtQTDE_GotFocus()
   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE.Text)
   txtQTDE.BackColor = &HC0FFFF
End Sub

Private Sub cmdSERVIcO_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSServicoConsulta.Show 1
   If SQL3 <> "" Then _
      txtServico.Text = SQL3
   SQL3 = ""
   txtServico.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSERVIcO_Click"
End Sub

Private Sub TXTSERVICO_GotFocus()
   txtDescontoTarefa.SelStart = 0
   txtDescontoTarefa.SelLength = Len(txtDescontoTarefa.Text)
   txtDescontoTarefa.BackColor = &HC0FFFF
End Sub

Private Sub txtServico_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
      Case vbKeyF7
         SQL3 = ""
         frmOSServicoConsulta.Show 1
         If SQL3 <> "" Then _
            txtServico.Text = SQL3
         SQL3 = ""
         txtServico.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtServico_KeyDown"
End Sub

Private Sub TXTSERVICO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescTarefa.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSERVICO_KeyPress"
End Sub

Private Sub txtDescTarefa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbMecanico.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescTarefa_KeyPress"
End Sub

Private Sub cmbmecanico_Click()
On Error Resume Next

   cmbMecanicoAUX.ListIndex = cmbMecanico.ListIndex

Err.Clear
End Sub

Private Sub cmbMecanico_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSitServ.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMecanico_KeyPress"
End Sub

Private Sub cmbSitServ_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtInicio.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbSitServ_KeyPress"
End Sub

Private Sub cmbSitServ_LostFocus()
   txtDtFim.Enabled = True
   If Trim(cmbSitServ.Text) <> "" Then
      If Left(cmbSitServ.Text, 1) <> "F" Then _
         txtDtFim.Enabled = False
      Else: cmbSitServ.Text = "P-Pendente"
   End If
End Sub

Private Sub txtDescontoTarefa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      VALOR_DESCONTO_N = 0 & txtDescontoTarefa.Text

      txtValorTarefa.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescontoTarefa_KeyPress"
End Sub

Private Sub txtdescontotarefa_LostFocus()
   If txtDescontoTarefa.Text = "" Then _
      txtDescontoTarefa.Text = 0

   txtDescontoTarefa.Text = Format(txtDescontoTarefa.Text, strFormatacao2Digitos)
   txtDescontoTarefa.BackColor = &HFFFFFF
End Sub

Private Sub txtTotalTarefa_LostFocus()
   txtTotalTarefa.BackColor = &HFFFFFF
End Sub

Private Sub txtTotDescontoProduto_GotFocus()
   txtServico.SetFocus
End Sub

Private Sub txtTotDescontoServico_GotFocus()
   txtServico.SetFocus
End Sub

Private Sub txtTotGeralProduto_GotFocus()
   txtServico.SetFocus
End Sub

Private Sub txtTotGeralServico_GotFocus()
On Error Resume Next
   txtServico.SetFocus
End Sub

Private Sub txtTotOS_GotFocus()
   txtServico.SetFocus
End Sub

Private Sub txtTotProduto_GotFocus()
   txtServico.SetFocus
End Sub

Private Sub txtTotServico_GotFocus()
   txtServico.SetFocus
End Sub

Private Sub txtValorProduto_GotFocus()
   txtValorProduto.SelStart = 0
   txtValorProduto.SelLength = Len(txtValorProduto.Text)
   txtValorProduto.BackColor = &HC0FFFF
End Sub

Private Sub txtValorTarefa_GotFocus()
   txtValorTarefa.SelStart = 0
   txtValorTarefa.SelLength = Len(txtValorTarefa.Text)
   txtValorTarefa.BackColor = &HC0FFFF
End Sub

Private Sub txtValorTarefa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then

      If CHECA_DADOS_OS = True Then _
         GRAVA_SERVIÇO

      KeyAscii = 0
      txtServico.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorTarefa_KeyPress"
End Sub

Private Sub txtValorTarefa_LostFocus()

   If txtValorTarefa.Text = "" Then _
      txtValorTarefa.Text = 0

   txtValorTarefa.Text = Format(txtValorTarefa.Text, strFormatacao2Digitos)

   TOTALIZA_CAMPOS
   txtValorTarefa.BackColor = &HFFFFFF
End Sub

Private Sub TXTSERVICO_LostFocus()
'On Error GoTo ERRO_TRATA

   INDR_PRI = False
   If Trim(txtServico.Text) <> "" Then
      If IsNumeric(txtServico.Text) Then
         Else: INDR_PRI = True
      End If
      Else: INDR_PRI = True
   End If

If Trim(txtOs.Text) <> "" Then _
   If INDR_PRI = True Then _
      txtServico.Text = MAX_ID("osservico_id", "OSSERVICO", "OS_ID", txtOs.Text, "", "")

   MOSTRA_TAREFA

   txtServico.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSERVICO_LostFocus"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub cmbVENDEDOR_Click()
On Error Resume Next

   cmbVendedorAUX.ListIndex = cmbVendedor.ListIndex

Err.Clear
End Sub

Private Sub cmbvendedor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtGarantia.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVendedor_KeyPress"
End Sub

Private Sub txtGARANTIA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtQTDE.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtGARANTIA_KeyPress"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDESCONTOPRODUTO.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
   txtQTDE.BackColor = &HFFFFFF
   If Trim(txtQTDE.Text) = "" Then _
      txtQTDE.Text = 1

   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
End Sub

Private Sub txtDescontoProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescontoProduto_KeyPress"
End Sub

Private Sub txtdescontoproduto_LostFocus()
   txtDESCONTOPRODUTO.BackColor = &HFFFFFF
   If txtDESCONTOPRODUTO.Text = "" Then _
      txtDESCONTOPRODUTO.Text = 0

   txtDESCONTOPRODUTO.Text = Format(txtDESCONTOPRODUTO.Text, strFormatacao2Digitos)
End Sub

Private Sub txtValorProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then

      If CHECA_DADOS_OS = True Then _
         GRAVA_PECA

      KeyAscii = 0
      txtProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorProduto_KeyPress"
End Sub

Private Sub txtvalorproduto_LostFocus()
   txtValorProduto.BackColor = &HFFFFFF
   If txtValorProduto.Text = "" Then _
      txtValorProduto.Text = 0

   txtValorProduto.Text = Format(txtValorProduto.Text, strFormatacao2Digitos)

   TOTALIZA_CAMPOS
End Sub

Sub LIMPA_OS()
'On Error GoTo ERRO_TRATA

   OSPECA_ID_N = 0
   VEICULO_ID_N = 0
   EQUIPAMENTO_ID_N = 0
   optEqp.Value = True
   optVeiculo.Value = False

   Toolbar1.Buttons(3).Enabled = True
   Toolbar1.Buttons(5).Enabled = True

   INDR_RECEITA = 0
   VENDEDOR_ID_N = 0
   PEDIDO_ID_N = 0
   OS_ID_N = 0
   CLIENTE_ID_N = 0
   txtCNPJCPF.PromptInclude = False

   Situação_Ordem_Serviço = ""
   txtTotDescontoServico.Text = Format(0, strFormatacao2Digitos)
   txtTotServico.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralServico.Text = Format(0, strFormatacao2Digitos)
   txtTotDescontoProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotOS.Text = Format(0, strFormatacao2Digitos)
   txtDesc.Text = ""

   lstServiço.ListItems.Clear
   lstProduto.ListItems.Clear

   EQUIPAMENTO_ID_N = 0
   txtOs.Text = ""
   txtDtOS.PromptInclude = False
   txtDtOS.Text = Now
   txtDtOS.PromptInclude = True
   cmbConsultor.Text = ""
   cmbConsultorAUX.Text = ""
   cmbTipoOS.Text = ""
   cmbTipoOSAUX.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   txtEqp.Text = ""
   txtPlaca.Text = ""
   txtCNPJCPF.Text = ""
   txtCliente.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   txtMarca.Text = ""
   txtTotOS.Text = ""
   txtGarantia.PromptInclude = False
   txtGarantia.Text = ""

   LIMPA_PRODUTO
   LIMPA_SERVIÇO
   CARREGA_COMBOS
   HABILITA_TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_OS"
End Sub

Sub LIMPA_OS_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(3).Enabled = True
   Toolbar1.Buttons(5).Enabled = True

   INDR_RECEITA = 0
   VENDEDOR_ID_N = 0
   CLIENTE_ID_N = 0
   txtCNPJCPF.PromptInclude = False

   Situação_Ordem_Serviço = ""
   txtTotDescontoServico.Text = Format(0, strFormatacao2Digitos)
   txtTotServico.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralServico.Text = Format(0, strFormatacao2Digitos)
   txtTotDescontoProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotOS.Text = Format(0, strFormatacao2Digitos)
   txtDesc.Text = ""

   lstServiço.ListItems.Clear
   lstProduto.ListItems.Clear
   OSPECA_ID_N = 0
   EQUIPAMENTO_ID_N = 0
   txtOs.Text = ""
   txtDtOS.PromptInclude = False
   txtDtOS.Text = Now
   txtDtOS.PromptInclude = True
   cmbConsultor.Text = ""
   cmbConsultorAUX.Text = ""
   cmbTipoOS.Text = ""
   cmbTipoOSAUX.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   txtEqp.Text = ""
   txtPlaca.Text = ""
   txtCNPJCPF.Text = ""
   txtCliente.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   txtMarca.Text = ""
   txtTotOS.Text = ""

   LIMPA_PRODUTO
   LIMPA_SERVIÇO
   CARREGA_COMBOS
   HABILITA_TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_OS_QUASE_TUDO"
End Sub

Sub LIMPA_SERVIÇO()
   txtDtInicio.PromptInclude = False
   txtDtInicio.Text = ""
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtServico.Text = ""
   txtDescTarefa.Text = ""
   cmbMecanicoAUX.Text = ""
   cmbMecanico.Text = ""
   txtDescontoTarefa.Text = ""
   txtValorTarefa.Text = ""
   txtTotalTarefa.Text = ""
   cmbSitServ.Text = ""
End Sub

Sub LIMPA_PRODUTO()
   OSPECA_ID_N = 0
   PRODUTO_ID_N = 0
   txtProduto.Text = ""
   txtDESCPRODUTO.Text = ""
   cmbVendedorAUX.Text = ""
   cmbVendedor.Text = ""
   txtQTDE.Text = ""
   txtDESCONTOPRODUTO.Text = ""
   txtValorProduto.Text = ""
   txtTOTALPRODUTO.Text = ""
   txtGarantia.PromptInclude = False
   txtGarantia.Text = ""
End Sub

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

'parametros combos x tabela descr
'8 = consultor tecnico
'9 = mecanico
   
   cmbSitServ.Clear

   cmbSitServ.AddItem "E-Execução"
   cmbSitServ.AddItem "P-Pendente"
   cmbSitServ.AddItem "O-Orçamento"
   cmbSitServ.AddItem "C-Cancelada"
   cmbSitServ.AddItem "F-Finalizada"

   cmbConsultorAUX.Clear
   cmbConsultor.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 8 or tipo = 5 "   'consultor tecnico
   SQL = SQL & " and status = 1"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbConsultorAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbConsultor.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbMecanicoAUX.Clear
   cmbMecanico.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 9 "   'mecanico/tecnico
   SQL = SQL & " and status = 1"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbMecanicoAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbMecanico.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbVendedorAUX.Clear
   cmbVendedor.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select vendedor_id, descricao from vwVendedor "
   SQL = SQL & " where status = 'A' "   'vendedor
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbVendedorAUX.AddItem TabDESCR.Fields("vendedor_id").Value
      cmbVendedor.AddItem Trim(TabDESCR.Fields("descricao").Value) & "-" & Trim(TabDESCR.Fields("vendedor_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbTipoOSAUX.Clear
   cmbTipoOS.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'H' "
   SQL = SQL & "order by codigo "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipoOS.AddItem Trim(TabDESCR!DESCRICAO)
      cmbTipoOSAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbSituacao.Clear

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'Z' "
   SQL = SQL & "order by codigo "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbSituacao.AddItem Trim(TabDESCR!DESCRICAO)
      cmbSituacaoAUX.AddItem Trim(TabDESCR.Fields("CODIGO").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   If INDR_RECEITA = 2 Then
      cmbSituacao.Clear
      cmbSituacao.AddItem "Fechamento"
      cmbSituacao.Text = "Fechamento"
      cmbSituacaoAUX.Text = 9
      cmbSituacaoAUX.AddItem 9
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      txtDtOS.Enabled = False
      cmbConsultor.Enabled = False
      cmbTipoOS.Enabled = False
      cmbSituacao.Enabled = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Sub MOSTRA_OS()
'On Error GoTo ERRO_TRATA

   LIMPA_OS_QUASE_TUDO

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtOs.Text = OS_ID_N
      txtPlaca.Text = "" & TabTemp.Fields("placa").Value
      VEICULO_ID_N = 0 & TabTemp.Fields("veiculo_ID").Value
      EQUIPAMENTO_ID_N = 0 & TabTemp.Fields("EQUIPAMENTO_ID").Value
      txtEqp.Text = "" & EQUIPAMENTO_ID_N
      PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value

      If VEICULO_ID_N > 0 Then
         optVeiculo.Value = True
         optEqp.Value = False
         txtPlaca.Visible = True
         txtEqp.Visible = False
         txtEqp.Text = ""
         lblEqp.Caption = "Placa: "
         Else
            optVeiculo.Value = False
            optEqp.Value = True
            txtEqp.Visible = True
            txtPlaca.Visible = False
            txtPlaca.Text = ""
            lblEqp.Caption = "Eqp: "
      End If

      txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("CNPJCPF").Value)
      txtCNPJCPF.PromptInclude = True

      txtCliente.Text = "" & Trim(TabTemp.Fields("nome_cliente").Value)
      If Not IsNull(TabTemp.Fields("cliente").Value) Then _
         txtCliente.Text = "" & Trim(TabTemp.Fields("cliente").Value)

      txtDtOS.PromptInclude = False
         txtDtOS.Text = "" & TabTemp.Fields("dt_os").Value
         If Len(txtDtOS.Text) = 8 Then _
            txtDtOS.Text = "" & DMA(TabTemp.Fields("dt_os").Value, "I")
      txtDtOS.PromptInclude = True

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      If Not IsNull(TabTemp.Fields("ct_id").Value) Then
         If IsNumeric(TabTemp.Fields("ct_id").Value) Then
            cmbConsultorAUX.Text = Trim(TabTemp.Fields("ct_id").Value)

            SQL = "select nome from USUARIO "
            SQL = SQL & " where usuario_id = " & TabTemp.Fields("ct_id").Value
            SQL = SQL & " and status = 1"
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then _
               cmbConsultor.Text = "" & Trim(TabDESCR.Fields(0).Value)
         End If
      End If

      txtDesc.Text = "" & Trim(TabTemp.Fields("nome_eqp").Value)

      cmbSituacao.Text = "" & TRAZ_DESCRITOR("Z", TabTemp.Fields("SITUACAO_os").Value)
      cmbSituacaoAUX.Text = "" & TabTemp.Fields("SITUACAO_os").Value

      Situação_Ordem_Serviço = TabTemp.Fields("situacao_os").Value
      INDR_RECEITA = 0 & TabTemp.Fields("situacao_os").Value

      cmbTipoOS.Text = "" & TRAZ_DESCRITOR("H", TabTemp.Fields("tipo_os").Value)
      cmbTipoOSAUX.Text = "" & Trim(TabTemp.Fields("tipo_os").Value)

      txtANO.Text = "" & Trim(TabTemp.Fields("ano").Value)
      txtMODELO.Text = "" & Trim(TabTemp.Fields("modelo").Value)

      If Not IsNull(TabTemp.Fields("marca_id").Value) Then _
         txtMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)

      txtTotOS.Text = ""
      DT_FECHAMENTO_OS = 0 & TabTemp.Fields("dt_fecha").Value
      Else
         INDR_PRI = True
         txtOs.Text = OS_ID_N
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SETA_GRID_SERVIÇO
   SETA_GRID_PRODUTO
   TOTALIZA_CAMPOS

   If Situação_Ordem_Serviço = "4" Then
      'DESAABILITA_TELA
      MsgBox "Ordem de Serviço CANCELADA, permitido somente consulta."
   End If
   If Situação_Ordem_Serviço = "2" Then
      'DESAABILITA_TELA
      MsgBox "Ordem de Serviço FECHADA, permitido somente consulta."
   End If

   If Trim(txtEqp.Text) <> "" Then _
      MOSTRA_EQP

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OS"
End Sub

Private Sub MOSTRA_EQP()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   VEICULO_ID_N = 0
   EQUIPAMENTO_ID_N = 0
   If Trim(txtEqp.Text) <> "" Then
      EQUIPAMENTO_ID_N = 0 & txtEqp.Text
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwEQUIPAMENTO "
      SQL = SQL & " where equipamento_id = " & EQUIPAMENTO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value

         txtCNPJCPF.PromptInclude = False
            If Trim(txtCNPJCPF.Text) = "" Then _
               txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         txtCliente.Text = "" & Trim(TabTemp.Fields("descpessoa").Value)
         txtDesc.Text = "" & Trim(TabTemp.Fields("descricao").Value)
         txtANO.Text = "" & TabTemp!Ano
         txtMODELO.Text = "" & TabTemp!MODELO

         If Not IsNull(TabTemp.Fields("marca_id").Value) Then _
            If IsNumeric(TabTemp.Fields("marca_id").Value) Then _
               txtMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)
         Else
            MsgBox "Não encontrado, verifique."
            txtEqp.SetFocus
            txtEqp.Text = ""
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_EQP"
End Sub

Sub GRAVA_OS()
'On Error GoTo ERRO_TRATA

   Dim KM_N As Long

   OS_ID_N = 0 & txtOs.Text
   KM_N = 0 & txtKM.Text

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from OS "
   SQL = SQL & " where os_id = " & OS_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      SQL = "update OS set "
         SQL = SQL & " TIPO_OS = " & cmbTipoOSAUX.Text                       'TIPO_OS
         SQL = SQL & ", SITUACAO_OS = " & cmbSituacaoAUX.Text                 'SITUACAO_OS
         SQL = SQL & ", CT_ID = " & cmbConsultorAUX.Text                      'CT_ID
         SQL = SQL & ", CLIENTE = '" & Trim(Left(txtCliente.Text, 50)) & "'"  'CLIENTE
         SQL = SQL & ", pessoa_id = " & PESSOA_ID_N                           'pessoa_id
         SQL = SQL & ", KM = " & KM_N                           'KM
         If Trim(cmbSituacaoAUX.Text) = "2" Then _
            SQL = SQL & ", dt_fecha = '" & Now & "'"                          'dt_fecha
      SQL = SQL & " where os_id = " & OS_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      Else
         SQL = "insert into OS "
            SQL = SQL & "(OS_ID,ESTABELECIMENTO_ID,DT_OS,TIPO_OS,"
            SQL = SQL & " SITUACAO_OS,CT_ID,CLIENTE,PESSOA_ID,KM)"
         SQL = SQL & " values ( "
            SQL = SQL & OS_ID_N                                   'OS_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                   'EMPRESA_ID
            SQL = SQL & ",'" & Trim(txtDtOS.Text) & "'"               'DT_OS
            SQL = SQL & "," & cmbTipoOSAUX.Text                      'TIPO_OS
            SQL = SQL & ",0" & cmbSituacaoAUX.Text                   'SITUACAO_OS
            SQL = SQL & "," & cmbConsultorAUX.Text                   'CT_ID
            SQL = SQL & ",'" & Trim(Left(txtCliente.Text, 50)) & "'" 'CLIENTE
            SQL = SQL & "," & PESSOA_ID_N                            'PESSOA_ID
            SQL = SQL & "," & KM_N                           'KM
         SQL = SQL & "  )"
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL

'--------OSVEICEQP
   GRAVA_OSVEICEQP

   'fechada
   If cmbSituacaoAUX.Text = "2" Then _
      GERA_FATURA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_OS"
End Sub

Sub GRAVA_SERVIÇO()
'On Error GoTo ERRO_TRATA

   Dim DtFecha_D  As Date
   Dim DtOS_D     As Date
   Dim DATA_INI_A As String
   Dim DATA_FIM_A As String

   If Trim(cmbSitServ.Text) = "" Then
      MsgBox "Informar etapa da tarefa."
      cmbSitServ.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtServico.Text) Then
      MsgBox "Serviço inválido."
      txtServico.SetFocus
      Exit Sub
   End If
   If Trim(txtDescTarefa.Text) = "" Then
      MsgBox "Descrição Serviço inválida."
      txtDescTarefa.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(cmbMecanicoAUX.Text) Then
      cmbMecanico.ListIndex = 1
      'cmbMecanicoAUX.ListIndex = cmbMecanico.ListIndex
      'MsgBox "Mecanico não informado."
      'cmbMecanico.SetFocus
      'Exit Sub
   End If
   If Not IsNumeric(txtDescontoTarefa.Text) Then _
      txtDescontoTarefa.Text = 0

   If Trim(txtValorTarefa.Text) = "" Then
      MsgBox "Valor serviço inválido."
      txtValorTarefa.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtValorTarefa.Text) Then
      MsgBox "Valor serviço inválido."
      txtValorTarefa.SetFocus
      Exit Sub
   End If

   DtOS_D = txtDtOS.Text

   DtFecha_D = 0

   txtDtFim.PromptInclude = True
   If IsDate(txtDtFim.Text) Then
      DtFecha_D = txtDtFim.Text
'MsgBox CDate(DtFecha_D) & "      " & CDate(DtOS_D)
      If CDate(DtFecha_D) < CDate(DtOS_D) Then
         'MsgBox "Data de fechamento do serviço menor que data da Ordem de Serviço, não permitido."
         'Exit Sub
      End If
   End If
   
   DATA_INI_A = ""
   If IsDate(txtDtInicio.Text) Then _
      DATA_INI_A = txtDtInicio.Text
   DATA_FIM_A = ""
   If IsDate(txtDtFim.Text) Then _
      DATA_FIM_A = txtDtFim.Text

   GRAVA_OS

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from OSServico "
   SQL = SQL & " where os_id = " & OS_ID_N
   SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      SQL = "update OSSERVICO set "
         SQL = SQL & " RESPONSAVEL_ID = " & cmbMecanicoAUX.Text                  'RESPONSAVEL_ID
         SQL = SQL & ", VALOR_SERVICO = " & tpMOEDA(txtValorTarefa.Text)         'VALOR_SERVICO
         SQL = SQL & ", DESCRICAO = '" & Trim(txtDescTarefa.Text) & "'"          'DESCRICAO
         SQL = SQL & ", DESCONTO_SERVICO = " & tpMOEDA(txtDescontoTarefa.Text)   'DESCONTO_SERVICO
         
         If Trim(DATA_FIM_A) = "" Then
            SQL = SQL & ", DT_fim = null"                                        'DT_fim
            Else: SQL = SQL & ", DT_fim = '" & Trim(DATA_FIM_A) & "'"            'DT_fim
         End If

         If Trim(DATA_INI_A) = "" Then
            SQL = SQL & ", DT_inicio = null"                                        'DT_inicio
            Else: SQL = SQL & ", DT_inicio = '" & Trim(DATA_INI_A) & "'"            'DT_inicio
         End If

         SQL = SQL & ", situacao = '" & Left(cmbSitServ.Text, 1) & "'"           'SitServ
      SQL = SQL & " where os_id = " & OS_ID_N
      SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text                       'OSSERVICO_ID
      Else
         SQL = "insert into OSSERVICO "
            SQL = SQL & "(OSSERVICO_ID,OS_ID,OSTAREFA_ID,DT_CAD,RESPONSAVEL_ID,"
            SQL = SQL & "VALOR_SERVICO,DESCRICAO,desconto_servico,DT_fim,situacao,dt_inicio) "
         SQL = SQL & " values ( "
            SQL = SQL & MAX_ID("osservico_id", "OSSERVICO", "OS_ID", txtOs.Text, "", "")  'OSSERVICO_ID
            SQL = SQL & "," & OS_ID_N                                                 'OS_ID
            SQL = SQL & "," & txtServico.Text                                             'OSTAREFA_ID
            SQL = SQL & ",'" & Now & "'"                                            'DT_CAD
            SQL = SQL & "," & cmbMecanicoAUX.Text                                         'RESPONSAVEL_ID
            SQL = SQL & "," & tpMOEDA(txtValorTarefa.Text)                                'VALOR_SERVICO
            SQL = SQL & ",'" & Trim(txtDescTarefa.Text) & "'"                             'DESCRICAO
            SQL = SQL & "," & tpMOEDA(txtDescontoTarefa.Text)                             'DESCONTO_SERVICO

            If Trim(DATA_FIM_A) = "" Then
               SQL = SQL & ",null"                                        'DT_fim
               Else: SQL = SQL & ",'" & Trim(DATA_FIM_A) & "'"            'DT_fim
            End If

            SQL = SQL & ",'" & Left(cmbSitServ.Text, 1) & "'"                             'SitServ

            If Trim(DATA_INI_A) = "" Then
               SQL = SQL & ",null"                                        'DT_ini
               Else: SQL = SQL & ",'" & Trim(DATA_INI_A) & "'"            'DT_ini
            End If

         SQL = SQL & "  )"
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID_SERVIÇO
   LIMPA_SERVIÇO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_SERVIÇO"
End Sub

Sub GRAVA_PECA()
'On Error GoTo ERRO_TRATA

   If PRODUTO_ID_N <= 0 Then
      MsgBox "produto inválido."
      txtProduto.SetFocus
      Exit Sub
   End If
   If Trim(txtDESCPRODUTO.Text) = "" Then
      MsgBox "Descrição produto inválida."
      txtDESCPRODUTO.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(cmbVendedorAUX.Text) Then
      MsgBox "Vendedor não informado."
      cmbVendedor.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtDESCONTOPRODUTO.Text) Then _
      txtDESCONTOPRODUTO.Text = 0

   If Trim(txtQTDE.Text) = "" Then
      MsgBox "Quantidade informada inválida."
      txtQTDE.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtQTDE.Text) Then
      MsgBox "Quantidade informada inválida."
      txtQTDE.SetFocus
      Exit Sub
   End If
   If Trim(txtValorProduto.Text) = "" Then
      MsgBox "Valor produto inválido."
      txtValorProduto.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtValorProduto.Text) Then
      MsgBox "Valor produto inválido."
      txtValorProduto.SetFocus
      Exit Sub
   End If

   GRAVA_OS

   If OSPECA_ID_N <= 0 Then
      OSPECA_ID_N = 0 & MAX_ID("OSPECA_id", "OSPECA", "", "", "", "")
   End If

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from OSPECA "
   SQL = SQL & " where os_id = " & OS_ID_N
   SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N
   SQL = SQL & " and OSPECA_ID = " & OSPECA_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      SQL = "update OSPECA set "
         SQL = SQL & " DT_CAD = '" & DMA(Date) & "'"                             'DT_CAD
         SQL = SQL & ", SOLICITANTE_ID = " & cmbVendedorAUX.Text                 'SOLICITANTE_ID
         SQL = SQL & ", VALOR_ITEM = " & tpMOEDA(txtValorProduto.Text)           'VALOR_ITEM
         SQL = SQL & ", DESCONTO_PRODUTO = " & tpMOEDA(txtDESCONTOPRODUTO.Text)  'DESCONTO_PRODUTO
         SQL = SQL & ", QTDE = " & tpMOEDA(txtQTDE.Text)                         'QTDE
         SQL = SQL & ", DT_GARANTIA = '" & DMA(txtGarantia.Text) & "'"           'DT_GARANTIA
      SQL = SQL & " where os_id = " & OS_ID_N
      SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N
      SQL = SQL & " and OSPECA_ID = " & OSPECA_ID_N
      Else
         SQL = "insert into OSPECA "
            SQL = SQL & "(OSPECA_ID,OS_ID,PRODUTO_ID,DT_CAD,SOLICITANTE_ID,VALOR_ITEM,DESCONTO_PRODUTO,QTDE,DT_GARANTIA) "
         SQL = SQL & " values ( "
            SQL = SQL & OSPECA_ID_N                                              'OSPECA_ID
            SQL = SQL & "," & OS_ID_N                                         'OS_ID
            SQL = SQL & "," & PRODUTO_ID_N                                    'PRODUTO_ID
            SQL = SQL & ",'" & DMA(Date) & "'"                                'DT_CAD
            SQL = SQL & "," & cmbVendedorAUX.Text                             'SOLICITANTE_ID
            SQL = SQL & "," & tpMOEDA(txtValorProduto.Text)                   'VALOR_ITEM
            SQL = SQL & "," & tpMOEDA(txtDESCONTOPRODUTO.Text)                 'DESCONTO_PRODUTO
            SQL = SQL & "," & tpMOEDA(txtQTDE.Text)                           'QTDE
            SQL = SQL & ",'" & DMA(txtGarantia.Text) & "'"                    'DT_GARANTIA
         SQL = SQL & "  )"
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL
   OSPECA_ID_N = 0

   SETA_GRID_PRODUTO
   LIMPA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PECA"
End Sub

Sub SETA_GRID_SERVIÇO()
'On Error GoTo ERRO_TRATA

   lstServiço.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSSERVICO "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstServiço.ListItems.Add(, "seq." & TabTemp.Fields("OSSERVICO_ID").Value, TabTemp.Fields("OSSERVICO_ID").Value)

      item.SubItems(1) = "" & Trim(TabTemp!DESCRICAO)
      item.SubItems(2) = "" & Format(TabTemp!VALOR_SERVICO, strFormatacao2Digitos)
      item.SubItems(3) = "" & Format(TabTemp!DESCONTO_SERVICO, strFormatacao2Digitos)
      item.SubItems(4) = "" & Format(TabTemp!VALOR_SERVICO - TabTemp!DESCONTO_SERVICO, strFormatacao2Digitos)
      item.SubItems(6) = "" & Trim(TabTemp.Fields("DT_INICIO").Value)
      item.SubItems(7) = "" & Trim(TabTemp.Fields("DT_fim").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome from USUARIO "
      SQL = SQL & " where USUARIO_ID = " & TabTemp.Fields("RESPONSAVEL_ID").Value
      SQL = SQL & " and status = 1"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         item.SubItems(5) = "" & Trim(TabConsulta.Fields("nome").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_SERVIÇO"
End Sub

Sub SETA_GRID_PRODUTO()
'On Error GoTo ERRO_TRATA

   lstProduto.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select OSPECA.OSPECA_ID, OSPECA.OS_ID, OSPECA.PRODUTO_ID, "
   SQL = SQL & " OSPECA.DT_CAD, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, "
   SQL = SQL & " OSPECA.DESCONTO_PRODUTO, OSPECA.DT_GARANTIA, "
   SQL = SQL & " OSPECA.QTDE, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
   SQL = SQL & " from OSPECA "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstProduto.ListItems.Add(, "seq." & TabTemp.Fields("OSPECA_ID").Value, _
                                                     TabTemp.Fields("CODG_PRODUTO").Value)

      item.SubItems(1) = "" & Trim(TabTemp!DESCRICAO)
      item.SubItems(2) = "" & Format(TabTemp.Fields("QTDE").Value, strFormatacao2Digitos)
      item.SubItems(3) = "" & Format(TabTemp!Valor_Item, strFormatacao2Digitos)
      item.SubItems(4) = "" & Format(TabTemp!DESCONTO_PRODUTO, strFormatacao2Digitos)
      item.SubItems(5) = "" & Format((TabTemp!Valor_Item - TabTemp!DESCONTO_PRODUTO) _
                                    * TabTemp.Fields("QTDE").Value, strFormatacao3Digitos)

      item.SubItems(6) = "" & TabTemp.Fields("OSPECA_ID").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select descricao from vwVendedor"
      SQL = SQL & " where VENDEDOR_ID = " & TabTemp.Fields("solicitante_ID").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         item.SubItems(7) = "" & Trim(TabConsulta.Fields("descricao").Value)
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      item.SubItems(8) = "" & TabTemp.Fields("dt_garantia").Value
      item.SubItems(9) = "" & TabTemp.Fields("PRODUTO_ID").Value

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PRODUTO"
End Sub

Private Sub MOSTRA_TAREFA()
'On Error GoTo ERRO_TRATA

   If IsNumeric(txtServico.Text) Then
      txtDescTarefa.Text = ""
      cmbMecanico.Text = ""
      cmbMecanicoAUX.Text = ""
      txtDescontoTarefa.Text = ""
      txtValorTarefa.Text = ""
      txtTotalTarefa.Text = ""

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from OSServico "
      SQL = SQL & " where os_id = " & OS_ID_N
      SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         If Not IsNull(TabConsulta.Fields("dt_inicio").Value) Then
            If IsDate(TabConsulta.Fields("dt_inicio").Value) Then
               txtDtInicio.PromptInclude = False
                  txtDtInicio.Text = TabConsulta.Fields("dt_inicio").Value
               txtDtInicio.PromptInclude = True
            End If
         End If
         If Not IsNull(TabConsulta.Fields("dt_fim").Value) Then
            If IsDate(TabConsulta.Fields("dt_fim").Value) Then
               txtDtFim.PromptInclude = False
                  txtDtFim.Text = TabConsulta.Fields("dt_fim").Value
               txtDtFim.PromptInclude = True
            End If
         End If

         TAREFA_ID_N = 0 & TabConsulta.Fields("ostarefa_id").Value

         txtDescTarefa.Text = "" & Trim(TabConsulta.Fields("descricao").Value)

         VALOR_ITEM_N = 0 & TabConsulta.Fields("VALOR_servico").Value
         txtValorTarefa.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         VALOR_DESCONTO_N = 0 & TabConsulta.Fields("desconto_servico").Value
         txtDescontoTarefa.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

         txtTotalTarefa.Text = "" & Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

         cmbMecanicoAUX.Text = "" & TabConsulta.Fields("responsavel_id").Value

         If Not IsNull(TabConsulta.Fields("situacao").Value) Then
            If Trim(TabConsulta.Fields("situacao").Value) <> "" Then
               If Trim(TabConsulta.Fields("situacao").Value) = "E" Then _
                  cmbSitServ.Text = "E-Execução"
               If Trim(TabConsulta.Fields("situacao").Value) = "P" Then _
                  cmbSitServ.Text = "P-Pendente"
               If Trim(TabConsulta.Fields("situacao").Value) = "O" Then _
                  cmbSitServ.Text = "O-Orçamento"
               If Trim(TabConsulta.Fields("situacao").Value) = "C" Then _
                  cmbSitServ.Text = "C-Cancelada"
               If Trim(TabConsulta.Fields("situacao").Value) = "F" Then _
                  cmbSitServ.Text = "F-Finalizada"
            End If
         End If

         If TabTemp.State = 1 Then _
            TabTemp.Close
   
         SQL = "select nome from USUARIO "
         SQL = SQL & " where USUARIO_ID = " & TabConsulta.Fields("responsavel_id").Value
         SQL = SQL & " and status = 1"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            cmbMecanico.Text = "" & Trim(TabTemp.Fields("nome").Value)

         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'se não achou na tabela serviço busca na tabela ostarefa que é a de cadastro de serviço
      If Trim(txtDescTarefa.Text) = "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from OSTAREFA "
         SQL = SQL & " where OSTAREFA_ID = " & txtServico.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            txtDescTarefa.Text = "" & Trim(TabTemp.Fields("DESCRICAO").Value)
            txtValorTarefa.Text = "" & Format(TabTemp.Fields("VALOR").Value, strFormatacao2Digitos)
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TAREFA"
End Sub

Private Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PRODUTO_ID_N = 0

      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDESCPRODUTO.Text = "" & Trim(TabConsulta.Fields("DESCRICAO").Value)
         txtValorProduto.Text = "" & Format(TabConsulta.Fields("PRECO_VENDA").Value, strFormatacao2Digitos)
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
         Else
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            MsgBox "Produto não cadastrado."
            txtProduto.SetFocus
            Exit Sub
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from vwOSPECA "
      SQL = SQL & " where os_id = " & OS_ID_N
      SQL = SQL & " and OSPECA_ID = " & OSPECA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         PRODUTO_ID_N = 0 & TabConsulta.Fields("PRODUTO_id").Value

         txtDESCPRODUTO.Text = "" & Trim(TabConsulta.Fields("descricaoproduto").Value)

         VALOR_ITEM_N = 0 & TabConsulta.Fields("VALOR_ITEM").Value
         txtValorProduto.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         VALOR_DESCONTO_N = 0 & TabConsulta.Fields("desconto_produto").Value
         txtDESCONTOPRODUTO.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

         QTDE_PEDIDO = 0 & TabConsulta.Fields("qtde").Value
         txtQTDE.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)

         txtTOTALPRODUTO.Text = "" & Format((VALOR_ITEM_N - VALOR_DESCONTO_N) * QTDE_PEDIDO, strFormatacao3Digitos)

         cmbVendedorAUX.Text = "" & TabConsulta.Fields("solicitante_id").Value

         If TabTemp.State = 1 Then _
            TabTemp.Close
   
         SQL = "select nome from USUARIO "
         SQL = SQL & " where USUARIO_ID = " & TabConsulta.Fields("solicitante_id").Value
         SQL = SQL & " and status = 1"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            cmbVendedor.Text = "" & Trim(TabTemp.Fields("nome").Value)

         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If
   cmbVendedor.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub TOTALIZA_CAMPOS()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral desconto serviço
   SQL = "select sum(desconto_servico) from OSSERVICO "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotDescontoServico.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral serviço
   SQL = "select sum(valor_servico) from OSSERVICO "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotServico.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   VALOR_ITEM_N = 0 & txtTotServico.Text
   VALOR_DESCONTO_N = 0 & txtTotDescontoServico.Text
   txtTotGeralServico.Text = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

'total geral desconto produto
   SQL = "select sum(desconto_produto) from OSPECA "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotDescontoProduto.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral produto
   SQL = "select sum(valor_item*QTDE) from OSPECA "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotProduto.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'PEÇA
   VALOR_ITEM_N = 0 & txtTotProduto.Text
   VALOR_DESCONTO_N = 0 & txtTotDescontoProduto.Text
   txtTotGeralProduto.Text = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

'SERVIÇO
   VALOR_ITEM_N = 0 & txtTotServico.Text
   VALOR_DESCONTO_N = 0 & txtTotDescontoServico.Text

'TOTAL
   VALOR_ITEM_N = 0 & txtTotGeralServico.Text
   VALOR_DESCONTO_N = 0 & txtTotGeralProduto.Text
   txtTotOS.Text = "" & Format(VALOR_ITEM_N + VALOR_DESCONTO_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOTALIZA_CAMPOS"
End Sub

Sub GERA_FATURA()
'On Error GoTo ERRO_TRATA

   Dim Vlr_OS_n As Double

   Vlr_OS_n = 0 & txtTotOS.Text

   If Vlr_OS_n <= 0 Then _
      Exit Sub

   Dim TabOS               As New ADODB.Recordset
   Dim TabPedido           As New ADODB.Recordset
   Dim TIPO_REGISTRO_A     As String
   Dim STATUS_N            As Integer
   Dim DESCONTO_PEÇA_N     As Double
   Dim DESCONTO_SERVIÇO_N  As Double

   If USUARIO_ID_N <= 0 Then _
      USUARIO_ID_N = 144

   VENDEDOR_ID_N = 0
   CLIENTE_ID_N = 0
   txtCNPJCPF.PromptInclude = False
   STATUS_N = 2
   TIPO_REGISTRO_A = "OS"
   VALOR_TOTAL_DESCONTO_N = 0 & DESCONTO_PEÇA_N + DESCONTO_SERVIÇO_N
'===========================
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select vendedor_id from vwVendedor "
   SQL = SQL & " where descricao = 'BALCAO' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      VENDEDOR_ID_N = 0 & TabConsulta.Fields(0).Value
      Else
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select vendedor_id from vwVendedor "
         SQL = SQL & " where descricao = 'BALCÃO' "
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            VENDEDOR_ID_N = 0 & TabConsulta.Fields(0).Value
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select cliente_id from CLIENTE "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      CLIENTE_ID_N = 0 & TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
'===========================
   If TabOS.State = 1 Then _
      TabOS.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & OS_ID_N
   TabOS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabOS.EOF Then
      If TabPedido.State = 1 Then _
         TabPedido.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where pedido_id = " & OS_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         If TabPedido.State = 1 Then _
            TabPedido.Close
         If TabOS.State = 1 Then _
            TabOS.Close

         MsgBox "Pedido já existente, verificar."
         Exit Sub
         Else
            SQL = "INSERT INTO PEDIDO "
               SQL = SQL & "("
                  SQL = SQL & " PEDIDO_ID,Empresa_id, CGCCPF, Vendedor_id, Dt_Req, "
                  SQL = SQL & " Nome_Cliente, Status, Tipo_Registro,USUARIO_ID, "
                  SQL = SQL & " CLIENTE_ID, Valor_ToTal, valor_desconto,perc_desc,estabelecimento_id"
               SQL = SQL & " )"
            SQL = SQL & " VALUES ("
               SQL = SQL & OS_ID_N
               SQL = SQL & "," & EMPRESA_ID_N
               SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
               SQL = SQL & "," & VENDEDOR_ID_N & ","
               SQL = SQL & "'" & Now & "'"
               SQL = SQL & ",'" & Trim(txtCliente.Text) & "'"
               SQL = SQL & "," & STATUS_N
               SQL = SQL & ",'" & TIPO_REGISTRO_A & "'"
               SQL = SQL & "," & USUARIO_ID_N
               'SQL = SQL & "," & 9999
               SQL = SQL & "," & CLIENTE_ID_N
               SQL = SQL & "," & tpMOEDA(txtTotOS.Text)
               SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
               SQL = SQL & "," & tpMOEDA(0)
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close
'======================================
      'PRODUTOS ORDEM DE SERVIÇO
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select OSPECA.*,PRODUTO.CODG_PRODUTO,PRODUTO.DESCRICAO,PRODUTO.PRECO_CUSTO,PRODUTO.SITUACAO_TRIBUTARIA"
      SQL = SQL & " from OS "
      SQL = SQL & " INNER JOIN OSPECA "
      SQL = SQL & " ON OS.OS_ID = OSPECA.OS_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where OSPECA.os_id = " & OS_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         SEQ_ID_N = MAX_ID("seq_id", "PEDIDOITEM", "", "", "", "")
         QTDE_PEDIDO = 0 & TabTemp.Fields("qtde").Value

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("os_id").Value
         SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabConsulta.EOF Then
            SQL = "INSERT INTO PEDIDOITEM "
            SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Qtd_Pedida, "
            SQL = SQL & " Valor_item, valor_desconto, status,preco_custo,TIPO_REG) "
            SQL = SQL & " VALUES ("
               SQL = SQL & TabTemp.Fields("os_id").Value                            'PEDIDO_id
               SQL = SQL & "," & SEQ_ID_N                                           'SEQ_ID
               SQL = SQL & "," & TabTemp.Fields("produto_id").Value                 'produto_id
               SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                               'Qtd_Pedida
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_item").Value)        'Valor_item
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("desconto_produto").Value)  'Valor_desconto
               SQL = SQL & ", 'P'"                                                  'status
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("preco_custo").Value)       'PRECO_CUSTO
               SQL = SQL & ", 'PC'"                                                 'TIPO_REG
               'SQL = SQL & ",'" & Trim(TabTemp.Fields("SITUACAO_TRIBUTARIA").Value) & "'" 'stributaria
            SQL = SQL & ")"
            Else
               SQL = "UPDATE PEDIDOITEM SET "
                  SQL = SQL & " Qtd_Pedida = " & tpMOEDA(QTDE_PEDIDO)                                    'Qtd_Pedida
                  SQL = SQL & ", Valor_item = " & tpMOEDA(TabTemp.Fields("valor_item").Value)            'Valor_item
                  SQL = SQL & ", Valor_desconto = " & tpMOEDA(TabTemp.Fields("desconto_produto").Value)  'Valor_desconto
                  SQL = SQL & ", status = 'P'"                                                           'status
                  SQL = SQL & ", PRECO_CUSTO = " & tpMOEDA(TabTemp.Fields("preco_custo").Value)          'PRECO_CUSTO
                  SQL = SQL & ", TIPO_REG = 'PC'"                                                        'TIPO_REG
                  SQL = SQL & ", stributaria = '" & Trim(TabTemp.Fields("SITUACAO_TRIBUTARIA").Value)    'stributaria
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("os_id").Value
               SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
               SQL = SQL & " and seq_id = " & SEQ_ID_N
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         CONECTA_RETAGUARDA.Execute SQL

         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close
'======================================
      'SERVIÇO ORDEM DE SERVIÇO

      PRODUTO_ID_N = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select produto_id from PRODUTO "
      SQL = SQL & " where descricao = 'SERVICO'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         PRODUTO_ID_N = TabTemp.Fields(0).Value
         Else
            MsgBox "No cadastro de produto deve conter registro 'SERVICO' "
            Exit Sub
      End If

      SQL = "delete from PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabOS.Fields("os_id").Value
      SQL = SQL & " and tipo_reg = 'OS'"
      CONECTA_RETAGUARDA.Execute SQL

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select OSSERVICO.* from OS "
      SQL = SQL & " INNER JOIN OSSERVICO "
      SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID"
      SQL = SQL & " where OSSERVICO.os_id = " & OS_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         SEQ_ID_N = MAX_ID("seq_id", "PEDIDOITEM", "", "", "", "")
         QTDE_PEDIDO = 1

         SQL = "INSERT INTO PEDIDOITEM "
         SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Qtd_Pedida, "
         SQL = SQL & " Valor_item, valor_desconto, status,preco_custo,TIPO_REG) "
         SQL = SQL & " VALUES ("
            SQL = SQL & TabTemp.Fields("os_id").Value                            'PEDIDO_id
            SQL = SQL & "," & SEQ_ID_N                                           'SEQ_ID
            SQL = SQL & "," & PRODUTO_ID_N                                       'produto_id
            SQL = SQL & "," & tpMOEDA(1)                                         'Qtd_Pedida
            SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_servico").Value)     'valor_servico
            SQL = SQL & "," & tpMOEDA(TabTemp.Fields("desconto_servico").Value)  'Valor_desconto
            SQL = SQL & ", 'P'"                                                  'status
            SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_servico").Value)     'PRECO_CUSTO
            SQL = SQL & ", 'SV'"                                                 'TIPO_REG
         SQL = SQL & ")"

         CONECTA_RETAGUARDA.Execute SQL

         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close
'======================================
   End If   'If Not TabOS.EOF Then
   If TabOS.State = 1 Then _
      TabOS.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FATURA"
End Sub

Sub HABILITA_TELA()
   If INDR_RECEITA = 2 Then
      cmbSituacao.Clear
      cmbSituacao.AddItem "Fechamento"
      cmbSituacao.Text = "Fechamento"
      cmbSituacaoAUX.AddItem 9
      cmbSituacaoAUX.Text = 9
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      txtDtOS.Enabled = False
      cmbConsultor.Enabled = False
      cmbTipoOS.Enabled = False
      cmbSituacao.Enabled = False
      Else
         Frame2.Enabled = True
         Frame3.Enabled = True
         Frame4.Enabled = True
         txtDtOS.Enabled = True
         cmbConsultor.Enabled = True
         cmbTipoOS.Enabled = True
         cmbSituacao.Enabled = True
   End If
End Sub

Sub DESAABILITA_TELA()
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   txtDtOS.Enabled = False
   cmbConsultor.Enabled = False
   cmbTipoOS.Enabled = False
   cmbSituacao.Enabled = False
   Toolbar1.Buttons(3).Enabled = False
   'Toolbar1.Buttons(5).Enabled = False
End Sub

Public Function CHECA_DADOS_OS() As Boolean
'On Error GoTo ERRO_TRATA

   CHECA_DADOS_OS = False

   If Trim(txtOs.Text) = "" Then
      MsgBox "Número de Ordem de Serviço inválida."
      txtOs.SetFocus
      Exit Function
   End If
   If Not IsNumeric(txtOs.Text) Then
      MsgBox "Número de Ordem de Serviço inválida."
      txtOs.SetFocus
      Exit Function
   End If
   If Not IsDate(txtDtOS.Text) Then
      MsgBox "Data de Ordem de Serviço inválida."
      txtDtOS.SetFocus
      Exit Function
   End If
   If Trim(cmbConsultorAUX.Text) = "" Then
      MsgBox "Consultor inválido."
      cmbConsultor.SetFocus
      Exit Function
   End If
   If Not IsNumeric(cmbConsultorAUX.Text) Then
      MsgBox "Consultor inválido."
      cmbConsultor.SetFocus
      Exit Function
   End If
   If Trim(cmbTipoOSAUX.Text) = "" Then
      MsgBox "Tipo de Ordem de Serviço inválido."
      cmbTipoOS.SetFocus
      Exit Function
   End If
   If Not IsNumeric(cmbTipoOSAUX.Text) Then
      MsgBox "Tipo de Ordem de Serviço inválido."
      cmbTipoOS.SetFocus
      Exit Function
   End If
   If Trim(cmbSituacao.Text) = "" Then
      MsgBox "Situação da Ordem de Serviço inválida."
      cmbSituacao.SetFocus
      Exit Function
   End If
   
   If EQP_VEICULO = False Then
      If Trim(txtPlaca.Text) = "" Then
         MsgBox "Veículo não informado, Ordem de Serviço inválida."
         txtEqp.SetFocus
         Exit Function
      End If
      Else
         If Trim(txtEqp.Text) = "" Then
            MsgBox "Equipamento não informado, Ordem de Serviço inválida."
            txtEqp.SetFocus
            Exit Function
         End If
   End If
   
   CHECA_DADOS_OS = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_DADOS_OS"
End Function

Sub EXCLUIR_SERVIÇO_ITEM()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma Exclusão desse serviço ?"
   Style = vbYesNo + 32
   Title = "Atenção."
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then

      SQL = "Delete from OSSERVICO "
      SQL = SQL & " Where OSSERVICO_id = " & lstServiço.SelectedItem.Text
      SQL = SQL & " and os_id = " & txtOs.Text
      CONECTA_RETAGUARDA.Execute SQL

      SETA_GRID_SERVIÇO
      TOTALIZA_CAMPOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_SERVIÇO_ITEM"
End Sub

Sub EXCLUIR_PRODUTO_ITEM()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma Exclusão desse produto ?"
   Style = vbYesNo + 32
   Title = "Atenção."
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then

      SQL = "Delete from ospeca "
      SQL = SQL & " Where ospeca_id = " & Trim(lstProduto.SelectedItem.ListSubItems(6).Text)
      SQL = SQL & " and os_id = " & txtOs.Text
      CONECTA_RETAGUARDA.Execute SQL

      SETA_GRID_PRODUTO
      TOTALIZA_CAMPOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_PRODUTO_ITEM"
End Sub

Sub EXCLUIR_OS()
'On Error GoTo ERRO_TRATA

   If Trim(txtOs) <> "" Then
      If IsNumeric(txtOs) Then
         Msg = "Confirma Exclusão desta ordem de serviço ?"
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then

            SQL = "Delete from OSPECA "
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OSSERVICO "
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OSrelitem"
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OSrel"
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from ostermo"
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from osOBS"
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OSveiceqp"
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OS "
            SQL = SQL & " Where os_id = " & txtOs.Text
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_OS
            txtOs.SetFocus
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_OS"
End Sub

Private Sub MOSTRA_VEICULO()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   VEICULO_ID_N = 0
   EQUIPAMENTO_ID_N = 0
   If Trim(txtPlaca.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwVeiculo WITH (NOLOCK) "
      SQL = SQL & " where placa = '" & Trim(txtPlaca.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         VEICULO_ID_N = TabTemp.Fields("veiculo_id").Value
         PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value

         txtCliente.Text = "" & Trim(TabTemp.Fields("DescPessoa").Value)

         txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         txtDesc.Text = "" & TabTemp!DESCRICAO
         txtANO.Text = "" & TabTemp!Ano
         txtMODELO.Text = "" & TabTemp!MODELO

         If Not IsNull(TabTemp.Fields("marca_id").Value) Then _
            If IsNumeric(TabTemp.Fields("marca_id").Value) Then _
               txtMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VEICULO"
End Sub

Sub IMPRIMIR_ORDEM_SERVICO(NUMR_OS_N As Long, TIPO_REL As String, NOME_CLI As String)
'On Error GoTo ERRO_TRATA

   If NUMR_OS_N <= 0 Then _
      Exit Sub

   Dim NOME_CT_A        As String
   Dim CGC_A            As String
   Dim RAZAO_SOCIAL_A   As String
   Dim NOME_FANT_A      As String
   Dim ENDERECO_EMP_A   As String
   Dim CEP_EMP_A        As String
   Dim COMP_EMP_A       As String
   Dim NUMERO_EMP_A     As String
   Dim BAIRRO_EMP_A     As String
   Dim CIDADE_EMP_A     As String
   Dim UF_EMP_A         As String
   Dim FONE_EMP_A       As String
   Dim FONE_CLIENTE_A   As String
   Dim DT_FECHA_A       As String
   Dim RESPONSAVEL_A    As String
   Dim DT_GARANTIA_D    As String
   Dim COR_ID_N
   Dim MARCA_ID_N
   Dim TIPO_EQP_ID_N

   CGC_A = ""
   RAZAO_SOCIAL_A = ""
   NOME_FANT_A = ""
   NOME_CT_A = ""
   ENDERECO_EMP_A = ""
   CEP_EMP_A = ""
   COMP_EMP_A = ""
   NUMERO_EMP_A = ""
   BAIRRO_EMP_A = ""
   CIDADE_EMP_A = ""
   UF_EMP_A = ""
   FONE_EMP_A = ""
   FONE_CLIENTE_A = ""
   DT_FECHA_A = ""

   SQL = "delete from OSRELITEM where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from OSREL where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & NUMR_OS_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'CONSULTOR TECNICO
      SQL = "select nome from USUARIO "
      SQL = SQL & " where usuario_id = " & TabTemp.Fields("ct_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            If Trim(TabConsulta.Fields(0).Value) <> "" Then _
               NOME_CT_A = "" & Trim(TabConsulta.Fields(0).Value)
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'ENDEREÇO EMPRESA
      SQL = "select ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
      SQL = SQL & " ENDERECO.NUMERO, CEP.Cidade, CEP.UF, CEP.IBGE_ID, CEP.Cep_ID"
      SQL = SQL & " from ENDERECO "
      SQL = SQL & " INNER JOIN EMPRESA "
      SQL = SQL & " ON ENDERECO.PESSOA_ID = EMPRESA.PESSOA_ID "
      SQL = SQL & " LEFT OUTER JOIN CEP "
      SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"
      SQL = SQL & " Where EMPRESA.empresa_ID = " & EMPRESA_ID_N
      SQL = SQL & " and endereco.tipo = 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         ENDERECO_EMP_A = "" & Trim(TabConsulta.Fields("rua").Value)
         CEP_EMP_A = "" & Trim(TabConsulta.Fields("cep_id").Value)
         COMP_EMP_A = "" & Trim(TabConsulta.Fields("COMPLEMENTO").Value)
         NUMERO_EMP_A = "" & Trim(TabConsulta.Fields("NUMERO").Value)
         BAIRRO_EMP_A = "" & Trim(TabConsulta.Fields("BAIRRO").Value)
         CIDADE_EMP_A = "" & Trim(TabConsulta.Fields("cidade").Value)
         UF_EMP_A = "" & Trim(TabConsulta.Fields("uf").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE EMPRESA
      SQL = "select FONE.NUMERO, FONE.DDD, FONE.LOCAL"
      SQL = SQL & " from EMPRESA "
      SQL = SQL & " INNER JOIN PESSOA "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_EMP_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_EMP_A = FONE_EMP_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_EMP_A = FONE_EMP_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE cliente
      SQL = "select FONE.NUMERO, FONE.DDD, FONE.LOCAL "
      SQL = SQL & " from CLIENTE "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON CLIENTE.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " Where CLIENTE.PESSOA_ID = " & TabTemp.Fields("PESSOA_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_CLIENTE_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'EMPRESA
      SQL = "SELECT EMPRESA_ID, PESSOA.PESSOA_ID, CNPJCPF, DESCRICAO, RAZAO"
      SQL = SQL & " FROM EMPRESA "
      SQL = SQL & " INNER JOIN PESSOA "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID"
      SQL = SQL & " Where empresa_ID = " & EMPRESA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         CGC_A = "" & Trim(TabConsulta.Fields("CNPJCPF").Value)
         RAZAO_SOCIAL_A = "" & Trim(TabConsulta.Fields("RAZAO").Value)
         NOME_FANT_A = "" & Trim(TabConsulta.Fields("descricao").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      DT_FECHA_A = "" & TabTemp.Fields("DT_FECHA").Value
      If Trim(DT_FECHA_A) = "" Then
         DT_FECHA_A = ""
         Else: DT_FECHA_A = DMA(DT_FECHA_A)
      End If

COR_ID_N = 0 & TabTemp.Fields("COR_ID").Value
MARCA_ID_N = 0 & TabTemp.Fields("MARCA_ID").Value
TIPO_EQP_ID_N = 0 & TabTemp.Fields("TIPO_EQP").Value

      SQL = "insert into OSREL "
         SQL = SQL & "("
         SQL = SQL & "OS_ID,DT_OS,TIPO_OS,SITUACAO_OS,CONSULTOR_OS,"
         SQL = SQL & "KM_OS,PLACA_OS,estabelecimento_ID,DT_OS_FEHCA,NUMR_FROTA_OS,"
         SQL = SQL & "NOME_EMP,CNPJ_EMP,ENDERECO_EMP,NUMERO_EMP,COMPLEM_EMP,"
         SQL = SQL & "CEP_EMP,BAIRRO_EMP,CIDADE_EMP,UF_EMP,FONE_EMP,NOME_CLI,"
         SQL = SQL & "CNPJCPF_CLI,FONE_CLI,DESC_VEICULO,COR_VEICULO,MARCA_VEICULO,"
         SQL = SQL & "TIPO_VEICULO,ANO_VEICULO,MODELO_VEICULO,COMB_VEICULO,"
         SQL = SQL & "CHASSI_VEICULO,MOTOR_VEICULO,PESSOA_ID_CLIENTE"
         SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & NUMR_OS_N                                                               'OS_ID
         SQL = SQL & ",'" & DMA(TabTemp.Fields("dt_os").Value) & "'"                         'DT_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("H", TabTemp.Fields("tipo_os").Value) & "'"       'TIPO_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("Z", TabTemp.Fields("SITUACAO_OS").Value) & "'"   'SITUACAO_OS
         SQL = SQL & ",'" & Trim(Left(NOME_CT_A, 20)) & "'"                                           'CONSULTOR_OS
         SQL = SQL & "," & Trim(TabTemp.Fields("km").Value)                                  'KM_OS

         SQL = SQL & ",'" & Trim(txtPlaca.Text) & "'"             'PLACA_OS

         SQL = SQL & "," & ESTABELECIMENTO_ID_N                                              'estabelecimento_ID
         SQL = SQL & ",'" & DT_FECHA_A & "'"                                                 'DT_OS_FEHCA
         SQL = SQL & ",0"                                                                    'NUMR_FROTA_OS

         SQL = SQL & ",'" & Trim(Left(NOME_FANT_A, 100)) & "'"                                         'NOME_EMP
         SQL = SQL & ",'" & Trim(CGC_A) & "'"                                                'CNPJ_EMP

         SQL = SQL & ",'" & Trim(Replace(ENDERECO_EMP_A, ",", ".")) & "'"                    'ENDERECO_EMP
         SQL = SQL & "," & Trim(NUMERO_EMP_A)                                                'NUMERO_EMP
         SQL = SQL & ",'" & Trim(Replace(COMP_EMP_A, ",", ".")) & "'"                        'COMPLEM_EMP
         SQL = SQL & ",'" & Trim(CEP_EMP_A) & "'"                                            'CEP_EMP
         SQL = SQL & ",'" & Trim(Replace(BAIRRO_EMP_A, ",", ".")) & "'"                      'BAIRRO_EMP
         SQL = SQL & ",'" & Trim(CIDADE_EMP_A) & "'"                                         'CIDADE_EMP
         SQL = SQL & ",'" & Trim(UF_EMP_A) & "'"                                             'UF_EMP
         SQL = SQL & ",'" & Trim(FONE_EMP_A) & "'"                                           'FONE_EMP

         SQL = SQL & ",'" & Trim(NOME_CLI) & "'"                                             'NOME_CLI

         SQL = SQL & ",'" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"                      'CNPJCPF_CLI
         SQL = SQL & ",'" & Trim(FONE_CLIENTE_A) & "'"                                       'FONE_CLI
         SQL = SQL & ",'" & Trim(TabTemp.Fields("nome_eqp").Value) & "'"                     'DESC_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("S", Str(COR_ID_N)) & "'"                               'COR_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("W", Str(MARCA_ID_N)) & "'"                            'MARCA_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("A", Str(TIPO_EQP_ID_N)) & "'"                         'TIPO_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("ano").Value) & "'"                                 'ANO_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("modelo").Value) & "'"                              'MODELO_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("U", Str(TIPO_EQP_ID_N)) & "'"                         'COMB_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("identificacao").Value) & "'"                'CHASSI_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("EQUIPAMENTO_ID").Value) & "'"               'MOTOR_VEICULO
         SQL = SQL & "," & TabTemp.Fields("PESSOA_id").Value                                 'PESSOA_ID_CLIENTE
      SQL = SQL & ")"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

'ITENS SERVIÇO
      SQL = "select * from OSSERVICO "
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "select nome from USUARIO "
         SQL = SQL & " where usuario_id = " & TabConsulta.Fields("responsavel_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)
         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSSERVICO_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'S' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            If TabItem.State = 1 Then _
               TabItem.Close

            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,"
               SQL = SQL & "DT_CAD,DESCRICAO,VALR_ITEM,VALR_DESCONTO,QTDE,"
               SQL = SQL & " RESPONSAVEL, CODG_PRODUTO "
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSSERVICO_ID").Value              'OSRELITEM_ID
               SQL = SQL & ",'S'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("responsavel_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSTAREFA_ID").Value               'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("DESCRICAO").Value) & "'"    'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_servico").Value)    'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_servico").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(1)                                            'QTDE
               SQL = SQL & ",'" & Trim(Left(RESPONSAVEL_A, 20)) & "'"                  'RESPONSAVEL
               SQL = SQL & ",''"                                                       'CODG_PRODUTO
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

'ITENS PRODUTO
      SQL = "select OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
      SQL = SQL & " from OSPECA "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         DT_GARANTIA_D = 0
         If Not IsNull(TabConsulta.Fields("dt_garantia").Value) Then _
            DT_GARANTIA_D = TabConsulta.Fields("dt_garantia").Value

         NOME_A = Replace(TabConsulta.Fields("DESCRICAO").Value, ",", ".")
         NOME_A = Replace(NOME_A, "'", "´")

         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "select descricao from vwVendedor "
         SQL = SQL & " where vendedor_id = " & TabConsulta.Fields("SOLICITANTE_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)
         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSPECA_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'P' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            If TabItem.State = 1 Then _
               TabItem.Close

            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,DT_CAD,DESCRICAO,"
               SQL = SQL & "VALR_ITEM,VALR_DESCONTO,QTDE,RESPONSAVEL,CODG_PRODUTO,dt_garantia"
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'OSRELITEM_ID
               SQL = SQL & ",'P'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("SOLICITANTE_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(NOME_A) & "'"                                   'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_ITEM").Value)       'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_PRODUTO").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTDE").Value)             'QTDE
               SQL = SQL & ",'" & Trim(RESPONSAVEL_A) & "'"                            'RESPONSAVEL
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("CODG_PRODUTO").Value) & "'" 'CODG_PRODUTO
               SQL = SQL & ",'" & DMA(DT_GARANTIA_D) & "'"                              'DT_garantia
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If

         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

'Sleep 3000

   FORMULA_REL = "{OSREL.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = FORMULA_REL & " and {OSREL.OS_ID} = " & NUMR_OS_N

   If chkImp.Value = 1 Then
      If ESCOLHE_IMPRESSORA(NOME_BANCO_DADOS) = True Then
         If EQP_VEICULO = False Then
            Nome_Relatorio = "REL_OFICINA.rpt"
            Else: Nome_Relatorio = "REL_SERVICO.rpt"
         End If
         Nome_Relatorio = "REL_OFICINA.rpt"
         frmRELATORIO10.Show 1
      End If
      Else
         If EQP_VEICULO = False Then
            Nome_Relatorio = "REL_OFICINA.rpt"
            Else: Nome_Relatorio = "REL_SERVICO.rpt"
         End If
         Nome_Relatorio = "REL_OFICINA.rpt"
         frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIMIR_ORDEM_SERVICO"
End Sub

Function EQP_VEICULO() As Boolean
'On Error GoTo ERRO_TRATA

   EQP_VEICULO = True
   If optEqp.Value = True Then
      txtPlaca.Visible = False
      txtEqp.Visible = True
      lblEqp.Caption = "Eqp: "
      Else  'AQUI É VEÍCULO
         EQP_VEICULO = False
         txtPlaca.Visible = True
         txtEqp.Visible = False
         lblEqp.Caption = "Placa: "
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EQP_VEICULO"
End Function

Sub GRAVA_OSVEICEQP()
'On Error GoTo ERRO_TRATA

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from OSVEICEQP"
   SQL = SQL & " where os_id = " & OS_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      SQL = "update OSVEICEQP set "
         'SQL = SQL & " VEICULO_ID = " & VEICULO_ID_N              'VEICULO_ID
         'SQL = SQL & ", EQUIPAMENTO_ID_n = " & EQUIPAMENTO_ID_N   'EQUIPAMENTO_ID

         If VEICULO_ID_N > 0 Then
            SQL = SQL & " VEICULO_ID = " & VEICULO_ID_N      'VEICULO_ID
            Else: SQL = SQL & " VEICULO_ID = NULL"           'VEICULO_ID
         End If
         If EQUIPAMENTO_ID_N > 0 Then
            SQL = SQL & ",EQUIPAMENTO_ID = " & EQUIPAMENTO_ID_N  'EQUIPAMENTO_ID
            Else: SQL = SQL & ",EQUIPAMENTO_ID = NULL"           'EQUIPAMENTO_ID
         End If
      SQL = SQL & " where os_id = " & OS_ID_N
      Else
         SQL = "insert into OSVEICEQP "
            SQL = SQL & "(OS_ID,VEICULO_ID,EQUIPAMENTO_ID)"
         SQL = SQL & " values ( "
            SQL = SQL & OS_ID_N                    'OS_ID
            If VEICULO_ID_N > 0 Then
               SQL = SQL & "," & VEICULO_ID_N      'VEICULO_ID
               Else: SQL = SQL & ",NULL"           'VEICULO_ID
            End If
            If EQUIPAMENTO_ID_N > 0 Then
               SQL = SQL & "," & EQUIPAMENTO_ID_N  'EQUIPAMENTO_ID
               Else: SQL = SQL & ",NULL"           'EQUIPAMENTO_ID
            End If
         SQL = SQL & " )"
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_OSVEICEQP"
End Sub


