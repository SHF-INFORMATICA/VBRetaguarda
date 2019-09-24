VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOSVEICULO 
   Caption         =   "Abertura de Ordem de Serviço"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   11865
   Icon            =   "OSABRE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotGeralServico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8115
      TabIndex        =   74
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox txtTotGeralProduto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   8115
      TabIndex        =   73
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox txtTotProduto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   5280
      TabIndex        =   71
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox txtTotServico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5280
      TabIndex        =   69
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox txtTotDescontoServico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2250
      TabIndex        =   67
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox txtTotDescontoProduto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2250
      TabIndex        =   65
      Top             =   8520
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Produtos Ordem de Serviço"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2535
      Left            =   50
      TabIndex        =   56
      Top             =   5400
      Width           =   11775
      Begin VB.CommandButton cmdProduto 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2444
         Picture         =   "OSABRE.frx":47C4A
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.ComboBox cmbVendedorAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtTotalProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         MaxLength       =   6
         TabIndex        =   23
         ToolTipText     =   "Digite a Quantidade"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cmbVendedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         TabIndex        =   22
         ToolTipText     =   "Responsável Venda"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDescontoProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   24
         ToolTipText     =   "Desconto Produto"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtValorProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   25
         ToolTipText     =   "Valor Venda Produto"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   20
         ToolTipText     =   "Informe Código Produto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtDescProduto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   5775
      End
      Begin MSComctlLib.ListView lstProduto 
         Height          =   1125
         Left            =   45
         TabIndex        =   57
         Top             =   1320
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   1984
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
         NumItems        =   7
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
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Item ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   9000
         TabIndex        =   63
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtde ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   240
         TabIndex        =   62
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   8730
         TabIndex        =   61
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   2940
         TabIndex        =   60
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Item = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   6000
         TabIndex        =   59
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   90
         TabIndex        =   58
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serviços Ordem de Serviço"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2535
      Left            =   50
      TabIndex        =   48
      Top             =   2880
      Width           =   11775
      Begin VB.CommandButton cmdServico 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   1920
         Picture         =   "OSABRE.frx":4864C
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.ComboBox cmbMecanicoAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9120
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtServico 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   14
         ToolTipText     =   "Digite Código Tarefa ou 0 para Diversar"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDescTarefa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         ToolTipText     =   "Descrição Serviço"
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txtValorTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         ToolTipText     =   "Informe Valor Serviço"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDescontoTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         ToolTipText     =   "Desconto Serviço"
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbMecanico 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9120
         TabIndex        =   16
         ToolTipText     =   "Selecione Mecanico para tarefa"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtTotalTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin MSComctlLib.ListView lstServiço 
         Height          =   1125
         Left            =   45
         TabIndex        =   55
         Top             =   1320
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   1984
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
         NumItems        =   6
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
      End
      Begin VB.Label lblTarefa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serviço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   0
         TabIndex        =   53
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Serviço ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4920
         TabIndex        =   52
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   720
         TabIndex        =   51
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mecânico:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8040
         TabIndex        =   50
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Serviço ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8160
         TabIndex        =   49
         Top             =   840
         Width           =   1965
      End
   End
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
      Left            =   10560
      TabIndex        =   39
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem de Serviço"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   45
      TabIndex        =   32
      Top             =   600
      Width           =   11775
      Begin VB.ComboBox cmbSituacaoAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         TabIndex        =   80
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbTipoOSAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbConsultorAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdOS 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   1500
         Picture         =   "OSABRE.frx":4904E
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   480
         Width           =   405
      End
      Begin VB.ComboBox cmbConsultor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   2
         ToolTipText     =   "Selecione Consultor Tecnico"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtOS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "<<Enter>> gerar nova O.S."
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipoOS 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   3
         ToolTipText     =   "Selecione Tipo Ordem de Serviço"
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox cmbSituacao 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         TabIndex        =   4
         ToolTipText     =   "Selecione situação Ordem de Serviço"
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDtOS 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   37
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblCt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consultor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblOs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número O.S."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo O.S."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         TabIndex        =   34
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10080
         TabIndex        =   33
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados Veículo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1335
      Left            =   50
      TabIndex        =   27
      Top             =   1560
      Width           =   11775
      Begin PVMaskEditLib.PVMaskEdit txtPlaca 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Informe a Placa do Veículo"
         Top             =   360
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
      Begin VB.TextBox txtFrota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtCombustivel 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtMODELO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5160
         MaxLength       =   4
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtANO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdCadPlaca 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3000
         Picture         =   "OSABRE.frx":49A50
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Consulta Cadastro Veículo"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtKM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Informe Kilometragem atual"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   8
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton cmdConsultaPlaca 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2550
         Picture         =   "OSABRE.frx":4BB7A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   360
         Width           =   405
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Frota:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   3720
         TabIndex        =   79
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Combustível:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   8640
         TabIndex        =   44
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   6360
         TabIndex        =   43
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ano/Modelo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   3015
         TabIndex        =   42
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km Atual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   165
         TabIndex        =   31
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   495
         TabIndex        =   30
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   5460
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4C57C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4C9D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4CCEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4D140
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4D594
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4D8B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSABRE.frx":4DD08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
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
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
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
      DesignHeight    =   8910
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Serviço = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6525
      TabIndex        =   76
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Produto = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6495
      TabIndex        =   75
      Top             =   8520
      Width           =   1530
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal Produto = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3285
      TabIndex        =   72
      Top             =   8520
      Width           =   1905
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal Serviço = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3315
      TabIndex        =   70
      Top             =   8160
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desc. Serviço = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   68
      Top             =   8160
      Width           =   2070
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desc. Produto = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   66
      Top             =   8520
      Width           =   2100
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
      Left            =   9180
      TabIndex        =   40
      Top             =   8280
      Width           =   1230
   End
End
Attribute VB_Name = "frmOSVEICULO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim EQUIPAMENTO_ID_N       As Long
   Dim TAREFA_ID_N            As Long
   Dim Situação_Ordem_Serviço As String
   Dim DT_FECHAMENTO_OS       As Date

Private Sub Form_Load()
   ABRE_BANCO_MEGASIM NOME_BANCO_DADOS

   LIMPA_OS
   CARREGA_COMBOS

   If SINAL_INDICADOR_N = 2 Then
      Me.Caption = "Fechamento Ordem de Serviço"
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERRO_TRATA

   Select Case Button.key
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
               IMPRIMIR_ORDEM_SERVIÇO txtOs.Text
      Case "excluir"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbSituacao_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbSituacao_KeyPress"
End Sub

Private Sub cmbConsultor_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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

Err.Clear
End Sub

Private Sub txtDescontoProduto_GotFocus()
   txtDESCONTOPRODUTO.SelStart = 0
   txtDESCONTOPRODUTO.SelLength = Len(txtDESCONTOPRODUTO.Text)
End Sub

Private Sub txtDescontoTarefa_GotFocus()
   txtDescontoTarefa.SelStart = 0
   txtDescontoTarefa.SelLength = Len(txtDescontoTarefa.Text)
End Sub

Private Sub txtDescTarefa_GotFocus()
   txtDescTarefa.SelStart = 0
   txtDescTarefa.SelLength = Len(txtDescTarefa.Text)
End Sub

Private Sub txtDtOS_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtKM_GotFocus()
   txtKM.SelStart = 0
   txtKM.SelLength = Len(txtKM.Text)
End Sub

Private Sub txtOs_GotFocus()
   txtOs.SelStart = 0
   txtOs.SelLength = Len(txtOs.Text)
End Sub

Private Sub txtOS_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
On Error GoTo ERRO_TRATA

   INDR_PRI = False
   If Trim(txtOs.Text) <> "" Then
      If IsNumeric(txtOs.Text) Then
         NUMR_REQ_N = txtOs.Text

         MOSTRA_OS

         Exit Sub
      End If
   End If

   If INDR_PRI = False Then
      GERA_NUMR_REQ
      txtOs.Text = NUMR_REQ_N
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOS_KeyPress"
End Sub

Private Sub txtKM_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtKM_KeyPress"
End Sub

Private Sub cmdos_Click()
On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSCONSULTA.Show 1
   If SQL3 <> "" Then _
      txtOs.Text = SQL3
   SQL3 = ""
   txtOs.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdos_Click"
End Sub

Private Sub cmdConsultaPlaca_Click()
On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSVEICULOCONSULTA.Show 1
   If SQL3 <> "" Then _
      txtPLACA.Text = SQL3
   SQL3 = ""
   txtPLACA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsultaPlaca_Click"
End Sub

Private Sub cmdCadPlaca_Click()
   frmOSVEICULOCADASTRO.Show 1
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSVEICULOCONSULTA.Show 1
         If Trim(SQL3) <> "" Then
            If TabAUX.State = 1 Then _
               TabAUX.Close

            SQL = "select placa from VEICULO "
            SQL = SQL & " where placa = '" & Trim(SQL3) & "'"
            TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabAUX.EOF Then _
               MOSTRA_VEICULO
            If TabAUX.State = 1 Then _
               TabAUX.Close
         End If
         SQL3 = ""
         txtPLACA.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPLACA_KeyDown"
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtplaca_KeyPress"
End Sub

Private Sub txtPLACA_LostFocus()
   If Trim(txtPLACA.Text) <> "" Then _
      MOSTRA_VEICULO
End Sub

Private Sub cmdProduto_Click()
On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmPRODUTOCONSULTA.Show 1
   If SQL3 <> "" Then _
      txtPRODUTO.Text = SQL3
   SQL3 = ""
   txtPRODUTO.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdPRODUTO_Click"
End Sub

Private Sub txtProduto_GotFocus()
   txtPRODUTO.SelStart = 0
   txtPRODUTO.SelLength = Len(txtPRODUTO.Text)
End Sub

Private Sub txtQTDE_GotFocus()
   txtQtde.SelStart = 0
   txtQtde.SelLength = Len(txtQtde.Text)
End Sub

Private Sub cmdSERVIcO_Click()
On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSSERVICOCONSULTA.Show 1
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
End Sub

Private Sub TXTSERVICO_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMecanico_KeyPress"
End Sub

Private Sub txtDescontoTarefa_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      VALOR_DESCONTO_N = 0 & txtDescontoTarefa.Text

      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Then
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
End Sub

Private Sub txtTotDescontoProduto_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtTotDescontoServico_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtTotGeralProduto_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtTotGeralServico_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtTotOS_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtTotProduto_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtTotServico_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtValorProduto_GotFocus()
   txtValorProduto.SelStart = 0
   txtValorProduto.SelLength = Len(txtValorProduto.Text)
End Sub

Private Sub txtValorTarefa_GotFocus()
   txtValorTarefa.SelStart = 0
   txtValorTarefa.SelLength = Len(txtValorTarefa.Text)
End Sub

Private Sub txtValorTarefa_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then

      If CHECA_DADOS_OS = True Then _
         GRAVA_SERVIÇO

      KeyAscii = 0
      'SendKeys ("{tab}")
      txtServico.SetFocus
      Else
         If KeyAscii = 8 Then
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
End Sub

Private Sub TXTSERVICO_LostFocus()
On Error GoTo ERRO_TRATA

   INDR_PRI = False
   If Trim(txtServico.Text) <> "" Then
      If IsNumeric(txtServico.Text) Then
         Else: INDR_PRI = True
      End If
      Else: INDR_PRI = True
   End If

   If INDR_PRI = True Then _
      txtServico.Text = MAX_ID("osservico_id", "OSSERVICO", "OS_ID", txtOs.Text, "", "")

   MOSTRA_TAREFA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSERVICO_LostFocus"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      MOSTRA_PRODUTO

      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub cmbVENDEDOR_Click()
On Error Resume Next

   cmbVendedorAUX.ListIndex = cmbVENDEDOR.ListIndex

Err.Clear
End Sub

Private Sub cmbvendedor_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVendedor_KeyPress"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
   If txtQtde.Text = "" Then _
      txtQtde.Text = 0

   txtQtde.Text = Format(txtQtde.Text, strFormatacao2Digitos)
End Sub

Private Sub txtDescontoProduto_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescontoProduto_KeyPress"
End Sub

Private Sub txtdescontoproduto_LostFocus()
   If txtDESCONTOPRODUTO.Text = "" Then _
      txtDESCONTOPRODUTO.Text = 0

   txtDESCONTOPRODUTO.Text = Format(txtDESCONTOPRODUTO.Text, strFormatacao2Digitos)
End Sub

Private Sub txtValorProduto_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then

      If CHECA_DADOS_OS = True Then _
         GRAVA_PRODUTO

      KeyAscii = 0
      'SendKeys ("{tab}")
      txtPRODUTO.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorProduto_KeyPress"
End Sub

Private Sub txtvalorproduto_LostFocus()
   If txtValorProduto.Text = "" Then _
      txtValorProduto.Text = 0

   txtValorProduto.Text = Format(txtValorProduto.Text, strFormatacao2Digitos)

   TOTALIZA_CAMPOS
End Sub

Sub LIMPA_OS()
On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(5).Enabled = True
   Toolbar1.Buttons(9).Enabled = True

   VENDEDOR_ID_N = 0
   PEDIDO_ID_N = 0
   CLIENTE_ID_N = 0
   PEDIDO_ID_N = NUMR_REQ_N
   txtCNPJCPF.PromptInclude = False

   Situação_Ordem_Serviço = ""
   txtTotDescontoServico.Text = Format(0, strFormatacao2Digitos)
   txtTotServico.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralServico.Text = Format(0, strFormatacao2Digitos)
   txtTotDescontoProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotOS.Text = Format(0, strFormatacao2Digitos)
   txtFrota.Text = ""

   lstServiço.ListItems.Clear
   lstProduto.ListItems.Clear

   EQUIPAMENTO_ID_N = 0
   txtOs.Text = ""
   txtDtOS.PromptInclude = False
   txtDtOS.Text = Date
   txtDtOS.PromptInclude = True
   cmbConsultor.Text = ""
   cmbConsultorAUX.Text = ""
   cmbTipoOS.Text = ""
   cmbTipoOSAUX.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   txtPLACA.Text = ""
   txtCNPJCPF.Text = ""
   txtCliente.Text = ""
   txtKM.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   txtMarca.Text = ""
   txtCombustivel.Text = ""
   txtTotOS.Text = ""

   LIMPA_PRODUTO
   LIMPA_SERVIÇO
   CARREGA_COMBOS

   HABILITA_TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_OS"
End Sub

Sub LIMPA_SERVIÇO()
   txtServico.Text = ""
   txtDescTarefa.Text = ""
   cmbMecanicoAUX.Text = ""
   cmbMecanico.Text = ""
   txtDescontoTarefa.Text = ""
   txtValorTarefa.Text = ""
   txtTotalTarefa.Text = ""
End Sub

Sub LIMPA_PRODUTO()
   PRODUTO_ID_N = 0
   txtPRODUTO.Text = ""
   txtDESCPRODUTO.Text = ""
   cmbVendedorAUX.Text = ""
   cmbVENDEDOR.Text = ""
   txtQtde.Text = ""
   txtDESCONTOPRODUTO.Text = ""
   txtValorProduto.Text = ""
   txtTOTALPRODUTO.Text = ""
End Sub

Sub CARREGA_COMBOS()
On Error GoTo ERRO_TRATA

'parametros combos x tabela descr
'8 = consultor tecnico
'9 = mecanico

   cmbConsultorAUX.Clear
   cmbConsultor.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 8 "   'consultor tecnico
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
   SQL = SQL & " where tipo = 9 "   'mecanico
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbMecanicoAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbMecanico.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbVendedorAUX.Clear
   cmbVENDEDOR.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select vendedor_id, nome_vend from VENDEDOR "
   SQL = SQL & " where status = 'A' "   'vendedor
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbVendedorAUX.AddItem TabDESCR.Fields("vendedor_id").Value
      cmbVENDEDOR.AddItem Trim(TabDESCR.Fields("nome_vend").Value) & "-" & Trim(TabDESCR.Fields("vendedor_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbTipoOSAUX.Clear
   cmbTipoOS.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo_a = 'H' "
   SQL = SQL & "order by desc_a"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipoOS.AddItem Trim(TabDESCR!desc_a)
      cmbTipoOSAUX.AddItem TabDESCR!codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbSituacao.Clear

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo_a = 'Z' "
   SQL = SQL & "order by desc_a"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbSituacao.AddItem Trim(TabDESCR!desc_a)
      cmbSituacaoAUX.AddItem Trim(TabDESCR.Fields("CODIGO").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   If SINAL_INDICADOR_N = 2 Then
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
On Error GoTo ERRO_TRATA

   LIMPA_OS

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value
      Situação_Ordem_Serviço = TabTemp.Fields("situacao_os").Value
      EQUIPAMENTO_ID_N = 0 & TabTemp.Fields("EQUIPAMENTO_ID").Value
      txtOs.Text = NUMR_REQ_N
      txtDtOS.PromptInclude = False
         txtDtOS.Text = "" & TabTemp.Fields("dt_os").Value
      txtDtOS.PromptInclude = True

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      If Not IsNull(TabTemp.Fields("ct_id").Value) Then
         If IsNumeric(TabTemp.Fields("ct_id").Value) Then
            cmbConsultorAUX.Text = Trim(TabTemp.Fields("ct_id").Value)

            SQL = "select nome from USUARIO "
            SQL = SQL & " where usuario_id = " & TabTemp.Fields("ct_id").Value
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then _
               cmbConsultor.Text = "" & Trim(TabDESCR.Fields(0).Value)
         End If
      End If

      cmbTipoOS.Text = "" & TRAZ_DESCRITOR("H", TabTemp.Fields("tipo_os").Value)
      cmbTipoOSAUX.Text = "" & Trim(TabTemp.Fields("tipo_os").Value)
      txtFrota.Text = "" & TabTemp.Fields("numr_frota").Value

      cmbSituacao.Text = "" & TRAZ_DESCRITOR("Z", TabTemp.Fields("SITUACAO_os").Value)

      txtPLACA.Text = "" & Trim(TabTemp.Fields("placa").Value)
      txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("CNPJCPF").Value)
      txtCliente.Text = "" & Trim(TabTemp.Fields("nome_cliente").Value)
      txtKM.Text = "" & Trim(TabTemp.Fields("km").Value)
      txtANO.Text = "" & Trim(TabTemp.Fields("ano").Value)
      txtMODELO.Text = "" & Trim(TabTemp.Fields("modelo").Value)

      txtMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)
      txtCombustivel.Text = "" & TRAZ_DESCRITOR("U", TabTemp.Fields("combustivel_id").Value)

      txtTotOS.Text = ""
      DT_FECHAMENTO_OS = 0 & TabTemp.Fields("dt_fecha").Value
      Else
         INDR_PRI = True
         txtOs.Text = NUMR_REQ_N
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SETA_GRID_SERVIÇO
   SETA_GRID_PRODUTO

   TOTALIZA_CAMPOS

   If Situação_Ordem_Serviço = "C" Then
      DESAABILITA_TELA
      MsgBox "Ordem de Serviço CANCELADA, permitido somente consulta."
   End If
   If Situação_Ordem_Serviço = "F" Then
      DESAABILITA_TELA
      MsgBox "Ordem de Serviço FECHADA, permitido somente consulta."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OS"
End Sub

Private Sub MOSTRA_VEICULO()
On Error GoTo ERRO_TRATA

   If Trim(txtPLACA.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwRel_VEICULO "
      SQL = SQL & " where placa = '" & Trim(txtPLACA.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         EQUIPAMENTO_ID_N = TabTemp.Fields("EQUIPAMENTO_ID").Value

         txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         txtCliente.Text = "" & Trim(TabTemp.Fields("nome_cliente").Value)
         PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value

         'txtCHASSI.Text = "" & Trim(TabTemp!chassi)
         'txtDescricao.Text = "" & TabTemp!Descricao
         'txtMotor.Text = "" & TabTemp!motor
         txtFrota.Text = "" & TabTemp.Fields("numr_frota").Value
         txtANO.Text = "" & TabTemp!Ano
         txtMODELO.Text = "" & TabTemp!modelo

         'If Not IsNull(TabTemp.Fields("cor_id").Value) Then
         '   If IsNumeric(TabTemp.Fields("cor_id").Value) Then
               'cmbCor.Text = "" & TRAZ_DESCRITOR("S", TabTemp.Fields("cor_id").Value)
               'cmbCorAUX.Text = "" & TabTemp.Fields("cor_id").Value
         '   End If
         'End If

         If Not IsNull(TabTemp.Fields("marca_id").Value) Then _
            If IsNumeric(TabTemp.Fields("marca_id").Value) Then _
               txtMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)

         If Not IsNull(TabTemp!COMBUSTIVEL_ID) Then _
            If IsNumeric(TabTemp.Fields("combustivel_id").Value) Then _
               txtCombustivel.Text = "" & TRAZ_DESCRITOR("U", TabTemp.Fields("combustivel_id").Value)

         'If Not IsNull(TabTemp!tipo_eqp) Then
         '   If IsNumeric(TabTemp!tipo_eqp) Then
         '      cmbTipo.Text = "" & TRAZ_DESCRITOR("A", TabTemp!tipo_eqp)
         '      cmbTipoAUX.Text = "" & TabTemp!tipo_eqp
         '   End If
         'End If
         Else
            MsgBox "Placa não encontrada."
            txtPLACA.SetFocus
            txtPLACA.Text = ""
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VEICULO"
End Sub

Sub GRAVA_OS()
On Error GoTo ERRO_TRATA

   NUMR_REQ_N = 0 & txtOs.Text

   If SINAL_INDICADOR_N = 1 Then
      If TabCABECA.State = 1 Then _
         TabCABECA.Close
   
      SQL = "select * from vwOS "
      SQL = SQL & " where os_id = " & NUMR_REQ_N
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCABECA.EOF Then
         SQL = "update OS set "
            SQL = SQL & " EQUIPAMENTO_ID = " & EQUIPAMENTO_ID_N         'EQUIPAMENTO_ID
            SQL = SQL & ", DT_OS = '" & DMA(txtDtOS.Text) & "'"         'DT_OS
            SQL = SQL & ", TIPO_OS = " & cmbTipoOSAUX.Text              'TIPO_OS
            SQL = SQL & ", SITUACAO_OS = 0" & cmbSituacaoAUX.Text       'SITUACAO_OS
            SQL = SQL & ", KM = " & txtKM.Text                          'KM
            SQL = SQL & ", CT_ID = " & cmbConsultorAUX.Text             'CT_ID
            SQL = SQL & ", NUMR_FROTA = '" & Trim(txtFrota.Text) & "'"  'NUMR_FROTA
         SQL = SQL & " where os_id = " & NUMR_REQ_N
         Else
            SQL = "insert into OS "
               SQL = SQL & "(OS_ID,EMPRESA_ID,EQUIPAMENTO_ID,DT_OS,TIPO_OS,SITUACAO_OS,KM,CT_ID,NUMR_FROTA)"
            SQL = SQL & " values ( "
               SQL = SQL & NUMR_REQ_N                               'OS_ID
               SQL = SQL & "," & EMPRESA_ID_N                       'EMPRESA_ID
               SQL = SQL & "," & EQUIPAMENTO_ID_N                   'EQUIPAMENTO_ID
               SQL = SQL & ",'" & DMA(txtDtOS.Text) & "'"           'DT_OS
               SQL = SQL & "," & cmbTipoOSAUX.Text                  'TIPO_OS
               SQL = SQL & ",0" & cmbSituacaoAUX.Text               'SITUACAO_OS
               SQL = SQL & "," & txtKM.Text                         'KM
               SQL = SQL & "," & cmbConsultorAUX.Text               'CT_ID
               SQL = SQL & ",'" & Trim(txtFrota.Text) & "'"         'NUMR_FROTA
            SQL = SQL & "  )"
      End If
      
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If
   If SINAL_INDICADOR_N = 2 Then
      FECHA_OS
      MsgBox "Ordem de Serviço fechada com sucesso.   "
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_OS"
End Sub

Sub GRAVA_SERVIÇO()
On Error GoTo ERRO_TRATA

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
      MsgBox "Mecanico não informado."
      cmbMecanico.SetFocus
      Exit Sub
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

   GRAVA_OS

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      SQL = "update OSSERVICO set "
         SQL = SQL & " RESPONSAVEL_ID = " & cmbMecanicoAUX.Text                     'RESPONSAVEL_ID
         SQL = SQL & ", VALOR_SERVICO = " & tpMOEDA(txtValorTarefa.Text)            'VALOR_SERVICO
         SQL = SQL & ", DESCRICAO = '" & Trim(txtDescTarefa.Text) & "'"             'DESCRICAO
         SQL = SQL & ", DESCONTO_SERVICO = " & tpMOEDA(txtDescontoTarefa.Text)      'DESCONTO_SERVICO
      SQL = SQL & " where os_id = " & NUMR_REQ_N
      SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text                          'OSSERVICO_ID
      Else
         SQL = "insert into OSSERVICO "
            SQL = SQL & "(OSSERVICO_ID,OS_ID,OSTAREFA_ID,DT_CAD,RESPONSAVEL_ID,VALOR_SERVICO,DESCRICAO,desconto_servico) "
         SQL = SQL & " values ( "
            SQL = SQL & MAX_ID("osservico_id", "OSSERVICO", "", "", "", "")   'OSSERVICO_ID
            SQL = SQL & "," & NUMR_REQ_N                                      'OS_ID
            SQL = SQL & "," & txtServico.Text                                 'OSTAREFA_ID
            SQL = SQL & ",'" & DMA(Date) & "'"                                'DT_CAD
            SQL = SQL & "," & cmbMecanicoAUX.Text                             'RESPONSAVEL_ID
            SQL = SQL & "," & tpMOEDA(txtValorTarefa.Text)                    'VALOR_SERVICO
            SQL = SQL & ",'" & Trim(txtDescTarefa.Text) & "'"                 'DESCRICAO
            SQL = SQL & "," & tpMOEDA(txtDescontoTarefa.Text)                 'DESCONTO_SERVICO
         SQL = SQL & "  )"
   End If
   
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID_SERVIÇO
   LIMPA_SERVIÇO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_SERVIÇO"
End Sub

Sub GRAVA_PRODUTO()
On Error GoTo ERRO_TRATA

   If PRODUTO_ID_N <= 0 Then
      MsgBox "produto inválido."
      txtPRODUTO.SetFocus
      Exit Sub
   End If
   If Trim(txtDESCPRODUTO.Text) = "" Then
      MsgBox "Descrição produto inválida."
      txtDESCPRODUTO.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(cmbVendedorAUX.Text) Then
      MsgBox "Vendedor não informado."
      cmbVENDEDOR.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtDESCONTOPRODUTO.Text) Then _
      txtDESCONTOPRODUTO.Text = 0

   If Trim(txtQtde.Text) = "" Then
      MsgBox "Quantidade informada inválida."
      txtQtde.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtQtde.Text) Then
      MsgBox "Quantidade informada inválida."
      txtQtde.SetFocus
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

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      SQL = "update OSPECA set "
         SQL = SQL & " PRODUTO_ID = " & PRODUTO_ID_N                             'PRODUTO_ID
         SQL = SQL & ", DT_CAD = '" & DMA(Date) & "'"                            'DT_CAD
         SQL = SQL & ", SOLICITANTE_ID = " & cmbVendedorAUX.Text                 'SOLICITANTE_ID
         SQL = SQL & ", VALOR_ITEM = " & tpMOEDA(txtValorProduto.Text)           'VALOR_ITEM
         SQL = SQL & ", DESCONTO_PRODUTO = " & tpMOEDA(txtDescontoTarefa.Text)   'DESCONTO_PRODUTO
         SQL = SQL & ", QTDE = " & tpMOEDA(txtQtde.Text)                         'QTDE
      SQL = SQL & " where os_id = " & NUMR_REQ_N
      SQL = SQL & " and OSPECA_ID = " & TabCABECA.Fields("OSPECA_ID").Value      'OSPECA_ID
      SQL = SQL & " PRODUTO_ID = " & PRODUTO_ID_N
      Else
         SQL = "insert into OSPECA "
            SQL = SQL & "(OSPECA_ID,OS_ID,PRODUTO_ID,DT_CAD,SOLICITANTE_ID,VALOR_ITEM,DESCONTO_PRODUTO,QTDE) "
         SQL = SQL & " values ( "
            SQL = SQL & MAX_ID("OSPECA_id", "OSPECA", "", "", "", "")         'OSPECA_ID
            SQL = SQL & "," & NUMR_REQ_N                                      'OS_ID
            SQL = SQL & "," & PRODUTO_ID_N                                    'PRODUTO_ID
            SQL = SQL & ",'" & DMA(Date) & "'"                                'DT_CAD
            SQL = SQL & "," & cmbVendedorAUX.Text                             'SOLICITANTE_ID
            SQL = SQL & "," & tpMOEDA(txtValorProduto.Text)                   'VALOR_ITEM
            SQL = SQL & "," & tpMOEDA(txtDescontoTarefa.Text)                 'DESCONTO_PRODUTO
            SQL = SQL & "," & tpMOEDA(txtQtde.Text)                           'QTDE
         SQL = SQL & "  )"
   End If
   
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Debug.Print SQL

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID_PRODUTO
   LIMPA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO"
End Sub

Sub SETA_GRID_SERVIÇO()
On Error GoTo ERRO_TRATA

   lstServiço.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSSERVICO "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set Item = lstServiço.ListItems.Add(, "seq." & TabTemp.Fields("OSSERVICO_ID").Value, TabTemp.Fields("OSSERVICO_ID").Value)

      Item.SubItems(1) = "" & Trim(TabTemp!Descricao)
      Item.SubItems(2) = "" & Format(TabTemp!VALOR_SERVICO, strFormatacao2Digitos)
      Item.SubItems(3) = "" & Format(TabTemp!DESCONTO_SERVICO, strFormatacao2Digitos)
      Item.SubItems(4) = "" & Format(TabTemp!VALOR_SERVICO - TabTemp!DESCONTO_SERVICO, strFormatacao2Digitos)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome from USUARIO "
      SQL = SQL & " where USUARIO_ID = " & TabTemp.Fields("RESPONSAVEL_ID").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         Item.SubItems(5) = "" & Trim(TabConsulta.Fields("nome").Value)

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
On Error GoTo ERRO_TRATA

   lstProduto.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT OSPECA.OSPECA_ID, OSPECA.OS_ID, OSPECA.PRODUTO_ID, OSPECA.DT_CAD, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, "
   SQL = SQL & " OSPECA.QTDE, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.QTDE AS QTDE_ESTOQUE"
   SQL = SQL & " FROM OSPECA "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set Item = lstProduto.ListItems.Add(, "seq." & TabTemp.Fields("OSPECA_ID").Value, _
                                                     TabTemp.Fields("CODG_PRODUTO").Value)

      Item.SubItems(1) = "" & Trim(TabTemp!Descricao)
      Item.SubItems(2) = "" & Format(TabTemp.Fields("QTDE").Value, strFormatacao2Digitos)
      Item.SubItems(3) = "" & Format(TabTemp!Valor_Item, strFormatacao2Digitos)
      Item.SubItems(4) = "" & Format(TabTemp!DESCONTO_PRODUTO, strFormatacao2Digitos)
      Item.SubItems(5) = "" & Format((TabTemp!Valor_Item - TabTemp!DESCONTO_PRODUTO) _
                                    * TabTemp.Fields("QTDE").Value, strFormatacao3Digitos)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome from USUARIO "
      SQL = SQL & " where USUARIO_ID = " & TabTemp.Fields("solicitante_ID").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         Item.SubItems(6) = "" & Trim(TabConsulta.Fields("nome").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PRODUTO"
End Sub

Private Sub MOSTRA_TAREFA()
On Error GoTo ERRO_TRATA

   If IsNumeric(txtServico.Text) Then
      txtDescTarefa.Text = ""
      cmbMecanico.Text = ""
      cmbMecanicoAUX.Text = ""
      txtDescontoTarefa.Text = ""
      txtValorTarefa.Text = ""
      txtTotalTarefa.Text = ""

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from vwOS "
      SQL = SQL & " where os_id = " & NUMR_REQ_N
      SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         TAREFA_ID_N = 0 & TabConsulta.Fields("ostarefa_id").Value

         txtDescTarefa.Text = "" & Trim(TabConsulta.Fields("descricao_serviÇo").Value)

         VALOR_ITEM_N = 0 & TabConsulta.Fields("VALOR_servico").Value
         txtValorTarefa.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         VALOR_DESCONTO_N = 0 & TabConsulta.Fields("desconto_servico").Value
         txtDescontoTarefa.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

         txtTotalTarefa.Text = "" & Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

         cmbMecanicoAUX.Text = "" & TabConsulta.Fields("mecanico_id").Value

         If TabTemp.State = 1 Then _
            TabTemp.Close
   
         SQL = "select nome from USUARIO "
         SQL = SQL & " where USUARIO_ID = " & TabConsulta.Fields("mecanico_id").Value
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
On Error GoTo ERRO_TRATA

   If Trim(txtPRODUTO.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PRODUTO_ID_N = 0

      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(txtPRODUTO.Text) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDESCPRODUTO.Text = "" & Trim(TabConsulta.Fields("DESCRICAO").Value)
         txtValorProduto.Text = "" & Format(TabConsulta.Fields("PRECO_VENDA").Value, strFormatacao2Digitos)
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
         Else
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            MsgBox "Produto não cadastrado."
            txtPRODUTO.SetFocus
            Exit Sub
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from vwOS "
      SQL = SQL & " where os_id = " & NUMR_REQ_N
      SQL = SQL & " and OSPECA_ID = " & PRODUTO_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         PRODUTO_ID_N = 0 & TabConsulta.Fields("PRODUTO_id").Value

         txtDESCPRODUTO.Text = "" & Trim(TabConsulta.Fields("desc_PRODUTO").Value)

         VALOR_ITEM_N = 0 & TabConsulta.Fields("VALOR_ITEM").Value
         txtValorProduto.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         VALOR_DESCONTO_N = 0 & TabConsulta.Fields("desconto_produto").Value
         txtDESCONTOPRODUTO.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

         QTDE_PEDIDO = 0 & TabConsulta.Fields("qtde").Value
         txtQtde.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)

         txtTOTALPRODUTO.Text = "" & Format((VALOR_ITEM_N - VALOR_DESCONTO_N) * QTDE_PEDIDO, strFormatacao3Digitos)

         cmbVendedorAUX.Text = "" & TabConsulta.Fields("VENDEDOR_id").Value

         If TabTemp.State = 1 Then _
            TabTemp.Close
   
         SQL = "select nome from USUARIO "
         SQL = SQL & " where USUARIO_ID = " & TabConsulta.Fields("VENDEDOR_id").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            cmbVENDEDOR.Text = "" & Trim(TabTemp.Fields("nome").Value)

         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub TOTALIZA_CAMPOS()
On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral desconto serviço
   SQL = "select sum(desconto_servico) from OSSERVICO "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotDescontoServico.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral serviço
   SQL = "select sum(valor_servico) from OSSERVICO "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
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
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotDescontoProduto.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral produto
   SQL = "select sum(valor_item*QTDE) from OSPECA "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
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

Sub FECHA_OS()
On Error GoTo ERRO_TRATA

   Dim TabOS               As New ADODB.Recordset
   Dim TabPedido           As New ADODB.Recordset
   Dim TIPO_REGISTRO_A     As String
   Dim STATUS_N            As Integer
   Dim DESCONTO_PEÇA_N     As Double
   Dim DESCONTO_SERVIÇO_N  As Double

   VENDEDOR_ID_N = 0
   CLIENTE_ID_N = 0
   PEDIDO_ID_N = NUMR_REQ_N
   txtCNPJCPF.PromptInclude = False
   STATUS_N = 2
   TIPO_REGISTRO_A = "OS"
   VALOR_TOTAL_DESCONTO_N = 0 & DESCONTO_PEÇA_N + DESCONTO_SERVIÇO_N

   If TabOS.State = 1 Then _
      TabOS.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   TabOS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabOS.EOF Then
      If TabPedido.State = 1 Then _
         TabPedido.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where pedido_id = " & NUMR_REQ_N
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
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select vendedor_id from VENDEDOR "
            SQL = SQL & " where nome_vend = 'BALCAO' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               VENDEDOR_ID_N = 0 & TabConsulta.Fields(0).Value
               Else
                  If TabConsulta.State = 1 Then _
                     TabConsulta.Close

                  SQL = "select vendedor_id from VENDEDOR "
                  SQL = SQL & " where nome_vend = 'BALCÃO' "
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

            SQL = "INSERT INTO PEDIDO "
               SQL = SQL & "(PEDIDO_ID,Empresa_id, numr_req, CGCCPF, Vendedor_id, Dt_Req, "
               SQL = SQL & " Nome_Cliente, Status, Tipo_Registro,Codg_USU, TIPOVENDA_ID, "
               SQL = SQL & " CLIENTE_ID, Valor_ToTal, valor_desconto,perc_desc) "
            SQL = SQL & " VALUES ("
               SQL = SQL & PEDIDO_ID_N
               SQL = SQL & "," & EMPRESA_ID_N
               SQL = SQL & "," & PEDIDO_ID_N
               SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
               SQL = SQL & "," & VENDEDOR_ID_N & ","
               SQL = SQL & "'" & DMA(Date) & "'"
               SQL = SQL & ",'" & Trim(txtCliente.Text) & "'"
               SQL = SQL & "," & STATUS_N
               SQL = SQL & ",'" & TIPO_REGISTRO_A & "'"
               SQL = SQL & "," & CODG_USU_N
               SQL = SQL & "," & 9999
               SQL = SQL & "," & CLIENTE_ID_N
               SQL = SQL & "," & tpMOEDA(txtTotOS.Text)
               SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
               SQL = SQL & "," & tpMOEDA(0)
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close
'======================================
      'PRODUTOS ORDEM DE SERVIÇO
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "SELECT OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.PRECO_CUSTO"
      SQL = SQL & " FROM OS "
      SQL = SQL & " INNER JOIN OSPECA "
      SQL = SQL & " ON OS.OS_ID = OSPECA.OS_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where OSPECA.os_id = " & NUMR_REQ_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         SEQ_ID_N = MAX_ID("seq_id", "PEDIDOITEM", "", "", "", "")
         QTDE_PEDIDO = 0 & TabTemp.Fields("qtde").Value

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("os_id").Value
         SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabConsulta.EOF Then
            SQL = "INSERT INTO PEDIDOITEM "
            SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Numr_req, Codg_Prod, Qtd_Pedida, "
            SQL = SQL & " Valor_item, valor_desconto, status,preco_custo,TIPO_REG) "
            SQL = SQL & " VALUES ("
               SQL = SQL & TabTemp.Fields("os_id").Value                            'PEDIDO_id
               SQL = SQL & "," & SEQ_ID_N                                           'SEQ_ID
               SQL = SQL & "," & TabTemp.Fields("produto_id").Value                 'produto_id
               SQL = SQL & "," & TabTemp.Fields("os_id").Value                      'Numr_req
               SQL = SQL & ",'" & Trim(TabTemp.Fields("codg_produto").Value) & "'"  'Codg_Prod
               SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                               'Qtd_Pedida
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_item").Value)        'Valor_item
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("desconto_produto").Value)  'Valor_desconto
               SQL = SQL & ", 'P'"                                                  'status
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("preco_custo").Value)       'PRECO_CUSTO
               SQL = SQL & ", 'PC'"                                                 'TIPO_REG
            SQL = SQL & ")"
            Else
               SQL = "UPDATE PEDIDOITEM SET "
                  SQL = SQL & " Qtd_Pedida = " & tpMOEDA(QTDE_PEDIDO)                                    'Qtd_Pedida
                  SQL = SQL & ", Valor_item = " & tpMOEDA(TabTemp.Fields("valor_item").Value)            'Valor_item
                  SQL = SQL & ", Valor_desconto = " & tpMOEDA(TabTemp.Fields("desconto_produto").Value)  'Valor_desconto
                  SQL = SQL & ", status = 'P'"                                                           'status
                  SQL = SQL & ", PRECO_CUSTO = " & tpMOEDA(TabTemp.Fields("preco_custo").Value)          'PRECO_CUSTO
                  SQL = SQL & ", TIPO_REG = 'PC'"                                                        'TIPO_REG
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("os_id").Value
               SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
               SQL = SQL & " and seq_id = " & SEQ_ID_N
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         CONECTA_RETAGUARDA.Execute SQL

         'Atualiza Qt Balcao
         SQL = "UPDATE Produto SET "
         SQL = SQL & " qtde_retido = qtde_retido + " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
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
      SQL = SQL & " where descricao = 'SERVIÇO'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         PRODUTO_ID_N = TabTemp.Fields(0).Value
         Else
            MsgBox "No cadastro de produto deve conter registro 'SERVIÇO' "
            Exit Sub
      End If

      SQL = "delete from PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabOS.Fields("os_id").Value
      SQL = SQL & " and tipo_reg = 'OS'"
      CONECTA_RETAGUARDA.Execute SQL

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "SELECT OSSERVICO.* FROM OS "
      SQL = SQL & " INNER JOIN OSSERVICO "
      SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID"
      SQL = SQL & " where OSSERVICO.os_id = " & NUMR_REQ_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         SEQ_ID_N = MAX_ID("seq_id", "PEDIDOITEM", "", "", "", "")
         QTDE_PEDIDO = 1

         SQL = "INSERT INTO PEDIDOITEM "
         SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Numr_req, Codg_Prod, Qtd_Pedida, "
         SQL = SQL & " Valor_item, valor_desconto, status,preco_custo,TIPO_REG) "
         SQL = SQL & " VALUES ("
            SQL = SQL & TabTemp.Fields("os_id").Value                            'PEDIDO_id
            SQL = SQL & "," & SEQ_ID_N                                           'SEQ_ID
            SQL = SQL & "," & PRODUTO_ID_N                                       'produto_id
            SQL = SQL & "," & TabTemp.Fields("os_id").Value                      'Numr_req
            SQL = SQL & "," & PRODUTO_ID_N                                       'Codg_Prod
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

   SQL = "update OS set "
   SQL = SQL & " dt_fecha = '" & DMA(Date) & "'"
   SQL = SQL & ", situacao_os = 9"
   SQL = SQL & " where os_id = " & NUMR_REQ_N
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_OS"
End Sub

Sub HABILITA_TELA()
   If SINAL_INDICADOR_N = 2 Then
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
   Toolbar1.Buttons(5).Enabled = False
   Toolbar1.Buttons(9).Enabled = False
End Sub

Public Function CHECA_DADOS_OS() As Boolean
On Error GoTo ERRO_TRATA

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
   If Trim(txtPLACA.Text) = "" Then
      MsgBox "Veículo não informado, Ordem de Serviço inválida."
      txtPLACA.SetFocus
      Exit Function
   End If
   If EQUIPAMENTO_ID_N <= 0 Then
      MsgBox "Veículo não informado, Ordem de Serviço inválida."
      txtPLACA.SetFocus
      Exit Function
   End If
   If Trim(txtKM.Text) = "" Then
      MsgBox "KM não informado, Ordem de Serviço inválida."
      txtKM.SetFocus
      Exit Function
   End If
   If Not IsNumeric(txtKM.Text) Then
      MsgBox "KM inválido."
      txtKM.SetFocus
      Exit Function
   End If
   CHECA_DADOS_OS = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_DADOS_OS"
End Function
