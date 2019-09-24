VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSSERVI�O 
   Caption         =   "O.S. (frmOSSERVI�O)"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "OSSERVI�O.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   11865
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
      Caption         =   "Produtos Ordem de Servi�o"
      ForeColor       =   &H00008000&
      Height          =   2535
      Left            =   50
      TabIndex        =   52
      Top             =   5400
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
         Picture         =   "OSSERVI�O.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Consulta Cadastro Ve�culo"
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
         TabIndex        =   17
         ToolTipText     =   "Informe C�digo Produto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtValorProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7320
         TabIndex        =   21
         ToolTipText     =   "Valor Venda Produto"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDescontoProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4200
         TabIndex        =   20
         ToolTipText     =   "Desconto Produto"
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbVendedor 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10080
         TabIndex        =   18
         ToolTipText     =   "Respons�vel Venda"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   960
         MaxLength       =   6
         TabIndex        =   19
         ToolTipText     =   "Digite a Quantidade"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtTotalProduto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   10320
         TabIndex        =   55
         Top             =   840
         Width           =   1335
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
         Picture         =   "OSSERVI�O.frx":7D3C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Pesquisa Ve�culo"
         Top             =   360
         Width           =   405
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descri��o"
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
            Text            =   "id"
            Object.Width           =   2
         EndProperty
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
         Left            =   6000
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
         Left            =   2940
         TabIndex        =   61
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Respons�vel:"
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
         Left            =   240
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
         Left            =   9000
         TabIndex        =   58
         Top             =   840
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Servi�os Ordem de Servi�o"
      ForeColor       =   &H00C00000&
      Height          =   2535
      Left            =   50
      TabIndex        =   42
      Top             =   2880
      Width           =   11775
      Begin VB.CommandButton cmdCadServi�o 
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
         Picture         =   "OSSERVI�O.frx":873E
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Consulta Cadastro Ve�culo"
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
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbMecanico 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9600
         TabIndex        =   13
         ToolTipText     =   "Selecione Mecanico para tarefa"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtDescontoTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         ToolTipText     =   "Desconto Servi�o"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtValorTarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7080
         TabIndex        =   16
         ToolTipText     =   "Informe Valor Servi�o"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDescTarefa 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         ToolTipText     =   "Descri��o Servi�o"
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txtServico 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   11
         ToolTipText     =   "Digite C�digo Tarefa ou 0 para Diversar"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cmbMecanicoAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   9120
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   495
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
         Picture         =   "OSSERVI�O.frx":A868
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Pesquisa Ve�culo"
         Top             =   360
         Width           =   405
      End
      Begin MSComctlLib.ListView lstServi�o 
         Height          =   1125
         Left            =   45
         TabIndex        =   46
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Servi�o"
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
            Text            =   "DtEncerramento"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDtFecha 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Fechamento:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   135
         TabIndex        =   82
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Servi�o ="
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8160
         TabIndex        =   51
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T�cnico:"
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
         Left            =   2880
         TabIndex        =   49
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Servi�o ="
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5400
         TabIndex        =   48
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label lblTarefa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Servi�o:"
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
      TabIndex        =   33
      Top             =   1560
      Width           =   11775
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
         Picture         =   "OSSERVI�O.frx":B26A
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Consulta Cadastro Ve�culo"
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
         Left            =   6500
         Picture         =   "OSSERVI�O.frx":D394
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Pesquisa Ve�culo"
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
         Width           =   3495
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   9720
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtMODELO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   7440
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
         Left            =   6600
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
         Left            =   3240
         Picture         =   "OSSERVI�O.frx":DD96
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Consulta Cadastro Ve�culo"
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   7440
         TabIndex        =   35
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
         Left            =   2750
         Picture         =   "OSSERVI�O.frx":FEC0
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Pesquisa Ve�culo"
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
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descri��o:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   405
         TabIndex        =   41
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
         Left            =   8520
         TabIndex        =   40
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ano/Modelo:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   5295
         TabIndex        =   39
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identifica��o:"
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   165
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem de Servi�o"
      Height          =   975
      Left            =   50
      TabIndex        =   23
      Top             =   600
      Width           =   11775
      Begin VB.ComboBox cmbSituacao 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10080
         TabIndex        =   4
         ToolTipText     =   "Selecione situa��o Ordem de Servi�o"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbTipoOS 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6720
         TabIndex        =   3
         ToolTipText     =   "Selecione Tipo Ordem de Servi�o"
         Top             =   480
         Width           =   3135
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
         Width           =   3015
      End
      Begin VB.CommandButton cmdOS 
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
         Left            =   1500
         Picture         =   "OSSERVI�O.frx":108C2
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Pesquisa Ve�culo"
         Top             =   480
         Width           =   405
      End
      Begin VB.ComboBox cmbConsultorAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   3600
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbTipoOSAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6720
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbSituacaoAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   10080
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   615
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situa��o"
         Height          =   240
         Left            =   10080
         TabIndex        =   32
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo O.S."
         Height          =   240
         Left            =   6720
         TabIndex        =   31
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblOs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� O.S."
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblCt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consultor"
         Height          =   240
         Left            =   3600
         TabIndex        =   29
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   240
         Left            =   2040
         TabIndex        =   28
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":112C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":11718
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":11A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":11E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":122DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":125FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":12A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSSERVI�O.frx":12D70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Object.ToolTipText     =   "Limpar formul�rio"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "cons"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetiva��o da comiss�o"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impress�o"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cliente"
            Key             =   "cli"
            ImageIndex      =   8
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
      Caption         =   "Total Desc. Servi�o = "
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
      Caption         =   "SubTotal Servi�o = "
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
      Caption         =   "Total Servi�o = "
      Height          =   240
      Left            =   6585
      TabIndex        =   71
      Top             =   8160
      Width           =   1500
   End
End
Attribute VB_Name = "frmOSSERVI�O"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim EQUIPAMENTO_ID_N       As Long
   Dim TAREFA_ID_N            As Long
   Dim Situa��o_Ordem_Servi�o As String
   Dim DT_FECHAMENTO_OS       As Date

Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   LIMPA_OS
   CARREGA_COMBOS

   If SINAL_INDICADOR_N = 2 Then _
      Me.Caption = "Fechamento Ordem de Servi�o"

End Sub

Private Sub Form_Unload(Cancel As Integer)
   MOSTRA_RODAPE "", "", "", "", ""
End Sub

Private Sub lstProduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Trim(txtOS) <> "" Then
            If IsNumeric(txtOS) Then
               If Trim(lstServi�o.SelectedItem.Text) <> "" Then
                  If IsNumeric(lstServi�o.SelectedItem.Text) Then
                     EXCLUIR_PRODUTO_ITEM
                  End If
               End If
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_KeyDown"
End Sub

Private Sub lstServi�o_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Trim(txtOS) <> "" Then
            If IsNumeric(txtOS) Then
               If Trim(lstServi�o.SelectedItem.Text) <> "" Then
                  If IsNumeric(lstServi�o.SelectedItem.Text) Then
                     EXCLUIR_SERVI�O_ITEM
                  End If
               End If
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstServi�o_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         EXCLUIR_OS
      Case "cons"
         SQL3 = ""
         frmOSCONSULTA.Show 1
         If SQL3 <> "" Then _
            If IsNumeric(SQL3) Then _
                  txtOS.Text = SQL3
         SQL3 = ""
         txtOS.SetFocus
      Case "sair"
        Unload Me
      Case "limpar"
         LIMPA_OS
         txtOS.SetFocus
      Case "gravar"
         If CHECA_DADOS_OS = True Then
            GRAVA_OS
            LIMPA_OS
            txtOS.SetFocus
         End If
      Case "imprimir"
         If Trim(txtOS.Text) <> "" Then _
            If IsNumeric(txtOS.Text) Then _
               IMPRIMIR_ORDEM_SERVI�O txtOS.Text, "SERVI�O"
      Case "excluir"
      Case "cli"
         frmCADASTROCLIENTE.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdCadProduto_Click()
   frmCADASTROPRODUTO.Show 1
End Sub

Private Sub cmdCadCli_Click()
   frmCADASTROCLIENTE.Show 1
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmdCadServi�o_Click()
   frmOSSERVICOCADASTRO.Show 1
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Text = ""
   frmDISPLAYCLIENTE.Show 1
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

Private Sub cmbConsultor_LostFocus()
   If Trim(cmbConsultor.Text) = "" Then _
      cmbConsultor.ListIndex = 0
End Sub

Private Sub cmbTipoOS_LostFocus()
   If Trim(cmbTipoOS.Text) = "" Then _
      cmbTipoOS.ListIndex = 0
End Sub

Private Sub cmbSituacao_LostFocus()
   If Trim(cmbSituacao.Text) = "" Then _
      cmbSituacao.ListIndex = 0
End Sub

Private Sub cmbMecanico_LostFocus()
   If Trim(cmbMecanico.Text) = "" Then _
      cmbMecanico.ListIndex = 0
End Sub

Private Sub cmbVendedor_LostFocus()
   If Trim(cmbVendedor.Text) = "" Then _
      cmbVendedor.ListIndex = 0
End Sub

Private Sub cmbSituacao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

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
   SINAL_INDICADOR_N = cmbSituacaoAUX.Text

Err.Clear
End Sub

'==================cgccpf
Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC-SAIR", "F7-Consulta Clientes", "Inform CNPJ/CPF Cliente e Tecle <<Enter>>", "", ""
   txtCliente.Enabled = True
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then _
      txtCNPJCPF.Mask = "##############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
         frmDISPLAYCLIENTE.Show 1
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

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         txtCliente.Enabled = True
         txtCliente.SetFocus
         Else
            txtServico.SetFocus
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

Private Sub TXTCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Text = "99999999999"
      Else
         If Trim(txtCNPJCPF.Text) <> "99999999999" Then _
            TRATA_CLIENTE
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtDescontoProduto_GotFocus()
   txtDescontoProduto.SelStart = 0
   txtDescontoProduto.SelLength = Len(txtDescontoProduto.Text)
End Sub

Private Sub txtDescontoTarefa_GotFocus()
   txtDescontoTarefa.SelStart = 0
   txtDescontoTarefa.SelLength = Len(txtDescontoTarefa.Text)
End Sub

Private Sub txtDescTarefa_GotFocus()
   txtDescTarefa.SelStart = 0
   txtDescTarefa.SelLength = Len(txtDescTarefa.Text)
End Sub

Private Sub txtdtfecha_GotFocus()
   txtDtFecha.PromptInclude = True
End Sub

Private Sub txtDtFecha_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

End Sub

Private Sub txtdtfecha_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtFecha.PromptInclude = True
   If Not IsDate(txtDtFecha.Text) Then
      txtDtFecha.PromptInclude = False
         txtDtFecha.Text = ""
      txtDtFecha.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtfecha_LostFocus"
End Sub

Private Sub txtDtOS_GotFocus()
   SendKeys ("{tab}")
End Sub

Private Sub txtOs_GotFocus()
   txtOS.SelStart = 0
   txtOS.SelLength = Len(txtOS.Text)
End Sub

Private Sub txtOS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

   INDR_PRI = False
   If Trim(txtOS.Text) <> "" Then
      If IsNumeric(txtOS.Text) Then
         PEDIDO_ID_N = txtOS.Text

         MOSTRA_OS

         Exit Sub
      End If
   End If

   If INDR_PRI = False Then
      GERA_NUMR_REQ
      txtOS.Text = PEDIDO_ID_N
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOS_KeyPress"
End Sub

Private Sub cmdos_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSCONSULTA.Show 1
   If SQL3 <> "" Then _
      txtOS.Text = SQL3
   SQL3 = ""
   txtOS.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdos_Click"
End Sub

Private Sub cmdConsultaPlaca_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSEqpCONSULTA.Show 1
   If SQL3 <> "" Then _
      txtEqp.Text = SQL3
   SQL3 = ""
   txtEqp.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsultaPlaca_Click"
End Sub

Private Sub cmdCadPlaca_Click()
   frmOSEQPCADASTRO.Show 1
End Sub

Private Sub txtEqp_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSVEICULOCONSULTA.Show 1
         If Trim(SQL3) <> "" Then
            txtEqp.Text = SQL3
            MOSTRA_EQP
         End If
         SQL3 = ""
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
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txtQTDE_GotFocus()
   txtQtde.SelStart = 0
   txtQtde.SelLength = Len(txtQtde.Text)
End Sub

Private Sub cmdSERVIcO_Click()
'On Error GoTo ERRO_TRATA

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

Private Sub txtServico_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
      Case vbKeyF7
         SQL3 = ""
         frmOSSERVICOCONSULTA.Show 1
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
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      VALOR_DESCONTO_N = 0 & txtDescontoTarefa.Text

      SendKeys ("{tab}")
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
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then

      If CHECA_DADOS_OS = True Then _
         GRAVA_SERVI�O

      KeyAscii = 0
      'SendKeys ("{tab}")
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

   If INDR_PRI = True Then _
      txtServico.Text = MAX_ID("osservico_id", "OSSERVICO", "OS_ID", txtOS.Text, "", "")

   MOSTRA_TAREFA

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

      SendKeys ("{tab}")
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
'On Error GoTo ERRO_TRATA

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
   If Trim(txtQtde.Text) = "" Then _
      txtQtde.Text = 1

   txtQtde.Text = Format(txtQtde.Text, strFormatacao3Digitos)
End Sub

Private Sub txtDescontoProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
   If txtDescontoProduto.Text = "" Then _
      txtDescontoProduto.Text = 0

   txtDescontoProduto.Text = Format(txtDescontoProduto.Text, strFormatacao2Digitos)
End Sub

Private Sub txtValorProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then

      If CHECA_DADOS_OS = True Then _
         GRAVA_PRODUTO

      KeyAscii = 0
      'SendKeys ("{tab}")
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
   If txtValorProduto.Text = "" Then _
      txtValorProduto.Text = 0

   txtValorProduto.Text = Format(txtValorProduto.Text, strFormatacao2Digitos)

   TOTALIZA_CAMPOS
End Sub

Sub LIMPA_OS()
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(5).Enabled = True
   Toolbar1.Buttons(9).Enabled = True

   SINAL_INDICADOR_N = 0
   VENDEDOR_ID_N = 0
   PEDIDO_ID_N = 0
   CLIENTE_ID_N = 0
   txtCNPJCPF.PromptInclude = False

   Situa��o_Ordem_Servi�o = ""
   txtTotDescontoServico.Text = Format(0, strFormatacao2Digitos)
   txtTotServico.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralServico.Text = Format(0, strFormatacao2Digitos)
   txtTotDescontoProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotOS.Text = Format(0, strFormatacao2Digitos)
   txtDesc.Text = ""

   lstServi�o.ListItems.Clear
   lstProduto.ListItems.Clear

   EQUIPAMENTO_ID_N = 0
   txtOS.Text = ""
   txtDtOS.PromptInclude = False
   txtDtOS.Text = Date
   txtDtOS.PromptInclude = True
   cmbConsultor.Text = ""
   cmbConsultorAUX.Text = ""
   cmbTipoOS.Text = ""
   cmbTipoOSAUX.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   txtEqp.Text = ""
   txtCNPJCPF.Text = ""
   txtCliente.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   txtMarca.Text = ""
   txtTotOS.Text = ""

   LIMPA_PRODUTO
   LIMPA_SERVI�O
   CARREGA_COMBOS
   HABILITA_TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_OS"
End Sub

Sub LIMPA_OS_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(5).Enabled = True
   Toolbar1.Buttons(9).Enabled = True

   SINAL_INDICADOR_N = 0
   VENDEDOR_ID_N = 0
   CLIENTE_ID_N = 0
   txtCNPJCPF.PromptInclude = False

   Situa��o_Ordem_Servi�o = ""
   txtTotDescontoServico.Text = Format(0, strFormatacao2Digitos)
   txtTotServico.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralServico.Text = Format(0, strFormatacao2Digitos)
   txtTotDescontoProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotGeralProduto.Text = Format(0, strFormatacao2Digitos)
   txtTotOS.Text = Format(0, strFormatacao2Digitos)
   txtDesc.Text = ""

   lstServi�o.ListItems.Clear
   lstProduto.ListItems.Clear

   EQUIPAMENTO_ID_N = 0
   txtOS.Text = ""
   txtDtOS.PromptInclude = False
   txtDtOS.Text = Date
   txtDtOS.PromptInclude = True
   cmbConsultor.Text = ""
   cmbConsultorAUX.Text = ""
   cmbTipoOS.Text = ""
   cmbTipoOSAUX.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   txtEqp.Text = ""
   txtCNPJCPF.Text = ""
   txtCliente.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   txtMarca.Text = ""
   txtTotOS.Text = ""

   LIMPA_PRODUTO
   LIMPA_SERVI�O
   CARREGA_COMBOS
   HABILITA_TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_OS_QUASE_TUDO"
End Sub

Sub LIMPA_SERVI�O()
   txtDtFecha.PromptInclude = False
   txtDtFecha.Text = ""
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
   txtProduto.Text = ""
   txtDescProduto.Text = ""
   cmbVendedorAUX.Text = ""
   cmbVendedor.Text = ""
   txtQtde.Text = ""
   txtDescontoProduto.Text = ""
   txtValorProduto.Text = ""
   txtTotalProduto.Text = ""
End Sub

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

'parametros combos x tabela descr
'8 = consultor tecnico
'9 = mecanico

   cmbConsultorAUX.Clear
   cmbConsultor.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 8 or tipo = 5 "   'consultor tecnico
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

   SQL = "select vendedor_id, nome_vend from VENDEDOR "
   SQL = SQL & " where status = 'A' "   'vendedor
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbVendedorAUX.AddItem TabDESCR.Fields("vendedor_id").Value
      cmbVendedor.AddItem Trim(TabDESCR.Fields("nome_vend").Value) & "-" & Trim(TabDESCR.Fields("vendedor_id").Value)

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
      cmbTipoOSAUX.AddItem TabDESCR!codigo
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
'On Error GoTo ERRO_TRATA

   LIMPA_OS_QUASE_TUDO

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS_Servico "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value

      txtCNPJCPF.PromptInclude = False
      'txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("CNPJCPF").Value)
      txtCliente.Text = "" & Trim(TabTemp.Fields("cliente").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select descricao,cnpjcpf from PESSOA "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtCNPJCPF.Text = "" & Trim(TabConsulta.Fields("cnpjcpf").Value)
         If Trim(txtCliente.Text) = "" Then _
            txtCliente.Text = "" & Trim(TabConsulta.Fields("descricao").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      txtCNPJCPF.PromptInclude = True

      EQUIPAMENTO_ID_N = 0 & TabTemp.Fields("EQUIPAMENTO_ID").Value
      txtOS.Text = PEDIDO_ID_N
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

      'txtDesc.Text = "" & Trim(TabTemp.Fields("descricao").Value)
      txtDesc.Text = "" & Trim(TabTemp.Fields("Nome_Equipamento").Value)

      cmbSituacao.Text = "" & TRAZ_DESCRITOR("Z", TabTemp.Fields("SITUACAO_os").Value)
      cmbSituacaoAUX.Text = "" & TabTemp.Fields("SITUACAO_os").Value

      Situa��o_Ordem_Servi�o = TabTemp.Fields("situacao_os").Value
      SINAL_INDICADOR_N = 0 & TabTemp.Fields("situacao_os").Value

      cmbTipoOS.Text = "" & TRAZ_DESCRITOR("H", TabTemp.Fields("tipo_os").Value)
      cmbTipoOSAUX.Text = "" & Trim(TabTemp.Fields("tipo_os").Value)

      txtEqp.Text = "" & Trim(TabTemp.Fields("equipamento_id").Value)
      txtANO.Text = "" & Trim(TabTemp.Fields("ano").Value)
      txtMODELO.Text = "" & Trim(TabTemp.Fields("modelo").Value)

      txtMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)
      'txtCombustivel.Text = "" & TRAZ_DESCRITOR("U", TabTemp.Fields("combustivel_id").Value)

      txtTotOS.Text = ""
      DT_FECHAMENTO_OS = 0 & TabTemp.Fields("dt_fecha").Value
      Else
         INDR_PRI = True
         txtOS.Text = PEDIDO_ID_N
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SETA_GRID_SERVI�O
   SETA_GRID_PRODUTO
   TOTALIZA_CAMPOS

   If Situa��o_Ordem_Servi�o = "4" Then
      DESAABILITA_TELA
      MsgBox "Ordem de Servi�o CANCELADA, permitido somente consulta."
   End If
   If Situa��o_Ordem_Servi�o = "2" Then
      DESAABILITA_TELA
      MsgBox "Ordem de Servi�o FECHADA, permitido somente consulta."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OS"
End Sub

Private Sub MOSTRA_EQP()
'On Error GoTo ERRO_TRATA

   If Trim(txtEqp.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "SELECT * from vwRel_EQUIPAMENTO "
      SQL = SQL & " where equipamento_id = " & txtEqp.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         EQUIPAMENTO_ID_N = TabTemp.Fields("EQUIPAMENTO_ID").Value

         txtCNPJCPF.PromptInclude = False
            If Trim(txtCNPJCPF.Text) = "" Then _
               txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         If Trim(txtCliente.Text) = "" Then _
            txtCliente.Text = "" & Trim(TabTemp.Fields("nome_cliente").Value)

         PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value

         'txtCHASSI.Text = "" & Trim(TabTemp!chassi)
         'txtDescricao.Text = "" & TabTemp!Descricao
         'txtMotor.Text = "" & TabTemp!motor
         txtDesc.Text = "" & Trim(TabTemp.Fields("descricao").Value)

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

         'If Not IsNull(TabTemp!COMBUSTIVEL_ID) Then _
            If IsNumeric(TabTemp.Fields("combustivel_id").Value) Then _
               txtCombustivel.Text = "" & TRAZ_DESCRITOR("U", TabTemp.Fields("combustivel_id").Value)

         'If Not IsNull(TabTemp!tipo_eqp) Then
         '   If IsNumeric(TabTemp!tipo_eqp) Then
         '      cmbTipo.Text = "" & TRAZ_DESCRITOR("A", TabTemp!tipo_eqp)
         '      cmbTipoAUX.Text = "" & TabTemp!tipo_eqp
         '   End If
         'End If
         Else
            MsgBox "N�o encontrado, verifique."
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

   PEDIDO_ID_N = 0 & txtOS.Text

   'If SINAL_INDICADOR_N <> 2 Then
      If TabCABECA.State = 1 Then _
         TabCABECA.Close
   
      SQL = "select * from vwOS_Servico "
      SQL = SQL & " where os_id = " & PEDIDO_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCABECA.EOF Then
         SQL = "update OS set "
            SQL = SQL & " EQUIPAMENTO_ID = " & EQUIPAMENTO_ID_N                  'EQUIPAMENTO_ID
            SQL = SQL & ", TIPO_OS = " & cmbTipoOSAUX.Text                       'TIPO_OS
            SQL = SQL & ", SITUACAO_OS = " & cmbSituacaoAUX.Text                'SITUACAO_OS
            SQL = SQL & ", KM = 0"                                               'KM
            SQL = SQL & ", CT_ID = " & cmbConsultorAUX.Text                      'CT_ID
            SQL = SQL & ", CLIENTE = '" & Trim(Left(txtCliente.Text, 50)) & "'"  'CLIENTE
            SQL = SQL & ", pessoa_id = " & PESSOA_ID_N                           'pessoa_id
         SQL = SQL & " where os_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         Else
            SQL = "insert into OS "
               SQL = SQL & "(OS_ID,ESTABELECIMENTO_ID,EQUIPAMENTO_ID,DT_OS,TIPO_OS,"
               SQL = SQL & " SITUACAO_OS,KM,CT_ID,NUMR_FROTA,CLIENTE,PESSOA_ID)"
            SQL = SQL & " values ( "
               SQL = SQL & PEDIDO_ID_N                                   'OS_ID
               SQL = SQL & "," & ESTABELECIMENTO_ID_N                   'EMPRESA_ID
               SQL = SQL & "," & EQUIPAMENTO_ID_N                       'EQUIPAMENTO_ID
               SQL = SQL & ",'" & DMA(txtDtOS.Text) & "'"               'DT_OS
               SQL = SQL & "," & cmbTipoOSAUX.Text                      'TIPO_OS
               SQL = SQL & ",0" & cmbSituacaoAUX.Text                   'SITUACAO_OS
               SQL = SQL & ",0"                                         'KM
               SQL = SQL & "," & cmbConsultorAUX.Text                   'CT_ID
               SQL = SQL & ",''"                                        'NUMR_FROTA
               SQL = SQL & ",'" & Trim(Left(txtCliente.Text, 50)) & "'" 'CLIENTE
               SQL = SQL & "," & PESSOA_ID_N                            'PESSOA_ID
            SQL = SQL & "  )"
      End If
      
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      CONECTA_RETAGUARDA.Execute SQL
   'End If

   If SINAL_INDICADOR_N = 2 Then
      Msg = "Deseja gerar financeiro ? "
      PERGUNTA Msg, vbYesNo + 32, "Aten��o !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then _
         GERA_PEDIDO
      MsgBox "Ordem de Servi�o fechada com sucesso."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_OS"
End Sub

Sub GRAVA_SERVI�O()
'On Error GoTo ERRO_TRATA

   Dim DtFecha_D  As Date
   Dim DtOS_D     As Date

   If Not IsNumeric(txtServico.Text) Then
      MsgBox "Servi�o inv�lido."
      txtServico.SetFocus
      Exit Sub
   End If
   If Trim(txtDescTarefa.Text) = "" Then
      MsgBox "Descri��o Servi�o inv�lida."
      txtDescTarefa.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(cmbMecanicoAUX.Text) Then
      MsgBox "Mecanico n�o informado."
      cmbMecanico.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtDescontoTarefa.Text) Then _
      txtDescontoTarefa.Text = 0

   If Trim(txtValorTarefa.Text) = "" Then
      MsgBox "Valor servi�o inv�lido."
      txtValorTarefa.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtValorTarefa.Text) Then
      MsgBox "Valor servi�o inv�lido."
      txtValorTarefa.SetFocus
      Exit Sub
   End If

   DtOS_D = txtDtOS.Text

   DtFecha_D = 0

   txtDtFecha.PromptInclude = True
   If IsDate(txtDtFecha.Text) Then
      DtFecha_D = txtDtFecha.Text
      If DtFecha_D < DtOS_D Then
         MsgBox "Data de fechamento do servi�o menor que data da Ordem de Servi�o, n�o permitido."
         Exit Sub
      End If
   End If
   SQL3 = DtFecha_D

   GRAVA_OS

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from vwOS_Servico "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      SQL = "update OSSERVICO set "
         SQL = SQL & " RESPONSAVEL_ID = " & cmbMecanicoAUX.Text                  'RESPONSAVEL_ID
         SQL = SQL & ", VALOR_SERVICO = " & tpMOEDA(txtValorTarefa.Text)         'VALOR_SERVICO
         SQL = SQL & ", DESCRICAO = '" & Trim(txtDescTarefa.Text) & "'"          'DESCRICAO
         SQL = SQL & ", DESCONTO_SERVICO = " & tpMOEDA(txtDescontoTarefa.Text)   'DESCONTO_SERVICO
         SQL = SQL & ", DT_fecha = '" & DMA(SQL3) & "'"                          'DT_fecha
      SQL = SQL & " where os_id = " & PEDIDO_ID_N
      SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text                       'OSSERVICO_ID
      Else
         SQL = "insert into OSSERVICO "
            SQL = SQL & "(OSSERVICO_ID,OS_ID,OSTAREFA_ID,DT_CAD,RESPONSAVEL_ID,"
            SQL = SQL & "VALOR_SERVICO,DESCRICAO,desconto_servico,dt_fecha) "
         SQL = SQL & " values ( "
            SQL = SQL & MAX_ID("osservico_id", "OSSERVICO", "OS_ID", txtOS.Text, "", "")  'OSSERVICO_ID
            SQL = SQL & "," & PEDIDO_ID_N                                                  'OS_ID
            SQL = SQL & "," & txtServico.Text                                             'OSTAREFA_ID
            SQL = SQL & ",'" & DMA(Date) & "'"                                            'DT_CAD
            SQL = SQL & "," & cmbMecanicoAUX.Text                                         'RESPONSAVEL_ID
            SQL = SQL & "," & tpMOEDA(txtValorTarefa.Text)                                'VALOR_SERVICO
            SQL = SQL & ",'" & Trim(txtDescTarefa.Text) & "'"                             'DESCRICAO
            SQL = SQL & "," & tpMOEDA(txtDescontoTarefa.Text)                             'DESCONTO_SERVICO
            SQL = SQL & ",'" & DMA(SQL3) & "'"                                            'DT_fecha
         SQL = SQL & "  )"
   End If

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID_SERVI�O
   LIMPA_SERVI�O

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_SERVI�O"
End Sub

Sub GRAVA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If PRODUTO_ID_N <= 0 Then
      MsgBox "produto inv�lido."
      txtProduto.SetFocus
      Exit Sub
   End If
   If Trim(txtDescProduto.Text) = "" Then
      MsgBox "Descri��o produto inv�lida."
      txtDescProduto.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(cmbVendedorAUX.Text) Then
      MsgBox "Vendedor n�o informado."
      cmbVendedor.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtDescontoProduto.Text) Then _
      txtDescontoProduto.Text = 0

   If Trim(txtQtde.Text) = "" Then
      MsgBox "Quantidade informada inv�lida."
      txtQtde.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtQtde.Text) Then
      MsgBox "Quantidade informada inv�lida."
      txtQtde.SetFocus
      Exit Sub
   End If
   If Trim(txtValorProduto.Text) = "" Then
      MsgBox "Valor produto inv�lido."
      txtValorProduto.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtValorProduto.Text) Then
      MsgBox "Valor produto inv�lido."
      txtValorProduto.SetFocus
      Exit Sub
   End If

   GRAVA_OS

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from vwOS_Servico "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      SQL = "update OSPECA set "
         SQL = SQL & " DT_CAD = '" & DMA(Date) & "'"                            'DT_CAD
         SQL = SQL & ", SOLICITANTE_ID = " & cmbVendedorAUX.Text                 'SOLICITANTE_ID
         SQL = SQL & ", VALOR_ITEM = " & tpMOEDA(txtValorProduto.Text)           'VALOR_ITEM
         SQL = SQL & ", DESCONTO_PRODUTO = " & tpMOEDA(txtDescontoTarefa.Text)   'DESCONTO_PRODUTO
         SQL = SQL & ", QTDE = " & tpMOEDA(txtQtde.Text)                         'QTDE
      SQL = SQL & " where os_id = " & PEDIDO_ID_N
      SQL = SQL & " and OSPECA_ID = " & TabCABECA.Fields("OSPECA_ID").Value      'OSPECA_ID
      SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N
      Else
         SQL = "insert into OSPECA "
            SQL = SQL & "(OSPECA_ID,OS_ID,PRODUTO_ID,DT_CAD,SOLICITANTE_ID,VALOR_ITEM,DESCONTO_PRODUTO,QTDE) "
         SQL = SQL & " values ( "
            SQL = SQL & MAX_ID("OSPECA_id", "OSPECA", "", "", "", "")         'OSPECA_ID
            SQL = SQL & "," & PEDIDO_ID_N                                      'OS_ID
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

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID_PRODUTO
   LIMPA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO"
End Sub

Sub SETA_GRID_SERVI�O()
'On Error GoTo ERRO_TRATA

   lstServi�o.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSSERVICO "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set Item = lstServi�o.ListItems.Add(, "seq." & TabTemp.Fields("OSSERVICO_ID").Value, TabTemp.Fields("OSSERVICO_ID").Value)

      Item.SubItems(1) = "" & Trim(TabTemp!DESCRICAO)
      Item.SubItems(2) = "" & Format(TabTemp!VALOR_SERVICO, strFormatacao2Digitos)
      Item.SubItems(3) = "" & Format(TabTemp!DESCONTO_SERVICO, strFormatacao2Digitos)
      Item.SubItems(4) = "" & Format(TabTemp!VALOR_SERVICO - TabTemp!DESCONTO_SERVICO, strFormatacao2Digitos)
      Item.SubItems(6) = "" & Trim(TabTemp.Fields("dt_fecha").Value)

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
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_SERVI�O"
End Sub

Sub SETA_GRID_PRODUTO()
'On Error GoTo ERRO_TRATA

   lstProduto.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT OSPECA.OSPECA_ID, OSPECA.OS_ID, OSPECA.PRODUTO_ID, OSPECA.DT_CAD, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, "
   SQL = SQL & " OSPECA.QTDE, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
   SQL = SQL & " FROM OSPECA "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set Item = lstProduto.ListItems.Add(, "seq." & TabTemp.Fields("OSPECA_ID").Value, _
                                                     TabTemp.Fields("CODG_PRODUTO").Value)

      Item.SubItems(1) = "" & Trim(TabTemp!DESCRICAO)
      Item.SubItems(2) = "" & Format(TabTemp.Fields("QTDE").Value, strFormatacao2Digitos)
      Item.SubItems(3) = "" & Format(TabTemp!Valor_Item, strFormatacao2Digitos)
      Item.SubItems(4) = "" & Format(TabTemp!DESCONTO_PRODUTO, strFormatacao2Digitos)
      Item.SubItems(5) = "" & Format((TabTemp!Valor_Item - TabTemp!DESCONTO_PRODUTO) _
                                    * TabTemp.Fields("QTDE").Value, strFormatacao3Digitos)
      Item.SubItems(6) = "" & TabTemp.Fields("OSPECA_ID").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome_vend from VENDEDOR"
      SQL = SQL & " where VENDEDOR_ID = " & TabTemp.Fields("solicitante_ID").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         Item.SubItems(6) = "" & Trim(TabConsulta.Fields("nome_vend").Value)
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

      SQL = "select * from vwOS_Servico "
      SQL = SQL & " where os_id = " & PEDIDO_ID_N
      SQL = SQL & " and OSSERVICO_ID = " & txtServico.Text
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         TAREFA_ID_N = 0 & TabConsulta.Fields("ostarefa_id").Value

         txtDescTarefa.Text = "" & Trim(TabConsulta.Fields("descricao_servi�o").Value)

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

      'se n�o achou na tabela servi�o busca na tabela ostarefa que � a de cadastro de servi�o
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
         txtDescProduto.Text = "" & Trim(TabConsulta.Fields("DESCRICAO").Value)
         txtValorProduto.Text = "" & Format(TabConsulta.Fields("PRECO_VENDA").Value, strFormatacao2Digitos)
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
         Else
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            MsgBox "Produto n�o cadastrado."
            txtProduto.SetFocus
            Exit Sub
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from vwOS_Servico "
      SQL = SQL & " where os_id = " & PEDIDO_ID_N
      SQL = SQL & " and OSPECA_ID = " & PRODUTO_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         PRODUTO_ID_N = 0 & TabConsulta.Fields("PRODUTO_id").Value

         txtDescProduto.Text = "" & Trim(TabConsulta.Fields("desc_PRODUTO").Value)

         VALOR_ITEM_N = 0 & TabConsulta.Fields("VALOR_ITEM").Value
         txtValorProduto.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         VALOR_DESCONTO_N = 0 & TabConsulta.Fields("desconto_produto").Value
         txtDescontoProduto.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

         QTDE_PEDIDO = 0 & TabConsulta.Fields("qtde").Value
         txtQtde.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)

         txtTotalProduto.Text = "" & Format((VALOR_ITEM_N - VALOR_DESCONTO_N) * QTDE_PEDIDO, strFormatacao3Digitos)

         cmbVendedorAUX.Text = "" & TabConsulta.Fields("VENDEDOR_id").Value

         If TabTemp.State = 1 Then _
            TabTemp.Close
   
         SQL = "select nome from USUARIO "
         SQL = SQL & " where USUARIO_ID = " & TabConsulta.Fields("VENDEDOR_id").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            cmbVendedor.Text = "" & Trim(TabTemp.Fields("nome").Value)

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
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral desconto servi�o
   SQL = "select sum(desconto_servico) from OSSERVICO "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotDescontoServico.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral servi�o
   SQL = "select sum(valor_servico) from OSSERVICO "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
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
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotDescontoProduto.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'total geral produto
   SQL = "select sum(valor_item*QTDE) from OSPECA "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtTotProduto.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

'PE�A
   VALOR_ITEM_N = 0 & txtTotProduto.Text
   VALOR_DESCONTO_N = 0 & txtTotDescontoProduto.Text
   txtTotGeralProduto.Text = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

'SERVI�O
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

Sub GERA_PEDIDO()
'On Error GoTo ERRO_TRATA

   Dim TabOS               As New ADODB.Recordset
   Dim TabPedido           As New ADODB.Recordset
   Dim TIPO_REGISTRO_A     As String
   Dim STATUS_N            As Integer
   Dim DESCONTO_PE�A_N     As Double
   Dim DESCONTO_SERVI�O_N  As Double

   VENDEDOR_ID_N = 0
   CLIENTE_ID_N = 0
   txtCNPJCPF.PromptInclude = False
   STATUS_N = 2
   TIPO_REGISTRO_A = "OS"
   VALOR_TOTAL_DESCONTO_N = 0 & DESCONTO_PE�A_N + DESCONTO_SERVI�O_N

   If TabOS.State = 1 Then _
      TabOS.Close

   SQL = "select * from vwOS_Servico "
   SQL = SQL & " where os_id = " & PEDIDO_ID_N
   TabOS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabOS.EOF Then
      If TabPedido.State = 1 Then _
         TabPedido.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         If TabPedido.State = 1 Then _
            TabPedido.Close
         If TabOS.State = 1 Then _
            TabOS.Close

         MsgBox "Pedido j� existente, verificar."
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
                  SQL = SQL & " where nome_vend = 'BALC�O' "
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
               SQL = SQL & "(PEDIDO_ID,Empresa_id, CGCCPF, Vendedor_id, Dt_Req, "
               SQL = SQL & " Nome_Cliente, Status, Tipo_Registro,Codg_USU, TIPOVENDA_ID, "
               SQL = SQL & " CLIENTE_ID, Valor_ToTal, valor_desconto,perc_desc,estabelecimento_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & PEDIDO_ID_N
               SQL = SQL & "," & EMPRESA_ID_N
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
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close
'======================================
      'PRODUTOS ORDEM DE SERVI�O
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "SELECT OSPECA.*,PRODUTO.CODG_PRODUTO,PRODUTO.DESCRICAO,PRODUTO.PRECO_CUSTO,PRODUTO.SITUACAO_TRIBUTARIA"
      SQL = SQL & " FROM OS "
      SQL = SQL & " INNER JOIN OSPECA "
      SQL = SQL & " ON OS.OS_ID = OSPECA.OS_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where OSPECA.os_id = " & PEDIDO_ID_N
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
            SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Codg_Prod, Qtd_Pedida, "
            SQL = SQL & " Valor_item, valor_desconto, status,preco_custo,TIPO_REG) "
            SQL = SQL & " VALUES ("
               SQL = SQL & TabTemp.Fields("os_id").Value                            'PEDIDO_id
               SQL = SQL & "," & SEQ_ID_N                                           'SEQ_ID
               SQL = SQL & "," & TabTemp.Fields("produto_id").Value                 'produto_id
               SQL = SQL & ",'" & Trim(TabTemp.Fields("codg_produto").Value) & "'"  'Codg_Prod
               SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                               'Qtd_Pedida
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_item").Value)        'Valor_item
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("desconto_produto").Value)  'Valor_desconto
               SQL = SQL & ", 'P'"                                                  'status
               SQL = SQL & "," & tpMOEDA(TabTemp.Fields("preco_custo").Value)       'PRECO_CUSTO
               SQL = SQL & ", 'PC'"                                                 'TIPO_REG
               SQL = SQL & ",'" & Trim(TabTemp.Fields("SITUACAO_TRIBUTARIA").Value) 'stributaria
            SQL = SQL & ")"
            Else
               SQL = "UPDATE PEDIDOITEM SET "
                  SQL = SQL & " Qtd_Pedida = " & tpMOEDA(QTDE_PEDIDO)                                    'Qtd_Pedida
                  SQL = SQL & ", Valor_item = " & tpMOEDA(TabTemp.Fields("valor_item").Value)            'Valor_item
                  SQL = SQL & ", Valor_desconto = " & tpMOEDA(TabTemp.Fields("desconto_produto").Value)  'Valor_desconto
                  SQL = SQL & ", status = 'P'"                                                           'status
                  SQL = SQL & ", PRECO_CUSTO = " & tpMOEDA(TabTemp.Fields("preco_custo").Value)          'PRECO_CUSTO
                  SQL = SQL & ", TIPO_REG = 'PC'"                                                        'TIPO_REG
                  SQL = SQL & ", stributaria = '" & Trim(TabTemp.Fields("SITUACAO_TRIBUTARIA").Value) 'stributaria
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
      'SERVI�O ORDEM DE SERVI�O

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

      SQL = "SELECT OSSERVICO.* FROM OS "
      SQL = SQL & " INNER JOIN OSSERVICO "
      SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID"
      SQL = SQL & " where OSSERVICO.os_id = " & PEDIDO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         SEQ_ID_N = MAX_ID("seq_id", "PEDIDOITEM", "", "", "", "")
         QTDE_PEDIDO = 1

         SQL = "INSERT INTO PEDIDOITEM "
         SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Codg_Prod, Qtd_Pedida, "
         SQL = SQL & " Valor_item, valor_desconto, status,preco_custo,TIPO_REG) "
         SQL = SQL & " VALUES ("
            SQL = SQL & TabTemp.Fields("os_id").Value                            'PEDIDO_id
            SQL = SQL & "," & SEQ_ID_N                                           'SEQ_ID
            SQL = SQL & "," & PRODUTO_ID_N                                       'produto_id
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_PEDIDO"
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
'On Error GoTo ERRO_TRATA

   CHECA_DADOS_OS = False

   If Trim(txtOS.Text) = "" Then
      MsgBox "N�mero de Ordem de Servi�o inv�lida."
      txtOS.SetFocus
      Exit Function
   End If
   If Not IsNumeric(txtOS.Text) Then
      MsgBox "N�mero de Ordem de Servi�o inv�lida."
      txtOS.SetFocus
      Exit Function
   End If
   If Not IsDate(txtDtOS.Text) Then
      MsgBox "Data de Ordem de Servi�o inv�lida."
      txtDtOS.SetFocus
      Exit Function
   End If
   If Trim(cmbConsultorAUX.Text) = "" Then
      MsgBox "Consultor inv�lido."
      cmbConsultor.SetFocus
      Exit Function
   End If
   If Not IsNumeric(cmbConsultorAUX.Text) Then
      MsgBox "Consultor inv�lido."
      cmbConsultor.SetFocus
      Exit Function
   End If
   If Trim(cmbTipoOSAUX.Text) = "" Then
      MsgBox "Tipo de Ordem de Servi�o inv�lido."
      cmbTipoOS.SetFocus
      Exit Function
   End If
   If Not IsNumeric(cmbTipoOSAUX.Text) Then
      MsgBox "Tipo de Ordem de Servi�o inv�lido."
      cmbTipoOS.SetFocus
      Exit Function
   End If
   If Trim(cmbSituacao.Text) = "" Then
      MsgBox "Situa��o da Ordem de Servi�o inv�lida."
      cmbSituacao.SetFocus
      Exit Function
   End If
   If Trim(txtEqp.Text) = "" Then
      MsgBox "Ve�culo n�o informado, Ordem de Servi�o inv�lida."
      txtEqp.SetFocus
      Exit Function
   End If
   If EQUIPAMENTO_ID_N <= 0 Then
      MsgBox "Ve�culo n�o informado, Ordem de Servi�o inv�lida."
      txtEqp.SetFocus
      Exit Function
   End If
   
   CHECA_DADOS_OS = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_DADOS_OS"
End Function

Sub TRATA_CLIENTE()
'On Error GoTo ERRO_TRATA

   Dim rstCliente       As New ADODB.Recordset
   Dim rstAux           As New ADODB.Recordset
   Dim rstEndereco      As New ADODB.Recordset
   Dim rstCep           As New ADODB.Recordset

   Dim VALOR_LIMITE_N   As Double
   Dim VALOR_PENDENTE_N As Double

   ENDERECO_A = ""
   PESSOA_ID_N = 0
   If txtCNPJCPF.Text = "" Then
      txtCNPJCPF.Text = "99999999999"
      Else
         txtCNPJCPF.PromptInclude = False
         If CHECA_CNPJCPF(Trim(txtCNPJCPF.Text)) = True Then
            CRITERIO = txtCNPJCPF.Text
            Else
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCNPJCPF.PromptInclude = False
               txtCNPJCPF.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
         End If
   End If

   If Trim(txtCNPJCPF.Text) <> "" Then
      CRITERIO = Trim(txtCNPJCPF.Text)
      If Not IsNull(txtCNPJCPF.Text) Then
         If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
      End If
      txtCNPJCPF.Text = CRITERIO
   End If

   txtCliente.Enabled = True

   If rstCliente.State = 1 Then _
      rstCliente.Close

   SQL = "select pessoa_id,nome,cliente_id,limite_credito,tipo_cliente,cgccpf,ie,Status from CLIENTE "
   SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
   SQL = SQL & " and status = 'A'"
   rstCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstCliente.EOF Then
      If rstCliente.State = 1 Then _
         rstCliente.Close

      Beep
      MsgBox "CPF n�o Cadastrado.", vbOKOnly, "Aten��o."
      txtCNPJCPF.SetFocus
      Exit Sub
      Else
         txtCliente.Enabled = False
         PESSOA_ID_N = rstCliente.Fields("pessoa_id").Value

         If Trim(rstCliente!NOME) <> "" And Trim(txtCNPJCPF.Text) <> "99999999999" Then _
            txtCliente.Text = rstCliente!NOME

         SQL = "update PEDIDO set nome_cliente = '" & Trim(txtCliente.Text) & "'"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         CONECTA_RETAGUARDA.Execute SQL

         CLIENTE_ID_N = rstCliente.Fields("cliente_id").Value
         'If Not IsNull(rstCliente!limite_credito) Then _
            txtLIMITE.Text = Format(rstCliente!limite_credito, strFormatacao2Digitos)

         'Pegou o tipo do cliente
         'If Not IsNull(rstCliente!TIPO_CLIENTE) Then _
            dblTipoCliente = rstCliente!TIPO_CLIENTE

         'If Not IsNull(rstCliente!CGCCPF) Then _
            strCPFCNPJ = rstCliente!CGCCPF

         'If Not IsNull(rstCliente!IE) Then 'O Cara ja tem no Cadastro de Cliente
         '   strInscEstadual = rstCliente!IE
         '   Else ' Se ele nao tiver no Cadastro de Cliente pega aqui!
         '      If rstCliente.State = 1 Then _
         '         rstCliente.Close
         '      MsgBox "Inscri��o estatual invalida para este cliente, atualizar."
         '      Exit Sub
         'End If

         If rstAux.State = 1 Then _
            rstAux.Close

         SQL = "select sum(i.valor_item) from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = l.numr_doc "
         SQL = SQL & " and l.pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and i.status = 'A' "
         SQL = SQL & " and l.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = 1"
         SQL = SQL & " and i.formapagto_id <> 1"
         rstAux.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not rstAux.EOF Then
            If Not IsNull(rstAux.Fields(0).Value) Then
               VALOR_PENDENTE_N = 0 & rstAux.Fields(0).Value
               'txtPAGAR.Text = Format(rstAux.Fields(0).Value, strFormatacao2Digitos)
               'txtPAGAR.Refresh
            End If
         End If
         If rstAux.State = 1 Then _
            rstAux.Close

         VALOR_LIMITE_N = 0 & rstCliente.Fields("LIMITE_CREDITO").Value

         If VALOR_LIMITE_N > 0 Then
            If VALOR_PENDENTE_N >= VALOR_LIMITE_N Then
               MsgBox "Valor limite de credito para esse cliente ultrapassado, n�o permitido venda, verificar com departamento financeiro."
               txtCNPJCPF.Text = ""
               txtCliente.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If

         If rstEndereco.State = 1 Then _
            rstEndereco.Close

         SQL = "select * from ENDERECO "
         SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & " and tipo = 'C'"
         rstEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not rstEndereco.EOF Then
            If Not IsNull(rstEndereco!Rua) Then _
               ENDERECO_A = rstEndereco!Rua
            If Not IsNull(rstEndereco!Complemento) Then _
               ENDERECO_A = ENDERECO_A & "," & rstEndereco!Complemento
            If Not IsNull(rstEndereco!Bairro) Then _
               ENDERECO_A = ENDERECO_A & "," & rstEndereco!Bairro

            'Pegou o CEP do cliente
            If IsNull(rstEndereco!CEP) Then
               If rstEndereco.State = 1 Then _
                  rstEndereco.Close
   
               MsgBox "O Cadastro do cliente n�o est� completo. Verique os dados (CEP, UF, Endere�o, Insc. Estadual etc..)" & vbCrLf & "O sitema n�o pode continuar sem o devido acerto.", vbCritical
               txtCNPJCPF.Text = ""
               txtCliente.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
            If rstCep.State = 1 Then _
               rstCep.Close
      
            'Pegar a uf do cliente
            rstCep.Open "Select * From CEP Where CEP = " & rstEndereco!CEP, CONECTA_RETAGUARDA, , , adCmdText
            If Not rstCep.EOF Then
               If Not IsNull(rstCep!UF) Then
                  'UF_CLIENTE = rstCep!UF
                  Else 'UF nao localizada
                     If rstCep.State = 1 Then _
                        rstCep.Close
                     MsgBox "O Cadastro do cliente n�o est� completo. Verique os dados (CEP, UF, Endere�o, Insc. Estadual etc..)" & vbCrLf & "O sitema n�o pode continuar sem o devido acerto.", vbCritical
                     txtCNPJCPF.Text = ""
                     txtCliente.Text = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
               End If
               Else
                  If rstCep.State = 1 Then _
                     rstCep.Close

                  MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais completos. Verique-os, principalmente o Estado(UF) da empresa"
                  txtCNPJCPF.Text = ""
                  txtCliente.Text = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End If
            If rstCep.State = 1 Then _
               rstCep.Close
         End If
         If rstEndereco.State = 1 Then _
            rstEndereco.Close

         If rstCliente!Status = "C" Then
            If rstCliente.State = 1 Then _
               rstCliente.Close

            Beep
            MsgBox "Cliente Esta Bloqueado!, Verifique Cadastro!.", vbOKOnly, "Aten��o."
            txtCNPJCPF.Text = ""
            txtCliente.Text = ""
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
   End If
   If rstCliente.State = 1 Then _
      rstCliente.Close

   SQL = "select nome_cliente from PEDIDO "
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   rstCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not rstCliente.EOF Then _
      If Not IsNull(rstCliente.Fields(0).Value) Then _
         txtCliente.Text = Trim(rstCliente.Fields(0).Value)
   If rstCliente.State = 1 Then _
      rstCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_CLIENTE"
End Sub

Sub EXCLUIR_SERVI�O_ITEM()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma Exclus�o desse servi�o ?"
   Style = vbYesNo + 32
   Title = "Aten��o."
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then

      SQL = "Delete FROM OSSERVICO "
      SQL = SQL & " Where OSSERVICO_id = " & lstServi�o.SelectedItem.Text
      SQL = SQL & " and os_id = " & txtOS.Text
      CONECTA_RETAGUARDA.Execute SQL

      SETA_GRID_SERVI�O
      TOTALIZA_CAMPOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_SERVI�O_ITEM"
End Sub

Sub EXCLUIR_PRODUTO_ITEM()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma Exclus�o desse produto ?"
   Style = vbYesNo + 32
   Title = "Aten��o."
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then

      SQL = "Delete FROM ospeca "
      SQL = SQL & " Where produto_id = " & lstProduto.SelectedItem.Text
      SQL = SQL & " and os_id = " & txtOS.Text
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

   If Trim(txtOS) <> "" Then
      If IsNumeric(txtOS) Then
         Msg = "Confirma Exclus�o desta ordem de servi�o ?"
         Style = vbYesNo + 32
         Title = "Aten��o."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then

            SQL = "Delete from OSPEDIDO "
            SQL = SQL & " Where os_id = " & txtOS.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OSPECA "
            SQL = SQL & " Where os_id = " & txtOS.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OSSERVICO "
            SQL = SQL & " Where os_id = " & txtOS.Text
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from OS "
            SQL = SQL & " Where os_id = " & txtOS.Text
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_OS
            txtOS.SetFocus
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_OS"
End Sub
