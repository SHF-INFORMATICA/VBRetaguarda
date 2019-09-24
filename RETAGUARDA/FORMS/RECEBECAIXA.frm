VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRECEBECAIXA 
   BackColor       =   &H000000C0&
   Caption         =   "Recebimento Caixa "
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   1845
   ClientWidth     =   10950
   Icon            =   "RECEBECAIXA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   0
      TabIndex        =   12
      Top             =   550
      Width           =   10935
      Begin VB.TextBox txtPercEntrada 
         Alignment       =   2  'Center
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtValorEntrada 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         MaxLength       =   12
         TabIndex        =   0
         Top             =   1800
         Width           =   2040
      End
      Begin VB.ComboBox cmbAuxTIPOVENDA 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   2760
         TabIndex        =   32
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTIPOVENDA 
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
         Left            =   2760
         TabIndex        =   2
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
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
         Left            =   9000
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPercDesconto 
         Alignment       =   2  'Center
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         TabIndex        =   28
         Top             =   1200
         Width           =   2040
      End
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
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
         Height          =   405
         Left            =   9000
         TabIndex        =   26
         Top             =   2160
         Width           =   1800
      End
      Begin VB.TextBox txtVendaComDesconto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   9000
         TabIndex        =   24
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox txtRecebido 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   405
         Left            =   9000
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtLanc 
         Alignment       =   2  'Center
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
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtVendaSemDesconto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   9000
         TabIndex        =   15
         Top             =   720
         Width           =   1800
      End
      Begin VB.TextBox txtCli 
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
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtVendedor 
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
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblPRAZO 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6360
         TabIndex        =   37
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6000
         TabIndex        =   36
         Top             =   1800
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6000
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Entrada = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   34
         Top             =   1800
         Width           =   2340
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo Venda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   33
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Desconto = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   885
         TabIndex        =   29
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Troco = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7305
         TabIndex        =   27
         Top             =   2160
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Venda c/ Desconto = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   25
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Recebido = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6885
         TabIndex        =   23
         Top             =   1680
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Venda = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7245
         TabIndex        =   21
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Dt.Emissão:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2670
         TabIndex        =   18
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   10935
      Begin VB.CommandButton cmdCadProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   4850
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Cadastro Cheque"
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton cmdMata 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   800
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtDias 
         Alignment       =   2  'Center
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
         Left            =   6960
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cmbModalidadeAUX 
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
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtValorItem 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbModalidade 
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
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtDTVENC 
         Height          =   360
         Left            =   9480
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox txtDTEMIS 
         Height          =   360
         Left            =   8040
         TabIndex        =   7
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Vencimento"
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
         Index           =   6
         Left            =   9480
         TabIndex        =   43
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Emissão"
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
         Index           =   5
         Left            =   8040
         TabIndex        =   42
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
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
         Index           =   4
         Left            =   6960
         TabIndex        =   41
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor "
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
         Index           =   3
         Left            =   5280
         TabIndex        =   40
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Modalidade Recebimento "
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
         Index           =   2
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Seq.:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView ListaLanc 
      Height          =   2385
      Left            =   0
      TabIndex        =   11
      Top             =   4350
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   4207
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   14737632
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Doc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modalidade"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Dt.Lanç."
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Dt.Venc."
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Juros"
         Object.Width           =   3528
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10950
      DesignHeight    =   6765
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RECEBECAIXA.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1164
      ButtonWidth     =   1614
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confirmar"
            Key             =   "conf"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmRECEBECAIXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
   Dim VALOR_RECEBIDO_N          As Double
   Dim NUMR_PARCELA              As Integer
   Dim VALOR_TROCO_N             As Double
   Dim VALOR_TOTAL_LANÇADO       As Double
   Dim VALOR_ENTRADA             As Double
   Dim PERC_JUROS_N              As Double
   Dim DIAS_PRAZO                As Integer
   Dim TabTipoVenda              As New ADODB.Recordset
   Dim INDR_FINALIZA_RECEBIMENTO As Boolean
   Dim VALOR_DESCONTO_CABECA_N   As Double
   Dim TOTAL_DESCONTO_N          As Double
   Dim VALOR_DIFERENCA_N         As Double
   Dim VALOR_DESCONTO_ITEM_N     As Double
   Dim VALOR_VENDA_N             As Double
   Dim VALOR_REC_N               As Double

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0

   INDR_FORM_ABERTO = True
   Me.Caption = Me.Caption & " - " & Me.Name & " SINAL_INDICADOR_N = " & SINAL_INDICADOR_N & " nr " & NUMR_REQ_N
      
   VALOR_TOTAL_N = 0
   'Frame1.Enabled = False
   NUMR_PARCELA = 0

   LIMPA_LANCAMENTO

   txtData.Text = Now
   txtLanc.Text = NUMR_REQ_N
   If NUMR_REQ_N > 0 Then
      SETA_GRID
      Else
         MsgBox "Número de lançamento não foi informado. verifique."
         Unload Me
   End If

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * FROM PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome_vend from VENDEDOR v, EQUIPE e "
      SQL = SQL & " where v.vendedor_id = " & TabCABECA!VENDEDOR_ID
      SQL = SQL & " and e.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and v.codg_eq = e.codg_eq "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtVendedor.Text = TabConsulta!NOME_VEND
         txtVendedor.Refresh
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome,pessoa_id from CLIENTE "
      SQL = SQL & " where cgccpf = '" & TabCABECA!CGCCPF & "'"
      SQL = SQL & " and status = 'A'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtCli.Text = Trim(TabConsulta!NOME)
         CNPJCPF_A = TabCABECA!CGCCPF
         PESSOA_ID_N = TabConsulta.Fields("pessoa_id").Value
         txtCli.Refresh
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select nome_cliente from PEDIDO "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            txtCli.Text = Trim(TabConsulta.Fields(0).Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PERC_DESCONTO_N = 0
      VALOR_DESCONTO_N = 0
      VALOR_DESCONTO_CABECA_N = 0
      TOTAL_DESCONTO_N = 0
      VALOR_DESCONTO_ITEM_N = 0

      VALOR_DESCONTO_CABECA_N = 0 & TabCABECA.Fields("valor_desconto").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'Fazendo duas leituras no mesmo banco de dados
      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
      SQL = SQL & " and tipo_reg = 'PC' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_ITEM_N = TabConsulta.Fields(0).Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'BUSCA VALOR TOTAL VENDA
      VALOR_ITEM_N = 0
      SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
      SQL = SQL & " and tipo_reg = 'PC' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_ITEM_N = TabConsulta.Fields(0).Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TOTAL_DESCONTO_N = VALOR_DESCONTO_CABECA_N + VALOR_DESCONTO_ITEM_N

      VALOR_TOTAL_N = VALOR_ITEM_N - TOTAL_DESCONTO_N

      PERC_DESCONTO_N = 0 & (TOTAL_DESCONTO_N / VALOR_ITEM_N) * 100

      txtVendaSemDesconto.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      txtVendaComDesconto.Text = Format((VALOR_ITEM_N - TOTAL_DESCONTO_N), strFormatacao2Digitos)
      txtDesconto.Text = Format(TOTAL_DESCONTO_N, strFormatacao2Digitos)
      txtPercDesconto.Text = Format(PERC_DESCONTO_N, strFormatacao2Digitos)
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   BUSCA_LANCAMENTO

   txtVendaSemDesconto.Refresh

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF2
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KEYDOWN"
End Sub

Private Sub listalanc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView ListaLanc, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         EXCLUIR_TUDO
         SETA_GRID
         LIMPA_BODY
         VALOR_ITEM_N = 0
         VALOR_ENTRADA = 0
         txtValorEntrada.Text = ""
         txtPercEntrada.Text = ""
         cmbTIPOVENDA.Text = ""
         cmbAuxTIPOVENDA.Text = ""
         'Frame1.Enabled = False
         txtTroco.Text = ""
         cmbTIPOVENDA.SetFocus
      Case "conf"
         CONFIRMAR_RECEBIMENTO_PARCELADO
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbTIPOVENDA_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbAuxTIPOVENDA.Text) <> "" Then
      cmbModalidade.Clear
      cmbModalidadeAUX.Clear

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where formapagto_id < 9999 "
      SQL = SQL & " and status = 'true' "
      If Trim(cmbAuxTIPOVENDA.Text) <> "" Then _
         If IsNumeric(cmbAuxTIPOVENDA.Text) Then _
            SQL = SQL & " and formapagto_id >= 1 "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         cmbModalidade.AddItem TabConsulta!Descricao
         cmbModalidadeAUX.AddItem TabConsulta!formapagto_id
         TabConsulta.MoveNext
      Wend

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_LostFocus"
End Sub

Private Sub cmbTipoVenda_Click()
'On Error GoTo ERRO_TRATA

   lblPRAZO.Caption = ""
   cmbAuxTIPOVENDA.ListIndex = cmbTIPOVENDA.ListIndex
   VALOR_ITEM_N = 0
   VALOR_ENTRADA = 0

   EXCLUIR_TUDO

   NUMR_PARCELA = 0
   DIAS_PRAZO = 0

   If cmbAuxTIPOVENDA.Text <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         If Not IsNull(TabConsulta!parcela) Then _
            NUMR_PARCELA = TabConsulta!parcela
         If Not IsNull(TabConsulta!PRAZO) Then _
            DIAS_PRAZO = TabConsulta!PRAZO
      End If
      Else
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         MsgBox "Selecione tipo de venda."
         Exit Sub
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Frame1.Enabled = True
   If DIAS_PRAZO > 0 Then
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      'GERA TITULOS
      SQL = "select * FROM PEDIDO "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCABECA.EOF Then
         GERA_FATURAMENTO
         SETA_GRID
      End If
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      Else
         NUMR_SEQ_N = 1

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select max(i.seq) as ultimo_reg  from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
         SQL = SQL & " and i.numr_doc = l.numr_doc "
         SQL = SQL & " and i.lancamento_id = l.lancamento_id "
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = " & SINAL_INDICADOR_N
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then _
            If Not IsNull(TabLancamento!ultimo_reg) Then _
               NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         txtSeq.Text = NUMR_SEQ_N
         VALOR_ITEM_N = 0 & txtVendaComDesconto.Text
         VALOR_ENTRADA = 0 & txtValorEntrada.Text
         cmbModalidadeAUX.Text = 1
         txtDTVENC.PromptInclude = False
            txtDTVENC.Text = Date
         txtDTVENC.PromptInclude = True

         If Str(txtRecebido.Text) <= 0 Then
            txtRecebido.Text = Format(txtVendaComDesconto.Text, strFormatacao2Digitos)
            txtTroco.Text = Format(0, strFormatacao2Digitos)
         End If

         GRAVAR_TUDO
   End If

   SETA_GRID
   LIMPA_BODY

   CONFIRMAR_RECEBIMENTO_PARCELADO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_Click"
End Sub

Private Sub cmbTIPOVENDA_GotFocus()
'On Error GoTo ERRO_TRATA

   'Frame1.Enabled = False
   cmbTIPOVENDA.Clear
   cmbAuxTIPOVENDA.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA "
   SQL = SQL & " order by tipovenda_id desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbTIPOVENDA.AddItem Trim(TabTemp!Descricao)
      cmbAuxTIPOVENDA.AddItem Trim(TabTemp!TipoVenda_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MOSTRA_RODAPE "ESC - SAIR", "Selecione Tipo Venda", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_GotFocus"
End Sub

Private Sub cmbTIPOVENDA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Frame1.Enabled = True Then _
         txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_KeyPress"
End Sub

Private Sub cmdCadProd_Click()
'On Error GoTo ERRO_TRATA

   cmbModalidadeAUX.ListIndex = cmbModalidade.ListIndex
   
   If Trim(cmbModalidade.Text) <> "" Then
      If Left(UCase(cmbModalidade.Text), 6) = "CHEQUE" Then
         INDR_PRI = True
         frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = False
            frmCHEQUECADASTRO.txtPORTADOR.Text = CNPJCPF_A
         frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = True
         frmCHEQUECADASTRO.Show 1
         INDR_PRI = False
      End If
   End If
   txtValorItem.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCadProd_Click"
End Sub

Private Sub cmbmodalidade_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "Selecione Forma de Pagto.", "ESC - Confirma", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmodalidade_GotFocus"
End Sub

Private Sub cmbModalidade_LostFocus()
   If Trim(cmbModalidade.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where descricao = '" & Trim(cmbModalidade.Text) & "'"
      SQL = SQL & " and status = 'true' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         cmbModalidadeAUX.Text = "" & TabConsulta!formapagto_id

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If
End Sub

Private Sub txtDias_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtDias.Text <> "" Then _
      If IsNumeric(txtDias.Text) Then _
         DIAS_PRAZO = txtDias.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDias_LostFocus"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEmis.PromptInclude = True
   If Not IsDate(txtDtEmis.Text) Then
      txtDtEmis.PromptInclude = False
         txtDtEmis.Text = Date
      txtDtEmis.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_GotFocus"
End Sub

Private Sub txtDTEMIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDTVENC.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMIS_KeyPress"
End Sub

Private Sub txtDTEMIS_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtEmis.PromptInclude = True
   If Not IsDate(txtDtEmis.Text) Then
      txtDtEmis.PromptInclude = False
         txtDtEmis.Text = Date
      txtDtEmis.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMIS_LostFocus"
End Sub

Private Sub txtDTVENC_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDTVENC.PromptInclude = True

   MOSTRA_RODAPE "Informe Data Vencimento da parcela", "ESC - Confirma", "", "", ""

   If DIAS_PRAZO > 0 Then
      NUMR_SEQ_N = 0 & txtSeq.Text
      DATA_INI = txtDtEmis.Text
      txtDTVENC.Text = DATA_INI + DIAS_PRAZO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_GotFocus"
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtDTVENC.PromptInclude = True
      If Not IsDate(txtDTVENC.Text) Then
         txtDTVENC.SetFocus
         txtDTVENC.PromptInclude = False
            txtDTVENC.Text = Date
         txtDTVENC.PromptInclude = True
         Exit Sub
      End If
      If txtSeq.Text = "" Then
         MsgBox "Seqüência deve ser gerada ou informada."
         txtSeq.SetFocus
         Exit Sub
      End If
      If cmbModalidadeAUX.Text = "" Then
         MsgBox "Selecione Forma de Pagamento !!!"
         cmbModalidade.SetFocus
         Exit Sub
      End If
      If txtValorItem.Text = "" Then
         MsgBox "Valor Incorreto !!!"
         txtValorItem.SetFocus
         Exit Sub
      End If
      txtDtEmis.PromptInclude = True
      If Not IsDate(txtDtEmis.Text) Then
         MsgBox "Data de emissão inválida !!!"
         txtDTVENC.SetFocus
         Exit Sub
      End If
      txtDTVENC.PromptInclude = True
      If CDate(txtDTVENC.Text) < CDate(txtDtEmis.Text) Then
         MsgBox "Data de vencimento não pode ser menor que data de emissão !!!"
         txtDTVENC.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      VALOR_ITEM_N = txtValorItem.Text

      GRAVAR_TUDO
      LIMPA_BODY
      SETA_GRID

      VALOR_VENDA_N = 0 & txtVendaComDesconto.Text
      VALOR_REC_N = 0 & txtRecebido.Text

CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA

      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_KeyPress"
End Sub

Private Sub txtLanc_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtVENDEDOR_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtDATA_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtCLI_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtVendaSemDesconto_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtVendaComDesconto_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtRecebido_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txttROCO_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtDesconto_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtPercDesconto_GotFocus()
   cmbTIPOVENDA.SetFocus
End Sub

Private Sub txtValorItem_GotFocus()
'On Error GoTo ERRO_TRATA

   Dim VALOR_TOTAL_VENDA As Double

   MOSTRA_RODAPE "Informe o valor da parcela", "ESC - Confirma", "", "", ""

   If Trim(txtVendaSemDesconto.Text) <> "" Then
      VALOR_ITEM_N = txtVendaSemDesconto.Text
      txtValorItem.Text = Format(VALOR_ITEM_N - VALR_DESCONTO_N, strFormatacao2Digitos)

      If Trim(txtVendaComDesconto.Text) <> "" Then
         VALOR_ITEM_N = txtVendaComDesconto.Text
         txtValorItem.Text = Format(VALOR_ITEM_N - VALR_DESCONTO_N, strFormatacao2Digitos)
      End If
   End If
   VALOR_ITEM_N = 0
   VALR_DESCONTO_N = 0

   BUSCA_LANCAMENTO

VALOR_TOTAL_VENDA = 0 & txtVendaComDesconto.Text
txtValorItem.Text = Format(VALOR_TOTAL_VENDA - VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
txtValorItem.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_GotFocus"
End Sub

Private Sub txtValorItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtValorItem.Text <> "" Then
         VALOR_ITEM_N = txtValorItem.Text
         VALOR_ITEM_N = Format(VALOR_ITEM_N, strFormatacao2Digitos)
         VALOR_TOTAL_N = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
         If VALOR_ITEM_N >= VALOR_TOTAL_N Then
            VALOR_TROCO_N = VALOR_ITEM_N - VALOR_TOTAL_N
            txtTroco.Text = Format(VALOR_TROCO_N, strFormatacao2Digitos)
         End If
      End If
      KeyAscii = 0
      txtDias.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_KeyPress"
End Sub

Private Sub txtdias_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "Informe quantidade de dias sua vaca ", "ESC - Confirma", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdias_GotFocus"
End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtEmis.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdias_KeyPress"
End Sub

Private Sub cmdMata_Click()
   If Trim(txtSeq.Text) <> "" Then
      If IsNumeric(txtSeq.Text) Then
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select lancamento_id from LANCAMENTO "

         SQL = SQL & " where numr_doc = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then
            If Not IsNull(TabLancamento.Fields(0).Value) Then

               Msg = "Confirma Exclusão do Item =  ?" & txtSeq.Text
               Style = vbYesNo + 32
               Title = "Atenção !!!"
               Help = "DEMO.HLP"
               Ctxt = 1000
               RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
               If RESPOSTA = vbYes Then

                  SQL = "delete from ITEMLANCAMENTO "
                  SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
                  SQL = SQL & " and seq = " & txtSeq.Text
                  CONECTA_RETAGUARDA.Execute SQL

                  SETA_GRID
               End If
            End If
         End If
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         BUSCA_LANCAMENTO
      End If
   End If
   txtSeq.SetFocus
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   SETA_GRID

   VALOR_DIFERENCA_N = 0

   MOSTRA_RODAPE "Tecle <<ENTER>> para nova seqüência, ou selecione", "ESC - Confirma", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtSeq.Text = "" Then
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         NUMR_SEQ_N = 1
         SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
         SQL = SQL & " and i.numr_doc = l.numr_doc "
         SQL = SQL & " and i.lancamento_id = l.lancamento_id "
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = " & SINAL_INDICADOR_N
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then _
            If Not IsNull(TabLancamento!ultimo_reg) Then _
               NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         txtSeq.Text = NUMR_SEQ_N
         Else
            If TabLancamento.State = 1 Then _
               TabLancamento.Close

            SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
            SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
            SQL = SQL & " and i.numr_doc = l.numr_doc "
            SQL = SQL & " and i.lancamento_id = l.lancamento_id "
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and seq = " & txtSeq.Text
            SQL = SQL & " and l.tipo_lancamento = " & SINAL_INDICADOR_N
            TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLancamento.EOF Then
               'valor lançamento
               txtValorItem.Text = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
               VALOR_DIFERENCA_N = TabLancamento!Valor_Item

               If TabDESCR.State = 1 Then _
                  TabDESCR.Close

               'descrição da modalidade
               SQL = "select * from FORMAPAGTO "
               SQL = SQL & " where formapagto_id = " & TabLancamento!formapagto_id
               SQL = SQL & " and status = 'true' "
               TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabDESCR.EOF Then
                  cmbModalidade.Text = TabDESCR!Descricao
                  cmbModalidadeAUX.Text = TabDESCR!formapagto_id
               End If
               If TabDESCR.State = 1 Then _
                  TabDESCR.Close

               txtDTVENC.PromptInclude = False
               txtDtEmis.PromptInclude = False
               txtDTVENC.Text = TabLancamento!DT_VENCIMENTO
               'txtDTEMIS.Text = data_lancamento
               'else
            End If
            If TabLancamento.State = 1 Then _
               TabLancamento.Close
      End If
      cmbModalidade.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyPress"
End Sub

Private Sub cmbMODALIDADE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorItem.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_KeyPress"
End Sub

Private Sub txtValorItem_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorItem.Text <> "" Then _
      txtValorItem.Text = Format(txtValorItem.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_LostFocus"
End Sub

Private Sub txtValorEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Confirma", "Informe Valor da Entrada", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_GotFocus"
End Sub

Private Sub txtpercEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Confirma", "Informe Percentual(%) da Entrada", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercEntrada_GotFocus"
End Sub

Private Sub txtPercEntrada_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPercEntrada.Text <> "" Then
         VALOR_ITEM_N = txtPercEntrada.Text
         txtValorEntrada.Text = Format(((VALOR_ITEM_N * VALOR_TOTAL_N) / 100), strFormatacao2Digitos)
         txtValorEntrada.Refresh
      End If
      cmbTIPOVENDA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercEntrada_KeyPress"
End Sub

Private Sub txtValorEntrada_keypress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtValorEntrada.Text) <> "" Then
         
         VALOR_RECEBIDO_N = 0 & txtValorEntrada.Text
         txtRecebido.Text = "" & Format(VALOR_RECEBIDO_N, strFormatacao2Digitos)
         txtTroco.Text = "" & Format(VALOR_RECEBIDO_N - (VALOR_ITEM_N - TOTAL_DESCONTO_N), strFormatacao2Digitos)

         If VALOR_RECEBIDO_N >= (VALOR_ITEM_N - TOTAL_DESCONTO_N) Then
            cmbModalidade.Clear
            cmbModalidadeAUX.Clear

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from FORMAPAGTO "
            SQL = SQL & " where formapagto_id < 9999 "
            SQL = SQL & " and status = 'true' "
            If Trim(cmbAuxTIPOVENDA.Text) <> "" Then _
               If IsNumeric(cmbAuxTIPOVENDA.Text) Then _
                  SQL = SQL & " and formapagto_id >= 1 "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            While Not TabConsulta.EOF
               cmbModalidade.AddItem TabConsulta!Descricao
               cmbModalidadeAUX.AddItem TabConsulta!formapagto_id
               TabConsulta.MoveNext
            Wend

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            txtPercEntrada.Text = Format(VALOR_RECEBIDO_N / VALOR_TOTAL_N * 100, strFormatacao2Digitos)
            cmbTIPOVENDA.Text = "A VISTA"
            cmbTIPOVENDA.Refresh
            cmbAuxTIPOVENDA.Text = 9999
            cmbAuxTIPOVENDA.Refresh
            txtSeq.Text = 1
            cmbModalidade.Text = "Dinheiro"
            cmbModalidadeAUX.Text = 1
            txtValorItem.Text = VALOR_RECEBIDO_N
            txtValorItem.Text = VALOR_ITEM_N
            txtDias.Text = 0
            txtDtEmis.PromptInclude = False
            txtDtEmis.Text = Date
            txtDTVENC.PromptInclude = False
            txtDTVENC.Text = Date

            Frame1.Enabled = True

'Call txtDTVENC_KeyPress(13)

      VALOR_ITEM_N = txtValorItem.Text

      GRAVAR_TUDO
      'LIMPA_BODY
      'SETA_GRID

      VALOR_VENDA_N = 0 & txtVendaComDesconto.Text
      VALOR_REC_N = 0 & txtRecebido.Text
MsgBox "vai chamar"
CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA
MsgBox "vortou confirma"

            Else
               txtPercEntrada.Text = Format(((VALOR_ITEM_N / VALOR_TOTAL_N) * 100), strFormatacao2Digitos)
               cmbTIPOVENDA.SetFocus
         End If
         Else: txtPercEntrada.SetFocus
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_keypress"
End Sub

Private Sub txtValorEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   'If txtValorEntrada.Text <> "" Then _
      txtValorEntrada.Text = Format(txtValorEntrada.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_LostFocus"
End Sub

Private Sub txtpercEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtPercEntrada.Text <> "" Then _
      txtPercEntrada.Text = Format(txtPercEntrada.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercEntrada_LostFocus"
End Sub
'============================================================='subrotinas
Private Sub GRAVAR_TUDO()
'On Error GoTo ERRO_TRATA

   Dim STATUS_A As String
   Dim DT_BAIXA As String

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      NUMR_ID_N = TabLancamento!Lancamento_id

      SQL = "UPDATE LANCAMENTO SET "
      SQL = SQL & "Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Prop = '" & CNPJCPF_A & "'"
      SQL = SQL & ", dt_lanc = '" & DMA(Date) & "'"
      SQL = SQL & ", Valor_Lanc = " & Str(txtRecebido.Text)
      SQL = SQL & ", Total_Desconto = " & Str(TOTAL_DESCONTO_N)
      SQL = SQL & ", Tipo_pagto = " & cmbAuxTIPOVENDA.Text

      SQL = SQL & " WHERE Empresa_Id = " & EMPRESA_ID_N
      SQL = SQL & " and Numr_Doc = " & NUMR_REQ_N
      SQL = SQL & " and Tipo_Lancamento = " & SINAL_INDICADOR_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      Else
         NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

         SQL = "INSERT INTO LANCAMENTO "
         SQL = SQL & " ("
            SQL = SQL & " Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, Total_Desconto, "
            SQL = SQL & " Tipo_Lancamento, Empresa_id, Tipo_pagto,pessoa_id,estabelecimento_id"
         SQL = SQL & " ) "
         SQL = SQL & " VALUES ("
            SQL = SQL & NUMR_ID_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & ",'" & CNPJCPF_A & "'"
            SQL = SQL & ",'" & Date & "'"
            SQL = SQL & "," & Str(txtVendaComDesconto.Text)
            SQL = SQL & "," & Str(TOTAL_DESCONTO_N)
            SQL = SQL & "," & SINAL_INDICADOR_N
            SQL = SQL & "," & EMPRESA_ID_N
            SQL = SQL & "," & cmbAuxTIPOVENDA.Text
            SQL = SQL & "," & PESSOA_ID_N
            SQL = SQL & "," & ESTABELECIMENTO_ID_N
         SQL = SQL & ")"
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   CONECTA_RETAGUARDA.Execute SQL

   STATUS_A = "A"
   DT_BAIXA = 0
   CRITERIO = ""

   'ITENS
   If Left(UCase(cmbModalidade.Text), 8) = UCase("Dinheiro") Or _
      Left(UCase(cmbModalidade.Text), 8) = "CHEQUE" Then
      STATUS_A = "B"
      DT_BAIXA = Date
   End If

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and seq = " & txtSeq.Text
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SqL2 = "UPDATE ITEMLANCAMENTO SET "
      SqL2 = SqL2 & " usu_alt = " & CODG_USU_N
      SqL2 = SqL2 & ", Dt_Alt = '" & DMA(Date) & "'"
      SqL2 = SqL2 & ", Dt_Cad = '" & DMA(Date) & "'"
      SqL2 = SqL2 & ", lancamento_id = " & NUMR_ID_N
      SqL2 = SqL2 & ", Numr_doc = " & NUMR_REQ_N
      SqL2 = SqL2 & ", Numr_Dp = " & NUMR_REQ_N
      SqL2 = SqL2 & ", Seq = " & txtSeq.Text
      SqL2 = SqL2 & ", Valor_Item = " & Str(VALOR_ITEM_N)
      SqL2 = SqL2 & ", Status = '" & STATUS_A & "'"
      SqL2 = SqL2 & ", formapagto_id = " & cmbModalidadeAUX.Text
      SqL2 = SqL2 & ", DT_VENCIMENTO = '" & DMA(txtDTVENC.Text) & "'"
      SqL2 = SqL2 & ", dt_baixa = '" & DMA(DT_BAIXA) & "'"
      SqL2 = SqL2 & " Where Lancamento_id = " & NUMR_ID_N
      SqL2 = SqL2 & " and Seq = " & txtSeq.Text
      Else
         SqL2 = "INSERT INTO ITEMLANCAMENTO "
         SqL2 = SqL2 & " (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, seq, "
         SqL2 = SqL2 & " Valor_Item, Status, formapagto_id, DT_VENCIMENTO, acerto,dt_baixa) "
         SqL2 = SqL2 & " VALUES ("
            SqL2 = SqL2 & CODG_USU_N
            SqL2 = SqL2 & ",'" & DMA(Date) & "'"
            SqL2 = SqL2 & ",'" & DMA(Date) & "'"
            SqL2 = SqL2 & "," & NUMR_ID_N
            SqL2 = SqL2 & "," & NUMR_REQ_N
            SqL2 = SqL2 & "," & NUMR_REQ_N
            SqL2 = SqL2 & "," & txtSeq.Text
            SqL2 = SqL2 & "," & Str(VALOR_ITEM_N)
            SqL2 = SqL2 & ",'" & STATUS_A & "'"
            SqL2 = SqL2 & "," & cmbModalidadeAUX.Text
            SqL2 = SqL2 & ",'" & DMA(txtDTVENC.Text) & "'"
            SqL2 = SqL2 & "," & 1
            SqL2 = SqL2 & ",'" & DMA(DT_BAIXA) & "'"
         SqL2 = SqL2 & " )"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SqL2

   If TabLancamento.State = 1 Then _
      TabLancamento.Close
'================================================
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select formapagto_id from ITEMLANCAMENTO"
   SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and Seq = " & txtSeq.Text
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where formapagto_id = " & TabLANCAMENTOITEM.Fields("formapagto_id").Value
      SQL = SQL & " and status = 'true' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         CRITERIO = Trim(TabConsulta.Fields("descricao").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      If Left(UCase(CRITERIO), 8) = UCase("Dinheiro") _
      Or Left(UCase(CRITERIO), 6) = "CHEQUE" _
      Or Left(UCase(CRITERIO), 6) = "CARTAO" _
      Or Left(UCase(CRITERIO), 6) = "CARTÃO" Then

         SQL = "UPDATE ITEMLANCAMENTO SET "
         SQL = SQL & " Status = 'B'"
         SQL = SQL & ", DT_BAIXA = '" & DMA(Date) & "'"
         SQL = SQL & ", CODG_USU_BAIXA = " & CODG_USU_N
         SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
         SQL = SQL & " and Seq = " & txtSeq.Text
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVAR_TUDO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   ListaLanc.ListItems.Clear

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLANCAMENTOITEM.EOF
      'sequencia

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set Item = ListaLanc.ListItems.Add(, "seq." & NUMR_SEQ_N, TabLANCAMENTOITEM!SEQ)
      'numero documento
      Item.SubItems(1) = TabLANCAMENTOITEM.Fields("numr_doc").Value
      'valor lançamento
      Item.SubItems(2) = Format(TabLANCAMENTOITEM!Valor_Item, strFormatacao2Digitos)

      'descrição da modalidade
      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where formapagto_id = " & TabLANCAMENTOITEM!formapagto_id
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and status = 'true' "
      If TabDESCR.State = 1 Then TabDESCR.Close
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         Item.SubItems(3) = TabDESCR!Descricao
      TabDESCR.Close

      Item.SubItems(4) = Date
      Item.SubItems(5) = TabLANCAMENTOITEM!DT_VENCIMENTO

      If cmbAuxTIPOVENDA.Text <> "" Then
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from TIPOVENDA "
         SQL = SQL & " where TIPOVENDA_id = " & cmbAuxTIPOVENDA.Text
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If Not IsNull(TabAUX!PERC_JUROS) Then
               Item.SubItems(6) = TabAUX!PERC_JUROS & "%"
               Else: Item.SubItems(6) = "00,00 %"
            End If
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close
      End If
      TabLANCAMENTOITEM.MoveNext
   Wend
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   BUSCA_LANCAMENTO

   txtVendaSemDesconto.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   lblPRAZO.Caption = ""
   cmbTIPOVENDA.Text = ""
   cmbAuxTIPOVENDA.Text = ""
   txtLanc.Text = ""
   txtVendedor.Text = ""
   txtData.Text = ""
   txtVendaSemDesconto.Text = ""
   txtVendaComDesconto.Text = ""
   txtRecebido.Text = ""
   txtTroco.Text = ""
   txtCli.Text = ""
   txtPercDesconto.Text = ""
   txtDesconto.Text = ""
   cmbModalidadeAUX.Clear
   cmbModalidade.Clear
   txtValorItem.Text = ""
   txtDtEmis.PromptInclude = False
   txtDTVENC.PromptInclude = False
   txtDtEmis.Text = ""
   txtDTVENC.Text = ""
   ListaLanc.ListItems.Clear
   
   txtSeq.Text = ""
   VALOR_TOTAL_LANÇADO = 0
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_LANCAMENTO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   VALOR_DIFERENCA_N = 0
   VALOR_ITEM_N = 0
   txtSeq.Text = ""
   cmbModalidadeAUX.Text = ""
   cmbModalidade.Text = ""
   txtValorItem.Text = ""
   txtDias.Text = ""
   txtDtEmis.PromptInclude = False
   txtDTVENC.PromptInclude = False
   txtDtEmis.Text = ""
   txtDTVENC.Text = ""
   VALOR_TOTAL_LANÇADO = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub GERA_FATURAMENTO()
'On Error GoTo ERRO_TRATA

   Dim Valor_Tot_n As Double

   NUMR_PARCELA = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA "
   SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then

      If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
         If TabTemp.Fields("contabiliza").Value = False Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Esse Tipo de Venda está configurado para não contabilizar !!!"

            SQL = "update PEDIDO set "
            SQL = SQL & "status = 6 " 'não contabiliza
            SQL = SQL & " where numr_req = " & NUMR_REQ_N
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            Exit Sub
            Unload Me
         End If
      End If


      NUMR_PARCELA = TabTemp!parcela
      
      VALOR_DESCONTO_N = 0

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select perc_desc FROM PEDIDO "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & TabCABECA!numr_req
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         PERC_DESCONTO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
      
      VALOR_DESCONTO_CABECA_N = 0
      SQL = "select valor_desconto FROM PEDIDO "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & TabCABECA!numr_req
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_CABECA_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM "
      SQL = SQL & " where numr_req = " & TabCABECA!numr_req
      SQL = SQL & " and tipo_reg = 'PC' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      'BUSCA VALOR TOTAL VENDA
      Valor_Tot_n = 0
      SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM "
      SQL = SQL & " where numr_req = " & TabCABECA!numr_req
      SQL = SQL & " and tipo_reg = 'PC' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         Valor_Tot_n = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
      
      'VALOR_DESCONTO_N = VALOR_DESCONTO_N + (Valor_Tot_n * IIf(PERC_DESCONTO_N > 0, PERC_DESCONTO_N / 100, 1))
      VALOR_DESCONTO_N = VALOR_DESCONTO_N + VALOR_DESCONTO_CABECA_N
      
      VALOR_ITEM_N = 0
      DATA_INI = Date
      If NUMR_PARCELA > 0 Then _
         VALOR_ITEM_N = (Valor_Tot_n - VALOR_DESCONTO_N) / NUMR_PARCELA

      'CABEÇA
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      SQL = "select * from LANCAMENTO "
      SQL = SQL & " where numr_doc = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         NUMR_ID_N = TabLancamento!Lancamento_id

         SQL = "UPDATE LANCAMENTO SET "
            SQL = SQL & " Numr_doc = " & NUMR_REQ_N
            SQL = SQL & ", Prop = '" & CNPJCPF_A & "'"
            SQL = SQL & ", dt_lanc = '" & TabCABECA!dt_req & "'"
            SQL = SQL & ", Valor_Lanc = " & Str(Format(Valor_Tot_n, strFormatacao2Digitos))
            SQL = SQL & ", Total_Desconto = " & Str(Format(VALOR_DESCONTO_N, strFormatacao2Digitos))
            SQL = SQL & ", Tipo_Lancamento = " & SINAL_INDICADOR_N
            SQL = SQL & ", Empresa_Id = " & EMPRESA_ID_N
            SQL = SQL & ", Tipo_pagto = " & cmbAuxTIPOVENDA.Text

         SQL = SQL & " WHERE Empresa_Id = " & EMPRESA_ID_N
         SQL = SQL & " and Numr_Doc = " & NUMR_REQ_N
         SQL = SQL & " and Tipo_Lancamento = " & SINAL_INDICADOR_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         CONECTA_RETAGUARDA.Execute SQL
         Else
            NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

            SQL = "INSERT INTO LANCAMENTO "
            SQL = SQL & " ("
               SQL = SQL & " Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, Total_Desconto, "
               SQL = SQL & " Tipo_Lancamento, Empresa_id, Tipo_pagto,pessoa_id,estabelecimento_id"
            SQL = SQL & " ) "
            SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & NUMR_REQ_N
               SQL = SQL & ",'" & CNPJCPF_A & "'"

               SQL = SQL & ",'" & DMA(Date) & "'"

               SQL = SQL & "," & Str(Format(Valor_Tot_n, strFormatacao2Digitos))
               SQL = SQL & "," & Str(Format(TOTAL_DESCONTO_N, strFormatacao2Digitos))
               SQL = SQL & "," & SINAL_INDICADOR_N
               SQL = SQL & "," & EMPRESA_ID_N
               SQL = SQL & "," & cmbAuxTIPOVENDA.Text
               SQL = SQL & "," & PESSOA_ID_N
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      SQL3 = NUMR_REQ_N
      SqL2 = EMPRESA_ID_N
      CONT_N = 0
      'ITENS
      While CONT_N < NUMR_PARCELA
         GRAVA_LANÇAMENTO
         CONT_N = CONT_N + 1
      Wend
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FATURAMENTO"
End Sub

Private Sub GRAVA_LANÇAMENTO()
'On Error GoTo ERRO_TRATA

   Dim Situacao_A As String

   Situacao_A = "A"
   If Trim(cmbAuxTIPOVENDA.Text) = "9999" Then _
      Situacao_A = "B"

   NUMR_SEQ_N = 1

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento!ultimo_reg) Then _
         NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   'tem que dividir os dias de prazo
   If NUMR_PARCELA <= 0 Then
      DATA_INI = DATA_INI + DIAS_PRAZO
      Else: DATA_INI = DATA_INI + TabTemp!PRAZO
   End If

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where seq = " & NUMR_SEQ_N
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
         SQL = SQL & "  usu_alt = " & CODG_USU_N
         SQL = SQL & ", Dt_Alt = '" & DMA(Date) & "'"
         SQL = SQL & ", Numr_doc = " & NUMR_REQ_N
         SQL = SQL & ", Seq = " & NUMR_SEQ_N
         SQL = SQL & ", Valor_Item = " & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA / NUMR_PARCELA))
         SQL = SQL & ", Status = '" & Situacao_A & "'"
         SQL = SQL & ", formapagto_id = " & TabTemp!formapagto_id
         SQL = SQL & ", DT_VENCIMENTO = '" & DATA_INI & "'"
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & NUMR_SEQ_N
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
            SQL = SQL & " (Usu_Alt, Dt_Alt, Lancamento_id, Numr_doc, "
            SQL = SQL & " NUMR_DP, seq, Valor_Item, Status, formapagto_id, "
            SQL = SQL & " DT_VENCIMENTO, ACERTO) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CODG_USU_N
            SQL = SQL & ",'" & DMA(Date) & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_SEQ_N
            SQL = SQL & "," & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA / NUMR_PARCELA))
            SQL = SQL & ",'" & Situacao_A & "'"
            SQL = SQL & "," & TabTemp!formapagto_id
            SQL = SQL & ",'" & DATA_INI & "'"
            SQL = SQL & "," & 0
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

'================================================

   SQL = "select formapagto_id from ITEMLANCAMENTO"
   SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and Seq = " & NUMR_SEQ_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where formapagto_id = " & TabLANCAMENTOITEM.Fields("formapagto_id").Value
      SQL = SQL & " and status = 'true' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         CRITERIO = Trim(TabConsulta.Fields("descricao").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      If Left(UCase(CRITERIO), 8) = UCase("Dinheiro") _
      Or Left(UCase(CRITERIO), 6) = "CHEQUE" _
      Or Left(UCase(CRITERIO), 6) = "CARTAO" _
      Or Left(UCase(CRITERIO), 6) = "CARTÃO" Then

         SQL = "UPDATE ITEMLANCAMENTO SET "
         SQL = SQL & " Status = 'B'"
         SQL = SQL & ", DT_BAIXA = '" & DMA(Date) & "'"
         SQL = SQL & ", CODG_USU_BAIXA = " & CODG_USU_N
         SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
         SQL = SQL & " and Seq = " & NUMR_SEQ_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close
'=========================================

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_LANÇAMENTO"
End Sub

Private Sub BUSCA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_LANÇADO = 0
   VALOR_RECEBIDO_N = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select sum(valor_item) from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento.Fields(0).Value) Then _
         VALOR_TOTAL_LANÇADO = TabLancamento.Fields(0).Value
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   txtRecebido.Text = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   txtRecebido.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_LANCAMENTO"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO

   VALOR_TOTAL_LANÇADO = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   VALOR_TOTAL_N = Format((VALOR_TOTAL_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

   If tpMOEDA(VALOR_TOTAL_LANÇADO) = tpMOEDA(VALOR_TOTAL_N) Then
      If Left(UCase(cmbModalidade.Text), 8) = UCase("Dinheiro") Then
         Msg = "Confirma recebimento ?"
         PERGUNTA Msg, vbYesNo + 32, "Recebimento ", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
   
            SQL = "select * FROM PEDIDO "
            SQL = SQL & " where numr_req = " & txtLanc.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCABECA.EOF Then
               If TabCABECA!Status = 2 Then
                  SQL = "UPDATE PEDIDO set "
                  SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
                  SQL = SQL & " where numr_req = " & txtLanc.Text
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
            Unload Me
            Exit Sub
         End If
         Else
            If VALOR_TOTAL_LANÇADO >= VALOR_TOTAL_N Then
               If txtTroco.Text <> "" Then _
                  If VALOR_TROCO_N > 0 Then _
                     REGISTRA_TROCO
   
               Msg = "Confirma recebimento ?"
               PERGUNTA Msg, vbYesNo + 32, "Recebimento ", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                 If TabCABECA.State = 1 Then _
                    TabCABECA.Close
   
                  SQL = "select * FROM PEDIDO "
                  SQL = SQL & " where numr_req = " & txtLanc.Text
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabCABECA.EOF Then
                     If TabCABECA!Status = 2 Then
                        SQL = "UPDATE PEDIDO set "
                        SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
                        SQL = SQL & " where numr_req = " & txtLanc.Text
                        SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                        SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                        CONECTA_RETAGUARDA.Execute SQL
                     End If
                  End If
                  If TabCABECA.State = 1 Then _
                     TabCABECA.Close
                  Unload Me
                  Exit Sub
               End If
            End If
      End If
   End If

   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA"
End Sub

Private Sub REGISTRA_TROCO()
'On Error GoTo ERRO_TRATA

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select lancamento_id from LANCAMENTO "

   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      If Not IsNull(TabLancamento.Fields(0).Value) Then
         SQL = "update itemLANCAMENTO set "
         SQL = SQL & " valor_desconto = valor_desconto + " & Replace(VALOR_TROCO_N, ",", ".")
         SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
         SQL = SQL & " and seq = " & NUMR_SEQ_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REGISTRA_TROCO"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_PARCELADO()
'On Error GoTo ERRO_TRATA

   'BUSCA_LANCAMENTO

   If VALOR_TOTAL_LANÇADO <= 0 Then _
      VALOR_TOTAL_LANÇADO = 0 & txtRecebido.Text

   VALOR_TOTAL_LANÇADO = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   VALOR_TOTAL_N = Format(VALOR_TOTAL_N, strFormatacao2Digitos)

   If Round(VALOR_TOTAL_LANÇADO) >= Round(VALOR_TOTAL_N) Then
   'If Format(Round(VALOR_TOTAL_LANÇADO), strFormatacao2Digitos) >= Format(Round(VALOR_TOTAL_N), strFormatacao2Digitos) Then
      INDR_FINALIZA_RECEBIMENTO = False
      Msg = "Confirma recebimento ?"
      PERGUNTA Msg, vbYesNo + 32, "Recebimento NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         txtLanc.Text = NUMR_REQ_N
         If txtLanc.Text = "" Then _
            txtLanc.Text = 1

         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select * FROM PEDIDO "
         SQL = SQL & " where numr_req = " & txtLanc.Text
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If TabCABECA!Status = 2 Then
               SQL = "UPDATE PEDIDO set "
               SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
               SQL = SQL & ", TIPOVENDA_ID = " & cmbAuxTIPOVENDA.Text
               SQL = SQL & " where numr_req = " & txtLanc.Text
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         LIMPA_LANCAMENTO

         Me.Hide
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub

Private Sub EXCLUIR_TUDO()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select lancamento_id from LANCAMENTO "

   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      NUMR_ID_N = TabConsulta.Fields(0).Value

      SQL = "delete from ITEMLANCAMENTO "
      SQL = SQL & " where lancamento_id = " & NUMR_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_TUDO"
End Sub
