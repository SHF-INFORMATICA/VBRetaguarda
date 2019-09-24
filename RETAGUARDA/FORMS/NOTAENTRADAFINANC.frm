VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNOTAENTRADAFINANC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Títulos à Pagar"
   ClientHeight    =   7560
   ClientLeft      =   2280
   ClientTop       =   2460
   ClientWidth     =   10950
   Icon            =   "NOTAENTRADAFINANC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10950
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   9000
         TabIndex        =   28
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cmbAuxTIPOVENDA 
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
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTIPOVENDA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox txtData 
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
         Height          =   345
         Left            =   9000
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtVendaComDesconto 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   9000
         TabIndex        =   22
         Top             =   2280
         Width           =   1800
      End
      Begin VB.TextBox txtRecebido 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   9000
         TabIndex        =   15
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtLanc 
         Alignment       =   2  'Center
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
         Height          =   345
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtVendaSemDesconto 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   9000
         TabIndex        =   13
         Top             =   1320
         Width           =   1800
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
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox txtVendedor 
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
         Left            =   6360
         TabIndex        =   11
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   29
         Top             =   1845
         Width           =   1335
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
         TabIndex        =   27
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Forma Faturamento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   1880
         Width           =   1665
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Compra com Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   23
         Top             =   2325
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Pagamento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido Compra:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Compra sem Desconto: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   19
         Top             =   1365
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Entrada:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   17
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   5235
         TabIndex        =   16
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Width           =   11055
      Begin VB.TextBox txtSeq 
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
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cmbAuxLanc 
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
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtValorItem 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbMODALIDADE 
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
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   4335
      End
      Begin MSMask.MaskEdBox txtDTVENC 
         Height          =   345
         Left            =   9240
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDTEMIS 
         Height          =   345
         Left            =   7320
         TabIndex        =   4
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Vencimento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9240
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Emissão:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7320
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Modalidade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Seq:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   435
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
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
            Picture         =   "NOTAENTRADAFINANC.frx":47C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADAFINANC.frx":4809E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADAFINANC.frx":483BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADAFINANC.frx":4880E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADAFINANC.frx":48C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADAFINANC.frx":48F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADAFINANC.frx":493D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1270
      ButtonWidth     =   2672
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
            Caption         =   "&Confirmar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Sequência"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5880
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADAFINANC.frx":496F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADAFINANC.frx":4A881
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADAFINANC.frx":4BF7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADAFINANC.frx":4D00D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListaLanc 
      Height          =   2625
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   4630
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   ""
      SmallIcons      =   ""
      ColHdrIcons     =   ""
      ForeColor       =   4194304
      BackColor       =   14737632
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Doc."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Modalidade"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Dt.Lanç."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dt.Venc."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Juros"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmNOTAENTRADAFINANC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_RECEBIDO_N As Double, NUMR_PARCELA As Integer
   Dim VALOR_TROCO_N As Double, VALOR_TOTAL_LANÇADO As Double
   Dim VALOR_ENTRADA As Double, PERC_JUROS_N As Double, DIAS_PRAZO As Integer
   Dim TabTipoVenda As Recordset
   Dim VALOR_ICMS_SUB_N As Currency

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela2(frmNOTAENTRADAFINANC)

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", ""

   If SINAL <> 2 Then
      MsgBox "Registro não é do contas a pagar, não permitido, chamar suporte."
      Unload Me
      Exit Sub
   End If

   Frame1.Enabled = False
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

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where NUMR_PEDIDO_COMPRA = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      NUMR_ID_N = TabCABECA.Fields("entrada_id").Value

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select nome from USUARIO "
      SQL = SQL & " where usuario_id = " & TabCABECA!CODG_USU
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         txtVendedor.Text = TabTemp!NOME
         txtVendedor.Refresh
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select nome,cgccpf from FORNECEDOR "
      SQL = SQL & " where fornecedor_id = " & TabCABECA!fornecedor_id
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         CPF_N = TabTemp!CGCCPF
         txtCli.Text = TabTemp!NOME
         txtCli.Refresh
         Else
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Fornecedor não encontrado !!!"
            Unload Me
            Exit Sub
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      VALOR_TOTAL_N = 0
      SQL = "select sum(preco_custo*qtd_entrada) "
      SQL = SQL & " from NOTAENTRADAITEM i, NOTAENTRADA n "
      SQL = SQL & " where n.entrada_id = " & NUMR_ID_N
      SQL = SQL & " and n.entrada_id = i.entrada_id "
      SQL = SQL & " and n.empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_TOTAL_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close

      VALOR_DESCONTO_N = 0
      SQL = "select valor_desconto from NOTAENTRADA "
      SQL = SQL & " where entrada_id = " & NUMR_ID_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_DESCONTO_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      VALOR_ICMS_SUB_N = 0
      SQL = "select valor_icms_subst from NOTAENTRADA "
      SQL = SQL & " where entrada_id = " & NUMR_ID_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         VALOR_ICMS_SUB_N = TabTemp.Fields(0).Value
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      VALOR_IPI_N = 0
      SQL = "select valor_ipi from NOTAENTRADA "
      SQL = SQL & " where entrada_id = " & NUMR_ID_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         VALOR_IPI_N = TabTemp.Fields(0).Value
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select sum(valor_desconto*qtd_entrada) from NOTAENTRADAITEM "
      SQL = SQL & " where entrada_id = " & NUMR_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_DESCONTO_N = 0 & TabTemp.Fields(0).Value + VALOR_DESCONTO_N
      If TabTemp.State = 1 Then _
         TabTemp.Close

      VALOR_TOTAL_N = (VALOR_TOTAL_N + VALOR_ICMS_SUB_N + VALOR_IPI_N - VALOR_DESCONTO_N)
      txtVendaSemDesconto.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
      If VALOR_DESCONTO_N > 0 Then _
         txtVendaComDesconto.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   BUSCA_LANCAMENTO

   txtRecebido.Refresh
   txtVendaSemDesconto.Refresh

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
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         Msg = "Deseja Cancelar todo lançamento ?"
         Style = vbYesNo + 32
         Title = "Atenção !!!"
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then _
            MATA_LANCAMENTO
         SETA_GRID
         LIMPA_BODY
         VALOR_ITEM_N = 0
         VALOR_ENTRADA = 0
         cmbTIPOVENDA.Text = ""
         cmbAuxTIPOVENDA.Text = ""
         Frame1.Enabled = False
         cmbTIPOVENDA.SetFocus
      Case "matar"
         If txtSeq.Text <> "" Then
            Msg = "Confirma Exclusão do Item =  ?" & txtSeq.Text
            Style = vbYesNo + 32
            Title = "Atenção !!!"
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then _
               MATA_LANCAMENTO
            SETA_GRID
            Else: MsgBox "Informe número da seqüência."
         End If
      Case "voltar"
         CONFIRMAR_RECEBIMENTO_PARCELADO
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbTIPOVENDA_LostFocus()
'On Error GoTo ERRO_TRATA

   If cmbAuxTIPOVENDA.Text <> "" Then
      cmbMODALIDADE.Clear
      cmbAuxLanc.Clear

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabDESCR.EOF
         cmbMODALIDADE.AddItem TabDESCR!Descricao
         cmbAuxLanc.AddItem TabDESCR!FORMA_ID
         TabDESCR.MoveNext
      Wend
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
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

   SETA_GRID
   If cmbAuxTIPOVENDA.Text <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
   
      SQL = "select * from TIPOVENDA "
      SQL = SQL & " where TIPOVENDA_id = " & cmbAuxTIPOVENDA.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         NUMR_PARCELA = 0
         DIAS_PRAZO = 0
         If Not IsNull(TabTemp!parcela) Then _
            NUMR_PARCELA = TabTemp!parcela
         If Not IsNull(TabTemp!Prazo) Then _
            DIAS_PRAZO = TabTemp!Prazo
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      Else
         MsgBox "Selecione tipo de venda."
         Exit Sub
   End If

   Frame1.Enabled = True
   txtSeq.SetFocus
   Exit Sub
   
   If cmbAuxTIPOVENDA.Text = 2 Then
      If TabTipoVenda!parcela = 0 Then
         Frame1.Enabled = True
         txtSeq.SetFocus
      Else
         lblPRAZO.Caption = TabTipoVenda!Prazo & " dias"
         lblPRAZO.Refresh
         Frame1.Enabled = False
         If NUMR_PARCELA = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue parcelas. " & TabTipoVenda!tipovenda_id & " - " & TabTipoVenda!Descricao

            If TabTipoVenda.State = 1 Then _
               TabTipoVenda.Close

            cmbTIPOVENDA.SetFocus
            Exit Sub
         End If
         If DIAS_PRAZO = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue dias de vencimento. " & TabTipoVenda!tipovenda_id & " - " & TabTipoVenda!Descricao

            If TabTipoVenda.State = 1 Then _
               TabTipoVenda.Close

            cmbTIPOVENDA.SetFocus
            Exit Sub
         End If
         NUMR_SEQ_N = 1
         CONT_N = 0

         'GERA TITULOS
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select * from NOTAENTRADA "
         SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            GERA_FATURAMENTO
            SETA_GRID
         End If
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         CONFIRMAR_RECEBIMENTO_PARCELADO
      End If
      Else   'se parcelas parametrizadas não for 0 entra aqui
         lblPRAZO.Caption = TabTemp!Prazo & " dias"
         lblPRAZO.Refresh
         If NUMR_PARCELA = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue parcelas. " & TabTemp!tipovenda_id & " - " & TabTemp!Descricao

            If TabTemp.State = 1 Then _
               TabTemp.Close

            cmbTIPOVENDA.SetFocus
            Exit Sub
         End If
         If DIAS_PRAZO = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue dias de vencimento. " & TabTemp!tipovenda_id & " - " & TabTemp!Descricao

            If TabTemp.State = 1 Then _
               TabTemp.Close

            cmbTIPOVENDA.SetFocus
            Exit Sub
         End If
         Frame1.Enabled = False
         CONT_N = 0

         'GERA TITULOS
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select * from NOTAENTRADA "
         SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            GERA_FATURAMENTO
            SETA_GRID
         End If
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         CONFIRMAR_RECEBIMENTO_PARCELADO
   End If

   If TabFORNEC.State = 1 Then _
      TabFORNEC.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_Click"
End Sub

Private Sub cmbTIPOVENDA_GotFocus()
'On Error GoTo ERRO_TRATA

   Frame1.Enabled = False
   cmbTIPOVENDA.Clear
   cmbAuxTIPOVENDA.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbTIPOVENDA.AddItem TabTemp!Descricao
      cmbAuxTIPOVENDA.AddItem TabTemp!tipovenda_id
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
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_KeyPress"
End Sub

Private Sub cmbMODALIDADE_Click()
'On Error GoTo ERRO_TRATA
   cmbAuxLanc.ListIndex = cmbMODALIDADE.ListIndex
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub cmbmodalidade_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione Forma de Pagto."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmodalidade_GotFocus"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA
   txtDtEmis.PromptInclude = True
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
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Data Vencimento da parcela"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_GotFocus"
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtDTVENC.PromptInclude = True
      If Not IsDate(txtDTVENC.Text) Then
         'MsgBox "Data Informada Inválida !!!"
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
      If cmbAuxLanc.Text = "" Then
         MsgBox "Selecione Forma de Pagamento !!!"
         cmbMODALIDADE.SetFocus
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
      
      GRAVAR_TUDO
      
      LIMPA_BODY
      SETA_GRID
      txtSeq.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_KeyPress"
End Sub

Private Sub txtDTVENC_LostFocus()
'On Error GoTo ERRO_TRATA

   CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_LostFocus"
End Sub

Private Sub txtValorItem_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe o valor da parcela"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_GotFocus"
End Sub

Private Sub txtValorItem_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_KeyPress"
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   SETA_GRID
   VALOR_DIFERENCA_N = 0
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Tecle <<ENTER>> para nova seqüência, ou selecione"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtSeq.Text = "" Then
         NUMR_SEQ_N = 1

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select max(seq) as ultimo_reg "
         SQL = SQL & " from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
         SQL = SQL & " and i.numr_doc = l.numr_doc "
         SQL = SQL & " and i.lancamento_id = l.lancamento_id "
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = " & SINAL
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then _
            If Not IsNull(TabLancamento!ultimo_reg) Then _
               NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         txtSeq.Text = NUMR_SEQ_N
         Else
            SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
            SQL = SQL & " where seq = " & txtSeq.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and i.numr_doc = l.numr_doc "
            SQL = SQL & " and i.lancamento_id = l.lancamento_id "
            SQL = SQL & " and i.numr_doc = " & NUMR_REQ_N
            SQL = SQL & " and l.tipo_lancamento = " & SINAL
            TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLancamento.EOF Then
               'valor lançamento
               txtValorItem.Text = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
               VALOR_DIFERENCA_N = TabLancamento!Valor_Item

               If TabDESCR.State = 1 Then _
                  TabDESCR.Close

               'descrição da modalidade
               SQL = "select * from FORMAPAGTO "
               SQL = SQL & " where forma_id = " & TabLancamento!FORMA_ID
               'SQL = SQL & " and empresa_id = " & EMPRESA_ID
               TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabDESCR.EOF Then
                  cmbMODALIDADE.Text = TabDESCR!Descricao
                  cmbAuxLanc.Text = TabDESCR!FORMA_ID
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
      cmbMODALIDADE.SetFocus
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

Private Sub txtLanc_GotFocus()
'On Error GoTo ERRO_TRATA
   txtSeq.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLanc_GotFocus"
End Sub

Private Sub txtrecebido_gotfocus()
'On Error GoTo ERRO_TRATA
   txtSeq.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrecebido_gotfocus"
End Sub

Private Sub txtValorItem_LostFocus()
'On Error GoTo ERRO_TRATA
   If txtValorItem.Text <> "" Then
      txtValorItem.Text = Format(txtValorItem.Text, strFormatacao2Digitos)
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_LostFocus"
End Sub

Private Sub txtvendasemdesconto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, " txtvendasemdesconto_GotFocus"
End Sub
'subrotinas
Private Sub GRAVAR_TUDO()
'On Error GoTo ERRO_TRATA

   VALOR_ITEM_N = 0 & txtValorItem.Text
   VALOR_DESCONTO_N = 0 & txtDesconto.Text

   SINAL = 2

   SQL = "select * from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      NUMR_ID_N = TabLancamento!LANCAMENTO_ID
      SqL2 = "UPDATE LANCAMENTO SET Lancamento_id = " & NUMR_ID_N & ", Numr_doc = " & NUMR_REQ_N & ", Prop = '" & CPF_N & "',"
      SqL2 = SqL2 & " dt_lanc = '" & DMA(Date) & "', Valor_Lanc = " & Str(txtVendaSemDesconto.Text) & ", Total_Desconto = " & Str(VALOR_DESCONTO_N) & ", Tipo_pagto = " & cmbAuxTIPOVENDA.Text & " WHERE Empresa_Id = " & EMPRESA_ID_N & " and Numr_Doc = " & NUMR_REQ_N & " and Tipo_Lancamento = " & SINAL
      CONECTA_RETAGUARDA.Execute SqL2
      Else
         NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")


         SqL2 = "INSERT INTO LANCAMENTO (Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, Total_Desconto, Tipo_Lancamento, Empresa_id, Tipo_pagto) "
         SqL2 = SqL2 & " VALUES (" & NUMR_ID_N & "," & NUMR_REQ_N & ",'" & CPF_N & "','" & DMA(Date) & "'," & Str(txtVendaSemDesconto.Text) & "," & Str(VALOR_DESCONTO_N) & "," & SINAL & "," & EMPRESA_ID_N & "," & cmbAuxTIPOVENDA.Text & ")"
         CONECTA_RETAGUARDA.Execute SqL2
      SQL3 = NUMR_REQ_N
   End If
  
   'ITENS
   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where seq = " & txtSeq.Text
   SQL = SQL & " and numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
       SqL2 = "UPDATE ITEMLANCAMENTO SET usu_alt = " & CODG_USU_N & ", Dt_Alt = '" & Date & "', Dt_Cad = '" & Date & "',"
       SqL2 = SqL2 & " lancamento_id = " & NUMR_ID_N & ", Numr_doc = " & NUMR_REQ_N & ",  Numr_Dp = " & NUMR_REQ_N & ", Seq = " & txtSeq.Text & ", Valor_Item = " & Str(VALOR_ITEM_N) & ","
       SqL2 = SqL2 & " Status = '" & "A" & "', FORMA_ID = " & cmbAuxLanc.Text & ", DT_VENCIMENTO = '" & txtDTVENC.Text & "' Where Lancamento_id = " & NUMR_ID_N & " and Seq = " & txtSeq.Text
       CONECTA_RETAGUARDA.Execute SqL2
   Else
       SqL2 = "INSERT INTO ITEMLANCAMENTO (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, seq, Valor_Item, Status, FORMA_ID, DT_VENCIMENTO, CODG_USU_BAIXA, Acerto) "
       SqL2 = SqL2 & " VALUES (" & CODG_USU_N & ",'" & DMA(Date) & "','" & DMA(Date) & "'," & NUMR_ID_N & "," & NUMR_REQ_N & "," & NUMR_REQ_N & "," & txtSeq.Text & "," & Str(VALOR_ITEM_N) & ",'" & "A" & "'," & cmbAuxLanc.Text & ",'" & DMA(txtDTVENC.Text) & "'," & CODG_USU_N & "," & 1 & ")"
       CONECTA_RETAGUARDA.Execute SqL2
   End If
   
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVAR_TUDO"
End Sub

Private Sub MATA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   SQL = "select lancamento_id from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      If Not IsNull(TabLancamento.Fields(0).Value) Then
         SQL = "delete from LANCAMENTO "
         SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and tipo_lancamento = " & SINAL
         If txtSeq.Text <> "" Then
            If IsNumeric(txtSeq.Text) Then
               SQL = SQL & " and seq = " & txtSeq.Text
            End If
         End If
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   BUSCA_LANCAMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_LANCAMENTO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   ListaLanc.ListItems.Clear
   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where l.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and l.tipo_lancamento = " & SINAL
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLANCAMENTOITEM.EOF
      'sequencia
      Set Item = ListaLanc.ListItems.Add(, "seq." & TabLANCAMENTOITEM!SEQ, TabLANCAMENTOITEM!SEQ)
      'numero documento
      Item.SubItems(1) = TabLANCAMENTOITEM!NUMR_DOC
      'valor lançamento
      Item.SubItems(2) = Format(TabLANCAMENTOITEM!Valor_Item, strFormatacao2Digitos)
      Item.SubItems(3) = Format(0 & TabLANCAMENTOITEM!PERC_DESCONTO * TabLANCAMENTOITEM!Valor_Item / 100, strFormatacao2Digitos)
      Item.SubItems(4) = Format(TabLANCAMENTOITEM!Valor_Item - (TabLANCAMENTOITEM!PERC_DESCONTO * TabLANCAMENTOITEM!Valor_Item / 100), strFormatacao2Digitos)

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      'descrição da modalidade
      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where forma_id = " & TabLANCAMENTOITEM!FORMA_ID
      'SQL = SQL & " and empresa_id = " & EMPRESA_ID
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         Item.SubItems(5) = TabDESCR!Descricao
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      Item.SubItems(6) = Date
      Item.SubItems(7) = TabLANCAMENTOITEM!DT_VENCIMENTO

      If cmbAuxTIPOVENDA.Text <> "" Then
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from TIPOVENDA "
         SQL = SQL & " where TIPOVENDA_id = " & cmbAuxTIPOVENDA.Text
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If Not IsNull(TabAUX!PERC_JUROS) Then
               Item.SubItems(8) = TabAUX!PERC_JUROS & "%"
               Else: Item.SubItems(8) = "00,00 %"
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
   txtRecebido.Refresh
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
   txtCli.Text = ""
   cmbAuxLanc.Clear
   cmbMODALIDADE.Clear
   txtValorItem.Text = ""
   txtDtEmis.PromptInclude = False
   txtDTVENC.PromptInclude = False
   txtDtEmis.Text = ""
   txtDTVENC.Text = ""
   ListaLanc.ListItems.Clear
   txtSeq.Text = ""
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
   cmbAuxLanc.Text = ""
   cmbMODALIDADE.Text = ""
   txtValorItem.Text = ""
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

   SINAL = 2
   NUMR_PARCELA = 0
   VALOR_DESCONTO_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA "
   SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      NUMR_PARCELA = TabTemp!parcela
      VALOR_DESCONTO_N = 0 & txtDesconto.Text
      VALOR_ITEM_N = 0
      DATA_INI = Date
      VALOR_ITEM_N = VALOR_TOTAL_N / NUMR_PARCELA

      'CABEÇA
      If TabFORNEC.State = 1 Then _
         TabFORNEC.Close

      SQL = "Select * from Fornecedor where fornecedor_id = " & TabCABECA!fornecedor_id
      TabFORNEC.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      
      SQL = "select * from LANCAMENTO "
      SQL = SQL & " where numr_doc = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and tipo_lancamento = " & SINAL
      If TabLancamento.State = 1 Then TabLancamento.Close
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         NUMR_ID_N = TabLancamento!LANCAMENTO_ID
         SqL2 = "UPDATE LANCAMENTO SET Lancamento_id = " & NUMR_ID_N & ", Numr_doc = " & NUMR_REQ_N & ", Prop = '" & TabFORNEC!CGCCPF & "',"
         SqL2 = SqL2 & " dt_lanc = '" & DMA(Date) & "', Valor_Lanc = " & Str(txtVendaSemDesconto.Text) & ", Total_Desconto = " & Str(VALOR_DESCONTO_N) & ", Tipopagto = " & cmbAuxTIPOVENDA.Text & " WHERE Empresa_Id = " & EMPRESA_ID_N & " and Numr_Doc = " & NUMR_REQ_N & " and Tipo_Lancamento = " & SINAL
         CONECTA_RETAGUARDA.Execute SqL2
      Else
         NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")
         SqL2 = "INSERT INTO LANCAMENTO (Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, Total_Desconto, Tipo_Lancamento, Empresa_id, Tipo_pagto) "
         SqL2 = SqL2 & " VALUES (" & NUMR_ID_N & "," & NUMR_REQ_N & ",'" & TabFORNEC!CGCCPF & "','" & Date & "'," & Str(txtVendaSemDesconto.Text) & "," & Str(VALOR_DESCONTO_N) & "," & SINAL & "," & EMPRESA_ID_N & "," & cmbAuxTIPOVENDA.Text & ")"
         CONECTA_RETAGUARDA.Execute SqL2
      End If
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      VALOR_DESCONTO_N = VALOR_DESCONTO_N / NUMR_PARCELA
      While CONT_N <> NUMR_PARCELA
         GRAVA_LANÇAMENTO
         CONT_N = CONT_N + 1
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   If TabFORNEC.State = 1 Then _
      TabFORNEC.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FATURAMENTO"
End Sub

Private Sub GRAVA_LANÇAMENTO()
'On Error GoTo ERRO_TRATA

   Dim DtEmissao_D      As String
   Dim DtVencimento_D   As String

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   NUMR_SEQ_N = 1
   SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and l.tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento!ultimo_reg) Then _
         NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   'ITENS
   DATA_INI = DATA_INI + TabTemp!Prazo

   DtEmissao_D = Date
   DtVencimento_D = Date

   txtDtEmis.PromptInclude = True
   txtDTVENC.PromptInclude = True

   If IsDate(txtDtEmis.Text) Then _
      DtEmissao_D = txtDtEmis.Text

   If IsDate(txtDTVENC.Text) Then _
      DtVencimento_D = txtDTVENC.Text

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close
   
   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where seq = " & NUMR_SEQ_N
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " usu_alt = " & CODG_USU_N
      SQL = SQL & ", Dt_Alt = '" & Date & "'"
      SQL = SQL & ", lancamento_id = " & NUMR_ID_N
      SQL = SQL & ", Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Numr_Dp = " & NUMR_REQ_N
      SQL = SQL & ", Seq = " & txtSeq.Text
      SQL = SQL & ", Valor_Item = " & Str(VALOR_ITEM_N - (VALOR_ENTRADA / NUMR_PARCELA))
      SQL = SQL & ", Status = 'A'"
      SQL = SQL & ", FORMA_ID = " & TabTemp!FORMA_ID
      SQL = SQL & ", DT_VENCIMENTO = '" & DMA(DtVencimento_D) & "'"
      SQL = SQL & ", DT_cad = '" & DMA(DtEmissao_D) & "'"
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & NUMR_SEQ_N
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
            SQL = SQL & " (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, seq, Valor_Item, Status, FORMA_ID, DT_VENCIMENTO, CODG_USU_BAIXA, Acerto) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CODG_USU_N
            SQL = SQL & ",'" & Date & "'"
            SQL = SQL & ",'" & DMA(DtEmissao_D) & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_SEQ_N
            SQL = SQL & "," & Str(VALOR_ITEM_N - (VALOR_ENTRADA / NUMR_PARCELA))
            SQL = SQL & ",'A'"
            SQL = SQL & "," & TabTemp!FORMA_ID
            SQL = SQL & ",'" & DMA(DtVencimento_D) & "'"
            SQL = SQL & "," & CODG_USU_N
            SQL = SQL & "," & 1
         SQL = SQL & ")"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_LANÇAMENTO"
End Sub

Private Sub BUSCA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_LANÇADO = 0
   VALOR_RECEBIDO_N = 0
   txtRecebido.Text = ""

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select sum(valor_item) from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
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
   VALOR_DESCONTO_N = 0 & txtDesconto.Text
   If Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos) >= (Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)) Then
      Msg = "Confirma recebimento ?"
      PERGUNTA Msg, vbYesNo + 32, "Recebimento Entrada de Mercadoria", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         Unload Me
         Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_PARCELADO()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO
   If VALOR_TOTAL_LANÇADO >= (VALOR_TOTAL_N - VALOR_DESCONTO_N) Then
      Msg = "Confirma lançamento ?"
      PERGUNTA Msg, vbYesNo + 32, "Recebimento Entrada", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         Me.Hide
         Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub
