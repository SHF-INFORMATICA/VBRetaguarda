VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFINGERALANC 
   Caption         =   "Contas à Receber e Contas à Pagar"
   ClientHeight    =   6735
   ClientLeft      =   1740
   ClientTop       =   2025
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GERACONTAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbCCaux 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   1560
      TabIndex        =   69
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbCC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   67
      Top             =   3840
      Width           =   4095
   End
   Begin VB.CommandButton cmdRecibo 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5760
      Picture         =   "GERACONTAS.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Recibo"
      Top             =   3840
      Width           =   405
   End
   Begin VB.CommandButton cmdMata 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   480
      Picture         =   "GERACONTAS.frx":69A5
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2160
      Width           =   405
   End
   Begin VB.CommandButton cmdLanc 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   1390
      Picture         =   "GERACONTAS.frx":77E6
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   1080
      Width           =   405
   End
   Begin VB.CommandButton cmdForCli 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   4320
      Picture         =   "GERACONTAS.frx":81E8
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   1560
      Width           =   405
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tela Usada para Desconto de Cheques"
      ForeColor       =   &H00400000&
      Height          =   1245
      Left            =   120
      TabIndex        =   54
      Top             =   7320
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox txtCheque 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   61
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtrep 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3240
         TabIndex        =   55
         Top             =   330
         Width           =   4815
      End
      Begin MSMask.MaskEdBox txtCGCREP 
         Height          =   360
         Left            =   1080
         TabIndex        =   56
         ToolTipText     =   "Selecione Fornecedor Para Descontar Cheque"
         Top             =   330
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox txtdtdesc 
         Height          =   360
         Left            =   1080
         TabIndex        =   57
         ToolTipText     =   "Data do Desconto do Cheque!"
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
      Begin VB.Label Label22 
         Caption         =   "Nº Cheque:"
         Height          =   255
         Left            =   3330
         TabIndex        =   60
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblrepasse 
         Caption         =   "Repasse:"
         Height          =   255
         Left            =   270
         TabIndex        =   59
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Dt. Desc.:"
         Height          =   255
         Left            =   180
         TabIndex        =   58
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   53
      Top             =   2640
      Width           =   615
      Begin VB.OptionButton optPercDesc 
         Caption         =   "%"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   630
      End
      Begin VB.OptionButton optValorDesc 
         Caption         =   "R$"
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   0
         TabIndex        =   14
         Top             =   195
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      TabIndex        =   52
      Top             =   2640
      Width           =   615
      Begin VB.OptionButton optValorJuros 
         Caption         =   "R$"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optPercJuros 
         Caption         =   "%"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCCusto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9120
      TabIndex        =   21
      Top             =   7380
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbCCustoAux 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   9120
      TabIndex        =   49
      Top             =   7755
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbAuxGridcc 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   10800
      TabIndex        =   48
      Top             =   7755
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCCustoItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   9120
      TabIndex        =   22
      Top             =   7755
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   1575
      Left            =   9120
      TabIndex        =   47
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codg."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtJuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   4680
      MaxLength       =   12
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtTotalItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   10320
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtNumrDoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4680
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cmbStatusLancItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   16
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtPessoa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   1560
      Width           =   6855
   End
   Begin VB.TextBox txtToTDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7680
      TabIndex        =   39
      Text            =   " "
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9840
      TabIndex        =   36
      Text            =   " "
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtValorItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmbModalidadeAux 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   3000
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbModalidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
   End
   Begin VB.ComboBox cmbTipoRegistroAUX 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   4320
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtLanc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      MaxLength       =   9
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipoRegistro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   4320
      TabIndex        =   19
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtHistorico 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3270
      Width           =   4935
   End
   Begin MSMask.MaskEdBox txtDtEmis 
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
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
   Begin MSMask.MaskEdBox txtDtVenc 
      Height          =   360
      Left            =   7680
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
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
   Begin MSMask.MaskEdBox txtDtBaixa 
      Height          =   360
      Left            =   10320
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
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
   Begin MSComctlLib.ListView ListaTitulos 
      Height          =   2325
      Left            =   60
      TabIndex        =   41
      Top             =   4380
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   4101
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Título"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Modalidade"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Juros"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Dt.Baixa"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Dt.Venc."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Situação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Dt.Cancelamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Histórico"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CentroCusto"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1270
      ButtonWidth     =   2858
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
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
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Carnê"
            Key             =   "carne"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
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
         Left            =   10680
         TabIndex        =   62
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11850
      DesignHeight    =   6735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":8BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":A012
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":B0A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":C309
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":D414
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":EB11
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":FCAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":10EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GERACONTAS.frx":120FD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Nota"
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
      Left            =   1920
      TabIndex        =   70
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "CentroCusto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   68
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Grp.CC:"
      Height          =   225
      Left            =   8430
      TabIndex        =   51
      Top             =   7440
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Item CC:"
      Height          =   225
      Left            =   8370
      TabIndex        =   50
      Top             =   7800
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Shape Shape3 
      Height          =   795
      Left            =   8250
      Top             =   7365
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Juros:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4080
      TabIndex        =   46
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   225
      Left            =   9810
      TabIndex        =   45
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label lblCob 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   225
      Left            =   6360
      TabIndex        =   44
      Top             =   3750
      Width           =   315
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Nº Doc.:"
      Height          =   240
      Left            =   3960
      TabIndex        =   43
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblCli 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cliente/Fornecedor:"
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
      Left            =   -90
      TabIndex        =   42
      Top             =   1620
      Width           =   1890
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total Desconto"
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
      Left            =   7680
      TabIndex        =   40
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Valor Desc.:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   6600
      TabIndex        =   38
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total Título"
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
      Left            =   9840
      TabIndex        =   37
      Top             =   840
      Width           =   1650
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Situação:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor Título: "
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
      Left            =   225
      TabIndex        =   34
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   33
      Top             =   2160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00000000&
      Height          =   1335
      Left            =   45
      Top             =   720
      Width           =   11715
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Modl.:"
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
      Left            =   2460
      TabIndex        =   30
      Top             =   2220
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Dt.Baixa:"
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
      Left            =   9300
      TabIndex        =   28
      Top             =   2220
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
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
      Left            =   4320
      TabIndex        =   27
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Venc.:"
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
      Left            =   6465
      TabIndex        =   26
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NºIdent.:"
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
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Histórico:"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   3390
      Width           =   885
   End
   Begin VB.Label Label3 
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
      Left            =   3000
      TabIndex        =   23
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "frmFINGERALANC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim LANCAMENTO_ID_N  As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   cmdRecibo.Visible = False
   If CNPJ_EMPRESA_N = "15333554000188" Then _
      cmdRecibo.Visible = True

   cmbTipoRegistroAUX.Clear
   cmbTipoRegistro.Clear

   If INDR_RECEITA = 1 Then
      cmbTipoRegistro.AddItem "À Receber"
      cmbTipoRegistro.Text = "À Receber"
      frmFINGERALANC.Caption = "Manutenção em títulos Contas a Receber"
      cmbTipoRegistroAUX.Text = 1
      lblCli.Caption = "Cliente : "
      Else
         If INDR_RECEITA = 2 Then
            cmbTipoRegistro.AddItem "À Pagar"
            cmbTipoRegistro.Text = "À Pagar"
            frmFINGERALANC.Caption = "Manutenção em títulos Contas a Pagar"
            cmbTipoRegistroAUX.Text = 2
            lblCli.Caption = "Fornecedor : "
         End If
   End If
   frmFINGERALANC.Refresh
   lblCli.Refresh

   txtDtVenc.Mask = "##/##/####"
   If TITULO_N > 0 Then
      txtLanc.Text = TITULO_N
      txtLanc_KeyPress 13
   End If

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   cmbCCAux.Clear
   cmbCC.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR WITH (NOLOCK) "
   SQL = SQL & " where TIPO = 'O'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCC.AddItem Trim(TabTemp!DESCRICAO)
      cmbCCAux.AddItem TabTemp!Codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

cmbStatusLancItem.AddItem "Aberto"
cmbStatusLancItem.AddItem "Baixado"
cmbStatusLancItem.AddItem "Cancelado"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF9
         LIMPA_TUDO
         txtLanc.SetFocus
      Case vbKeyF10
         If INDR_GRAVA = True Then
            GRAVA_TITULO
            LIMPA_TUDO
            txtLanc.SetFocus
         End If
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Resize()
'On Error GoTo ERRO_TRATA

   'MODALIDADE
   cmbModalidadeAux.Clear
   cmbModalidade.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " where formapagto_id < 9999 "
   SQL = SQL & " and status = 'true' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbModalidade.AddItem TabDESCR!DESCRICAO
      cmbModalidadeAux.AddItem TabDESCR!FORMAPAGTO_ID
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   'CENTRO DE CUSTO
   cmbCCustoAux.Clear
   cmbCCusto.Clear

   SQL = "select * from DESCR WITH (NOLOCK) "
   SQL = SQL & " where TIPO = 'O' "
   SQL = SQL & " order by DESCRICAO "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbCCusto.AddItem Trim(TabDESCR!DESCRICAO)
      cmbCCustoAux.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Resize"
End Sub

Private Sub cmdForCli_Click()
   CONSULTA_PESSOA
End Sub

Private Sub cmdLanc_Click()
   Indr_Consulta = True
   frmFINCONSULTAFATURA.Show 1
   If CRITERIO_A <> "" Then
      If IsNumeric(CRITERIO_A) Then
         txtLanc.Enabled = True
         txtLanc.Text = CRITERIO_A
         txtLanc_KeyPress 13
         Exit Sub
      End If
   End If
   CRITERIO_A = ""
End Sub

Private Sub cmbmodalidade_GotFocus()
   cmbModalidade.SelStart = 0
   cmbModalidade.SelLength = Len(cmbModalidade)
   cmbModalidade.BackColor = &HC0FFFF

   If Trim(cmbModalidade.Text) = "" Then _
       cmbStatusLancItem.Text = "Aberto"

End Sub

Private Sub cmbMODALIDADE_LostFocus()
'On Error GoTo ERRO_TRATA

   If cmbModalidade.Text <> "" Then
      If Left(UCase(cmbModalidade.Text), 6) = "CHEQUE" Then
         frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = False
         frmCHEQUECADASTRO.txtPORTADOR.Text = CNPJCPF_A
         frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = True
         frmCHEQUECADASTRO.Show 1
      End If
   End If
   cmbModalidade.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lista_LostFocus()
'On Error GoTo ERRO_TRATA

   LISTA.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lista_LostFocus"
End Sub

Private Sub cmbCCusto_Click()
'On Error GoTo ERRO_TRATA

   cmbCCustoAux.ListIndex = cmbCCusto.ListIndex
   If cmbCCustoAux.Text <> "" Then
      SQL = "select * from CCUSTO WITH (NOLOCK) "
      SQL = SQL & " where codg_cc = " & cmbCCustoAux.Text
      If INDR_RECEITA = 2 Then
         SQL = SQL & " and tipo_cc = 'D' "
         Else: SQL = SQL & " and tipo_cc = 'C' "
      End If
      SQL = SQL & " order by descr_cc "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      LISTA.ListItems.Clear
      If Not TabTemp.EOF Then
         LISTA.Visible = True
         While Not TabTemp.EOF
            Set item = LISTA.ListItems.Add(, "seq." & TabTemp!Codg_cc, TabTemp!Codg_cc)
            item.SubItems(1) = TabTemp!descr_cc
            item.SubItems(2) = TabTemp!tipo_cc
            TabTemp.MoveNext
         Wend
         Else: LISTA.Visible = False
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCCusto_Click"
End Sub

Private Sub Lista_DblClick()
'On Error GoTo ERRO_TRATA

   If LISTA.SelectedItem.Text <> "" Then
      cmbAuxGridcc.Text = LISTA.SelectedItem.Text

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CCUSTO WITH (NOLOCK) "
      SQL = SQL & " where codg_cc = " & Trim(LISTA.SelectedItem.Text)
      If INDR_RECEITA = 2 Then
         SQL = SQL & " and tipo_cc = 'D' "
         Else: SQL = SQL & " and tipo_cc = 'C' "
      End If
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         txtCCustoItem.Text = TabTemp!descr_cc
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   LISTA.Visible = False
   txtDtVenc.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTA_DblClick"
End Sub

Private Sub listatitulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView ListaTitulos, ColumnHeader
End Sub

Private Sub ListaTitulos_DblClick()
On Error Resume Next
   txtSeq.Text = ListaTitulos.SelectedItem.Text
   txtSeq.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "carne"
         IMPRIME_CARNE
      Case "print"
         If txtLanc.Text <> "" Then
            FORMULA_REL = "{ITEMLANCAMENTO.numr_doc} = " & txtLanc.Text

            If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_fin01.rpt"
            frmRELATORIO10.Show 1
         End If
      Case "matar"
         Msg = "Confirma exclusão do lançamento ?"
         PERGUNTA Msg, vbYesNo + 32, "Lancamento Fincanceiro", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then _
            MATA_LANCAMENTO
      Case "consultar"
         Indr_Consulta = True
         frmFINCONSULTAFATURA.Show 1
         If CRITERIO_A <> "" Then
            If IsNumeric(CRITERIO_A) Then
               txtLanc.Enabled = True
               txtLanc.Text = CRITERIO_A
               txtLanc_KeyPress 13
               Exit Sub
            End If
         End If
         CRITERIO_A = ""
      Case "gravar"
         SQL = "update LANCAMENTO set "
         SQL = SQL & "pessoa_id = " & PESSOA_ID_N
         SQL = SQL & ",nome_pessoa = '" & Trim(Left(txtPessoa.Text, 30)) & "'"

         SQL = SQL & " where NUMR_DOC = " & txtLanc.Text
         SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         If INDR_GRAVA = True Then
            If Trim(txtLanc.Text) <> "" Then
               If IsNumeric(txtSeq.Text) Then
                  GRAVA_TITULO
                  LIMPA_TUDO
                  txtLanc.SetFocus
                  Else
                     'MsgBox "Impossivel Gravar Titulo Sem Lancamento de Sequencia!", vbExclamation, "MEGASIM
                     'Exit Sub
                     LIMPA_TUDO
                     txtLanc.SetFocus
               End If
            End If
         End If
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
         txtLanc.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCNPJCPF_LostFocus()
   txtCNPJCPF.BackColor = &HFFFFFF
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF)
   txtCNPJCPF.BackColor = &HC0FFFF

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_GotFocus"
End Sub

Private Sub TXTCGCREP_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCGCREP.PromptInclude = False
   If txtCGCREP.Text = "" Then
      txtCGCREP.Mask = "##############"
   End If
   txtCGCREP.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCREP_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CONSULTA_PESSOA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyDown"
End Sub

Private Sub TXTCGCREP_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCGCREP.PromptInclude = False
            txtCGCREP.Text = CNPJCPF_A
            txtCGCREP.PromptInclude = True
         End If
         txtCGCREP.SetFocus
   End Select
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCREP_KeyDown"
End Sub

Private Sub txtDesconto_LostFocus()
'On Error GoTo ERRO_TRATA

   CALCULA_JUROS_DESCONTO
   txtDesconto.Text = Format(txtDesconto.Text, strFormatacao2Digitos)
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_LostFocus"
End Sub

Private Sub txtJuros_GotFocus()
'On Error GoTo ERRO_TRATA

   txtJuros.SelStart = 0
   txtJuros.SelLength = Len(txtJuros)
   txtJuros.BackColor = &HC0FFFF

   If ((optPercJuros.Value = False) And (optValorJuros.Value = False)) Then
      optPercJuros.Value = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtJuros_GotFocus"
End Sub

Private Sub txtJuros_LostFocus()
'On Error GoTo ERRO_TRATA

   CALCULA_JUROS_DESCONTO
   txtJuros.Text = Format(txtJuros.Text, strFormatacao2Digitos)
   txtJuros.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtJuros_LostFocus"
End Sub

Private Sub txtLanc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If IsNumeric(txtLanc.Text) Then
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select tipo_lancamento from LANCAMENTO WITH (NOLOCK) "
         SQL = SQL & " where numr_doc = " & txtLanc.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then
            If INDR_RECEITA <> TabLancamento.Fields(0).Value Then
               If TabLancamento.State = 1 Then _
                  TabLancamento.Close
               MsgBox "Não permitido, título em uso."
               txtLanc.Text = ""
               Exit Sub
            End If
         End If

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         MOSTRA_LANCAMENTO
      End If

      If TITULO_N = 0 Then
         txtNOTA.SetFocus
         Else: TITULO_N = 0
      End If

      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLanc_KeyPress"
End Sub

Private Sub txtLanc_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtLanc.Text = "" Then
      GERA_NUMERO_LANCAMENTO
      txtLanc.Text = NUMR_LANCAMENTO_N
   End If
   INDR_GRAVA = True
   txtLanc.Enabled = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLanc_LostFocus"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEmis.PromptInclude = False
   If txtDtEmis.Text = "" Then
      txtDtEmis.Mask = "##/##/####"
      Else
         txtDtEmis.PromptInclude = True
         If Not IsDate(txtDtEmis.Text) Then _
            txtDtEmis.Mask = "##/##/####"
   End If
   txtDtEmis.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_GotFocus"
End Sub

Private Sub txtDTEMIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            'Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
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
      txtDtEmis.Mask = "##/##/####"
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

   txtDtVenc.SelStart = 0
   txtDtVenc.SelLength = Len(txtDtVenc)
   txtDtVenc.BackColor = &HC0FFFF

   txtDtVenc.PromptInclude = False
   If txtDtVenc.Text = "" Then
      txtDtVenc.Mask = "##/##/####"
      Else
         txtDtVenc.PromptInclude = True
         If Not IsDate(txtDtVenc.Text) Then _
            txtDtVenc.Mask = "##/##/####"
   End If
   txtDtVenc.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_GotFocus"
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtBaixa.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_KeyPress"
End Sub

Private Sub txtDTVENC_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtVenc.PromptInclude = True
   txtDtEmis.PromptInclude = True
   If IsDate(txtDtVenc.Text) And IsDate(txtDtEmis.Text) Then
      If DMA(txtDtVenc.Text, "i") < DMA(txtDtEmis.Text, "f") Then
         'MsgBox "Data de vencimento menor que data de emissão,não permitido."
         'txtDtEmis.SetFocus
         'Exit Sub
      End If
   End If
   txtDtVenc.PromptInclude = False
   If txtDtVenc.Text = "" Then
      txtDtVenc.Mask = "##/##/####"
      txtDtVenc.PromptInclude = False
         txtDtVenc.Text = Date
      txtDtVenc.PromptInclude = True
   End If
   txtDtVenc.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_LostFocus"
End Sub

Private Sub txtDtBaixa_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtBaixa.SelStart = 0
   txtDtBaixa.SelLength = Len(txtDtBaixa)
   txtDtBaixa.BackColor = &HC0FFFF

   txtDtBaixa.PromptInclude = False
   If txtDtBaixa.Text = "" Then
      txtDtBaixa.Mask = "##/##/####"
      Else
         txtDtBaixa.PromptInclude = True
         If Not IsDate(txtDtBaixa.Text) Then _
            txtDtBaixa.Mask = "##/##/####"
   End If
   txtDtBaixa.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtBaixa_GotFocus"
End Sub

Private Sub txtDtBaixa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorItem.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtBaixa_KeyPress"
End Sub

Private Sub txtDtdesc_GotFocus()
'On Error GoTo ERRO_TRATA

   txtdtdesc.PromptInclude = False
   If txtdtdesc.Text = "" Then
      txtdtdesc.Mask = "##/##/####"
      Else
         txtdtdesc.PromptInclude = True
         If Not IsDate(txtdtdesc.Text) Then _
            txtdtdesc.Mask = "##/##/####"
   End If
   txtdtdesc.PromptInclude = True
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtdesc_GotFocus"
End Sub

Private Sub txtDtdesc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCheque.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtdesc_KeyPress"
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCheque.Text <> "" Then _
         GRAVA_CHEQUE
      txtHistorico.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCheque_KeyPress"
End Sub

Private Sub txtDtbaixa_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtVenc.PromptInclude = True
   txtDtBaixa.PromptInclude = True
   If IsDate(txtDtBaixa.Text) And IsDate(txtDtVenc.Text) Then
   If Not IsDate(txtDtEmis.Text) Then
      MsgBox "Conferir Data de Emissao do Titulo!", vbCritical, "MEGASIM"
      Exit Sub
   End If
   'If CDate(txtDtBaixa.Text) < CDate(txtDtEmis.Text) Then
   '   MsgBox "Data de baixa menor que data de vencimento,não permitido."
   '   txtDtVenc.SetFocus
   '   Exit Sub
   'End If
   End If
   
   txtDtBaixa.PromptInclude = False
   If txtDtBaixa.Text <> "" Then
      txtDtBaixa.PromptInclude = True
      If Not IsDate(txtDtBaixa.Text) Then
         txtDtBaixa.Mask = "##/##/####"
         txtDtBaixa.PromptInclude = False
            txtDtBaixa.Text = ""
         txtDtBaixa.PromptInclude = True
      End If
   End If
   If IsDate(txtDtBaixa.Text) Then _
      cmbStatusLancItem.Text = "Baixado"
   txtDtVenc.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtbaixa_LostFocus"
End Sub

Private Sub cmbTipoRegistro_Click()
'On Error GoTo ERRO_TRATA

   cmbTipoRegistroAUX.ListIndex = cmbTipoRegistro.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoRegistro_Click"
End Sub

Private Sub cmbTipoRegistro_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoRegistro_KeyPress"
End Sub

Private Sub txtNumrDoc_GotFocus()
'On Error GoTo ERRO_TRATA

   txtNumrDoc.SelStart = 0
   txtNumrDoc.SelLength = Len(txtNumrDoc)
   txtNumrDoc.BackColor = &HC0FFFF

   If txtNumrDoc.Text = "" Then
      If txtNOTA.Text <> "" Then
         txtNumrDoc.Text = Trim(txtNOTA.Text) & "-" & Trim(TIPO_DP_EMPRESA)
         Else
            If txtLanc.Text <> "" Then _
               txtNumrDoc.Text = Trim(txtLanc.Text) & "-" & Trim(TIPO_DP_EMPRESA)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNumrDoc_GotFocus"
End Sub

Private Sub txtnumrdoc_LostFocus()
   txtNumrDoc.BackColor = &HFFFFFF
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtNOTA.Text <> "" Then
         If INDR_RECEITA = 1 Then
            If TabNOTA.State = 1 Then _
               TabNOTA.Close

            SQL = "select * from NF "
            SQL = SQL & " where numr_nota = " & txtNOTA.Text
            TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabNOTA.EOF Then
               INDR_GRAVA = True
               If Not IsNull(TabNOTA!PEDIDO_ID) Then _
                  txtLanc.Text = TabNOTA!PEDIDO_ID
               If Not IsNull(TabNOTA!NUMR_NOTA) Then _
                  txtNOTA.Text = TabNOTA!NUMR_NOTA

               If IsNumeric(txtLanc.Text) Then _
                  MOSTRA_LANCAMENTO

               txtDtEmis.SetFocus
               Else: MsgBox "Nota não encontrada."
            End If
         Else
            If TabNOTA.State = 1 Then _
               TabNOTA.Close

            FORNEC_ID_N = 0
            SQL = "select fornecedor_id from vwFornecedor "
            SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
            TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabNOTA.EOF Then _
               If Not IsNull(TabNOTA.Fields(0).Value) Then _
                  FORNEC_ID_N = TabNOTA.Fields(0).Value
            If TabNOTA.State = 1 Then _
               TabNOTA.Close

            SQL = "select * from NOTAENTRADA "
            SQL = SQL & " where numr_nota = " & txtLanc.Text
            'SQL = SQL & " where numr_nota = " & txtNota.Text
            SQL = SQL & " and estabelecimento_id = " & EMPRESA_ID_N
            SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
            TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabNOTA.EOF Then
               INDR_GRAVA = True
               If Not IsNull(TabNOTA!numr_pedido_compra) Then _
                  txtLanc.Text = TabNOTA!numr_pedido_compra
               If Not IsNull(TabNOTA!NUMR_NOTA) Then _
                  txtNOTA.Text = TabNOTA!NUMR_NOTA
               If IsNumeric(txtLanc.Text) Then _
                  MOSTRA_LANCAMENTO
               txtDtEmis.SetFocus
               Else: MsgBox "Nota não encontrada."
            End If
         End If

         On Error Resume Next

         If TabNOTA.State = 1 Then _
            TabNOTA.Close
      End If
      txtDtEmis.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnota_KeyPress"
End Sub

Private Sub cmbcc_GotFocus()
   cmbCC.SelStart = 0
   cmbCC.SelLength = Len(cmbCC)
   cmbCC.BackColor = &HC0FFFF
End Sub

Private Sub cmbcc_LostFocus()
   cmbCC.BackColor = &HFFFFFF
End Sub

Private Sub cmbCC_Click()
'On Error GoTo ERRO_TRATA

   cmbCCAux.ListIndex = cmbCC.ListIndex
   txtHistorico.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub cmdRecibo_Click()
'On Error GoTo ERRO_TRATA

'abrir para enviar via email
   If Trim(txtLanc.Text) <> "" And Trim(txtSeq.Text) <> "" Then
      FORMULA_REL = "{ITEMLANCAMENTO.numr_doc} = " & txtLanc.Text
      FORMULA_REL = FORMULA_REL & " and {ITEMLANCAMENTO.seq} = " & txtSeq.Text

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "reciboshf.rpt"
      Nome_Relatorio = "recibo.rpt"
      frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdRecibo_Click"
End Sub

Private Sub cmdMata_Click()
   MATA_ITEM
End Sub

Private Sub txtseq_GotFocus()
   txtSeq.SelStart = 0
   txtSeq.SelLength = Len(txtSeq)
   txtSeq.BackColor = &HC0FFFF
End Sub

Private Sub txtseq_LostFocus()
   txtSeq.BackColor = &HFFFFFF
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         MATA_ITEM
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyDown"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtSeq.Text) = "" And Trim(txtLanc.Text) <> "" Then
         NUMR_SEQ_N = 0 & MAX_ID("seq", "itemlancamento", "numr_doc", Trim(txtLanc.Text), "", "")
         txtSeq.Text = "" & NUMR_SEQ_N
      End If

      MOSTRA_ITEM

      cmbModalidade.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyPress"
End Sub

Private Sub cmbMODALIDADE_Click()
'On Error GoTo ERRO_TRATA

   cmbModalidadeAux.ListIndex = cmbModalidade.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub cmbMODALIDADE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(cmbModalidade.Text) <> "" Then _
         txtDtVenc.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_KeyPress"
End Sub

Private Sub cmbccusto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCGCREP.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbccusto_KeyPress"
End Sub

Private Sub txtValorItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtValorItem.Text) <> "" Then _
         optPercJuros.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_KeyPress"
End Sub

Private Sub optPercdesc_Click()
'On Error GoTo ERRO_TRATA

   txtDesconto.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optPercdesc_Click"
End Sub

Private Sub optValordesc_Click()
'On Error GoTo ERRO_TRATA

   txtDesconto.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optValordesc_Click"
End Sub

Private Sub optPercdesc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDesconto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optPercdesc_KeyPress"
End Sub

Private Sub optValordesc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDesconto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optValordesc_KeyPress"
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbStatusLancItem.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_KeyPress"
End Sub

Private Sub optPercjuros_Click()
'On Error GoTo ERRO_TRATA

   txtJuros.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optPercjuros_Click"
End Sub

Private Sub optValorjuros_Click()
'On Error GoTo ERRO_TRATA

   txtJuros.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optValorjuros_Click"
End Sub

Private Sub optPercjuros_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtJuros.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optPercjuros_KeyPress"
End Sub

Private Sub optValorjuros_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtJuros.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optValorjuros_KeyPress"
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      optPercDesc.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtjuros_KeyPress"
End Sub

Private Sub cmbStatusLancItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNumrDoc.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbModalidade_KeyPress"
End Sub

Private Sub txtnumrdoc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnumrdoc_KeyPress"
End Sub

Private Sub cmbCC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtHistorico.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCC_KeyPress"
End Sub

Private Sub cmbStatusLancItem_GotFocus()
   cmbStatusLancItem.SelStart = 0
   cmbStatusLancItem.SelLength = Len(cmbStatusLancItem)
   cmbStatusLancItem.BackColor = &HC0FFFF
End Sub

Private Sub cmbStatusLancItem_LostFocus()
'On Error GoTo ERRO_TRATA

   If Left(cmbStatusLancItem.Text, 1) <> "C" Then
      If cmbStatusLancItem.Text = "Baixado" Then
         txtDtBaixa.PromptInclude = True
         If Not IsDate(txtDtBaixa.Text) Then
            MsgBox "Data da baixa inválida."
            txtDtBaixa.PromptInclude = False
               txtDtBaixa.Text = Date
            txtDtBaixa.PromptInclude = True
            Exit Sub
         End If
      End If
      txtDtBaixa.PromptInclude = True
      If IsDate(txtDtBaixa.Text) Then _
         cmbStatusLancItem.Text = "Baixado"
   End If
   cmbStatusLancItem.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbModalidade_LostFocus"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False

      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         If Len(txtCNPJCPF.Text) > 0 Then
            Select Case Len(txtCNPJCPF.Text)
               Case Is = 11
                  If Not CALCULACPF(txtCNPJCPF.Text) Then
                     MsgBox "CPF com DV incorreto !!!"
                     txtCNPJCPF.PromptInclude = False
                     txtCNPJCPF = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
                   End If
               Case Is = 14
                  If Not VALIDACGC(txtCNPJCPF.Text) Then
                     MsgBox "CNPJ com DV incorreto !!! "
                     txtCNPJCPF.PromptInclude = False
                     txtCNPJCPF = ""
                     txtCNPJCPF.SetFocus
                      Exit Sub
                  End If
               Case Is > 14
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCNPJCPF = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
               Case Is < 11
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCNPJCPF = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End Select
            Else
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCNPJCPF = ""
               txtCNPJCPF.SetFocus
               Exit Sub
         End If
         txtCNPJCPF.PromptInclude = False
         PESSOA_ID_N = 0
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         SQL = "select pessoa_id,descricao,razao from PESSOA WITH (NOLOCK) "
         SQL = SQL & " where cnpjcpf ='" & Trim(txtCNPJCPF.Text) & "'"
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then
            PESSOA_ID_N = 0 & TabPessoa.Fields("pessoa_id").Value
            txtPessoa.Text = Trim(TabPessoa.Fields("descricao").Value)
         End If
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         If Trim(txtLanc.Text) <> "" Then
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            SQL = "select cliente from OS WITH (NOLOCK) "
            SQL = SQL & " where os_id = " & txtLanc.Text
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then _
               txtPessoa.Text = Trim(TabPessoa.Fields("cliente").Value)
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
         End If
      End If
      txtCNPJCPF.PromptInclude = False

      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         PESSOA_ID_N = 0 & TRAZ_ID_TABELA("PESSOA", "PESSOA_ID", "CNPJCPF", Trim(txtCNPJCPF.Text))
         If TabPessoa.State = 1 Then _
            TabPessoa.Close
         SQL = "select nome_pessoa from LANCAMENTO WITH (NOLOCK) "
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then _
            txtPessoa.Text = Trim(TabPessoa.Fields(0).Value)
         If TabPessoa.State = 1 Then _
            TabPessoa.Close
      End If

      If txtCNPJCPF.Text <> "" Then
         CRITERIO_A = txtCNPJCPF.Text
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
         End If
      End If
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         txtPessoa.SetFocus
         Else: txtSeq.SetFocus
      End If
      txtCNPJCPF.PromptInclude = True
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub txtCGCREP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCGCREP.PromptInclude = False
      If txtCGCREP.Text = "" Then
         txtHistorico.SetFocus
         Else: txtdtdesc.SetFocus
      End If
        txtCGCREP.PromptInclude = False
        If txtCGCREP.Text <> "" Then
           If Len(txtCGCREP.Text) > 0 Then
              Select Case Len(txtCGCREP.Text)
                 Case Is = 11
                    If Not CALCULACPF(txtCGCREP.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtCGCREP.PromptInclude = False
                       txtCGCREP = ""
                       txtCGCREP.SetFocus
                       Exit Sub
                     End If
                 Case Is = 14
                    If Not VALIDACGC(txtCGCREP.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtCGCREP.PromptInclude = False
                       txtCGCREP = ""
                       txtCGCREP.SetFocus
                        Exit Sub
                    End If
                 Case Is > 14
                    MsgBox "CNPJ/CPF com DV incorreto !!! "
                    txtCGCREP = ""
                    txtCGCREP.SetFocus
                    Exit Sub
                 Case Is < 11
                    MsgBox "CNPJ/CPF com DV incorreto !!! "
                    txtCGCREP = ""
                    txtCGCREP.SetFocus
                    Exit Sub
              End Select
              Else
                 MsgBox "CNPJ/CPF com DV incorreto !!! "
                 txtCGCREP = ""
                 txtCGCREP.SetFocus
                 Exit Sub
           End If
           txtCGCREP.PromptInclude = False

           If TabAUX.State = 1 Then _
              TabAUX.Close

           SQL = "select * from vwFornecedor WITH (NOLOCK) "
           SQL = SQL & " where cnpjcpf = '" & txtCGCREP.Text & "'"
           TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
           If Not TabAUX.EOF Then
              txtrep.Text = TabAUX!NOME
              Else
                  MsgBox "Fornecedor não cadastrado."
                  txtCGCREP.SetFocus
           End If
           If TabAUX.State = 1 Then _
              TabAUX.Close
        End If
        txtCGCREP.PromptInclude = False
        If txtCGCREP.Text <> "" Then
           CRITERIO_A = txtCGCREP.Text
           If Not IsNull(txtCGCREP.Text) Then
              If Len(txtCGCREP.Text) <= 11 Then
                 txtCGCREP.Mask = "###.###.###-##"
                 Else: txtCGCREP.Mask = "##.###.###/####-##"
              End If
           End If
        End If
      txtCGCREP.PromptInclude = True
      txtdtdesc.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCGCREP_KeyPress"
End Sub

Private Sub txtNota_GotFocus()
   txtNOTA.SelStart = 0
   txtNOTA.SelLength = Len(txtNOTA)
   txtNOTA.BackColor = &HC0FFFF
End Sub

Private Sub txtNOTA_LostFocus()
   txtNOTA.BackColor = &HFFFFFF
End Sub

Private Sub txtpessoa_GotFocus()
   txtPessoa.SelStart = 0
   txtPessoa.SelLength = Len(txtPessoa)
   txtPessoa.BackColor = &HC0FFFF
End Sub

Private Sub txtpessoa_LostFocus()
   txtPessoa.BackColor = &HFFFFFF
End Sub

Private Sub TXTPESSOA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtHistorico.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPESSOA_KeyPress"
End Sub

Private Sub txthistorico_GotFocus()
   txtHistorico.SelStart = 0
   txtHistorico.SelLength = Len(txtHistorico)
   txtHistorico.BackColor = &HC0FFFF
End Sub

Private Sub txthistorico_LostFocus()
   txtHistorico.BackColor = &HFFFFFF
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtLanc.Text) = "" Then
         MsgBox "Número de lançamento deve ser informado ou gerado."
         txtLanc.SetFocus
         Exit Sub
      End If
      txtDtEmis.PromptInclude = True
      If Not IsDate(txtDtEmis.Text) Then
         MsgBox "Informe data emissão válida."
         txtDtEmis.SetFocus
         Exit Sub
      End If
      If txtSeq.Text = "" Then
         MsgBox "Informe seqüência do lançamento."
         txtSeq.SetFocus
         Exit Sub
      End If
      If cmbModalidadeAux.Text = "" Then
         MsgBox "Informe modalidade do lançamento."
         cmbModalidade.SetFocus
         Exit Sub
      End If
      txtDtVenc.PromptInclude = True
      If Not IsDate(txtDtVenc.Text) Then
         MsgBox "Informe data de vencimento do lançamento."
         txtDtVenc.SetFocus
         Exit Sub
      End If
      'txtDTBaixa.PromptInclude = True
      'If Not IsDate(txtDTBaixa.Text) Then
      'End If
      If cmbModalidade.Text = "" Then
         MsgBox "Informe status do lançamento."
         cmbModalidade.SetFocus
         Exit Sub
      End If
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         MsgBox "Informe cliente."
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
      If txtDesconto.Text <> "" Then
         VALOR_DESCONTO_N = txtDesconto.Text
         Else: VALOR_DESCONTO_N = 0
      End If

      If Trim(txtValorItem.Text) = "" Then
         MsgBox "Informe valor do lançamento."
         txtValorItem.SetFocus
         Exit Sub
      End If
      If Not IsNumeric(txtValorItem.Text) Then
         MsgBox "Informe valor do lançamento."
         txtValorItem.SetFocus
         Exit Sub
      End If
      VALOR_ITEM_N = txtValorItem.Text
      If VALOR_ITEM_N <= 0 Then
         MsgBox "Informe valor do lançamento."
         txtValorItem.SetFocus
         Exit Sub
      End If

      'VALOR_ITEM_N = VALOR_ITEM_N - VALOR_DIFERENCA_N + VLR_DESCT_DIF_N
      'VALOR_ITEM_N = VALOR_ITEM_N - VALOR_DESCONTO_N + VLR_DESCT_DIF_N
      'VALOR_ITEM_N = VALOR_ITEM_N

      VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N - VALOR_DIFERENCA_N
      VALOR_TOTAL_DESCONTO_N = VALOR_TOTAL_DESCONTO_N + VALOR_DESCONTO_N - VLR_DESCT_DIF_N
      txtValorTotal.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtValorTotal.Refresh
      txtToTDesc.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      txtToTDesc.Refresh

      GRAVA_OBS

      'Msg = "Deseja gravar lancamento Definitivo?"
      'PERGUNTA Msg, vbYesNo + 32, "Hitorico Lancamentos", "DEMO.HLP", 1000
      'If RESPOSTA = vbYes Then
         GRAVA_TITULO
         SETA_GRID
         LIMPA_BODY

         txtSeq.SetFocus
      '   Else
      '       SETA_GRID
      '       LIMPA_BODY
      '       txtSeq.SetFocus
      'End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTHISTORICO_KeyPress"
End Sub

Private Sub txtValorItem_GotFocus()
   txtValorItem.SelStart = 0
   txtValorItem.SelLength = Len(txtValorItem)
   txtValorItem.BackColor = &HC0FFFF
End Sub

Private Sub txtValorItem_LostFocus()
'On Error GoTo ERRO_TRATA

   CALCULA_JUROS_DESCONTO
   txtValorItem.Text = Format(txtValorItem.Text, strFormatacao2Digitos)
   txtValorItem.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_LostFocus"
End Sub

Private Sub txtValorTotal_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   KeyAscii = 0
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorTotal_KeyPress"
End Sub
'====================================
Private Sub GERA_NUMERO_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   NUMR_LANCAMENTO_N = 1
   GERA_PEDIDO_ID
   NUMR_LANCAMENTO_N = PEDIDO_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NUMERO_LANCAMENTO"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   ListaTitulos.ListItems.Clear
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtValorTotal.Text = ""
   cmbAuxGridcc.Text = ""
   txtToTDesc.Text = ""
   txtLanc.Enabled = True
   txtLanc.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtEmis.Text = ""
   txtPessoa.Text = ""
   LIMPA_BODY
   VALOR_TOTAL_N = 0
   VALOR_ITEM_N = 0
   VALOR_DIFERENCA_N = 0
   VLR_DESCT_DIF_N = 0
   txtNOTA.Text = ""
   INDR_GRAVA = False
   PESSOA_ID_N = 0
   LANCAMENTO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   cmbCCAux.Text = ""
   cmbCC.Text = ""
   cmbCCustoAux.Text = ""
   cmbCCusto.Text = ""
   txtCCustoItem.Text = ""
   cmbAuxGridcc.Text = ""
   txtTotalItem.Text = ""
   txtJuros.Text = ""
   lblCob.Caption = ""
   txtNumrDoc.Text = ""
   txtSeq.Text = ""
   cmbModalidade.Text = ""
   cmbModalidadeAux.Text = ""
   txtDtVenc.PromptInclude = False
   txtDtVenc.Text = ""
   txtDtBaixa.PromptInclude = False
   txtDtBaixa.Text = ""
   txtValorItem.Text = ""
   txtDesconto.Text = ""
   cmbStatusLancItem.Text = ""
   txtHistorico.Text = ""
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   VALOR_DIFERENCA_N = 0
   VLR_DESCT_DIF_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub GRAVA_OBS()
'On Error GoTo ERRO_TRATA

   'ITENS
   txtDtBaixa.PromptInclude = True
   
   PERC_DESCONTO_N = 0
   If txtDesconto.Text <> "" Then
      VALOR_DESCONTO_N = txtDesconto.Text
      If VALOR_DESCONTO_N > 0 Then
         If optValorDesc.Value = True Then
            PERC_DESCONTO_N = Format(((VALOR_DESCONTO_N / VALOR_ITEM_N) * 100), strFormatacao2Digitos)
            Else: PERC_DESCONTO_N = txtDesconto.Text
         End If
      End If
   End If
   PERC_JUROS_N = 0
   If txtJuros.Text <> "" Then
      VALOR_DESCONTO_N = txtJuros.Text
      If VALOR_DESCONTO_N > 0 Then
         If optValorJuros.Value = True Then
            PERC_JUROS_N = Format(((VALOR_DESCONTO_N / VALOR_ITEM_N) * 100), strFormatacao2Digitos)
            Else: PERC_JUROS_N = txtJuros.Text
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_OBS"
End Sub

Private Sub GRAVA_TITULO()
'On Error GoTo ERRO_TRATA

   VALOR_ITEM_N = 0 & txtValorItem.Text
   If VALOR_ITEM_N <= 0 Then
      MsgBox "Informar valor corretamente."
      txtValorItem.SetFocus
      Exit Sub
   End If

   If Trim(cmbModalidadeAux.Text) = "" Then _
      cmbModalidadeAux.Text = 1
   If Not IsNumeric(txtCCustoItem.Text) Then _
      txtCCustoItem.Text = 0
   If Not IsNumeric(txtDesconto.Text) Then _
      txtDesconto.Text = 0
   If Not IsNumeric(cmbCCustoAux.Text) Then _
      cmbCCustoAux.Text = 0
   If Not IsNumeric(txtJuros.Text) Then _
      txtJuros.Text = 0
   If Trim(cmbCCAux.Text) = "" Then _
      cmbCCAux.Text = "NULL"

   NUMR_ID_N = 0
   If PESSOA_ID_N <= 0 Then _
      MsgBox "Pessoa não vinculada ao título !!!"

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from LANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " where NUMR_DOC = " & txtLanc.Text
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      txtCNPJCPF.PromptInclude = False
      txtDtEmis.PromptInclude = True
      NUMR_ID_N = TabLancamento!LANCAMENTO_ID

      SQL = "update LANCAMENTO set "
      SQL = SQL & "pessoa_id = " & PESSOA_ID_N
      SQL = SQL & ",nome_pessoa = '" & Trim(Left(txtPessoa.Text, 30)) & "'"

      SQL = SQL & " where NUMR_DOC = " & txtLanc.Text
      SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      CONECTA_RETAGUARDA.Execute SQL
      Else
         NUMR_ID_N = MAX_ID("lancamento_ID", "lancamento", "", "", "", "")

         SQL = "INSERT INTO LANCAMENTO "
         SQL = SQL & " ("
            SQL = SQL & " lancamento_id, numr_doc, dt_cad, "
            SQL = SQL & " TIPO_LANCAMENTO, tipovenda_id,pessoa_id,estabelecimento_id,nome_pessoa ) "
         SQL = SQL & " VALUES ("
            SQL = SQL & NUMR_ID_N                           'lancamento_id
            SQL = SQL & "," & txtLanc.Text                  'numr_doc
            SQL = SQL & ",'" & DMA(txtDtEmis.Text) & "'"    'dt_cad
            SQL = SQL & "," & INDR_RECEITA                  'TIPO_LANCAMENTO
            SQL = SQL & "," & cmbTipoRegistroAUX.Text       'tipovenda_id
            SQL = SQL & "," & PESSOA_ID_N                   'pessoa_id
            SQL = SQL & "," & ESTABELECIMENTO_ID_N          'estabelecimento_id
            SQL = SQL & ",'" & Trim(Left(txtPessoa.Text, 30)) & "'"
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   Dim Dt_Baixa_Titulo As String

   Dt_Baixa_Titulo = ""

   txtDtBaixa.PromptInclude = False
   If Trim(txtDtBaixa.Text) <> "" Then
      txtDtBaixa.PromptInclude = True
      If IsDate(txtDtBaixa.Text) Then _
         Dt_Baixa_Titulo = txtDtBaixa.Text
      Dt_Baixa_Titulo = Format(Dt_Baixa_Titulo, "dd/mm/yyyy")
   End If

   cmbModalidade.Refresh

  'ITENS
   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " where seq = " & txtSeq.Text
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabAUX.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
         SQL = SQL & " Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
         SQL = SQL & ", Status = '" & Left(cmbStatusLancItem.Text, 1) & "'"
         SQL = SQL & ", formapagto_id = " & cmbModalidadeAux.Text
         SQL = SQL & ", DT_VENCIMENTO = '" & DMA(txtDtVenc.Text) & "'"

         If Trim(Dt_Baixa_Titulo) <> "" Then
            SQL = SQL & ", DT_baixa = '" & Dt_Baixa_Titulo & "'"
            Else: SQL = SQL & ", DT_baixa = null"
         End If

         SQL = SQL & ", PERC_DESCONTO = " & tpMOEDA(txtDesconto.Text)
         SQL = SQL & ", PERC_JUROS = " & tpMOEDA(txtJuros.Text)
         SQL = SQL & ", CODG_USU_BAIXA = " & 0
         SQL = SQL & ", NUMR_DP = '" & Trim(TabAUX!NUMR_DP) & "'"
         SQL = SQL & ", GRP = " & 0
         SQL = SQL & ", ItemGRP = " & 0
         SQL = SQL & ", cc_id = " & Trim(cmbCCAux.Text)
         SQL = SQL & ", historico = '" & Trim(txtHistorico.Text) & "'"
      SQL = SQL & " where lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and seq = " & TabAUX!SEQ
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
         SQL = SQL & " (lancamento_id, numr_doc, seq, Valor_Item, Status, formapagto_id, "
         SQL = SQL & " DT_VENCIMENTO, PERC_DESCONTO, PERC_JUROS, NUMR_DP, GRP, itemGRP, "
         SQL = SQL & " usu_cad, DT_CAD,acerto,cc_id,historico)"
         SQL = SQL & " VALUES ("
            SQL = SQL & NUMR_ID_N                                       'lancamento_id
            SQL = SQL & "," & txtLanc.Text                              'numr_doc
            SQL = SQL & "," & txtSeq.Text                               'seq
            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                     'Valor_Item
            SQL = SQL & ",'" & Left(cmbStatusLancItem.Text, 1) & "'"    'Status
            SQL = SQL & "," & cmbModalidadeAux.Text                     'formapagto_id
            SQL = SQL & ",'" & DMA(txtDtVenc.Text) & "'"                'DT_VENCIMENTO
            SQL = SQL & "," & tpMOEDA(txtDesconto.Text)                 'PERC_DESCONTO
            SQL = SQL & "," & tpMOEDA(txtJuros.Text)                    'PERC_JUROS
            SQL = SQL & ",'" & txtNumrDoc.Text & "'"                    'NUMR_DP
            SQL = SQL & "," & cmbCCustoAux.Text                         'GRP
            SQL = SQL & "," & txtCCustoItem.Text                        'itemGRP
            SQL = SQL & "," & USUARIO_ID_N                              'usu_cad
            SQL = SQL & ",'" & Now & "'"                                'DT_CAD
            SQL = SQL & "," & 0                                         'acerto
            SQL = SQL & "," & Trim(cmbCCAux.Text)                       'cc_id
            SQL = SQL & ",'" & Trim(txtHistorico.Text) & "'"            'historico
         SQL = SQL & ")"
   End If
   If TabAUX.State = 1 Then _
      TabAUX.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TITULO"
End Sub

Private Sub MOSTRA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   Dim TabLocal   As New ADODB.Recordset

   LANCAMENTO_ID_N = 0
   VALOR_TOTAL_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   txtToTDesc.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

   If TabLocal.State = 1 Then _
      TabLocal.Close

   SQL = "select LANCAMENTO.*, PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO"
   SQL = SQL & " from LANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where numr_doc = " & txtLanc.Text
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabLocal.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLocal.EOF Then
      LANCAMENTO_ID_N = 0 & TabLocal.Fields("lancamento_id").Value
      PESSOA_ID_N = 0 & TabLocal.Fields("pessoa_id").Value
      If Trim(txtPessoa.Text) = "" Then _
         txtPessoa.Text = Trim(TabLocal.Fields("descricao").Value)

      If Trim(txtLanc.Text) <> "" Then
         If TabPessoa.State = 1 Then _
            TabPessoa.Close
         SQL = "select cliente from OS WITH (NOLOCK) "
         SQL = SQL & " where os_id = " & txtLanc.Text
         'TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         'If Not TabPessoa.EOF Then _
            txtPessoa.Text = Trim(TabPessoa.Fields("cliente").Value)
         'If TabPessoa.State = 1 Then _
            TabPessoa.Close
      End If

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = TabLocal.Fields("cnpjcpf").Value

      If Not IsNull(TabLocal!TIPO_LANCAMENTO) Then
         cmbTipoRegistro.Text = Trim(TRAZ_DESCRITOR("B", TabLocal!TIPO_LANCAMENTO))
         cmbTipoRegistroAUX.Text = TabLocal!TIPO_LANCAMENTO
      End If

      VALOR_TOTAL_N = 0 & TRAZ_VALOR_ITEMLANCAMENTO(LANCAMENTO_ID_N)
      txtValorTotal.Text = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtDtEmis.Text = "" & TabLocal!DT_CAD

      SETA_GRID
   End If
   If TabLocal.State = 1 Then _
      TabLocal.Close
   If txtLanc.Text <> "" Then
      If INDR_RECEITA = 1 Then
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         SQL = "select numr_nota from NF "
         SQL = SQL & " INNER JOIN PEDIDONF "
         SQL = SQL & " ON NF.NF_ID = PEDIDONF.NF_ID"
         SQL = SQL & " where pedido_id = " & txtLanc.Text
         
         TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabNOTA.EOF Then
            If Not IsNull(TabNOTA!NUMR_NOTA) Then _
               txtNOTA.Text = TabNOTA!NUMR_NOTA
         End If
         If TabNOTA.State = 1 Then _
            TabNOTA.Close
         Else
            If TabNOTA.State = 1 Then _
               TabNOTA.Close

            SQL = "select * from NOTAENTRADA "
               SQL = SQL & " where numr_nota = " & txtLanc.Text
               SQL = SQL & " and estabelecimento_id = " & EMPRESA_ID_N
               SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
            TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabNOTA.EOF Then _
               If Not IsNull(TabNOTA!NUMR_NOTA) Then _
                  txtNOTA.Text = TabNOTA!NUMR_NOTA
            If TabNOTA.State = 1 Then _
               TabNOTA.Close
      End If
   End If
   
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_LANCAMENTO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim strSQL As String
   ListaTitulos.ListItems.Clear

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   strSQL = "select CNPJCPF, DESCRICAO, RAZAO, LANCAMENTO.LANCAMENTO_ID, LANCAMENTO.PESSOA_ID, LANCAMENTO.NUMR_DOC, TIPO_LANCAMENTO, LANCAMENTO.dt_cad, "
   strSQL = strSQL & " tipovenda_id, ESTABELECIMENTO_ID, SEQ, FORMAPAGTO_ID, VALOR_ITEM, ITEMLANCAMENTO.STATUS,PERC_JUROS,PERC_DESCONTO,"
   strSQL = strSQL & " DT_VENCIMENTO , DT_BAIXA, DT_CANCELA, Valor_Desconto, CC_ID, HISTORICO,NUMR_DP"

   strSQL = strSQL & " from LANCAMENTO WITH (NOLOCK) "
   strSQL = strSQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   strSQL = strSQL & " ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID "
   strSQL = strSQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
   strSQL = strSQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

   strSQL = strSQL & " where itemLANCAMENTO.numr_doc = " & txtLanc.Text
   strSQL = strSQL & " and LANCAMENTO.tipo_lancamento = " & INDR_RECEITA
   strSQL = strSQL & " order by seq desc "
   TabLANCAMENTOITEM.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLANCAMENTOITEM.EOF
      Set item = ListaTitulos.ListItems.Add(, "seq." & TabLANCAMENTOITEM.Fields("SEQ").Value, TabLANCAMENTOITEM.Fields("SEQ").Value)
      PERC_JUROS_N = 0 & TabLANCAMENTOITEM!PERC_JUROS
      PERC_DESCONTO_N = 0 & TabLANCAMENTOITEM!PERC_DESCONTO

      item.SubItems(1) = "" & Trim(TabLANCAMENTOITEM!NUMR_DP)
      item.SubItems(2) = "" & TRAZ_DESCRICAO_FORMAPAGTO(TabLANCAMENTOITEM!FORMAPAGTO_ID)
      item.SubItems(3) = "" & Format(TabLANCAMENTOITEM!Valor_Item, strFormatacao2Digitos)
      item.SubItems(4) = "" & Format(TabLANCAMENTOITEM!Valor_Item * PERC_JUROS_N / 100, strFormatacao2Digitos)
      item.SubItems(5) = "" & Format(TabLANCAMENTOITEM!Valor_Item * PERC_DESCONTO_N / 100, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(TabLANCAMENTOITEM!Valor_Item + (TabLANCAMENTOITEM!Valor_Item * PERC_JUROS_N / 100) - (TabLANCAMENTOITEM!Valor_Item * PERC_DESCONTO_N / 100), strFormatacao2Digitos)
      item.SubItems(9) = "Aberto"

      If Not IsNull(TabLANCAMENTOITEM!DT_BAIXA) Then
         If IsDate(TabLANCAMENTOITEM!DT_BAIXA) Then
            If Year(TabLANCAMENTOITEM!DT_BAIXA) > 2000 Then
               item.SubItems(7) = "" & TabLANCAMENTOITEM!DT_BAIXA
               item.SubItems(9) = "Baixado"
            End If
         End If
      End If

      If IsDate(TabLANCAMENTOITEM!DT_VENCIMENTO) Then _
         item.SubItems(8) = "" & TabLANCAMENTOITEM!DT_VENCIMENTO

      If Not IsNull(TabLANCAMENTOITEM!DT_CANCELA) Then
         If IsDate(TabLANCAMENTOITEM!DT_CANCELA) Then
            item.SubItems(10) = "" & TabLANCAMENTOITEM!DT_CANCELA
            item.SubItems(9) = "Cancelado"
         End If
      End If
SQL = "" & TabLANCAMENTOITEM.Fields("CC_ID").Value
      item.SubItems(11) = "" & Trim(TabLANCAMENTOITEM.Fields("historico").Value)
      item.SubItems(12) = "" & Trim(TRAZ_DESCRITOR("O", SQL))

      If Not IsNull(TabLANCAMENTOITEM!STATUS) Then
         If TabLANCAMENTOITEM!STATUS = "A" Then
            item.SubItems(9) = "Aberto"
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
         End If
         If TabLANCAMENTOITEM!STATUS = "B" Then
            item.SubItems(9) = "Baixado"
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
         End If
         If TabLANCAMENTOITEM!STATUS = "C" Then
            item.SubItems(9) = "Cancelado"
            item.ForeColor = vbRed
            item.ListSubItems(1).ForeColor = vbRed
            item.ListSubItems(2).ForeColor = vbRed
            item.ListSubItems(3).ForeColor = vbRed
            item.ListSubItems(4).ForeColor = vbRed
            item.ListSubItems(5).ForeColor = vbRed
            item.ListSubItems(6).ForeColor = vbRed
            item.ListSubItems(7).ForeColor = vbRed
            item.ListSubItems(8).ForeColor = vbRed
         End If
      End If

      TabLANCAMENTOITEM.MoveNext
   Wend
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Public Sub MOSTRA_ITEM()
'On Error GoTo ERRO_TRATA

   If txtSeq.Text <> "" Then
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select LANCAMENTO.LANCAMENTO_ID, LANCAMENTO.PESSOA_ID, LANCAMENTO.ESTABELECIMENTO_ID, "
      SQL = SQL & " LANCAMENTO.TIPOVENDA_ID, LANCAMENTO.NUMR_DOC, LANCAMENTO.TIPO_LANCAMENTO, "
      SQL = SQL & " LANCAMENTO.DT_CAD, ITEMLANCAMENTO.SEQ, ITEMLANCAMENTO.FORMAPAGTO_ID, "
      SQL = SQL & " ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.DT_VENCIMENTO, ITEMLANCAMENTO.DT_BAIXA,"
      SQL = SQL & " ITEMLANCAMENTO.DT_CANCELA , ITEMLANCAMENTO.Valor_Desconto, ITEMLANCAMENTO.NUMR_DP, "
      SQL = SQL & " ITEMLANCAMENTO.CC_ID, ITEMLANCAMENTO.HISTORICO,PERC_JUROS,ITEMLANCAMENTO.Status"
      SQL = SQL & " from LANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

      SQL = SQL & " where lancamento.numr_doc = " & txtLanc.Text
      SQL = SQL & " and seq = " & txtSeq.Text
      SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then
         If Not IsNull(TabAUX.Fields("CC_ID").Value) Then
            cmbCCAux.Text = TabAUX.Fields("CC_ID").Value
            cmbCC.Text = "" & Trim(TRAZ_DESCRITOR("O", TabAUX.Fields("CC_ID").Value))
         End If
         VALOR_ITEM_N = 0 & TabAUX!Valor_Item
         txtJuros.Text = "" & TabAUX!PERC_JUROS
         txtNumrDoc.Text = "" & Trim(TabAUX!NUMR_DP)

         If Not IsNull(TabAUX!STATUS) Then
            If TabAUX!STATUS = "A" Then _
               cmbStatusLancItem.Text = "Aberto"
            If TabAUX!STATUS = "B" Then _
               cmbStatusLancItem.Text = "Baixado"
            If TabAUX!STATUS = "C" Then _
               cmbStatusLancItem.Text = "Cancelado"
         End If

         cmbModalidadeAux.Text = "" & TabAUX!FORMAPAGTO_ID
         cmbModalidade.Text = "" & Trim(TRAZ_DESCRICAO_FORMAPAGTO(TabAUX!FORMAPAGTO_ID))

         txtDtVenc.PromptInclude = False
            txtDtVenc.Text = "" & TabAUX!DT_VENCIMENTO
         txtDtVenc.PromptInclude = True

         txtDtBaixa.PromptInclude = False
         If IsDate(TabAUX!DT_BAIXA) Then _
            If Year(TabAUX!DT_BAIXA) > 2000 Then _
               txtDtBaixa.Text = "" & TabAUX!DT_BAIXA
         txtDtBaixa.PromptInclude = True

         txtHistorico.Text = "" & Trim(TabAUX.Fields("historico").Value)

         VALOR_DESCONTO_N = 0 & TabAUX!Valor_Desconto
         VALOR_DIFERENCA_N = 0 & TabAUX!Valor_Item
         VLR_DESCT_DIF_N = 0 & VALOR_DESCONTO_N
         txtDesconto.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
         txtTotalItem.Text = "" & Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
         txtValorItem = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      End If
   End If
   If TabAUX.State = 1 Then _
      TabAUX.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ITEM"
End Sub

Private Sub MATA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtLanc.Text) <> "" Then
      Dim TabLocal As New ADODB.Recordset

      If TabLocal.State = 1 Then _
         TabLocal.Close

      SQL = "select numr_doc,TIPO_LANCAMENTO from LANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " where numr_doc = " & txtLanc.Text
      SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabLocal.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLocal.EOF Then
         If TabLocal!TIPO_LANCAMENTO <> INDR_RECEITA Then
            MsgBox "Operação não permitida, verifique se esse título é realmente do " & frmFINGERALANC.Caption
            Exit Sub
         End If

         PEDIDO_ID_N = 0 & TabLocal.Fields("numr_doc").Value

         If TabLocal.State = 1 Then _
            TabLocal.Close

         If INDR_RECEITA = 1 Then
            SQL = "select * from NF WITH (NOLOCK) "
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            TabLocal.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLocal.EOF Then
               If TabLocal.State = 1 Then _
                  TabLocal.Close
               Msg = "Existe Nota Fiscal para este Lancamento, Deseja Excluir?"
               PERGUNTA Msg, vbYesNo + 32, "Lancamentos", "DEMO.HLP", 1000
               If RESPOSTA = vbNo Then _
                  Exit Sub
            End If
            If TabLocal.State = 1 Then _
               TabLocal.Close
         End If

         SQL = "Delete from itemLANCAMENTO "
         SQL = SQL & " Where Numr_doc = " & PEDIDO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "Delete from LANCAMENTO "
         SQL = SQL & " Where Numr_doc = " & PEDIDO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         LIMPA_TUDO
      End If
      If TabLocal.State = 1 Then _
         TabLocal.Close

      txtLanc.SetFocus
   End If      'If Trim(txtLanc.Text) <> "" Then

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_LANCAMENTO"
End Sub

Private Sub CALCULA_JUROS_DESCONTO()
'On Error GoTo ERRO_TRATA

   Dim Valor_Item As Double
   Dim Perc_Item As Double
   Dim Valor_Desconto As Double

   Valor_Item = 0
   Perc_Item = 0
   Valor_Desconto = 0
   VALOR_ITEM_N = 0
   If txtValorItem.Text <> "" Then _
      Valor_Item = txtValorItem.Text
   If txtJuros.Text <> "" Then
      Perc_Item = txtJuros.Text
      If optPercJuros.Value = True Then
         If Perc_Item > 0 Then _
            Valor_Item = (Valor_Item * Perc_Item / 100) + Valor_Item
         Else
            If Perc_Item > 0 Then _
               Valor_Item = (Valor_Item + Perc_Item) + Valor_Item
      End If
   End If
   If txtDesconto.Text <> "" Then
      Valor_Desconto = txtDesconto.Text
      VALOR_ITEM_N = txtValorItem.Text
      If optPercDesc.Value = True Then
         If Valor_Desconto > 0 Then _
            Valor_Item = Valor_Item - (VALOR_ITEM_N * Valor_Desconto / 100)
         Else
            If Valor_Desconto > 0 Then _
               Valor_Item = Valor_Item - Valor_Desconto
      End If
   End If

   txtTotalItem.Text = Format(Valor_Item, strFormatacao2Digitos)

   If Valor_Item < 0 Then
      MsgBox "Valor digitado inválido."
      txtValorItem.Text = ""
      txtJuros.Text = ""
      txtDesconto.Text = ""
      txtValorItem.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CALCULA_JUROS_DESCONTO"
End Sub

Private Sub GRAVA_CHEQUE()
'On Error GoTo ERRO_TRATA

   If TabCHEQUE.State = 1 Then _
      TabCHEQUE.Close

   CRITERIO_A = Trim(Replace(txtCNPJCPF.Text, ",", ""))
   CRITERIO_A = Trim(Replace(CRITERIO_A, "/", ""))
   CRITERIO_A = Trim(Replace(CRITERIO_A, "-", ""))
   CRITERIO_A = Trim(Replace(CRITERIO_A, ",", ""))
   CRITERIO_A = Trim(Replace(CRITERIO_A, ".", ""))

   SQL = "select * from CHEQUE WITH (NOLOCK) "
   'Verificar por onde pegar o cheque
   SQL = SQL & " where CGCCPFRESP = '" & CRITERIO_A & "'"
   SQL = SQL & " and NUMR_CHEQUE = '" & txtCheque.Text & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCHEQUE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCHEQUE.EOF Then
      CRITERIO_A = Trim(Replace(txtCGCREP.Text, ",", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, "/", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, "-", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, ",", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, ".", ""))
      SQL = ""
      SQL = "update CHEQUE set "
      SQL = SQL & "CGCCPFREPASSE = '" & CRITERIO_A & "'"
      If Not IsDate(txtdtdesc.Text) Then
         SQL = SQL & ",DT_DESCONTO = " & "'" & Format(0, "dd/mm/yyyy") & "'"
         Else: SQL = SQL & ",DT_DESCONTO = " & "'" & Format(txtdtdesc.Text, "dd/mm/yyyy") & "'"
      End If
      CRITERIO_A = Trim(Replace(txtCNPJCPF.Text, ",", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, "/", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, "-", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, ",", ""))
      CRITERIO_A = Trim(Replace(CRITERIO_A, ".", ""))
      SQL = SQL & " where CGCCPFRESP = '" & CRITERIO_A & "'"
      SQL = SQL & " and NUMR_CHEQUE = '" & txtCheque.Text & "'"

      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabCHEQUE.State = 1 Then _
      TabCHEQUE.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CHEQUE"
End Sub

Sub IMPRIME_CARNE()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = ""
   FORMULA_REL = "{vw_BOLETO.estabelecimento_id} = " & ESTABELECIMENTO_ID_N

   If Trim(txtLanc.Text) <> "" Then _
      If IsNumeric(txtLanc.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {vw_BOLETO.numr_doc} = " & txtLanc.Text

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_carne.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIME_CARNE"
End Sub

Sub CONSULTA_PESSOA()
'On Error GoTo ERRO_TRATA

   If INDR_RECEITA = 1 Then _
      TIPO_PESSOA_CADASTRO = "CLIENTE"

   If INDR_RECEITA = 2 Then _
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"

   frmPessoaConsulta.Show 1

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Mask = "##############"
      txtCNPJCPF.Text = ""
         txtCNPJCPF.Text = CNPJCPF_A
      TXTCNPJCPF_KeyPress 13
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = ""
   End If

   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PESSOA"
End Sub

Sub MATA_ITEM()
'On Error GoTo ERRO_TRATA

   If Trim(txtSeq.Text) <> "" And Trim(txtLanc.Text) <> "" Then
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select LANCAMENTO.LANCAMENTO_ID,seq from LANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

      SQL = SQL & " where LANCAMENTO.numr_doc = " & Trim(txtLanc.Text)
      SQL = SQL & " and seq = " & Trim(txtSeq.Text)
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then
         Msg = "Confirma exclusão desse item ?"
         PERGUNTA Msg, vbYesNo + 32, "Exclui Lancamentos", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            SQL = "delete from ITEMLANCAMENTO "
              SQL = SQL & " where lancamento_id = " & TabAUX.Fields("lancamento_id").Value
              SQL = SQL & " and seq = " & TabAUX.Fields("seq").Value
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_BODY
            SETA_GRID
         End If
      End If
      If TabAUX.State = 1 Then _
         TabAUX.Close
   End If
   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_ITEM"
End Sub
