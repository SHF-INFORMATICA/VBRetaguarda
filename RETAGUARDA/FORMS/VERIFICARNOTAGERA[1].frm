VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNOTAGERA 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissor de Nota Fiscal"
   ClientHeight    =   8820
   ClientLeft      =   1725
   ClientTop       =   2340
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VERIFICARNOTAGERA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   2840
      Width           =   1333
   End
   Begin VB.CommandButton cmdNFE 
      Caption         =   "Gerar Nota Eletrônica"
      Height          =   855
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   2840
      Width           =   1500
   End
   Begin VB.Frame Frame9 
      Caption         =   "Quem pagará o frete?"
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   5880
      TabIndex        =   86
      Top             =   2820
      Width           =   3135
      Begin VB.OptionButton OptFreteEmitente 
         Caption         =   "Emitente"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   405
         Width           =   1215
      End
      Begin VB.OptionButton OptFreteDestinatario 
         Caption         =   "Destinatário"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   405
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Outras Informações"
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   0
      TabIndex        =   73
      Top             =   2820
      Width           =   5895
      Begin VB.TextBox TxtEspecie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "UN"
         ToolTipText     =   "Informe a especie"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtPesoLiquido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   21
         Text            =   "1"
         ToolTipText     =   "Peso liquido da nota"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtPesoBruto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "1"
         ToolTipText     =   "Peso bruto da nota"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtQuantidadeRodapeNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   18
         Text            =   "1"
         ToolTipText     =   "Informe a quantidade"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso líquido:"
         Height          =   240
         Left            =   4440
         TabIndex        =   83
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso bruto:"
         Height          =   240
         Left            =   3000
         TabIndex        =   82
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie:"
         Height          =   240
         Left            =   1560
         TabIndex        =   81
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   240
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   67
      Top             =   6360
      Width           =   11925
      Begin VB.TextBox txtDadosAdicionais 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2310
         MultiLine       =   -1  'True
         TabIndex        =   26
         Tag             =   " "
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   930
         Width           =   9525
      End
      Begin VB.TextBox txtDescDesconto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         TabIndex        =   24
         Top             =   210
         Width           =   9525
      End
      Begin VB.TextBox txtMSG 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         TabIndex        =   25
         Text            =   "DOCUMENTO EMITIDO POR EPP OPTANTE PELO SIMPLES NACIONAL E NAO GERA DIREITO A CREDITO FISCAL DE ICMS"
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   570
         Width           =   9525
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dados Adicionais:"
         Height          =   240
         Left            =   240
         TabIndex        =   70
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem Desconto:"
         Height          =   240
         Left            =   30
         TabIndex        =   69
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem Rodapé:"
         Height          =   240
         Left            =   165
         TabIndex        =   68
         Top             =   600
         Width           =   2130
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transportadora"
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   0
      TabIndex        =   66
      Top             =   2160
      Width           =   11895
      Begin VB.ComboBox cmbAuxCNPJCPF_TRANSP 
         BackColor       =   &H80000000&
         Height          =   345
         ItemData        =   "VERIFICARNOTAGERA.frx":47C4A
         Left            =   10230
         List            =   "VERIFICARNOTAGERA.frx":47C4C
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbCNPJCPF_TRANSP 
         BackColor       =   &H8000000E&
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "VERIFICARNOTAGERA.frx":47C4E
         Left            =   3960
         List            =   "VERIFICARNOTAGERA.frx":47C50
         TabIndex        =   37
         Top             =   240
         Width           =   7845
      End
      Begin MSMask.MaskEdBox txtCNPJCPF_TRANSP 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         ToolTipText     =   "Se houver uma transportadora informe aqui. F7 Consulta transportadora"
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         PromptInclude   =   0   'False
         MaxLength       =   18
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
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF:"
         Height          =   240
         Left            =   120
         TabIndex        =   71
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Frame fraEmitente 
      ForeColor       =   &H00400000&
      Height          =   1335
      Left            =   0
      TabIndex        =   57
      Top             =   840
      Width           =   11895
      Begin VB.TextBox txtIBGE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4665
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox txtEmitente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtCep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5340
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   600
         Width           =   2715
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtUF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtFone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6660
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         Top             =   960
         Width           =   1395
      End
      Begin VB.ComboBox cmbIE 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         ItemData        =   "VERIFICARNOTAGERA.frx":47C52
         Left            =   9720
         List            =   "VERIFICARNOTAGERA.frx":47C54
         TabIndex        =   16
         Top             =   960
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   330
         Left            =   9360
         TabIndex        =   8
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483634
         ForeColor       =   8388608
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IBGE:"
         Height          =   240
         Left            =   4095
         TabIndex        =   89
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emitente:"
         Height          =   240
         Left            =   195
         TabIndex        =   85
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF:"
         Height          =   240
         Left            =   8145
         TabIndex        =   65
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   240
         Left            =   15
         TabIndex        =   64
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         Height          =   240
         Left            =   9210
         TabIndex        =   63
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Height          =   240
         Left            =   4650
         TabIndex        =   62
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Município:"
         Height          =   240
         Left            =   -15
         TabIndex        =   61
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         Height          =   240
         Left            =   6075
         TabIndex        =   60
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   240
         Left            =   8160
         TabIndex        =   59
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblInsc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc.Estadual:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8070
         TabIndex        =   58
         Top             =   960
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fatura"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      TabIndex        =   55
      Top             =   3720
      Width           =   11895
      Begin MSComctlLib.ListView ListaDP 
         Height          =   1035
         Left            =   90
         TabIndex        =   56
         Top             =   210
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   1826
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Seq."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor da Duplicata"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "NºDocumento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Venc.DP"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Histórico"
            Object.Width           =   10583
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Produtos"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      TabIndex        =   53
      Top             =   5040
      Width           =   11895
      Begin MSComctlLib.ListView ListaProdutos 
         Height          =   1035
         Left            =   90
         TabIndex        =   54
         Top             =   210
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   1826
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codg."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Unidade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Aliquota"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qtd."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Valor Unitário"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Valor Total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "ST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "CFOP"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Cálculo Imposto"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      TabIndex        =   46
      Top             =   7680
      Width           =   11925
      Begin VB.TextBox txtValorOutros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8700
         MaxLength       =   12
         TabIndex        =   35
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txtVlrIcmsSub 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6780
         MaxLength       =   10
         TabIndex        =   34
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txtBaseIcmsSub 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6780
         MaxLength       =   10
         TabIndex        =   29
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txtvaloripi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   10740
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   31
         Top             =   210
         Width           =   1080
      End
      Begin VB.TextBox txtValorTotalNota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   10740
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   36
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txtValorProdutos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   32
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   28
         Top             =   210
         Width           =   1080
      End
      Begin VB.TextBox txtBaseCalculo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   27
         Top             =   210
         Width           =   1080
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   33
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txtFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8700
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   30
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr ICMS Sub."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5640
         TabIndex        =   93
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outros:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8040
         TabIndex        =   92
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. Icms Sub."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5625
         TabIndex        =   91
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VLR IPI.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10005
         TabIndex        =   84
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Nota:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9945
         TabIndex        =   52
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total dos Produtos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   51
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do ICMS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3105
         TabIndex        =   50
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base cálculo ICMS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   49
         Top             =   270
         Width           =   1620
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3480
         TabIndex        =   48
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8115
         TabIndex        =   47
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame fraNota 
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
      TabIndex        =   38
      Top             =   -120
      Width           =   11895
      Begin VB.ComboBox cmbCFOPAux 
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
         Left            =   6960
         TabIndex        =   90
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNaturezaOperacao 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Descrição Natureza Operação Fiscal"
         Top             =   550
         Width           =   4215
      End
      Begin VB.TextBox txtDtEmis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Data Emissão Nota Fiscal"
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtDtSaida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4410
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Data Saida Nota Fiscal"
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtHoraSaida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Hora de Saída Nota Fiscal"
         Top             =   200
         Width           =   1095
      End
      Begin VB.ComboBox cmbCFOP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6990
         TabIndex        =   6
         Text            =   "-- Selecione --"
         ToolTipText     =   "CFOP"
         Top             =   550
         Width           =   4815
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         ToolTipText     =   "Número Nota Fiscal"
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   10920
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "Série Nota Fiscal"
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   240
         Left            =   6330
         TabIndex        =   45
         Top             =   615
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Natureza Operação:"
         Height          =   240
         Left            =   30
         TabIndex        =   44
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data da Emissão:"
         Height          =   240
         Left            =   60
         TabIndex        =   43
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Saída:"
         Height          =   240
         Left            =   3525
         TabIndex        =   42
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Saída:"
         Height          =   240
         Left            =   5685
         TabIndex        =   41
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº  Nota:"
         Height          =   240
         Left            =   8040
         TabIndex        =   40
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série NF:"
         Height          =   240
         Left            =   9930
         TabIndex        =   39
         Top             =   255
         Width           =   885
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   960
      TabIndex        =   74
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   18
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   315
      Left            =   0
      TabIndex        =   77
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   18
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
      DesignWidth     =   11910
      DesignHeight    =   8820
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   80
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   79
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   78
      Top             =   0
      Width           =   1020
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ/CPF:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   75
      Top             =   360
      Width           =   810
   End
End
Attribute VB_Name = "frmNOTAGERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Tipo_Endereço             As String * 1
   Dim rstEmpresa                As New ADODB.Recordset
   Dim UNIDADE_FEDERAÇÃO_EMPRESA As String * 2
   Dim NOME_VEND_A               As String
   Dim TIPO_DOC                  As String
   Dim NF_DEV_ENTRADA            As Long
   Dim intTributacao             As Integer
   Dim strCNPJEMPRESA            As String
   Dim NossoNumero               As String
   Dim CODG_SUFRAMA_A            As String
   Dim booOptanteSimples         As Boolean
   'Valores com devolucao de entrada
   Dim VLR_ICMS_SUB_DEV As Double, VLR_FRETE_DEV As Double, VLR_OUTROS_DEV As Double, VLR_IPI_DEV As Double

Private Sub Form_Load()
   INICIALIZA_NF
   Frame7.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo erro_trata

   Select Case KeyCode
      Case vbKeyF10
         GERAR_NFE
   End Select

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub cmbCFOP_Click()
   txtNaturezaOperacao.Text = "" & Trim(cmbCFOP.Text)
   txtNaturezaOperacao.Refresh

On Error Resume Next
   cmbCFOPAux.ListIndex = cmbCFOP.ListIndex

   If Trim(cmbCFOPAux.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from CFOP "
      SQL = SQL & " where codigo = '" & Trim(cmbCFOPAux.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            txtNaturezaOperacao.Text = "" & TabTemp.Fields(0).Value

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
End Sub

Private Sub TxtEspecie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
       KeyAscii = 0
       TxtPesoBruto.SetFocus
   End If
End Sub

Private Sub TxtPesoBruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        TxtPesoLiquido.SetFocus
    End If
End Sub

Private Sub TxtPesoLiquido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDadosAdicionais.SetFocus
    End If
End Sub

Private Sub TxtQuantidadeRodapeNota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         KeyAscii = 0
         TxtEspecie.SetFocus
     End If
End Sub

Private Sub txtDadosAdicionais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
   End If
End Sub

Private Sub cmdNFE_Click()
   AtualizaTotalNota
   GERAR_NFE
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub OptFreteDestinatario_Click()
On Error GoTo erro_trata

   OptFreteDestinatario.Value = True
   OptFreteEmitente = False

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "OptFreteDestinatario_Click"
End Sub

Private Sub OptFreteEmitente_Click()
On Error GoTo erro_trata

   OptFreteDestinatario.Value = False
   OptFreteEmitente = True

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "OptFreteEmitente_Click"
End Sub

Private Sub txtCNPJCPF_TRANSP_GotFocus()
On Error GoTo erro_trata

   txtCNPJCPF_TRANSP.PromptInclude = False
   If txtCNPJCPF_TRANSP.Text = "" Then
      txtCNPJCPF_TRANSP.Mask = "##############"
      Else
         If Len(txtCNPJCPF_TRANSP.Text) <= 11 Then
            txtCNPJCPF_TRANSP.Mask = "###.###.###-##"
            Else
               If Len(txtCNPJCPF_TRANSP.Text) > 11 Then
                  txtCNPJCPF_TRANSP.Mask = "##.###.###/####-##"
               End If
         End If
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_GotFocus"
End Sub

Private Sub txtCNPJCPF_TRANSP_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo erro_trata

    Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYTRANSPORTADORA.Show 1
         If CPF_N <> "" Then
            txtCNPJCPF_TRANSP.PromptInclude = False
            txtCNPJCPF_TRANSP.Text = CPF_N
         End If

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select cgccpf,nome from TRANSPORTADORA "
         SQL = SQL & " where cgccpf = '" & txtCNPJCPF_TRANSP.Text & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               cmbCNPJCPF_TRANSP.Text = TabTemp!NOME

         If TabTemp.State = 1 Then _
            TabTemp.Close

         txtCNPJCPF_TRANSP.PromptInclude = True
         CPF_N = ""
   End Select

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_KeyDown"
End Sub

Private Sub txtCNPJCPF_TRANSP_KeyPress(KeyAscii As Integer)
On Error GoTo erro_trata

   If KeyAscii = 13 Then
      KeyAscii = 0
      TxtQuantidadeRodapeNota.SetFocus
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_KeyPress"
End Sub

Private Sub txtCNPJCPF_TRANSP_LostFocus()
On Error GoTo erro_trata

   txtCNPJCPF_TRANSP.PromptInclude = False

   If txtCNPJCPF_TRANSP.Text <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select nome from TRANSPORTADORA "
      SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         MsgBox "Informar trasportadora."
         txtCNPJCPF_TRANSP.Text = ""
         cmbCNPJCPF_TRANSP.Text = ""
         Exit Sub
         Else: If Not IsNull(TabTemp.Fields(0).Value) _
               Then cmbCNPJCPF_TRANSP.Text = Trim(TabTemp.Fields(0).Value)
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   txtCNPJCPF_TRANSP.PromptInclude = True

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_LostFocus"
End Sub

Private Sub cmbCNPJCPF_TRANSP_GotFocus()
On Error GoTo erro_trata

   cmbCNPJCPF_TRANSP.Clear
   cmbAuxCNPJCPF_TRANSP.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TRANSPORTADORA "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by nome"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCNPJCPF_TRANSP.AddItem Trim(TabTemp!CGCCPF) & " - " & Trim(TabTemp!NOME)
      cmbAuxCNPJCPF_TRANSP.AddItem Trim(TabTemp!CGCCPF)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "cmbCNPJCPF_TRANSP_GotFocus"
End Sub

Private Sub cmbCNPJCPF_TRANSP_Click()
On Error GoTo erro_trata

   cmbAuxCNPJCPF_TRANSP.ListIndex = cmbCNPJCPF_TRANSP.ListIndex
   txtCNPJCPF_TRANSP.PromptInclude = False
      Select Case Len(cmbAuxCNPJCPF_TRANSP.Text)
         Case Is <= 11
            txtCNPJCPF_TRANSP.Mask = "###.###.###-##"
         Case Is = 14
            txtCNPJCPF_TRANSP.Mask = "##.###.###/####-##"
      End Select
      txtCNPJCPF_TRANSP.Text = "" & cmbAuxCNPJCPF_TRANSP.Text
   txtCNPJCPF_TRANSP.PromptInclude = True

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "cmbCNPJCPF_TRANSP_Click"
End Sub

'============= S U B R O T I N A S

Private Sub MONTA_NOTA_SAIDA()
On Error GoTo erro_trata

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from CABECAREQ "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then  'CHECA TIPO VENDA
      If ((TabCABECA!TIPO_REGISTRO <> "R") And (TabCABECA!TIPO_REGISTRO <> "S")) Then
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         MsgBox "Número de Pedido inválido !!!"
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         ''Frame7.Enabled = False
         Exit Sub
      End If
      If TabCABECA!Status = 1 Then
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         MsgBox "Não é permitido emitir nota para Orçamento."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         ''Frame7.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If
      If TabCABECA!Status = 2 Then
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         MsgBox "É necessário fazer faturamento antes de emitir nota fiscal."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         ''Frame7.Enabled = False
         Frame8.Enabled = False
         Unload Me
         Exit Sub
      End If
      If TabCABECA!Status = 4 Then
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         MsgBox "Cupom fiscal já emitido para essa Pedido."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         ''Frame7.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      'passou do cabecareq, checar na tabela nf agora
      SQL = "select * from NF "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         txtNota.Text = TabNOTA!numr_Nota
         txtSerie.Text = TabNOTA!serie_nota
         txtDtEmis.Text = Format(TabNOTA!dt_emissao, "dd/mm/yyyy")
         txtDtSaida.Text = Format(TabNOTA!DT_ENTRASAI, "dd/mm/yyyy")

         If Not IsNull(TabNOTA!TRANSP_ID) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cgccpf,nome from TRANSPORTADORA "
            SQL = SQL & " where cgccpf = '" & TabNOTA!TRANSP_ID & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp.Fields(0).Value) Then
                  cmbCNPJCPF_TRANSP.Text = Trim(TabTemp!CGCCPF) & " - " & Trim(TabTemp!NOME)
                  txtCNPJCPF_TRANSP.Text = TabTemp.Fields(0).Value
                  'Volumes
                  If TabNOTA!qtd_volume <> "" Then _
                     TxtQuantidadeRodapeNota.Text = TabNOTA!qtd_volume

                  If TabNOTA!TIPO_ESPECIE <> "" Then _
                     TxtEspecie.Text = TabNOTA!TIPO_ESPECIE

                  If TabNOTA!PESO_BRUTO <> "" Then _
                     TxtPesoBruto.Text = TabNOTA!PESO_BRUTO

                  If TabNOTA!PESO_LIQUIDO <> "" Then _
                     TxtPesoLiquido.Text = TabNOTA!PESO_LIQUIDO
               End If
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         If Not IsNull(TabNOTA!cfop) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from CFOP "
            SQL = SQL & " where CODIGO = '" & TabNOTA!cfop & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbCFOP.Text = TabTemp!Codigo & "-" & Trim(TabTemp!Descricao)
               txtNaturezaOperacao.Text = Trim(TabTemp!Descricao)
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         MOSTRA_NOTA_TELA

         If TabCABECA.State = 1 Then _
            TabCABECA.Close
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         'aqui
         MsgBox "Já existe nota fiscal emitida para Pedido = " & NUMR_REQ_N & " ; Nota Fiscal = " & txtNota & " ; Empresa = " & EMPRESA_ID_N

         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         ''Frame7.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      MOSTRA_NOTA_TELA
      Else
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         MsgBox "Registro de venda não encontrado."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         ''Frame7.Enabled = False
         Exit Sub
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_NOTA_SAIDA"
End Sub

Private Sub MONTA_NOTA_Devolução() 'Devolução de Entrada de mecadorias
On Error GoTo erro_trata

   If TabCABENTRA.State = 1 Then _
      TabCABENTRA.Close

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABENTRA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABENTRA.EOF Then
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "select * from NF "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         txtNota.Text = TabNOTA!numr_Nota
         txtSerie.Text = TabNOTA!serie_nota
         txtDtEmis.Text = Format(TabNOTA!dt_emissao, "dd/mm/yyyy")
         txtDtSaida.Text = Format(TabNOTA!DT_ENTRASAI, "dd/mm/yyyy")
         If Not IsNull(TabNOTA!TRANSP_ID) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cgccpf,nome from TRANSPORTADORA "
            SQL = SQL & " where cgccpf = '" & TabNOTA!TRANSP_ID & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp.Fields(0).Value) Then
                  cmbCNPJCPF_TRANSP.Text = Trim(TabTemp!CGCCPF) & " - " & Trim(TabTemp!NOME)
                  txtCNPJCPF_TRANSP.Text = TabTemp.Fields(0).Value
               End If
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
         If Not IsNull(TabNOTA!cfop) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from CFOP "
            SQL = SQL & " where codigo = '" & TabNOTA!cfop & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbCFOP.Text = TabTemp!Codigo & "-" & Trim(TabTemp!Descricao)
               txtNaturezaOperacao.Text = Trim(TabTemp!Descricao)
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         MOSTRA_Devolução_TELA

         If TabNOTA.State = 1 Then _
            TabNOTA.Close
         If TabCABENTRA.State = 1 Then _
            TabCABENTRA.Close

         MsgBox "Já existe nota fiscal emitida para esta Devolução = " & NUMR_REQ_N & " ; Nota Fiscal = " & txtNota.Text & " ; Empresa = " & EMPRESA_ID_N
    
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'Frame6.Enabled = False
         'Frame7.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      MOSTRA_Devolução_TELA
      Else
         If TabCABENTRA.State = 1 Then _
            TabCABENTRA.Close

         MsgBox "Registro de Compra não encontrado."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'Frame6.Enabled = False
         'Frame7.Enabled = False
         Exit Sub
   End If
   If TabCABENTRA.State = 1 Then _
      TabCABENTRA.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_NOTA_Devolução"
End Sub

Private Sub MOSTRA_Devolução_TELA() 'Devolução de Entrada
On Error GoTo erro_trata

   If txtDtEmis.Text = "" Then _
      txtDtEmis.Text = Format(Date, "dd/mm/yyyy")
   If txtDtSaida.Text = "" Then _
      txtDtSaida.Text = Format(Date, "dd/mm/yyyy")

'==========================================
   If Trim(txtNota.Text) = "" Then
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "select seq_nota_saida, serie_nota_saida from EMPRESA"
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         txtNota.Text = TabNOTA.Fields(0).Value + 1
         txtNota.Refresh
         txtSerie.Text = TabNOTA.Fields(1).Value
      End If
      txtSerie.Refresh

      If TabNOTA.State = 1 Then _
         TabNOTA.Close
   End If
'==========================================

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      NF_DEV_ENTRADA = TabTemp!numr_Nota

      If Not IsNull(TabTemp!cfop) Then
         If TabCEP.State = 1 Then _
            TabCEP.Close

         SQL = "select * from CFOP "
         SQL = SQL & " where codigo = '" & TabTemp!cfop & "'"
         TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCEP.EOF Then
            cmbCFOP.Text = Trim(TabCEP!Codigo) & "-" & Trim(TabCEP!Descricao)
            txtNaturezaOperacao.Text = Trim(TabCEP!Descricao)

            cmbCFOPAux.Text = Trim(TabCEP!Codigo)
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & " , Nota fiscal de devolução referente a NFe : " & NF_DEV_ENTRADA
'===========================================

   MOSTRA_FORNECEDOR

   TOTAIS_NOTA_Devolução

   GRID_DP_DEV

   GRID_PRODUTOS_DEV

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_Devolução_TELA"
End Sub

Private Sub MOSTRA_NOTA_TELA()
On Error GoTo erro_trata

   Dim NOME_A        As String
   Dim VENDEDOR_ID_N   As Long
   Dim TIPO_NOTA_A   As String

   NOME_A = ""

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from VENDEDOR "
   SQL = SQL & " where vendedor_id = " & TabCABECA!VENDEDOR_ID
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      NOME_A = Trim(TabUSU!NOME_VEND)
      VENDEDOR_ID_N = TabUSU!VENDEDOR_ID
      If Trim(TIPO_DOC) = "S" Then _
         Me.Caption = Me.Caption & "Emissão Nota Fiscal de Saída" & " ; Vendedor = " & Trim(TabUSU!NOME_VEND)
         
      If Trim(TIPO_DOC) = "DV" Then _
         Me.Caption = Me.Caption & "Emissão Nota Fiscal de Devolução" & " ; Vendedor = " & Trim(TabUSU!NOME_VEND)
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If txtDtEmis.Text = "" Then _
      txtDtEmis.Text = Format(Date, "dd/mm/yyyy")
   If txtDtSaida.Text = "" Then _
      txtDtSaida.Text = Format(Date, "dd/mm/yyyy")

   If txtSerie.Text = "" Then
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "select seq_nota_saida, serie_nota_saida from EMPRESA"
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         txtNota.Text = TabNOTA.Fields(0).Value + 1
         txtNota.Refresh
         txtSerie.Text = TabNOTA.Fields(1).Value
         txtSerie.Refresh
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close
   End If

   MOSTRA_CLIENTE

   TOTAIS_NOTA

   GRID_DP

   GRID_PRODUTOS

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_NOTA_TELA"
End Sub

Private Sub MOSTRA_CLIENTE()
On Error GoTo erro_trata

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from CLIENTE "
   SQL = SQL & " where cgccpf='" & TabCABECA!CGCCPF & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      cmbIE.Clear
      If Not IsNull(TabCliente!IE) Then
         If Trim(TabCliente!IE) <> "" Then
            cmbIE.Text = TabCliente!IE
            Else: MsgBox "Inscrição Estadual inválida !!!"
         End If
         Else: MsgBox "Inscrição Estadual inválida !!!"
      End If

      txtEmitente.Text = TabCliente!NOME
      Select Case Len(TabCliente!CGCCPF)
         Case Is <= 11
            txtCNPJCPF.Mask = "###.###.###-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = TabCliente!CGCCPF
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "R"
         Case Is = 14
            txtCNPJCPF.Mask = "##.###.###/####-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = TabCliente!CGCCPF
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "C"
      End Select

      If TabEND.State = 1 Then _
         TabEND.Close

      SQL = "select * from FONE "
      SQL = SQL & " where prop = '" & Trim(TabCliente.Fields("cgccpf").Value) & "'"
      SQL = SQL & " and numero <> ''"
      TabEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEND.EOF Then
         txtFone.Text = "" & Trim(TabEND.Fields("numero").Value)
         Else: MsgBox "Fone de cliente não encontrado !!!"
      End If

      If TabEND.State = 1 Then _
         TabEND.Close

      'endereço
      SQL = "select * from ENDERECO "
      SQL = SQL & " where prop='" & TabCliente!CGCCPF & "'"
      SQL = SQL & " and tipo = '" & Tipo_Endereço & "'"
      TabEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEND.EOF Then
         If Not IsNull(TabEND!Rua) Then _
            txtEnd.Text = TabEND!Rua

         If Not IsNull(TabEND!Complemento) Then
            If txtEnd.Text = "" Then
               txtEnd.Text = TabEND!Complemento
               Else: txtEnd.Text = txtEnd.Text & " , " & TabEND!Complemento
            End If
         End If

         If Not IsNull(TabEND!Bairro) Then _
            txtBairro.Text = TabEND!Bairro

         If Not IsNull(TabEND!CEP) Then  'CEP
            txtCep.Text = "" & TabEND!CEP

            If TabCEP.State = 1 Then _
               TabCEP.Close

            SQL = "select * from CEP "
            SQL = SQL & " where cep = '" & Trim(txtCep.Text) & "'"
            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCEP.EOF Then
               If Not IsNull(TabCEP!Cidade) Then _
                  txtCidade.Text = TabCEP!Cidade
               If Not IsNull(TabCEP!UF) Then _
                  txtUF.Text = TabCEP!UF
               If IsNull(TabCEP!codigo_ibge) Then
                  MsgBox "Código IBGE inválido !!!"
                  Else
                     If Trim(TabCEP!codigo_ibge) = "" Then
                        MsgBox "Código IBGE inválido !!!"
                        Else: txtIBGE.Text = TabCEP!codigo_ibge
                     End If
               End If
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
         End If
      End If
      If TabEND.State = 1 Then _
         TabEND.Close

      If TabEMP.State = 1 Then _
         TabEMP.Close

      SQL = "select * from EMPRESA "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabEMP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEMP.EOF Then
         'SERIE NOTA SAIDA
         If Not IsNull(TabEMP!serie_NOTA_SAIDA) Then _
            txtSerie.Text = TabEMP!serie_NOTA_SAIDA

         If Not IsNull(TabEMP!Instrucao_Fisco) Then _
            txtMSG.Text = (TabEMP!Instrucao_Fisco)

         ' = Left(TABEMP!Instrucao_Fisco, 50)
         ' = Mid(TABEMP!Instrucao_Fisco, 51, 50)
            
         txtNota.Refresh
         Else
            MsgBox "Erro no arquivo de empresa."
            Exit Sub
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      PEGA_DADOS_EMPRESA 'Atualiza o buffer com informacoes da empresa

      intTributacao = rstEmpresa!Tipo_Enquadramento_Simples
      If UNIDADE_FEDERAÇÃO_EMPRESA = "" Then
         MsgBox "Impossivel continuar, inconsitencia na UNIDADE FEDERAÇÃO cadastrada para empresa."
         Exit Sub
         Else
            If Trim(TIPO_DOC) = "S" Then
               If Trim(txtUF.Text) = Trim(UNIDADE_FEDERAÇÃO_EMPRESA) Then     'dentro do estado
                  cmbCFOPAux.Text = rstEmpresa!CFOP_SAIDA_DE
'===================================== 'para cupom fiscal vinculado a nota fiscal
                  If Not IsNull(TabCABECA.Fields("status").Value) Then
                     If TabCABECA.Fields("status").Value = 7 Then
                        cmbCFOPAux.Text = 5929
                        txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & "Nota fiscal referente ao cupom fiscal nº: " & TabCABECA.Fields("numr_cupom").Value

                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close

                        SQL = "select nf_id from NF "
                        SQL = SQL & " where numr_req = " & NUMR_REQ_N
                        TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If Not TabConsulta.EOF Then
                           SQL = "update NFITEM set cfop = 5929"
                           SQL = SQL & " where nf_id = " & TabConsulta.Fields("nf_id").Value
                           CONECTA_RETAGUARDA.Execute SQL
                        End If

                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close
                     End If
                  End If
'=====================================
                  Else                                                        'fora do estado
                     cmbCFOPAux.Text = rstEmpresa!CFOP_SAIDA_FE
'===================================== 'para cupom fiscal vinculado a nota fiscal
                     If Not IsNull(TabCABECA.Fields("status").Value) Then
                        If TabCABECA.Fields("status").Value = 7 Then
                           cmbCFOPAux.Text = 6929
                           txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & "Nota fiscal referente ao cupom fiscal nº: " & TabCABECA.Fields("numr_cupom").Value
                        
                           If TabConsulta.State = 1 Then _
                              TabConsulta.Close

                           SQL = "select nf_id from NF "
                           SQL = SQL & " where numr_req = " & NUMR_REQ_N
                           TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabConsulta.EOF Then
                              SQL = "update NFITEM set cfop = 6929"
                              SQL = SQL & " where nf_id = " & TabConsulta.Fields("nf_id").Value
                              CONECTA_RETAGUARDA.Execute SQL
                           End If

                           If TabConsulta.State = 1 Then _
                              TabConsulta.Close
                        End If
                     End If
'=====================================
               End If
               Else  'não é nota de saida
                  If Trim(txtUF.Text) = Trim(UNIDADE_FEDERAÇÃO_EMPRESA) Then
                     cmbCFOPAux.Text = rstEmpresa!CFOP_DV_SAI_DE
                     Else: cmbCFOPAux.Text = rstEmpresa!CFOP_DV_SAI_FE
                  End If
            End If
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

'===========suframa
      If Not IsNull(TabCliente.Fields("codg_suframa").Value) Then
         If Trim(TabCliente.Fields("codg_suframa").Value) <> "" Then
            If IsNumeric(TabCliente.Fields("codg_suframa").Value) Then
               If Len(TabCliente.Fields("codg_suframa").Value) > 4 Then
                  CODG_SUFRAMA_A = Trim(TabCliente.Fields("codg_suframa").Value)
                  cmbCFOPAux.Text = 6110
                  cmbCFOP.Text = "VENDA MERCAD ADQ. OU RECEBIDA TERCEIROS" & "-" & 6110
                  txtNaturezaOperacao.Text = "VENDA MERCAD ADQ. OU RECEBIDA TERCEIROS" & "-" & 6110
                  txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & " " & "Codigo Suframa: " & Trim(CODG_SUFRAMA_A)
               End If
            End If
         End If
      End If

      SQL = "select * from CFOP "
      SQL = SQL & " where CODIGO = '" & Trim(cmbCFOPAux.Text) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         cmbCFOP.Text = TabConsulta!Codigo & "-" & TabConsulta!Descricao
         txtNaturezaOperacao.Text = Trim(TabConsulta!Descricao)
         Else
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            MsgBox "CFOP não cadastrado."
            Exit Sub
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      Else
         If TabCliente.State = 1 Then _
            TabCliente.Close

         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         MsgBox "Cliente não cadastrado !!!"
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'Frame6.Enabled = False
         'Frame7.Enabled = False
         Exit Sub
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CLIENTE"
End Sub

Private Sub MOSTRA_FORNECEDOR()
On Error GoTo erro_trata

   If TabFOR.State = 1 Then _
      TabFOR.Close

   SQL = "select * from FORNECEDOR "
   SQL = SQL & " where fornecedor_id = " & TabCABENTRA!FORNECEDOR_ID
   TabFOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFOR.EOF Then
      cmbIE.Clear
      cmbIE.Text = TabFOR!IE
      cmbIE.AddItem TabFOR!IE
      txtEmitente.Text = TabFOR!NOME
      Select Case Len(TabFOR!CGCCPF)
         Case Is <= 11
            txtCNPJCPF.Mask = "###.###.###-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = TabFOR!CGCCPF
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "R"
         Case Is = 14
            txtCNPJCPF.Mask = "##.###.###/####-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = TabFOR!CGCCPF
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "C"
      End Select

      If TabEND.State = 1 Then _
         TabEND.Close

      'endereço
      SQL = "select * from ENDERECO "
      SQL = SQL & " where prop='" & TabFOR!CGCCPF & "'"
      SQL = SQL & " and tipo = '" & Tipo_Endereço & "'"
      TabEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEND.EOF Then
         If Not IsNull(TabEND!Rua) Then _
            txtEnd.Text = TabEND!Rua
         If Not IsNull(TabEND!Complemento) Then
            If txtEnd.Text = "" Then
               txtEnd.Text = TabEND!Complemento
               Else: txtEnd.Text = txtEnd.Text & " , " & TabEND!Complemento
            End If
         End If

         If Not IsNull(TabEND!Bairro) Then _
            txtBairro.Text = TabEND!Bairro

         If Not IsNull(TabEND!CEP) Then  'CEP
            txtCep.Text = "" & TabEND!CEP

            If TabCEP.State = 1 Then _
               TabCEP.Close

            SQL = "select * from CEP "
            SQL = SQL & " where cep = '" & Trim(txtCep.Text) & "'"
            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCEP.EOF Then
               If Not IsNull(TabCEP!Cidade) Then _
                  txtCidade.Text = TabCEP!Cidade
               If Not IsNull(TabCEP!UF) Then _
                  txtUF.Text = TabCEP!UF
               If IsNull(TabCEP!codigo_ibge) Then
                  MsgBox "Código IBGE inválido !!!"
                  Else
                     If Trim(TabCEP!codigo_ibge) = "" Then
                        MsgBox "Código IBGE inválido !!!"
                        Else: txtIBGE.Text = TabCEP!codigo_ibge
                     End If
               End If
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
'=================
         End If
      End If
      If TabEND.State = 1 Then _
         TabEND.Close

      'Transportadora

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from EMPRESA "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabEMP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEMP.EOF Then
         'SERIE NOTA SAIDA
         If Not IsNull(TabEMP!serie_NOTA_SAIDA) Then _
            txtSerie.Text = TabEMP!serie_NOTA_SAIDA

         If Not IsNull(TabEMP!Instrucao_Fisco) Then _
            txtMSG.Text = TabEMP!Instrucao_Fisco
            txtNota.Refresh
            Else
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               MsgBox "Erro no arquivo de empresa."
               Exit Sub
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      PEGA_DADOS_EMPRESA 'Atualiza o buffer com informacoes da empresa
      
      If UNIDADE_FEDERAÇÃO_EMPRESA = "" Then
         MsgBox "Impossivel continuar, inconsitencia na UNIDADE FEDERAÇÃO cadastrada para empresa."
         Exit Sub
         Else
            If Trim(TIPO_DOC) = "DC" Then
               If Trim(txtUF.Text) = Trim(UNIDADE_FEDERAÇÃO_EMPRESA) Then
                  cmbCFOPAux.Text = 5411        'rstEmpresa!CFOP_DV_ENT_DE  '5411
                  Else: cmbCFOPAux.Text = 6411  'rstEmpresa!CFOP_DV_ENT_FE  '6411
                End If
            End If
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from CFOP "
      SQL = SQL & " where codigo = '" & cmbCFOPAux.Text & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         cmbCFOP.Text = TabConsulta!Codigo & "-" & TabConsulta!Descricao
         txtNaturezaOperacao.Text = Trim(TabConsulta!Descricao)
         Else
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            MsgBox "CFOP não cadastrado."
            Exit Sub
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      Else
         If TabFOR.State = 1 Then _
            TabFOR.Close

         If TabCABENTRA.State = 1 Then _
            TabCABENTRA.Close

         MsgBox "Fornecedor Não Cadastrado !!!"
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         Frame6.Enabled = False
         'Frame7.Enabled = False
         Exit Sub
   End If
   If TabFOR.State = 1 Then _
      TabFOR.Close

   txtCNPJCPF.PromptInclude = False

   If TabEND.State = 1 Then _
      TabEND.Close

   SQL = "select * from FONE "
   SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
   SQL = SQL & " and numero <> ''"
   TabEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEND.EOF Then
      txtFone.Text = "" & Trim(TabEND.Fields("numero").Value)
      Else: MsgBox "Fone de cliente não encontrado !!!"
   End If

   If TabEND.State = 1 Then _
      TabEND.Close

   txtCNPJCPF.PromptInclude = True

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_FORNECEDOR"
End Sub

Private Sub TOTAIS_NOTA()
On Error GoTo erro_trata

   Dim dblVlrBaseICMS   As Double
   Dim dblVlrICMS       As Double
   Dim Desconto_Cabeça  As Double

   dblVlrBaseICMS = 0
   dblVlrICMS = 0
   VALOR_DESCONTO_N = 0

  'valor de desconto na cabeça
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'desconto individual por item
   SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from ITEMREQ "
   SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select valor_desconto from CABECAREQ "
   SQL = SQL & " where numr_req = " & TabCABECA!numr_req
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabConsulta.Fields(0).Value

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_ITEM_N = 0

   SQL = "select sum(valor_item*qtd_pedida) from ITEMREQ "
   SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_ITEM_N = TabConsulta.Fields(0).Value

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N

   txtValorTotalNota.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
   'txtValorProdutos.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   txtValorProdutos.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtBaseCalculo.Text = Format(dblVlrBaseICMS, strFormatacao2Digitos)
   txtValorICMS.Text = Format(dblVlrICMS, strFormatacao2Digitos)

   If VALOR_DESCONTO_N > 0 Then _
      txtDescDesconto.Text = "DESCONTO ESPECIAL DE " & _
         Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

   txtDescDesconto.Refresh

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "TOTAIS_NOTA"
End Sub

Private Sub TOTAIS_NOTA_Devolução() 'Devolução de Entrada
On Error GoTo erro_trata

   PERC_DESCONTO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0

   If TABITEM.State = 1 Then _
      TABITEM.Close

   SQL = "select sum(preco_custo*qtd_entrada) "

   SQL = SQL & " FROM NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"

   SQL = SQL & " where numr_pedido_compra = " & TabCABENTRA!numr_pedido_compra
   TABITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TABITEM.EOF Then _
      If Not IsNull(TABITEM.Fields(0).Value) Then _
         VALOR_ITEM_N = TABITEM.Fields(0).Value
   If TABITEM.State = 1 Then _
      TABITEM.Close

   txtValorTotalNota.Text = Format(VALOR_ITEM_N + TabCABENTRA!VALOR_IPI, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
   txtValorProdutos.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   txtBaseCalculo.Text = Format(TabCABENTRA!BASE_CALC_ICMS, strFormatacao2Digitos)
   txtValorICMS.Text = Format(TabCABENTRA!VALOR_ICMS, strFormatacao2Digitos)
   txtvaloripi.Text = Format(TabCABENTRA!VALOR_IPI, strFormatacao2Digitos)

   txtBaseIcmsSub.Text = Format(0, strFormatacao2Digitos)
   txtVlrIcmsSub.Text = Format(0, strFormatacao2Digitos)
   txtFrete.Text = Format(0, strFormatacao2Digitos)
   txtValorOutros.Text = Format(0, strFormatacao2Digitos)

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "TOTAIS_NOTA_Devolução"
End Sub

Private Sub GRID_DP()
On Error GoTo erro_trata

   Dim Desconto_Item_Fat   As Double
   Dim Valr_Item_Fat       As Double

   ListaDP.ListItems.Clear

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where l.numr_doc = " & TabCABECA!numr_req
   SQL = SQL & " and l.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and l.lancamento_id = i.lancamento_id "
   SQL = SQL & " and l.tipo_lancamento = 1 "
   SQL = SQL & " order by i.seq"
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLancamento.EOF
      If TabLancamento!FORMA_ID = 9999 Then
         ListaDP.ForeColor = vbRed
         Set Item = ListaDP.ListItems.Add(, "seq." & TabLancamento!seq, TabLancamento!seq)
         Else
            ListaDP.ForeColor = vbBlue
            Set Item = ListaDP.ListItems.Add(, "seq." & TabLancamento!seq, TabLancamento!seq)
      End If

Desconto_Item_Fat = 0 & TabLancamento!Valor_Desconto
Valr_Item_Fat = 0 & TabLancamento!Valor_Item

      Item.SubItems(1) = Format(Valr_Item_Fat - Desconto_Item_Fat, "currency")
      Item.SubItems(2) = TabLancamento!Numr_doc & "-" & TabLancamento!seq
      Item.SubItems(3) = TabLancamento!dt_Vencimento

      If Not IsNull(TabLancamento!FORMA_ID) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from FORMAPAGTO "
         SQL = SQL & " where forma_id = " & TabLancamento!FORMA_ID
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            Item.SubItems(4) = TabTemp!Descricao & " ; "
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from OBS "
      SQL = SQL & " where prop = '" & Trim(TabLancamento!Numr_doc) & "'"
      SQL = SQL & " and seq = " & TabLancamento!seq
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         Item.SubItems(4) = Trim(Item.SubItems(4)) & TabTemp!obs
      If TabTemp.State = 1 Then _
         TabTemp.Close

      TabLancamento.MoveNext
   Wend
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRID_DP"
End Sub

Private Sub GRID_DP_DEV() 'Devolução de Entrada
On Error GoTo erro_trata
  
Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRID_DP_DEV"
End Sub

Private Sub GRID_PRODUTOS()
On Error GoTo erro_trata

   Dim dblSequencia As Double
   Dim VALR_DIF_PERC As Double
   
   ListaProdutos.ListItems.Clear

   dblSequencia = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from ITEMREQ "
   SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      TabLancamento.MoveFirst
      
      While Not TabLancamento.EOF
         DoEvents
         dblSequencia = dblSequencia + 1

         If TabProduto.State = 1 Then _
            TabProduto.Close

         Set Item = ListaProdutos.ListItems.Add(, "seq." & dblSequencia, TabLancamento!Codg_Prod)

         SQL = "select * from PRODUTO "
         SQL = SQL & " where codg_produto = '" & TabLancamento!Codg_Prod & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            Item.SubItems(1) = Trim(TabProduto!Descricao)
            If Not IsNull(TabProduto!Unidade_Medida) Then _
               Item.SubItems(2) = TabProduto!Unidade_Medida
         End If

         If Not IsNull(TabLancamento!PERCICMS) Then _
            Item.SubItems(3) = TabLancamento!PERCICMS

         Item.SubItems(4) = TabLancamento!qtd_pedida

         If VALR_DIF_PERC > 0 Then
            Item.SubItems(5) = Format(TabLancamento!Valor_Item * VALR_DIF_PERC / 100, strFormatacao2Digitos)
            Item.SubItems(6) = Format(TabLancamento!qtd_pedida * TabLancamento!VALOR_TOTAL_ITEM * VALR_DIF_PERC / 100, strFormatacao2Digitos)
            Else
               Item.SubItems(5) = Format(TabLancamento!Valor_Item - (TabLancamento!Valor_Item * TabLancamento!PERC_desc / 100), strFormatacao2Digitos)
               Item.SubItems(6) = Format(TabLancamento!qtd_pedida * TabLancamento!Valor_Item - (TabLancamento!qtd_pedida * TabLancamento!Valor_Item * TabLancamento!PERC_desc / 100), strFormatacao2Digitos)
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

         Item.SubItems(7) = "" & TabLancamento.Fields("stributaria").Value
         Item.SubItems(8) = "" & TabLancamento.Fields("cfop").Value

         TabLancamento.MoveNext
      Wend
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRID_PRODUTOS"
End Sub

Private Sub GRID_PRODUTOS_DEV()
On Error GoTo erro_trata

   Dim dblSequencia As Double
   
   ListaProdutos.ListItems.Clear

   dblSequencia = 0
   If Trim(TIPO_DOC) = "DC" Then
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      SQL = "select NOTAENTRADAITEM.* "

   SQL = SQL & " FROM NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"

      SQL = SQL & " where numr_pedido_compra = " & TabCABENTRA!numr_pedido_compra
      SQL = SQL & " order by seq"
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         TabLancamento.MoveFirst

         While Not TabLancamento.EOF
            DoEvents
            dblSequencia = dblSequencia + 1
            Set Item = ListaProdutos.ListItems.Add(, "seq." & dblSequencia, TabLancamento!Codg_Prod)
      
            If TabProduto.State = 1 Then _
               TabProduto.Close
      
            SQL = "select * from PRODUTO "
            SQL = SQL & " where codg_produto = '" & TabLancamento!Codg_Prod & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then
               Item.SubItems(1) = Trim(TabProduto!Descricao)
               If Not IsNull(TabProduto!Unidade_Medida) Then _
                  Item.SubItems(2) = TabProduto!Unidade_Medida
            End If
            
            Item.SubItems(3) = TabLancamento!PERC_ICMS
            Item.SubItems(4) = TabLancamento!qtd_entrada
            Item.SubItems(5) = Format(TabLancamento!Preco_Custo, strFormatacao2Digitos)
            Item.SubItems(6) = Format(TabLancamento!qtd_entrada * TabLancamento!Preco_Custo, strFormatacao2Digitos)

            If TabProduto.State = 1 Then _
               TabProduto.Close

            TabLancamento.MoveNext
         Wend
      End If
      If TabLancamento.State = 1 Then _
         TabLancamento.Close
      Else 'Devolução de Vendas tipo "DV"
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select * from DEVITEMSAI "
         SQL = SQL & " where numr_req = " & TabCABECA!numr_req
         SQL = SQL & " order by seq"
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then
            TabLancamento.MoveFirst

            While Not TabLancamento.EOF
               DoEvents
               dblSequencia = dblSequencia + 1
               Set Item = ListaProdutos.ListItems.Add(, "seq." & dblSequencia, TabLancamento!Codg_Prod)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select * from PRODUTO "
               SQL = SQL & " where codg_prod = '" & TabLancamento!Codg_Prod & "'"
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and situacao <> 'C' "
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then
                  Item.SubItems(1) = Trim(TabProduto!Descricao)
                  If Not IsNull(TabProduto!Unidade_Medida) Then _
                     Item.SubItems(2) = TabProduto!Unidade_Medida
               End If
               
               Item.SubItems(3) = TabLancamento!PERC_ICMS
               Item.SubItems(4) = TabLancamento!qtd_devolucao
               Item.SubItems(5) = Format(TabLancamento!Preco_Venda, strFormatacao2Digitos)
               Item.SubItems(6) = Format(TabLancamento!qtd_devolucao * TabLancamento!Preco_Venda, strFormatacao2Digitos)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               TabLancamento.MoveNext
            Wend
         End If
         If TabLancamento.State = 1 Then _
            TabLancamento.Close
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRID_PRODUTOS_DEV"
End Sub

Private Sub GERA_NUMERO_NF()
On Error GoTo erro_trata

   Dim RstTemp          As New ADODB.Recordset
   Dim NUMR_NOTA_FISCAL As Double

   CRITERIO = EMPRESA_ID_N

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select numr_nota, NF_ID from NF "
   SQL = SQL & " where numr_nota = " & Trim(txtNota.Text)
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      NUMR_NOTA_FISCAL = TabNOTA.Fields(0).Value
      NUMR_ID_N = TabNOTA.Fields(1).Value
      Else
         NUMR_NOTA_FISCAL = MAX_ID("seq_nota_saida", "empresa", "empresa_id", CRITERIO, "", "")

         SQL = "update EMPRESA set "
         SQL = SQL & " seq_nota_saida = " & NUMR_NOTA_FISCAL
         SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         NUMR_ID_N = 1
         NUMR_ID_N = MAX_ID("nf_id", "nf", "", "", "", "")
   End If

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   txtNota.Text = NUMR_NOTA_FISCAL
   txtNota.Refresh
   txtNota.Enabled = False

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NUMERO_NF"
End Sub

Private Sub GRAVA_NOTA()
On Error GoTo erro_trata

   Dim CFOP_A As String

   CFOP_A = Left(cmbCFOP.Text, 4)

   GERA_NUMERO_NF

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF_TRANSP.PromptInclude = False

   PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NF "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      NUMR_ID_N = TabNOTA.Fields("nf_id").Value

      SQL = "delete from NFITEM "
      SQL = SQL & " where nf_id = " & NUMR_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "delete from NF "
      SQL = SQL & " where numr_nota = " & Trim(txtNota.Text)
      SQL = SQL & " and nf_id = " & NUMR_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value

   SQL = "INSERT INTO NF ("
      SQL = SQL & " PEDIDO_ID,EMPRESA_ID, NF_ID, NF_TIPO, NUMR_NOTA, SERIE_NOTA, "
      SQL = SQL & " PROP, NUMR_REQ, DT_EMISSAO, DT_ENTRASAI, TRANSP_ID, "
      SQL = SQL & " Qtd_Volume, Tipo_especie, Peso_Bruto, Peso_Liquido, status, CFOP,pessoa_id"
   SQL = SQL & " )"
   SQL = SQL & " VALUES ("
      SQL = SQL & NUMR_REQ_N
      SQL = SQL & "," & EMPRESA_ID_N
      SQL = SQL & "," & NUMR_ID_N
      SQL = SQL & ",'" & Trim(TIPO_DOC) & "'"
      SQL = SQL & "," & txtNota.Text
      SQL = SQL & ",'" & Trim(txtSerie.Text) & "'"
      SQL = SQL & ",'" & txtCNPJCPF.Text & "'"
      SQL = SQL & "," & NUMR_REQ_N
      SQL = SQL & ",'" & DMA(Date) & "'"
      SQL = SQL & ",'" & DMA(Date) & "'"
      
      'SQL = SQL & ",'" & txtCNPJCPF_TRANSP.Text & "'"
      SQL = SQL & ",1"
      
      SQL = SQL & "," & Replace(TxtQuantidadeRodapeNota.Text, ",", ".")
      SQL = SQL & ",'" & Replace(TxtEspecie.Text, ",", ".") & "'"
      SQL = SQL & "," & Replace(TxtPesoBruto.Text, ",", ".")
      SQL = SQL & "," & Replace(TxtPesoLiquido.Text, ",", ".")
      SQL = SQL & ",'" & "i" & "'"
      SQL = SQL & ",'" & CFOP_A & "'"
      SQL = SQL & "," & PESSOA_ID_N
   SQL = SQL & " )"
   CONECTA_RETAGUARDA.Execute SQL

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * from ITEMREQ "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      TabPedidoItem.MoveFirst
      While Not TabPedidoItem.EOF
         If TabAUX.State = 1 Then _
            TabAUX.Close

If Not IsNull(TabPedidoItem!cfop) Then _
   If Trim(TabPedidoItem!cfop) <> "" Then _
      If IsNumeric(TabPedidoItem!cfop) Then _
         CFOP_A = Trim(TabPedidoItem!cfop)

If Trim(TIPO_DOC) = "DV" Then _
   CFOP_A = Left(cmbCFOP.Text, 4)

         SQL = "select * from NFITEM "
         SQL = SQL & " where nf_id = " & NUMR_ID_N
         SQL = SQL & " and codg_prod = '" & Trim(TabPedidoItem!Codg_Prod) & "'"
         SQL = SQL & " and seq_id = " & TabPedidoItem.Fields("seq_id").Value
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            SQL = "UPDATE NFITEM SET "
               SQL = SQL & " Valor = " & tpMoeda(TabPedidoItem!Valor_Item - (TabPedidoItem!Valor_Item * TabPedidoItem!PERC_desc / 100))
               SQL = SQL & ", Desconto = " & tpMoeda((TabPedidoItem!Valor_Item * TabPedidoItem!qtd_pedida) * TabPedidoItem!PERC_desc / 100)
               SQL = SQL & ", Qtde = " & tpMoeda(TabPedidoItem!qtd_pedida)
               SQL = SQL & ", Cfop = '" & CFOP_A & "'"
               SQL = SQL & ", STRIBUTARIA = " & tpMoeda(TabPedidoItem!STRIBUTARIA)
               SQL = SQL & ", VlrBaseIcms = " & tpMoeda(TabPedidoItem!vlrbaseicms)
               SQL = SQL & ", PERCICMS = " & tpMoeda(TabPedidoItem!PERCICMS)
               SQL = SQL & ", VlrICMS = " & tpMoeda(TabPedidoItem!VlrIcms)
               SQL = SQL & ", VLRBASEICMSSUBST = " & tpMoeda(TabPedidoItem!VLRBASEICMSSUBST)
               SQL = SQL & ", PERCICMSSUBST = " & tpMoeda(TabPedidoItem!PERCICMSSUBST)
               SQL = SQL & ", VLRICMSSUBST = " & tpMoeda(TabPedidoItem!VLRICMSSUBST)
               SQL = SQL & ", PERCREDUCAOICMS = " & tpMoeda(TabPedidoItem!PERCREDUCAOICMS)
               SQL = SQL & ", PERCIVA = " & tpMoeda(TabPedidoItem!PercIva)
            SQL = SQL & " where nf_id = " & NUMR_ID_N
            SQL = SQL & " and codg_prod = '" & TabPedidoItem!Codg_Prod & "'"
            SQL = SQL & " and seq_id = " & TabPedidoItem.Fields("seq_id").Value
            Else
               SQL = "INSERT INTO NFITEM ("
                  SQL = SQL & "nf_id, seq_id, produto_id, Codg_Prod, Valor, Desconto, Qtde, Cfop, STRIBUTARIA, "
                  SQL = SQL & "VlrBaseIcms, PERCICMS, VlrICMS,  VLRBASEICMSSUBST, PERCICMSSUBST, "
                  SQL = SQL & "VLRICMSSUBST, PERCREDUCAOICMS, PERCIVA, PERC_IPI"
               SQL = SQL & ")"
               SQL = SQL & " VALUES ("
                  SQL = SQL & NUMR_ID_N                                                                                     'nf_id
                  SQL = SQL & "," & TabPedidoItem.Fields("seq_id").Value
                  SQL = SQL & "," & TabPedidoItem.Fields("produto_id").Value
                  SQL = SQL & ",'" & Trim(TabPedidoItem!Codg_Prod) & "'"                                                             'Codg_Prod
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!Valor_Item - (TabPedidoItem!Valor_Item * TabPedidoItem!PERC_desc / 100))   'Valor
                  SQL = SQL & "," & tpMoeda((TabPedidoItem!Valor_Item * TabPedidoItem!qtd_pedida) * TabPedidoItem!PERC_desc / 100)   'Desconto
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!qtd_pedida)                                                          'Qtde
                  SQL = SQL & ",'" & CFOP_A & "'"                                                                           'Cfop
                  SQL = SQL & ",'" & TabPedidoItem!STRIBUTARIA & "'"                                                           'STRIBUTARIA
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!vlrbaseicms)                                                         'VlrBaseIcms
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!PERCICMS)                                                            'PERCICMS
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!VlrIcms)                                                             'VlrICMS
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!VLRBASEICMSSUBST)                                                    'VLRBASEICMSSUBST
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!PERCICMSSUBST)                                                       'PERCICMSSUBST
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!VLRICMSSUBST)                                                        'VLRICMSSUBST
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!PERCREDUCAOICMS)                                                     'PERCREDUCAOICMS
                  SQL = SQL & "," & tpMoeda(TabPedidoItem!PercIva)                                                             'PERCIVA
                  SQL = SQL & "," & 0                                                                                       'PERC_IPI
               SQL = SQL & ")"
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         CONECTA_RETAGUARDA.Execute SQL

         'SQL = "update ITEMREQ set "
         'SQL = SQL & " status = 'N'"
         'SQL = SQL & " where pedido_id = " & NUMR_REQ_N
         'SQL = SQL & " and seq_id = " & TabPedidoItem.Fields("seq_id").Value
         'CONECTA_RETAGUARDA.Execute SQL

         TabPedidoItem.MoveNext
      Wend

      GRAVA_STATUS_EMITIDO
   End If
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_NOTA"
End Sub

Private Sub GRAVA_NOTA_Devolução_COMPRA() 'Devolução entrada
On Error GoTo erro_trata

   Dim strTributacao As String

   SQL3 = EMPRESA_ID_N
   NUMR_ID_N = 1
   NUMR_ID_N = MAX_ID("nf_id", "nf", "empresa_id", SQL3, "", "")

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF_TRANSP.PromptInclude = False
   
   If Not IsDate(txtDtEmis.Text) Then _
      txtDtEmis.Text = Date

   If Not IsDate(txtDtSaida.Text) Then _
      txtDtSaida.Text = Date

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF_TRANSP.PromptInclude = False

'=============================
   PESSOA_ID_N = 0

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "SELECT fornecedor.pessoa_id, FORNECEDOR.CGCCPF, FORNECEDOR.NOME "
   SQL = SQL & " FROM NOTAENTRADA "
   SQL = SQL & " INNER JOIN FORNECEDOR "
   SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

   SQL = SQL & " where tipoentrada_id = 1 "
   SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
   SQL = SQL & " and NOTAENTRADA.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and NOTAENTRADA.NUMR_PEDIDO_COMPRA = " & NUMR_REQ_N

   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then _
      PESSOA_ID_N = TabNOTA.Fields("pessoa_id").Value
'=============================

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NF "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabNOTA.EOF Then
      GERA_NUMERO_NF

      SQL = "INSERT INTO NF "
         SQL = SQL & " (EMPRESA_ID, NF_ID, NF_TIPO, NUMR_NOTA, SERIE_NOTA, PROP, "
         SQL = SQL & " NUMR_REQ, DT_EMISSAO, DT_ENTRASAI, TRANSP_ID, Qtd_Volume, "
         SQL = SQL & " Tipo_especie, Peso_bruto, Peso_Liquido, status, CFOP,"
         SQL = SQL & " pessoa_id, pedido_id)"
      SQL = SQL & " VALUES ("
         SQL = SQL & EMPRESA_ID_N
         SQL = SQL & "," & NUMR_ID_N
         SQL = SQL & ",'" & Trim(TIPO_DOC) & "'"
         SQL = SQL & "," & txtNota.Text
         SQL = SQL & ",'" & Trim(txtSerie.Text) & "'"
         SQL = SQL & ",'" & txtCNPJCPF.Text & "'"
         SQL = SQL & "," & NUMR_REQ_N
         SQL = SQL & ",'" & DMA(txtDtEmis.Text) & "'"
         SQL = SQL & ",'" & DMA(txtDtSaida.Text) & "'"
         SQL = SQL & ",'" & txtCNPJCPF_TRANSP.Text & "'"
         SQL = SQL & "," & Replace(TxtQuantidadeRodapeNota.Text, ",", ".")
         SQL = SQL & ",'" & Replace(TxtEspecie.Text, ",", ".") & "'"
         SQL = SQL & "," & Replace(TxtPesoBruto.Text, ",", ".")
         SQL = SQL & "," & Replace(TxtPesoLiquido.Text, ",", ".")
         SQL = SQL & ",'" & "D" & "'"
         SQL = SQL & ",'" & Left(cmbCFOP.Text, 4) & "'"
         SQL = SQL & "," & PESSOA_ID_N
         SQL = SQL & "," & NUMR_REQ_N
      SQL = SQL & " )"
      Else
         SQL = "UPDATE NF SET "
         SQL = SQL & "EMPRESA_ID = " & EMPRESA_ID_N
         SQL = SQL & ", NF_ID = " & NUMR_ID_N
         SQL = SQL & ", NF_TIPO = " & Trim(TIPO_DOC)
         SQL = SQL & ", NUMR_NOTA = " & txtNota.Text
         SQL = SQL & ", SERIE_NOTA = '" & Trim(txtSerie.Text) & "'"
         SQL = SQL & ", PROP = '" & txtCNPJCPF.Text & "'"
         SQL = SQL & ", DT_EMISSAO = '" & DMA(Date) & "'"
         SQL = SQL & ", DT_ENTRASAI = '" & DMA(Date) & "'"
         SQL = SQL & ", TRANSP_ID = '" & txtCNPJCPF_TRANSP.Text & "'"
         SQL = SQL & ", Qtd_Volume = " & Replace(TxtQuantidadeRodapeNota.Text, ",", ".")
         SQL = SQL & ", Tipo_Especie = '" & Replace(TxtEspecie.Text, ",", ".") & "'"
         SQL = SQL & ", Peso_bruto = " & Replace(TxtPesoBruto.Text, ",", ".")
         SQL = SQL & ", Peso_Liquido = " & Replace(TxtPesoLiquido.Text, ",", ".")
         SQL = SQL & ", Status = '" & "i" & "'"
         SQL = SQL & ", CFOP = '" & Left(cmbCFOP.Text, 4) & "'"
         SQL = SQL & " where numr_req = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * "

   SQL = SQL & " FROM NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"

   SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
   SQL = SQL & " order by seq "

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      TabTemp.MoveFirst
      While Not TabTemp.EOF
         PRODUTO_ID_N = 0

         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select * from PRODUTO "
         SQL = SQL & " where codg_produto = '" & Trim(TabTemp!Codg_Prod) & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            strTributacao = TabProduto!Situacao_Tributaria
            PRODUTO_ID_N = TabProduto.Fields("produto_id").Value
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from NFITEM "
         SQL = SQL & " where nf_id = " & NUMR_ID_N
         SQL = SQL & " and codg_prod = '" & Trim(TabTemp!Codg_Prod) & "'"
         SQL = SQL & " and seq_id = " & TabTemp.Fields("seq").Value
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            SQL = "UPDATE NFITEM SET "
            SQL = SQL & " nf_id = " & NUMR_ID_N
            SQL = SQL & ", Codg_Prod = '" & Trim(TabTemp!Codg_Prod) & "'"
            SQL = SQL & ", Valor = " & tpMoeda(TabTemp!Preco_Custo)
            SQL = SQL & ", Qtde = " & tpMoeda(TabTemp!qtd_entrada)
            SQL = SQL & ", Cfop = '" & Left(cmbCFOP.Text, 4) & "'"
            SQL = SQL & ", PERCICMS = " & tpMoeda(TabTemp!PERC_ICMS)
            SQL = SQL & ", Perc_ipi " & tpMoeda(TabTemp!perc_ipi)
            SQL = SQL & " STRIBUTARIA = '" & strTributacao & "'"
            SQL = SQL & " where nf_id = " & NUMR_ID_N
            SQL = SQL & " and codg_prod = '" & Trim(TabTemp!Codg_Prod) & "'"
            SQL = SQL & " and seq_id = " & TabTemp.Fields("seq").Value
            Else
               SQL3 = NUMR_ID_N
               NUMR_SEQ_N = MAX_ID("seq_id", "nfitem", "nf_id", SQL3, "", "")

               SQL = "INSERT INTO NFITEM "
               SQL = SQL & "(nf_id,seq_id,produto_id,Codg_Prod,Valor,Qtde,Cfop,PERCICMS,Perc_IPI,STRIBUTARIA) "
               SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & NUMR_SEQ_N
               SQL = SQL & "," & PRODUTO_ID_N
               SQL = SQL & ",'" & Trim(TabTemp!Codg_Prod) & "'"
               SQL = SQL & "," & tpMoeda(TabTemp!Preco_Custo)
               SQL = SQL & "," & tpMoeda(TabTemp!qtd_entrada)
               SQL = SQL & ",'" & Left(cmbCFOP.Text, 4) & "'"
               SQL = SQL & "," & tpMoeda(TabTemp!PERC_ICMS)
               SQL = SQL & "," & tpMoeda(TabTemp!perc_ipi)
               SQL = SQL & ",'" & strTributacao & "'"
               SQL = SQL & ")"
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         CONECTA_RETAGUARDA.Execute SQL

         BAIXA_ESTOQUE

         TabTemp.MoveNext
      Wend
      GRAVA_STATUS_EMITIDO
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_NOTA_Devolução_COMPRA"
End Sub

Private Sub GRAVA_STATUS_EMITIDO()
On Error GoTo erro_trata

   If Trim(TIPO_DOC) = "S" Then
      SQL = "update CABECAREQ set "
      SQL = SQL & "status = 3 " 'foi Gerado Nota Fiscal
      SQL = SQL & ", numr_doc = " & Trim(txtNota.Text)
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If Trim(TIPO_DOC) = "DC" Then
      SQL = "update NOTAENTRADA set "
      SQL = SQL & "tipoentrada_id = " & 2  'Foi devolvida com sucesso
      SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If Trim(TIPO_DOC) = "DV" Then
      SQL = "update CABECAREQ set "
      SQL = SQL & "status = 3 " 'foi Gerado Nota Fiscal
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and tipo_registro = 'D'"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_STATUS_EMITIDO"
End Sub

Private Sub BAIXA_ESTOQUE()
On Error GoTo erro_trata

   If Trim(TIPO_DOC) = "S" Then
      SQL = "update PRODUTO set "
      SQL = SQL & " qtde = qtde - " & Replace(TabTemp!qtd_pedida, ",", ".")
      SQL = SQL & ", qtde_retido = qtde_retido - " & Replace(TabTemp!qtd_pedida, ",", ".")
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(TabTemp!Codg_Prod) & "'"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If Trim(TIPO_DOC) = "DC" Then
      SQL = "update PRODUTO set "
      SQL = SQL & "qtde = qtde - " & Replace(TabTemp!qtd_entrada, ",", ".")
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(TabTemp!Codg_Prod) & "'"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If Trim(TIPO_DOC) = "DV" Then
      SQL = "update PRODUTO set "
      SQL = SQL & "qtde = qtde + " & Replace(TabTemp!Qtde, ",", ".")
      SQL = SQL & ", qtde_retido = qtde_retido - " & Replace(TabTemp!Qtde, ",", ".")
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(TabTemp!Codg_Prod) & "'"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_ESTOQUE"
End Sub
'=================== ROTINA DE IMPRESSÃO DE NOTA FISCAL
Private Sub IMPRESSAO_NF()
On Error GoTo erro_trata

   Dim Qtde_Item_Nf As Long
   Dim strCaminhoNF As String

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NF "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and numr_req = " & NUMR_REQ_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabNOTA.EOF Then _
      MsgBox "NF não encontrada."
   If Not TabNOTA.EOF Then
      If Not IsNull(TabNOTA!Status) Then
         If TabEMP.State = 1 Then _
            TabEMP.Close

         SQL = "Select CaminhoNFE from Empresa "
         SQL = SQL & " Where Empresa_id = " & EMPRESA_ID_N
         TabEMP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabEMP.EOF Then _
            strCaminhoNF = TabEMP!CaminhoNFE

         If TabEMP.State = 1 Then _
            TabEMP.Close

         Open strCaminhoNF & "nf" & txtNota.Text & "-nfe.txt" For Output As #1
         'Open "c:\nf" & txtNota.Text & "-nfe.txt" For Output As #1

         GERAR_CABEÇALHO_NFe

         GERAR_PRODUTOS_NFe

         GERAR_TOTAIS_NFe

         GERAR_TRANSPORTADORA_NFe

         GERAR_FATURAS_NFe

         GERAR_RODAPE_NFe

         Close #1

         SQL = "update NF set "
         SQL = SQL & "status = 'E' "
         SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and numr_req = " & NUMR_REQ_N
         CONECTA_RETAGUARDA.Execute SQL

         MsgBox "Arquivo Gerado com Sucesso!? ", vbExclamation, "NFE"
      End If
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "IMPRESSAO_NF"
End Sub
'=============NOTA TRANSFERENCIA
Private Sub MONTA_TRANSFERENCIA()
On Error GoTo erro_trata

   SQL = "select * from PEDIDONOTA "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      TabNOTA!EMPRESA_ID = EMPRESA_ID_N
      TabNOTA!pedido_id = NUMR_REQ_N
      TabNOTA!DT_EMIS = txtDtEmis.Text
      If Trim(TIPO_DOC) = "T" Then _
         TabNOTA!TIPO_PEDIDO = 1 '1 no parametro é transferencia
      txtCNPJCPF.PromptInclude = False
      TabNOTA!CGCCPF = txtCNPJCPF.Text
      TabNOTA!Status = "A"
      TabNOTA!CODG_USU = CODG_USU_N
      TabNOTA.Update
      Else: MsgBox "Pedido de N.F. não registrado"
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_TRANSFERENCIA"
End Sub

Private Sub MONTA_NOTA_Devolução_SAIDA() 'Devolução de saida de mecadorias
On Error GoTo erro_trata

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from CABECAREQ "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and tipo_registro = '" & "D" & "'" 'Devolução de saida
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      'CHECA TIPO VENDA
      If Indr_Consulta <> True Then
         If TabCABECA!Status = 3 Then
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            MsgBox "Nota fiscal já emitida para essa Devolução."
            fraNota.Enabled = False
            fraEmitente.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = True
            Frame5.Enabled = True
            'Frame7.Enabled = False
            Frame8.Enabled = False
            Exit Sub
         End If
      End If

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      'passou do cabecareq, checar na tabela nf agora
      SQL = "select * from NF "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         txtNota.Text = TabNOTA!numr_Nota
         txtSerie.Text = TabNOTA!serie_nota
         txtDtEmis.Text = Format(TabNOTA!dt_emissao, "dd/mm/yyyy")
         txtDtSaida.Text = Format(TabNOTA!DT_ENTRASAI, "dd/mm/yyyy")
         If Not IsNull(TabNOTA!TRANSP_ID) Then
           If TabTemp.State = 1 Then _
              TabTemp.Close

           SQL = "select cgccpf,nome from TRANSPORTADORA "
           SQL = SQL & " where cgccpf = '" & TabNOTA!TRANSP_ID & "'"
           TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
           If Not TabTemp.EOF Then
              If Not IsNull(TabTemp.Fields(0).Value) Then
                 cmbCNPJCPF_TRANSP.Text = Trim(TabTemp!CGCCPF) & " - " & Trim(TabTemp!NOME)
                 txtCNPJCPF_TRANSP.Text = TabTemp.Fields(0).Value
                 'Volumes
                 If TabNOTA!qtd_volume <> "" Then _
                    TxtQuantidadeRodapeNota.Text = TabNOTA!qtd_volume

                 If TabNOTA!TIPO_ESPECIE <> "" Then _
                    TxtEspecie.Text = TabNOTA!TIPO_ESPECIE

                 If TabNOTA!PESO_BRUTO <> "" Then _
                    TxtPesoBruto.Text = TabNOTA!PESO_BRUTO

                 If TabNOTA!PESO_LIQUIDO <> "" Then _
                    TxtPesoLiquido.Text = TabNOTA!PESO_LIQUIDO
              End If
           End If
           If TabTemp.State = 1 Then _
              TabTemp.Close
         End If
         If Not IsNull(TabNOTA!cfop) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from CFOP "
            SQL = SQL & " where CODIGO = '" & TabNOTA!cfop & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbCFOP.Text = TabTemp!Codigo & "-" & Trim(TabTemp!Descricao)
               txtNaturezaOperacao.Text = Trim(TabTemp!Descricao)
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         MOSTRA_NOTA_TELA

         MsgBox "Já existe nota fiscal emitida para Esta Devolução = " & NUMR_REQ_N & " ; Nota Fiscal = " & TabNOTA!numr_Nota & " ; Empresa = " & EMPRESA_ID_N

         If TabNOTA.State = 1 Then _
            TabNOTA.Close
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

          fraNota.Enabled = False
          fraEmitente.Enabled = False
          Frame3.Enabled = False
          Frame4.Enabled = True
          Frame5.Enabled = True
          'Frame6.Enabled = False
          'Frame7.Enabled = False
          Frame8.Enabled = False
          Exit Sub
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      MOSTRA_NOTA_TELA

      Else
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      MsgBox "Registro de Devolução não encontrado."

      fraNota.Enabled = False
      fraEmitente.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = True
      Frame5.Enabled = True
      Frame6.Enabled = False
      'Frame7.Enabled = False
      Exit Sub
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_NOTA_Devolução_SAIDA"
End Sub

Private Sub MOSTRA_Devolução_TELA_SAIDA() 'Devolução de saida
On Error GoTo erro_trata

   If txtDtEmis.Text = "" Then txtDtEmis.Text = Format(Date, "dd/mm/yyyy")
   If txtDtSaida.Text = "" Then txtDtSaida.Text = Format(Date, "dd/mm/yyyy")
       
   If txtSerie.Text = "" Then
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & NUMR_REQ_N
      SQL = "select seq_nota_saida,serie_nota_saida from EMPRESA "
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then _
         txtSerie.Text = TabNOTA.Fields(1).Value

      If TabNOTA.State = 1 Then _
         TabNOTA.Close
   End If

   MOSTRA_CLIENTE
   TOTAIS_NOTA_Devolução_SAIDA
   GRID_DP
   GRID_PRODUTOS_DEV

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_Devolução_TELA_SAIDA"
End Sub
Private Sub TOTAIS_NOTA_Devolução_SAIDA() 'Devolução de saida
On Error GoTo erro_trata

   PERC_DESCONTO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
    
   SQL = "select sum(preco_venda*qtd_devolucao) from DEVITEMSAI "
   SQL = SQL & " where numr_req = " & TabCABECA!numr_req
   TABITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TABITEM.EOF Then _
      If Not IsNull(TABITEM.Fields(0).Value) Then _
         VALOR_ITEM_N = TABITEM.Fields(0).Value
   If TABITEM.State = 1 Then _
      TABITEM.Close
   
   txtValorTotalNota.Text = Format(VALOR_ITEM_N + TabCABECA!VALOR_IPI, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
   txtValorProdutos.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   txtBaseCalculo.Text = Format(TabCABECA!BASE_CALC_ICMS, strFormatacao2Digitos)
   txtValorICMS.Text = Format(TabCABECA!VALOR_ICMS, strFormatacao2Digitos)
   txtvaloripi.Text = Format(TabCABECA!VALOR_IPI, strFormatacao2Digitos)

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "TOTAIS_NOTA_Devolução_SAIDA"
End Sub

Private Sub PEGA_DADOS_EMPRESA()
On Error GoTo erro_trata

   If rstEmpresa.State = 1 Then _
      rstEmpresa.Close

   SQL = "Select * From EMPRESA "
   SQL = SQL & " Where EMPRESA_ID = " & EMPRESA_ID_N
   rstEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstEmpresa.EOF Then
      If rstEmpresa.State = 1 Then _
         rstEmpresa.Close

      MsgBox "O sistema não localizaou a empresa de codigo " & EMPRESA_ID_N & "não posso continuar.", vbCritical
      Exit Sub
      Unload Me
      Else
         booOptanteSimples = rstEmpresa!Empresa_Optante_Simples

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from OBS"
         SQL = SQL & " where prop = '" & Trim(rstEmpresa.Fields("CGC").Value) & "'"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabConsulta.EOF
            If TabConsulta.Fields("seq").Value = 1 Then _
               txtDescDesconto.Text = "" & Trim(TabConsulta.Fields("obs").Value)
            If TabConsulta.Fields("seq").Value = 2 Then _
               txtMSG.Text = "" & Trim(TabConsulta.Fields("obs").Value)
            If TabConsulta.Fields("seq").Value = 3 Then _
               txtDadosAdicionais.Text = "" & Trim(TabConsulta.Fields("obs").Value)
            TabConsulta.MoveNext
         Wend

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "PEGA_DADOS_EMPRESA"
End Sub

Sub GERAR_NFE()
On Error GoTo erro_trata

   Dim VALOR_01      As Double
   Dim VALOR_02      As Double
   Dim strTributacao As String

   'validar transportadora
   If Trim(cmbCNPJCPF_TRANSP.Text) = "" Then
      MsgBox "Informar trasportadora."
      txtCNPJCPF_TRANSP.SetFocus
      Exit Sub
   End If

   txtCNPJCPF_TRANSP.PromptInclude = False

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TRANSPORTADORA "
   SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      MsgBox "Informar trasportadora."
      txtCNPJCPF_TRANSP.SetFocus
      Exit Sub
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'validar cliente
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "CNPJ/CPF inválido !!!"
      Exit Sub
   End If
   If Trim(txtEnd.Text) = "" Then
      MsgBox "Endereço inválido !!!"
      txtEnd.SetFocus
      Exit Sub
   End If
   If Trim(txtBairro.Text) = "" Then
      MsgBox "Bairro inválido !!!"
      txtBairro.SetFocus
      Exit Sub
   End If
   If Trim(txtUF.Text) = "" Then
      MsgBox "UF inválido !!!"
      txtUF.SetFocus
      Exit Sub
   End If
   If Trim(txtCep.Text) = "" Then
      MsgBox "CEP inválido !!!"
      txtCep.SetFocus
      Exit Sub
   End If

   CRITERIO = txtCep.Text
   CRITERIO = Replace(CRITERIO, "-", "")
   If Len(Trim(CRITERIO)) < 8 Then
      MsgBox "CEP inválido, deve conter 8 digitos !!!"
      txtCep.SetFocus
      Exit Sub
   End If
   If Trim(cmbIE.Text) = "" Then
      MsgBox "IE inválido !!!"
      cmbIE.SetFocus
      Exit Sub
   End If
   If Trim(txtCidade.Text) = "" Then
      MsgBox "Cidade inválido !!!"
      txtCidade.SetFocus
      Exit Sub
   End If

   If Trim(txtIBGE.Text) = "" Then
      MsgBox "Código IBGE inválido !!!"
      txtIBGE.SetFocus
      Exit Sub
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from IBGE"
         SQL = SQL & " where codigo_ibge = " & txtIBGE.Text
         'SQL = SQL & " and municipio = '" & Trim(txtCidade.Text) & "'"
         'SQL = SQL & " and estado = '" & Trim(txtUF.Text) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            MsgBox "Erro no IBGE, verificar."
            Exit Sub
            Else
               If Not IsNull(TabTemp.Fields("estado").Value) Then
                  If Trim(UCase(TabTemp.Fields("estado").Value)) <> Trim(UCase(txtUF.Text)) Then
                     MsgBox "Erro no IBGE, verificar."
                     Exit Sub
                  End If
                  Else
                     MsgBox "Erro no IBGE, verificar."
                     Exit Sub
               End If
               If Not IsNull(TabTemp.Fields("municipio").Value) Then
                  'If Trim(UCase(TABTEMP.Fields("municipio").Value)) <> Trim(UCase(txtCidade.Text)) Then
                  '   MsgBox "Erro no IBGE, verificar."
                  '   Exit Sub
                  'End If
                  Else
                     MsgBox "Erro no IBGE, verificar."
                     Exit Sub
               End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
   End If

   If Trim(txtFone.Text) = "" Then
      MsgBox "Fone inválido !!!"
      txtFone.Enabled = True
      txtFone.SetFocus
      Exit Sub
      Else
         If Len(Trim(txtFone.Text)) <> 8 Then
            MsgBox "Fone inválido, deve conter 8 digitos !!!"
            txtFone.SetFocus
            Exit Sub
         End If
   End If

   If TxtQuantidadeRodapeNota.Text = "" Then
      MsgBox "Digite Quantidade de Produtos a Transportar!"
      TxtQuantidadeRodapeNota.SetFocus
      Exit Sub
      Else
         If Not IsNumeric(TxtQuantidadeRodapeNota.Text) Then
            MsgBox "Favor Digite Quantidade de Produtos a Transportar!"
            TxtQuantidadeRodapeNota.SetFocus
            Exit Sub
         End If
   End If

   If TxtEspecie.Text = "" Then
      MsgBox "Digite a Especie a Transportar!"
      TxtEspecie.SetFocus
      Exit Sub
   End If

   If TxtPesoBruto.Text = "" Then
      MsgBox "Digite o Peso Bruto a Transportar!"
      TxtPesoBruto.SetFocus
      Exit Sub
   End If

   If TxtPesoLiquido.Text = "" Then
      MsgBox "Digite o Peso Liquido a Transportar!"
      TxtPesoLiquido.SetFocus
      Exit Sub
   End If

   If Trim(TIPO_DOC) = "DV" Then
      If Indr_Consulta = False Then
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select * from CABECAREQ "
         SQL = SQL & " where numr_req = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and tipo_registro = 'D'"
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            txtCNPJCPF.PromptInclude = False

            SQL = "select * from CLIENTE "
            SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
            SQL = SQL & " and status = 'A'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then
               GRAVA_NOTA

               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from NFITEM"
               SQL = SQL & " where nf_id = " & NUMR_ID_N

               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  TabTemp.MoveFirst
                  While Not TabTemp.EOF
                     If TabProduto.State = 1 Then _
                        TabProduto.Close
         
                     SQL = "select * from PRODUTO " 'Ler Tab. Produtos Para pegar tributacao e nacionalidade
                     SQL = SQL & " where codg_produto = '" & Trim(TabTemp!Codg_Prod) & "'"
                     SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                     SQL = SQL & " and situacao <> 'C' "
                     TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabProduto.EOF Then _
                        strTributacao = TabProduto!Situacao_Tributaria
         
                     If TabProduto.State = 1 Then _
                        TabProduto.Close
         
                     If TabAUX.State = 1 Then _
                        TabAUX.Close
         
BAIXA_ESTOQUE

                     TabTemp.MoveNext
                  Wend
                  GRAVA_STATUS_EMITIDO
               End If

               IMPRESSAO_NF
               Unload Me
            End If
         End If
         Else
            IMPRESSAO_NF
            Unload Me
      End If
   End If

   If Trim(TIPO_DOC) = "DC" Then
      GRAVA_NOTA_Devolução_COMPRA

      IMPRESSAO_NF

      Unload Me
   End If

   If Trim(TIPO_DOC) = "S" Then
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      'chegando itens
      SQL = "select * from CABECAREQ "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCABECA.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from ITEMREQ "
         SQL = SQL & " where numr_req = " & TabCABECA.Fields("numr_req").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close
            MsgBox "Pedido não possue itens, não permitido."
            Exit Sub
         End If
         While Not TabTemp.EOF
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from PRODUTO "
            SQL = SQL & " where codg_produto = '" & Trim(TabTemp.Fields("codg_prod").Value) & "'"
            SQL = SQL & " and situacao <> 'C' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabConsulta.EOF Then
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               MsgBox "Pedido Produto não cadastrado, não permitido."
               Exit Sub
               Else
                  If IsNull(TabConsulta.Fields("situacao").Value) Then
                     MsgBox "Situação do produto inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        If Trim(TabConsulta.Fields("situacao").Value) <> "A" Then
                           MsgBox "Situação do produto inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                        End If
                  End If
                  QTDE_PEDIDO = TabConsulta.Fields("qtde").Value
                  If Indr_Consulta = False Then
                     If IsNull(TabConsulta.Fields("qtde").Value) Then
                        MsgBox "Qtde disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                        Exit Sub
                        Else
                           If QTDE_PEDIDO <= 0 Then
                              MsgBox "Qtde disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                              Exit Sub
                           End If
                     End If
                  End If
                  If IsNull(TabConsulta.Fields("PRECO_VENDA").Value) Then
                     MsgBox "Valor_venda disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        VALOR_ITEM_N = TabConsulta.Fields("PRECO_VENDA").Value
                        If VALOR_ITEM_N <= 0 Then
                           MsgBox "Valor_venda disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                           Else
                              VALOR_01 = TabConsulta.Fields("PRECO_VENDA").Value
                              VALOR_02 = TabConsulta.Fields("PRECO_CUSTO").Value
                              If VALOR_01 <= VALOR_02 Then
                                 MsgBox "Produto : " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value) & " , venda abaixo preço de custo."
                              End If
                        End If
                  End If
                  If IsNull(TabConsulta.Fields("CODG_NCM").Value) Then
                     MsgBox "CODG_NCM disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        VALOR_01 = TabConsulta.Fields("CODG_NCM").Value
                        If VALOR_01 <= 0 Then
                           MsgBox "CODG_NCM disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                        End If
                  End If
            End If

            TabTemp.MoveNext

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Wend

         If TabTemp.State = 1 Then _
            TabTemp.Close
         Else
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
            MsgBox "Pedido não encontrado."
            Exit Sub
      End If
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      'If Indr_Consulta = False Then
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select * from CABECAREQ "
         SQL = SQL & " where numr_req = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            txtCNPJCPF.PromptInclude = False

            SQL = "select * from CLIENTE "
            SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
            SQL = SQL & " and status = 'A'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then
               GRAVA_NOTA
               IMPRESSAO_NF
            End If
            If TabCliente.State = 1 Then _
               TabCliente.Close
         End If
         If booUsaCobranca = True Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            CONECTA_RETAGUARDA.Execute "UPDATE Lancamento SET CodigoContaCorrente = 1 WHERE Tipo_Lancamento = 1"
            TabTemp.Open "SELECT ItemLancamento.forma_id FROM  Lancamento INNER JOIN ContaCorrente ON Lancamento.empresa_id = ContaCorrente.Empresa_id AND  Lancamento.CodigoContaCorrente = ContaCorrente.CodigoContaCorrente INNER JOIN ItemLancamento ON Lancamento.Lancamento_id = ItemLancamento.Lancamento_Id AND Lancamento.Numr_doc = ItemLancamento.Numr_doc WHERE (Lancamento.Empresa_id = " & EMPRESA_ID_N & ") AND (ItemLancamento.Numr_doc in (" & NUMR_REQ_N & "))  and (ItemLancamento.Forma_id = 5) and (ContaCorrente.EmiteBoleto = 1)", CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               If TabTemp!FORMA_ID <> 5 Then _
                  Exit Sub
            If TabTemp.State = 1 Then _
               TabTemp.Close
            Else
               Unload Me
               Exit Sub
         End If

         txtCNPJCPF.PromptInclude = False
         If Trim(txtCNPJCPF.Text) <> "99999999999" Then
            Msg = "Deseja Emitir Boleto Desta Venda? " & NUMR_REQ_N
            Style = vbYesNo + 32
            Title = "Atenção."
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then
               'Atualizando Codigo da Conta onde for boleto de contas a receber
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               TabTemp.Open "SELECT ItemLancamento.Lancamento_id, ItemLancamento.Numr_doc, ItemLancamento.Seq, ContaCorrente.COMRegistro, Lancamento.CodigoContaCorrente, isnull(ItemLancamento.NossoNumero, '') as NossoNumero, ItemLancamento.Dt_Vencimento FROM  Lancamento INNER JOIN ContaCorrente ON Lancamento.empresa_id = ContaCorrente.Empresa_id AND  Lancamento.CodigoContaCorrente = ContaCorrente.CodigoContaCorrente INNER JOIN ItemLancamento ON Lancamento.Lancamento_id = ItemLancamento.Lancamento_Id AND Lancamento.Numr_doc = ItemLancamento.Numr_doc WHERE (Lancamento.Empresa_id = " & EMPRESA_ID_N & ") AND (ItemLancamento.Numr_doc in (" & NUMR_REQ_N & "))  and (ItemLancamento.Forma_id = 5) and (ContaCorrente.EmiteBoleto = 1)", CONECTA_RETAGUARDA, , , adCmdText
               Do While Not TabTemp.EOF
                  If TabTemp!NossoNumero = "" Or TabTemp!NossoNumero = "0" Then
                     NossoNumero = CalculaNossonumero(TabTemp!CodigoContaCorrente, TabTemp!Numr_doc & Right(TabTemp!seq, 1), TabTemp!dt_Vencimento)
                     If NossoNumero = "" Then
                        MsgBox "Atenção. Boleto não gerou nosso número", 48, Me.Caption

                        If TabTemp.State = 1 Then _
                           TabTemp.Close
                        Exit Sub
                        Else
                           CONECTA_RETAGUARDA.Execute "UPDATE ItemLancamento SET NossoNumero = '" & NossoNumero & "' WHERE Lancamento_id = " & TabTemp!lancamento_id & " and Numr_doc = " & TabTemp!Numr_doc & " and Seq = " & TabTemp!seq
                           CONECTA_RETAGUARDA.Execute "UPDATE ContaCorrente SET SequenciaBoleta = SequenciaBoleta + 1 WHERE Empresa_id = " & EMPRESA_ID_N & " and CodigoContaCorrente = " & TabTemp!CodigoContaCorrente
                     End If
                  End If
                  TabTemp.MoveNext
               Loop
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               FORMULA_REL = "{LANCAMENTO.Tipo_Lancamento} = " & 1
               FORMULA_REL = FORMULA_REL & " and {ITEMLANCAMENTO.Numr_Doc} = " & NUMR_REQ_N
ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
               Nome_Relatorio = "BoletoItau.rpt"
               frmRELATORIO10.Show 1
            End If
         End If
         Unload Me
      'End If
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_NFE"
End Sub

Sub INICIALIZA_NF()
On Error GoTo erro_trata

   Dim TOTAL_N As Long

   TIPO_DOC = ""
   fraNota.Enabled = True
   fraEmitente.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   Frame5.Enabled = True
   Frame6.Enabled = True
   Frame7.Enabled = True

   TOTAL_N = 0

   Me.Caption = Me.Caption & " - " & Me.Name

   cmbCFOPAux.Clear
   cmbCFOP.Clear

   If TabEMP.State = 1 Then _
      TabEMP.Close

   SQL = "select * from EMPRESA "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   TabEMP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEMP.EOF Then _
      strCNPJEMPRESA = TabEMP!CGC

   If TabEMP.State = 1 Then _
      TabEMP.Close
   
   CARREGA_TRANSPORTADORA
   CARREGA_CFOP

   OptFreteDestinatario.Value = False
   OptFreteEmitente = True

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT CEP.UF FROM (ENDERECO INNER JOIN EMPRESA ON ENDERECO.PROP = EMPRESA.CGC) "
   SQL = SQL & " INNER JOIN CEP ON ENDERECO.CEP = CEP.Cep "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         UNIDADE_FEDERAÇÃO_EMPRESA = TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT TIPO_DOC, TIPO_REGISTRO From CABECAREQ "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and Numr_Req = " & NUMR_REQ_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If TabTemp!TIPO_REGISTRO <> "D" Then
         If Not IsNull(TabTemp!TIPO_DOC) Then
            TIPO_DOC = TabTemp!TIPO_DOC
            Else: TIPO_DOC = "S"           'Tipo Saida
         End If
         Else: TIPO_DOC = "DV"        'Tipo Devolução de Saida
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TIPO_DOC = "" Then 'Sera uma Devolução de Entrada
      SQL = "SELECT status From notaentrada "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and Numr_pedido_compra = " & NUMR_REQ_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp!Status) Then _
            TIPO_DOC = "DC"

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   'NOTA VENDA
   If Trim(TIPO_DOC) = "S" Then _
      If (NUMR_REQ_N > 0) Then _
         MONTA_NOTA_SAIDA

   'NOTA DE TRANSFERENCIA
   If Trim(TIPO_DOC) = "T" Then _
      If ((NUMR_REQ_N > 0) And (EMPRESA_ID_N > 0)) Then _
         MONTA_TRANSFERENCIA

   'NOTA DE DEVOLUÇÃO
   If Trim(TIPO_DOC) = "DC" Then _
      If ((NUMR_REQ_N > 0) And (EMPRESA_ID_N > 0)) Then _
         MONTA_NOTA_Devolução

   'NOTA DE Devolução de saida
   If Trim(TIPO_DOC) = "DV" Then _
      If ((NUMR_REQ_N > 0) And (EMPRESA_ID_N > 0)) Then _
         MONTA_NOTA_Devolução_SAIDA

   'If Trim(TIPO_DOC) = "E" Then   'NOTA DE ENTRADA
   'End If
   'If Trim(TIPO_DOC) = "R" Then   'NOTA DE SIMPLES REMESSA
   'End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_NF"
End Sub

Private Sub CARREGA_TRANSPORTADORA()
On Error GoTo erro_trata

   If TabFORNEC.State = 1 Then _
      TabFORNEC.Close

   TabFORNEC.Open "Select * from TRANSPORTADORA Where CGCCPF = '" & strCNPJEMPRESA & "'", CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFORNEC.EOF Then
      txtCNPJCPF_TRANSP.PromptInclude = False
      txtCNPJCPF_TRANSP.Text = TabFORNEC!CGCCPF
      cmbCNPJCPF_TRANSP.Text = TabFORNEC!NOME
      txtCNPJCPF_TRANSP.PromptInclude = True
   End If
   If TabFORNEC.State = 1 Then _
      TabFORNEC.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_TRANSPORTADORA"
End Sub

Private Sub CARREGA_CFOP()
On Error GoTo erro_trata

   'CFOP
   cmbCFOPAux.Clear
   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      Do Until TabDESCR.EOF
         DoEvents
         'cmbCFOPAUX.AddItem TABDESCR!Descricao
         cmbCFOPAux.AddItem Trim(TabDESCR!Codigo)
         cmbCFOP.AddItem Trim(TabDESCR!Codigo) & "-" & Trim(TabDESCR!Descricao)
         TabDESCR.MoveNext
      Loop
      Else
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         MsgBox "Cadastro de CFOP com problemas. Não foi localizado nenhum codigo de CFOP cadastrado", vbCritical
         Exit Sub
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_CFOP"
End Sub
'================================================================
Private Sub GERAR_CABEÇALHO_NFe()
On Error GoTo erro_trata

   Dim RstTemp          As New ADODB.Recordset
   Dim rstTemp2         As New ADODB.Recordset
   Dim NOME_A           As String
   Dim lngCont          As Long
   Dim Numero_A         As String
   Dim DDD_N            As Integer
   Dim CODG_IBGE_A      As String
   Dim Tipo_Endereço    As String * 1, FONE_A As String, IE_A As String

   Print #1, Tab(1); "NOTAFISCAL|1";   'saida
   Print #1, Tab(1); "A|2.00|NFe52110310628919000188550010000000011203740010|";

   'PEGA_DADOS_EMPRESA

   SP_PROCURA_ENDEREÇONFE rstEmpresa!CGC, "C", "", 0

   If TabEND.EOF Then
      MsgBox "Não achou endereço da empresa, verifique."
      Exit Sub
   End If

   SQL = "select * From Endereco Where prop = '" & rstEmpresa!CGC & "'"

   SP_PROCURA_CEP TabEND!CEP

   CODG_IBGE_A = "" & TabCEP!codigo_ibge

   If Len(CODG_IBGE_A) < 7 Then
      MsgBox "IBGE errado para o cep = " & TabEND!CEP
      Unload Me
      Exit Sub
   End If

   If TIPO_DOC = "S" Then
      Print #1, Tab(1); "B|52|" & NUMR_REQ_N & "|" & Trim(txtNaturezaOperacao.Text) & "|0|55|1|" & TabNOTA!numr_Nota & "|" & Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & "|" & Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & "|" & Time & "|1|" & Trim(CODG_IBGE_A) & "|1|1||2|1|3|2.0.7|||";
   ElseIf TIPO_DOC = "DC" Then
      Print #1, Tab(1); "B|52|" & NUMR_REQ_N & "|" & Trim(txtNaturezaOperacao.Text) & "|0|55|1|" & TabNOTA!numr_Nota & "|" & Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & "|" & Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & "|" & Time & "|1|" & Trim(CODG_IBGE_A) & "|1|1||2|1|3|2.0.7|||";
   ElseIf TIPO_DOC = "DV" Then
      Print #1, Tab(1); "B|52|" & NUMR_REQ_N & "|" & Trim(txtNaturezaOperacao.Text) & "|0|55|1|" & TabNOTA!numr_Nota & "|" & Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & "|" & Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & "|" & Time & "|1|" & Trim(CODG_IBGE_A) & "|1|1||2|1|3|2.0.7|||";
   End If

   'encerra Registro B52
   'Começa registro B13 a Chave que depois eu vou gerar aqui
   'Print #1, Tab(1); "B|13|0000000000000000000000000000000000000000000"
   'Registro Codigo do estado ano e mes conforme yuri nao passar
   'Print #1, Tab(1); "B|14|35|" & Mid(Date, 7, 4) & Mid(Date, 4, 2);
   'Dados do Emitente
   'Colocar Campo Inscricao Municipaç empresa
   Print #1, Tab(1); "C|" & rstEmpresa!RAZAO_SOCIAL & "|" & rstEmpresa!RAZAO_SOCIAL & "|" & rstEmpresa!IE & "||      ||" & rstEmpresa!Tipo_Regime_Empresa & "|"
   Print #1, Tab(1); "C02|" & rstEmpresa!CGC

   'Buscar Endereco
   'arrumae isso TABEND!Numero
   Print #1, Tab(1); "C05|" & TabEND!Rua & "|" & "0" & "||" & TabEND!Bairro & "|" & TabCEP!codigo_ibge & "|" & TabCEP!Cidade & "|" & TabCEP!UF & "|" & Trim(txtCep.Text) & "|1058|BRASIL|" & rstEmpresa!fone;

   'Dados do Destinatario
   If TIPO_DOC = "DC" Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where cgccpf = '" & TabNOTA!prop & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         Print #1, Tab(1); "E|" & TabCliente!NOME & "|" & UCase(TabCliente!IE) & "|||";
         If Len(TabCliente!CGCCPF) > 11 Then
            Print #1, Tab(1); "E02|" & TabCliente!CGCCPF;
            Else 'Nesta Caso e CNPJ
               Print #1, Tab(1); "E03|" & TabCliente!CGCCPF;
         End If

         SP_PROCURA_ENDEREÇONFE TabNOTA!prop, "C", "", 0
         If Not TabEND.EOF Then
            Numero_A = "000"
            If Not IsNull(TabEND.Fields("numero").Value) Then _
               If Trim(TabEND.Fields("numero").Value) <> "" Then _
                  Numero_A = Trim(TabEND.Fields("numero").Value)

            SP_PROCURA_CEP TabEND!CEP
            SP_PROCURA_FONE TabEND!prop, 0

            Print #1, Tab(1); "E05|" & TabEND!Rua & "|" & Numero_A & "| |" & TabEND!Bairro & "|" & TabCEP!codigo_ibge & "|" & TabCEP!Cidade & "|" & TabCEP!UF & "|" & Trim(txtCep.Text) & "|1058|BRASIL|" & Right(TabFone!DDD, 2) & TabFone!Numero;
         End If
         If TabEND.State = 1 Then _
            TabEND.Close
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close
      Else
         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select * from CLIENTE "
         SQL = SQL & " where cgccpf = '" & TabNOTA!prop & "'"
         SQL = SQL & " and status = 'A'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then
            CODG_SUFRAMA_A = "" & Trim(TabCliente.Fields("codg_suframa").Value)
            If Len(CODG_SUFRAMA_A) <= 3 Then _
               CODG_SUFRAMA_A = ""

            'Print #1, Tab(1); "E|" & TabCliente!Nome & "|" & UCase(TabCliente!IE) & "|||";
            Print #1, Tab(1); "E|" & TabCliente!NOME & "|" & UCase(TabCliente!IE) & "|" & Trim(CODG_SUFRAMA_A) & "||";

            If Len(TabCliente!CGCCPF) > 11 Then
               Print #1, Tab(1); "E02|" & TabCliente!CGCCPF;
               Else 'Nesta Caso e CNPJ
                  Print #1, Tab(1); "E03|" & TabCliente!CGCCPF;
            End If

            SP_PROCURA_ENDEREÇONFE TabNOTA!prop, "C", "", 0
            If Not TabEND.EOF Then
               Numero_A = "000"
               If Not IsNull(TabEND.Fields("numero").Value) Then _
                  If Trim(TabEND.Fields("numero").Value) <> "" Then _
                     Numero_A = Trim(TabEND.Fields("numero").Value)

               SP_PROCURA_CEP TabEND!CEP

               If TabCEP.EOF Then _
                  MsgBox "NÃO ENCONTROU CEP"

               SP_PROCURA_FONE TabEND!prop, 0

               If TabFone.EOF Then
                  MsgBox "NÃO ENCONTROU TABFONE"
                  Else: DDD_N = 0 & TabFone!DDD
               End If

               Print #1, Tab(1); _
               "E05|" & _
               Trim(txtEnd.Text) & "|" & _
               Numero_A & "| |" & _
               Trim(txtBairro.Text) & "|" & _
               Trim(txtIBGE.Text) & "|" & _
               Trim(txtCidade.Text) & "|" & _
               Trim(txtUF.Text) & "|" & _
               Trim(txtCep.Text) & _
               "|1058|BRASIL|" & _
               Right(DDD_N, 2) & _
               Trim(txtFone.Text);
            End If
            If TabEND.State = 1 Then _
               TabEND.Close
         End If
         If TabCliente.State = 1 Then _
            TabCliente.Close
   End If

   If TabEND.State = 1 Then _
      TabEND.Close

   'Dados Local da retirada
   SQL = "select * from ENDERECO "
   SQL = SQL & " where prop = '" & Trim(rstEmpresa!CGC) & "'"
   TabEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEND.EOF Then
      SP_PROCURA_CEP TabEND!CEP

      Numero_A = "000"
      If Not IsNull(TabEND.Fields("numero").Value) Then _
         If Trim(TabEND.Fields("numero").Value) <> "" Then _
            Numero_A = Trim(TabEND.Fields("numero").Value)

      'Print #1, Tab(1); "F|" & TABEND!Rua & "|" & "000|  " & "|" & TABEND!Bairro & "|" & TABCEP!codigo_ibge & "|" & TABCEP!Cidade & "|" & TABCEP!UF & "|";
      Print #1, Tab(1); "F|" & TabEND!Rua&; "|" & Numero_A & "|  " & "|" & TabEND!Bairro & "|" & TabCEP!codigo_ibge & "|" & TabCEP!Cidade & "|" & TabCEP!UF & "|";
      Print #1, Tab(1); "F02|" & rstEmpresa!CGC;
      'Print #1, Tab(1); "G|" & TABEND!Rua & "|" & "000|  " & "|" & TABEND!Bairro & "|" & TABCEP!codigo_ibge & "|" & TABCEP!Cidade & "|" & TABCEP!UF & "|";
      Print #1, Tab(1); "G|" & TabEND!Rua & "|" & Numero_A & "|  " & "|" & TabEND!Bairro & "|" & TabCEP!codigo_ibge & "|" & TabCEP!Cidade & "|" & TabCEP!UF & "|";
      Print #1, Tab(1); "G02|" & rstEmpresa!CGC;
   End If
   If TabEND.State = 1 Then _
      TabEND.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_CABEÇALHO_NFe"
End Sub

Private Sub GERAR_PRODUTOS_NFe()
On Error GoTo erro_trata

   Dim Linhas_Impressas_N  As Long
   Dim strCFOP_ITEM        As String
   Dim CODIGO_NCM_A        As String

   NUMR_SEQ_N = 0

   If TabItemNota.State = 1 Then _
      TabItemNota.Close

   SQL = "SELECT NF.PEDIDO_ID, NF.NUMR_REQ, NF.NF_TIPO, NFITEM.*, "
   SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.UNIDADE_MEDIDA, PRODUTO.CODG_NCM, PRODUTO.TIPO_PROD, "
   SQL = SQL & " Produto.Situacao_Tributaria"
   SQL = SQL & " FROM NF "
   SQL = SQL & " INNER JOIN NFITEM "
   SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where NF.nf_id = " & TabNOTA!nf_id
   SQL = SQL & " and NF.empresa_id = " & EMPRESA_ID_N

   TabItemNota.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabItemNota.EOF Then
      TabItemNota.MoveFirst
      While Not TabItemNota.EOF

         strCFOP_ITEM = "" & Trim(TabItemNota!cfop)

         'quando é cupom com nfe
         If Trim(cmbCFOPAux.Text) = 5929 Or Trim(cmbCFOPAux.Text) = 6929 Then
            SQL = "update NFITEM set cfop = " & Trim(cmbCFOPAux.Text)
            SQL = SQL & " where nf_id = " & TabItemNota.Fields("nf_id").Value
            CONECTA_RETAGUARDA.Execute SQL

            strCFOP_ITEM = "" & Trim(cmbCFOPAux.Text)
            Else
               'quando é substituição tributária
               If Trim(TabItemNota!Situacao_Tributaria) = 10 Or _
                  Trim(TabItemNota!Situacao_Tributaria) = 60 Or _
                  Trim(TabItemNota!Situacao_Tributaria) = 70 Then

                  If Trim(TIPO_DOC) <> "DC" And Trim(TIPO_DOC) <> "DV" Then
                     If UNIDADE_FEDERAÇÃO_EMPRESA = Trim(txtUF.Text) Then
                        strCFOP_ITEM = "5405"
                        Else: strCFOP_ITEM = "6404"
                     End If
                  End If

                  Else: strCFOP_ITEM = "" & Trim(cmbCFOPAux.Text) 'segue o cfop informado do combo da tela
               End If
         End If
'=====================================

         frmINTEGRA.INTEGRA_CFOP Int(strCFOP_ITEM)

         NUMR_SEQ_N = NUMR_SEQ_N + 1
         CODIGO_NCM_A = "" & Trim(TabItemNota!codg_ncm)

         If Len(CODIGO_NCM_A) < 8 Then _
            CODIGO_NCM_A = "00"

         intTributacao = BUSCA_TRIBUTAÇÃO_PRODUTO(TabItemNota!Situacao_Tributaria)

         Print #1, Tab(1); "H|" & NUMR_SEQ_N & "|" & Trim(TabItemNota!Descricao);
         Print #1, Tab(1); "I|" & Trim(TabItemNota!Codg_Prod) & "|   |" & Trim(TabItemNota!Descricao) & "|" & Trim(CODIGO_NCM_A) & "||" & strCFOP_ITEM & "|" & TabItemNota!Unidade_Medida & "|" & Format$(TabItemNota!Qtde, strFormatacao4Digitos) & "|" & Format(TabItemNota!Valor, strFormatacao2Digitos) & "|" & Format(TabItemNota!Qtde * TabItemNota!Valor, strFormatacao2Digitos) & "| |" & TabItemNota!Unidade_Medida & "|" & Format$(TabItemNota!Qtde, strFormatacao4Digitos) & "|" & Format(TabItemNota!Valor, strFormatacao2Digitos) & "| | | ||||";

         'Parte de impostos dos produtos
         Print #1, Tab(1); "M";
         Print #1, Tab(1); "N";

         Print #1, Tab(1); "N10d|0|" & intTributacao & "|";

         'Finaliza Itens PIS
         Print #1, Tab(1); "Q";
         Print #1, Tab(1); "Q05|99|0,00";
         Print #1, Tab(1); "Q07|0,00|0,00";

         'Finaliza Itens Cofins
         Print #1, Tab(1); "S";
         Print #1, Tab(1); "S05|99|0,00";
         Print #1, Tab(1); "S07|0,00|0,00";

'=============baixa estoque INICIO
         If TabItemNota.Fields("nf_tipo").Value = "S" Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select status from ITEMREQ "
            SQL = SQL & " where pedido_id = " & TabItemNota.Fields("pedido_id").Value
            SQL = SQL & " and produto_id = " & TabItemNota.Fields("produto_id").Value
            SQL = SQL & " and status <> 'B' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               PRODUTO_ID_N = TabItemNota.Fields("produto_id").Value
               QTDE_PEDIDO = TabItemNota.Fields("qtde").Value
               QTDE_RETIDO = TabItemNota.Fields("qtde").Value

               BAIXA_ESTOQUE_PRODUTO PRODUTO_ID_N, QTDE_PEDIDO, QTDE_RETIDO

               SQL = "update ITEMREQ set "
               SQL = SQL & " status = 'B' "
               SQL = SQL & " where pedido_id = " & TabItemNota.Fields("pedido_id").Value
               SQL = SQL & " and produto_id = " & TabItemNota.Fields("produto_id").Value
               SQL = SQL & " and status <> 'B' "
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
'=============baixa estoque FIM

         'Atualiza Data da Ultima Venda na tabela produto
         SQL = "Update Produto Set "
         SQL = SQL & " Dt_ult_Venda = '" & DMA(Date) & "'"
         SQL = SQL & " Where Empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and Codg_Produto = '" & Trim(TabItemNota!Codg_Prod) & "'"
         CONECTA_RETAGUARDA.Execute SQL

         TabItemNota.MoveNext
      Wend
   End If
   If TabItemNota.State = 1 Then _
      TabItemNota.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_PRODUTOS_NFe"
End Sub

Private Sub GERAR_TOTAIS_NFe()
On Error GoTo erro_trata

   Dim dblVlrBaseICMS As Double
   Dim dblVlrICMS     As Double

   dblVlrBaseICMS = 0
   dblVlrICMS = 0
   VALOR_TOTAL_N = 0
   VALOR_DESCONTO_N = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum((valor*qtde)) from NFITEM "
   SQL = SQL & " where nf_id = " & TabNOTA!nf_id
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_TOTAL_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum(desconto) from NFITEM "
   SQL = SQL & " where nf_id = " & TabNOTA!nf_id
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

'============================================================
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

VALOR_DESCONTO_N = 0

   'desconto individual por item
   SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from ITEMREQ "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select valor_desconto from CABECAREQ "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
'============================================================

   Print #1, Tab(1); "W";

'============================================================
   If booOptanteSimples = True Then
      Print #1, Tab(1); "W";

      'W02|vBC|vICMS|vBCST|vST|vProd                                                                  |vFrete                                              |vSeg|vDesc                             |vII|vIPI|vPIS|vCOFINS|vOutro|vNF|
'eu
'Print #1, Tab(1); "W02|0,00|0,00|0,00|0,00|" & Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos) & "|0,00|0,00|" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos) & "|0,00|0,00|0,00|0,00|0,00|" & Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos) & "|0,00";

      If Not IsNumeric(txtFrete.Text) Then
         If txtBaseIcmsSub.Text > 0 Then
            Print #1, Tab(1); "W02|0,00|0,00|" & txtBaseIcmsSub.Text & "|" & txtVlrIcmsSub.Text & "|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|" & Format(txtFrete.Text, strFormatacao2Digitos) & "|0,00|" & Format(txtDesconto.Text, strFormatacao2Digitos) & "|0,00|0,00|0,00|0,00|" & Format(VLR_OUTROS_DEV, strFormatacao2Digitos) & "|" & Format(VALOR_TOTAL_N + VLR_ICMS_SUB_DEV + VLR_FRETE_DEV + VLR_OUTROS_DEV + VLR_IPI_DEV, strFormatacao2Digitos) & "|0,00";
            Else: Print #1, Tab(1); "W02|0,00|0,00|0,00|0,00|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|" & Format(txtFrete.Text, strFormatacao2Digitos) & "|0,00|" & Format(txtDesconto.Text, strFormatacao2Digitos) & "|0,00|0,00|0,00|0,00|" & Format(VLR_OUTROS_DEV, strFormatacao2Digitos) & "|" & Format(VALOR_TOTAL_N + VLR_ICMS_SUB_DEV + VLR_FRETE_DEV + VLR_OUTROS_DEV + VLR_IPI_DEV, strFormatacao2Digitos) & "|0,00";
         End If
         Else
            If txtBaseIcmsSub.Text > 0 Then
               Print #1, Tab(1); "W02|0,00|0,00|" & txtBaseIcmsSub.Text & "|" & txtVlrIcmsSub.Text & "|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|" & Format(txtFrete.Text, strFormatacao2Digitos) & "|0,00|" & Format(txtDesconto.Text, strFormatacao2Digitos) & "|0,00|0,00|0,00|0,00|" & Format(VLR_OUTROS_DEV, strFormatacao2Digitos) & "|" & Format(VALOR_TOTAL_N + VLR_ICMS_SUB_DEV + VLR_FRETE_DEV + VLR_OUTROS_DEV + VLR_IPI_DEV, strFormatacao2Digitos) & "|0,00";
               Else: Print #1, Tab(1); "W02|0,00|0,00|0,00|0,00|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|" & Format(txtFrete.Text, strFormatacao2Digitos) & "|0,00|" & Format(txtDesconto.Text, strFormatacao2Digitos) & "|0,00|0,00|0,00|0,00|" & Format(VLR_OUTROS_DEV, strFormatacao2Digitos) & "|" & Format(VALOR_TOTAL_N + VLR_ICMS_SUB_DEV + VLR_FRETE_DEV + VLR_OUTROS_DEV + VLR_IPI_DEV, strFormatacao2Digitos) & "|0,00";
            End If
      End If
      Else 'Nao e Optante pelo Simples
         'W02|vBC                        |vICMS                    |                      vBCST|                       vST|                                               vProd|                                            vFrete|vSeg|                                                  vDesc| vII|                    vIPI|vPIS|vCOFINS|vOutro|                                                          vNF|

         If Not IsNumeric(txtFrete.Text) Then
            Print #1, Tab(1); "W02|" & txtBaseCalculo.Text & "|" & txtValorICMS.Text & "|0,00|" & txtVlrIcmsSub.Text & "|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|0,00|0,00|" & Format(txtDesconto.Text, strFormatacao2Digitos) & "|0,00|" & txtvaloripi.Text & "|0,00|0,00|0,00|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|0,00";
            Else: Print #1, Tab(1); "W02|" & txtBaseCalculo.Text & "|" & txtValorICMS.Text & "|" & txtBaseIcmsSub.Text & "|" & txtVlrIcmsSub.Text & "|" & Format(VALOR_TOTAL_N, strFormatacao2Digitos) & "|" & Format(txtFrete.Text, strFormatacao2Digitos) & "|0,00|" & Format(txtDesconto.Text, strFormatacao2Digitos) & "|0,00|" & txtvaloripi.Text & "|0,00|0,00|0,00|" & Format(VALOR_TOTAL_N + txtFrete.Text, strFormatacao2Digitos) & "|0,00";
         End If
   End If

'============================================================

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_TOTAIS_NFe"
End Sub

Private Sub GERAR_TRANSPORTADORA_NFe()
On Error GoTo erro_trata

   Dim strMarcarPagamentoFrete As String
       
   If txtCNPJCPF_TRANSP.Text <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TRANSPORTADORA "
      SQL = SQL & " where cgccpf = '" & txtCNPJCPF_TRANSP.Text & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            SP_PROCURA_ENDEREÇONFE TabTemp!CGCCPF, "C", "", 0

            If OptFreteEmitente.Value = True Then _
               strMarcarPagamentoFrete = "1"

            If OptFreteDestinatario.Value = True Then _
               strMarcarPagamentoFrete = "2"

            Print #1, Tab(1); "X|1";
            If Not TabEND.EOF Then
               Print #1, Tab(1); "X03|" & Trim(TabTemp!NOME) & "|" & TabTemp!IE & "|" & TabEND!Rua & "|" & TabEND!UF & "|" & TabEND!Cidade;
               Else
               Print #1, Tab(1); "X03|" & Trim(TabTemp!NOME) & "|" & "ISENTO" & "|" & "GOIANIA" & "|" & "GO" & "|" & "GOIANIA";
            End If
            Print #1, Tab(1); "X04|" & Trim(TabTemp!CGCCPF);
            
            'ESTAVA INVERTIDO PESO LIQUIDO COM BRUTO
            'Print #1, Tab(1); "X26|" & strMarcarPagamentoFrete & "|" & TxtEspecie.Text & "|PROPRIO|0,000|" & Format(TxtPesoBruto.Text, strFormatacao3Digitos) & "|" & Format(TxtPesoLiquido.Text, strFormatacao3Digitos)
            Print #1, Tab(1); "X26|" & strMarcarPagamentoFrete & "|" & TxtEspecie.Text & "|PROPRIO|0,000|" & Format(TxtPesoLiquido.Text, strFormatacao3Digitos) & "|" & Format(TxtPesoBruto.Text, strFormatacao3Digitos)
         End If
         Else
            MsgBox "Transportadora inexistente, impossível continuar !!!"
            Unload Me
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      If TabEND.State = 1 Then _
         TabEND.Close
      If TabCEP.State = 1 Then _
         TabCEP.Close
   End If

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_TRANSPORTADORA_NFe"
End Sub

Private Sub GERAR_FATURAS_NFe()
On Error GoTo erro_trata

   Print #1, Tab(1); "Y";
   Print #1, Tab(1); "Y02|" & NUMR_REQ_N & "|" & Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos) & "|" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos) & "|" & Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where l.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and l.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and l.lancamento_id = i.lancamento_id "
   SQL = SQL & " and l.tipo_lancamento = 1 "
   SQL = SQL & " and forma_id > 1 "
   SQL = SQL & " order by i.seq"
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLancamento.EOF
      Print #1, Tab(1); "Y07|" & txtNota.Text & "|" & Mid(TabLancamento!dt_Vencimento, 7, 4) & "-" & Mid(TabLancamento!dt_Vencimento, 4, 2) & "-" & Mid(TabLancamento!dt_Vencimento, 1, 2) & "|" & Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
      TabLancamento.MoveNext
   Wend
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_FATURAS_NFe"
End Sub

Private Sub GERAR_RODAPE_NFe()
On Error GoTo erro_trata

   Dim TABTIPOVENDA As New ADODB.Recordset
   Dim descricao_pgto As String

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select v.nome_vend,c.tipovenda_id from VENDEDOR v, CABECAREQ c "
   SQL = SQL & " where v.vendedor_id = c.vendedor_id "
   SQL = SQL & " and c.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and c.numr_req = " & NUMR_REQ_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         NOME_VEND_A = TabConsulta.Fields(0).Value

   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   
   'Condicoes de Pgto
   SQL = "select tipo_pagto from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & 1 'RECEBER
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      If Not IsNull(TabConsulta!tipo_pagto) Then
         If IsNumeric(TabConsulta!tipo_pagto) Then
            If TABTIPOVENDA.State = 1 Then _
               TABTIPOVENDA.Close

            SqL2 = "select descricao from TIPOVENDA "
            SqL2 = SqL2 & " where tipovenda_id = " & TabConsulta!tipo_pagto
            TABTIPOVENDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TABTIPOVENDA.EOF Then _
               If Not IsNull(TABTIPOVENDA.Fields(0).Value) Then _
                  descricao_pgto = TABTIPOVENDA.Fields(0).Value

            If TABTIPOVENDA.State = 1 Then _
               TABTIPOVENDA.Close
         End If
      End If
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'PEGA_DADOS_EMPRESA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      NF_DEV_ENTRADA = TabTemp!numr_Nota

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   'PEGA_DADOS_EMPRESA

   Print #1, Tab(1); "Z|" & Trim(txtMSG.Text) & ", Numero Pedido: " & NUMR_REQ_N & ", Vendedor: " & Trim(NOME_VEND_A) & ", " & Trim(txtDadosAdicionais.Text);
   If TIPO_DOC = "DC" Then
      Print #1, Tab(1); "Z04|" & "DV ENT REF. " & NF_DEV_ENTRADA & "|0";
      Else
         If TIPO_DOC = "DV" Then
            Print #1, Tab(1); "Z04|" & "DV SAI REF. " & 0 & "|0";
            Else: Print #1, Tab(1); "Z04|" & "VOLTE SEMPRE" & "|0";
         End If
   End If

   Print #1, Tab(1); "Z10|" & "PROCESSO INTERNO" & "|1";

Exit Sub
erro_trata:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_RODAPE_NFe"
End Sub

Private Sub AtualizaTotalNota()

   VLR_ICMS_SUB_DEV = 0 & txtVlrIcmsSub.Text

   If VLR_FRETE_DEV > 0 Then
      VLR_FRETE_DEV = 0 & txtFrete.Text
      Else: txtFrete.Text = 0
   End If

   If VLR_OUTROS_DEV > 0 Then
      VLR_OUTROS_DEV = 0 & txtValorOutros.Text
      Else: txtValorOutros.Text = 0
   End If

   If VLR_IPI_DEV > 0 Then
      VLR_IPI_DEV = 0 & txtvaloripi.Text
      Else: txtvaloripi.Text = 0
   End If

   VALOR_ITEM_N = 0 & txtValorProdutos.Text
   txtValorTotalNota.Text = 0 & Format(VLR_ICMS_SUB_DEV + VLR_FRETE_DEV + VLR_OUTROS_DEV + VLR_IPI_DEV + VALOR_ITEM_N, strFormatacao2Digitos)

End Sub
