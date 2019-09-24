VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProdutoConsulta 
   Caption         =   "Consulta Produtos"
   ClientHeight    =   7230
   ClientLeft      =   1185
   ClientTop       =   2250
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PRODUTOCONSULTA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   13575
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkProdBalanca 
      Caption         =   "Produto Balança"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   29
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtLocacao 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      MaxLength       =   40
      TabIndex        =   27
      Top             =   1800
      Width           =   1575
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
      Left            =   7800
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbEstab 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      TabIndex        =   24
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtValor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12120
      MaxLength       =   30
      TabIndex        =   22
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox cmbTamanhoAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10080
      TabIndex        =   21
      ToolTipText     =   "Aliquota de ICMS do produto."
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbTamanho 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10080
      TabIndex        =   19
      ToolTipText     =   "Aliquota de ICMS do produto."
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtBarra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9480
      TabIndex        =   17
      ToolTipText     =   "Digite Codigo de Barra do Produto"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox cmbAuxFamilia 
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
      Left            =   4680
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbFamilia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   14
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton optMP 
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12840
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.OptionButton optPA 
      Caption         =   "PA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12840
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtREF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4665
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox cmbAuxFORNEC 
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
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.ComboBox cmbFORNEC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      MaxLength       =   40
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   1270
      ButtonWidth     =   3572
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rel.Preço Custo"
            Key             =   "custo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rel.Preço Venda"
            Key             =   "venda"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rel.Custo/Venda"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkCancelados 
         Caption         =   "Imprimir Cancelados?"
         Height          =   240
         Left            =   10440
         TabIndex        =   35
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   10440
         TabIndex        =   18
         Top             =   0
         Width           =   2655
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9240
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":941B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":AC2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":C007
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":D317
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOCONSULTA.frx":E422
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstProduto 
      Height          =   4500
      Left            =   45
      TabIndex        =   8
      Top             =   2280
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   7938
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   16777215
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Estab"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qtde."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Pr.Venda"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pr.Atacado"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Pr.Custo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fornecedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "+ Est."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "- Est."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Referência"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Grupo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ST"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Codg.Barras"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "NCM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Locação"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "UN"
         Object.Width           =   2540
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
      DesignWidth     =   13575
      DesignHeight    =   7230
   End
   Begin VB.Label lblMP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   225
      Left            =   10200
      TabIndex        =   34
      Top             =   6960
      Width           =   105
   End
   Begin VB.Label lblPA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   225
      Left            =   7800
      TabIndex        =   33
      Top             =   6960
      Width           =   105
   End
   Begin VB.Label lblProdRevenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   225
      Left            =   5160
      TabIndex        =   32
      Top             =   6960
      Width           =   105
   End
   Begin VB.Label lblProdBalanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   225
      Left            =   2280
      TabIndex        =   31
      Top             =   6960
      Width           =   105
   End
   Begin VB.Label lblQtde 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   225
      Left            =   240
      TabIndex        =   30
      Top             =   6960
      Width           =   105
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Locação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7800
      TabIndex        =   28
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label15 
      Caption         =   "Estabelecimento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7800
      TabIndex        =   25
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Venda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12120
      TabIndex        =   23
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Medida Peça:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10245
      TabIndex        =   20
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Label5 
      Caption         =   "Codg.Barras:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9480
      TabIndex        =   16
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Referência:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Fornecedor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "Família de Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   6
      Top             =   1560
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Código Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1740
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "frmProdutoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

'======
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
   cmbEstab.Text = "" & TRAZ_ESTABELECIMENTO(cmbEstabAUX.Text)

   lblQtde.Caption = ""
   lblProdBalanca.Caption = ""
   lblProdRevenda.Caption = ""
   lblPA.Caption = ""
   lblMP.Caption = ""

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Activate()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 1
   lstProduto.ListItems.Clear

   MOSTRA_RODAPE "ESC - SAIR", "CLICK DUAS VEZES PARA SELECIONAR PRODUTO", "", "", ""

   cmbFamilia.Clear
   cmbAuxFamilia.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFamilia.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("familiaproduto_id").Value)
      cmbAuxFamilia.AddItem Trim(TabDESCR.Fields("familiaproduto_id").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close
   
   cmbFORNEC.Clear
   cmbAuxFORNEC.Clear

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select CNPJCPF, DESCRICAO from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " order by descricao "
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabFornecedor.EOF
      If Not IsNull(TabFornecedor!CNPJCPF) Then
        cmbFORNEC.AddItem Trim(TabFornecedor!DESCRICAO) & "-" & Trim(TabFornecedor!CNPJCPF)
        cmbAuxFORNEC.AddItem TabFornecedor!CNPJCPF
      End If
      TabFornecedor.MoveNext
   Wend
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   cmbTamanho.Clear
   cmbTamanhoAUX.Clear
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'N'"
   SQL = SQL & " order by codigo "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbTamanho.AddItem TabTemp.Fields("codigo").Value & " - " & TabTemp.Fields("DESCRICAO").Value
      cmbTamanhoAUX.AddItem TabTemp.Fields("codigo").Value

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TIPO_USUARIO < 4 Or TIPO_USUARIO > 5 Then
      cmbEstab.Visible = False
      Label15.Visible = False
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         lblProdBalanca.Caption = ""
         lblProdRevenda.Caption = ""
         lblPA.Caption = ""
         lblMP.Caption = ""
         lblQtde.Caption = ""
         lstProduto.Visible = True
         txtItem.Text = ""
         txtDesc.Text = ""
         txtDesc.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
   FORMULA_REL = ""
   MOSTRA_RODAPE "", "", "", "", ""
      
Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub cmbTAMANHO_Click()
'On Error GoTo ERRO_TRATA

   cmbTamanhoAUX.ListIndex = cmbTamanho.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTAMANHO_Click"
End Sub

Private Sub lstProduto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   OrdenaListView lstProduto, ColumnHeader
End Sub

Private Sub lstProduto_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstProduto.SelectedItem.Text) Then
      SQL3 = lstProduto.SelectedItem.Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_DblClick"
End Sub

Private Sub cmbFORNEC_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxFORNEC.ListIndex = cmbFORNEC.ListIndex
   If cmbAuxFORNEC.Text <> "" Then
      Busca_fornec

      'CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFORNEC_Click"
End Sub

Private Sub lstProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then _
      lstProduto_DblClick

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_KeyPress"
End Sub

Private Sub optPA_Click()
'On Error GoTo ERRO_TRATA

   'CONSULTA_TUDO

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "optPA_Click"
End Sub

Private Sub optmp_Click()
'On Error GoTo ERRO_TRATA

   'CONSULTA_TUDO

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "optmp_Click"
End Sub

Private Sub txtDesc_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDesc.SelStart = 0
   txtDesc.SelLength = Len(txtDesc.Text)

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_GotFocus"
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      UCase (txtDesc.Text)
      KeyAscii = 0

      CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtRef_KeyPress"
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      UCase (txtDesc.Text)
      KeyAscii = 0
      NUMR_SEQ_N = 1

      CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_KeyPress"
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtItem_KeyPress"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         CONSULTA_TUDO
      Case "venda"
         MONTA_FORMULA_REL

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_prod_venda.rpt"
         frmRELATORIO10.Show 1
      Case "custo"
         MONTA_FORMULA_REL

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_prod_custo.rpt"
         frmRELATORIO10.Show 1
      Case "print"
         MONTA_FORMULA_REL

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Estoque.rpt"
         frmRELATORIO10.Show 1
      Case "limpar"
         chkCancelados.Value = 0
         lblProdBalanca.Caption = ""
         lblProdRevenda.Caption = ""
         lblPA.Caption = ""
         lblMP.Caption = ""
         lblQtde.Caption = ""
         lstProduto.Visible = True
         cmbTamanhoAUX.Text = ""
         cmbTamanho.Text = ""
         cmbAuxFORNEC.Text = ""
         cmbFORNEC.Text = ""
         lstProduto.ListItems.Clear
         txtItem.Text = ""
         txtDesc.Text = ""
         cmbAuxFamilia.Text = ""
         txtBarra.Text = ""
         cmbFamilia.Text = ""
         optMP.Value = False
         optPA.Value = False
         NUMR_SEQ_N = 1
         lstProduto.ListItems.Clear
         txtRef.Text = ""
         txtDesc.SetFocus
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxFamilia.ListIndex = cmbFamilia.ListIndex

   'CONSULTA_TUDO

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub cmbestab_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "cmbestab_Click"
End Sub

Private Sub MONTA_FORMULA_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{PRODUTO.empresa_id} = " & EMPRESA_ID_N

   If Trim(cmbEstabAUX.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {ESTOQUE.ESTABELECIMENTO_id} = " & cmbEstabAUX.Text

   FORMULA_REL = FORMULA_REL & " and {PRODUTO.produto_ID} = {ESTOQUE.produto_ID}"

   If optPA.Value = True Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.tipo_prod} = 1"
   If optMP.Value = True Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.tipo_prod} = 0"

   If chkCancelados.Value <> 1 Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.situacao} = 'A' "

   If chkProdBalanca.Value = 1 Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.produto_balanca} = 1"

   If Trim(txtItem.Text) <> "" Then
      SqL2 = Chr$(39) & Trim(txtItem.Text) & "%" & Chr(39)
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.CODG_PRODUTO} = '" & Trim(txtItem.Text) & "'"
      Else
         If txtDesc.Text <> "" Then
            SqL2 = Chr$(39) & Trim(txtDesc.Text) & "%" & Chr(39)
            FORMULA_REL = FORMULA_REL & " and {PRODUTO.descricao} like " & SqL2
         End If
   End If
   If Trim(cmbAuxFamilia.Text) <> "" Then
      If FORMULA_REL = "" Then
         FORMULA_REL = "{PRODUTO.familiaproduto_id} = " & Trim(cmbAuxFamilia.Text)
         Else: FORMULA_REL = FORMULA_REL & " and {PRODUTO.familiaproduto_id} = " & Trim(cmbAuxFamilia.Text)
      End If
   End If
   If Trim(txtRef.Text) <> "" Then
      SqL2 = Chr$(39) & txtRef.Text & Chr(39)
      SqL2 = Chr$(39) & txtRef.Text & "%" & Chr(39)
      If FORMULA_REL = "" Then
         FORMULA_REL = "{PRODUTO.referencia} >= " & SqL2
         Else: FORMULA_REL = FORMULA_REL & " and {PRODUTO.referencia} like " & SqL2
      End If
   End If
   If Trim(cmbAuxFORNEC.Text) <> "" Then
      If FORMULA_REL = "" Then
         FORMULA_REL = "{PRODUTO.fornecedor_id} = " & FORNEC_ID_N
         Else: FORMULA_REL = FORMULA_REL & " and {PRODUTO.fornecedor_id} = " & FORNEC_ID_N
      End If
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_formula_rel_REL"
End Sub

Private Sub Busca_fornec()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select fornecedor_id from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(cmbAuxFORNEC.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      FORNEC_ID_N = TabTemp!FORNECEDOR_ID

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Busca_fornec"
End Sub

Sub CONSULTA_TUDO()
   CONT_N = 0
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select count(PRODUTO.produto_id) from PRODUTO WITH (NOLOCK)"

   SQL = SQL & " INNER JOIN ESTOQUE WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = ESTOQUE.PRODUTO_ID"

   SQL = SQL & " where situacao <> 'C' "

   If Trim(cmbEstabAUX.Text) <> "" Then _
      SQL = SQL & " and ESTABELECIMENTO_id = " & cmbEstabAUX.Text

   If Trim(txtDesc.Text) <> "" Then _
      SQL = SQL & " and descricao like '" & UCase(Trim(txtDesc.Text)) & "%" & "'"

   If optPA.Value = True Then _
      SQL = SQL & " and tipo_prod = 1"

   If optMP.Value = True Then _
      SQL = SQL & " and tipo_prod = 0"

   If Trim(txtValor.Text) <> "" Then _
      If IsNumeric(txtValor.Text) Then _
         SQL = SQL & " and preco_venda = " & tpMOEDA(txtValor.Text)

   If Trim(txtRef.Text) <> "" Then _
      SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

   If cmbAuxFORNEC.Text <> "" Then _
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   If Trim(txtItem.Text) <> "" Then _
      SQL = SQL & " and codg_produto like '" & UCase(Trim(txtItem.Text)) & "%'"

   If Trim(cmbAuxFamilia.Text) <> "" Then _
      SQL = SQL & " and familiaproduto_id = " & Trim(cmbAuxFamilia.Text)

   If Trim(txtBarra.Text) <> "" Then _
      SQL = SQL & " and codg_barra like '" & Trim(txtBarra.Text) & "%" & "'"

   If Trim(cmbTamanhoAUX.Text) <> "" Then _
      SQL = SQL & " and tamanho = " & Trim(cmbTamanhoAUX.Text)

   If Trim(txtLocacao.Text) <> "" Then _
      SQL = SQL & " and locacao like '" & Trim(txtLocacao.Text) & "%" & "'"

   If chkProdBalanca.Value = 1 Then _
      SQL = SQL & " and produto_balanca = 1"

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then _
      If Not IsNull(TabProduto.Fields(0).Value) Then _
         CONT_N = 0 & TabProduto.Fields(0).Value
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If CONT_N > 500 Then
      Msg = "Esta operação irá processar todos produtos cadastrado, deseja continuar ? " & CONT_N & " registros"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      Msg = ""
      If RESPOSTA = vbNo Then _
         Exit Sub
   End If

   SQL = "select * from PRODUTO WITH (NOLOCK)"

   SQL = SQL & " INNER JOIN ESTOQUE WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = ESTOQUE.PRODUTO_ID"
   
   SQL = SQL & " where situacao <> 'C' "

   If Trim(cmbEstabAUX.Text) <> "" Then _
      SQL = SQL & " and ESTABELECIMENTO_id = " & cmbEstabAUX.Text

   If Trim(txtDesc.Text) <> "" Then _
      SQL = SQL & " and descricao like '" & UCase(Trim(txtDesc.Text)) & "%" & "'"

   If optPA.Value = True Then _
      SQL = SQL & " and tipo_prod = 1"

   If optMP.Value = True Then _
      SQL = SQL & " and tipo_prod = 0"

   If Trim(txtValor.Text) <> "" Then _
      If IsNumeric(txtValor.Text) Then _
         SQL = SQL & " and preco_venda = " & tpMOEDA(txtValor.Text)

   If Trim(txtRef.Text) <> "" Then _
      SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

   If cmbAuxFORNEC.Text <> "" Then _
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   If Trim(txtItem.Text) <> "" Then _
      SQL = SQL & " and codg_produto like '" & UCase(Trim(txtItem.Text)) & "%'"

   If Trim(cmbAuxFamilia.Text) <> "" Then _
      SQL = SQL & " and familiaproduto_id = " & Trim(cmbAuxFamilia.Text)

   If Trim(txtBarra.Text) <> "" Then _
      SQL = SQL & " and codg_barra like '" & Trim(txtBarra.Text) & "%" & "'"

   If Trim(cmbTamanhoAUX.Text) <> "" Then _
      SQL = SQL & " and tamanho = " & Trim(cmbTamanhoAUX.Text)

   If Trim(txtLocacao.Text) <> "" Then _
      SQL = SQL & " and locacao like '" & Trim(txtLocacao.Text) & "%" & "'"

   SQL = SQL & " order by descricao"

   SETA_GRID
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      MsgBox "Nenhum registro encontrado", vbExclamation
      txtDesc.SetFocus
      Exit Sub
      Else: TabProduto.MoveFirst
   End If
   Me.Enabled = False

   Dim dblContador   As Double
   Dim VALOR_CUSTO_N As Double
   Dim Qtde_Revenda  As Long
   Dim Qtde_PA       As Long
   Dim Qtde_MP       As Long
   Dim QTDE_BALANCA  As Long
   Dim PRECO_VENDA_N As Double
   Dim TAB_PRECO_ID_N   As Integer

   TAB_PRECO_ID_N = 1
   lstProduto.Visible = False
   lstProduto.ListItems.Clear
   dblContador = 0
   Qtde_Revenda = 0
   Qtde_PA = 0
   Qtde_MP = 0
   QTDE_BALANCA = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select tabelapreco_id from TABELAPRECO WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TAB_PRECO_ID_N = 0 & TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   While Not TabProduto.EOF
      DoEvents
      dblContador = dblContador + 1
      lblQtde.Caption = "QtdeProdutos = " & dblContador

      Me.Caption = "Aguarde, Processando ...  "

      Set item = lstProduto.ListItems.Add(, "seq." & dblContador, Trim(TabProduto.Fields("codg_produto").Value))

      SqL2 = ""
      If Trim(TabProduto.Fields("referencia").Value) <> "" Then _
         SqL2 = "  |  " & Trim(TabProduto.Fields("referencia").Value)

      item.SubItems(1) = "" & Trim(TabProduto!DESCRICAO) & SqL2
      item.SubItems(2) = "" & TRAZ_ESTABELECIMENTO(TabProduto.Fields("estabelecimento_id").Value)
      item.SubItems(3) = "" & Format(TRAZ_QTDE_ESTOQUE(TabProduto.Fields("estabelecimento_id").Value, TabProduto.Fields("produto_id").Value), strFormatacao3Digitos)

      item.SubItems(4) = "" & Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
      item.SubItems(5) = "" & Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)

PRECO_VENDA_N = 0 & (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(TabProduto.Fields("produto_id").Value, TAB_PRECO_ID_N, 1))
If PRECO_VENDA_N > 0 Then _
   item.SubItems(4) = "" & Format(PRECO_VENDA_N, strFormatacao2Digitos)

      VALOR_CUSTO_N = 0 & TabProduto!PRECO_CUSTO

      item.SubItems(6) = "" & Format(0, strFormatacao2Digitos)
      If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
         item.SubItems(6) = "" & Format(VALOR_CUSTO_N, strFormatacao2Digitos)

      If Not IsNull(TabProduto.Fields("produto_id").Value) Then
         If TabProduto.Fields("produto_id").Value > 0 Then
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

            If Not IsNull(TabProduto.Fields("fornecedor_id").Value) Then
               SQL = "select Descricao from vwFornecedor WITH (NOLOCK)"
               SQL = SQL & " where fornecedor_id = " & TabProduto.Fields("fornecedor_id").Value
               TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabFornecedor.EOF Then _
                  item.SubItems(7) = "" & Trim(TabFornecedor.Fields(0).Value)
            End If
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

         End If
      End If

      item.SubItems(8) = "" & Format(TabProduto!Qtd_minimo, strFormatacao3Digitos)
      item.SubItems(9) = "" & Format(TabProduto!qtd_maximo, strFormatacao3Digitos)
      item.SubItems(10) = "" & Trim(TabProduto!REFERENCIA)
      item.SubItems(11) = "SEM GRUPO"

      NUMR_ID_N = 0 & TabProduto!FAMILIAPRODUTO_ID

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where familiaproduto_id = " & NUMR_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp!DESCRICAO) Then _
            item.SubItems(11) = TabTemp!DESCRICAO
      If TabTemp.State = 1 Then _
         TabTemp.Close

      item.SubItems(12) = "" & TabProduto.Fields("situacao_tributaria").Value
      item.SubItems(13) = "" & TabProduto.Fields("codg_barra").Value
      item.SubItems(14) = "" & TabProduto.Fields("codg_NCM").Value
      item.SubItems(15) = "" & TabProduto.Fields("locacao").Value
      item.SubItems(16) = "" & TabProduto.Fields("UNIDADE_MEDIDA").Value

      If TabProduto.Fields("situacao").Value = "A" Then
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
         item.ListSubItems(10).ForeColor = vbBlue
         item.ListSubItems(11).ForeColor = vbBlue
         item.ListSubItems(12).ForeColor = vbBlue
         item.ListSubItems(13).ForeColor = vbBlue
      End If
      If TabProduto.Fields("situacao").Value = "P" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(3).ForeColor = vbRed
         item.ListSubItems(4).ForeColor = vbRed
         item.ListSubItems(5).ForeColor = vbRed
         item.ListSubItems(6).ForeColor = vbRed
         item.ListSubItems(7).ForeColor = vbRed
         item.ListSubItems(8).ForeColor = vbRed
         item.ListSubItems(9).ForeColor = vbRed
         item.ListSubItems(10).ForeColor = vbRed
         item.ListSubItems(11).ForeColor = vbRed
         item.ListSubItems(12).ForeColor = vbRed
         item.ListSubItems(13).ForeColor = vbRed
      End If

      If TabProduto.Fields("produto_balanca").Value = True Then
         QTDE_BALANCA = QTDE_BALANCA + 1
         Else: Qtde_Revenda = Qtde_Revenda + 1
      End If
      If TabProduto.Fields("tipo_prod").Value = 1 Then
         Qtde_PA = Qtde_PA + 1
         Else
            Qtde_MP = Qtde_MP + 1
            item.ForeColor = &H404080
            item.ListSubItems(1).ForeColor = &H404080
            item.ListSubItems(2).ForeColor = &H404080
            item.ListSubItems(3).ForeColor = &H404080
            item.ListSubItems(4).ForeColor = &H404080
            item.ListSubItems(5).ForeColor = &H404080
            item.ListSubItems(6).ForeColor = &H404080
            item.ListSubItems(7).ForeColor = &H404080
            item.ListSubItems(8).ForeColor = &H404080
            item.ListSubItems(9).ForeColor = &H404080
            item.ListSubItems(10).ForeColor = &H404080
            item.ListSubItems(11).ForeColor = &H404080
            item.ListSubItems(12).ForeColor = &H404080
            item.ListSubItems(13).ForeColor = &H404080
      End If
      lblMP.Caption = "Mp = " & Qtde_MP
      lblPA.Caption = "PA = " & Qtde_PA
      lblProdRevenda.Caption = "Produto Revenda = " & Qtde_Revenda
      lblProdBalanca.Caption = "Produto Produção = " & QTDE_BALANCA
'========================
      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   Me.Enabled = True
   lstProduto.Visible = True

   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Public Sub ABRE_DATA_BASE(DATA_BASE As String, SERVIDOR_DEPOSITO As String)
'On Error GoTo ERRO_TRATA

   If Not FSO.FileExists(App.Path & DATA_BASE) Then _
      Exit Sub

   Dim ConnectString       As String

   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

   SENHA_ADM_SQLSERVER = "ejsnenas"
   USUARIO_ADM_SQLSERVER = "sa"

   ConnectString = "uid=" & USUARIO_ADM_SQLSERVER & _
                   ";pwd=" & SENHA_ADM_SQLSERVER & _
                   ";Provider=SQLOLEDB.1;Server=" & SERVIDOR_DEPOSITO & _
                   ";database=" & DATA_BASE & _
                   ";dsn='" & SERVIDOR_DEPOSITO & _
                   "';connection=adConnectAsync"

   With CONECTA_AUXILIAR
      .ConnectionString = ConnectString
      .ConnectionTimeout = 10
      .Open
   End With


Exit Sub
ERRO_TRATA:
   MsgBox "não conectou no deposito"
   TRATA_ERROS Err.Description, Me.Name, "ABRE_DATA_BASE"
End Sub

Function LE_REFERENCIA(REF_A As String) As Long
'On Error GoTo ERRO_TRATA

   LE_REFERENCIA = 0
   If Trim(REF_A) = "" Then _
      Exit Function

   Dim TabProdFornec As New ADODB.Recordset
   PRODUTO_ID_N = 0

   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

   SQL = "select PRODUTOFORNECEDOR.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
   SQL = SQL & " from PRODUTOFORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTOFORNECEDOR.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where codg_prod_fornec = '" & Trim(REF_A) & "'"
   SQL = SQL & " and PRODUTOFORNECEDOR.fornecedor_id = " & FORNEC_ID_N
   TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProdFornec.EOF Then
      txtProduto.Text = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
      PRODUTO_ID_N = 0 & Trim(TabProdFornec.Fields("produto_id").Value)
      Else
         If TabProdFornec.State = 1 Then _
            TabProdFornec.Close

         SQL = "select codg_produto,descricao,produto_id,preco_custo,codg_ncm,unidade_medida,aliquota_icms,situacao_tributaria "
         SQL = SQL & " from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where referencia = '" & Trim(REF_A) & "'"
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProdFornec.EOF Then
            txtProduto.Text = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
            PRODUTO_ID_N = 0 & Trim(TabProdFornec.Fields("produto_id").Value)
         End If
   End If
   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_REFERENCIA"
End Function

Function TRAZ_ESTOQUE_TRANSITO(PROD_ID_N As Long, Estab_ID_N As Integer) As Double
'On Error GoTo ERRO_TRATA

   TRAZ_ESTOQUE_TRANSITO = 0
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum(ESTOQUETRANSF.QTDE_TRANSF) "

   SQL = SQL & " from ESTOQUETRANSF WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUETRANSF.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where estab_destino_id = " & Estab_ID_N
   SQL = SQL & " and ESTOQUETRANSF.situacao = 'T' "

   SQL = SQL & " and ESTOQUETRANSF.produto_id = " & PROD_ID_N

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TRAZ_ESTOQUE_TRANSITO = 0 & TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRAZ_ESTOQUE_TRANSITO"
End Function
