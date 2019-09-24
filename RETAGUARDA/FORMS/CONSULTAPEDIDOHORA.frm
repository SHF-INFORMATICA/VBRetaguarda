VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCONSULTAPEDIDOHORA 
   Caption         =   "Consulta Pedido Venda Por Hora"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
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
   Icon            =   "CONSULTAPEDIDOHORA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTicketMedio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   6855
      TabIndex        =   41
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame fraFiltro 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   30
      TabIndex        =   16
      Top             =   1320
      Width           =   11775
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   7695
         Picture         =   "CONSULTAPEDIDOHORA.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1320
         Width           =   405
      End
      Begin VB.ComboBox cmbEstab 
         Height          =   360
         Left            =   1455
         TabIndex        =   30
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cmbFamilia 
         Height          =   360
         Left            =   9015
         TabIndex        =   29
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3330
         Picture         =   "CONSULTAPEDIDOHORA.frx":6614
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   840
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
         Left            =   1455
         TabIndex        =   27
         Top             =   840
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
         Left            =   3795
         MaxLength       =   100
         TabIndex        =   26
         Top             =   840
         Width           =   3825
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
         Left            =   4335
         TabIndex        =   25
         Top             =   360
         Width           =   1815
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
         Left            =   8175
         MaxLength       =   100
         TabIndex        =   24
         Top             =   1320
         Width           =   3495
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
         Left            =   8175
         TabIndex        =   23
         Top             =   360
         Width           =   3495
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
         Left            =   1455
         TabIndex        =   22
         Top             =   360
         Width           =   1785
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
         Left            =   8175
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   735
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
         Left            =   1455
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   735
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
         Left            =   9015
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   975
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
         Left            =   4335
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   735
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
         Left            =   1455
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   870
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   5535
         TabIndex        =   32
         Top             =   1320
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Estab.:"
         Height          =   240
         Left            =   735
         TabIndex        =   39
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Família:"
         Height          =   240
         Left            =   8175
         TabIndex        =   38
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   240
         Left            =   495
         TabIndex        =   37
         Top             =   885
         Width           =   810
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   240
         Left            =   3375
         TabIndex        =   36
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   240
         Left            =   4755
         TabIndex        =   35
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faturamento:"
         Height          =   240
         Left            =   6795
         TabIndex        =   34
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor(a):"
         Height          =   240
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.OptionButton optPeriodo 
      Caption         =   "Por Período"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      TabIndex        =   15
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton optDia 
      Caption         =   "Por Dia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtQtdeAtende 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3480
      TabIndex        =   12
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtQtdeProd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1800
      TabIndex        =   3
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtTotVendas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5160
      TabIndex        =   2
      Top             =   6840
      Width           =   1575
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "print"
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
         Left            =   6840
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
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
               Picture         =   "CONSULTAPEDIDOHORA.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTAPEDIDOHORA.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTAPEDIDOHORA.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTAPEDIDOHORA.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTAPEDIDOHORA.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTAPEDIDOHORA.frx":D2E1
               Key             =   ""
            EndProperty
         EndProperty
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
      DesignWidth     =   11865
      DesignHeight    =   7290
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3015
      Left            =   30
      TabIndex        =   6
      Top             =   3360
      Width           =   11730
      _ExtentX        =   20690
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Hora"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde.Pedidos"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "QtdeTotal.Itens"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Qtde.Itens.Atendente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total.Venda"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Tiket Médio"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   10200
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket.Médio"
      Height          =   240
      Index           =   1
      Left            =   6840
      TabIndex        =   42
      Top             =   6555
      Width           =   1515
   End
   Begin VB.Label lblDtFim 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Fim:"
      Height          =   240
      Left            =   9180
      TabIndex        =   40
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItensAtendente"
      Height          =   240
      Left            =   3480
      TabIndex        =   13
      Top             =   6555
      Width           =   1425
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.TotalVendas"
      Height          =   240
      Index           =   0
      Left            =   5145
      TabIndex        =   10
      Top             =   6555
      Width           =   1515
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos"
      Height          =   240
      Left            =   150
      TabIndex        =   9
      Top             =   6555
      Width           =   765
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
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lblDtIni 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dia:"
      Height          =   240
      Left            =   6840
      TabIndex        =   8
      Tag             =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Itens"
      Height          =   240
      Left            =   1800
      TabIndex        =   7
      Top             =   6555
      Width           =   1290
   End
End
Attribute VB_Name = "frmCONSULTAPEDIDOHORA"
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

Private Sub optDia_Click()
   lblDtIni.Caption = "Dia:"
   fraFiltro.Enabled = True
   lblDtFim.Visible = False
   txtDtFim.Visible = False
   'txtDtIni.SetFocus
End Sub

Private Sub optPeriodo_Click()
   lblDtIni.Caption = "Dt.Ini:"
   fraFiltro.Enabled = True
   lblDtFim.Visible = True
   txtDtFim.Visible = True
   txtDtFim.PromptInclude = False
   txtDtFim.Text = Date
   txtDtIni.PromptInclude = False
   txtDtIni.Text = Date
   txtDtIni.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
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
      Case vbKeyF11
         'frmSenha.Show 1

         'If UCase(CRITERIO_A) = UCase("acerto") Then
            PEDIDO_ID_N = 0
            If Not IsNull(lstPedido.SelectedItem.Text) Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  PEDIDO_ID_N = lstPedido.SelectedItem.Text

                  frmPedidoClienteAcerto.Show 1
                  MONTA_CONSULTA_SQL True
                  PEDIDO_ID_N = 0
               End If
            End If
         'End If
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
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_LostFocus"
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
   Me.Enabled = True
   Me.KeyPreview = True
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date)
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   fraFiltro.Enabled = False
   optDia.Value = True
   lblDtFim.Visible = False
   txtDtFim.Visible = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   optPeriodo.Value = False
   CLIENTE_ID_N = 0
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   lstPedido.ListItems.Clear
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""
   txtTotVendas.Text = ""
   txtQtdeAtende.Text = ""
   txtTicketMedio = ""

   If cmbVend.Enabled = True Then _
      cmbVend.Text = ""

   cmbForma.Text = ""
   cmbAuxForma.Text = ""
   txtDtIni.PromptInclude = False
   txtDtIni.Text = ""
   txtTotVendas.Text = ""
   txtReg.Text = ""
   txtQtdeProd.Text = ""

   lstPedido.Visible = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
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

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date)
   txtDtIni.PromptInclude = True

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

   cmbSituacao.AddItem "Encomenda"
   cmbSituacaoAUX.AddItem "'8'"

   cmbSituacao.AddItem "Cancelado"
   cmbSituacaoAUX.AddItem "'9'"

   cmbSituacao.Text = "Faturado"
   cmbSituacaoAUX.Text = "'3','5','7'"

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

      MONTA_CONSULTA_SQL True
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

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   lstPedido.Visible = False

   Dim QTDE_PEDIDOS_N      As Long
   Dim QTDE_ITENS_N        As Long
   Dim QTDE_ITENS_ATENDE_N As Long
   Dim HORA_QUEBRA_N         As Long

   HORA_INI = Time
   CRITERIO_A = ""

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDO_HORA", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "drop table PEDIDO_HORA"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDO_HORA", "") = False Then
      SQL = "CREATE TABLE [dbo].[PEDIDO_HORA]("
      SQL = SQL & " [DIA] [datetime] NOT NULL,"
      SQL = SQL & " [HORA_QUEBRA] [BIGINT] NOT NULL,"
      SQL = SQL & " [QTDE_PEDIDOS] [BIGINT] NOT NULL,"
      SQL = SQL & " [QTDE_ITENS] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE_ITENS_ATENDE] [bigint] NOT NULL,"
      SQL = SQL & " [TOTAL_VENDAS] [Float]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   SQL = "delete pedido_hora"
   CONECTA_RETAGUARDA.Execute SQL

   VALOR_TOTAL_N = 0
   txtReg.Text = ""
   txtQtdeProd.Text = ""
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   HORA_QUEBRA_N = 0
   txtTotVendas.Text = ""
   txtQtdeAtende.Text = ""
   PEDIDO_ID_N = 0
   VALOR_TOTAL_N = 0
   CONTA_REGISTRO_N = 0

   Me.Enabled = False
   Me.KeyPreview = False

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

   If optDia.Value = False And optPeriodo.Value = False Then _
      optDia.Value = True

   If TabTemp.State = 1 Then _
      TabTemp.Close

   QTDE_PEDIDOS_N = 0
   QTDE_ITENS_N = 0
   QTDE_ITENS_ATENDE_N = 0
   CONTA_REGISTRO_N = 0

   SQL = "select PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.VENDEDOR_ID, "
   SQL = SQL & " PEDIDO.USUARIO_ID, PEDIDO.DT_REQ, PEDIDO.STATUS as StatusPedido,"
   SQL = SQL & " PEDIDO.ESTABELECIMENTO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID,"
   SQL = SQL & " PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.STATUS AS StatusItem, PEDIDOITEM.USU_ATENDE"
   SQL = SQL & " from PEDIDO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"
   SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"

   SQL = SQL & " where PEDIDO.pedido_id Is Not Null"
   SQL = SQL & " and estabelecimento_id = " & cmbEstabAUX.Text

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and produto_id = " & PRODUTO_ID_N

   If CLIENTE_ID_N > 0 Then _
      SQL = SQL & " and cliente_id = " & CLIENTE_ID_N

   If Trim(cmbVend.Text) <> "" Then _
      SQL = SQL & " and vendedor_id = " & cmbVendAux.Text

   If Trim(cmbSituacaoAUX.Text) <> "" Then _
      SQL = SQL & " and PEDIDO.status in (" & Trim(cmbSituacaoAUX.Text) & ")"

   'If Trim(cmbAuxForma.Text) <> "" Then _
      SQL = SQL & " and  = " & cmbAuxForma.Text

   If optDia.Value = True Then
      If IsDate(txtDtIni.Text) Then
         SQL = SQL & " and dt_req >= '" & DMA(txtDtIni.Text, "i") & "'"
         SQL = SQL & " and dt_req <= '" & DMA(txtDtIni.Text, "f") & "'"
      End If
      Else
         If optPeriodo.Value = True Then
            If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
               SQL = SQL & " and dt_req >= '" & DMA(txtDtIni.Text, "i") & "'"
               SQL = SQL & " and dt_req <= '" & DMA(txtDtFim.Text, "f") & "'"
            End If
         End If
   End If
   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

SQL = SQL & " order by pedido.dt_req "

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If
   While Not TabTemp.EOF
      DoEvents

      If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
         PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value
         CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1

            If HORA_QUEBRA_N <> Hour(TabTemp.Fields("dt_req").Value) Then
               HORA_QUEBRA_N = Hour(TabTemp.Fields("dt_req").Value)
               QTDE_PEDIDOS_N = 1
               QTDE_ITENS_N = 0
               QTDE_ITENS_ATENDE_N = 0
               VALOR_TOTAL_N = 0
            End If

            'quantidade de itens no pedido
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
               SQL = "select count(produto_id) from PEDIDOITEM "
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
               SQL = SQL & " and status <> 'C'"
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then _
                  If Not IsNull(TabConsulta.Fields(0).Value) Then _
                     QTDE_ITENS_N = TabConsulta.Fields(0).Value
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            'quantidade de itens no pedido por atendente
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
               SQL = "select count(produto_id) from PEDIDOITEM "
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
               SQL = SQL & " and usu_atende > 0 "
               SQL = SQL & " and status <> 'C'"
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then _
                  If Not IsNull(TabConsulta.Fields(0).Value) Then _
                     QTDE_ITENS_ATENDE_N = TabConsulta.Fields(0).Value
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            'valor total pedido
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
               SQL = "select sum(qtd_pedida*valor_item) from PEDIDOITEM "
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
               SQL = SQL & " and status <> 'C'"
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then _
                  If Not IsNull(TabConsulta.Fields(0).Value) Then _
                     VALOR_TOTAL_N = TabConsulta.Fields(0).Value
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from PEDIDO_HORA "
            SQL = SQL & " where hora_quebra = " & HORA_QUEBRA_N
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabConsulta.EOF Then
               SQL = "insert into PEDIDO_HORA "
                  SQL = SQL & "(DIA,HORA_QUEBRA,QTDE_PEDIDOS,QTDE_ITENS,QTDE_ITENS_ATENDE,TOTAL_VENDAS)"
               SQL = SQL & " values("

                  SQL = SQL & "'" & TabTemp.Fields("dt_req").Value & "'"         'dia
                  SQL = SQL & "," & HORA_QUEBRA_N                                'HORA_QUEBRA
                  SQL = SQL & "," & 1  'QTDE_PEDIDOS_N                               'QTDE_PEDIDOS
                  SQL = SQL & "," & QTDE_ITENS_N                                 'QTDE_ITENS
                  SQL = SQL & "," & QTDE_ITENS_ATENDE_N                          'QTDE_ITENS_ATENDE
                  SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)                       'TOTAL_VENDAS

               SQL = SQL & ")"
               Else
                  'QTDE_PEDIDOS_N = QTDE_PEDIDOS_N + 1
                  SQL = "update PEDIDO_HORA set "
                     SQL = SQL & "  QTDE_PEDIDOS = QTDE_PEDIDOS + 1" '& QTDE_PEDIDOS_N                  'QTDE_PEDIDOS
                     SQL = SQL & ", QTDE_ITENS = QTDE_ITENS + " & QTDE_ITENS_N                        'QTDE_ITENS
                     SQL = SQL & ", QTDE_ITENS_ATENDE = QTDE_ITENS_ATENDE + " & QTDE_ITENS_ATENDE_N   'QTDE_ITENS_ATENDE
                     SQL = SQL & ", TOTAL_VENDAS = TOTAL_VENDAS + " & tpMOEDA(VALOR_TOTAL_N)          'TOTAL_VENDAS
                  SQL = SQL & " where hora_quebra = " & HORA_QUEBRA_N
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
            CONECTA_RETAGUARDA.Execute SQL

      End If
      txtReg.Text = CONTA_REGISTRO_N
      txtReg.Refresh

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONTA_REGISTRO_N = 0

   SQL = "select * from PEDIDO_HORA "
   SQL = SQL & " order by hora_quebra "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      DoEvents

      Set item = lstPedido.ListItems.Add(, "seq." & Hour(TabTemp.Fields("dia").Value), Hour(TabTemp.Fields("dia").Value))
      item.SubItems(1) = "" & Hour(TabTemp.Fields("dia").Value)
      item.SubItems(2) = "" & TabTemp.Fields("qtde_pedidos").Value
      item.SubItems(3) = "" & TabTemp.Fields("QTDE_ITENS").Value
      item.SubItems(4) = "" & TabTemp.Fields("QTDE_ITENS_ATENDE").Value
      item.SubItems(5) = "" & Format(TabTemp.Fields("TOTAL_VENDAS").Value, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(TabTemp.Fields("TOTAL_VENDAS").Value / TabTemp.Fields("qtde_pedidos").Value, strFormatacao2Digitos)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select sum(qtde_pedidos) from PEDIDO_HORA "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      txtReg.Text = "" & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(qtde_itens) from PEDIDO_HORA "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      txtQtdeProd.Text = "" & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(qtde_itens_atende) from PEDIDO_HORA "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      txtQtdeAtende.Text = "" & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(total_vendas) from PEDIDO_HORA "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      txtTotVendas.Text = "" & Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   QTDE_N = 0 & txtReg.Text
   VALOR_TOTAL_N = 0 & txtTotVendas.Text
   txtTicketMedio = Format(VALOR_TOTAL_N / QTDE_N, strFormatacao2Digitos)

DoEvents

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub
