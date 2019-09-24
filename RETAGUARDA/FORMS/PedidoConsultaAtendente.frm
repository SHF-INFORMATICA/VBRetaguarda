VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoConsultaAtendente 
   Caption         =   "Consulta Pedido Atendente"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoConsultaAtendente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Atualiza Cadastro Atendente (bc.xls)"
      Height          =   495
      Left            =   6480
      TabIndex        =   37
      Top             =   6840
      Width           =   1455
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
      Left            =   6600
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
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
      Left            =   1680
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   870
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
      Left            =   8160
      TabIndex        =   18
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1335
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
      Picture         =   "PedidoConsultaAtendente.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtQtdeProd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1200
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   1440
      TabIndex        =   7
      Top             =   2760
      Width           =   3255
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
      Picture         =   "PedidoConsultaAtendente.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtDescProd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   14
      Top             =   2280
      Width           =   5655
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
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtCli 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   13
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtTotalVenda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   8010
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cmbForma 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ComboBox cmbVend 
      Height          =   360
      Left            =   6480
      TabIndex        =   4
      Top             =   1320
      Width           =   3105
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
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
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
               Picture         =   "PedidoConsultaAtendente.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaAtendente.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaAtendente.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaAtendente.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaAtendente.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoConsultaAtendente.frx":D2E1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
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
      DesignWidth     =   9690
      DesignHeight    =   7335
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3015
      Left            =   50
      TabIndex        =   20
      Top             =   3360
      Width           =   9570
      _ExtentX        =   16880
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Atendente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Vlr.Venda"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "QtdeItensPedido"
         Object.Width           =   3919
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "QtdeItensAtendente"
         Object.Width           =   3919
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "VlrAtendente"
         Object.Width           =   3919
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "TICKETEMEDIO"
         Object.Width           =   3919
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "DtPedido"
         Object.Width           =   3919
      EndProperty
   End
   Begin Threed.SSOption optSintetico 
      Height          =   270
      Left            =   6960
      TabIndex        =   21
      Top             =   2640
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
      Left            =   6960
      TabIndex        =   22
      Top             =   2880
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.Faturado"
      Height          =   240
      Index           =   1
      Left            =   8040
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Produtos"
      Height          =   240
      Left            =   1230
      TabIndex        =   33
      Top             =   6600
      Width           =   1785
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Família:"
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   240
      Left            =   3480
      TabIndex        =   31
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   2280
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   6840
      TabIndex        =   28
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   240
      Left            =   4740
      TabIndex        =   26
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos"
      Height          =   240
      Left            =   150
      TabIndex        =   25
      Top             =   6600
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Balconista:"
      Height          =   240
      Left            =   5145
      TabIndex        =   23
      Top             =   1320
      Width           =   1230
   End
End
Attribute VB_Name = "frmPedidoConsultaAtendente"
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

Private Sub txtReg_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtQtdeProd_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtTotalVenda_GotFocus()
   txtPedido.SetFocus
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      MONTA_CONSULTA_SQL
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
      'Unload Me
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imprimir"
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         FORMULA_REL = ""

         If optAnalitico.Value = True Then
            Nome_Relatorio = "rel_atende_analitico.rpt"
            Else: Nome_Relatorio = "rel_atende_sintetico.rpt"
         End If

         frmRELATORIO10.Show 1
      Case "consultar"
         MONTA_CONSULTA_SQL
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

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   lstPedido.ListItems.Clear
   txtPedido.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""
   txtQtdeProd.Text = ""

   If cmbVend.Enabled = True Then _
      cmbVend.Text = ""

   cmbForma.Text = ""
   cmbAuxForma.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtTotalVenda.Text = ""
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

         SQL = "select nome,usuario_id from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and tipo = 8"
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            cmbVend.Text = TabUSU!NOME
            cmbVendAux.Text = TabUSU!USUARIO_ID
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
      Label7(0).Visible = False
      Label7(1).Visible = False
      Label18.Visible = False
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
   cmbVendAux.Clear

   SQL = "select nome,usuario_id from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where tipo = 8"
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbVend.AddItem Trim(TabTemp!NOME) & " - " & Trim(TabTemp!USUARIO_ID)
      cmbVendAux.AddItem Trim(TabTemp!USUARIO_ID)
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

Private Sub MONTA_CONSULTA_SQL()
'On Error GoTo ERRO_TRATA

   HORA_INI = Time

   CRIA_TABELA_TEMPORARIA

   CHECA_ULTIMO_DIA_MES

   Dim TabPedido                       As New ADODB.Recordset
   Dim QTDE_ITENS_PEDIDO_N             As Long
   Dim QTDE_ITENS_ATENDENTE_N          As Long
   Dim QTDE_VENDIDA_ATENDENTE_N        As Double
   Dim VALOR_VENDA_ATENDENTE_N         As Double
   Dim PERC_PARTICIPACAO_ATENDENTE_N   As Double
   Dim VALOR_TOTAL_PEDIDO_N            As Double
   Dim TOT_VENDA_QUE_TEVE_ATENDENTE_N  As Double

   QTDE_ITENS_PEDIDO_N = 0
   QTDE_ITENS_ATENDENTE_N = 0
   QTDE_VENDIDA_ATENDENTE_N = 0
   VALOR_VENDA_ATENDENTE_N = 0
   VALOR_TOTAL_PEDIDO_N = 0
   TOT_VENDA_QUE_TEVE_ATENDENTE_N = 0

   txtTotalVenda.Text = ""
   txtReg.Text = ""
   txtQtdeProd.Text = ""

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
SQL = ""
   sql_cabeca = "select PEDIDOITEM.USU_ATENDE, PEDIDO.CLIENTE_ID, PEDIDO.VENDEDOR_ID, "
   sql_cabeca = sql_cabeca & " PEDIDO.DT_REQ, PEDIDO.STATUS AS StatusPedido, PEDIDOITEM.PEDIDO_ID, "
   sql_cabeca = sql_cabeca & " PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA, "
   sql_cabeca = sql_cabeca & " PEDIDOITEM.VALOR_ITEM as PrVenda, PEDIDOITEM.STATUS AS StatusItem,"
   sql_cabeca = sql_cabeca & " USUARIO.PESSOA_ID , PESSOA.CNPJCPF, PESSOA.DESCRICAO"
   sql_cabeca = sql_cabeca & " from PEDIDO WITH (NOLOCK)"
   sql_cabeca = sql_cabeca & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   sql_cabeca = sql_cabeca & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   sql_cabeca = sql_cabeca & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   sql_cabeca = sql_cabeca & " INNER JOIN USUARIO WITH (NOLOCK)"
   sql_cabeca = sql_cabeca & " ON PEDIDOITEM.USU_ATENDE = USUARIO.USUARIO_ID "
   sql_cabeca = sql_cabeca & " INNER JOIN PESSOA WITH (NOLOCK)"
   sql_cabeca = sql_cabeca & " ON USUARIO.PESSOA_ID = PESSOA.PESSOA_ID"
   sql_cabeca = sql_cabeca & " where PEDIDOITEM.pedido_id > 0"

SQL_COUNT = "select count(pedido_id) from PEDIDO "

   'SQL = SQL & " where PEDIDOITEM.usu_atende > 0"

   SQL_COUNT = SQL_COUNT & " where PEDIDO.pedido_id > 0"

   If Trim(txtPedido.Text) <> "" Then _
      SQL = SQL & " and PEDIDO.pedido_id = " & txtPedido.Text

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then _
      If CLIENTE_ID_N > 0 Then _
         SQL = SQL & " and cliente_id = " & CLIENTE_ID_N
   txtCNPJCPF.PromptInclude = True

   If Trim(cmbSituacaoAUX.Text) <> "" Then _
      SQL = SQL & " and PEDIDO.status in (" & Trim(cmbSituacaoAUX.Text) & ")"

   'If Trim(cmbAuxForma.Text) <> "" Then _
      SQL = SQL & " and PEDIDO.tipovenda_id = " & cmbAuxForma.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_req <= '" & txtDtFim.Text & "'"
   End If

SQL_COUNT = SQL_COUNT & SQL
'=================
   If Trim(cmbVend.Text) <> "" Then _
      SQL = SQL & " and usu_atende = " & cmbVendAux.Text

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and PEDIDOITEM.produto_id = " & PRODUTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   SQL = SQL & " order by pedidoitem.usu_atende, pedidoitem.pedido_id "

sql_cabeca = sql_cabeca & SQL
   CONTA_REGISTRO_N = 0

   If TabPedido.State = 1 Then _
      TabPedido.Close

   TabPedido.Open sql_cabeca, CONECTA_RETAGUARDA, , , adCmdText
   If TabPedido.EOF Then
      If TabPedido.State = 1 Then _
         TabPedido.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   While Not TabPedido.EOF
   'SET
      ATENDENTE_ID_N = 0 & TabPedido.Fields("usu_atende").Value
      PEDIDO_ID_N = 0 & TabPedido.Fields("PEDIDO_ID").Value
'==============
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select * from REL_ATENDENTE WITH (NOLOCK)"
      SQL = SQL & " where PEDIDO_ID = " & PEDIDO_ID_N
      SQL = SQL & " and ATENDENTE_ID = " & ATENDENTE_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
'==============='pegando qtde de itens no pedido
         QTDE_ITENS_PEDIDO_N = 0
         If TabAUX.State = 1 Then _
            TabAUX.Close
         SQL = "select count(produto_id) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and status <> 'C'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            QTDE_ITENS_PEDIDO_N = 0 & TabAUX.Fields(0).Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
'===============
'==============='pegando qtde de itens no pedido POR ATENDENTE
         QTDE_ITENS_ATENDENTE_N = 0
         If TabAUX.State = 1 Then _
            TabAUX.Close
         SQL = "select count(produto_id) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and usu_atende = " & ATENDENTE_ID_N
         SQL = SQL & " and status <> 'C'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            QTDE_ITENS_ATENDENTE_N = 0 & TabAUX.Fields(0).Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
'===============
'==============='pegando qtde VENDIA no pedido POR ATENDENTE
         QTDE_VENDIDA_ATENDENTE_N = 0
         If TabAUX.State = 1 Then _
            TabAUX.Close
         SQL = "select sum(qtd_pedida) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and usu_atende = " & ATENDENTE_ID_N
         SQL = SQL & " and status <> 'C'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            QTDE_VENDIDA_ATENDENTE_N = 0 & TabAUX.Fields(0).Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
'===============
'==============='pegando qtde VENDIA no pedido POR ATENDENTE
         VALOR_VENDA_ATENDENTE_N = 0
         If TabAUX.State = 1 Then _
            TabAUX.Close
         SQL = "select sum(qtd_pedida*valor_item) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and usu_atende = " & ATENDENTE_ID_N
         SQL = SQL & " and status <> 'C'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            VALOR_VENDA_ATENDENTE_N = 0 & TabAUX.Fields(0).Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
'===============
'==============='pegando total do pedido
         VALOR_TOTAL_PEDIDO_N = 0
         If TabAUX.State = 1 Then _
            TabAUX.Close
         SQL = "select sum(qtd_pedida*valor_item) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and status <> 'C'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            VALOR_TOTAL_PEDIDO_N = 0 & TabAUX.Fields(0).Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
'===============
         TOT_VENDA_QUE_TEVE_ATENDENTE_N = 0
         If TabAUX.State = 1 Then _
            TabAUX.Close
         SQL = "select sum(qtd_pedida*valor_item) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and usu_atende > 0 "
         SQL = SQL & " and status <> 'C'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            TOT_VENDA_QUE_TEVE_ATENDENTE_N = 0 & TabAUX.Fields(0).Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
'==============='calculo tickete medio
'PERC_PARTICIPACAO_ATENDENTE_N = 0 & (VALOR_TOTAL_PEDIDO_N / VALOR_VENDA_ATENDENTE_N)
PERC_PARTICIPACAO_ATENDENTE_N = 0 & (VALOR_VENDA_ATENDENTE_N / TOT_VENDA_QUE_TEVE_ATENDENTE_N)

'===============
         SQL = "insert into REL_ATENDENTE "
   
         SQL = SQL & "("
            SQL = SQL & "estabelecimento_ID,PEDIDO_ID,ATENDENTE_ID,DT_VENDA,"
            SQL = SQL & "QTDE_VENDIDA,VALOR_VENDIDO_ATENDENTE,TICKETEMEDIO,DT_INI,DT_FIM,"
            SQL = SQL & "QTDE_ITENS_ATENDENTE,QTDE_ITENS_PEDIDO "
         SQL = SQL & ")"
   
         SQL = SQL & " values("
            SQL = SQL & ESTABELECIMENTO_ID_N                                                 'estabelecimento_ID
            SQL = SQL & "," & PEDIDO_ID_N                                                    'PEDIDO_ID
            SQL = SQL & "," & ATENDENTE_ID_N                                                 'ATENDENTE_ID
            SQL = SQL & ",'" & DMA(TabPedido.Fields("DT_REQ").Value) & "'"                   'DT_VENDA
            SQL = SQL & "," & tpMOEDA(QTDE_VENDIDA_ATENDENTE_N)                              'QTDE_VENDIDA
            SQL = SQL & "," & tpMOEDA(VALOR_VENDA_ATENDENTE_N)                               'VALOR_VENDIDO_ATENDENTE
            SQL = SQL & "," & tpMOEDA(PERC_PARTICIPACAO_ATENDENTE_N * VALOR_TOTAL_PEDIDO_N)  'TICKETEMEDIO
            SQL = SQL & ",'" & Trim(txtDtIni.Text) & "'"                                     'DT_INI
            SQL = SQL & ",'" & Trim(txtDtFim.Text) & "'"                                     'DT_FIM
            SQL = SQL & "," & QTDE_ITENS_ATENDENTE_N                                         'QTDE_ITENS_atendente
            SQL = SQL & "," & QTDE_ITENS_PEDIDO_N                                            'QTDE_ITENS_PEDIDo
         SQL = SQL & ")"

         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
      txtReg.Text = CONTA_REGISTRO_N
      txtReg.Refresh
      DoEvents
      TabPedido.MoveNext
   Wend

   SETA_GRID

  If TabTemp.State = 1 Then _
      TabTemp.Close
   TabTemp.Open SQL_COUNT, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabTemp.Fields(0).Value) Then
      txtReg.Text = TabTemp.Fields(0).Value
      txtReg.Refresh
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False

   Dim TabAtendente                     As New ADODB.Recordset
   Dim VALOR_TOTAL_PEDIDO_N            As Double
   Dim ATENDENTE_ID_N                  As Long
   Dim VALOR_VENDA_ATENDENTE_N         As Double
   Dim NOME_ATENDENTE_A                As String
   Dim TOT_PEDIDO_ATENDENTE_N          As Long

   VALOR_TOTAL_PEDIDO_N = 0
   VALOR_DESCONTO_N = 0
   PEDIDO_ID_N = 0
   CLIENTE_ID_N = 0
   PEDIDO_ID_N = 0
   VALOR_TOTAL_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_VENDA_ATENDENTE_N = 0

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

'============================
   If TabAtendente.State = 1 Then _
      TabAtendente.Close

   SQL = "select * from REL_ATENDENTE WITH (NOLOCK)"
   SQL = SQL & " order by pedido_id,ATENDENTE_ID "
   TabAtendente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabAtendente.EOF Then
      If TabAtendente.State = 1 Then _
         TabAtendente.Close
      MsgBox "Nenhum registro encontrado."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If
   ATENDENTE_ID_N = 0
   PEDIDO_ID_N = 0

   While Not TabAtendente.EOF
      PEDIDO_ID_N = TabAtendente.Fields("pedido_id").Value
      ATENDENTE_ID_N = TabAtendente.Fields("ATENDENTE_ID").Value
      NOME_ATENDENTE_A = ""

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select nome from USUARIO "
      SQL = SQL & " where usuario_id = " & ATENDENTE_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         NOME_ATENDENTE_A = "" & Trim(TabConsulta.Fields(0).Value)
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
      Set item = lstPedido.ListItems.Add(, "seq." & CONTA_REGISTRO_N, TabAtendente.Fields("PEDIDO_ID").Value)
      item.SubItems(1) = "" & Trim(NOME_ATENDENTE_A)

      'BUSCA VALOR TOTAL VENDA
      VALOR_TOTAL_PEDIDO_N = 0
      SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabAtendente.Fields("pedido_id").Value
      SQL = SQL & " and status <> 'C' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_TOTAL_PEDIDO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_TOTAL_PEDIDO_N

      item.SubItems(2) = "" & Format(VALOR_TOTAL_PEDIDO_N, strFormatacao2Digitos)
      item.SubItems(3) = "" & TabAtendente.Fields("QTDE_ITENS_PEDIDO").Value
      item.SubItems(4) = "" & TabAtendente.Fields("QTDE_ITENS_ATENDENTE").Value

      'pegando total venda que o atendente participou
      VALOR_VENDA_ATENDENTE_N = 0 & TabAtendente.Fields("VALOR_VENDIDO_ATENDENTE").Value

      item.SubItems(5) = "" & Format(VALOR_VENDA_ATENDENTE_N, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(TabAtendente.Fields("ticketemedio").Value, strFormatacao2Digitos)
      item.SubItems(7) = ""

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select dt_req from PEDIDO "
      SQL = SQL & " where PEDIDO_id = " & PEDIDO_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         item.SubItems(7) = "" & TabConsulta.Fields("dt_req").Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
      txtReg.Text = CONTA_REGISTRO_N
      txtReg.Refresh
      DoEvents

      TabAtendente.MoveNext
   Wend
   If TabAtendente.State = 1 Then _
      TabAtendente.Close
'=====================
   SQL = "select distinct(ATENDENTE_ID) from REL_ATENDENTE WITH (NOLOCK)"
   TabAtendente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabAtendente.EOF
      ATENDENTE_ID_N = TabAtendente.Fields("ATENDENTE_ID").Value

'===================== pega total de pedidos do atendente
      TOT_PEDIDO_ATENDENTE_N = 0
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select DISTINCT COUNT(pedido_id) from REL_ATENDENTE"
      SQL = SQL & " where atendente_id = " & ATENDENTE_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            TOT_PEDIDO_ATENDENTE_N = 0 & TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
'===================== pega total VALOR dOS pedidos do atendente
      VALOR_VENDA_ATENDENTE_N = 0
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select sum(ticketemedio) from REL_ATENDENTE"
      SQL = SQL & " where atendente_id = " & ATENDENTE_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            VALOR_VENDA_ATENDENTE_N = 0 & TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = " update REL_ATENDENTE set "
         SQL = SQL & " tot_ticketemedio = " & tpMOEDA((VALOR_VENDA_ATENDENTE_N / TOT_PEDIDO_ATENDENTE_N))
      SQL = SQL & " where atendente_id = " & ATENDENTE_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      TabAtendente.MoveNext
   Wend
   If TabAtendente.State = 1 Then _
      TabAtendente.Close

'=====================
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   SQL = "select count(pedido_id) from REL_ATENDENTE"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         txtReg.Text = "" & TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum(qtde_itens_pedido) from REL_ATENDENTE WITH (NOLOCK)"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         txtQtdeProd.Text = "" & TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub CRIA_TABELA_TEMPORARIA()
'On Error GoTo ERRO_TRATA

   Dim strSQL As String

   If EXISTE_OBJ_BANCO("RETAGUARDA", "REL_ATENDENTE", "U") = True Then
      strSQL = "drop table REL_ATENDENTE"
      CONECTA_RETAGUARDA.Execute strSQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "REL_ATENDENTE", "U") = False Then
      strSQL = "create table REL_ATENDENTE"
      strSQL = strSQL & " ("
         strSQL = strSQL & " ESTABELECIMENTO_ID       int        null,"
         strSQL = strSQL & " PEDIDO_ID                bigint     null,"
         strSQL = strSQL & " ATENDENTE_ID             bigint     null,"
         strSQL = strSQL & " DT_VENDA                 datetime   null,"
         strSQL = strSQL & " QTDE_VENDIDA             float      null,"
         strSQL = strSQL & " QTDE_ITENS_ATENDENTE     INT        null,"
         strSQL = strSQL & " QTDE_ITENS_PEDIDO        INT        null,"
         strSQL = strSQL & " VALOR_VENDIDO_ATENDENTE  float      null,"
         strSQL = strSQL & " TICKETEMEDIO             float      null,"
         strSQL = strSQL & " TOT_TICKETEMEDIO         float      null,"
         strSQL = strSQL & " DT_INI                   datetime   null,"
         strSQL = strSQL & " DT_FIM                   datetime   null,"
      strSQL = strSQL & " )"
      CONECTA_RETAGUARDA.Execute strSQL
   End If

   strSQL = "delete from REL_ATENDENTE"
   CONECTA_RETAGUARDA.Execute strSQL

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CRIA_TABELA_TEMPORARIA"
End Sub
'============
Private Sub Command1_Click()

   Proc_n = 0
   Novos_n = 0
   At_n = 0

   Set oConn = New ADODB.Connection
   oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                      "FIL=excel 8.0;" & _
                      "DefaultDir=c:\MEGASIM\txt\;" & _
                      "MaxBufferSize=2048;" & _
                      "PageTimeout=5;" & _
                      "DBQ=c:\MEGASIM\txt\bc.xls;"

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'aabre o recordset pelo nome da planilha
   TabConsulta.Open "[Plan1$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

   TabConsulta.MoveFirst

   If TabConsulta.EOF Then
      MsgBox "Planilha incorreta !!!"
      Exit Sub
   End If

   While Not TabConsulta.EOF
      Proc_n = Proc_n + 1
      Command1.Caption = "Processados = " & Proc_n

'MsgBox TabConsulta(2).Value

      If Not IsNull(TabConsulta(2).Value) Then
         If Trim(TabConsulta(2).Value) <> "" Then
            If Not IsNull(TabConsulta(0).Value) Then
               If Trim(TabConsulta(0).Value) <> "" Then
                  If IsNumeric(TabConsulta(0).Value) Then
                     CRITERIO_A = Replace(TabConsulta(2).Value, ".", "")
                     CRITERIO_A = Replace(CRITERIO_A, "-", "")
                     CRITERIO_A = Replace(CRITERIO_A, "/", "")

'MsgBox CRITERIO_A

                     If TabTemp.State = 1 Then _
                        TabTemp.Close

                     SQL = "select usuario_id from USUARIO "
                     SQL = SQL & " where cpf = '" & Trim(CRITERIO_A) & "'"
                     TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabTemp.EOF Then
                        SQL = "delete from estabelecimentoacesso "
                        SQL = SQL & " where usuario_id = " & TabTemp.Fields(0).Value
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from permissao "
                        SQL = SQL & " where usuid = " & TabTemp.Fields(0).Value
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "update usuario set "
                        SQL = SQL & " usuario_id = " & TabConsulta(0).Value
                        SQL = SQL & " where cpf = '" & Trim(CRITERIO_A) & "'"
                        CONECTA_RETAGUARDA.Execute SQL
                        Else
                        '=========================PESSOA
                           PESSOA_ID_N = 0
                           If TabPessoa.State = 1 Then _
                              TabPessoa.Close
                           SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
                           SQL = SQL & " where CNPJcpf = '" & Trim(CRITERIO_A) & "'"
                           TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabPessoa.EOF Then _
                              If Not IsNull(TabPessoa.Fields(0).Value) Then _
                                 PESSOA_ID_N = Trim(TabPessoa.Fields(0).Value)
                           If TabPessoa.State = 1 Then _
                              TabPessoa.Close

                           'executa stored procedure spPessoa
                           CONT_N = 1
                           If PESSOA_ID_N <= 0 Then _
                              spPessoa 1, _
                              PESSOA_ID_N, _
                              Trim(CRITERIO_A), _
                              Trim(TabConsulta(1).Value), _
                              "", _
                              "A"

                           PESSOA_ID_N = 0
                           If TabCliente.State = 1 Then _
                              TabCliente.Close
                           SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
                           SQL = SQL & " where CNPJcpf = '" & Trim(CRITERIO_A) & "'"
                           TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabCliente.EOF Then _
                              If Not IsNull(TabCliente.Fields(0).Value) Then _
                                 PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
                           If TabCliente.State = 1 Then _
                              TabCliente.Close
                        '=========================

                           SQL = "INSERT INTO USUARIO "
                              SQL = SQL & " (empresa_id, usuario_id, Nome, Senha, Cpf, "
                              SQL = SQL & " DtNasc, Tipo, Perc_desconto, Perc_Comissao, "
                              SQL = SQL & " Status, Logon, Classe, Pessoa_id,FUNCIONARIO) "
                           SQL = SQL & " VALUES ("
                              SQL = SQL & EMPRESA_ID_N                        'empresa_id
                              SQL = SQL & "," & TabConsulta(0).Value
                              SQL = SQL & ",'" & Trim(TabConsulta(1).Value) & "'"     'Nome
                              SQL = SQL & ",'" & Trim("123") & "'"    'Senha
                              SQL = SQL & ",'" & Trim(CRITERIO_A) & "'"      'Cpf
                              SQL = SQL & ",'" & DMA(Date) & "'"    'DtNasc
                              SQL = SQL & "," & 8                'Tipo
                              SQL = SQL & "," & tpMOEDA(0)     'Perc_desconto
                              SQL = SQL & "," & tpMOEDA(0)        'Perc_Comissao
                              SQL = SQL & "," & 1                             'Status
                              SQL = SQL & ",'" & Trim(Left(TabConsulta(1).Value, 5)) & "'"   'Logon
                              SQL = SQL & ",'A'"                              'Classe
                              SQL = SQL & "," & PESSOA_ID_N                   'Pessoa_id
                              SQL = SQL & ",'true'"              'FUNCIONARIO
                              SQL = SQL & ")"
                           CONECTA_RETAGUARDA.Execute SQL
                     End If
                     If TabTemp.State = 1 Then _
                        TabTemp.Close
                  End If
               End If
            End If
         End If
      End If
      DoEvents
On Error Resume Next
      TabConsulta.MoveNext
Err.Clear
   Wend

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

MsgBox "OK"

End Sub

Private Sub SETA_GRID_old()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False

   Dim TabTempGrid                     As New ADODB.Recordset
   Dim Conta_Produto_N                 As Long
   Dim ValorTabela_N                   As Double
   Dim VALOR_TOTAL_PEDIDO_N            As Double
   Dim ATENDENTE_ID_N                  As Long
   Dim VALOR_VENDA_ATENDENTE_N         As Double
   Dim PERC_PARTICIPACAO_ATENDENTE_N   As Double
   Dim NOME_ATENDENTE_A                As String

   VALOR_TOTAL_PEDIDO_N = 0
   VALOR_DESCONTO_N = 0
   PEDIDO_ID_N = 0
   CLIENTE_ID_N = 0
   ValorTabela_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   PERC_PARTICIPACAO_ATENDENTE_N = 0
   VALOR_VENDA_ATENDENTE_N = 0

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

'============================
   If TabTempGrid.State = 1 Then _
      TabTempGrid.Close

   SQL = "select * from REL_ATENDENTE WITH (NOLOCK)"
   SQL = SQL & " order by ATENDENTE_ID "
   TabTempGrid.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTempGrid.EOF Then
      If TabTempGrid.State = 1 Then _
         TabTempGrid.Close
      MsgBox "Nenhum registro encontrado."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If
   ATENDENTE_ID_N = 0
   PEDIDO_ID_N = 0

   While Not TabTempGrid.EOF
      If PEDIDO_ID_N <> TabTempGrid.Fields("pedido_id").Value And ATENDENTE_ID_N <> TabTempGrid.Fields("ATENDENTE_ID").Value Then
         PEDIDO_ID_N = TabTempGrid.Fields("pedido_id").Value
         ATENDENTE_ID_N = TabTempGrid.Fields("ATENDENTE_ID").Value
         NOME_ATENDENTE_A = ""

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select nome from USUARIO "
         SQL = SQL & " where usuario_id = " & ATENDENTE_ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            NOME_ATENDENTE_A = "" & Trim(TabConsulta.Fields(0).Value)
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
         Set item = lstPedido.ListItems.Add(, "seq." & CONTA_REGISTRO_N, TabTempGrid.Fields("PEDIDO_ID").Value)
         item.SubItems(1) = "" & Trim(NOME_ATENDENTE_A)

         'BUSCA VALOR TOTAL VENDA
         VALOR_TOTAL_PEDIDO_N = 0

         SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabTempGrid.Fields("pedido_id").Value
         SQL = SQL & " and status <> 'C' "
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
            VALOR_TOTAL_PEDIDO_N = TabPedidoItem.Fields(0).Value
         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_TOTAL_PEDIDO_N

         item.SubItems(2) = "" & Format(VALOR_TOTAL_PEDIDO_N, strFormatacao2Digitos)
         item.SubItems(3) = "" & TabTempGrid.Fields("QTDE_ITENS_PEDIDO").Value
         item.SubItems(4) = "" & TabTempGrid.Fields("QTDE_ITENS_ATENDENTE").Value

'===============
         'pegando qtde de itens no pedido
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select count(produto_id) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & TabTempGrid.Fields("pedido_id").Value
         SQL = SQL & " and status <> 'C'"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            item.SubItems(3) = "" & TabConsulta.Fields(0).Value
            SQL = "update REL_ATENDENTE set "
            SQL = SQL & " qtde_itens_pedido = " & TabConsulta.Fields(0).Value
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         'pegando total venda que o atendente participou
         VALOR_VENDA_ATENDENTE_N = 0 & TabTempGrid.Fields("VALOR_VENDIDO_ATENDENTE").Value
         'If TabConsulta.State = 1 Then _
            TabConsulta.Close

         'SQL = "select sum(qtd_pedida*valor_item) from PEDIDOITEM "
         'SQL = SQL & " where pedido_id = " & TabTempGrid.Fields("pedido_id").Value
         'SQL = SQL & " and usu_atende = " & TabTempGrid.Fields("atendente_id").Value
         'TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         'If Not TabConsulta.EOF Then _
            VALOR_VENDA_ATENDENTE_N = 0 & TabConsulta.Fields(0).Value
         'If TabConsulta.State = 1 Then _
            TabConsulta.Close

PERC_PARTICIPACAO_ATENDENTE_N = 0 & (VALOR_TOTAL_PEDIDO_N / VALOR_VENDA_ATENDENTE_N)

         'TICKETEMEDIO
         'If TabConsulta.State = 1 Then _
            TabConsulta.Close

         'SQL = "select TICKETEMEDIO from PEDIDOITEM "
         'SQL = SQL & " where pedido_id = " & TabTempGrid.Fields("pedido_id").Value
         'SQL = SQL & " and usu_atende = " & TabTempGrid.Fields("atendente_id").Value
         'TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         'If Not TabConsulta.EOF Then
            SQL = "update REL_ATENDENTE set "
               SQL = SQL & " TICKETEMEDIO = " & tpMOEDA(PERC_PARTICIPACAO_ATENDENTE_N * VALOR_TOTAL_PEDIDO_N)
            SQL = SQL & " where pedido_id = " & TabTempGrid.Fields("pedido_id").Value
            SQL = SQL & " and atendente_id = " & TabTempGrid.Fields("atendente_id").Value
            CONECTA_RETAGUARDA.Execute SQL
         'End If
         'If TabConsulta.State = 1 Then _
            TabConsulta.Close
'===================
         item.SubItems(5) = "" & Format(VALOR_VENDA_ATENDENTE_N, strFormatacao2Digitos)
         item.SubItems(6) = "" & Format(PERC_PARTICIPACAO_ATENDENTE_N * VALOR_TOTAL_PEDIDO_N, strFormatacao2Digitos)
         item.SubItems(7) = ""

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select dt_req from PEDIDO "
         SQL = SQL & " where PEDIDO_id = " & ATENDENTE_ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            item.SubItems(7) = "" & TabConsulta.Fields("dt_req").Value
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If   'If PEDIDO_ID_N <> TabTempGrid.Fields("pedido_id").Value  Then

      CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
      txtReg.Text = CONTA_REGISTRO_N
      txtReg.Refresh
      DoEvents

      TabTempGrid.MoveNext
   Wend
   If TabTempGrid.State = 1 Then _
      TabTempGrid.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
