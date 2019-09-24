VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoItemConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Faturamento Produtos "
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoItemConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   10080
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   735
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
      Left            =   10080
      TabIndex        =   7
      Top             =   960
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
      Left            =   3900
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1440
      Width           =   3945
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
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3435
      Picture         =   "PedidoItemConsulta.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   405
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   9120
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
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
      Left            =   9120
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
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
         NumButtons      =   3
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
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Imp"
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
         Left            =   8880
         TabIndex        =   15
         Top             =   360
         Width           =   735
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
               Picture         =   "PedidoItemConsulta.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoItemConsulta.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoItemConsulta.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoItemConsulta.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoItemConsulta.frx":A8FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoItemConsulta.frx":C8DF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12000
      DesignHeight    =   6705
   End
   Begin MSComctlLib.ListView lstItens 
      Height          =   4455
      Left            =   45
      TabIndex        =   9
      Top             =   2040
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   7858
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "QtdeEntrada"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "QtdeSaída"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Entrada-Saída"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblConta 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   240
      Left            =   7920
      TabIndex        =   17
      Top             =   960
      Width           =   105
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   12000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   9120
      TabIndex        =   14
      Top             =   960
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   12000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   12000
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   240
      Left            =   600
      TabIndex        =   13
      Top             =   1485
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Família:"
      Height          =   240
      Left            =   8280
      TabIndex        =   10
      Top             =   1440
      Width           =   780
   End
End
Attribute VB_Name = "frmPedidoItemConsulta"
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
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         MONTA_CONSULTA_SQL True
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         PEDIDO_ID_N = 0
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
      Case "print"
   End Select
   PEDIDO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
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
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub LSTITENS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstItens, ColumnHeader
End Sub

Private Sub cmbSituacao_Click()
'On Error GoTo ERRO_TRATA

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbsituacao_Click"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
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
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   lstItens.ListItems.Clear
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   lstItens.Visible = True

Exit Sub
ERRO_TRATA:
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
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
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

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

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

   cmbSituacao.AddItem "Faturado"
   cmbSituacaoAUX.AddItem "'3','5','7'"

   cmbSituacao.AddItem "Pendente"
   cmbSituacaoAUX.AddItem "'1','2','4'"

   cmbSituacao.AddItem "Cupom Fiscal"
   cmbSituacaoAUX.AddItem "'7'"

   cmbSituacao.AddItem "Nota Eletrônica"
   cmbSituacaoAUX.AddItem "'7','5','3'"

   cmbSituacao.AddItem "Cancelado"
   cmbSituacaoAUX.AddItem "'9'"

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

   VALOR_TOTAL_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   Dim SQL_ITENS        As String
   Dim TabTemp          As New ADODB.Recordset
   Dim Conta_Produto_N  As Long
   Dim Peso_N           As Double
   Dim QTDE_ENTRADA_N   As Double
   Dim QTDE_SAIDA_N     As Double

   Peso_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   SQL_ITENS = ""
   NUMR_ID_N = 0
   CONT_N = 0
   QTDE_ENTRADA_N = 0
   Qtde_Saida_S = 0
   VALOR_TOTAL_N = 0

   lstItens.Visible = False
   lstItens.ListItems.Clear

   CHECA_ULTIMO_DIA_MES

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FATURAPRODUTO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[FATURAPRODUTO]("
      SQL = SQL & " [FATURAPRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DT_INI] [datetime] NOT NULL,"
      SQL = SQL & " [DT_FIM] [datetime] NOT NULL,"
      SQL = SQL & " [QTDE_SAIDA] [float] NOT NULL,"
      SQL = SQL & " [QTDE_ENTRADA] [float] NOT NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

'=============SAÍDA
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.ESTABELECIMENTO_ID, "
   SQL = SQL & " PEDIDO.VENDEDOR_ID, PEDIDO.CGCCPF, PEDIDO.DT_REQ, PEDIDO.STATUS, "
   SQL = SQL & " PEDIDO.NOME_CLIENTE, PEDIDO.PREFIXO, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, "
   SQL = SQL & " PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, Produto.CODG_NCM"
   SQL = SQL & " from PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where pedido.status in (7,5,3,8)"

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_req <= '" & txtDtFim.Text & "'"
   End If

   If Trim(cmbSituacaoAUX.Text) <> "" Then _
      SQL = SQL & " and status = " & Trim(cmbSituacaoAUX.Text)

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and produto_id = " & PRODUTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   NUMR_ID_N = 0 & MAX_ID("FATURAPRODUTO_ID", "FATURAPRODUTO", "", "", "", "")

   While Not TabTemp.EOF
      DoEvents
      lblConta.Caption = TabTemp.Fields("produto_id").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select * from FATURAPRODUTO "
      SQL = SQL & " where FATURAPRODUTO_ID = " & NUMR_ID_N
      SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         SQL = "insert into FATURAPRODUTO "
            SQL = SQL & " (FATURAPRODUTO_ID,PRODUTO_ID,DT_INI,DT_FIM,QTDE_SAIDA,QTDE_ENTRADA)"
         SQL = SQL & " values("
            SQL = SQL & NUMR_ID_N
            SQL = SQL & "," & TabTemp.Fields("produto_id").Value
            SQL = SQL & ",'" & txtDtIni.Text & "'"
            SQL = SQL & ",'" & txtDtFim.Text & "'"
            SQL = SQL & ",'" & tpMOEDA(TabTemp.Fields("QTD_PEDIDA").Value) & "'"
            SQL = SQL & ",'" & tpMOEDA(0) & "'"
         SQL = SQL & ")"
         Else
            SQL = "update FATURAPRODUTO set "
               SQL = SQL & " QTDE_SAIDA = QTDE_SAIDA + '" & tpMOEDA(TabTemp.Fields("QTD_PEDIDA").Value) & "'"
            SQL = SQL & " where FATURAPRODUTO_ID = " & NUMR_ID_N
            SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

'=============ENTRADA
   If NUMR_ID_N <= 0 Then _
      NUMR_ID_N = 0 & MAX_ID("FATURAPRODUTO_ID", "FATURAPRODUTO", "", "", "", "")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select NOTAENTRADA.ENTRADA_ID, NOTAENTRADA.FORNECEDOR_ID, NOTAENTRADA.TRANSP_ID, NOTAENTRADA.TIPOENTRADA_ID, "
   SQL = SQL & " NOTAENTRADA.NUMR_NOTA, NOTAENTRADA.SERIE_NOTA, NOTAENTRADA.DT_ENTRADA, NOTAENTRADA.DT_EMISSAO, "
   SQL = SQL & " NOTAENTRADA.STATUS, NOTAENTRADAITEM.SEQ_ID, NOTAENTRADAITEM.PRODUTO_ID, NOTAENTRADAITEM.PRECO_CUSTO,"
   SQL = SQL & " NOTAENTRADAITEM.QTDE_ENTRADA, NOTAENTRADAITEM.STATUS AS StatusNotaItem, NOTAENTRADAITEM.CFOP_ID, "
   SQL = SQL & " NOTAENTRADAITEM.NCM, NOTAENTRADAITEM.CST, NOTAENTRADAITEM.UN, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.SITUACAO, PRODUTO.CODG_NCM, PRODUTO.PRECO_CUSTO AS PrCustoProd, "
   SQL = SQL & " PRODUTO.PRECO_ATACADO,Produto.PRECO_Venda"
   SQL = SQL & " from NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
   SQL = SQL & " AND NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where NOTAENTRADA.status = 'E' "

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and DT_ENTRADA >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and DT_ENTRADA <= '" & txtDtFim.Text & "'"
   End If

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and produto_id = " & PRODUTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      DoEvents
      lblConta.Caption = TabTemp.Fields("produto_id").Value

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select * from FATURAPRODUTO "
      SQL = SQL & " where FATURAPRODUTO_ID = " & NUMR_ID_N
      SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         SQL = "insert into FATURAPRODUTO "
            SQL = SQL & " (FATURAPRODUTO_ID,PRODUTO_ID,DT_INI,DT_FIM,QTDE_SAIDA,QTDE_ENTRADA)"
         SQL = SQL & " values("
            SQL = SQL & NUMR_ID_N
            SQL = SQL & "," & TabTemp.Fields("produto_id").Value
            SQL = SQL & ",'" & txtDtIni.Text & "'"
            SQL = SQL & ",'" & txtDtFim.Text & "'"
            SQL = SQL & ",'" & tpMOEDA(TabTemp.Fields("QTDE_ENTRADA").Value) & "'"
            SQL = SQL & ",'" & tpMOEDA(0) & "'"
         SQL = SQL & ")"
         Else
            SQL = "update FATURAPRODUTO set "
               SQL = SQL & " QTDE_ENTRADA = QTDE_ENTRADA + '" & tpMOEDA(TabTemp.Fields("QTDE_ENTRADA").Value) & "'"
            SQL = SQL & " where FATURAPRODUTO_ID = " & NUMR_ID_N
            SQL = SQL & " and produto_id = " & TabTemp.Fields("produto_id").Value
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select FATURAPRODUTO.*, PRODUTO.EMPRESA_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.UNIDADE_MEDIDA, PRODUTO.CODG_BARRA, PRODUTO.SITUACAO, "
   SQL = SQL & " PRODUTO.SITUACAO_TRIBUTARIA, PRODUTO.ALIQUOTA_ICMS, PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, "
   SQL = SQL & " PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, PRODUTO.DT_CADASTRO, "
   SQL = SQL & " PRODUTO.DT_ULT_VENDA, PRODUTO.DT_ULT_COMPRA, PRODUTO.PESO_LIQUIDO, PRODUTO.PESO_BRUTO, "
   SQL = SQL & " PRODUTO.PRODUTO_BALANCA, FAMILIAPRODUTO.CODG_FAMILIA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia"
   SQL = SQL & " from FATURAPRODUTO "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON FATURAPRODUTO.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"
   SQL = SQL & " where FATURAPRODUTO_ID = " & NUMR_ID_N
   SQL = SQL & " order by descricao "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      QTDE_ENTRADA_N = 0 & TabTemp.Fields("QTDE_ENTRADA").Value
      QTDE_SAIDA_N = 0 & TabTemp.Fields("QTDE_SAIDA").Value

      Set item = lstItens.ListItems.Add(, "seq." & TabTemp.Fields("produto_id").Value, TabTemp.Fields("CODG_PRODUTO").Value)

      item.SubItems(1) = "" & TabTemp.Fields("DESCRICAO").Value
      item.SubItems(2) = "" & Format(QTDE_ENTRADA_N, strFormatacao2Digitos)
      item.SubItems(3) = "" & Format(QTDE_SAIDA_N, strFormatacao2Digitos)
      item.SubItems(4) = "" & Format(QTDE_ENTRADA_N - QTDE_SAIDA_N, strFormatacao2Digitos)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstItens.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub
