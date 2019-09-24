VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPedidoLibera 
   Caption         =   "Consulta Liberação Pedido"
   ClientHeight    =   7695
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
   Icon            =   "PedidoLibera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbFormaAUX 
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
      Left            =   1680
      TabIndex        =   17
      Top             =   1320
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
      Left            =   8640
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   735
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
      Left            =   8460
      TabIndex        =   3
      Top             =   840
      Width           =   3225
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
      Left            =   1500
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
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
      Left            =   8460
      MaxLength       =   100
      TabIndex        =   8
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5940
      MaxLength       =   6
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsCli 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   7980
      Picture         =   "PedidoLibera.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   405
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
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
            Caption         =   "Imprimir"
            Key             =   "pedido"
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
               Picture         =   "PedidoLibera.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoLibera.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoLibera.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoLibera.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoLibera.frx":A8FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoLibera.frx":C8DF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   360
      Left            =   5940
      TabIndex        =   5
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
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   300
      Left            =   1140
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Height          =   300
      Left            =   3540
      TabIndex        =   1
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3015
      Left            =   0
      TabIndex        =   15
      Top             =   1920
      Width           =   11850
      _ExtentX        =   20902
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cupom"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NFe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr.Venda"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Vlr.Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dt.Emisão"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Faturamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ID"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "NºCaixa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "TpRegistro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Comanda"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      Height          =   240
      Left            =   7365
      TabIndex        =   14
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   240
      Left            =   5040
      TabIndex        =   12
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   240
      Left            =   5160
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Inicial:"
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Final:"
      Height          =   240
      Left            =   2700
      TabIndex        =   9
      Top             =   840
      Width           =   795
   End
End
Attribute VB_Name = "frmPEDIDOLIBERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CRIA_VW
   CARREGA_COMBOS
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         MONTA_CONSULTA_SQL
         SETA_GRID
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         CRITERIO = ""
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

Sub CRIA_VW()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False
   If ExisteTabela("RETAGUARDA", "vwPEDIDO_LIBERACAO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwPEDIDO_LIBERACAO]"

   SQL = "CREATE VIEW [dbo].[vwPEDIDO_LIBERACAO] AS "
   SQL = SQL & " SELECT PEDIDO.PEDIDO_ID, PEDIDO.VENDEDOR_ID, VENDEDOR.NOME_VEND, "
   SQL = SQL & " PEDIDO.TIPOVENDA_ID, PEDIDO.USUARIO_ID, PEDIDO.ESTABELECIMENTO_ID,"
   SQL = SQL & " PEDIDO.TABELAPRECO_ID, PEDIDO.CGCCPF, PEDIDO.DT_REQ, PEDIDO.STATUS, "
   SQL = SQL & " PEDIDO.USUARIO_LIBERA_VENDA, PEDIDO.VALOR_DESCONTO, PEDIDO.PERC_DESC, "
   SQL = SQL & " PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_RECEBIDO, PEDIDO.VALOR_TOTAL, "
   SQL = SQL & " USUARIO_Venda.NOME as RespVenda, USUARIO_Libera.NOME AS RespLibVenda,"
   SQL = SQL & " TABELAPRECO.CODG_TABELA, TABELAPRECO.DESCRICAO"

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN USUARIO AS USUARIO_Venda "
   SQL = SQL & " ON PEDIDO.USUARIO_ID = USUARIO_Venda.USUARIO_ID "
   SQL = SQL & " INNER JOIN USUARIO AS USUARIO_Libera "
   SQL = SQL & " ON PEDIDO.USUARIO_LIBERA_VENDA = USUARIO_Libera.USUARIO_ID "
   SQL = SQL & " INNER JOIN VENDEDOR "
   SQL = SQL & " ON PEDIDO.VENDEDOR_ID = VENDEDOR.VENDEDOR_ID "
   SQL = SQL & " INNER JOIN TABELAPRECO "
   SQL = SQL & " ON VENDEDOR.TABELAPRECO_ID = TABELAPRECO.TABELAPRECO_ID"

   Me.Enabled = True
   Me.KeyPreview = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False

   cmbForma.Clear
   cmbAuxForma.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbForma.AddItem TabTemp!DESCRICAO & " - " & TabTemp!TipoVenda_ID
      cmbAuxForma.AddItem TabTemp!TipoVenda_ID
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbVend.Clear
   SQL = "select vendedor_id,nome_vend from VENDEDOR WITH (NOLOCK)"
   SQL = SQL & " order by nome_vend "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbVend.AddItem Trim(TabTemp!NOME_VEND) & " - " & Trim(TabTemp!vendedor_id)
      cmbVendAux.AddItem Trim(TabTemp!vendedor_id)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA
set
   Me.Enabled = False
   CHECA_ULTIMO_DIA_MES

   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False

   VALOR_TOTAL_N = 0
   If Indr_Consulta = True Then
      txtTotalVenda.Text = ""
      txtReg.Text = ""
      txtQtdeProd.Text = ""
   End If
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SqL2 = "SELECT  * "
   SQL3 = "SELECT count(vwPEDIDO_LIBERACAO.cliente_ID) "

   SQL = " FROM vwPEDIDO_LIBERACAO WITH (NOLOCK) "

   SQL = SQL & " where pedido_id Is Not Null"
   SQL = SQL & " and estabelecimento_id = " & cmbEstabAUX.Text

   If Trim(cmbCPUaux.Text) <> "" Then _
      If IsNumeric(cmbCPUaux.Text) Then _
         SQL = SQL & " and numero_caixa_cpu = " & cmbCPUaux.Text

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and produto_id = " & PRODUTO_ID_N

   If Trim(txtCupom.Text) <> "" Then _
      SQL = SQL & " and numr_cupom = " & txtCupom.Text

   If Trim(txtNOTA.Text) <> "" Then _
      SQL = SQL & " and numr_nota = " & txtNOTA.Text

   If Trim(txtPedido.Text) <> "" Then _
      SQL = SQL & " and pedido_id = " & txtPedido.Text

   txtCGCCPF.PromptInclude = False
   If Trim(txtCGCCPF.Text) <> "" Then _
      If CLIENTE_ID_N > 0 Then _
         SQL = SQL & " and cliente_id = " & CLIENTE_ID_N
         'SQL = SQL & " and cgccpf = '" & Trim(txtCGCCPF.Text) & "'"
   txtCGCCPF.PromptInclude = True

   If Trim(cmbVend.Text) <> "" Then _
      SQL = SQL & " and vendedor_id = " & cmbVendAux.Text

   If Trim(cmbSituacaoAUX.Text) <> "" Then
      If Trim(cmbSituacaoAUX.Text) = "'7','5','3'" Then _
         SQL = SQL & " and numr_nota > 0 "

      SQL = SQL & " and SIT_PEDIDO in (" & Trim(cmbSituacaoAUX.Text) & ")"
   End If

   If Trim(cmbAuxForma.Text) <> "" Then _
      SQL = SQL & " and tipovenda_id = " & cmbAuxForma.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & Format(txtDtIni.Text, "dd/mm/yyyy") & "'"
      SQL = SQL & " and dt_req <= '" & Format(txtDtFim.Text, "dd/mm/yyyy") & "'"
   End If

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   If Trim(txtComanda.Text) <> "" Then _
      If IsNumeric(txtComanda.Text) Then _
         SQL = SQL & " and cartaobarra_id = " & txtComanda.Text

   SQL3 = SQL3 & " " & SQL

   SQL = SQL & " order by PEDIDO_ID desc"

   SQL = SqL2 & " " & SQL

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

   Me.Enabled = True
   Me.KeyPreview = True
Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False
   HORA_INI = Time

   Dim TabTemp          As New ADODB.Recordset
   Dim Conta_Produto_N  As Long

   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0

   Me.Enabled = False
   Me.KeyPreview = False

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
'MsgBox SQL3
   TabTemp.Open SQL3, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then _
      CONTA_REG_PROGRESSO = TabTemp.Fields(0).Value
'============================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If CONTA_REG_PROGRESSO > 0 Then
      ProgressBar1.Min = 0                   'Indica o valor inicial
      ProgressBar1.Max = CONTA_REG_PROGRESSO 'Indica o valor final
      'frmProgresso.Show 1
   End If
   CONT_N = 0

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then
      While Not TabTemp.EOF
         DoEvents

         If CONT_N < CONTA_REG_PROGRESSO Then
            CONT_N = CONT_N + 1
            ProgressBar1.Value = CONT_N
         End If

         If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
            CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
            txtReg.Text = CONTA_REGISTRO_N
            txtReg.Refresh

            PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value

            NUMR_SEQ_N = NUMR_SEQ_N + 1
            Set Item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp.Fields("PEDIDO_ID").Value)

            Item.SubItems(11) = "" & TabTemp.Fields("PEDIDO_ID").Value
            Item.SubItems(1) = "" & TabTemp.Fields("numr_cupom").Value
            Item.SubItems(2) = "" & TabTemp.Fields("numr_nota").Value
            Item.SubItems(3) = "" & Trim(TabTemp!NOME_CLIENTE) & " - " & Trim(TabTemp.Fields("CNPJCPF").Value)

            If IsNull(TabTemp!NOME_CLIENTE) Or Trim(TabTemp!NOME_CLIENTE) = "" Then
               If TabCliente.State = 1 Then _
                  TabCliente.Close

               SQL = "select nome from CLIENTE WITH (NOLOCK)"
               SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"
               TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCliente.EOF Then _
                  Item.SubItems(3) = "" & TabCliente!NOME

               If TabCliente.State = 1 Then _
                  TabCliente.Close
            End If

            Item.SubItems(7) = TabTemp!DT_REQ
            Item.SubItems(8) = ""

            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            SQL = "select * from TIPOVENDA WITH (NOLOCK)"
            SQL = SQL & " where tipovenda_id = " & TabTemp!TipoVenda_ID
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then _
               Item.SubItems(8) = TabDESCR!DESCRICAO
            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            If TabUSU.State = 1 Then _
               TabUSU.Close

            Item.SubItems(9) = ""
   
            SQL = "select * from VENDEDOR WITH (NOLOCK)"
            SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id").Value
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               Item.SubItems(9) = TabUSU!NOME_VEND
            If TabUSU.State = 1 Then _
               TabUSU.Close

            Item.SubItems(10) = ""

            If Not IsNull(TabTemp.Fields("Status")) Then
               If TabTemp.Fields("Status") = 2 Then
                  If TabTemp.Fields("tipo_registro") = "O" Then
                     Item.SubItems(10) = "Orcamento"
                     Else: Item.SubItems(10) = "Pedido"
                  End If
               End If
               If TabTemp.Fields("Status").Value = 3 Then _
                  Item.SubItems(10) = "3-Faturado"
               If TabTemp.Fields("Status").Value = 4 Then _
                  Item.SubItems(10) = "4-Cupom"
               If TabTemp.Fields("Status").Value = 5 Then _
                  Item.SubItems(10) = "5-Faturado"
               If TabTemp.Fields("Status").Value = 7 Then _
                  Item.SubItems(10) = "7-Cupom Fiscal"
               If TabTemp.Fields("Status").Value = 9 Then _
                  Item.SubItems(10) = "9-Cancelado"
            End If

            If Not IsNull(TabTemp.Fields("numero_caixa_cpu").Value) Then _
               Item.SubItems(12) = TabTemp.Fields("numero_caixa_cpu").Value

            Item.SubItems(13) = TabTemp.Fields("tipo_registro").Value

            VALOR_DESCONTO_N = 0
            'VALOR_TOTAL_DESCONTO_N = 0

            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPedidoItem.EOF Then _
               If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
                  VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

           If Not IsNull(TabTemp.Fields("desccabeca").Value) Then _
              VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabTemp.Fields("desccabeca").Value

            VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

            'BUSCA VALOR TOTAL VENDA
            VALOR_ITEM_N = 0

            SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            'SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            SQL = "select sum(qtd_pedida) FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               Conta_Produto_N = Conta_Produto_N + TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            txtQtdeProd.Text = Conta_Produto_N
            txtQtdeProd.Refresh
            'VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N
            VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

            txtTotDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
            txtTotVendas.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
            txtTotalVenda.Text = Format(VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

            Item.SubItems(4) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
            Item.SubItems(5) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
            Item.SubItems(6) = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)

            If TabTemp.Fields("SIT_PEDIDO").Value = 1 Then
               Item.ForeColor = vbRed
               Item.ListSubItems(1).ForeColor = vbRed
               Item.ListSubItems(2).ForeColor = vbRed
               Item.ListSubItems(3).ForeColor = vbRed
               Item.ListSubItems(4).ForeColor = vbRed
               Item.ListSubItems(5).ForeColor = vbRed
               Item.ListSubItems(6).ForeColor = vbRed
               Item.ListSubItems(7).ForeColor = vbRed
               Item.ListSubItems(8).ForeColor = vbRed
               Item.SubItems(10) = "" & "Em Aberto - 1"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 2 Then
               Item.ForeColor = vbBlue
               Item.ListSubItems(1).ForeColor = vbBlue
               Item.ListSubItems(2).ForeColor = vbBlue
               Item.ListSubItems(3).ForeColor = vbBlue
               Item.ListSubItems(4).ForeColor = vbBlue
               Item.ListSubItems(5).ForeColor = vbBlue
               Item.ListSubItems(6).ForeColor = vbBlue
               Item.ListSubItems(7).ForeColor = vbBlue
               Item.ListSubItems(8).ForeColor = vbBlue
               Item.ListSubItems(9).ForeColor = vbBlue
               Item.SubItems(10) = "" & "A Faturar - 2"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 3 Then
               Item.ForeColor = vbBlack
               Item.ListSubItems(1).ForeColor = vbBlack
               Item.ListSubItems(2).ForeColor = vbBlack
               Item.ListSubItems(3).ForeColor = vbBlack
               Item.ListSubItems(4).ForeColor = vbBlack
               Item.ListSubItems(5).ForeColor = vbBlack
               Item.ListSubItems(6).ForeColor = vbBlack
               Item.ListSubItems(7).ForeColor = vbBlack
               Item.ListSubItems(8).ForeColor = vbBlack
               Item.ListSubItems(9).ForeColor = vbBlack
               Item.SubItems(10) = "" & "Faturado - 3"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 5 Then
               Item.ForeColor = vbBlack
               Item.ListSubItems(1).ForeColor = vbBlack
               Item.ListSubItems(2).ForeColor = vbBlack
               Item.ListSubItems(3).ForeColor = vbBlack
               Item.ListSubItems(4).ForeColor = vbBlack
               Item.ListSubItems(5).ForeColor = vbBlack
               Item.ListSubItems(6).ForeColor = vbBlack
               Item.ListSubItems(7).ForeColor = vbBlack
               Item.ListSubItems(8).ForeColor = vbBlack
               Item.ListSubItems(9).ForeColor = vbBlack
               Item.SubItems(10) = "" & "Faturado - 5"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 6 Then
               Item.ForeColor = vbBlack
               'Item.ListSubItems(1).ForeColor = vbYellow
               'Item.ListSubItems(2).ForeColor = vbYellow
               'Item.ListSubItems(3).ForeColor = vbYellow
               'Item.ListSubItems(4).ForeColor = vbYellow
               'Item.ListSubItems(5).ForeColor = vbYellow
               'Item.ListSubItems(6).ForeColor = vbYellow
               'Item.ListSubItems(7).ForeColor = vbYellow
               'Item.ListSubItems(8).ForeColor = vbYellow
               Item.ListSubItems(10).ForeColor = vbYellow
               Item.SubItems(10) = "" & "Não Contabilizado"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 7 Then
               Item.ForeColor = vbMagenta
               Item.ListSubItems(1).ForeColor = vbMagenta
               Item.ListSubItems(2).ForeColor = vbMagenta
               Item.ListSubItems(3).ForeColor = vbMagenta
               Item.ListSubItems(4).ForeColor = vbMagenta
               Item.ListSubItems(5).ForeColor = vbMagenta
               Item.ListSubItems(6).ForeColor = vbMagenta
               Item.ListSubItems(7).ForeColor = vbMagenta
               Item.ListSubItems(8).ForeColor = vbMagenta
               Item.ListSubItems(9).ForeColor = vbMagenta
               Item.SubItems(10) = "" & "Cupom Fiscal - 7"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 9 Then
               'Item.ForeColor = &HC0E0FF '&HC0C0C0
               'Item.ListSubItems(1).ForeColor = &HC0E0FF
               'Item.ListSubItems(2).ForeColor = &HC0E0FF
               'Item.ListSubItems(3).ForeColor = &HC0E0FF
               'Item.ListSubItems(4).ForeColor = &HC0E0FF
               'Item.ListSubItems(5).ForeColor = &HC0E0FF
               'Item.ListSubItems(6).ForeColor = &HC0E0FF
               'Item.ListSubItems(7).ForeColor = &HC0E0FF
               Item.ListSubItems(8).ForeColor = &HC0E0FF
               Item.ListSubItems(9).ForeColor = &HC0E0FF
               Item.ListSubItems(10).ForeColor = &HC0E0FF
               Item.SubItems(10) = "" & "Cancelado - 9"
            End If
         End If

         Item.ListSubItems(2).ForeColor = vbRed
         Item.ListSubItems(1).ForeColor = vbRed

'verificando se é venda com comanda eletronica
         Item.SubItems(14) = "" & TabTemp.Fields("cartaobarra_id").Value

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select cartaobarra_id from PEDIDOTEMP WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("PEDIDO_ID").Value
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            'Item.ForeColor = vbBlack
            'Item.ListSubItems(1).ForeColor = vbBlack
            'Item.ListSubItems(2).ForeColor = vbBlack
            'Item.ListSubItems(3).ForeColor = vbBlack
            'Item.ListSubItems(4).ForeColor = vbBlack
            'Item.ListSubItems(5).ForeColor = vbBlack
            'Item.ListSubItems(6).ForeColor = vbBlack
            'Item.ListSubItems(7).ForeColor = vbBlack
            Item.ListSubItems(8).ForeColor = vbBlack
            Item.ListSubItems(9).ForeColor = vbBlack
            Item.SubItems(14) = "" & TabDESCR.Fields(0).Value
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         TabTemp.MoveNext
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

   Me.Enabled = True
   Me.KeyPreview = True
Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
