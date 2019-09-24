VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAGENDASERVICO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atendimentos Pendentes"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AGENDASERVICO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbSituacaoAUX 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbTipoOSAUX 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbSituacao 
      Height          =   360
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cmbTIPOOS 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstAgenda 
      Height          =   5385
      Left            =   45
      TabIndex        =   0
      Top             =   1200
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   9499
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2252
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CNPJ/CPF"
         Object.Width           =   4505
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   8259
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TipoVenda"
         Object.Width           =   2151
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Valor Pedido"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Desconto"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total Pedido"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dt.Emissão"
         Object.Width           =   3754
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Vendedor"
         Object.Width           =   4129
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tipo"
         Object.Width           =   1502
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Pedido Entrada"
         Object.Width           =   375
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "entrada_id"
         Object.Width           =   1
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "tipovenda_id"
         Object.Width           =   72
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   360
      Left            =   9480
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lblRep 
      AutoSize        =   -1  'True
      Caption         =   "Situação:"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo O.S.:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmAGENDASERVICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name

   SETA_GRID

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub


Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

'parametros combos x tabela descr
'8 = consultor tecnico
'9 = mecanico

   cmbTipoOSAUX.Clear
   cmbTipoOS.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'H' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipoOS.AddItem Trim(TabDESCR!DESCRICAO)
      cmbTipoOSAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbSituacao.Clear
   cmbSituacaoAUX.Clear

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'Z' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbSituacao.AddItem Trim(TabDESCR!DESCRICAO)
      cmbSituacaoAUX.AddItem Trim(TabDESCR.Fields("CODIGO").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

cmbSituacao.Text = "Aberta"
cmbSituacaoAUX.Text = 1

   cmbConsultorAUX.Clear
   cmbConsultor.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 8 "   'consultor tecnico
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
   SQL = SQL & " where tipo = 9 "   'mecanico
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbMecanicoAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbMecanico.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   'cmbVendedorAUX.Clear
   'cmbVendedor.Clear

   'If TabDESCR.State = 1 Then _
      TabDESCR.Close

   'SQL = "select vendedor_id, descricao from vwVendedor "
   'SQL = SQL & " where status = 'A' "   'vendedor
   'TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'While Not TabDESCR.EOF

   '   cmbVendedorAUX.AddItem TabDESCR.Fields("vendedor_id").Value
   '   cmbVendedor.AddItem Trim(TabDESCR.Fields("descricao").Value) & "-" & Trim(TabDESCR.Fields("vendedor_id").Value)

   '   TabDESCR.MoveNext
   'Wend


   cmbProduto.Clear
   cmbProdutoAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select OSPECA.PRODUTO_ID, PRODUTO.Descricao "
   SQL = SQL & " from OS "
   SQL = SQL & " INNER JOIN OSPECA "
   SQL = SQL & " ON OS.OS_ID = OSPECA.OS_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbProdutoAUX.AddItem TabDESCR.Fields("PRODUTO_id").Value
      cmbProduto.AddItem Trim(TabDESCR.Fields("DESCRICAO").Value) & "-" & Trim(TabDESCR.Fields("PRODUTO_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbServico.Clear
   cmbServicoAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select OSSERVICO.DESCRICAO, OSSERVICO.OSSERVICO_ID"
   SQL = SQL & " from OS "
   SQL = SQL & " INNER JOIN OSSERVICO "
   SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbServicoAUX.AddItem TabDESCR.Fields("OSSERVICO_id").Value
      cmbServico.AddItem Trim(TabDESCR.Fields("DESCRICAO").Value) & "-" & Trim(TabDESCR.Fields("OSSERVICO_id").Value)

      TabDESCR.MoveNext
   Wend

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim TabOS      As New ADODB.Recordset
   Dim TabServico As New ADODB.Recordset
   Dim TabPeca    As New ADODB.Recordset
   Dim DtFecha_D  As Date

   lstAgenda.ListItems.Clear

   CONT_N = 0
   SQL3 = "Endereço"
   QTD_COTAS = 0

   If TabOS.State = 1 Then _
      TabOS.Close

   SQL = "select * from vwOSServico "
   SQL = SQL & " where os_id > 0 "

   'If Trim(txtEqp.Text) <> "" Then _
      SQL = SQL & " and equipamento_id = " & txtEqp.Text

   'If Trim(txtOS.Text) <> "" Then _
      If IsNumeric(txtOS.Text) Then _
         SQL = SQL & " and os_id = " & txtOS.Text

   'If Trim(txtCHASSI.Text) <> "" Then _
      SQL = SQL & " and chassi = '" & Trim(txtCHASSI.Text) & "'"

   'If Trim(cmbTipoOSAUX.Text) <> "" Then _
      If IsNumeric(cmbTipoOSAUX.Text) Then _
         SQL = SQL & " and tipo_os = " & Trim(cmbTipoOSAUX.Text)

   'If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao_os = " & cmbSituacaoAUX.Text

   If Trim(cmbConsultorAUX.Text) <> "" Then _
      If IsNumeric(cmbConsultorAUX.Text) Then _
         SQL = SQL & " and ct_id = " & cmbConsultorAUX.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optDtAbertura.Value = True Then
         SQL = SQL & " and dt_os >= '" & (txtDtIni.Text) & "'"
         SQL = SQL & " and dt_os <= '" & (txtDtFim.Text) & "'"
      End If
      If optDtFechamento.Value = True Then
         SQL = SQL & " and dt_fecha >= '" & (txtDtIni.Text) & "'"
         SQL = SQL & " and dt_fecha <= '" & (txtDtFim.Text) & "'"
      End If
   End If

   SQL = SQL & " order by OS_ID desc"

   TabOS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If TabOS.EOF Then _
      MsgBox "Não existe O.S. para essa pesquisa."

   While Not TabOS.EOF
      If PEDIDO_ID_N <> TabOS.Fields("OS_ID").Value Then
         '====================================cliente
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = TabOS.Fields("cnpjcpf").Value
         If Len(Trim(TabOS.Fields("cnpjcpf").Value)) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If

         txtCNPJCPF.Text = "" & TabOS.Fields("cnpjcpf").Value
         NOME_A = "" & Trim(TabOS.Fields("cliente").Value)
         PESSOA_ID_N = 0 & TabOS.Fields("pessoa_id").Value

         '======================

         txtCNPJCPF.PromptInclude = True
'==============
         '========================      'totais serviço para cabeça do grid
         VALOR_TOTAL_SERVICO_N = 0
         If TabServico.State = 1 Then _
            TabServico.Close

         SQL = "select sum(valor_servico-desconto_servico) from OSSERVICO "
         SQL = SQL & " where os_id = " & TabOS.Fields("os_id").Value
         TabServico.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabServico.EOF Then _
            VALOR_TOTAL_SERVICO_N = 0 & TabServico.Fields(0).Value
         If TabServico.State = 1 Then _
            TabServico.Close

         'totais produto para cabeça do grid
         VALOR_TOTAL_PRODUTO_N = 0
         If TabPeca.State = 1 Then _
            TabPeca.Close
   
         SQL = "select sum((valor_item-desconto_produto) * qtde) from OSPECA "
         SQL = SQL & " where os_id = " & TabOS.Fields("os_id").Value
         TabPeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPeca.EOF Then _
            VALOR_TOTAL_PRODUTO_N = 0 & TabPeca.Fields(0).Value
         If TabPeca.State = 1 Then _
            TabPeca.Close
'======================
         SqL2 = "" & TRAZ_DESCRITOR("Z", TabOS.Fields("situacao_os").Value)

         PEDIDO_ID_N = TabOS.Fields("OS_ID").Value

         QTD_COTAS = QTD_COTAS + 1
         Set item = lstPedidos.ListItems.Add(, "seq." & Trim(TabOS.Fields("OS_ID").Value), Trim(TabOS.Fields("OS_ID").Value))

         item.SubItems(1) = "" & Format(VALOR_TOTAL_PRODUTO_N + VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos)
         txtCNPJCPF.PromptInclude = True
         item.SubItems(2) = "" & Trim(txtCNPJCPF.Text) & " - " & Trim(NOME_A)
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
      End If
      TabOS.MoveNext
   Wend
   lstAgenda.Refresh
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

