VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEmissorNFCe 
   Caption         =   "Emissor NFCe"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   Icon            =   "NFCe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   114.565
   ScaleMode       =   0  'User
   ScaleWidth      =   210.873
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView LISTAITEM 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   5265
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648384
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2252
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   15930
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "QtdeVendida"
         Object.Width           =   2389
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valr.Unitário"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Desconto"
         Object.Width           =   1877
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valr.Total"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ST_Item"
         Object.Width           =   573
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedidos 
      Height          =   5385
      Left            =   0
      TabIndex        =   1
      Top             =   0
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
      DesignWidth     =   11955
      DesignHeight    =   6495
   End
   Begin MSComctlLib.ListView lstTotais 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1296
      View            =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1434
      EndProperty
   End
   Begin MSComctlLib.StatusBar barRodape 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6120
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "NFCe.frx":5C12
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmEmissorNFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================
Private Declare Function Bematech_FI_AdicionaInformacoesCombustivel Lib "BEMAFI32.DLL" (ByVal itemIndex As String, ANPProductCode As String, ByVal percentMixGN As String, ByVal CODIF As String, ByVal quantity As String, ByVal consumeUF As String, ByVal BCProductCIDE As String, ByVal taxProductCIDE As String, ByVal valueCIDE As String, ByVal fuelNozzleNumber As String, ByVal fuelPumpNumber As String, ByVal fuelTankNumber As String, ByVal fuelGaugeInitial As String, ByVal fuelGaugeFinal As String) As Integer
Private Declare Function Bematech_FI_ChaveAcessoNFCe Lib "BEMAFI32.DLL" (ByVal Index As String, ByVal counter As String, ByRef accessKey As String) As Integer
Private Declare Function Bematech_FI_DadosConsumidorNFCe Lib "BEMAFI32.DLL" (ByVal CPF As String, ByVal name As String, ByVal address As String, ByVal complement As String, ByVal number As String, ByVal neighborhood As String, ByVal IBGECode As String, ByVal city As String, ByVal UF As String, ByVal CEP As String, ByVal countyCode As String, ByVal country As String, ByVal phone As String, ByVal stateRegistrationIndicator As String, ByVal stateRegistration As String, ByVal SUFRAMACode As String, ByVal email As String) As Integer
Private Declare Function Bematech_FI_DadosEnvioNFCe Lib "BEMAFI32.DLL" (ByVal TipoLayout As String, ByVal TipoEmissao As String, ByVal cEmail As String) As Integer
Private Declare Function Bematech_FI_EfetuaFormaPagamentoNFCeEx Lib "BEMAFI32.DLL" (ByVal descBandeira As String, ByVal ValorForma As String, ByVal CNPJCrede As String, ByVal bandeira As String, ByVal CodAuto As String, ByVal CodIntegra As String) As Integer
Private Declare Function Bematech_FI_NumeroNotaNFCe Lib "BEMAFI32.DLL" (ByVal noteNumber As String) As Integer
Private Declare Function Bematech_FI_NumeroSerieNFCe Lib "BEMAFI32.DLL" (ByVal serialNumber As String) As Integer
Private Declare Function Bematech_FI_ProgramaContadorNFCe Lib "BEMAFI32.DLL" (ByVal Index As String, ByVal counter As String) As Integer
Private Declare Function Bematech_FI_ProtocoloUltimaNFCe Lib "BEMAFI32.DLL" (ByVal protocol As String, ByVal datehour As String) As Integer
Private Declare Function Bematech_FI_RetornaInformacoesNFCe Lib "BEMAFI32.DLL" (ByVal paramType As String, ByVal paramValue As String, ByVal retChaveAcesso As String, ByVal retSerie As String, ByVal retNumNFCe As String, ByVal retCancelled As String, ByVal retSendStatus As String, ByVal retSendProtocol As String, ByVal retSendProtocolDatetime As String, ByVal retCancellationStatus As String, ByVal retCancellationProtocol As String) As Integer
Private Declare Function Bematech_FI_StatusUltimaNFCe Lib "BEMAFI32.DLL" (ByVal Status As String) As Integer
Private Declare Function Bematech_FI_StatusUltimoCancelamentoNFCe Lib "BEMAFI32.DLL" (ByVal Status As String) As Integer
Private Declare Function Bematech_FI_UltimaChaveAcessoNFCe Lib "BEMAFI32.DLL" (ByVal accessKey As String) As Integer
Private Declare Function Bematech_FI_VendeItemCompleto Lib "BEMAFI32.DLL" (ByVal sParametros As String) As Integer
Private Declare Function Bematech_FI_VendeItemCompletoJSON Lib "BEMAFI32.DLL" (ByVal sParametros As String) As Integer
Private Declare Function Bematech_FI_TerminaFechamentoCupomNFCe Lib "BEMAFI32.DLL" (ByVal Mensagem As String, ByVal Taxas As String) As Integer

Option Explicit
   Dim VALOR_DESCONTO_CABECA_N   As Double
   Dim NUMEROCUPOM               As String
   Dim NOME_VENDEDOR             As String
   Dim NUMEROCUPOMCancelado      As String
   Dim NUMR_CUPOM_ABERTO         As Long
   Dim INDR_PERGUNTA             As Boolean
   Dim Mensagem_Final            As String
   Dim NOME_CLI                  As String
   Dim CONTA_TENTATIVA           As Long
   Dim TOTAL_DESCONTO_N          As Double
   Dim Parametros                As Variant
   Dim OperacaoECFOK             As Boolean
   Dim ITEM_DESCONTO_N           As Double
   Dim Descr_Forma_Pagto         As String
   Dim VALOR_TOTAL_IMPOSTO       As Double
   Dim LocalRetorno              As String
   Dim CFOP_A                    As String
   Dim DESCRICAO_CFOP_A          As String

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.name

   PESQUISA_VENDA

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF3
         IMPRIME_TELA
      Case vbKeyF6
         If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
            frmPedidoCancela.txtPedido.Text = 0 & lstPedidos.SelectedItem.ListSubItems.item(10).Text
            frmPedidoCancela.Show 1
            CRITERIO = ""
            Else: MsgBox "Não permitido."
         End If

         PESQUISA_VENDA
      Case vbKeyF7
         If Not IsNull(lstPedidos.SelectedItem.Text) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cgccpf from PEDIDO "
            SQL = SQL & " where pedido_id = " & lstPedidos.SelectedItem.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp!CGCCPF) Then
                  txtCNPJCPF.PromptInclude = False
                     txtCNPJCPF.Text = TabTemp!CGCCPF
                  txtCNPJCPF.PromptInclude = True
               End If
            End If

            If TabTemp.State = 1 Then _
               TabTemp.Close

            LISTAITEM.ListItems.Clear

            SQL = "SELECT PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.valor_desconto,"
            SQL = SQL & " produto.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
            SQL = SQL & " descricao, situacao_tributaria"

            SQL = SQL & " from PEDIDO WITH (NOLOCK) "
            SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
            SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
            SQL = SQL & " INNER JOIN PRODUTO "
            SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

            SQL = SQL & " where pedido.pedido_id = " & lstPedidos.SelectedItem.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            'SQL = SQL & " and I.tipo_reg = 'PC' "
            'SQL = SQL & " and pedidoitem.status <> 'C' "
            SQL = SQL & " order by descricao"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               MOSTRA_RODAPE_AQUI "Duplo Click no grid ocultar", " ", " ", " ", ""
               LISTAITEM.Visible = True
            End If
            While Not TabTemp.EOF
               Set item = LISTAITEM.ListItems.Add(, "seq." & TabTemp.Fields("seq_id"), Trim(TabTemp.Fields("codg_produto").Value))
               item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
               item.SubItems(2) = "" & Format(TabTemp.Fields("qtd_pedida").Value, strFormatacao3Digitos)
               item.SubItems(3) = "" & Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
               item.SubItems(4) = "" & Format(TabTemp.Fields("valor_desconto").Value, strFormatacao2Digitos)
               item.SubItems(5) = "" & Format((TabTemp.Fields("valor_item").Value - TabTemp.Fields("valor_desconto").Value) * TabTemp.Fields("qtd_pedida").Value, strFormatacao2Digitos)
               item.SubItems(6) = "" & Trim(TabTemp.Fields("situacao_tributaria").Value)

               If Trim(TabTemp.Fields("situacao_tributaria").Value) = "A" Then
                  item.ForeColor = vbBlue
                  item.ListSubItems(1).ForeColor = vbBlue
                  item.ListSubItems(2).ForeColor = vbBlue
                  item.ListSubItems(3).ForeColor = vbBlue
                  item.ListSubItems(4).ForeColor = vbBlue
                  item.ListSubItems(5).ForeColor = vbBlue
                  item.ListSubItems(6).ForeColor = vbBlue
                  Else
                     If Trim(TabTemp.Fields("situacao_tributaria").Value) = "P" Then
                        item.ForeColor = vbRed
                        item.ListSubItems(1).ForeColor = vbRed
                        item.ListSubItems(2).ForeColor = vbRed
                        item.ListSubItems(3).ForeColor = vbRed
                        item.ListSubItems(4).ForeColor = vbRed
                        item.ListSubItems(5).ForeColor = vbRed
                        item.ListSubItems(6).ForeColor = vbRed
                        Else
                           If Trim(TabTemp.Fields("situacao_tributaria").Value) = "B" Then
                              item.ForeColor = vbMagenta
                              item.ListSubItems(1).ForeColor = vbMagenta
                              item.ListSubItems(2).ForeColor = vbMagenta
                              item.ListSubItems(3).ForeColor = vbMagenta
                              item.ListSubItems(4).ForeColor = vbMagenta
                              item.ListSubItems(5).ForeColor = vbMagenta
                              item.ListSubItems(6).ForeColor = vbMagenta
                           End If
                     End If
               End If

               TabTemp.MoveNext
               CRITERIO = ""
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close

            LISTAITEM.Refresh
         End If
      Case vbKeyF9
         PESQUISA_VENDA
      Case vbKeyF10
         Call lstPedidos_DblClick
      Case vbKeyF11
         FORMULA_REL = ""
         If Not IsNull(lstPedidos.SelectedItem.Text) Then
            FORMULA_REL = lstPedidos.SelectedItem.Text

            BlockInput False  'Desbloqueia o teclado
            If Not IsNumeric(FORMULA_REL) Then _
               Exit Sub

            PEDIDO_ID_N = FORMULA_REL

            FORMULA_REL = "{vwRelVenda.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & PEDIDO_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.statusitem} <> 'C' "

            'If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_pedido_venda.rpt"
            If CNPJ_EMPRESA_N = "15333554000188" Then _
               Nome_Relatorio = "pedido_shf.rpt"

            frmRELATORIO10.Show 1
         End If

      Case vbKeyEscape
         Unload Me
   End Select

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "Form_Unload"
End Sub

Private Sub lstPedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedidos, ColumnHeader
End Sub

Private Sub lstPedidos_Click()
'On Error GoTo ERRO_TRATA

   Me.Caption = "ESC-Sair " & "F6-Cancelar " & " F7-Ver Itens" & " F9-Atutalizar" & " F10-Recebimento | F11-Imprimir Pedido"
   MOSTRA_RODAPE_AQUI " ESC-Sair", "F6-Cancelar", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento | F11-Imprimir Pedido"

   If Not IsNull(lstPedidos.SelectedItem.Text) Then
      If Trim(lstPedidos.SelectedItem.Text) <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select cgccpf from PEDIDO "
         SQL = SQL & " where pedido_id = " & lstPedidos.SelectedItem.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp!CGCCPF) Then
               txtCNPJCPF.PromptInclude = False
                  txtCNPJCPF.Text = TabTemp!CGCCPF
               txtCNPJCPF.PromptInclude = True
            End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "lstPedidos_Click"
End Sub

Private Sub lstPedidos_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedidos.SelectedItem.ListSubItems.item(10).Text) Then
      If Trim(lstPedidos.SelectedItem.ListSubItems.item(10).Text) <> "" Then

         PEDIDO_ID_N = lstPedidos.SelectedItem.ListSubItems.item(10).Text

         '================================== PEDIDO DE VENDA
         If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("R") Or _
            UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("OS") Then
            TIPO_NFe_GERAR = "R"          'Tipo Saida
            FAZ_RECEBIMENTO
         End If
         If USA_NFe = True Then
            '================================== DEVOLUÇÃO DE ENTRADA
            If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("DC") Then
               TIPO_NFe_GERAR = "DC"

               If TabCABECA.State = 1 Then _
                  TabCABECA.Close

               SQL = "select * from PEDIDO "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCABECA.EOF Then
                  If TabCABECA!Status = 2 Then
                     CFOP_A = ""
                     DESCRICAO_CFOP_A = ""

                     Msg = "Processar Devolução de Compra NFe ?"
                     PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
                     If RESPOSTA = vbYes Then _
                        frmNOTAGERA.Show 1
                  End If
               End If
               If TabCABECA.State = 1 Then _
                  TabCABECA.Close
            End If
            '================================== DEVOLUÇÃO DE SAIDA
            If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("DV") Then
               TIPO_NFe_GERAR = "DV"          'DEVOLUÇÃO VENDA

               If TabCABECA.State = 1 Then _
                  TabCABECA.Close

               SQL = "select * from PEDIDO "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCABECA.EOF Then
                  If TabCABECA!Status = 2 Then
                     CFOP_A = ""
                     DESCRICAO_CFOP_A = ""

                     Msg = "Processar Devolução de Venda NFe ?"
                     PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
                     If RESPOSTA = vbYes Then _
                        frmNOTAGERA.Show 1
                  End If
               End If
               If TabCABECA.State = 1 Then _
                  TabCABECA.Close
            End If
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   '================================== DEVOLUÇÃO DE TRANSFERENCIA
   If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("T") Then _
      TIPO_NFe_GERAR = "T"

   PESQUISA_VENDA

   Me.Caption = "ESC-Sair " & "F6-Cancelar " & " F7-Ver Itens" & " F9-Atutalizar" & " F10-Recebimento | F11-Imprimir Pedido"
   MOSTRA_RODAPE_AQUI " ESC-Sair", "F6-Cancelar", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento | F11-Imprimir Pedido"

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "lstPedidos_DblClick"
End Sub

Private Sub LISTAITEM_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE_AQUI " ESC-Sair", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento", ""
   LISTAITEM.Visible = False

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "LISTAITEM_DblClick"
End Sub

Private Sub PESQUISA_VENDA()
'On Error GoTo ERRO_TRATA

   SETA_GRID

   MOSTRA_RODAPE_AQUI " ESC-Sair", "F6-Cancelar", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento | F11-Imprimir Pedido"

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "PESQUISA_VENDA"
End Sub

Private Sub IMPRIME_TELA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from PEDIDO "
   SQL = SQL & " where tipo_registro in ('S','R','D') "
   SQL = SQL & " and status in (2)" 'gerado somente Pedido"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by pedido_id DESC "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      FORMULA_REL = "{PEDIDO.status} =  2"
      FORMULA_REL = FORMULA_REL & " and {PEDIDO.tipo_registro} = ('R')"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_tela_nf.rpt"
   frmRELATORIO10.Show 1

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "IMPRIME_TELA"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstPedidos.ListItems.Clear
   NUMR_SEQ_N = 0
   NUMR_CONSULTA_N = 0

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PEDIDO "

   SQL = SQL & " where tipo_registro in ('S','R','DC','DV','OS') "
   SQL = SQL & " and status = 2" 'gerado somente Pedido"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by dt_req DESC "
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCABECA.EOF
      CRITERIO = ""

      txtCNPJCPF.PromptInclude = False
      If Len(Trim(TabCABECA.Fields("cgccpf").Value)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If

      txtCNPJCPF.Text = TabCABECA.Fields("cgccpf").Value
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = TabCABECA.Fields("cgccpf").Value

'========================================cliente
      If Not IsNull(TabCABECA.Fields("nome_cliente").Value) Then
         If Trim(TabCABECA.Fields("nome_cliente").Value) <> "" Then
            CRITERIO = Trim(TabCABECA!NOME_CLIENTE)
            Else: BUSCA_CLIENTE (TabCABECA.Fields("CLIENTE_ID").Value)
         End If
         Else: BUSCA_CLIENTE (TabCABECA.Fields("CLIENTE_ID").Value)
      End If
'========================================setando grid
      Set item = lstPedidos.ListItems.Add(, "seq." & Trim(TabCABECA.Fields("pedido_id").Value), Trim(TabCABECA.Fields("pedido_id").Value))

      item.SubItems(1) = "" & txtCNPJCPF.Text
      item.SubItems(2) = "" & CRITERIO

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select descricao from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & TabCABECA.Fields("tipovenda_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            item.SubItems(3) = "" & TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PERC_DESCONTO_N = 0 & TabCABECA.Fields("perc_desc").Value
      VALOR_DESCONTO_N = 0 & TabCABECA.Fields("valor_desconto").Value
      VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N

'========================================parceiro, tem que ver se pega pelo valor do desconto ou percentual
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
      SQL = SQL & " and pedidoitem.status <> 'C' "
      'SQL = SQL & " and tipo_reg = 'PC' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
'========================================

      VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

      'BUSCA VALOR TOTAL VENDA
      VALOR_ITEM_N = 0

      SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
      SQL = SQL & " and pedidoitem.status <> 'C' "
      'SQL = SQL & " and tipo_reg = 'PC' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
'========================================

      item.SubItems(4) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      item.SubItems(5) = "" & Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      item.SubItems(7) = "" & Trim(TabCABECA!dt_req)

'========================================
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select descricao from vwVendedor "
      SQL = SQL & " where vendedor_id = " & TabCABECA!VENDEDOR_ID
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then _
         item.SubItems(8) = "" & TabVENDEDOR!DESCRICAO
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
'========================================

      item.SubItems(9) = "" & TabCABECA!TIPO_REGISTRO
      item.SubItems(10) = "" & TabCABECA.Fields("pedido_id").Value
      item.SubItems(12) = "" & TabCABECA.Fields("tipovenda_id").Value

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      CONT_N = CONT_N + 1
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      If Trim(UCase(TabCABECA.Fields("tipo_registro").Value)) = "DV" Then
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
      If Trim(UCase(TabCABECA.Fields("tipo_registro").Value)) = "DC" Then
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
      End If
      If Trim(UCase(TabCABECA.Fields("tipo_registro").Value)) = "OS" Then
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

'========================================FUNCIONARIO
      If Trim(TabCABECA.Fields("cgccpf").Value) <> "99999999999" Then
         If TabVENDEDOR.State = 1 Then _
            TabVENDEDOR.Close
         SQL = "select usuario_id, funcionario from USUARIO "
         SQL = SQL & " where cpf = '" & Trim(TabCABECA.Fields("cgccpf").Value) & "'"
         TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVENDEDOR.EOF Then
            If Not IsNull(TabVENDEDOR.Fields(1).Value) Then
               If TabVENDEDOR.Fields(1).Value = True Then
                  item.ForeColor = vbGreen
                  item.ListSubItems(1).ForeColor = vbMagenta
                  item.ListSubItems(2).ForeColor = vbMagenta
                  item.ListSubItems(3).ForeColor = vbMagenta
                  item.ListSubItems(4).ForeColor = vbMagenta
                  item.ListSubItems(5).ForeColor = vbMagenta
                  item.ListSubItems(6).ForeColor = vbMagenta
                  item.ListSubItems(7).ForeColor = vbMagenta
                  item.ListSubItems(8).ForeColor = vbMagenta
                  item.ListSubItems(9).ForeColor = vbMagenta
               End If
            End If
         End If
         If TabVENDEDOR.State = 1 Then _
            TabVENDEDOR.Close
      End If

      TabCABECA.MoveNext
   Wend
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   lstPedidos.Refresh

   MOSTRA_TOTAIS

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "SETA_GRID"
End Sub

Private Sub CHECA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   STATUS_A = ""
   SQL = "select * from PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & lstPedidos.SelectedItem.Text
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      SP_PROCURA_PRODUTO EMPRESA_ID_N, Trim(TabPedidoItem!Codg_Produto), 0, "", FORNEC_ID_N, "", 1
      If Not TabProduto.EOF Then _
         QTDE_ESTOQUE = TabProduto!QTD 'Recebe so qtd. porque ja esta retido no pedido
      If TabProduto.State = 1 Then _
         TabProduto.Close

      If QTDE_ESTOQUE < TabPedidoItem!QTD_PEDIDA Then _
         STATUS_A = "V"

      TabPedidoItem.MoveNext
   Wend
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close
    
   If STATUS_A = "V" Then 'status de que  a Itens sem Quantidade!
      MsgBox "Pedido com Items Aquardando ordem de Producao , Impossivel Emitir nota!"
      LISTAITEM.Refresh
      Else: FAZ_RECEBIMENTO
   End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "CHECA_ESTOQUE"
End Sub

Private Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   Dim TabPedido As New ADODB.Recordset
   INDR_VENDA = False

   If Not IsNull(lstPedidos.SelectedItem.Text) Then
      If Trim(lstPedidos.SelectedItem.Text) <> "" Then
         PEDIDO_ID_N = lstPedidos.SelectedItem.Text
         INDR_RECEITA = 1

         If INDR_FORM_ABERTO = True Then
            Unload frmCADRECEBVENDA
            INDR_FORM_ABERTO = False
         End If
'===================================
         If Not IsNull(lstPedidos.SelectedItem.ListSubItems.item(12).Text) Then
            If Trim(lstPedidos.SelectedItem.ListSubItems.item(12).Text) <> "" Then
               If IsNumeric(lstPedidos.SelectedItem.ListSubItems.item(12).Text) Then

                  If TabTemp.State = 1 Then _
                     TabTemp.Close

                  SQL = "select contabiliza from TIPOVENDA "
                  SQL = SQL & " where tipovenda_id = " & lstPedidos.SelectedItem.ListSubItems.item(12).Text
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
                        If TabTemp.Fields("contabiliza").Value = True Then
                           If TabTemp.State = 1 Then _
                              TabTemp.Close
         
                           frmCADRECEBVENDA.Show 1
   
                           BlockInput False  'Desbloqueia o teclado
                           Else
                              SQL = "update PEDIDO set "
                              SQL = SQL & "status = 6 " 'não contabiliza
                              SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                              SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                              CONECTA_RETAGUARDA.Execute SQL
                        End If
                     End If
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
               End If
            End If
         End If
'===================================
         PEDIDO_ID_N = lstPedidos.SelectedItem.Text
         BlockInput False  'Desbloqueia o teclado
         If INDR_CONTROLA_ESTOQUE = False Then _
            Exit Sub
'===================================
         If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And USUARIO_ID_N = 144) Then
            SQL3 = IMPRESSORA_FISCAL_N
            CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
            Select Case CRITERIO
               Case "BEMATECH"
                  'Verifica se a Impressa esta ligada ou nao
                  RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
                  If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
                     BlockInput False  'Desbloqueia o teclado
                     RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
                     MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!", vbCritical, "MEGASIM"
                     Exit Sub
                     Else
                        INDR_VENDA = True
                        INDR_CUPOM_ABERTO = False
                        Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
                        If INDR_CUPOM_ABERTO = True Then _
                           CANCELA_CUPOM_ABERTO

                        Msg = ""
                        Indr_Erro = False
                        Call VerificaRetornoImpressora("", "", "Checando ECF")
                        If Indr_Erro = True Then
                           BlockInput False  'Desbloqueia o teclado
                           If Trim(Msg) <> "" Then
                           MsgBox Msg
                           End If
                           Exit Sub
                        End If
                  End If
                  INDR_VENDA = True
               Case "DARUMA"
               Case "Sweda"
            End Select
         End If
'===================================
         If TabPedido.State = 1 Then _
            TabPedido.Close

         SQL = "select * from PEDIDO "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedido.EOF Then
            PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
            VALOR_RECEBIDO_N = 0 & TabPedido.Fields("valor_recebido").Value

            If TabPedido!Status = 5 Then
               CNPJCPF_A = Trim(TabPedido!CGCCPF)
'=============================================================================
               If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And USUARIO_ID_N = 144) Then
                  INDR_PERGUNTA = True

                  If TabTemp.State = 1 Then _
                     TabTemp.Close

                  SQL = "select descricao from TIPOVENDA "
                  SQL = SQL & " where tipovenda_id = " & TabPedido.Fields("tipovenda_id").Value
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then _
                     If Left(UCase(Trim(TabTemp.Fields(0).Value)), 6) = "CARTAO" Or _
                     Left(UCase(Trim(TabTemp.Fields(0).Value)), 6) = "CARTÃO" Then _
                        INDR_PERGUNTA = False
                  If TabTemp.State = 1 Then _
                     TabTemp.Close

                  If INDR_PERGUNTA = True Then
                     Msg = "Confirma Faturamento ?"
                     PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                     Else: RESPOSTA = vbYes
                  End If
                  If RESPOSTA = vbYes Then
                     INDR_VENDA = True
'====================
                     MOSTRA_RODAPE_AQUI "Aguarde, imprimindo cupom fiscal ...", "", "", "", ""

                     'não fechar cupom
                     If INDR_ERRO_TEF = True Then
                        BlockInput False  'Desbloqueia o teclado
                        MOSTRA_RODAPE_AQUI "ERRO TEF ...", "", "", "", ""
                        Exit Sub
                     End If
'==============================
                     SQL3 = IMPRESSORA_FISCAL_N
                     CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
                     Select Case CRITERIO
                        Case "BEMATECH"
                           'Verifica se a Impressa esta ligada ou nao
                           RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
                           If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
                              BlockInput False  'Desbloqueia o teclado
                              RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
                              MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!!!", vbCritical, "MEGASIM"
                              Exit Sub
                              Else
                           End If
                        Case "DARUMA"
                        Case "Sweda"
                     End Select

                     'COMEÇA AQUI DE ACORDO O TIPO DA IMPRESSORA FISCAL
                     'ESSE AQUI É A ROTINA QUE NÃO É COM COMITANCIA
                     'DEPOIS TEM QUE FAZER O MESMO COM A TELA DE COMITANCIA

                     'BlockInput True   'Bloqueia o teclado
                     'incluindo esse teste aqui para que não cancele o cupom e imprima novamente quando
                     'cliente errar a senha ou ficar testando se tem crédito nos cartões dele

                     Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
                     'If INDR_CUPOM_ABERTO = True And INDR_ERRO_TEF = False Then
                     If INDR_ERRO_TEF = True Then
                        'chamando rotina do TEF
                        Msg = "Chamando TEF"
                        MOSTRA_RODAPE_AQUI Msg & " ...", "", "", "", ""
                        frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""

                        INDR_ERRO_TEF = False
                        If USA_TEF = True Then _
                           CHAMA_EASYTEF  'VERIFICA SE TEM CARTÃO

                        'chamando fechamento cupom fiscal
                        frmDISPLAYEMISSOR.FECHA_CUPOM_BEMATECH

                        Else: IMPRIME_CUPOM_FISCAL
                     End If

                     BlockInput False  'Desbloqueia o teclado

                     '=======================
                     Me.WindowState = 0

                     If Trim(NUMEROCUPOM) <> "" And PEDIDO_ID_N > 0 Then
                        SQL = "update PEDIDO set "
                        SQL = SQL & " status = 7 " 'CUPOM FISCAL
                        SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                        SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                        CONECTA_RETAGUARDA.Execute SQL
                     End If

If USA_NFe = True Then _
   frmNOTAGERA.Show 1

                     Else
                        If Trim(CNPJCPF_A) <> "99999999999" Then _
                           If USA_NFe = True Then _
                              frmNOTAGERA.Show 1
                  End If
                  Else
                     If Trim(CNPJCPF_A) <> "99999999999" Then _
                        If USA_NFe = True Then _
                           frmNOTAGERA.Show 1
               End If
'====================
ATUALIZA_ESTOQUE 0, PEDIDO_ID_N
'====================
            End If
         End If
         If TabPedido.State = 1 Then _
            TabPedido.Close
      End If   'If Not IsNull(lstPedidos.SelectedItem.Text) Then
   End If      'If Trim(lstPedidos.SelectedItem.Text) <> "" Then
   If TabPedido.State = 1 Then _
      TabPedido.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "FAZ_RECEBIMENTO"
End Sub

Sub GRAVA_CUPOM(Numero_Cupom As String)
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N <= 0 Then _
      Exit Sub
   If Trim(Numero_Cupom) = "" Then _
      Exit Sub
   If Not IsNumeric(Numero_Cupom) Then _
      Exit Sub

   If IMPRESSORA_ID_N <= 0 Then _
      IMPRESSORA_ID_N = 1

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'GRAVA TABELA CUPOM
   SQL = "select * from CUPOM"
   SQL = SQL & " where numr_cupom = " & Numero_Cupom
   SQL = SQL & " and Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO
   SQL = SQL & " and IMPRESSORA_ID = " & IMPRESSORA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      SQL = "update CUPOM set "
         SQL = SQL & " VALOR_CUPOM = " & tpMOEDA(VALOR_TOTAL_N)               'VALOR_CUPOM
      SQL = SQL & " where numr_cupom = " & Numero_Cupom
      SQL = SQL & " and Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO
      SQL = SQL & " and IMPRESSORA_ID = " & IMPRESSORA_ID_N
      Else
         SQL = "insert into CUPOM "
         SQL = SQL & " (CUPOM_ID,PEDIDO_ID,IMPRESSORA_ID,NUMR_CUPOM,Numr_Contador_Reinicio,VALOR_CUPOM)"
         SQL = SQL & " VALUES("

            SQL = SQL & MAX_ID("cupom_id", "cupom", "", "", "", "")  'CUPOM_ID
            SQL = SQL & "," & PEDIDO_ID_N                            'PEDIDO_ID
            SQL = SQL & "," & IMPRESSORA_ID_N                        'IMPRESSORA_ID
            SQL = SQL & "," & Numero_Cupom                           'NUMR_CUPOM
            SQL = SQL & "," & NUMR_CONTADOR_REINICIO                 'Numr_Contador_Reinicio
            SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)                 'VALOR_CUPOM
         SQL = SQL & ")"
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   CONECTA_RETAGUARDA.Execute SQL

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "GRAVA_CUPOM"
End Sub

Sub BUSCA_CLIENTE(CLIENTE_ID As Long)
'On Error GoTo ERRO_TRATA

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select CGCCPF, nome from CLIENTE "
   SQL = SQL & " where cliente_id = " & CLIENTE_ID
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      CRITERIO = Trim(TabCliente!NOME)

      txtCNPJCPF.PromptInclude = False
      If Len(Trim(TabCliente!CGCCPF)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If

      txtCNPJCPF.Text = TabCliente!CGCCPF
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = TabCliente!CGCCPF
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "BUSCA_CLIENTE"
End Sub

Sub CANCELA_CUPOM_ABERTO()
'On Error GoTo ERRO_TRATA

   SQL3 = IMPRESSORA_FISCAL_N
   CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
   Select Case CRITERIO
      Case "BEMATECH"

'tem que criar cancelamento cupom

         'RETORNO_ECF = Bematech_FI_NumeroCupom(NUMR_CUPOM_ABERTO)
         RETORNO_ECF = NUMR_CUPOM_ABERTO

         Indr_Erro = False

         RETORNO_ECF = Bematech_FI_CancelaCupom()
         'Função que analisa o retorno da impressora
         Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      Case "DARUMA"
         'RETORNO_ECF = Bematech_FI_NumeroCupom(NUMR_CUPOM_ABERTO)
         RETORNO_ECF = NUMR_CUPOM_ABERTO

         Indr_Erro = False

         'RETORNO_ECF = iCFCancelar_ECF_Daruma()
         'Função que analisa o retorno da impressora
         Call VerificaRetornoImpressoraDaruma("", "", "Emissão de Cupom Fiscal")
   End Select

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select pedido_id from CUPOM "
   SQL = SQL & " where numr_cupom = " & NUMR_CUPOM_ABERTO
   SQL = SQL & " and Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields(0).Value) Then
         NUMR_ID_N = TabTemp.Fields(0).Value

         SQL = "update PEDIDO set "
         SQL = SQL & " status = 9"
         SQL = SQL & " where pedido_id = " & NUMR_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "CANCELA_CUPOM_ABERTO"
End Sub

Sub IMPRIME_CUPOM_FISCAL()
'On Error GoTo ERRO_TRATA

   SQL3 = IMPRESSORA_FISCAL_N
   CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
   Select Case CRITERIO
      Case "BEMATECH"
         ROTINA_CUPOM_FISCAL_BEMATECH
      Case "DARUMA"
         ROTINA_CUPOM_FISCAL_DARUMA
      Case "Sweda"
         ROTINA_CUPOM_FISCAL_SWEDA
   End Select

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "IMPRIME_CUPOM_FISCAL"
End Sub

Sub ROTINA_CUPOM_FISCAL_BEMATECH()
'On Error GoTo ERRO_TRATA

   Dim TabPedidoItem    As New ADODB.Recordset
   Dim sParametros
   Dim CODG_PRODUTO_A   As String
   Dim EAN13_A          As String
   Dim DESCRICAO_A      As String
   Dim UN_A             As String
   Dim casasDecimaisQuantidade   As String
   Dim TipoAcrescimoDesconto     As String
   Dim ValorAcrescimo            As String
   Dim ValorDesconto             As String
   Dim tipoCalculo               As String
   Dim CODG_NCM_A                As String
   Dim informacoesAdicionais     As String
   Dim Qtde_N
   Dim casasDecimaisValorUnitario
Dim IndiceDepartamento
Dim CST_ICMS
Dim origemProduto
Dim itemListaServico
Dim codigoISS
Dim naturezaOperacaoISS
Dim indicadorIncentivoISS
Dim codigoIBGE
Dim CSOSN
Dim baseCalculoValorRetido
Dim ICMS_ValorRetido
Dim modoBaseCalculo
Dim percentualReducaoBaseCalculo
Dim ICMS_ST_ModoBaseCalculo
Dim ICMS_ST_PercentualMargemAdicionado
Dim ICMS_ST_PercentualReducaoBaseCalculo
Dim ICMS_ST_ValorReducaoBaseCalculo
Dim ICMS_ST_Aliquota
Dim ICMS_ST_Valor
Dim valorDesoneracaoICMS
Dim motivoDesoneracaoICMS
Dim aliquotaCalculoCredito
Dim creditoICMSSimples
Dim impostosIncidentes
Dim CST_PIS
Dim PIS_BaseCalculo
Dim PIS_Aliquota
Dim PIS_Valor
Dim PIS_QuantidadeVendida
Dim PIS_ValorAliquotaReais
Dim CST_COFINS
Dim COFINS_BaseCalculo
Dim COFINS_Aliquota
Dim COFINS_Valor
Dim COFINS_QuantidadeVendida
Dim COFINS_ValorAliquotaReais
Dim CEST

   'Símbolos identificadores dos totalizadores
   '(os mesmos usados nas máquinas registradoras)
   '----- Tnn . Tributado (sujeito ao ICMS)
   '----- ISnn . Tributado (sujeito ao ISS)
   '----- F . Substituição Tributária
   '----- i .Isenção
   '----- N . Não incidência;
   'Parâmetro8-Alíquota com o índice no tamanho de 2 caracteres,
   'ou FF (Substituição Tributária)
   'ou II (Isenção)
   'ou NN (Não Incidência)

   Dim Aliquota_N          As String
   Dim ALIQUOTA_ICMS_NORMAL_DENTRO_UF       As Long
   Dim CNPJCPF_CLIENTE     As String
   Dim INDR_ECF_ABERTO     As Boolean
   Dim ALIQ_IBPT_N         As Double
   
   Dim ORIGEM_MERDACO_N    As Integer

   Dim Tipo_Venda          As String
   Dim sTemp               As String
   Dim TIPO_QTDE_A         As String * 1
   Dim INDR_PROD_BALANCA   As Boolean

   NOME_CLI = SQL

   Indr_Erro = False
   INDR_VENDA = True
   NUMR_CUPOM_ABERTO = 0
   CONTA_TENTATIVA = 0
   INDR_ECF_ABERTO = False

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "SELECT PEDIDO.PEDIDO_ID,PEDIDO.EMPRESA_ID, PEDIDO.vendedor_id, PEDIDO.DT_REQ, PEDIDO.STATUS, "
   SQL = SQL & " PEDIDO.TIPO_REGISTRO, PEDIDO.VALOR_DESCONTO, PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_TOTAL, PEDIDO.cgccpf, "

   SQL = SQL & " produto.CODG_PRODuto, PEDIDOITEM.PERCICMS, PEDIDOITEM.VALOR_DESCONTO AS Desconto_item, "
   SQL = SQL & " PEDIDOITEM.STATUS AS Situacao_item, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "

   SQL = SQL & " Produto.Descricao, PEDIDOITEM.produto_id, PRODUTO.produto_balanca,"
   SQL = SQL & " PRODUTO.codg_ncm,PRODUTO.situacao_tributaria,Produto.Aliquota_Icms, "
   SQL = SQL & " PRODUTO.codg_barra, PRODUTO.origem_mercado,PRODUTO.unidade_medida,produto.codg_ncm"

   SQL = SQL & " from PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and pedidoitem.status <> 'C' "

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      PEDIDO_ID_N = TabPedidoItem.Fields("pedido_id").Value
      CNPJCPF_CLIENTE = "" & Trim(TabPedidoItem.Fields("cgccpf").Value)
      NOME_CLI = "" & Trim(TabPedidoItem.Fields("nome_cliente").Value)

      If CNPJCPF_CLIENTE = "99999999999" Then _
         CNPJCPF_CLIENTE = ""

      Msg = "Abrindo Gaveta"
      Me.Caption = Msg
      RETORNO_ECF = Bematech_FI_AcionaGaveta()

      INDR_PRI = True
      Indr_Erro = False

ABRINDO_CUPOM_FISCAL:

      RETORNO_ECF = Bematech_FI_AbreCupom(CNPJCPF_CLIENTE)
      Msg = "Abrindo Cupom Fiscal"
      MOSTRA_RODAPE_AQUI Msg & " ...", "", "", "", ""
      frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""
      Me.Caption = Msg

      Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
'tem que criar rotina de gravar cupom cancelado
      If INDR_CUPOM_ABERTO = True Then
         CANCELA_CUPOM_ABERTO
         GoTo ABRINDO_CUPOM_FISCAL
      End If

      If Indr_Erro = True Then
         If Indr_Cancela_Cupom = True Then
            GRAVANDO_CUPOM_ERRO
            Exit Sub
         End If

         If CONTA_TENTATIVA >= 3 Then
            If Indr_Erro = True Then
               MsgBox "Erro ao tentar processar cupom fiscal."
               GRAVANDO_CUPOM_ERRO
               Exit Sub
            End If
            Else
               Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
               PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                  Sleep 1000
                  GoTo ABRINDO_CUPOM_FISCAL
                  Else
                     Call frmINICIO.EasyTEF.CancelarVendasPendentes
                     GRAVANDO_CUPOM_ERRO
                     Exit Sub
               End If
         End If
      End If

      'pegando numero cupom aberto
      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOM = Space(1)
         Else: NUMEROCUPOM = Space(6)
      End If

      RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOM)
      Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
      If Trim(NUMEROCUPOM) <> "" Then
         NUMR_CUPOM_ABERTO = NUMEROCUPOM
         frmINICIO.NumeroCupomFiscal = NUMEROCUPOM

         TOTAL_DESCONTO_N = 0 & TabPedidoItem.Fields("VALOR_DESCONTO").Value

         GRAVA_CUPOM NUMEROCUPOM
         INDR_ECF_ABERTO = True
      End If
      If Indr_Erro = True Then _
         Me.Caption = "Erro na leitura do Cupom Fiscal, Cupom Fiscal = " & NUMR_CUPOM_ABERTO

      'lei 12.741
      ALIQ_IBPT_N = 0
      VALOR_TOTAL_IMPOSTO = 0
      VALOR_TOTAL_N = 0

      Indr_Erro = False
      While Not TabPedidoItem.EOF
         ORIGEM_MERDACO_N = 0 & TabPedidoItem.Fields("origem_mercado").Value
         ITEM_DESCONTO_N = 0 & TabPedidoItem.Fields("Desconto_item").Value
         VALOR_ITEM_N = 0 & TabPedidoItem.Fields("VALOR_ITEM").Value
         QTDE_PEDIDO = 0 & TabPedidoItem.Fields("QTD_PEDIDA").Value

         'VALOR_TOTAL_N = VALOR_ITEM_N * QTDE_PEDIDO + VALOR_TOTAL_N

         Aliquota_N = "01"
         ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0

         If Not IsNull(TabPedidoItem.Fields("situacao_tributaria").Value) Then _
            SITUACAO_TRIBUT_A = TabPedidoItem.Fields("situacao_tributaria").Value

         If Not IsNull(TabPedidoItem.Fields("Aliquota_Icms").Value) Then _
            ALIQUOTA_ICMS_NORMAL_DENTRO_UF = TabPedidoItem.Fields("Aliquota_Icms").Value

         'Tributada  e com cobrança do ICMS por substituição tributária
         If SITUACAO_TRIBUT_A = 10 Then _
            Aliquota_N = "FF"

         'Com redução de base de cálculo
         If SITUACAO_TRIBUT_A = 20 Then _
            Aliquota_N = "FF"
            'Aliquota_n = "NN"

         'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
         If SITUACAO_TRIBUT_A = 30 Then _
            Aliquota_N = "FF"

         'Isenta
         If SITUACAO_TRIBUT_A = 40 Then _
            Aliquota_N = "II"

         'Não tributada
         If SITUACAO_TRIBUT_A = 41 Then _
            Aliquota_N = "II"

         'Suspensão
         If SITUACAO_TRIBUT_A = 50 Then _
            Aliquota_N = "NN"

         If SITUACAO_TRIBUT_A = 60 Then _
            Aliquota_N = "FF"

         'Outras
         If SITUACAO_TRIBUT_A = 90 Then _
            Aliquota_N = "FF"
            'Aliquota_n = "NN"

         If IsNumeric(Aliquota_N) Then
            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "SELECT INDICE_ID from IMPRESSORA "
            SQL = SQL & " INNER JOIN INDICE "
            SQL = SQL & " ON IMPRESSORA.IMPRESSORA_ID = INDICE.IMPRESSORA_ID"

            SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and aliquota = " & ALIQUOTA_ICMS_NORMAL_DENTRO_UF
            SQL = SQL & " and IMPRESSORA.impressora_id = " & IMPRESSORA_ID_N
            SQL = SQL & " and numr_caixa = " & NUMERO_CAIXA_ECF
            SQL = SQL & " and numr_SERIE_IMP = '" & Trim(NUMERO_SERIE_ECF) & "'"

            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               If Not IsNull(TabUSU.Fields(0).Value) Then _
                  Aliquota_N = TabUSU.Fields(0).Value
            If TabUSU.State = 1 Then _
               TabUSU.Close

            CONT_N = 0
            If IsNumeric(Aliquota_N) Then _
               CONT_N = Aliquota_N

            If ((CONT_N < 10) And (CONT_N > 0)) Then _
               Aliquota_N = "0" & CONT_N
         End If

         CONT_N = 0

         CONTA_TENTATIVA = 0

TENTATIVAS:

         Sleep 700
         Msg = "Aguarde, Imprimindo Cupom Fiscal, Imprimindo Produto(s) " & Trim(TabPedidoItem.Fields("descricao").Value)
         MOSTRA_RODAPE_AQUI Msg & " ...", "", "", "", ""
         frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""
         Me.Caption = Msg

         TIPO_QTDE_A = "F"

         If Not IsNull(TabPedidoItem.Fields("produto_balanca").Value) Then _
            INDR_PROD_BALANCA = TabPedidoItem.Fields("produto_balanca").Value

         If INDR_PROD_BALANCA = True Then
            VALOR_ITEM_N = Format(VALOR_ITEM_N * QTDE_PEDIDO, strFormatacao2Digitos)
            QTDE_PEDIDO = 1
         End If

         'RETORNO_ECF = Bematech_FI_VendeItem(Replace(Trim(Left(TabPedidoItem.Fields("codg_produto").Value, 13)), ",", "."), _
            Trim(Replace(Left(Trim(TabPedidoItem.Fields("descricao").Value), 29), ",", ".")), _
            Trim(Aliquota_N), _
            TIPO_QTDE_A, _
            Replace(Format$(QTDE_PEDIDO, strFormatacao3Digitos), ",", "."), _
            "2", _
            Replace(Format$(VALOR_ITEM_N, strFormatacao2Digitos), ",", "."), _
            "$", _
            Replace(Format$(ITEM_DESCONTO_N, strFormatacao2Digitos), ",", ".") _
            )
'======================================
' Exemplo em Visual Basic

CODG_PRODUTO_A = Replace(Trim(Left(TabPedidoItem.Fields("codg_produto").Value, 13)), ",", ".")

EAN13_A = ""
If Trim(Len(TabPedidoItem.Fields("codg_barra").Value)) = 13 Then _
   EAN13_A = "7891000081754"

DESCRICAO_A = Trim(Replace(Left(Trim(TabPedidoItem.Fields("descricao").Value), 29), ",", "."))
IndiceDepartamento = "01"

'Aliquota = "FF"
'Trim (Aliquota_N)

'UnidadeMedida = "UN"
UN_A = Left(Trim(TabPedidoItem.Fields("unidade_medida").Value), 2)

'TipoQuantidade = "I"
'TIPO_QTDE_A

casasDecimaisQuantidade = "2"

'Quantidade = "1000"
Qtde_N = Replace(Format$(QTDE_PEDIDO, strFormatacao3Digitos), ",", ".")

casasDecimaisValorUnitario = "2"

'ValorUnitario = "1,00"
VALOR_ITEM_N = Replace(Format$(VALOR_ITEM_N, strFormatacao2Digitos), ",", ".")

TipoAcrescimoDesconto = "$"
ValorAcrescimo = "00,00"
ValorDesconto = "00,00"
tipoCalculo = "A"

'NCM = "09011200"
CODG_NCM_A = Left(Trim(TabPedidoItem.Fields("unidade_medida").Value), 8)

CFOP_A = "5102"

informacoesAdicionais = ""

CST_ICMS = "00"
origemProduto = "0"
itemListaServico = ""
codigoISS = ""
naturezaOperacaoISS = ""
indicadorIncentivoISS = ""
codigoIBGE = "5208707"
CSOSN = "500"
baseCalculoValorRetido = ""
ICMS_ValorRetido = ""
modoBaseCalculo = ""
percentualReducaoBaseCalculo = ""
ICMS_ST_ModoBaseCalculo = ""
ICMS_ST_PercentualMargemAdicionado = ""
ICMS_ST_PercentualReducaoBaseCalculo = ""
ICMS_ST_ValorReducaoBaseCalculo = ""
ICMS_ST_Aliquota = ""
ICMS_ST_Valor = ""
valorDesoneracaoICMS = ""
motivoDesoneracaoICMS = ""
aliquotaCalculoCredito = ""
creditoICMSSimples = ""
impostosIncidentes = "0,00"
'CST_PIS = "04"
CST_PIS = "00"
PIS_BaseCalculo = "000,00"
PIS_Aliquota = "00,00"
PIS_Valor = "0,00"
PIS_QuantidadeVendida = ""
PIS_ValorAliquotaReais = ""
'CST_COFINS = "04"
CST_COFINS = "00"
COFINS_BaseCalculo = "000,00"
COFINS_Aliquota = "00,00"
COFINS_Valor = "0,00"
COFINS_QuantidadeVendida = ""
COFINS_ValorAliquotaReais = ""
CEST = "0100100"

' Opção para concatenar variáveis

sParametros = "{" & Chr(34) & "CODG_PRODUTO_A" & Chr(34) & ":" & Chr(34) & CODG_PRODUTO_A & Chr(34) & "," & Chr(34) & "EAN13_A" & Chr(34) & ":" & Chr(34) & EAN13_A & Chr(34) & "," & Chr(34) & "descricao" & Chr(34) & ":" & Chr(34) & DESCRICAO_A & Chr(34) & "," & Chr(34) & "indiceDepartamento" & Chr(34) & ":" & Chr(34) & IndiceDepartamento & Chr(34) & "," _
& Chr(34) & "aliquota_n" & Chr(34) & ":" & Chr(34) & Aliquota_N & Chr(34) & "," & Chr(34) & "unidadeMedida" & Chr(34) & ":" & Chr(34) & UN_A & Chr(34) & "," & Chr(34) & "tipoQuantidade" & Chr(34) & ":" & Chr(34) & TIPO_QTDE_A & Chr(34) & "," & Chr(34) & "casasDecimaisQuantidade" & Chr(34) & ":" & Chr(34) & casasDecimaisQuantidade & Chr(34) & "," _
& Chr(34) & "Qtde_N" & Chr(34) & ":" & Chr(34) & Qtde_N & Chr(34) & "," & Chr(34) & "casasDecimaisValorUnitario" & Chr(34) & ":" & Chr(34) & casasDecimaisValorUnitario & Chr(34) & "," & Chr(34) & "VALOR_ITEM_N" & Chr(34) & ":" & Chr(34) & VALOR_ITEM_N & Chr(34) & "," & Chr(34) & "tipoAcrescimoDesconto" & Chr(34) & ":" & Chr(34) & TipoAcrescimoDesconto & Chr(34) & "," _
& Chr(34) & "valorAcrescimo" & Chr(34) & ":" & Chr(34) & ValorAcrescimo & Chr(34) & "," & Chr(34) & "valorDesconto" & Chr(34) & ":" & Chr(34) & ValorDesconto & Chr(34) & "," & Chr(34) & "tipoCalculo" & Chr(34) & ":" & Chr(34) & tipoCalculo & Chr(34) & "," & Chr(34) & "NCM" & Chr(34) & ":" & Chr(34) & CODG_NCM_A & Chr(34) & "," & Chr(34) & "CFOP_a" & Chr(34) & ":" _
& Chr(34) & CFOP_A & Chr(34) & "," & Chr(34) & "informacoesAdicionais" & Chr(34) & ":" & Chr(34) & informacoesAdicionais & Chr(34) & "," & Chr(34) & "CST_ICMS" & Chr(34) & ":" & Chr(34) & CST_ICMS & Chr(34) & "," & Chr(34) & "origemProduto" & Chr(34) & ":" & Chr(34) & origemProduto & Chr(34) & "," & Chr(34) & "itemListaServico" & Chr(34) & ":" _
& Chr(34) & itemListaServico & Chr(34) & "," & Chr(34) & "codigoISS" & Chr(34) & ":" & Chr(34) & codigoISS & Chr(34) & "," & Chr(34) & "naturezaOperacaoISS" & Chr(34) & ":" & Chr(34) & naturezaOperacaoISS & Chr(34) & "," & Chr(34) & "indicadorIncentivoISS" & Chr(34) & ":" & Chr(34) & indicadorIncentivoISS & Chr(34) & "," & Chr(34) & "codigoIBGE" & Chr(34) & ":" _
& Chr(34) & codigoIBGE & Chr(34) & "," & Chr(34) & "CSOSN" & Chr(34) & ":" & Chr(34) & CSOSN & Chr(34) & "," & Chr(34) & "baseCalculoValorRetido" & Chr(34) & ":" & Chr(34) & baseCalculoValorRetido & Chr(34) & "," & Chr(34) & "ICMS_ValorRetido" & Chr(34) & ":" & Chr(34) & ICMS_ValorRetido & Chr(34) & "," & Chr(34) & "modoBaseCalculo" & Chr(34) & ":" & Chr(34) & modoBaseCalculo & Chr(34) & "," _
& Chr(34) & "percentualReducaoBaseCalculo" & Chr(34) & ":" & Chr(34) & percentualReducaoBaseCalculo & Chr(34) & "," & Chr(34) & "ICMS_ST_ModoBaseCalculo" & Chr(34) & ":" & Chr(34) & ICMS_ST_ModoBaseCalculo & Chr(34) & "," & Chr(34) & "ICMS_ST_PercentualMargemAdicionado" & Chr(34) & ":" & Chr(34) & ICMS_ST_PercentualMargemAdicionado & Chr(34) & "," & Chr(34) & "ICMS_ST_PercentualReducaoBaseCalculo" & Chr(34) & ":" _
& Chr(34) & ICMS_ST_PercentualReducaoBaseCalculo & Chr(34) & "," & Chr(34) & "ICMS_ST_ValorReducaoBaseCalculo" & Chr(34) & ":" & Chr(34) & ICMS_ST_ValorReducaoBaseCalculo & Chr(34) & "," & Chr(34) & "ICMS_ST_Aliquota" & Chr(34) & ":" & Chr(34) & ICMS_ST_Aliquota & Chr(34) & "," & Chr(34) & "ICMS_ST_Valor" & Chr(34) & ":" & Chr(34) & ICMS_ST_Valor & Chr(34) & "," & Chr(34) & "valorDesoneracaoICMS" & Chr(34) & ":" _
& Chr(34) & valorDesoneracaoICMS & Chr(34) & "," & Chr(34) & "motivoDesoneracaoICMS" & Chr(34) & ":" & Chr(34) & motivoDesoneracaoICMS & Chr(34) & "," & Chr(34) & "aliquotaCalculoCredito" & Chr(34) & ":" & Chr(34) & aliquotaCalculoCredito & Chr(34) & "," & Chr(34) & "creditoICMSSimples" & Chr(34) & ":" & Chr(34) & creditoICMSSimples & Chr(34) & "," & Chr(34) & "impostosIncidentes" & Chr(34) & ":" & Chr(34) & impostosIncidentes & Chr(34) & "," _
& Chr(34) & "CST_PIS" & Chr(34) & ":" & Chr(34) & CST_PIS & Chr(34) & "," & Chr(34) & "PIS_BaseCalculo" & Chr(34) & ":" & Chr(34) & PIS_BaseCalculo & Chr(34) & "," & Chr(34) & "PIS_Aliquota" & Chr(34) & ":" & Chr(34) & PIS_Aliquota & Chr(34) & "," & Chr(34) & "PIS_Valor" & Chr(34) & ":" & Chr(34) & PIS_Valor & Chr(34) & "," & Chr(34) & "PIS_QuantidadeVendida" & Chr(34) & ":" & Chr(34) & PIS_QuantidadeVendida & Chr(34) & "," _
& Chr(34) & "PIS_ValorAliquotaReais" & Chr(34) & ":" & Chr(34) & PIS_ValorAliquotaReais & Chr(34) & "," & Chr(34) & "CST_COFINS" & Chr(34) & ":" & Chr(34) & CST_COFINS & Chr(34) & "," & Chr(34) & "COFINS_BaseCalculo" & Chr(34) & ":" & Chr(34) & COFINS_BaseCalculo & Chr(34) & "," & Chr(34) & "COFINS_Aliquota" & Chr(34) & ":" & Chr(34) & COFINS_Aliquota & Chr(34) & "," & Chr(34) & "COFINS_Valor" & Chr(34) & ":" & Chr(34) & COFINS_Valor & Chr(34) & "," _
& Chr(34) & "COFINS_QuantidadeVendida" & Chr(34) & ":" & Chr(34) & COFINS_QuantidadeVendida & Chr(34) & "," _
& Chr(34) & "COFINS_ValorAliquotaReais" & Chr(34) & ":" & Chr(34) & COFINS_ValorAliquotaReais & Chr(34) & "," & Chr(34) & "CEST" & Chr(34) & ":" & Chr(34) & CEST & Chr(34) & "}"

RETORNO_ECF = Bematech_FI_VendeItemCompletoJSON(sParametros)

         Msg = ""
         Call VerificaRetornoImpressora("Bematech_FI_VendeItemCompletoJSON", "", "Emissão de Cupom Fiscal")
         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, Imprimindo Produto(s) " & Trim(TabPedidoItem.Fields("descricao").Value)

         If Indr_Erro = True Then
            If CONTA_TENTATIVA >= 3 Then
               If Indr_Erro = True Then
                  MsgBox "Erro ao imprimir produto = " & Trim(TabPedidoItem.Fields("codg_produto").Value) & " - " & TabPedidoItem.Fields("descricao").Value & " , verificar."
                  GRAVANDO_CUPOM_ERRO
                  Exit Sub
               End If
               Else
                  Me.Caption = "Imprimindo Produto(s), Tentativas = " & CONTA_TENTATIVA & " ; Erro  " & Trim(TabPedidoItem.Fields("codg_produto").Value)
Sleep 1000

                  Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
                  PERGUNTA Msg, vbYesNo + 32, "Cupom Fiscal", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                     GoTo TENTATIVAS
                     Else
                        Call frmINICIO.EasyTEF.CancelarVendasPendentes
                        GRAVANDO_CUPOM_ERRO
                        Exit Sub
                  End If
            End If
         End If
'========================================================
'=============baixa estoque INICIO
         PRODUTO_ID_N = TabPedidoItem.Fields("produto_id").Value
         QTDE_PEDIDO = TabPedidoItem.Fields("QTD_PEDIDA").Value

'=============baixa estoque FIM
'==========================
         'CALCULO IMPOSTO LEI 12.741 (BUSCA CODIGO NCM DO CADASTRO DO PRODUTO,
         'LÊ TABELA 'IBPTax' QUE CONTEM A ALIQUOTA RELACIONADA AO NCM DO PRODUTO
         If INDR_LEI_12741 = True Then
            If Not IsNull(TabPedidoItem.Fields("codg_ncm").Value) Then
               If Trim(TabPedidoItem.Fields("codg_ncm").Value) <> "" Then

                  If TabTemp.State = 1 Then _
                     TabTemp.Close

                  SQL = "select ALIQNAC,ALIQIMP from IBPTax "
                  SQL = SQL & " where codg_ncm = '" & Trim(TabPedidoItem.Fields("codg_ncm").Value) & "'"
                  SQL = SQL & " and tabela = " & ORIGEM_MERDACO_N
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     ALIQ_IBPT_N = 0
   
                     If ORIGEM_MERDACO_N = 0 Then _
                        ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqnac").Value
                     If ORIGEM_MERDACO_N = 1 Then _
                        ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqimp").Value
   
                     VALOR_TOTAL_IMPOSTO = VALOR_TOTAL_IMPOSTO + ((VALOR_ITEM_N * QTDE_PEDIDO) * ALIQ_IBPT_N / 100)
   
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
               End If
            End If
         End If 'If INDR_LEI_12741 = True Then
'==========================

         PEDIDO_ID_N = TabPedidoItem.Fields("pedido_id").Value

         TabPedidoItem.MoveNext
      Wend
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

'FECHANDO CUPOM
'chamando rotina do TEF
      Msg = "Chamando TEF"
      MOSTRA_RODAPE_AQUI Msg & " ...", "", "", "", ""
      frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""

      INDR_ERRO_TEF = False
      If USA_TEF = True Then _
         CHAMA_EASYTEF  'VERIFICA SE TEM CARTÃO

'==============================
      FECHA_CUPOM_BEMATECH
'=====================================

      Else
         BlockInput False  'Desbloqueia o teclado
         MsgBox "Pedido Venda não encontrado."
   End If

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   BlockInput False  'Desbloqueia o teclado

Exit Sub

'SE DU ERRO ENTRA AQUI
'SE DU ERRO ENTRA AQUI
'SE DU ERRO ENTRA AQUI

   GRAVANDO_CUPOM_ERRO

Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "ROTINA_CUPOM_FISCAL_BEMATECH"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Sub ROTINA_CUPOM_FISCAL_DARUMA()
'On Error GoTo ERRO_TRATA

   Dim TabPedidoItem       As New ADODB.Recordset

   'Símbolos identificadores dos totalizadores
   '(os mesmos usados nas máquinas registradoras)
   '----- Tnn . Tributado (sujeito ao ICMS)
   '----- ISnn . Tributado (sujeito ao ISS)
   '----- F . Substituição Tributária
   '----- i .Isenção
   '----- N . Não incidência;
   'Parâmetro8-Alíquota com o índice no tamanho de 2 caracteres,
   'ou FF (Substituição Tributária)
   'ou II (Isenção)
   'ou NN (Não Incidência)

   Dim Aliquota_N          As String
   Dim ALIQUOTA_ICMS_NORMAL_DENTRO_UF       As Long
   Dim CONTA_TENTATIVA     As Long
   Dim Mensagem_Final      As String
   Dim f                   As Variant
   Dim LocalRetorno        As String
   Dim TOTAL_DESCONTO_N    As Double
   Dim ITEM_DESCONTO_N     As Double
   Dim Descr_Forma_Pagto   As String
   Dim NOME_CLI            As String
   Dim CNPJCPF_CLIENTE     As String
   Dim INDR_ECF_ABERTO     As Boolean

   NOME_CLI = SQL

   Indr_Erro = False
   INDR_VENDA = True

   NUMR_CUPOM_ABERTO = 0

   CONTA_TENTATIVA = 0

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "SELECT PEDIDO.EMPRESA_ID, PEDIDO.pedido_id, PEDIDO.vendedor_id, PEDIDO.DT_REQ, PEDIDO.STATUS, "
   SQL = SQL & " PEDIDO.TIPO_REGISTRO, PEDIDO.VALOR_DESCONTO, PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_TOTAL, PEDIDO.cgccpf,"
   SQL = SQL & " PRODUTO.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.PERCICMS, Produto.Aliquota_Icms, "
   SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO AS Desconto_item, PEDIDOITEM.STATUS AS Situacao_item, Produto.Descricao, PEDIDOITEM.produto_id "

   SQL = SQL & " from PEDIDO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and pedidoitem.status <> 'C' "

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then

      CNPJCPF_CLIENTE = "" & Trim(TabPedidoItem.Fields("cgccpf").Value)
      NOME_CLI = "" & Trim(TabPedidoItem.Fields("nome_cliente").Value)

      If CNPJCPF_CLIENTE = "99999999999" Then _
         CNPJCPF_CLIENTE = ""

      Msg = "Abrindo Gaveta"
      Me.Caption = Msg
      Call VerificaRetornoImpressoraDaruma("", "", "Acionamento da Gaveta")

      INDR_PRI = True
      Indr_Erro = False

ABRINDO_CUPOM_FISCAL:

      NUMEROCUPOM = Space(6)
 
      'Verifica se Impressora Esta Ligada
      'RETORNO_ECF = rVerificarImpressoraLigada_ECF_Daruma()
      If RETORNO_ECF <> 1 Then
         BlockInput False  'Desbloqueia o teclado
         MsgBox "ECF Desligado, verifique!", vbCritical, "MEGASIM"
         INDR_CUPOM_ABERTO = False
         Exit Sub
      End If

      Msg = "Abrindo Cupom Fiscal"
      Me.Caption = Msg

      'RETORNO_ECF = iCFAbrir_ECF_Daruma(CNPJCPF_CLIENTE, NOME_CLI, "")

      Call VerificaRetornoImpressoraDaruma("Número do Último Cupom: ", NUMEROCUPOM, "Informações da Impressora")

If INDR_CUPOM_ABERTO = True Then
   CANCELA_CUPOM_ABERTO
   GoTo ABRINDO_CUPOM_FISCAL
End If

      If Indr_Erro = True Then
         If Indr_Cancela_Cupom = True Then _
            GoTo GRAVANDO_ERRO_EMISSAO_CUPOM

         If CONTA_TENTATIVA >= 3 Then
            If Indr_Erro = True Then
               MsgBox "Erro ao abrir cupom fiscal"
               GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
            End If
            Else
               Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
               PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                  Sleep 1000
                  GoTo ABRINDO_CUPOM_FISCAL
                  Else
                     'Call frmINICIO.EasyTEF.CancelarVendasPendentes
                     GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
               End If
         End If
      End If

      'pegando numero cupom aberto
      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOM = Space(1)
         Else: NUMEROCUPOM = Space(6)
      End If

      'RETORNO_ECF = Daruma_FI_NumeroCupom(NUMEROCUPOM)
      
      Call VerificaRetornoImpressoraDaruma("", "", "Emissão de Cupom Fiscal")
      If Trim(NUMEROCUPOM) <> "" Then
         NUMR_CUPOM_ABERTO = NUMEROCUPOM
         'frmINICIO.NumeroCupomFiscal = NUMEROCUPOM

         VALOR_TOTAL_N = 0
         TOTAL_DESCONTO_N = 0 & TabPedidoItem.Fields("VALOR_DESCONTO").Value

         GRAVA_CUPOM NUMEROCUPOM
         INDR_ECF_ABERTO = True
      End If
      If Indr_Erro = True Then _
         Me.Caption = "Erro na leitura do Cupom Fiscal, Cupom Fiscal = " & NUMR_CUPOM_ABERTO

      Indr_Erro = False
      While Not TabPedidoItem.EOF
         ITEM_DESCONTO_N = 0 & TabPedidoItem.Fields("Desconto_item").Value
         VALOR_ITEM_N = 0 & TabPedidoItem.Fields("VALOR_ITEM").Value
         QTDE_PEDIDO = 0 & TabPedidoItem.Fields("QTD_PEDIDA").Value

         VALOR_TOTAL_N = VALOR_ITEM_N * QTDE_PEDIDO + VALOR_TOTAL_N

         Aliquota_N = "01"
         ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select situacao_tributaria,aliquota_icms from PRODUTO "
         SQL = SQL & " where produto_id = " & TabPedidoItem.Fields("produto_id").Value
         SQL = SQL & " and situacao = 'A' "
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               SITUACAO_TRIBUT_A = TabConsulta.Fields(0).Value

            If Not IsNull(TabConsulta.Fields(1).Value) Then _
               ALIQUOTA_ICMS_NORMAL_DENTRO_UF = TabConsulta.Fields(1).Value
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         'Tributada  e com cobrança do ICMS por substituição tributária
         If SITUACAO_TRIBUT_A = 10 Then _
            Aliquota_N = "FF"

         'Com redução de base de cálculo
         If SITUACAO_TRIBUT_A = 20 Then _
            Aliquota_N = "FF"
            'Aliquota_n = "NN"

         'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
         If SITUACAO_TRIBUT_A = 30 Then _
            Aliquota_N = "FF"

         'Isenta
         If SITUACAO_TRIBUT_A = 40 Then _
            Aliquota_N = "II"

         'Não tributada
         If SITUACAO_TRIBUT_A = 41 Then _
            Aliquota_N = "II"

         'Suspensão
         If SITUACAO_TRIBUT_A = 50 Then _
            Aliquota_N = "NN"

         If SITUACAO_TRIBUT_A = 60 Then _
            Aliquota_N = "FF"

         'Outras
         If SITUACAO_TRIBUT_A = 90 Then _
            Aliquota_N = "FF"
            'Aliquota_n = "NN"

         If IsNumeric(Aliquota_N) Then
            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "select indice_id from INDICE "
            SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and aliquota = " & ALIQUOTA_ICMS_NORMAL_DENTRO_UF
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               If Not IsNull(TabUSU.Fields(0).Value) Then _
                  Aliquota_N = TabUSU.Fields(0).Value
            If TabUSU.State = 1 Then _
               TabUSU.Close

            CONT_N = 0
            If IsNumeric(Aliquota_N) Then _
               CONT_N = Aliquota_N

            If ((CONT_N < 10) And (CONT_N > 0)) Then _
               Aliquota_N = "0" & CONT_N
         End If

         CONT_N = 0
         CONTA_TENTATIVA = 0

TENTATIVAS:

Sleep 700

         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, Imprimindo Produto(s) " & Trim(TabPedidoItem.Fields("descricao").Value)

'RETORNO_ECF = Daruma_FI_VendeItem(Replace(Trim(Left(TabPedidoItem.Fields("codg_prod").Value, 13)), ",", "."), _
                              Trim(Replace(Left(Trim(TabPedidoItem.Fields("descricao").Value), 29), ",", ".")), _
                              Trim(Aliquota_N), _
                              UCase("f"), _
                              Replace(Format$(QTDE_PEDIDO, strFormatacao3Digitos), ",", "."), _
                              "2", _
                              Replace(Format$(VALOR_ITEM_N, strFormatacao2Digitos), ",", "."), _
                              "$", _
                              Replace(Format$(ITEM_DESCONTO_N, strFormatacao2Digitos), ",", ".") _
                              )

         Msg = ""
         Call VerificaRetornoImpressoraDaruma("", "", "Emissão de Cupom Fiscal")
         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, Imprimindo Produto(s) " & Trim(TabPedidoItem.Fields("descricao").Value)

         If Indr_Erro = True Then
            If CONTA_TENTATIVA >= 3 Then
               If Indr_Erro = True Then
                  MsgBox "Erro ao imprimir produto = " & Trim(TabPedidoItem.Fields("codg_produto").Value) & " - " & TabPedidoItem.Fields("descricao").Value & " , verificar."
                  GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
               End If
               Else
                  Me.Caption = "Imprimindo Produto(s), Tentativas = " & CONTA_TENTATIVA & " ; Erro  " & Trim(TabPedidoItem.Fields("codg_produto").Value)
Sleep 1000

                  Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
                  PERGUNTA Msg, vbYesNo + 32, "Cupom Fiscal", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                     GoTo TENTATIVAS
                     Else
                        'Call frmINICIO.EasyTEF.CancelarVendasPendentes
                        GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
                  End If
            End If
         End If

'=============baixa estoque INICIO
         PRODUTO_ID_N = TabPedidoItem.Fields("produto_id").Value
         QTDE_PEDIDO = TabPedidoItem.Fields("QTD_PEDIDA").Value

'=============baixa estoque FIM

         TabPedidoItem.MoveNext
      Wend
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

'FECHANDO CUPOM
      If INDR_PRI = True Then
         INDR_PRI = False
         VALOR_ITEM_N = 0

         Mensagem_Final = "Obrigado, Volte Sempre."
         While Len(Mensagem_Final) < 48
            Mensagem_Final = Mensagem_Final & " "
         Wend

'=======================
         txtCNPJCPF.PromptInclude = False
         If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
            If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
            txtCNPJCPF.PromptInclude = True

            SQL = "Cliente: " & Trim(txtCNPJCPF.Text)

            While Len(SQL) < 48
               SQL = SQL & " "
            Wend

            Mensagem_Final = Mensagem_Final & SQL
         End If

         If Trim(UCase(NOME_CLI)) <> UCase("Consumidor Final") Then
            SQL = Trim(Left(NOME_CLI, 48))

            While Len(SQL) < 48
               SQL = SQL & " "
            Wend

            Mensagem_Final = Mensagem_Final & SQL
         End If
'=======================

         SQL = "NºPedido =  " & PEDIDO_ID_N
         While Len(SQL) < 48
            SQL = SQL & " "
         Wend

         Mensagem_Final = Mensagem_Final & SQL

         NOME_VENDEDOR = "Balcão"

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "SELECT DESCRICAO FROM PEDIDO "
         SQL = SQL & " INNER JOIN vwVendedor "
         SQL = SQL & " ON PEDIDO.VENDEDOR_ID = vwVendedor.VENDEDOR_ID"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               NOME_VENDEDOR = Trim(TabTemp.Fields(0).Value)
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "Vendedor: " & Trim(NOME_VENDEDOR)
         While Len(SQL) < 48
            SQL = SQL & " "
         Wend

         Mensagem_Final = Mensagem_Final & SQL

         CONTA_TENTATIVA = 0

INICIANDO_FECHAMENTO_CUPOM_FISCAL:

         'RETORNO_ECF = Daruma_FI_IniciaFechamentoCupom("D", "$", Replace(Format(TOTAL_DESCONTO_N, strFormatacao2Digitos), ",", "."))
         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Iniciando Fechamento Cupom Fiscal"
            Call VerificaRetornoImpressoraDaruma("Daruma_FI_IniciaFechamentoCupom", "", "Emissão de Cupom Fiscal")
         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Iniciando Fechamento Cupom Fiscal"

         If Indr_Erro = True Then
            If CONTA_TENTATIVA >= 10 Then
               If Indr_Erro = True Then
                  MsgBox Msg & " Iniciando Fechamento Cupom Fiscal ; ERRO"
                  GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
               End If
               Else
                  Me.Caption = "Iniciando Fechamento Cupom Fiscal ; ERRO, Tentativas = " & CONTA_TENTATIVA & " ; Erro  "
Sleep 1000
                  Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
                  PERGUNTA Msg, vbYesNo + 32, "Cupom Fiscal", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                     GoTo INICIANDO_FECHAMENTO_CUPOM_FISCAL
                     Else
                        'Call frmINICIO.EasyTEF.CancelarVendasPendentes
                        GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
                  End If
            End If
         End If

         CONTA_TENTATIVA = 0

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "SELECT ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
         SQL = SQL & " from LANCAMENTO "
         SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
         SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
         SQL = SQL & " INNER JOIN FORMAPAGTO "
         SQL = SQL & " ON ITEMLANCAMENTO.formapagto_id = FORMAPAGTO.formapagto_id"

         SQL = SQL & " where LANCAMENTO.numr_doc = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF

EFETUANDO_FORMA_DE_PAGAMENTO:

            ITEM_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
            Descr_Forma_Pagto = "" & Trim(TabTemp.Fields("descricao").Value)
            If UCase(TabTemp.Fields("descricao").Value) = UCase("Dinheiro") Then _
               Descr_Forma_Pagto = "Dinheiro"

            'RETORNO_ECF = Daruma_FI_EfetuaFormaPagamento( _
                      Trim(Left(Descr_Forma_Pagto, 15)), _
                      Replace(Format(TabTemp.Fields("valor_item").Value - ITEM_DESCONTO_N, strFormatacao2Digitos), ",", "."))

            Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Efetuando Forma de Pagamento"
               Call VerificaRetornoImpressoraDaruma("daruma_FI_EfetuaFormaPagamentoIndice", "", "Emissão de Cupom Fiscal")
            Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Efetuando Forma de Pagamento"

            If Indr_Erro = True Then
               If CONTA_TENTATIVA >= 10 Then
                  If Indr_Erro = True Then
                     MsgBox Msg & " Efetuando_Forma_de_Pagamento ; ERRO"
                     GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
                  End If
                  Else
                     Me.Caption = "Efetuando_Forma_de_Pagamento ; ERRO , Tentativas = " & CONTA_TENTATIVA & " ; Erro  "
Sleep 1000
                  Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
                  PERGUNTA Msg, vbYesNo + 32, "Cupom Fiscal", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                     GoTo EFETUANDO_FORMA_DE_PAGAMENTO
                     Else
                        'Call frmINICIO.EasyTEF.CancelarVendasPendentes
                        GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
                  End If
               End If
            End If
            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONTA_TENTATIVA = 0

Finalizando_Fechamento_Cupom_Fiscal:

         'RETORNO_ECF = Daruma_FI_TerminaFechamentoCupom(Mensagem_Final)
         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Finalizando Fechamento Cupom Fiscal"
            Call VerificaRetornoImpressoraDaruma("daruma_FI_TerminaFechamentoCupom", "", "Emissão de Cupom Fiscal")
         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Finalizando Fechamento Cupom Fiscal"

         If Indr_Erro = True Then
            If CONTA_TENTATIVA >= 10 Then
               If Indr_Erro = True Then
                  MsgBox Msg & " Finalizando_Fechamento_Cupom_Fiscal ; ERRO"
                  GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
               End If
               Else
                  Me.Caption = "Finalizando_Fechamento_Cupom_Fiscal ; ERRO , Tentativas = " & CONTA_TENTATIVA & " ; Erro  "
Sleep 1000
                  Msg = "Impressora não responde. Tentar novamente? <Sim>/<Não>"
                  PERGUNTA Msg, vbYesNo + 32, "Cupom Fiscal", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     CONTA_TENTATIVA = CONTA_TENTATIVA + 1
                     GoTo Finalizando_Fechamento_Cupom_Fiscal
                     Else
                        'Call frmINICIO.EasyTEF.CancelarVendasPendentes
                        GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
                  End If
            End If
         End If

         If Indr_Erro = True Then
            MsgBox "Erro no Fechamento do Cupom Fiscal."
            GoTo GRAVANDO_ERRO_EMISSAO_CUPOM
         End If
      End If

      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOM = Space(1)
         Else: NUMEROCUPOM = Space(6)
      End If

      NUMR_SEQ_N = 0

LE_ULTIMO_ECF:

Sleep 1000

      Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Ultimo Cupom Impresso"
         'RETORNO_ECF = Daruma_FI_NumeroCupom(NUMEROCUPOM)
      Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Ultimo Cupom Impresso"

      If NUMEROCUPOM = "" Then
         If Not IsNumeric(NUMEROCUPOM) Then
            MsgBox "Erro na leitura do ultimo cupom impresso.  \" & NUMEROCUPOM
            NUMR_SEQ_N = NUMR_SEQ_N + 1
            If NUMR_SEQ_N < 10 Then
               Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Tentando ler Ultimo Cupom Impresso"
Sleep 1000
               GoTo LE_ULTIMO_ECF
            End If
         End If
      End If

      If Trim(NUMEROCUPOM) = "" Then _
         MsgBox "Atenção, erro de comunicação com impressora. Cupom Fiscal não gravado."

      GRAVA_CUPOM NUMEROCUPOM

      BlockInput False  'Desbloqueia o teclado
      Me.Caption = "OK, " & Msg & " " & "Fim Impressão"
      Exit Sub

   '----------------------------------------------------------------------------------------------------------------------
   'SE DU ERRO ENTRA AQUI
GRAVANDO_ERRO_EMISSAO_CUPOM:
      'Close #F

      If Indr_Erro = True Then
         MsgBox "Ocorreu erro, cupom fiscal " & NUMR_CUPOM_ABERTO & " será cancelado. " & Msg

         Dim RETORNOSTATUS As String

         NUMR_SEQ_N = 0

LE_ULTIMO_ECF2:

         If (LocalRetorno = "1") Then 'Grava retorno em arquivo
            NUMEROCUPOMCancelado = Space(1)
            Else: NUMEROCUPOMCancelado = Space(6)
         End If

         'RETORNO_ECF = Daruma_FI_NumeroCupom(NUMEROCUPOMCancelado)
         'Função que analisa o retorno da impressora
         'Call VerificaRetornoImpressoradaruma("Número do Último Cupom: ", _
              NumeroCupomCancelado, "Informações da Impressora")
         Me.Caption = "ERRO, " & Msg & " Ultimo Cupom Impresso"

         If NUMEROCUPOMCancelado = "" Then
            If Not IsNumeric(NUMEROCUPOMCancelado) Then
               MsgBox "Erro na leitura do ultimo cupom impresso.  \" & NUMEROCUPOM
               NUMR_SEQ_N = NUMR_SEQ_N + 1
               'If NUMR_SEQ_N < 3 Then _
                  GoTo LE_ULTIMO_ECF2
            End If
         End If

         If IsNumeric(NUMEROCUPOMCancelado) Then
            If NUMEROCUPOMCancelado = NUMR_CUPOM_ABERTO Then
               'RETORNO_ECF = iCFCancelar_ECF_Daruma()
               'Função que analisa o retorno da impressora
               Call VerificaRetornoImpressoraDaruma("iCFCancelar_ECF_Daruma", "", "Emissão de Cupom Fiscal")
               Me.Caption = "ERRO, " & Msg & " Cancelando Cupom Fiscal"

               GRAVA_CUPOM NUMEROCUPOMCancelado

               NUMR_ID_N = 0

               Else: MsgBox "Erro, cupom fiscal diferente do impresso, não cancelado."
            End If
         End If
      End If

      Me.Caption = "OK, " & Msg & " / " & "Fim Impressão, ECF cancelado  " & NUMEROCUPOMCancelado
      INDR_VENDA = False

      Else: MsgBox "Pedido Venda não encontrado."
   End If

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "ROTINA_CUPOM_FISCAL_DARUMA"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Sub ROTINA_CUPOM_FISCAL_SWEDA()
'On Error GoTo ERRO_TRATA

   If frmINICIO.Sweda.PortOpen = False Then _
      frmINICIO.Sweda.PortOpen = True

   RETORNO_ECF = frmINICIO.Sweda.Input

   'Efetua a leitura do próximo cupom
   frmINICIO.Sweda.Output = Chr(27) & ".271}"
   Tempo 0.8
   RETORNO_ECF = frmINICIO.Sweda.Input

   'If InStr(Trim(a), "+") = 0 Then _
      GoTo Novamente

   'If InStr(a, ".") Then
   '   ID_Cupom = Mid(ret, 14, 4) + 1
   '   Else: ID_Cupom = Mid(ret, 13, 4) + 1
   'End If

   'Abri cupom fiscal
   frmINICIO.Sweda.Output = Chr(27) & ".17}"
   Tempo 2.5
   RETORNO_ECF = frmINICIO.Sweda.Input
   'parametros
   'Sweda.Output = Chr(27) & ".09|2|01   " & _
                  Format(usuarioatual, "00") & "}"
   RETORNO_ECF = frmINICIO.Sweda.Input

'============ITENS
   'If Len(Trim(LblDescricao)) > 24 Then LblDescricao = Left(Trim(LblDescricao), 24)
   
   'Desc = Trim(LblDescricao) & Space(24 - Len(Trim(LblDescricao)))
   
   RETORNO_ECF = frmINICIO.Sweda.Input

NovamenteImpressao:
   'frmINICIO.Sweda.Output = Chr(27) & ".01" & _
                          Format(CodigoEan13, "0000000000000") & _
                          Qtde & _
                          VrU & _
                          vrt & _
                          UCase(Desc) & _
                          Aliquota & "}"
   
   Tempo 0.7
   RETORNO_ECF = frmINICIO.Sweda.Input
               
   'If Trim(a) <> "" Then
   '    If Mid(ret, 2, 1) = "+" Then
           'Ret = Frminicio.Sweda.Input
           'Sweda.Output = Chr(27) & ".28"
           'Tempo 1
           'Ret = Frminicio.Sweda.Input
           'If Trim(a) <> "" Then MsgBox mid(ret, 58, 35)
   '    ElseIf Mid(ret, 2, 1) = "-" Then
           'Ret = Frminicio.Sweda.Input
           'Sweda.Output = Chr(27) & ".28"
           'Tempo 0.3
           'Ret = Frminicio.Sweda.Input
           'If Trim(a) <> "" Then
            '   If mid(ret, 10, 1) = "C" Then
            '   Else
             '  End If
           'End If
   '    End If
       
   '    If InStr(1, a, "QUANT X UNIT") <> 0 Then
           '.-0002ERRO-QUANT X UNIT. DIFERENTE}
   '        vrt = Format(CCur(vrt) - 1, "000000000000")
   '        GoTo NovamenteImpressao
   '    End If
   'End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "ROTINA_CUPOM_FISCAL_SWEDA"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   Dim TIPO_VENDA_ID_N        As Long
   Dim Valor_Tipo_Venda_N     As Double
   Dim Descrição_Tipo_Venda   As String

   lstTotais.ListItems.Clear
   TIPO_VENDA_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT distinct(TIPOVENDA_ID) from PEDIDO "
   SQL = SQL & " where tipo_registro in ('S','R','D','OS') "
   SQL = SQL & " and status = 2" 'gerado somente Pedido"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by TIPOVENDA_ID "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum(valor_total-valor_desconto) from PEDIDO "
      SQL = SQL & " where tipo_registro in ('S','R','D','OS') "
      SQL = SQL & " and status = 2" 'gerado somente Pedido"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and tipovenda_id = " & TabTemp.Fields(0).Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         If Not IsNull(TabConsulta.Fields(0).Value) Then
            If TabConsulta.Fields(0).Value > 0 Then
               SqL2 = TabTemp.Fields(0).Value & "-" & Mostra_Descrição_TipoVenda(TabTemp.Fields(0).Value) & " = " & Format(TabConsulta.Fields(0).Value, strFormatacao2Digitos)
               Set item = lstTotais.ListItems.Add(, "seq." & TabConsulta.Fields(0).Value, SqL2)
            End If
         End If
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, Me.name, "MOSTRA_TOTAIS"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Sub MOSTRA_RODAPE_AQUI(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
'On Error GoTo ERRO_TRATA

   If Trim(Msg1) <> "" Then
      barRodape.Panels.Clear
      barRodape.Panels.Add (1)
      barRodape.Panels(1).Text = Trim(Msg1)
      barRodape.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         barRodape.Panels.Add (2)
         barRodape.Panels(2).Text = Trim(Msg2)
         barRodape.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            barRodape.Panels.Add (3)
            barRodape.Panels(3).Text = Trim(Msg3)
            barRodape.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               barRodape.Panels.Add (4)
               barRodape.Panels(4).Text = Trim(Msg4)
               barRodape.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  barRodape.Panels.Add (5)
                  barRodape.Panels(5).Text = Trim(Msg5)
                  barRodape.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.description, "mdlGeral", "MOSTRA_RODAPE_AQUI"
End Sub

Sub GRAVANDO_CUPOM_ERRO()

GRAVANDO_ERRO_EMISSAO_CUPOM:
   'Close #F

   If Indr_Erro = True Then
      MsgBox "Ocorreu erro, cupom fiscal " & NUMR_CUPOM_ABERTO & " será cancelado. " & Msg

      Dim RETORNOSTATUS As String

      NUMR_SEQ_N = 0

LE_ULTIMO_ECF2:

      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOMCancelado = Space(1)
         Else: NUMEROCUPOMCancelado = Space(6)
      End If

      RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOMCancelado)
      'Função que analisa o retorno da impressora
      'Call VerificaRetornoImpressora("Número do Último Cupom: ", _
           NumeroCupomCancelado, "Informações da Impressora")
      Me.Caption = "ERRO, " & Msg & " Ultimo Cupom Impresso"

      If NUMEROCUPOMCancelado = "" Then
         If Not IsNumeric(NUMEROCUPOMCancelado) Then
            MsgBox "Erro na leitura do ultimo cupom impresso.  \" & NUMEROCUPOM
            NUMR_SEQ_N = NUMR_SEQ_N + 1
            'If NUMR_SEQ_N < 3 Then _
               GoTo LE_ULTIMO_ECF2
         End If
      End If

      If IsNumeric(NUMEROCUPOMCancelado) Then
         If NUMEROCUPOMCancelado = NUMR_CUPOM_ABERTO Then
            RETORNO_ECF = Bematech_FI_CancelaCupom()
            'Função que analisa o retorno da impressora
            Call VerificaRetornoImpressora("Bematech_FI_CancelaCupom", "", "Emissão de Cupom Fiscal")
            Me.Caption = "ERRO, " & Msg & " Cancelando Cupom Fiscal"

            GRAVA_CUPOM NUMEROCUPOMCancelado

            NUMR_ID_N = 0

            Else: MsgBox "Erro, cupom fiscal diferente do impresso, não cancelado."
         End If
      End If
   End If

   Me.Caption = "OK, " & Msg & " / " & "Fim Impressão, ECF cancelado  " & NUMEROCUPOMCancelado
   INDR_VENDA = False
End Sub

Sub FECHA_CUPOM_BEMATECH()
   'não fechar cupom
   BlockInput False  'Desbloqueia o teclado

      If INDR_ERRO_TEF = True Then
         SQL = "update PEDIDO set "
         SQL = SQL & "status = 2 " 'não passou cartão volta para situação não faturado
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
         Exit Sub
      End If

      If INDR_PRI = True Then
         INDR_PRI = False
         VALOR_ITEM_N = 0

         Mensagem_Final = "Obrigado, Volte Sempre."
         While Len(Mensagem_Final) < 48
            Mensagem_Final = Mensagem_Final & " "
         Wend

'=======================
         txtCNPJCPF.PromptInclude = False
         If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
            If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
            txtCNPJCPF.PromptInclude = True

            SQL = "Cliente: " & Trim(txtCNPJCPF.Text)

            While Len(SQL) < 48
               SQL = SQL & " "
            Wend

            Mensagem_Final = Mensagem_Final & SQL
         End If

         If Trim(UCase(NOME_CLI)) <> UCase("Consumidor Final") Then
            SQL = Trim(Left(NOME_CLI, 48))

            While Len(SQL) < 48
               SQL = SQL & " "
            Wend

            Mensagem_Final = Mensagem_Final & SQL
         End If

'=======================
         SQL = "NºPedido =  " & PEDIDO_ID_N
         While Len(SQL) < 48
            SQL = SQL & " "
         Wend

         Mensagem_Final = Mensagem_Final & SQL
'=======================
         NOME_VENDEDOR = "Balcão"

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "SELECT DESCRICAO FROM PEDIDO "
         SQL = SQL & " INNER JOIN vwVendedor "
         SQL = SQL & " ON PEDIDO.VENDEDOR_ID = vwVendedor.VENDEDOR_ID"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               NOME_VENDEDOR = Trim(TabTemp.Fields(0).Value)
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "Vendedor: " & Trim(NOME_VENDEDOR)
         While Len(SQL) < 48
            SQL = SQL & " "
         Wend

         Mensagem_Final = Mensagem_Final & SQL

         CONTA_TENTATIVA = 0

INICIANDO_FECHAMENTO_CUPOM_FISCAL:
         Msg = "Aguarde, iniciando fechamento cupom fiscal ..."
         MOSTRA_RODAPE_AQUI Msg, "", "", "", ""
         frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""

         If USA_TEF = True Then
            ' desconto cielo premia
            Parametros = Array("D", "$", Format(frmINICIO.EasyTEF.ValorCampo709_000 + TOTAL_DESCONTO_N, "#0.00"))
            Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeIniciarFechamentoCupomFiscal, Parametros, OperacaoECFOK)

            If OperacaoECFOK = False Then
               BlockInput False  'Desbloqueia o teclado
               MsgBox "Não foi possível iniciar o fechamento do cupom fiscal.", vbCritical
               Exit Sub
            End If
            Else
               RETORNO_ECF = Bematech_FI_IniciaFechamentoCupom("D", "$", Replace(Format(TOTAL_DESCONTO_N, strFormatacao2Digitos), ",", "."))
               Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Iniciando Fechamento Cupom Fiscal"
                  Call VerificaRetornoImpressora("Bematech_FI_IniciaFechamentoCupom", "", "Emissão de Cupom Fiscal")
               Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Iniciando Fechamento Cupom Fiscal"
         End If

         Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Iniciando Fechamento Cupom Fiscal"

         CONTA_TENTATIVA = 0

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "SELECT ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
         SQL = SQL & " from LANCAMENTO "
         SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
         SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
         SQL = SQL & " INNER JOIN FORMAPAGTO "
         SQL = SQL & " ON ITEMLANCAMENTO.formapagto_id = FORMAPAGTO.formapagto_id"

         SQL = SQL & " where LANCAMENTO.numr_doc = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF

EFETUANDO_FORMA_DE_PAGAMENTO:

            Msg = "Aguarde, Efetuando Formas PAGTO ..."
            MOSTRA_RODAPE_AQUI Msg, "", "", "", ""
            frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""

            ITEM_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
            Descr_Forma_Pagto = "" & Trim(TabTemp.Fields("descricao").Value)
            If UCase(TabTemp.Fields("descricao").Value) = UCase("Dinheiro") Then _
               Descr_Forma_Pagto = "Dinheiro"

            If USA_TEF = True Then               ' Formas de pagamento que NÃO são de cartão
               If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CARTAO") = 0 Then

                   Parametros = Array(Trim(TabTemp.Fields("descricao").Value), _
                       Replace(Format(TabTemp.Fields("valor_item").Value - ITEM_DESCONTO_N, strFormatacao2Digitos), ",", "."))

                   Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeEfetuarFormaPagamento, Parametros, OperacaoECFOK)

                   ' A variável operacaoECFOK retorna se o comando da ECF foi executado
                   ' com sucesso ou não
                   If Not OperacaoECFOK Then
                       BlockInput False  'Desbloqueia o teclado
                       MsgBox "Não foi possível efetuar a forma de pagamento.", vbCritical
                       Exit Sub
                   End If
                   VALOR_TOTAL_N = VALOR_TOTAL_N + (TabTemp.Fields("valor_item").Value - ITEM_DESCONTO_N)
               End If
               Else            'MsgBox "aqui efetura forma pagto  " & TabTemp.Fields("valor_item").Value
                  VALOR_TOTAL_N = VALOR_TOTAL_N + (TabTemp.Fields("valor_item").Value - ITEM_DESCONTO_N)

                  RETORNO_ECF = Bematech_FI_EfetuaFormaPagamento( _
                            Trim(Left(Descr_Forma_Pagto, 15)), _
                            Replace(Format(TabTemp.Fields("valor_item").Value - ITEM_DESCONTO_N, strFormatacao2Digitos), ",", "."))

                  Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Efetuando Forma de Pagamento"
                     Call VerificaRetornoImpressora("Bematech_FI_EfetuaFormaPagamentoIndice", "", "Emissão de Cupom Fiscal")
            End If

            Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Efetuando Forma de Pagamento"

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         If USA_TEF = True Then
            ' se houve pagamento com cartão
            ' usa o método automático para efetuar as formas de pagamento de maneira
            ' simples, ou seja, somente descrição da forma de pagamento de cartão
            ' e o valor de cada forma de pagamento
            If Not (frmINICIO.EasyTEF.OperacaoTEFAtual = ttCheque) Then
               If Not frmINICIO.EasyTEF.EfetuarFormasPagamentoCartao Then
                  BlockInput False  'Desbloqueia o teclado
                  MsgBox "Não foi possível efetuar a(s) forma(s) de pagamento de cartão.", vbCritical
                  Exit Sub
               End If
            End If
         End If
         CONTA_TENTATIVA = 0

Finalizando_Fechamento_Cupom_Fiscal:
         Msg = "Aguarde, Finalizando cupom fiscal ..."
         MOSTRA_RODAPE_AQUI Msg, "", "", "", ""
         frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", ""

'=======================INDR_LEI_12741
         If INDR_LEI_12741 = True Then
            If VALOR_TOTAL_IMPOSTO > 0 And VALOR_TOTAL_N > 0 Then
               SQL = "Lei 12.741, Valor Aprox. Imposto R$ " & Format(VALOR_TOTAL_IMPOSTO, strFormatacao2Digitos) & _
                     "(" & Format((VALOR_TOTAL_IMPOSTO / VALOR_TOTAL_N), strFormatacao2Digitos) & "%)"
               Mensagem_Final = Mensagem_Final & SQL
            End If
         End If
'=======================

         If USA_TEF = True Then
            Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeTerminarFechamentoCupomFiscal, Array(Mensagem_Final), OperacaoECFOK)

            If Not OperacaoECFOK Then
               BlockInput False  'Desbloqueia o teclado
               MsgBox "Não foi possível terminar o fechamento do cupom fiscal.", vbCritical
               Exit Sub
            End If

            ' imprime todos os cupons tef de transações aprovadas
            Call frmINICIO.EasyTEF.ImprimirCuponsECF

            Else
               RETORNO_ECF = Bematech_FI_TerminaFechamentoCupom(Mensagem_Final)
                  Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Finalizando Fechamento Cupom Fiscal"
                     Call VerificaRetornoImpressora("Bematech_FI_TerminaFechamentoCupom", "", "Emissão de Cupom Fiscal")
                  Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Finalizando Fechamento Cupom Fiscal"
         End If
      End If   'If INDR_PRI = True Then

      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOM = Space(1)
         Else: NUMEROCUPOM = Space(6)
      End If

      NUMR_SEQ_N = 0

LE_ULTIMO_ECF:
Sleep 1000

      Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Ultimo Cupom Impresso"
         RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOM)
      Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Ultimo Cupom Impresso"

      If NUMEROCUPOM = "" Then
         If Not IsNumeric(NUMEROCUPOM) Then
            MsgBox "Erro na leitura do ultimo cupom impresso.  \" & NUMEROCUPOM
            NUMR_SEQ_N = NUMR_SEQ_N + 1
            If NUMR_SEQ_N < 10 Then
               Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Tentando ler Ultimo Cupom Impresso"
Sleep 1000
               GoTo LE_ULTIMO_ECF
            End If
         End If
      End If

      If Trim(NUMEROCUPOM) = "" Then _
         MsgBox "Atenção, erro de comunicação com impressora. Cupom Fiscal não gravado."

      GRAVA_CUPOM NUMEROCUPOM

      SQL = "update PEDIDO set "
      SQL = SQL & " status = 7 "                            'CUPOM FISCAL
      SQL = SQL & ",NUMERO_CAIXA_CPU = " & NUMERO_CAIXA_CPU 'NUMERO_CAIXA_CPU
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      INDR_VENDA = True

      Me.Caption = "OK, " & Msg & " " & "Fim Impressão"

   BlockInput False  'Desbloqueia o teclado
End Sub
