VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmDISPLAYEMISSOR 
   BackColor       =   &H80000008&
   Caption         =   "Emissor de Documento Fiscal"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DISPLAYEMISSOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleMode       =   0  'User
   ScaleWidth      =   33699.57
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAdmTEF 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8160
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtReg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10800
      TabIndex        =   6
      Text            =   "25"
      Top             =   5520
      Width           =   495
   End
   Begin MSComctlLib.ListView LISTAITEM 
      Height          =   2985
      Left            =   0
      TabIndex        =   2
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
      NumItems        =   8
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
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "NCM"
         Object.Width           =   1877
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedidos 
      Height          =   5385
      Left            =   0
      TabIndex        =   0
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
      NumItems        =   14
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
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "PESSOA_ID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   120
      TabIndex        =   1
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
      Width           =   9495
      _ExtentX        =   16748
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
            Picture         =   "DISPLAYEMISSOR.frx":5C12
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros = "
      Height          =   240
      Left            =   9600
      TabIndex        =   5
      Top             =   5520
      Width           =   1110
   End
End
Attribute VB_Name = "frmDISPLAYEMISSOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents EasyTEF As EasyTEF.EasyTEFDiscado
Attribute EasyTEF.VB_VarHelpID = -1

Private TotalDescontoCielo As Double
Private TotalSaqueCielo As Double
Private BufferTransacoesTEF() As String

Private gerentePersist As GeraXmlNFECupomFiscal.XmlNFECupomFiscal
Private gerentePersistInterface As GeraXmlNFECupomFiscal.IGeraXmlCupomFiscal

Private Declare Function func Lib "C:\Windows\System32\GeraXmlNFECupomFiscal.dll" () As String

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

   Me.Caption = Me.Caption & " - " & Me.Name
   cmdAdmTEF.Visible = False

   If USA_NFC_E = True Then
      If USA_TEF = True Then
         Call CarregarEasyTEF
         ReDim BufferTransacoesTEF(1)

         If TIPO_USUARIO < 4 Or TIPO_USUARIO > 5 Then _
            cmdAdmTEF.Visible = True

      End If
      Set gerentePersist = New GeraXmlNFECupomFiscal.XmlNFECupomFiscal
      Set gerentePersistInterface = gerentePersist
   End If

   PESQUISA_VENDA
   Text1.Text = ""

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
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
            CRITERIO_A = ""
            Else: MsgBox "Não permitido."
         End If

         PESQUISA_VENDA
      Case vbKeyF7
         If Not IsNull(lstPedidos.SelectedItem.Text) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cgccpf from PEDIDO WITH (NOLOCK) "
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

            SQL = "select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.valor_desconto,"
            SQL = SQL & " produto.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
            SQL = SQL & " descricao, situacao_tributaria, codg_ncm "

            SQL = SQL & " from PEDIDO WITH (NOLOCK) "
            SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
            SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
            SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
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
               item.SubItems(7) = "" & Trim(TabTemp.Fields("codg_ncm").Value)

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
               CRITERIO_A = ""
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
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
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

         SQL = "select cgccpf from PEDIDO WITH (NOLOCK) "
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
   TRATA_ERROS Err.Description, Me.Name, "lstPedidos_Click"
End Sub

Private Sub lstPedidos_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedidos.SelectedItem.ListSubItems.item(10).Text) Then
      If Trim(lstPedidos.SelectedItem.ListSubItems.item(10).Text) <> "" Then

         PEDIDO_ID_N = lstPedidos.SelectedItem.ListSubItems.item(10).Text

         TIPO_PEDIDO_A = ""
         If Not IsNull(lstPedidos.SelectedItem.ListSubItems.item(3).Text) Then _
            If Trim(lstPedidos.SelectedItem.ListSubItems.item(3).Text) <> "" Then _
               TIPO_PEDIDO_A = "" & lstPedidos.SelectedItem.ListSubItems.item(3).Text

         '================================== PEDIDO DE VENDA
         If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("R") Or _
            UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("OS") Or _
            UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("P") Then
            TIPO_NFe_GERAR = "R"          'Tipo Saida
            INDR_Tela_Chamada_NFC = Me.Name

            If Not IsNull(lstPedidos.SelectedItem.Text) Then
               If Trim(lstPedidos.SelectedItem.Text) <> "" Then
                  PEDIDO_ID_N = lstPedidos.SelectedItem.Text

                  If Trim(lstPedidos.SelectedItem.ListSubItems.item(12).Text) <> "" Then _
                     If IsNumeric(lstPedidos.SelectedItem.ListSubItems.item(12).Text) Then _
                        FAZ_RECEBIMENTO lstPedidos.SelectedItem.ListSubItems.item(12).Text
               End If
            End If
         End If

         If USA_DOC_FISCAL = True Then
            If USA_NFe = True Then
               '================================== DEVOLUÇÃO DE ENTRADA
               If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("DC") Then
                  TIPO_NFe_GERAR = "DC"
   
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close
   
                  SQL = "select * from PEDIDO WITH (NOLOCK) "
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabCabeca.EOF Then
                     If TabCabeca!STATUS = 2 Then
                        CFOP_A = ""
                        DESCRICAO_CFOP_A = ""
   
                        Msg = "Processar Devolução de Compra ?"
                        PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
                        Msg = ""
                        If RESPOSTA = vbYes Then _
                           frmNOTAGERA.Show 1
                        Msg = ""
                     End If
                  End If
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close
               End If
               '================================== DEVOLUÇÃO DE SAIDA
               If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("DV") Then
                  TIPO_NFe_GERAR = "DV"          'DEVOLUÇÃO VENDA
   
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close

NF_ID_N = 0 & Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(10).Text)

                  SQL = "select * from NF WITH (NOLOCK) "
                  SQL = SQL & " where nf_id = " & NF_ID_N
                  'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabCabeca.EOF Then
                     If TabCabeca!STATUS = "A" Then
                        PESSOA_ID_N = 0 & TabCabeca.Fields("pessoa_id").Value

                        CFOP_A = ""
                        DESCRICAO_CFOP_A = ""
                        Msg = "Processar Devolução de Venda ?"
                        PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
                        Msg = ""
                        If RESPOSTA = vbYes Then
                           CRITERIO_A = "" & TabCabeca.Fields("numr_nota").Value

GERA_FINANC_DEVOLUCAO CRITERIO_A
PEDIDO_ID_N = 0

                           frmNOTAGERA.Show 1
                        End If
                        Msg = ""
                     End If
                  End If
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close
               End If
            End If
         End If
      End If
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

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
   TRATA_ERROS Err.Description, Me.Name, "lstPedidos_DblClick"
End Sub

Private Sub LISTAITEM_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE_AQUI " ESC-Sair", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento", ""
   LISTAITEM.Visible = False

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "LISTAITEM_DblClick"
End Sub

Private Sub PESQUISA_VENDA()
'On Error GoTo ERRO_TRATA

   SETA_GRID_VENDA
   SETA_GRID_DIVERSAS

   MOSTRA_RODAPE_AQUI " ESC-Sair", "F6-Cancelar", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento | F11-Imprimir Pedido"

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "PESQUISA_VENDA"
End Sub

Private Sub IMPRIME_TELA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from PEDIDO WITH (NOLOCK) "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status in (2)" 'gerado somente Pedido"
   SQL = SQL & " and tipo_registro in ('S','R','D') "
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
   TRATA_ERROS Err.Description, Me.Name, "IMPRIME_TELA"
End Sub

Private Sub SETA_GRID_VENDA()
'On Error GoTo ERRO_TRATA

   lstPedidos.ListItems.Clear
   NUMR_SEQ_N = 0
   NUMR_CONSULTA_N = 0
   TABELAPRECO_ID_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   If Trim(txtReg.Text) = "" Then _
      txtReg.Text = 25

   If Not IsNumeric(txtReg.Text) Then _
      txtReg.Text = 25

   SQL = "select top(" & txtReg.Text & ") *, PEDIDOFATURA_ID, TABELAPRECO_ID, FORMAPAGTO_ID, TIPOVENDA_ID"
   SQL = SQL & " from PEDIDO WITH (NOLOCK) "

   SQL = SQL & " INNER JOIN PEDIDOFATURA "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOFATURA.PEDIDO_ID"

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status in (2,8) " 'gerado somente Pedido e encomendas
   SQL = SQL & " and tipo_registro in ('S','R','DC','DV','OS','P') "
   SQL = SQL & " order by dt_req DESC "
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      CRITERIO_A = ""
      PEDIDO_ID_N = 0 & TabCabeca.Fields("pedido_id").Value

      txtCNPJCPF.PromptInclude = False
      If Len(Trim(TabCabeca.Fields("cgccpf").Value)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If

      txtCNPJCPF.Text = TabCabeca.Fields("cgccpf").Value
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = TabCabeca.Fields("cgccpf").Value

'========================================cliente
      If Not IsNull(TabCabeca.Fields("nome_cliente").Value) Then
         If Trim(TabCabeca.Fields("nome_cliente").Value) <> "" Then
            CRITERIO_A = Trim(TabCabeca!NOME_CLIENTE)
            Else: TRAZ_NOME_CLIENTE (TabCabeca.Fields("CLIENTE_ID").Value)
         End If
         Else: TRAZ_NOME_CLIENTE (TabCabeca.Fields("CLIENTE_ID").Value)
      End If

'========================================setando grid
      Set item = lstPedidos.ListItems.Add(, "seq." & NUMR_SEQ_N, Trim(TabCabeca.Fields("pedido_id").Value))

      item.SubItems(1) = "" & txtCNPJCPF.Text
      item.SubItems(2) = "" & CRITERIO_A

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

TABELAPRECO_ID_N = 0 & TabCabeca.Fields("tabelapreco_id").Value

      'If TABELAPRECO_ID_N > 0 Then
      '   SQL = "select descricao from FORMAPAGTO WITH (NOLOCK) "
      '   SQL = SQL & " where FORMAPAGTO_id = " & TabCABECA.Fields("tipovenda_id").Value
      '   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      '   If Not TabConsulta.EOF Then _
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               item.SubItems(3) = "" & TabConsulta.Fields(0).Value
      '   Else
            SQL = "select descricao from TIPOVENDA WITH (NOLOCK) "
            SQL = SQL & " where tipovenda_id = " & TabCabeca.Fields("tipovenda_id").Value
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               If Not IsNull(TabConsulta.Fields(0).Value) Then _
                  item.SubItems(3) = "" & TabConsulta.Fields(0).Value
      'End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PERC_DESCONTO_N = 0 & TabCabeca.Fields("perc_desc").Value
      VALOR_DESCONTO_N = 0 & TabCabeca.Fields("valor_desconto").Value
      VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N

'========================================parceiro, tem que ver se pega pelo valor do desconto ou percentual
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK) "
      SQL = SQL & " where pedido_id = " & TabCabeca.Fields("pedido_id").Value
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

      SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK) "
      SQL = SQL & " where pedido_id = " & TabCabeca.Fields("pedido_id").Value
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
      item.SubItems(7) = "" & Trim(TabCabeca!DT_REQ)

'========================================
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select descricao from vwVendedor WITH (NOLOCK) "
      SQL = SQL & " where vendedor_id = " & TabCabeca!VENDEDOR_ID
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then _
         item.SubItems(8) = "" & TabVENDEDOR!DESCRICAO
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
'========================================

      item.SubItems(9) = "" & TabCabeca!TIPO_REGISTRO
      item.SubItems(10) = "" & TabCabeca.Fields("pedido_id").Value

      item.SubItems(12) = ""
      SQL = "select formapagto_id from PEDIDOFATURA WITH (NOLOCK) "
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         item.SubItems(12) = "" & TabVENDEDOR.Fields(0).Value
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      item.SubItems(13) = "" & TRAZ_ID_TABELA("CLIENTE", "PESSOA_ID", "CGCCPF", TabCabeca.Fields("CGCCPF").Value)

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      CONT_N = CONT_N + 1
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      If Trim(UCase(TabCabeca.Fields("tipo_registro").Value)) = "DV" Then
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
      If Trim(UCase(TabCabeca.Fields("tipo_registro").Value)) = "DC" Then
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
      If Trim(UCase(TabCabeca.Fields("tipo_registro").Value)) = "OS" Then
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
      If Trim(TabCabeca.Fields("cgccpf").Value) <> "99999999999" Then
         If TabVENDEDOR.State = 1 Then _
            TabVENDEDOR.Close
         SQL = "select usuario_id, funcionario from USUARIO WITH (NOLOCK) "
         SQL = SQL & " where cpf = '" & Trim(TabCabeca.Fields("cgccpf").Value) & "'"
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

      TabCabeca.MoveNext
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstPedidos.Refresh

   MOSTRA_TOTAIS

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_VENDA"
End Sub

Private Sub CHECA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   STATUS_A = ""
   SQL = "select * from PEDIDOITEM WITH (NOLOCK) "
   SQL = SQL & " where pedido_id = " & lstPedidos.SelectedItem.Text
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      SP_PROCURA_PRODUTO EMPRESA_ID_N, Trim(TabPedidoItem!Codg_Produto), 0, "", FORNEC_ID_N, "", 1
      If Not TabProduto.EOF Then _
         QTDE_ESTOQUE_N = TabProduto!QTD 'Recebe so qtd. porque ja esta retido no pedido
      If TabProduto.State = 1 Then _
         TabProduto.Close

      If QTDE_ESTOQUE_N < TabPedidoItem!QTD_PEDIDA Then _
         STATUS_A = "V"

      TabPedidoItem.MoveNext
   Wend
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close
    
   If STATUS_A = "V" Then 'status de que  a Itens sem Quantidade!
      MsgBox "Pedido com Items Aquardando ordem de Producao , Impossivel Emitir nota!"
      LISTAITEM.Refresh
      Else
         If Not IsNull(lstPedidos.SelectedItem.Text) Then
            If Trim(lstPedidos.SelectedItem.Text) <> "" Then
               PEDIDO_ID_N = lstPedidos.SelectedItem.Text

               If Trim(lstPedidos.SelectedItem.ListSubItems.item(12).Text) <> "" Then _
                  If IsNumeric(lstPedidos.SelectedItem.ListSubItems.item(12).Text) Then _
                     FAZ_RECEBIMENTO lstPedidos.SelectedItem.ListSubItems.item(12).Text
            End If
         End If
   End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "CHECA_ESTOQUE"
End Sub

Public Sub FAZ_RECEBIMENTO(TIPO_VENDA_N As Integer)
'On Error GoTo ERRO_TRATA

   Dim TabPedido As New ADODB.Recordset
   INDR_VENDA = False

   INDR_RECEITA = 1

   If INDR_FORM_ABERTO = True Then
      Unload frmFatura
      INDR_FORM_ABERTO = False
   End If
'===================================
   If TIPO_VENDA_N > 0 Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select contabiliza from TIPOVENDA WITH (NOLOCK) "
      SQL = SQL & " where tipovenda_id = " & TIPO_VENDA_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
            If TabTemp.Fields("contabiliza").Value = True Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               frmFatura.Show 1

               BlockInput False  'Desbloqueia o teclado

               Else
                  SQL = "update PEDIDO set "
                  SQL = SQL & "status = 6 " 'não contabiliza
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
         Else: MsgBox "Tipo venda não informado."
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
'===================================
   BlockInput False  'Desbloqueia o teclado
   If INDR_CONTROLA_ESTOQUE = False Then _
      Exit Sub
'===================================
'===================================
   If TabPedido.State = 1 Then _
      TabPedido.Close

   SQL = "select * from PEDIDO WITH (NOLOCK) "
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedido.EOF Then
      PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
      VALOR_RECEBIDO_N = 0 & TabPedido.Fields("valor_recebido").Value

      If TabPedido!STATUS = 5 Then
         CNPJCPF_A = Trim(TabPedido!CGCCPF)

         If USA_DOC_FISCAL = True Then
            RESPOSTA = ""
            INDR_Tela_Chamada_NFC = Me.Name

            'SE TRABALHA COM CUPOM ELETRONICO
            If USA_NFC_E = True Then
               Msg = ""
               If INDR_VENDA_CARTAO = True Then
                  RESPOSTA = vbYes
                  Else
                     Msg = "Deseja Gerar Cupom Eletrônico ?"
                     PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               End If

               If RESPOSTA = vbYes Then
                  frmDISPLAYEMISSOR.ROTINA_NFC
                  Else: RESPOSTA = ""
               End If
            End If
            If RESPOSTA = "" Then
               If Trim(CNPJCPF_A) <> "99999999999" Then
                  RESPOSTA = ""
                  Msg = "Confirma GERAR Nota Fiscal Eletrônica ? "
                  PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                  Msg = ""
                  If RESPOSTA = vbYes Then _
                     If USA_NFe = True Then _
                        frmNOTAGERA.Show 1
                  Msg = ""
               End If
            End If
         End If
         Msg = ""
'====================
ATUALIZA_ESTOQUE 0, PEDIDO_ID_N
'====================
      End If   'If TabPedido!Status = 5 Then
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

Msg = ""

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub

Sub TRAZ_NOME_CLIENTE(CLIENTE_ID As Long)
'On Error GoTo ERRO_TRATA

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select CGCCPF, nome from CLIENTE WITH (NOLOCK) "
   SQL = SQL & " where cliente_id = " & CLIENTE_ID
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      CRITERIO_A = Trim(TabCliente!NOME)

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
   TRATA_ERROS Err.Description, Me.Name, "TRAZ_NOME_CLIENTE"
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

   SQL = "select distinct(TIPOVENDA_ID) "
   SQL = SQL & " from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOFATURA WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOFATURA.PEDIDO_ID"

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status = 2" 'gerado somente Pedido"
   SQL = SQL & " and tipo_registro in ('S','R','D','OS') "
   SQL = SQL & " order by TIPOVENDA_ID "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum(valor_total-valor_desconto) from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDOFATURA WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOFATURA.PEDIDO_ID"

      SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and tipovenda_id = " & TabTemp.Fields(0).Value
      SQL = SQL & " and status = 2" 'gerado somente Pedido"
      SQL = SQL & " and tipo_registro in ('S','R','D','OS') "
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
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAIS"
   INDR_VENDA = False
   CRITERIO_A = ""
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
   TRATA_ERROS Err.Description, "mdlGeral", "MOSTRA_RODAPE_AQUI"
End Sub
'================================
Sub ROTINA_NFC()
'On Error GoTo ERRO_TRATA

'EXCLUIR_400

   Msg = "Aguarde, Gerando NFC-e ..............."
   If INDR_Tela_Chamada_NFC = "frmDISPLAYEMISSOR" Then
      MOSTRA_RODAPE_AQUI Msg, "", "", "", ""
      Else: frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & "", "", "", "", ""
   End If

   If PESSOA_ID_N = 0 Then _
      Exit Sub
   If PEDIDO_ID_N = 0 Then _
      Exit Sub

   Dim TabPedido                 As New ADODB.Recordset
   Dim TabNCM                    As New ADODB.Recordset
   Dim TRANSP_ID_N               As Long
   Dim ID_NF_N                   As Long
   Dim DESC_NATUREZA_OPERACAO_A  As String
   Dim CFOP_N                    As String
   Dim NOME_CLIENTE_A            As String
   Dim ALIQ_IBPT_N               As Long
   Dim VALOR_TOTAL_IMPOSTO_N     As Double
   Dim NUMR_DOC_N                As String
   Dim MSG_MetodoGeraXmlNfeCupomFiscalCOM                  As String
   Dim MFASEQUENCIA_N            As Long
   Dim EMPRESSA_A                As String
   Dim ESTABELECIMENTO_A         As String
   Dim IMPOSTO_A                 As String
   Dim MFACODSTAT_A              As String
   Dim VALIDA_NCM_B              As Boolean

   CLIENTE_ID_N = 0
   ID_NF_N = 0
   VALOR_TOTAL_IMPOSTO_N = 0
   IMPOSTO_A = ""
   TRANSP_ID_N = 0
   DESC_NATUREZA_OPERACAO_A = ""
   CFOP_N = ""
   NOME_CLIENTE_A = ""
   VALIDA_NCM_B = False

   If TabPedido.State = 1 Then _
      TabPedido.Close
   SQL = "select PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.ESTABELECIMENTO_ID, "
   SQL = SQL & " PEDIDO.VENDEDOR_ID, PEDIDO.CGCCPF, PEDIDO.USUARIO_ID, PEDIDO.DT_REQ, "
   SQL = SQL & " PEDIDO.STATUS AS StatusPedido, PEDIDO.TIPO_REGISTRO, PEDIDO.NOME_CLIENTE, PEDIDOITEM.SEQ_ID,"
   SQL = SQL & " PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.CFOP_ID, "
   SQL = SQL & " PEDIDOITEM.STATUS AS StatusItem, PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.VALOR_DESCONTO, "
   SQL = SQL & " PEDIDOITEM.TIPO_REG, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, Produto.SITUACAO_TRIBUTARIA, "
   SQL = SQL & " Produto.Aliquota_Icms, Produto.CODG_NCM, PRODUTO.situacao"
   SQL = SQL & " from PEDIDO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N

   TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedido.EOF Then
      'PESSOA_ID_N = 0 & TabPedido.Fields("pessoa_id").Value
      CLIENTE_ID_N = 0 & TabPedido.Fields("cliente_id").Value
      NOME_CLIENTE_A = "" & Trim(TabPedido.Fields("NOME_CLIENTE").Value)
      CFOP_N = "" & TabPedido.Fields("cfop_id").Value

      'SE FOR DIFERENTE DE CONSUMIDOR FINAL INTEGRA
      If Trim(Left(TabPedido.Fields("cgccpf").Value, 11)) <> "99999999999" Then _
         Call frmINTEGRA.CLIENTE_INTEGRA(TabPedido.Fields("cgccpf").Value)

      While Not TabPedido.EOF
         If Trim(TabPedido.Fields("situacao").Value) = "A" Then
            If Not IsNull(TabPedido.Fields("codg_ncm").Value) Then
               If Trim(TabPedido.Fields("codg_ncm").Value) <> "" Then
                  If Trim(Len(TabPedido.Fields("codg_ncm").Value)) < 8 Then
                     VALIDA_NCM_B = False
                     MsgBox "Produto: " & Trim(TabPedido.Fields("codg_produto").Value) & "-" & Trim(TabPedido.Fields("descricao").Value) & " ; NCM incorreto : " & Trim(TabPedido.Fields("codg_ncm").Value)
                     Exit Sub
                     Else: VALIDA_NCM_B = True
                  End If
               End If
            End If
         End If
         TabPedido.MoveNext
      Wend

      If VALIDA_NCM_B = True Then
         NUMR_DOC_N = "" & GERA_NUMERO_NFC_N("NFC")

         GRAVA_NOTA NUMR_DOC_N, "1", "NFC", "P", "", "", "", "1", "1", CFOP_N, ""

         Else: Exit Sub
      End If
   End If

'=========
      SQL = "insert into PEDIDOTIME "
         SQL = SQL & "(PEDIDO_ID,DT_IN,DT_FIM,TIPO_DOC,NUMR_DOC)"
      SQL = SQL & " values("
         SQL = SQL & PEDIDO_ID_N
         SQL = SQL & ",'" & Now & "'"
         SQL = SQL & ",''"
         SQL = SQL & ",'NFC'"
         SQL = SQL & "," & NUMR_DOC_N
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
'=========

   If TabPedido.State = 1 Then _
      TabPedido.Close
   SQL = "select NFITEM.NF_ID, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID, qtde,valor"
   SQL = SQL & " from NF WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN NFITEM WITH (NOLOCK) "
   SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedido.EOF
      If TabNCM.State = 1 Then _
         TabNCM.Close

      SQL = "select codg_ncm from PRODUTO WITH (NOLOCK) "
      SQL = SQL & " where produto_id = " & TabPedido.Fields("produto_id").Value
      TabNCM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNCM.EOF Then
         If Not IsNull(TabNCM.Fields("codg_ncm").Value) Then
            If Trim(TabNCM.Fields("codg_ncm").Value) <> "" Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select ALIQNAC,ALIQIMP from IBPTax WITH (NOLOCK) "
               SQL = SQL & " where codg_ncm = '" & Trim(TabNCM.Fields("codg_ncm").Value) & "'"
               'SQL = SQL & " and tabela = " & 1 'ORIGEM_MERCADORIA_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqimp").Value

                  'If ORIGEM_MERCADORIA_N = 0 Then _
                     ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqnac").Value
                  'If ORIGEM_MERCADORIA_N = 1 Then _
                     ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqimp").Value

                  VALOR_TOTAL_IMPOSTO_N = VALOR_TOTAL_IMPOSTO_N + ((TabPedido.Fields("valor").Value * TabPedido.Fields("qtde").Value) * ALIQ_IBPT_N / 100)

               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close
            End If
         End If
      End If
      If TabNCM.State = 1 Then _
         TabNCM.Close

      ID_NF_N = 0 & TabPedido.Fields("nf_id").Value

      frmINTEGRA.INTEGRA_PRODUTO (TabPedido.Fields("produto_id").Value)

      TabPedido.MoveNext
   Wend
   If TabPedido.State = 1 Then _
      TabPedido.Close

'aqui é só para o cabeçalho, mas beleza
   DESC_NATUREZA_OPERACAO_A = ""
   If Trim(CFOP_N) = "6102" Then _
      CFOP_N = "5102"

   SQL = "select distinct(cfop_id) from NFITEM WITH (NOLOCK) "
   SQL = SQL & " where nf_id = " & ID_NF_N
   TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedido.EOF Then
      CFOP_N = "" & TabPedido.Fields(0).Value

      If Trim(CFOP_N) = "" Then _
         CFOP_N = "5102"

      If TabPedido.State = 1 Then _
         TabPedido.Close

         SQL = "select descricao from CFOP WITH (NOLOCK) "
         SQL = SQL & " where cfop_id = " & CFOP_N
         TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedido.EOF Then _
            DESC_NATUREZA_OPERACAO_A = "" & Trim(TabPedido.Fields(0).Value)

      If TabPedido.State = 1 Then _
         TabPedido.Close
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

   IMPOSTO_A = "" & Trim("Tributos Totais Incidentes(Lei Federal 12.741/2012): R$ " & Format(VALOR_TOTAL_IMPOSTO_N, strFormatacao2Digitos))

   TRANSP_ID_N = 0 & TRAZ_ID_TABELA("vwTRANSPORTADORA", "transp_id", "cnpjcpf", CNPJ_EMPRESA_N)

   Call frmINTEGRA.PEDIDO_INTEGRA_MFA010(ID_NF_N, _
                                         TRANSP_ID_N, _
                                         "NFC", _
                                         IMPOSTO_A, _
                                         "1", _
                                         "1", _
                                         "1", _
                                         "", _
                                         "1", _
                                         "0", _
                                         "0", _
                                         "9", _
                                         DESC_NATUREZA_OPERACAO_A, _
                                         "0", _
                                         "0", _
                                         "0", _
                                         "1", _
                                         "N", _
                                           0)

   MFASEQUENCIA_N = 0
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then
      MsgBox "Banco GLOBAL não conectado."
      Exit Sub
   End If

   'EMPRESSA_A = "0" & EMPRESA_ID_N
   EMPRESSA_A = "01"
   ESTABELECIMENTO_A = "0" & ESTABELECIMENTO_ID_N

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MFASEQUENCIA from MFA010 WITH (NOLOCK) "
   SQL = SQL & " where mfadoc = '" & Trim(NUMR_DOC_N) & "'"
   SQL = SQL & " and mfaprefixo = 'NFC'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         MFASEQUENCIA_N = 0 & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MSG_MetodoGeraXmlNfeCupomFiscalCOM = ""

''''''''''''''''''''''''JÁ PASSOU DA INTEGRAÇÃO

HORA_INI = Time

   If MFASEQUENCIA_N > 0 Then

         Set gerentePersist = New GeraXmlNFECupomFiscal.XmlNFECupomFiscal
         Set gerentePersistInterface = gerentePersist

         MSG_MetodoGeraXmlNfeCupomFiscalCOM = gerentePersistInterface.MetodoGeraXmlNfeCupomFiscalCOM(MFASEQUENCIA_N, _
                                                                           "NFC", _
                                                                           EMPRESSA_A, _
                                                                           ESTABELECIMENTO_A)

      Else: MsgBox "Sequencia não encontrada. " & MFASEQUENCIA_N
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

'MsgBox "VOLTOU DO GERAXML  :  " & MSG_MetodoGeraXmlNfeCupomFiscalCOM

   HORA_FIM = Time

'=========
   SQL = "select * from PEDIDOTIME WITH (NOLOCK) "
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and numr_doc = " & NUMR_DOC_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into PEDIDOTIME "
         SQL = SQL & "(PEDIDO_ID,DT_IN,DT_FIM,TIPO_DOC,NUMR_DOC)"
      SQL = SQL & " values("
         SQL = SQL & PEDIDO_ID_N
         SQL = SQL & ",'" & Now & "'"
         SQL = SQL & ",''"
         SQL = SQL & ",'NFC'"
         SQL = SQL & "," & NUMR_DOC_N
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "update PEDIDOTIME set "
            SQL = SQL & " dt_fim = '" & Now & "'"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and numr_doc = " & NUMR_DOC_N
         CONECTA_RETAGUARDA.Execute SQL
   End If
'=========

'MsgBox "Duração da consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss")

   Msg = ""
   If INDR_Tela_Chamada_NFC = "frmDISPLAYEMISSOR" Then
      MOSTRA_RODAPE_AQUI Msg, "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")
      Else: frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg & " ...", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")
   End If
INDR_VENDA_CARTAO = False
   Msg = ""
   MFACODSTAT_A = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select mfacodstat from MFA010 WITH (NOLOCK) "
   SQL = SQL & " where mfadoc = '" & Trim(NUMR_DOC_N) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         MFACODSTAT_A = "" & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

SQL3 = "" & MFACODSTAT_A

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select motivo from MENSAGEMSEFAZ WITH (NOLOCK) "
   SQL = SQL & " where erro_id = " & MFACODSTAT_A
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         MFACODSTAT_A = "" & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(SQL3) <> "100" Then _
      MsgBox "Retorno : " & MSG_MetodoGeraXmlNfeCupomFiscalCOM & " : " & MFACODSTAT_A & " ; Processado em " & Format((HORA_FIM - HORA_INI), "hh:mm:ss")

'EXCLUIR_400

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "ROTINA_NFC"
   INDR_VENDA = False
   SQL = ""
   SqL2 = ""
   SQL3 = ""
   CRITERIO_A = ""
   INDR_VENDA_CARTAO = False
End Sub

Function BUSCA_IntPos() As Boolean
'On Error GoTo ERRO_TRATA

   BUSCA_IntPos = False
   Text1.Text = ""

   If Not FSO.FileExists(Path_IntPos_Entrada_A & "IntPos.001") Then
      MsgBox "Arquivo de inicialização do sistema não encontrado, entre em contato com suporte."
      End
   End If

   Dim f
   Dim sLine
   Dim Tipo_Operacao
   Dim INDR_VAI      As Boolean
   Dim NSU_A         As String
   Dim CODG_TRANSA   As String
   Dim sComando
   Dim iRetorno

   f = FreeFile
   INDR_PRI = True
   Tipo_Operacao = ""
   INDR_VAI = False
   NSU_A = ""

   If Not FSO.FolderExists(Path_IntPos_Saida_A) Then _
      FSO.CreateFolder (Path_IntPos_Saida_A)

   Open Path_IntPos_Entrada_A & "IntPos.001" For Input As f
   'Open App.Path & "C:\TEF_DISC\RESP\IntPos.001" For Input As f

   Do While Not EOF(f)
      Line Input #f, sLine

      If INDR_PRI = True Then
         INDR_PRI = False
         Tipo_Operacao = "" & Trim(Mid(sLine, 11, 3))
      End If

      If Tipo_Operacao = "CRT" Then
         If Trim(Mid(sLine, 1, 3)) = "012" Then
            NSU_A = Trim(Mid(sLine, 11, Len(sLine) - 1))
         End If
         If Trim(Mid(sLine, 1, 3)) = "013" Then _
            CODG_TRANSA = Trim(Mid(sLine, 11, Len(sLine) - 11))

         If Trim(Mid(sLine, 1, 3)) = "029" And INDR_VAI = False Then
            INDR_VAI = True

            Open Path_IntPos_Saida_A & NSU_A & ".txt" For Output As #2
         End If

         'ACIONAMENTO DA GUILHOTINA
         If Trim(Mid(sLine, 1, 3)) = "712" Then
            'Para finalizar nosso arquivo de exemplo, podemos enviar o comando de corte de papel,
            'para isso inserimos na última linha o comando <ALT> 27 mais <ALT> 119.

            'sComando = Chr(27) + Chr(119)

            'iRetorno = ComandoTX(sComando, Len(sComando))
            Print #2, sComando
            GoTo SAI_PRI_COMPROVANTE
         End If

         If INDR_VAI = True And Trim(Mid(sLine, 1, 3)) <> "710" Then
            If Trim(Mid(sLine, 1, 3)) <> "999" Then
               Print #2, Mid(sLine, 12, Len(sLine) - 12)
               Else: Print #2, Trim(Mid(sLine, 12, Len(sLine) - 1))
            End If
         End If
      End If

      DoEvents
   Loop

   Print #2, ""

SAI_PRI_COMPROVANTE:

   Close #f
   Close #1
   Close #2

   Open Path_IntPos_Saida_A & NSU_A & ".txt" For Input As #1
   Text1.Text = Input(LOF(1), #1)

   Close
   
   'If MsgBox("Confirma Impressão do Arquivo ? ", vbYesNo, "Imprimindo um arquivo com Print") = vbYes Then
      Printer.Print Text1.Text
      Printer.EndDoc
   'End If
   Text1.Text = ""


'=============== CLIENTE
   INDR_VAI = False

   Open Path_IntPos_Entrada_A & "IntPos.001" For Input As f
   'Open App.Path & "C:\TEF_DISC\RESP\IntPos.001" For Input As f

   Do While Not EOF(f)
      Line Input #f, sLine

      If INDR_PRI = True Then
         INDR_PRI = False
         Tipo_Operacao = "" & Trim(Mid(sLine, 11, 3))
      End If

      If Tipo_Operacao = "CRT" Then
         If Trim(Mid(sLine, 1, 3)) = "012" Then
            NSU_A = Trim(Mid(sLine, 11, Len(sLine) - 1))
         End If
         If Trim(Mid(sLine, 1, 3)) = "013" Then _
            CODG_TRANSA = Trim(Mid(sLine, 11, Len(sLine) - 11))

         'ACIONAMENTO DA GUILHOTINA
         If Trim(Mid(sLine, 1, 3)) = "712" Then
            Open Path_IntPos_Saida_A & NSU_A & "CLI.txt" For Output As #2
            INDR_VAI = True
         End If

         If INDR_VAI = True Then
            If Trim(Mid(sLine, 1, 3)) <> "999" Then
               Print #2, Mid(sLine, 12, Len(sLine) - 12)
               Else: Print #2, Trim(Mid(sLine, 12, Len(sLine) - 1))
            End If
         End If
      End If

      DoEvents
   Loop

   Print #2, ""

   Text1.Text = ""
   Close #f
   Close #1
   Close #2

   Open Path_IntPos_Saida_A & NSU_A & "CLI.txt" For Input As #1
   Text1.Text = Input(LOF(1), #1)

   Close
   
   'If MsgBox("Confirma Impressão do Arquivo ? ", vbYesNo, "Imprimindo um arquivo com Print") = vbYes Then
      Printer.Print Text1.Text
      Printer.EndDoc
   'End If
   Text1.Text = ""

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "BUSCA_IntPos"
End Function

Sub IMPRIME_COMPROVANTES_TEF()
'On Error GoTo ERRO_TRATA

   Text1.Text = ""

   If Not FSO.FileExists(Path_IntPos_Entrada_A & "IntPos.001") Then
      MsgBox "Arquivo de inicialização do sistema não encontrado, entre em contato com suporte."
      End
   End If

   Dim f
   Dim sLine
   Dim Tipo_Operacao
   Dim INDR_VAI      As Boolean
   Dim NSU_A         As String
   Dim CODG_TRANSA   As String
   Dim sComando
   Dim iRetorno

   f = FreeFile
   INDR_PRI = True
   Tipo_Operacao = ""
   INDR_VAI = False
   NSU_A = ""

   If Not FSO.FolderExists(Path_IntPos_Saida_A) Then _
      FSO.CreateFolder (Path_IntPos_Saida_A)

   Open Path_IntPos_Entrada_A & "IntPos.001" For Input As f
   'Open App.Path & "C:\TEF_DISC\RESP\IntPos.001" For Input As f

   Do While Not EOF(f)
      Line Input #f, sLine

      If INDR_PRI = True Then
         INDR_PRI = False
         Tipo_Operacao = "" & Trim(Mid(sLine, 11, 3))
      End If

      If Tipo_Operacao = "CRT" Then
         If Trim(Mid(sLine, 1, 3)) = "012" Then
            NSU_A = Trim(Mid(sLine, 11, Len(sLine) - 1))
         End If
         If Trim(Mid(sLine, 1, 3)) = "013" Then _
            CODG_TRANSA = Trim(Mid(sLine, 11, Len(sLine) - 11))

         If Trim(Mid(sLine, 1, 3)) = "029" And INDR_VAI = False Then
            INDR_VAI = True

            Open Path_IntPos_Saida_A & NSU_A & ".txt" For Output As #2
         End If

         'ACIONAMENTO DA GUILHOTINA
         If Trim(Mid(sLine, 1, 3)) = "712" Then
            'Para finalizar nosso arquivo de exemplo, podemos enviar o comando de corte de papel,
            'para isso inserimos na última linha o comando <ALT> 27 mais <ALT> 119.

            'sComando = Chr(27) + Chr(119)

            'iRetorno = ComandoTX(sComando, Len(sComando))
            Print #2, sComando
            GoTo SAI_PRI_COMPROVANTE
         End If

         If INDR_VAI = True And Trim(Mid(sLine, 1, 3)) <> "710" Then
            If Trim(Mid(sLine, 1, 3)) <> "999" Then
               Print #2, Mid(sLine, 12, Len(sLine) - 12)
               Else: Print #2, Trim(Mid(sLine, 12, Len(sLine) - 1))
            End If
         End If
      End If

      DoEvents
   Loop

   Print #2, ""

SAI_PRI_COMPROVANTE:

   Close #f
   Close #1
   Close #2

   Open Path_IntPos_Saida_A & NSU_A & ".txt" For Input As #1
   Text1.Text = Input(LOF(1), #1)

   Close
   
   'If MsgBox("Confirma Impressão do Arquivo ? ", vbYesNo, "Imprimindo um arquivo com Print") = vbYes Then
      Printer.Print Text1.Text
      Printer.EndDoc
   'End If
   Text1.Text = ""


'=============== CLIENTE
   INDR_VAI = False

   Open Path_IntPos_Entrada_A & "IntPos.001" For Input As f
   'Open App.Path & "C:\TEF_DISC\RESP\IntPos.001" For Input As f

   Do While Not EOF(f)
      Line Input #f, sLine

      If INDR_PRI = True Then
         INDR_PRI = False
         Tipo_Operacao = "" & Trim(Mid(sLine, 11, 3))
      End If

      If Tipo_Operacao = "CRT" Then
         If Trim(Mid(sLine, 1, 3)) = "012" Then
            NSU_A = Trim(Mid(sLine, 11, Len(sLine) - 1))
         End If
         If Trim(Mid(sLine, 1, 3)) = "013" Then _
            CODG_TRANSA = Trim(Mid(sLine, 11, Len(sLine) - 11))

         'ACIONAMENTO DA GUILHOTINA
         If Trim(Mid(sLine, 1, 3)) = "712" Then
            Open Path_IntPos_Saida_A & NSU_A & "CLI.txt" For Output As #2
            INDR_VAI = True
         End If

         If INDR_VAI = True Then
            If Trim(Mid(sLine, 1, 3)) <> "999" Then
               Print #2, Mid(sLine, 12, Len(sLine) - 12)
               Else: Print #2, Trim(Mid(sLine, 12, Len(sLine) - 1))
            End If
         End If
      End If

      DoEvents
   Loop

   Print #2, ""

   Text1.Text = ""
   Close #f
   Close #1
   Close #2

   Open Path_IntPos_Saida_A & NSU_A & "CLI.txt" For Input As #1
   Text1.Text = Input(LOF(1), #1)

   Close
   
   'If MsgBox("Confirma Impressão do Arquivo ? ", vbYesNo, "Imprimindo um arquivo com Print") = vbYes Then
      Printer.Print Text1.Text
      Printer.EndDoc
   'End If
   Text1.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIME_COMPROVANTES_TEF"
End Sub
'=============================
'=============================
'=============================
Public Sub CarregarEasyTEF()
   Dim f As New StdFont
   Dim ini As String

    f.Name = "Tahoma"
    f.Size = 9
    ini = App.Path & "\tef.ini"

    Set EasyTEF = New EasyTEF.EasyTEFDiscado
    
    EasyTEF.ImpressaoNaoFiscal = True
    
    EasyTEF.Gerenciador = tgGerenciadorPadrao
    EasyTEF.FormMsgOperador.Fonte = f
    EasyTEF.FormMsgOperador.ALTURA = 110
    EasyTEF.FormMsgOperador.LARGURA = 400
    EasyTEF.FormMsgOperador.BotaoOK.ALTURA = 25
    EasyTEF.FormMsgOperador.BotaoOK.LARGURA = 75
    EasyTEF.ContraSenha = ReadINI(ini, "TEF", "contraSenhaDisc", "")
    EasyTEF.AutoAtivarGerenciador = True
    'se usar tef dedicado D-TEF, com Client D-TEF, setar True

   '1=DICADO ; 2=IP ; 3=DEDICADO
   EasyTEF.UsarAuttar = False
   If TIPO_TEF_N = 2 Then
      If USA_AUTTAR = True Then _
         EasyTEF.UsarAuttar = True
      Else
         'ReadINI(ini, "TEF", "usarDTEF", "0")

         If TIPO_TEF_N = 1 Then _
            EasyTEF.UsarDTEF = ReadINI(ini, "TEF", "usarDTEF", "0")
   End If

    EasyTEF.UsarDTEF = ReadINI(ini, "TEF", "usarDTEF", "0")
    EasyTEF.Somente1RelGerencial = True
    ' configurações para Cielo Premia
   EasyTEF.CieloPremia.RazaoSocialSW = "SERGIO HORACIO FERREIRA 59008148153 ME"
   EasyTEF.CieloPremia.VersaoSW = "Sistema de Gestão Comercial v2.0"

    'EasyTEF.CieloPremia.RazaoSocialSW = "Razão Social da Software House"
    'EasyTEF.CieloPremia.VersaoSW = "Nome da Automação e Versão"
    EasyTEF.CieloPremia.TIPO = tcpAmbas
    
    If Not EasyTEF.AutoVerificarTEF Then
        EasyTEF.AutoVerificarTEF = True
    End If
End Sub

Public Function TratarPagamentoComCartoes(Valores As Variant) As Boolean
Dim resultado As Boolean
Dim valorCartao As Double
Dim i As Integer
Dim Pedido_Id_A As String

    resultado = True
    Pedido_Id_A = "" & PEDIDO_ID_N

    EasyTEF.NumeroDeCartoes = 0

    If IsArray(Valores) Then
        EasyTEF.ImprimirComprovante = False
        EasyTEF.NumeroDeCartoes = UBound(Valores) + 1
        For i = 1 To EasyTEF.NumeroDeCartoes
            valorCartao = Valores(i - 1)

'Call EasyTEF.PagarNoCartao(valorCartao, tmReal,  PegarValor, i = 1, i = EasyTEF.NumeroDeCartoes, "TEF")
'Call EasyTEF.PagarNoCartao(valorCartao, tmReal, valorCartao, i = 1, i = EasyTEF.NumeroDeCartoes, "TEF")
INDR_ERRO_TEF = True
Call EasyTEF.PagarNoCartao(valorCartao, tmReal, Pedido_Id_A, i = 1, i = EasyTEF.NumeroDeCartoes, "Cartao")

            resultado = EasyTEF.TransacaoAprovada
            If Not EasyTEF.TransacaoAprovada Then
               INDR_ERRO_TEF = True
                MsgBox "Não foi possível finalizar com sucesso o pagamento com cartão", _
                    vbCritical
                Exit For
            Else
                TotalDescontoCielo = TotalDescontoCielo + EasyTEF.ValorCampo709_000
                TotalSaqueCielo = TotalSaqueCielo + EasyTEF.ValorCampo708_000
                ReDim Preserve BufferTransacoesTEF(UBound(BufferTransacoesTEF))
                ' nome da rede + NSU + finalização
                BufferTransacoesTEF(UBound(BufferTransacoesTEF)) = EasyTEF.ValorCampo010_000 & ";" _
                    & EasyTEF.ValorCampo012_000 & ";" & EasyTEF.ValorCampo027_000
            End If
        Next i
    End If
    
    TratarPagamentoComCartoes = resultado
End Function

Private Sub EmitirSatOuNFCe()
  ' este método representa o método de seu sistema que gera e transmite
  ' o SAT ou NFC-e

End Sub

Private Sub cmdAdmTEF_Click()

    Screen.MousePointer = vbHourglass
    
    EasyTEF.ImprimirComprovante = True
    Call EasyTEF.FazerRequisicaoAdministrativa
    
    Screen.MousePointer = vbDefault

End Sub

Private Function PegarValor() As Double
    Randomize
    PegarValor = (Int(Rnd * 10000) + 1) / 100
End Function

Private Function PegarSequencial() As String
    Randomize
    PegarSequencial = (Int(Rnd * 100000) + 1)
End Function

Private Function ArrayToStr(a As Variant)
    Dim i As Integer
    Dim s As String
    s = ""
    For i = LBound(a) To UBound(a)
        s = s & a(i) & vbCrLf
    Next i
    
    ArrayToStr = s
End Function

Private Sub EasyTEF_OnGerarIdentificador(identificacao As Long)
    Randomize
    identificacao = Int(Rnd * 1000) + 1
End Sub

Private Sub EasyTEF_OnImpressaoNaoFiscal(ByVal ImagemCupomTEF As Variant, ImpressaoOK As Boolean)

    Call ImprimirCupomEmSuaImpressoraNaoFiscal(ImagemCupomTEF)

    ImpressaoOK = MsgBox("A impressão foi efetuada totalmente com sucesso?", _
        vbYesNo + vbQuestion, "Impressão") = vbYes

End Sub

Public Function TRATA_RECEBIMENTO_CARTAO(Descricao_A As String, Valor_Cartao_n As Double) As Boolean
'On Error GoTo ERRO_TRATA

   If Valor_Cartao_n <= 0 Then _
      Exit Function

   TRATA_RECEBIMENTO_CARTAO = False
   INDR_ERRO_TEF = True

   If Left(UCase(Trim(Descricao_A)), 6) = "CARTAO" Or Left(UCase(Trim(Descricao_A)), 6) = "CARTÃO" Then
      If cmdPagar(Valor_Cartao_n) = True Then 'se  aqui [e true entao validou cartao
         TRATA_RECEBIMENTO_CARTAO = True
         INDR_ERRO_TEF = False
      End If
   End If

BlockInput False  'Desbloqueia o teclado
Exit Function
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "TRATA_RECEBIMENTO_CARTAO "
End Function

Public Function cmdPagar(Valor_Cartao_n As Double) As Boolean

   If Valor_Cartao_n <= 0 Then _
      Exit Function

   cmdPagar = False
   Dim i As Integer
   Dim CountCartoes As Integer
   Dim Cartoes() As Variant
   Dim Formas              As Variant
   Dim Valores             As Variant

'==========
   i = 0
   Formas = Array("")
   Valores = Array("")

   'If TabTemp.State = 1 Then _
      TabTemp.Close

   '   SQL = "select ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
   '   SQL = SQL & " from LANCAMENTO "
   '   SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   '   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   '   SQL = SQL & " INNER JOIN FORMAPAGTO "
   '   SQL = SQL & " ON ITEMLANCAMENTO.formapagto_id = FORMAPAGTO.formapagto_id"

   '   SQL = SQL & " where LANCAMENTO.numr_doc = " & PEDIDO_ID_N
   '   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   '   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   '   While Not TabTemp.EOF
         ' se for uma forma de pagamento de cartão
   '      If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CARTAO") > 0 Then
   '          If i > 0 Then
   '              ReDim Preserve Formas(UBound(Formas) + 1)
   '              ReDim Preserve Valores(UBound(Valores) + 1)
   '          End If

   '          Formas(i) = Left(UCase(Trim(TabTemp.Fields("descricao").Value)), 6)
   '          Valores(i) = Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
   '          i = i + 1
   '      End If

   '      TabTemp.MoveNext
   '   Wend
   'If TabTemp.State = 1 Then _
      TabTemp.Close

   '===========================================================
   'Call EmitirSatOuNFCe
   'CountCartoes = txtQtdCartoes.Text
   CountCartoes = 1

   ReDim Cartoes(CountCartoes - 1)

   Dim VALOR_ZERO As Double
   'VALOR_ZERO = 0 & Valores(0)
   VALOR_ZERO = 0 & Valor_Cartao_n

   If VALOR_ZERO <= 0 Then
      MsgBox "Valor Cartao nao informado."
      Else
         For i = 0 To CountCartoes - 1
            Cartoes(i) = VALOR_ZERO
         Next i

         INDR_ERRO_TEF = True
         Call EasyTEF.IniciarTransacaoTEF

         If TratarPagamentoComCartoes(Cartoes) Then
      
             For i = LBound(EasyTEF.ValoresCartoes) To UBound(EasyTEF.ValoresCartoes)
                 Call ImprimirCupomEmSuaImpressoraNaoFiscal(EasyTEF.CuponsDisponiveis(i))
             Next i

             If INDR_ERRO_TEF = False Then
               'If MsgBox("A impressão foi efetuada totalmente com sucesso?", vbYesNo + vbQuestion, "Impressão") = vbYes Then

               cmdPagar = True

               Call EasyTEF.ConfirmacaoVendaImpressaoCupom(EasyTEF.ValorCampo010_000, EasyTEF.ValorCampo012_000, EasyTEF.ValorCampo027_000, EasyTEF.ValorCampo002_000)

               Else
                  Call EasyTEF.CancelarVendasPendentes
                  INDR_ERRO_TEF = True
             End If

             Call EasyTEF.FinalizarTransacaoTEF
             TotalDescontoCielo = 0
             TotalSaqueCielo = 0
         End If

         ReDim bufferInfoTransacoesTEF(0)
   End If

End Function

Private Sub ImprimirCupomEmSuaImpressoraNaoFiscal(Cupom As Variant)
    ' aqui deve ser adicionado o comando de qualquer impressora não fiscal que
    ' fará a impressão do Cupom TEF

   'If INDR_ERRO_TEF = False Then
      'Call MsgBox(ArrayToStr(Cupom), vbInformation)

      Text1.Text = "" & ArrayToStr(Cupom)
   
      Printer.Print Text1.Text
      Printer.EndDoc
   
      INDR_ERRO_TEF = False
      INDR_VENDA_CARTAO = True

   'End If
End Sub

Sub GERA_FINANC_DEVOLUCAO(Numero_NF_A As String)
'On Error GoTo ERRO_TRATA

'MsgBox Trim(lstPedidos.SelectedItem.ListSubItems.item(1).Text)
'MsgBox Trim(lstPedidos.SelectedItem.ListSubItems.item(13).Text)
   If Trim(Numero_NF_A) = "" Then
      MsgBox "Numero Nota não informado!!!"
      Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where tipovenda_id = 9999"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      'CABEÇA lançamento
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      SQL = "select * from LANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " where numr_doc = " & Trim(Numero_NF_A)
      SQL = SQL & " and tipo_lancamento = 1"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         NUMR_ID_N = TabLancamento!LANCAMENTO_ID
         Else
            NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

            SQL = "INSERT INTO LANCAMENTO "
            SQL = SQL & " ("
               SQL = SQL & " Lancamento_id, Numr_doc, dt_cad, Tipo_Lancamento, tipovenda_id,pessoa_id,estabelecimento_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & Trim(Numero_NF_A)
               SQL = SQL & ",'" & Date & "'"
               SQL = SQL & "," & 1
               SQL = SQL & "," & 9999
               SQL = SQL & "," & PESSOA_ID_N
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      SQL3 = PEDIDO_ID_N
      SqL2 = EMPRESA_ID_N
      CONT_N = 0

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      'BUSCA VALOR TOTAL VENDA
      VALOR_ITEM_N = 0
      SQL = "select sum(valor*qtde) from NFITEM WITH (NOLOCK)"
      SQL = SQL & " where nf_id = " & NF_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      'ITENS
      If TabLANCAMENTOITEM.State = 1 Then _
         TabLANCAMENTOITEM.Close

      SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " where seq = " & 1
      SQL = SQL & " and lancamento_id = " & NUMR_ID_N
      TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLANCAMENTOITEM.EOF Then
         SQL = "UPDATE ITEMLANCAMENTO SET "
            SQL = SQL & "  usu_alt = " & USUARIO_ID_N
            SQL = SQL & ", Dt_Alt = '" & Date & "'"
            SQL = SQL & ", Numr_doc = " & Trim(Numero_NF_A)
            SQL = SQL & ", Seq = " & 1
            SQL = SQL & ", Valor_Item = '" & tpMOEDA(VALOR_ITEM_N) & "'"
            SQL = SQL & ", Status = 'B'"
            SQL = SQL & ", formapagto_id = " & 1
            SQL = SQL & ", DT_VENCIMENTO = '" & Date & "'"
         SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
         SQL = SQL & " and Seq = " & 1
         Else
            SQL = "INSERT INTO ITEMLANCAMENTO "
               SQL = SQL & " (Usu_cad, Dt_cad, Lancamento_id, Numr_doc, "
               SQL = SQL & " NUMR_DP, seq, Valor_Item, Status, formapagto_id, "
               SQL = SQL & " DT_VENCIMENTO, ACERTO,cc_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & USUARIO_ID_N
               SQL = SQL & ",'" & Date & "'"
               SQL = SQL & "," & NUMR_ID_N
               SQL = SQL & "," & Trim(Numero_NF_A)
               SQL = SQL & "," & Trim(Numero_NF_A)
               SQL = SQL & "," & 1
               SQL = SQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"
               SQL = SQL & ",'B'"
               SQL = SQL & "," & 1
               SQL = SQL & ",'" & Date & "'"
               SQL = SQL & "," & 0
               SQL = SQL & "," & 0
            SQL = SQL & ")"
      End If
      If TabLANCAMENTOITEM.State = 1 Then _
         TabLANCAMENTOITEM.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close
   If TabTemp.State = 1 Then _
      TabTemp.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "GERA_FINANC_DEVOLUCAO"
End Sub

Private Sub SETA_GRID_DIVERSAS()
'On Error GoTo ERRO_TRATA

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * FROM NF WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON NF.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status = 'A'"

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      CRITERIO_A = ""
      NF_ID_N = 0 & TabCabeca.Fields("nf_id").Value
      VALOR_TOTAL_N = 0

      txtCNPJCPF.PromptInclude = False
      If Len(Trim(TabCabeca.Fields("cnpjcpf").Value)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If

      txtCNPJCPF.Text = TabCabeca.Fields("cnpjcpf").Value
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = TabCabeca.Fields("cnpjcpf").Value

'========================================cliente
      If Not IsNull(TabCabeca.Fields("descricao").Value) Then
         If Trim(TabCabeca.Fields("descricao").Value) <> "" Then
            CRITERIO_A = Trim(TabCabeca!DESCRICAO)
            Else: TRAZ_NOME_CLIENTE (TabCabeca.Fields("CLIENTE_ID").Value)
         End If
         Else: TRAZ_NOME_CLIENTE (TabCabeca.Fields("CLIENTE_ID").Value)
      End If

'========================================setando grid
      Set item = lstPedidos.ListItems.Add(, "seq." & NUMR_SEQ_N, Trim(TabCabeca.Fields("numr_nota").Value))

      item.SubItems(1) = "" & txtCNPJCPF.Text
      item.SubItems(2) = "" & CRITERIO_A
      item.SubItems(3) = "NFe"

'========================================
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum(valor*qtde) from NFITEM WITH (NOLOCK) "
      SQL = SQL & " where nf_id = " & NF_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_TOTAL_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
'========================================
      item.SubItems(4) = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      item.SubItems(5) = "" & Format(0, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      item.SubItems(7) = "" & Trim(TabCabeca!DT_EMISSAO)
      item.SubItems(8) = ""
'========================================

      item.SubItems(9) = "" & TabCabeca.Fields("nf_tipo").Value
      item.SubItems(10) = "" & TabCabeca.Fields("nf_id").Value

      item.SubItems(12) = ""
      SQL = "select formapagto_id from PEDIDOFATURA WITH (NOLOCK) "
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then _
         item.SubItems(12) = "" & TabVENDEDOR.Fields(0).Value
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      item.SubItems(13) = "" & TRAZ_ID_TABELA("CLIENTE", "PESSOA_ID", "pessoa_id", TabCabeca.Fields("cnpjcpf").Value)

      NUMR_SEQ_N = NUMR_SEQ_N + 1

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      If Trim(UCase(TabCabeca.Fields("nf_tipo").Value)) = "DV" Then
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
      If Trim(UCase(TabCabeca.Fields("nf_tipo").Value)) = "DC" Then
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
      If Trim(UCase(TabCabeca.Fields("nf_tipo").Value)) = "OS" Then
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

      TabCabeca.MoveNext
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstPedidos.Refresh

   'MOSTRA_TOTAIS

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_DIVERSAS"
End Sub
