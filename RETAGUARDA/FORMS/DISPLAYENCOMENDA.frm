VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmDISPLAYENCOMENDA 
   Caption         =   "Encomendas Agendadas"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DISPLAYENCOMENDA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView LISTAITEM 
      Height          =   2985
      Left            =   45
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   11340
      _ExtentX        =   20003
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
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
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
      DesignWidth     =   11400
      DesignHeight    =   6495
   End
   Begin MSComctlLib.ListView lstTotais 
      Height          =   735
      Left            =   45
      TabIndex        =   3
      Top             =   5400
      Width           =   11295
      _ExtentX        =   19923
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
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "DISPLAYENCOMENDA.frx":5C12
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
Attribute VB_Name = "frmDISPLAYENCOMENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_DESCONTO_CABECA_N   As Double
   Dim DEVOLUCAO_VENDA_N         As String
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

   PESQUISA_VENDA
   PESQUISA_Devolução_ENTRADA

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
         If Trim(lstPedidos.SelectedItem.ListSubItems.item(9).Text) = "DC" Then 'Devolução de Entrada

            PEDIDO_ID_N = lstPedidos.SelectedItem.ListSubItems.item(10).Text

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select NOTAENTRADA.pedidocompra_id, NOTAENTRADA.NUMR_NOTA, NOTAENTRADA.SERIE_NOTA, "
            SQL = SQL & " NOTAENTRADA.DT_ENTRADA, NOTAENTRADA.DT_EMISSAO, NOTAENTRADA.VALOR_FRETE"
            SQL = SQL & " from NOTAENTRADA "
            SQL = SQL & " INNER JOIN FORNECEDOR "
            SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

            SQL = SQL & " where tipoentrada_id = 1 "
            SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
            SQL = SQL & " and NOTAENTRADA.estabelecimento_ID = " & EMPRESA_ID_N
            SQL = SQL & " and NOTAENTRADA.pedidocompra_id = " & PEDIDO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               Msg = "Deseja realmente cancelar nota de devolução número: " & Trim(TabTemp.Fields("numr_nota").Value) & " Serie: " & Trim(TabTemp.Fields("serie_nota").Value) & " , Fornecedor: " & TabTemp.Fields("descricao").Value & " ?"
               PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  SQL = "update NOTAENTRADA set "
                  SQL = SQL & " status = 'C' "
                  SQL = SQL & " where tipoentrada_id = 1 "
                  SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
                  SQL = SQL & " and NOTAENTRADA.estabelecimento_ID = " & EMPRESA_ID_N
                  SQL = SQL & " and NOTAENTRADA.pedidocompra_id = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

                  MsgBox "Operação realizada com sucesso !!!"
               End If
            End If

            If TabTemp.State = 1 Then _
               TabTemp.Close
            Else
               If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
                  frmPedidoCancela.txtPedido.Text = 0 & lstPedidos.SelectedItem.ListSubItems.item(10).Text
                  frmPedidoCancela.Show 1
                  CRITERIO_A = ""
                  Else: MsgBox "Não permitido."
               End If
         End If

         PESQUISA_VENDA
         PESQUISA_Devolução_ENTRADA
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

            SQL = "select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.valor_desconto,"
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
               CRITERIO_A = ""
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close

            LISTAITEM.Refresh
         End If
      Case vbKeyF9
         PESQUISA_VENDA
         PESQUISA_Devolução_ENTRADA
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
   TRATA_ERROS Err.Description, Me.Name, "lstPedidos_Click"
End Sub

Private Sub lstPedidos_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedidos.SelectedItem.ListSubItems.item(10).Text) Then

      PEDIDO_ID_N = lstPedidos.SelectedItem.ListSubItems.item(10).Text

      '================================== PEDIDO DE VENDA
      If UCase(Trim(frmDISPLAYENCOMENDA.lstPedidos.SelectedItem.ListSubItems.item(9).Text)) = UCase("PEDIDO") Then
         TIPO_NFe_GERAR = "R"          'Tipo Saida

         FAZ_RECEBIMENTO

      End If

   End If

   PESQUISA_VENDA
   PESQUISA_Devolução_ENTRADA

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

   SETA_GRID

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

   SQL = "select * from PEDIDO "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status = 8 " 'somente encomendas
   SQL = SQL & " and tipo_registro in ('S','R','D') "

   SQL = SQL & " order by pedido_id DESC "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      FORMULA_REL = "{PEDIDO.status} = " & 2
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

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstPedidos.ListItems.Clear
   NUMR_SEQ_N = 0
   NUMR_CONSULTA_N = 0

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PEDIDO "

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status = 8 " 'somente encomendas
   SQL = SQL & " and tipo_registro in ('S','R','D','OS') "
   SQL = SQL & " order by dt_req DESC "
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCABECA.EOF
      CRITERIO_A = ""

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
            CRITERIO_A = Trim(TabCABECA!NOME_CLIENTE)
            Else: BUSCA_CLIENTE (TabCABECA.Fields("CLIENTE_ID").Value)
         End If
         Else: BUSCA_CLIENTE (TabCABECA.Fields("CLIENTE_ID").Value)
      End If

'========================================setando grid
      Set item = lstPedidos.ListItems.Add(, "seq." & Trim(TabCABECA.Fields("pedido_id").Value), Trim(TabCABECA.Fields("pedido_id").Value))

      item.SubItems(1) = "" & txtCNPJCPF.Text
      item.SubItems(2) = "" & CRITERIO_A

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

      If TabCABECA!TIPO_REGISTRO <> "D" Then
         item.SubItems(9) = "Pedido"
         Else: item.SubItems(9) = "DV"
      End If

      item.SubItems(10) = "" & TabCABECA.Fields("pedido_id").Value
      item.SubItems(12) = "" & TabCABECA.Fields("tipovenda_id").Value

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      CONT_N = CONT_N + 1
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      If Trim(UCase(TabCABECA.Fields("tipo_registro").Value)) = "D" Then
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
      If Trim(UCase(TabCABECA.Fields("tipo_registro").Value)) = "D" Then
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
'vbBlack  vbRed  vbGreen  vbYellow  vbBlue  vbMagenta  vbCyan  vbWhite
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
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   If USA_DOC_FISCAL = True Then
      If USA_NFe = True Then

         SQL = "select status from PEDIDO "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If Not IsNull(TabCABECA!Status) Then
               If TabCABECA!Status <> "" Then
                  If DEVOLUCAO_VENDA_N <> "D" Then
                     If TabCABECA!Status = 5 Or TabCABECA!Status = 7 Then
                        CRITERIO_A = PEDIDO_ID_N
                        If TabCABECA.State = 1 Then _
                           TabCABECA.Close
                        frmNOTAGERA.Show 1
                     End If
                  Else
                     If TabCABECA!Status = 2 Then
                        CRITERIO_A = PEDIDO_ID_N
                        If TabCABECA.State = 1 Then _
                           TabCABECA.Close
                        frmNOTAGERA.Show 1
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
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
      Else: FAZ_RECEBIMENTO
   End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "CHECA_ESTOQUE"
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
         If TabPedido.State = 1 Then _
            TabPedido.Close

         SQL = "select * from PEDIDO "
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedido.EOF Then
            PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
            If TabPedido!Status = 5 Then
               CNPJCPF_A = Trim(TabPedido!CGCCPF)

'====================
ATUALIZA_ESTOQUE 0, PEDIDO_ID_N
'====================
            End If
         End If
         If TabPedido.State = 1 Then _
            TabPedido.Close
      End If   'If Not IsNull(lstPedidos.selectedItem.Text) Then
   End If      'If Trim(lstPedidos.selectedItem.Text) <> "" Then
   If TabPedido.State = 1 Then _
      TabPedido.Close

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub

Private Sub PESQUISA_Devolução_ENTRADA()
'On Error GoTo ERRO_TRATA

   SETA_GRID_ENTRADA

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "PESQUISA_Devolução_ENTRADA"
End Sub

Private Sub SETA_GRID_ENTRADA()
'On Error GoTo ERRO_TRATA

   NOTAENTRADA_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select NOTAENTRADA.TIPOENTRADA_ID, NOTAENTRADA.STATUS, NOTAENTRADA.ENTRADA_ID, NOTAENTRADA.ESTABELECIMENTO_ID, NOTAENTRADA.FORNECEDOR_ID, "
   SQL = SQL & " NOTAENTRADA.TRANSP_ID, NOTAENTRADA.PEDIDOCOMPRA_ID, NOTAENTRADA.NUMR_NOTA, NOTAENTRADA.SERIE_NOTA, NOTAENTRADA.DT_ENTRADA,"
   SQL = SQL & " NOTAENTRADA.DT_EMISSAO, NOTAENTRADA.VALOR_FRETE, NOTAENTRADA.VALOR_DESCONTO, FORNECEDOR.PESSOA_ID, PESSOA.CNPJCPF,"
   SQL = SQL & " PESSOA.DESCRICAO , PESSOA.RAZAO"
   SQL = SQL & " from NOTAENTRADA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where tipoentrada_id = 1 "
   SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
   SQL = SQL & " and NOTAENTRADA.estabelecimento_ID = " & EMPRESA_ID_N
   SQL = SQL & " order by entrada_id asc "

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NOTAENTRADA_ID_N = TabTemp.Fields("entrada_id").Value

      CRITERIO_A = Trim(TabTemp!NOME)
      txtCNPJCPF.PromptInclude = False
      If Len(Trim(TabTemp!CNPJCPF)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If

      txtCNPJCPF.Text = TabTemp!CNPJCPF
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = TabTemp!CNPJCPF

      Set item = lstPedidos.ListItems.Add(, "seq." & Trim(TabTemp!numr_pedido_compra), Trim(TabTemp!numr_pedido_compra))
      item.SubItems(1) = txtCNPJCPF.Text
      item.SubItems(2) = CRITERIO_A

      PERC_DESCONTO_N = 0
      VALOR_DESCONTO_N = 0
      VALOR_TOTAL_DESCONTO_N = 0
      VALOR_DESCONTO_CABECA_N = 0 & TabTemp.Fields("valor_desconto").Value
      VALOR_ITEM_N = 0

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((qtde_entrada*preco_custo)) from NOTAENTRADAITEM "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      VALOR_DESCONTO_N = VALOR_DESCONTO_CABECA_N
      
      item.SubItems(4) = Format(Trim(VALOR_ITEM_N), strFormatacao2Digitos)
      item.SubItems(5) = Format(Trim(VALOR_DESCONTO_N), strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      item.SubItems(7) = Trim(TabTemp!DT_EMISSAO)
      item.SubItems(8) = NOME_EMPRESA_A
      item.SubItems(9) = "Devolução Entrada"
      item.SubItems(9) = "DC"
      item.SubItems(10) = TabTemp.Fields("ENTRADA_ID").Value

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      CONT_N = CONT_N + 1

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedidos.Refresh

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ENTRADA"
End Sub

Private Sub GERA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   If USA_DOC_FISCAL = True Then
      If USA_NFe = True Then

         PEDIDO_ID_N = lstPedidos.SelectedItem.ListSubItems.item(10).Text

         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select status from NOTAENTRADA "
         SQL = SQL & " where entrada_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & EMPRESA_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If Not IsNull(TabCABECA!Status) Then
               If TabCABECA!Status <> "" Then
                  If TabCABECA!Status = "D" Then
                     If TabCABECA.State = 1 Then _
                        TabCABECA.Close
      
                     frmNOTAGERA.Show 1
                  End If
               End If
            End If
         End If
         If TabCABECA.State = 1 Then _
            TabCABECA.Close
      End If
   End If

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA_ENTRADA"
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
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_CLIENTE"
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

   SQL = "select distinct(TIPOVENDA_ID) from PEDIDO "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status = 8 " 'somente encomendas
   SQL = SQL & " and tipo_registro in ('S','R','D','OS') "
   SQL = SQL & " order by TIPOVENDA_ID "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum(valor_total-valor_desconto) from PEDIDO "
      SQL = SQL & " where tipo_registro in ('S','R','D','OS') "
      SQL = SQL & " and status = 8 " 'somente encomendas
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

