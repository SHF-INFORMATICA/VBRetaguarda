VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVENDACUSTO 
   Caption         =   "Relatório Analise Venda/Custo"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VENDACUSTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   -120
      TabIndex        =   2
      Top             =   1200
      Width           =   12735
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Data Liquidação Título"
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.OptionButton optPedido 
         Caption         =   "Por Data Pedido Venda"
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   420
         Left            =   4320
         TabIndex        =   0
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Height          =   420
         Left            =   6840
         TabIndex        =   1
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         Height          =   285
         Left            =   10920
         TabIndex        =   8
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6840
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1080
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1270
      ButtonWidth     =   2461
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
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8760
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACUSTO.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACUSTO.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACUSTO.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACUSTO.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACUSTO.frx":9EFB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstItem 
      Height          =   5535
      Left            =   45
      TabIndex        =   9
      Top             =   2520
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   9763
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vlr.Desc."
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr.Item"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Tot.Item"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Dt.Emisão"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "PrCustoVenda"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "PrCustoTabela"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "PrCustoProd"
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
      DesignWidth     =   12615
      DesignHeight    =   8085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Relatório Venda/Custo/Produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   12660
   End
End
Attribute VB_Name = "frmVENDACUSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VALOR_CUSTO_TABELA_N   As Double
Dim Nome_Vendedor_A        As String

Private Sub Form_Load()

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   CRITERIO_A = Month(Date)
   If Len(CRITERIO_A) = 1 Then _
      CRITERIO_A = "0" & CRITERIO_A

   txtDtIni.PromptInclude = False
   txtDtIni.Text = "01/" & CRITERIO_A & "/" & Year(Date) & "00:00:00"
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   CRITERIO_A = FimDoMes(txtDtIni.Text, False)
   CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
   txtDtFim.Text = CRITERIO_A & "23:59:59"
   txtDtFim.PromptInclude = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imprimir"
         CRIA_TABELA_TEMPORARIA
         lstItem.Visible = False
         If optPedido.Value = True Then
            GERA_REL
            Else: GERA_REL_FINAC
         End If
         lstItem.Visible = True
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   CRITERIO_A = FimDoMes(txtDtIni.Text, False)
   CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
   txtDtFim.Text = CRITERIO_A & "23:59:59"
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Sub GERA_REL()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False
   lstItem.ListItems.Clear
   CONT_N = 0
   CRITERIO_A = ""
   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      QTDE_N = 0
      VALOR_ITEM_N = 0
      VALOR_CUSTO_N = 0
      VALOR_DESCONTO_N = 0
      NUMR_ID_N = 0
      PEDIDO_ID_N = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwRelVenda "
      SQL = SQL & " where status in (3,5,7)"
      SQL = SQL & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_req <= '" & txtDtFim.Text & "'"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         VALOR_DESCONTO_N = 0 & TabTemp.Fields("descontoitem").Value
         CONT_N = CONT_N + 1
         QTDE_N = 0 & TabTemp.Fields("qtd_pedida").Value
         VALOR_ITEM_N = 0 & TabTemp.Fields("valor_item").Value

         'esse valor aqui vem da tabela pedidoitem.preco_custo
         VALOR_CUSTO_N = 0 & TabTemp.Fields("CustoItem").Value

         VALOR_DESCONTO_N = 0
         CLIENTE_ID_N = 0
         PRODUTO_ID_N = 0 & TabTemp.Fields("produto_id").Value
         TABELAPRECO_ID_N = 0 & TabTemp.Fields("tabelapreco_id").Value

         lblConta.Caption = CONT_N
         lblConta.Refresh
         DoEvents

         If Not IsNull(TabTemp.Fields("cliente_id").Value) Then _
            CLIENTE_ID_N = TabTemp.Fields("cliente_id").Value

         If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
            VALOR_DESCONTO_N = 0 & Format(TabTemp.Fields("valor_desconto").Value, strFormatacao2Digitos)
            PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value
         End If
         If Not IsNull(TabTemp.Fields("descontoitem").Value) Then _
            VALOR_DESCONTO_N = VALOR_DESCONTO_N + Format(TabTemp.Fields("descontoitem").Value, strFormatacao2Digitos)

         If Not IsNull(TabTemp.Fields("CustoProduto").Value) Then _
            If TabTemp.Fields("CustoProduto").Value > 0 Then _
               VALOR_CUSTO_N = 0 & TabTemp.Fields("CustoProduto").Value

'set aqui ele vai pegar o preco de custo da tabela, ou seja,
'se vendemos a tres meses atras e o custo aumentou agora,
'ta errado, tem que calcular com o preço
'de custo da epoca da venda
''''''''''''''''''''''''''''''''''''
'MsgBox TabTemp.Fields("produto_id").Value
VALOR_CUSTO_N = 0 & TabTemp.Fields("CustoItem").Value

VALOR_CUSTO_TABELA_N = 0 & BUSCA_FATURAMENTO_TABELAPRECO(TabTemp.Fields("pedido_id").Value)

If VALOR_CUSTO_N <= 0 Then
   If VALOR_CUSTO_TABELA_N > 0 Then _
      VALOR_CUSTO_N = VALOR_CUSTO_TABELA_N
End If

'tabela produto
If VALOR_CUSTO_N <= 0 Then _
   VALOR_CUSTO_N = 0 & TabTemp.Fields("CustoProduto").Value

VENDEDOR_ID_N = 0 & TabTemp.Fields("vendedor_id").Value
Nome_Vendedor_A = "" & TRAZ_NOME_VENDEDOR(VENDEDOR_ID_N)

         Set item = lstItem.ListItems.Add(, "seq." & CONT_N, TabTemp.Fields("PEDIDO_ID").Value)
         item.SubItems(1) = "" & Trim(TabTemp.Fields("codg_produto").Value) & "-" & Trim(TabTemp.Fields("descproduto").Value)
         item.SubItems(2) = "" & Format(TabTemp.Fields("qtd_pedida").Value, strFormatacao3Digitos)
         item.SubItems(3) = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
         item.SubItems(4) = "" & Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
         item.SubItems(5) = "" & Format((TabTemp.Fields("valor_item").Value * TabTemp.Fields("qtd_pedida").Value) - VALOR_DESCONTO_N, strFormatacao2Digitos)
         item.SubItems(6) = "" & Trim(TabTemp.Fields("dt_req").Value)
         item.SubItems(7) = "" & Nome_Vendedor_A
         item.SubItems(8) = "" & Format(TabTemp.Fields("CustoItem").Value, strFormatacao2Digitos)
         item.SubItems(9) = "" & Format(VALOR_CUSTO_TABELA_N, strFormatacao2Digitos)
         item.SubItems(10) = "" & Format(TabTemp.Fields("CustoProduto").Value, strFormatacao2Digitos)

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from VENDACUSTO "
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabConsulta.EOF Then
            NUMR_ID_N = MAX_ID("vendacusto_id", "vendacusto", "", "", "", "")

            SQL = "insert into VENDACUSTO values("
               SQL = SQL & NUMR_ID_N                                                      'VENDACUSTO_ID
               SQL = SQL & "," & EMPRESA_ID_N                                             'EMPRESA_ID
               SQL = SQL & "," & CLIENTE_ID_N                                             'CLIENTE_ID
               SQL = SQL & ",'" & Trim(TabTemp.Fields("dt_req").Value) & "'"              'DT_VENDA
               SQL = SQL & "," & TabTemp.Fields("pedido_id").Value                        'PEDIDO_ID
               SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N * QTDE_N)                           'VLR_TOT_VENDA
               SQL = SQL & "," & tpMOEDA(VALOR_CUSTO_N * QTDE_N)                          'VLR_TOT_CUSTO
               SQL = SQL & "," & tpMOEDA(VALOR_DESCONTO_N)                                'VLR_TOT_DESCONTO
            SQL = SQL & ",'" & Trim(Left(TabTemp.Fields("nome_cliente").Value, 50)) & "'" 'CLIENTE
            SQL = SQL & ",'" & Trim(Left(Nome_Vendedor_A, 50)) & "'"                      'VENDEDOR
            SQL = SQL & ")"
            Else
               SQL = "update VENDACUSTO set "
                  SQL = SQL & " vlr_tot_venda = vlr_tot_venda + " & tpMOEDA(VALOR_ITEM_N * QTDE_N)
                  SQL = SQL & ", vlr_tot_custo = vlr_tot_custo + " & tpMOEDA(VALOR_CUSTO_N * QTDE_N)
                  SQL = SQL & ", vlr_tot_desconto = vlr_tot_desconto + " & tpMOEDA(VALOR_DESCONTO_N)
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         CONECTA_RETAGUARDA.Execute SQL

         DoEvents

         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   Me.Enabled = True

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "venda_custo.rpt"
   frmRELATORIO10.Show 1

   Me.Enabled = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "GERA_REL"
End Sub

Sub GERA_REL_FINAC()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False
   CONT_N = 0
   CRITERIO_A = ""
   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      QTDE_N = 0
      VALOR_ITEM_N = 0
      VALOR_CUSTO_N = 0
      VALOR_DESCONTO_N = 0
      NUMR_ID_N = 0
      PEDIDO_ID_N = 0

      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      SQL = "select distinct(LANCAMENTO.NUMR_DOC) from LANCAMENTO"
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO"
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
      SQL = SQL & " AND LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

      SQL = SQL & " where ITEMLANCAMENTO.status = 'B'"
      SQL = SQL & " and ITEMLANCAMENTO.dt_baixa >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and ITEMLANCAMENTO.dt_baixa <= '" & txtDtFim.Text & "'"
      SQL = SQL & " and LANCAMENTO.tipo_lancamento = 1"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabCabeca.EOF

         If Not IsNull(TabCabeca.Fields(0).Value) Then
            If TabCabeca.Fields(0).Value > 0 Then
               PROCESSA_VWRELVENDA TabCabeca.Fields(0).Value

            SQL = "update VENDACUSTO set "
            SQL = SQL & " vlr_tot_venda = " & tpMOEDA(BUSCA_FATURAMENTO_QUITADO(TabCabeca.Fields(0).Value))
            SQL = SQL & " where pedido_id = " & TabCabeca.Fields(0).Value
            CONECTA_RETAGUARDA.Execute SQL
            End If
         End If

         TabCabeca.MoveNext
      Wend
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
   End If

   Me.Enabled = True

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
'MsgBox FORMULA_REL
   Nome_Relatorio = "venda_custo_finac.rpt"
   frmRELATORIO10.Show 1

   Me.Enabled = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "GERA_REL_FINAC"
End Sub

Sub PROCESSA_VWRELVENDA(NUMR_PEDIDO_N As Long)
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
         TabTemp.Close

   SQL = "select * from vwRelVenda "
   SQL = SQL & " where pedido_id = " & NUMR_PEDIDO_N
   SQL = SQL & " and status in (3,5,7)"

   'SQL = SQL & " and dt_req >= '" & Format(txtDtIni.Text, "dd/mm/yyyy") & "'"
   'SQL = SQL & " and dt_req <= '" & Format(txtDtFim.Text, "dd/mm/yyyy") & "'"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CONT_N = CONT_N + 1
      QTDE_N = 0 & TabTemp.Fields("qtd_pedida").Value
      VALOR_ITEM_N = 0 & TabTemp.Fields("valor_item").Value
      VALOR_CUSTO_N = 0 & TabTemp.Fields("preco_custo").Value
      VALOR_DESCONTO_N = 0
      CLIENTE_ID_N = 0
      PRODUTO_ID_N = 0 & TabTemp.Fields("produto_id").Value
      TABELAPRECO_ID_N = 0 & TabTemp.Fields("tabelapreco_id").Value

      lblConta.Caption = CONT_N
      lblConta.Refresh
      DoEvents

      If Not IsNull(TabTemp.Fields("cliente_id").Value) Then _
         CLIENTE_ID_N = TabTemp.Fields("cliente_id").Value

      If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
         VALOR_DESCONTO_N = 0 & Format(TabTemp.Fields("valor_desconto").Value, strFormatacao2Digitos)
         PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value
      End If
      If Not IsNull(TabTemp.Fields("descontoitem").Value) Then _
         VALOR_DESCONTO_N = VALOR_DESCONTO_N + Format(TabTemp.Fields("descontoitem").Value, strFormatacao2Digitos)

      If Not IsNull(TabTemp.Fields("PreçoCusto").Value) Then _
         If TabTemp.Fields("PreçoCusto").Value > 0 Then _
            VALOR_CUSTO_N = 0 & TabTemp.Fields("PreçoCusto").Value

'''''''''''''''''''''''''''''''''''
VALOR_CUSTO_TABELA_N = 0 & BUSCA_FATURAMENTO_TABELAPRECO(TabTemp.Fields("pedido_id").Value)
If VALOR_CUSTO_TABELA_N > 0 Then _
   VALOR_CUSTO_N = VALOR_CUSTO_TABELA_N

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from VENDACUSTO "
      SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         NUMR_ID_N = MAX_ID("vendacusto_id", "vendacusto", "", "", "", "")

         SQL = "insert into VENDACUSTO values("
            SQL = SQL & NUMR_ID_N                                                'VENDACUSTO_ID
            SQL = SQL & "," & EMPRESA_ID_N                                       'EMPRESA_ID
            SQL = SQL & "," & CLIENTE_ID_N                                       'CLIENTE_ID
            SQL = SQL & ",'" & Trim(TabTemp.Fields("dt_req").Value) & "'"         'DT_VENDA
            SQL = SQL & "," & TabTemp.Fields("pedido_id").Value                   'PEDIDO_ID
            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N * QTDE_N)                     'VLR_TOT_VENDA
            SQL = SQL & "," & tpMOEDA(VALOR_CUSTO_N * QTDE_N)                    'VLR_TOT_CUSTO
            SQL = SQL & "," & tpMOEDA(VALOR_DESCONTO_N)                          'VLR_TOT_DESCONTO
            SQL = SQL & ",'" & Trim(Left(TabTemp.Fields("nome_cliente").Value, 50)) & "'" 'CLIENTE
            SQL = SQL & ",'" & Trim(Left(Nome_Vendedor_A, 50)) & "'"     'VENDEDOR
         SQL = SQL & ")"
         Else
            SQL = "update VENDACUSTO set "
               SQL = SQL & " vlr_tot_venda = vlr_tot_venda + " & tpMOEDA(VALOR_ITEM_N * QTDE_N)
               SQL = SQL & ", vlr_tot_custo = vlr_tot_custo + " & tpMOEDA(VALOR_CUSTO_N * QTDE_N)
               SQL = SQL & ", vlr_tot_desconto = vlr_tot_desconto + " & tpMOEDA(VALOR_DESCONTO_N)
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      
      CONECTA_RETAGUARDA.Execute SQL

      DoEvents

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_VWRELVENDA"
End Sub

Function BUSCA_FATURAMENTO_QUITADO(NUMR_PEDIDO_N As Long) As Double
'On Error GoTo ERRO_TRATA

   BUSCA_FATURAMENTO_QUITADO = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SqL2 = "select sum(ITEMLANCAMENTO.VALOR_ITEM) from ITEMLANCAMENTO "
   SqL2 = SqL2 & " where numr_doc = " & NUMR_PEDIDO_N
   SqL2 = SqL2 & " and status = 'B' "
   TabLancamento.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento.Fields(0).Value) Then _
         BUSCA_FATURAMENTO_QUITADO = 0 & TabLancamento.Fields(0).Value
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_FATURAMENTO_QUITADO"
End Function

Function BUSCA_FATURAMENTO_TABELAPRECO(NUMR_PEDIDO_N As Long) As Double
'On Error GoTo ERRO_TRATA

   BUSCA_FATURAMENTO_TABELAPRECO = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SqL2 = "select  ITEMLANCAMENTO.FORMAPAGTO_ID from LANCAMENTO "
   SqL2 = SqL2 & " INNER JOIN ITEMLANCAMENTO "
   SqL2 = SqL2 & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   SqL2 = SqL2 & " AND LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

   SqL2 = SqL2 & " where ITEMLANCAMENTO.numr_doc = " & NUMR_PEDIDO_N
   SqL2 = SqL2 & " and ITEMLANCAMENTO.status = 'B' "
   SqL2 = SqL2 & " order by ITEMLANCAMENTO.FORMAPAGTO_ID"

   TabLancamento.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      If Not IsNull(TabLancamento.Fields(0).Value) Then

         BUSCA_FATURAMENTO_TABELAPRECO = 0 & TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PRODUTO_ID_N, TABELAPRECO_ID_N, TabLancamento.Fields(0).Value)

      End If
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_FATURAMENTO_QUITADO"
End Function

Sub CRIA_TABELA_TEMPORARIA()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "VENDACUSTO", "U") = True Then
      SQL = "drop table VENDACUSTO"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "VENDACUSTO", "U") = False Then
      SQL = "create table VENDACUSTO"
      SQL = SQL & " ("
         SQL = SQL & " VENDACUSTO_ID     bigint       not null,"
         SQL = SQL & " EMPRESA_ID        bigint       not null,"
         SQL = SQL & " CLIENTE_ID        bigint       not null,"
         SQL = SQL & " DT_VENDA          datetime     not null,"
         SQL = SQL & " PEDIDO_ID         bigint       not null,"
         SQL = SQL & " VLR_TOT_VENDA     float        not null,"
         SQL = SQL & " VLR_TOT_CUSTO     float        null    ,"
         SQL = SQL & " VLR_TOT_DESCONTO  float        null    ,"
         SQL = SQL & " CLIENTE           varchar(50)  null    ,"
         SQL = SQL & " VENDEDOR          varchar(50)  null    ,"
         SQL = SQL & " constraint PK_VENDACUSTO primary key (VENDACUSTO_ID)"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   SQL = "delete from VENDACUSTO"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CRIA_TABELA_TEMPORARIA"
End Sub
