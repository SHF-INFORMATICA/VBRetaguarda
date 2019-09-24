VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCHAMADOS 
   BackColor       =   &H80000008&
   Caption         =   "Emissor de Documento Fiscal"
   ClientHeight    =   6495
   ClientLeft      =   1635
   ClientTop       =   2475
   ClientWidth     =   11400
   Icon            =   "CHAMADOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleMode       =   0  'User
   ScaleWidth      =   28051.34
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstPedidos 
      Height          =   5745
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   10134
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
      TabIndex        =   1
      Top             =   6000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1270
      ButtonWidth     =   2249
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6000
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
               Picture         =   "CHAMADOS.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHAMADOS.frx":19B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHAMADOS.frx":2A41
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHAMADOS.frx":3BDB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHAMADOS.frx":4E5F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHAMADOS.frx":5F90
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCHAMADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name

   'PESQUISA_chamado

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF3
         IMPRIME_TELA
      Case vbKeyF6
         If Trim(lstPedidos.SelectedItem.ListSubItems.Item(9).Text) = "DVE" Then 'Devolução de Entrada

            NUMR_REQ_N = lstPedidos.SelectedItem.ListSubItems.Item(10).Text

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "SELECT NOTAENTRADA.NUMR_PEDIDO_COMPRA, NOTAENTRADA.NUMR_NOTA, NOTAENTRADA.SERIE_NOTA, "
            SQL = SQL & " NOTAENTRADA.DT_ENTRADA, NOTAENTRADA.DT_EMISSAO, NOTAENTRADA.VALOR_FRETE, "
            SQL = SQL & " FORNECEDOR.CGCCPF, FORNECEDOR.NOME "
            SQL = SQL & " FROM NOTAENTRADA "
            SQL = SQL & " INNER JOIN FORNECEDOR "
            SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

            SQL = SQL & " where tipoentrada_id = 1 "
            SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
            SQL = SQL & " and NOTAENTRADA.empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and NOTAENTRADA.NUMR_PEDIDO_COMPRA = " & NUMR_REQ_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               Msg = "Deseja realmente cancelar nota de devolução número: " & Trim(TabTemp.Fields("numr_nota").Value) & " Serie: " & Trim(TabTemp.Fields("serie_nota").Value) & " , Fornecedor: " & TabTemp.Fields("NOME").Value & " ?"
               PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  SQL = "update NOTAENTRADA set "
                  SQL = SQL & " status = 'C' "
                  SQL = SQL & " where tipoentrada_id = 1 "
                  SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
                  SQL = SQL & " and NOTAENTRADA.empresa_id = " & EMPRESA_ID_N
                  SQL = SQL & " and NOTAENTRADA.NUMR_PEDIDO_COMPRA = " & NUMR_REQ_N
                  CONECTA_RETAGUARDA.Execute SQL

                  MsgBox "Operação realizada com sucesso !!!"
               End If
            End If

            If TabTemp.State = 1 Then _
               TabTemp.Close
            Else
               If UCase(Trim(lstPedidos.SelectedItem.ListSubItems.Item(9).Text)) = UCase("PEDIDO") Then
                  frmPedidoCancela.txtPedido.Text = 0 & lstPedidos.SelectedItem.ListSubItems.Item(10).Text
                  frmPedidoCancela.Show 1
                  CRITERIO = ""
               End If
         End If

         PESQUISA_VENDA
         PESQUISA_Devolução_ENTRADA
      Case vbKeyF7
         If Not IsNull(lstPedidos.SelectedItem.Text) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cgccpf from PEDIDO "
            SQL = SQL & " where numr_req = " & lstPedidos.SelectedItem.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
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

            SQL = "select i.codg_prod,p.descricao,i.qtd_pedida,i.Valor_Item,"
            SQL = SQL & " i.valor_desconto,i.status as st_item,i.seq_id "
            SQL = SQL & " FROM PEDIDOITEM i, PRODUTO p "
            SQL = SQL & " where i.codg_prod = p.CODG_PRODUTO "
            SQL = SQL & " and i.numr_req = " & lstPedidos.SelectedItem.Text
            SQL = SQL & " and p.empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and I.tipo_reg = 'PC' "
            SQL = SQL & " order by p.referencia,p.descricao"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               MOSTRA_TOP "Duplo Click no grid ocultar", " ", " ", " ", ""
               LISTAITEM.Visible = True
            End If
            While Not TabTemp.EOF
               Set Item = LISTAITEM.ListItems.Add(, "seq." & TabTemp.Fields("seq_id"), Trim(TabTemp.Fields("codg_prod").Value))
               Item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
               Item.SubItems(2) = "" & Format(TabTemp.Fields("qtd_pedida").Value, strFormatacao3Digitos)
               Item.SubItems(3) = "" & Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
               Item.SubItems(4) = "" & Format(TabTemp.Fields("valor_desconto").Value, strFormatacao2Digitos)
               Item.SubItems(5) = "" & Format((TabTemp.Fields("valor_item").Value - TabTemp.Fields("valor_desconto").Value) * TabTemp.Fields("qtd_pedida").Value, strFormatacao2Digitos)
               Item.SubItems(6) = "" & Trim(TabTemp.Fields("st_item").Value)

               If Trim(TabTemp.Fields("st_item").Value) = "A" Then
                  Item.ForeColor = vbBlue
                  Item.ListSubItems(1).ForeColor = vbBlue
                  Item.ListSubItems(2).ForeColor = vbBlue
                  Item.ListSubItems(3).ForeColor = vbBlue
                  Item.ListSubItems(4).ForeColor = vbBlue
                  Item.ListSubItems(5).ForeColor = vbBlue
                  Item.ListSubItems(6).ForeColor = vbBlue
                  Else
                     If Trim(TabTemp.Fields("st_item").Value) = "P" Then
                        Item.ForeColor = vbRed
                        Item.ListSubItems(1).ForeColor = vbRed
                        Item.ListSubItems(2).ForeColor = vbRed
                        Item.ListSubItems(3).ForeColor = vbRed
                        Item.ListSubItems(4).ForeColor = vbRed
                        Item.ListSubItems(5).ForeColor = vbRed
                        Item.ListSubItems(6).ForeColor = vbRed
                        Else
                           If Trim(TabTemp.Fields("st_item").Value) = "B" Then
                              Item.ForeColor = vbMagenta
                              Item.ListSubItems(1).ForeColor = vbMagenta
                              Item.ListSubItems(2).ForeColor = vbMagenta
                              Item.ListSubItems(3).ForeColor = vbMagenta
                              Item.ListSubItems(4).ForeColor = vbMagenta
                              Item.ListSubItems(5).ForeColor = vbMagenta
                              Item.ListSubItems(6).ForeColor = vbMagenta
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
         PESQUISA_Devolução_ENTRADA
      Case vbKeyF10
         Call lstPedidos_DblClick
      Case vbKeyF11
         FORMULA_REL = ""
         If Not IsNull(lstPedidos.SelectedItem.Text) Then
            FORMULA_REL = lstPedidos.SelectedItem.Text

            If Not IsNumeric(FORMULA_REL) Then _
               Exit Sub

            NUMR_REQ_N = FORMULA_REL

            FORMULA_REL = "{vwRelVenda.empresa_id} = " & EMPRESA_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.numr_req} = " & NUMR_REQ_N

            'If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_pedido_venda.rpt"
            frmRELATORIO10.Show 1
         End If

      Case vbKeyEscape
         Unload Me
   End Select
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub lstPedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedidos, ColumnHeader
End Sub

Private Sub lstPedidos_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedidos.SelectedItem.Text) Then
      If Trim(lstPedidos.SelectedItem.Text) <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select cgccpf from PEDIDO "
         SQL = SQL & " where numr_req = " & lstPedidos.SelectedItem.Text
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidos_Click"
End Sub

Private Sub lstPedidos_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedidos.SelectedItem.ListSubItems.Item(10).Text) Then
      If Trim(lstPedidos.SelectedItem.ListSubItems.Item(10).Text) <> "" Then

         '================================== PEDIDO DE VENDA
         If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.Item(9).Text)) = UCase("PEDIDO") Then _
            TIPO_NFe_GERAR = "S"          'Tipo Saida

         '================================== DEVOLUÇÃO DE SAIDA
         If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.Item(9).Text)) = UCase("DV") Then _
            TIPO_NFe_GERAR = "DV"          'DEVOLUÇÃO VENDA

         '================================== DEVOLUÇÃO DE ENTRADA
         If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.Item(9).Text)) = UCase("DVE") Then _
            TIPO_NFe_GERAR = "DC"

         '================================== DEVOLUÇÃO DE TRANSFERENCIA
         If UCase(Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.Item(9).Text)) = UCase("T") Then _
            TIPO_NFe_GERAR = "T"

         NUMR_REQ_N = lstPedidos.SelectedItem.ListSubItems.Item(10).Text
      
         If Trim(lstPedidos.SelectedItem.ListSubItems.Item(9).Text) = "DVE" Then 'Devolução de Entrada
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            SQL = "select * from NOTAENTRADA "
            SQL = SQL & " where entrada_id = " & NUMR_REQ_N
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCABECA.EOF Then
               If TabCABECA!Status = "D" Then
                  Msg = "Deseja Processar Devolução de Entrada?"
                  PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then _
                     GERA_NOTA_ENTRADA
               End If
            End If
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
         End If

         If Left(lstPedidos.SelectedItem.ListSubItems.Item(9).Text, 2) = "DV" Then 'Devolução de Saida
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            SQL = "select * from PEDIDO "
            SQL = SQL & " where numr_req = " & NUMR_REQ_N
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCABECA.EOF Then
               If TabCABECA!Status = 2 Then
                  Msg = "Deseja Processar Devolução de Venda NFE ?"
                  PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     DEVOLUCAO_VENDA_N = "D"

                     If USA_NFe = True Then _
                        GERA_NOTA
                  End If
               End If
            End If
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
         End If

         If UCase(lstPedidos.SelectedItem.ListSubItems.Item(9).Text) = UCase("pedido") Then 'PEDIDO VENDA
            DEVOLUCAO_VENDA_N = "N"
            FAZ_RECEBIMENTO
         End If

      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   PESQUISA_VENDA
   PESQUISA_Devolução_ENTRADA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidos_DblClick"
End Sub

Private Sub LISTAITEM_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP " ESC-Sair", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento", ""
   LISTAITEM.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTAITEM_DblClick"
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Private Sub PESQUISA_VENDA()
'On Error GoTo ERRO_TRATA

   SETA_GRID

   MOSTRA_TOP " ESC-Sair", "F6-Cancelar", " F7-Ver Itens", " F9-Atutalizar", " F10-Recebimento | F11-Imprimir Pedido"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PESQUISA_VENDA"
End Sub

Private Sub IMPRIME_TELA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from PEDIDO c"
   SQL = SQL & " where tipo_registro in ('S','R','D') "
   SQL = SQL & " and status in (2)" 'gerado somente Pedido"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by NUMR_REQ DESC "
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

Exit Sub
ERRO_TRATA:
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
   SQL = SQL & " where tipo_registro in ('S','R','D','OS') "
   SQL = SQL & " and status = 2" 'gerado somente Pedido"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by NUMR_REQ DESC "
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
      Set Item = lstPedidos.ListItems.Add(, "seq." & Trim(TabCABECA!NUMR_REQ), Trim(TabCABECA!NUMR_REQ))

      Item.SubItems(1) = "" & txtCNPJCPF.Text
      Item.SubItems(2) = "" & CRITERIO

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select descricao from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & TabCABECA.Fields("tipovenda_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            Item.SubItems(3) = "" & TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PERC_DESCONTO_N = 0 & TabCABECA.Fields("perc_desc").Value
      VALOR_DESCONTO_N = 0 & TabCABECA.Fields("valor_desconto").Value
      VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N

'========================================parceiro, tem que ver se pega pelo valor do desconto ou percentual
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
      SQL = SQL & " and tipo_reg = 'PC' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
'========================================

      VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

      'BUSCA VALOR TOTAL VENDA
      VALOR_ITEM_N = 0

      SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM "
      SQL = SQL & " where pedido_id = " & TabCABECA.Fields("pedido_id").Value
      SQL = SQL & " and tipo_reg = 'PC' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
'========================================

      Item.SubItems(4) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      Item.SubItems(5) = "" & Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      Item.SubItems(6) = "" & Format(VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      Item.SubItems(7) = "" & Trim(TabCABECA!DT_REQ)

'========================================
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select nome_vend from VENDEDOR v, EQUIPE e "
      SQL = SQL & " where v.vendedor_id = " & TabCABECA!VENDEDOR_ID
      SQL = SQL & " and e.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and v.codg_eq = e.codg_eq "
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then _
         Item.SubItems(8) = "" & TabVENDEDOR!NOME_VEND
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
'========================================

      If TabCABECA!TIPO_REGISTRO <> "D" Then
         Item.SubItems(9) = "Pedido"
         Else: Item.SubItems(9) = "DV"
      End If

      Item.SubItems(10) = "" & TabCABECA!NUMR_REQ
      Item.SubItems(12) = "" & TabCABECA.Fields("tipovenda_id").Value

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      CONT_N = CONT_N + 1
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""
      TabCABECA.MoveNext
   Wend
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   lstPedidos.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      If Not IsNull(TabCABECA!Status) Then
         If TabCABECA!Status <> "" Then
            If DEVOLUCAO_VENDA_N <> "D" Then
               If TabCABECA!Status = 5 Then
                  CRITERIO = NUMR_REQ_N
                  If TabCABECA.State = 1 Then _
                     TabCABECA.Close
                  frmNOTAGERA.Show 1
               End If
            Else
               If TabCABECA!Status = 2 Then
                  CRITERIO = NUMR_REQ_N
                  If TabCABECA.State = 1 Then _
                     TabCABECA.Close
                  frmNOTAGERA.Show 1
               End If
            End If
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
End Sub

Private Sub CHECA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   STATUS_A = ""
   SQL = "select * FROM PEDIDOITEM "
   SQL = SQL & " where numr_req = " & lstPedidos.SelectedItem.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      SP_PROCURA_PRODUTO EMPRESA_ID_N, Trim(TabPedidoItem!Codg_Prod), 0, "", "", "", 1
      If Not TabProduto.EOF Then _
         QTDE_ESTOQUE = TabProduto!Qtd 'Recebe so qtd. porque ja esta retido no pedido
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_ESTOQUE"
End Sub

Private Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   Dim TabPedido As New ADODB.Recordset

   If Not IsNull(lstPedidos.SelectedItem.Text) Then
      If Trim(lstPedidos.SelectedItem.Text) <> "" Then
         NUMR_REQ_N = lstPedidos.SelectedItem.Text
         SINAL_INDICADOR_N = 1

         If INDR_FORM_ABERTO = True Then
            Unload frmCADRECEBVENDA
            INDR_FORM_ABERTO = False
         End If

'===================================
'===================================
      If Not IsNull(lstPedidos.SelectedItem.ListSubItems.Item(12).Text) Then
         If Trim(lstPedidos.SelectedItem.ListSubItems.Item(12).Text) <> "" Then
            If IsNumeric(lstPedidos.SelectedItem.ListSubItems.Item(12).Text) Then

               If TabTemp.State = 1 Then _
                  TabTemp.Close
   
               SQL = "select contabiliza from TIPOVENDA "
               SQL = SQL & " where tipovenda_id = " & lstPedidos.SelectedItem.ListSubItems.Item(12).Text
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
      
                  If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
                     If TabTemp.Fields("contabiliza").Value = True Then
                        If TabTemp.State = 1 Then _
                           TabTemp.Close
      
               frmCADRECEBVENDA.Show 1
      
                        'Exit Sub
                        Else
                           SQL = "update PEDIDO set "
                           SQL = SQL & "status = 6 " 'não contabiliza
                           SQL = SQL & " where numr_req = " & NUMR_REQ_N
                           SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
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
'===================================
'===================================

         NUMR_REQ_N = lstPedidos.SelectedItem.Text
      
         If INDR_CONTROLA_ESTOQUE = False Then _
            Exit Sub

'===================================
         If USA_ECF = True Then
            CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
            Select Case CRITERIO
               Case "BEMATECH"
                  'Verifica se a Impressa esta ligada ou nao
                  Retorno = Bematech_FI_VerificaImpressoraLigada()
                  If Retorno <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
                     Retorno = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
                     MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!", vbCritical, "SHFSYS"
                     Exit Sub
                  End If

                  INDR_CUPOM_ABERTO = False
                  Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
                  If INDR_CUPOM_ABERTO = True Then _
                     CANCELA_CUPOM_ABERTO

                  Msg = ""
                  Indr_Erro = False
                  Call VerificaRetornoImpressora("", "", "Checando ECF")
                  If Indr_Erro = True Then
                     MsgBox Msg
                     Exit Sub
                  End If
               Case "DARUMA"
                  
               Case "Sweda"
                  
            End Select
         End If
'===================================
         If TabPedido.State = 1 Then _
            TabPedido.Close

         SQL = "select * from PEDIDO "
         SQL = SQL & " where numr_req = " & NUMR_REQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedido.EOF Then
            PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
            If TabPedido!Status = 5 Then
               CNPJCPF_A = Trim(TabPedido!CGCCPF)
   
               '====================================
               If USA_ECF = True Then
                  Msg = "Confirma Faturamento ?"
                  PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     '==============================
                     CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
                     Select Case CRITERIO
                        Case "BEMATECH"
                           'Verifica se a Impressa esta ligada ou nao
                           Retorno = Bematech_FI_VerificaImpressoraLigada()
                           If Retorno <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
                              Retorno = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
                              MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!!!", vbCritical, "SHFSYS"
                              Exit Sub
                              Else
                           End If
                        Case "DARUMA"
                           
                        Case "Sweda"
                           
                     End Select


'SET
'COMEÇA AQUI DE ACORDO O TIPO DA IMPRESSORA FISCAL
'ESSE AQUI É A ROTINA QUE NÃO É COM COMITANCIA
'DEPOIS TEM QUE FAZER O MESMO COM A TELA DE COMITANCIA

                     'BlockInput True   'Bloqueia o teclado
                        IMPRIME_CUPOM_FISCAL
                     BlockInput False  'Desbloqueia o teclado

                     '=======================
                     Me.WindowState = 0

                     If Trim(NUMEROCUPOM) <> "" And NUMR_REQ_N > 0 Then
                        SQL = "update PEDIDO set "
                        SQL = SQL & "status = 7 " 'CUPOM FISCAL
                        SQL = SQL & ", numr_cupom =  " & NUMEROCUPOM
                        SQL = SQL & " where numr_req = " & NUMR_REQ_N
                        SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                        CONECTA_RETAGUARDA.Execute SQL
                     End If
                     Else
                        If Trim(CNPJCPF_A) <> "99999999999" Then _
                           If USA_NFe = True Then _
                              GERA_NOTA
                  End If
                  Else
                     If Trim(CNPJCPF_A) <> "99999999999" Then _
                        If USA_NFe = True Then _
                           GERA_NOTA
               End If
'====================
CONTROLE_ESTOQUE_2  'CONTROLE
'====================
            End If
         End If
         If TabPedido.State = 1 Then _
            TabPedido.Close
      End If
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub

Private Sub PESQUISA_Devolução_ENTRADA()
'On Error GoTo ERRO_TRATA

   SETA_GRID_ENTRADA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PESQUISA_Devolução_ENTRADA"
End Sub

Private Sub SETA_GRID_ENTRADA()
'On Error GoTo ERRO_TRATA

   NOTAENTRADA_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT NOTAENTRADA.TIPOENTRADA_ID, NOTAENTRADA.STATUS, NOTAENTRADA.ENTRADA_ID, "
   SQL = SQL & " NOTAENTRADA.EMPRESA_ID, NOTAENTRADA.FORNECEDOR_ID, NOTAENTRADA.TRANSP_ID, "
   SQL = SQL & " NOTAENTRADA.NUMR_PEDIDO_COMPRA, NOTAENTRADA.NUMR_NOTA, NOTAENTRADA.SERIE_NOTA, "
   SQL = SQL & " NOTAENTRADA.DT_ENTRADA, NOTAENTRADA.DT_EMISSAO, NOTAENTRADA.VALOR_FRETE, NOTAENTRADA.VALOR_DESCONTO,"
   SQL = SQL & " FORNECEDOR.PESSOA_ID, FORNECEDOR.CGCCPF, FORNECEDOR.NOME, FORNECEDOR.RAZAO_SOCIAL "
   SQL = SQL & " FROM NOTAENTRADA "
   SQL = SQL & " INNER JOIN FORNECEDOR "
   SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

   SQL = SQL & " where tipoentrada_id = 1 "
   SQL = SQL & " and NOTAENTRADA.STATUS = 'D'" 'Devolução de Entrada
   SQL = SQL & " and NOTAENTRADA.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by entrada_id asc "

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NOTAENTRADA_ID_N = TabTemp.Fields("entrada_id").Value

      CRITERIO = Trim(TabTemp!NOME)
      txtCNPJCPF.PromptInclude = False
      If Len(Trim(TabTemp!CGCCPF)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If

      txtCNPJCPF.Text = TabTemp!CGCCPF
      txtCNPJCPF.PromptInclude = True
      CNPJCPF_A = TabTemp!CGCCPF

      Set Item = lstPedidos.ListItems.Add(, "seq." & Trim(TabTemp!NUMR_PEDIDO_COMPRA), Trim(TabTemp!NUMR_PEDIDO_COMPRA))
      Item.SubItems(1) = txtCNPJCPF.Text
      Item.SubItems(2) = CRITERIO

      PERC_DESCONTO_N = 0
      VALOR_DESCONTO_N = 0
      VALOR_TOTAL_DESCONTO_N = 0
      VALOR_DESCONTO_CABECA_N = 0 & TabTemp.Fields("valor_desconto").Value
      VALOR_ITEM_N = 0

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((qtd_entrada*preco_custo)) from NOTAENTRADAITEM "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      VALOR_DESCONTO_N = VALOR_DESCONTO_CABECA_N
      
      Item.SubItems(4) = Format(Trim(VALOR_ITEM_N), strFormatacao2Digitos)
      Item.SubItems(5) = Format(Trim(VALOR_DESCONTO_N), strFormatacao2Digitos)
      Item.SubItems(6) = "" & Format(VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      Item.SubItems(7) = Trim(TabTemp!DT_EMISSAO)
      Item.SubItems(8) = NOME_EMPRESA
      Item.SubItems(9) = "Devolução Entrada"
      Item.SubItems(9) = "DVE"
      Item.SubItems(10) = TabTemp.Fields("ENTRADA_ID").Value

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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ENTRADA"
End Sub

Private Sub GERA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   NUMR_REQ_N = lstPedidos.SelectedItem.ListSubItems.Item(10).Text

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status from NOTAENTRADA "
   SQL = SQL & " where entrada_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA_ENTRADA"
End Sub

Sub GRAVA_CUPOM(Numero_Pedido As String, Numero_Cupom As String)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'GRAVA TABELA CUPOM
   SQL = "select * from CUPOM"
   SQL = SQL & " where numr_cupom = " & Numero_Pedido
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      SqL2 = "update CUPOM set "
         SqL2 = SqL2 & " VALOR_CUPOM = " & tpMOEDA(VALOR_TOTAL_N)               'VALOR_CUPOM
         SqL2 = SqL2 & ", IMPRESSORA_ID = " & IMPRESSORA_ID_N                   'IMPRESSORA_ID
         SqL2 = SqL2 & ", Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO   'Numr_Contador_Reinicio
      SqL2 = SqL2 & " where numr_cupom = " & Numero_Cupom
      SqL2 = SqL2 & " and Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO   'Numr_Contador_Reinicio
      Else
         SqL2 = "insert into CUPOM "
         SqL2 = SqL2 & " (CUPOM_ID,NUMR_CUPOM,PEDIDO_ID,VALOR_CUPOM,IMPRESSORA_ID,Numr_Contador_Reinicio)"
         SqL2 = SqL2 & " VALUES("
            SqL2 = SqL2 & MAX_ID("cupom_id", "cupom", "", "", "", "")  'CUPOM_ID
            SqL2 = SqL2 & "," & Numero_Cupom                           'NUMR_CUPOM
            SqL2 = SqL2 & "," & Numero_Pedido                          'PEDIDO_ID
            SqL2 = SqL2 & "," & tpMOEDA(VALOR_TOTAL_N)                 'VALOR_CUPOM
            SqL2 = SqL2 & "," & IMPRESSORA_ID_N                        'IMPRESSORA_ID
            SqL2 = SqL2 & "," & NUMR_CONTADOR_REINICIO                 'Numr_Contador_Reinicio
         SqL2 = SqL2 & ")"
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   CONECTA_RETAGUARDA.Execute SqL2
End Sub

Sub BUSCA_CLIENTE(CLIENTE_ID As Long)
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
End Sub


Private Sub CONTROLE_ESTOQUE()
'On Error GoTo ERRO_TRATA

   Dim QTDE_BALCAO_N       As Double
   Dim QTDE_ESTOQUE_N      As Double

   Dim QTDE_PEDIDA_N       As Double
   Dim QTDE_DISPONIVEL_N   As Double

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * FROM PEDIDOITEM "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      QTDE_PEDIDA_N = 0 & TabPedidoItem!QTD_PEDIDA

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(TabPedidoItem!Codg_Prod) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         QTDE_DISPONIVEL_N = 0 & TabProduto!QTDE

         If QTDE_DISPONIVEL_N >= QTDE_PEDIDA_N Then

            QTDE_BALCAO_N = 0 & TabProduto!QTDE_RETIDO

            If QTDE_BALCAO_N > 0 Then
               QTDE_BALCAO_N = QTDE_BALCAO_N - QTDE_PEDIDA_N
               QTDE_ESTOQUE_N = QTDE_DISPONIVEL_N - QTDE_PEDIDA_N
               Else: QTDE_BALCAO_N = 0 'Retido Negativo
            End If
            Else: QTDE_ESTOQUE_N = 0 ' Quantida em Estoque e Menor que Quantidade Pedida
         End If

         If QTDE_BALCAO_N < 0 Then _
            QTDE_BALCAO_N = 0
         If QTDE_ESTOQUE_N < 0 Then _
            QTDE_ESTOQUE_N = 0
      End If

      If TabProduto.State = 1 Then _
         TabProduto.Close

      TabPedidoItem.MoveNext
   Wend
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

Exit Sub
ERRO_TRATA:
   MsgBox "ATENÇÃO, ERRO NA BAIXA DE ESTOQUE, INFORMAR IMEDIATAMENTE AO SUPORTE"
   TRATA_ERROS Err.Description, Me.Name, "CONTROLE_ESTOQUE"
End Sub

Sub CONTROLE_ESTOQUE_2()
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select produto_id, qtd_pedida, seq_id FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and status <> 'B' "
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF

'=============baixa estoque INICIO
      PRODUTO_ID_N = TabPedidoItem.Fields("produto_id").Value
      QTDE_PEDIDO = TabPedidoItem.Fields("QTD_PEDIDA").Value
      QTDE_RETIDO = TabPedidoItem.Fields("QTD_PEDIDA").Value

      BAIXA_ESTOQUE_PRODUTO PRODUTO_ID_N, QTDE_PEDIDO, QTDE_RETIDO
'=============baixa estoque FIM

      SQL = "UPDATE PEDIDOITEM set "
      SQL = SQL & " status = 'B' "
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and status <> 'B' "
      SQL = SQL & " and produto_id = " & Trim(TabPedidoItem.Fields("produto_id").Value)
      CONECTA_RETAGUARDA.Execute SQL

      TabPedidoItem.MoveNext
   Wend
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close
End Sub

Sub CANCELA_CUPOM_ABERTO()
'On Error GoTo ERRO_TRATA

   CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
   Select Case CRITERIO
      Case "BEMATECH"
         'Retorno = Bematech_FI_NumeroCupom(NUMR_CUPOM_ABERTO)
         Retorno = NUMR_CUPOM_ABERTO

         Indr_Erro = False

         Retorno = Bematech_FI_CancelaCupom()
         'Função que analisa o retorno da impressora
         Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      Case "DARUMA"
         'Retorno = Bematech_FI_NumeroCupom(NUMR_CUPOM_ABERTO)
         Retorno = NUMR_CUPOM_ABERTO

         Indr_Erro = False

         'Retorno = iCFCancelar_ECF_Daruma()
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
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CANCELA_CUPOM_ABERTO"
End Sub

Sub IMPRIME_CUPOM_FISCAL()
   CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
   Select Case CRITERIO
      Case "BEMATECH"
         ROTINA_CUPOM_FISCAL_BEMATECH
      Case "DARUMA"
         ROTINA_CUPOM_FISCAL_DARUMA
      Case "Sweda"
         ROTINA_CUPOM_FISCAL_SWEDA
   End Select
End Sub

