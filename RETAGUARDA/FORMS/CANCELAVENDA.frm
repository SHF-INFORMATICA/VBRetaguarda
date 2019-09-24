VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVENDACANCELA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento Pedido"
   ClientHeight    =   3840
   ClientLeft      =   3525
   ClientTop       =   2895
   ClientWidth     =   8490
   Icon            =   "CANCELAVENDA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8490
   Begin VB.TextBox txtDoc 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtVendedor 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3360
      Width           =   6495
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtForma 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   3570
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   2760
      Width           =   4215
   End
   Begin MSMask.MaskEdBox DTEMIS 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin VB.TextBox txtPedido 
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
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin MSMask.MaskEdBox DTCANCELA 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox CGCCPF 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4830
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":47C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":4809E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":483BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":4880E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":48C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":48F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":493D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":496F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":4A108
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":4AB1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CANCELAVENDA.frx":4B52C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C&onsultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6000
         Top             =   240
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
               Picture         =   "CANCELAVENDA.frx":4BF3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CANCELAVENDA.frx":4D366
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CANCELAVENDA.frx":4E3F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CANCELAVENDA.frx":4FAF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CANCELAVENDA.frx":50BFD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NºDoc.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5010
      TabIndex        =   15
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cancela:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   915
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmVENDACANCELA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim INDR_CONTINUA As Boolean

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
         CRITERIO = ""
         frmVENDACONSULTA.Show 1
         If IsNumeric(CRITERIO) Then _
            txtPedido.Text = CRITERIO
         CRITERIO = ""
      Case "gravar"
         If Not IsNumeric(txtPedido.Text) Then
            MsgBox "Informe número de Pedido."
            Exit Sub
         End If

         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "select * from PEDIDO "
         SQL = SQL & " where numr_req = " & txtPedido.Text
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If TabCABECA!Status = 9 Then
               
               If TabCABECA.State = 1 Then _
                  TabCABECA.Close

               MsgBox "Esse registro já foi cancelado !!!"
               txtPedido.SetFocus
               Exit Sub
               Else
                  Msg = "Confirma cancelamento ?"
                  Style = vbYesNo + 32
                  Title = "Atenção !!!"
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                     If USA_ECF = True Then
                        'dai é cupom por isso pode fazer o check
                        If TabCABECA.Fields("status").Value = 7 Then
                           INDR_CONTINUA = False
                           CANCELA_CUPOM_FISCAL
                           'If INDR_CONTINUA = False Then _
                              Exit Sub
                        End If
                     End If

                     'VOLTANDO PRODUTO PARA ESTOQUE
                     If TabCABECA!TIPO_REGISTRO = "R" Then
                        If INDR_CONTROLA_ESTOQUE = True Then

                           If TabPedidoItem.State = 1 Then _
                              TabPedidoItem.Close

                           SQL = "select * FROM PEDIDOITEM "
                           SQL = SQL & " where pedido_id = " & txtPedido.Text
                           SQL = SQL & " and tipo_reg = 'PC' "
                           TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           While Not TabPedidoItem.EOF
                              If TabProduto.State = 1 Then _
                                 TabProduto.Close

                              SP_PROCURA_PRODUTO EMPRESA_ID_N, TabPedidoItem!Codg_Prod, 0, "", "", "", -1
                              If Not TabProduto.EOF Then
                                 'se baixa estoque na Pedido então tem que
                                 'subtrair da quantidade em estoque diretamente
                                 If INDR_BAIXA_ESTQ_PEDIDO = True Then
                                    If Not IsNull(TabProduto!QTDE) Then

                                       'If TabProduto!qtde >= TabPedidoItem!qtd_pedida Then
                                          SQL = "update PRODUTO "
                                          SQL = SQL & "set qtde =  " & tpMOEDA(TabProduto!QTDE + TabPedidoItem!QTD_PEDIDA)
                                          SQL = SQL & " where codg_produto = '" & TabPedidoItem!Codg_Prod & "'"
                                          SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                                          CONECTA_RETAGUARDA.Execute SQL
                                       'End If

                                    End If
                                    Else  'se INDR_BAIXA_ESTQ_PEDIDO = False então é venda com documento fiscal
                                       If TabNOTA.State = 1 Then _
                                          TabNOTA.Close

                                       SQL = "select * from NF "
                                       SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
                                       SQL = SQL & " and numr_req = " & txtPedido.Text
                                       TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                                       If Not TabNOTA.EOF Then 'se a nota foi emitida baixa na quantidade
                                          If TabNOTA!Status = "E" Then
                                             If Not IsNull(TabProduto!QTDE) Then

                                                'If TabPedidoItem!qtd_pedida <= TABPRODUTO!Qtde Then
                                                   SQL = "update PRODUTO "
                                                   SQL = SQL & "set qtde =  " & tpMOEDA(TabProduto!QTDE + TabPedidoItem!QTD_PEDIDA)
                                                   SQL = SQL & " where codg_produto = '" & TabPedidoItem!Codg_Prod & "'"
                                                   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                                                   CONECTA_RETAGUARDA.Execute SQL
                                                'End If

                                             End If
                                          End If
                                          Else 'se a nota não foi emitida baixa retido somente
                                             If TabCABECA!Status = 1 Then  'baixando somente retido
                                                If Not IsNull(TabProduto!QTDE_RETIDO) Then
                                                   SQL = "update PRODUTO "
                                                   SQL = SQL & "set qtde_retido = qtde_retido - " & tpMOEDA(TabPedidoItem!QTD_PEDIDA)
                                                   SQL = SQL & " where codg_produto = '" & TabPedidoItem!Codg_Prod & "'"
                                                   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                                                   CONECTA_RETAGUARDA.Execute SQL
                                                End If
                                                Else
                                                   If TabCABECA!Status = 2 Then
                                                      If Not IsNull(TabProduto!QTDE_RETIDO) Then
                                                         SQL = "update PRODUTO "
                                                         SQL = SQL & "set qtde_retido = " & tpMOEDA(TabProduto!QTDE_RETIDO - TabPedidoItem!QTD_PEDIDA)
                                                         SQL = SQL & " where codg_produto = '" & TabPedidoItem!Codg_Prod & "'"
                                                         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                                                         CONECTA_RETAGUARDA.Execute SQL
                                                      End If
                                                      Else 'Se ja foi emitido soma no estoque atual
                                                         SQL = "update PRODUTO "
                                                         SQL = SQL & "set qtde = " & tpMOEDA(TabProduto!QTDE + TabPedidoItem!QTD_PEDIDA)
                                                         SQL = SQL & " where codg_produto = '" & TabPedidoItem!Codg_Prod & "'"
                                                         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                                                         CONECTA_RETAGUARDA.Execute SQL
                                                   End If
                                             End If
                                       End If
                                 End If
                              End If
                              If TabProduto.State = 1 Then _
                                 TabProduto.Close

                              TabPedidoItem.MoveNext
                           Wend
                           If TabPedidoItem.State = 1 Then _
                              TabPedidoItem.Close

                           'Gravando Status de Pedido Cancelado
                           SQL = "UPDATE PEDIDO SET "
                           SQL = SQL & " Status = " & 9
                           SQL = SQL & " where numr_req = " & txtPedido.Text
                           SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                           CONECTA_RETAGUARDA.Execute SQL

                           SQL = "UPDATE NF SET "
                           SQL = SQL & " Status = 'C'"
                           SQL = SQL & " where numr_req = " & txtPedido.Text
                           SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                           CONECTA_RETAGUARDA.Execute SQL
                        End If

                        'LANÇAMENTOS
                        If TabLancamento.State = 1 Then _
                           TabLancamento.Close

                        SQL = "select * from LANCAMENTO "
                        SQL = SQL & " where numr_doc = " & txtPedido.Text
                        SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                        TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If Not TabLancamento.EOF Then
                           Msg = "Deseja Cancelar Lançamento Contas a Receber?"
                           Style = vbYesNo + 32
                           Title = "Atenção !!!"
                           Help = "DEMO.HLP"
                           Ctxt = 1000
                           RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                           If RESPOSTA = vbYes Then
                              'SQL = "delete  from CAIXADIAITEM c, CAIXADIA p "
                              'SQL = SQL & " where c.numr_doc = " & txtPedido.Text
                              'SQL = SQL & " and p.empresa_id = " & EMPRESA_ID_n_n
                              'SQL = SQL & " and p.caixa_id = c.caixa_id "
                              'CONECTA_RETAGUARDA.execute SQL

                              'isso aqui nao precisa pois ja esta cancelando os itens
                              'CONECTA_RETAGUARDA.execute "UPDATE LANCAMENTO SET Tipo_Lancamento = " & 9 & " where numr_doc = " & txtPedido.Text & " and empresa_id = " & EMPRESA_ID_n_n

                              If TabTemp.State = 1 Then _
                                 TabTemp.Close

                              SQL = "select * from ITEMLANCAMENTO "
                              SQL = SQL & " where lancamento_id = " & TabLancamento!LANCAMENTO_ID
                              TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                              While Not TabTemp.EOF
                                 SQL = "UPDATE ITEMLANCAMENTO SET "
                                 SQL = SQL & " usu_alt = " & CODG_USU_N
                                 SQL = SQL & ", dt_alt = '" & DMA(Date) & "'"
                                 SQL = SQL & ", dt_baixa = '" & DMA(Date) & "'"
                                 SQL = SQL & ", Status = '" & "C" & "'"
                                 SQL = SQL & ", CODG_USU_BAIXA = " & CODG_USU_N
                                 SQL = SQL & " where lancamento_id = " & TabLancamento!LANCAMENTO_ID
                                 CONECTA_RETAGUARDA.Execute SQL
                                 TabTemp.MoveNext
                              Wend
                              If TabTemp.State = 1 Then _
                                 TabTemp.Close
                           End If
                           If TabLancamento.State = 1 Then _
                              TabLancamento.Close
                        End If
                     End If
                     If TabCABECA!TIPO_REGISTRO = "O" Then _
                        RESPOSTA = "Orçamento"
                     If TabCABECA!TIPO_REGISTRO = "R" Then _
                        RESPOSTA = "Pedido"
                     MsgBox RESPOSTA & " foi cancelado com sucesso."
                  End If
            End If
            Else
               If TabCABECA.State = 1 Then _
                  TabCABECA.Close

               MsgBox "Orçamento ou Pedido inexistente !!!"
               LIMPA_CANC
               txtPedido.SetFocus
               Exit Sub
         End If
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         LIMPA_CANC
      Case "limpar"
         LIMPA_CANC
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtPedido.Text = "" Then _
         Exit Sub

      If IsNull(txtPedido.Text) Then _
         Exit Sub

      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where numr_req = " & txtPedido.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCABECA.EOF Then
         DTEMIS.Text = TabCABECA!DT_REQ
         DTCANCELA.PromptInclude = False
            DTCANCELA.Text = Date
         DTCANCELA.PromptInclude = True
         CGCCPF.Text = TabCABECA!CGCCPF

txtDoc.Text = "" & TabCABECA.Fields("numr_doc").Value

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select * from CLIENTE "
         SQL = SQL & " where CGCCPF = '" & TabCABECA!CGCCPF & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then _
            txtNome.Text = TabCliente!NOME
         If TabCliente.State = 1 Then _
            TabCliente.Close

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select * from TIPOVENDA "
         SQL = SQL & " where tipovenda_id = " & TabCABECA!TIPOVENDA_ID
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            txtForma.Text = TabDESCR!Descricao
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         If TabCABECA!Status = 1 Then _
            txtStatus.Text = "ABERTO"
         If TabCABECA!Status = 3 Then _
            txtStatus.Text = "ATUALIZADO"
         If TabCABECA!Status = 3 Then _
            txtStatus.Text = "RECEBIDO"
         If TabCABECA!Status = 9 Then _
            txtStatus.Text = "CANCELADO"

         txtVendedor.Text = ""

         SQL = "select nome_vend from VENDEDOR "
         SQL = SQL & " where vendedor_id = " & TabCABECA.Fields("vendedor_id").Value
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            txtVendedor.Text = "" & TabDESCR.Fields(0).Value
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select * FROM PEDIDOITEM i, PRODUTO p "
         SQL = SQL & " where i.numr_req = " & txtPedido.Text
         SQL = SQL & " and i.codg_prod = p.CODG_PRODUTO "
         SQL = SQL & " and p.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and i.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and i.tipo_reg = 'PC' "
         Else
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            MsgBox "Orçamento ou Pedido inexistente !!!"
            txtPedido.SetFocus
      End If
   End If
End Sub

Private Sub LIMPA_CANC()
   txtDoc.Text = ""
   txtVendedor.Text = ""
   txtPedido.Text = ""
   DTEMIS.Text = ""
   DTCANCELA.Text = ""
   txtForma.Text = ""
   CGCCPF.Text = ""
   txtNome.Text = ""
   txtStatus.Text = ""
   txtPedido.SetFocus
End Sub

Sub CANCELA_CUPOM_FISCAL()
'On Error GoTo ERRO_TRATA

   Dim NUMEROCUPOM As String
   Dim NUMR_CUPOM_CABECA As Long
   Dim NUMR_ULTIMO_CUPOM_TABELA As Long

   NUMEROCUPOM = 0

   Dim RETORNOSTATUS As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NUMEROCUPOM = Space(1)
      Else: NUMEROCUPOM = Space(6)
   End If

   Retorno = Bematech_FI_NumeroCupom(NUMEROCUPOM)
   'Função que analisa o retorno da impressora
   Call VerificaRetornoImpressora("Número do Último Cupom: ", _
        NUMEROCUPOM, "Informações da Impressora")

   If TABCUPOM.State = 1 Then _
      TABCUPOM.Close

   SQL = "select max(numr_cupom) from CUPOM"
   SQL = SQL & " where Numr_Contador_Reinicio = " & Numr_Contador_Reinicio
   TABCUPOM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TABCUPOM.EOF Then
      NUMR_CUPOM_CABECA = TabCABECA.Fields("numr_cupom").Value
      NUMR_ULTIMO_CUPOM_TABELA = TABCUPOM.Fields(0).Value
      If NUMR_CUPOM_CABECA = NUMR_ULTIMO_CUPOM_TABELA Then
         If Str(NUMEROCUPOM) <> Str(TABCUPOM.Fields(0).Value) Then _
            Exit Sub

         PERGUNTA "Confirma cancelamento cupom fiscal número = " & NUMEROCUPOM, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
         If RESPOSTA = vbNo Then _
            Exit Sub

         Indr_Erro = False

         Retorno = Bematech_FI_CancelaCupom()
         'Função que analisa o retorno da impressora
         Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")

         If Indr_Erro = False Then _
            NUMR_ID_N = 0
      End If
   End If
   If TABCUPOM.State = 1 Then _
      TABCUPOM.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CANCELA_CUPOM_FISCAL"
End Sub
