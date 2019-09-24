VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoCancela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento Pedido"
   ClientHeight    =   3840
   ClientLeft      =   3525
   ClientTop       =   2895
   ClientWidth     =   8490
   Icon            =   "PEDIDOCANCELA.frx":0000
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
   Begin MSMask.MaskEdBox txtCNPJCPF 
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
            Picture         =   "PEDIDOCANCELA.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":739E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":76BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":80D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":8AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PEDIDOCANCELA.frx":94F4
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
      ButtonWidth     =   2858
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
               Picture         =   "PEDIDOCANCELA.frx":9F06
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCANCELA.frx":B32E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCANCELA.frx":C3BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCANCELA.frx":DABA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCANCELA.frx":EBC5
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
Attribute VB_Name = "frmPedidoCancela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim INDR_CONTINUA As Boolean

Private Sub Form_Activate()

   If Trim(txtPedido.Text) <> "" Then _
      If IsNumeric(txtPedido.Text) Then _
         Call txtpedido_KeyPress(13)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         CRITERIO_A = ""
         CNPJCPF_A = ""
         frmPedidoConsulta.Show 1
         If IsNumeric(CRITERIO_A) Then _
            txtPedido.Text = CRITERIO_A
         CRITERIO_A = ""
      Case "gravar"
         If Not IsNumeric(txtPedido.Text) Then
            MsgBox "Informe número de Pedido."
            Exit Sub
         End If

         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from PEDIDO "
         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            If TabCabeca!STATUS = 9 Then
               
               If TabCabeca.State = 1 Then _
                  TabCabeca.Close

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
                     'VOLTANDO PRODUTO PARA ESTOQUE
                     If TabCabeca!TIPO_REGISTRO = "P" Or TabCabeca!TIPO_REGISTRO = "R" Or TabCabeca!TIPO_REGISTRO = "OS" Or Left(TabCabeca!TIPO_REGISTRO, 1) = "D" Then
                        If INDR_CONTROLA_ESTOQUE = True Then
                           'só vai bulir quando for nota, cupom ou faturado
                           If TabCabeca!STATUS = 3 Or TabCabeca!STATUS = 5 Or TabCabeca!STATUS = 7 Then
                              If TabPedidoItem.State = 1 Then _
                                 TabPedidoItem.Close
   
                              SQL = "select * from PEDIDOITEM "
                              SQL = SQL & " where pedido_id = " & txtPedido.Text
                              SQL = SQL & " and tipo_reg = 'PC' "
                              SQL = SQL & " and pedidoitem.status <> 'C' "
                              TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                              While Not TabPedidoItem.EOF
                                 SQL = "update ESTOQUE set "
                                 SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(TabPedidoItem!QTD_PEDIDA)
                                 SQL = SQL & " where produto_id = " & TabPedidoItem.Fields("produto_id").Value
                                 SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                                 CONECTA_RETAGUARDA.Execute SQL
   
                                 TabPedidoItem.MoveNext
                              Wend
                              If TabPedidoItem.State = 1 Then _
                                 TabPedidoItem.Close
                           End If   'If TabCABECA!Status = 3 Or TabCABECA!Status = 5 Or TabCABECA!Status = 7 Then

                           'Gravando Status de Pedido Cancelado
                           SQL = "UPDATE PEDIDO SET "
                           SQL = SQL & " Status = " & 9
                           SQL = SQL & " where pedido_id = " & txtPedido.Text
                           SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                           CONECTA_RETAGUARDA.Execute SQL

                           SQL = "UPDATE NF SET "
                           SQL = SQL & " Status = 'C'"
                           SQL = SQL & " from Nf "
                           SQL = SQL & " INNER JOIN PEDIDONF"
                           SQL = SQL & " ON NF.NF_ID = PEDIDONF.NF_ID"
                           SQL = SQL & " where pedido_id = " & txtPedido.Text
                           CONECTA_RETAGUARDA.Execute SQL
                        End If   'If INDR_CONTROLA_ESTOQUE = True Then

                        'LANÇAMENTOS
                        If TabLancamento.State = 1 Then _
                           TabLancamento.Close

                        SQL = "select * from LANCAMENTO "
                        SQL = SQL & " where numr_doc = " & txtPedido.Text
                        SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                        TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If Not TabLancamento.EOF Then
                           Msg = "Deseja Cancelar Lançamento Contas a Receber?"
                           Style = vbYesNo + 32
                           Title = "Atenção !!!"
                           Help = "DEMO.HLP"
                           Ctxt = 1000
                           'RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                           RESPOSTA = vbYes
                           If RESPOSTA = vbYes Then
                              If TabTemp.State = 1 Then _
                                 TabTemp.Close

                              SQL = "select * from ITEMLANCAMENTO "
                              SQL = SQL & " where lancamento_id = " & TabLancamento!LANCAMENTO_ID
                              TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                              While Not TabTemp.EOF
                                 SQL = "UPDATE ITEMLANCAMENTO SET "
                                 SQL = SQL & " usu_alt = " & USUARIO_ID_N
                                 SQL = SQL & ", dt_alt = '" & Now & "'"
                                 SQL = SQL & ", dt_cancela = '" & Now & "'"
                                 SQL = SQL & ", Status = '" & "C" & "'"
                                 SQL = SQL & ", CODG_USU_BAIXA = " & USUARIO_ID_N
                                 SQL = SQL & " where lancamento_id = " & TabLancamento!LANCAMENTO_ID
                                 CONECTA_RETAGUARDA.Execute SQL

                                 TabTemp.MoveNext
                              Wend
                              If TabTemp.State = 1 Then _
                                 TabTemp.Close
                           End If
                           If TabLancamento.State = 1 Then _
                              TabLancamento.Close
                        End If   'If TabCABECA!Tipo_Registro = "R" Then
                     End If
                     If TabCabeca!TIPO_REGISTRO = "O" Then
                        RESPOSTA = "Orçamento"
                        MsgBox RESPOSTA & " foi cancelado com sucesso."
                     End If
                     If TabCabeca!TIPO_REGISTRO = "R" Then
                        RESPOSTA = "Pedido"
                        MsgBox RESPOSTA & " foi cancelado com sucesso."
                     End If
                     If TabCabeca!TIPO_REGISTRO = "OS" Then
                        RESPOSTA = "Pedido"
                        MsgBox RESPOSTA & " foi cancelado com sucesso."
                     End If
                  End If   'If RESPOSTA = vbYes Then
            End If
            Else
               If TabCabeca.State = 1 Then _
                  TabCabeca.Close

               MsgBox "Orçamento ou Pedido inexistente !!!"
               LIMPA_CANC
               txtPedido.SetFocus
               Exit Sub
         End If
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         LIMPA_CANC
         Unload Me
      Case "limpar"
         LIMPA_CANC
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtPedido.Text = "" Then _
         Exit Sub

      If IsNull(txtPedido.Text) Then _
         Exit Sub

      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where pedido_id = " & txtPedido.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabeca.EOF Then
         DTEMIS.Text = TabCabeca!DT_REQ
         DTCANCELA.PromptInclude = False
            DTCANCELA.Text = Date
         DTCANCELA.PromptInclude = True
         txtCNPJCPF.Text = TabCabeca!CGCCPF

         txtDoc.Text = ""

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select * from CLIENTE "
         SQL = SQL & " where CGCCPF = '" & TabCabeca!CGCCPF & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then _
            txtNome.Text = Trim(TabCliente!NOME)
         If TabCliente.State = 1 Then _
            TabCliente.Close

         If Not IsNull(TabCabeca.Fields("nome_cliente").Value) Then _
            If Trim(TabCabeca.Fields("nome_cliente").Value) <> "" Then _
               txtNome.Text = Trim(TabCabeca.Fields("nome_cliente").Value)

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         'SQL = "select * from TIPOVENDA "
         'SQL = SQL & " where tipovenda_id = " & TabCabeca!TIPOVENDA_ID
         'TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         'If Not TabDESCR.EOF Then _
            txtForma.Text = TabDESCR!DESCRICAO
         'If TabDESCR.State = 1 Then _
            TabDESCR.Close

         If TabCabeca!STATUS = 1 Then _
            txtStatus.Text = "ABERTO"
         If TabCabeca!STATUS = 3 Then _
            txtStatus.Text = "ATUALIZADO"
         If TabCabeca!STATUS = 3 Then _
            txtStatus.Text = "RECEBIDO"
         If TabCabeca!STATUS = 9 Then _
            txtStatus.Text = "CANCELADO"

         txtVendedor.Text = ""

         SQL = "select descricao from vwVendedor "
         SQL = SQL & " where vendedor_id = " & TabCabeca.Fields("vendedor_id").Value
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            txtVendedor.Text = "" & TabDESCR.Fields(0).Value
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select * from PEDIDOITEM "
         SQL = SQL & " INNER JOIN PRODUTO "
         SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "

         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and tipo_reg = 'PC' "
         SQL = SQL & " and status <> 'C' "
         Else
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            MsgBox "Orçamento ou Pedido inexistente !!!"
            txtPedido.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpedido_KeyPress"
End Sub

Private Sub LIMPA_CANC()
   txtDoc.Text = ""
   txtVendedor.Text = ""
   txtPedido.Text = ""
   DTEMIS.Text = ""
   DTCANCELA.Text = ""
   txtForma.Text = ""
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtStatus.Text = ""
   txtPedido.SetFocus
End Sub

