VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPEDIDODETALHE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destalhamento Itens Produção"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PEDIDODETALHE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRef 
      Height          =   1815
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4080
      Width           =   7575
   End
   Begin VB.TextBox txtObs 
      Height          =   1815
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   7575
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   1230
      TabIndex        =   0
      ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   3720
      MaxLength       =   29
      TabIndex        =   1
      Top             =   1440
      Width           =   5655
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   6
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
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDODETALHE.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDODETALHE.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDODETALHE.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDODETALHE.frx":97C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDODETALHE.frx":A8D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDODETALHE.frx":BA6B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   10560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Referência:"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   405
      TabIndex        =   7
      Top             =   4080
      Width           =   1320
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Observações:"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   2040
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Destalhamento Itens Produção"
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
      TabIndex        =   5
      Top             =   720
      Width           =   9540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   10560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   1440
      Width           =   1020
   End
End
Attribute VB_Name = "frmPEDIDODETALHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   LIMPA_TUDO
   MOSTRA_DADOS
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "gravar"
         If Trim(txtOBS.Text) <> "" And Trim(txtProduto.Text) <> "" Then
            GRAVA_DADOS
            Else
               If Trim(txtRef.Text) <> "" And Trim(txtProduto.Text) <> "" Then _
                  GRAVA_DADOS
         End If
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
   End Select
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtRef.Text = ""
   txtOBS.Text = ""
   'txtObs.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub MOSTRA_DADOS()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
   SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_ID_N
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtProduto.Text = "" & Trim(TabTemp.Fields("codg_produto").Value)
      txtDescricao.Text = "" & Trim(TabTemp.Fields("descricao").Value)
      txtProduto.Text = "" & Trim(TabTemp.Fields("codg_produto").Value)

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select PEDIDOITEMOBS.QTDE, PEDIDOITEMOBS.Valor, "
      SQL = SQL & " PEDIDOITEMOBS.OBS , PEDIDOITEMOBS.Referencia"
      SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDOITEMOBS WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = PEDIDOITEMOBS.PEDIDO_ID "
      SQL = SQL & " AND PEDIDOITEM.SEQ_ID = PEDIDOITEMOBS.SEQ_ID "
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where PEDIDOITEM.pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_ID_N
      'sql=sql & " and pedidoitem.statud <> 'C'
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         txtRef.Text = "" & Trim(TabTemp.Fields("referencia").Value)
         txtOBS.Text = "" & Trim(TabTemp.Fields("obs").Value)
      End If
      Else: MsgBox "Registro não encontrado."
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS"
End Sub

Sub GRAVA_DADOS()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, "
   SQL = SQL & " PRODUTO.DESCRICAO, PEDIDOITEMOBS.QTDE, PEDIDOITEMOBS.Valor, "
   SQL = SQL & " PEDIDOITEMOBS.OBS , PEDIDOITEMOBS.Referencia"
   SQL = SQL & " from PEDIDOITEM "
   SQL = SQL & " INNER JOIN PEDIDOITEMOBS "
   SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = PEDIDOITEMOBS.PEDIDO_ID "
   SQL = SQL & " AND PEDIDOITEM.SEQ_ID = PEDIDOITEMOBS.SEQ_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_ID_N
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into PEDIDOITEMOBS "
         SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID,OBS,Referencia,QTDE,Valor)"
      SQL = SQL & " values("
         SQL = SQL & PEDIDO_ID_N
         SQL = SQL & "," & SEQ_ID_N
         SQL = SQL & "," & PRODUTO_ID_N
         SQL = SQL & ",'" & Trim(txtOBS.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRef.Text) & "'"
         SQL = SQL & ",0"
         SQL = SQL & ",0"
      SQL = SQL & ")"
      Else
         SQL = "update PEDIDOITEMOBS set "
            SQL = SQL & " OBS = '" & Trim(txtOBS.Text) & "'"
            SQL = SQL & ",Referencia = '" & Trim(txtRef.Text) & "'"
            SQL = SQL & ",QTDE = 0"
            SQL = SQL & ",Valor = 0"
         SQL = SQL & " where PEDIDOITEMOBS.pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and PEDIDOITEMOBS.seq_id = " & SEQ_ID_N
   End If

   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close

Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_DADOS"
End Sub
