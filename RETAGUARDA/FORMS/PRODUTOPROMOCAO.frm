VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPRODUTOPROMOCAO 
   Caption         =   "Cadastro Promoção Produtos"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PRODUTOPROMOCAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mensagem Promocional"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   10455
      Begin VB.TextBox txtMSGPromocao 
         Height          =   960
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "PRODUTOPROMOCAO.frx":5C12
         Top             =   360
         Width           =   10215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Programação dia e horário de exibição"
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   10455
      Begin VB.TextBox txtMSGDuracao 
         Height          =   840
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "PRODUTOPROMOCAO.frx":5C3B
         Top             =   840
         Width           =   8415
      End
      Begin VB.TextBox txtDuracao 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   8520
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Height          =   360
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MSG Display:"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblMinuto 
         AutoSize        =   -1  'True
         Caption         =   "Minuto(s)"
         Height          =   240
         Left            =   9360
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Duração:"
         Height          =   240
         Left            =   7560
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Hora Final:"
         Height          =   240
         Left            =   3930
         TabIndex        =   12
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Hora Inicial:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3525
      Picture         =   "PRODUTOPROMOCAO.frx":5C5D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   405
   End
   Begin VB.TextBox txtDesc 
      Enabled         =   0   'False
      Height          =   360
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   8
      ToolTipText     =   "Informe "
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Informe o código do produto."
      Top             =   1320
      Width           =   3375
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOPROMOCAO.frx":665F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOPROMOCAO.frx":77F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOPROMOCAO.frx":8888
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUTOPROMOCAO.frx":9993
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstProgramacao 
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   2778
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Produto"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DtIni"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DtFim"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Duração"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MSG Display"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "MSG Promocional"
         Object.Width           =   8819
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
      DesignWidth     =   10665
      DesignHeight    =   6780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Produtos Promoção"
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
      TabIndex        =   6
      Top             =   720
      Width           =   10620
   End
End
Attribute VB_Name = "frmPRODUTOPROMOCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim PRODUTOPROMOCAO_ID_N   As Long

Private Sub Form_Load()
   CHECA_TABELA
   LIMPA_TELA
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TELA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstProgramacao_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstProgramacao.SelectedItem.Text) Then
      If Trim(lstProgramacao.SelectedItem.Text) <> "" Then
         TRAZ_DADOS_PROMOCAO lstProgramacao.SelectedItem.ListSubItems.item(1).Text, lstProgramacao.SelectedItem.Text

         If Not IsNull(lstProgramacao.SelectedItem.Text) Then _
            PRODUTOPROMOCAO_ID_N = 0 & lstProgramacao.SelectedItem.Text
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProgramacao_DblClick"
End Sub

Private Sub lstProgramacao_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(lstProgramacao.SelectedItem.Text) Then
            If Trim(lstProgramacao.SelectedItem.Text) <> "" Then
               SQL = "delete PRODUTOPROMOCAO "
               SQL = SQL & " where PRODUTOPROMOCAO_id = " & lstProgramacao.SelectedItem.Text
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
      
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProgramacao_KeyDown"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      Call txtProduto_KeyPress(13)
      txtProduto.SetFocus
   End If
   SQL3 = ""
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub txtDesc_GotFocus()
   txtDtIni.SetFocus
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtini_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni)
   txtDtIni.BackColor = &HC0FFFF
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Now
End Sub

Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub TXTDTFIM_GotFocus()
   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim)
   txtDtFim.BackColor = &HC0FFFF
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Now
End Sub

Private Sub txtDtFim_LostFocus()
   txtDtFim.BackColor = &HFFFFFF
   TRAZ_DADOS_PROMOCAO PRODUTO_ID_N, 0
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then _
         txtDuracao.Text = "" & CONVERTE_TEMPO(txtDtIni.Text, txtDtFim.Text)

      txtMSGDuracao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtfim_KeyPress"
End Sub

Private Sub txtMSGDuracao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtMSGPromocao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMSGDuracao_KeyPress"
End Sub

Private Sub txtMSGPromocao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_DADOS
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMSGPromocao_KeyPress"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPRODUTO_GotFocus"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   If Trim(txtProduto.Text) <> "" Then
      txtDesc.Text = "" & TRAZ_DESCRICAO_PRODUTO(0, Trim(txtProduto.Text))
      PRODUTO_ID_N = 0 & TRAZ_ID_TABELA("PRODUTO", "PRODUTO_ID", "CODG_PRODUTO", Trim(txtProduto.Text))
   End If
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPRODUTO_KeyPress"
End Sub

Sub LIMPA_TELA()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0
   PRODUTOPROMOCAO_ID_N = 0
   txtProduto.Text = ""
   txtDesc.Text = ""
   txtDtIni.Text = ""
   txtDtFim.Text = ""
   txtDuracao.Text = ""
   txtMSGDuracao.Text = ""
   txtMSGPromocao.Text = ""
   lstProgramacao.ListItems.Clear
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TELA"
End Sub

Sub CHECA_TABELA()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUTOPROMOCAO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[PRODUTOPROMOCAO]("
      SQL = SQL & " [PRODUTOPROMOCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DATAINI] [datetime] NULL,"
      SQL = SQL & " [DATAFIM] [datetime] NULL,"
      SQL = SQL & " [MSGDISPLAY] [nvarchar](max) NULL,"
      SQL = SQL & " [MSGPROMOCAO] [nvarchar](max) NULL,"
      SQL = SQL & " CONSTRAINT [PK_PRODUTOPROMOCAO] PRIMARY KEY CLUSTERED([PRODUTOPROMOCAO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) "
      SQL = SQL & " ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[PRODUTOPROMOCAO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTOPROMOCAO_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[PRODUTOPROMOCAO] CHECK CONSTRAINT [FK_PRODUTOPROMOCAO_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_TABELA"
End Sub

Sub TRAZ_DADOS_PROMOCAO(PROD_ID_N As Long, PROD_PROMO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PRODUTOPROMOCAO.*, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " PRODUTO.UNIDADE_MEDIDA, Produto.Tipo_Prod"
   SQL = SQL & " from PRODUTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTOPROMOCAO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOPROMOCAO.PRODUTO_ID"

   SQL = SQL & " where PRODUTOPROMOCAO.produto_id = " & PROD_ID_N

   If PROD_PROMO_ID_N > 0 Then _
      SQL = SQL & " and PRODUTOPROMOCAO_Id = " & PROD_PROMO_ID_N

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " where produto_id >= '" & DMA(txtDtIni.Text) & "'"
      SQL = SQL & " where produto_id <= '" & DMA(txtDtFim.Text) & "'"
   End If

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtProduto.Text = "" & Trim(TabTemp.Fields("codg_produto").Value)
      Call TXTPRODUTO_LostFocus
      txtDtIni.Text = "" & Trim(TabTemp.Fields("dAtAini").Value)
      txtDtFim.Text = "" & Trim(TabTemp.Fields("dAtAFIM").Value)
      txtDuracao.Text = "" & TabTemp.Fields("dAtAFIM").Value - TabTemp.Fields("dAtAini").Value
      txtMSGDuracao.Text = "" & Trim(TabTemp.Fields("MSGDISPLAY").Value)
      txtMSGPromocao.Text = "" & Trim(TabTemp.Fields("MSGPROMOCAO").Value)
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRAZ_DADOS_PROMOCAO"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstProgramacao.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PRODUTOPROMOCAO.*, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " PRODUTO.UNIDADE_MEDIDA, Produto.Tipo_Prod"
   SQL = SQL & " from PRODUTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTOPROMOCAO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOPROMOCAO.PRODUTO_ID"

   SQL = SQL & " where PRODUTOPROMOCAO_ID > 0 "

   'SQL = SQL = " and dataini >= '" & DMA(Date) & "'"

   If PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and PRODUTOPROMOCAO.produto_id = " & PRODUTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstProgramacao.ListItems.Add(, "seq." & TabTemp.Fields("PRODUTOPROMOCAO_id").Value, TabTemp.Fields("PRODUTOPROMOCAO_id").Value)
      item.SubItems(1) = "" & Trim(TabTemp.Fields("codg_produto").Value)
      item.SubItems(2) = "" & Trim(TabTemp.Fields("descricao").Value)
      item.SubItems(3) = "" & Trim(TabTemp.Fields("DATAINI").Value)
      item.SubItems(4) = "" & Trim(TabTemp.Fields("dAtAfim").Value)
      item.SubItems(5) = ""
      If Not IsNull(TabTemp.Fields("dAtAfim").Value) And Not IsNull(TabTemp.Fields("dAtAini").Value) Then _
         item.SubItems(5) = "" & CONVERTE_TEMPO(TabTemp.Fields("dAtAini").Value, TabTemp.Fields("dAtAfim").Value)
      item.SubItems(6) = "" & Trim(TabTemp.Fields("msgdisplay").Value)
      item.SubItems(7) = "" & Trim(TabTemp.Fields("msgpromocao").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub GRAVA_DADOS()
'On Error GoTo ERRO_TRATA

   If PRODUTO_ID_N <= 0 Then _
      Exit Sub

   Dim DATA_INICIO_D As String
   Dim DATA_FIM_D    As String
   
   DATA_INICIO_D = "" & Format(txtDtIni.Text, "##/##/#### ##:##:##")
   DATA_FIM_D = "" & Format(txtDtFim.Text, "##/##/#### ##:##:##")

   If Trim(DATA_INICIO_D) = "" Then _
      DATA_INICIO_D = "NULL"
   If Trim(DATA_FIM_D) = "" Then _
      DATA_FIM_D = "NULL"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PRODUTOPROMOCAO_ID from PRODUTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTOPROMOCAO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOPROMOCAO.PRODUTO_ID"

   SQL = SQL & " where PRODUTOPROMOCAO.produto_id = " & PRODUTO_ID_N

   If PRODUTOPROMOCAO_ID_N > 0 Then _
      SQL = SQL & " and PRODUTOPROMOCAO_Id = " & PRODUTOPROMOCAO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update PRODUTOPROMOCAO set "
         SQL = SQL & " PRODUTO_ID = " & PRODUTO_ID_N
         SQL = SQL & ",DATAINI = '" & DATA_INICIO_D & "'"
         SQL = SQL & ",DATAFIM = '" & DATA_FIM_D & "'"
         SQL = SQL & ",MSGDISPLAY = '" & Trim(txtMSGDuracao.Text) & "'"
         SQL = SQL & ",MSGPROMOCAO = '" & Trim(txtMSGPromocao.Text) & "'"
      SQL = SQL & " and PRODUTOPROMOCAO_Id = " & PRODUTOPROMOCAO_ID_N
      Else
         SQL = "insert into PRODUTOPROMOCAO"
            SQL = SQL & "(PRODUTOPROMOCAO_ID,PRODUTO_ID,DATAINI,DATAFIM,MSGDISPLAY,MSGPROMOCAO)"
         SQL = SQL & " values("
            SQL = SQL & MAX_ID("PRODUTOPROMOCAO_ID", "PRODUTOPROMOCAO", "", "", "", "")
            SQL = SQL & "," & PRODUTO_ID_N

            If DATA_INICIO_D <> "NULL" Then
               SQL = SQL & ",'" & Format(txtDtIni.Text, "##/##/#### ##:##:##") & "'"
               Else: SQL = SQL & "," & DATA_INICIO_D '& "'"
            End If
            If DATA_FIM_D <> "NULL" Then
               SQL = SQL & ",'" & Format(txtDtFim.Text, "##/##/#### ##:##:##") & "'"
               Else: SQL = SQL & "," & DATA_FIM_D              '& "'"
            End If

            SQL = SQL & ",'" & Trim(txtMSGDuracao.Text) & "'"
            SQL = SQL & ",'" & Trim(txtMSGPromocao.Text) & "'"
         SQL = SQL & ")"
   End If

   CONECTA_RETAGUARDA.Execute SQL
   LIMPA_TELA
   SETA_GRID
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_DADOS"
End Sub

Function CONVERTE_TEMPO(DATA_INI_D As String, DATA_FIM_D As String) As String
   Dim d1, d2, d3, Segundo_n, d5  As Single
   Dim Minuto_N            As Long
   Dim Hora_N              As Long
   Dim Dia_N               As Long
   Dim Data_ini_Vaca       As Date
   Dim Data_fim_Vaca       As Date

   If IsDate(DATA_INI_D) Then
      Data_ini_Vaca = DATA_INI_D
      Else: Data_ini_Vaca = Format(DATA_INI_D, "##/##/#### ##:##:##")
   End If
   If IsDate(DATA_FIM_D) Then
      Data_fim_Vaca = DATA_FIM_D
      Else: Data_fim_Vaca = Format(DATA_FIM_D, "##/##/#### ##:##:##")
   End If

   d1 = DateDiff("d", Data_ini_Vaca, Data_fim_Vaca)
   d2 = DateDiff("m", Data_ini_Vaca, Data_fim_Vaca)
   d3 = DateDiff("yyyy", Data_ini_Vaca, Data_fim_Vaca)
   Segundo_n = DateDiff("s", Data_ini_Vaca, Data_fim_Vaca)

   Msg = " Sua idade e : " & vbCrLf
   Msg = Msg & " ============================== " & vbCrLf
   Msg = Msg & " Em dias : " & d1 & " dias " & vbCrLf
   Msg = Msg & " Em meses : " & d2 & " meses " & vbCrLf
   Msg = Msg & " Em anos : " & d3 & " anos " & vbCrLf
   Msg = Msg & " Em segundos : " & Segundo_n & " segundos " & vbCrLf

   CONVERTE_TEMPO = ""

'converção para minutos
   Minuto_N = 0
   While Segundo_n >= 60
      Segundo_n = Segundo_n - 60
      Minuto_N = Minuto_N + 1
   Wend
   
'converção para horas
   Hora_N = 0
   While Minuto_N >= 60
      Minuto_N = Minuto_N - 60
      Hora_N = Hora_N + 1
   Wend

'converção para dia
   Dia_N = 0
   While Hora_N >= 60
      Hora_N = Hora_N - 60
      Dia_N = Dia_N + 1
   Wend

   'txtDuracao.Text = Minuto_N & ":" & Segundo_n
   CONVERTE_TEMPO = "" & Minuto_N & ":" & Segundo_n
   
'MsgBox "Dia = " & Dia_n & " ; Horas = " & Hora_n & " ; Minutos = " & Minuto_n & " ; Segundos = " & Segundo_n

End Function
