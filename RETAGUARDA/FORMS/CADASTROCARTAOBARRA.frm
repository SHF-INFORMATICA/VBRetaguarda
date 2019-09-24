VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROCARTAOBARRA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cartão Código de Barras"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROCARTAOBARRA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStatus 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2025
      TabIndex        =   8
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtDescricao 
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox txtCodg 
      Height          =   405
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   6255
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Description     =   "Voltar para Tela Início"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   6
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
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Cadastro"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":9331
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":A2E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":B9E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCARTAOBARRA.frx":CB7D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstBarras 
      Height          =   2415
      Left            =   50
      TabIndex        =   2
      Top             =   2520
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   4260
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DtCad"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Situação: "
      Height          =   285
      Left            =   810
      TabIndex        =   7
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label lblDtCad 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   8115
      TabIndex        =   6
      Top             =   1440
      Width           =   60
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   285
      Left            =   690
      TabIndex        =   5
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código Barras:"
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   1755
   End
End
Attribute VB_Name = "frmCADASTROCARTAOBARRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   'ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
   CHECA_TABELA
   SETA_GRID
   LIMPA_BARRA

   cmbSTATUS.Clear
   cmbSTATUS.AddItem "Ativo"
   cmbSTATUS.AddItem "Desativado"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
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
      Case "voltar"
         Unload Me
      Case "matar"
         MATA_BARRA Trim(txtCodg.Text)
         txtCodg.SetFocus
      Case "limpar"
         LIMPA_BARRA
         txtCodg.SetFocus
      Case "gravar"
         GRAVA_BARRA
         SETA_GRID
         txtCodg.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCodg_LostFocus()
   MOSTRA_BARRA
End Sub

Private Sub txtcodg_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescricao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodg_KeyPress"
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSTATUS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodg_KeyPress"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCodg.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodg_KeyPress"
End Sub

Private Sub cmbStatus_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodg.Text) = "" Then
      txtCodg.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Trim(cmbSTATUS.Text) = "" Then
      cmbSTATUS.SetFocus
      Exit Sub
   End If

   GRAVA_BARRA
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub lstBarras_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstBarras, ColumnHeader
End Sub

Private Sub lstBarras_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(lstBarras.SelectedItem.Text) Then _
            If Trim(lstBarras.SelectedItem.Text) <> "" Then _
               MATA_BARRA Trim(lstBarras.SelectedItem.Text)

      Case vbKeyF6
         If Not IsNull(lstBarras.SelectedItem.Text) Then _
            If Trim(lstBarras.SelectedItem.Text) <> "" Then _
               MATA_BARRA Trim(lstBarras.SelectedItem.Text)
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstBarras_KeyDown"
End Sub

Sub LIMPA_BARRA()
'On Error GoTo ERRO_TRATA

   txtCodg.Text = ""
   txtDescricao.Text = ""
   lblDtCad.Caption = ""
   cmbSTATUS.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BARRA"
End Sub

Sub MOSTRA_BARRA()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodg.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CARTAOBARRA WITH (NOLOCK)"
      SQL = SQL & " where codigo_barra = '" & Trim(txtCodg.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         txtCodg.Text = "" & Trim(TabTemp.Fields("CODIGO_BARRA").Value)
         txtDescricao.Text = "" & Trim(TabTemp.Fields("DESCRICAO").Value)
         lblDtCad.Caption = "" & Trim(TabTemp.Fields("DTCAD").Value)
         cmbSTATUS.Text = "" & Trim(TabTemp.Fields("status").Value)
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_BARRA"
End Sub

Sub MATA_BARRA(CODG_BARRA_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(CODG_BARRA_A) = "" Then
      MsgBox "Código de Barras inválido."
      txtCodg.SetFocus
      Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CARTAOBARRA WITH (NOLOCK)"
   SQL = SQL & " where codigo_barra = '" & Trim(CODG_BARRA_A) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      Msg = "Confirma exclusão do registro selecionado ? " & "  " & CODG_BARRA_A
      PERGUNTA Msg, vbYesNo + 32, "Desconto Pedido Venda", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "delete from CARTAOBARRA "
         SQL = SQL & " where codigo_barra = '" & Trim(CODG_BARRA_A) & "'"
         CONECTA_RETAGUARDA.Execute SQL

         SETA_GRID
         txtCodg.SetFocus
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_BARRA"
End Sub

Sub GRAVA_BARRA()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodg.Text) = "" Then
      MsgBox "Código de Barras inválido."
      txtCodg.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Descrição inválida."
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Trim(cmbSTATUS.Text) = "" Then
      MsgBox "Selecione a situação."
      cmbSTATUS.SetFocus
      Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CARTAOBARRA WITH (NOLOCK)"
   SQL = SQL & " where codigo_barra = '" & Trim(txtCodg.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update CARTAOBARRA set "
         SQL = SQL & " descricao = '" & Trim(txtDescricao.Text) & "'"
         SQL = SQL & ", status = '" & Trim(Left(cmbSTATUS.Text, 1)) & "'"
      SQL = SQL & " where codigo_barra = '" & Trim(txtCodg.Text) & "'"
      Else
         SQL = "insert into CARTAOBARRA "
         SQL = SQL & "(CARTAOBARRA_ID,ESTABELECIMENTO_ID,CODIGO_BARRA,DESCRICAO,DTCAD,status)"
         SQL = SQL & " VALUES ("
            SQL = SQL & MAX_ID("CARTAOBARRA_ID", "CARTAOBARRA", "", "", "", "")  'CARTAOBARRA_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                               'ESTABELECIMENTO_ID
            SQL = SQL & ",'" & Trim(txtCodg.Text) & "'"                          'CODIGO_BARRA
            SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"                     'DESCRICAO
            SQL = SQL & ",'" & Now & "'"                                   'DTCAD
            SQL = SQL & ",'" & Trim(Left(cmbSTATUS.Text, 1)) & "'"               'status
         SQL = SQL & ")"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   LIMPA_BARRA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_BARRA"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstBarras.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CARTAOBARRA WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by cartaobarra_id desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstBarras.ListItems.Add(, "seq." & TabTemp.Fields("CARTAOBARRA_id").Value, Trim(TabTemp.Fields("codigo_barra").Value))
      item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
      item.SubItems(2) = "" & Trim(TabTemp.Fields("dtcad").Value)
      item.SubItems(3) = "" & Trim(TabTemp.Fields("status").Value)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub CHECA_TABELA()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CARTAOBARRA", "") = False Then
      SQL = "CREATE TABLE [dbo].[CARTAOBARRA]("
      SQL = SQL & " [CARTAOBARRA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [CODIGO_BARRA] [nvarchar](50) NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](50) NOT NULL,"
      SQL = SQL & " [DTCAD] [datetime] NOT NULL,"
      SQL = SQL & " [STATUS] [char] (1) ,"
      SQL = SQL & " CONSTRAINT [PK_CARTAOBARRA] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([CARTAOBARRA_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, "
      SQL = SQL & " ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[CARTAOBARRA]  WITH CHECK ADD CONSTRAINT [FK_CARTAOBARRA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[CARTAOBARRA] CHECK CONSTRAINT [FK_CARTAOBARRA_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOTEMP", "") = False Then
      SQL = "CREATE TABLE [dbo].[PEDIDOTEMP]("
      SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [CARTAOBARRA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_PEDIDO] [date] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_PEDIDOTEMP] PRIMARY KEY CLUSTERED("
      SQL = SQL & " [PEDIDO_ID] ASC, [ESTABELECIMENTO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOTEMP]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOTEMP_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOTEMP]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOTEMP_CARTAOBARRA] FOREIGN KEY([CARTAOBARRA_ID])"
      SQL = SQL & " References [dbo].[CARTAOBARRA]([CARTAOBARRA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOTEMP] CHECK CONSTRAINT [FK_PEDIDOTEMP_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      'SQL = " ALTER TABLE [dbo].[PEDIDOTEMP]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOTEMP_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
      'SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
      'CONECTA_RETAGUARDA.Execute SQL

      'SQL = " ALTER TABLE [dbo].[PEDIDOTEMP] CHECK CONSTRAINT [FK_PEDIDOTEMP_PEDIDO]"
      'CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOTEMP]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOTEMP_USUARIO] FOREIGN KEY([USUARIO_ID])"
      SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOTEMP] CHECK CONSTRAINT [FK_PEDIDOTEMP_USUARIO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_TABELA"
End Sub
