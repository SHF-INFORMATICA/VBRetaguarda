VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProducaoRegistroPerdaConsulta 
   Caption         =   "Consulta Perda Produto Produção"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INDControlePerdaConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optAna 
      Caption         =   "Analítico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9360
      TabIndex        =   26
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optSint 
      Caption         =   "Sintético"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8160
      TabIndex        =   25
      Top             =   840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      Height          =   240
      Left            =   9360
      TabIndex        =   24
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtUsuCodg 
      BackColor       =   &H80000003&
      Height          =   360
      Left            =   7080
      MaxLength       =   50
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbFamiliaAUX 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7800
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   7800
      TabIndex        =   9
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CheckBox chkC 
      Caption         =   "&Cancelados"
      Height          =   240
      Left            =   9360
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox cmbEstabAUX 
      BackColor       =   &H80000003&
      Height          =   360
      ItemData        =   "INDControlePerdaConsulta.frx":5C12
      Left            =   6960
      List            =   "INDControlePerdaConsulta.frx":5C1C
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtProduto 
      Height          =   360
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdProd 
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
      Left            =   960
      Picture         =   "INDControlePerdaConsulta.frx":5C31
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Consulta Produto"
      Top             =   2280
      Width           =   405
   End
   Begin VB.TextBox txtDescricao 
      Enabled         =   0   'False
      Height          =   360
      Left            =   3720
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2280
      Width           =   3975
   End
   Begin VB.ComboBox cmbEstab 
      Height          =   360
      ItemData        =   "INDControlePerdaConsulta.frx":6633
      Left            =   6960
      List            =   "INDControlePerdaConsulta.frx":663D
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdUsu 
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
      Left            =   6480
      Picture         =   "INDControlePerdaConsulta.frx":6652
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Consulta Produto"
      Top             =   1680
      Width           =   405
   End
   Begin VB.ComboBox cmbSt 
      Height          =   360
      ItemData        =   "INDControlePerdaConsulta.frx":7054
      Left            =   5040
      List            =   "INDControlePerdaConsulta.frx":705E
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lote"
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   4815
      Begin VB.OptionButton optFechamento 
         Caption         =   "Dt.Fecham."
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optAbertura 
         Caption         =   "Dt.Abertura"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Final:"
         Height          =   240
         Left            =   300
         TabIndex        =   15
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Inicial:"
         Height          =   240
         Left            =   195
         TabIndex        =   14
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.TextBox txtUsu 
      Height          =   360
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1680
      Width           =   3975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
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
         NumButtons      =   7
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
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7080
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
               Picture         =   "INDControlePerdaConsulta.frx":7073
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControlePerdaConsulta.frx":820D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControlePerdaConsulta.frx":929C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControlePerdaConsulta.frx":A3A7
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstPerda 
      Height          =   4065
      Left            =   30
      TabIndex        =   11
      Top             =   2760
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   7170
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lote"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Seq"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CodgProd"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descrição"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qtde"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ST"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Familia"
         Object.Width           =   5292
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
      DesignWidth     =   10950
      DesignHeight    =   6840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Situação"
      Height          =   240
      Index           =   1
      Left            =   5040
      TabIndex        =   27
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Família de Produto"
      Height          =   240
      Left            =   7800
      TabIndex        =   21
      Top             =   2040
      Width           =   1830
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label lblNome 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Responsável:"
      Height          =   240
      Left            =   5160
      TabIndex        =   12
      Top             =   1680
      Width           =   1260
   End
End
Attribute VB_Name = "frmProducaoRegistroPerdaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CARREGA_ESTAB
   CARREGA_FAMILIA
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
      Case "limpar"
         LIMPA_CONSULTA
      Case "consultar"
         MONTA_CONSULTA
      Case "print"
         GERA_IMPRESSAO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstPerda_DblClick()
On Error Resume Next

   CRITERIO_A = ""
   If Not IsNull(lstPerda.SelectedItem.Text) Then
      If IsNumeric(lstPerda.SelectedItem.Text) Then
         CRITERIO_A = lstPerda.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub cmdProd_Click()
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         'If Trim(txtProduto.Text) <> "" Then _
            If Trim(txtLote.Text) <> "" Then _
               If Trim(txtSeq.Text) <> "" Then _
                  MATA_ITEM txtSeq.Text
      Case vbKeyF7
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtProduto.ForeColor = vbBlue
   txtDescricao.ForeColor = vbBlue

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCESSA_DADOS_PRODUTOS
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub cmbestab_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbestab_Click"
End Sub

Sub LIMPA_CONSULTA()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.Text = ""
   optAbertura.Value = True
   cmbSt.Text = ""
   cmbEstab.Text = ""
   txtUsu.Text = ""
   txtUsuCodg.Text = ""
   txtProduto.Text = ""
   txtDescricao.Text = ""
   lstPerda.ListItems.Clear
   cmbFamiliaAUX.Text = ""
   cmbFamilia.Text = ""
   PRODUTO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_CONSULTA"
End Sub

Sub CARREGA_ESTAB()
'On Error GoTo ERRO_TRATA

   cmbEstab.Clear
   cmbEstabAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ESTABELECIMENTO "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbEstab.AddItem Trim(TabTemp.Fields("descricao").Value)
      cmbEstabAUX.AddItem TabTemp.Fields("estabelecimento_id").Value

      If ESTABELECIMENTO_ID_N = TabTemp.Fields("estabelecimento_id").Value Then _
         cmbEstab.Text = Trim(TabTemp.Fields("descricao").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_ESTAB"
End Sub

Sub CARREGA_FAMILIA()
'On Error GoTo ERRO_TRATA

   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FAMILIAPRODUTO "
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFamilia.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("familiaproduto_id").Value)
      cmbFamiliaAUX.AddItem Trim(TabDESCR.Fields("familiaproduto_id").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_FAMILIA"
End Sub

Sub MONTA_CONSULTA()
'On Error GoTo ERRO_TRATA

   Dim TabRegistro   As New ADODB.Recordset

   CODG_PRODUTO_A = ""
   NUMR_ID_N = 0
   CONT_N = 0
   lstPerda.ListItems.Clear

   If TabRegistro.State = 1 Then _
      TabRegistro.Close

   SQL = "select CONTROLEPERDA.*, CONTROLEPERDAITEM.SEQ_ID, CONTROLEPERDAITEM.PRODUTO_ID, "
   SQL = SQL & " CONTROLEPERDAITEM.QTDE, CONTROLEPERDAITEM.valor, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " PRODUTO.SITUACAO, PRODUTO.TIPO_PROD,"
   SQL = SQL & " FAMILIAPRODUTO.PRODUCAO, FAMILIAPRODUTO.CODG_FAMILIA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia "
   SQL = SQL & " from CONTROLEPERDAITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON CONTROLEPERDAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " INNER JOIN CONTROLEPERDA WITH (NOLOCK)"
   SQL = SQL & " ON CONTROLEPERDAITEM.CONTROLEPERDA_ID = CONTROLEPERDA.CONTROLEPERDA_ID"

   SQL = SQL & " where estabelecimento_id = " & cmbEstabAUX.Text

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True

      If optAbertura.Value = True Then
         SQL = SQL & " and dt_registro >= '" & txtDtIni.Text & "'"
         SQL = SQL & " and dt_registro <= '" & txtDtFim.Text & "'"
      End If
      If optFechamento.Value = True Then
         SQL = SQL & " and dt_fecha >= '" & txtDtIni.Text & "'"
         SQL = SQL & " and dt_fecha <= '" & txtDtFim.Text & "'"
      End If
   End If

   If Trim(cmbSt.Text) <> "" Then _
      SQL = SQL & " and status = '" & Left(cmbSt.Text, 1) & "'"
   
   If Trim(txtUsuCodg.Text) <> "" Then _
      SQL = SQL & " and usuario_id = " & txtUsuCodg.Text

   If PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and CONTROLEPERDAITEM.PRODUTO_ID = " & PRODUTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      SQL = SQL & " and familiaproduto_ID = " & cmbFamiliaAUX.Text

   SQL = SQL & " order by controleperda.controleperda_id, CONTROLEPERDAITEM.PRODUTO_ID"

   TabRegistro.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabRegistro.EOF Then
      If TabRegistro.State = 1 Then _
         TabRegistro.Close
      MsgBox "Nenhum Registro Encotrado."
      Else
         While Not TabRegistro.EOF
            CONT_N = CONT_N + 1
            Set item = lstPerda.ListItems.Add(, "A" & CONT_N, TabRegistro.Fields("controleperda_id").Value)

            item.SubItems(1) = "" & TabRegistro.Fields("seq_id").Value
            item.SubItems(2) = "" & Trim(TabRegistro.Fields("codg_produto").Value)
            item.SubItems(3) = "" & Trim(TabRegistro.Fields("descricao").Value)
            item.SubItems(4) = "" & Format(TabRegistro.Fields("qtde").Value, strFormatacao3Digitos)
            item.SubItems(5) = "" & Format(TabRegistro.Fields("valor").Value, strFormatacao3Digitos)
            item.SubItems(6) = "" & Trim(TabRegistro.Fields("status").Value)
            item.SubItems(7) = "" & Trim(TabRegistro.Fields("descfamilia").Value)

            TabRegistro.MoveNext
         Wend
   End If
   If TabRegistro.State = 1 Then _
      TabRegistro.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub

Sub GERA_IMPRESSAO()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{controleperda.estabelecimento_id} = " & ESTABELECIMENTO_ID_N

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True

      DATA_INI = txtDtIni.Text
      DATA_FIM = txtDtFim.Text
      If optAbertura.Value = True Then
         FORMULA_REL = FORMULA_REL & " and {controleperda.dt_registro} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
         FORMULA_REL = FORMULA_REL & " and {controleperda.dt_registro} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"
      End If
      If optFechamento.Value = True Then
         FORMULA_REL = FORMULA_REL & " and {controleperda.dt_fecha} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
         FORMULA_REL = FORMULA_REL & " and {controleperda.dt_fecha} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"
      End If
   End If

   If Trim(cmbSt.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {controleperda.status} = '" & Left(cmbSt.Text, 1) & "'"

   If Trim(txtUsuCodg.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {controleperda.usuario_id} = " & txtUsuCodg.Text

   If PRODUTO_ID_N > 0 Then _
      FORMULA_REL = FORMULA_REL & " and {controleperdaitem.produto_id} = " & PRODUTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {controleperda.familiaproduto_ID} = " & cmbFamiliaAUX.Text

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If optSint.Value = True Then
      Nome_Relatorio = "REL_PERDA_SINT.rpt"
      Else: Nome_Relatorio = "REL_PERDA_ANALIT.rpt"
   End If

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_IMPRESSAO"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

   If TabProduto.State = 1 Then _
      TabProduto.Close

   'txtQtde.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtProduto.Text = Trim(CODG_PRODUTO_A)
   txtDescricao.Text = DESC_PRODUTO_A
   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
