VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmInventarioConsulta 
   Caption         =   "Consulta Inventário"
   ClientHeight    =   6270
   ClientLeft      =   225
   ClientTop       =   555
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
   Icon            =   "InventarioConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   8760
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   8760
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.ComboBox cmbTipoMov 
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
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3000
      Picture         =   "InventarioConsulta.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   405
   End
   Begin VB.TextBox txtDescricao 
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
      Left            =   3480
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cmbSituacao 
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
      Left            =   6840
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtProduto 
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
      Height          =   360
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   1270
      ButtonWidth     =   3149
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
            Caption         =   "Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Lote"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7560
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
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
               Picture         =   "InventarioConsulta.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioConsulta.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioConsulta.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioConsulta.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioConsulta.frx":A8FD
               Key             =   ""
            EndProperty
         EndProperty
      End
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
      DesignHeight    =   6270
   End
   Begin MSComctlLib.ListView ListaInventario 
      Height          =   3945
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   6959
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   14737632
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "req"
         Text            =   "Lote"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Lote"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Seq."
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Produto"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "cli"
         Text            =   "Saldo Anterior"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "valor"
         Text            =   "1ºContagem"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   "desconto"
         Text            =   "2ºContagem"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   "total"
         Text            =   "QtdeAtual"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "status"
         Text            =   "Situação"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "TipoMov."
         Object.Width           =   2646
      EndProperty
   End
   Begin MSMask.MaskEdBox txtDtLote 
      Height          =   360
      Left            =   10320
      TabIndex        =   3
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   6000
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TipoMov.:"
      Height          =   240
      Left            =   2520
      TabIndex        =   15
      Top             =   960
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Lote:"
      Height          =   240
      Left            =   9480
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6240
      Y2              =   6240
   End
End
Attribute VB_Name = "frmInventarioConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   LIMPA_TUDO
   MONTA_FAMILIAPRODUTO

   cmbTipoMov.Clear
   cmbTipoMov.AddItem "Entrada"
   cmbTipoMov.AddItem "Saída"

   cmbSituacao.Clear
   cmbSituacao.AddItem "Aberta"
   cmbSituacao.AddItem "Fechada"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "print"
         MONTA_CONSULTA

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_inventario.rpt"
         frmRELATORIO10.Show 1
      Case "consultar"
         MONTA_CONSULTA
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub ListaInventario_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView ListaInventario, ColumnHeader
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Sub MONTA_CONSULTA()
'On Error GoTo ERRO_TRATA

   ListaInventario.ListItems.Clear
   ListaInventario.Visible = False

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select INVENTARIO.*, PRODUTO.DESCRICAO,produto.codg_produto "
   SQL = SQL & " from INVENTARIO "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON INVENTARIO.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where INVENTARIO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = "{INVENTARIO.estabelecimento_id} = " & ESTABELECIMENTO_ID_N

   If Trim(txtLOTE.Text) <> "" Then
      If IsNumeric(txtLOTE.Text) Then
         SQL = SQL & " and numr_lote = " & Trim(txtLOTE.Text)
         FORMULA_REL = FORMULA_REL & " and {INVENTARIO.numr_lote} = " & Trim(txtLOTE.Text)
      End If
   End If

   If Trim(cmbSituacao.Text) <> "" Then
      SQL = SQL & " and status = '" & Left(cmbSituacao.Text, 1) & "'"
      FORMULA_REL = FORMULA_REL & " and {INVENTARIO.STATUS} = '" & Left(cmbSituacao.Text, 1) & "'"
   End If

   txtDtLote.PromptInclude = False
      If IsDate(txtDtLote.Text) Then
         SQL = SQL & " and dt_lote = '" & DMA(txtDtLote.Text) & "'"
         FORMULA_REL = FORMULA_REL & " and {INVENTARIO.dt_lote} = " & DMA(txtDtLote.Text)
      End If
   txtDtLote.PromptInclude = True

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and INVENTARIO.produto_id = " & PRODUTO_ID_N
      FORMULA_REL = FORMULA_REL & " and {INVENTARIO.produto_id} = " & PRODUTO_ID_N
   End If

   If Trim(cmbTipoMov.Text) <> "" Then
      SQL = SQL & " and tipo_mov = '" & Left(cmbTipoMov.Text, 1) & "'"
      FORMULA_REL = FORMULA_REL & " and {INVENTARIO.tipo_mov} = '" & Trim(Left(cmbTipoMov.Text, 1)) & "'"
   End If

   SQL = SQL & " order by numr_lote, seq "

   NUMR_SEQ_N = 0

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If TabConsulta.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      Me.Enabled = True
      MsgBox "Não há dados para esta pesquisa."
      ListaInventario.Visible = True
      Exit Sub
   End If

   Me.Enabled = False
'====================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL3 = "select count(numr_lote) from INVENTARIO "
   SQL3 = SQL3 & " where INVENTARIO.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(txtLOTE.Text) <> "" Then _
      If IsNumeric(txtLOTE.Text) Then _
         SQL3 = SQL3 & " and numr_lote = " & Trim(txtLOTE.Text)

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL3 = SQL3 & " and status = '" & Left(cmbSituacao.Text, 1) & "'"

   txtDtLote.PromptInclude = False
      If IsDate(txtDtLote.Text) Then _
         SQL3 = SQL3 & " and dt_lote = '" & DMA(txtDtLote.Text) & "'"
   txtDtLote.PromptInclude = True

   If Trim(txtProduto.Text) <> "" Then _
      SQL3 = SQL3 & " and produto_id = " & PRODUTO_ID_N

   If Trim(cmbTipoMov.Text) <> "" Then _
      SQL3 = SQL3 & " and tipo_mov = '" & Left(cmbTipoMov.Text, 1) & "'"

   TabTemp.Open SQL3, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      CONTA_REG_PROGRESSO = TabTemp.Fields(0).Value

   If CONTA_REG_PROGRESSO > 0 Then
      ProgressBar1.Min = 0                   'Indica o valor inicial
      ProgressBar1.Max = CONTA_REG_PROGRESSO 'Indica o valor final
   End If
'====================
   While Not TabConsulta.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      DoEvents

      If CONT_N < CONTA_REG_PROGRESSO Then
         CONT_N = CONT_N + 1
         ProgressBar1.Value = CONT_N
      End If

      Set item = ListaInventario.ListItems.Add(, "seq." & NUMR_SEQ_N & TabConsulta.Fields("numr_lote").Value, TabConsulta.Fields("numr_lote").Value)

      item.SubItems(1) = "" & TabConsulta.Fields("DT_LOTE").Value
      item.SubItems(2) = "" & TabConsulta.Fields("seq").Value
      item.SubItems(3) = "" & Trim(TabConsulta.Fields("CODG_PRODUTO").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
      item.SubItems(4) = "" & Format(TabConsulta.Fields("QTD_ANTERIOR").Value, strFormatacao3Digitos)
      item.SubItems(5) = "" & Format(TabConsulta.Fields("QTD_PRIMEIRA").Value, strFormatacao3Digitos)
      item.SubItems(6) = "" & Format(TabConsulta.Fields("QTD_SEGUNDA").Value, strFormatacao3Digitos)
      item.SubItems(7) = "" & Format(TabConsulta.Fields("QTD_ATUAL").Value, strFormatacao3Digitos)
      item.SubItems(8) = "" & TabConsulta.Fields("STATUS").Value
      item.SubItems(9) = "" & TabConsulta.Fields("tipo_mov").Value

      If Trim(TabConsulta.Fields("STATUS").Value) <> "" Then
         If Trim(TabConsulta.Fields("STATUS").Value) = "F" Then _
            item.SubItems(8) = "Atualizado"
         If Trim(TabConsulta.Fields("STATUS").Value) = "A" Then _
            item.SubItems(8) = "Aberto"
         If Trim(TabConsulta.Fields("STATUS").Value) = "C" Then _
            item.SubItems(8) = "Cancelado"
      End If

      If Trim(TabConsulta.Fields("tipo_mov").Value) <> "" Then
         If Trim(TabConsulta.Fields("tipo_mov").Value) = "E" Then _
            item.SubItems(9) = "Entrada"
         If Trim(TabConsulta.Fields("tipo_mov").Value) = "S" Then _
            item.SubItems(9) = "Saída"
         If Trim(TabConsulta.Fields("tipo_mov").Value) = "C" Then _
            item.SubItems(9) = "Cancelado"
      End If

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Me.Enabled = True
   ListaInventario.Visible = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   ListaInventario.Visible = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   ListaInventario.ListItems.Clear

   txtLOTE.Text = ""
   cmbSituacao.Text = ""
   txtDtLote.PromptInclude = False
      txtDtLote.Mask = "##/##/####"
      txtDtLote.Text = ""
   txtDtLote.PromptInclude = True
   txtProduto.Text = ""
   txtDescricao.Text = ""
   cmbTipoMov.Text = ""
   cmbFamilia.Text = ""
   cmbFamiliaAUX.Text = ""


Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub MONTA_FAMILIAPRODUTO()
'On Error GoTo ERRO_TRATA

   cmbFamilia.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from FAMILIAPRODUTO "
   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbFamilia.AddItem Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("familiaproduto_id").Value)

      TabTemp.MoveNext
   Wend

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_FAMILIAPRODUTO"
End Sub

Sub PROCURA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select produto_id,descricao from PRODUTO "
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDescricao.Text = TabConsulta.Fields("descricao").Value
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
End Sub
