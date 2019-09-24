VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoqueTransfConsulta 
   Caption         =   "Consulta Transferências Estoque"
   ClientHeight    =   6000
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EstoqueTransfConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbEstabDestinoAUX 
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
      Left            =   4560
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbEstabDestino 
      Height          =   360
      Left            =   4560
      TabIndex        =   20
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   16
      ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3105
      Picture         =   "EstoqueTransfConsulta.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   405
   End
   Begin VB.ComboBox cmbSituacao 
      Height          =   360
      Left            =   7800
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox cmbEstabOrigemAUX 
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
      Left            =   4560
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbEstabOrigem 
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtID 
      Height          =   405
      Left            =   7800
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
      Height          =   405
      Left            =   3600
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1270
      ButtonWidth     =   2593
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
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin Threed.SSCheck chkOrdem 
         Height          =   270
         Left            =   6600
         TabIndex        =   17
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   476
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ordem Decresente"
         Value           =   1
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
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
         Left            =   8640
         TabIndex        =   11
         Top             =   280
         Width           =   1455
      End
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
               Picture         =   "EstoqueTransfConsulta.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueTransfConsulta.frx":7A3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueTransfConsulta.frx":8ACB
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueTransfConsulta.frx":9C65
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueTransfConsulta.frx":BADC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueTransfConsulta.frx":D7E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10560
      DesignHeight    =   6000
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView lstTransf 
      Height          =   3585
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Clique para selecionar um produto ja gravado."
      Top             =   2400
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   6324
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
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lote"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Origem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qtde.Transf."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Qtde.Origem"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Dt.Transferência"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Dt.Entrada"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Estab.Destino:"
      Height          =   240
      Left            =   3120
      TabIndex        =   19
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Situacão:"
      Height          =   240
      Left            =   6810
      TabIndex        =   14
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Estab.Origem:"
      Height          =   240
      Left            =   3135
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      Index           =   0
      X1              =   0
      X2              =   11415
      Y1              =   720
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      Index           =   1
      X1              =   0
      X2              =   11415
      Y1              =   2280
      Y2              =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "DtInicial:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "DtFinal:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Lote:"
      Height          =   240
      Index           =   2
      Left            =   7320
      TabIndex        =   8
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Produto:"
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   810
   End
End
Attribute VB_Name = "frmEstoqueTransfConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   cmbEstabOrigemAUX.Clear
   cmbEstabOrigem.Clear
   cmbEstabOrigem.AddItem "Todos"
   cmbEstabOrigemAUX.AddItem ""

   cmbEstabDestinoAUX.Clear
   cmbEstabDestino.Clear
   cmbEstabDestino.AddItem "Todos"
   cmbEstabDestinoAUX.AddItem ""

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbEstabOrigem.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      cmbEstabOrigemAUX.AddItem Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)

      cmbEstabDestino.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      cmbEstabDestinoAUX.AddItem Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)

      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbEstabDestinoAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstabDestino.Text = TRAZ_ESTABELECIMENTO(cmbEstabDestinoAUX.Text)

   cmbEstabOrigemAUX.Text = ""
   cmbEstabOrigem.Text = "Todos"

   cmbEstabDestino.Enabled = False
   'If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
   If TIPO_USUARIO = 5 Then _
      cmbEstabDestino.Enabled = True

   cmbSituacao.Clear
   cmbSituacao.AddItem "Aberto"
   cmbSituacao.AddItem "Fechado"
   cmbSituacao.AddItem "Transito"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         MONTA_CONSULTA
      Case "print"
         MONTA_REL
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTDTINI_GotFocus()

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

End Sub

Private Sub TXTDTFIM_GotFocus()

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtProduto.Text) <> "" Then _
         PROCESSA_DADOS_PRODUTOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub cmbEstabOrigem_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabOrigemAUX.ListIndex = cmbEstabOrigem.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbEstabOrigem_Click"
End Sub

Private Sub cmbEstabdestino_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabDestinoAUX.ListIndex = cmbEstabDestino.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbEstabdestino_Click"
End Sub

Private Sub cmdConsProd_Click()
'On Error GoTo ERRO_TRATA

   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Enabled = True
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub
Sub MONTA_CONSULTA()
'On Error GoTo ERRO_TRATA

   Dim LONTE_N As Long
   lote_n = 0
   PRODUTO_ID_N = 0
   CONT_N = 0
   lstTransf.Visible = False
   lstTransf.ListItems.Clear
   If chkOrdem.Value = 0 Then
      SQL3 = ""
      Else: SQL3 = "desc"
   End If

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from vwConsTransf WITH (NOLOCK)"

   SQL = SQL & " where codg_produto is not null "

   If Trim(cmbEstabDestinoAUX.Text) <> "" Then
      If IsNumeric(cmbEstabDestinoAUX.Text) Then
         SQL = SQL & " and estab_destino_id = " & cmbEstabDestinoAUX.Text
         SQL = SQL & " and estabelecimento_id = " & cmbEstabDestinoAUX.Text
      End If
   End If

   If Trim(cmbEstabOrigemAUX.Text) <> "" Then
      If IsNumeric(cmbEstabOrigemAUX.Text) Then
         SQL = SQL & " and estab_origem_id = " & cmbEstabOrigemAUX.Text
         SQL = SQL & " and estabelecimento_id = " & cmbEstabOrigemAUX.Text
      End If
   End If

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True
      SQL = SQL & " and DT_TRANSF >= '" & (txtDtIni.Text) & "'"
      SQL = SQL & " and DT_TRANSF <= '" & (txtDtFim.Text) & "'"
   End If

   If Trim(txtID.Text) <> "" Then _
      If IsNumeric(txtID.Text) Then _
         SQL = SQL & " and TRANSF_ID = " & txtID.Text

   If PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and vwConsTransf.produto_ID = " & PRODUTO_ID_N

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and vwConsTransf.situacao = '" & Left(Trim(cmbSituacao.Text), 1) & "'"

   If Trim(txtDesc.Text) <> "" Then
      CRITERIO_A = Chr$(39) & txtDesc.Text & "%" & Chr(39)
      SQL = SQL & " and descricao like " & CRITERIO_A
   End If

   SQL = SQL & " order by transf_id " & SQL3

'Debug.Print SQL

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabConsulta.EOF Then _
      MsgBox "Nenhum registro encontrado."
   While Not TabConsulta.EOF
      CONT_N = CONT_N + 1

'If TabConsulta.Fields("lote").Value = 1287 Then
'   MsgBox "e ai"
'End If
'If PRODUTO_ID_N <> TabConsulta.Fields("produto_id").Value And
      'If lote_n <> TabConsulta.Fields("lote").Value Then
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
         lote_n = TabConsulta.Fields("lote").Value

         Set item = lstTransf.ListItems.Add(, "seq." & CONT_N, TabConsulta.Fields("lote").Value)
         item.SubItems(1) = "" & Trim(TabConsulta.Fields("codg_produto").Value)
         item.SubItems(2) = "" & Trim(TabConsulta.Fields("descricao").Value)
         item.SubItems(3) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estab_origem_id").Value)
         item.SubItems(4) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estab_destino_id").Value)
         item.SubItems(5) = "" & Format(TabConsulta.Fields("qtde_transf").Value, strFormatacao3Digitos)
         item.SubItems(6) = "" & Format(TabConsulta.Fields("QTDE_ESTOQUE").Value, strFormatacao3Digitos)
         item.SubItems(7) = ""
         item.SubItems(8) = "" & TabConsulta.Fields("dt_transf").Value
         item.SubItems(9) = "" & TabConsulta.Fields("dt_entrada").Value

         SqL2 = ""
         If Not IsNull(TabConsulta.Fields("SITUACAO").Value) Then
            If Trim(TabConsulta.Fields("SITUACAO").Value) = "A" Then _
               SqL2 = "Aberto"
            If Trim(TabConsulta.Fields("SITUACAO").Value) = "T" Then
               SqL2 = "Transito"
               item.ForeColor = vbBlue
               item.ListSubItems(1).ForeColor = vbBlue
               item.ListSubItems(2).ForeColor = vbRed
               item.ListSubItems(3).ForeColor = vbRed
               item.ListSubItems(4).ForeColor = vbRed
               item.ListSubItems(5).ForeColor = vbRed
               item.ListSubItems(6).ForeColor = vbRed
               item.ListSubItems(7).ForeColor = vbRed
               item.ListSubItems(8).ForeColor = vbRed
               item.ListSubItems(9).ForeColor = vbRed
            End If
            If Trim(TabConsulta.Fields("SITUACAO").Value) = "F" Then
               SqL2 = "Fechado"
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
         End If
         item.SubItems(7) = "" & Trim(SqL2)
      'End If
      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   lstTransf.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub

Sub LIMPA_TUDO()
   cmbEstabOrigemAUX.Text = ""
   cmbEstabOrigem.Text = ""
   cmbEstabDestinoAUX.Text = ""
   cmbEstabDestino.Text = ""
   SQL3 = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.Text = ""
   txtID.Text = ""
   txtDesc.Text = ""
   chkOrdem.Value = 1
   txtProduto.Text = ""
   PRODUTO_ID_N = 0
   cmbSituacao.Text = ""
   cmbEstabOrigem.Text = "Todos"
   cmbEstabOrigemAUX.Text = ""
   cmbEstabDestino.Text = "Todos"
   cmbEstabDestinoAUX.Text = ""

   cmbEstabDestino.Enabled = False
   'If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
   If TIPO_USUARIO = 5 Then _
      cmbEstabDestino.Enabled = True

   cmbEstabDestinoAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstabDestino.Text = TRAZ_ESTABELECIMENTO(cmbEstabDestinoAUX.Text)
End Sub

Sub MONTA_REL()
   DATA_INI = DMA(txtDtIni.Text, "i")
   DATA_FIM = DMA(txtDtFim.Text, "f")

   FORMULA_REL = "{ESTOQUEtransf.produto_id} > 0 "

   If Trim(txtID.Text) <> "" Then _
      If IsNumeric(txtID.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.TRANSF_ID} = " & txtID.Text

   If Trim(txtDtIni.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.DT_TRANSF} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"

   If Trim(txtDtFim.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.DT_TRANSF} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"

   If Trim(txtDesc.Text) <> "" Then
      CRITERIO_A = Chr$(39) & txtDesc.Text & "%" & Chr(39)
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.descricao} like " & CRITERIO_A
   End If

   If Trim(cmbEstabDestinoAUX.Text) <> "" Then _
      If IsNumeric(cmbEstabDestinoAUX.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.estab_destino_id} = " & cmbEstabDestinoAUX.Text

   If Trim(cmbEstabOrigemAUX.Text) <> "" Then _
      If IsNumeric(cmbEstabOrigemAUX.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.estab_origem_id} = " & cmbEstabOrigemAUX.Text

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

'MsgBox FORMULA_REL

   Nome_Relatorio = "estoque_transf.rpt"
   frmRELATORIO10.Show 1
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

   txtDesc.Text = "" & Trim(DESC_PRODUTO_A)
   txtProduto.Text = "" & CODG_PRODUTO_A
   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
