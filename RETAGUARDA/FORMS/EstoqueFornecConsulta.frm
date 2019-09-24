VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoqueFornecConsulta 
   Caption         =   "Consulta MP Fornecedor"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EstoqueFornecConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optRetorno 
      Caption         =   "Dt.Retorno"
      Height          =   255
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   2295
   End
   Begin VB.OptionButton optMov 
      Caption         =   "Dt.Movimento"
      Height          =   255
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   720
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      Picture         =   "EstoqueFornecConsulta.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
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
      Left            =   3720
      MaxLength       =   100
      TabIndex        =   20
      Top             =   1320
      Width           =   4575
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
      Left            =   9240
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDesc 
      Height          =   360
      Left            =   3720
      TabIndex        =   8
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txtID 
      Height          =   360
      Left            =   9240
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cmbEstabOrigem 
      Height          =   360
      Left            =   9240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbSituacao 
      Height          =   360
      Left            =   9240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      Picture         =   "EstoqueFornecConsulta.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
      Top             =   1800
      Width           =   1935
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   360
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
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
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   6
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
               Picture         =   "EstoqueFornecConsulta.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueFornecConsulta.frx":843E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueFornecConsulta.frx":94CD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueFornecConsulta.frx":A667
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueFornecConsulta.frx":C4DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EstoqueFornecConsulta.frx":E1E8
               Key             =   ""
            EndProperty
         EndProperty
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
      Begin Threed.SSCheck chkOrdem 
         Height          =   270
         Left            =   6600
         TabIndex        =   10
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
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10545
      DesignHeight    =   6075
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   360
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSComctlLib.ListView lstTransf 
      Height          =   3615
      Left            =   60
      TabIndex        =   12
      ToolTipText     =   "Clique para selecionar um produto ja gravado."
      Top             =   2400
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6376
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
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Origem"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fornecedor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qtde.Envio"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Qtde.Retorno"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Situação"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Dt.Movimento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Dt.Retorno"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   360
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
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
   Begin VB.Label Label1 
      Caption         =   "Fornec.:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Produto:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Lote:"
      Height          =   240
      Index           =   2
      Left            =   8775
      TabIndex        =   17
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "DtFinal:"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "DtInicial:"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   15
      Top             =   840
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   4
      Index           =   1
      X1              =   0
      X2              =   11415
      Y1              =   2280
      Y2              =   2295
   End
   Begin VB.Label Label15 
      Caption         =   "Origem:"
      Height          =   240
      Left            =   8490
      TabIndex        =   14
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Situacão:"
      Height          =   240
      Left            =   8355
      TabIndex        =   13
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "frmEstoqueFornecConsulta"
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

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbEstabOrigem.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      cmbEstabOrigemAUX.AddItem Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)

      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbEstabOrigemAUX.Text = ""
   cmbEstabOrigem.Text = "Todos"

   cmbSituacao.Clear
   cmbSituacao.AddItem "Transito"
   cmbSituacao.AddItem "Reservado"
   cmbSituacao.AddItem "Fechado"
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

Private Sub cmdConsProd_Click()
'On Error GoTo ERRO_TRATA

   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
      Case vbKeyBack
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            txtCNPJCPF.Mask = "##############"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
       
      If txtCNPJCPF.Text <> "" Then _
         txtNome.SetFocus

   ElseIf KeyAscii = vbKeyDelete Then
      If Not IsNumeric(txtCNPJCPF.Text) Then
         txtCNPJCPF.Mask = "##############"
      End If
   ElseIf KeyAscii = vbKeyBack Then
      If Not IsNumeric(txtCNPJCPF.Text) Then
         txtCNPJCPF.Mask = "##############"
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_DADOS_PESSOA

   txtCNPJCPF.PromptInclude = False
   If Len(Trim(txtCNPJCPF.Text)) > 0 Then
      If CInt(Len(Trim(txtCNPJCPF.Text))) = 11 Then
         If Not ValidaCPF(Trim(txtCNPJCPF.Text)) Then
            MsgBox "CPF com DV incorreto !!!"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF = ""
            'ssTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      ElseIf CInt(Len(Trim(txtCNPJCPF.Text))) = 14 Then
         If Not VALIDACNPJ(Trim(txtCNPJCPF.Text)) Then
            MsgBox "CNPJ com DV incorreto !!! "
            txtCNPJCPF.PromptInclude = False
            'ssTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "CNPJ/CPF com DV incorreto !!! "
         txtCNPJCPF = ""
         'ssTab.Tab = 0
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   ElseIf Len(Trim(txtCNPJCPF.Text)) <> 0 Then
       MsgBox "CNPJ/CPF com DV incorreto !!! "
       txtCNPJCPF = ""
       'ssTab.Tab = 0
       TXTCNPJCPF_GotFocus
       txtCNPJCPF.SetFocus
       Exit Sub
   End If
   
   txtCNPJCPF.PromptInclude = False
   CRITERIO_A = Trim(txtCNPJCPF.Text)
   txtCNPJCPF.PromptInclude = False
   
   If Trim(txtCNPJCPF.Text) <> "" Then
      CRITERIO_A = Trim(txtCNPJCPF.Text)

      If Not IsNull(Trim(txtCNPJCPF.Text)) Then
          If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(Trim(txtCNPJCPF.Text)) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.Text = CRITERIO_A
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_LostFocus"
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
   ESTOQUEFORNEC_ID_n = 0
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

   SQL = "select ESTOQUEFORNEC.ESTOQUEFORNEC_ID, ESTOQUEFORNEC.SEQ_ID, ESTOQUEFORNEC.ESTAB_ORIGEM_ID, ESTOQUEFORNEC.FORNECEDOR_ID, "
   SQL = SQL & " ESTOQUEFORNEC.PRODUTO_ID, ESTOQUEFORNEC.QTDE_ENVIO, ESTOQUEFORNEC.QTDE_RETORNO, ESTOQUEFORNEC.DT_MOVIMENTO,"
   SQL = SQL & " ESTOQUEFORNEC.DT_RETORNO, ESTOQUEFORNEC.SITUACAO, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, FORNECEDOR.PESSOA_ID,"
   SQL = SQL & " PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFornec, ESTABELECIMENTO.EMPRESA_ID, ESTABELECIMENTO.DESCRICAO AS NomeEstab"
   SQL = SQL & " from ESTOQUEFORNEC WITH (NOLOCK)"
   SQL = SQL & " Inner Join PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUEFORNEC.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUEFORNEC.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUEFORNEC.ESTAB_ORIGEM_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID"

   SQL = SQL & " where ESTOQUEFORNEC_ID is not null "

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

      If optMov.Value = True Then
         SQL = SQL & " and DT_MOVIMENTO >= '" & (txtDtIni.Text) & "'"
         SQL = SQL & " and DT_MOVIMENTO <= '" & (txtDtFim.Text) & "'"
         Else
            SQL = SQL & " and DT_RETORNO >= '" & (txtDtIni.Text) & "'"
            SQL = SQL & " and DT_RETORNO <= '" & (txtDtFim.Text) & "'"
      End If
   End If

   If Trim(txtID.Text) <> "" Then _
      If IsNumeric(txtID.Text) Then _
         SQL = SQL & " and ESTOQUEFORNEC_ID = " & txtID.Text

   If PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and ESTOQUEFORNEC.produto_ID = " & PRODUTO_ID_N

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao = '" & Left(Trim(cmbSituacao.Text), 1) & "'"

   If Trim(txtDesc.Text) <> "" Then
      CRITERIO_A = Chr$(39) & txtDesc.Text & "%" & Chr(39)
      SQL = SQL & " and descricao like " & CRITERIO_A
   End If

   If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N

   SQL = SQL & " order by ESTOQUEFORNEC_ID " & SQL3

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabConsulta.EOF Then _
      MsgBox "Nenhum registro encontrado."
   While Not TabConsulta.EOF
      CONT_N = CONT_N + 1

      If ESTOQUEFORNEC_ID_n <> TabConsulta.Fields("ESTOQUEFORNEC_ID").Value Then
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
         ESTOQUEFORNEC_ID_n = TabConsulta.Fields("ESTOQUEFORNEC_ID").Value

         Set item = lstTransf.ListItems.Add(, "seq." & CONT_N, TabConsulta.Fields("ESTOQUEFORNEC_ID").Value)
         item.SubItems(1) = "" & Trim(TabConsulta.Fields("codg_produto").Value)
         item.SubItems(2) = "" & Trim(TabConsulta.Fields("descricao").Value)
         item.SubItems(3) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estab_origem_id").Value)
         item.SubItems(4) = "" & TRAZ_NOME_FORNECEDOR(TabConsulta.Fields("fornecedor_id").Value, TabConsulta.Fields("PESSOA_id").Value)
         item.SubItems(5) = "" & Format(TabConsulta.Fields("qtde_envio").Value, strFormatacao3Digitos)
         item.SubItems(6) = "" & Format(TabConsulta.Fields("QTDE_retorno").Value, strFormatacao3Digitos)
         item.SubItems(7) = ""
         item.SubItems(8) = "" & TabConsulta.Fields("dt_movimento").Value
         item.SubItems(9) = "" & TabConsulta.Fields("dt_retorno").Value

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
      End If
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
   PESSOA_ID_N = 0
   FORNEC_ID_N = 0
   txtNome.Text = ""
   txtNome.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   cmbEstabOrigemAUX.Text = ""
   cmbEstabOrigem.Text = ""
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
End Sub

Sub MONTA_REL()
   DATA_INI = DMA(txtDtIni.Text, "i")
   DATA_FIM = DMA(txtDtFim.Text, "f")

   FORMULA_REL = "{ESTOQUEtransf.produto_id} > 0 "

   If Trim(txtID.Text) <> "" Then _
      If IsNumeric(txtID.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.ESTOQUEFORNEC_ID} = " & txtID.Text

   If Trim(txtDtIni.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.dt_movimento} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"

   If Trim(txtDtFim.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.dt_movimento} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"

   If Trim(txtDesc.Text) <> "" Then
      CRITERIO_A = Chr$(39) & txtDesc.Text & "%" & Chr(39)
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.descricao} like " & CRITERIO_A
   End If

   If Trim(cmbEstabOrigemAUX.Text) <> "" Then _
      If IsNumeric(cmbEstabOrigemAUX.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {ESTOQUEtransf.estab_origem_id} = " & cmbEstabOrigemAUX.Text

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

'MsgBox FORMULA_REL

   Nome_Relatorio = "estoque_transf.rpt"
   frmRELATORIO10.Show 1
End Sub

Sub MOSTRA_DADOS_PESSOA()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   FORNEC_ID_N = 0
   txtNome.Text = ""

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If IsNumeric(txtCNPJCPF.Text) Then
         Dim TabPessoa     As New ADODB.Recordset

         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         SQL = "select  PESSOA.PESSOA_ID, PESSOA.CNPJCPF, FORNECEDOR.FORNECEDOR_ID, FORNECEDOR.STATUS, PESSOA.DESCRICAO"
         SQL = SQL & " from PESSOA  WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN FORNECEDOR  WITH (NOLOCK)"
         SQL = SQL & "ON PESSOA.PESSOA_ID = FORNECEDOR.PESSOA_ID"

         SQL = SQL & " where CNPJCPF = '" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & " and status = 'A' "
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then
            PESSOA_ID_N = 0 & TabPessoa.Fields("pessoa_id").Value
            FORNEC_ID_N = 0 & TabPessoa.Fields("fornecedor_id").Value
            txtNome.Text = "" & Trim(TabPessoa.Fields("descricao").Value)
         End If
         If TabPessoa.State = 1 Then _
            TabPessoa.Close
         Else: Exit Sub
      End If
      Else: Exit Sub
   End If
   If PESSOA_ID_N <= 0 Then _
      Exit Sub

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PESSOA"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

   If TabProduto.State = 1 Then _
      TabProduto.Close

   'txtQTDE.Text = Format(Qtde_N, strFormatacao3Digitos)
   txtProduto.Text = Trim(CODG_PRODUTO_A)
   txtDesc.Text = DESC_PRODUTO_A
   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
