VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMGV 
   Caption         =   "Gerar Arquivo Texto para Balança Toledo"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MGV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbFamiliaAux 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   7440
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   7440
      TabIndex        =   1
      ToolTipText     =   "Selecione o grupo do produto."
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CheckBox chkBalanca 
      Caption         =   "Balança?"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Height          =   240
      Left            =   10440
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   240
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Informe o código do produto."
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdConsProd2 
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
      Left            =   2190
      Picture         =   "MGV.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Consulta Produto"
      Top             =   1080
      Width           =   405
   End
   Begin VB.TextBox txtDesc2 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2640
      MaxLength       =   29
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11895
      DesignHeight    =   6090
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "cons"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Key             =   "ex"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   -120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":89E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":9C7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":AD87
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":BE16
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":CDCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":E048
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":F2CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MGV.frx":1076F
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstProduto 
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   7858
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "QTD.BC."
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qtde.Dep."
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pr.Venda"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Pr.Atacado"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Pr.Custo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Fornecedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "+ Est."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "- Est."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "%"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Referência"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Grupo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ST"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "produto_id"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Familia"
      Height          =   240
      Left            =   7440
      TabIndex        =   11
      Top             =   840
      Width           =   720
   End
   Begin VB.Label txtQtde 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   7440
      Width           =   60
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Produto:"
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   810
   End
End
Attribute VB_Name = "frmMGV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim Tipo_Produto_N      As Integer
   Dim CODG_PRODUTO_A      As String
   Dim Descricao_Produto_A As String
   Dim UNIDADE_MEDIDA_A    As String
   Dim Preco_Venda_A       As String
   Dim Linha_Produto_A     As String
   Dim Extra01             As String

Private Sub Form_Load()
    CARREGA_FAMILIA_PRODUTO
    Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "cons"
         CONSULTA_TUDO
      Case "ex"
         VAI_TXT
      Case "limpar"
         Toolbar1.Buttons(4).Enabled = False
         txtCodigo.Text = ""
         txtDesc2.Text = ""
         cmbFamilia.Text = ""
         cmbFamiliaAUX.Text = ""
         chkBalanca.Value = 0
         chkTodos.Value = 0
         lstProduto.ListItems.Clear
         txtCodigo.SetFocus
      Case "sair"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbFamilia_Click()
On Error Resume Next

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmbFamilia_LostFocus()
   cmbFamilia.BackColor = &HFFFFFF
End Sub

Private Sub cmbFamilia_GotFocus()
   cmbFamilia.SelStart = 0
   cmbFamilia.SelLength = Len(cmbFamilia)
   cmbFamilia.BackColor = &HC0FFFF
End Sub

Private Sub chkTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstProduto.ListItems.Count > 0 Then
      For i = lstProduto.ListItems.Count To 1 Step -1
         If chkTodos.Value = 1 Then
            lstProduto.ListItems(i).Checked = True
            Else: lstProduto.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "chkTodos_Click"
End Sub

Private Sub cmdConsProd2_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txttxtcodigo.Text = SQL3
      txtCodigo.SetFocus
   End If
   SQL3 = ""
   txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Sub GERA_ITENSMGV(PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Tipo_Produto_N = 0
   CODG_PRODUTO_A = ""
   Descricao_Produto_A = ""
   UNIDADE_MEDIDA_A = ""
   Preco_Venda_A = ""
   Linha_Produto_A = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select codg_produto,descricao,UNIDADE_MEDIDA,preco_venda,produto_balanca "
   SQL = SQL & " from produto"
   SQL = SQL & " where produto_id = " & PROD_ID_N
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then _
      Exit Sub

   UNIDADE_MEDIDA_A = "KG"
   If Not IsNull(TabProduto.Fields("Unidade_Medida").Value) Then
      If Trim(TabProduto.Fields("Unidade_Medida").Value) <> "" Then
         UNIDADE_MEDIDA_A = Trim(TabProduto.Fields("Unidade_Medida").Value)

         If UCase(UNIDADE_MEDIDA_A) = UCase("KG") Then
            Tipo_Produto_N = 0
            Else
               If UCase(UNIDADE_MEDIDA_A) = UCase("UN") Then _
                  Tipo_Produto_N = 1
         End If
      End If
   End If

   CODG_PRODUTO_A = "" & Trim(Left(TabProduto.Fields("codg_produto").Value, 6))
   Descricao_Produto_A = "" & Trim(Left(TabProduto.Fields("descricao").Value, 25))
   Preco_Venda_A = Format(TabProduto.Fields("preco_venda").Value, strFormatacao2Digitos)

   If TabProduto.State = 1 Then _
      TabProduto.Close

   Preco_Venda_A = Replace(Preco_Venda_A, ",", "")
   Preco_Venda_A = Replace(Preco_Venda_A, ".", "")

   Preco_Venda_A = Trim(Preco_Venda_A)
   If Len(Preco_Venda_A) < 6 Then _
      SQL = INSERIR_0(6, Preco_Venda_A)

   CODG_PRODUTO_A = Trim(CODG_PRODUTO_A)
   If Len(CODG_PRODUTO_A) < 6 Then _
      SQL = INSERIR_0(6, CODG_PRODUTO_A)

   Descricao_Produto_A = Trim(Descricao_Produto_A)
   If Len(Descricao_Produto_A) < 25 Then _
      SQL = INSERIR_BRANCO(25, Descricao_Produto_A)

   Linha_Produto_A = ""

'DD(2)   Código do departamento
   Linha_Produto_A = "01"

'T(1)    Tipo de produto
   '[0] => Venda por peso
   '[1] => Venda por unidade
   '[2] => EAN-13 por peso
   '[3] => Venda por peso glaciado
   '[4] => Venda por peso drenado
   '[5] => EAN-13 por unidade
   Linha_Produto_A = Linha_Produto_A & Tipo_Produto_N

'CCCCCC(6)     Código do Item
   Linha_Produto_A = Linha_Produto_A & CODG_PRODUTO_A

'PPPPPP(6)     Preço/kg ou Preço/Unid. do item
   Linha_Produto_A = Linha_Produto_A & Preco_Venda_A

'VVV(3)        Dias de validade do produto
   Linha_Produto_A = Linha_Produto_A & "000"

'D1(25)        Descritivo do Item – Primeira Linha
   Linha_Produto_A = Linha_Produto_A & Descricao_Produto_A

'D2(25)     Descritivo do Item – Segunda Linha
   'Linha_Produto_A = Linha_Produto_A & "                         "
   Linha_Produto_A = Linha_Produto_A & Descricao_Produto_A

'RRRRRR(6)     Código da Informação Extra do item
   Linha_Produto_A = Linha_Produto_A & "000000"

'FFFF(4)       Código da Imagem do Item
   'Linha_Produto_A = Linha_Produto_A & "0000"
   Linha_Produto_A = Linha_Produto_A & "000"

'IIIIII(6)     Código da Informação Nutricional
   'Linha_Produto_A = Linha_Produto_A & "000000"
   Linha_Produto_A = Linha_Produto_A & "00000"

'DV(1)         Impressão da Data de Validade
'[1] => Imprime Data de Validade
'[0] => Não Imprime Data de Validade
   Linha_Produto_A = Linha_Produto_A & "0"

'DE(1)         Impressão da Data de Embalagem
'[1] => Imprime Data de Embalagem
'[0] => Não Imprime Data de Embalagem
   Linha_Produto_A = Linha_Produto_A & "0"

'CF(4)         Código do Fornecedor
   Linha_Produto_A = Linha_Produto_A & "0000"

'L(12)         Lote
   'Linha_Produto_A = Linha_Produto_A & "            "
   Linha_Produto_A = Linha_Produto_A & "000000000000"

'G(11)         Código EAN-13 Especial
   'Linha_Produto_A = Linha_Produto_A & "           "
   Linha_Produto_A = Linha_Produto_A & "00000000000"

'Z(1)          Versão do preço
   'Linha_Produto_A = Linha_Produto_A & " "
   Linha_Produto_A = Linha_Produto_A & "0"

GoTo VAZA

'CS(4)         Código do Som
   Linha_Produto_A = Linha_Produto_A & "0000"

'CT(4)         Código de Tara Pré-determinada
   Linha_Produto_A = Linha_Produto_A & "0000"

'FR(4)         Código do Fracionador
   Linha_Produto_A = Linha_Produto_A & "0000"

'CE1(4)        Código do Campo Extra 1
   Linha_Produto_A = Linha_Produto_A & "0000"

'CE2(4)        Código do Campo Extra 2
   Linha_Produto_A = Linha_Produto_A & "0000"

'CON(4)        Código da Conservação
   Linha_Produto_A = Linha_Produto_A & "0000"

'EAN(12)       EAN-13 de Fornecedor
   Linha_Produto_A = Linha_Produto_A & "            "
   'Linha_Produto_A = Linha_Produto_A & ""

'GL(6)         Percentual de Glaciamento
   Linha_Produto_A = Linha_Produto_A & "000000"

'|DA           Sequencia de departamentos associados
'Ex: Para associar departamentos 2 e 5: |0205|
   Linha_Produto_A = Linha_Produto_A & "|  |"

'|D3(35)       Descritivo do Item – Terceira Linha
   Linha_Produto_A = Linha_Produto_A & "                                   "

'D4(35)        Descritivo do Item – Terceira Linha
   Linha_Produto_A = Linha_Produto_A & "                                   "
   'Linha_Produto_A = Linha_Produto_A & ""

'CE3(6)        Código do Campo Extra 3
   Linha_Produto_A = Linha_Produto_A & "000000"

'CE4(6)        Código do Campo Extra 4
   Linha_Produto_A = Linha_Produto_A & "000000"

'MIDIA(6)      Código da mídia (Prix 6 Touch)
   Linha_Produto_A = Linha_Produto_A & "000000"

'PPPPPP(6)     Preço Promocional - Preço/kg ou Preço/Unid. do item
   Linha_Produto_A = Linha_Produto_A & "      "
   'Linha_Produto_A = Linha_Produto_A & ""

'SF(1)
'[0] = Utiliza o fornecedor associado
'[1] = Balança solicita fornecedor após chamada do PLU
   Linha_Produto_A = Linha_Produto_A & "0"

'|FFFFFFFF(n)  Código de Fornecedor Associado
'Ex: Para associar fornecedores 2 e 5: |000002000005|
   'Linha_Produto_A = Linha_Produto_A & "|000000000000"
   Linha_Produto_A = Linha_Produto_A & "|000000|"

'|ST(1)
'[0] = Não solicita tara na balança
'[1] = Solicita Tara na Balança
   Linha_Produto_A = Linha_Produto_A & "0"

'| BNA(n)      Sequência de balanças onde o item não estará ativo.
'Ex: Para associar balanças 2 e 5 com itens inativos: |0205|
   'Linha_Produto_A = Linha_Produto_A & "|0000"
   Linha_Produto_A = Linha_Produto_A & "|00"

'| (+CR+LF)
   Linha_Produto_A = Linha_Produto_A & "|(+CR+LF)"

VAZA:

   Print #1, Tab(1); Linha_Produto_A


Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_ITENSMGV"
End Sub


'=========================================
Sub GERA_ITENSMGV_old(PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Tipo_Produto_N = 0
   CODG_PRODUTO_A = ""
   Descricao_Produto_A = ""
   UNIDADE_MEDIDA_A = ""
   Preco_Venda_A = ""
   Linha_Produto_A = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select codg_produto,descricao,UNIDADE_MEDIDA,preco_venda,produto_balanca "
   SQL = SQL & " from produto"
   SQL = SQL & " where produto_id = " & PROD_ID_N
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then _
      Exit Sub

   UNIDADE_MEDIDA_A = "KG"
   If Not IsNull(TabProduto.Fields("Unidade_Medida").Value) Then
      If Trim(TabProduto.Fields("Unidade_Medida").Value) <> "" Then
         UNIDADE_MEDIDA_A = Trim(TabProduto.Fields("Unidade_Medida").Value)

         If UCase(UNIDADE_MEDIDA_A) = UCase("KG") Then
            Tipo_Produto_N = 0
            Else
               If UCase(UNIDADE_MEDIDA_A) = UCase("UN") Then _
                  Tipo_Produto_N = 1
         End If
      End If
   End If

   CODG_PRODUTO_A = "" & Trim(Left(TabProduto.Fields("codg_produto").Value, 6))
   Descricao_Produto_A = "" & Trim(Left(TabProduto.Fields("descricao").Value, 25))
   Preco_Venda_A = Format(TabProduto.Fields("preco_venda").Value, strFormatacao2Digitos)

   If TabProduto.State = 1 Then _
      TabProduto.Close

   Preco_Venda_A = Replace(Preco_Venda_A, ",", "")
   Preco_Venda_A = Replace(Preco_Venda_A, ".", "")

   Preco_Venda_A = Trim(Preco_Venda_A)
   If Len(Preco_Venda_A) < 6 Then _
      SQL = INSERIR_0(6, Preco_Venda_A)

   CODG_PRODUTO_A = Trim(CODG_PRODUTO_A)
   If Len(CODG_PRODUTO_A) < 6 Then _
      SQL = INSERIR_0(6, CODG_PRODUTO_A)

   Descricao_Produto_A = Trim(Descricao_Produto_A)
   If Len(Descricao_Produto_A) < 25 Then _
      SQL = INSERIR_BRANCO(25, Descricao_Produto_A)

   Linha_Produto_A = ""

'DD(2)   Código do departamento
   Linha_Produto_A = "00"
   Linha_Produto_A = Linha_Produto_A & "00"

'T(1)    Tipo de produto
   '[0] => Venda por peso
   '[1] => Venda por unidade
   '[2] => EAN-13 por peso
   '[3] => Venda por peso glaciado
   '[4] => Venda por peso drenado
   '[5] => EAN-13 por unidade
   Linha_Produto_A = Linha_Produto_A & Format(Mid(Tipo_Produto_N, 1, 1), "0")

'CCCCCC(6)     Código do Item
   Linha_Produto_A = Linha_Produto_A & CODG_PRODUTO_A

'PPPPPP(6)     Preço/kg ou Preço/Unid. do item
   Linha_Produto_A = Linha_Produto_A & Preco_Venda_A

'VVV(3)        Dias de validade do produto
   Linha_Produto_A = Linha_Produto_A & "000"

'D1(25)        Descritivo do Item – Primeira Linha
   Linha_Produto_A = Linha_Produto_A & Descricao_Produto_A

'D2(25)     Descritivo do Item – Segunda Linha
   Linha_Produto_A = Linha_Produto_A & "                         "
   'Linha_Produto_A = Linha_Produto_A & ""
'set
'RRRRRR(6)     Código da Informação Extra do item
   Linha_Produto_A = Linha_Produto_A & Mid(Extra01, 1, 50)

'RRRRRR(6)     Código da Informação Extra do item
   Linha_Produto_A = Linha_Produto_A & Mid(Extra01, 1, 50)

'RRRRRR(6)     Código da Informação Extra do item
   Linha_Produto_A = Linha_Produto_A & Mid(Extra01, 1, 50)

'RRRRRR(6)     Código da Informação Extra do item
   Linha_Produto_A = Linha_Produto_A & Mid(Extra01, 1, 50)

'RRRRRR(6)     Código da Informação Extra do item
   Linha_Produto_A = Linha_Produto_A & Mid(Extra01, 1, 50)
'''''''''''''''''

Open PATH_TXT & "Itensmgv.txt" For Output As #1
   Print #1, Tab(1); Linha_Produto_A

Close #1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_ITENSMGV"
End Sub

Public Function Gera_Balanca(Optional Arquivo As String)
   Dim Linha As String
   Dim DescricaoFilizola As String * 22

   On Error GoTo Erro

   Set TbProduto = Banco.Execute("select * from Produto Where Referencia='" & "1" & "'")
   If TbProduto.EOF And TbProduto.BOF Then
       MsgBox "Não foi possivel localizar nenhum produto com o campo Balança = 1. Para criar o arquivo correto, altere os produto desejados.", vbCritical, "PROGma"
   Else
      TbProduto.MoveFirst
      Open Arquivo For Output As #1
      Do While TbProduto.EOF = False

         Linha = ""
         Departamento = "01"
         Etiqueta = "00"
         TIPO = "0"
         If Not IsNull(TbProduto("CODIGO")) Then CodProduto = CCur(TbProduto("CODIGO"))
         If Not IsNull(TbProduto("VVenda")) Then Preco = Replace(Replace(Format(TbProduto("VVenda"), "##,##0.00"), ",", ""), ".", "")
         Validade = "000"
         If Not IsNull(TbProduto("DESCRICAO")) Then Descricao01 = Mid(TbProduto("DESCRICAO"), 1, 25)
         If Not IsNull(TbProduto("DESCRICAO")) Then Descricao02 = Mid(TbProduto("DESCRICAO"), 26, 25)
         If Not IsNull(TbProduto("DESCRICAO")) Then DescricaoFilizola = Mid(TbProduto("DESCRICAO"), 1, 22)
         
         If TipBalanca = "TOLEDO" Then
            Linha = Format(Mid(Departamento, 1, 2), "00")
            Linha = Linha & Format(Mid(Etiqueta, 1, 2), "00")
            Linha = Linha & Format(Mid(TIPO, 1, 1), "0")
            Linha = Linha & Format(Mid(CodProduto, 1, 6), "000000")
            Linha = Linha & Format(Mid(Preco, 1, 6), "000000")
            Linha = Linha & Format(Mid(Validade, 1, 3), "000")
            Linha = Linha & Mid(Descricao01, 1, 25)
            Linha = Linha & Mid(Descricao02, 1, 25)
            Linha = Linha & Mid(Extra01, 1, 50)
            Linha = Linha & Mid(Extra02, 1, 50)
            Linha = Linha & Mid(Extra03, 1, 50)
            Linha = Linha & Mid(Extra04, 1, 50)
            Linha = Linha & Mid(Extra05, 1, 50)
            Print #1, Linha
         
         
         ElseIf TipBalanca = "FILIZOLA" Then
            Linha = Linha & Format(Mid(CodProduto, 1, 6), "000000")
            Linha = Linha & "p"
            Linha = Linha & Mid(DescricaoFilizola, 1, 22)
            Linha = Linha & Format(Mid(Preco, 1, 7), "0000000") & "000"
            Print #1, Linha
         End If
         TbProduto.MoveNext
      Loop
      Close #1
      MsgBox "Arquivo de balança gerado com sucesso no: " & Arquivo, vbInformation, "PROGma"
   End If

Exit Function
Erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "PROGma": Exit Function
End Function

Sub CONSULTA_TUDO()

   CONT_N = 0
   Toolbar1.Buttons(4).Enabled = False

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select count(produto_id) from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where situacao = 'A'"

   If chkBalanca.Value = 0 Then
      SQL = SQL & " and produto_balanca = 'FALSE'"
      Else: SQL = SQL & " and produto_balanca = 'TRUE'"
   End If

   If Trim(txtCodigo.Text) <> "" Then _
      SQL = SQL & " and codg_produto = '" & Trim(txtCodigo.Text) & "'"

   If Trim(txtDesc2.Text) <> "" Then _
      SQL = SQL & " and descricao like '" & UCase(Trim(txtDesc2.Text)) & "%" & "'"

    If Trim(cmbFamiliaAUX.Text) <> "" Then _
        SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then _
      If Not IsNull(TabProduto.Fields(0).Value) Then _
         CONT_N = 0 & TabProduto.Fields(0).Value

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If CONT_N > 500 Then
      Msg = "Esta operação irá processar todos produtos cadastrado, deseja continuar ? " & CONT_N & " registros"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbNo Then _
         Exit Sub
   End If

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where situacao = 'A'"

   If chkBalanca.Value = 0 Then
      SQL = SQL & " and produto_balanca = 'FALSE'"
      Else: SQL = SQL & " and produto_balanca = 'TRUE'"
   End If

   If Trim(txtCodigo.Text) <> "" Then _
      SQL = SQL & " and codg_produto = '" & Trim(txtCodigo.Text) & "'"

   If Trim(txtDesc2.Text) <> "" Then _
      SQL = SQL & " and descricao like '" & UCase(Trim(txtDesc2.Text)) & "%" & "'"

    If Trim(cmbFamiliaAUX.Text) <> "" Then _
        SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   SQL = SQL & " order by descricao"

   SETA_GRID
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim dblContador   As Double
   Dim VALOR_CUSTO_N As Double

   lstProduto.Visible = False
   lstProduto.ListItems.Clear
   dblContador = 0

   If TabProduto.State = 1 Then _
      TabProduto.Close

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      MsgBox "Não encontrei nenhum produto com o CRITERIO_A de procura especificado", vbExclamation
      txtDesc2.SetFocus
      Exit Sub
      Else
         Toolbar1.Buttons(4).Enabled = True
         TabProduto.MoveFirst
   End If
   Me.Enabled = False
   While Not TabProduto.EOF
      DoEvents
      dblContador = dblContador + 1

      Me.Caption = "Aguarde, Processando ...  " & dblContador

      Set item = lstProduto.ListItems.Add(, "seq." & dblContador, Trim(TabProduto.Fields("codg_produto").Value))
      
      item.SubItems(1) = "" & Trim(TabProduto!DESCRICAO)
      item.SubItems(2) = "" & Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value), strFormatacao3Digitos)

      item.SubItems(4) = "" & Format(0, strFormatacao3Digitos)
      item.SubItems(4) = "-"

      If CONECTA_AUXILIAR.State = 1 Then _
         item.SubItems(4) = "" & Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("PRODUTO_ID").Value), strFormatacao3Digitos)

      item.SubItems(5) = "" & Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)

      VALOR_CUSTO_N = 0 & TabProduto!PRECO_CUSTO

      item.SubItems(7) = "" & Format(0, strFormatacao2Digitos)
      If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
         item.SubItems(7) = "" & Format(VALOR_CUSTO_N, strFormatacao2Digitos)

      If Not IsNull(TabProduto.Fields("produto_id").Value) Then
         If TabProduto.Fields("produto_id").Value > 0 Then
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

            If Not IsNull(TabProduto.Fields("fornecedor_id").Value) Then
               SQL = "select Descricao from vwFornecedor WITH (NOLOCK)"
               SQL = SQL & " where fornecedor_id = " & TabProduto.Fields("fornecedor_id").Value
               TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabFornecedor.EOF Then _
                  item.SubItems(8) = "" & Trim(TabFornecedor.Fields(0).Value)
            End If
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

         End If
      End If

      item.SubItems(9) = "" & Format(TabProduto!Qtd_minimo, strFormatacao3Digitos)
      item.SubItems(10) = "" & Format(TabProduto!qtd_maximo, strFormatacao3Digitos)
      item.SubItems(11) = "0"
      item.SubItems(12) = "" & Trim(TabProduto!REFERENCIA)
      item.SubItems(14) = "" & TabProduto.Fields("situacao_tributaria").Value
      item.SubItems(15) = "" & TabProduto.Fields("produto_id").Value

      NUMR_ID_N = 0 & TabProduto!FAMILIAPRODUTO_ID

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where familiaproduto_id = " & NUMR_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp!DESCRICAO) Then
            item.SubItems(13) = TabTemp!DESCRICAO
            Else: item.SubItems(13) = "SEM GRUPO"
         End If
         Else: item.SubItems(13) = "SEM GRUPO"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If TabProduto.Fields("situacao").Value = "A" Then
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
         item.ListSubItems(10).ForeColor = vbBlue
         item.ListSubItems(11).ForeColor = vbBlue
         item.ListSubItems(12).ForeColor = vbBlue
         item.ListSubItems(13).ForeColor = vbBlue
      End If
      If TabProduto.Fields("situacao").Value = "P" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(3).ForeColor = vbRed
         item.ListSubItems(4).ForeColor = vbRed
         item.ListSubItems(5).ForeColor = vbRed
         item.ListSubItems(6).ForeColor = vbRed
         item.ListSubItems(7).ForeColor = vbRed
         item.ListSubItems(8).ForeColor = vbRed
         item.ListSubItems(9).ForeColor = vbRed
         item.ListSubItems(10).ForeColor = vbRed
         item.ListSubItems(11).ForeColor = vbRed
         item.ListSubItems(12).ForeColor = vbRed
         item.ListSubItems(13).ForeColor = vbRed
      End If

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   Me.Enabled = True
   lstProduto.Visible = True

   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub VAI_TXT()
'On Error GoTo ERRO_TRATA

   Dim i                   As Integer
   Dim INDR_VAI_TXT   As Boolean

   INDR_PRI = True
   INDR_VAI_TXT = False

   If lstProduto.ListItems.Count > 0 Then
      For i = lstProduto.ListItems.Count To 1 Step -1
         If lstProduto.ListItems(i).Checked = True Then
            If Trim(lstProduto.SelectedItem.Text) <> "" Then

               If INDR_PRI = True Then
                  INDR_PRI = False
                  Msg = "Confirma Selecionado(s) ? "
                  Style = vbYesNo + 32
                  Title = "Atenção."
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                     INDR_VAI_TXT = True
                     Open PATH_TXT & "Itensmgv.txt" For Output As #1
                     Else
                        INDR_VAI_TXT = False
                        Exit Sub
                  End If
               End If
               If INDR_VAI_TXT = True Then
                  PRODUTO_ID_N = lstProduto.ListItems(i).SubItems(15)

                  GERA_ITENSMGV (PRODUTO_ID_N)
               End If
               txtQTDE.Caption = Trim(lstProduto.ListItems(i).Text)
               DoEvents
            End If
         End If
      Next i
      If INDR_PRI = False Then
         Close #1
         MsgBox "Processo realizado com sucesso !!!"
      End If
   End If

PRODUTO_ID_N = 0
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_TITULOS_SELECIONADOS"
End Sub

Sub CARREGA_FAMILIA_PRODUTO()
'On Error GoTo ERRO_TRATA

   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & "order by descricao "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFamilia.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("codg_familia").Value)
      cmbFamiliaAUX.AddItem Trim(TabDESCR.Fields("FAMILIAPRODUTO_ID").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_FAMILIA_PRODUTO"
End Sub
