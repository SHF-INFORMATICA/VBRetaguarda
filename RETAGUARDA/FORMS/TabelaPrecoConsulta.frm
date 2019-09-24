VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTabelaPrecoConsulta 
   Caption         =   "Consulta Tabela Preço Cadastrada"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "TabelaPrecoConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
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
      Left            =   3480
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbFamilia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtDescricao 
      DataSource      =   "DataCep"
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
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   3
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
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPrecoConsulta.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPrecoConsulta.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPrecoConsulta.frx":80C9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstTabela 
      Height          =   4545
      Left            =   45
      TabIndex        =   3
      Top             =   1560
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   8017
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   3263
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   9172
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Validade"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSMask.MaskEdBox txtValidade 
      Height          =   360
      Left            =   7500
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valida:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7560
      TabIndex        =   8
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Família de Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3480
      TabIndex        =   6
      Top             =   840
      Width           =   2040
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1125
   End
End
Attribute VB_Name = "frmTabelaPrecoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

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
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9: LIMPA_TUDO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstTabela_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   OrdenaListView lstTabela, ColumnHeader
End Sub

Private Sub lstTabela_DblClick()
'On Error GoTo ERRO_TRATA

   If Trim(lstTabela.SelectedItem.Text) <> "" Then
      CRITERIO_A = lstTabela.SelectedItem.Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstTabela_DblClick"
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescricao_KeyPress"
End Sub

Private Sub txtvALIDADE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtvALIDADE_KeyPress"
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex
   SETA_GRID

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtDescricao.Text = ""
   cmbFamilia.Text = ""
   cmbFamiliaAUX.Text = ""
   txtValidade.PromptInclude = False
   txtValidade.Text = ""
   lstTabela.ListItems.Clear
   txtDescricao.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub
   
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstTabela.ListItems.Clear
   CONT_N = 0
   SqL2 = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select TABELAPRECO.CODG_TABELA, TABELAPRECO.DESCRICAO, "
   SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.TABELAPRECOITEM_ID, "
   SQL = SQL & " TABELAPRECOITEM.TABELAPRECO_ID, TABELAPRECOITEM.PRODUTO_ID, "
   SQL = SQL & " TABELAPRECOITEM.FORMAPAGTO_ID, TABELAPRECOITEM.VALOR_VENDA,"
   SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO, TABELAPRECOITEM.PERC_COMISSAO, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO,PRODUTO.DESCRICAO AS DescProduto,Produto.FAMILIAPRODUTO_ID"
   SQL = SQL & " from TABELAPRECO "
   SQL = SQL & " INNER JOIN TABELAPRECOITEM "
   SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where TABELAPRECO.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(txtDescricao.Text) <> "" Then
      CRITERIO_A = Chr$(39) & Trim(txtDescricao.Text) & "%" & Chr(39)
      SQL = SQL & " and descricao like " & CRITERIO_A
   End If

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      SQL = SQL & " and Produto.FAMILIAPRODUTO_ID = " & Trim(cmbFamiliaAUX.Text)

   txtValidade.PromptInclude = False
   If Trim(txtValidade.Text) <> "" Then
      txtValidade.PromptInclude = True
      SQL = SQL & " and dt_validade <= '" & DMA(txtValidade.Text) & "'"
   End If
   txtValidade.PromptInclude = True

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If Trim(SqL2) <> Trim(TabTemp.Fields("codg_tabela").Value) Then
         Set item = lstTabela.ListItems.Add(, "seq." & CONT_N, Trim(TabTemp.Fields("codg_tabela").Value))
         item.SubItems(1) = Trim(TabTemp.Fields("descricao").Value)
         item.SubItems(2) = Trim(TabTemp.Fields("dt_validade").Value)
         SqL2 = Trim(TabTemp.Fields("codg_tabela").Value)
         CONT_N = CONT_N + 1
      End If
      TabTemp.MoveNext
   Wend
   CONT_N = 0
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

