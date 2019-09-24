VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDISPLAYFORNECEDOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Fornecedores"
   ClientHeight    =   6960
   ClientLeft      =   1950
   ClientTop       =   3120
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DISPLAYFORNECEDOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10995
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFone 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
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
      Left            =   4440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DISPLAYFORNECEDOR.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
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
            Key             =   "todos"
            Object.ToolTipText     =   "Consultar Todos"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6840
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
               Picture         =   "DISPLAYFORNECEDOR.frx":76BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DISPLAYFORNECEDOR.frx":8858
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DISPLAYFORNECEDOR.frx":98E7
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkC 
         Caption         =   "&Cancelados"
         Height          =   195
         Left            =   8280
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   5625
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   9922
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CNPJ/CPF"
         Object.Width           =   3379
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   6006
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Razão Social"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Cadastro"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone:"
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
      Left            =   8040
      TabIndex        =   5
      Top             =   870
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "CGC/CPF:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   7245
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
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
      Left            =   150
      TabIndex        =   3
      Top             =   870
      Width           =   675
   End
End
Attribute VB_Name = "frmDISPLAYFORNECEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

   If KeyAscii = 13 Then
      If Not IsNull(DBGrid1.Columns(0).Text) Then _
         CNPJCPF_A = DBGrid1.Columns(0).Text
      If Not IsNull(DBGrid1.Columns(1).Text) Then _
         NOME_A = DBGrid1.Columns(1).Text

      'If CONECTA_RETAGUARDA.State = 1 Then _
         CONECTA_RETAGUARDA.Close

      Unload Me
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         txtFone.Text = ""
         txtNome.Text = ""
         txtCPF.Text = ""
         txtFone.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub chkC_Click()
'On Error GoTo ERRO_TRATA

   CRITERIO = "A"

   If chkC.Value = 1 Then
      CRITERIO = "C"
      Else: CRITERIO = "A"
   End If

   SQL = "SELECT * from vwFornecedor "
   SQL = SQL & " where status = '" & CRITERIO & "'"
   SQL = SQL & " order by descricao asc"

   SETA_FORNECEDOR

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "chkC_Click"
End Sub

Private Sub optJ_Click()
'On Error GoTo ERRO_TRATA

   SQL = "SELECT * from vwFornecedor "
   SQL = SQL & " where cnpjcpf <> '' "

   If chkC.Value = 1 Then
      SQL = SQL & " and status='C' "
      Else: SQL = SQL & " and status='A' "
   End If

   SQL = SQL & " order by descricao asc"
   SETA_FORNECEDOR
   txtNome.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optJ_Click"
End Sub

Private Sub optN_Click()

   SQL = "SELECT * from vwFornecedor "
   SQL = SQL & " where descricao <> '' "
   If chkC.Value = 1 Then
      SQL = SQL & " and status='C' "
      Else: SQL = SQL & " and status='A' "
   End If
   SQL = SQL & " order by descricao asc"
   SETA_FORNECEDOR
   txtNome.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.key
      Case "todos"
         SQL = "SELECT * from vwFornecedor "
         SQL = SQL & " order by descricao asc"
         SETA_FORNECEDOR
      Case "limpar"
         txtNome.Text = ""
         txtCPF.Text = ""
         txtFone.Text = ""
         LISTA.ListItems.Clear
         chkC.Value = 0
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next

  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
On Error Resume Next

   'NOME_A = LISTA.SelectedItem.Text
   CNPJCPF_A = LISTA.SelectedItem.Text
   Unload Me
End Sub

Private Sub txtCpf_Click()
   txtNome.SetFocus
End Sub

Private Sub txtfone_KeyPress(KeyAscii As Integer) 'ver rotina com sergio

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtFone.Text <> "" Then
         CRITERIO = Chr$(39) & txtFone.Text & "%" & Chr(39)
         SQL = "select * from FONE "
         SQL = SQL & " where numero like " & CRITERIO
         SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " order by descricao"
         Else
            CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)
            SQL = "SELECT * from vwFornecedor "
            SQL = SQL & " where descricao LIKE " & CRITERIO
            If chkC.Value = 1 Then _
               SQL = SQL & " and status='C' "
            SQL = SQL & " order by descricao asc"
      End If
      SETA_FORNECEDOR
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)

      SQL3 = "A"
      If chkC.Value = 1 Then _
         SQL3 = "C"

      SQL = "SELECT * from vwFornecedor "
      SQL = SQL & " where descricao LIKE " & CRITERIO
      SQL = SQL & " and status ='" & SQL3 & "'"
      SQL = SQL & " order by descricao asc"

      SETA_FORNECEDOR
   End If
End Sub

Private Sub txtNome_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyUp: txtFone.SetFocus
   End Select
End Sub

Private Sub SETA_FORNECEDOR()

   LISTA.ListItems.Clear
   CONT_N = 0
   If TabTemp.State = 1 Then _
      TabTemp.Close
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CONT_N = CONT_N + 1
      If Not IsNull(TabTemp.Fields("cnpjcpf").Value) Then
        Set Item = LISTA.ListItems.Add(, "seq." & CONT_N, Trim(TabTemp.Fields("cnpjcpf").Value))
        Item.SubItems(1) = TabTemp!DESCRICAO
        Item.SubItems(2) = TabTemp!RAZAO
        Item.SubItems(3) = TabTemp!DT_CAD
        Item.SubItems(4) = "" & TabTemp!Status
      End If
      TabTemp.MoveNext
   Wend
   TabTemp.Close
End Sub
