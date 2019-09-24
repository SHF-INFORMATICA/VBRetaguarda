VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCONSULTAUSUARIO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Funcionários"
   ClientHeight    =   4335
   ClientLeft      =   3975
   ClientTop       =   2565
   ClientWidth     =   7395
   Icon            =   "consultausuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7395
   Begin VB.Frame Frame1 
      Caption         =   " Nome do Funcionário "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   7335
      Begin VB.TextBox txtCodg 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   7095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   120
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
            Picture         =   "consultausuario.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultausuario.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultausuario.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultausuario.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultausuario.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultausuario.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultausuario.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
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
            Key             =   "consultar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5760
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
               Picture         =   "consultausuario.frx":76BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultausuario.frx":8AE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultausuario.frx":9B75
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   2745
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   4842
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Códg."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCONSULTAUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call CentralizaJanela(frmCONSULTAUSUARIO)

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
End Sub

Private Sub Form_Resize()
'On Error GoTo ERRO_TRATA

   SQL = "select * from USUARIO WITH (NOLOCK)"
'SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by nome"

   GRID_USU

   MOSTRA_RODAPE "ESC - Sair", "", "", "", ""

   txtNome.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Resize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
      Case vbKeyF9
         txtNome.Text = ""
         txtNome.SetFocus
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

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
On Error Resume Next

   CRITERIO_A = LISTA.SelectedItem
   Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "consultar"
         CRITERIO_A = Chr$(39) & txtNome.Text & "%" & Chr(39)

         SQL = "select * from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where nome like " & CRITERIO_A
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N

         GRID_USU
      Case "limpar"
         txtNome.Text = ""
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub DBGrid1_DblClick()
   If txtCodg.Text <> "" Then _
      CRITERIO_A = txtCodg.Text
   Unload Me
End Sub

Private Sub txtCodg_GotFocus()
   txtNome.SetFocus
End Sub

Private Sub txtNome_Change()
   CRITERIO_A = Chr$(39) & txtNome.Text & "%" & Chr(39)

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where nome like " & CRITERIO_A
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N

   GRID_USU
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtCodg.Text <> "" Then _
         CRITERIO_A = txtCodg.Text
      Unload Me
   End If
End Sub

Private Sub GRID_USU()
   LISTA.ListItems.Clear
   If TabTemp.State = 1 Then TabTemp.Close
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = LISTA.ListItems.Add(, "seq." & TabTemp!USUARIO_ID, TabTemp!USUARIO_ID)
      item.SubItems(1) = TabTemp!NOME
      item.SubItems(2) = ""

      If Not IsNull(TabTemp.Fields("status").Value) Then
         If TabTemp.Fields("status").Value = 0 Then
            item.SubItems(2) = "Desativado"
            Else: item.SubItems(2) = "Ativo"
         End If
      End If

      TabTemp.MoveNext
   Wend
   TabTemp.Close
End Sub
