VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDISPLAYFUNC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Funcionário"
   ClientHeight    =   7020
   ClientLeft      =   1950
   ClientTop       =   2565
   ClientWidth     =   11025
   Icon            =   "frmDISPLAYFUNC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11025
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
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
      Height          =   330
      Left            =   5280
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   1270
      ButtonWidth     =   2646
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
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5160
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
               Picture         =   "frmDISPLAYFUNC.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":7E3B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3090
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":8F46
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":939A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":96B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":9B0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":9F5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":A27E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":A6D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":A9F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":B404
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDISPLAYFUNC.frx":BE16
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkC 
         Caption         =   "&Cancelados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9570
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   5625
      Left            =   60
      TabIndex        =   4
      Top             =   1320
      Width           =   10890
      _ExtentX        =   19209
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
         Object.Width           =   2196
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Cadastro"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Situação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Código"
         Object.Width           =   1561
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCgcCpf 
      Height          =   345
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2100
      TabIndex        =   8
      Top             =   870
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   615
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4680
      TabIndex        =   5
      Top             =   870
      Width           =   510
   End
End
Attribute VB_Name = "frmDISPLAYFUNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         txtCodigo.Text = ""
         intCodigo = 0
         txtNome.Text = ""
         txtCPF.Text = ""
         txtCodigo.SetFocus
   End Select
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   Call CentralizaJanela(frmDISPLAYFUNC)
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_load"
End Sub

Private Sub chkC_Click()
'On Error GoTo ERRO_TRATA

   If chkC.Value = 1 Then
      CRITERIO = "0"
      Else: CRITERIO = "1"
   End If

   SQL = "SELECT * FROM FUNCIONARIOCONVENIO "
   SQL = SQL & " where status = '" & CRITERIO & "'"

   If intCodigo > 0 Then _
      SQL = SQL & " and  Codigo_Cliente = " & intCodigo

   SQL = SQL & " order by nome asc"

   SETA_FUNCIONARIOCONVENIO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "chkC_Click"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "todos"
         SQL = "SELECT * FROM FUNCIONARIOCONVENIO "
         If chkC.Value = 1 Then
            CRITERIO = "0"
         Else: CRITERIO = "1"
         End If
         SQL = SQL & " where status='" & CRITERIO & "'"
         If intCodigo > 0 Then
            SQL = SQL & " and  Codigo_Cliente = " & intCodigo
         End If
         SQL = SQL & " order by nome asc"
         SETA_FUNCIONARIOCONVENIO
      Case "limpar"
         txtNome.Text = ""
         intCodigo = 0
         txtCGCCPF.PromptInclude = False
         txtCGCCPF.Text = ""
         txtCGCCPF.PromptInclude = True
         txtCodigo.Text = ""
         LISTA.ListItems.Clear
         chkC.Value = 0
      Case "voltar"
         Unload Me
   End Select
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
'On Error GoTo ERRO_TRATA
   'CNPJCPF_a = LISTA.SelectedItem.ListSubItems.Item(4)
   intCodigo = LISTA.SelectedItem.ListSubItems.Item(4)
   Unload Me
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTA_DblClick"
End Sub

Private Sub txtCpf_Click()
'On Error GoTo ERRO_TRATA
   txtNome.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcpf_click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Not IsNull(DBGrid1.Columns(0).Text) Then _
         CNPJCPF_A = DBGrid1.Columns(0).Text
      If Not IsNull(DBGrid1.Columns(1).Text) Then _
         NOME_A = DBGrid1.Columns(1).Text

      'If CONECTA_RETAGUARDA.State = 1 Then _
         CONECTA_RETAGUARDA.Close

      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer) 'ver rotina com sergio
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCodigo.Text <> "" Then
         CRITERIO = Chr$(39) & txtCodigo.Text & "%" & Chr(39)
         SQL = "select * from FUNCIONARIOCONVENIO F "
         SQL = SQL & " where F.Codigo = " & txtCodigo.Text
         If intCodigo > 0 Then
            SQL = SQL & " and  Codigo_Cliente = " & intCodigo
         End If
         SQL = SQL & " order by F.Nome "
         Else
             CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)
             SQL = "SELECT * FROM FUNCIONARIOCONVENIO "
             SQL = SQL & " where nome LIKE " & CRITERIO
             If chkC.Value = 1 Then
                SQL = SQL & " and status='0' "
             Else
                SQL = SQL & " and status='1' "
             End If
             If intCodigo > 0 Then
                SQL = SQL & " and  Codigo_Cliente = " & intCodigo
             End If
             SQL = SQL & " order by nome asc"
      End If
      SETA_FUNCIONARIOCONVENIO
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodigo_KeyPress"
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCGCCPF.Text <> "" Then
         CRITERIO = Chr$(39) & txtCGCCPF.Text & "%" & Chr(39)
         SQL = "select * from FUNCIONARIOCONVENIO F "
         SQL = SQL & " where F.CPF = '" & txtCGCCPF.Text & "'"
         If intCodigo > 0 Then
            SQL = SQL & " and  Codigo_Cliente = " & intCodigo
         End If
         SQL = SQL & " order by F.Nome "
         Else
             CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)
             SQL = "SELECT * FROM FUNCIONARIOCONVENIO "
             SQL = SQL & " where nome LIKE " & CRITERIO
             If chkC.Value = 1 Then
                SQL = SQL & " and status='0' "
             Else
                SQL = SQL & " and status='1' "
             End If
             If intCodigo > 0 Then
                SQL = SQL & " and  Codigo_Cliente = " & intCodigo
             End If
             SQL = SQL & " order by nome asc"
      End If
      SETA_FUNCIONARIOCONVENIO
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCgcCpf_KeyPress"
End Sub


Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)
      SQL = "SELECT * FROM FUNCIONARIOCONVENIO "
      SQL = SQL & " where nome LIKE " & CRITERIO
      If chkC.Value = 1 Then
         SQL = SQL & " and status='0' "
      Else
         SQL = SQL & " and status='1' "
      End If
      If intCodigo > 0 Then
         SQL = SQL & " and  Codigo_Cliente = " & intCodigo
      End If
      SQL = SQL & " order by nome asc"
      SETA_FUNCIONARIOCONVENIO
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub

Private Sub txtNome_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA
   Select Case KeyCode
      Case vbKeyUp: txtCodigo.SetFocus
   End Select
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyUp"
End Sub

Private Sub SETA_FUNCIONARIOCONVENIO()
'On Error GoTo ERRO_TRATA
   LISTA.ListItems.Clear
   CONT_N = 0
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CONT_N = CONT_N + 1
      If Not IsNull(TabTemp!CPF) Then
        Set Item = LISTA.ListItems.Add(, "seq." & CONT_N, Trim(TabTemp!CPF))
        Item.SubItems(1) = TabTemp!NOME
        Item.SubItems(2) = TabTemp!DT_CAD
        Item.SubItems(3) = "" & TabTemp!Status
        Item.SubItems(4) = "" & TabTemp!codigo
      End If
      TabTemp.MoveNext
   Wend
   TabTemp.Close


   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Duplo click selecionar"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_FUNCIONARIOCONVENIO"
End Sub


