VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDISPLAYTRANSPORTADORA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Cadastro de Transportadora"
   ClientHeight    =   6930
   ClientLeft      =   1950
   ClientTop       =   2235
   ClientWidth     =   10965
   Icon            =   "displaytransportadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10965
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox txtFone 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      MaxLength       =   10
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
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
            Picture         =   "displaytransportadora.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "displaytransportadora.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "displaytransportadora.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "displaytransportadora.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "displaytransportadora.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "displaytransportadora.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "displaytransportadora.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCPF 
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
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
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
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4200
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
               Picture         =   "displaytransportadora.frx":76BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "displaytransportadora.frx":8858
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "displaytransportadora.frx":98E7
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkC 
         Caption         =   "&Cancelados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8400
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   5625
      Left            =   0
      TabIndex        =   5
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
         Name            =   "Calibri"
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
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2070
      TabIndex        =   8
      Top             =   900
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "CGC/CPF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7095
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone:"
      BeginProperty Font 
         Name            =   "Calibri"
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
      Top             =   900
      Width           =   435
   End
End
Attribute VB_Name = "frmDISPLAYTRANSPORTADORA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         txtFone.Text = ""
         txtNome.Text = ""
         txtCPF.Text = ""
         txtFone.SetFocus
   End Select
End Sub

Private Sub Form_Load()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   Call CentralizaJanela(frmDISPLAYTRANSPORTADORA)
End Sub

Private Sub chkC_Click()
   If chkC.Value = 1 Then
      CRITERIO = "C"
      Else: CRITERIO = "A"
   End If
   SQL = "SELECT * from TRANSPORTADORA "
   SQL = SQL & " where status='" & CRITERIO & "'"
   SQL = SQL & " order by nome asc"
   SETA_TRANSPORTADORA
End Sub

Private Sub optJ_Click()
   SQL = "SELECT * from TRANSPORTADORA "
   SQL = SQL & " where cgccpf <> '' "
   If chkC.Value = 1 Then
      SQL = SQL & " and status='C' "
      Else: SQL = SQL & " and status='A' "
   End If
   SQL = SQL & " order by nome asc"
   SETA_TRANSPORTADORA
   txtNome.SetFocus
End Sub

Private Sub optN_Click()
   SQL = "SELECT * from TRANSPORTADORA "
   SQL = SQL & " where nome <> '' "
   If chkC.Value = 1 Then
      SQL = SQL & " and status='C' "
      Else: SQL = SQL & " and status='A' "
   End If
   SQL = SQL & " order by nome asc"
   SETA_TRANSPORTADORA
   txtNome.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "todos"
         SQL = "SELECT * from TRANSPORTADORA "
         SQL = SQL & " order by nome asc"
         SETA_TRANSPORTADORA
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

Private Sub Form_Unload(Cancel As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(DBGrid1.Columns(0).Text) Then _
         CNPJCPF_A = DBGrid1.Columns(0).Text
      If Not IsNull(DBGrid1.Columns(1).Text) Then _
         NOME_A = DBGrid1.Columns(1).Text
      Unload Me
   End If

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub txtfone_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtFone.Text <> "" Then
         CRITERIO = Chr$(39) & txtFone.Text & "%" & Chr(39)
         SQL = "select * from TRANSPORTADORA c left join FONE f "
         SQL = SQL & " on c.CGCCPF = f.prop "
         SQL = SQL & " where f.numero like " & CRITERIO
         SQL = SQL & " order by c.nome "
         Else
             CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)
             SQL = "SELECT * from TRANSPORTADORA "
             SQL = SQL & " where nome LIKE " & CRITERIO
             If chkC.Value = 1 Then _
                SQL = SQL & " and status='C' "
                SQL = SQL & " order by nome asc"
      End If
      SETA_TRANSPORTADORA
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      CRITERIO = Chr$(39) & txtNome.Text & "%" & Chr(39)
      SQL = "SELECT * from TRANSPORTADORA "
      SQL = SQL & " where nome LIKE " & CRITERIO
      If chkC.Value = 1 Then _
         SQL = SQL & " and status='C' "
      SQL = SQL & " order by nome asc"
      SETA_TRANSPORTADORA
   End If
End Sub

Private Sub txtNome_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyUp: txtFone.SetFocus
   End Select
End Sub

Private Sub SETA_TRANSPORTADORA()
   LISTA.ListItems.Clear
   CONT_N = 0
   If TabTemp.State = 1 Then TabTemp.Close
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CONT_N = CONT_N + 1
      Set Item = LISTA.ListItems.Add(, "seq." & CONT_N, Trim(TabTemp!CGCCPF))
      Item.SubItems(1) = TabTemp!NOME
      Item.SubItems(2) = TabTemp!razao_social
      Item.SubItems(3) = TabTemp!DT_CAD
      Item.SubItems(4) = "" & TabTemp!Status
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
End Sub


