VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSVeiculoConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Veículo"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "OSVeiculoConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPLACA 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox cmbAuxTipo 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCHASSI 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtNome 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox txtANO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtMODELO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox cmbTIPO 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
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
            Picture         =   "OSVeiculoConsulta.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVeiculoConsulta.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVeiculoConsulta.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVeiculoConsulta.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVeiculoConsulta.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVeiculoConsulta.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVeiculoConsulta.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView lstVeiculo 
      Height          =   3825
      Left            =   0
      TabIndex        =   13
      Top             =   2280
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   6747
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777152
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Placa"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Chassi"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ANO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "MODELO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "TIPO"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblCpf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   12
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Veículo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1800
      Width           =   1320
   End
End
Attribute VB_Name = "frmOSVeiculoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
   LIMPA_TELA
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "limpar"
         LIMPA_TELA
      Case "consultar"
         SETA_GRID_VEICULO
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID_VEICULO
   End If
End Sub

Private Sub LSTVEICULO_DblClick()
   If Not IsNull(lstVeiculo.SelectedItem.Text) Then
      SQL3 = lstVeiculo.SelectedItem.Text
      Unload Me
   End If
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID_VEICULO
   End If
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      'SETA_GRID_VEICULO
   End If
End Sub

Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID_VEICULO
   End If
End Sub

Private Sub txtCNPJCPF_LostFocus()
   txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) <> "" Then _
         txtNome.Text = "" & TRAZ_NOME_PESSOA(0, Trim(txtCNPJCPF.Text))
   txtCNPJCPF.PromptInclude = True
End Sub

Private Sub txtMODELO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID_VEICULO
   End If
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID_VEICULO
   End If

End Sub

Private Sub SETA_GRID_VEICULO()
'On Error GoTo ERRO_TRATA

   Dim Marcador_A          As Long
   Dim TIPO_VEICULO_ID_N   As Integer

   lstVeiculo.ListItems.Clear
   Marcador_A = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT * from vwVEICULO WITH (NOLOCK)"
   SQL = SQL & " where veiculo_id > 0"

   If Trim(txtPlaca.Text) <> "" Then _
      SQL = SQL & " and placa like '" & Replace(Trim(txtPlaca.Text), "-", "") & "%'"

   If Trim(txtCHASSI.Text) <> "" Then _
      SQL = SQL & "  and chassi like '" & Trim(txtCHASSI.Text) & "%'"

   If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N

   If Trim(txtANO.Text) <> "" Then _
      SQL = SQL & " and ano like '" & Trim(txtANO.Text) & "%'"

   If Trim(txtMODELO.Text) <> "" Then _
      SQL = SQL & " and modelo like '" & Trim(txtMODELO.Text) & "%'"

   If Trim(cmbAuxTipo.Text) <> "" Then _
      SQL = SQL & " and tipo_eqp like '" & Trim(cmbAuxTipo.Text) & "%'"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

'Debug.Print SQL

   While Not TabTemp.EOF
      NOME_A = "" & Trim(TabTemp.Fields("descpessoa").Value)
      TIPO_VEICULO_ID_N = 0 & TabTemp!tipo_veiculo_id
      If Trim(NOME_A) = "" Then _
         NOME_A = "" & TRAZ_NOME_PESSOA(TabTemp.Fields("pessoa_id").Value, "")

      Set item = lstVeiculo.ListItems.Add(, "seq." & TabTemp.Fields("veiculo_id").Value, TabTemp!PLACA)

      item.SubItems(1) = "" & TabTemp!CHASSI
      item.SubItems(2) = "" & Trim(NOME_A)
      item.SubItems(3) = "" & TabTemp!Ano
      item.SubItems(4) = "" & TabTemp!MODELO
      item.SubItems(5) = "" & TRAZ_DESCRITOR("A", Str(TIPO_VEICULO_ID_N))

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Sub SETA_GRID_VEICULO"
End Sub

Sub LIMPA_TELA()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   txtPlaca.Text = ""
   txtCHASSI.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   cmbTIPO.Text = ""

   SETA_GRID_VEICULO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TELA"
End Sub
