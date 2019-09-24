VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCONSULTAVENDEDOR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Vendedores"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   885
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "consultavendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox txtFone 
         Height          =   360
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   360
         ItemData        =   "consultavendedor.frx":5C12
         Left            =   7680
         List            =   "consultavendedor.frx":5C1C
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbEquipeAUX 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   1080
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbEquipe 
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox cmbUF 
         Height          =   360
         Left            =   7680
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cmbCidade 
         Height          =   360
         Left            =   1080
         TabIndex        =   6
         Top             =   1680
         Width           =   5415
      End
      Begin VB.TextBox txtNome 
         Height          =   360
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   0
         Top             =   240
         Width           =   5415
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   360
         Left            =   7680
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCPF 
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         Height          =   240
         Left            =   7035
         TabIndex        =   17
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   240
         Left            =   525
         TabIndex        =   16
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6675
         TabIndex        =   15
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equipe:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   255
         TabIndex        =   13
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   240
         Left            =   7140
         TabIndex        =   12
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         Height          =   240
         Left            =   7140
         TabIndex        =   10
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblNome 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultavendedor.frx":5C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultavendedor.frx":6084
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultavendedor.frx":63A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultavendedor.frx":67F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultavendedor.frx":6C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultavendedor.frx":709C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "sair"
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
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkSituacao 
         Caption         =   "Ativos"
         Height          =   240
         Left            =   3600
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   4185
      Left            =   0
      TabIndex        =   19
      Top             =   720
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   7382
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777215
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
         Text            =   "Códg."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Equipe"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dt.Início"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dt.Desligamento"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Situação"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmCONSULTAVENDEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   SETA_GRID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub cmbCidade_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Selecione uma cidade", "", "", ""
End Sub

Private Sub cmbEquipe_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Selecione um representante", "", "", ""
End Sub

Private Sub cmbUF_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Selecione um estado", "", "", ""
End Sub

Private Sub GRIDRESP_DblClick()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = GRIDRESP.Columns(0).Value
   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRIDRESP_DblClick"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
On Error Resume Next

   CRITERIO_A = ""
   CRITERIO_A = LISTA.SelectedItem.Text
   Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         SETA_GRID
      Case "print"
      Case "limpar"
         LIMPA_VEND
      Case "sair"
         CRITERIO_A = ""
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCep_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe CEP", "", "", ""
End Sub

Private Sub txtFone_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe número do telefone", "", "", ""
End Sub

Private Sub txtcpf_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe número do CPF", "", "", ""
   
   txtCPF.PromptInclude = False
   If txtCPF.Text = "" Then _
      txtCPF.Mask = "##############"
   txtCPF.PromptInclude = True
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii >= 32 And KeyAscii <= 47 Or _
            KeyAscii >= 58 And KeyAscii <= 64 Or _
            KeyAscii >= 91 And KeyAscii <= 96 Or _
            KeyAscii >= 123 And KeyAscii <= 127 Then KeyAscii = 0
   End If
End Sub

Private Sub txtNome_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe o nome", "", "", ""
End Sub

Private Sub cmbEquipe_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxEquipe.ListIndex = cmbEquipe.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbEquipe_Click"
End Sub

Private Sub cmbUF_Click()
'On Error GoTo ERRO_TRATA

   If cmbUF.Text <> "" Then
      If cmbCidade.Enabled = True Then
         SQL = "select distinct(cidade) from CEP "
         SQL = SQL & " where uf='" & cmbUF.Text & "'"
         SQL = SQL & " order by cidade"

         If TabTemp.State = 1 Then _
            TabTemp.Close
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

         While Not TabTemp.EOF
            cmbCidade.AddItem TabTemp!CIDADE
            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbUF_Click"
End Sub

Private Sub MONTA_UF()
'On Error GoTo ERRO_TRATA

   cmbUF.Clear
   SQL = "select distinct(uf) from CEP "
   SQL = SQL & "order by uf"

   If TabTemp.State = 1 Then _
      TabTemp.Close
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbUF.AddItem TabTemp!UF
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_UF"
End Sub

Private Sub LIMPA_VEND()
'On Error GoTo ERRO_TRATA

   GRIDRESP.Enabled = False
   txtNome.Enabled = False
   cmbEquipe.Enabled = False
   txtFone.Enabled = False
   txtCep.Enabled = False
   cmbCidade.Enabled = False
   cmbUF.Enabled = False
   txtCPF.Enabled = False
   txtFone.Enabled = False
   txtNome.Text = ""
   cmbAuxEquipe.Text = ""
   cmbEquipe.Text = ""
   txtCep.Text = ""
   cmbSTATUS.Text = ""
   txtFone.Text = ""
   txtCPF.PromptInclude = False
   txtCPF.Text = ""
   cmbCidade.Text = ""
   cmbCidade.Clear
   cmbUF.Text = ""
   cmbUF.Clear
   CRITERIO_A = ""

   SQL = "select empresa.* from EMPRESA "

   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and razao_social = ''"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_VEND"
End Sub

Private Sub CHECAR_CPF()
'On Error GoTo ERRO_TRATA

   txtCPF.PromptInclude = False
   If txtCPF.Text = "" Then
      MsgBox "CNPJ/CPF com DV incorreto !!! "
      txtCPF = ""
      txtCPF.SetFocus
      Exit Sub
   End If
   If Len(txtCPF.Text) > 0 Then
      Select Case Len(txtCPF.Text)
         Case Is = 11
           If Not CALCULACPF(txtCPF.Text) Then
              MsgBox "CPF com DV incorreto !!!"
              txtCPF.PromptInclude = False
              txtCPF = ""
              txtCPF.SetFocus
              Exit Sub
           End If
         Case Is = 14
           If Not VALIDACGC(txtCPF.Text) Then
              MsgBox "CNPJ com DV incorreto !!! "
              txtCPF.PromptInclude = False
              txtCPF = ""
              txtCPF.SetFocus
              Exit Sub
           End If
         Case Is > 14
            MsgBox "CNPJ/CPF com DV incorreto !!! "
            txtCPF = ""
            txtCPF.SetFocus
            Exit Sub
         Case Is < 11
            MsgBox "CNPJ/CPF com DV incorreto !!! "
            txtCPF = ""
            txtCPF.SetFocus
            Exit Sub
      End Select
      Else
         MsgBox "CNPJ/CPF com DV incorreto !!! "
         txtCPF = ""
         txtCPF.SetFocus
         Exit Sub
   End If
   txtCPF.PromptInclude = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECAR_CPF"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   LISTA.ListItems.Clear
   txtCPF.PromptInclude = False
   CONT_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwVendedor "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If chkSituacao.Value = 1 Then
      SQL = SQL & " and status = 'A'"
      Else: SQL = SQL & " and status = 'C'"
   End If

   If Trim(txtCPF.Text) <> "" Then _
      SQL = SQL & " and cnpjcpf = '" & Trim(txtCPF.Text) & "'"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If CONT_N <> TabTemp.Fields("vendedor_id").Value Then
         CONT_N = TabTemp.Fields("vendedor_id").Value
         NUMR_SEQ_N = NUMR_SEQ_N + 1
         Set item = LISTA.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp!VENDEDOR_ID)
         item.SubItems(1) = Trim(TabTemp.Fields("descricao").Value)
         If Not IsNull(TabTemp!DESCRICAO) Then _
            item.SubItems(2) = TabTemp!DESCRICAO

         CRITERIO_A = ""
         If Not IsNull(TabTemp!DATA_CAD) Then _
            CRITERIO_A = TabTemp!DATA_CAD
         item.SubItems(3) = CRITERIO_A

         CRITERIO_A = ""
         If Not IsNull(TabTemp!DT_BAIXA) Then _
            CRITERIO_A = TabTemp!DT_BAIXA
         item.SubItems(4) = CRITERIO_A

         CRITERIO_A = ""
         If Not IsNull(TabTemp!STATUS) Then
            If TabTemp!STATUS = "A" Then
               CRITERIO_A = "Ativo"
               Else: CRITERIO_A = "Desativado"
            End If
         End If
         item.SubItems(5) = CRITERIO_A
      End If
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
