VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCHEQUEPAGTO 
   Caption         =   "Borderô Pagamento Cheque"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CHEQUEPAGTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   45
      TabIndex        =   0
      Top             =   3720
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Repasse"
      TabPicture(0)   =   "CHEQUEPAGTO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtCNPJCPF_REPASSE"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstRepasse"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRepasse"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtRepasse"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdVAI"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdVAI 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10560
         Picture         =   "CHEQUEPAGTO.frx":5C2E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtRepasse 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4200
         MaxLength       =   100
         TabIndex        =   6
         Top             =   480
         Width           =   6255
      End
      Begin VB.CommandButton cmdRepasse 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3720
         Picture         =   "CHEQUEPAGTO.frx":B230
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   405
      End
      Begin MSComctlLib.ListView lstRepasse 
         Height          =   2055
         Left            =   45
         TabIndex        =   3
         Top             =   960
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   3625
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N.Cheque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dt.Vencimento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Banco"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "cheque_id"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "repasse_id"
            Object.Width           =   176
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCNPJCPF_REPASSE 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   192
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repasse:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Incluir"
            Key             =   "incluir"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "imp"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7440
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUEPAGTO.frx":BC32
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUEPAGTO.frx":CDCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUEPAGTO.frx":DE5B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUEPAGTO.frx":EF66
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUEPAGTO.frx":FF1B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   8160
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11655
      DesignHeight    =   6960
   End
   Begin MSComctlLib.ListView lstCheque 
      Height          =   2775
      Left            =   45
      TabIndex        =   4
      Top             =   840
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   4895
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N.Cheque"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dt.Vencimento"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Banco"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Repasse"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "cnpjcpfter"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "cheque_id"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Repasse_id"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCHEQUEPAGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim PESSOA_REPASSE_ID_N    As Long
   Dim PESSOA_PORTADOR_ID_N   As Long

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "incluir"
      Case "imp"
      Case "voltar"
         Unload Me
      Case "limpar"
      Case "print"
      Case "consultar"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstcheque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstCheque, ColumnHeader
End Sub

Private Sub txtCNPJCPF_REPASSE_GotFocus()
   txtCNPJCPF_REPASSE.PromptInclude = False
      If Trim(txtCNPJCPF_REPASSE.Text) = "" Then _
         txtCNPJCPF_REPASSE.Text = "99999999999"
   txtCNPJCPF_REPASSE.PromptInclude = True
End Sub

Private Sub txtCNPJCPF_REPASSE_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF_REPASSE.PromptInclude = False
               txtCNPJCPF_REPASSE.Text = CNPJCPF_A
            txtCNPJCPF_REPASSE.PromptInclude = True
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_REPASSE_KeyDown"
End Sub

Private Sub txtCNPJCPF_REPASSE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCNPJCPF_REPASSE.PromptInclude = False

      If Trim(txtCNPJCPF_REPASSE.Text) <> "" Then _
         txtRepasse.Text = PROCURA_REPASSE(Trim(txtCNPJCPF_REPASSE.Text))

      txtCNPJCPF_REPASSE.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_REPASSE_KeyPress"
End Sub

Private Sub txtCNPJCPF_REPASSE_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_REPASSE.PromptInclude = False

   If Trim(txtCNPJCPF_REPASSE.Text) <> "" Then _
      txtRepasse.Text = PROCURA_REPASSE(Trim(txtCNPJCPF_REPASSE.Text))

   txtCNPJCPF_REPASSE.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrepasse_LostFocus"
End Sub

Private Sub cmdRepasse_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF_REPASSE.PromptInclude = False
      txtCNPJCPF_REPASSE.Text = CNPJCPF_A

      If Trim(txtCNPJCPF_REPASSE.Text) <> "" Then _
         txtRepasse.Text = PROCURA_REPASSE(Trim(txtCNPJCPF_REPASSE.Text))

   End If
   CNPJCPF_A = ""
   txtCNPJCPF_REPASSE.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCli_Click"
End Sub

Private Sub cmdVAI_Click()
'On Error GoTo ERRO_TRATA

   Dim INDR_VAI As Boolean
   Dim i
   INDR_VAI = False

   For i = lstCheque.ListItems.Count To 1 Step -1
      If lstCheque.ListItems(i).Checked = True Then
         Set item = lstRepasse.ListItems.Add(, "a" & lstCheque.ListItems(i).SubItems(7), Trim(lstCheque.ListItems(i).Text))   'N.Cheque
         item.SubItems(1) = "" & lstCheque.ListItems(i).SubItems(1)  'Valor
         item.SubItems(2) = "" & lstCheque.ListItems(i).SubItems(2)  'Dt.Vencimento
         item.SubItems(3) = "" & lstCheque.ListItems(i).SubItems(3)  'banco
         item.SubItems(4) = "" & lstCheque.ListItems(i).SubItems(7)  'cheque_id
         item.SubItems(5) = "" & lstCheque.ListItems(i).SubItems(8)  'repasse_ID
         item.Checked = lstCheque.ListItems(i).Checked
         INDR_VAI = True
      End If
   Next i

   If INDR_VAI = True Then
      Msg = "Cofirma repasse dos cheques para : " & Trim(txtRepasse.Text) & " ?"
      PERGUNTA Msg, vbYesNo + 32, "Desconto Pedido Venda", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         For i = lstRepasse.ListItems.Count To 1 Step -1
            If lstRepasse.ListItems(i).Checked = True Then

               SQL = "update CHEQUE set "
                  SQL = SQL & " repasse_id = " & PESSOA_REPASSE_ID_N
                  SQL = SQL & " , repasse = '" & Trim(txtRepasse.Text) & "'"
               SQL = SQL & " where cheque_id = " & lstRepasse.ListItems(i).SubItems(4)
               CONECTA_RETAGUARDA.Execute SQL

            End If
         Next i

         MsgBox "Processo realizado com sucesso."
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Function PROCURA_REPASSE(CNPJCPF_A As String) As String
'On Error GoTo ERRO_TRATA

   PROCURA_REPASSE = ""
   PESSOA_REPASSE_ID_N = 0

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select descricao,pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then
      PROCURA_REPASSE = "" & Trim(TabPessoa.Fields("descricao").Value)
      PESSOA_REPASSE_ID_N = TabPessoa.Fields("pessoa_id").Value
      Else
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         MsgBox "CNPJ/CPF não encontrado"
         Exit Function
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_REPASSE"
End Function
