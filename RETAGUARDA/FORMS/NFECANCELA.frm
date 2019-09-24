VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNFECANCELA 
   Caption         =   "Cancela Entrada de Produto Estoque"
   ClientHeight    =   6360
   ClientLeft      =   2085
   ClientTop       =   2355
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NFECANCELA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   10980
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   10935
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtSERIE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5880
         TabIndex        =   8
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox txtTotalNota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   9360
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtDtEntrada 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtDTEMISSAO 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Série: "
         Height          =   225
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nota: "
         Height          =   225
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pedido: "
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor: "
         Height          =   225
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor Nota Fiscal:"
         Height          =   225
         Left            =   7800
         TabIndex        =   11
         Top             =   1020
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data Entrada:"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão:"
         Height          =   225
         Left            =   3960
         TabIndex        =   9
         Top             =   1020
         Width           =   1170
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
            Picture         =   "NFECANCELA.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECANCELA.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECANCELA.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECANCELA.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECANCELA.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECANCELA.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECANCELA.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "gravar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Con&sultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECANCELA.frx":76BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECANCELA.frx":8858
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECANCELA.frx":98E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECANCELA.frx":AFE4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTAITENS 
      Height          =   2025
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   3572
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   12648447
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qtd."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr. Item"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valr.Total "
         Object.Width           =   2645
      EndProperty
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   2025
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   3572
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   16777152
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Nota"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Série"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Fornecedor"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "fornec_id"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10920
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmNFECANCELA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmNFECANCELA)
   txtTotalNota.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTotalNota.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF9
         LIMPA_NOTA_ENTRADA
         txtNOTA.SetFocus
      Case vbKeyF10
         GRAVA_CABECA_NOTA
         LIMPA_NOTA_ENTRADA
         txtNOTA.SetFocus
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub LISTA_Click()
'On Error GoTo ERRO_TRATA

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   If Not IsNull(LISTA.SelectedItem.Text) Then
      SQL = "select * from NOTAENTRADA "
      SQL = SQL & " where numr_nota = " & LISTA.SelectedItem.ListSubItems.item(1).Text
      SQL = SQL & " and serie_nota = '" & Trim(LISTA.SelectedItem.ListSubItems.item(2).Text) & "'"
      SQL = SQL & " and estabelecimento_id = " & EMPRESA_ID_N
      SQL = SQL & " and fornecedor_id = " & LISTA.SelectedItem.ListSubItems.item(6).Text
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then _
         SETA_GRID_ITENS
   End If

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTA_Click"
End Sub

Private Sub listaitens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTAITENS, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         LISTA.ListItems.Clear
         LISTAITENS.ListItems.Clear

         MOSTRAR_NOTA
      Case "print"
         FORMULA_REL = ""

      If Trim(LISTA.SelectedItem.ListSubItems.item(1).Text) <> "" And _
         Trim(LISTA.SelectedItem.ListSubItems.item(2).Text) <> "" And _
         LISTA.SelectedItem.ListSubItems.item(6).Text <> "" Then

         FORMULA_REL = "{NOTAENTRADA.numr_nota} = " & Trim(Trim(LISTA.SelectedItem.ListSubItems.item(1).Text))
         FORMULA_REL = FORMULA_REL & " {NOTAENTRADA.serie_nota} = '" & Trim(Trim(LISTA.SelectedItem.ListSubItems.item(2).Text)) & "'"
         FORMULA_REL = FORMULA_REL & " {NOTAENTRADA.estabelecimento_ID} = " & EMPRESA_ID_N
         FORMULA_REL = FORMULA_REL & " {NOTAENTRADA.fornecedor_id} = " & Trim(LISTA.SelectedItem.ListSubItems.item(6).Text)
         Else
            If Trim(txtPedido.Text) <> "" Then _
               If IsNumeric(txtPedido.Text) Then _
                  FORMULA_REL = "{NOTAENTRADA.pedidocompra_id} = " & Trim(txtPedido.Text)
      End If

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "nf_entra.rpt"
         frmRELATORIO10.Show 1
      Case "voltar"
         Unload Me
      Case "gravar"
         If LISTA.ListItems.Count > 0 Then
            If LISTA.SelectedItem.Text <> "" Then

               If TabNOTA.State = 1 Then _
                  TabNOTA.Close

               SQL = "select entrada_id from NOTAENTRADA "
               SQL = SQL & " where numr_nota = " & LISTA.SelectedItem.ListSubItems.item(1).Text
               SQL = SQL & " and serie_nota = '" & Trim(LISTA.SelectedItem.ListSubItems.item(2).Text) & "'"
               SQL = SQL & " and estabelecimento_id = " & EMPRESA_ID_N
               SQL = SQL & " and fornecedor_id = " & LISTA.SelectedItem.ListSubItems.item(6).Text
               TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabNOTA.EOF Then
                  Msg = "Confirma cancelamento de nota referente ao pedido = " & LISTA.SelectedItem.Text & " ? "
                  Style = vbYesNo + 32
                  Title = "Atenção !!!"
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                     GRAVA_CABECA_NOTA
                     MOSTRAR_NOTA
                  End If
               End If
               If TabNOTA.State = 1 Then _
                  TabNOTA.Close
            End If
         End If
      Case "limpar"
         LIMPA_NOTA_ENTRADA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub
'==================CNPJcpf
Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Fornecedores", "", "", ""

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Mask = "##############"
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         MsgBox "Informe CNPJ/CPF corretamente"
         txtCNPJCPF.SetFocus
         Exit Sub
         Else
            If Len(txtCNPJCPF.Text) > 0 Then
               Select Case Len(txtCNPJCPF.Text)
                  Case Is = 11
                    If Not CALCULACPF(txtCNPJCPF.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtCNPJCPF.PromptInclude = False
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    End If
                  Case Is = 14
                    If Not VALIDACGC(txtCNPJCPF.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtCNPJCPF.PromptInclude = False
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    End If
                  Case Is > 14
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCNPJCPF = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
                  Case Is < 11
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCNPJCPF = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
               End Select
               Else
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCNPJCPF = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End If
            txtCNPJCPF.PromptInclude = False
            CRITERIO_A = txtCNPJCPF.Text
      End If
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         CRITERIO_A = txtCNPJCPF.Text
         txtCNPJCPF.Mask = "##############"
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then _
               txtCNPJCPF.Mask = "###.###.###-##"
            If Len(txtCNPJCPF.Text) > 11 Then _
               txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
         txtCNPJCPF.Text = CRITERIO_A
      End If
      txtCNPJCPF.PromptInclude = False

      If TabCliente.State = 1 Then _
         TabCliente.Close
      SQL = "select * from vwFornecedor "
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
         Else
            txtNome.Text = TabCliente!DESCRICAO
            FORNEC_ID_N = TabCliente!FORNECEDOR_ID
      End If
      TabCliente.Close
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSerie.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnota_KeyPress"
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtserie_KeyPress"
End Sub
'=============================
Private Sub LIMPA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   FORNEC_ID_N = 0
   txtSerie.Text = ""
   txtNOTA.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtDtEntrada.PromptInclude = False
   txtDtEntrada.Text = ""
   txtDtEmissao.PromptInclude = False
   txtDtEmissao.Text = ""
   txtTotalNota.Text = ""
   VALOR_TOTAL_N = 0
   txtTotalNota.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTotalNota.Refresh
   LISTA.ListItems.Clear
   LISTAITENS.ListItems.Clear

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_NOTA_ENTRADA"
End Sub

Private Sub GRAVA_CABECA_NOTA()
'On Error GoTo ERRO_TRATA

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   If Trim(LISTA.SelectedItem.ListSubItems.item(1).Text) <> "" And _
      Trim(LISTA.SelectedItem.ListSubItems.item(2).Text) <> "" And _
      LISTA.SelectedItem.ListSubItems.item(6).Text <> "" Then

      SQL = "select * from NOTAENTRADA "
      SQL = SQL & " where numr_nota = " & LISTA.SelectedItem.ListSubItems.item(1).Text
      SQL = SQL & " and serie_nota = '" & Trim(LISTA.SelectedItem.ListSubItems.item(2).Text) & "'"
      SQL = SQL & " and estabelecimento_id = " & EMPRESA_ID_N
      SQL = SQL & " and fornecedor_id = " & LISTA.SelectedItem.ListSubItems.item(6).Text
      Else
         If Trim(txtPedido.Text) <> "" Then
            If IsNumeric(txtPedido.Text) Then
               SQL = "select * from NOTAENTRADA "
               SQL = SQL & " where numr_pedido_compra = " & LISTA.SelectedItem.Text
            End If
         End If
   End If

   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      If (INDR_CONTROLA_ESTOQUE = True) And (TabNOTA!STATUS = "E") Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from NOTAENTRADAITEM "
         SQL = SQL & " where ENTRADA_ID = " & TabNOTA.Fields("entrada_id").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select produto_id from PRODUTO "
            SQL = SQL & " where codg_produto = '" & Trim(TabTemp!Codg_Produto) & "'"
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then

               SQL = "UPDATE ESTOQUE SET "
               SQL = SQL & " qtde_estoque = " & Str(TabProduto!QTDE_ESTOQUE - TabTemp!QTDE_ENTRADA)
               SQL = SQL & " Where estabelecimento_id = " & ESTABELECIMENTO_ID_N
               SQL = SQL & " and produto_id = " & TabProduto.Fields("produto_id").Value
               CONECTA_RETAGUARDA.Execute SQL

               SQL = "UPDATE NOTAENTRADAITEM SET "
               SQL = SQL & " Status = 'C' "
               SQL = SQL & " where ENTRADA_ID = " & TabNOTA.Fields("entrada_id").Value
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabProduto.State = 1 Then _
               TabProduto.Close

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      SQL = "UPDATE NOTAENTRADA SET "
      SQL = SQL & " Status = 'C'"
      SQL = SQL & ", Codg_usu = " & USUARIO_ID_N
      SQL = SQL & " where ENTRADA_ID = " & TabNOTA.Fields("entrada_id").Value
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA_NOTA"
End Sub

Private Sub MOSTRAR_NOTA()
'On Error GoTo ERRO_TRATA

   LISTA.ListItems.Clear
   LISTAITENS.ListItems.Clear

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where estabelecimento_id = " & EMPRESA_ID_N

   If Trim(txtNOTA.Text) <> "" Then _
      SQL = SQL & " and numr_nota = " & txtNOTA.Text
   
   If Trim(txtSerie.Text) <> "" Then _
      SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"

   If FORNEC_ID_N > 0 Then _
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   If Trim(txtPedido.Text) <> "" Then _
      If IsNumeric(txtPedido.Text) Then _
         SQL = SQL & " where numr_pedido_compra = " & Trim(txtPedido.Text)

   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   While Not TabNOTA.EOF
      SETA_GRID_CABECA
      TabNOTA.MoveNext
   Wend
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   FORNEC_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRAR_NOTA"
End Sub

Private Sub SETA_GRID_CABECA()
'On Error GoTo ERRO_TRATA

   Set item = LISTA.ListItems.Add(, "seq." & TabNOTA.Fields("entrada_id").Value, TabNOTA!numr_pedido_compra)
   item.SubItems(1) = TabNOTA!NUMR_NOTA
   item.SubItems(2) = TabNOTA!SERIE_NOTA

   SQL3 = ""
   FORNEC_ID_N = 0

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select cnpjcpf, descricao, fornecedor_id from vwFornecedor "
   SQL = SQL & " where fornecedor_id = " & TabNOTA!FORNECEDOR_ID
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      SQL3 = TabFornecedor!CNPJCPF & " - " & TabFornecedor!DESCRICAO
      FORNEC_ID_N = TabFornecedor.Fields("fornecedor_id").Value
   End If

   item.SubItems(6) = "" & FORNEC_ID_N

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   item.SubItems(3) = Trim(SQL3)

   VALOR_TOTAL_N = 0
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(preco_custo-NOTAENTRADAITEM.valor_desconto) "

   SQL = SQL & " from NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"

   SQL = SQL & " where numr_pedido_compra = " & TabNOTA!numr_pedido_compra

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      VALOR_TOTAL_N = 0 & TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   item.SubItems(4) = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   SqL2 = ""
   If TabNOTA!STATUS = "C" Then _
      SqL2 = TabNOTA!STATUS & " - " & "Cancelada"
   If TabNOTA!STATUS = "E" Then _
      SqL2 = TabNOTA!STATUS & " - " & "Emitida"
   If TabNOTA!STATUS = "A" Then _
      SqL2 = TabNOTA!STATUS & " - " & "Ativo"
   item.SubItems(5) = SqL2

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_CABECA"
End Sub

Private Sub SETA_GRID_ITENS()
'On Error GoTo ERRO_TRATA

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   LISTAITENS.ListItems.Clear

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"
   SQL = SQL & " where numr_pedido_compra = " & TabNOTA!numr_pedido_compra
   SQL = SQL & " order by seq desc"
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      Set item = LISTAITENS.ListItems.Add(, "seq." & TabPedidoItem!SEQ, TabPedidoItem!SEQ)
      item.SubItems(1) = TabPedidoItem!Codg_Produto

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(TabPedidoItem!Codg_Produto) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         item.SubItems(2) = TabTemp!DESCRICAO
      If TabTemp.State = 1 Then _
         TabTemp.Close

      item.SubItems(3) = TabPedidoItem!QTDE_ENTRADA
      item.SubItems(4) = Format(TabPedidoItem!PRECO_CUSTO, strFormatacao2Digitos)
      item.SubItems(5) = Format(TabPedidoItem!PRECO_CUSTO * TabPedidoItem!QTDE_ENTRADA, strFormatacao2Digitos)
      TabPedidoItem.MoveNext
   Wend
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ITENS"
End Sub
