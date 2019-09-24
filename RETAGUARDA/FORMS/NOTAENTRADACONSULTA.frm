VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNOTAENTRADACONSULTA 
   Caption         =   "Consulta Nota Fiscal"
   ClientHeight    =   8340
   ClientLeft      =   3090
   ClientTop       =   2700
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NOTAENTRADACONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   14400
   Begin VB.Frame FraReq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   45
      TabIndex        =   8
      Top             =   720
      Width           =   14295
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   7680
         Picture         =   "NOTAENTRADACONSULTA.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1680
         Width           =   405
      End
      Begin VB.ComboBox cmbProdutoAux 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   360
         Left            =   8160
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   0
         Top             =   495
         Width           =   1215
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   8400
         MaxLength       =   6
         TabIndex        =   1
         Top             =   495
         Width           =   1215
      End
      Begin VB.OptionButton optDTS 
         Caption         =   "Data Emissão"
         Height          =   240
         Left            =   12240
         TabIndex        =   11
         Top             =   225
         Width           =   1575
      End
      Begin VB.OptionButton optDTE 
         Caption         =   "Data Entrada"
         Height          =   240
         Left            =   10560
         TabIndex        =   10
         Top             =   225
         Width           =   1695
      End
      Begin VB.TextBox txtCli 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   9
         Top             =   1200
         Width           =   6135
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   5520
         TabIndex        =   5
         Top             =   1695
         Width           =   2055
      End
      Begin VB.ComboBox cmbProduto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8160
         TabIndex        =   6
         Top             =   1680
         Width           =   5655
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   360
         Left            =   10560
         TabIndex        =   2
         Top             =   615
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   360
         Left            =   12480
         TabIndex        =   3
         Top             =   615
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   5520
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
         Height          =   240
         Left            =   4320
         TabIndex        =   15
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NºNota:"
         Height          =   240
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NºPedido Compra:"
         Height          =   240
         Left            =   6630
         TabIndex        =   13
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código Produto:"
         Height          =   240
         Left            =   3960
         TabIndex        =   12
         Top             =   1680
         Width           =   1545
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   240
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
            Picture         =   "NOTAENTRADACONSULTA.frx":6614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADACONSULTA.frx":6A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADACONSULTA.frx":6D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADACONSULTA.frx":71D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADACONSULTA.frx":762C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADACONSULTA.frx":794C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTAENTRADACONSULTA.frx":7DA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   1270
      ButtonWidth     =   2593
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
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
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   6840
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6240
         Top             =   0
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
               Picture         =   "NOTAENTRADACONSULTA.frx":80C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADACONSULTA.frx":925A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADACONSULTA.frx":A2E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADACONSULTA.frx":B29E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5295
      Left            =   45
      TabIndex        =   17
      Top             =   3000
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   9340
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Text            =   "N.Nota"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Série"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Vlr.Compra"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DtEntrada"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DtEmissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Situação"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fornecedor"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Transportadora"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CNPJFORNC"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmNOTAENTRADACONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmNOTAENTRADACONSULTA)

   CARREGA_COMBO

   FORMULA_REL = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FORMULA_REL = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub cmbProduto_Click()
On Error GoTo ERRO_TRATA

   cmbProdutoAux.ListIndex = cmbProduto.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbProduto_Click"
End Sub

Private Sub Lista_DblClick()
   If Not IsNull(LISTA.SelectedItem.Text) Then
      If Trim(LISTA.SelectedItem.Text) <> "" Then
         frmNOTAENTRADA.txtNOTA.Text = "" & Trim(LISTA.SelectedItem.Text)
         frmNOTAENTRADA.txtSerie.Text = "" & LISTA.SelectedItem.ListSubItems.Item(1).Text
         frmNOTAENTRADA.txtCNPJCPF.PromptInclude = False
         frmNOTAENTRADA.txtCNPJCPF.Text = "" & LISTA.SelectedItem.ListSubItems.Item(8).Text
         Unload Me
      End If
   End If
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub cmdConsProd_Click()
   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
End Sub

Private Sub optDTE_Click()
On Error GoTo ERRO_TRATA

   txtDtIni.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optDTE_Click"
End Sub

Private Sub optDTs_Click()
On Error GoTo ERRO_TRATA

   txtDtIni.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optDTs_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imprimir"
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_nf_entrada.rpt"
         frmRELATORIO10.Show 1
      Case "consultar"
         FORMULA_REL = ""
         CONSULTA_TUDO
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         FORMULA_REL = ""
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCNPJCPF_GotFocus()
On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""
   
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_GotFocus"
End Sub

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyDown"
End Sub

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCNPJCPF.Text = "" Then _
         txtCNPJCPF.Mask = "##############"

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         If TabCliente.State = 1 Then _
            TabCliente.Close

         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
         Else: If TabCliente!NOME <> "" _
               Then txtCli.Text = TabCliente!NOME
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub txtDTfim_GotFocus()
On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDTINI_GotFocus()
On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(txtProduto.Text) <> "" Then
         KeyAscii = 0
         MOSTRA_PRODUTO
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodigo_KeyDown"
End Sub

Private Sub CONSULTA_TUDO()
On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SQL = "select NUMR_NOTA as Nota,SERIE_NOTA as Serie,DT_ENTRADA as DtEntrada,DT_EMISSAO as DtEmissão,"
   SQL = SQL & " Status_Nota as Situação,CFOP,NOME As Fornecedor,NOME_TRANSP as Transportadora, entrada_id,"
   SQL = SQL & " CGCCPF "

   SQL = SQL & " from vwRel_Nf_Entrada "
   SQL = SQL & " where numr_nota > 0 "

   FORMULA_REL = "{vwRel_Nf_Entrada.entrada_id} > 0 "

   If Trim(txtNOTA.Text) <> "" Then
      If IsNumeric(txtNOTA.Text) Then
         SQL = SQL & " and numr_nota = " & txtNOTA.Text

         FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.numr_nota} = " & Trim(txtNOTA.Text)
      End If
   End If

   If Trim(txtPedido.Text) <> "" Then
      If IsNumeric(txtPedido.Text) Then
         SQL = SQL & " and numr_pedido_compra = " & Trim(txtPedido.Text)

         FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.NUMR_PEDIDO_COMPRA} = " & Trim(txtPedido.Text)
      End If
   End If

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      SQL = SQL & " and CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.CGCCPF} = '" & Trim(txtCNPJCPF.Text) & "'"
   End If
   txtCNPJCPF.PromptInclude = True

   If optDTE.Value = True Then
      If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
         SQL = SQL & "and dt_entrada >= '" & DMA(txtDtIni.Text) & "'"
         SQL = SQL & "and dt_entrada <= '" & DMA(txtDtFim.Text) & "'"

FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_entrada} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_entrada} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
      End If
   End If

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.codg_prod} = '" & Trim(txtProduto.Text) & "'"
   End If

   SQL = SQL & " ORDER BY entrada_id desc"

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_TUDO"
End Sub

Private Sub SETA_GRID()
On Error GoTo ERRO_TRATA

   LISTA.ListItems.Clear
   NUMR_SEQ_N = 0
   SQL3 = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If Trim(SQL3) <> Trim(TabTemp.Fields("Nota").Value) Then
         NUMR_SEQ_N = NUMR_SEQ_N + 1

         VALOR_ITEM_N = 0

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select sum(preco_custo*qtd_entrada) from NOTAENTRADAITEM "
         SQL = SQL & " Where ENTRADA_ID = " & TabTemp.Fields("entrada_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               VALOR_ITEM_N = 0 & TabConsulta.Fields(0).Value
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         Set Item = LISTA.ListItems.Add(, "seq." & NUMR_SEQ_N, Trim(TabTemp.Fields("Nota").Value))

         Item.SubItems(1) = "" & Trim(TabTemp.Fields("SERIE").Value)

         Item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         Item.SubItems(3) = "" & Trim(TabTemp.Fields("DtEntrada").Value)
         Item.SubItems(4) = "" & Trim(TabTemp.Fields("DtEmissão").Value)
         Item.SubItems(5) = "" & Trim(TabTemp.Fields("Situação").Value)
         Item.SubItems(6) = "" & Trim(TabTemp.Fields("Fornecedor").Value)
         Item.SubItems(7) = "" & Trim(TabTemp.Fields("Transportadora").Value)
         Item.SubItems(8) = "" & Trim(TabTemp.Fields("CGCCPF").Value)
      End If
      SQL3 = Trim(TabTemp.Fields("Nota").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_TUDO()
On Error GoTo ERRO_TRATA

   FORMULA_REL = ""
   LISTA.ListItems.Clear
   txtProduto.Text = ""
   txtPedido.Text = ""
   txtNOTA.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   optDTE.Value = False
   optDTS.Value = False
   txtNOTA.SetFocus
   cmbProdutoAux.Text = ""
   cmbProduto.Text = ""

   SQL = "select * from NOtaentrada "
   SQL = SQL & " where numr_nota < 0 "

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub MOSTRA_PRODUTO()
On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select produto_id,codg_produto,descricao from PRODUTO "
   SQL = SQL & " where situacao = 'A' "

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"

   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      cmbProduto.Text = Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("codg_produto").Value)

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub CARREGA_COMBO()
On Error GoTo ERRO_TRATA

   cmbProduto.Clear
   cmbProdutoAux.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select produto_id,codg_produto,descricao from PRODUTO "
   SQL = SQL & " where situacao = 'A' "
   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbProduto.AddItem Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("codg_produto").Value)
      cmbProdutoAux.AddItem Trim(TabTemp.Fields("codg_produto").Value)

      TabTemp.MoveNext
   Wend

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub
