VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNOTACONSULTA 
   Caption         =   "Consulta Nota Fiscal"
   ClientHeight    =   6900
   ClientLeft      =   3090
   ClientTop       =   2700
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NOTACONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin VB.Frame fraProduto 
      Caption         =   "Produto"
      Height          =   735
      Left            =   6480
      TabIndex        =   32
      Top             =   2520
      Width           =   7215
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
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   615
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
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   255
         Width           =   2055
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2160
         Picture         =   "NOTACONSULTA.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame fraPessoa 
      Caption         =   "Fornecedor"
      Height          =   735
      Left            =   6480
      TabIndex        =   30
      Top             =   1800
      Width           =   7215
      Begin VB.CommandButton cmdConsPessoa 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2160
         Picture         =   "NOTACONSULTA.frx":6614
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtPessoa 
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
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   31
         Top             =   240
         Width           =   4455
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   240
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
   End
   Begin VB.Frame fraDoc 
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   6255
      Begin VB.OptionButton optProtocolo 
         Caption         =   "Protocolo NFe"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optChaveNFe 
         Caption         =   "Chave NFe"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2040
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optPedidoCompra 
         Caption         =   "Pedido Compra"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optNota 
         Caption         =   "Nº Nota"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPedidoVenda 
         Caption         =   "Pedido Venda"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtDoc 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   120
         MaxLength       =   6
         TabIndex        =   14
         Top             =   960
         Width           =   6015
      End
   End
   Begin VB.Frame fraStatus 
      Height          =   1095
      Left            =   12120
      TabIndex        =   25
      Top             =   720
      Width           =   1575
      Begin VB.ComboBox cmbSituacaoAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbSituacao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Text            =   "-- Selecione --"
         ToolTipText     =   "CFOP"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Situação"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   345
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraData 
      Height          =   1095
      Left            =   7560
      TabIndex        =   22
      Top             =   720
      Width           =   4455
      Begin VB.OptionButton optDtEntrada 
         Caption         =   "Data Entrada"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optDtEmissao 
         Caption         =   "Data Emissão"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   360
         Left            =   840
         TabIndex        =   6
         Top             =   600
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
         Left            =   3000
         TabIndex        =   7
         Top             =   600
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
      Begin VB.Label lblFim 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fim:"
         Height          =   240
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblIni 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.Frame fraCFOP 
      Height          =   1095
      Left            =   50
      TabIndex        =   21
      Top             =   720
      Width           =   7455
      Begin VB.ComboBox cmbCFOPAux 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optCFOP 
         Caption         =   "&Todas"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optCFOP 
         Caption         =   "&Entrada"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optCFOP 
         Caption         =   "&Saida"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbCFOP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Text            =   "-- Selecione --"
         ToolTipText     =   "CFOP"
         Top             =   600
         Width           =   7215
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
            Picture         =   "NOTACONSULTA.frx":7016
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTACONSULTA.frx":746A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTACONSULTA.frx":7786
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTACONSULTA.frx":7BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTACONSULTA.frx":802E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTACONSULTA.frx":834E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTACONSULTA.frx":87A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   1270
      ButtonWidth     =   2461
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
         TabIndex        =   20
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
               Picture         =   "NOTACONSULTA.frx":8AC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTACONSULTA.frx":9C5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTACONSULTA.frx":ACEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTACONSULTA.frx":BCA0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstNota 
      Height          =   3495
      Left            =   45
      TabIndex        =   18
      Top             =   3360
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   6165
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
      NumItems        =   10
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
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "id"
         Object.Width           =   2
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   13695
      DesignHeight    =   6900
   End
End
Attribute VB_Name = "frmNOTACONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   CARREGA_COMBO
   CARREGA_CFOP

   FORMULA_REL = ""
   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FORMULA_REL = ""
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

Private Sub lstNota_Click()
On Error Resume Next

   If Not IsNull(lstNota.SelectedItem.ListSubItems.item(9).Text) Then _
      If Trim(lstNota.SelectedItem.ListSubItems.item(9).Text) <> "" Then _
         If IsNumeric(lstNota.SelectedItem.ListSubItems.item(9).Text) Then _
            NUMR_ID_N = 0 & Trim(lstNota.SelectedItem.ListSubItems.item(9).Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imprimir"
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
         If NUMR_ID_N > 0 Then _
            FORMULA_REL = "{vwRel_Nf_Entrada.entrada_id} = " & NUMR_ID_N

         Nome_Relatorio = "rel_nf_entrada.rpt"
         frmRELATORIO10.Show 1
         Err.Clear
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
   Err.Clear
End Sub

Private Sub cmbCFOP_Click()
On Error Resume Next

   cmbCFOPAux.ListIndex = cmbCFOP.ListIndex
End Sub

Private Sub cmbProduto_Click()
On Error Resume Next

   cmbProdutoAUX.ListIndex = cmbProduto.ListIndex
End Sub

Private Sub cmbSituacao_Click()
On Error Resume Next

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex
End Sub

Private Sub LSTNOTA_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstNota.SelectedItem.Text) Then
      If Trim(lstNota.SelectedItem.Text) <> "" Then
         frmNOTAENTRADA.txtNOTA.Text = "" & Trim(lstNota.SelectedItem.Text)
         frmNOTAENTRADA.txtSerie.Text = "" & lstNota.SelectedItem.ListSubItems.item(1).Text
         frmNOTAENTRADA.txtCNPJCPF.PromptInclude = False
         frmNOTAENTRADA.txtCNPJCPF.Text = "" & lstNota.SelectedItem.ListSubItems.item(8).Text
         Unload Me
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LSTNOTA_DblClick"
End Sub

Private Sub LSTNOTA_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstNota, ColumnHeader
End Sub

Private Sub cmdConsProd_Click()
   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      Call txtProduto_KeyPress(13)
      txtProduto.SetFocus
   End If
End Sub

Private Sub optCFOP_Click(Index As Integer)
'On Error GoTo ERRO_TRATA

   Select Case Index
      Case 0
         FraPessoa.Caption = "Pessoa:"
         optPedidoVenda.Visible = True
         optDtEntrada.Visible = True
         optPedidoCompra.Visible = True
      Case 1   'entrada
         FraPessoa.Caption = "Fornecedor:"
         optDtEntrada.Visible = True
         optPedidoCompra.Visible = True
         optPedidoVenda.Visible = False
      Case 2   'saida
         FraPessoa.Caption = "Cliente:"
         optDtEntrada.Visible = False
         optPedidoCompra.Visible = False
         optPedidoVenda.Visible = True
   End Select

   CARREGA_CFOP
   cmbCFOP.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optCFOP_Click"
End Sub

Private Sub optDtEntrada_Click()
'On Error GoTo ERRO_TRATA

   txtDtIni.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optDtEntrada_Click"
End Sub

Private Sub optDtEmissao_Click()
'On Error GoTo ERRO_TRATA

   txtDtIni.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optDtEmissao_Click"
End Sub

Private Sub optNota_Click()
   txtDoc.SetFocus
End Sub

Private Sub optPedidoVenda_Click()
   txtDoc.SetFocus
End Sub

Private Sub optPedidoCompra_Click()
   txtDoc.SetFocus
End Sub

Private Sub optChaveNFe_Click()
   txtDoc.SetFocus
End Sub

Private Sub optProtocolo_Click()
   txtDoc.SetFocus
End Sub

Private Sub cmdConsPessoa_Click()
'On Error GoTo ERRO_TRATA

   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      txtCNPJCPF.SetFocus
      Call TXTCNPJCPF_KeyPress(13)
   End If
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsPessoa_Click"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""
   
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCNPJCPF.Text = "" Then _
         txtCNPJCPF.Mask = "##############"

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select descricao,pessoa_id from vwFornecedor "
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         If TabCliente.State = 1 Then _
            TabCliente.Close

         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
         Else
            txtPessoa.Text = "" & Trim(TabCliente.Fields("descricao").Value)
            PESSOA_ID_N = 0 & Trim(TabCliente.Fields("pessoa_id").Value)
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

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

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

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SQL = "select * from vwRel_Nf_Entrada "
   SQL = SQL & " where numr_nota > 0 "

   FORMULA_REL = "{vwRel_Nf_Entrada.entrada_id} > 0 "

   If Trim(txtDoc.Text) <> "" Then
      If IsNumeric(txtDoc.Text) Then
         If optNota.Value = True Then
            SQL = SQL & " and numr_nota = " & txtDoc.Text
            FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.numr_nota} = " & Trim(txtDoc.Text)
         End If
         If optPedidoCompra.Value = True Then
            SQL = SQL & " and pedidocompra_id = " & Trim(txtDoc.Text)
            FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.pedidocompra_id} = " & Trim(txtDoc.Text)
         End If
      End If
   End If

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      SQL = SQL & " and cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.cnpjcpf} = '" & Trim(txtCNPJCPF.Text) & "'"
   End If
   txtCNPJCPF.PromptInclude = True

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optDtEntrada.Value = True Then
         SQL = SQL & "and dt_entrada >= '" & DMA(txtDtIni.Text, "i") & "'"
         SQL = SQL & "and dt_entrada <= '" & DMA(txtDtFim.Text, "f") & "'"

DATA_INI = DMA(txtDtIni.Text, "i")
DATA_FIM = DMA(txtDtFim.Text, "f")
FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_entrada} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_entrada} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"

         'FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_entrada} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
         'FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_entrada} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
         Else
            If optDtEmissao.Value = True Then
               SQL = SQL & "and dt_emissao >= '" & DMA(txtDtIni.Text, "i") & "'"
               SQL = SQL & "and dt_emissao <= '" & DMA(txtDtFim.Text, "f") & "'"

DATA_INI = DMA(txtDtIni.Text, "i")
DATA_FIM = DMA(txtDtFim.Text, "f")
FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_emissao} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.dt_emissao} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"

               'FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_emissao.dt_emissao} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
               'FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_emissao.dt_emissao} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
            End If
      End If
   End If

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.CODG_PRODUTO} = '" & Trim(txtProduto.Text) & "'"
   End If

   If Trim(cmbSituacaoAUX.Text) <> "" Then
      SQL = SQL & " and vwRel_Nf_Entrada.status_nota = '" & Trim(cmbSituacaoAUX.Text) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Entrada.status} = '" & Trim(cmbSituacaoAUX.Text) & "'"
   End If

   SQL = SQL & " ORDER BY entrada_id desc"

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_TUDO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstNota.ListItems.Clear
   NUMR_SEQ_N = 0
   SQL3 = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close
   
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If Trim(SQL3) <> Trim(TabTemp.Fields("numr_nota").Value) Then
         NUMR_SEQ_N = NUMR_SEQ_N + 1

         VALOR_ITEM_N = 0

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select sum(preco_custo*qtde_entrada) from NOTAENTRADAITEM "
         SQL = SQL & " Where ENTRADA_ID = " & TabTemp.Fields("entrada_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               VALOR_ITEM_N = 0 & TabConsulta.Fields(0).Value
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         Set item = lstNota.ListItems.Add(, "seq." & NUMR_SEQ_N, Trim(TabTemp.Fields("numr_nota").Value))

         item.SubItems(1) = "" & Trim(TabTemp.Fields("serie_nota").Value)

         item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         item.SubItems(3) = "" & Trim(TabTemp.Fields("Dt_Entrada").Value)
         item.SubItems(4) = "" & Trim(TabTemp.Fields("Dt_Emissao").Value)
         item.SubItems(5) = "" & Trim(TabTemp.Fields("status_nota").Value)
         item.SubItems(6) = "" & Trim(TabTemp.Fields("nomeFornecedor").Value)
         'Item.SubItems(7) = "" & Trim(TabTemp.Fields("Transportadora").Value)
         item.SubItems(8) = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         item.SubItems(9) = "" & Trim(TabTemp.Fields("entrada_id").Value)
      End If
      SQL3 = Trim(TabTemp.Fields("numr_nota").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   NUMR_ID_N = 0
   optCFOP(0).Value = True
   cmbCFOP.Text = ""
   cmbCFOPAux.Text = ""
   optDtEmissao.Value = True
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   optNota.Value = True
   txtDoc.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtPessoa.Text = ""
   cmbProdutoAUX.Text = ""
   cmbProduto.Text = ""
   PESSOA_ID_N = 0
   FORMULA_REL = ""
   lstNota.ListItems.Clear

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

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
'On Error GoTo ERRO_TRATA

   cmbProduto.Clear
   cmbProdutoAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select produto_id,codg_produto,descricao from PRODUTO "
   SQL = SQL & " where situacao = 'A' "
   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbProduto.AddItem Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("codg_produto").Value)
      cmbProdutoAUX.AddItem Trim(TabTemp.Fields("codg_produto").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbSituacao.Clear
   cmbSituacaoAUX.Clear

   cmbSituacao.AddItem "Processada"
   cmbSituacaoAUX.AddItem "E"

   cmbSituacao.AddItem "Pendente"
   cmbSituacaoAUX.AddItem "A"

   cmbSituacao.AddItem "Cancelada"
   cmbSituacaoAUX.AddItem "C"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub

Private Sub CARREGA_CFOP()
'On Error GoTo ERRO_TRATA

   'CFOP
   cmbCFOPAux.Clear
   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP With (NOLOCK)"
   SQL = SQL & " where cfop_id > 0 "

   If optCFOP(1).Value = True Then _
      SQL = SQL & " and left(CFOP_id,1) <= 2 "

   If optCFOP(2).Value = True Then _
      SQL = SQL & " and left(CFOP_id,1) >= 5 "

   SQL = SQL & " order by cfop_id "

   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      Do Until TabDESCR.EOF
         DoEvents
         cmbCFOPAux.AddItem Trim(TabDESCR!CFOP_ID)
         cmbCFOP.AddItem Trim(TabDESCR!CFOP_ID) & "-" & Trim(TabDESCR!DESCRICAO)
         TabDESCR.MoveNext
      Loop
      Else
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         MsgBox "Cadastro de CFOP com problemas. Não foi localizado nenhum codigo de CFOP cadastrado", vbCritical
         Exit Sub
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_CFOP"
End Sub

Sub MONTA_CONSULTA()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_GLOBAL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MFADOC from MFA010"
   SQL = SQL & " where mfadoc = '" & Trim(txtNOTA.Text) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabTemp.EOF Then
End If

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub
