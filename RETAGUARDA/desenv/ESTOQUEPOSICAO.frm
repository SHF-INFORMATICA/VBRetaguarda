VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmESTOQUEPOSICAO 
   Caption         =   "Posição Estoque"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ESTOQUEPOSICAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin PVMaskEditLib.PVMaskEdit txtCNPJCPF 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524288
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   253
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BorderStyle     =   1
      Text            =   ""
   End
   Begin VB.CheckBox chkFamilia 
      Caption         =   "Por Família"
      Height          =   240
      Left            =   9960
      TabIndex        =   28
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7200
      MaxLength       =   100
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdConsFor 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   6720
      Picture         =   "ESTOQUEPOSICAO.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3050
      Picture         =   "ESTOQUEPOSICAO.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1320
      Width           =   405
   End
   Begin VB.OptionButton optSomente0 
      Caption         =   "Somente Qtde Zerada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7440
      TabIndex        =   9
      Top             =   1440
      Width           =   2175
   End
   Begin VB.OptionButton optConsiderar0 
      Caption         =   "Considerar Qtde Zerada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7440
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   15
      TabIndex        =   18
      Top             =   2400
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Saída"
      TabPicture(0)   =   "ESTOQUEPOSICAO.frx":7016
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstGeral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView lstGeral 
         Height          =   3945
         Left            =   60
         TabIndex        =   30
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6959
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Produto"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Referencia"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   8259
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qtde.Atual"
            Object.Width           =   3003
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qtde.Compra"
            Object.Width           =   3003
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Inventário"
            Object.Width           =   3003
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Familia"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fornecedor"
            Object.Width           =   3754
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Dt.Ult.Venda"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Dt.Ult.Compra"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cmbFamiliaAUX 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Left            =   1080
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtLocacao 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5160
      MaxLength       =   6
      TabIndex        =   4
      ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtDescProd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      MaxLength       =   100
      TabIndex        =   12
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cmbSituacao 
      Height          =   360
      Left            =   9960
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1270
      ButtonWidth     =   3387
      ButtonHeight    =   1111
      Appearance      =   1
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
            Caption         =   "&Imprimir Tela"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   285
         Left            =   8400
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8760
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
               Picture         =   "ESTOQUEPOSICAO.frx":7032
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEPOSICAO.frx":81CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEPOSICAO.frx":925B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEPOSICAO.frx":A210
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEPOSICAO.frx":B31B
               Key             =   ""
            EndProperty
         EndProperty
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
      DesignWidth     =   11940
      DesignHeight    =   7035
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTotUN 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   240
      Left            =   6600
      TabIndex        =   29
      Top             =   6720
      Width           =   315
   End
   Begin VB.Label lblTotKG 
      AutoSize        =   -1  'True
      Caption         =   "000"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   7440
      TabIndex        =   27
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label lblTotProduto 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   240
      Left            =   9960
      TabIndex        =   26
      Top             =   1920
      Width           =   105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   0
      X2              =   11880
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label lblTop 
      Caption         =   "+ vendidos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Final:"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   11880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Locação:"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Fornecedor:"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Família:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   255
      Left            =   9000
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmESTOQUEPOSICAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   CARREGA_FAMILIA_PRODUTO
   CARREGA_SITUAÇÃO_PRODUTO

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   VALOR_ITEM_N = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imprimir"
         MONTA_REL
      Case "consultar"
         lblTotKG.Caption = ""
         lblTotUN.Caption = ""
         Me.Enabled = False

         CRIA_TAB_TEMP
         COMPARA_GERAL

         Me.Enabled = True
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

   lblTop.Caption = ""
   If SSTab1.Tab = 0 Then _
      lblTop.Caption = "+Vendidos"
   If SSTab1.Tab = 1 Then _
      lblTop.Caption = "+Comprados"
   'If SSTab1.Tab = 2 Then _


   lblTop.Refresh
End Sub

Private Sub cmbFamilia_Click()
On Error Resume Next

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

End Sub

Private Sub lstGeral_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstGeral, ColumnHeader
End Sub

Private Sub cmdConsFor_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmConsultaPessoa.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      
      MOSTRA_FORNECEDOR
   End If
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsFor_Click"
End Sub

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
        txtCNPJCPF.PromptInclude = False
        If Trim(txtCNPJCPF.Text) = "" Then
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
                         MsgBox "CGC com DV incorreto !!! "
                         txtCNPJCPF.PromptInclude = False
                         txtCNPJCPF = ""
                         txtCNPJCPF.SetFocus
                         Exit Sub
                      End If
                    Case Is > 14
                       MsgBox "CGC/CPF com DV incorreto !!! "
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    Case Is < 11
                       MsgBox "CGC/CPF com DV incorreto !!! "
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                 End Select
                 Else
                    MsgBox "CGC/CPF com DV incorreto !!! "
                    txtCNPJCPF = ""
                    txtCNPJCPF.SetFocus
                    Exit Sub
              End If
              txtCNPJCPF.PromptInclude = False
              CRITERIO = txtCNPJCPF.Text
        End If
        txtCNPJCPF.PromptInclude = False
        If Trim(txtCNPJCPF.Text) <> "" Then
           CRITERIO = txtCNPJCPF.Text
           If Not IsNull(txtCNPJCPF.Text) Then
              If Len(txtCNPJCPF.Text) <= 11 Then
                 txtCNPJCPF.Mask = "###.###.###-##"
                 Else: txtCNPJCPF.Mask = "##.###.###/####-##"
              End If
           End If
           txtCNPJCPF.Text = CRITERIO
           Else: txtCNPJCPF.Mask = "##############"
        End If
        txtCNPJCPF.PromptInclude = False

      MOSTRA_FORNECEDOR

      txtRazao.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub cmdConsProd_Click()
   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
End Sub

Private Sub txtDTINI_GotFocus()

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDTfim_GotFocus()

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTop_KeyPress"
End Sub

Sub CARREGA_FAMILIA_PRODUTO()
'On Error GoTo ERRO_TRATA

   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & "order by descricao "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFamilia.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("FAMILIAPRODUTO_ID").Value)
      cmbFamiliaAUX.AddItem Trim(TabDESCR.Fields("FAMILIAPRODUTO_ID").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMB_FAMILIA_PRODUTO"
End Sub

Sub CARREGA_SITUAÇÃO_PRODUTO()
'On Error GoTo ERRO_TRATA

   cmbSituacao.Clear

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'P'"
   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabAUX.EOF
      cmbSituacao.AddItem Trim(TabAUX.Fields("DESCRICAO").Value)
      TabAUX.MoveNext
   Wend
   If TabAUX.State = 1 Then _
      TabAUX.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMB_FAMILIA_PRODUTO"
End Sub

Sub MOSTRA_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   FORNEC_ID_N = 0

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      FORNEC_ID_N = 0 & TabCliente!FORNECEDOR_ID
      txtNome.Text = TabCliente!NOME
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_FORNECEDOR"
End Sub

Sub LIMPA_TUDO()
   PESSOA_ID_N = 0
   chkFamilia.Value = 0
   lblTotProduto.Caption = ""
   lblTotKG.Caption = ""
   lblTotUN.Caption = ""
   FORNEC_ID_N = 0
   PRODUTO_ID_N = 0
   cmbFamiliaAUX.Text = ""
   cmbFamilia.Text = ""
   txtLocacao.Text = ""
   cmbSituacao.Text = ""
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   optConsiderar0.Value = False
   optSomente0.Value = False
   txtProduto.Text = ""
   txtDescProd.Text = ""
   txtDtIni.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtTop.Text = "0"
   lstGeral.ListItems.Clear
End Sub

Sub PROCURA_PRODUTO()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0

   If Trim(txtProduto.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select produto_id,descricao from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDescProd.Text = TabConsulta.Fields("descricao").Value
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
End Sub

Sub MONTA_CONSULTA_SQL()

   SQL = "select produto_id,estabelecimento_id,qtd_pedida from vwPOSICAOESTOQUE WITH (NOLOCK)"

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and (statusPedido >= 3 and statuspedido < 9) "
   SQL = SQL & " and vwPOSICAOESTOQUE.TIPO_REGISTRO = 'R' "
   SQL = SQL & " and vwPOSICAOESTOQUE.statusitem <> 'C' "

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & Trim(cmbFamiliaAUX.Text)

   If Trim(txtLocacao.Text) <> "" Then _
      SQL = SQL & " and locacao = '" & Trim(txtLocacao.Text) & "'"

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao = '" & Left(cmbSituacao.Text, 1) & "'"

   If Trim(txtCNPJCPF.Text) <> "" Then _
      If IsNumeric(txtCNPJCPF.Text) Then _
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"
      Else
         If Trim(txtDescProd.Text) <> "" Then _
            SQL = SQL & " and produtodescrição = '" & Trim(txtDescProd.Text) & "'"
   End If

   If optConsiderar0.Value = True Then _
      SQL = SQL & " and qtdeatualproduto > 0"

   If optSomente0.Value = True Then _
      SQL = SQL & " and qtdeatualproduto = 0"

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True

      SQL = SQL & " and dt_req >= '" & DMA(txtDtIni.Text, "i") & "'"
      SQL = SQL & " and dt_req <= '" & DMA(txtDtFim.Text, "f") & "'"
   End If

   SQL = SQL & " order by codg_produto desc "

End Sub

Sub BUSCA_INVENTARIO(Tipo_Mov_A As String)
'On Error GoTo ERRO_TRATA

   Dim QTDE_SAIDA_N   As Double
   Dim QTDE_ENTRADA_N As Double
   CONTA_REG_PROGRESSO = 0
   CONT_N = 0

   If TabInventario.State = 1 Then _
      TabInventario.Close

   SQL = "SELECT INVENTARIO.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.SITUACAO, PRODUTO.DT_ULT_COMPRA, "
   SQL = SQL & " produto.referencia, PRODUTO.DT_ULT_VENDA, FORNECEDOR.FORNECEDOR_ID, "
   SQL = SQL & " FORNECEDOR.PESSOA_ID, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS Nome "
   SQL = SQL & " FROM PESSOA "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON PESSOA.PESSOA_ID = FORNECEDOR.PESSOA_ID "
   SQL = SQL & " RIGHT OUTER JOIN INVENTARIO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON INVENTARIO.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " ON FORNECEDOR.FORNECEDOR_ID = PRODUTO.FORNECEDOR_ID"

   SQL = SQL & " where INVENTARIO.status = 'F' "
   SQL = SQL & " and Tipo_Mov = '" & Trim(Tipo_Mov_A) & "'"
   SQL = SQL & " and INVENTARIO.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"
      Else
         If Trim(txtDescProd.Text) <> "" Then _
            SQL = SQL & " and produto.descricao = '" & Trim(txtDescProd.Text) & "'"
   End If

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True

      SQL = SQL & " and dt_lote >= '" & (txtDtIni.Text) & "'"
      SQL = SQL & " and dt_lote <= '" & (txtDtFim.Text) & "'"
   End If

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & Trim(cmbFamiliaAUX.Text)

   If Trim(txtLocacao.Text) <> "" Then _
      SQL = SQL & " and txtLocacao = '" & Trim(txtLocacao.Text) & "'"

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao = '" & Left(cmbSituacao.Text, 1) & "'"

   If Trim(txtCNPJCPF.Text) <> "" Then _
      If IsNumeric(txtCNPJCPF.Text) Then _
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   If optConsiderar0.Value = True Then _
      SQL = SQL & " and qtde_estoque > 0"

   If optSomente0.Value = True Then _
      SQL = SQL & " and qtde_estoque = 0"

   TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabInventario.EOF
      QTDE_SAIDA_N = 0
      QTDE_ENTRADA_N = 0

      If Trim(Tipo_Mov_A) = "S" Then _
         QTDE_SAIDA_N = 0 & TabInventario.Fields("qtd_primeira").Value
      If Trim(Tipo_Mov_A) = "E" Then _
         QTDE_ENTRADA_N = 0 & TabInventario.Fields("qtd_primeira").Value

      GRAVA_POSICAOESTOQUE TabInventario.Fields("produto_id").Value, _
                           TabInventario.Fields("estabelecimento_id").Value, _
                           0, _
                           QTDE_SAIDA_N, _
                            0, _
                           0, _
                           QTDE_ENTRADA_N, _
                           0

      DoEvents
      TabInventario.MoveNext
   Wend
   If TabInventario.State = 1 Then _
      TabInventario.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_INVENTARIO"
End Sub

Sub ATUALIZA_UNIDADE_MEDIDA_TABELA_TEMP()
'On Error GoTo ERRO_TRATA

   SQL = "UPDATE POSICAOESTOQUE "
   SQL = SQL & " SET POSICAOESTOQUE.UNIDADE_MEDIDA = PRODUTO.UNIDADE_MEDIDA "
   SQL = SQL & " from POSICAOESTOQUE"
   SQL = SQL & " inner join produto"
   SQL = SQL & " on POSICAOESTOQUE.produto_id = produto.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "UPDATE vwRel_EstoqueEntrada "
   SQL = SQL & " SET vwRel_EstoqueEntrada.UNIDADE_MEDIDA = PRODUTO.UNIDADE_MEDIDA "
   SQL = SQL & " from vwRel_EstoqueEntrada "
   SQL = SQL & " inner join produto"
   SQL = SQL & " on vwRel_EstoqueEntrada.produto_id = produto.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select sum(qtde_vendida) from POSICAOESTOQUE "
   SQL = SQL & " where UPPER(unidade_medida) = 'KG'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields(0).Value) Then
         lblTotKG.Caption = "Total Kg = " & Format(TabTemp.Fields(0).Value, strFormatacaoKilo)
         lblTotKG.Refresh
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(qtde_vendida) from POSICAOESTOQUE "
   SQL = SQL & " where UPPER(unidade_medida) = 'UN'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields(0).Value) Then
         lblTotUN.Caption = "Total UN = " & TabTemp.Fields(0).Value
         lblTotUN.Refresh
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_UNIDADE_MEDIDA_TABELA_TEMP"
End Sub

Sub COMPARA_GERAL()
'On Error GoTo ERRO_TRATA

   Dim TabGeral   As New ADODB.Recordset

   lstGeral.ListItems.Clear
   lstGeral.Visible = False

   lblTotProduto.Caption = NUMR_SEQ_N = 0
   Me.Enabled = False

'====PEDIDO VENDA, ATUALIZANDO (QTDE_SAIDA_VENDA)
   If TabGeral.State = 1 Then _
      TabGeral.Close

   MONTA_CONSULTA_SQL

   TabGeral.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabGeral.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      lblTotProduto.Caption = "Processando ... " & NUMR_SEQ_N
      DoEvents

      GRAVA_POSICAOESTOQUE TabGeral.Fields("produto_id").Value, _
                           TabGeral.Fields("estabelecimento_id").Value, _
                           TabGeral.Fields("qtd_pedida").Value, _
                           0, _
                           0, _
                           0, _
                           0, _
                           0
      TabGeral.MoveNext
   Wend
   If TabGeral.State = 1 Then _
      TabGeral.Close

'====INVENTARIO SAIDA, ATUALIZANDO (QTDE_SAIDA_INVENTARIO)
   BUSCA_INVENTARIO "S"

'====TRANSFERENCIA SAIDA, ATUALIZANDO (QTDE_SAIDA_TRANSFERENCIA)

'====NOTA ENTRADA, ATUALIZANDO (QTDE_ENTRADA_NOTA)
   SQL = "select produto_id,estabelecimento_id,qtde_entrada from vwRel_Nf_Entrada WITH (NOLOCK)"

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and status_nota = 'E' "

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         SQL = SQL & " and familiaproduto_id = " & Trim(cmbFamiliaAUX.Text)

   If Trim(txtLocacao.Text) <> "" Then _
      SQL = SQL & " and txtLocacao = '" & Trim(txtLocacao.Text) & "'"

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao = '" & Left(cmbSituacao.Text, 1) & "'"

   If Trim(txtCNPJCPF.Text) <> "" Then _
      If IsNumeric(txtCNPJCPF.Text) Then _
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"
      Else
      If Trim(txtDescProd.Text) <> "" Then _
         SQL = SQL & " and descricao = '" & Trim(txtDescProd.Text) & "'"
   End If

   If optConsiderar0.Value = True Then _
      SQL = SQL & " and qtde > 0"

   If optSomente0.Value = True Then _
      SQL = SQL & " and qtde = 0"

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   If Trim(txtDtIni.Text) <> "" And Trim(txtDtFim.Text) <> "" Then
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True

      SQL = SQL & " and dt_entrada >= '" & (txtDtIni.Text) & "'"
      SQL = SQL & " and dt_entrada <= '" & (txtDtFim.Text) & "'"
   End If

   SQL = SQL & " order by codg_produto desc "

TabGeral.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabGeral.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      lblTotProduto.Caption = "Processando ... " & NUMR_SEQ_N
      DoEvents

      GRAVA_POSICAOESTOQUE TabGeral.Fields("produto_id").Value, _
                           TabGeral.Fields("estabelecimento_id").Value, _
                           0, _
                           0, _
                           0, _
                           TabGeral.Fields("qtde_entrada").Value, _
                           0, _
                           0
      TabGeral.MoveNext
   Wend
   If TabGeral.State = 1 Then _
      TabGeral.Close
   
'====INVENTARIO ENTRADA, ATUALIZANDO (QTDE_ENTRADA_INVENTARIO)
   BUSCA_INVENTARIO "S"

'====TRANSFERENCIA ENTRADA, ATUALIZANDO (QTDE_ENTRADA_TRANSFERENCIA_N)


   lstGeral.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "COMPARA_GERAL"
End Sub

Sub CRIA_TAB_TEMP()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "POSICAOESTOQUE", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP table [dbo].[POSICAOESTOQUE]"

   SQL = "CREATE TABLE [dbo].[POSICAOESTOQUE]("

   SQL = SQL & " [PRODUTO_ID]                   [bigint] NOT NULL,"
   SQL = SQL & " [ESTABELECIMENTO_ID]           [int]    NOT NULL,"

   SQL = SQL & " [QTDE_ATUAL]                   [FLOAT]  NULL,"

   SQL = SQL & " [QTDE_SAIDA_VENDA]             [FLOAT]  NULL,"
   SQL = SQL & " [QTDE_SAIDA_INVENTARIO]        [FLOAT]  NULL,"
   SQL = SQL & " [QTDE_SAIDA_TRANSFERENCIA]     [FLOAT]  NULL,"

   SQL = SQL & " [QTDE_ENTRADA_NOTA]            [FLOAT]  NULL,"
   SQL = SQL & " [QTDE_ENTRADA_INVENTARIO]      [FLOAT]  NULL,"
   SQL = SQL & " [QTDE_ENTRADA_TRANSFERENCIA]   [FLOAT]  NULL"

   SQL = SQL & " CONSTRAINT [PK_POSICAOESTOQUE] PRIMARY KEY CLUSTERED("
   SQL = SQL & " [PRODUTO_ID],[ESTABELECIMENTO_ID] Asc)"
   SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, "
   SQL = SQL & " ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"

   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from POSICAOESTOQUE"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CRIA_TAB_TEMP"
End Sub

Sub GRAVA_POSICAOESTOQUE(PROD_ID_N As Long, _
                         Estab_ID_N As Long, _
                         QTDE_SAIDA_VENDA_N As Double, _
                         QTDE_SAIDA_INVENTARIO_N As Double, _
                         QTDE_SAIDA_TRANSFERENCIA_N As Double, _
                         QTDE_ENTRADA_NOTA_N As Double, _
                         QTDE_ENTRADA_INVENTARIO_N As Double, _
                         QTDE_ENTRADA_TRANSFERENCIA_N As Double)
'On Error GoTo ERRO_TRATA
set
   Dim strsql        As String
   Dim QTDE_ATUAL_N  As Double
   Dim TabPosicao    As New ADODB.Recordset

   QTDE_ATUAL_N = TRAZ_QTDE_ESTOQUE(Estab_ID_N, PROD_ID_N)

      If TabPosicao.State = 1 Then _
         TabPosicao.Close

      strsql = "select * from POSICAOESTOQUE WITH (NOLOCK)"
      strsql = strsql & " where produto_id = " & PROD_ID_N
      strsql = strsql & " and estabelecimento_id = " & Estab_ID_N
      TabPosicao.Open strsql, CONECTA_RETAGUARDA, , , adCmdText
      If TabPosicao.EOF Then
         strsql = "insert into POSICAOESTOQUE "
            strsql = strsql & "(PRODUTO_ID , ESTABELECIMENTO_ID, QTDE_ATUAL, "
            strsql = strsql & "QTDE_SAIDA_VENDA, QTDE_SAIDA_INVENTARIO, QTDE_SAIDA_TRANSFERENCIA, "
            strsql = strsql & "QTDE_ENTRADA_NOTA, QTDE_ENTRADA_INVENTARIO, QTDE_ENTRADA_TRANSFERENCIA)"
         strsql = strsql & " values("

            strsql = strsql & PROD_ID_N                                    'PRODUTO_ID
            strsql = strsql & "," & Estab_ID_N                             'ESTABELECIMENTO_ID
            strsql = strsql & "," & tpMOEDA(QTDE_ATUAL_N)                  'QTDE_ATUAL
            strsql = strsql & "," & tpMOEDA(QTDE_SAIDA_VENDA_N)            'QTDE_SAIDA_VENDA
            strsql = strsql & "," & tpMOEDA(QTDE_SAIDA_INVENTARIO_N)       'QTDE_SAIDA_INVENTARIO
            strsql = strsql & "," & tpMOEDA(QTDE_SAIDA_TRANSFERENCIA_N)    'QTDE_SAIDA_TRANSFERENCIA
            strsql = strsql & "," & tpMOEDA(QTDE_ENTRADA_NOTA_N)           'QTDE_ENTRADA_NOTA
            strsql = strsql & "," & tpMOEDA(QTDE_ENTRADA_INVENTARIO_N)     'QTDE_ENTRADA_INVENTARIO
            strsql = strsql & "," & tpMOEDA(QTDE_ENTRADA_TRANSFERENCIA_N)  'QTDE_ENTRADA_TRANSFERENCIA

         strsql = strsql & ")"
         Else
            strsql = "update POSICAOESTOQUE set "
               strsql = strsql & "QTDE_ATUAL = " & tpMOEDA(QTDE_ATUAL_N)
               strsql = strsql & ",QTDE_SAIDA_VENDA = QTDE_SAIDA_VENDA + " & tpMOEDA(QTDE_SAIDA_VENDA_N)
               strsql = strsql & ",QTDE_SAIDA_INVENTARIO = QTDE_SAIDA_INVENTARIO + " & tpMOEDA(QTDE_SAIDA_INVENTARIO_N)
               strsql = strsql & ",QTDE_SAIDA_TRANSFERENCIA = QTDE_SAIDA_TRANSFERENCIA + " & tpMOEDA(QTDE_SAIDA_TRANSFERENCIA_N)
               strsql = strsql & ",QTDE_ENTRADA_NOTA = QTDE_ENTRADA_NOTA + " & tpMOEDA(QTDE_ENTRADA_NOTA_N)
               strsql = strsql & ",QTDE_ENTRADA_INVENTARIO = QTDE_ENTRADA_INVENTARIO + " & tpMOEDA(QTDE_ENTRADA_INVENTARIO_N)
               strsql = strsql & ",QTDE_ENTRADA_TRANSFERENCIA = QTDE_ENTRADA_TRANSFERENCIA + " & tpMOEDA(QTDE_ENTRADA_TRANSFERENCIA_N)
            strsql = strsql & " where produto_id = " & PROD_ID_N
            strsql = strsql & " and estabelecimento_id = " & Estab_ID_N
         End If
      If TabPosicao.State = 1 Then _
         TabPosicao.Close

      CONECTA_RETAGUARDA.Execute strsql

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_POSICAOESTOQUE"
End Sub

Sub MONTA_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = ""
   FORMULA_REL = "{POSICAOESTOQUE.estabelecimento_id} = " & ESTABELECIMENTO_ID_N

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      If IsNumeric(cmbFamiliaAUX.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {POSICAOESTOQUE.familiaproduto_id} = " & Trim(cmbFamiliaAUX.Text)

   If Trim(txtProduto.Text) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {POSICAOESTOQUE.produto_id} = " & PRODUTO_ID_N
      Else
         If Trim(txtDescProd.Text) <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {POSICAOESTOQUE.decricao} = '" & Trim(txtDescProd.Text) & "'"
   End If

   If optConsiderar0.Value = True Then _
      FORMULA_REL = FORMULA_REL & " and {POSICAOESTOQUE.QTDE_VENDIDA} > 0 "

   If optSomente0.Value = True Then _
      FORMULA_REL = FORMULA_REL & " and {POSICAOESTOQUE.QTDE_VENDIDA} = 0 "

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_venda_item.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA



Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
