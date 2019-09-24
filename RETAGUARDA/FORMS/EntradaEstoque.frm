VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEntradaEstoque 
   Caption         =   "Entrada Produto Estoque"
   ClientHeight    =   7620
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EntradaEstoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7620
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDtCad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   9720
      TabIndex        =   18
      ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8160
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtValorDig 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3465
      Picture         =   "EntradaEstoque.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   405
   End
   Begin VB.TextBox txtQTDE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Informe a quantidade de venda deste produto."
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      MaxLength       =   29
      TabIndex        =   9
      Top             =   1560
      Width           =   7215
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5835
      TabIndex        =   3
      ToolTipText     =   "Informe a quantidade de venda deste produto."
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
      Enabled         =   0   'False
      Height          =   405
      Left            =   4560
      MaxLength       =   80
      TabIndex        =   6
      Top             =   960
      Width           =   6615
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3060
      Picture         =   "EntradaEstoque.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   405
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11250
      DesignHeight    =   7620
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
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
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Impressão"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":C455
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":C8A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":E71E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":FDD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoque.frx":11DB6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   405
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   0
      TabIndex        =   16
      Top             =   2640
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8493
      _Version        =   393216
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   11400
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblQtde 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   165
      TabIndex        =   14
      Top             =   2040
      Width           =   1410
   End
   Begin VB.Label lblCodgProduto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   765
      TabIndex        =   13
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Preço Compra = "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3945
      TabIndex        =   12
      Top             =   2040
      Width           =   1905
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      X1              =   0
      X2              =   11400
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3750
      TabIndex        =   7
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frmEntradaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private LastRow         As Long ' Ultima linha em que se editou
   Private LastCol         As Long ' ultima coluna em que se editou
   Private ControlVisible  As Boolean

Private Sub Form_Load()
   CRITERIO_A = ""
   txtDtCad.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   LIMPA_TUDO
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         If FORNEC_ID_N > 0 And Trim(txtID.Text) <> "" Then
            Msg = "Confirma atualização do estoque ?"
            PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               FECHA_ENTRADA
               FINANCEIRO_FORM
               LIMPA_TUDO
               txtCNPJCPF.SetFocus
            End If
         End If
      Case "consultar"
         SQL3 = ""
         frmEntradaEstoqueConsulta.Show 1
         txtID.Text = SQL3
         MOSTRA_TUDO
         SQL3 = ""
      Case "print"
      Case "limpar"
         LIMPA_TUDO
         txtCNPJCPF.SetFocus
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   Dim strTemp As String
   Dim dblTemp As Double

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then _
      Exit Sub

   If VALIDA_CNPJCPF(Trim(txtCNPJCPF.Text)) = False Then
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   FORNEC_ID_N = 0

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select * from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabFornecedor.EOF Then
      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      Beep
      MsgBox "CNPJ/CPF não Cadastrado.", vbOKOnly, "Atenção."
      txtCNPJCPF.SetFocus
      Exit Sub
      Else
         txtNome.Text = Trim(TabFornecedor.Fields("descricao").Value)

         If Not IsNull(TabFornecedor!STATUS) Then
            If TabFornecedor!STATUS <> "A" Then
               If TabFornecedor.State = 1 Then _
                  TabFornecedor.Close
               MsgBox "Fornecedor Desativado, Favor Atualizar Cadastro!"
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If
         PESSOA_ID_N = Trim(TabFornecedor.Fields("pessoa_id").Value)
         FORNEC_ID_N = Trim(TabFornecedor.Fields("fornecedor_id").Value)
   End If   'If TabFornecedor.EOF Then
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   txtCNPJCPF.PromptInclude = False
   CRITERIO_A = txtCNPJCPF.Text

   If Trim(CRITERIO_A) <> "" Then
      If Len(txtCNPJCPF.Text) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CRITERIO_A
   End If
   txtCNPJCPF.PromptInclude = True
   txtCNPJCPF.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_LostFocus"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Mask = "###############"
   INDR_FUNCIONARIO = False

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Text)
   txtCNPJCPF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
            txtProduto.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      txtProduto.SetFocus
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
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   PESSOA_ID_N = 0
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      txtCNPJCPF.PromptInclude = True
      Call txtCNPJCPF_LostFocus
      txtProduto.SetFocus
   End If
   CNPJCPF_A = ""
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtProduto.ForeColor = vbBlue
   txtDescricao.ForeColor = vbBlue

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCESSA_DADOS_PRODUTOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = Empty Then
      txtProduto.SetFocus
      Exit Sub
   End If
   QTDE_N = 0 & txtQTDE.Text
   If QTDE_N <= 0 Then _
      txtQTDE.Text = 1

   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE)
   txtQTDE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Len(Trim(txtQTDE.Text)) > 10 Then
         txtProduto.SetFocus
         Exit Sub
      End If
      QTDE_N = 0 & txtQTDE.Text
      If QTDE_N < 0 Then _
         txtQTDE.Text = 1

      txtValor.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
'On Error GoTo ERRO_TRATA

   If Len(Trim(txtQTDE.Text)) >= 10 Then
      txtProduto.SetFocus
      Exit Sub
   End If

   If Trim(txtQTDE.Text) = "" Then
      txtQTDE.Text = 1
      Else
         If IsNumeric(txtQTDE.Text) Then
            QTDE_N = txtQTDE.Text
            If QTDE_N <= 0 Then _
               txtQTDE.Text = 1
         End If
   End If
   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
   txtQTDE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_LostFocus"
End Sub

Private Sub txtvalor_GotFocus()
   txtValor.SelStart = 0
   txtValor.SelLength = Len(txtValor)
   txtValor.BackColor = &HC0FFFF
   txtValor.Text = Format(txtValor.Text, strFormatacao2Digitos)
   VALOR_ITEM_N = 0
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(txtQTDE.Text) <> "" Then
         VALOR_ITEM_N = 0 & txtValor.Text

         GRAVA_TUDO "A"
         LIMPA_BODY

         txtProduto.SetFocus
         KeyAscii = 0
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_KeyPress"
End Sub

Private Sub txtValor_LostFocus()

   txtValor.BackColor = &HFFFFFF
   txtValor.Text = Format(txtValor.Text, strFormatacao2Digitos)

End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorDig.SelStart = 0
   txtValorDig.SelLength = Len(txtValorDig)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_GotFocus"
End Sub

Private Sub txtValorDig_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         OcultarControles
         MSFlexGrid1.SetFocus
      Case vbKeyUp
         OcultarControles
         'move para a cima celula.
         With MSFlexGrid1
            If .Row > 1 Then
                .Row = .Row - 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
      Case vbKeyDown
         OcultarControles
         With MSFlexGrid1
             If .Row + 1 < .Rows Then
                .Row = .Row + 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyDown"
End Sub

Private Sub txtValorDig_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   ' ao pressionar ENTER aceitar a entrada de dados
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      If LastCol > 3 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      Dim QTDE_RETIDO_ESTORNO As Double

      QTDE_RETIDO_ESTORNO = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)

      AtribuiValorCelula
      'ProximaCelula
      OcultarControles

'==========ATUALIZAR GRID colunas
'3 = qtde
'4 = valor venda
'5 = desconto

      QTDE_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 3)
      VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
      VALOR_DESCONTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 5)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)
      'PRECO_CUSTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 8)
      CODG_PRODUTO_A = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 0))
      PRODUTO_ID_N = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 12))

      If QTDE_N > 0 And VALOR_ITEM_N > 0 And VALOR_DESCONTO_N >= 0 And SEQ_ID_N > 0 Then

         MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N) - VALOR_DESCONTO_N), strFormatacao2Digitos)  'total item
         'lucro MSFlexGrid1.TextMatrix(LastRow, 9) = Format(((VALOR_ITEM_N - PRECO_CUSTO_N) * QTDE_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

         If INDR_ESTQ_NEGATIVO = False Then
            QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

            If QTDE_ESTOQUE_N < 0 Then
               Beep
               MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
               txtQTDE.SetFocus
               Exit Sub
            End If
         End If
      End If

      With MSFlexGrid1
         If .Row + 1 < .Rows Then
            .Row = .Row + 1
            '.Col = 0
            Else
               .Row = 1
               '.Col = 0
         End If
      End With
      txtValorDig.Text = ""
      MSFlexGrid1.SetFocus
      LIMPA_BODY
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            'ControlVisible = False
            Else
               If KeyAscii = 8 Or KeyAscii = 44 Then
                  Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyPress"
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    'OcultarControles

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_Click"
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   OcultarControles

   ExibirCelula

   txtProduto.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 0)
   txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 11)
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then
            If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then
               If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then
                  If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then
                     If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) <> "" Then                'codg Produto
                        If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then             'pedido_id
                           If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then          'seq_id
                              If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then       'produto_id
                                 txtProduto.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
                                 txtSeq.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
                                 'EXCLUIR_ITEM Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Case vbKeyF12
         'frmobs.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

   Select Case KeyAscii
      Case vbKeyReturn  ' Editar ao teclar ENTER
         KeyAscii = 0
         ExibirCelula
      Case vbKeyEscape  ' Cancelar ao pressionar ESC
         KeyAscii = 0
         AtribuiValorCelula
      Case 32 To 255    ' Editar ao pressinar qualquer tecla
         ExibirCelula
         With txtValorDig
            If .Visible Then
             .Text = Chr$(KeyAscii)
             .SelStart = Len(.Text) + 1
           End If
         End With
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyPress"
End Sub

Private Sub ExibirCelula()
'On Error GoTo ERRO_TRATA

   Static OK As Boolean

   If MSFlexGrid1.Col >= 3 And MSFlexGrid1.Col <= 5 Then

      ' Se for celula fixa , sair
      If MSFlexGrid1.Col <= MSFlexGrid1.FixedCols - 1 Or MSFlexGrid1.Row <= MSFlexGrid1.FixedRows - 1 Then _
         Exit Sub
   
      If OK Then _
         Exit Sub

      OK = True

      OcultarControles

      LastRow = MSFlexGrid1.Row
      LastCol = MSFlexGrid1.Col

      Select Case LastCol
         Case Else
            txtValorDig.Move MSFlexGrid1.CellLeft - Screen.TwipsPerPixelX, MSFlexGrid1.CellTop + MSFlexGrid1.Top - Screen.TwipsPerPixelY, MSFlexGrid1.CellWidth + Screen.TwipsPerPixelX * 2, MSFlexGrid1.CellHeight + Screen.TwipsPerPixelY * 2
            txtValorDig.Text = MSFlexGrid1.Text

            If Len(MSFlexGrid1.Text) = 0 Then _
               If LastRow > 1 Then _
                  txtValorDig.Text = MSFlexGrid1.TextMatrix(LastRow - 1, LastCol)

            txtValorDig.Visible = True

            If txtValorDig.Visible Then
               txtValorDig.ZOrder
               txtValorDig.SetFocus
            End If
      End Select
   
      ControlVisible = True

      OK = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ExibirCelula"
End Sub

Private Sub ProximaCelula()
'On Error GoTo ERRO_TRATA

   If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      Else
         MSFlexGrid1.Col = 1
         If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
             MSFlexGrid1.Row = MSFlexGrid1.Row + 1
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ProximaCelula"
End Sub

Private Sub AtribuiValorCelula()
'On Error GoTo ERRO_TRATA

   Dim texto As String

   ' atribuir o texto anterior a celula
   Select Case LastCol
      Case 3 To 5
         texto = txtValorDig.Text

         If LastCol = 3 Then
            MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
            Else: MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao2Digitos)
         End If

         VALOR_VAREJO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
         VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, LastCol)

'&H80C0FF = LARANJA
'&H8000000F = CINZA
'&HFF& = VERMELHO
'vbBlack 0x0
'vbRed 0xFF
'vbGreen 0xFF00
'vbYellow 0xFFFF
'vbBlue 0xFF0000
'vbMagenta 0xFF00FF
'vbCyan 0xFFFF00
'vbWhite 0xFFFFFF

         If VALOR_ITEM_N < VALOR_VAREJO_N Then
            MSFlexGrid1.CellForeColor = vbRed
            MSFlexGrid1.CellFontBold = True
            MSFlexGrid1.CellBackColor = &H8000000F
            Else
               If VALOR_ITEM_N = VALOR_VAREJO_N Then
                  MSFlexGrid1.CellForeColor = vbBlack
                  MSFlexGrid1.CellFontBold = True
                  MSFlexGrid1.CellBackColor = vbCyan
                  Else
                     MSFlexGrid1.CellForeColor = vbBlue
                     MSFlexGrid1.CellFontBold = True
                     MSFlexGrid1.CellBackColor = vbWhite
               End If
         End If
      Case Else
         'texto = txtValorDig.Text
         'MSFlexGrid1.TextMatrix(LastRow, LastCol) = texto
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   'Ocultar o controle textbox
   txtValorDig.Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(8).Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Toolbar1.Buttons(9).Visible = True
      Toolbar1.Buttons(8).Visible = True
   End If
   If MULT_EMPRESA_B = True Then _
      Toolbar1.Buttons(9).Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Sub LIMPA_TUDO()
   NOTAENTRADA_ID_N = 0
   PESSOA_ID_N = 0
   FORNEC_ID_N = 0
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtID.Text = ""
   txtValorDig.Text = ""
   MSFlexGrid1.Clear
   LIMPA_BODY
End Sub

Sub LIMPA_BODY()
   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtQTDE.Text = ""
   txtValor.Text = ""
   txtSeq.Text = ""
   PRODUTO_ID_N = 0
   QTDE_N = 0
   VALOR_ITEM_N = 0
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If Trim(txtID.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtID.Text) Then _
      Exit Sub

   CONT_N = 0

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear

   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   ' define linhas fixas igual a uma e não usa colunas fixas
   MSFlexGrid1.Rows = 2
   'MSFlexGrid1.FixedRows = 3
   MSFlexGrid1.FixedCols = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select ENTRADAESTOQUEITEM.ENTRADAESTOQUE_ID, ENTRADAESTOQUE.estabelecimento_id, ENTRADAESTOQUE.fornecedor_id,"
   SQL = SQL & " ENTRADAESTOQUE.usuario_id, ENTRADAESTOQUE.dt_cadastro, ENTRADAESTOQUE.situacao, ENTRADAESTOQUEITEM.ENTRADAESTOQUEITEM_ID,"
   SQL = SQL & " ENTRADAESTOQUEITEM.PRODUTO_ID, PRODUTO.CODG_PRODUTO as Código, PRODUTO.DESCRICAO, ENTRADAESTOQUEITEM.QTDE, "
   SQL = SQL & " ENTRADAESTOQUEITEM.PRECO as PreçoCompra "

   SQL = SQL & " from ENTRADAESTOQUE WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ENTRADAESTOQUEITEM WITH (NOLOCK) "
   SQL = SQL & " ON ENTRADAESTOQUE.ENTRADAESTOQUE_ID = ENTRADAESTOQUEITEM.ENTRADAESTOQUE_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON ENTRADAESTOQUEITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where ENTRADAESTOQUE.ENTRADAESTOQUE_id = " & txtID.Text

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      txtDtCad.Text = "" & Trim(TabConsulta.Fields("DT_CADASTRO").Value)
      'txtCNPJCPF.PromptInclude = False
      'txtCNPJCPF.Text = "" & Trim(TabConsulta.Fields("cnpjcpf").Value)
      'txtCNPJCPF.PromptInclude = True
      'txtNome.Text = "" & Trim(TabConsulta.Fields("descricao").Value)

      'FORNEC_ID_N = Trim(TabConsulta.Fields("fornecedor_id").Value)

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabConsulta.Fields.Count

      ReDim largura_coluna(0 To TabConsulta.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabConsulta.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabConsulta.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabConsulta.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabConsulta.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabConsulta.Fields.Count - 1
            'If Coluna = 3 Or Coluna = 7 Then
            If Coluna = 10 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabConsulta.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna = 11 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabConsulta.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabConsulta.Fields(Coluna).Value)
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabConsulta.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabConsulta.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabConsulta.MoveNext
         Linha = Linha + 1
      Loop

      'define a largura das colunas do grid
      For Coluna = 0 To MSFlexGrid1.Cols - 1
         MSFlexGrid1.ColWidth(Coluna) = largura_coluna(Coluna) + 240
      Next Coluna

      MSFlexGrid1.ColWidth(0) = 0
      MSFlexGrid1.Refresh

      MSFlexGrid1.BackColor = vbWhite
      MSFlexGrid1.ForeColor = vbBlue

      'ENTRADAESTOQUE.ENTRADAESTOQUE_ID
         MSFlexGrid1.ColWidth(0) = 0

      'ENTRADAESTOQUE.ESTABELECIMENTO_ID
         MSFlexGrid1.ColWidth(1) = 0

      'ENTRADAESTOQUE.FORNECEDOR_ID
         MSFlexGrid1.ColWidth(2) = 0

      'ENTRADAESTOQUE.USUARIO_ID
         MSFlexGrid1.ColWidth(3) = 0

      'ENTRADAESTOQUE.DT_CADASTRO
         MSFlexGrid1.ColWidth(4) = 0

      'ENTRADAESTOQUE.SITUACAO
         MSFlexGrid1.ColWidth(5) = 0

      'ENTRADAESTOQUEITEM.ENTRADAESTOQUEITEM_ID
         MSFlexGrid1.ColWidth(6) = 0

      'ENTRADAESTOQUEITEM.PRODUTO_ID
         MSFlexGrid1.ColWidth(7) = 0
'===================
      'PRODUTO.CODG_PRODUTO
         MSFlexGrid1.ColWidth(8) = 2000
         MSFlexGrid1.ColAlignment(8) = 0

      'PRODUTO.DESCRICAO
         MSFlexGrid1.ColWidth(9) = 8000
         MSFlexGrid1.ColAlignment(9) = 0

      'ENTRADAESTOQUEITEM.QTDE
         MSFlexGrid1.ColWidth(10) = 2000
         MSFlexGrid1.ColAlignment(10) = 7

      'ENTRADAESTOQUEITEM.PRECO
         MSFlexGrid1.ColWidth(11) = 3000
         MSFlexGrid1.ColAlignment(11) = 7
'===================
      'FORNECEDOR.PESSOA_ID
         'MSFlexGrid1.ColWidth(12) = 0

      'FORNECEDOR.cnpjcpf
         'MSFlexGrid1.ColWidth(13) = 0

      'FORNECEDOR.NOME
         'MSFlexGrid1.ColWidth(14) = 0
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GRAVA_TUDO(SITUACAO_A As String)
'On Error GoTo ERRO_TRATA

   txtDtCad.Text = Format(Date, "dd/mm/yyyy")

   If Trim(txtID.Text) = "" Then _
      txtID.Text = MAX_ID("ENTRADAESTOQUE_ID", "ENTRADAESTOQUE", "", "", "", "")

   If FORNEC_ID_N <= 0 Then
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If PRODUTO_ID_N <= 0 Then
      txtProduto.SetFocus
      Exit Sub
   End If
   If QTDE_N <= 0 Then
      txtQTDE.SetFocus
      Exit Sub
   End If
   If QTDE_N <= 0 Then
      txtQTDE.SetFocus
      Exit Sub
   End If

   SQL3 = "NULL"
   If Trim(Left(SITUACAO_A, 1)) = "E" Then _
      SQL3 = Date

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select dt_baixa,situacao from ENTRADAESTOQUE WITH (NOLOCK)"
   SQL = SQL & " where ENTRADAESTOQUE_id = " & txtID.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCabeca.EOF Then
      SQL = "insert into ENTRADAESTOQUE "
         SQL = SQL & " (ENTRADAESTOQUE_ID,ESTABELECIMENTO_ID,FORNECEDOR_ID,USUARIO_ID,DT_CADASTRO,SITUACAO)"
      SQL = SQL & " values ("
         SQL = SQL & txtID.Text                          'ENTRADAESTOQUE_ID
         SQL = SQL & "," & ESTABELECIMENTO_ID_N          'ESTABELECIMENTO_ID
         SQL = SQL & "," & FORNEC_ID_N                   'FORNECEDOR_ID
         SQL = SQL & "," & USUARIO_ID_N                  'USUARIO_ID
         SQL = SQL & ",'" & DMA(txtDtCad.Text) & "'"     'DT_CADASTRO
         SQL = SQL & ",'A'"                              'SITUACAO
      SQL = SQL & ")"
      Else
         If Trim(TabCabeca.Fields("situacao").Value) = "E" Then
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            MsgBox "Entrada já realizada."
            LIMPA_TUDO
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
         SQL = "update ENTRADAESTOQUE SET"
            SQL = SQL & " situacao = 'A'"  'SITUACAO
            SQL = SQL & ",USUARIO_ID = " & USUARIO_ID_N              'USUARIO_ID
            SQL = SQL & ",FORNECEDOR_ID = " & FORNEC_ID_N            'FORNECEDOR_ID
            SQL = SQL & ",DT_BAIXA = " & SQL3                        'DT_BAIXA
         SQL = SQL & " where ENTRADAESTOQUE_id = " & txtID.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL

   If PRODUTO_ID_N > 0 And QTDE_N > 0 Then _
      GRAVA_ITEM

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub GRAVA_ITEM()
'On Error GoTo ERRO_TRATA

   If TabItem.State = 1 Then _
      TabItem.Close

   SQL = "select ENTRADAESTOQUE_ID from ENTRADAESTOQUEITEM WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and ENTRADAESTOQUE_id = " & txtID.Text
   SQL = SQL & " and ENTRADAESTOQUEITEM_ID = " & txtSeq.Text
   TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabItem.EOF Then
      SQL = "insert into ENTRADAESTOQUEITEM "
         SQL = SQL & " (ENTRADAESTOQUE_ID,ENTRADAESTOQUEITEM_ID,PRODUTO_ID,PRECO,QTDE) "
      SQL = SQL & " values ("
        SQL = SQL & txtID.Text                     'ENTRADAESTOQUE_ID
        SQL = SQL & "," & Trim(txtSeq.Text)        'ENTRADAESTOQUEITEM_ID
        SQL = SQL & "," & PRODUTO_ID_N             'PRODUTO_ID
        SQL = SQL & "," & tpMOEDA(txtValor.Text)   'PRECO
        SQL = SQL & "," & tpMOEDA(txtQTDE.Text)    'QTDE
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabItem.State = 1 Then _
      TabItem.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITEM"
End Sub

Sub FECHA_ENTRADA()
'On Error GoTo ERRO_TRATA

   If TabItem.State = 1 Then _
      TabItem.Close

   SQL = "select produto_id,qtde,ENTRADAESTOQUE_id from ENTRADAESTOQUEITEM WITH (NOLOCK)"
   SQL = SQL & " where ENTRADAESTOQUE_id = " & txtID.Text
   TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabItem.EOF Then
      NOTAENTRADA_ID_N = TabItem.Fields("ENTRADAESTOQUE_id").Value
      SQL = "update ENTRADAESTOQUE SET"
         SQL = SQL & " situacao = 'E'"                   'SITUACAO
         SQL = SQL & ",USUARIO_ID = " & USUARIO_ID_N     'USUARIO_ID
         SQL = SQL & ",FORNECEDOR_ID = " & FORNEC_ID_N   'FORNECEDOR_ID
         SQL = SQL & ",DT_BAIXA = '" & Now & "'"         'DT_BAIXA
      SQL = SQL & " where ENTRADAESTOQUE_id = " & txtID.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If
   While Not TabItem.EOF
      '================estoque
      SQL = "update ESTOQUE set "
      SQL = SQL & " QTDE_ESTOQUE = QTDE_ESTOQUE + " & tpMOEDA(TabItem.Fields("QTDe").Value)

      SQL = SQL & " from EMPRESA "
      SQL = SQL & " INNER JOIN ESTABELECIMENTO "
      SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID "
      SQL = SQL & " INNER JOIN ESTOQUE "
      SQL = SQL & " ON ESTABELECIMENTO.ESTABELECIMENTO_ID = ESTOQUE.ESTABELECIMENTO_ID"

      SQL = SQL & " where produto_id = " & TabItem.Fields("PRODUTO_ID").Value
      SQL = SQL & " and ESTOQUE.estabelecimento_id = " & ESTABELECIMENTO_ID_N

      CONECTA_RETAGUARDA.Execute SQL
      '=======================

      SQL = "update produto set dt_ult_compra = '" & Now & "'"
      SQL = SQL & " where produto_id = " & TabItem.Fields("PRODUTO_ID").Value
      CONECTA_RETAGUARDA.Execute SQL

      TabItem.MoveNext
   Wend
   If TabItem.State = 1 Then _
      TabItem.Close

   MsgBox "Processo realizado com sucesso."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_ENTRADA"
End Sub

Sub MOSTRA_TUDO()
'On Error GoTo ERRO_TRATA

   If Trim(txtID.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtID.Text) Then _
      Exit Sub

   Dim rstConsulta   As New ADODB.Recordset

   If rstConsulta.State = 1 Then _
      rstConsulta.Close

   SQL = "select ENTRADAESTOQUE.*, "
   SQL = SQL & " vwFornecedor.PESSOA_ID, vwFornecedor.DESCRICAO, vwFornecedor.NUMR_IE, "
   SQL = SQL & " vwFornecedor.cnpjcpf, vwFornecedor.descricao "

   SQL = SQL & " from ENTRADAESTOQUE WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN vwFornecedor WITH (NOLOCK) "
   SQL = SQL & " ON ENTRADAESTOQUE.FORNECEDOR_ID = vwFornecedor.FORNECEDOR_ID"

   SQL = SQL & " where ENTRADAESTOQUE_id = " & txtID.Text
   SQL = SQL & " and ENTRADAESTOQUE.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   rstConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not rstConsulta.EOF Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = "" & Trim(rstConsulta.Fields("cnpjcpf").Value)
      txtNome.Text = "" & Trim(rstConsulta.Fields("descricao").Value)
      txtID.Text = "" & Trim(rstConsulta.Fields("entradaestoque_id").Value)
      txtDtCad.Text = "" & Trim(rstConsulta.Fields("dt_cadastro").Value)
      FORNEC_ID_N = 0 & Trim(rstConsulta.Fields("fornecedor_id").Value)
      PESSOA_ID_N = Trim(rstConsulta.Fields("pessoa_id").Value)

      SETA_GRID

      If Trim(rstConsulta.Fields("situacao").Value) = "E" Then
         If rstConsulta.State = 1 Then _
            rstConsulta.Close

         MsgBox "Entrada já realizada."
         LIMPA_TUDO
         txtCNPJCPF.SetFocus
         Exit Sub
      End If

      txtProduto.SetFocus
   End If
   If rstConsulta.State = 1 Then _
      rstConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_ENTRADA"
End Sub

Private Sub FINANCEIRO_FORM()
'On Error GoTo ERRO_TRATA

   If NOTAENTRADA_ID_N > 0 Then
      INDR_RECEITA = 2
      TIPO_ENTRADA_N = 2
      frmNOTAENTRADAFINANC.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FINANCEIRO_FORM"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto)
      txtProduto.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If

   txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtProduto.Text = Trim(CODG_PRODUTO_A)
   txtDescricao.Text = DESC_PRODUTO_A
   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            txtProduto.Enabled = True
            txtProduto.SetFocus
            Exit Sub
         End If
   End If
   txtValor.Text = "" & Format(PR_VAREJO_N, strFormatacao2Digitos)
   If PR_VAREJO_N < 0 Then
      MsgBox "Valor do produto invalido !!!"
      Exit Sub
   End If

   QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

   If Not IsNull(CODG_NCM_A) Then
      If Len(CODG_NCM_A) > 2 Then
         If Len(CODG_NCM_A) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"

            LIMPA_BODY

            txtProduto.Enabled = True
            txtProduto.SetFocus
         End If
      End If
   End If
'=====================
   If Trim(txtID.Text) = "" Then _
      txtID.Text = MAX_ID("ENTRADAESTOQUE_ID", "ENTRADAESTOQUE", "", "", "", "")

   If Trim(txtID.Text) = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("ENTRADAESTOQUEITEM_id", "ENTRADAESTOQUEITEM", "ENTRADAESTOQUE_id", Trim(txtID.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("ENTRADAESTOQUEITEM_id", "ENTRADAESTOQUEITEM", "ENTRADAESTOQUE_id", Trim(txtID.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then
      txtValor.SetFocus
      Call txtvalor_KeyPress(13)
      txtProduto.SetFocus
      Else
         If Trim(txtProduto.Text) <> "" Then _
            txtQTDE.SetFocus
   End If
   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
