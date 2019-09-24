VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEntradaEstoqueConsulta 
   Caption         =   "Consulta Entrada Estoque"
   ClientHeight    =   7695
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EntradaEstoqueConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbSituacao 
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
      Left            =   9120
      TabIndex        =   16
      ToolTipText     =   "Selecione a situação para este produto"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdFornec 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3600
      Picture         =   "EntradaEstoqueConsulta.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1485
      Width           =   405
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3600
      Picture         =   "EntradaEstoqueConsulta.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1965
      Width           =   405
   End
   Begin VB.TextBox txtPreco 
      Alignment       =   2  'Center
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
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   13
      ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
      Top             =   2475
      Width           =   1335
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
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
      Left            =   1755
      MaxLength       =   6
      TabIndex        =   2
      ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
      Top             =   975
      Width           =   1815
   End
   Begin VB.TextBox txtFornecDesc 
      DataField       =   "Nome"
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
      Left            =   4080
      MaxLength       =   80
      TabIndex        =   12
      Top             =   1440
      Width           =   7455
   End
   Begin VB.TextBox txtProdutoDesc 
      DataField       =   "Nome"
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
      Left            =   4080
      MaxLength       =   80
      TabIndex        =   11
      Top             =   1920
      Width           =   7455
   End
   Begin VB.TextBox txtQtde 
      Alignment       =   1  'Right Justify
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
      Left            =   1755
      MaxLength       =   6
      TabIndex        =   10
      ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
      Top             =   2475
      Width           =   1815
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
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
      Left            =   1755
      MaxLength       =   15
      TabIndex        =   4
      ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
      Top             =   1965
      Width           =   1815
   End
   Begin VB.TextBox txtTotalPedido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   10080
      TabIndex        =   9
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtQtdeItensPedidos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   5745
      TabIndex        =   8
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtQtdeItensRelacionados 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1545
      TabIndex        =   7
      Top             =   7200
      Width           =   1455
   End
   Begin VB.OptionButton optInclusao 
      Caption         =   "Dt.Inclusão"
      Height          =   240
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optBaixa 
      Caption         =   "Dt.&Entrada"
      Height          =   240
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11595
      DesignHeight    =   7695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      DisabledImageList=   "ImageList"
      HotImageList    =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   7680
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoqueConsulta.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoqueConsulta.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoqueConsulta.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoqueConsulta.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoqueConsulta.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EntradaEstoqueConsulta.frx":D2E1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6840
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox txtFornec 
      Height          =   390
      Left            =   1755
      TabIndex        =   3
      Top             =   1485
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   390
      Left            =   7920
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Height          =   375
      Left            =   10200
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.TreeView lstPedido 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   3000
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      ImageList       =   "ILTw"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Situação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   8040
      TabIndex        =   30
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblpreco 
      Alignment       =   1  'Right Justify
      Caption         =   "Preço Mercadoria="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4110
      TabIndex        =   29
      Top             =   2520
      Width           =   2025
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Registro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   28
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fornecedor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   390
      TabIndex        =   27
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Inicial:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7140
      TabIndex        =   26
      Top             =   960
      Width           =   675
   End
   Begin VB.Label lblproduto 
      Alignment       =   1  'Right Justify
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   780
      TabIndex        =   25
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label lblquantidade 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   405
      TabIndex        =   24
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "ValorTotalPedidos = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   7590
      TabIndex        =   23
      Top             =   7200
      Width           =   2205
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "QtdeItensPedidos = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   3540
      TabIndex        =   22
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "QtdeItens = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   21
      Top             =   7200
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Final:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9525
      TabIndex        =   20
      Top             =   960
      Width           =   570
   End
End
Attribute VB_Name = "frmEntradaEstoqueConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         SETA_GRID
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
      Case "print"
         FORMULA_REL = "{ENTRADAESTOQUE.ENTRADAESTOQUE_id} > 0 "

         If Trim(cmbSituacao.Text) <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.situacao} = " & Left(cmbSituacao.Text, 1)

         If Trim(txtPedido.Text) <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.ENTRADAESTOQUE_id} = " & Trim(txtPedido.Text)

         txtDtIni.PromptInclude = True
         txtDtFim.PromptInclude = True
         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
            If optInclusao.Value = True Then
               FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.dt_cadastro} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
               FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.dt_cadastro} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
            End If
            If optBaixa.Value = True Then
               FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.dt_baixa} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
               FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.dt_baixa} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
            End If
         End If

         If FORNEC_ID_N > 0 Then _
            FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUE.fornecedor_id} = " & FORNEC_ID_N

         If PRODUTO_ID_N > 0 Then _
            FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUEitem.produto_id} = " & PRODUTO_ID_N

         If Trim(txtQTDE.Text) <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUEitem.qtde} = " & txtQTDE.Text
   
         If Trim(txtPreco.Text) <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {ENTRADAESTOQUEitem.preco} = " & txtPreco.Text

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "Pedido_Compra.rpt"
         frmRELATORIO10.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstPedido_DblClick()
   SQL3 = Mid(lstPedido.SelectedItem.key, 2, Len(lstPedido.SelectedItem.key) - 1)
   Unload Me
End Sub

Private Sub optInclusao_Click()
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False

      txtDtIni.Text = Date
      txtDtFim.Text = Date

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   txtDtIni.SetFocus
End Sub

Private Sub optBaixa_Click()
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False

      txtDtIni.Text = Date
      txtDtFim.Text = Date

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   txtDtIni.SetFocus
End Sub

Private Sub cmdFornec_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   CRITERIO_A = ""
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtFornec.PromptInclude = False
      txtFornec.Text = CNPJCPF_A
   End If
   txtFornec.SetFocus
   CNPJCPF_A = ""
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdFornec_Click"
End Sub

Private Sub txtfornec_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         CRITERIO_A = ""
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtFornec.PromptInclude = False
            txtFornec.Text = CNPJCPF_A
         End If
         txtFornec.SetFocus
         CNPJCPF_A = ""
         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFornec_KeyDown"
End Sub

Private Sub txtFornec_LostFocus()
'On Error GoTo ERRO_TRATA

   txtFornec.PromptInclude = False
   If Trim(txtFornec.Text) <> "" Then
      If VALIDA_CNPJCPF(Trim(txtFornec.Text)) = False Then _
         Exit Sub

      CRITERIO_A = txtFornec.Text

      If Trim(CRITERIO_A) <> "" Then
         If Len(txtFornec.Text) <= 11 Then
            txtFornec.Mask = "###.###.###-##"
            Else: txtFornec.Mask = "##.###.###/####-##"
         End If
         txtFornec.Text = CRITERIO_A
      End If
      FORNEC_ID_N = 0

      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      SQL = "select descricao,fornecedor_id from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & Trim(CRITERIO_A) & "'"
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabFornecedor.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção."
         txtFornec.SetFocus
         Exit Sub
         Else
            txtFornecDesc.Text = Trim(TabFornecedor.Fields("descricao").Value)
            FORNEC_ID_N = Trim(TabFornecedor.Fields("fornecedor_id").Value)
      End If
      CRITERIO_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFornec_LostFocus"
End Sub

Private Sub cmdConsProd_Click()
   CONSULTA_PRODUTO
   txtProduto.SetFocus
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtProduto.Text = "" Then
      txtProduto.Text = 0
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_GotFocus"
End Sub

Private Sub TXTPRODUTO_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      PRODUTO_ID_N = 0

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select descricao,produto_id from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProduto.EOF Then
         MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)
         txtProduto.SetFocus
         Else
            txtProdutoDesc.Text = Trim(TabProduto!DESCRICAO)
            PRODUTO_ID_N = TabProduto.Fields("produto_id").Value
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_LostFocus"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CONSULTA_PRODUTO
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then _
      txtProduto.Text = SQL3
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Sub LIMPA_TUDO()
   FORNEC_ID_N = 0
   PRODUTO_ID_N = 0
   lstPedido.Nodes.Clear
   txtPedido.Text = ""
   txtDtIni.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtFornec.Text = ""
   txtFornecDesc.Text = ""
   txtProduto.Text = ""
   txtProdutoDesc.Text = ""
   txtQTDE.Text = ""
   txtPreco.Text = ""
   cmbSituacao.Text = ""
   txtQtdeItensRelacionados.Text = ""
   txtQtdeItensPedidos.Text = ""
   txtTotalPedido.Text = ""
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim TabConsulta         As New ADODB.Recordset
   lstPedido.Nodes.Clear
   lstPedido.Visible = False
   NUMR_ID_N = 0
   CRITERIO_A = ""
   CONT_N = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select ENTRADAESTOQUEITEM.ENTRADAESTOQUE_ID, ENTRADAESTOQUE.ESTABELECIMENTO_ID, ENTRADAESTOQUE.FORNECEDOR_ID, "
   SQL = SQL & " ENTRADAESTOQUE.USUARIO_ID, ENTRADAESTOQUE.DT_CADASTRO, ENTRADAESTOQUE.SITUACAO, ENTRADAESTOQUEITEM.ENTRADAESTOQUEITEM_ID,"
   SQL = SQL & " ENTRADAESTOQUEITEM.PRODUTO_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, ENTRADAESTOQUEITEM.QTDE, ENTRADAESTOQUEITEM.PRECO,"
   SQL = SQL & " PESSOA.DESCRICAO AS NomePessoa, FORNECEDOR.PESSOA_ID"
   SQL = SQL & " from ENTRADAESTOQUE WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ENTRADAESTOQUEITEM WITH (NOLOCK)"
   SQL = SQL & " ON ENTRADAESTOQUE.ENTRADAESTOQUE_ID = ENTRADAESTOQUEITEM.ENTRADAESTOQUE_ID"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON ENTRADAESTOQUEITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " INNER JOIN FORNECEDOR  WITH (NOLOCK)"
   SQL = SQL & " ON ENTRADAESTOQUE.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"
   SQL = SQL & " INNER JOIN PESSOA  WITH (NOLOCK)"
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where ENTRADAESTOQUE.ENTRADAESTOQUE_ID > 0 "

   If Trim(txtPedido.Text) <> "" Then _
      SQL = SQL & " and ENTRADAESTOQUE.ENTRADAESTOQUE_ID = " & txtPedido.Text

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optInclusao.Value = True Then
         SQL = SQL & " and ENTRADAESTOQUE.dt_cadastro >= '" & Format(txtDtIni.Text, "dd/mm/yyyy") & "'"
         SQL = SQL & " and ENTRADAESTOQUE.dt_cadastro <= '" & Format(txtDtFim.Text, "dd/mm/yyyy") & "'"
      End If
      If optBaixa.Value = True Then
         SQL = SQL & " and ENTRADAESTOQUE.dt_baixa >= '" & Format(txtDtIni.Text, "dd/mm/yyyy") & "'"
         SQL = SQL & " and ENTRADAESTOQUE.dt_baixa <= '" & Format(txtDtFim.Text, "dd/mm/yyyy") & "'"
      End If
   End If

   If FORNEC_ID_N > 0 Then _
      SQL = SQL & " and ENTRADAESTOQUE.fornecedor_id = " & FORNEC_ID_N

   If PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and ENTRADAESTOQUEitem.produto_id = " & PRODUTO_ID_N

   If Trim(txtQTDE.Text) <> "" Then _
      SQL = SQL & " and ENTRADAESTOQUEitem.qtde = " & txtQTDE.Text
   
   If Trim(txtPreco.Text) <> "" Then _
      SQL = SQL & " and ENTRADAESTOQUEitem.preco = " & txtPreco.Text
   
   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and ENTRADAESTOQUE.situacao = " & Left(cmbSituacao.Text, 1)

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      DoEvents
      CONT_N = CONT_N + 1
      CRITERIO_A = ""

      If NUMR_ID_N <> TabConsulta.Fields("ENTRADAESTOQUE_ID").Value Then
         Set Nodx = lstPedido.Nodes.Add(, , "C" & TabConsulta.Fields("ENTRADAESTOQUE_ID").Value, "Pedido Compra: " & TabConsulta.Fields("ENTRADAESTOQUE_ID").Value & " ; Dt.: " & TabConsulta.Fields("dt_cadastro").Value & " - " & Trim(TabConsulta.Fields("nomepessoa").Value))
         NUMR_ID_N = TabConsulta.Fields("ENTRADAESTOQUE_ID").Value
      End If

      CRITERIO_A = "Produto: "
      CRITERIO_A = CRITERIO_A & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)

      CRITERIO_A = CRITERIO_A & " ; " & "Qtde = " & Format(TabConsulta.Fields("qtde").Value, strFormatacao3Digitos)
      CRITERIO_A = CRITERIO_A & " ; " & "Preço = " & Format(TabConsulta.Fields("preco").Value, strFormatacao2Digitos)

      Set Nodx = lstPedido.Nodes.Add("C" & TabConsulta.Fields("ENTRADAESTOQUE_ID").Value, tvwChild, "itens" & CONT_N, CRITERIO_A)

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   lstPedido.Visible = True
   NUMR_ID_N = 0
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub


