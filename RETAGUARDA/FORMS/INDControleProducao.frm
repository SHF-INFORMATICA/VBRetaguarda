VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmINDControleProducao 
   Caption         =   "Controle de Produção"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INDControleProducao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtQtdeItens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   240
      TabIndex        =   21
      Top             =   7020
      Width           =   1455
   End
   Begin VB.TextBox txtPesoTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   525
      Left            =   9360
      TabIndex        =   20
      Top             =   7020
      Width           =   1455
   End
   Begin VB.Frame FraReq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   50
      TabIndex        =   15
      Top             =   720
      Width           =   6855
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDtLote 
         Height          =   360
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         Enabled         =   0   'False
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
      Begin VB.Label lblPedido 
         Alignment       =   1  'Right Justify
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   495
         TabIndex        =   19
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblDtEmis 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3600
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FraSeq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   50
      TabIndex        =   5
      Top             =   1440
      Width           =   11295
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
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4065
         Picture         =   "INDControleProducao.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtQTDE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   9
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
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
         Left            =   4530
         MaxLength       =   29
         TabIndex        =   8
         Top             =   240
         Width           =   6615
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   5355
         TabIndex        =   6
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblQtde 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantidade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   525
         TabIndex        =   14
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblCodgProduto 
         Alignment       =   1  'Right Justify
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   885
         TabIndex        =   13
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4680
         TabIndex        =   12
         Top             =   720
         Width           =   570
      End
   End
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   525
      Left            =   6555
      TabIndex        =   4
      Top             =   7020
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6960
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      Begin VB.OptionButton optSint 
         Caption         =   "Sintético"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
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
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton optAna 
         Caption         =   "Analítico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1270
      ButtonWidth     =   3016
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
            Caption         =   "Fechar Lote"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cad.Produto"
            Key             =   "produto"
            ImageIndex      =   6
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
               Picture         =   "INDControleProducao.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":A8FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":BA53
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":BEA5
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":DD1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":F3D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDControleProducao.frx":113B4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11385
      DesignHeight    =   7620
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   45
      TabIndex        =   23
      Top             =   2760
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
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
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   3840
      X2              =   3840
      Y1              =   6720
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   6120
      X2              =   6120
      Y1              =   6720
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   8880
      X2              =   8880
      Y1              =   6720
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   1800
      X2              =   1800
      Y1              =   6720
      Y2              =   7560
   End
   Begin VB.Label lblItensPedido 
      Alignment       =   1  'Right Justify
      Caption         =   "Qtde.Itens"
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
      Left            =   705
      TabIndex        =   26
      Top             =   6735
      Width           =   960
   End
   Begin VB.Label lblTotKg 
      Alignment       =   1  'Right Justify
      Caption         =   "Peso Total (Kg)"
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
      Left            =   9270
      TabIndex        =   25
      Top             =   6730
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      Height          =   870
      Left            =   120
      Top             =   6720
      Width           =   10935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Total"
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
      Left            =   7215
      TabIndex        =   24
      Top             =   6735
      Width           =   1140
   End
End
Attribute VB_Name = "frmINDControleProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim NUMR_LOTE_N            As Long
   Dim TabFamilia             As New ADODB.Recordset
   Dim VALOR_UNITARIO_N       As Double
   Dim Valr_Venda_Produto_n   As Double
   Dim Qtde_N                 As Double
   Dim PESO_ITEM_N            As Double
   Dim CODIGO_BARRAS          As String
   Dim PRECO_CUSTO_N          As Double
   Private LastRow            As Long ' Ultima linha em que se editou
   Private LastCol            As Long ' ultima coluna em que se editou
   Private ControlVisible     As Boolean

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   CHECA_TABELAS
   txtDtLote = Format(Date, "dd/mm/yyyy")
   Call txtLote_LostFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         frmINDPRODUCAOConsulta.Show 1
         If Trim(CRITERIO) <> "" Then
            txtLote.Text = CRITERIO
            NUMR_LOTE_N = CRITERIO
            PROCURA_REGISTRO txtLote.Text
         End If
         FraSeq.Enabled = True

         txtProduto.Enabled = True
         txtProduto.SetFocus
      Case "print"
         GERA_IMPRESSAO
      Case "limpar"
         LIMPA_TUDO

         Call txtLote_LostFocus
         FraSeq.Enabled = True

         txtProduto.Enabled = True
         txtProduto.SetFocus
      Case "voltar"
         Unload Me
      Case "produto"
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
            frmCADASTROPRODUTO.Show 1
            Else: CHAMA_PRODUTO_SIMPLIFICADO
         End If
      Case "gravar"
         GRAVA_REGISTRO txtLote.Text, Date, "F"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtLote_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtLote.Text) = "" Then _
      NUMR_LOTE_N = MAX_ID("PRODUCAO_ID", "PRODUCAO", "", "", "", "")

   txtLote.Text = NUMR_LOTE_N

   If Trim(txtLote.Text) <> "" Then
      If IsNumeric(txtLote.Text) Then

         NUMR_LOTE_N = txtLote.Text

         PROCURA_REGISTRO txtLote.Text
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLote_LostFocus"
End Sub

Private Sub cmdConsProd_Click()
'On Error GoTo ERRO_TRATA

   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Private Sub txtITENS_GotFocus()
'On Error GoTo ERRO_TRATA

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtITENS_GotFocus"
End Sub

Private Sub txtPesoTotal_GotFocus()
'On Error GoTo ERRO_TRATA

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPesoTotal_GotFocus"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDescricao.Enabled = False
   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto.Text)
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtLote.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then _
            EXCLUIR_ITEM Trim(txtLote.Text), Trim(txtSeq.Text)

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus
      Case vbKeyF7
         CONSULTA_PRODUTO
         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus
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
      LE_PRODUTO
      If Trim(txtProduto.Text) <> "" And Trim(txtLote.Text) <> "" Then _
         txtQTDE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtProduto_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA
   
   If Trim(txtProduto.Text) = Empty Then
   '   MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
   '   txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   Qtde_N = 0 & txtQTDE.Text
   If Qtde_N <= 0 Then _
      txtQTDE.Text = 1

   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE.Text)
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
      Qtde_N = 0 & txtQTDE.Text
      If Qtde_N < 0 Then _
         txtQTDE.Text = 1

      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus
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

   If Trim(txtQTDE.Text) = "" Then
      txtQTDE.Text = 1
      Else
         If IsNumeric(txtQTDE.Text) Then
            Qtde_N = txtQTDE.Text
            If Qtde_N <= 0 Then _
               txtQTDE.Text = 1
         End If
   End If
   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)

   VALOR_ITEM_N = 0 & txtValor.Text
   Qtde_N = 0 & txtQTDE.Text

   GRAVA_REGISTRO txtLote.Text, "", "A"
   txtQTDE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_LostFocus"
End Sub
'============================subrotinas
Private Sub EXCLUIR_ITEM(PRODUCAO_ID_N As Long, SEQ_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Trim(PRODUCAO_ID_N) > 0 And Trim(SEQ_ID_N) > 0 Then
      Msg = "Deseja Excluir Esse Item?"
      Style = vbYesNo + 32
      Title = "Atenção."
      Help = "DEMO.HLP"
      Ctxt = 1000
      RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
      If RESPOSTA = vbYes Then
         SQL = "Delete FROM PRODUCAOITEM "
         SQL = SQL & " Where PRODUCAO_id = " & PRODUCAO_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         LIMPA_BODY
         SETA_GRID
      End If
      Else: MsgBox "Produto não encontrado."
   End If

   FraSeq.Enabled = True
   txtProduto.Enabled = True

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   Aliquota_Icms = 0
   Valr_Venda_Produto_n = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""
   txtQTDE.Text = ""
   txtValor.Text = ""

   QTDE_ESTOQUE = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_DIFERENCA_N = 0
   PRODUTO_ID_N = 0
   NUMR_SEQ_N = 0

   txtQTDE.Text = Format(0, strFormatacao3Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   INDR_ERRO_TEF = False
   INDR_VENDA = False

   If TabUSU.State = 1 Then _
      TabUSU.Close

   CRITERIO = ""

   MSFlexGrid1.Clear

   FraSeq.Enabled = False

   txtPesoTotal.Text = ""
   txtQtdeItens.Text = "" & Format(0, strFormatacao3Digitos)

   optSint.Value = True
   optAna.Value = False
   PRODUTO_ID_N = 0
   Aliquota_Icms = 0
   NUMR_SEQ_N = 0
   txtLote.Text = ""
   txtDtLote = Format(Date, "dd/mm/yyyy")
   txtValorTotal.Text = ""
   LIMPA_BODY
   
   VALOR_TOTAL_N = 0
   NUMR_LOTE_N = 0
   QTDE_ESTOQUE = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_TOTAL_N = 0
   USU_LIBERA_VENDA_N = 0
   INDR_RECEITA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo
   Dim TabGridVaca   As New ADODB.Recordset

   MSFlexGrid1.Clear
   PESO_ITEM_N = 0
   CONT_N = 0
   VALOR_ITEM_N = 0

   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   If TabGridVaca.State = 1 Then _
      TabGridVaca.Close

   SQL = "SELECT PRODUCAOITEM.SEQ_ID as Seq, PRODUTO.CODG_PRODUTO as Codg, "
   SQL = SQL & " PRODUTO.DESCRICAO as Descrição, PRODUCAOITEM.QTDE, PRODUCAOITEM.valor as Valor,"
   SQL = SQL & " FAMILIAPRODUTO.DESCRICAO AS DescFamilia, FAMILIAPRODUTO.PRODUCAO, "
   SQL = SQL & " PRODUCAOITEM.PRODUTO_ID, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " FAMILIAPRODUTO.CODG_FAMILIA, PRODUCAOITEM.PRODUCAO_ID, "
   SQL = SQL & " PRODUTO.PESO_LIQUIDO"
   SQL = SQL & " FROM PRODUCAOITEM "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PRODUCAOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " INNER JOIN PRODUCAO "
   SQL = SQL & " ON PRODUCAOITEM.PRODUCAO_ID = PRODUCAO.PRODUCAO_ID"

   SQL = SQL & " where PRODUCAO.PRODUCAO_id = " & txtLote.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " order by seq_id desc"

   TabGridVaca.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGridVaca.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabGridVaca.Fields.Count

      ReDim largura_coluna(0 To TabGridVaca.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabGridVaca.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabGridVaca.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabGridVaca.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabGridVaca.EOF
'========= verificando se o produto é de produção
         INDR_PRI = False
         CONT_N = CONT_N + 1

         If Not IsNull(TabGridVaca.Fields("producao").Value) Then
            INDR_PRI = TabGridVaca.Fields("producao").Value
            PESO_ITEM_N = TabGridVaca.Fields("qtde").Value + PESO_ITEM_N
            Else
               If Not IsNull(TabGridVaca.Fields("peso_liquido").Value) Then _
                  PESO_ITEM_N = TabGridVaca.Fields("peso_liquido").Value + PESO_ITEM_N
         End If
         txtPesoTotal.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
         txtQtdeItens.Text = CONT_N

         If Not IsNull(TabGridVaca.Fields("valor").Value) Then _
            VALOR_ITEM_N = TabGridVaca.Fields("valor").Value + VALOR_ITEM_N

         txtValorTotal.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
'=========

         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabGridVaca.Fields.Count - 1
            If Coluna = 3 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGridVaca.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna = 4 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGridVaca.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabGridVaca.Fields(Coluna).Value)
                  End If
            End If

'=========se o produto for de produção pintar linha
            If INDR_PRI = True Then
               MSFlexGrid1.Row = Linha
               MSFlexGrid1.Col = Coluna
               MSFlexGrid1.CellForeColor = &H4000&   '&H40&
            End If
'=========

            ' verifica o tamanho dos campos
            If Not IsNull(TabGridVaca.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGridVaca.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabGridVaca.MoveNext
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

      'seq_id
      MSFlexGrid1.ColWidth(0) = 500
      MSFlexGrid1.ColAlignment(0) = 0

      'Codigo Produto
      MSFlexGrid1.ColWidth(1) = 2000
      MSFlexGrid1.ColAlignment(1) = 0

      'Descrição Produto
      MSFlexGrid1.ColWidth(2) = 7000
      MSFlexGrid1.ColAlignment(2) = 0

      'QTDE
      MSFlexGrid1.ColWidth(3) = 2000
      MSFlexGrid1.ColAlignment(3) = 7

      'valor
      MSFlexGrid1.ColWidth(4) = 2000
      MSFlexGrid1.ColAlignment(4) = 7

      'descrição familia
      MSFlexGrid1.ColWidth(5) = 2000
      MSFlexGrid1.ColAlignment(5) = 7

      'produto_id
      MSFlexGrid1.ColWidth(6) = 0
      MSFlexGrid1.ColAlignment(6) = 0

      'familiaproduto_id
      MSFlexGrid1.ColWidth(7) = 0
      MSFlexGrid1.ColAlignment(7) = 0

      'familiaproduto_id
      MSFlexGrid1.ColWidth(8) = 0
      MSFlexGrid1.ColAlignment(8) = 0

      'FAMILIAPRODUTO.PRODUCAO
      MSFlexGrid1.ColWidth(9) = 0
      MSFlexGrid1.ColAlignment(9) = 0

      '
      MSFlexGrid1.ColWidth(10) = 0
      MSFlexGrid1.ColAlignment(10) = 0

      '
      MSFlexGrid1.ColWidth(11) = 0
      MSFlexGrid1.ColAlignment(11) = 0
   End If
   ' fecha o recordset e a conexao
   If TabGridVaca.State = 1 Then _
      TabGridVaca.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub GERA_IMPRESSAO()
'On Error GoTo ERRO_TRATA

   If Trim(txtLote.Text) <> "" Then
      NUMR_LOTE_N = txtLote.Text
      Else: NUMR_LOTE_N = InputBox(SQL3, "Informe número de PRODUCAO a ser impressa ")
   End If

   If NUMR_LOTE_N > 0 Then
      FORMULA_REL = "{PRODUCAO.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
      FORMULA_REL = FORMULA_REL & " and {PRODUCAO.PRODUCAO_id} = " & NUMR_LOTE_N
   End If

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If optSint.Value = True Then
      Nome_Relatorio = "REL_PERDA_SINT.rpt"
      Else: Nome_Relatorio = "REL_PERDA_ANALIT.rpt"
   End If

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_IMPRESSAO"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      FraSeq.Enabled = True
      txtProduto.Enabled = True

      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value
   txtProduto.Text = Trim(TabProduto.Fields("codg_produto").Value)
   STATUS_PROD = TabProduto!SITUACAO
   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      txtProduto.Text = TabProduto!Codg_Produto
      txtDescricao.Text = TabProduto!DESCRICAO
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            FraSeq.Enabled = True
            txtProduto.Enabled = True

            txtProduto.SetFocus
            Exit Sub
            Else: txtDescricao.Text = Trim(TabProduto!DESCRICAO)
         End If
   End If

   STATUS_PROD = TabProduto!SITUACAO

   If Not IsNull(TabProduto!PRECO_VENDA) Then
      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_VENDA

      txtValor.Text = "" & Format(Valr_Venda_Produto_n, strFormatacao2Digitos)

      VLR_ANTERIOR_N = TabProduto!PRECO_VENDA
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   If txtLote.Text = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

   QTDE_ESTOQUE = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value)

   CODG_PRODUTO_A = Trim(txtProduto.Text)

   If Not IsNull(TabProduto.Fields("codg_ncm").Value) Then
      If Len(TabProduto.Fields("codg_ncm").Value) > 2 Then
         If Len(TabProduto.Fields("codg_ncm").Value) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"

            LIMPA_BODY

            FraSeq.Enabled = True
            txtProduto.Enabled = True
            txtProduto.SetFocus
         End If
      End If
   End If

   If Trim(txtLote.Text) = "" Then
      MsgBox "Falta numero PRODUCAO."
      Exit Sub
   End If

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "PRODUCAOITEM", "PRODUCAO_id", Trim(txtLote.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "PRODUCAOITEM", "PRODUCAO_id", Trim(txtLote.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================

   If TabProduto.State = 1 Then _
      TabProduto.Close

   CODIGO_BARRAS = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PRODUTO"
End Sub

Sub LE_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   Dim INDR_PROD_BALANCA  As Boolean
   CODG_PRODUTO_A = Trim(txtProduto.Text)
   INDR_PROD_BALANCA = False

   'LE POR CODIGO DE PRODUTO
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close

      Exit Sub
   End If

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   Qtde_N = 0
   CRITERIO = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close
'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   SQL = "select count(produto_id) from PRODUTO "
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO = Trim(CODIGO_BARRAS)

            'frmPRODUCAOBARRAS.Show 1

            If Trim(CRITERIO) <> "" Then
               txtProduto.Text = Trim(CRITERIO)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select * from PRODUTO "
               SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and situacao <> 'C' "
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then _
                  MOSTRA_DADOS_PRODUTO

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               CRITERIO = ""
               Exit Sub
            End If
         End If
      End If
   End If

   CRITERIO = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close

      txtQTDE.Text = 1

      Call txtQtde_LostFocus

      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS) = 13 Then
      '2 = produtos "in store" (sempre será 2)     1
      'C = código do produto (4,5 ou 6 dígitos)    2 a 8
      'T = total a pagar (sempre 6 dígitos)        9 a 13
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      'pegando codigo do produto no codigo de barras da etiqueta de balança
      txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, TamanhoCodgProdBarra_N))

      If INDR_PANIFIC = True Then
         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select familiaproduto_id,produto_id,produto_balanca,codg_produto,SITUACAO,descricao,"
         SQL = SQL & " peso_liquido,PRECO_ATACADO,PRECO_VENDA,preco_custo,codg_ncm, unidade_medida "
         SQL = SQL & " from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then
               INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value
               Qtde_N = 1

               If Trim(CODIGO_BARRAS) <> "" And INDR_PROD_BALANCA = True Then
                  VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS, 8, TamanhoPesoValorBarra_N) / 100
   
                  If UCase(PESO_VALOR_A) = UCase("GRAMAS") Then
                     Qtde_N = 0 & Int(Mid(CODIGO_BARRAS, 8, TamanhoPesoValorBarra_N))   'gramas
                     Qtde_N = Qtde_N / 1000
                     VALOR_ITEM_N = 0 & (TabProduto.Fields("PRECO_VENDA").Value * Qtde_N)
   
   'regra: se o produto é de balança e unidade medida UN dai vai pegar unidade ao invez de peso
                     If Not IsNull(TabProduto.Fields("unidade_medida").Value) Then
                        If UCase(Trim(TabProduto.Fields("unidade_medida").Value)) = "UN" Then
                           Qtde_N = 0 & Int(Mid(CODIGO_BARRAS, 8, TamanhoPesoValorBarra_N))   'unidade
                           If Not IsNull(TabProduto.Fields("preco_venda").Value) Then _
                              VALOR_ITEM_N = TabProduto.Fields("preco_venda").Value
                        End If
                     End If
                     Else: Qtde_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value) 'sta
                  End If
   
                  PESO_ITEM_N = Qtde_N
                  txtQTDE.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
               End If

               MOSTRA_DADOS_PRODUTO

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               Call txtQtde_LostFocus

               Exit Sub
               Else: MsgBox "Verificar cadastro produto."
            End If
            Else: MsgBox "Verificar cadastro produto."
         End If
         Else
            txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 6))  'PORTO SEGURO

            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select * from PRODUTO "
            SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then
               Qtde_N = 0 & Int(Mid(CODIGO_BARRAS, 8, 5))   'gramas
               PESO_ITEM_N = Qtde_N

               MOSTRA_DADOS_PRODUTO

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               Exit Sub
            End If
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MsgBox "Produto não cadastrado."

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_PRODUTO"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then
            EXCLUIR_ITEM Trim(txtLote.Text), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
End Sub

Sub CHECA_TABELAS()
'On Error GoTo ERRO_TRATA

   If ExisteTabela("RETAGUARDA", "PRODUCAO", "") = False Then
      SQL = "CREATE TABLE [dbo].[PRODUCAO]("
      SQL = SQL & " [PRODUCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_REGISTRO] [datetime] NOT NULL,"
      SQL = SQL & " [DT_FECHA] [datetime] NULL,"
      SQL = SQL & " [STATUS] [nchar](1) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_PRODUCAO] PRIMARY KEY CLUSTERED([PRODUCAO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) "
      SQL = SQL & " ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAO] CHECK CONSTRAINT [FK_PRODUCAO_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAO_USUARIO] FOREIGN KEY([USUARIO_ID])"
      SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAO] CHECK CONSTRAINT [FK_PRODUCAO_USUARIO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If ExisteTabela("RETAGUARDA", "PRODUCAOITEM", "") = False Then
      SQL = "CREATE TABLE [dbo].[PRODUCAOITEM]("
      SQL = SQL & " [PRODUCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE] [Float] not null,"
      SQL = SQL & " [VALOR] [Float] not null"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAOITEM_PRODUCAO] FOREIGN KEY([PRODUCAO_ID])"
      SQL = SQL & " References [dbo].[PRODUCAO]([PRODUCAO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOITEM] CHECK CONSTRAINT [FK_PRODUCAOITEM_PRODUCAO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAOITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOITEM] CHECK CONSTRAINT [FK_PRODUCAOITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_TABELAS"
End Sub

Sub PROCURA_REGISTRO(NUMR_PRODUCAO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   CRITERIO = ""

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PRODUCAO "
   SQL = SQL & " where PRODUCAO_id = " & NUMR_PRODUCAO_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      txtDtLote.Text = Format(TabCABECA.Fields("DT_REGISTRO").Value, "dd/mm/yyyy")

      SETA_GRID

      If TabCABECA.Fields("status").Value = "C" Then _
         MsgBox "Lote Cancelado !!!"
      If TabCABECA.Fields("status").Value = "B" Then _
         MsgBox "Lote Baixado !!!"
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_REGISTRO"
End Sub

Sub GRAVA_REGISTRO(NUMR_PRODUCAO_ID_N As Long, Dt_Fehca As String, SIT_ATUAL_LOTE As String)
'On Error GoTo ERRO_TRATA

   If NUMR_PRODUCAO_ID_N <= 0 Then _
      Exit Sub

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PRODUCAO "
   SQL = SQL & " where PRODUCAO_id = " & NUMR_PRODUCAO_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      If TabCABECA.Fields("status").Value <> "F" Then
         SQL = "update PRODUCAO set "
            SQL = SQL & " USUARIO_ID = " & USUARIO_ID_N        'USUARIO_ID
            SQL = SQL & ", DT_FECHA = '" & DMA(Date) & "'"     'DT_FECHA
            SQL = SQL & ", Status = '" & SIT_ATUAL_LOTE & "'"  'Status
         SQL = SQL & " where PRODUCAO_id = " & NUMR_PRODUCAO_ID_N
         Else
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
            MsgBox "Lote já fechado."
            Exit Sub
      End If
      Else
         SQL = "insert into PRODUCAO "
            SQL = SQL & "(PRODUCAO_ID,ESTABELECIMENTO_ID,USUARIO_ID,DT_REGISTRO,Status)"
         SQL = SQL & " values("

            SQL = SQL & NUMR_PRODUCAO_ID_N       'PRODUCAO_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N    'ESTABELECIMENTO_ID
            SQL = SQL & "," & USUARIO_ID_N            'USUARIO_ID
            SQL = SQL & ",'" & DMA(Date) & "'"        'DT_REGISTRO
            SQL = SQL & ",'" & SIT_ATUAL_LOTE & "'"   'Status

         SQL = SQL & " )"
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   CONECTA_RETAGUARDA.Execute SQL

'=========================
   If Trim(txtSeq.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtSeq.Text) Then _
      Exit Sub
   If Trim(txtProduto.Text) = "" Then _
      Exit Sub
   If Trim(txtQTDE.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtQTDE.Text) Then _
      Exit Sub

   If TabItem.State = 1 Then _
      TabItem.Close

   SQL = "select * from PRODUCAOITEM "
   SQL = SQL & " where PRODUCAO_id = " & NUMR_PRODUCAO_ID_N
   SQL = SQL & " and seq_id = " & txtSeq.Text
   TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabItem.EOF Then
      If TabItem.Fields("status").Value <> "F" Then
         SQL = "update PRODUCAO set "
            SQL = SQL & " qtde = " & tpMOEDA(Qtde_N)
            SQL = SQL & ", VALOR = " & tpMOEDA(Qtde_N * VALOR_ITEM_N)
         SQL = SQL & " where PRODUCAO_id = " & NUMR_PRODUCAO_ID_N
         Else
            If TabItem.State = 1 Then _
               TabItem.Close
            MsgBox "Lote já fechado."
            Exit Sub
      End If
      Else
         SQL = "insert into PRODUCAOITEM "
            SQL = SQL & "(PRODUCAO_ID,SEQ_ID,PRODUTO_ID,QTDE,VALOR)"
         SQL = SQL & " values("

            SQL = SQL & NUMR_PRODUCAO_ID_N       'PRODUCAO_ID
            SQL = SQL & "," & txtSeq.Text             'SEQ_ID
            SQL = SQL & "," & PRODUTO_ID_N            'PRODUTO_ID
            SQL = SQL & "," & tpMOEDA(txtQTDE.Text)   'Qtde
            SQL = SQL & "," & tpMOEDA(Qtde_N * VALOR_ITEM_N) 'VALOR

         SQL = SQL & " )"
   End If
   If TabItem.State = 1 Then _
      TabItem.Close

   CONECTA_RETAGUARDA.Execute SQL

   LIMPA_BODY
   SETA_GRID
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_TABELAS"
End Sub

Sub GERA_SEQUENCIA()

   NUMR_SEQ_N = MAX_ID("SEQ_ID", "PRODUCAOITEM", "PRODUCAO_ID", txtLote.Text, "", "")
   txtSeq.Text = NUMR_SEQ_N

End Sub
