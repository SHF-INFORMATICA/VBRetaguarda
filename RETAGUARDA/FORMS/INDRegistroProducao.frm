VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProducaoRegistroProducaoCadastro 
   Caption         =   "Registro de Produção"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INDRegistroProducao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtQtdeItens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Begin VB.ComboBox cmbTurnoAUX 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbTurno 
         Height          =   360
         Left            =   4560
         TabIndex        =   27
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   720
         MaxLength       =   8
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDtLote 
         Height          =   360
         Left            =   2520
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   240
         Left            =   3960
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPedido 
         Alignment       =   1  'Right Justify
         Caption         =   "Lote:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblDtEmis 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Lote:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1680
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
      TabIndex        =   7
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
         Height          =   350
         Left            =   4065
         Picture         =   "INDRegistroProducao.frx":5C12
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
         TabIndex        =   1
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
         TabIndex        =   9
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
         TabIndex        =   0
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
         TabIndex        =   8
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblQtde 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantidade:"
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
      ForeColor       =   &H00C0C000&
      Height          =   525
      Left            =   6555
      TabIndex        =   6
      Top             =   7020
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão"
      Height          =   705
      Left            =   6960
      TabIndex        =   2
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
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   285
         Left            =   2760
         TabIndex        =   4
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
         TabIndex        =   3
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
      Width           =   11310
      _ExtentX        =   19950
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
               Picture         =   "INDRegistroProducao.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":A8FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":BA53
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":BEA5
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":DD1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":F3D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INDRegistroProducao.frx":113B4
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
      DesignWidth     =   11310
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
Attribute VB_Name = "frmProducaoRegistroProducaoCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim NUMR_LOTE_N            As Long
   Dim TabProd             As New ADODB.Recordset
   Dim VALOR_UNITARIO_N       As Double
   Dim Valr_Venda_Produto_n   As Double
   Dim PESO_ITEM_N            As Double
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
         frmProducaoRegistroProducaoConsulta.Show 1
         If Trim(CRITERIO_A) <> "" Then
            txtLOTE.Text = CRITERIO_A
            NUMR_LOTE_N = CRITERIO_A
            PROCURA_REGISTRO txtLOTE.Text
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
         GRAVA_REGISTRO txtLOTE.Text, Date, "F", cmbTurnoAUX.Text
         
   MsgBox "Processo realizado com sucesso."

   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbTurno_Click()
'On Error GoTo ERRO_TRATA

   cmbTurnoAUX.ListIndex = cmbTurno.ListIndex

   If Trim(cmbTurnoAUX.Text) = "" Then _
      Exit Sub

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTurno_Click"
End Sub

Private Sub cmbTurno_GotFocus()

   cmbTurno.SelStart = 0
   cmbTurno.SelLength = Len(cmbTurno)
   cmbTurno.BackColor = &HC0FFFF

End Sub

Private Sub cmbTurno_LostFocus()
   cmbTurno.BackColor = &HFFFFFF
End Sub

Private Sub txtLote_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtLOTE.Text) = "" Then _
      NUMR_LOTE_N = MAX_ID("REGISTROPRODUCAO_ID", "REGISTROPRODUCAO", "", "", "", "")

   txtLOTE.Text = NUMR_LOTE_N

   If Trim(txtLOTE.Text) <> "" Then
      If IsNumeric(txtLOTE.Text) Then

         NUMR_LOTE_N = txtLOTE.Text

         PROCURA_REGISTRO txtLOTE.Text
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

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtLOTE.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then _
            EXCLUIR_ITEM Trim(txtLOTE.Text), Trim(txtSeq.Text)

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
      PROCESSA_DADOS_PRODUTOS
      If Trim(txtProduto.Text) <> "" And Trim(txtLOTE.Text) <> "" Then _
         txtQTDE.SetFocus
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
   '   MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
   '   txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   QTDE_N = 0 & txtQTDE.Text
   If QTDE_N <= 0 Then _
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
      QTDE_N = 0 & txtQTDE.Text
      If QTDE_N < 0 Then _
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
            QTDE_N = txtQTDE.Text
            If QTDE_N <= 0 Then _
               txtQTDE.Text = 1
         End If
   End If
   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)

   VALOR_ITEM_N = 0 & txtValor.Text
   QTDE_N = 0 & txtQTDE.Text

   GRAVA_REGISTRO txtLOTE.Text, "", "A", cmbTurnoAUX.Text
   txtQTDE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_LostFocus"
End Sub
'============================subrotinas
Private Sub EXCLUIR_ITEM(REGISTROPRODUCAO_ID_N As Long, SEQ_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Trim(REGISTROPRODUCAO_ID_N) > 0 And Trim(SEQ_ID_N) > 0 Then
      Msg = "Deseja Excluir Esse Item?"
      Style = vbYesNo + 32
      Title = "Atenção."
      Help = "DEMO.HLP"
      Ctxt = 1000
      RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
      If RESPOSTA = vbYes Then
         SQL = "Delete from REGISTROPRODUCAOITEM "
         SQL = SQL & " Where REGISTROPRODUCAO_id = " & REGISTROPRODUCAO_ID_N
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

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   Valr_Venda_Produto_n = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""
   txtQTDE.Text = ""
   txtValor.Text = ""

   QTDE_ESTOQUE_N = 0
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

   INDR_VENDA = False

   If TabUSU.State = 1 Then _
      TabUSU.Close

   CRITERIO_A = ""

   MSFlexGrid1.Clear

   FraSeq.Enabled = False

   txtPesoTotal.Text = ""
   txtQtdeItens.Text = "" & Format(0, strFormatacao3Digitos)

   optSint.Value = True
   optAna.Value = False
   PRODUTO_ID_N = 0
   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0
   cmbTurnoAUX.Text = ""
   cmbTurno.Text = ""
   NUMR_SEQ_N = 0
   txtLOTE.Text = ""
   txtDtLote = Format(Date, "dd/mm/yyyy")
   txtValorTotal.Text = ""
   LIMPA_BODY
   
   VALOR_TOTAL_N = 0
   NUMR_LOTE_N = 0
   QTDE_ESTOQUE_N = 0
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

   SQL = "select REGISTROPRODUCAOITEM.SEQ_ID as Seq, PRODUTO.CODG_PRODUTO as Codg, "
   SQL = SQL & " PRODUTO.DESCRICAO as Descrição, REGISTROPRODUCAOITEM.QTDE, REGISTROPRODUCAOITEM.valor as Valor,"
   SQL = SQL & " FAMILIAPRODUTO.DESCRICAO AS DescFamilia, FAMILIAPRODUTO.PRODUCAO, "
   SQL = SQL & " REGISTROPRODUCAOITEM.PRODUTO_ID, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " FAMILIAPRODUTO.CODG_FAMILIA, REGISTROPRODUCAOITEM.REGISTROPRODUCAO_ID, "
   SQL = SQL & " PRODUTO.PESO_LIQUIDO"
   SQL = SQL & " from REGISTROPRODUCAOITEM "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON REGISTROPRODUCAOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " INNER JOIN REGISTROPRODUCAO "
   SQL = SQL & " ON REGISTROPRODUCAOITEM.REGISTROPRODUCAO_ID = REGISTROPRODUCAO.REGISTROPRODUCAO_ID"

   SQL = SQL & " where REGISTROPRODUCAO.REGISTROPRODUCAO_id = " & txtLOTE.Text
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

         If Not IsNull(TabGridVaca.Fields("PRODUCAO").Value) Then
            INDR_PRI = TabGridVaca.Fields("PRODUCAO").Value
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

   If Trim(txtLOTE.Text) <> "" Then
      NUMR_LOTE_N = txtLOTE.Text
      Else: NUMR_LOTE_N = InputBox(SQL3, "Informe número de PRODUCAO a ser impressa ")
   End If

   If NUMR_LOTE_N > 0 Then
      FORMULA_REL = "{REGISTROPRODUCAO.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
      FORMULA_REL = FORMULA_REL & " and {REGISTROPRODUCAO.REGISTROPRODUCAO_id} = " & NUMR_LOTE_N
   End If

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If optSint.Value = True Then
      Nome_Relatorio = "REL_producao_SINT.rpt"
      Else: Nome_Relatorio = "REL_producao_ANALIT.rpt"
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

   If Not IsNull(TabProduto!PRECO_Venda) Then
      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_Venda

      txtValor.Text = "" & Format(Valr_Venda_Produto_n, strFormatacao2Digitos)

      VLR_ANTERIOR_N = TabProduto!PRECO_Venda
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   If txtLOTE.Text = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

   QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value)

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

   If Trim(txtLOTE.Text) = "" Then
      MsgBox "Falta numero REGISTROPRODUCAO."
      Exit Sub
   End If

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "REGISTROPRODUCAOITEM", "REGISTROPRODUCAO_id", Trim(txtLOTE.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "REGISTROPRODUCAOITEM", "REGISTROPRODUCAO_id", Trim(txtLOTE.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================

   If TabProduto.State = 1 Then _
      TabProduto.Close

   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PRODUTO"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then
            EXCLUIR_ITEM Trim(txtLOTE.Text), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
End Sub

Sub CHECA_TABELAS()
'On Error GoTo ERRO_TRATA

   Dim TabTurno As New ADODB.Recordset

   cmbTurno.Clear
   cmbTurnoAUX.Clear

   If TabTurno.State = 1 Then _
      TabTurno.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where tipo = 'A2'"
   TabTurno.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTurno.EOF
      cmbTurno.AddItem Trim(TabTurno!DESCRICAO) & "-" & Trim(TabTurno.Fields("codigo").Value)
      cmbTurnoAUX.AddItem Trim(TabTurno.Fields("codigo").Value)
      TabTurno.MoveNext
   Wend
   If TabTurno.State = 1 Then _
      TabTurno.Close

   If EXISTE_OBJ_BANCO("RETAGUARDA", "REGISTROPRODUCAO", "") = False Then
      SQL = "CREATE TABLE [dbo].[REGISTROPRODUCAO]("
      SQL = SQL & " [REGISTROPRODUCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [TURNO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_REGISTRO] [datetime] NOT NULL,"
      SQL = SQL & " [DT_FECHA] [datetime] NULL,"
      SQL = SQL & " [STATUS] [nchar](1) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_REGISTROPRODUCAO] PRIMARY KEY CLUSTERED([REGISTROPRODUCAO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) "
      SQL = SQL & " ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAO]  WITH CHECK ADD  CONSTRAINT [FK_REGISTROPRODUCAO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAO] CHECK CONSTRAINT [FK_REGISTROPRODUCAO_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAO]  WITH CHECK ADD  CONSTRAINT [FK_REGISTROPRODUCAO_USUARIO] FOREIGN KEY([USUARIO_ID])"
      SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAO] CHECK CONSTRAINT [FK_REGISTROPRODUCAO_USUARIO]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "TURNO_ID", "REGISTROPRODUCAO") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE REGISTROPRODUCAO ADD TURNO_ID INT"
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "REGISTROPRODUCAOITEM", "") = False Then
      SQL = "CREATE TABLE [dbo].[REGISTROPRODUCAOITEM]("
      SQL = SQL & " [REGISTROPRODUCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE] [Float] not null,"
      SQL = SQL & " [VALOR] [Float] not null"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_REGISTROPRODUCAOITEM_REGISTROPRODUCAO] FOREIGN KEY([REGISTROPRODUCAO_ID])"
      SQL = SQL & " References [dbo].[REGISTROPRODUCAO]([REGISTROPRODUCAO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAOITEM] CHECK CONSTRAINT [FK_REGISTROPRODUCAOITEM_REGISTROPRODUCAO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_REGISTROPRODUCAOITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[REGISTROPRODUCAOITEM] CHECK CONSTRAINT [FK_REGISTROPRODUCAOITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_TABELAS"
End Sub

Sub PROCURA_REGISTRO(NUMR_REGISTROPRODUCAO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from REGISTROPRODUCAO "
   SQL = SQL & " where REGISTROPRODUCAO_id = " & NUMR_REGISTROPRODUCAO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      txtDtLote.Text = Format(TabCabeca.Fields("DT_REGISTRO").Value, "dd/mm/yyyy")

      SETA_GRID

      If TabCabeca.Fields("status").Value = "C" Then _
         MsgBox "Lote Cancelado !!!"
      If TabCabeca.Fields("status").Value = "B" Then _
         MsgBox "Lote Baixado !!!"
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_REGISTRO"
End Sub

Sub GRAVA_REGISTRO(NUMR_REGISTROPRODUCAO_ID_N As Long, Dt_Fehca As String, SIT_ATUAL_LOTE As String, TURNO_ID_N As String)
'On Error GoTo ERRO_TRATA

   If Trim(TURNO_ID_N) = "" Then
      MsgBox "Informe Turno para registro."
      cmbTurno.SetFocus
      Exit Sub
   End If

   Dim TabItem As New ADODB.Recordset

   If NUMR_REGISTROPRODUCAO_ID_N <= 0 Then _
      Exit Sub

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from REGISTROPRODUCAO "
   SQL = SQL & " where REGISTROPRODUCAO_id = " & NUMR_REGISTROPRODUCAO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      If TabCabeca.Fields("status").Value <> "F" Then
         SQL = "update REGISTROPRODUCAO set "
            SQL = SQL & " USUARIO_ID = " & USUARIO_ID_N        'USUARIO_ID
            SQL = SQL & ", DT_FECHA = '" & Now & "'"     'DT_FECHA
            SQL = SQL & ", Status = '" & SIT_ATUAL_LOTE & "'"  'Status
         SQL = SQL & " where REGISTROPRODUCAO_id = " & NUMR_REGISTROPRODUCAO_ID_N
         Else
            If TabCabeca.State = 1 Then _
               TabCabeca.Close
            MsgBox "Lote já fechado."
            Exit Sub
      End If
      Else
         SQL = "insert into REGISTROPRODUCAO "
            SQL = SQL & "(REGISTROPRODUCAO_ID,ESTABELECIMENTO_ID,USUARIO_ID,DT_REGISTRO,Status,TURNO_ID)"
         SQL = SQL & " values("

            SQL = SQL & NUMR_REGISTROPRODUCAO_ID_N    'REGISTROPRODUCAO_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N    'ESTABELECIMENTO_ID
            SQL = SQL & "," & USUARIO_ID_N            'USUARIO_ID
            SQL = SQL & ",'" & Now & "'"              'DT_REGISTRO
            SQL = SQL & ",'" & SIT_ATUAL_LOTE & "'"   'Status
            SQL = SQL & "," & cmbTurnoAUX.Text        'TURNO_ID

         SQL = SQL & " )"
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL

   If Trim(SIT_ATUAL_LOTE) = "F" Then _
      MsgBox "Processo realizado com sucesso."

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

   SQL = "select * from REGISTROPRODUCAOITEM "
   SQL = SQL & " where REGISTROPRODUCAO_id = " & NUMR_REGISTROPRODUCAO_ID_N
   SQL = SQL & " and seq_id = " & txtSeq.Text
   TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabItem.EOF Then
      'If TabItem.Fields("status").Value <> "F" Then
      '   SQL = "update REGISTROPRODUCAO set "
      '      SQL = SQL & " qtde = " & tpMOEDA(Qtde_N)
      '      SQL = SQL & ", VALOR = " & tpMOEDA(Qtde_N * VALOR_ITEM_N)
      '   SQL = SQL & " where REGISTROPRODUCAO_id = " & NUMR_REGISTROPRODUCAO_ID_N
         'Else
         '   If TabItem.State = 1 Then _
         '      TabItem.Close
         '   MsgBox "Lote já fechado."
         '   Exit Sub
      'End If
      Else
         SQL = "insert into REGISTROPRODUCAOITEM "
            SQL = SQL & "(REGISTROPRODUCAO_ID,SEQ_ID,PRODUTO_ID,QTDE,VALOR)"
         SQL = SQL & " values("

            SQL = SQL & NUMR_REGISTROPRODUCAO_ID_N       'REGISTROPRODUCAO_ID
            SQL = SQL & "," & txtSeq.Text             'SEQ_ID
            SQL = SQL & "," & PRODUTO_ID_N            'PRODUTO_ID
            SQL = SQL & "," & tpMOEDA(txtQTDE.Text)   'Qtde
            SQL = SQL & "," & tpMOEDA(QTDE_N * VALOR_ITEM_N) 'VALOR

         SQL = SQL & " )"
   End If
   If TabItem.State = 1 Then _
      TabItem.Close
'MsgBox SQL
   CONECTA_RETAGUARDA.Execute SQL

   LIMPA_BODY
   SETA_GRID
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_REGISTRO"
End Sub

Sub GERA_SEQUENCIA()

   NUMR_SEQ_N = MAX_ID("SEQ_ID", "REGISTROPRODUCAOITEM", "REGISTROPRODUCAO_ID", txtLOTE.Text, "", "")
   txtSeq.Text = NUMR_SEQ_N

End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

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
            FraSeq.Enabled = True
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

   If txtLOTE.Text = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

   QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

   If Not IsNull(CODG_NCM_A) Then
      If Len(CODG_NCM_A) > 2 Then
         If Len(CODG_NCM_A) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"

            LIMPA_BODY

            FraSeq.Enabled = True
            txtProduto.Enabled = True
            txtProduto.SetFocus
         End If
      End If
   End If
   If Trim(txtLOTE.Text) = "" Then
      MsgBox "Falta numero REGISTROPRODUCAO."
      Exit Sub
   End If
'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "REGISTROPRODUCAOITEM", "REGISTROPRODUCAO_id", Trim(txtLOTE.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "REGISTROPRODUCAOITEM", "REGISTROPRODUCAO_id", Trim(txtLOTE.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then
      txtQTDE.SetFocus
      txtProduto.SetFocus
   End If
   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
