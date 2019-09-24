VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoManutencao 
   Caption         =   "Manutenção Pedido Venda"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PEDIDOMANUT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbVend 
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
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
      Width           =   3585
   End
   Begin VB.ComboBox cmbForma 
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
   End
   Begin VB.ComboBox cmbAuxForma 
      BackColor       =   &H80000000&
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1380
      Visible         =   0   'False
      Width           =   735
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
      Left            =   6120
      MaxLength       =   100
      TabIndex        =   6
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbVendAux 
      BackColor       =   &H80000000&
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDescProd 
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
      Left            =   3780
      MaxLength       =   100
      TabIndex        =   3
      Top             =   2280
      Width           =   3945
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3315
      Picture         =   "PEDIDOMANUT.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   405
   End
   Begin VB.CommandButton cmdConsCli 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5640
      Picture         =   "PEDIDOMANUT.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   405
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
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
         NumButtons      =   3
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
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
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
               Picture         =   "PEDIDOMANUT.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOMANUT.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOMANUT.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOMANUT.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOMANUT.frx":B2FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOMANUT.frx":D2E1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   360
      Left            =   3600
      TabIndex        =   11
      Top             =   840
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11865
      DesignHeight    =   7290
   End
   Begin MSComctlLib.ListView lstPedidoItem 
      Height          =   1905
      Left            =   45
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   3360
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648384
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2252
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   7508
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   1668
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valr.Unitário"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Desconto"
         Object.Width           =   1877
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valr.Total"
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "NCM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Pedido"
         Object.Width           =   1502
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3015
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   5318
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      Appearance      =   1
      MousePointer    =   99
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cupom"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NFe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr.Venda"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Vlr.Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Dt.Emisão"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Faturamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ID"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "NºCaixa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "TpRegistro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Comanda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ipPedidoStatus"
         Object.Width           =   18
      EndProperty
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   9000
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor(a):"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   22
      Top             =   1800
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   21
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   20
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2820
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   18
      Top             =   2325
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8160
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmPedidoManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim SQL_CONSULTA        As String
   Dim SQL_CONSULTA_CORPO  As String
   Dim SQL_CONSULTA3       As String
   Dim CANCELA_LOOP        As Boolean

Private Sub Form_Load()
   CANCELA_LOOP = False
   CARREGA_COMBOS
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         CANCELA_LOOP = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         MONTA_CONSULTA_SQL True
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         PEDIDO_ID_N = 0
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
   End Select
   PEDIDO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  CRITERIO_A = ""
                  CRITERIO_A = Trim(InputBox("Informe CPF/CNPJ do cliente", "Atualização de Dados Pedido Venda", CRITERIO_A))

                  If Trim(CRITERIO_A) <> "" Then
                     SQL = "update PEDIDO set "
                     SQL = SQL & " cgccpf = '" & Trim(CRITERIO_A) & "'"
                     SQL = SQL & " where pedido_id = " & lstPedido.SelectedItem.Text
                     CONECTA_RETAGUARDA.Execute SQL
                  End If

                  SQL = ""
                  CRITERIO_A = ""
                  MONTA_CONSULTA_SQL True
               End If
            End If
         End If
      Case vbKeyF6
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then

                  NUMR_SEQ_N = 0 & Trim(lstPedido.SelectedItem.ListSubItems.item(15).Text)
                  If NUMR_SEQ_N < 3 Then
                     CRITERIO_A = ""
                     frmPedidoCancela.txtPedido.Text = 0 & lstPedido.SelectedItem.Text
                     frmPedidoCancela.Show 1
                     SQL = ""
                     CRITERIO_A = ""
                     MONTA_CONSULTA_SQL True
                     Else
                        If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
                           CRITERIO_A = ""
                           frmPedidoCancela.txtPedido.Text = 0 & lstPedido.SelectedItem.Text
                           frmPedidoCancela.Show 1
                           SQL = ""
                           CRITERIO_A = ""
                           MONTA_CONSULTA_SQL True
                           Else: MsgBox "Não permitido."
                        End If
                  End If

               End If
            End If
         End If
      Case vbKeyF7
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.ListItems.Clear

            SQL = "select PRODUTO.CODG_PRODUTO, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "
            SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.PRECO_CUSTO, pedidoitem.seq_id,"
            SQL = SQL & " PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.CFOP_id, pedidoitem.status, "
            SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, Produto.FORNECEDOR_ID"
            SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
            SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

            SQL = SQL & " where pedidoitem.produto_id = produto.produto_id "
            SQL = SQL & " and PEDIDO_ID = " & lstPedido.SelectedItem.ListSubItems.item(11).Text
            'SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               MOSTRA_TOP "Duplo Click no grid ocultar", " ", " ", " ", ""
               lstPedidoItem.Visible = True
            End If
            While Not TabTemp.EOF
               VALOR_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
               VALOR_ITEM_N = TabTemp.Fields("qtd_pedida").Value * (TabTemp.Fields("valor_item").Value - VALOR_DESCONTO_N)

               Set item = lstPedidoItem.ListItems.Add(, "seq." & TabTemp.Fields("seq_id").Value, Trim(TabTemp.Fields("codg_produto").Value))
               item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
               item.SubItems(2) = "" & Format(Trim(TabTemp.Fields("qtd_pedida").Value), strFormatacao3Digitos)
               item.SubItems(3) = "" & Format(Trim(TabTemp.Fields("valor_item").Value), strFormatacao2Digitos)
               item.SubItems(4) = "" & Format(Trim(TabTemp.Fields("valor_desconto").Value), strFormatacao2Digitos)
               item.SubItems(5) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
               item.SubItems(6) = "" & Trim(TabTemp.Fields("CODG_ncm").Value)
               item.SubItems(7) = "" & lstPedido.SelectedItem.Text

               If Trim(TabTemp.Fields("status").Value) = "A" Then
                  item.ForeColor = vbBlue
                  item.ListSubItems(1).ForeColor = vbBlue
                  item.ListSubItems(2).ForeColor = vbBlue
                  item.ListSubItems(3).ForeColor = vbBlue
                  item.ListSubItems(4).ForeColor = vbBlue
                  item.ListSubItems(5).ForeColor = vbBlue
                  item.ListSubItems(6).ForeColor = vbBlue
               End If
               If Trim(TabTemp.Fields("status").Value) = "P" Then
                  item.ForeColor = vbBlack
                  item.ListSubItems(1).ForeColor = vbBlack
                  item.ListSubItems(2).ForeColor = vbBlack
                  item.ListSubItems(3).ForeColor = vbBlack
                  item.ListSubItems(4).ForeColor = vbBlack
                  item.ListSubItems(5).ForeColor = vbBlack
                  item.ListSubItems(6).ForeColor = vbBlack
               End If
               If Trim(TabTemp.Fields("status").Value) = "C" Then
                  item.ForeColor = vbRed
                  item.ListSubItems(1).ForeColor = vbRed
                  item.ListSubItems(2).ForeColor = vbRed
                  item.ListSubItems(3).ForeColor = vbRed
                  item.ListSubItems(4).ForeColor = vbRed
                  item.ListSubItems(5).ForeColor = vbRed
                  item.ListSubItems(6).ForeColor = vbRed
               End If
               TabTemp.MoveNext
               CRITERIO_A = ""
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.Refresh
         End If
      Case vbKeyF11
         'frmSenha.Show 1

         'If UCase(CRITERIO_A) = UCase("acerto") Then
            'PEDIDO_ID_N = 0
            'If Not IsNull(lstPedido.selectedItem.Text) Then
            '   If IsNumeric(lstPedido.selectedItem.Text) Then
            '      PEDIDO_ID_N = lstPedido.selectedItem.Text

            '      frmPedidoClienteAcerto.Show 1
            '      MONTA_CONSULTA_SQL True
            '      PEDIDO_ID_N = 0
            '   End If
            'End If
         'End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoVENDA_KeyDown"
End Sub

Private Sub cmdConsCli_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then _
      txtCNPJCPF.Text = CNPJCPF_A
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus
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

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Text = "99999999999"

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome,cliente_id from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCliente.EOF Then
      Beep
      MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
      txtCNPJCPF.SetFocus
      Exit Sub
      Else:
         CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
         If TabCliente!NOME <> "" Then _
            txtCli.Text = TabCliente!NOME
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_LostFocus"
End Sub

Private Sub TXTDTFIM_LostFocus()
   CHECA_ULTIMO_DIA_MES
End Sub

Private Sub TXTPRODUTO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_PRODUTO
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
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      MONTA_CONSULTA_SQL True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpedido_KeyPress"
End Sub

Private Sub lstPedido_DblClick()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""
   If Not IsNull(lstPedido.SelectedItem.Text) Then
      NUMR_SEQ_N = 0 & Trim(lstPedido.SelectedItem.ListSubItems.item(15).Text)
      If NUMR_SEQ_N = 3 Or NUMR_SEQ_N = 5 Or NUMR_SEQ_N = 7 Or NUMR_SEQ_N = 9 Then
         MsgBox "Permitido somente consulta."
         PEDIDO_ID_N = 0
         Exit Sub
      End If
      CRITERIO_A = lstPedido.SelectedItem.Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_DblClick"
End Sub

Private Sub lstpedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedido, ColumnHeader
End Sub

Private Sub lstPedido_Click()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 0
   If Not IsNull(lstPedido.SelectedItem.Text) Then _
      If IsNumeric(lstPedido.SelectedItem.Text) Then _
         PEDIDO_ID_N = lstPedido.SelectedItem.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_Click"
End Sub

Private Sub lstPedidoitem_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoitem_DblClick"
End Sub

Private Sub lstpedidoitem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedidoItem, ColumnHeader
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxForma.ListIndex = cmbForma.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFORMA_Click"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_Click"
End Sub

Private Sub txtCNPJCPF_GotFocus()
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

Private Sub txtCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyDown"
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
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

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
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

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   CANCELA_LOOP = False
   CLIENTE_ID_N = 0
   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   lstPedido.ListItems.Clear
   txtPedido.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""

   If cmbVend.Enabled = True Then _
      cmbVend.Text = ""

   cmbForma.Text = ""
   cmbAuxForma.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""

   lstPedido.Visible = True
   txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = lstPedido.SelectedItem.Text
   CNPJCPF_A = ""

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status, cgccpf from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      CNPJCPF_A = "" & Trim(TabCABECA.Fields("cgccpf").Value)
      If Not IsNull(TabCABECA!Status) Then
         If TabCABECA!Status <> "9" Then
            If Trim(CNPJCPF_A) = "99999999999" Then
               Msg = "Para geração de nota fiscal eletrônica, os dados do cliente devem ser cadastrados, deseja continuar essa operação ?"
               PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CNPJCPF_A = ""
                  TIPO_PESSOA_CADASTRO = "CLIENTE"
                  frmPessoaConsulta.Show 1
                  If Trim(CNPJCPF_A) <> "" Then
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

SQL = "select nome,cgccpf from CLIENTE WITH (NOLOCK)"
SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_A) & "'"
TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
If TabConsulta.EOF Then
   Msg = "CNPF/CPF informado não cadastrado, deseja realizar cadastro de cliente agora ?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaCadastro.Show 1

      'frmCADASTROCLIENTE.Show 1
      MsgBox "Repetir operação."
      Else
         If TabCABECA.State = 1 Then _
            TabCABECA.Close
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Exit Sub
   End If
   Else
      If TabCABECA.State = 1 Then _
         TabCABECA.Close
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
End If
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

                     SQL = "update PEDIDO set "
                     SQL = SQL & " cgccpf = '" & Trim(CNPJCPF_A) & "'"
                     SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                     SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
                     Else
                        If TabCABECA.State = 1 Then _
                           TabCABECA.Close
                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close
                        Exit Sub
                  End If
                  Else
                     If TabCABECA.State = 1 Then _
                        TabCABECA.Close
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
                     Exit Sub
               End If
            End If

            CRITERIO_A = PEDIDO_ID_N
            'TIPO_NFe_GERAR = "R"
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            If USA_DOC_FISCAL = True Then _
               If USA_NFe = True Then _
                  frmNOTAGERA.Show 1
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
End Sub

Sub PROCURA_PRODUTO()
'On Error GoTo ERRO_TRATA

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
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Sub CHECA_ULTIMO_DIA_MES()
   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
      txtDtFim.Text = ""

      txtDtIni.PromptInclude = True
      If IsDate(txtDtIni.Text) Then
         CRITERIO_A = FimDoMes(DMA(txtDtIni.Text), False)
         CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
         txtDtFim.Text = CRITERIO_A
         txtDtFim.PromptInclude = True
      End If
   End If
End Sub

Sub CARREGA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(6).Visible = False 'NFe
   If USA_NFe = True Then _
      Toolbar1.Buttons(6).Visible = True

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbVend.Enabled = True

      Else
         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select logon from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            CRITERIO_A = Chr$(39) & Trim(TabUSU.Fields("logon").Value) & "%" & Chr(39)

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
            SQL = SQL & " where descricao like " & CRITERIO_A
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVend.Text = TabVENDEDOR!DESCRICAO
               cmbVendAux.Text = TabVENDEDOR!VENDEDOR_ID
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_VENDEDOR"
End Sub

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   VALOR_TOTAL_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   HORA_INI = Time
   Me.Caption = "Aguarde, Pesquisando ..."
   SQL_CONSULTA = ""
   SQL_CONSULTA_CORPO = ""
   SQL_CONSULTA3 = ""
   CRITERIO_A = ""

   CHECA_ULTIMO_DIA_MES

   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False

   VALOR_TOTAL_N = 0
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SQL_CONSULTA = ""
   SQL_CONSULTA_CORPO = ""
   SQL_CONSULTA3 = ""
   CRITERIO_A = ""

   Dim SQL_COUNT  As String
   SQL_COUNT = ""

   SQL_CONSULTA = "select * from vwCONSULTA_PEDIDO WITH (NOLOCK) "
SQL_COUNT = "select count(pedido_id) from PEDIDO "

   SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " where pedido_id Is Not Null"

'itens pedido
SQL = SQL & " and status <> 'C'"

   If Trim(txtPedido.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and pedido_id = " & txtPedido.Text

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then _
      If CLIENTE_ID_N > 0 Then _
         SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and cliente_id = " & CLIENTE_ID_N
   txtCNPJCPF.PromptInclude = True

   If Trim(cmbVend.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and vendedor_id = " & cmbVendAux.Text


   If Trim(cmbAuxForma.Text) <> "" Then _
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and tipovenda_id = " & cmbAuxForma.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and dt_req >= '" & txtDtIni.Text & "'"
      SQL_CONSULTA_CORPO = SQL_CONSULTA_CORPO & " and dt_req <= '" & txtDtFim.Text & "'"
   End If

'===============
SQL_COUNT = SQL_COUNT & SQL_CONSULTA_CORPO
'===============

   Dim SQL_ITENS        As String
   Dim TabTemp          As New ADODB.Recordset
   Dim Conta_Produto_N  As Long
   Dim Peso_N           As Double

   Peso_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   SQL_ITENS = ""

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

   CONT_N = 0
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL_CONSULTA = SQL_CONSULTA & SQL_CONSULTA_CORPO & " order by PEDIDO_ID desc"

   TabTemp.Open SQL_CONSULTA, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Exit Sub
   End If
   If Not TabTemp.EOF Then
      While Not TabTemp.EOF
         DoEvents

         If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
            CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1

            PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value

            NUMR_SEQ_N = NUMR_SEQ_N + 1
            Set item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp.Fields("PEDIDO_ID").Value)

            item.SubItems(11) = "" & TabTemp.Fields("PEDIDO_ID").Value
            item.SubItems(1) = "" & TabTemp.Fields("numr_cupom").Value
            item.SubItems(2) = "" & TabTemp.Fields("numr_nota").Value
            item.SubItems(3) = "" & Trim(TabTemp!NOME_CLIENTE) & " - " & Trim(TabTemp.Fields("CNPJCPF").Value)

            If IsNull(TabTemp!NOME_CLIENTE) Or Trim(TabTemp!NOME_CLIENTE) = "" Then
               If TabCliente.State = 1 Then _
                  TabCliente.Close

               SQL = "select nome from CLIENTE WITH (NOLOCK)"
               SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"
               TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCliente.EOF Then _
                  item.SubItems(3) = "" & TabCliente!NOME

               If TabCliente.State = 1 Then _
                  TabCliente.Close
            End If

            item.SubItems(7) = TabTemp!dt_req
            item.SubItems(8) = ""

            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            SQL = "select * from TIPOVENDA WITH (NOLOCK)"
            SQL = SQL & " where tipovenda_id = " & TabTemp!TipoVenda_ID
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then _
               item.SubItems(8) = TabDESCR!DESCRICAO
            If TabDESCR.State = 1 Then _
               TabDESCR.Close

            If TabUSU.State = 1 Then _
               TabUSU.Close

            item.SubItems(9) = ""
   
            SQL = "select * from vwVendedor WITH (NOLOCK)"
            SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id").Value
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               item.SubItems(9) = TabUSU!DESCRICAO
            If TabUSU.State = 1 Then _
               TabUSU.Close

            item.SubItems(10) = ""

            If Not IsNull(TabTemp.Fields("Status")) Then
               If TabTemp.Fields("Status") = 2 Then
                  If TabTemp.Fields("tipo_registro") = "O" Then
                     item.SubItems(10) = "Orcamento"
                     Else: item.SubItems(10) = "Pedido"
                  End If
               End If
               If TabTemp.Fields("Status").Value = 3 Then _
                  item.SubItems(10) = "3-Faturado"
               If TabTemp.Fields("Status").Value = 4 Then _
                  item.SubItems(10) = "4-Cupom"
               If TabTemp.Fields("Status").Value = 5 Then _
                  item.SubItems(10) = "5-Faturado"
               If TabTemp.Fields("Status").Value = 7 Then _
                  item.SubItems(10) = "7-Cupom Fiscal"
               If TabTemp.Fields("Status").Value = 9 Then _
                  item.SubItems(10) = "9-Cancelado"
            End If

            If Not IsNull(TabTemp.Fields("numero_caixa_cpu").Value) Then _
               item.SubItems(12) = TabTemp.Fields("numero_caixa_cpu").Value

            item.SubItems(13) = TabTemp.Fields("tipo_registro").Value

            VALOR_DESCONTO_N = 0
            'VALOR_TOTAL_DESCONTO_N = 0

            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPedidoItem.EOF Then _
               If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
                  VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

           If Not IsNull(TabTemp.Fields("desccabeca").Value) Then _
              VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabTemp.Fields("desccabeca").Value

            VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

            'BUSCA VALOR TOTAL VENDA
            VALOR_ITEM_N = 0

            SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
            'SQL = SQL & " and tipo_reg = 'PC' "
            SQL = SQL & " and pedidoitem.status <> 'C' "
            TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
            If TabPedidoItem.State = 1 Then _
               TabPedidoItem.Close

            VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

            item.SubItems(4) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
            item.SubItems(5) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
            item.SubItems(6) = Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
            item.SubItems(15) = "" & TabTemp.Fields("sit_pedido")
            If TabTemp.Fields("SIT_PEDIDO").Value = 1 Then
               item.ForeColor = vbRed
               item.ListSubItems(1).ForeColor = vbRed
               item.ListSubItems(2).ForeColor = vbRed
               item.ListSubItems(3).ForeColor = vbRed
               item.ListSubItems(4).ForeColor = vbRed
               item.ListSubItems(5).ForeColor = vbRed
               item.ListSubItems(6).ForeColor = vbRed
               item.ListSubItems(7).ForeColor = vbRed
               item.ListSubItems(8).ForeColor = vbRed
               item.SubItems(10) = "" & "Em Aberto - 1"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 2 Then
               item.ForeColor = vbBlue
               item.ListSubItems(1).ForeColor = vbBlue
               item.ListSubItems(2).ForeColor = vbBlue
               item.ListSubItems(3).ForeColor = vbBlue
               item.ListSubItems(4).ForeColor = vbBlue
               item.ListSubItems(5).ForeColor = vbBlue
               item.ListSubItems(6).ForeColor = vbBlue
               item.ListSubItems(7).ForeColor = vbBlue
               item.ListSubItems(8).ForeColor = vbBlue
               item.ListSubItems(9).ForeColor = vbBlue
               item.SubItems(10) = "" & "A Faturar - 2"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 3 Then
               item.ForeColor = vbBlack
               item.ListSubItems(1).ForeColor = vbBlack
               item.ListSubItems(2).ForeColor = vbBlack
               item.ListSubItems(3).ForeColor = vbBlack
               item.ListSubItems(4).ForeColor = vbBlack
               item.ListSubItems(5).ForeColor = vbBlack
               item.ListSubItems(6).ForeColor = vbBlack
               item.ListSubItems(7).ForeColor = vbBlack
               item.ListSubItems(8).ForeColor = vbBlack
               item.ListSubItems(9).ForeColor = vbBlack
               item.SubItems(10) = "" & "Faturado - 3"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 5 Then
               item.ForeColor = vbBlack
               item.ListSubItems(1).ForeColor = vbBlack
               item.ListSubItems(2).ForeColor = vbBlack
               item.ListSubItems(3).ForeColor = vbBlack
               item.ListSubItems(4).ForeColor = vbBlack
               item.ListSubItems(5).ForeColor = vbBlack
               item.ListSubItems(6).ForeColor = vbBlack
               item.ListSubItems(7).ForeColor = vbBlack
               item.ListSubItems(8).ForeColor = vbBlack
               item.ListSubItems(9).ForeColor = vbBlack
               item.SubItems(10) = "" & "Faturado - 5"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 6 Then
               item.ForeColor = vbBlack
               'Item.ListSubItems(1).ForeColor = vbYellow
               'Item.ListSubItems(2).ForeColor = vbYellow
               'Item.ListSubItems(3).ForeColor = vbYellow
               'Item.ListSubItems(4).ForeColor = vbYellow
               'Item.ListSubItems(5).ForeColor = vbYellow
               'Item.ListSubItems(6).ForeColor = vbYellow
               'Item.ListSubItems(7).ForeColor = vbYellow
               'Item.ListSubItems(8).ForeColor = vbYellow
               item.ListSubItems(10).ForeColor = vbYellow
               item.SubItems(10) = "" & "Não Contabilizado"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 7 Then
               item.ForeColor = vbMagenta
               item.ListSubItems(1).ForeColor = vbMagenta
               item.ListSubItems(2).ForeColor = vbMagenta
               item.ListSubItems(3).ForeColor = vbMagenta
               item.ListSubItems(4).ForeColor = vbMagenta
               item.ListSubItems(5).ForeColor = vbMagenta
               item.ListSubItems(6).ForeColor = vbMagenta
               item.ListSubItems(7).ForeColor = vbMagenta
               item.ListSubItems(8).ForeColor = vbMagenta
               item.ListSubItems(9).ForeColor = vbMagenta
               item.SubItems(10) = "" & "Cupom Fiscal - 7"
            End If
            If TabTemp.Fields("SIT_PEDIDO").Value = 9 Then
               item.ListSubItems(8).ForeColor = &HC0E0FF
               item.ListSubItems(9).ForeColor = &HC0E0FF
               item.ListSubItems(10).ForeColor = &HC0E0FF
               item.SubItems(10) = "" & "Cancelado - 9"
            End If
         End If

         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed

         'verificando se é venda com comanda eletronica
         item.SubItems(14) = "" & TabTemp.Fields("cartaobarra_id").Value

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select cartaobarra_id from PEDIDOTEMP WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("PEDIDO_ID").Value
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            item.ListSubItems(8).ForeColor = vbBlack
            item.ListSubItems(9).ForeColor = vbBlack
            item.SubItems(14) = "" & TabDESCR.Fields(0).Value
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         TabTemp.MoveNext

         If CANCELA_LOOP = True Then _
            Exit Sub
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
''''''''''''''''''''''''''''''''
   SQL_ITENS = "select count(qtd_pedida) from vwCONSULTA_PEDIDO WITH (NOLOCK) "
   SQL_ITENS = SQL_ITENS & SQL_CONSULTA_CORPO
   SQL_ITENS = SQL_ITENS & " and vwCONSULTA_PEDIDO.tipo_reg = 'PC' "
   SQL_ITENS = SQL_ITENS & " and vwCONSULTA_PEDIDO.status <> 'C' "
   SQL = SQL & " and status <> 'C'"
   TabTemp.Open SQL_ITENS, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabTemp.Fields(0).Value) Then _
      Conta_Produto_N = Conta_Produto_N + TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

''''''''''''''''''''''''''''''''
   HORA_FIM = Time

   MOSTRA_TOP "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

