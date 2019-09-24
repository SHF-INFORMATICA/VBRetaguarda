VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPEDIDOFATURA 
   Caption         =   "Pedido Faturamento"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PEDIDOFATURA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbFaturaAUX 
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
      Left            =   4800
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFatura 
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
      Left            =   4800
      TabIndex        =   16
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cmbDocAUX 
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
      Left            =   7680
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbDoc 
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
      TabIndex        =   13
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtValor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   7290
      TabIndex        =   11
      Top             =   6600
      Width           =   2295
   End
   Begin VB.ComboBox cmbSituacaoAUX 
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
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1320
      TabIndex        =   3
      Top             =   6600
      Width           =   855
   End
   Begin VB.ComboBox cmbSITUACAO 
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
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
               Picture         =   "PEDIDOFATURA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOFATURA.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOFATURA.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOFATURA.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOFATURA.frx":9EFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOFATURA.frx":BEDD
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
      DesignWidth     =   9690
      DesignHeight    =   7050
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   360
      Left            =   7680
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid grdPedido 
      Bindings        =   "PEDIDOFATURA.frx":D426
      Height          =   4095
      Left            =   30
      TabIndex        =   10
      Top             =   1920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoPedido 
      Height          =   330
      Left            =   30
      Top             =   6120
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Grid Pedido"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3720
      TabIndex        =   19
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fatura:"
      Height          =   240
      Left            =   3720
      TabIndex        =   17
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Doc.:"
      Height          =   240
      Left            =   6600
      TabIndex        =   14
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor = "
      Height          =   240
      Left            =   6360
      TabIndex        =   12
      Top             =   6600
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos = "
      Height          =   240
      Left            =   150
      TabIndex        =   9
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   240
      Left            =   6480
      TabIndex        =   6
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "frmPEDIDOFATURA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CARREGA_COMBOS
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub txtDtFim_LostFocus()
   CHECA_ULTIMO_DIA_MES
   txtDtFim.BackColor = &HFFFFFF
End Sub

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
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub


Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub cmbSituacao_Click()
'On Error GoTo ERRO_TRATA

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbsituacao_Click"
End Sub

Private Sub cmbFATURA_Click()
'On Error GoTo ERRO_TRATA

   cmbFaturaAUX.ListIndex = cmbFatura.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFATURA_Click"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni.Text)
   txtDtIni.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.PromptInclude = True
      If IsDate(txtDtIni.Text) Then
         txtDtFim.PromptInclude = False
            txtDtFim.Text = UltimoDiaMes(Month(txtDtIni.Text), Year(txtDtIni.Text)) & "23:59:59"
         txtDtFim.PromptInclude = True
      End If
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim.Text)
   txtDtFim.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   CLIENTE_ID_N = 0
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   cmbDoc.Text = ""
   cmbDocAUX.Text = ""
   PRODUTO_ID_N = 0
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtReg.Text = ""
   txtValor.Text = ""

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
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

   cmbSituacao.Clear
   cmbSituacaoAUX.Clear

   cmbSituacao.AddItem "Todos"
   cmbSituacaoAUX.AddItem "1"

   'cmbSITUACAO.AddItem "Todos"
   'cmbSituacaoAUX.AddItem "1"

   cmbSituacao.AddItem "Autorizados/NFCe;NFe"
   cmbSituacaoAUX.AddItem "100"

   cmbSituacao.AddItem "Pendentes/Erros"
   cmbSituacaoAUX.AddItem ""

   cmbSituacao.Text = "Autorizados"
   cmbSituacaoAUX.Text = "100"

   cmbDoc.Clear
   cmbDocAUX.Clear

   cmbDoc.AddItem "NFCe"
   cmbDocAUX.AddItem "NFC"

   cmbDoc.AddItem "NFe"
   cmbDocAUX.AddItem "NFE"

   Me.Enabled = True
   Me.KeyPreview = True
   VALOR_TOTAL_N = 0


   cmbFatura.Clear
   cmbFaturaAUX.Clear

   cmbFatura.AddItem "Dinheiro"
   cmbFaturaAUX.AddItem "01"

   cmbFatura.AddItem "Cheque"
   cmbFaturaAUX.AddItem "02"

   cmbFatura.AddItem "Cartão de Crédito"
   cmbFaturaAUX.AddItem "03"

   cmbFatura.AddItem "Cartão de Débito"
   cmbFaturaAUX.AddItem "04"

   cmbFatura.AddItem "Crédito Loja"
   cmbFaturaAUX.AddItem "05"

   cmbFatura.AddItem "Vale Alimentação"
   cmbFaturaAUX.AddItem "10"

   cmbFatura.AddItem "Vale Refeição"
   cmbFaturaAUX.AddItem "11"

   cmbFatura.AddItem "Vale Presente"
   cmbFaturaAUX.AddItem "12"

   cmbFatura.AddItem "Vale Combustível"
   cmbFaturaAUX.AddItem "13"

   cmbFatura.AddItem "Duplicata Mercantil"
   cmbFaturaAUX.AddItem "14"

   cmbFatura.AddItem "Sem pagamento"
   cmbFaturaAUX.AddItem "90"

   cmbFatura.AddItem "Outros"
   cmbFaturaAUX.AddItem "99"

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Me.Enabled = False

   Dim TabMFA010  As New ADODB.Recordset
   Dim TabPedidos As New ADODB.Recordset
   Dim sqlCabeca  As String
   Dim sqlCorpo   As String

   CHECA_ULTIMO_DIA_MES

   lblPeriodo.Caption = ""
   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then

      lblPeriodo.Caption = "De: " & _
      TRAZ_NOME_MES(Month(txtDtIni.Text)) & "/" & Year(txtDtIni.Text) & _
      " à " & _
      TRAZ_NOME_MES(Month(txtDtFim.Text)) & "/" & Year(txtDtFim.Text)

      lblPeriodo.Refresh
      DoEvents
   End If

   txtReg.Text = ""
   sqlCabeca = ""
   sqlCorpo = ""
   txtValor.Text = ""
   txtReg.Text = ""
   grdPedido.Visible = True
   txtReg.BackColor = &HFFFFFF
   txtValor.BackColor = &HFFFFFF

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   If cmbSituacaoAUX.Text = "100" Then
      If CONECTA_GLOBAL.State = 1 Then _
         CONECTA_GLOBAL.Close

      ABRE_BANCO_GLOBAL

      If CONECTA_GLOBAL.State <> 1 Then
         'MsgBox "Banco GLOBAL não conectado."
         Exit Sub
      End If

      sqlCabeca = "select MFADOC,MFASEQUENCIA,MFAPREFIXO,MFAEMISSAO,MFAVALLIQUI,MFACODSTAT,"
      sqlCabeca = sqlCabeca & " MFACODPROT,MFACHAVENFE,MFAMOTRESU,MFACODSITT,MFAREGISTRO,MFACODEMP,"
      sqlCabeca = sqlCabeca & " MFALOJA, MFAFILIAL, MFACLIENTE,MFAOBSNOTA,MFANOMECONSUMIDOR,MFACPFCONSUMIDOR"
   
      sqlCorpo = " from MFA010 WITH (NOLOCK) "
   
      sqlCorpo = sqlCorpo & " INNER JOIN SE1010 WITH (NOLOCK) "
      sqlCorpo = sqlCorpo & " ON MFA010.MFADOC = SE1010.E1_NUMNOTA"
   
      sqlCorpo = sqlCorpo & " where CONVERT(INT,MFADOC) > 0"
   
         sqlCorpo = sqlCorpo & " and MFALOJA = '" & ESTABELECIMENTO_ID_N & "'"
         sqlCorpo = sqlCorpo & " and MFAfilial = '" & ESTABELECIMENTO_ID_N & "'"
   
      If Trim(cmbSituacaoAUX.Text) = "100" Then
         sqlCorpo = sqlCorpo & " and MFACODSTAT = '" & Trim(cmbSituacaoAUX.Text) & "'"
         Else: sqlCorpo = sqlCorpo & " and MFACODSTAT <> 100 "
      End If
   
      If Trim(cmbDocAUX.Text) <> "" Then _
         sqlCorpo = sqlCorpo & " and MFAPREFIXO = '" & Trim(UCase(cmbDocAUX.Text)) & "'"
   
      If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
         sqlCorpo = sqlCorpo & " and MFAEMISSAO >= '" & txtDtIni.Text & "'"
         sqlCorpo = sqlCorpo & " and MFAEMISSAO <= '" & txtDtFim.Text & "'"
         Else
            MsgBox "Informe Data Válida."
            txtDtIni.SetFocus
            Exit Sub
      End If
   
      If Trim(cmbFaturaAUX.Text) <> "" Then _
         sqlCorpo = sqlCorpo & " and E1_TIPO = '" & Trim(cmbFaturaAUX.Text) & "'"

      'Campo E1_TIPO=Forma de Pagamento gravar os seguintes numeros :
      '01=Dinheiro
      '02=Cheque
      '03=Cartão de Crédito
      '04=Cartão de Débito
      '05=Crédito Loja
      '10=Vale Alimentação
      '11=Vale Refeição
      '12=Vale Presente
      '13=Vale Combustível
      '14=Duplicata Mercantil
      '90= Sem pagamento
      '99=Outros
      'onde os numeros informados 03 e 04 e obrigatorios preencher os campos criados no item 03.
      'ver relação pois essas formas dependem de cada empresa
      'MsgBox sqlCabeca & sqlCorpo
      'Debug.Print sqlCabeca & sqlCorpo
   
      adoPedido.ConnectionString = AUTENTICA_GRID_GLOBAL
      adoPedido.CommandType = adCmdText
   
      adoPedido.RecordSource = sqlCabeca & sqlCorpo
      adoPedido.Enabled = True
      adoPedido.Refresh
   
      grdPedido.Columns(0).DataField = "MFADOC"
      grdPedido.Columns(0).Caption = "Doc."
      grdPedido.Columns(0).Width = 1111
      grdPedido.Columns(0).Alignment = dbgLeft
   
      grdPedido.Columns(1).DataField = "MFASEQUENCIA"
      grdPedido.Columns(1).Caption = "Sequencia."
      grdPedido.Columns(1).Width = 900
      grdPedido.Columns(1).Alignment = dbgLeft
   
      grdPedido.Columns(2).DataField = "MFAPREFIXO"
      grdPedido.Columns(2).Caption = "TpDoc."
      grdPedido.Columns(2).Width = 1000
      grdPedido.Columns(2).Alignment = dbgLeft
   
      grdPedido.Columns(3).DataField = "MFAEMISSAO"
      grdPedido.Columns(3).Caption = "Dt.Emissão"
      grdPedido.Columns(3).Width = 3000
      grdPedido.Columns(3).Alignment = dbgLeft
   
      grdPedido.Columns(4).DataField = "MFAVALLIQUI"
      grdPedido.Columns(4).Caption = "Valor"
      grdPedido.Columns(4).Width = 2000
      grdPedido.Columns(4).Alignment = dbgRight
   
      grdPedido.Columns(5).DataField = "MFACODSTAT"
      grdPedido.Columns(5).Caption = "Situação"
      grdPedido.Columns(5).Width = 800
      grdPedido.Columns(5).Alignment = dbgLeft
   
      grdPedido.Columns(6).DataField = "MFACODPROT"
      grdPedido.Columns(6).Caption = "Protocolo"
      grdPedido.Columns(6).Width = 2000
      grdPedido.Columns(6).Alignment = dbgLeft
   
      grdPedido.Columns(7).DataField = "Chave"
      grdPedido.Columns(7).Caption = "MFACHAVENFE"
      grdPedido.Columns(7).Width = 2000
      grdPedido.Columns(7).Alignment = dbgLeft
   
      grdPedido.Columns(8).DataField = "MFAMOTRESU"
      grdPedido.Columns(8).Caption = "Retorno"
      grdPedido.Columns(8).Width = 2000
      grdPedido.Columns(8).Alignment = dbgLeft
   
      grdPedido.Columns(9).DataField = "MFACODSITT"
      grdPedido.Columns(9).Caption = "Fat."
      grdPedido.Columns(9).Width = 2000
      grdPedido.Columns(9).Alignment = dbgLeft
   
      grdPedido.Columns(10).DataField = "MFAREGISTRO"
      grdPedido.Columns(10).Caption = "Registro"
      grdPedido.Columns(10).Width = 1000
      grdPedido.Columns(10).Alignment = dbgLeft
   
      grdPedido.Columns(11).DataField = "MFACODEMP"
      grdPedido.Columns(11).Caption = "MFACODEMP"
      grdPedido.Columns(11).Width = 800
      grdPedido.Columns(11).Alignment = dbgLeft
   
      grdPedido.Columns(12).DataField = "MFALOJA"
      grdPedido.Columns(12).Caption = "MFACODEMP"
      grdPedido.Columns(12).Width = 800
      grdPedido.Columns(12).Alignment = dbgLeft
   
      grdPedido.Columns(13).DataField = "MFAFILIAL"
      grdPedido.Columns(13).Caption = "MFAFILIAL"
      grdPedido.Columns(13).Width = 800
      grdPedido.Columns(13).Alignment = dbgLeft
    
      grdPedido.Columns(14).DataField = "MFACLIENTE"
      grdPedido.Columns(14).Caption = "MFACLIENTE"
      grdPedido.Columns(14).Width = 800
      grdPedido.Columns(14).Alignment = dbgLeft
    
      grdPedido.Columns(15).DataField = "MFAOBSNOTA"
      grdPedido.Columns(15).Caption = "MFAOBSNOTA"
      grdPedido.Columns(15).Width = 800
      grdPedido.Columns(15).Alignment = dbgLeft
   
      grdPedido.Columns(16).DataField = "MFANOMECONSUMIDOR"
      grdPedido.Columns(16).Caption = "MFANOMECONSUMIDOR"
      grdPedido.Columns(16).Width = 800
      grdPedido.Columns(16).Alignment = dbgLeft
   
      grdPedido.Columns(17).DataField = "MFACPFCONSUMIDOR"
      grdPedido.Columns(17).Caption = "MFACPFCONSUMIDOR"
      grdPedido.Columns(17).Width = 800
      grdPedido.Columns(17).Alignment = dbgLeft

      If TabMFA010.State = 1 Then _
         TabMFA010.Close
      sqlCabeca = "select SUM(MFAVALLIQUI) " & sqlCorpo
      TabMFA010.Open sqlCabeca, CONECTA_GLOBAL, , , adCmdText
      If Not TabMFA010.EOF Then _
         If Not IsNull(TabMFA010.Fields(0).Value) Then _
            txtValor.Text = "" & Format(TabMFA010.Fields(0).Value, strFormatacao2Digitos)
   
      If TabMFA010.State = 1 Then _
         TabMFA010.Close
      sqlCabeca = "select count(MFAdoc) " & sqlCorpo
      TabMFA010.Open sqlCabeca, CONECTA_GLOBAL, , , adCmdText
      If Not TabMFA010.EOF Then _
         If Not IsNull(TabMFA010.Fields(0).Value) Then _
            txtReg.Text = "" & TabMFA010.Fields(0).Value
      If TabMFA010.State = 1 Then _
         TabMFA010.Close
   
      If CONECTA_GLOBAL.State = 1 Then _
         CONECTA_GLOBAL.Close
      Else  'VAI LER MEGASIM SOMENTE
         grdPedido.Visible = False

         If TabPedidos.State = 1 Then _
            TabPedidos.Close

         sqlCabeca = "select count(pedido_id) as QtdePedidos from PEDIDO WITH (NOLOCK) "
         sqlCabeca = sqlCabeca & " Inner Join LANCAMENTO "
         sqlCabeca = sqlCabeca & " ON PEDIDO.PEDIDO_ID = LANCAMENTO.NUMR_DOC"

         sqlCabeca = sqlCabeca & " where PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N

         If Trim(cmbSituacaoAUX.Text) = "1" Then _
            sqlCabeca = sqlCabeca & " and status in (3,5,7) "

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
            sqlCabeca = sqlCabeca & " and dt_req >= '" & txtDtIni.Text & "'"
            sqlCabeca = sqlCabeca & " and dt_req <= '" & txtDtFim.Text & "'"
            Else
               MsgBox "Informe Data Válida."
               txtDtIni.SetFocus
               Exit Sub
         End If

         If Trim(cmbFaturaAUX.Text) <> "" Then _
            sqlCabeca = sqlCabeca & " and prefixo = '" & Trim(cmbFaturaAUX.Text) & "'"

         TabPedidos.Open sqlCabeca, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedidos.EOF Then
            If Not IsNull(TabPedidos.Fields(0).Value) Then
               txtReg.Text = "" & TabPedidos.Fields("QtdePedidos").Value

               txtReg.SelStart = 0
               txtReg.SelLength = Len(txtReg.Text)
               txtReg.BackColor = &HC0FFFF
            End If
         End If
         If TabPedidos.State = 1 Then _
            TabPedidos.Close
   
         sqlCabeca = " SELECT sum(qtd_pedida*valor_item) as TotItens"
         sqlCabeca = sqlCabeca & " FROM PEDIDO "
         sqlCabeca = sqlCabeca & " INNER JOIN PEDIDOITEM "
         sqlCabeca = sqlCabeca & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"

         sqlCabeca = sqlCabeca & " Inner Join LANCAMENTO "
         sqlCabeca = sqlCabeca & " ON PEDIDO.PEDIDO_ID = LANCAMENTO.NUMR_DOC"

         sqlCabeca = sqlCabeca & " where PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
         sqlCabeca = sqlCabeca & " and PEDIDOitem.status = 'P'"

         If Trim(cmbSituacaoAUX.Text) = "1" Then _
            sqlCabeca = sqlCabeca & " and PEDIDO.status in (3,5,7) "

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
            sqlCabeca = sqlCabeca & " and dt_req >= '" & txtDtIni.Text & "'"
            sqlCabeca = sqlCabeca & " and dt_req <= '" & txtDtFim.Text & "'"
         End If

         TabPedidos.Open sqlCabeca, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedidos.EOF Then
            If Not IsNull(TabPedidos.Fields(0).Value) Then
               txtValor.Text = "" & Format(TabPedidos.Fields("TotItens").Value, strFormatacao2Digitos)

               txtValor.SelStart = 0
               txtValor.SelLength = Len(txtValor.Text)
               txtValor.BackColor = &HC0FFFF
            End If
         End If
         If TabPedidos.State = 1 Then _
            TabPedidos.Close
   End If

   Me.Enabled = True
   Me.KeyPreview = True
   DoEvents

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
