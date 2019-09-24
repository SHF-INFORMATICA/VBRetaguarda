VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProducaoPerdaVenda 
   Caption         =   "RESUMO/PRODUÇAO/PERDA/VENDAS"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PRODUCAOPERDAVENA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkEncomenda 
      Caption         =   "Encomendas"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox chkPlan 
      Caption         =   "Planilha"
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
      Left            =   6120
      TabIndex        =   19
      Top             =   840
      Width           =   1095
   End
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
      Left            =   1440
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbTurno 
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ComboBox cmbFamiliaAUX 
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
      Left            =   9000
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDescProd 
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
      Height          =   375
      Left            =   3840
      MaxLength       =   100
      TabIndex        =   10
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      Picture         =   "PRODUCAOPERDAVENA.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.ComboBox cmbFamilia 
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
      Left            =   9000
      TabIndex        =   4
      Top             =   1680
      Width           =   2775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
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
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   3
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
               Picture         =   "PRODUCAOPERDAVENA.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUCAOPERDAVENA.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUCAOPERDAVENA.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUCAOPERDAVENA.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUCAOPERDAVENA.frx":A8FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRODUCAOPERDAVENA.frx":C8DF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   1200
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   1200
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid grdProd 
      Bindings        =   "PRODUCAOPERDAVENA.frx":DE28
      Height          =   3495
      Left            =   45
      TabIndex        =   5
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoProd 
      Height          =   330
      Left            =   9000
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "Grid"
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11865
      DesignHeight    =   5880
   End
   Begin MSComctlLib.ListView lstProduto 
      Height          =   3495
      Left            =   45
      TabIndex        =   20
      Top             =   2280
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   6165
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "P.Liquido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ValorVenda"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qtde.Produção"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qtde.PerdaKG"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "PerdaR$"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "VendaEstimada"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "VendaSistema"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "TotalVenda"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "%PerdaVenda"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "%PerdaProdução"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      Caption         =   "000000000000"
      Height          =   285
      Left            =   4320
      TabIndex        =   18
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label lblConta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   11235
      TabIndex        =   17
      Top             =   840
      Width           =   60
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Turno:"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Data de Registro Etiquetas"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   840
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Família:"
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final:"
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicial:"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmProducaoPerdaVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   lblDescricao.Caption = ""
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
      Case vbKeyEscape: Unload Me
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
         Unload Me
      Case "print"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub chkPlan_Click()
   If chkPlan.Value = 1 Then
      MONTA_PLANILHA
      Else: SETA_GRID
   End If
End Sub

Private Sub cmbTurno_Click()
'On Error GoTo ERRO_TRATA

   cmbTurnoAUX.ListIndex = cmbTurno.ListIndex

   If Trim(cmbTurnoAUX.Text) = "" Then _
      Exit Sub

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   txtDtIni.SetFocus

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

Private Sub txtDescProd_GotFocus()
   txtProduto.SetFocus
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni)
   txtDtIni.BackColor = &HC0FFFF

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

Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim)
   txtDtFim.BackColor = &HC0FFFF

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

Private Sub txtDtFim_LostFocus()
   CHECA_ULTIMO_DIA_MES
   txtDtFim.BackColor = &HFFFFFF
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

Private Sub txtProduto_GotFocus()

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.BackColor = &HC0FFFF

End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCESSA_DADOS_PRODUTO
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

Private Sub cmbFamilia_GotFocus()

   cmbFamilia.SelStart = 0
   cmbFamilia.SelLength = Len(cmbFamilia)
   cmbFamilia.BackColor = &HC0FFFF

End Sub

Private Sub cmbFamilia_LostFocus()
   cmbFamilia.BackColor = &HFFFFFF
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub
'===========================
Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0
   cmbFamilia.Text = ""
   cmbFamiliaAUX.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtProduto.Text = ""
   txtDescProd.Text = ""

   adoProd.ConnectionString = AUTENTICA_GRID
   adoProd.CommandType = adCmdText

   SQL = "select PRODUCAOPERDA_ID as ' ' from vwProducaoPerda WITH (NOLOCK) "
   SQL = SQL & " where PRODUCAOPERDA_ID = 0"

   adoProd.ConnectionString = AUTENTICA_GRID
   adoProd.CommandType = adCmdText

   adoProd.RecordSource = SQL
   adoProd.Enabled = True
   adoProd.Refresh
   grdProd.Refresh

Exit Sub
ERRO_TRATA:
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

   Dim TabProd As New ADODB.Recordset

   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabProd.State = 1 Then _
      TabProd.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " order by DESCRICAO"
   TabProd.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProd.EOF
      cmbFamilia.AddItem Trim(TabProd!DESCRICAO) & "-" & Trim(TabProd.Fields("familiaproduto_id").Value)
      cmbFamiliaAUX.AddItem Trim(TabProd.Fields("familiaproduto_id").Value)
      TabProd.MoveNext
   Wend
   If TabProd.State = 1 Then _
      TabProd.Close

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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Sub PROCESSA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0
   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then
      txtProduto.Enabled = True
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto)
      Exit Sub
   End If

   txtDescProd.Text = "" & Trim(DESC_PRODUTO_A)

   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescProd.ForeColor = vbRed
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            Exit Sub
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTO"
End Sub
'============
Private Sub MONTA_CONSULTA_SQL(Indr_Consulta As Boolean)
'On Error GoTo ERRO_TRATA

   Dim TabRegistro   As New ADODB.Recordset
   Dim TabGrava      As New ADODB.Recordset
   Dim Conta_N       As Long
   Dim UN_A          As String
   Dim PesoLiquidoA  As Long
   Dim TURNO_ID_N    As Integer

   SQL = "delete PRODUCAOPERDA"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PRODUCAOPERDAvenda"
   CONECTA_RETAGUARDA.Execute SQL

   adoProd.ConnectionString = AUTENTICA_GRID
   adoProd.CommandType = adCmdText

   SQL = "select PRODUCAOPERDA_ID as ' ' from vwProducaoPerda WITH (NOLOCK) "
   SQL = SQL & " where PRODUCAOPERDA_ID = 0"

   adoProd.ConnectionString = AUTENTICA_GRID
   adoProd.CommandType = adCmdText

   adoProd.RecordSource = SQL
   'adoProd.RecordSource = Nothing
   adoProd.Enabled = True
   adoProd.Refresh
   grdProd.Refresh
   adoProd.Enabled = False

'REGISTRO PRODUÇÃO
   Conta_N = 0
   lblConta.Caption = ""
   lblDescricao.Caption = ""
   If TabRegistro.State = 1 Then _
      TabRegistro.Close

   SQL = "select * from vwRegistroProducao WITH (NOLOCK) " 'REGISTRO PRODUÇÃO

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 'F' "

   If Trim(cmbTurnoAUX.Text) <> "" Then _
      SQL = SQL & " and TURNO_ID = " & cmbTurnoAUX.Text

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   If Trim(PRODUTO_ID_N) > 0 Then _
      SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_registro >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_registro <= '" & txtDtFim.Text & "'"
   End If

   TabRegistro.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabRegistro.EOF
      UN_A = "" & TabRegistro.Fields("UN").Value
      PesoLiquidoA = 0 & TabRegistro.Fields("peso_liquido").Value

      TURNO_ID_N = 0 & TabRegistro.Fields("TURNO_ID").Value
      If TURNO_ID_N = 0 Then
         TURNO_ID_N = 1
      End If

      If TabGrava.State = 1 Then _
         TabGrava.Close

      SQL = "select * from PRODUCAOPERDA WITH (NOLOCK) "
      SQL = SQL & " where TURNO_ID = " & TURNO_ID_N
      SQL = SQL & " and PRODUTO_ID = " & TabRegistro.Fields("produto_id").Value
      SQL = SQL & " and tiporegistro = 'PROD' "
      TabGrava.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabGrava.EOF Then
         SQL = "insert into PRODUCAOPERDA "
            SQL = SQL & " (PRODUCAOPERDA_ID,ESTABELECIMENTO_ID,DTREGISTRO,TURNO_ID,PRODUTO_ID,QTDE,VALOR,tiporegistro,un,pesoliquido)"
         SQL = SQL & " values ("
            SQL = SQL & MAX_ID("PRODUCAOPERDA_ID", "producaoperda", "", "", "", "") 'PRODUCAOPERDA_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                                  'ESTABELECIMENTO_ID
            SQL = SQL & ",'" & TabRegistro.Fields("dt_registro").Value & "'"        'DT_REGISTRO
            SQL = SQL & ",'" & TURNO_ID_N & "'"            'TURNO
            SQL = SQL & ",'" & TabRegistro.Fields("produto_id").Value & "'"         'PRODUTO_ID
            SQL = SQL & ",'" & tpMOEDA(TabRegistro.Fields("qtde").Value) & "'"      'QTDE
            SQL = SQL & ",'" & tpMOEDA(TabRegistro.Fields("valor").Value) & "'"     'VALOR
            SQL = SQL & ",'PROD' "
            SQL = SQL & ",'" & UN_A & "'"                                              'un
            SQL = SQL & ",'" & tpMOEDA(PesoLiquidoA) & "'"  'pesoliquido
         SQL = SQL & ")"
         Else
            SQL = "update PRODUCAOPERDA set "
               SQL = SQL & " QTDE = QTDE + '" & tpMOEDA(TabRegistro.Fields("qtde").Value) & "'"       'QTDE
               SQL = SQL & ", VALOR = VALOR + '" & tpMOEDA(TabRegistro.Fields("valor").Value) & "'"   'VALOR
            SQL = SQL & " where TURNO_ID = " & TURNO_ID_N
            SQL = SQL & " and PRODUTO_ID = " & TabRegistro.Fields("produto_id").Value
            SQL = SQL & " and tiporegistro = 'PROD' "
      End If
      If TabGrava.State = 1 Then _
         TabGrava.Close
      
      CONECTA_RETAGUARDA.Execute SQL
      Conta_N = Conta_N + 1
      lblConta.Caption = Conta_N
      lblDescricao.Caption = "Produção"
      DoEvents
      TabRegistro.MoveNext
   Wend

'============
'REGISTRO PERDA
   Conta_N = 0
   If TabRegistro.State = 1 Then _
      TabRegistro.Close

   SQL = "select * from vwRegistroPerda WITH (NOLOCK) " 'REGISTRO PERDA

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 'F' "

   If Trim(cmbTurnoAUX.Text) <> "" Then _
      SQL = SQL & " and TURNO_ID = " & cmbTurnoAUX.Text

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   If Trim(PRODUTO_ID_N) > 0 Then _
      SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_registro >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_registro <= '" & txtDtFim.Text & "'"
   End If

   TabRegistro.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabRegistro.EOF
      UN_A = "" & TabRegistro.Fields("UN").Value
      PesoLiquidoA = 0 & TabRegistro.Fields("peso_liquido").Value

      TURNO_ID_N = 0 & TabRegistro.Fields("TURNO_ID").Value
      If TURNO_ID_N = 0 Then
         TURNO_ID_N = 1
      End If

      If TabGrava.State = 1 Then _
         TabGrava.Close

      SQL = "select * from PRODUCAOPERDA WITH (NOLOCK) "
      SQL = SQL & " where TURNO_ID = " & TURNO_ID_N
      SQL = SQL & " and PRODUTO_ID = " & TabRegistro.Fields("produto_id").Value
      SQL = SQL & " and tiporegistro = 'PERDA' "
      TabGrava.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabGrava.EOF Then
         SQL = "insert into PRODUCAOPERDA "
            SQL = SQL & " (PRODUCAOPERDA_ID,ESTABELECIMENTO_ID,DtRegistro,TURNO_ID,PRODUTO_ID,QTDE,VALOR,TipoRegistro,un,pesoliquido)"
         SQL = SQL & " values ("
            SQL = SQL & MAX_ID("PRODUCAOPERDA_ID", "producaoperda", "", "", "", "") 'PRODUCAOPERDA_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                                  'ESTABELECIMENTO_ID
            SQL = SQL & ",'" & TabRegistro.Fields("dt_registro").Value & "'"        'DT_REGISTRO
            SQL = SQL & ",'" & TURNO_ID_N & "'"            'TURNO
            SQL = SQL & ",'" & TabRegistro.Fields("produto_id").Value & "'"         'PRODUTO_ID
            SQL = SQL & ",'" & tpMOEDA(TabRegistro.Fields("qtde").Value) & "'"      'QTDE
            SQL = SQL & ",'" & tpMOEDA(TabRegistro.Fields("valor").Value) & "'"     'VALOR
            SQL = SQL & ",'PERDA' "
            SQL = SQL & ",'" & UN_A & "'" 'un
            SQL = SQL & ",'" & tpMOEDA(PesoLiquidoA) & "'"  'pesoliquido
         SQL = SQL & ")"
         Else
            SQL = "update PRODUCAOPERDA set "
               SQL = SQL & " QTDE = QTDE + '" & tpMOEDA(TabRegistro.Fields("qtde").Value) & "'"       'QTDE
               SQL = SQL & ", VALOR = VALOR + '" & tpMOEDA(TabRegistro.Fields("valor").Value) & "'"   'VALOR
            SQL = SQL & " where TURNO_ID = " & TURNO_ID_N
            SQL = SQL & " and PRODUTO_ID = " & TabRegistro.Fields("produto_id").Value
            SQL = SQL & " and tiporegistro = 'PERDA' "
      End If
      If TabGrava.State = 1 Then _
         TabGrava.Close
      
      CONECTA_RETAGUARDA.Execute SQL
      Conta_N = Conta_N + 1
      lblConta.Caption = Conta_N
      lblDescricao.Caption = "Perda"
      DoEvents
      TabRegistro.MoveNext
   Wend
   If TabRegistro.State = 1 Then _
      TabRegistro.Close

'============
'REGISTRO VENDA
   Conta_N = 0
   PEDIDO_ID_N = 0
   If TabRegistro.State = 1 Then _
      TabRegistro.Close

   If chkEncomenda.Value = 1 Then
      SQL = "SELECT PEDIDOENCOMENDA.PEDIDO_ID, PEDIDOENCOMENDA.PEDIDOENCOMENDA_ID, PEDIDOENCOMENDA.DT_RECEBE, "
      SQL = SQL & " PEDIDOENCOMENDA.USUARIO_ID, PEDIDOENCOMENDA.VLR_TX_ENTREGA, vwRegistroVenda.EMPRESA_ID, "
      SQL = SQL & " vwRegistroVenda.ESTABELECIMENTO_ID, vwRegistroVenda.DESCRICAO, vwRegistroVenda.DT_REGISTRO, "
      SQL = SQL & " vwRegistroVenda.STATUS, vwRegistroVenda.SEQ_ID, vwRegistroVenda.PRODUTO_ID, vwRegistroVenda.QTDE,"
      SQL = SQL & " vwRegistroVenda.VALOR, vwRegistroVenda.CODG_PRODUTO, vwRegistroVenda.DescProduto, "
      SQL = SQL & " vwRegistroVenda.PESO_LIQUIDO, vwRegistroVenda.UN, vwRegistroVenda.PRODUCAO , "
      SQL = SQL & " vwRegistroVenda.PRECO_VENDA"
      SQL = SQL & " FROM PEDIDOENCOMENDA WITH (NOLOCK) "
      SQL = SQL & " LEFT OUTER JOIN vwRegistroVenda WITH (NOLOCK) "
      SQL = SQL & " ON PEDIDOENCOMENDA.PEDIDO_ID = vwRegistroVenda.PEDIDO_ID "
      Else: SQL = "select * from vwRegistroVenda WITH (NOLOCK) "  'REGISTRO PERDA
   End If

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status in (7,5,3)"
   SQL = SQL & " and producao = 1"

   'If Trim(cmbTurnoAUX.Text) <> "" Then _
      SQL = SQL & " and TURNO_ID = " & cmbTurnoAUX.Text

   If Trim(cmbFamiliaAUX.Text) <> "" Then _
      SQL = SQL & " and familiaproduto_id = " & cmbFamiliaAUX.Text

   If Trim(PRODUTO_ID_N) > 0 Then _
      SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_registro >= '" & txtDtIni.Text & "'"
      SQL = SQL & " and dt_registro <= '" & txtDtFim.Text & "'"
   End If

   TabRegistro.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabRegistro.EOF
      If PEDIDO_ID_N <> TabRegistro.Fields("pedido_id").Value Then
         PEDIDO_ID_N = TabRegistro.Fields("pedido_id").Value
         UN_A = "" & TabRegistro.Fields("UN").Value
         PesoLiquidoA = 0 & TabRegistro.Fields("peso_liquido").Value
   
         If Mid(TabRegistro.Fields("dt_registro").Value, 12, 2) < 13 Then
            TURNO_ID_N = 1
            Else: TURNO_ID_N = 2
         End If
   
         If TabGrava.State = 1 Then _
            TabGrava.Close
   
         SQL = "select * from PRODUCAOPERDA WITH (NOLOCK) "
         SQL = SQL & " where TURNO_ID = " & TURNO_ID_N
         SQL = SQL & " and PRODUTO_ID = " & TabRegistro.Fields("produto_id").Value
         SQL = SQL & " and tiporegistro = 'VENDA' "
         TabGrava.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabGrava.EOF Then
            SQL = "insert into PRODUCAOPERDA "
               SQL = SQL & " (PRODUCAOPERDA_ID,ESTABELECIMENTO_ID,DtRegistro,TURNO_ID,PRODUTO_ID,QTDE,VALOR,TipoRegistro,un,pesoliquido)"
            SQL = SQL & " values ("
               SQL = SQL & MAX_ID("PRODUCAOPERDA_ID", "producaoperda", "", "", "", "") 'PRODUCAOPERDA_ID
               SQL = SQL & "," & ESTABELECIMENTO_ID_N                                  'ESTABELECIMENTO_ID
               SQL = SQL & ",'" & TabRegistro.Fields("dt_registro").Value & "'"        'DT_REGISTRO
               SQL = SQL & ",'" & TURNO_ID_N & "'"            'TURNO
               SQL = SQL & ",'" & TabRegistro.Fields("produto_id").Value & "'"         'PRODUTO_ID
               SQL = SQL & ",'" & tpMOEDA(TabRegistro.Fields("qtde").Value) & "'"      'QTDE
               SQL = SQL & ",'" & tpMOEDA(TabRegistro.Fields("valor").Value) & "'"     'VALOR
               SQL = SQL & ",'VENDA' "
               SQL = SQL & ",'" & UN_A & "'"                                           'un
               SQL = SQL & ",'" & tpMOEDA(PesoLiquidoA) & "'"  'pesoliquido
            SQL = SQL & ")"
            Else
               SQL = "update PRODUCAOPERDA set "
                  SQL = SQL & " QTDE = QTDE + '" & tpMOEDA(TabRegistro.Fields("qtde").Value) & "'"       'QTDE
                  SQL = SQL & ", VALOR = VALOR + '" & tpMOEDA(TabRegistro.Fields("valor").Value) & "'"   'VALOR
               SQL = SQL & " where TURNO_ID = " & TURNO_ID_N
               SQL = SQL & " and PRODUTO_ID = " & TabRegistro.Fields("produto_id").Value
               SQL = SQL & " and tiporegistro = 'VENDA' "
         End If
         If TabGrava.State = 1 Then _
            TabGrava.Close
   
         CONECTA_RETAGUARDA.Execute SQL
         Conta_N = Conta_N + 1
         lblConta.Caption = Conta_N
         lblDescricao.Caption = "Venda"
      End If
      DoEvents
      TabRegistro.MoveNext
   Wend
   If TabRegistro.State = 1 Then _
      TabRegistro.Close

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstProduto.Visible = False
   lblConta.Caption = ""
   lblDescricao.Caption = ""
   lstProduto.Visible = False

   adoProd.Enabled = False
   adoProd.ConnectionString = AUTENTICA_GRID
   adoProd.CommandType = adCmdText

   SQL = "select codg_produto as Codigo,DescProduto,peso_liquido as 'PesoLiquido',valor as Valor,qtde as Peso,"
   SQL = SQL & " TURNO_ID as Turno,TipoRegistro  from vwProducaoPerda WITH (NOLOCK) "
   SQL = SQL & " order by codg_produto,tiporegistro desc"

   adoProd.ConnectionString = AUTENTICA_GRID
   adoProd.CommandType = adCmdText

   adoProd.RecordSource = SQL
   adoProd.Enabled = True
   adoProd.Refresh
   grdProd.Refresh

   grdProd.Columns(0).DataField = "CODG_PRODUTO"
   grdProd.Columns(0).Caption = "Código"
   grdProd.Columns(0).Width = 1500
   grdProd.Columns(0).Alignment = dbgLeft

   grdProd.Columns(1).DataField = "descproduto"
   grdProd.Columns(1).Caption = "Produto"
   grdProd.Columns(1).Width = 6000
   grdProd.Columns(1).Alignment = dbgLeft

   grdProd.Columns(2).DataField = "peso_liquido"
   grdProd.Columns(2).Caption = "PesoLiquido"
   grdProd.Columns(2).Width = 2000
   grdProd.Columns(2).Alignment = dbgRight

   grdProd.Columns(3).DataField = "valor"
   grdProd.Columns(3).Caption = "Valor"
   grdProd.Columns(3).Width = 2000
   grdProd.Columns(3).Alignment = dbgRight
   grdProd.Columns(3).NumberFormat = "#,##0.00;#,##0.00;#,##0.00"

   grdProd.Columns(4).DataField = "qtde"
   grdProd.Columns(4).Caption = "Qtde"
   grdProd.Columns(4).Width = 2000
   grdProd.Columns(4).Alignment = dbgRight

   grdProd.Columns(5).DataField = "TURNO_ID"
   grdProd.Columns(5).Caption = "Turno"
   grdProd.Columns(5).Width = 1000
   grdProd.Columns(5).Alignment = dbgCenter

   lblConta.Caption = ""
   lblDescricao.Caption = ""
   grdProd.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Sub MONTA_PLANILHA()
'On Error GoTo ERRO_TRATA

   Dim TabProd             As New ADODB.Recordset
   Dim TabTemp             As New ADODB.Recordset
   Dim QtdeProducao_N      As Double
   Dim QtdePerda_N         As Double
   Dim QtdeVenda_N         As Double
   Dim QtdeVendaEstimada_N As Double
   Dim QtdeVendaSistema_N  As Double
   Dim TotalVenda_N        As Double
   Dim PercVenda_N         As Double
   Dim PercProducao_N      As Double
   Dim PROD_ID_N           As Long

   lblConta.Caption = ""
   lblDescricao.Caption = ""
   lstProduto.ListItems.Clear
   grdProd.Visible = False
   QtdeProducao_N = 0
   QtdePerda_N = 0
   QtdeVenda_N = 0
   QtdeVendaEstimada_N = 0
   QtdeVendaSistema_N = 0
   TotalVenda_N = 0
   PercVenda_N = 0
   PercProducao_N = 0
   PROD_ID_N = 0

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUCAOPERDAVENDA", "U") = False Then
      SQL = " CREATE TABLE [dbo].[PRODUCAOPERDAVENDA]("
      SQL = SQL & " [PRODUTO_ID] [nchar](10) NOT NULL,"
      SQL = SQL & " [QtdeProducao] [float] NULL,"
      SQL = SQL & " [QtdePerda] [float] NULL,"
      SQL = SQL & " [QtdeVenda] [float] NULL,"
      SQL = SQL & " [QtdeVendaEstimada] [float] NULL,"
      SQL = SQL & " [QtdeVendaSistema] [float] NULL,"
      SQL = SQL & " [TotalVenda] [float] NULL,"
      SQL = SQL & " [PercVenda] [float] NULL,"
      SQL = SQL & " [PercProducao] [float] NULL,"
      SQL = SQL & " CONSTRAINT [PK_PRODUCAOPERDAVENDA] PRIMARY KEY CLUSTERED([PRODUTO_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   SQL = "delete PRODUCAOPERDAVENDA"
   CONECTA_RETAGUARDA.Execute SQL

   If TabProd.State = 1 Then _
      TabProd.Close

   SQL = "select * from vwProducaoPerda WITH (NOLOCK) "
   SQL = SQL & " order by codg_produto,tiporegistro "
   TabProd.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProd.EOF
      QtdeProducao_N = 0
      QtdePerda_N = 0
      QtdeVenda_N = 0
      QtdeVendaEstimada_N = 0
      QtdeVendaSistema_N = 0
      TotalVenda_N = 0
      PercVenda_N = 0
      PercProducao_N = 0

      If TabProd.Fields("tiporegistro").Value = "PROD" Then _
         QtdeProducao_N = 0 & TabProd.Fields("qtde").Value

      If TabProd.Fields("tiporegistro").Value = "PERDA" Then _
         QtdePerda_N = 0 & TabProd.Fields("qtde").Value

      If TabProd.Fields("tiporegistro").Value = "VENDA" Then _
         QtdeVenda_N = 0 & TabProd.Fields("qtde").Value

      If TabTemp.State = 1 Then _
         TabTemp.Close
   
      SQL = "select * from PRODUCAOPERDAVENDA"
      SQL = SQL & " where produto_id = " & TabProd.Fields("PRODUTO_ID").Value
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "insert into PRODUCAOPERDAVENDA"
            SQL = SQL & " (PRODUTO_ID,QtdeProducao,QtdePerda,QtdeVenda,QtdeVendaEstimada,QtdeVendaSistema,TotalVenda,PercVenda,PercProducao)"
         SQL = SQL & " values("
            SQL = SQL & TabProd.Fields("PRODUTO_ID").Value
            SQL = SQL & ",'" & tpMOEDA(QtdeProducao_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(QtdePerda_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(QtdeVenda_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(QtdeVendaEstimada_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(QtdeVendaSistema_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(TotalVenda_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(PercVenda_N) & "'"
            SQL = SQL & ",'" & tpMOEDA(PercProducao_N) & "'"
         SQL = SQL & ")"
         Else
            SQL = "update PRODUCAOPERDAVENDA set "
               SQL = SQL & " QtdeProducao = QtdeProducao + '" & tpMOEDA(QtdeProducao_N) & "'"
               SQL = SQL & ",QtdePerda = QtdePerda + '" & tpMOEDA(QtdePerda_N) & "'"
               SQL = SQL & ",QtdeVenda = QtdeVenda + '" & tpMOEDA(QtdeVenda_N) & "'"
               SQL = SQL & ",QtdeVendaEstimada = QtdeVendaEstimada + '" & tpMOEDA(QtdeVendaEstimada_N) & "'"
               SQL = SQL & ",QtdeVendaSistema = QtdeVendaSistema + '" & tpMOEDA(QtdeVendaSistema_N) & "'"
               SQL = SQL & ",TotalVenda = TotalVenda + '" & tpMOEDA(TotalVenda_N) & "'"
               SQL = SQL & ",PercVenda = '" & tpMOEDA(PercVenda_N) & "'"
               SQL = SQL & ",PercProducao = '" & tpMOEDA(PercProducao_N) & "'"
            SQL = SQL & " where produto_id = " & TabProd.Fields("PRODUTO_ID").Value
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

CONECTA_RETAGUARDA.Execute SQL

      TabProd.MoveNext
   Wend
   If TabProd.State = 1 Then _
      TabProd.Close

   SQL = "select PRODUCAOPERDAVENDA.*, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO,PRODUTO.PESO_LIQUIDO,PRODUTO.PRECO_VENDA"
   SQL = SQL & " from PRODUCAOPERDAVENDA "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PRODUCAOPERDAVENDA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " order by descricao "
   TabProd.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProd.EOF
      Set item = lstProduto.ListItems.Add(, "seq." & TabProd.Fields("produto_id").Value, TabProd.Fields("codg_produto").Value)
      item.SubItems(1) = "" & TabProd.Fields("descricao").Value
      item.SubItems(2) = "" & TabProd.Fields("PESO_LIQUIDO").Value
      item.SubItems(3) = "" & TabProd.Fields("preco_venda").Value

      item.SubItems(4) = "" & TabProd.Fields("QtdeProducao").Value

      item.SubItems(5) = "" & TabProd.Fields("Qtdeperda").Value
      item.SubItems(6) = "" & TabProd.Fields("Qtdeperda").Value * TabProd.Fields("preco_venda").Value

      item.SubItems(7) = "" & TabProd.Fields("QtdeVendaEstimada").Value
      item.SubItems(8) = "" & TabProd.Fields("QtdeVendaSistema").Value
      
      item.SubItems(9) = "" & TabProd.Fields("TotalVenda").Value

      item.SubItems(10) = "" & TabProd.Fields("PercVenda").Value
      item.SubItems(11) = "" & TabProd.Fields("PercProducao").Value
'converter tudo que é unidade para kg
'peso esta sendo cadastrado no cadastro de produtos
'pega a qtde e multiplicar pelo peso no cadastro de produto quando o produto é unidade
'colar dados da venda tambem
'      End If

      TabProd.MoveNext
   Wend
   If TabProd.State = 1 Then _
      TabProd.Close

   lblConta.Caption = ""
   lblDescricao.Caption = ""
   lstProduto.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_PLANILHA"
End Sub
