VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmERROS 
   Caption         =   "Relação erros operacionais do Sistema"
   ClientHeight    =   6330
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ERROSDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin PVMaskEditLib.PVMaskEdit txtDtFim 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1095
      _Version        =   524288
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   253
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Text            =   ""
      Mask            =   "##/##/####"
   End
   Begin PVMaskEditLib.PVMaskEdit txtDtIni 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1095
      _Version        =   524288
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   253
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Text            =   ""
      Mask            =   "##/##/####"
   End
   Begin Threed.SSCommand optID 
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   262144
      Caption         =   "ID"
   End
   Begin VB.TextBox txtDesc 
      Height          =   405
      Left            =   7200
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtID 
      Height          =   405
      Left            =   4320
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdERROS 
      Bindings        =   "ERROSDB.frx":5C12
      Height          =   4455
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   1270
      ButtonWidth     =   2249
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
            ImageIndex      =   3
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
      EndProperty
      BorderStyle     =   1
      Begin Threed.SSCheck chkOrdem 
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   262144
         Caption         =   "Ordem Decresente"
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6000
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ERROSDB.frx":5C29
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ERROSDB.frx":7051
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ERROSDB.frx":80E0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc adoERROS 
      Height          =   735
      Left            =   10320
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      DesignWidth     =   11370
      DesignHeight    =   6330
   End
   Begin Threed.SSCommand optData 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   262144
      Caption         =   "Por Data"
   End
   Begin Threed.SSCommand optDesc 
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   262144
      Caption         =   "Descrição"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   13
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Top             =   840
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "à"
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      Index           =   1
      X1              =   0
      X2              =   11415
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      Index           =   0
      X1              =   0
      X2              =   11415
      Y1              =   720
      Y2              =   735
   End
End
Attribute VB_Name = "frmERROS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   SQL = "select * from ERRO WITH (NOLOCK)"
   SQL = SQL & " where erro_id > 0 "
   SQL = SQL & " order by erro_id desc"

   SETA_GRID
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         SQL3 = ""
         txtDtIni.Text = ""
         txtDtFim.Text = ""
         txtID.Text = ""
         txtDesc.Text = ""
         optData.Value = False
         optID.Value = False
         optDesc.Value = False
         chkOrdem.Value = 0

         SQL = "select * from ERRO WITH (NOLOCK)"
         SQL = SQL & " where erro_id > 0 "
         SQL = SQL & " order by erro_id desc"

         SETA_GRID
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub chkOrdem_Click(Value As Integer)
   If chkOrdem.Value = 0 Then
      SQL3 = ""
      Else: SQL3 = "desc"
   End If
End Sub

Private Sub optData_Click()
   SQL = "select * from ERRO WITH (NOLOCK)"
   SQL = SQL & " where erro_id > 0 "

   If Trim(txtDtIni.Text) <> "" Then _
      If IsDate(txtDtIni.Text) Then _
         SQL = SQL & " and data >= '" & DMA(txtDtIni.Text) & "'"

   If Trim(txtDtFim.Text) <> "" Then _
      If IsDate(txtDtFim.Text) Then _
         SQL = SQL & " and data <= '" & DMA(txtDtFim.Text) & "'"

   SQL = SQL & " order by data " & SQL3

   SETA_GRID
End Sub

Private Sub optID_Click()
   SQL = "select * from ERRO WITH (NOLOCK)"
   SQL = SQL & " where erro_id > 0 "

   If Trim(txtID.Text) <> "" Then _
      If IsNumeric(txtID.Text) Then _
         SQL = SQL & " and erro_id = " & txtID.Text

   SQL = SQL & " order by erro_id " & SQL3

   SETA_GRID
End Sub

Private Sub optDesc_Click()
   SQL = "select * from ERRO WITH (NOLOCK)"
   SQL = SQL & " where erro_id > 0 "

   If Trim(txtDesc.Text) <> "" Then
      CRITERIO_A = Chr$(39) & txtDesc.Text & "%" & Chr(39)
      SQL = SQL & " and descricao like " & CRITERIO_A
   End If

   SQL = SQL & " order by descricao " & SQL3

   SETA_GRID
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   adoERROS.Enabled = True
   adoERROS.ConnectionString = AUTENTICA_GRID

   adoERROS.RecordSource = SQL
   adoERROS.Enabled = True
   adoERROS.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcpf_KeyDown"
End Sub
