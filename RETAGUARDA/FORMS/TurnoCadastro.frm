VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTurnoCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Turno Trabalho"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TurnoCadastro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
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
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbTurno 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   975
      TabIndex        =   1
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   2895
      TabIndex        =   2
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   1270
      ButtonWidth     =   2223
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4560
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":9331
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":A2E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":B9E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TurnoCadastro.frx":CB7D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid grdTurno 
      Bindings        =   "TurnoCadastro.frx":DDAF
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
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
         Size            =   11.25
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
   Begin MSAdodcLib.Adodc adoTurno 
      Height          =   330
      Left            =   0
      Top             =   720
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
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Turno:"
      Height          =   240
      Left            =   285
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicial:"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   750
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final:"
      Height          =   285
      Left            =   2145
      TabIndex        =   4
      Top             =   1440
      Width           =   645
   End
End
Attribute VB_Name = "frmTurnoCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   CARREGA_COMBO
   MOSTRA_TURNO

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         GRAVA_TURNO
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

Private Sub cmbTurno_Click()
'On Error GoTo ERRO_TRATA

   cmbTurnoAUX.ListIndex = cmbTurno.ListIndex

   If Trim(cmbTurnoAUX.Text) = "" Then _
      Exit Sub

   Dim TabTurno   As New ADODB.Recordset
   Dim strIni     As String
   Dim strFim     As String

   If TabTurno.State = 1 Then _
      TabTurno.Close

   SQL = "select * from TURNO "
   SQL = SQL & " where TURNO_ID = " & cmbTurnoAUX.Text
   TabTurno.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTurno.EOF Then

      txtDtIni.PromptInclude = False
      txtDtFim.PromptInclude = False

      txtDtIni.Text = "" & TabTurno.Fields("horaini").Value
      txtDtFim.Text = "" & TabTurno.Fields("horafim").Value

   End If
   If TabTurno.State = 1 Then _
      TabTurno.Close

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

Private Sub TXTDTINI_GotFocus()

   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni)
   txtDtIni.BackColor = &HC0FFFF

End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
End Sub

Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub TXTDTFIM_GotFocus()

   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim)
   txtDtFim.BackColor = &HC0FFFF

End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_TURNO
      cmbTurno.SetFocus
   End If
End Sub

Private Sub txtDtFim_LostFocus()
   txtDtFim.BackColor = &HFFFFFF
End Sub

Sub CARREGA_COMBO()
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub

Sub GRAVA_TURNO()
'On Error GoTo ERRO_TRATA

   If Trim(cmbTurnoAUX.Text) = "" Then _
      Exit Sub

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False

   If Trim(txtDtIni.Text) = "" Then _
      Exit Sub

   If Trim(txtDtFim.Text) = "" Then _
      Exit Sub

   Dim TabTurno   As New ADODB.Recordset
   Dim strIni     As String
   Dim strFim     As String

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   strIni = "" & txtDtIni.Text & ":00"
   strFim = "" & txtDtFim.Text & ":00"

   If TabTurno.State = 1 Then _
      TabTurno.Close

   SQL = "select * from TURNO "
   SQL = SQL & " where TURNO_ID = " & cmbTurnoAUX.Text
   TabTurno.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTurno.EOF Then
      SQL = "insert into TURNO "
         SQL = SQL & " (TURNO_ID,HoraIni,HoraFim)"
      SQL = SQL & " values("
         SQL = SQL & cmbTurnoAUX.Text     'TURNO_ID
         SQL = SQL & ",'" & strIni & "'"  'HoraIni
         SQL = SQL & ",'" & strFim & "'"  'HoraFim
      SQL = SQL & ")"
      Else
         SQL = "update TURNO set"
            SQL = SQL & " HoraIni = '" & strIni & "'"    'HoraIni
            SQL = SQL & ", HoraFim = '" & strFim & "'"   'HoraFim
         SQL = SQL & " where TURNO_ID = " & cmbTurnoAUX.Text
   End If
   If TabTurno.State = 1 Then _
      TabTurno.Close

   CONECTA_RETAGUARDA.Execute SQL

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   MOSTRA_TURNO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TURNO"
End Sub

Sub MOSTRA_TURNO()
'On Error GoTo ERRO_TRATA

   adoTurno.ConnectionString = AUTENTICA_GRID
   adoTurno.CommandType = adCmdText

   SQL = "select * from TURNO WITH (NOLOCK) "

   adoTurno.RecordSource = SQL
   adoTurno.Enabled = True
   adoTurno.Refresh
   grdTurno.Refresh

   grdTurno.Columns(0).DataField = "TURNO_ID"
   grdTurno.Columns(0).Caption = "Turno"
   grdTurno.Columns(0).Width = 800
   grdTurno.Columns(0).Alignment = dbgLeft

   grdTurno.Columns(1).DataField = "HoraIni"
   grdTurno.Columns(1).Caption = "HoraIni"
   grdTurno.Columns(1).Width = 1100
   grdTurno.Columns(1).Alignment = dbgLeft

   grdTurno.Columns(2).DataField = "HoraFim"
   grdTurno.Columns(2).Caption = "HoraFim"
   grdTurno.Columns(2).Width = 1100
   grdTurno.Columns(2).Alignment = dbgLeft

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TURNO"
End Sub

Sub LIMPA_TUDO()

   cmbTurno.Text = ""
   cmbTurnoAUX.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.Text = ""

   MOSTRA_TURNO
End Sub
