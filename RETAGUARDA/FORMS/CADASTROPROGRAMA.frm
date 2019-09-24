VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROPROGRAMA 
   Caption         =   "Cadastro de Programa Menu"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROPROGRAMA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkImprime 
      Caption         =   "Gera Impressão?"
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtDesc 
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtID 
      Height          =   405
      Left            =   100
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1270
      ButtonWidth     =   2487
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
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Excluir"
            Key             =   "matar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6480
         Top             =   240
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
               Picture         =   "CADASTROPROGRAMA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPROGRAMA.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPROGRAMA.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPROGRAMA.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPROGRAMA.frx":9EFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPROGRAMA.frx":B051
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPROGRAMA.frx":B4A3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "CADASTROPROGRAMA.frx":C70B
      Height          =   5655
      Left            =   105
      TabIndex        =   5
      Top             =   2160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "MENUID"
         Caption         =   "ID"
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
         DataField       =   "DescMenu"
         Caption         =   "Descrição"
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
         Caption         =   "Gera Impressão"
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
            ColumnWidth     =   2234,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4919,811
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADOMenu 
      Height          =   330
      Left            =   4200
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Grid Menu"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Menu ID"
      Height          =   285
      Left            =   100
      TabIndex        =   4
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "frmCADASTROPROGRAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
      Case "gravar"
         GRAVA_TUDO
         SETA_GRID
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTID_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_DADOS
      txtDesc.SetFocus
   End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0

      GRAVA_TUDO
      SETA_GRID

      txtID.SetFocus
   End If
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   ADOMenu.Enabled = True
   ADOMenu.ConnectionString = AUTENTICA_GRID

   SQL = "select * from MENU"
   SQL = SQL & " order by DESCMENU"

   ADOMenu.RecordSource = SQL
   ADOMenu.Enabled = True
   ADOMenu.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtID.Text = ""
   txtDesc.Text = ""
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   If Trim(txtID.Text) <> "" And Trim(txtDesc.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from MENU "
      SQL = SQL & " where menuid = '" & Trim(txtID.Text) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         SQL = "insert into MENU values("
            SQL = SQL & "'" & Trim(txtID.Text) & "'"
            SQL = SQL & ",'" & Trim(txtDesc.Text) & "'"
         SQL = SQL & ")"
         Else
            SQL = "update MENU set "
            SQL = SQL & " descmenu = '" & Trim(txtDesc.Text) & "'"
            SQL = SQL & " where menuid = '" & Trim(txtID.Text) & "'"
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      LIMPA_TUDO
      Else: MsgBox "Informe dados corretamente."
   End If
   txtID.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Sub MOSTRA_DADOS()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from MENU "
   SQL = SQL & " where menuid = '" & Trim(txtID.Text) & "'"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      txtDesc.Text = Trim(TabConsulta.Fields("descmenu").Value)

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS"
End Sub
