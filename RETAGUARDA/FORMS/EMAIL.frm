VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastra e-mail"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EMAIL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7935
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   -240
      TabIndex        =   3
      Top             =   720
      Width           =   8535
      Begin VB.TextBox txtEmail 
         Height          =   405
         Left            =   1395
         TabIndex        =   0
         Top             =   840
         Width           =   6855
      End
      Begin MSDataGridLib.DataGrid GridEmail 
         Bindings        =   "EMAIL.frx":5C12
         Height          =   2175
         Left            =   285
         TabIndex        =   1
         Top             =   1320
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3836
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
         ColumnCount     =   1
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   7395,024
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc ADOeMail 
         Height          =   330
         Left            =   360
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Caption         =   "Grid Cabeça"
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
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   555
         TabIndex        =   4
         Top             =   840
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   1270
      ButtonWidth     =   2223
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
            Caption         =   "Voltar"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "excluir"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4440
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EMAIL.frx":5C29
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EMAIL.frx":6DC3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EMAIL.frx":7E52
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EMAIL.frx":8F5D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EMAIL.frx":A1C5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   If Trim(CNPJCPF_A) <> "" Then _
      MOSTRA_DADOS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub GridEmail_DblClick()
On Error Resume Next

   txtEmail = Trim(GridEmail.Columns(0).Text)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "excluir"
         If Trim(txtEmail.Text) <> "" Then
            If TabEmail.State = 1 Then _
               TabEmail.Close

               SQL = "select * from EMAIL "
               SQL = SQL & " where email = '" & Trim(txtEmail.Text) & "'"
               TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabEmail.EOF Then
                  Msg = "Confirma exclusão de registro ? "
                  PERGUNTA Msg, vbYesNo + 32, "Cadastro e-mail", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then

                     spEmail 3, TabEmail.Fields("email_ID").Value, Trim(txtEmail.Text), TabEmail.Fields("PESSOA_ID").Value

                     MOSTRA_DADOS
                  End If
               End If
            If TabEmail.State = 1 Then _
               TabEmail.Close
         End If
      Case "sair"
         Unload Me
      Case "limpar"
         txtEmail.Text = ""
      Case "gravar"
         GRAVA_EMAIL
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_EMAIL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEmail_KeyPress"
End Sub

Sub MOSTRA_DADOS()
'On Error GoTo ERRO_TRATA

   HORA_INI = Time
   NUMR_SEQ_N = 0

   MOSTRA_RODAPE "Aguarde, Pesquisando ...", "", "", "", ""

   SETA_GRID

   HORA_FIM = Time

   MOSTRA_RODAPE "OK", "Duração Consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss"), "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0
   txtProp.Text = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from PESSOA"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtProp.Text = Trim(TabTemp.Fields("descricao").Value) & " - " & Trim(TabTemp.Fields("CNPJCPF").Value)
      PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select email from EMAIL "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N

   ADOeMail.ConnectionString = AUTENTICA_GRID
   ADOeMail.CommandType = adCmdText

   ADOeMail.RecordSource = SQL
   ADOeMail.Enabled = True
   ADOeMail.Refresh

   GridEmail.Columns(0).DataField = "EMAIL"
   GridEmail.Columns(0).Caption = "e-mail"
   GridEmail.Columns(0).Width = 7395.024

   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub GRAVA_EMAIL()
'On Error GoTo ERRO_TRATA

   If Trim(txtProp.Text) = "" Then
      MsgBox "Realizar inclusão do CNPJ ou CPF antes de cadastrar e-mail."
      Exit Sub
   End If
   If Trim(CNPJCPF_A) = "" Then
      MsgBox "Impossível prosseguir, cnpjcpf inválido."
      Exit Sub
   End If
   If Trim(txtEmail.Text) = "" Then
      MsgBox "Impossível prosseguir, e-mail inválido."
      Exit Sub
   End If

   PESSOA_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select pessoa_id from PESSOA "
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      PESSOA_ID_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from EMAIL "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and email = '" & Trim(txtEmail.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      spEmail 2, TabEmail.Fields("email_ID").Value, Trim(txtEmail.Text), TabEmail.Fields("PESSOA_ID").Value
      Else: spEmail 1, 0, Trim(txtEmail.Text), PESSOA_ID_N
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtEmail.Text = ""
   MOSTRA_DADOS
   txtEmail.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_EMAIL"
End Sub
