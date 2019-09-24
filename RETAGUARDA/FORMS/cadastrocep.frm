VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmCADASTROCEP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Cep"
   ClientHeight    =   1890
   ClientLeft      =   1620
   ClientTop       =   3015
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cadastrocep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   8295
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   -210
      TabIndex        =   5
      Top             =   760
      Width           =   11505
      Begin VB.TextBox txtCidade 
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
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   1
         Top             =   510
         Width           =   4095
      End
      Begin VB.TextBox txtUF 
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
         Left            =   6150
         MaxLength       =   2
         TabIndex        =   2
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txtIbge 
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
         Left            =   6990
         TabIndex        =   3
         Top             =   510
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   390
         TabIndex        =   9
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1950
         TabIndex        =   8
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6150
         TabIndex        =   7
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código IBGE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6990
         TabIndex        =   6
         Top             =   180
         Width           =   1260
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1270
      ButtonWidth     =   3466
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Salvar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consulta IBGE"
            Key             =   "IBGE"
            Object.ToolTipText     =   "Exluir Cadastro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   120
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
               Picture         =   "cadastrocep.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrocep.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrocep.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrocep.frx":9331
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrocep.frx":A43C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrocep.frx":BB39
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrocep.frx":CCD3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmCADASTROCEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         LIMPA_CEP
         txtCep.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         If txtCep.Text <> "" Then
            SQL = "select * from CEP "
            SQL = SQL & " where cep_ID = '" & txtCep.Text & "'"
            If TabCEP.State = 1 Then TabCEP.Close
            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabCEP.EOF Then
               MsgBox "Registro não encontrado.", vbOKOnly, "Atenção !!!"
               txtCep.SetFocus
               Exit Sub
               Else
                  Msg = "Confirma Exclusão?"
                  Style = vbYesNo + 32
                  Title = "Atenção !!!"
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                     TabCEP.Delete
                     LIMPA_CEP
                  End If
            End If
         End If
         txtCep.SetFocus
      Case "gravar"
         GRAVA_CEP
         txtCep.SetFocus
      Case "voltar"
        Unload Me
      Case "limpar"
         LIMPA_CEP
         txtCep.SetFocus
      Case "IBGE"
         CRITERIO_A = ""
         frmIBGECONSULTA.Show 1
         txtIBGE.Text = CRITERIO_A
         CRITERIO_A = ""
         If Trim(SQL3) <> "" Then _
            txtCidade.Text = "" & SQL3
         SQL3 = ""
         txtCep.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCep_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCep.SelStart = 0
   txtCep.SelLength = Len(txtCep)
   txtCep.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_GotFocus"
End Sub

Private Sub txtcep_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtCep.Text = "" Then
         MsgBox "Digito Inválido.", vbOKOnly, "ERRO !!!"
         txtCep.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      txtCidade.SetFocus
      Else
         If KeyAscii = 8 Then
            Else
               If KeyAscii = 45 Then
                  Else
                     If KeyAscii = 32 Then
                        Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
                     End If
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcep_KeyPress"
End Sub

Private Sub txtCep_LostFocus()
   MOSTRA_CEP
   txtCep.BackColor = &HFFFFFF
End Sub

Private Sub txtcidade_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCidade.SelStart = 0
   txtCidade.SelLength = Len(txtCidade)
   txtCidade.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtcidade_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_KeyPress"
End Sub

Private Sub txtCidade_LostFocus()
   txtCidade.BackColor = &HFFFFFF
End Sub

Private Sub txtuf_GotFocus()
'On Error GoTo ERRO_TRATA

   txtUF.SelStart = 0
   txtUF.SelLength = Len(txtUF)
   txtUF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_GotFocus"
End Sub

Private Sub txtuf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      'GRAVA_CEP
      txtIBGE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_KeyPress"
End Sub

Private Sub txtUF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtUF.Text = UCase(txtUF.Text)
   txtUF.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtuf_LostFocus"
End Sub

Private Sub txtibge_GotFocus()
'On Error GoTo ERRO_TRATA

   txtIBGE.SelStart = 0
   txtIBGE.SelLength = Len(txtIBGE)
   txtIBGE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIbge_GotFocus"
End Sub

Private Sub txtIBGE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_CEP
      txtCep.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIbge_KeyPress"
End Sub

Private Sub txtIbge_LostFocus()
'On Error GoTo ERRO_TRATA

   txtIBGE.Text = txtIBGE.Text
   txtIBGE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIbge_LostFocus"
End Sub

Private Sub MOSTRA_CEP()
'On Error GoTo ERRO_TRATA

   txtCep.PromptInclude = False
   If Trim(txtCep.Text) <> "" Then
      If CONSULTA_CEP_WEB(Trim(txtCep.Text)) = True Then
         txtCidade.Text = "" & Trim(Xcidade_A)
         txtUF.Text = "" & Trim(Xuf_A)
      End If
   End If

   If TabCEP.State = 1 Then _
      TabCEP.Close

   SQL = "select * from CEP "
   SQL = SQL & " where cep_ID = '" & Trim(txtCep.Text) & "'"
   TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCEP.EOF Then
      txtCep.Text = "" & TabCEP!CEP_ID
      txtCidade.Text = "" & TabCEP!CIDADE
      txtUF.Text = "" & TabCEP!UF
      txtIBGE.Text = "" & TabCEP!IBGE_ID
   End If
   If TabCEP.State = 1 Then _
      TabCEP.Close

   'If Trim(txtUF.Text) <> "" And Trim(txtCidade.Text) <> "" Then
   '   SQL = "select ibge_id from IBGE "
   '   SQL = SQL & " WHERE estado = '" & Trim(txtUF.Text) & "'"
   '   SQL = SQL & " and municipio = '" & Trim(txtCidade.Text) & "'"
   '   TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   '   If Not TabCEP.EOF Then _
   '      txtIBGE.Text = "" & TabCEP.Fields(0).Value
   '   If TabCEP.State = 1 Then _
   '      TabCEP.Close
   'End If
   txtCep.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Mostra_Cep"
End Sub

Private Sub LIMPA_CEP()
'On Error GoTo ERRO_TRATA

   txtCep.PromptInclude = False
   txtCep.Text = ""
   txtCep.PromptInclude = True
   txtCidade.Text = ""
   txtUF.Text = ""
   txtIBGE.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_CEP"
End Sub

Sub GRAVA_CEP()
'On Error GoTo ERRO_TRATA

   txtCep.PromptInclude = False
   If txtCep.Text = "" Then
      MsgBox "Informe o Cep"
      txtCep.SetFocus
      Exit Sub
   End If
   If txtCidade.Text = "" Then
      MsgBox "Informe a Cidade"
      txtCidade.SetFocus
      Exit Sub
   End If
   If txtUF.Text = "" Then
      MsgBox "Informe o Estado"
      txtUF.SetFocus
      Exit Sub
   End If
   If txtIBGE.Text = "" Then
      txtIBGE.Text = 0
   End If

   Acao_N = 0

   If TabCEP.State = 1 Then _
      TabCEP.Close

   SQL = "select * from CEP "
   SQL = SQL & " where cep_ID = '" & Trim(txtCep.Text) & "'"
   TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCEP.EOF Then
      Acao_N = 1
      Else: Acao_N = 2
   End If
   If TabCEP.State = 1 Then _
      TabCEP.Close

   spCEP Acao_N, Trim(txtCep.Text), Trim(txtCidade.Text), Trim(txtUF.Text), Trim(txtIBGE.Text)

   CRITERIO_A = txtCep.Text
   LIMPA_CEP

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CEP"
End Sub
