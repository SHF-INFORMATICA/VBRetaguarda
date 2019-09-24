VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCAIXATESORARIAAFREABRE 
   Caption         =   "Reabertura Caixa"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "CAIXATESORARIAAFREABRE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8055
      Begin MSMask.MaskEdBox txtDtDig 
         Height          =   450
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   794
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reabertura:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   360
         Width           =   1365
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXATESORARIAAFREABRE.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXATESORARIAAFREABRE.frx":6DAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1270
      ButtonWidth     =   2910
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "$Abrir Caixa"
            Key             =   "abrir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmCAIXATESORARIAAFREABRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   INICIALIZA_TELA
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
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
   Select Case Button.key
      Case "abrir"
         REABRE_ABRE_CAIXA
      Case "sair"
         Unload Me
   End Select
End Sub

Private Sub txtdtdig_GotFocus()
   txtDtDig.PromptInclude = True
End Sub

Private Sub txtdtdig_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtdtdig_LostFocus()
   txtDtDig.PromptInclude = True
   If Not IsDate(txtDtDig.Text) Then
      txtDtDig.PromptInclude = False
         txtDtDig.Text = Date
      txtDtDig.PromptInclude = True
   End If
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyPress"
End Sub

Sub INICIALIZA_TELA()
'On Error GoTo ERRO_TRATA

   If USUARIO_ID_N <= 0 Then
      MsgBox "Usuário inexistente."
      Unload Me
      Exit Sub
   End If
   txtDtDig.Text = Date

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_TELA"
End Sub

Sub REABRE_ABRE_CAIXA()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma ?"
   PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
   If RESPOSTA <> vbYes Then _
      Exit Sub

   SQL = "update " & SQL & " set "
   SQL = SQL & " dt_fechamento = NULL"
   SQL = SQL & ",status = '' "
   SQL = SQL & " where dt_fechamento >= '" & DMA(txtDtDig.Text, "I") & "'"
   SQL = SQL & " and dt_fechamento <= '" & DMA(txtDtDig.Text, "F") & "'"
   CONECTA_RETAGUARDA.Execute SQL

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_TELA"
End Sub
