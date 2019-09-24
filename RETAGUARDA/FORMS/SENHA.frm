VERSION 5.00
Begin VB.Form frmSenha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorização"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SENHA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "@"
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe Senha:"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   1755
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   txtSenha.Text = ""
End Sub

Private Sub cmdOK_Click()
   CRITERIO_A = Trim(txtSenha.Text)
   Unload Me
End Sub

Private Sub cmdCancela_Click()
   CRITERIO_A = ""
   Unload Me
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Call cmdOK_Click
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtSenha_KeyPress"
End Sub


