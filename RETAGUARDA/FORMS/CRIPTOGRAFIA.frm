VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCRIPTO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desbloquear MEGASIM"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CRIPTOGRAFIA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVerifica 
      Caption         =   "Verifica Sistama"
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   2280
      Width           =   2175
   End
   Begin MSMask.MaskEdBox txtPARDecodifica 
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtPAR 
      Height          =   405
      Left            =   2040
      TabIndex        =   10
      Text            =   "MEGASIM"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtCriptoBanco 
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   405
      Left            =   2040
      TabIndex        =   8
      Text            =   "MEGASIM"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdAtivar 
      Caption         =   "Ativar"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Encerrar"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeCripto 
      Caption         =   "Descriptografar"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtCripto 
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdCripto 
      Caption         =   "Criptografar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin MSMask.MaskEdBox txtOrigem 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDeCripto 
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DesCrypBanco:"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CrypBanco:"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   11
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   5640
      TabIndex        =   5
      Top             =   360
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   165
   End
End
Attribute VB_Name = "frmCRIPTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   txtOrigem.PromptInclude = False
   txtOrigem.Text = "" & Date
   txtCripto.Text = ""
   txtDeCripto.PromptInclude = False
   txtDeCripto.Text = ""
   txtCNPJ.Text = "" & CNPJ_EMPRESA_N
   txtPARDecodifica.PromptInclude = False
   txtPARDecodifica.Text = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select par from EMPRESA WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtPAR.Text = "" & Trim(TabTemp.Fields(0).Value)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtPARDecodifica.Text = "" & EncryptString(CNPJ_EMPRESA_N, txtPAR.Text, DECRYPT)
   txtPARDecodifica.PromptInclude = True
   txtPARDecodifica.Refresh

End Sub

Private Sub cmdVerifica_Click()
   VERIFICA_SISTEMA
End Sub

Private Sub cmdAtivar_Click()

   If Trim(txtCripto.Text) = "" Then
      MsgBox "Informar dados !!!"
      Exit Sub
   End If

   GRAVA_PRIMEIRA_DATA Trim(txtCripto.Text)
End Sub

Private Sub cmdCripto_Click()
   CODIFICA
End Sub

Private Sub cmdDeCripto_Click()
   DECODIFICA (txtCripto.Text)
End Sub

Private Sub cmdSair_Click()
   End
End Sub

Private Sub Command1_Click()
   txtOrigem.PromptInclude = False
   txtOrigem.Text = ""
   txtCripto.Text = ""
   txtDeCripto.PromptInclude = False
   txtDeCripto.Text = ""
End Sub

Sub CODIFICA()
'On Error GoTo ERRO_TRATA

   txtOrigem.PromptInclude = False
   txtCripto.Text = ""
   txtCripto.Text = EncryptString(CNPJ_EMPRESA_N, Trim(txtOrigem.Text), ENCRYPT)
   txtOrigem.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Caption, "CODIFICA"
End Sub

Function DECODIFICA(VACA_VEIA As String) As String
'On Error GoTo ERRO_TRATA

   txtDeCripto.PromptInclude = False
   txtDeCripto.Text = ""
   DECODIFICA = "" & EncryptString(CNPJ_EMPRESA_N, VACA_VEIA, DECRYPT)
   txtDeCripto.Text = "" & DECODIFICA
   txtDeCripto.PromptInclude = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Caption, "DECODIFICA"
End Function
