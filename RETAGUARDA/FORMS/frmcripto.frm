VERSION 5.00
Begin VB.Form frmCRIPTOGRAFIA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criptografia no Visual Basic"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   Icon            =   "frmcripto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Encerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeCripto 
      Caption         =   "Descriptografar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtDeCripto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtCripto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtOrigem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "MEGASIM"
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdCripto 
      Caption         =   "Criptografar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2535
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
      Left            =   6600
      TabIndex        =   7
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
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   165
   End
End
Attribute VB_Name = "frmCriptografia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   txtOrigem.Text = "MEGASIM"
   txtCripto.Text = ""
   txtDeCripto.Text = ""
End Sub

Private Sub cmdCripto_Click()
   CODIFICA
End Sub

Private Sub cmdDeCripto_Click()
   DECODIFICA
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   txtOrigem.Text = ""
   txtCripto.Text = ""
   txtDeCripto.Text = ""
End Sub

Sub CODIFICA()
   txtCripto.Text = ""
   txtCripto.Text = EncryptString("KEY", txtOrigem.Text, ENCRYPT)
End Sub

Sub DECODIFICA()
   txtDeCripto.Text = ""
   txtDeCripto.Text = EncryptString("KEY", txtCripto.Text, DECRYPT)
End Sub
