VERSION 5.00
Begin VB.Form frmLojECFAliquota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programação de Alíquotas"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "LojAliquota.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "ISS"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ICMS"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtAliquota 
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota:"
         Height          =   195
         Left            =   285
         TabIndex        =   1
         Top             =   315
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmLojECFAliquota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim SITUACAOTRIBUTARIA As String

Private Sub cmdOK_Click()
'On Error GoTo ERRO_TRATA

    If Option1.Value = True Then
        SITUACAOTRIBUTARIA = "0"
    Else
        SITUACAOTRIBUTARIA = "1"
    End If
    
    RETORNO_ECF = Bematech_FI_ProgramaAliquota(txtAliquota, SITUACAOTRIBUTARIA)
    Call VerificaRetornoImpressora("", "", "Programação de Alíquotas")

    txtAliquota.Text = ""
    txtAliquota.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdFechar_Click()
'On Error GoTo ERRO_TRATA

    Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub txtAliquota_GotFocus()
'On Error GoTo ERRO_TRATA

    Call DestacaTexto(txtAliquota)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub
