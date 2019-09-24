VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLojApre 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5925
   ControlBox      =   0   'False
   Icon            =   "LojApre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3165
      Left            =   -105
      ScaleHeight     =   3105
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   -105
      Width           =   6165
      Begin VB.Timer TimerBar 
         Interval        =   100
         Left            =   5400
         Top             =   2160
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EJS INFORMÁTICA. Todos os direitos reservados."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4440
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrigth 2000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   3165
         Left            =   720
         Picture         =   "LojApre.frx":0CCA
         Top             =   120
         Width           =   4770
      End
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "sfsfsd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmLojApre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA_INI

Exit Sub
ERRO_TRATA_INI:

   If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

   MsgBox "Erro na inicialização : " & Err.Description
   End
End Sub

Private Sub TimerBar_Timer()
'On Error GoTo ERRO_TRATA

   Dim Aux As Integer
   Aux = Pbar.Value + 10
   If Aux <= 100 Then
      Pbar.Value = Aux
      Else
         Pbar.Value = 100
         TimerBar.Enabled = False

         INICIALIZA_SISTEMA
         INICIALIZA
         Unload Me
   End If
   DoEvents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TimerBar_Timer"
   End
End Sub

Sub INICIALIZA()
   INDR_INICIALIZA = False

   DoEvents

   frmLOGON.Show 1
   Unload frmLojApre
End Sub
