VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAPRESENTA 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13140
   ControlBox      =   0   'False
   Icon            =   "APRESENTA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   0
      ScaleHeight     =   1290
      ScaleMode       =   0  'User
      ScaleWidth      =   13077.03
      TabIndex        =   0
      Top             =   0
      Width           =   13140
      Begin MSComctlLib.ProgressBar Pbar 
         Height          =   150
         Left            =   0
         TabIndex        =   2
         Top             =   825
         Width           =   13110
         _ExtentX        =   23125
         _ExtentY        =   265
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Timer TimerBar 
         Interval        =   100
         Left            =   3120
         Top             =   120
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EJS/MEGA-SIM. Todos os direitos reservados."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   975
         Left            =   11640
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrigth 2011"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAPRESENTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA_INI

   'If Trim(Command) <> "" Then
      INICIALIZA_SISTEMA
      INICIALIZA
   '   Else
         'MsgBox "VERIFICAR INICIALIZAÇÃO DO EXECUTAVEL."
         'End
   'End If

Exit Sub
ERRO_TRATA_INI:

   If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

   MsgBox "Erro na inicialização : " & Err.Description
   End
End Sub

Sub INICIALIZA()
   INDR_INICIALIZA = False

   DoEvents

   frmLOGON.Show
   Unload frmAPRESENTA
End Sub
