VERSION 5.00
Begin VB.Form frmBACKUP 
   BackColor       =   &H80000007&
   Caption         =   "BACKUP"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   1050
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BACKUP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   600
   End
   Begin VB.Timer TimerBar 
      Interval        =   60000
      Left            =   6960
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   6360
      Picture         =   "BACKUP.frx":5C12
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdVai 
      Caption         =   "Confirma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   120
      Picture         =   "BACKUP.frx":4D85C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1260
   End
   Begin VB.CommandButton cmbSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   7560
      Picture         =   "BACKUP.frx":4E9D7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "lblTimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1920
   End
   Begin VB.Label lbllivro 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "COPIA DE SEGURANÇA BANCO DE DADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8970
   End
   Begin VB.Menu mnuFechar 
      Caption         =   ""
      Begin VB.Menu mnuFecharSair 
         Caption         =   "&Encerrar"
      End
   End
End
Attribute VB_Name = "frmBACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

   Dim CONT_N  As Integer
   Dim FSO     As New FileSystemObject

   Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
                            (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

   Private Type NOTIFYICONDATA
      cbSize As Long
      hwnd As Long
      uid As Long
      uFlags As Long
      uCallbackMessage As Long
      hIcon As Long
      szTip As String * 64
   End Type

   Private Const NIM_ADD = &H0
   Private Const NIM_MODIFY = &H1
   Private Const NIM_DELETE = &H2
   Private Const NIF_MESSAGE = &H1
   Private Const NIF_ICON = &H2
   Private Const NIF_TIP = &H4
   Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
   Private Const WM_MOUSEMOVE = &H200
   Private Const WM_LBUTTONDBLCLK = &H203
   Private Const WM_LBUTTONDOWN = &H201
   Private Const WM_RBUTTONDOWN = &H204

Private Sub Form_Load()
   CriaIcone
   Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ExcluiIcone
End Sub

Private Sub mnuFecharSair_Click()
   End
End Sub

Private Sub Timer1_Timer()
   lblTimer.Caption = Now
End Sub

Private Sub TimerBar_Timer()
   CONT_N = CONT_N + 1
   If CONT_N >= 30 Then
      Call cmdVAI_Click
      CONT_N = 0
   End If
End Sub

Private Sub cmbSair_Click()
   Unload Me
End Sub

Private Sub cmdVAI_Click()
   Me.Enabled = False

   If Not FSO.FolderExists(App.Path & "\Backup") Then
      MsgBox "Local backup não encontrado, entre em contato com suporte."
      Exit Sub
   End If

   If Not FSO.FileExists(App.Path & "\backup\BK_DB.bat") Then
      MsgBox "Script de backup não encontrado, entre em contato com suporte."
      Exit Sub
   End If

   Shell App.Path & "\backup\BK_DB.bat"

   Me.Enabled = True

   DoEvents
End Sub

Public Sub CriaIcone()
   Dim Tic As NOTIFYICONDATA

   Tic.cbSize = Len(Tic)
   Tic.hwnd = Picture1.hwnd
   Tic.uid = 1&
   Tic.uFlags = NIF_DOALL
   Tic.uCallbackMessage = WM_MOUSEMOVE
   Tic.hIcon = Me.Icon
   Tic.szTip = "Backup Automático" & vbNullChar
   erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub ExcluiIcone()
   Dim Tic As NOTIFYICONDATA

   Tic.cbSize = Len(Tic)
   Tic.hwnd = Picture1.hwnd
   Tic.uid = 1&
   erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   X = X / Screen.TwipsPerPixelX
   'Clique duas vezes com o botão direito do mouse sobre o icone exibido
   If X = WM_RBUTTONDOWN Then _
      Me.PopupMenu mnuFechar
End Sub
