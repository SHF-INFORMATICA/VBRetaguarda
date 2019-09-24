VERSION 5.00
Begin VB.Form frmRegistro 
   BackColor       =   &H00000000&
   Caption         =   "Formulário de registro"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   Icon            =   "frmRegistro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsair 
      Caption         =   "Sair do programa"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdregistrardepois 
      Caption         =   "Registrar Depois"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdregistraragora 
      Caption         =   "Registrar Agora"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtcodigoliberacao 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtcodigodoprograma 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtdiasquefaltampararegistrar 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   "Liberação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   "Código :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   465
      TabIndex        =   4
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   "dias para registrar o programa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   3450
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   "Faltam:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmRegistro.frx":47C4A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1125
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmregistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdregistraragora_Click()

If txtcodigoliberacao.Text = "" Then
   txtcodigoliberacao.SetFocus
   Exit Sub
End If

frmMenu.alock.LiberationKey = txtcodigoliberacao.Text

If Not frmMenu.alock.RegisteredUser Then
  MsgBox "Chave de LIBERAÇÃO INCORRETA", vbOKOnly + vbCritical, "Chave Liberação Incorreta"
   txtcodigoliberacao.SetFocus
Else
  MsgBox "REGISTRO EFETUADO COM SUCESSO !", vbExclamation, "Registro OK"
  frmMenu.lblAviso.Visible = False
  frmMenu.Caption = "VERSÃO REGISTRADA"
  frmMenu.lblregistro(2).Enabled = False
  Unload Me

End If
End Sub

Private Sub cmdregistrardepois_Click()
Unload Me
End Sub

Private Sub cmdsair_Click()
 End
End Sub

Private Sub Form_Load()

Dim diasQueFaltaParaRegistrar As Integer

diasQueFaltaParaRegistrar = 0

diasQueFaltaParaRegistrar = 30 - (frmMenu.alock.UsedDays)

txtdiasquefaltampararegistrar.Text = diasQueFaltaParaRegistrar


If diasQueFaltaParaRegistrar <= 0 Then
   cmdregistrardepois.Enabled = False
End If

txtcodigodoprograma.Text = frmMenu.alock.SoftwareCode
End Sub
