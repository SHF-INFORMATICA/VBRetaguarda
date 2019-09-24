VERSION 5.00
Object = "{EDF439C0-99E5-11CF-AFF3-004005100200}#8.0#0"; "PVMarq.ocx"
Begin VB.Form frmRegistro 
   BackColor       =   &H00000000&
   Caption         =   "Formulário de Registro"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   Icon            =   "RegistroEJS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdregistrardepois 
      Caption         =   "&Registrar Depois"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdregistraragora 
      Caption         =   "Registrar &Agora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtCodigoLiberacao 
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
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtCodigoDoPrograma 
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtDiasQueFaltamParaRegistrar 
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
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "30"
      Top             =   720
      Width           =   855
   End
   Begin PVMarqueeLib.PVMarquee PVMarquee1 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6135
      _Version        =   524288
      _ExtentX        =   10821
      _ExtentY        =   873
      _StockProps     =   29
      Text            =   "Informar código de liberação"
      ForeColor       =   16711680
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Frame           =   5
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Text            =   "Informar código de liberação"
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   2400
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
      TabIndex        =   5
      Top             =   1920
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
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
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
      TabIndex        =   2
      Top             =   720
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
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmregistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim diasQueFaltaParaRegistrar As Integer

   diasQueFaltaParaRegistrar = 0

   diasQueFaltaParaRegistrar = 30 - (frmINICIO.ActiveLock1.UsedDays)

   txtDiasQueFaltamParaRegistrar.Text = diasQueFaltaParaRegistrar

   If diasQueFaltaParaRegistrar <= 0 Then _
      cmdregistrardepois.Enabled = False

   txtCodigoDoPrograma.Text = frmINICIO.ActiveLock1.SoftwareCode

   Me.Caption = Me.Caption & " " & frmINICIO.ActiveLock1.LastRunDate
End Sub

Private Sub cmdregistraragora_Click()

   If txtCodigoLiberacao.Text = "" Then
      txtCodigoLiberacao.SetFocus
      Exit Sub
   End If

   frmINICIO.ActiveLock1.LiberationKey = txtCodigoLiberacao.Text

   If Not frmINICIO.ActiveLock1.RegisteredUser Then
      MsgBox "Chave de LIBERAÇÃO INCORRETA", vbOKOnly + vbCritical, "Chave Liberação Incorreta"
      txtCodigoLiberacao.SetFocus
      Else
         MsgBox "REGISTRO EFETUADO COM SUCESSO !", vbExclamation, "Registro OK"
         frmINICIO.lblAviso.Visible = False
         frmINICIO.Caption = "VERSÃO REGISTRADA"
         'Set frmINICIO.ActiveLock1.LastRunDate = Now
         Unload Me
   End If
End Sub

Private Sub cmdregistrardepois_Click()
   End
End Sub

Private Sub cmdSair_Click()
   End
End Sub
