VERSION 5.00
Begin VB.Form frmIMAGEM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imagem Produto"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IMAGEM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      Picture         =   "IMAGEM.frx":5C12
      ScaleHeight     =   495
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   4800
      Width           =   600
   End
   Begin VB.TextBox txtPath_foto 
      DataField       =   "Note"
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Top             =   4800
      Width           =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   4800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Image picimagem 
      Height          =   4575
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmIMAGEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   picimagem.Picture = LoadPicture("")

   If Trim(LOCAL_IMAGEM) <> "" Then
      txtPath_foto.Text = LOCAL_IMAGEM
      picimagem.Picture = LoadPicture(txtPath_foto.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Picture1_Click()
'On Error GoTo ERRO_TRATA

   frmINICIO.Dialogo.DialogTitle = "Selecione imagem com código do usuário !!!"
   frmINICIO.Dialogo.Filter = "*.jpg;*.gif;*.bmp;*.ico;*.cur"
   frmINICIO.Dialogo.ShowOpen
   If frmINICIO.Dialogo.FileName <> "" Then
      txtPath_foto.Text = frmINICIO.Dialogo.FileName
      picimagem.Picture = LoadPicture(txtPath_foto.Text)
      LOCAL_IMAGEM = Trim(txtPath_foto.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Picture1_Click"
End Sub
