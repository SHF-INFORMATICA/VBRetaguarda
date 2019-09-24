VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmNFeContador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar XML Contador"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NFeContador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "e-mail"
      TabPicture(0)   =   "NFeContador.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmbPeriodo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.ComboBox cmbPeriodo 
         Height          =   390
         Left            =   2880
         TabIndex        =   0
         Text            =   "Selecione mês/ano"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Height          =   3855
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   7815
         Begin VB.TextBox txtAnexo 
            Height          =   390
            Left            =   1800
            TabIndex        =   5
            Top             =   3360
            Width           =   5895
         End
         Begin VB.TextBox txtRemetente 
            Height          =   390
            Left            =   1800
            TabIndex        =   1
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   390
            Left            =   1800
            TabIndex        =   2
            Top             =   720
            Width           =   5895
         End
         Begin VB.TextBox txtAssunto 
            Height          =   390
            Left            =   1800
            TabIndex        =   3
            Top             =   1200
            Width           =   5895
         End
         Begin VB.TextBox txtTexto 
            Height          =   1590
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1680
            Width           =   5895
         End
         Begin Threed.SSCommand cmdEnviar 
            Height          =   975
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1720
            _Version        =   262144
            CaptionStyle    =   1
            ForeColor       =   255
            PictureFrames   =   1
            Picture         =   "NFeContador.frx":5C2E
            Caption         =   "&Enviar"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
            Height          =   270
            Index           =   4
            Left            =   720
            TabIndex        =   14
            Top             =   3360
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Remetente:"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Destinatário:"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
            Height          =   270
            Index           =   2
            Left            =   660
            TabIndex        =   10
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Destinatário:"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   1455
         End
      End
   End
   Begin MSMAPI.MAPISession MAPIConecta 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIEnvia 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Enviar XML Contador"
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
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8220
   End
End
Attribute VB_Name = "frmNFeContador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()

   txtRemetente.Text = "shfhoracio@gmail.com"
   txtDestinatario.Text = "shfhoracio@gmail.com"
   txtAssunto.Text = "E ai teste email"
   txtTexto.Text = "VACA VEIA 02"
   txtAnexo.Text = "c:\201311.RAR"

   If Trim(txtDestinatario.Text) = "" Then
      MsgBox "Destinatário inválido."
      Exit Sub
   End If

   
   EnviarEmail MAPIEnvia, MAPIConecta
   'Enviar_Email
End Sub

Sub Enviar_Email()
'On Error GoTo ERRO_TRATA

'enviando email
   'MAPIConecta.UserName = ""
   'MAPIConecta.UserName = "shfhoracio@gmail.com"
   'MAPIConecta.Password = "filhodaputavaca"

   'MAPIConecta.LogonUI = True
   MAPIConecta.DownLoadMail = False
   '---------------------------------------------------
   'Sign on Sessão
   '---------------------------------------------------
   If (MAPIEnvia.SessionID = 0) Then
      Call MAPIConecta.SignOn
   End If

   MAPIEnvia.SessionID = MAPIConecta.SessionID
   MAPIEnvia.AddressResolveUI = False

   MAPIEnvia.Compose

      MAPIEnvia.RecipIndex = 0
      MAPIEnvia.RecipType = 1
      MAPIEnvia.RecipAddress = "smtp: " & txtDestinatario.Text
      MAPIEnvia.MsgSubject = txtAssunto.Text
      MAPIEnvia.MsgNoteText = txtTexto.Text

      'anexa no final da mensagem
         MAPIEnvia.AttachmentPosition = Len(MAPIEnvia.MsgNoteText)

      'define o tipo de dados do anexo
         MAPIEnvia.AttachmentType = mapData

      'da um nome ao anexo
         MAPIEnvia.AttachmentName = "Anexos"

      'define o caminho e nome do arquivo a anexar
         MAPIEnvia.AttachmentPathName = txtAnexo.Text

   MAPIEnvia.send False

   MAPIConecta.SignOff

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Enviar_Email"
End Sub

Public Function EnviarEmail(ByRef MAPIMessages As MAPIMessages, ByRef MAPISession As MAPISession)
   MAPISession.DownLoadMail = False
   '---------------------------------------------------
   'Sign on Sessão
   '---------------------------------------------------
   If (MAPISession.SessionID = 0) Then
      Call MAPISession.SignOn
   End If

   With MAPIMessages
      .SessionID = MAPISession.SessionID
      .Compose
      .AddressResolveUI = False

      .MsgSubject = txtAssunto.Text
      .MsgNoteText = txtTexto.Text

      .RecipIndex = 0
      .RecipType = 1
      .RecipAddress = "smtp: " & txtDestinatario.Text
      .RecipDisplayName = "smtp: " & txtDestinatario.Text


      'anexa no final da mensagem
         .AttachmentPosition = Len(.MsgNoteText)

      'define o tipo de dados do anexo
         .AttachmentType = mapData

      'da um nome ao anexo
         .AttachmentName = "Anexos"

      'define o caminho e nome do arquivo a anexar
         .AttachmentPathName = txtAnexo.Text


      .send True
   End With
End Function
