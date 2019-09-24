VERSION 5.00
Begin VB.Form frmTranslator 
   Caption         =   "Tradutor"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TRADUTOR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      Begin VB.CommandButton Command2 
         Caption         =   "Ouvir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ouvir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   9
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   6255
      End
      Begin VB.CommandButton cmdTranslate 
         Caption         =   "Traduzir"
         Enabled         =   0   'False
         Height          =   345
         Left            =   6585
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtTranslate 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   6255
      End
      Begin VB.ComboBox cmbTo 
         Enabled         =   0   'False
         Height          =   360
         ItemData        =   "TRADUTOR.frx":5C12
         Left            =   4080
         List            =   "TRADUTOR.frx":5CDC
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Selecione"
         Top             =   290
         Width           =   2415
      End
      Begin VB.ComboBox cmbFrom 
         Enabled         =   0   'False
         Height          =   360
         ItemData        =   "TRADUTOR.frx":5F5C
         Left            =   780
         List            =   "TRADUTOR.frx":6026
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Selecione"
         Top             =   290
         Width           =   2415
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "Para:"
         Height          =   240
         Left            =   3480
         TabIndex        =   3
         Top             =   330
         Width           =   510
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "De :"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   330
         Width           =   375
      End
   End
   Begin VB.PictureBox WindowsMediaPlayer1 
      Height          =   3855
      Left            =   8040
      ScaleHeight     =   3795
      ScaleWidth      =   3555
      TabIndex        =   10
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Tradutor"
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
      TabIndex        =   12
      Top             =   0
      Width           =   7980
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4725
      Width           =   7845
   End
End
Attribute VB_Name = "frmTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ieGoogle As Object

Private Sub Command1_Click()
'WindowsMediaPlayer1.url = "http://translate.google.com/translate_tts?ie=UTF-8&tl=" & GetLanguage(cmbTo) & "&q=" & ieGoogle.document.getElementById("result_box").innertext
End Sub

Private Sub Command2_Click()
'WindowsMediaPlayer1.url = "http://translate.google.com/translate_tts?ie=UTF-8&tl=" & GetLanguage(cmbFrom) & "&q=" & txtTranslate.Text
End Sub

Private Sub cmdTranslate_Click()
On Error Resume Next
  Command1.Enabled = True
  Command2.Enabled = True
  txtResult.Enabled = True
  txtResult.Text = ieGoogle.Document.getElementById("result_box").innertext
  ChangeStatus "Carregando Fala/Som..."
  'WindowsMediaPlayer1.url = "http://translate.google.com/translate_tts?ie=UTF-8&tl=" & GetLanguage(cmbFrom) & "&q="
  ChangeStatus "Pronto"
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
  Me.Show
  ChangeStatus "Carregando o Programa , Aguarde..."
  
  Set ieGoogle = CreateObject("InternetExplorer.Application")
  
  With ieGoogle
    .Silent = True
    .Navigate "http://translate.google.com/"
    .Visible = False 'Put this to True if you want to see Internet Explorer Window
  End With

  If ieReady Then
    ChangeStatus "Pronto"
    cmbTo.Enabled = True
    cmbFrom.Enabled = True
    cmdTranslate.Enabled = True
    txtTranslate.Enabled = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ieGoogle.Quit
ieGoogle = Nothing
End Sub

Private Sub cmbFrom_Click()
  Command1.Enabled = False
  Command2.Enabled = False
  txtResult.Enabled = False
  txtTranslate.Text = ""
  ieGoogle.Document.getElementById("gt-submit").Click
  ieGoogle.Document.getElementById("gt-sl").Value = GetLanguage(cmbFrom)
  ieGoogle.Document.getElementById("gt-submit").Click
  ieGoogle.Document.Forms(0).submit
  'WindowsMediaPlayer1.URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=" & GetLanguage(cmbFrom) & "&q="
End Sub

Private Sub cmbTo_Click()
  Command1.Enabled = False
  Command2.Enabled = False
  txtResult.Enabled = False
  txtTranslate.Text = ""
  ieGoogle.Document.getElementById("gt-submit").Click
  ieGoogle.Document.getElementById("gt-tl").Value = GetLanguage(cmbTo)
  ieGoogle.Document.getElementById("gt-submit").Click
  ieGoogle.Document.Forms(0).submit
  'WindowsMediaPlayer1.URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=" & GetLanguage(cmbTo) & "&q="
End Sub

Private Sub txtTranslate_Change()
On Error Resume Next
  If cmbTo = "-Select-" Or cmbFrom = "-Select-" Then MsgBox "Please select langage to translate from/to.", vbInformation, "Select Language"
  Command1.Enabled = False
  Command2.Enabled = False
  txtResult.Enabled = False
  txtResult.Text = ""
  ieGoogle.Document.getElementById("gt-src-wrap").childNodes.Item(1).childNodes.Item(1).Value = txtTranslate.Text
End Sub

Private Function ieReady() As Boolean
Dim ie_Ready As Long
Dim doc_Ready As String

  ie_Ready = 4
  doc_Ready = "complete"
 
  Do Until ieGoogle.readyState = ie_Ready
    DoEvents
  Loop
  
  Do Until ieGoogle.Document.readyState = doc_Ready
    DoEvents
  Loop
  
  ieReady = True
  Exit Function
  
ErrExit:
  ieReady = False
End Function

Private Sub ChangeStatus(sStatus As String)
lblStatus.Caption = sStatus
End Sub

Private Function GetLanguage(sLanguage As String) As String
Select Case sLanguage
  Case "Afrikaans"
    GetLanguage = "af"
  Case "Albanian"
    GetLanguage = "sq"
  Case "Arabic"
    GetLanguage = "ar"
  Case "Armenian"
    GetLanguage = "hy"
  Case "Azerbaijani"
    GetLanguage = "zy"
  Case "Basque"
    GetLanguage = "eu"
  Case "Belarusian"
    GetLanguage = "be"
  Case "Bengali"
    GetLanguage = "bn"
  Case "Bulgarian"
    GetLanguage = "bg"
  Case "Catalan"
    GetLanguage = "ca"
  Case "Chinese(Simplified)"
    GetLanguage = "zh-CN"
  Case "Chinese(Traditional)"
    GetLanguage = "zh-TW"
  Case "Croatian"
    GetLanguage = "hr"
  Case "Czech"
    GetLanguage = "cs"
  Case "Danish"
    GetLanguage = "da"
  Case "Dutch"
    GetLanguage = "nl"
  Case "English"
    GetLanguage = "en"
  Case "Esperanto"
    GetLanguage = "eo"
  Case "Estonian"
    GetLanguage = "et"
  Case "Filipino"
    GetLanguage = "tl"
  Case "Finnish"
    GetLanguage "fi"
  Case "French"
    GetLanguage = "fr"
  Case "Galician"
    GetLanguage = "gl"
  Case "Georgian"
    GetLanguage = "ka"
  Case "German"
    GetLanguage = "de"
  Case "Greek"
    GetLanguage = "el"
  Case "Gujarati"
    GetLanguage = "gu"
  Case "Haitian Creole"
    GetLanguage = "ht"
  Case "Hebrew"
    GetLanguage = "iw"
  Case "Hindi"
    GetLanguage = "hi"
  Case "Hungarian"
    GetLanguage = "hu"
  Case "Icelandic"
    GetLanguage = "is"
  Case "Indonesian"
    GetLanguage = "id"
  Case "Irish"
    GetLanguage = "ga"
  Case "Italian"
    GetLanguage = "it"
  Case "Japanese"
    GetLanguage = "ja"
  Case "Kannada"
    GetLanguage = "kn"
  Case "Korean"
    GetLanguage = "ko"
  Case "Lao"
    GetLanguage = "lo"
  Case "Latin"
    GetLanguage = "la"
  Case "Latvian"
    GetLanguage = "lv"
  Case "Lithuanian"
    GetLanguage = "lt"
  Case "Macedonian"
    GetLanguage = "mk"
  Case "Malay"
    GetLanguage = "ms"
  Case "Maltese"
    GetLanguage = "mt"
  Case "Norwegian"
    GetLanguage = "no"
  Case "Persian"
    GetLanguage = "fa"
  Case "Polish"
    GetLanguage = "pl"
  Case "Portuguese"
    GetLanguage = "pt"
  Case "Romanian"
    GetLanguage = "ro"
  Case "Russian"
    GetLanguage = "ru"
  Case "Serbian"
    GetLanguage = "sr"
  Case "Slovak"
    GetLanguage = "sk"
  Case "Slovenian"
    GetLanguage = "sl"
  Case "Spanish"
    GetLanguage = "es"
  Case "Swahili"
    GetLanguage = "sw"
  Case "Swedish"
    GetLanguage = "sv"
  Case "Tamil"
    GetLanguage = "ta"
  Case "Telugu"
    GetLanguage = "te"
  Case "Thai"
    GetLanguage = "th"
  Case "Turkish"
    GetLanguage = "tr"
  Case "Ukrainian"
    GetLanguage = "uk"
  Case "Urdu"
    GetLanguage = "ur"
  Case "Vietnamese"
    GetLanguage = "vi"
  Case "Welsh"
    GetLanguage = "cy"
  Case "Yiddish"
    GetLanguage = "yi"
End Select
End Function

