VERSION 5.00
Begin VB.Form frmTROCASENHA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trocar Senha"
   ClientHeight    =   1275
   ClientLeft      =   5430
   ClientTop       =   3570
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TROCASENHA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3930
   Begin VB.TextBox txtNovaSenha 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   690
      Width           =   2175
   End
   Begin VB.TextBox txtSenhaAtual 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "@"
      TabIndex        =   0
      Top             =   210
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nova Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   300
      TabIndex        =   3
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha Atual:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   270
      Width           =   1230
   End
End
Attribute VB_Name = "frmTROCASENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtSenhaAtual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNovaSenha.SetFocus
   End If
End Sub

Private Sub txtnovasenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txtNovaSenha.Text) <> "" Then
         If Trim(txtSenhaAtual.Text) <> "" Then
            SQL = "select * from USUARIO "
            SQL = SQL & " where usuario_id = " & USUARIO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Trim(TabTemp.Fields("senha").Value) <> Trim(txtSenhaAtual.Text) Then
                  TabTemp.Close
                  MsgBox "Senha atual não confere."
                  Exit Sub
               End If

               CONECTA_RETAGUARDA.Execute "UPDATE USUARIO SET Senha = '" & Trim(txtNovaSenha.Text) & "' where usuario_id = " & USUARIO_ID_N

               txtNovaSenha.Text = ""
               txtSenhaAtual.Text = ""
               MsgBox "Senha alterada com sucesso."
               Unload Me
               Else
                  TabTemp.Close
                  MsgBox "Erro."
                  Exit Sub
            End If
            TabTemp.Close
         End If
      End If
   End If
End Sub
