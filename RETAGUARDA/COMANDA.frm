VERSION 5.00
Begin VB.Form frmCOMANDA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido Barra"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "COMANDA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtResp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtComanda 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Registro Comanda Eletrônica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Index           =   1
      Left            =   -480
      TabIndex        =   5
      Top             =   0
      Width           =   7500
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Atendente:"
      Height          =   285
      Left            =   315
      TabIndex        =   4
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Digite ou leia código de Barras Comanda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   6315
   End
End
Attribute VB_Name = "frmCOMANDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Consiste em dois momentos:
'1-BALCAO : Onde é registrado os produtos pelo atendente
'2-CAIXA :  Onde é efetuado a leitura da comanda e início do processo de composição do pedido venda

Private Sub Form_Load()

   LIMPA_TEMP
   txtAt.Text = "" & USUARIO_ID_N
   txtResp.Text = "" & TRAZ_NOME_USUARIO(txtAt.Text)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
   End Select

End Sub

Private Sub TXTCOMANDA_GotFocus()
   txtComanda.SelStart = 0
   txtComanda.SelLength = Len(txtComanda.Text)
   txtComanda.BackColor = &HC0FFFF
End Sub

Private Sub TXTCOMANDA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

   If KeyAscii = 13 Then
      KeyAscii = 0
      TRATA_COMANDA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtat_GotFocus()
   txtAt.SelStart = 0
   txtAt.SelLength = Len(txtAt.Text)
   txtAt.BackColor = &HC0FFFF
End Sub

Private Sub TXTAT_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

   If KeyAscii = 13 Then
      KeyAscii = 0
      TRATA_ATENDENTE
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Sub LIMPA_TEMP()
'On Error GoTo ERRO_TRATA

   txtComanda.Text = ""
   txtResp.Text = "" & USU_LOGADO
   CARTAOBARRA_ID_N = 0
   ATENDENTE_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TEMP"
End Sub

Sub TRATA_ATENDENTE()
'On Error GoTo ERRO_TRATA

      ATENDENTE_ID_N = 0

      If Trim(txtAt.Text) = "" Then
         MsgBox "Atendente informado inválido."
         Exit Sub
      End If
      If Not IsNumeric(txtAt.Text) Then
         MsgBox "Atendente informado inválido."
         Exit Sub
      End If

ATENDENTE_ID_N = 0 & txtAt.Text

      If TabUSU.State = 1 Then _
         TabUSU.Close

      SQL = "select usuario_id,logon,tipo,status,classe,pessoa_id,funcionario from USUARIO WITH (NOLOCK)"
      SQL = SQL & " where usuario_id = " & ATENDENTE_ID_N
      TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabUSU.EOF Then
         If TabUSU.State = 1 Then _
            TabUSU.Close

         ATENDENTE_ID_N = 0
         MsgBox "Atendente informado inválido."
         Exit Sub
      End If
      If TabUSU.State = 1 Then _
         TabUSU.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_ATENDENTE"
End Sub

Sub TRATA_COMANDA()
'On Error GoTo ERRO_TRATA

   If Trim(txtComanda.Text) = "" Then
      MsgBox "Informe comanda válida !!!"
      Exit Sub
   End If
   If Not IsNumeric(txtComanda.Text) Then
      MsgBox "Informe comanda válida !!!"
      Exit Sub
   End If

   Dim TabComanda As New ADODB.Recordset

   CARTAOBARRA_ID_N = 0 & Trim(txtComanda.Text)

   If TabComanda.State = 1 Then _
      TabComanda.Close

   SQL = "select * from CARTAOBARRA WITH (NOLOCK)"
   SQL = SQL & " where CARTAOBARRA_id = " & CARTAOBARRA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabComanda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabComanda.EOF Then
      If TabComanda.State = 1 Then _
         TabComanda.Close

      CARTAOBARRA_ID_N = 0
      MsgBox "Comanda inválida !!!"
      Exit Sub
   End If
   If TabComanda.State = 1 Then _
      TabComanda.Close

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_COMANDA"
End Sub
