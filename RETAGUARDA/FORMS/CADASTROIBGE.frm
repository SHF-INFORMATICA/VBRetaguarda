VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmIBGECADASTRO 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro IBGE"
   ClientHeight    =   1920
   ClientLeft      =   2625
   ClientTop       =   2685
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROIBGE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   10005
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
      Height          =   915
      Left            =   210
      TabIndex        =   4
      Top             =   840
      Width           =   9585
      Begin VB.TextBox txtUF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8700
         MaxLength       =   2
         TabIndex        =   2
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox txtCidade 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   1
         Top             =   390
         Width           =   6855
      End
      Begin VB.TextBox txtIbge 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   0
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8700
         TabIndex        =   7
         Top             =   150
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CIDADE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1740
         TabIndex        =   6
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IBGE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   150
         Width           =   525
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1270
      ButtonWidth     =   3466
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consulta IBGE"
            Key             =   "IBGE"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5280
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":9331
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":A43C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":BB39
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROIBGE.frx":CCD3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmIBGECADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         LIMPA_IBGE
         txtIBGE.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         If txtIBGE.Text <> "" Then
            SQL = "select * from IBGE "
            SQL = SQL & " where IBGE = " & txtIBGE.Text
            If TabIBGE.State = 1 Then TabIBGE.Close
            TabIBGE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabIBGE.EOF Then
               MsgBox "Registro não encontrado.", vbOKOnly, "Atenção !!!"
               txtIBGE.SetFocus
               Exit Sub
               Else
                  Msg = "Confirma Exclusão?"
                  Style = vbYesNo + 32
                  Title = "Atenção !!!"
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                     TabIBGE.Delete
                     LIMPA_IBGE
                  End If
            End If
         End If
         txtIBGE.SetFocus
      Case "gravar"
         GRAVA_IBGE
         txtIBGE.SetFocus
      Case "voltar"
        Unload Me
      Case "limpar"
         LIMPA_IBGE
         txtIBGE.SetFocus
      Case "IBGE"
         CRITERIO_A = ""
         frmIBGECONSULTA.Show 1
         txtIBGE.Text = CRITERIO_A
         CRITERIO_A = ""
         txtIBGE.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtIBGE_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtIBGE.Text & "%" & Chr(39)
   SQL = "select * from IBGE WHERE IBGE_ID like " & CRITERIO_A
   
   CRITERIO_A = txtIBGE.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIBGE_Change"
End Sub

Private Sub txtibge_GotFocus()
   txtIBGE.SelStart = 0
   txtIBGE.SelLength = Len(txtIBGE)
   txtIBGE.BackColor = &HC0FFFF
End Sub

Private Sub txtIbge_LostFocus()
   txtIBGE.BackColor = &HFFFFFF
End Sub

Private Sub txtIBGE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtIBGE.Text = "" Then
         MsgBox "Digito Inválido.", vbOKOnly, "ERRO !!!"
         txtIBGE.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      Mostra_IBGE
      txtCidade.SetFocus
      Else
         If KeyAscii = 8 Then
            Else
               If KeyAscii = 45 Then
                  Else
                     If KeyAscii = 32 Then
                        Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
                     End If
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIBGE_KeyPress"
End Sub

Private Sub txtcidade_GotFocus()
   txtCidade.SelStart = 0
   txtCidade.SelLength = Len(txtCidade)
   txtCidade.BackColor = &HC0FFFF
End Sub

Private Sub txtCidade_LostFocus()
   txtCidade.BackColor = &HFFFFFF
End Sub

Private Sub txtcidade_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_KeyPress"
End Sub

Private Sub txtuf_GotFocus()
   txtUF.SelStart = 0
   txtUF.SelLength = Len(txtUF)
   txtUF.BackColor = &HC0FFFF
End Sub

Private Sub txtUF_LostFocus()
   txtUF.Text = UCase(txtUF.Text)
   txtUF.BackColor = &HFFFFFF
End Sub

Private Sub txtuf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_IBGE
      txtIBGE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_KeyPress"
End Sub

Private Sub Mostra_IBGE()
'On Error GoTo ERRO_TRATA

   SQL = "select * from IBGE "
   SQL = SQL & " where IBGE_ID = " & txtIBGE.Text
   If TabIBGE.State = 1 Then TabIBGE.Close
   TabIBGE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabIBGE.EOF Then
      txtIBGE.Text = TabIBGE!IBGE_ID
      If Not IsNull(TabIBGE!Municipio) Then _
         txtCidade.Text = TabIBGE!Municipio
      If Not IsNull(TabIBGE!Estado) Then _
         txtUF.Text = TabIBGE!Estado
   End If
   TabIBGE.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Mostra_IBGE"
End Sub

Private Sub LIMPA_IBGE()
'On Error GoTo ERRO_TRATA

   txtIBGE.Text = ""
   txtCidade.Text = ""
   txtUF.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_IBGE"
End Sub

Sub GRAVA_IBGE()
'On Error GoTo ERRO_TRATA

   If txtIBGE.Text = "" Then
      MsgBox "Informe o IBGE"
      txtIBGE.SetFocus
      Exit Sub
   End If
   If txtCidade.Text = "" Then
      MsgBox "Informe a Cidade"
      txtCidade.SetFocus
      Exit Sub
   End If
   If txtUF.Text = "" Then
      MsgBox "Informe o Estado"
      txtUF.SetFocus
      Exit Sub
   End If

   SQL = "select * from IBGE "
   SQL = SQL & " where IBGE_ID = " & txtIBGE.Text
   If TabIBGE.State = 1 Then TabIBGE.Close
   TabIBGE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabIBGE.EOF Then
      SqL2 = "INSERT INTO IBGE (IBGE_ID, Municipio, Estado) "
      SqL2 = SqL2 & " VALUES (" & txtIBGE.Text & ",'" & txtCidade.Text & "','" & txtUF.Text & "')"
      CONECTA_RETAGUARDA.Execute SqL2
   Else
      CONECTA_RETAGUARDA.Execute "UPDATE IBGE SET Municipio = '" & txtCidade.Text & "', Estado = '" & txtUF.Text & "' where IBGE_ID = " & txtIBGE.Text
   End If
   CRITERIO_A = txtIBGE.Text
   TabIBGE.Close
   LIMPA_IBGE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_IBGE"
End Sub


