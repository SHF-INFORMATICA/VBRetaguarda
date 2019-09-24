VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTRORG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro RG"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTRORG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOrigem 
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
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtRg 
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
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin MSMask.MaskEdBox txtDTEXP 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1270
      ButtonWidth     =   2223
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "excluir"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4440
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTRORG.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTRORG.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTRORG.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTRORG.frx":8F46
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTRORG.frx":A1AE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   7920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   7920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dt.Emisão:"
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
      Left            =   330
      TabIndex        =   5
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label lblOrigem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Orgão Exp.:"
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
      Left            =   330
      TabIndex        =   4
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label lblRg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RG:"
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
      Left            =   1110
      TabIndex        =   3
      Top             =   840
      Width           =   345
   End
End
Attribute VB_Name = "frmCADASTRORG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   If Trim(CNPJCPF_A) <> "" Then _
      MOSTRA_RG

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "gravar"
         txtDTEXP.PromptInclude = True
         If Trim(txtRg.Text) <> "" Then _
            GRAVA_RG Trim(txtRg.Text), Trim(txtOrigem.Text), Trim(txtDTEXP.Text)
      Case "excluir"
         If TabRG.State = 1 Then _
            TabRG.Close

         SQL = "select * from RG "
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         TabRG.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabRG.EOF Then
            SQL = "delete from RG "
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            CONECTA_RETAGUARDA.Execute SQL
            LIMPA_RG
            txtRg.SetFocus
         End If
         If TabRG.State = 1 Then _
            TabRG.Close
      Case "limpar"
         LIMPA_RG
         txtRg.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtRG_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe Nº Registro de Identidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtRG_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      txtOrigem.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtRg_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(CNPJCPF_A) <> "" Then _
      MOSTRA_RG

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRg_LostFocus"
End Sub

Private Sub txtOrigem_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe origem do Registro de Identidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtOrigem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      txtDTEXP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDTEXP_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe data de emissão do Registro de Identidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDTEXP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtDTEXP.PromptInclude = True
      If Trim(txtRg.Text) <> "" Then _
         GRAVA_RG Trim(txtRg.Text), Trim(txtOrigem.Text), Trim(txtDTEXP.Text)

      txtRg.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Sub LIMPA_RG()
'On Error GoTo ERRO_TRATA

   txtRg.Text = ""
   txtOrigem.Text = ""
   txtDTEXP.PromptInclude = False
   txtDTEXP.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_RG"
End Sub

Sub MOSTRA_RG()
'On Error GoTo ERRO_TRATA

   If TabRG.State = 1 Then _
      TabRG.Close

   SQL = "select * from RG "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabRG.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabRG.EOF Then
      txtRg.Text = TabRG!numero_rg
      CRITERIO_A = ""
      If Not IsNull(TabRG!orgao) Then
         For i = 1 To Len(TabRG!orgao)
            If Mid(TabRG!orgao, i, 1) <> "-" And Mid(TabRG!orgao, i, 1) <> "/" Then
               CRITERIO_A = CRITERIO_A & Mid(TabRG!orgao, i, 1)
               Else: Exit For
            End If
         Next

         txtOrigem.Text = CRITERIO_A

         CRITERIO_A = ""
         For i = (Len(txtOrigem.Text) + 2) To Len(TabRG!orgao)
            If Mid(TabRG!orgao, i, 1) <> "-" And Mid(TabRG!orgao, i, 1) <> "/" Then
               CRITERIO_A = CRITERIO_A & Mid(TabRG!orgao, i, 1)
               Else: Exit For
            End If
         Next
      End If

      If Not IsNull(TabRG!Dt_Exp) Then _
         txtDTEXP.Text = Format(TabRG!Dt_Exp, "dd/MM/yyyy")
   End If
   If TabRG.State = 1 Then _
      TabRG.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_RG"
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub
