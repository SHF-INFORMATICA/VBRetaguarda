VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmECFLeituraMemoriaData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leitura Memoria Fiscal por Data"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtDataInicial 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDataFinal 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   555
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmECFLeituraMemoriaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdOK_Click()
'On Error GoTo ERRO_TRATA

   txtDataInicial.PromptInclude = True
   txtDataFinal.PromptInclude = True
   txtDataFinal.PromptInclude = False
   If Not txtDataFinal.Text = "" Then
      txtDataFinal.PromptInclude = True
      If Not IsDate(txtDataFinal.Text) Then
         MsgBox "Data Informada Inválida !!!"
         txtDataFinal.SetFocus
         Exit Sub
      End If
   End If
   txtDataInicial.PromptInclude = True
   txtDataFinal.PromptInclude = True
   If IsDate(txtDataInicial.Text) And IsDate(txtDataFinal.Text) Then
      If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
         MsgBox "Período Informado Inválido !!!"
         txtDataInicial.SetFocus
         Exit Sub
      End If
   End If

   Screen.MousePointer = vbHourglass
   sinal = 1
   If sinal = 1 Then 'Leitura Memoria Fiscal Data
       Retorno = Bematech_FI_LeituraMemoriaFiscalData(txtDataInicial, txtDataFinal)
   
   ElseIf sinal = 2 Then 'Leitura Memoria Fiscal Serial Data
       Retorno = Bematech_FI_LeituraMemoriaFiscalSerialData(txtDataInicial, txtDataFinal)
       If VerificaRetornoImpressora("", "", "Leitura da Memória Fiscal pela Serial por Data") Then
           Call ExibeArquivoRetorno
       End If
   
   ElseIf sinal = 3 Then 'Leitura de dados para geração do sintegra
       Retorno = Bematech_FI_DadosSintegra(txtDataInicial, txtDataFinal)
       If VerificaRetornoImpressora("", "", "Dados Sintegra") Then
           Call ExibeArquivoRetorno
       End If
   End If
   
   Screen.MousePointer = vbNormal
   
   If Retorno = 1 Then _
      Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdCancelar_Click()
'On Error GoTo ERRO_TRATA

    Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub txtDataFinal_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDataInicial.PromptInclude = True
   txtDataFinal.PromptInclude = True
   txtDataFinal.PromptInclude = False
   If Not IsDate(txtDataInicial.Text) Then
      MsgBox "Data Inicial Inválido !!!"
      txtDataInicial.SetFocus
      Exit Sub
   End If
   txtDataFinal.Text = DateSerial(Year(txtDataInicial.Text), Month(txtDataInicial.Text) + 1, 0)
   txtDataFinal.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub txtDataInicial_GotFocus()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Month(Format(Date, "dd/mm/yyyy"))
   CRITERIO_A = Trim(CRITERIO_A)
   If Len(CRITERIO_A) <= 1 Then _
      CRITERIO_A = "0" & CRITERIO_A
   txtDataInicial.PromptInclude = False
      txtDataInicial.Text = "01/" & CRITERIO_A & "/" & Year(Format(Date, "dd/mm/yyyy"))
   txtDataInicial.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub txtDatafinal_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmdOK.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub txtDataInicial_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDataInicial.PromptInclude = True
   CRITERIO_A = "01" & Right(txtDataInicial.Text, 7)
   txtDataInicial.PromptInclude = False
   txtDataInicial.Text = CRITERIO_A
   txtDataInicial.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub
