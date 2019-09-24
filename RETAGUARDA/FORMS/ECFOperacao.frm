VERSION 5.00
Begin VB.Form frmECFOperacao 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operações Cupom Fiscal"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   Icon            =   "ECFOperacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdMapaResumo 
      Caption         =   "Ma&pa Resumo"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      MaskColor       =   &H00FFFFC0&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Verifica Ultima Redução Z"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelaCupomCartão 
      Caption         =   "Cancelar Cupom &Venda Cartão Debito"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      MaskColor       =   &H00FFFFC0&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Programa Forma de Recebimento"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdContadoReinicio 
      Caption         =   "Contador Reinício"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   1500
   End
   Begin VB.CommandButton cmdNSerie 
      Caption         =   "Número Serie &Impressora"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdUltimoCupom 
      Caption         =   "Ultimo Cupom Imp&resso"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdFormaPagto 
      Caption         =   "Consulta &Formas Pagamento"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdSangria 
      Caption         =   "San&gria"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdVersao 
      Caption         =   "&Versão do Firmware"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdAliquota 
      Caption         =   "&Programa Alíquota"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelaCupom 
      Caption         =   "Cancelar Cupom Fiscal"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdHoraVerao 
      Caption         =   "Ativa/Desativa &Horário de Verão"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdLeMemoriaData 
      Caption         =   "Leitura &Memoria Fiscal por Período"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      MaskColor       =   &H00FFFFC0&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdReduçãoZ 
      Caption         =   "Redução &Z"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdSAI 
      Caption         =   "&Sair"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      Picture         =   "ECFOperacao.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1500
   End
   Begin VB.CommandButton cmdLeituraX 
      Caption         =   "Leitura &X"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MaskColor       =   &H00FFFFC0&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10077
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   4
      X1              =   8400
      X2              =   8400
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   6735
      X2              =   6735
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   5055
      X2              =   5055
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   3375
      X2              =   3375
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   20
      X2              =   20
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   10080
      X2              =   10080
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   1695
      X2              =   1695
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10077
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10077
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10077
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmECFOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   'Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
   'Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
   'Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
   'Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
   Dim LocalRetorno As String
   LocalRetorno = LeParametrosIni("Sistema", "Retorno")
   If LocalRetorno = "-2" Then
      LocalRetorno = "0" 'devolve o retorno na variavel
      Else: LocalRetorno = Left(LocalRetorno, 1)
   End If

   RETORNO_ECF = Bematech_FI_AbrePortaSerial()
   'Call VerificaRetornoImpressora("", "", "BemaFI32")

   If RETORNO_ECF = -4 Or RETORNO_ECF = -5 Then
      cmdCancelaCupom.Enabled = False
      'cmdLeituraX.Enabled = False
      cmdAliquota.Enabled = False
      cmdSangria.Enabled = False
      'cmdReduçãoZ.Enabled = False
      cmdLeMemoriaData.Enabled = False
      'cmdReseta.Enabled = False
      cmdHoraVerao.Enabled = False
      cmdNSerie.Enabled = False
      cmdUltimoCupom.Enabled = False
      cmdFormaPagto.Enabled = False
      cmdVersao.Enabled = False
      cmdMapaResumo.Enabled = False
      Command3.Enabled = False
      cmdContadoReinicio.Enabled = False
   End If

   If USUARIO_ID_N <> 144 Then
      cmdAliquota.Enabled = False
      'cmdReduçãoZ.Enabled = False
      cmdHoraVerao.Enabled = False
      Else
         cmdCancelaCupom.Enabled = True
         cmdLeituraX.Enabled = True
         cmdAliquota.Enabled = True
         cmdSangria.Enabled = True
         cmdReduçãoZ.Enabled = True
         cmdLeMemoriaData.Enabled = True
'         cmdReseta.Enabled = True
         cmdHoraVerao.Enabled = True
         cmdNSerie.Enabled = True
         cmdUltimoCupom.Enabled = True
         cmdFormaPagto.Enabled = True
         cmdVersao.Enabled = True
         cmdMapaResumo.Enabled = True
         Command3.Enabled = True
         cmdContadoReinicio.Enabled = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA
   
   Select Case KeyCode
      Case vbKeyEscape
         End
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdLeituraX_Click()
'On Error GoTo ERRO_TRATA

   PERGUNTA "Confirma emissão Leitura 'X' ", vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000

   SQL3 = IMPRESSORA_FISCAL_N
   CRITERIO_A = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
   Select Case CRITERIO_A
      Case "BEMATECH"
         If RESPOSTA = vbYes Then
            RETORNO_ECF = Bematech_FI_LeituraX()
            Call VerificaRetornoImpressora("", "", "Leitura X")
         End If
      Case "DARUMA"
         iRetorno = iLeituraX_ECF_Daruma
         'DarumaFramework_Mostrar_Retorno_ECF (iRetorno)
      Case "Sweda"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdReduçãoZ_Click()
'On Error GoTo ERRO_TRATA
   
   'Os parâmetros opcionais são para alterar
   'a hora da impressora em até + ou - 5 min.
   'para isso deve-se passar os parâmetros "Data" e "Hora"
   PERGUNTA "Após emitir redução 'Z' vendas com cupom só serão permitidas no dia seguinte. Confirma emissão Redução 'Z' ", vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
       RETORNO_ECF = Bematech_FI_ReducaoZ("", "")
       Call VerificaRetornoImpressora("", "", "Redução Z")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdLeMemoriaData_Click()
'On Error GoTo ERRO_TRATA

   If Libera_Acesso("cmdLeMemoriaData") Then
      frmECFLeituraMemoriaData.Caption = "Leitura da Memória Fiscal por Data"
      INDR_RECEITA = 1
      frmECFLeituraMemoriaData.Show 1
      INDR_RECEITA = 0
      Else: MsgBox "Acesso não permitido."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdHoraVerao_Click()
'On Error GoTo ERRO_TRATA

   RETORNO_ECF = Bematech_FI_ProgramaHorarioVerao()
   'Função que analisa o retorno da impressora
   Call VerificaRetornoImpressora("", "", "Programação do Horário de Verão")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdAliquota_Click()
'On Error GoTo ERRO_TRATA
'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

   Dim Aliquotas As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
       Aliquotas = Space(1)
      Else: Aliquotas = Space(79)
   End If

   RETORNO_ECF = Bematech_FI_RetornoAliquotas(Aliquotas)
   Call VerificaRetornoImpressora("Alíquotas Cadastradas: ", Aliquotas, "Informações da Impressora")

   If Libera_Acesso("cmdAliquota") Then
      frmECFAliquota.Show 1
      Else: MsgBox "Acesso não permitido."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdVersao_Click()
'On Error GoTo ERRO_TRATA

'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

   Dim VersaoFirmware As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
       VersaoFirmware = Space(1)
   Else
       VersaoFirmware = Space(4)
   End If
   
   RETORNO_ECF = Bematech_FI_VersaoFirmware(VersaoFirmware)
   VersaoFirmware = Mid(VersaoFirmware, 1, 2) & "." & Mid(VersaoFirmware, 3, 2)
   Call VerificaRetornoImpressora("Versão do Firmware: ", VersaoFirmware, "Informações da Impressora")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdSangria_Click()
'On Error GoTo ERRO_TRATA

   If Libera_Acesso("cmdSangria") Then
      frmLojECFSangria.Show 1
      Else: MsgBox "Acesso não permitido."
   End If


Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdFormaPAGTO_Click()
'On Error GoTo ERRO_TRATA

'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

    Dim Formas As String
    Dim FormasAux As String
    Dim LocalRetorno As String
    If (LocalRetorno = "1") Then 'Grava retorno em arquivo
        Formas = Space(1)
    Else
        Formas = Space(3016)
    End If
    
    RETORNO_ECF = Bematech_FI_VerificaFormasPagamento(Formas)
    FormasAux = vbCr & vbLf & vbLf & Formas
    Call VerificaRetornoImpressora("Formas de Pagamento: ", FormasAux, "Informações da Impressora")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdUltimoCupom_Click()
'On Error GoTo ERRO_TRATA

'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

    Dim NUMEROCUPOM As String
    Dim RETORNOSTATUS As String
    Dim LocalRetorno As String
    If (LocalRetorno = "1") Then 'Grava retorno em arquivo
        NUMEROCUPOM = Space(1)
    Else
        NUMEROCUPOM = Space(6)
    End If
    
    RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOM)
    'Função que analisa o retorno da impressora
    Call VerificaRetornoImpressora("Número do Último Cupom: ", _
         NUMEROCUPOM, "Informações da Impressora")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdNSerie_Click()
'On Error GoTo ERRO_TRATA

'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

   Dim NumeroSerie   As String
   Dim LocalRetorno  As String

   If USA_NFC_E = False Then
      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NumeroSerie = Space(1)
         Else: NumeroSerie = Space(20)
      End If

      'RETORNO_ECF = Bematech_FI_NumeroSerie(NumeroSerie)
      RETORNO_ECF = Bematech_FI_NumeroSerieMFD(NumeroSerie)
      Call VerificaRetornoImpressora("Número de Série: ", NumeroSerie, "Informações da Impressora")
      NUMERO_SERIE_ECF = CStr(NumeroSerie)
      Else
         NumeroSerie = Space(4)
         RETORNO_ECF = Bematech_FI_NumeroSerieNFCe(NumeroSerie)
         Call VerificaRetornoImpressora("Número de Série: ", NumeroSerie, "Informações da Impressora")

         MsgBox RETORNO_ECF
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdSAI_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Private Sub cmdMapaResumo_Click()
    Screen.MousePointer = vbHourglass
    RETORNO_ECF = Bematech_FI_MapaResumo()
    If VerificaRetornoImpressora("", "", "Mapa Resumo") Then
        Call ExibeArquivoRetorno
    End If
    Screen.MousePointer = vbNormal
    
    LoadEXE ("d:\windows\notepad.EXE " & "d:\LOJINHA\retorno.txt")
End Sub

Private Sub Command3_Click()
   If Libera_Acesso("frmvendaecf") Then
      frmLojECFProgramaFormaPagamentoMFD.Show 1
      Else: MsgBox "Acesso não permitido."
   End If
End Sub

Private Sub cmdContadoReinicio_Click()

   Dim NumeroIntervencao As String, CONTA_REINICIO As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NumeroIntervencao = Space(1)
      Else: NumeroIntervencao = Space(4)
   End If
   
   RETORNO_ECF = Bematech_FI_NumeroIntervencoes(NumeroIntervencao)
   
   If Trim(NumeroIntervencao) <> "" Then _
      CONTA_REINICIO = NumeroIntervencao
   
   MsgBox CONTA_REINICIO

End Sub
'Ler os Valores dos parâmetros nas seções do arquivo ini
Function LeParametrosIni(Secao As String, Label As String) As String
   Const TamanhoParametro = 80
   Dim ParametroIni As String * TamanhoParametro
   Dim RetornoFuncao
   Dim arquivoIni As String
   Dim Contador As Integer
   ParametroIni = ""
     
   RetornoFuncao = GetSystemDirectory(ParametroIni, TamanhoParametro)
   arquivoIni = Left(ParametroIni, RetornoFuncao) + "\BemaFI32.ini"
   ParametroIni = ""
   RetornoFuncao = GetPrivateProfileString(Secao, Label, "-2", ParametroIni, TamanhoParametro, arquivoIni)
   RetornoFuncao = Mid(ParametroIni, 1, 2)
   If Val(RetornoFuncao) <> -2 Then
       Contador = 1
       Do
           Tst = Mid(ParametroIni, Contador, 1)
           If Asc(Tst) <> 0 Then
               Contador = Contador + 1
           End If
       Loop While ((Asc(Tst) <> 0) And (Contador < Len(ParametroIni)))
       RetornoFuncao = Mid(ParametroIni, 1, Contador)
   End If
   LeParametrosIni = RetornoFuncao
End Function

Private Sub cmdCancelaCupom_Click()
'On Error GoTo ERRO_TRATA

   CANCELA_CUPOM_FISCAL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdCancelaCupom_Click"
End Sub

Private Sub cmdCancelaCupomCartão_Click()
'On Error GoTo ERRO_TRATA
  
   CRITERIO_A = InputBox("Informe número do cupom fiscal venda cartão de débito/credito.", "Cancelamente de Venda com cupom fiscal venda a cartão de debito/credito")
   If Trim(CRITERIO_A) <> "" Then
      If IsNumeric(CRITERIO_A) Then
         NUMR_ID_N = CRITERIO_A

         SQL = "select pedido_id,numr_cupom from CUPOM "
         SQL = SQL & " where numr_cupom = " & NUMR_ID_N
         SQL = SQL & " and CONTA_REINICIO = " & CONTA_REINICIO_N
         Set TabTemp = CONECTA_RETAGUARDA.OpenRecordset(SQL, 4)
         If Not TabTemp.EOF Then

            If Not IsNull(TabTemp.Fields("pedido_id").Value) Then

               PERGUNTA "Confirma cancelamento cupom fiscal número = " & TabTemp.Fields("numr_cupom").Value, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then

                  NUMR_ID_N = TabTemp.Fields("pedido_id").Value
                  TabTemp.Close

                     SQL = "update PEDIDO set "
                     SQL = SQL & " dt_cancela = '" & Now & "'"
                     SQL = SQL & " , status = 'C'"
                     SQL = SQL & " where pedido_id = " & NUMR_ID_N
                     SQL = SQL & " and CONTA_REINICIO = " & CONTA_REINICIO_N
                     SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                  MsgBox "Cupom fiscal nº " & NUMEROCUPOM & " foi cancelado com sucesso."

               End If
               Else: MsgBox "Cupom não encontrado."
            End If
            Else: MsgBox "Cupom não encontrado."
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdCancelaCupomCartão_Click"
End Sub

Private Sub Command5_Click()

   Flag = 0

   RETORNO_ECF = Bematech_FI_VerificaReducaoZAutomatica(Flag)
   'Função que analisa o retorno da impressora
   'Call VerificaRetornoImpressora("Verifica Redução Z: ",  , "Informações da Impressora")

End Sub

Sub CANCELA_CUPOM_FISCAL()
'On Error GoTo ERRO_TRATA

   Dim NUMEROCUPOM As String

   NUMEROCUPOM = 0

   Dim RETORNOSTATUS As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NUMEROCUPOM = Space(1)
      Else: NUMEROCUPOM = Space(6)
   End If

   RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOM)
   'Função que analisa o retorno da impressora
   Call VerificaRetornoImpressora("Número do Último Cupom: ", _
        NUMEROCUPOM, "Informações da Impressora")

   PERGUNTA "Confirma cancelamento cupom fiscal número = " & NUMEROCUPOM, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbNo Then _
      Exit Sub

   Indr_Erro = False

   RETORNO_ECF = Bematech_FI_CancelaCupom()
   'Função que analisa o retorno da impressora
   Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")

   If Indr_Erro = False Then
      NUMR_ID_N = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select pedido_id from CUPOM "
      SQL = SQL & " where numr_cupom = " & NUMEROCUPOM
      SQL = SQL & " and CONTA_REINICIO = " & CONTA_REINICIO_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            NUMR_ID_N = TabTemp.Fields(0).Value

            SQL = "UPDATE PEDIDO set "
            SQL = SQL & " status = 9"
            SQL = SQL & " where pedido_id = " & NUMR_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            MsgBox "Cupom fiscal nº " & NUMEROCUPOM & " foi cancelado com sucesso."
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "CANCELA_CUPOM_FISCAL"
End Sub

Public Function cImpressora_GeraRegistroCAT52(strArquivoMFD As String, strData As String, _
                                             strArquivoGerado As String) As Boolean
   
    'RETORNO_ECF = Bematech_FI_GeraRegistrosCAT52MFDEx(strArquivoMFD, strData, strArquivoGerado)
    'cImpressora_GeraRegistroCAT52 = VerificaRetornoEcf(Retorno)
   
End Function
