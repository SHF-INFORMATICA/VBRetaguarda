VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNFG 
   Caption         =   "Gera��o Arquivo TXT Nota Fiscal Goi�na"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NFG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNSerie 
      Caption         =   "N�mero Serie &Impressora"
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
      Height          =   615
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1500
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   315
      Left            =   3300
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   315
      Left            =   6030
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   2619
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar TXT"
            Key             =   "gerar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5040
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFG.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFG.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFG.frx":7E3B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
      Height          =   285
      Left            =   4620
      TabIndex        =   3
      Top             =   2670
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   2670
      Width           =   1335
   End
End
Attribute VB_Name = "frmNFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call cmdNSerie_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gerar"
         GERAR_NFG
      Case "voltar"
         Unload Me
      Case "limpar"
         'LIMPA_TELA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtDTfim_GotFocus()
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDtFim_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
         txtDtFim.Text = Date
      txtDtFim.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtfim_LostFocus"
End Sub

Private Sub txtDTINI_GotFocus()
   txtDtIni.PromptInclude = True
End Sub

Private Sub txtDtIni_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   If Not IsDate(txtDtIni.Text) Then
      txtDtIni.PromptInclude = False
         txtDtIni.Text = Date
      txtDtIni.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtIni_LostFocus"
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
   
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub
'========================================================================
Private Sub cmdNSerie_Click()
'On Error GoTo ERRO_TRATA

   Dim NumeroSerie   As String
   Dim LocalRetorno  As String

   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NumeroSerie = Space(1)
      Else: NumeroSerie = Space(20)
   End If

   RETORNO_ECF = Bematech_FI_NumeroSerieMFD(NumeroSerie)
   Call VerificaRetornoImpressora("N�mero de S�rie: ", NumeroSerie, "Informa��es da Impressora")
   NUMERO_SERIE_ECF = CStr(NumeroSerie)

   lblnumeroserie.Caption = ""
   If Trim(NUMERO_SERIE_ECF) = "" Then _
      lblnumeroserie.Caption = Trim(NUMERO_SERIE_ECF)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdNSerie"
End Sub

Sub GERAR_NFG()
   'Nova reda��o dada ao item 4 pelo Ato COTEPE/ICMS 08/07, efeitos a partir de 29/06/07.

   '4 - ESTRUTURA DO ARQUIVO:
   '4.1 - O arquivo comp�e-se dos seguintes tipos de registros:
   '4.1.1 - Registro tipo E01 � Identifica��o do ECF;
   '4.1.2 - Registro tipo E02 � Identifica��o do atual contribuinte usu�rio do ECF;
   '4.1.3 - Registro tipo E03 � Identifica��o dos prestadores de servi�o cadastrados no ECF;
   '4.1.4 - Registro tipo E04 � Rela��o dos usu�rios anteriores do ECF;
   '4.1.5 - Registro tipo E05 � Rela��o das codifica��es de GT;
   '4.1.6 - Registro tipo E06 � Rela��o dos s�mbolos da moeda;
   '4.1.7 - Registro tipo E07 � Rela��o das altera��es de vers�o do Software B�sico do ECF;
   '4.1.8 - Registro tipo E08 � Rela��o dos dispositivos de MFD utilizados;
   '4.1.9 - Registro tipo E09 � Rela��o de interven��es t�cnicas;
   '4.1.10 - Registro tipo E10 � Rela��o de Fitas-detalhe emitidas;
   '4.1.11 - Registro tipo E11 � Posi��o atual dos contadores e totalizadores;
   '4.1.12 - Registro tipo E12 � Rela��o de Redu��es Z;
   '4.1.13 - Registro tipo E13 � Detalhe da Redu��o Z � Totalizadores Parciais;
   '4.1.14 - Registro tipo E14 � Cupom Fiscal, Nota Fiscal de Venda a Consumidor ou Bilhete de Passagem;
   '4.1.15 - Registro tipo E15 � Detalhe do Cupom Fiscal, da Nota Fiscal de Venda a Consumidor ou do Bilhete de Passagem;
   '4.1.16 � Registro tipo E16 � Demais documentos emitidos pelo ECF;
   '4.1.17 � Registro tipo E17 � Detalhe da Redu��o Z � Totalizadores N�o Fiscais;
   '4.1.18 � Registro tipo E18 � Detalhe da Redu��o Z � Meios de Pagamento e Troco;
   '4.1.10 � Registro tipo E19 � Documento N�o Fiscal;
   '4.1.20 � Registro tipo E20 � Detalhe do Documento N�o Fiscal;
   '4.1.21 � Registro tipo E21 � Detalhe do Cupom Fiscal e do Documento N�o Fiscal � Meio de Pagamento.

'5 - GERA��O DO ARQUIVO:
'5.1 - O arquivo eletr�nico dever� ser gerado e assinado digitalmente por DLL (Dynamic Link Library) que atenda �s especifica��es estabelecidas no Anexo I deste ato, desenvolvida pelo fabricante do ECF para funcionamento com o programa aplicativo eECFc vers�o 3.03 ou posterior, disponibilizado pelo Fisco e que execute as seguintes fun��es de acordo com o comando previsto na tela de interface de usu�rio do programa eECFc, ap�s selecionado o fabricante e o modelo de ECF e a origem dos dados (Porta Serial ou Arquivo Bin�rio):
'5.1.1 - Comando �Gerar Arquivo Bin�rio�:
'5.1.1.1 - Bot�o de Sele��o: �MF - Leit. Dados da Mem�ria Fiscal�:
'5.1.1.1.1 - deve gerar arquivo bin�rio contendo todas as informa��es do per�odo solicitado gravadas na Mem�ria Fiscal e grav�-lo com o nome �xxxxxx_aaaammdd_hhmmss.MF�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos Bin�rios� existente no diret�rio do respectivo fabricante do ECF;
'5.1.1.2 - Bot�o de Sele��o: �MFD - Leit. Dados da Mem�ria Fita-Detalhe�:
'5.1.1.2.1 - deve gerar arquivo bin�rio contendo todas as informa��es do per�odo solicitado gravadas na Mem�ria de Fita Detalhe e grav�-lo com o nome �xxxxxx_aaaammdd_hhmmss.MFD�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos Bin�rios� existente no diret�rio do respectivo fabricante do ECF;
'5.1.1.3 - Bot�o de Sele��o: �TDM - Leit. Dados das Mem�rias do ECF�:
'5.1.1.3.1 - deve gerar dois arquivos bin�rios em conformidade com o previsto nos itens 5.1.1.1.1 e 5.1.1.2.1.
'5.1.2 - Comando �Gerar Arquivo Texto�:
'5.1.2.1 - Bot�o de Sele��o: �MF - Leit. Dados da Mem�ria Fiscal�:
'5.1.2.1.1 - deve abrir um arquivo bin�rio, previamente selecionado pelo usu�rio, com extens�o �.MF� gerado conforme disposto no item 5.1.1.1.1;
'5.1.2.1.2 - deve gerar arquivo texto conforme os itens 6 e 7 deste ato contendo os seguintes tipos de registro: E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12, E13, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato e grav�-lo com o nome �MFxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Formatados� existente no diret�rio do respectivo fabricante do ECF;
'5.1.2.2 - Bot�o de Sele��o: �MFD - Leit. Dados da Mem�ria Fita-Detalhe�:
'5.1.2.2.1 - deve abrir um arquivo bin�rio, previamente selecionado pelo usu�rio, com extens�o �.MFD� gerado conforme disposto no item 5.1.1.2.1;
'5.1.2.2.2 - deve gerar arquivo texto conforme os itens 6 e 7 deste ato contendo os seguintes tipos de registro: E01, E02, E14, E15, E16, E17, E18, E19, E20 e E21, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato e grav�-lo com o nome �MFDxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Formatados� existente no diret�rio do respectivo fabricante do ECF;
'5.1.2.3 - Bot�o de Sele��o: �TDM - Leit. Dados das Mem�rias do ECF�:
'5.1.2.3.1 - deve abrir dois arquivos bin�rios, previamente selecionados pelo usu�rio, com extens�o �MF� e �.MFD� gerados conforme disposto no item 5.1.1.3.1;
'5.1.2.3.2 - deve gerar arquivo texto conforme os itens 6 e 7 deste ato contendo os seguintes tipos de registro: E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12, E13, E14, E15, E16, E17, E18, E19, E20 e E21, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato e grav�-lo com o nome �TDMxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Formatados� existente no diret�rio do respectivo fabricante do ECF;
'5.1.2.4 - Bot�o de Sele��o: �RZ - Recup. Dados da Redu��o Z�:
'5.1.2.4.1 - deve abrir um arquivo bin�rio, previamente selecionado pelo usu�rio, com extens�o �.RZ� gerado conforme disposto no item 5.1.3.1 deste ato;
'5.1.2.4.2 - deve gerar arquivo texto conforme os itens 6 e 7 deste ato contendo os seguintes tipos de registro: E01, E02, E14, E15 e E16, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato e grav�-lo com o nome �RZxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Formatados� existente no diret�rio do respectivo fabricante do ECF;
'5.1.3 - Comando �Ler Bitmap RZ�:
'5.1.3.1 - deve gerar arquivo bin�rio contendo todas as informa��es representadas nos arquivos de imagem do BitMap e grav�-lo com o nome �xxxxxx_aaaammdd_hhmmss.RZ�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos Bin�rios� existente no diret�rio do respectivo fabricante do ECF;
'5.1.3.2 - deve gerar arquivo texto conforme os itens 6 e 7 deste ato contendo os seguintes tipos de registro: E01, E02, E14, E15 e E16, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato e grav�-lo com o nome �RZxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Formatados� existente no diret�rio do respectivo fabricante do ECF;
'5.1.4 - Comando �Gerar Espelho da LMF�:
'5.1.4.1 - deve abrir um arquivo bin�rio, previamente selecionado pelo usu�rio, com extens�o �.BIN� gerado conforme disposto no item 5.1.7;
'5.1.4.2 - deve possibilitar a sele��o da Leitura Simplificada ou Completa e o per�odo por data ou intervalos de CRZ;
'5.1.4.3 - deve gerar arquivo texto contendo a Leitura da Mem�ria Fiscal em formato de espelho do documento e grav�-lo com o nome �EMFxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Espelho� existente no diret�rio do respectivo fabricante do ECF;
'5.1.5 - Comando �Gerar Espelho da MFD�:
'5.1.5.1 - deve abrir um arquivo bin�rio, previamente selecionado pelo usu�rio, com extens�o �.MFD� gerado conforme disposto no item 5.1.1.2.1;
'5.1.5.2 - deve possibilitar a sele��o do per�odo por data ou intervalos de COO ou a impress�o total;
'5.1.5.3 - deve gerar arquivo texto contendo a Leitura da Mem�ria de Fita Detalhe em formato de espelho do documento e grav�-lo com o nome �EMFDxxxxxx_aaaammdd_hhmmss.TXT�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos TXT Espelho� existente no diret�rio do respectivo fabricante do ECF;
'5.1.6 - Comando �Leitura do Software B�sico�: deve gerar arquivo no formato bin�rio correspondente ao conte�do gravado no dispositivo de armazenamento do Software B�sico do ECF e grav�-lo com o nome �SBxxxxxx_aaaammdd_hhmmss.BIN�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos SB� existente no diret�rio do respectivo fabricante do ECF;
'5.1.7 - Comando �Leitura do Bin�rio da Mem�ria Fiscal�: deve gerar arquivo no formato bin�rio correspondente ao conte�do gravado no dispositivo de armazenamento da Mem�ria Fiscal do ECF e grav�-lo com o nome �MFxxxxxx_aaaammdd_hhmmss.BIN�, onde �xxxxxx� representa o n�mero de fabrica��o do ECF, �aaaammdd� representa a data de gera��o do arquivo e �hhmmss� representa o hor�rio de gera��o do arquivo, na pasta �Arquivos MF� existente no diret�rio do respectivo fabricante do ECF;
'5.1.8 - Comando �Leitura X�: deve enviar ao ECF comando para impress�o da Leitura X;
'5.1.9 - Comando �Leitura da Mem�ria Fiscal�: deve enviar ao ECF comando para impress�o da Leitura da Mem�ria Fiscal possibilitando selecionar Leitura Simplificada ou Completa e per�odo por data ou intervalos de CRZ;
'5.1.10 - Comando �Impress�o da Fita-Detalhe�: deve enviar ao ECF comando para impress�o da Fita Detalhe possibilitando selecionar per�odo por data ou intervalos de COO ou a impress�o total;
'5.2 - Quando n�o houver informa��o relativa ao tipo de registro que deve ser gerado dever� ser gerado apenas um registro do respectivo tipo devendo:
'5.2.1 - conter a informa��o dos quatro primeiros campos do registro, de modo a identificar o ECF;
'Nova reda��o dada ao item 5.2.2 pelo Ato COTEPE/ICMS 26/10, efeitos a partir de 01/11/10.
'5.2.2 - observar o disposto nos itens 3.1, 3.2, 3.3, 3.4 e 3.5 para os demais campos do registro;
'Reda��o anterior, efeitos de 16/04/08 a 31/10/10.
'5.2.2 - observar o disposto nos itens 3.1 e 3.2 para os demais campos do registro;

'Reda��o anterior dada ao item 5 pelo Ato COTEPE/ICMS 08/07, efeitos de 29/06/07 a 15/04/08.

'5 � GERA��O DO ARQUIVO:
'5.1 � O arquivo dever� ser gerado pela DLL (Dynamic Link Library) desenvolvida pelo fabricante do ECF que contenha as seguintes funcionalidades, devendo cada fun��o possuir comando �nico e exclusivo, para interface do fisco:
'5.1.1 � Leitura dos dados gravados na Mem�ria Fiscal, em conformidade com o disposto na cl�usula oitava do Conv�nio ICMS 85/01, de 28 de setembro de 2001 ou no � 2� da cl�usula vig�sima terceira do Conv�nio ICMS 156/94, de 7 de dezembro de 1994, conforme o caso, e no  item 20.1 da al�nea �b�, do  inciso III , da cl�usula s�tima,  do Protocolo  ICMS 41/06, de 15 de dezembro de 2006, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12, E13, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.2 � Leitura dos dados gravados na Mem�ria de Fita Detalhe, em conformidade com o disposto no inciso III da cl�usula d�cima segunda do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 20.2.1 da al�nea �b�, do inciso III, da cl�usula s�tima, do Protocolo  ICMS 41/06, de 15 de dezembro de 2006, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E14, E15, E16, E17, E18, E19, E20 e E21 observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.3 � Leitura de qualquer dado gravado nos dispositivos de mem�ria do ECF, em conformidade com o item 20.3 da al�nea �b�, do inciso III, da cl�usula s�tima, do Protocolo ICMS 41/06, de 15 de dezembro de 2006, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12, E13, E14, E15, E16, E17, E18, E19, E20 e E21, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.4 � Recupera��o dos dados constantes na Redu��o Z, em conformidade com o disposto nos incisos V e VI da cl�usula d�cima segunda do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 20.2.3 da al�nea �b�, do inciso III, da cl�usula s�tima, do Protocolo ICMS 41/06, de 15 de dezembro de 2006, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E14, E15 e E16, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.5 � Impress�o de Fita Detalhe, em conformidade com o disposto no inciso IV da cl�usula d�cima segunda do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 20.2.2 da al�nea �b�, do inciso III, da cl�usula s�tima, do Protocolo ICMS 41/06, de 15 de dezembro de 2006;
'5.1.6 � Leitura do Software B�sico do ECF, em conformidade com o disposto no inciso IX da cl�usula vig�sima s�tima do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 20.4 da al�nea �b�, do inciso III, da cl�usula s�tima, do Protocolo  ICMS 41/06, de 15 de dezembro de 2006;
'5.2 � Quando n�o houver informa��o relativa ao tipo de registro que deve ser gerado de acordo com o disposto no item anterior, dever� ser gerado apenas um registro do respectivo tipo devendo:
'5.2.1 � conter a informa��o dos quatro primeiros campos do registro, de modo a identificar o ECF;
'5.2.2 � observar o disposto nos itens 3.1 e 3.2 para os demais campos do registro;
 
'Reda��o original, efeitos at� 28/06/07.

'5 - GERA��O DO ARQUIVO:
'5.1 - O arquivo dever� ser gerado por programa aplicativo desenvolvido pelo fabricante do ECF que contenha as seguintes funcionalidades, devendo cada fun��o possuir comando �nico e exclusivo:
'5.1.1 - Leitura dos dados gravados na Mem�ria Fiscal, em conformidade com o disposto na cl�usula oitava do Conv�nio ICMS 85/01, de 28 de setembro de 2001 ou no � 2� da cl�usula vig�sima terceira do Conv�nio ICMS 156/94, de 7 de dezembro de 1994, conforme o caso, e no item 1 da al�nea �e� do inciso V da cl�usula quinta do Conv�nio ICMS 16/03, de 04 de abril de 2003, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12 e E13, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.2 - Leitura dos dados gravados na Mem�ria de Fita Detalhe, em conformidade com o disposto no inciso III da cl�usula d�cima segunda do Conv�nio ICMS 85/01, de 28 de setembro de 2001 e no item 2.1 da al�nea �e� do inciso V da cl�usula quinta do Conv�nio ICMS 16/03, de 04 de abril de 2003, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E14, E15 e E16, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.3 - Leitura de qualquer dado gravado nos dispositivos de mem�ria do ECF, em conformidade com o disposto no item 3 da al�nea �e� do inciso V da cl�usula quinta do Conv�nio ICMS 16/03, de 04 de abril de 2003, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12, E13, E14, E15 e E16, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.4 - Recupera��o dos dados constantes na Redu��o Z, em conformidade com o disposto nos incisos V e VI da cl�usula d�cima segunda do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 2.3 da al�nea �e� do inciso V da cl�usula quinta do Conv�nio ICMS 16/03, de 04 de abril de 2003, hip�tese em que o arquivo conter� os seguintes tipos de registro: E01, E02, E14, E15 e E16, observado o disposto nos itens 3.1, 3.2 e 5.2 deste ato;
'5.1.5 - Impress�o de Fita Detalhe, em conformidade com o disposto no inciso IV da cl�usula d�cima segunda do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 2.2 da al�nea �e� do inciso V da cl�usula quinta do Conv�nio ICMS 16/03, de 04 de abril de 2003;
'5.1.6 - Leitura do Software B�sico do ECF, em conformidade com o disposto no inciso IX da cl�usula vig�sima s�tima do Conv�nio ICMS 85/01, de 28 de setembro de 2001, e no item 4 da al�nea �e� do inciso V da cl�usula quinta do Conv�nio ICMS 16/03, de 04 de abril de 2003;

'Nova reda��o dada ao item 5.2 pelo Ato COTEPE/ICMS 43/05, efeitos a partir de 22/09/05.

'5.2 � Sendo obrigat�ria a gera��o do registro, considerando o disposto nos itens 7.3.1.1, 7.4.1.1, 7.5.1.1, 7.6.1.1, 7.7.1.1, 7.8.1.1, 7.10.1.1, 7.14.1.1, 7.15.1.1 e 7.16.1.1, e n�o houver informa��o relativa ao tipo de registro, dever� ser gerado apenas um registro do respectivo tipo devendo:
'5.2.1 � conter a informa��o dos quatro primeiros campos do registro, de modo a identificar o ECF;
'5.2.2 � observar o disposto nos itens 3.1 e 3.2 para os demais campos do registro;

'Reda��o original, efeitos at� 21/09/05.

'5.2 - Quando n�o houver informa��o relativa ao tipo de registro que deve ser gerado de acordo com o disposto no item anterior, dever� ser gerado apenas um registro do respectivo tipo devendo:
'5.2.1 - conter a informa��o dos quatro primeiros campos do registro, de modo a identificar o ECF;
'5.2.2 - observar o disposto nos itens 3.1 e 3.2 para os demais campos do registro;
'===============================
'===============================
''''''''''''''''''''''''INICIO
'N� Denomina��o do Campo       Conte�do                                                                   Tamanho  Posi��o  Formato
'1  Tipo do registro           "E01"                                                                         3     1-3      X
'2  N�mero de fabrica��o       N� de fabrica��o do ECF                                                       20    4-23     X
'3  MF adicional               Letra indicativa de MF adicional                                              1     24-24    X
'4  Tipo do ECF                Tipo do ECF                                                                   7     25-31    X
'5  Marca                      Marca do ECF                                                                  20    32-51    X
'6  Modelo                     Modelo do ECF                                                                 20    52-71    X
'7  Vers�o do SB               Vers�o atual do Software B�sico do ECF gravada na MF                          10    72-81    X
'8  Data da grava��o do SB     Data da grava��o na MF da vers�o do SB a que se refere o campo 07             8     82-89    D
'9  Hora da grava��o do SB     Hora da grava��o na MF da vers�o do SB a que se refere o campo 07             6     90-95    H
'10 N�mero Seq�encial do ECF   N� de ordem seq�encial do ECF no estabelecimento usu�rio                      3     96-98    N
'11 CNPJ do usu�rio            CNPJ do estabelecimento usu�rio do ECF                                        14    99-112   N
'12 Comando de gera��o         C�digo do comando utilizado para gerar o arquivo, conforme tabela abaixo      3     113-115  X
'13 CRZ inicial                Contador de Redu��es Z do in�cio do per�odo a ser capturado                   6     116-121  N
'14 CRZ final                  Contador de Redu��es Z do final do per�odo a ser capturado                    6     122-127  N
'15 Data Inicial               Data do In�cio do per�odo a ser capturado                                     8     128-135  D
'16 Data final                 Data do fim do per�odo a ser capturado                                        8     136-143  D
'17 Vers�o da biblioteca       Vers�o da biblioteca do fabricante do ECF geradora deste arquivo              8     144-151  X
'18 Vers�o do Ato/COTEPE       Vers�o do Ato/COTEPE                                                          15    152-166  X

   Open PATH_TXT & "nfg" & Month(Date) & Year(Date) & ".txt" For Output As #nfg

SQL = "E01"


   Close #nfg

End Sub
