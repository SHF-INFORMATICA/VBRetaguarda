Attribute VB_Name = "modTEFFuncoes"
' Projeto     : TEFVB
' Módulo      : modFuncoesTEF
' Versão      : 3.7.00
' Função      : Agrega todas as funções do TEF
' Autor       : Mariano Rosa em 11/07/2003
' Alterações  : Mariano Rosa em 17/10/2003
' Suporte     : suporte@microstar.com.br - www.microstar.com.br

' Nas operações de arquivos, os seguintes canais são usados :
' #1 = TEFParam
' #2 = Arquivos abertos e fechados dentro da mesma sub
' #3 = IntPos.tmp
' #4 = TEF.imp
' #5 = TEFParc.Txt

'Option Explicit
''variaveis do shell
'Public Const INFINITE = &HFFFF
''STARTINFO constants
'Private Const STARTF_USESHOWWINDOW = &H1
'Public Enum enSW
'    SW_HIDE = 0
'    SW_NORMAL = 1
'    SW_MAXIMIZE = 3
'SW_MINIMIZE = 6
'End Enum
'
'Private Type PROCESS_INFORMATION
'        hProcess As Long
'        hThread As Long
'        dwProcessId As Long
'        dwThreadId As Long
'End Type
'
'Private Type STARTUPINFO
'        cb As Long
'        lpReserved As String
'        lpDesktop As String
'        lpTitle As String
'        dwX As Long
'        dwY As Long
'        dwXSize As Long
'        dwYSize As Long
'        dwXCountChars As Long
'        dwYCountChars As Long
'        dwFillAttribute As Long
'        dwFlags As Long
'        wShowWindow As Integer
'        cbReserved2 As Integer
'        lpReserved2 As Byte
'        hStdInput As Long
'        hStdOutput As Long
'        hStdError As Long
'End Type
'
'Type SECURITY_ATTRIBUTES
'        nLength As Long
'        lpSecurityDescriptor As Long
'        bInheritHandle As Long
'End Type
'
'Public Enum enPriority_Class
'    NORMAL_PRIORITY_CLASS = &H20
'    IDLE_PRIORITY_CLASS = &H40
'    HIGH_PRIORITY_CLASS = &H80
'End Enum
'
'Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
'        (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
'        lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As _
'        SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal _
'        dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory _
'        As String, lpStartupInfo As STARTUPINFO, lpProcessInformation _
'        As PROCESS_INFORMATION) As Long
'Private Declare Function WaitForSingleObject Lib "kernel32" _
'        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'
''############################################
''##########variáveis de parametros###########
''############################################
'Public strEmiteCupomFiscal As String
'
'Public PAR_LOCAL_Id_Pdv As Byte             'nr do Checkout deste cliente
'
'Public IMPRESSORA_FISCAL_A As String       'qual modelo de impressora Daruma, Bematech, Corisco, Sweda
'Public PAR_LOCAL_ImpUsaPorta As Boolean     'Usa Porta Serial SIM ou NAO
'Public PAR_LOCAL_ImpPorta As Byte           'se usa qual a porta serial ?
'
'Public PAR_LOCAL_UsaBalanca As Boolean      'usa balanca SIM ou NAO?
'Public PAR_LOCAL_BalancaPorta As Byte       'se usa qual a porta da balança?
'Public PAR_LOCAL_BalancaDigitos As Integer  'quantos dígitos aceita a balança?
'
'Public PAR_LOCAL_TemGaveta As Boolean       'usa gaveta no caixa?
'
'Public PAR_LOCAL_UsaLeitorSerial As Boolean 'usa leitor serial SIM ou NAO?
'Public PAR_LOCAL_PortaLeitorSerial As Byte  'se usa qual a porta do leitor serial?
'
'Public PAR_LOCAL_TipoTEF As String          'TEF DIS - Discado ou DED - Dedicado
'Public PAR_LOCAL_VendaCartao As Boolean     'TEF vende no cartão (Crédito, Débito etc...) TEF ?
'Public PAR_LOCAL_Visa As Boolean            'TEF Usa cartão Visa
'Public PAR_LOCAL_MasterCard As Boolean      'TEF Usa cartão MasterCard
'Public PAR_LOCAL_AmericanExpress As Boolean 'TEF Usa cartão AmericanExpress
'Public PAR_LOCAL_TecBAN As Boolean          'TEF Usa TecBAN
'
'Public PAR_GERAL_Mensagem01 As String       'mensagem 01 que vai no cupom "OBRIGADO E VOLTE SEMPRE"
'Public PAR_GERAL_Mensagem02 As String       'mensagem 02 que vai no cupom "FONE (62) 3091-6340"
'
'Public PAR_GERAL_ControlaCaixa As Boolean   'a empresa vai ter controle de caixa  SIM ou NAO?
'Public PAR_GERAL_UsaOrcamento As Boolean    'a empresa vai fazer orçamento  SIM ou NAO?
'
'Public PAR_GERAL_TelaVenda As Integer       'Qual o modelo de tela de venda irá usar 1-Vendas01 ou 2-Vendas02 ?
'Public PAR_GERAL_TipoAbertCupom As Integer  'Quando abrir o cupom fiscal irá pedir (1 - informar o cpf, 2 - informar o nr mesa, 3 - nenhum)
'
'Public PAR_GERAL_DescontoValor As Boolean   'CUPOM Tipo de desconto é ($)-valor ou (%)-percentual
'Public PAR_GERAL_DescPorcentagem As Boolean 'CUPOM desconto porcentagem
'
'Public PAR_GERAL_CodigoFilial As Integer    'Codigo da empresa
'
''Public LocalRetorno As String
'
''############################################
''###########variáveis de controle############
''############################################
'Public ID_Cupom As String                   'último cupom fiscal
'Public UltimaVendaCartao As Boolean         'última venda realizada é cartão ou não ?
'Public CaixaCupomAberto As Boolean          'O Caixa está com o cupom aberto SIM ou NÃO ?
'Public lst As ListItem                      'variável que usa na listview
'
'Public FechouCupom As Boolean               '???????
'Public FlagGeral As Integer                 '???????
'Public FlagSim As Boolean                   '???????
'Public ArquivoTXT As String                 '???????
'Public NrRegistro As Long                   '???????
'
'Public FlagUsuario As Boolean               '???????
'Public strNomeMaquina As String             'Nome do computador que está usando o caixa
'
''############################################
''##########informações dos cheques###########
''############################################
'Public NrBanco As Integer
'Public NrAgencia As String * 10
'Public NrConta As String * 10
'Public NrCheque As String * 15
'Public DtDeposito As String * 10
'Public StrForma As String
'
'
'
'
'
'
'
'
'
'
'
'Public Barreira As Boolean
'
'Public Sangria As Boolean
'Public ChequeParaLancar As Boolean
'Public ValorRecebido As Double
'
'
'
'Public SqlRelatorio As String
'Public Incluindo As Boolean
'
'
'Global CodigoTemp As String
'Global QuantidadeTemp As Double
'Global TransferiuConsulta As Boolean
'Global VaiConsultar As Boolean
'Global VaiDigitar As Boolean
'
'Public Matricula As String
'Public PD As String
'Public Nome As String
'Public CnpjCpf As String
'Public Endereco1 As String
'Public Endereco2 As String
'Public Endereco3 As String
'Public MensagemPromocional As String
'
'Public ArrayAliquota(6) As Byte
'
'
'
''Variáveis
'Public dtDataSistema As Date
'Public strMsg As String
'
'
'Public VoltarParaConsulta As Boolean
'
'
''variáveis lixo
'Public Ret As Integer
'Public a As String
'Public VrDescontoOrcamento As Currency
'
''############### daqui para baixo é tudo do caixa ###############
'Public tpOrcamento As String
'Public ID_Cliente_Fornecedor As String
'Public ID_Orcamento As String
'Public SqlPesquisa As String
'
'Public f As Integer
'
''venda
'Public P1 As String
'Public Qtde As Currency
'Public VrU As Currency
'Public vrt As String
'Public Desc As Currency
'Public CasasDecimais As Currency
'Public TipoDesconto As String
'Public tt As Currency
'Public ProxIndice As Integer
'
''sweda
'Public VrTempEsc28 As String
'Public XX As Integer
'Public Indice As Integer
'Public Pos As Integer
'
''Public Style As String
''Public Title As String
'Public r As String
'
'
'
'Public AlteraValor As Boolean
'Public AcrescimoDesconto As String
'Public TipoAcrescimoDesconto As String
'Public Cpfcnpj As String
'Public Linha1 As String
'Public Linha2 As String
'Public Linha3 As String
'
'
'Public Flag_2P As Boolean
'
''############################################
''###############API do Windows###############
''############################################
'Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long 'pega o nome do computador
'
'Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long 'Desliga o windows
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpexecuta As String, ByVal lpWindowName As String) As Long
'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'
'
'
'
'
'
'
'
'
'' ---< Variáveis Gerais >---
'Public iTEFTempoEspera    As Integer ' Tempo em seg. para esperar resposta do GP
'Public iTEFTempoMensagem  As Integer ' Tempo em seg. para exibição de mensagens
'Public iTEFProximoNSU     As Integer ' Nº sequencial único da transação gerado pelo TEFVB
'Public iTEFTecban         As Integer ' 1 = Usa rede Tecban / 0 = Não usa Tecban
''Public strCaminhoGerenciadorPadrao           As String  ' Caminho do GP
''Public sTEFMsPath         As String  ' Caminho da MsTEF
'Public sTEFRetorno        As String
'
''Declaradas por Wanderley L. Costa em 12/07/2005
'Public sTEFQtdeVias       As Byte
'Public sTEFFormulario     As String
'Public ConcluiFechamento  As Boolean
'Public DesabilitaForm     As Boolean
'
'' ---< Variáveis da transação >---
'
'Public sTEFDoctoVinculado As String ' 002-000  DOCUMENTO FISCAL VINCULADO
'Public sTEFValorTotal     As String ' 003-000  VALOR TOTAL
'Public sTEFMoeda          As String ' 004-000  MOEDA - "0" = Real / "1" = Dollar
'Public sTEFCMC7           As String ' 005-000  CMC-7
'Public sTEFTipoDePessoa   As String ' 006-000  TIPO DE PESSOA - "F"ísica / "J"uridica
'Public sTEFDoctoPessoa    As String ' 007-000  DOCUMENTO DA PESSOA
'Public sTEFDataDoCheque   As String ' 008-000  DATA DO CHEQUE
'Public sTEFStatusTransac  As String ' 009-000  STATUS DA TRANSAÇÃO
'Public sTEFNomeDaRede     As String ' 010-000  NOME DA REDE
'Public sTEFTipoTransac    As String ' 011-000  TIPO DA TRANSAÇÃO
'Public sTEFNSUTransacao   As String ' 012-000  NÚMERO DA TRANSAÇÃO - NSU
'Public sTEFCodAutorizacao As String ' 013-000  CÓDIGO DE AUTORIZAÇÃO DA TRANSAÇÃO
'Public sTEFNumeroLote     As String ' 014-000  NÚMERO DO LOTE DA TRANSAÇÃO
'Public sTEFTsTransacaoH   As String ' 015-000  TIMESTAMP DA TRANSAÇÃO - HOST
'Public sTEFTsTransacaoL   As String ' 016-000  TIMESTAMP DA TRANSAÇÃO - LOCAL
'Public sTEFTipoParcela    As String ' 017-000  TIPO PARCELAMENTO
'Public sTEFDataTransacao  As String ' 022-000  DATA DA TRANSAÇÃO - COMPROVANTE
'Public sTEFHoraTransacao  As String ' 023-000  HORA DA TRANSAÇÃO - COMPROVANTE
'Public sTEFDataPreDatado  As String ' 024-000  DATA PRÉ-DATADO
'Public sTEFNumTransCanc   As String ' 025-000  NÚMERO DA TRANSAÇÃO CANCELADA - NSU
'Public sTEFTsTransCanc    As String ' 026-000  TIMESTAMP DA TRANSAÇÃO CANCELADA
'Public sTEFFinalizacao    As String ' 027-000  FINALIZAÇÃO
'Public sTEFMensOperador   As String ' 030-000  TEXTO ESPECIAL OPERADOR
'Public sTEFMensCliente    As String ' 031-000  TEXTO ESPECIAL CLIENTE
'Public sTEFAutenticacao   As String ' 032-000  AUTENTICAÇÃO
'Public sTEFBanco          As String ' 033-000  BANCO
'Public sTEFAgencia        As String ' 034-000  AGÊNCIA
'Public sTEFAgenciaDC      As String ' 035-000  AGÊNCIA - DC
'Public sTEFCtaCorrente    As String ' 036-000  CONTA CORRENTE
'Public sTEFCtaCorrenteDC  As String ' 037-000  CONTA CORRENTE - DC
'Public sTEFNumCheque      As String ' 038-000  NÚMERO DO CHEQUE
'Public sTEFNumChequeDC    As String ' 039-000  NÚMERO DO CHEQUE  - DC
'Public sTEFAdministradora As String ' 040-000  NOME DA ADMINISTRADORA
'
Public Type regTEFParam
    NSU As Integer
End Type
'
''---< Funções internas do TEFVB >---
'
'' ---< Retorna os caracteres a direita de sString a partir da posição iPos >---
'Public Function RightFromPos(ByVal sString As String, ByVal iPos As Integer) As String

'    RightFromPos = Trim(Mid(sString, iPos, (Len(sString) - iPos + 1)))
'
'    Exit Function
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Function
'
'' ---< Remove as aspas iniciais e finais de sString >---
'Public Function TEFRemoveAspas(ByVal sString As String) As String

'    TEFRemoveAspas = Trim(Replace(sString, Chr(34), ""))
'
'    Exit Function
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Function
'
'' ---< Suspende a execução por iSegundos >---
Public Sub TEFEspere(ByVal iSegundos As Integer)
On Error GoTo ERRO_TRATA
    Dim vInicio As Variant

    vInicio = Time
    While DateDiff("s", vInicio, Time) < iSegundos
    Wend

    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub
'
''---< Funções externas do TEFVB >---
'
'' ---< Exibe uma mensagem por iTEFTempoMensagem segundos e desativa sem intervenção do usuário >---
Public Sub TEFMensagemPopup(ByVal sLinha1 As String, _
                            ByVal sLinha2 As String, _
                            ByVal sLinha3 As String, _
                            ByVal sLinha4 As String)
On Error GoTo ERRO_TRATA

    With frmCaixaTEFMensagemPimpad
        .lblMensagem = sLinha1 & Chr(13) & Chr(10) & _
                       sLinha2 & Chr(13) & Chr(10) & _
                       sLinha3 & Chr(13) & Chr(10) & _
                       sLinha4
        If sTEFFormulario = "Menu" Then
            .Show 1, frmINICIO
        Else
            .Show
            '.Show 1, frmServicoAtendimentoFaturamentoSoFiltros
        End If
        .Refresh
    End With

    TEFEspere (iTEFTempoMensagem)
    Unload frmCaixaTEFMensagemPimpad

    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFMensagemPopup
'
'' ---< Verifica se existe algum arquivo temporário indevidamente >---
'Public Sub TEFVerificaArquivosPendentes(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    If Dir(strCaminhoGerenciadorPadrao & "\REQ\IntPos.Tmp") <> "" Then Kill (strCaminhoGerenciadorPadrao & "\REQ\IntPos.Tmp")
'
'    If Dir(strCaminhoGerenciadorPadrao & "\REQ\IntPos.001") <> "" Then Kill (strCaminhoGerenciadorPadrao & "\REQ\IntPos.001")
'
'    If Dir(strCaminhoGerenciadorPadrao & "\RESP\IntPos.Sts") <> "" Then Kill (strCaminhoGerenciadorPadrao & "\RESP\IntPos.Sts")
'
'    If Dir(App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEF.Imp") <> "" Then Kill (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEF.Imp")
'
'    If Dir(App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEFParc.Txt") <> "" Then Kill (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEFParc.Txt")
'
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVerificaArquivosPendentes
'
'' ---< Limpa as variáveis do TEF >---
'Public Sub TEFLimpaVariaveis()

'
'    sTEFDoctoVinculado = ""
'    sTEFValorTotal = ""
'    sTEFMoeda = ""
'    sTEFCMC7 = ""
'    sTEFTipoDePessoa = ""
'    sTEFDoctoPessoa = ""
'    sTEFDataDoCheque = ""
'    sTEFStatusTransac = ""
'    sTEFNomeDaRede = ""
'    sTEFTipoTransac = ""
'    sTEFNSUTransacao = ""
'    sTEFCodAutorizacao = ""
'    sTEFNumeroLote = ""
'    sTEFTsTransacaoH = ""
'    sTEFTsTransacaoL = ""
'    sTEFTipoParcela = ""
'    sTEFDataTransacao = ""
'    sTEFHoraTransacao = ""
'    sTEFDataPreDatado = ""
'    sTEFNumTransCanc = ""
'    sTEFTsTransCanc = ""
'    sTEFFinalizacao = ""
'    sTEFMensOperador = ""
'    sTEFMensCliente = ""
'    sTEFAutenticacao = ""
'    sTEFBanco = ""
'    sTEFAgencia = ""
'    sTEFAgenciaDC = ""
'    sTEFCtaCorrente = ""
'    sTEFCtaCorrenteDC = ""
'    sTEFNumCheque = ""
'    sTEFNumChequeDC = ""
'    sTEFAdministradora = ""
'
'    'strCaminhoGerenciadorPadrao = "C:\TEF_DISC\"
'    If iTEFTecban = 1 Then strCaminhoGerenciadorPadrao = strCaminhoGerenciadorPadraoTECBAN
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFLimpaVariaveis
'
'' ---< Verifica e abre arquivo TEF.tmp >---
'' sExibeErro : "S" = exibe mensagem, "N" = Não exibe
'Public Sub TEFVerificaTEFTmp(ByVal sExibeErro As String)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    If Dir(strCaminhoGerenciadorPadrao & "\RESP\IntPos.001") <> "" Then
'        Open (strCaminhoGerenciadorPadrao & "\RESP\IntPos.001") For Input As #2
'        If EOF(2) Then sTEFRetorno = "1"
'        Close #2
'    Else
'        sTEFRetorno = "1"
'    End If
'
'    If (sTEFRetorno <> "0") And (sExibeErro = "S") Then MsgBox ("Não existe nenhuma operação pendente")
'
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVerificaTEFTmp
'
'' ---< Verifica se existe o arquivo TEFParam e cria se necessário >---
'Public Sub TEFVerificaTEFParam(CodigoFilial As Integer)

'    Dim TEFParam As regTEFParam
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    sTEFRetorno = "0"
'
'    Open (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEFVB.dat") For Random As #1 Len = Len(TEFParam)
'    If EOF(1) Then
'        TEFParam.NSU = 0
'        Put #1, 1, TEFParam
'    End If
'    Close #1
'
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVerificaTEFParam
'
'' ---< Retorna o próximo NSU para arquivos de mensagens >---
'Public Sub TEFProximoNSU(CodigoFilial As Integer)

'    Dim TEFParam As regTEFParam
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Call TEFVerificaTEFParam(CodigoFilial)
'    If sTEFRetorno = "0" Then
'        Open (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEFVB.dat") For Random As #1 Len = Len(TEFParam)
'        Get #1, 1, TEFParam
'        TEFParam.NSU = TEFParam.NSU + 1
'        Put #1, 1, TEFParam
'        Close #1
'        iTEFProximoNSU = TEFParam.NSU
'    Else
'        sTEFRetorno = "1"
'    End If
'
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFProximoNSU
'
'' ---< Cria o arquivo IntPos.tmp para receber comandos >---
'Public Sub TEFCriaArquivoREQIntPos001()

'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'    sTEFRetorno = "0"
'    Open (strCaminhoGerenciadorPadrao & "\REQ\IntPos.tmp") For Output As #3
'    Close #3
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFCriaArquivoREQIntPos001
'
'' ---< Verifica se alguma operação ficou pendente (queda do sistema) >---
'Public Sub TEFVerificaOperacaoPendente(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Call TEFVerificaArquivosPendentes(CodigoFilial)
'    If sTEFRetorno = "0" Then
'        Call TEFVerificaTEFTmp("N")
'        If sTEFRetorno = "0" Then Call TEFNaoConfirmaOperacao(CodigoFilial)
'    End If
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVerificaOperacaoPendente
'
'' ---< Verifica a integridade do arquivo resposta IntPos.Sts >---
'Public Sub TEFVerificaIntPosSts(ByVal sOperacao As String, ByVal iNSU As Integer)

'    Dim sIntPosSts As String
'    Dim sLinha As String
'    Dim Qtde As Byte
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    sTEFRetorno = "0"
'    sIntPosSts = strCaminhoGerenciadorPadrao & "\RESP\IntPos.Sts"
'    Screen.MousePointer = vbHourglass
'    Qtde = 0
'
'    Do
'        Sleep (1000)
'        Qtde = Qtde + 1
'        If Qtde >= iTEFTempoEspera Then Exit Do
'    Loop Until ((Dir(sIntPosSts) <> ""))
'
'    If Dir(sIntPosSts) <> "" Then
'        Open sIntPosSts For Input As #2
'        If Not EOF(2) Then
'            Line Input #2, sLinha
'            If Trim(sLinha) <> ("000-000 = " & sOperacao) Then sTEFRetorno = "1"
'            Line Input #2, sLinha
'            If Left(sLinha, 7) = "001-000" Then If RightFromPos(sLinha, 11) <> Trim(iNSU) Then sTEFRetorno = "1"
'        Else
'            sTEFRetorno = "1"
'        End If
'        Close #2
'        Kill sIntPosSts
'    Else
'        sTEFRetorno = "1"
'    End If
'
'    Screen.MousePointer = vbDefault
'
'    'If sTEFRetorno <> "0" Then MsgBox "Não houve resposta do Gerenciador Padrão", vbInformation, "AVISO"
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVerificaIntPosSts
'
'' ---< Verifica a integridade do arquivo retorno IntPos.001 >---
'Public Sub TEFVerificaIntPos001(ByVal sOperacao As String, ByVal iNSU As Integer, CodigoFilial As Integer)

'    Dim iParcelas As Integer
'    Dim i As Long
'    Dim sArquivo As String
'    Dim sIntPos001 As String
'    Dim sLinha As String
'    Dim sDoctoRet As String
'    Dim sValorRet As String
'    Dim sRedeRet As String
'    Dim sLinhas As String
'    Dim sParcela As String
'    Dim sVencParcela As String
'    Dim sValorParcela As String
'    Dim sNSUParcela As String
'    Dim Qtde As Integer
'
'    sTEFRetorno = "0"
'    sDoctoRet = ""
'    sValorRet = ""
'    sRedeRet = ""
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    sIntPos001 = strCaminhoGerenciadorPadrao & "\RESP\IntPos.001"
'    Screen.MousePointer = vbHourglass
'
'    'Do
'    '    Tempo 1
'    '    qtde = qtde + 1
'    '    If qtde >= iTEFTempoEspera Then Exit Do
'    'Loop Until ((Dir(sIntPos001) <> ""))
'
'    Do
'        DoEvents
'    Loop Until (Dir(sIntPos001) <> "")
'    Close #2
'    Open sIntPos001 For Input As #2
'    If Not EOF(2) Then
'
'        Line Input #2, sLinha
'        If Trim(sLinha) <> ("000-000 = " & Trim(sOperacao)) Then sTEFRetorno = "1"
'
'        Line Input #2, sLinha
'        If Left(sLinha, 7) = "001-000" Then If RightFromPos(sLinha, 11) <> iNSU Then sTEFRetorno = "1"
'
'        If sTEFRetorno = "0" Then
'            Do
'                Line Input #2, sLinha
'                If Left(sLinha, 7) = "002-000" Then sTEFDoctoVinculado = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "003-000" Then sTEFValorTotal = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "004-000" Then sTEFMoeda = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "005-000" Then sTEFCMC7 = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "006-000" Then sTEFTipoDePessoa = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "007-000" Then sTEFDoctoPessoa = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "008-000" Then sTEFDataDoCheque = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "009-000" Then sTEFStatusTransac = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "010-000" Then sTEFNomeDaRede = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "011-000" Then sTEFTipoTransac = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "012-000" Then sTEFNSUTransacao = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "013-000" Then sTEFCodAutorizacao = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "014-000" Then sTEFNumeroLote = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "015-000" Then sTEFTsTransacaoH = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "016-000" Then sTEFTsTransacaoL = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "017-000" Then sTEFTipoParcela = RightFromPos(sLinha, 11)
'
'                ' ---> Se existe parcelamento
'                If Left(sLinha, 7) = "018-000" Then
'                    iParcelas = RightFromPos(sLinha, 11)
'                    Open (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEFParc.Txt") For Output As #5
'                    For i = 1 To iParcelas
'                        sParcela = i
'                        If i < 10 Then sParcela = "0" + i
'
'                        sVencParcela = ""
'                        sValorParcela = ""
'                        sNSUParcela = ""
'
'                        Line Input #2, sLinha
'                        If Left(sLinha, 3) = "019" Then sVencParcela = RightFromPos(sLinha, 11)
'
'                        Line Input #2, sLinha
'                        If Left(sLinha, 3) = "020" Then sValorParcela = RightFromPos(sLinha, 11)
'
'                        Line Input #2, sLinha
'                        If Left(sLinha, 3) = "021" Then sNSUParcela = RightFromPos(sLinha, 11)
'
'                        While Len(sValorParcela) < 12
'                          sValorParcela = sValorParcela + " "
'                        Wend
'
'                        While Len(sNSUParcela) < 12
'                          sNSUParcela = sNSUParcela + " "
'                        Wend
'
'                        If (sParcela <> "") And _
'                           (sVencParcela <> "") And _
'                           (sValorParcela <> "") And _
'                           (sNSUParcela <> "") Then Print #5, (sParcela + sVencParcela + sValorParcela + sNSUParcela)
'                    Next i
'                    Close #5
'                End If
'
'                If Left(sLinha, 7) = "022-000" Then sTEFDataTransacao = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "023-000" Then sTEFHoraTransacao = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "024-000" Then sTEFDataPreDatado = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "025-000" Then sTEFNumTransCanc = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "026-000" Then sTEFTsTransCanc = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "027-000" Then sTEFFinalizacao = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "028-000" Then sLinhas = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "030-000" Then sTEFMensOperador = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "031-000" Then sTEFMensCliente = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "032-000" Then sTEFAutenticacao = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "033-000" Then sTEFBanco = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "034-000" Then sTEFAgencia = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "035-000" Then sTEFAgenciaDC = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "036-000" Then sTEFCtaCorrente = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "037-000" Then sTEFCtaCorrenteDC = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "038-000" Then sTEFNumCheque = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "039-000" Then sTEFNumChequeDC = RightFromPos(sLinha, 11)
'                If Left(sLinha, 7) = "039-000" Then sTEFAdministradora = RightFromPos(sLinha, 11)
'            Loop Until EOF(2)
'        End If
'    Else
'        sTEFRetorno = "1"
'    End If
'
'    Close #2
'
'    Screen.MousePointer = vbDefault
'
'    If Trim(sTEFMensOperador) <> "" Then
'        If sTEFDoctoVinculado <> "" Then sDoctoRet = "Docto No: " & Trim(sTEFDoctoVinculado)
'        If sTEFValorTotal <> "" Then sValorRet = "Valor: " & Format((sTEFValorTotal / 100), cMoeda)
'        If sTEFNomeDaRede <> "" Then sRedeRet = "Rede: " & sTEFNomeDaRede
'        If sTEFStatusTransac <> "0" Then sTEFMensOperador = sTEFMensOperador & " - Status " & sTEFStatusTransac
'        If (sTEFStatusTransac <> "0") Or (sLinhas = "0") Then
'            MsgBox (sDoctoRet & Chr(13) & Chr(10) & _
'                    sValorRet & Chr(13) & Chr(10) & _
'                    sRedeRet & Chr(13) & Chr(10) & _
'                    sTEFMensOperador), vbInformation
'        Else
'            Call TEFMensagemPopup(sRedeRet, sDoctoRet, sValorRet, sTEFMensOperador)
'        End If
'    End If
'
'    If sTEFStatusTransac <> "0" Then
'        sTEFRetorno = "1"
'        Kill (sIntPos001)
'    End If
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVerificaIntPos001
'
'' ---< Grava a operação no arquivo IntPos.001 >---
' ---< Grava a operação no arquivo IntPos.001 >---
Public Sub TEFGravaOperacao(ByVal sOperacao As String)
On Error GoTo ERRO_TRATA
    Dim sIntPosTmp As String
    Dim sLinha As String
  
    sTEFRetorno = "0"
  
    Call TEFCriaArquivoREQIntPos001
  
    If sTEFRetorno = "0" Then
    
        ' ---> Verifica se existe operações pendentes
        If (sOperacao = "CNF") Or (sOperacao = "NCN") Then
            TEFVerificaTEFTmp ("S")
            If sTEFRetorno = "0" Then
                Open (sTEFPath & "Resp\IntPos.001") For Input As #2
                Do
                    Line Input #2, sLinha
                    If Left(sLinha, 7) = "002-000" Then sTEFDoctoVinculado = RightFROMPos(sLinha, 11)
                    If Left(sLinha, 7) = "010-000" Then sTEFNomeDaRede = RightFROMPos(sLinha, 11)
                    If Left(sLinha, 7) = "012-000" Then sTEFNSUTransacao = RightFROMPos(sLinha, 11)
                    If Left(sLinha, 7) = "027-000" Then sTEFFinalizacao = RightFROMPos(sLinha, 11)
                Loop Until EOF(2)
                Close #2
                Kill (sTEFPath & "Resp\IntPos.001")
            End If
        End If  ' (sOperacao = "CNF") Or (sOperacao = "NCN")
    
        sIntPosTmp = sTEFPath & "REQ\IntPos.tmp"
        Open sIntPosTmp For Output As #3
        Call TEFProximoNSU
        Print #3, ("000-000 = " & Trim(sOperacao))
        Print #3, ("001-000 = " & iTEFProximoNSU)
    
        If (sOperacao = "CHQ") Or _
           (sOperacao = "CRT") Or _
           (sOperacao = "CNC") Or _
           (sOperacao = "CNF") Or _
           (sOperacao = "NCN") Then
           If Trim(sTEFDoctoVinculado) <> "" Then Print #3, ("002-000 = " & Trim(sTEFDoctoVinculado))
        End If
    
        If (sOperacao = "CHQ") Or _
           (sOperacao = "CRT") Or _
           (sOperacao = "CNC") Then
          If Trim(sTEFValorTotal) <> "" Then Print #3, ("003-000 = " & Trim(sTEFValorTotal))
        End If
    
        If (sOperacao = "CRT") Then If Trim(sTEFMoeda) <> "" Then Print #3, ("004-000 = " & Trim(sTEFMoeda))

        If (sOperacao = "CHQ") Or _
           (sOperacao = "CNC") Then
            If Trim(sTEFCMC7) <> "" Then Print #3, ("005-000 = " & Trim(sTEFCMC7))
            If Trim(sTEFTipoDePessoa) <> "" Then Print #3, ("006-000 = " & Trim(sTEFTipoDePessoa))
            If Trim(sTEFDoctoPessoa) <> "" Then Print #3, ("007-000 = " & Trim(sTEFDoctoPessoa))
            If Trim(sTEFDataDoCheque) <> "" Then Print #3, ("008-000 = " & Trim(sTEFDataDoCheque))
        End If

        If (sOperacao = "CNC") Or _
           (sOperacao = "CNF") Or _
           (sOperacao = "NCN") Then
            If Trim(sTEFNomeDaRede) <> "" Then Print #3, ("010-000 = " & Trim(sTEFNomeDaRede))
            If Trim(sTEFNSUTransacao) <> "" Then Print #3, ("012-000 = " & Trim(sTEFNSUTransacao))
        End If

        If (sOperacao = "CNC") Then
            If Trim(sTEFDataTransacao) <> "" Then Print #3, ("022-000 = " & Trim(sTEFDataTransacao))
            If Trim(sTEFHoraTransacao) <> "" Then Print #3, ("023-000 = " & Trim(sTEFHoraTransacao))
        End If

        If (sOperacao = "CNF") Or _
           (sOperacao = "NCN") Then
            If Trim(sTEFFinalizacao) <> "" Then Print #3, ("027-000 = " & Trim(sTEFFinalizacao))
        End If

        If (sOperacao = "CHQ") Or _
           (sOperacao = "CNC") Then
            If Trim(sTEFBanco) <> "" Then Print #3, ("033-000 = " & Trim(sTEFBanco))
            If Trim(sTEFAgencia) <> "" Then Print #3, ("034-000 = " & Trim(sTEFAgencia))
            If Trim(sTEFAgenciaDC) <> "" Then Print #3, ("035-000 = " & Trim(sTEFAgenciaDC))
            If Trim(sTEFCtaCorrente) <> "" Then Print #3, ("036-000 = " & Trim(sTEFCtaCorrente))
            If Trim(sTEFCtaCorrenteDC) <> "" Then Print #3, ("037-000 = " & Trim(sTEFCtaCorrenteDC))
            If Trim(sTEFNumCheque) <> "" Then Print #3, ("038-000 = " & Trim(sTEFNumCheque))
            If Trim(sTEFNumChequeDC) <> "" Then Print #3, ("039-000 = " & Trim(sTEFNumChequeDC))
        End If
    
        Print #3, "999-999 = 0"
        Close #3
        Sleep 1000
        
        Call FileCopy(sTEFPath & "REQ\IntPos.tmp", sTEFPath & "REQ\IntPos.001")
        Kill (sTEFPath & "REQ\IntPos.tmp")
    
        ' ---> Verifica se houve resposta do GP
        Call TEFVerificaIntPosSts(sOperacao, iTEFProximoNSU)
        
        ' ---> Verifica arquivo de retorno do GP
        If sTEFRetorno = "0" And _
           ((sOperacao = "ADM") Or _
            (sOperacao = "CHQ") Or _
            (sOperacao = "CRT") Or _
            (sOperacao = "CNC")) Then
            Call TEFVerificaIntPos001(sOperacao, iTEFProximoNSU)
        End If
  
    End If  ' sTEFRetorno = "0"
    
    Exit Sub
ERRO_TRATA:
    If Err.Number = 76 Then
                
    ElseIf Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFGravaOperacao
'
'' ---< Imprime a transação >---
'Public Sub TEFImprimeTransacao()

'    sTEFRetorno = "0"
'    'frmImpressao.memEditor.FileName = App.Path & "\" & Format(CodigoFilial, "00") & "\" &"TEF.IMP"
'    'frmImpressao.Show
'    'frmImpressao.Refresh
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub
'
'' ---< Executa a confirmação da transação por meio de cartão >---
'Public Sub TEFConfirmaOperacao(CodigoFilial As Integer)

'    Dim sLinha As String
'    Dim iLinhas As Integer
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    sTEFRetorno = "0"
'
'    Call TEFVerificaTEFTmp("S")
'    If sTEFRetorno = "0" Then
'        iLinhas = 0
'        Open (strCaminhoGerenciadorPadrao & "\RESP\IntPos.001") For Input As #2
'        Open (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEF.Imp") For Output As #4
'        Do
'            Line Input #2, sLinha
'            If Left(sLinha, 3) = "029" Then
'                Print #4, TEFRemoveAspas(RightFromPos(sLinha, 11))
'                iLinhas = iLinhas + 1
'            End If
'        Loop Until EOF(2)
'        Close #2
'        Close #4
'        If iLinhas = 0 Then Kill (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEF.Imp")
'    End If
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub
'
'' ---< Executa a não confirmação da transação por meio de cartão >---
'Public Sub TEFNaoConfirmaOperacao(CodigoFilial As Integer)

'    Dim sLinha As String
'    Dim sDocto As String
'    Dim sValor As String
'    Dim sRede As String
'    Dim sTextoValor As String
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    sTEFRetorno = "0"
'
'    Call TEFVerificaTEFTmp("S")
'    If sTEFRetorno = "0" Then
'        Open (strCaminhoGerenciadorPadrao & "\RESP\IntPos.001") For Input As #2
'        Do
'            Line Input #2, sLinha
'            If Left(sLinha, 7) = "002-000" Then sDocto = RightFromPos(sLinha, 11)
'            If Left(sLinha, 7) = "003-000" Then sValor = RightFromPos(sLinha, 11)
'            If Left(sLinha, 7) = "010-000" Then sRede = RightFromPos(sLinha, 11)
'        Loop Until EOF(2)
'        Close #2
'        If sValor <> "" Then
'            sTextoValor = "Valor: " & (sValor / 100)
'        Else
'            sTextoValor = ""
'        End If
'        Call TEFMensagemPopup("Última Transação TEF Foi Cancelada:", ("Doc No: " & Trim(sDocto)), ("Rede: " & Trim(sRede)), sTextoValor)
'        Call TEFGravaOperacao("NCN", CodigoFilial)
'    End If
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFNaoConfirmaOperacao
'
' ---< Verifica se o Gerenciador Padrão está ativo >---
Public Sub TEFVerificaGerenciadorAtivo()
On Error GoTo ERRO_TRATA
    Dim sLinha As String
    Dim iTask As Long
  
    sTEFRetorno = "0"
  
    ' ---< Coloca valores default, se não existir >---
    
    If iTEFTempoEspera < 1 Then iTEFTempoEspera = 10
    If iTEFTempoMensagem < 1 Then iTEFTempoMensagem = 5
      
    Screen.MousePointer = vbHourglass
      
    ' ---< Se pasta não existir, cria >---
    If Dir(sTEFMsPath, vbDirectory) = "" Then MkDir (sTEFMsPath)
  
    Call TEFGravaOperacao("ATV")
    If sTEFRetorno <> "0" Then
        Call TEFMensagemPopup("Gerenciador Padrão não está ativo e será ativado automaticamente", "", "", "")
        If iTEFTecban = 0 Then
            Shell "C:\TEF_DIAL\TEF_DIAL.EXE", vbNormalFocus
        Else
            Shell "C:\TEF_DISC\TEF_DISC.EXE", vbNormalFocus
        End If
        sTEFRetorno = "0"
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub
'
'' ---< Executa o Módulo Administrativo do GP >---
'Public Sub TEFModoAdministrativo(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Call TEFVerificaGerenciadorAtivo(CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFGravaOperacao("ADM", CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFConfirmaOperacao(CodigoFilial)
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFModoAdministrativo
'
'' ---< Executa pedido de autorização para transação por meio de cartão >---
'Public Sub TEFPedidoAutorizacaoCartao(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'
'    Call TEFVerificaGerenciadorAtivo(CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFGravaOperacao("CRT", CodigoFilial)
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFPedidoAutorizacaoCartao
'
'' ---< Executa pedido de autorização para transação por meio de cheque >---
'Public Sub TEFPedidoAutorizacaoCheque(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Call TEFVerificaGerenciadorAtivo(CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFGravaOperacao("CHQ", CodigoFilial)
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFPedidoAutorizacaoCheque
'
'' ---< Executa a venda por meio de cartão >---
'Public Sub TEFVendaCartao(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'
'    Call TEFPedidoAutorizacaoCartao(CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFConfirmaOperacao(CodigoFilial)
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVendaCartao
'
'' ---< Executa a venda por meio de cheque >---
'Public Sub TEFVendaCheque(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Call TEFPedidoAutorizacaoCheque(CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFConfirmaOperacao(CodigoFilial)
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFVendaCheque
'
'' ---< Cancela Transação efetivada >---
'Public Sub TEFCancelaTransacao(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'
'    Call TEFVerificaGerenciadorAtivo(CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFGravaOperacao("CNC", CodigoFilial)
'    If sTEFRetorno = "0" Then Call TEFConfirmaOperacao(CodigoFilial)
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub ' TEFCancelaTransacao
'
'' ---< Confirma e encerra a transação do TEF >---
'Public Sub TEFFechaOperacao(CodigoFilial As Integer)

'    sTEFRetorno = "0"
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Call TEFGravaOperacao("CNF", CodigoFilial)
'    If (sTEFRetorno = "0") And (Dir(App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEF.Imp") <> "") Then Kill (App.Path & "\" & Format(CodigoFilial, "00") & "\" & "TEF.IMP")
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
'    End If
'End Sub
'
'
'Sub InicializandoVariaveis()
'    CaixaCupomAberto = False
'    'se caiu energia e o cupom estava aberto e quando retornar o sistema esta variavel estara false, mas na verdade é verdadeira
'    VerificaCaixaCupomAberto
'
'    ID_Cupom = 0
'
'    strNomeMaquina = GetComputer
'End Sub
'
'Sub VerificaStatusImpressora()

'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Select Case IMPRESSORA_FISCAL_A
'        Case "Daruma"
'            'strRetorno = 0 significa que está com algum problema (retorno impressora sem problemas)
'            'strRetorno = 1 significa que o retorno não tem problemas (retorno impressora com problemas)
'
'            'intRetorno = 0 status da impressora é falso (impressora com problemas)
'            'intRetorno = 1 status da impressora é true (impressora sem problemas)
'
'            'status cupom fiscal
''            strRetorno = Space(2)
''            intRetorno = Daruma_FI_StatusCupomFiscal(strRetorno)
''            If Trim(strRetorno) = "" Then strRetorno = 1
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                VerificaRetornoImpressoraDaruma "", "", "Status Cupom Fiscal"
''            End If
''
''            status relatório gerencial
''            strRetorno = Space(2)
''            intRetorno = Daruma_FI_StatusRelatorioGerencial(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                VerificaRetornoImpressoraDaruma "", "", "Status Cupom Fiscal"
''            End If
''
''            status comprovante não fiscal vinculado 'TEF
''            strRetorno = Space(2)
''            intRetorno = Daruma_FI_StatusComprovanteNaoFiscalVinculado(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                VerificaRetornoImpressoraDaruma "", "", "Status Cupom Fiscal"
''            End If
''
''            status comprovante não fiscal não vinculado
''            strRetorno = Space(2)
''            intRetorno = Daruma_FI_StatusComprovanteNaoFiscalNaoVinculado(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                VerificaRetornoImpressoraDaruma "", "", "Status Cupom Fiscal"
''            End If
'
'        Case "Bematech MP-4000 TH FI"
'
'            'status cupom fiscal
''            strRetorno = Space(20)
''            intRetorno = Bematech_FI_VerificaEstadoImpressoraMFD(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                'call VerificaRetornoImpressoraBematech "", "", "Status Cupom Fiscal"
''            End If
'
''            Dim Ack As Integer
''            Dim st1 As Integer
''            Dim st2 As Integer
'
''            Retorno = Bematech_FI_VerificaEstadoImpressora(Ack, st1, st2)
''            If Retorno = 1 Or Retorno = -27 Then
''                MsgBox "Estado da Impressora: " + Str(Ack) + "," + Str(st1) + "," + Str(st2), vbOKOnly, "Informações da Impressora"
''            Else
''                Call VerificaRetornoImpressora("Estado da Impressora: ", Str(Ack) + "," + Str(st1) + "," + Str(st2), "Informações da Impressora")
''            End If
'
'
'
'            'status relatório gerencial
''            strRetorno = Space(2)
''            intRetorno = Bematech_FI_StatusRelatorioGerencial(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                'call VerificaRetornoImpressoraBematech "", "", "Status Cupom Fiscal"
''            End If
'
'            'status comprovante não fiscal vinculado 'TEF
''            strRetorno = Space(2)
''            intRetorno = Bematech_FI_StatusComprovanteNaoFiscalVinculado(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                'call VerificaRetornoImpressoraBematech "", "", "Status Cupom Fiscal"
''            End If
'
'            'status comprovante não fiscal não vinculado
''            strRetorno = Space(2)
''            intRetorno = Bematech_FI_StatusComprovanteNaoFiscalNaoVinculado(strRetorno)
''            If CLng(CStr(strRetorno)) = 0 And intRetorno = 1 Then 'está tudo bem sem problemas
''
''            Else 'aqui tem problemas
''                'call VerificaRetornoImpressoraBematech "", "", "Status Cupom Fiscal"
''            End If
'
'        Case "Corisco"
'            MsgBox "Não Foi Desenvolvido."
'        Case "Sweda IF"
'            MsgBox "Não Foi Desenvolvido."
'    End Select
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        'trata_erros Err.Description, "Verifica Status da Impressora"
'    End If
'End Sub
'
'
'Sub HorarioDeVerao()
''    Sweda.PortOpen = True
''    Sweda.Output = Chr(27) & ".28}" 'Reducao Z
''
''    Tempo 1
''
''    a = Sweda.Input
''
''    Sta_Verao = Mid(a, 57, 1)
''
''    If Sta_Verao = "S" Then
''        If Date > CDate("01/02/2001") Then
''            r = MsgBox("A impressora está com o horário de verão programado. Deseja desprogramá-lo?", vbYesNo + vbExclamation + vbDefaultButton2, " A V I S O ")
''            If r = 6 Then Sweda.Output = Chr(27) & ".36N}"       'Sai do horário de verão
''            MsgBox "Mudança do horário concluída com sucesso."
''        Else
''            MsgBox "Data atual inferior a data permitida para desprogramação do horário de verão.", vbCritical + vbOKOnly, " A V I S O "
''        End If
''    Else
''        If Date > CDate("01/10/2001") Then
''            FrmAviso.LblMensagem.Caption = "Deseja programar o horário de verão?"
''            FrmAviso.Show 1
''            If FlagSim Then
''                Sweda.Output = Chr(27) & ".36S}"  'Entra no horário de verão
''                MsgBox "Mudança do horário concluída com sucesso."
''            End If
''        Else
''            MsgBox "Data atual inferior a data permitida para programação do horário de verão.", vbCritical + vbOKOnly, " A V I S O "
''        End If
''    End If
''    Tempo 0.8
''    a = Sweda.Input
''    Sweda.PortOpen = False
'End Sub
'
'
'Public Sub VerificaCaixaCupomAberto()

'
'    '
'    'rs.Open "SELECT * FROM VENDATEMPORARIA WHERE ID_PDV = " & PAR_LOCAL_Id_Pdv, CONECTA_RETAGUARDA , , , adCmdText
'    'If Not rs.EOF Then
'    '    CaixaCupomAberto = True
'    'Else
'    '    CaixaCupomAberto = False
'    'End If
'    'rs.Close
'
'    Exit Sub
'ERRO_TRATA:
'
'    'trata_erros Err.Description, Me.Name, "VerificaCaixaCupomAberto"
'End Sub
'
'Public Function GetComputer() As String
'    Dim BUFFER As String
'    Dim Size As Long
'    Dim dl As Long
'    Size = 199
'    BUFFER = String(200, 0)
'    dl = GetComputerName(BUFFER, Size)
'    If dl <> 0 Then
'      GetComputer = Left$(BUFFER, Size)
'    Else
'      GetComputer = ""
'    End If
'End Function
'
'Public Function SuperShell(ByVal App As String, ByVal WorkDir As String, _
'        dwMilliseconds As Long, ByVal start_size As enSW, ByVal Priority_Class _
'        As enPriority_Class) As Boolean
'
'
'    Dim pclass As Long
'    Dim sinfo As STARTUPINFO
'    Dim pinfo As PROCESS_INFORMATION
'    'Not used, but needed
'    Dim sec1 As SECURITY_ATTRIBUTES
'    Dim sec2 As SECURITY_ATTRIBUTES
'
'    sec1.nLength = Len(sec1)
'    sec2.nLength = Len(sec2)
'    sinfo.cb = Len(sinfo)
'
'    sinfo.dwFlags = STARTF_USESHOWWINDOW
'    sinfo.wShowWindow = start_size
'
'    pclass = Priority_Class
'
'    If CreateProcess(vbNullString, App, sec1, sec2, False, pclass, _
'        0&, WorkDir, sinfo, pinfo) Then
'        WaitForSingleObject pinfo.hProcess, dwMilliseconds
'        SuperShell = True
'    Else
'        SuperShell = False
'    End If
'
'End Function
' ---< Retorna o próximo NSU para arquivos de mensagens >---
Public Sub TEFProximoNSU()
On Error GoTo ERRO_TRATA
    Dim TEFParam As regTEFParam
  
    Call TEFVerificaTEFParam
    If sTEFRetorno = "0" Then
        Open (sTEFMsPath & "TEFVB.dat") For Random As #1 Len = Len(TEFParam)
        Get #1, 1, TEFParam
        TEFParam.NSU = TEFParam.NSU + 1
        Put #1, 1, TEFParam
        Close #1
        iTEFProximoNSU = TEFParam.NSU
    Else
        sTEFRetorno = "1"
    End If
  
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFProximoNSU

' ---< Verifica a integridade do arquivo resposta IntPos.Sts >---
Public Sub TEFVerificaIntPosSts(ByVal sOperacao As String, ByVal iNSU As Integer)
On Error GoTo ERRO_TRATA
    Dim sIntPosSts As String
    Dim sLinha As String
    Dim Qtde As Byte
  
    sTEFRetorno = "0"
    sIntPosSts = sTEFPath & "RESP\IntPos.sts"
    Screen.MousePointer = vbHourglass
    Qtde = 0
    
    Do
        Tempo 1
        Qtde = Qtde + 1
        If Qtde >= iTEFTempoEspera Then Exit Do
    Loop Until ((Dir(sIntPosSts) <> ""))
  
    If Dir(sIntPosSts) <> "" Then
        Open sIntPosSts For Input As #3
        If Not EOF(3) Then
            Line Input #3, sLinha
            If Trim(sLinha) <> ("000-000 = " & sOperacao) Then sTEFRetorno = "1"
            Line Input #3, sLinha
            If Left(sLinha, 7) = "001-000" Then If RightFROMPos(sLinha, 11) <> iNSU Then sTEFRetorno = "1"
        Else
            sTEFRetorno = "1"
        End If
        Close #3
        Kill sIntPosSts
    Else
        sTEFRetorno = "1"
    End If
  
    Screen.MousePointer = vbDefault

    If sTEFRetorno <> "0" Then MsgBox "Não houve resposta do Gerenciador Padrão", vbInformation, "AVISO"
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFVerificaIntPosSts

' ---< Verifica a integridade do arquivo retorno IntPos.001 >---
Public Sub TEFVerificaIntPos001(ByVal sOperacao As String, ByVal iNSU As Integer)
On Error GoTo ERRO_TRATA
    Dim iParcelas As Integer
    Dim i As Integer
    Dim sArquivo As String
    Dim sIntPos001 As String
    Dim sLinha As String
    Dim sDoctoRet As String
    Dim sValorRet As String
    Dim sRedeRet As String
    Dim sLinhas As String
    Dim sParcela As String
    Dim sVencParcela As String
    Dim sValorParcela As String
    Dim sNSUParcela As String
    Dim Qtde As Integer
  
    sTEFRetorno = "0"
    sDoctoRet = ""
    sValorRet = ""
    sRedeRet = ""
    sIntPos001 = sTEFPath & "RESP\IntPos.001"
    Screen.MousePointer = vbHourglass
    
    'Do
    '    Tempo 1
    '    qtde = qtde + 1
    '    If qtde >= iTEFTempoEspera Then Exit Do
    'Loop Until ((Dir(sIntPos001) <> ""))
    
    Do
        DoEvents
    Loop Until (Dir(sIntPos001) <> "")
    Tempo 3
    Open sIntPos001 For Input As #2
    Sleep 1000
    If Not EOF(2) Then
    
        Line Input #2, sLinha
        If Trim(sLinha) <> ("000-000 = " & Trim(sOperacao)) Then sTEFRetorno = "1"
        
        Line Input #2, sLinha
        If Left(sLinha, 7) = "001-000" Then If RightFROMPos(sLinha, 11) <> iNSU Then sTEFRetorno = "1"
    
        If sTEFRetorno = "0" Then
            Do
                Line Input #2, sLinha
                If Left(sLinha, 7) = "002-000" Then sTEFDoctoVinculado = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "003-000" Then sTEFValorTotal = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "004-000" Then sTEFMoeda = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "005-000" Then sTEFCMC7 = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "006-000" Then sTEFTipoDePessoa = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "007-000" Then sTEFDoctoPessoa = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "008-000" Then sTEFDataDoCheque = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "009-000" Then sTEFStatusTransac = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "010-000" Then sTEFNomeDaRede = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "011-000" Then sTEFTipoTransac = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "012-000" Then sTEFNSUTransacao = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "013-000" Then sTEFCodAutorizacao = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "014-000" Then sTEFNumeroLote = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "015-000" Then sTEFTsTransacaoH = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "016-000" Then sTEFTsTransacaoL = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "017-000" Then sTEFTipoParcela = RightFROMPos(sLinha, 11)
                    
                ' ---> Se existe parcelamento
                If Left(sLinha, 7) = "018-000" Then
                    iParcelas = RightFROMPos(sLinha, 11)
                    Open (sTEFMsPath & "TEFParc.Txt") For Output As #5
                    For i = 1 To iParcelas
                        sParcela = i
                        If i < 10 Then sParcela = "0" + i
                      
                        sVencParcela = ""
                        sValorParcela = ""
                        sNSUParcela = ""
                      
                        Line Input #2, sLinha
                        If Left(sLinha, 3) = "019" Then sVencParcela = RightFROMPos(sLinha, 11)
                        
                        Line Input #2, sLinha
                        If Left(sLinha, 3) = "020" Then sValorParcela = RightFROMPos(sLinha, 11)
                        
                        Line Input #2, sLinha
                        If Left(sLinha, 3) = "021" Then sNSUParcela = RightFROMPos(sLinha, 11)
                      
                        While Len(sValorParcela) < 12
                          sValorParcela = sValorParcela + " "
                        Wend
                        
                        While Len(sNSUParcela) < 12
                          sNSUParcela = sNSUParcela + " "
                        Wend
                      
                        If (sParcela <> "") And _
                           (sVencParcela <> "") And _
                           (sValorParcela <> "") And _
                           (sNSUParcela <> "") Then Print #5, (sParcela + sVencParcela + sValorParcela + sNSUParcela)
                    Next i
                    Close #5
                End If
        
                If Left(sLinha, 7) = "022-000" Then sTEFDataTransacao = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "023-000" Then sTEFHoraTransacao = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "024-000" Then sTEFDataPreDatado = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "025-000" Then sTEFNumTransCanc = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "026-000" Then sTEFTsTransCanc = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "027-000" Then sTEFFinalizacao = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "028-000" Then sLinhas = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "030-000" Then sTEFMensOperador = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "031-000" Then sTEFMensCliente = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "032-000" Then sTEFAutenticacao = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "033-000" Then sTEFBanco = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "034-000" Then sTEFAgencia = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "035-000" Then sTEFAgenciaDC = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "036-000" Then sTEFCtaCorrente = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "037-000" Then sTEFCtaCorrenteDC = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "038-000" Then sTEFNumCheque = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "039-000" Then sTEFNumChequeDC = RightFROMPos(sLinha, 11)
                If Left(sLinha, 7) = "039-000" Then sTEFAdministradora = RightFROMPos(sLinha, 11)
            Loop Until EOF(2)
        End If
    Else
        sTEFRetorno = "1"
    End If
  
    Close #2
  
    Screen.MousePointer = vbDefault
  
    If Trim(sTEFMensOperador) <> "" Then
        If sTEFDoctoVinculado <> "" Then sDoctoRet = "Docto No: " & Trim(sTEFDoctoVinculado)
        If sTEFValorTotal <> "" Then sValorRet = "Valor: " & Format((sTEFValorTotal / 100), cMoeda)
        If sTEFNomeDaRede <> "" Then sRedeRet = "Rede: " & sTEFNomeDaRede
        If sTEFStatusTransac <> "0" Then sTEFMensOperador = sTEFMensOperador & " - Status " & sTEFStatusTransac
        If (sTEFStatusTransac <> "0") Or (sLinhas = "0") Then
            MsgBox (sDoctoRet & Chr(13) & Chr(10) & _
                    sValorRet & Chr(13) & Chr(10) & _
                    sRedeRet & Chr(13) & Chr(10) & _
                    sTEFMensOperador), vbInformation
        Else
            Call TEFMensagemPopup(sRedeRet, sDoctoRet, sValorRet, sTEFMensOperador)
        End If
    End If
  
    If sTEFStatusTransac <> "0" Then
        sTEFRetorno = "1"
        Kill (sIntPos001)
    End If
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFVerificaIntPos001


Public Sub TEFVerificaTEFParam()
On Error GoTo ERRO_TRATA
    Dim TEFParam As regTEFParam
  
    sTEFRetorno = "0"
  
    Open (sTEFMsPath & "TEFVB.dat") For Random As #1 Len = Len(TEFParam)
    If EOF(1) Then
        TEFParam.NSU = 0
        Put #1, 1, TEFParam
    End If
    Close #1

    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFVerificaTEFParam
