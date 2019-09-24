Attribute VB_Name = "mdlDECLARA"
'    vbWhite = &HFFFFFF
'    vbLightGray = &HE0E0E0
'    vbGray = &HC0C0C0
'    vbMediumGray = &H808080
'    vbDarkGray = &H404040
'    vbBlack = &H0
'    vbPaleRed = &HC0C0FF
'    vbLightRed = &H8080FF
'    vbRed = &HFF
'    vbMediumRed = &HC0&
'    vbDarkRed = &H80&
'    vbBlackRed = &H40&
'    vbPaleOrange = &HC0E0FF
'    vbLightOrange = &H80C0FF
'    vbOrange = &H80FF&
'    vbMediumOrange = &H40C0&
'    vbDarkOrange = &H4080&
'    vbBlackOrange = &H404080
'    vbPaleYellow = &HC0FFFF
'    vbLightYellow = &H80FFFF
'    vbYellow = &HFFFF
'    vbMediumYellow = &HC0C0&
'    vbDarkYellow = &H8080&
'    vbBlackYellow = &H4040&
'    vbPaleGreen = &HC0FFC0
'    vbLightGreen = &H80FF80
'    vbGreen = &HFF00
'    vbMediumGreen = &HC000&
'    vbDarkGreen = &H8000&
'    vbBlackGreen = &H4000&
'    vbPaleCyan = &HFFFFC0
'    vbLightCyan = &HFFFF80
'    vbCyan = &HFFFF00
'    vbMediumCyan = &HC0C000
'    vbDarkCyan = &H808000
'    vbBlackCyan = &H404000
'    vbPaleBlue = &HFFC0C0
'    vbLightBlue = &HFF8080
'    vbBlue = &HFF0000
'    vbMediumBlue = &HC00000
'    vbDarkBlue = &H800000
'    vbBlackBlue = &H400000
'    vbPalePurple = &HFFC0FF
'    vbLightPurple = &HFF80FF
'    vbPurple = &HFF00FF
'    vbMagenta = &HFF00FF
'    vbMediumPurple = &HC000C0
'    vbDarkPurple = &H800080
'    vbBlackPurple = &H400040

   Public TabTempAccess       As New ADODB.Recordset
   Public TabCupom_ADO        As New ADODB.Recordset
   Public CONECTA_ACCESS      As ADODB.Connection

   Public CONECTA_RETAGUARDA  As New ADODB.Connection
   Public CONECTA_GLOBAL      As New ADODB.Connection
   Public CONECTA_AUXILIAR    As New ADODB.Connection

   Public Servidor_Global     As String
   Public VERSAO_APLICATIVO   As String

   Public cmdParametro        As ADODB.Command

   Public TabPessoa           As New ADODB.Recordset
   Public TabEmail            As New ADODB.Recordset
   Public tabEndereco         As New ADODB.Recordset
   Public TabNF               As New ADODB.Recordset
   Public TabRG               As New ADODB.Recordset
   Public tabEmpresa          As New ADODB.Recordset
   Public TabTipovenda        As New ADODB.Recordset
   Public TabCliente          As New ADODB.Recordset
   Public TabUSU              As New ADODB.Recordset
   Public TabProduto          As New ADODB.Recordset
   Public TabPedidoItem       As New ADODB.Recordset
   Public TabCabeca           As New ADODB.Recordset
   Public TabCEP              As New ADODB.Recordset
   Public TabFornecedor       As New ADODB.Recordset
   Public TabConsulta         As New ADODB.Recordset
   Public TabCONTA            As New ADODB.Recordset
   Public TabTemp             As New ADODB.Recordset
   Public TabLancamento       As New ADODB.Recordset
   Public TabIBGE             As New ADODB.Recordset
   Public TabCAIXA            As New ADODB.Recordset
   Public TabDESCR            As New ADODB.Recordset
   Public TabVENDEDOR         As New ADODB.Recordset
   Public TabFone             As New ADODB.Recordset
   Public TabEQUIPE           As New ADODB.Recordset
   Public TabAUX              As New ADODB.Recordset
   Public TabNOTA             As New ADODB.Recordset
   Public TabBANCO            As New ADODB.Recordset
   Public TabAGENCIA          As New ADODB.Recordset
   Public TabLANCAMENTOITEM   As New ADODB.Recordset
   Public TabCHEQUE           As New ADODB.Recordset
   Public TabCABENTRA         As New ADODB.Recordset
   Public TabItem             As New ADODB.Recordset
   Public TabDEVNOTA          As New ADODB.Recordset
   Public TABITEMDEV          As New ADODB.Recordset
   Public TabInventario       As New ADODB.Recordset
   Public TabNFITEM           As New ADODB.Recordset
   Public TaBPedidoCompraItem As New ADODB.Recordset
   Public TaBCompra           As New ADODB.Recordset

   Global crxApplication      As New CRAXDRT.Application
   Global crxReport           As New CRAXDRT.report
   
'==========TRIBUTAÇÃO=============================
   Public TP2_DE_CONTRIB                     As Double
   Public TP2_DE_NCONTRIB                    As Double
   Public TP2_DE_CMAQ_IMP                    As Double
   Public TP2_DE_NMAQ_IMP                    As Double
   Public TP2_FE_CMAQ_IMP                    As Double
   Public TP2_FE_NMAQ_IMP                    As Double
   Public TP2_FE_CAP_INDU                    As Double
   Public TP2_FE_NAP_INDU                    As Double
   Public CFOP_SAIDA_DENTRO_UF_N             As String
   Public CFOP_SAIDA_FORA_UF_N               As String
   Public CFOP_DEVOLUCAO_ENTRADA_FORA_UF_N   As String
   Public CFOP_DEVOLUCAO_ENTRADA_DENTRO_UF_N As String
   Public CFOP_ENTRADA_FE                    As String
   Public CFOP_ENTRADA_DE                    As String

   Public CFOP_DEVOLUCAO_SAI_DENTRO_UF_N     As String
   Public CFOP_DEVOLUCAO_SAI_FORA_UF_N       As String

   Public UF_EMPRESA_A                       As String
   Public SERIE_NFe_A                        As String
   Public TIPO_REGIME_EMPRESA_A              As String
   Public INDR_Tela_Chamada_NFC              As String
   Public CHAMADA_A                          As String
   Public OBS_A                              As String
   Public TIPO_PEDIDO_A                      As String

   Public ALIQUTOA_PIS_N                     As Double   '2 - PISPPIS = PERCENTUAL DO MPIS PADRAO 0,65%
   Public ALIQUTOA_COFINS_N                  As Double   '2 - COFINSPCOFINS, = 3% PADRAO
   Public CST_PIS_A                          As String
   Public CST_COFINS_A                       As String
   Public CST_ICMS_A                         As String
   Public PERC_BASE_REDUZ_N                  As Double
   Public CST_ORIG_ICMS_N                    As Integer
'=================================================
Public NF_ID_N As Long
'======== BOOLEANAS
   Public INDR_SEQUENCIA         As Boolean
   Public INDR_PreFatura         As Boolean
   Public INDR_OS_VEICULO        As Boolean
   Public INDR_ERRO_TEF          As Boolean
   Public INDR_VENDA_CARTAO      As Boolean
   Public INDR_CONTROLA_ESTOQUE  As Boolean
   Public INDR_FIM               As Boolean
   Public INDR_PRI               As Boolean
   Public INDR_VENDA             As Boolean
   Public Indr_Consulta          As Boolean
   Public INDR_CAIXA             As Boolean
   Public INDR_REMOTO            As Boolean
   Public Indr_Erro              As Boolean
   Public Indr_Cancela_Cupom     As Boolean
   Public INDR_LIBERA_DESCONTO   As Boolean
   Public INDR_DESCONTO_AUTORIZADO    As Boolean
   Public INDR_DESCONTO_CLIENTE       As Boolean
   Public INDR_DESCONTO_FUNCIONARIO   As Boolean
   Public INDR_LiberaPercDesconto   As Boolean
   Public INDR_RECEITA           As Integer
   Public INDR_CUPOM_ABERTO      As Boolean
   Public INDR_EMITE_REDUCAO     As Boolean
   Public MULT_EMPRESA_B           As Boolean
   Public INDR_FUNCIONARIO       As Boolean
   Public INDR_INDUSTRIA_B         As Boolean
   Public booUsaCobranca         As Boolean
   Public INDR_GRAVA             As Boolean
   Public INDR_FORM_ABERTO       As Boolean
   Public USA_TEF                As Boolean
   Public USA_NFe                As Boolean
   Public USA_DOC_FISCAL         As Boolean
   Public RECEBE_PEDIDO_VENDA    As Boolean
   Public LIMPA_PEDIDO           As Boolean
   Public INDR_ESTQ_NEGATIVO     As Boolean
   Public INDR_LEI_12741         As Boolean
   Public INDR_AT_VENDA_MKP      As Boolean
   Public INDR_PRODUTO_PRODUCAO_B  As Boolean
   Public INDR_LEU_POR_CODG_BARRAS  As Boolean
   Public INDR_PEDIDO_VENDA      As Boolean
   Public USA_TAB_PRECO_B         As Boolean 'INDICA QUAL TELA DE VENDA VAI SER CHAMADA
   Public ALTERA_FATURA_B     As Boolean

   Public VENDEDOR_ID_N          As Long
   Public TIPOVENDA_ID_N         As Long
   Public NOTAENTRADA_ID_N       As Long
   Public FORNEC_ID_N            As Long
   Public NUMR_PROD_N            As Long
   Public USUARIO_ATUAL          As Long
   Public USU_LIBERA_VENDA_N     As Long
   Public TITULO_N               As Long
   Public NUMR_LANCAMENTO_N      As Long
   Public TRANSP_ID              As Long
   Public NUMR_LOTE_N            As Long
   Public CLIENTE_ID_N           As Long

   Public PEDIDO_ID_N            As Long
   Public TABELAPRECO_ID_N       As Integer
   Public FORMAPAGTO_ID_N        As Integer
   Public Acao_N                 As Integer

   Public OS_ID_N                As Long
   Public SEQ_ID_N               As Long
   Public PRODUTO_ID_N           As Long
   Public Numero_Contador_Z      As Long
   Public CARTAOBARRA_ID_N       As Long

   Public TIPO_USUARIO           As Integer
   Public NUMR_PARCELA           As Integer
   Public DIAS_PRAZO             As Integer
   Public TIPO_TEF_N             As Integer '1=DICADO ; 2=IP ; 3=DEDICADO
   Public TIPO_ENTRADA_N         As Integer '1=com nota, 2=sem nota
   Public CTR_EMPRESA_N          As Integer
   Public DiasAtrazoCliente_N    As Integer

   Public USA_AUTTAR             As Boolean
   Public USA_POS                As Boolean
   Public INDR_TESTE             As Boolean
   Public USA_NFC_E              As Boolean 'se usa NFEc-e
'======== STRING
   Public Path_IntPos_Entrada_A  As String
   Public Path_IntPos_Saida_A    As String

   Public TIPO_PESSOA_CADASTRO   As String
   Public DATA_ABERTURA_CAIXA    As String
   Public TIPO_NFe_GERAR         As String
   Public CNPJCPF_N              As String
   Public NOME_BANCO_DADOS       As String
   Public PATH_TXT               As String
   Public SERVIDOR_MEGASIM       As String
   Public AUTENTICA_GRID         As String
   Public AUTENTICA_GRID_GLOBAL  As String
   Public SENHA_ADM_SQLSERVER    As String
   Public USUARIO_ADM_SQLSERVER  As String
   Public NOME_A                 As String
   Public SITUACAO_TRIBUT_A      As String
   Public REFERENCIA_A           As String
   Public STATUS_PROD            As String
   Public STATUS_A               As String
   Public ENDERECO_A             As String
   Public Crystaldsn             As String
   Public Crystaldsq             As String
   Public Crystaluid             As String
   Public Crystalpwd             As String
   Public LOCAL_IMAGEM           As String
   Public ESTACAO_CPU            As String
   Public USU_LOGADO             As String
   Public FORMULA_REL            As String
   Public selectionFormulaGrupo  As String
   Public CNPJ_EMPRESA_N         As String
   Public CNPJ_ESTABELECIMENTO_N As String
   Public CNPJ_CRED_CARTAO_ESTAB As String
   Public CCE_EMPRESA_N          As String
   Public PATH_REL               As String
   Public DT_EXP_D               As String
   Public strFormatacao2Digitos  As String
   Public strFormatacao3Digitos  As String
   Public strFormatacao4Digitos  As String
   Public strFormatacao5Digitos  As String
   Public strFormatacao6Digitos  As String
   Public strFormatacao7Digitos  As String
   Public strFormatacao8Digitos  As String
   Public strFormatacaoKilo      As String
   Public CCE_CLIENTE_A          As String
   Public UF_CLIENTE_A           As String
   Public NOME_CLIENTE_A         As String
   Public CODG_NCM_A             As String
   Public CODIGO_BARRAS_A        As String
   Public UNIDADE_MEDIDA_A       As String

'======== DOUBLE
   Public LIMITE_CREDITO_CLI_N   As Double
   Public VALOR_PENDENTE_N       As Double
   Public VALR_DESCONTO_N        As Double
   Public VALOR_RECEBIDO_N       As Double
   Public VALOR_ITEM_N           As Double
   Public VALOR_VAREJO_N         As Double
   Public VALOR_DESCONTO_N       As Double
   Public VALOR_TOTAL_N          As Double
   Public QTDE_PEDIDO            As Double
   Public QTDE_BALANCA           As Double
   Public QTDE_ESTOQUE_N           As Double
   Public VALOR_TOTAL_DESCONTO_N As Double
   Public PERC_DESCONTO_N        As Double
   Public PERC_DESCONTO_USUARIO_N  As Double
   Public Valor_Compra_Dia_Permitida   As Double
   Public PESO_ITEM_N            As Double
   Public QTDE_N                 As Double

   Public VLR_ANTERIOR_N         As Currency
   Public QTD_N                  As Currency
   Public VALOR_IPI_N            As Currency
   Public VALOR_ICMS_N           As Currency
   Public TIPO_DP_EMPRESA        As String
   Public VALOR_DIFERENCA_N      As Currency
   Public VLR_DESCT_DIF_N        As Currency
   Public PERC_JUROS_N           As Currency
   Public VALOR_DEBITO           As Currency
   Public VALOR_CREDITO          As Currency
   Public QTD_CONTROLE           As Currency
   Public VALOR_TOTAL_JUROS_N    As Currency
   Public VLR_JUROS_ATUALIZADO   As Currency
   Public VLR_TITULO_ATUALIZADO  As Currency
   Public PERC_ACUM_N            As Currency
   Public PERC_ACUM_VLR_N        As Currency
   Public MORA_JUROS             As Currency
   Public DIAS_ATRAZO            As Currency
   Public VALOR_PAGAR_FUTUROS    As Currency
   Public VALOR_RECEBER_FUTUROS  As Currency
   Public VALR_SALDO_ATUAL_N     As Currency

   Global OG_Formula_Field    As Object
   Global crxSubReport        As CRAXDRT.report

   Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
   Public Style, Title, Help, Ctxt, MyString

   Declare Function Inscricao Lib "DllInscE32.dll" Alias "ConsisteInscricaoEstadual" (ByVal VALOR As String, ByVal Retorno As String) As Integer 'API PARA INSCRICAO ESTADUAL
   Private Const HH_DISPLAY_TOPIC = &H0
   Private Const HH_HELP_CONTEXT = &HF
   
   Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

   Public Const cMoeda = "##,###,##0.00"

   Public NOME_LOJA As String
   Public RESPOSTA
   Public Nodx
   Public PATH_ARQ As String, PATH_SERVER As String, Nome_Relatorio As String
   Public CNPJCPF_A As String, NOME_EMPRESA_A As String, Msg As String

   Public SQL As String, SqL2 As String, SQL3 As String, CRITERIO_A As String
   Public CODG_PRODUTO_A As String, DESC_PRODUTO_A As String

   Public TIPO_PESSOA            As String

   Public ALIQUOTA_ICMS_NORMAL_DENTRO_UF  As Double
   Public ALIQUOTA_ICMS_NORMAL_FORA_UF    As Double
   Public PESO_LIQUIDO_N         As Double
   Public PESO_BRUTO_N           As Double
   Public PR_ATACADO_N           As Double
   Public PR_VAREJO_N            As Double
   Public PR_CUSTO_PRODUTO_N     As Double

   Public ALIQUOTA_ICMS_SUBST    As Double
   Public ALIQUOTA_ICMS_N        As Double
   Public IMPRESSORA_ID_N        As Long
   Public CONTA_REINICIO_N       As Long
   Public NUMERO_SERIE_ECF       As String
   Public NUMERO_CAIXA_ECF       As Long
   Public CAIXA_DIA_ID_N         As Long
   Public NumeroCancelamentos    As String
   Public NUMERO_CAIXA_CPU       As Long
   Public strRetorno             As String
   Public strConexaoGrid         As String

'======== LONG
   Public EMPRESA_ID_N           As Long
   Public ESTABELECIMENTO_ID_N   As Long
   Public PESSOA_ID_N            As Long
   Public PESSOA_ID_EMPRESA_N    As Long
   Public ENDERECO_ID_N          As Long
   Public MARCA_ID_N             As Long
   Public PESSOAENDERECO_ID_N    As Long

   Public CONTA_REGISTRO_N       As Long
   Public Numero_Pedido_N        As Long
   Public NUMR_SEQ_N             As Long
   Public CONT_N                 As Long
   Public NUMR_ID_N              As Long
   Public USUARIO_ID_N           As Long
   Public ATENDENTE_ID_N         As Long
   Public NUMR_CONSULTA_N        As Long
   Public CONTA_REG_PROGRESSO    As Long

   Public Cidade_Com             As String
   Public Estado_Com             As String
   Public CEP_ID_A               As String

   Public Xrua_A                 As String
   Public Xuf_A                  As String
   Public Xcidade_A              As String
   Public Xbairro_A              As String
   Public Xtipo_A                As String
   Public TxtLogradouro_A        As String

'======== DATETIME
   Public DATA_INI As Date, DATA_FIM As Date
   Public HORA_INI As Date, HORA_FIM As Date

   Public item As Object, ITEM2 As Object
   Public FSO As New FileSystemObject

'======== balança
   Public CasaInicioCodgProdBarra_N    As Integer
   Public TamanhoCodgProdBarra_N    As Integer
   Public TamanhoPesoValorBarra_N   As Integer
   Public PESO_VALOR_A              As String

   'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'=============== loucuras
   Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Public Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Public Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessId As Long
      dwThreadId As Long
   End Type

   Declare Function CreateProcessA Lib "kernel32" _
   (ByVal lpApplicationName As Long, ByVal lpCommandLine As _
   String, ByVal lpProcessAttributes As Long, ByVal _
   lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, lpStartupInfo As _
   STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) _
   As Long

   Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
   As Long) As Long

   Public Const NORMAL_PRIORITY_CLASS = &H20&
   Public Const INFINITE = -1&

   Public Declare Function BlockInput Lib "user32.dll" (ByVal Blk As Boolean) As Boolean

Public Sub ExecCmd(cmdline$)
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   
   'Inicia a strutura STARTUPINFO
   start.cb = Len(start)
   
   'Inicia a aplicação escolhida para ser executada
   'Ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
   NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
   
   'Aguarda a aplicação iniciada terminar
   'Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
   'Ret& = CloseHandle(proc.hProcess)
End Sub
