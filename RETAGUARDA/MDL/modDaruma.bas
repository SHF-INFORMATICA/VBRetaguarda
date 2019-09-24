Attribute VB_Name = "modDaruma"
Global Int_Retorno As Integer
Global Int_Ack As Integer
Global Int_St1 As Integer
Global Int_St2 As Integer
Global Int_Tipo_Aliquota As Integer
Global Str_Informacao As String
Global Int_Informacao As Integer
'======================================

Option Explicit
 '================================ DECLARACOES DARUMA FRAMEWORK ================================'
    '===========                           IMPRESSORAS FISCAL                          ============'

Public Declare Function Daruma_FIMFD_AcionarGuilhotina Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RetornoImpressora Lib "Daruma32.dll" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer
Public Declare Function Daruma_TEF_TravarTeclado Lib "Daruma32.dll" (ByVal Travar As String) As Integer

'Abertura de cupom fiscal
Public Declare Function iCFAbrir_ECF_Daruma Lib "DarumaFramework.dll" (ByVal CPF As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function iCFAbrirPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Registro de item
Public Declare Function iCFVender_ECF_Daruma Lib "DarumaFramework.dll" (ByVal Aliq As String, ByVal Qtd As String, ByVal PrecoUn As String, ByVal TipoDescAcresc As String, ByVal VlrDescAcresc As String, ByVal CodItem As String, ByVal Un As String, ByVal DescricaoItem As String) As Integer
Public Declare Function iCFVenderSemDesc_ECF_Daruma Lib "DarumaFramework.dll" (ByVal Aliq As String, ByVal Qtd As String, ByVal PrecoUn As String, ByVal CodItem As String, ByVal Un As String, ByVal DescricaoItem As String) As Integer
Public Declare Function iCFVenderResumido_ECF_Daruma Lib "DarumaFramework.dll" (ByVal Aliq As String, ByVal PrecoUn As String, ByVal CodItem As String, ByVal DescricaoItem As String) As Integer

'Desconto ou acrescimo  em item de cupom fiscal
Public Declare Function iCFLancarAcrescimoItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNumItem As String, ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String) As Integer
Public Declare Function iCFLancarDescontoItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNumItem As String, ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String) As Integer
Public Declare Function iCFLancarAcrescimoUltimoItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String) As Integer
Public Declare Function iCFLancarDescontoUltimoItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String) As Integer

'Cancelamento total de item em cupom fiscal
Public Declare Function iCFCancelarUltimoItem_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCFCancelarItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal NumItem As String) As Integer

'Cancelamento parcial de item em cupom fiscal
Public Declare Function iCFCancelarItemParcial_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNumItem As String, ByVal pszQuantidade As String) As Integer
Public Declare Function iCFCancelarUltimoItemParcial_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszQuantidade As String) As Integer

'Cancelamento de desconto em item
Public Declare Function iCFCancelarDescontoItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNumItem As String) As Integer
Public Declare Function iCFCancelarDescontoUltimoItem_ECF_Daruma Lib "DarumaFramework.dll" () As Integer


'Totalizacao de cupom fiscal
Public Declare Function iCFTotalizarCupom_ECF_Daruma Lib "DarumaFramework.dll" (ByVal TipoDescAcresc As String, ByVal VlrDescAcresc As String) As Integer
Public Declare Function iCFTotalizarCupomPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Cancelamento de desconto e acrescimo em subtotal de cupom fiscal
Public Declare Function iCFCancelarDescontoSubtotal_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCFCancelarAcrescimoSubtotal_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Descricao do meios de pagamento de cupom fiscal
Public Declare Function iCFEfetuarPagamentoPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCFEfetuarPagamentoFormatado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszFormaPgto As String, ByVal pszValor As String) As Integer
Public Declare Function iCFEfetuarPagamento_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszFormaPgto As String, ByVal pszValor As String, ByVal pszInfoAdicional As String) As Integer

'Saldo a Pagar
Public Declare Function rCFSaldoAPagar_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String) As Integer

'SubTotal
Public Declare Function rCFSubTotal_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String) As Integer

'Encerramento de cupom fiscal
Public Declare Function iCFEncerrarPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCFEncerrarConfigMsg_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszMensagem As String) As Integer
Public Declare Function iCFEncerrar_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszCupomAdicional As String, ByVal pszMensagem As String) As Integer
Public Declare Function iCFEncerrarResumido_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCFEmitirCupomAdicional_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Cancelamento de cupom fiscal
Public Declare Function iCFCancelar_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Status Cupom Fiscal
Public Declare Function rCFVerificarStatus_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cStatusCF As String, ByRef piStatusCF As Integer) As Integer
Public Declare Function rCFVerificarStatusInt_ECF_Daruma Lib "DarumaFramework.dll" (ByRef piStatusCF As Integer) As Integer
Public Declare Function rCFVerificarStatusStr_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cStatusCF As String) As Integer

'Identificar consumidor radape do Cupom fiscal
Public Declare Function iCFIdentificarConsumidor_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNome As String, ByVal pszEndereco As String, ByVal pszDoc As String) As Integer

'Cupom Mania
Public Declare Function rCMEfetuarCalculo_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszISS As String, ByVal pszICMS As String) As Integer

'Bilhete de Passagem
Public Declare Function iCFBPAbrir_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszOrigem As String, ByVal pszDestino As String, ByVal pszUFDestino As String, ByVal pszPercurso As String, ByVal pszPrestadora As String, ByVal pszPlataforma As String, ByVal pszPoltrona As String, ByVal pszModalidadetransp As String, ByVal pszCategoriaTransp As String, ByVal pszDataEmbarque As String, ByVal pszRGPassageiro As String, ByVal pszNomePassageiro As String, ByVal pszEnderecoPassageiro As String) As Integer
Public Declare Function iCFBPVender_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszAliquota As String, ByVal pszValor As String, ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String, ByVal pszDescricao As String) As Integer
Public Declare Function confCFBPProgramarUF_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszUF As String) As Integer

'Download Memórias
' binario
Public Declare Function rEfetuarDownloadMFD_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipo As String, ByVal pszInicial As String, ByVal pszFinal As String, ByVal pszNomeArquivo As String) As Integer
Public Declare Function rEfetuarDownloadMF_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNomeArquivo As String) As Integer
Public Declare Function rEfetuarDownloadTDM_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipo As String, ByVal pszInicial As String, ByVal pszFinal As String) As Integer

'Espelho MFD
Public Declare Function rGerarEspelhoMFD_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipo As String, ByVal pszInicial As String, ByVal pszFinal As String) As Integer

'Relatorios PAF-ECF
'Relatório PAF-ECF ON-line
Public Declare Function rGerarRelatorio_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszRelatorio As String, ByVal pszTipo As String, ByVal pszInicial As String, ByVal pszFinal As String) As Integer

'Relatório PAF-ECF Off-line
Public Declare Function rGerarRelatorioOffline_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszRelatorio As String, ByVal pszTipo As String, ByVal pszInicial As String, ByVal pszFinal As String, ByVal szArquivo_MF As String, ByVal szArquivo_MFD As String, ByVal szArquivo_INF As String) As Integer

'EAD PAF-ECF
Public Declare Function rAssinarRSA_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszPathArquivo As String, ByVal pszChavePrivada As String, ByVal pszAssinaturaGerada As String) As Integer


'MD5
Public Declare Function rCalcularMD5_ECF_Daruma Lib "DarumaFramework.dll" (ByRef pszPathArquivo As String, ByRef pszMD5GeradoHex As String, ByVal pszMD5GeradoAscii As String) As Integer

'Buscar GT Codificado
Public Declare Function rRetornarGTCodificado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszGT As String) As Integer

'Verifica GT Codificado
Public Declare Function rVerificarGTCodificado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszVerificaGT As String) As Integer

'Buscar Serial Codificado
Public Declare Function rRetornarNumeroSerieCodificado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszSerialCodificado As String) As Integer
'Verificar serial codificado
Public Declare Function rVerificarNumeroSerieCodificado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszSerialCriptografado As String) As Integer

'Codigo de Barras
Public Declare Function iImprimirCodigoBarras_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipo As String, ByVal pszCodigo As String, ByVal pszLargura As String, ByRef pszAltura As String, ByRef pszPosicao As String) As Integer

'--- ECF - Relatorio Gerencial - Inicio ---
Public Declare Function iRGAbrir_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNomeRG As String) As Integer
Public Declare Function iRGAbrirIndice_ECF_Daruma Lib "DarumaFramework.dll" (ByVal iIndiceRG As Integer) As Integer
Public Declare Function iRGAbrirPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iRGImprimirTexto_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTexto As String) As Integer
Public Declare Function iRGFechar_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
'--- ECF - Relatorio Gerencial - Fim ---

' --- ECF - Comprovante de CCD - Inicio ---
' Abertura de comprovante de credito e debito
Public Declare Function iCCDAbrir_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszFormaPgto As String, ByVal pszParcelas As String, ByVal pszDocOrigem As String, ByVal pszValor As String, ByVal pszCPF As String, ByVal pszNome As String, ByVal pszEndereco As String) As Integer
Public Declare Function iCCDAbrirSimplificado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszFormaPgto As String, ByVal pszParcelas As String, ByVal pszDocOrigem As String, ByVal pszValor As String) As Integer
Public Declare Function iCCDAbrirPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Impressao de texto no comprovante de credito e debito
Public Declare Function iCCDImprimirTexto_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTexto As String) As Integer
Public Declare Function iCCDImprimirArquivo_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszArqOrigem As String) As Integer
'Fechamento de texto no comprovante de credito e debito
Public Declare Function iCCDFechar_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
'Estorno de comprovante de credito e debito
Public Declare Function iCCDEstornarPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCCDEstornar_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszCOO As String, ByVal pszCPF As String, ByVal pszNome As String, ByVal pszEndereco As String) As Integer
'--- ECF - Comprovante de CCD - Fim ---

'Métodos para TEF
Public Declare Function iTEF_ImprimirResposta_ECF_Daruma Lib "DarumaFramework.dll" (ByVal szArquivo As String, ByVal bTravarTeclado As Boolean) As Integer
Public Declare Function iTEF_ImprimirRespostaCartao_ECF_Daruma Lib "DarumaFramework.dll" (ByVal szArquivo As String, ByVal bTravarTeclado As Boolean, ByVal szForma As String, ByVal szValor As String) As Integer
Public Declare Function iTEF_Fechar_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function eTEF_EsperarArquivo_ECF_Daruma Lib "DarumaFramework.dll" (ByVal szArquivo As String, ByVal iTempo As Integer, ByVal bTravar As Boolean) As Integer
Public Declare Function eTEF_TravarTeclado_ECF_Daruma Lib "DarumaFramework.dll" (ByVal bTravar As Boolean) As Integer
Public Declare Function eTEF_SetarFoco_ECF_Daruma Lib "DarumaFramework.dll" (ByVal szNomeTela As String) As Integer

'ECF - Leitura Memoria Fiscal - Inicio ---
Public Declare Function iMFLerSerial_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszInicial As String, ByVal pszFinal As String) As Integer
Public Declare Function iMFLer_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszInicial As String, ByVal pszFinal As String) As Integer
'ECF - Leitura Memoria Fiscal - Fim ---

'ECF - Comprovante não fiscal - Inicio ---
'Abertura de comprovante nao fiscal
Public Declare Function iCNFAbrir_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszCPF As String, ByVal pszNome As String, ByVal pszEndereco As String) As Integer
Public Declare Function iCNFAbrirPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Recebimento de itens
Public Declare Function iCNFReceber_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszIndice As String, ByVal pszValor As String, ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String) As Integer
Public Declare Function iCNFReceberSemDesc_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszIndice As String, ByVal pszValor As String) As Integer

'Cancelamento de item
Public Declare Function iCNFCancelarItem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszNumItem As String) As Integer
Public Declare Function iCNFCancelarUltimoItem_ECF Lib "DarumaFramework.dll" () As Integer

'Cancelamento de acrescimo em item
Public Declare Function iCNFCancelarAcrescimoItem_ECF Lib "DarumaFramework.dll" (ByVal pszNumItem As String) As Integer
Public Declare Function iCNFCancelarAcrescimoUltimoItem_ECF Lib "DarumaFramework.dll" () As Integer

'Cancelamento de desconto em item
Public Declare Function iCNFCancelarDescontoItem_ECF Lib "DarumaFramework.dll" (ByVal pszNumItem As String) As Integer
Public Declare Function iCNFCancelarDescontoUltimoItem_ECF Lib "DarumaFramework.dll" () As Integer

'Totalizacao de CNF
Public Declare Function iCNFTotalizarComprovante_ECF Lib "DarumaFramework.dll" (ByVal pszTipoDescAcresc As String, ByVal pszValorDescAcresc As String) As Integer
Public Declare Function iCNFTotalizarComprovantePadrao_ECF Lib "DarumaFramework.dll" () As Integer '

'Cancelamento de desconto e acrescimo em subtotal de CNF
Public Declare Function iCNFCancelarAcrescimoSubtotal_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCNFCancelarAcrescimoSubtotal_ECF Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCNFCancelarDescontoSubtotal_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iCNFCancelarDescontoSubtotal_ECF Lib "DarumaFramework.dll" () As Integer

'Descricao do meios de pagamento de CNF
Public Declare Function iCNFEfetuarPagamento_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszFormaPgto As String, ByVal pszValor As String, ByVal pszInfoAdicional As String) As Integer
Public Declare Function iCNFEfetuarPgtoFormatado_ECF Lib "DarumaFramework.dll" (ByVal pszFormaPgto As String, ByVal pszValor As String) As Integer
Public Declare Function iCNFEfetuarPagamentoPadrao_ECF Lib "DarumaFramework.dll" () As Integer

'Encerramento de CNF
Public Declare Function iCNFEncerrar_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszMensagem As String) As Integer
Public Declare Function iCNFEncerrarPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Cancelamento de CNF
Public Declare Function iCNFCancelar_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
'ECF - Comprovante não fiscal - Fim ---

'ECF - Funcoes Gerais - Inicio ---
Public Declare Function iEjetarCheque_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iEstornarPagamento_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszFormaPgtoEstornado As String, ByVal pszFormaPgtoEfetivado As String, ByVal pszValor As String, ByVal pszInfoAdicional As String) As Integer
Public Declare Function iAcionarGuilhotina_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszTipoCorte As String) As Integer

'Leitura X
Public Declare Function iLeituraX_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function rLeituraX_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function rLeituraXCustomizada_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszCaminho As String) As Integer

'Sangria
Public Declare Function iSangriaPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iSangria_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String, ByVal pszMensagem As String) As Integer

'Suprimento
Public Declare Function iSuprimentoPadrao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function iSuprimento_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String, ByVal pszMensagem As String) As Integer

'Reducao Z
Public Declare Function iReducaoZ_ECF_Daruma Lib "DarumaFramework.dll" (ByVal Inicial As String, ByVal Final As String) As Integer

'Programação do ECF
Public Declare Function confCadastrarPadrao_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszCadastrar As String, ByVal pszValor As String) As Integer
Public Declare Function confCadastrar_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszCadastrar As String, ByVal pszValor As String, ByVal pszSeparador As String) As Integer
Public Declare Function confHabilitarHorarioVerao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function confDesabilitarHorarioVerao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function confProgramarOperador_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String) As Integer
Public Declare Function confProgramarIDLoja_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String) As Integer
Public Declare Function confProgramarAvancoPapel_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszSepEntreLinhas As String, ByVal pszSepEntreDoc As String, ByVal pszLinhasGuilhotina As String, ByVal pszGuilhotina As String, ByVal pszImpClicheAntecipada As String) As Integer
Public Declare Function confHabilitarModoPreVenda_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function confDesabilitarModoPreVenda_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function confProgramarHorarioVerao_ECF Lib "DarumaFramework.dll" (ByVal iValor As Integer) As Integer

'Acionamento da Gaveta do ECF
Public Declare Function iAbrirGaveta_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

'Funcoes de retorno de status da impressora - Retorno
Public Declare Function rStatusImpressoraBinario_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszStatus As String) As Integer
Public Declare Function rConsultaStatusImpressoraStr_ECF_Daruma Lib "DarumaFramework.dll" (pszIndice As String, ByVal pszRetorno As String) As Integer
Public Declare Function rConsultaStatusImpressoraInt_ECF_Daruma Lib "DarumaFramework.dll" (pszIndice As Integer, ByVal pszRetorno As Integer) As Integer

'Funcoes - Retorno
Public Declare Function rVerificarImpressoraLigada_ECF_Daruma Lib "DarumaFramework.dll" () As Integer

Public Declare Function rLerAliquotas_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cAliquotas As String) As Integer
Public Declare Function rLerMeiosPagto_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszMeiosPgto As String) As Integer
Public Declare Function rLerRG_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszRelatorios As String) As Integer
Public Declare Function rLerDecimais_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszDecimalQtde As String, ByVal pszDecimalValor As String, ByRef piDecimalQtde As Integer, ByRef piDecimalValor As Integer) As Integer
Public Declare Function rLerDecimaisInt_ECF_Daruma Lib "DarumaFramework.dll" (ByRef piDecimalQtde As Integer, ByRef piDecimalValor As Integer) As Integer
Public Declare Function rLerDecimaisStr_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszDecimalQtde As String, ByVal pszDecimalValor As String) As Integer
Public Declare Function rDataHoraImpressora_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszData As String, ByVal pszHora As String) As Integer

Public Declare Function rInfoEstentida_ECF_Daruma Lib "DarumaFramework.dll" (ByRef NamelessParameter1 As Integer, ByVal NamelessParameter2 As String) As Integer
Public Declare Function rStatusImpressora_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszStatus As String, ByVal piStatusEcf As Integer) As Integer
Public Declare Function rStatusImpressoraStr_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszStatus As String) As Integer
Public Declare Function rStatusImpressoraInt_ECF_Daruma Lib "DarumaFramework.dll" (ByVal piStatusEcf As String) As Integer
Public Declare Function rInfoEstentida1_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cInfoEx As String) As Integer
Public Declare Function rInfoEstentida2_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cInfoEx As String) As Integer
Public Declare Function rInfoEstentida3_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cInfoEx As String) As Integer
Public Declare Function rInfoEstentida4_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cInfoEx As String) As Integer
Public Declare Function rInfoEstentida5_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cInfoEx As String) As Integer
Public Declare Function rVerificarReducaoZ_ECF_Daruma Lib "DarumaFramework.dll" (ByVal ZPendente As String) As Integer
Public Declare Function rStatusUltimoCmd_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszErro As String, ByVal pszAviso As String, ByRef piErro As Integer, ByRef piAviso As Integer) As Integer
Public Declare Function rStatusUltimoCmdInt_ECF_Daruma Lib "DarumaFramework.dll" (ByRef piErro As Long, ByRef piAviso As Long) As Integer
Public Declare Function rStatusUltimoCmdStr_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cErro As String, ByVal cAviso As String) As Integer
Public Declare Function rRetornarInformacao_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszIndice As String, ByVal pszRetornar As String) As Integer
Public Declare Function rRetornarNumeroSerie_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszSerialCriptografado As String, ByVal pszSerial As String) As Integer
Public Declare Function rCarregarNumeroSerie_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszSerial As String) As Integer
Public Declare Function rRetornarDadosReducaoZ_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszDados As String) As Integer
Public Declare Function rRetornarDadosReducaoZ_ECF Lib "DarumaFramework.dll" (ByRef pszDados As String) As Integer
Public Declare Function rRegistrarNumeroSerie_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
'ECF - Funcoes Gerais - Fim ---

'ECF - Especiais - Inicio ---
Public Declare Function eAguardarCompactacao_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function eBuscarPortaVelocidade_ECF Lib "DarumaFramework.dll" () As Integer
Public Declare Function eEnviarComando_ECF_Daruma Lib "DarumaFramework.dll" (ByVal cComando As String, ByVal iTamanhoComando As Integer, ByVal iType As Integer) As Integer
Public Declare Function eRetornarAviso_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function eRetornarErro_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
'ECF - Especiais - Fim ---

'ECF - Registro - Inicio ---
Public Declare Function regCCDDocOrigem_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCCDFormaPgto_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCCDLinhasTEF_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCCDParcelas_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCCDValor_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFFormaPgto_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFMensagemPromocional_ECF Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFQuantidade_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFTamanhoMinimoDescricao_ECF Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFTipoDescAcresc_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFUnidadeMedida_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFValorDescAcresc_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFCupomAdicional_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFCupomAdicionalDllConfig_ECF Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFCupomAdicionalDllTitulo_ECF Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regChequeXLinha1_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regChequeXLinha2_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regChequeXLinha3_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regChequeYLinha1_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regChequeYLinha2_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regChequeYLinha3_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCompatStatusFuncao_ECF Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regMaxFechamentoAutomatico_ECF Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFCupomAdicionalDLLConfig_ECF_Daruma Lib "DarumaFramework.dll" Alias "regCFCupomAdicionalDllConfig_ECF_Daruma" (ByVal pszParametro As String) As Integer
Public Declare Function regCFCupomMania Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFMensagemPromocional_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFAguardarImpressao_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regCFTamanhoMinimoDescricao_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFArquivoLeituraX_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFCaracterSeparador_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFAuditoria_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFReceberAvisoEmArquivo_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFReceberInfoEstendida_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFMaxFechamentoAutomatico_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regECFReceberErroEmArquivo_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function regAtocotepe_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro1 As String, ByVal pszParametro2 As String) As Integer
Public Declare Function regSintegra_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro1 As String, ByVal pszParametro2 As String) As Integer
Public Declare Function regECFReceberInfoEstendidaEmArquivo_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
Public Declare Function eDefinirModoRegistro_Daruma Lib "DarumaFramework.dll" (ByVal pszParametro As String) As Integer
'ECF - Registro - Fim ---'

'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                           IMPRESSORAS DUAL                            ==========='
Public Declare Function iEnviarBMP_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stArqOrigem As String) As Integer
Public Declare Function iAcionarGaveta_DUAL_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function iImprimirArquivo_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stPath As String) As Integer
Public Declare Function rStatusGaveta_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByRef iStatusGaveta As Integer) As Integer
Public Declare Function rStatusDocumento_DUAL_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function rStatusImpressora_DUAL_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function regVelocidade_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regTermica_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regTabulacao_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regPortaComunicacao_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regModoGaveta_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regLinhasGuilhotina_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regEnterFinal_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regAguardarProcesso_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function iImprimirTexto_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stTexto As String, ByVal iTam As Integer) As Integer
Public Declare Function iAutenticarDocumento_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stTexto As String, ByVal StLocal As String, ByVal stTimeOut As String) As Integer
Public Declare Function regCodePageAutomatico_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regZeroCortado_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer

'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                               TA2000                                  ==========='
Public Declare Function iEnviarDadosFormatados_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal szTexto As String, ByVal szRetorno As String) As Integer
Public Declare Function regPorta_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regAuditoria_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMensagemBoasVindasLinha1_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMensagemBoasVindasLinha2_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMarcadorOpcao_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMascara_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMascaraLetra_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMascaraNumero_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
Public Declare Function regMascaraEco_TA2000_Daruma Lib "DarumaFramework.dll" (ByVal stParametro As String) As Integer
                    
'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                               MIN-200                                 ==========='

Public Declare Function regLerApagar_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function regPorta_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function regThread_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function regVelocidade_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function regTempoAlertar_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function regCaptionWinAPP_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function regBandejaInicio_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer

Public Declare Function eInicializar_MODEM_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function eTrocarBandeja_MODEM_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function eApagarSms_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal iNumeroSMS As Integer) As Integer

Public Declare Function rListarSms_MODEM_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function rNivelSinalRecebido_MODEM_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function rReceberSms_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sIndiceSMS As String, ByVal sNumFone As String, ByVal sData As String, ByVal sHora As String, ByVal sMsg As String) As Integer
Public Declare Function rRetornarImei_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sIMEI As String) As Integer
Public Declare Function rRetornarOperadora_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sOperadora As String) As Integer
Public Declare Function tEnviarSms_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sNumeroTelefone As String, ByVal sMensagem As String) As Integer

Public Declare Function tEnviarDadosCsd_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function rReceberDadosCsd_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer
Public Declare Function eAtivarConexaoCsd_MODEM_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function eFinalizarChamadaCsd_MODEM_DarumaFramework Lib "DarumaFramework.dll" () As Integer
Public Declare Function eRealizarChamadaCsd_MODEM_DarumaFramework Lib "DarumaFramework.dll" (ByVal sParametro As String) As Integer



'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                          DARUMAFRAMEWORK                              ==========='

Public Declare Function eVerificarVersaoDLL_Daruma Lib "DarumaFramework.dll" (ByVal sVersaoDLL As String) As Integer
Public Declare Function eDefinirProduto_Daruma Lib "DarumaFramework.dll" (ByVal sProduto As String) As Integer
Public Declare Function regRetornaValorChave_DarumaFramework Lib "DarumaFramework.dll" (ByVal sProduto As String, sChave As String, ByVal sValor As String) As Integer

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public Declare Function eBuscarPortaVelocidade_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function eAcionarGuilhotina_ECF_Daruma Lib "DarumaFramework.dll" (ByVal sTipoCorte As String) As Integer
Public Declare Function eAbrirGaveta_ECF_Daruma Lib "DarumaFramework.dll" () As Integer
Public Declare Function regAlterarValor_Daruma Lib "DarumaFramework.dll" (ByVal sProduto_Chave As String, ByVal sValor As String) As Integer
Public Declare Function eInterpretarErro_ECF_Daruma Lib "DarumaFramework.dll" (ByVal iErro As Integer, ByVal sMsg_Erro As String) As Integer
Public Declare Function eInterpretarAviso_ECF_Daruma Lib "DarumaFramework.dll" (ByVal iAviso As Integer, ByVal sMsg_Aviso As String) As Integer


Public Declare Function rCodigoModeloFiscal_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszValor As String) As Integer
Public Declare Function eMemoriaFiscal_ECF_Daruma Lib "DarumaFramework.dll" (ByVal pszInicial As String, ByVal pszFinal As String, ByVal pszCompleta As Integer, ByVal pszTipo As String) As Integer



'Declaracoes globais'===========                         VARIAVEIS GLOBAIS                             ============'

   'Public iRetorno As Integer
   Public Int_NumErro As Long
   Public Int_NumAviso As Long
   Public Str_Msg_Retorno_Metodo As String

'================================    FUN��ES GLOBAIS    ================================'
          '===========            TRATAMENTO DE RETORNO IMPRESSORA FISCAL              ==========='
Public Function DarumaFramework_Mostrar_Retorno_ECF() As Boolean
On Error Resume Next

   Dim int_Ret As Integer
   Dim Str_Msg_NumErro As String
   Dim Str_Msg_NumAviso As String
   
   Str_Msg_NumErro = Space(50) '
   Str_Msg_NumAviso = Space(50)
   
   int_Ret = 0
   Int_NumErro = 0
   Int_NumAviso = 0

   'Retornos de M�todo
   If CStr(Retorno) <> "1" Then

       Select Case Retorno
           Case "0"
               Str_Msg_Retorno_Metodo = "[0] - Erro durante a execu��o"
           Case "-1"
               Str_Msg_Retorno_Metodo = "[-1] - Erro do M�todo"
           Case "-2"
               Str_Msg_Retorno_Metodo = "[-2] - Par�metro incorreto"
           Case "-3"
               Str_Msg_Retorno_Metodo = "[-3] - Al�quota (Situa��o tribut�ria) n�o programada"
           Case "-4"
               Str_Msg_Retorno_Metodo = "[-4] - Chave do Registry n�o encontrada"
           Case "-5"
               Str_Msg_Retorno_Metodo = "[-5] - Erro ao Abrir a porta de Comunica��o"
           Case "-6"
               Str_Msg_Retorno_Metodo = "[-6] - Impressora Desligada"
           Case "-7"
               Str_Msg_Retorno_Metodo = "[-7] - Erro no N�mero do Banco"
           Case "-8"
               Str_Msg_Retorno_Metodo = "[-8] - Erro ao Gravar as informa��es no arquivo de Status ou de Retorno de Info"
           Case "-9"
               Str_Msg_Retorno_Metodo = "[-9] - Erro ao Fechar a porta de Comunica��oo"
           Case "10"
               Str_Msg_Retorno_Metodo = "[10] - Se o ECF n�o tem forma de pagamento e n�o permite cadastrar esta forma"
           Case "24"
               Str_Msg_Retorno_Metodo = "[24] - Forma de Pagamento n�o Programada"
           Case "25"
               Str_Msg_Retorno_Metodo = "[25] - Totalizador nao ECF N�o Vinculado nao Programado"
           Case "27"
               Str_Msg_Retorno_Metodo = "[27] - Foi Detectado Erro ou Warning na Impressora"
           Case "28"
               Str_Msg_Retorno_Metodo = "[28] - Time-Out"
           Case "40"
               Str_Msg_Retorno_Metodo = "[40] - Tag XML Inv�lida"
           Case "50"
               Str_Msg_Retorno_Metodo = "[50] - Problemas ao Criar Chave no Registry"
           Case "51"
               Str_Msg_Retorno_Metodo = "[51] - Erro ao Gravar LOG"
           Case "52"
               Str_Msg_Retorno_Metodo = "[52] - Erro ao abrir arquivo"
           Case "53"
               Str_Msg_Retorno_Metodo = "[53] - Fim de arquivo"
           Case "60"
               Str_Msg_Retorno_Metodo = "[60] - Erro na tag de formatacao DHTML"
           Case "90"
               Str_Msg_Retorno_Metodo = "[90] - Erro Configurar a Porta de Comunica��o"
           Case "99"
               Str_Msg_Retorno_Metodo = "[99] - Par�metro inv�lido ou ponteiro nulo de par�metro"
           Case Else ' Se o Retorno ok
               Str_Msg_Retorno_Metodo = "[" + CStr(Retorno) + "] - Retorno Desconhecido!"
       End Select

       Sleep 1000
       'Verifico o status do ultimo comando e mostro o numero de erro e numero de aviso
       DoEvents
       If Not IsNumeric(Int_NumErro) Then
           Int_NumErro = 0
       End If
       
       If Not IsNumeric(Int_NumAviso) Then
           Int_NumAviso = 0
       End If
       
       int_Ret = rStatusUltimoCmdInt_ECF_Daruma(Int_NumErro, Int_NumAviso)
       DoEvents
       
       'Pe�o a interpreta��o do numero de erro e mostro na tela
       If (Int_NumErro <> 0) Then
       
          Select Case Int_NumErro

               Case "1"
                   Str_Msg_NumErro = "[1] - ECF com falha mec�nica"
               Case " 2"
                   Str_Msg_NumErro = "[2] - MF n�o conectada"
               Case "3"
                   Str_Msg_NumErro = "[3] - MFD n�o conectada"
               Case "4"
                   Str_Msg_NumErro = "[4] - MFD esgotada"
               Case "5"
                   Str_Msg_NumErro = "[5] - Erro na comunica��o com a MF"
               Case "6"
                   Str_Msg_NumErro = "[6] - Erro na comunica��o com a MFD"
               Case "7"
                   Str_Msg_NumErro = "[7] - MF n�o inicializada"
               Case "8"
                   Str_Msg_NumErro = "[8] - MFD n�o inicializada"
               Case "9"
                   Str_Msg_NumErro = "[9] - MFD j� inicializada"
               Case "10"
                   Str_Msg_NumErro = "[10] - MFD foi substitu�da"
               Case "11"
                   Str_Msg_NumErro = "[11] -  MFD j� cadastrada"
               Case "12"
                   Str_Msg_NumErro = "[12] -  Erro na inicializa��o da MFD"
               Case "13"
                   Str_Msg_NumErro = "[13] -  Faltam par�metros de inicializa��o na MF"
               Case "14"
                   Str_Msg_NumErro = "[14] -  Comando n�o suportado"
               Case "15"
                   Str_Msg_NumErro = "[15] -  Superaquecimento da cabe�a de impress�o"
               Case "16"
                   Str_Msg_NumErro = "[16] -  Perda de dados da MT"
               Case "17"
                   Str_Msg_NumErro = "[17] -  Opera��o habilitada apenas em MIT"
               Case "18"
                   Str_Msg_NumErro = "[18] -  Opera��o habilitada apenas em modo fiscal"
               Case "19"
                   Str_Msg_NumErro = "[19] -  Data inexistente "
               Case "20"
                   Str_Msg_NumErro = "[20] -  Data inferior ao do �ltimo documento"
               Case "21"
                   Str_Msg_NumErro = "[21] - Intervalo inconsistente"
               Case "22"
                   Str_Msg_NumErro = "[22] - N�o existem dados"
               Case "23"
                   Str_Msg_NumErro = "[23] - Clich� de formato inv�lido"
               Case "24"
                   Str_Msg_NumErro = "[24] - Erro no verificador da comunica��o"
               Case "25"
                   Str_Msg_NumErro = "[25] - Senha incorreta"
               Case "26"
                   Str_Msg_NumErro = "[26] - N�mero de decimais para quantidade inv�lido"
               Case "27"
                   Str_Msg_NumErro = "[27] - N�mero de decimais para valor unit�rio inv�lido"
               Case "28"
                   Str_Msg_NumErro = "[28] - Tipo de impress�o de FD inv�lido"
               Case "29"
                   Str_Msg_NumErro = "[29] - Caracter n�o estamp�vel"
               Case "30"
                   Str_Msg_NumErro = "[30] - Caracter n�o estamp�vel ou em branco"
               Case "31"
                   Str_Msg_NumErro = "[31] - Caracteres n�o podem ser repetidos"
               Case "32"
                   Str_Msg_NumErro = "[32] - Limite de itens atingido"
               Case "33"
                   Str_Msg_NumErro = "[33] - Todos os totalizadores fiscais j� est�o programados"
               Case "34"
                   Str_Msg_NumErro = "[34] - Totalizador fiscal j� programado"
               Case "35"
                   Str_Msg_NumErro = "[35] - Todos os totalizadores n�o fiscais j� est�o programados"
               Case "36"
                   Str_Msg_NumErro = "[36] - Totalizador n�o fiscal j� programado"
               Case "37"
                   Str_Msg_NumErro = "[37] - Todos os relat�rios gerenciais j� est�o programados"
               Case "38"
                   Str_Msg_NumErro = "[38] - Relat�rio gerencial j� programado"
               Case "39"
                   Str_Msg_NumErro = "[39] - Meio de pagamento j� programado"
               Case "40"
                   Str_Msg_NumErro = "[40] - �ndice inv�lido"
               Case "41"
                   Str_Msg_NumErro = "[41] - �ndice do meio de pagamento inv�lido"
               Case "42"
                   Str_Msg_NumErro = "[42] - Erro gravando n�mero de decimais na MF"
               Case "43"
                   Str_Msg_NumErro = "[43] - Erro gravando moeda na MF"
               Case "44"
                   Str_Msg_NumErro = "[44] - Erro gravando s�mbolos de decodifica��o do GT na MF"
               Case "45"
                   Str_Msg_NumErro = "[45] - Erro gravando n�mero de fabrica��o da MFD na MF"
               Case "46"
                   Str_Msg_NumErro = "[46] - Erro gravando usu�rio na MF"
               Case "47"
                   Str_Msg_NumErro = "[47] - Erro gravando GT do usu�rio anterior na MF"
               Case "48"
                   Str_Msg_NumErro = "[48] - Erro gravando registro de marca��o na MF"
               Case "49"
                   Str_Msg_NumErro = "[49] - Erro gravando CRO na MF"
               Case "50"
                   Str_Msg_NumErro = "[50] - Erro gravando impress�o de FD na MF"
               Case "51"
                   Str_Msg_NumErro = "[51] - Campo em branco ou zero n�o permitido"
               Case "52"
                   Str_Msg_NumErro = "[52] - Campo reservado a grava��o da moeda na MF esgotado"
               Case "53"
                   Str_Msg_NumErro = "[53] - Campo reservado a grava��o da tabela de GT na MF esgotado"
               Case "54"
                   Str_Msg_NumErro = "[54] - Campo reservado a grava��o do NS da MFD na MF esgotado"
               Case "55"
                   Str_Msg_NumErro = "[55] - Campo reservado a grava��o de usu�rio na MF esgotado"
               Case "56"
                   Str_Msg_NumErro = "[56] - CNPJ inv�lido"
               Case "57"
                   Str_Msg_NumErro = "[57] - CRZ e CRO em zero"
               Case "58"
                   Str_Msg_NumErro = "[58] - Intervalo invertido"
               Case "59"
                   Str_Msg_NumErro = "[59] - Utilize apenas 0 ou 1"
               Case "60"
                   Str_Msg_NumErro = "[60] - Configura��o permitida apenas imediatamente a RZ"
               Case "61"
                   Str_Msg_NumErro = "[61] - S�mbolo gr�fico inv�lido"
               Case "62"
                   Str_Msg_NumErro = "[62] - Falta pelo menos 1 campo no nome da moeda para cheque"
               Case "63"
                   Str_Msg_NumErro = "[63] - C�digo supera o valor 255"
               Case "64"
                   Str_Msg_NumErro = "[64] - Utilize valores entre 25 e 80"
               Case "65"
                   Str_Msg_NumErro = "[65] - Utilize valores entre 1 e 15"
               Case "66"
                   Str_Msg_NumErro = "[66] - Utilize valores entre 0 e 7250"
               Case "67"
                   Str_Msg_NumErro = "[67] - Data informada n�o coincide com a data do ECF"
               Case "68"
                   Str_Msg_NumErro = "[68] - Deve ajustar o rel�gio"
               Case "69"
                   Str_Msg_NumErro = "[69] - Erro ao ajustar o rel�gio"
               Case "70"
                   Str_Msg_NumErro = "[70] - Capacidade da MF esgotada"
               Case "71"
                   Str_Msg_NumErro = "[71] - Vers�o do SB gravado na MF incorreta"
               Case "72"
                   Str_Msg_NumErro = "[72] - Fim do papel"
               Case "73"
                   Str_Msg_NumErro = "[73] - Nenhum usu�rio programado"
               Case "74"
                   Str_Msg_NumErro = "[74] - Utilize apenas d�gitos num�ricos"
               Case "75"
                   Str_Msg_NumErro = "[75] - Campo n�o pode estar em zero"
               Case "76"
                   Str_Msg_NumErro = "[76] - Campo n�o pode estar em branco"
               Case "77"
                   Str_Msg_NumErro = "[77] - Valor da opera��o n�o pode ser zero"
               Case "78"
                   Str_Msg_NumErro = "[78] - CF aberto"
               Case "79"
                   Str_Msg_NumErro = "[79] - CNF aberto"
               Case "80"
                   Str_Msg_NumErro = "[80] - CCD aberto"
               Case "81"
                   Str_Msg_NumErro = "[81] - RG aberto"
               Case "82"
                   Str_Msg_NumErro = "[82] - CF n�o aberto"
               Case "83"
                   Str_Msg_NumErro = "[83] - CNF n�o aberto"
               Case "84"
                   Str_Msg_NumErro = "[84] - CCD n�o aberto"
               Case "85"
                   Str_Msg_NumErro = "[85] - RG n�o aberto"
               Case "86"
                   Str_Msg_NumErro = "[86] - CCD ou RG n�o aberto"
               Case "87"
                   Str_Msg_NumErro = "[87] - Documento j� totalizado"
               Case "88"
                   Str_Msg_NumErro = "[88] - RZ do movimento anterior pendente"
               Case "89"
                   Str_Msg_NumErro = "[89] - J� emitiu RZ de hoje"
               Case "90"
                   Str_Msg_NumErro = "[90] - Totalizador sem al�quota programada"
               Case "91"
                   Str_Msg_NumErro = "[91] - Campo de c�digo ausente"
               Case "92"
                   Str_Msg_NumErro = "[92] - Campo de descri��o ausente"
               Case "93"
                   Str_Msg_NumErro = "[93] - VU ou quantidade em zero"
               Case "94"
                   Str_Msg_NumErro = "[94] - Item ainda n�o vendido"
               Case "95"
                   Str_Msg_NumErro = "[95] - Desconto ou acr�scimo n�o pode ser zero"
               Case "96"
                   Str_Msg_NumErro = "[96] - Item j� possui desconto ou acr�scimo"
               Case "97"
                   Str_Msg_NumErro = "[97] - Item cancelado"
               Case "98"
                   Str_Msg_NumErro = "[98] - Opera��o inibida por configura��o"
               Case "99"
                   Str_Msg_NumErro = "[99] - Op��o n�o suportada"
               Case "100"
                   Str_Msg_NumErro = "[100] -  Desconto ou acr�scimo supera valor bruto"
               Case "101"
                   Str_Msg_NumErro = "[101] -  Desconto ou acr�scimo final de valor zero"
               Case "102"
                   Str_Msg_NumErro = "[102] -  Valor bruto zero"
               Case "103"
                   Str_Msg_NumErro = "[103] -  Overflow no valor do item"
               Case "104"
                   Str_Msg_NumErro = "[104] -  Overflou no valor do desconto ou acr�scimo"
               Case "105"
                   Str_Msg_NumErro = "[105] -  Overflow na capacidade do documento"
               Case "106"
                   Str_Msg_NumErro = "[106] -  Overflow na capacidade do totalizador"
               Case "107"
                   Str_Msg_NumErro = "[107] -  Item n�o possui desconto"
               Case "108"
                   Str_Msg_NumErro = "[108] -  Item j� possui desconto"
               Case "109"
                   Str_Msg_NumErro = "[109] -  Quantidade possui mais de 2 decimais"
               Case "110"
                   Str_Msg_NumErro = "[110] -  Valor unit�rio possui mais de 2 decimais"
               Case "111"
                   Str_Msg_NumErro = "[111] -  Quantidade a cancelar deve ser inferior a total"
               Case "112"
                   Str_Msg_NumErro = "[112] -  Campo de descri��o deste item n�o mais presente na MT"
               Case "113"
                   Str_Msg_NumErro = "[113] -  Subtotal n�o possui desconto ou acr�scimo"
               Case "114"
                   Str_Msg_NumErro = "[114] -  N�o em fase de totaliza��o"
               Case "115"
                   Str_Msg_NumErro = "[115] -  N�o em fase de venda ou totaliza��o"
               Case "116"
                   Str_Msg_NumErro = "[116] -  Mais de 1 desconto ou acr�scimo n�o permitido"
               Case "117"
                   Str_Msg_NumErro = "[117] -  Valor do desconto ou acr�scimo supera subtotal"
               Case "118"
                   Str_Msg_NumErro = "[118] -  Meio de pagamento n�o programado"
               Case "119"
                   Str_Msg_NumErro = "[119] -  N�o em fase de pagamento ou totaliza��o"
               Case "120"
                   Str_Msg_NumErro = "[120] -  N�o em fase de finaliza��o de documento"
               Case "121"
                   Str_Msg_NumErro = "[121] -  J� emitiu mais CCDs que poderia estornar"
               Case "122"
                   Str_Msg_NumErro = "[122] -  �ltimo documento n�o � cancel�vel"
               Case "123"
                   Str_Msg_NumErro = "[123] -  Estorne CCDs"
               Case "124"
                   Str_Msg_NumErro = "[124] -  �ltimo documento n�o foi CF"
               Case "125"
                   Str_Msg_NumErro = "[125] -  �ltimo documento n�o foi CNF"
               Case "126"
                   Str_Msg_NumErro = "[126] -  N�o pode cancelar"
               Case "127"
                   Str_Msg_NumErro = "[127] -  Pagamento n�o mais na MT"
               Case "128"
                   Str_Msg_NumErro = "[128] -  J� emitiu CCD deste pagamento"
               Case "129"
                   Str_Msg_NumErro = "[129] -  RG n�o programado"
               Case "130"
                   Str_Msg_NumErro = "[130] -  CNF n�o programado"
               Case "131"
                   Str_Msg_NumErro = "[131] -  C�pia n�o dispon�vel"
               Case "132"
                   Str_Msg_NumErro = "[132] -  J� emitiu segunda via"
               Case "133"
                   Str_Msg_NumErro = "[133] -  J� emitiu reimpress�o"
               Case "134"
                   Str_Msg_NumErro = "[134] -  Informa��es sobre o pagamento n�o dispon�veis"
               Case "135"
                   Str_Msg_NumErro = "[135] -  J� emitiu todas as parcelas"
               Case "136"
                   Str_Msg_NumErro = "[136] -  Parcelamento somente na sequ�ncia"
               Case "137"
                   Str_Msg_NumErro = "[137] -  CCD n�o encontrado"
               Case "138"
                   Str_Msg_NumErro = "[138] -  N�o pode utilizar SANGRIA ou SUPRIMENTO"
               Case "139"
                   Str_Msg_NumErro = "[139] -  Pagamento n�o admite CCD"
               Case "140"
                   Str_Msg_NumErro = "[140] -  Rel�gio inoperante"
               Case "141"
                   Str_Msg_NumErro = "[141] -  Usu�rio sem CNPJ"
               Case "142"
                   Str_Msg_NumErro = "[142] -  Usu�rio sem IM"
               Case "143"
                   Str_Msg_NumErro = "[143] -  N�o se passou 1 hora ap�s o fechamento do �ltimo documento"
               Case "144"
                   Str_Msg_NumErro = "[144] -  ECF OFF LINE"
               Case "145"
                   Str_Msg_NumErro = "[145] -  Documento em emiss�o"
               Case "146"
                   Str_Msg_NumErro = "[146] -  COO n�o coincide"
               Case "147"
                   Str_Msg_NumErro = "[147] -  Erro na autentica��o"
               Case "148"
                   Str_Msg_NumErro = "[148] -  Erro na impress�o de cheque"
               Case "149"
                   Str_Msg_NumErro = "[149] -  Data n�o pertence ao s�culo XXI"
               Case "150"
                   Str_Msg_NumErro = "[150] -  Usu�rio j� programado"
               Case "151"
                   Str_Msg_NumErro = "[151] -  Descri��o do pagamento j� utilizada"
               Case "152"
                   Str_Msg_NumErro = "[152] -  Descri��o do totalizador j� utilizada"
               Case "153"
                   Str_Msg_NumErro = "[153] -  Descri��o do RG j� utilizada"
               Case "154"
                   Str_Msg_NumErro = "[154] -  J� tem desconto ap�s acr�scimo ( ou vice versa )"
               Case "155"
                   Str_Msg_NumErro = "[155] -  J� programou 15 totalizadores para ICMS"
               Case "156"
                   Str_Msg_NumErro = "[156] -  J� programou 15 totalizadores para ISS"
               Case "157"
                   Str_Msg_NumErro = "[157] -  MFD com problemas"
               Case "158"
                   Str_Msg_NumErro = "[158] -  Raz�o social excede 48 caracteres"
               Case "159"
                   Str_Msg_NumErro = "[159] -  Nome fantasia excede 48 caracteres"
               Case "160"
                   Str_Msg_NumErro = "[160] -  Endere�o excede 120 caracteres"
               Case "161"
                   Str_Msg_NumErro = "[161] -  Identifica��o do programa aplicativo ausente"
               Case "162"
                   Str_Msg_NumErro = "[162] -  Valor de desconto supera valor acumulado em totalizador"
               Case "163"
                   Str_Msg_NumErro = "[163] -  N�mero de parcelas no pagamento n�o pode exceder 24"
               Case "164"
                   Str_Msg_NumErro = "[164] -  MFD n�o cadastrada"
               Case "165"
                   Str_Msg_NumErro = "[165] -  Excedeu limite de impress�o de FD ( capacidade na MF esgotada )"
               Case "166"
                   Str_Msg_NumErro = "[166] -  Efetivado � igual ao estornado"
               Case "167"
                   Str_Msg_NumErro = "[167] -  S�mbolo da moeda j� programado"
               Case "168"
                   Str_Msg_NumErro = "[168] -  UF inv�lida"
               Case "169"
                   Str_Msg_NumErro = "[169] -  UF j� programada"
               Case "170"
                   Str_Msg_NumErro = "[170] -  Erro gravando UF"
               Case "171"
                   Str_Msg_NumErro = "[171] - Leitor CMC-7 n�o instalado"
               Case "172"
                   Str_Msg_NumErro = "[172] -  Erro de leitura do c�digo CMC-7"
               Case "173"
                   Str_Msg_NumErro = "[173] -  Autentica��o n�o permitida"
               Case "174 "
                   Str_Msg_NumErro = "[174] -  Opera��o somente com mecanismo matricial de impacto"
               Case "175"
                   Str_Msg_NumErro = "[175] -  Coordenadas de cheque inv�lidas"
               Case "176"
                   Str_Msg_NumErro = "[176] -  Impress�o do verso do cheque somente ap�s a impress�o da frente"
               Case "177"
                   Str_Msg_NumErro = "[177] -  Indice do bitmap inv�lido"
               Case "178"
                   Str_Msg_NumErro = "[178] -  Bitmap de tamanho inv�lido"
               Case "179"
                   Str_Msg_NumErro = "[179] -  �ltima RZ a mais de 30 dias. Comando de RZ deve informar data correta"
               Case "184"
                   Str_Msg_NumErro = "[184] -  Par�metro s� pode ser �A� ou �T�"
               Case "185"
                   Str_Msg_NumErro = "[185] -  Falta unidade doproduto"
               Case "186"
                   Str_Msg_NumErro = "[186] -  Velocidade n�o permitida"
               Case "187"
                   Str_Msg_NumErro = "[187] -  C�digo repetido"
               Case "188"
                   Str_Msg_NumErro = "[188] -  Fora dos limites"
               Case "189"
                   Str_Msg_NumErro = "[189] -  J� identificou o consumidor"
               Case "190"
                   Str_Msg_NumErro = "[190] -  N�mero de Fabrica��o incorreto"
               Case "191"
                   Str_Msg_NumErro = "[191] -  Informa��o dispon�vel n�o corresponde a MF informada"
               Case "192"
                   Str_Msg_NumErro = "[192] -  MF j� em uso"
               Case "193"
                   Str_Msg_NumErro = "[193] -  Falha n�o recuper�vel durante a opera��o"
               Case "194"
                   Str_Msg_NumErro = "[194] -  Op��o inv�lida"
               Case "195"
                   Str_Msg_NumErro = "[195] -  Par�metros inv�lidos"
               Case "196"
                   Str_Msg_NumErro = "[196] -  Caracter HEXA inv�lido"
               Case "197"
                   Str_Msg_NumErro = "[197] -  Valor insuficiente de pagamento"
               Case "198"
                   Str_Msg_NumErro = "[198] -  IE inv�lido"
               Case "199"
                   Str_Msg_NumErro = "[199] -  IM inv�lido"
               Case "301"
                   Str_Msg_NumErro = "[301] -  CFBP Inibido"
               Case "302"
                   Str_Msg_NumErro = "[302] -  Modalidade de Transporte inv�lida"
               Case "303"
                   Str_Msg_NumErro = "[303] -  Categoria de Transporte inv�lida"
               Case "304"
                   Str_Msg_NumErro = "[304] -  UF incompat�vel"
               Case "305"
                   Str_Msg_NumErro = "[305] -  Comando dispon�vel apenas em CF gen�rico"
               Case "400"
                   Str_Msg_NumErro = "[400] -  Chave n�o carregada"
               Case "401"
                   Str_Msg_NumErro = "[401] -  Chave inv�lida"
               Case "402"
                   Str_Msg_NumErro = "[402] -  Erro na decodifica��o"
               Case "403"
                   Str_Msg_NumErro = "[403] -  Erro na codifica��o"
               Case Else ' Se o NumErro desconhecido
                   Str_Msg_NumErro = "[" + Str_Msg_NumErro + "] - Erro Desconhecido!"
           End Select
      
       Else
       'Se Erro = 0
           Str_Msg_NumErro = "[0] - Sem Erro"
       End If
   
      'solicito a interpreta��o do numero de aviso e mostro na tela
       If (Int_NumAviso <> 0) Then
       
           Select Case Int_NumAviso
           
               Case "1"
                   Str_Msg_NumAviso = "[1] - Papel Acabando"
               Case "2"
                    Str_Msg_NumAviso = "[2] - Tampa aberta"
               Case "3"
                   Str_Msg_NumAviso = "[4] - Bateria fraca"
               Case "4"
                   Str_Msg_NumAviso = "[40] - Compactando"
               Case Else ' Se o NumAviso desconhecido
                   Str_Msg_NumAviso = "[" + Str_Msg_NumAviso + "] - Aviso Desconhecido!"
           End Select
       Else
       'Se Aviso = 0 (ok)
           Str_Msg_NumAviso = "[0] - Sem Aviso"
       End If

       Dim strMenssagem As String
   If Len(Trim(Str_Msg_Retorno_Metodo)) + Len(Trim(Str_Msg_NumErro)) + Len(Trim(Str_Msg_NumAviso)) > 0 Then

      strMenssagem = "Retorno do Metodo = " + caracteres_validos(Str_Msg_Retorno_Metodo) & Chr(13) & Chr(10)
      strMenssagem = strMenssagem & "Num.Erro = " + caracteres_validos(Str_Msg_NumErro) & Chr(13) & Chr(10)
      strMenssagem = strMenssagem & "Num.Aviso = " + caracteres_validos(Str_Msg_NumAviso) & Chr(13) & Chr(10)

      MsgBox strMenssagem, vbCritical, "Mostrar Retorno ECF"

      DarumaFramework_Mostrar_Retorno_ECF = False
      Else: DarumaFramework_Mostrar_Retorno_ECF = True
   End If
       'frmMostraErroImpressora.lblRetorno.Caption = "Retorno do Metodo = " + Str_Msg_Retorno_Metodo
       'frmMostraErroImpressora.lblErro.Caption = "Num.Erro = " + Str_Msg_NumErro
       'frmMostraErroImpressora.lblAviso.Caption = "Num.Aviso = " + Str_Msg_NumAviso
               
       'frmMostraErroImpressora.Show 1
       
   Else 'SE RETORNO = 1 (OK)

       'frmMostraErroImpressora.Show
       'frmMostraErroImpressora.lblRetorno.Caption = "Retorno do Metodo = [1] - Opera��o realizada com sucesso"
       'frmMostraErroImpressora.lblErro.Caption = "Num.Erro = " + "[0] Sem Erros"
       'frmMostraErroImpressora.lblAviso.Caption = "Num.Aviso = " + "[0] Sem Avisos"
       
       'strMenssagem = "Retorno do Metodo = " + Str_Msg_Retorno_Metodo & Chr(13) & Chr(10)
       'strMenssagem = strMenssagem & "Num.Erro = " + Str_Msg_NumErro & Chr(13) & Chr(10)
       'strMenssagem = strMenssagem & "Num.Aviso = " + Str_Msg_NumAviso & Chr(13) & Chr(10)
       DarumaFramework_Mostrar_Retorno_ECF = True
   End If
End Function

'Declaracoes globais

' Funcoes globais
'Funcao que trata os retornos das Impressora Fiscais Varejo
Public Function VerificaRetornoImpressoraDaruma(Label As String, RetornoFuncao As String, TituloJanela As String)
'On Error GoTo ERRO_TRATA

    Dim Str_ErroExtendido As String
    Dim RetornaMensagem As Integer
    Dim iST1 As Integer, iST2 As Integer
    
    Int_Ack = 0
    Int_St1 = 0
    Int_St2 = 0
    
    Select Case intRetorno

        Case 0
            'MsgBox "Erro de comunica��o com a impressora.", vbOKOnly + vbCritical, TituloJanela
            'Retorno = Daruma_FI_AbrePortaSerial
            'Retorno = Daruma_FI_FechaPortaSerial
            GoTo entra
        Case 1
entra:
            RetornoStatus = Daruma_FI_RetornoImpressora(Ack, Int_St1, Int_St2)
            ValorRetorno = Str(Int_Ack) & "," & Str(Int_St1) & "," & Str(Int_St2)
            iST1 = Int_St1
            iST2 = Int_St2
            
            If Label <> "" And RetornoFuncao <> "" Then RetornaMensagem = 1

            If Ack = 21 Then
                MsgBox "Status da Impressora: 21" & vbCr & vbLf & "Comando n�o executado", vbOKOnly + vbInformation, TituloJanela
                Exit Function
            End If
            
            If ID_Cupom = 0 Then 'se for 0 � porque acabou de abrir o sistema e tem cupom em aberto
                'UltimoCupomImpresso
                If ID_Cupom = 0 Then 'se continuar 0 � algum erro ent�o dar mensagem e sair
                    'MsgBox "Cancelamento n�o permitido"
                    'Exit Function
                End If
            End If
            
            StringRetorno = ""
            If (Int_St1 <> 0 Or Int_St2 <> 0) Then
                If (Int_St1 >= 128) Then
                    StringRetorno = "Fim de Papel" & vbCr
                    
                    If Int_St1 >= 128 Then 'fim do papel
                        If MsgBox("-------------------------FIM DO PAPEL-------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "IMPRESSORA N�O RESPONDE. TENTAR NOVAMENTE?", vbYesNo + vbExclamation, TituloJanela) = vbNo Then
                            If CaixaCupomAberto = True Then
                                'CancelandoUltimoCupom
                            End If
                            GoTo ERRO_TRATA
                        Else: GoTo entra
                        End If
                    End If
                    Int_St1 = Int_St1 - 128
                End If
                    
                If (Int_St1 >= 64) Then
                    StringRetorno = StringRetorno & "Pouco Papel" & vbCr
                    Int_St1 = Int_St1 - 64
                End If
                
                If (Int_St1 >= 32) Then
                    StringRetorno = StringRetorno & "Erro no rel�gio" & vbCr
                    Int_St1 = Int_St1 - 32
                End If
                
                If (Int_St1 >= 16) Then
                    StringRetorno = StringRetorno & "Impressora em erro" & vbCr
                    Int_St1 = Int_St1 - 16
                End If
                    
                If (Int_St1 >= 8) Then
                    StringRetorno = StringRetorno & "Primeiro dado do comando n�o foi Esc" & vbCr
                    Int_St1 = Int_St1 - 8
                End If
                    
                If (Int_St1 >= 4) Then
                    StringRetorno = StringRetorno & "Comando inexistente" & vbCr
                    Int_St1 = Int_St1 - 4
                End If
                    
                If (Int_St1 >= 2) Then
                    StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
                    
                    Int_St1 = Int_St1 - 2
                End If
                
                If (Int_St1 >= 1) Then
                    StringRetorno = StringRetorno & "N�mero de par�metros inv�lido no comando" & vbCr
                    Int_St1 = Int_St1 - 1
                End If
                    
                If (Int_St2 >= 128) Then
                    StringRetorno = "Tipo de Par�metro de comando inv�lido" & vbCr
                    Int_St2 = Int_St2 - 128
                End If
                    
                If (Int_St2 >= 64) Then
                    StringRetorno = StringRetorno & "Mem�ria fiscal lotada" & vbCr
                    Int_St2 = Int_St2 - 64
                End If
                
                If (Int_St2 >= 32) Then
                    StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
                    Int_St2 = Int_St2 - 32
                End If
                
                If (Int_St2 >= 16) Then
                    StringRetorno = StringRetorno & "Al�quota n�o programada" & vbCr
                    Int_St2 = Int_St2 - 16
                End If
                    
                If (Int_St2 >= 8) Then
                    StringRetorno = StringRetorno & "Capacidade de al�quota program�veis lotada" & vbCr
                    Int_St2 = Int_St2 - 8
                End If
                    
                If (Int_St2 >= 4) Then
                    StringRetorno = StringRetorno & "Cancelamento n�o permitido" & vbCr
                    Int_St2 = Int_St2 - 4
                End If
                    
                If (Int_St2 >= 2) Then
                    StringRetorno = StringRetorno & "CGC/IE do propriet�rio n�o programados" & vbCr
                    Int_St2 = Int_St2 - 2
                End If
                
                If (Int_St2 >= 1) Then
                    StringRetorno = StringRetorno & "Comando n�o executado" & vbCr
                    Int_St2 = Int_St2 - 1
                End If
                    
                Str_ErroExtendido = Space(4)
                Daruma_FI_RetornaErroExtendido Str_ErroExtendido
                
                If RetornaMensagem Then
                    RetornaMensagem = "Status da Impressora: " & ValorRetorno & vbCr & vbLf & StringRetorno & vbCr & vbLf & Label & RetornoFuncao & Chr(13) & Chr(10) & st1(CStr(iST1)) + st2(CStr(iST2)) + "Erro Extendido = " + Str_ErroExtendido
                Else
                    RetornaMensagem = "Status da Impressora: " & ValorRetorno & vbCr & vbLf & StringRetorno & Chr(13) & Chr(10) & st1(CStr(iST1)) + st2(CStr(iST2)) + "Erro Extendido = " + Str_ErroExtendido
                End If
        
                MsgBox RetornaMensagem, vbOKOnly + vbInformation, TituloJanela
                Exit Function
            End If 'fim do Int_St1 <> 0 and Int_St2 <> 0
            
        Case -1
            MsgBox "Erro de execu��o da fun��o.", vbOKOnly + vbCritical, TituloJanela
        Case -2
            MsgBox "Par�metro inv�lido na fun��o.", vbOKOnly + vbExclamation, TituloJanela
        Case -3
            MsgBox "Al�quota n�o programada.", vbOKOnly + vbExclamation, TituloJanela
        Case -4
            MsgBox "Problemas com o arquivo de inicializa��o.", vbOKOnly + vbCritical, TituloJanela
        Case -5
            MsgBox "Erro ao abrir a porta de comunica��o.", vbOKOnly + vbExclamation, TituloJanela
        Case -6
            MsgBox "Impressora desligada ou cabo de comunica��o desconectado.", vbOKOnly + vbExclamation, TituloJanela
        Case -7
            MsgBox "Banco n�o encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
        Case -8
            MsgBox "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt.", vbOKOnly + vbExclamation, TituloJanela
        Case -18
            MsgBox "N�o foi poss�vel abrir arquivo INTPOS.001 !", vbOKOnly + vbExclamation, TituloJanela
        Case -19
            MsgBox "Par�metro diferentes !", vbOKOnly + vbExclamation, TituloJanela
        Case Is = -20
            MsgBox "Transa��o cancelada pelo Operador !", vbOKOnly + vbExclamation, TituloJanela
        Case -21
            MsgBox "A Transa��o n�o foi aprovada !", vbOKOnly + vbExclamation, TituloJanela
        Case -22
            MsgBox "N�o foi poss�vel terminal a Impress�o !", vbOKOnly + vbExclamation, TituloJanela
        Case -23
            MsgBox "N�o foi poss�vel terminal a Opera��o !", vbOKOnly + vbExclamation, TituloJanela
        Case -24
            MsgBox "Forma de pagamento n�o programada.", vbOKOnly + vbExclamation, TituloJanela
        Case -25
            MsgBox "Totalizador n�o fiscal n�o programado.", vbOKOnly + vbExclamation, TituloJanela
        Case -26
            MsgBox "Transa��o j� realizada.", vbOKOnly + vbExclamation, TituloJanela
        Case -27
            MsgBox "Status diferente de 6,0,0.", vbOKOnly + vbExclamation, TituloJanela
        Case -28
            MsgBox "N�o h� dados para serem impressos.", vbOKOnly + vbExclamation, TituloJanela
            
    End Select
    
    Exit Function
ERRO_TRATA:
    'If Err.Number <> 0 Then f_TrataErro
End Function

'Retornos da DUAL
Public Function Retorno_DUAL()
    If Int_Retorno = 1 Then
        MsgBox "1(um) - Impressora OK!", vbInformation, "Daruma Framework"
    End If
    If Int_Retorno = -50 Then
        MsgBox "-50 - Impressora OFF-LINE!", vbCritical, "Daruma Framework"
    End If
    If Int_Retorno = -51 Then
        MsgBox "-51 - Impressora Sem Papel!", vbCritical, "Daruma Framework"
    End If
    If Int_Retorno = -27 Then
        MsgBox "-27 - Erro Generico!", vbCritical, "Daruma Framework"
    End If
    If Int_Retorno = 0 Then
        MsgBox "0 - Impressora Desligada!", vbCritical, "Daruma Framework"
    End If
End Function

'Retornos do TA1000
Public Function Retorno_TA1000()

If Int_Retorno = 1 Then
    MsgBox "1(um) - Metodo Executado com Sucesso!", vbInformation, "Daruma Framework"
Else
    MsgBox CStr(Int_Retorno) + "   Erro Generico!", vbCritical, "Daruma Framework"
End If
End Function

'Funcao que trata os retornos das Impressora Fiscais Restaurante
Public Function Daruma_MostrarRetornoRestaurante()

Dim Str_ErroExtendido As String
Int_Ack = 0
Int_St1 = 0
Int_St2 = 0
Daruma_FIR_RetornoImpressora Int_Ack, Int_St1, Int_St2

If Int_St1 <> 0 And Int_St2 <> "0" Then
    Str_ErroExtendido = Space(4)
    Daruma_FI_RetornaErroExtendido Str_ErroExtendido

    MsgBox "Retorno do Metodo = " + CStr(Int_Retorno) + Chr(13) + Chr(10) _
            + "Ack = " + CStr(Int_Ack) + Chr(13) + Chr(10) _
            + st1(CStr(Int_St1)) _
            + st1(CStr(Int_St2)) _
            + "Erro Extendido = " + Str_ErroExtendido, , "Impressora Fiscal"
End If
End Function

Function st1(codigo As Integer) As String
    If codigo = 0 Then
        st1 = "" 'sem erro
    ElseIf codigo = 1 Then
        st1 = "st1 = " & codigo & " - N�mero de par�metros inv�lidos" + Chr(13) + Chr(10)
    ElseIf codigo = 2 Then
        st1 = "st1 = " & codigo & " - Cupom Fiscal Aberto e foi Cancelado" + Chr(13) + Chr(10)
        'cancelar o cupom
        'CancelandoUltimoCupom
    ElseIf codigo = 4 Then
        st1 = "st1 = " & codigo & " - M�todo inexistente" + Chr(13) + Chr(10)
    ElseIf codigo = 8 Then
        st1 = "st1 = " & codigo & " - Primeiro dado do m�todo n�o foi ESC (1Bh)" + Chr(13) + Chr(10)
    ElseIf codigo = 16 Then
        st1 = "st1 = " & codigo & " - Impressora em erro" + Chr(13) + Chr(10)
    ElseIf codigo = 32 Then
        st1 = "st1 = " & codigo & " - Erro no rel�gio da impressora" + Chr(13) + Chr(10)
    ElseIf codigo = 64 Then
        st1 = "st1 = " & codigo & " - O Papel est� acabando" + Chr(13) + Chr(10)
    ElseIf codigo = 128 Then
        st1 = "st1 = " & codigo & " - O Papel acabou" + Chr(13) + Chr(10)
    End If
End Function

Function st2(codigo As Integer) As String
    If codigo = 0 Then
        st2 = "" 'sem erro
    ElseIf codigo = 1 Then
        st2 = "st2 = " & codigo & " - M�todo n�o existente" + Chr(13) + Chr(10)
    ElseIf codigo = 2 Then
        st2 = "st2 = " & codigo & " - CNPJ/IE do propriet�rio n�o definidos" + Chr(13) + Chr(10)
    ElseIf codigo = 4 Then
        st2 = "st2 = " & codigo & " - Este cancelamento n�o � permitido" + Chr(13) + Chr(10)
    ElseIf codigo = 8 Then
        st2 = "st2 = " & codigo & " - Capacidade de al�quota esgotada" + Chr(13) + Chr(10)
    ElseIf codigo = 16 Then
        st2 = "st2 = " & codigo & " - Al�quota n�o definida" + Chr(13) + Chr(10)
    ElseIf codigo = 32 Then
        st2 = "st2 = " & codigo & " - Erro na mem�ria RAM n�o vol�til" + Chr(13) + Chr(10)
    ElseIf codigo = 64 Then
        st2 = "st2 = " & codigo & " - Mem�ria fiscal cheia" + Chr(13) + Chr(10)
    ElseIf codigo = 128 Then
        st2 = "st2 = " & codigo & " - Tipo de par�metro inv�lido" + Chr(13) + Chr(10)
    End If
End Function

' Fun��o: ImprimeTransacao
' Objetivo: Realiza a impress�o da Transa��o TEF
' Par�metros: string para a Forma de Pagamento
' string para a Valor da Forma de Pagamento
' string para o N�mero do Cupom Fiscal (COO)
' TDateTime para identificar o n�mero da transa��o
' Retorno: True para OK ou False para n�o OK

'Function ImprimeTransacao(ByVal cFormaPGTO As String, ByVal cValorPago As String, _
                          ByVal cCOO As String, ByVal hora As String, _
                          ByVal iConta As Integer, ByVal Gerencial As Boolean) As Boolean
Function ImprimeTransacao(cFormaPGTO As String, cValorPago As String, cCOO As String, cIdentificacao As String) As Integer
'On Error GoTo ERRO_TRATA

    Dim cLinhaArquivo As String
    Dim cLinha As String
    Dim cSaltaLinha As String
    Dim cConteudo As String
    Dim iVezes As Integer
    ' Bloqueia o teclado e o mouse para a impress�o do TEF
    'intRetorno = Daruma_FI_IniciaModoTEF()
    intRetorno = Daruma_TEF_TravarTeclado(1)
    cArquivoTemp = Dir(App.Path & "\IMPRIME.TXT")
    If cArquivoTemp <> "" Then
        intRetorno = Daruma_FI_AbreComprovanteNaoFiscalVinculado(cFormaPGTO, cValorPago, cCOO)
        VerificaRetornoImpressoraDaruma "", Trim(intRetorno), "Imprimindo Transa��o TEF"
    End If
    cConteudo = ""
    cLinha = ""
    
    Open App.Path & "\IMPRIME.TXT" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cLinha
        cConteudo = cConteudo + cLinha + Chr(13) + Chr(10)
        intRetorno = Daruma_FI_UsaComprovanteNaoFiscalVinculado(cLinha + Chr(13))
        VerificaRetornoImpressoraDaruma "", Trim(intRetorno), "Imprimindo Transa��o TEF"
        If EOF(1) Then
            cSaltaLinha = Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10)
            intRetorno = Daruma_FI_UsaComprovanteNaoFiscalVinculado(cSaltaLinha)
            VerificaRetornoImpressoraDaruma "", Trim(intRetorno), "Imprimindo Transa��o TEF"
            ' Est� sendo usado um form para a exibi��o desta mensagem
            frmMensagem.lblMensagem.Caption = "Por favor, destaque a 1� Via"
            frmMensagem.Show
            frmMensagem.Refresh
            Sleep (5000)
            Unload frmMensagem
            frmPrincipal.Refresh
            intRetorno = Daruma_FI_UsaComprovanteNaoFiscalVinculado(cConteudo)
            VerificaRetornoImpressoraDaruma "", Trim(intRetorno), "Imprimindo Transa��o TEF"
        End If
    Loop
    ' Desbloqeia o teclado e o mouse
    intRetorno = Daruma_TEF_TravarTeclado(0)
    Close #1
    Kill App.Path & "\IMPRIME.TXT"
    intRetorno = Daruma_FI_FechaComprovanteNaoFiscalVinculado()
    VerificaRetornoImpressoraDaruma "", Trim(intRetorno), "Imprimindo Transa��o TEF"
    
    Exit Function
ERRO_TRATA:
    'If Err.Number <> 0 Then f_TrataErro
End Function
