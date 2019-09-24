Attribute VB_Name = "modDarumaANT"
Global Int_Retorno As Integer
Global Int_Ack As Integer
Global Int_St1 As Integer
Global Int_St2 As Integer
'Global Str_ErroExtendido As String
Global Int_TIPOliquota As Integer
Global Str_Informacao As String
Global Int_Informacao As Integer

'**********************************************************************************************************************'
'                                                                                                                      '
'                                                       Release                                                '
'                                                                                                                      '
'**********************************************************************************************************************'
'Fun��o Release
Public Declare Function Daruma_FI_SaltarLinhas Lib "Daruma32.dll" (ByVal Linhas As Integer) As Integer

'**********************************************************************************************************************'
'                                                                                                                      '
'                                                       MFD                                                          '
'                                                                                                                      '
'**********************************************************************************************************************'
'Fun��o Codigo de barras da MFD

Public Declare Function Daruma_FIMFD_ImprimeCodigoBarras Lib "Daruma32.dll" (ByVal TIPO As String, ByVal codigo As String, ByVal Largura As String, ByVal Altura As String, ByVal Posicao As String) As Integer
Public Declare Function Daruma_FIMFD_DownloadDaMFD Lib "Daruma32.dll" (ByVal CoInicial As String, ByVal CoFinal As String) As Integer
Public Declare Function Daruma_FIMFD_CasasDecimaisProgramada Lib "Daruma32.dll" (ByVal Quantidade As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_FIMFD_IndicePrimeiroVinculado Lib "Daruma32.dll" (ByVal Indice As String) As Integer
Public Declare Function Daruma_FIMFD_RetornaInformacao Lib "Daruma32.dll" (ByVal Indice As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_FIMFD_TerminaFechamentoCupomCodigoBarras Lib "Daruma32.dll" (ByVal Mensagem As String, ByVal TIPO As String, ByVal codigo As String, ByVal Largura As String, ByVal Altura As String, ByVal Posicao As String) As Integer
Public Declare Function Daruma_FIMFD_SinalSonoro Lib "Daruma32.dll" (ByVal NumeroBeeps As String) As Integer
Public Declare Function Daruma_FIMFD_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StsCF_MFD As String) As Integer
Public Declare Function Daruma_FIMFD_ProgramaRelatoriosGerenciais Lib "Daruma32.dll" (ByVal NomeRelatorio As String) As Integer
Public Declare Function Daruma_FIMFD_VerificaRelatoriosGerenciais Lib "Daruma32.dll" (ByVal Relatorios As String) As Integer
Public Declare Function Daruma_FIMFD_AbreRelatorioGerencial Lib "Daruma32.dll" (ByVal NomeRelatorio As String) As Integer
Public Declare Function Daruma_FIMFD_EmitirCupomAdicional Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIMFD_AcionarGuilhotina Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIMFD_EqualizarVelocidade Lib "Daruma32.dll" (ByVal EqualizaVelocidade As String) As Integer
Public Declare Function Daruma_FIMFD_AbreRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal CPF As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Daruma_FIMFD_RecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal DescricaoTotalizador As String, ByVal AcresDesc As String, ByVal TipoAcresDesc As String, ByVal ValorAcresDesc As String, ByVal ValorRecebimento As String) As Integer
Public Declare Function Daruma_FIMFD_IniciaFechamentoNaoFiscal Lib "Daruma32.dll" (ByVal AcresDesc As String, ByVal TipoAcresDesc As String, ByVal ValorAcresDesc As String) As Integer
Public Declare Function Daruma_FIMFD_EfetuaFormaPagamentoNaoFiscal Lib "Daruma32.dll" (ByVal FormaPgto As String, ByVal Valor As String, ByVal Observacao As String) As Integer
Public Declare Function Daruma_FIMFD_TerminaFechamentoNaoFiscal Lib "Daruma32.dll" (ByVal MsgPromo As String) As Integer
Public Declare Function Daruma_FIMFD_ProgramarGuilhotina Lib "Daruma32.dll" (ByVal Separacao_entre_Documentos As String, ByVal Linhas_para_Acionamento_Guilhotina As String, ByVal Status_da_Guilhotina As String, ByVal Impressao_Antecipada_Cliche As String) As Integer
Public Declare Function Daruma_FIMFD_DescontoAcrescimoItem Lib "Daruma32.dll" (ByVal NumeroDoItem As String, ByVal DescontoOuAcrescimoItem As String, ByVal TipoDoDescontoOuAcrescimoItem As String, ByVal ValorDoDescontoOuAcrescimo As String) As Integer



'**********************************************************************************************************************'
'                                                                                                                      '
'                                                       FS345                                                          '
'                                                                                                                      '
'**********************************************************************************************************************'

'Metodos de Verificacao

Public Declare Function Daruma_FI_ResetaImpressora Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_LeituraXSerial Lib "Daruma32.dll" () As Integer

'M�todos Ato Cotepe PAF-ECF COO e Data
Public Declare Function Daruma_FIMFD_GerarAtoCotepePafData Lib "Daruma32.dll" (ByVal Dataini As String, ByVal DataFinal As String) As Integer
Public Declare Function Daruma_FIMFD_GerarAtoCotepePafCOO Lib "Daruma32.dll" (ByVal COOini As String, ByVal COOFinal As String) As Integer

'Metodos Cupom
Public Declare Function Daruma_FI_AbreCupom Lib "Daruma32.dll" (ByVal CPF_ou_CNPJ As String) As Integer
Public Declare Function Daruma_FI_VendeItem Lib "Daruma32.dll" (ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Daruma_FI_VendeItemDepartamento Lib "Daruma32.dll" (ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Daruma_FI_VendeItemTresDecimais Lib "Daruma32.dll" (ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal Desconto_ou_Acrescimo As String, ByVal Percentual_Desconto_ou_Acrescimo As String) As Integer
Public Declare Function Daruma_FI_FechaCupomResumido Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal TextoLivre As String) As Integer
Public Declare Function Daruma_FI_IdentificaConsumidor Lib "Daruma32.dll" (ByVal NomeConsumidor As String, ByVal Endereco As String, ByVal CPF_ou_CNPJ As String) As Integer
Public Declare Function Daruma_FI_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_FechaCupom Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_CancelaItemAnterior Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_CancelaItemGenerico Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FI_CancelaCupom Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AumentaDescricaoItem Lib "Daruma32.dll" (ByVal DESCRICAO As String) As Integer
Public Declare Function Daruma_FI_UsaUnidadeMedida Lib "Daruma32.dll" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Daruma_FI_EmitirCupomAdicional Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_EstornoFormasPagamento Lib "Daruma32.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer

'Metodos para Recebimentos e Relatorios
Public Declare Function Daruma_FI_AbreComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NUMEROCUPOM As String) As Integer
Public Declare Function Daruma_FI_UsaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal texto As String) As Integer
Public Declare Function Daruma_FI_FechaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioGerencial Lib "Daruma32.dll" (ByVal texto As String) As Integer
Public Declare Function Daruma_FI_AbreRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_EnviarTextoCNF Lib "Daruma32.dll" (ByVal texto As String) As Integer
Public Declare Function Daruma_FI_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal DescricaoTotalizador As String, ByVal ValorRecebimento As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_AbreRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal DescricaoTotalizador As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal ValorRecebimento As String, ByVal TextoLivreto As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamentoNaoFiscal Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal ObsLivre As String) As Integer
Public Declare Function Daruma_FI_Sangria Lib "Daruma32.dll" (ByVal Valor As String) As Integer
Public Declare Function Daruma_FI_Suprimento Lib "Daruma32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_FundoCaixa Lib "Daruma32.dll" (ByVal ValorPagamento As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_LeituraX Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ReducaoZ Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FI_ReducaoZAjustaDataHora Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalData Lib "Daruma32.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalReducao Lib "Daruma32.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalSerialData Lib "Daruma32.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Public Declare Function Daruma_FIMFD_GerarMFPAF_Data Lib "Daruma32.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalSerialReducao Lib "Daruma32.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String) As Integer
Public Declare Function Daruma_FIMFD_GerarMFPAF_CRZ Lib "Daruma32.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String) As Integer

'Metodos Modo Gaveta, Autentica e Outras
Public Declare Function Daruma_FI_VerificaDocAutenticacao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_Autenticacao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AutenticacaoStr Lib "Daruma32.dll" (ByVal AtenticacaoStr As String) As Integer
Public Declare Function Daruma_FI_VerificaEstadoGaveta Lib "Daruma32.dll" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Daruma_FI_VerificaEstadoGavetaStr Lib "Daruma32.dll" (ByVal EstadoGaveta As String) As Integer
Public Declare Function Daruma_FI_AcionaGaveta Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AbrePortaSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_FechaPortaSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AberturaDoDia Lib "Daruma32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_FechamentoDoDia Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ImprimeConfiguracoesImpressora Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RegistraNumeroSerie Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_VerificaNumeroSerie Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RetornaSerialCriptografado Lib "Daruma32.dll" (ByVal SerialCriptografado As String, ByVal NumeroSerial As String) As Integer
Public Declare Function Daruma_FI_ConfiguraHorarioVerao Lib "Daruma32.dll" (ByVal DataEntrada As String, ByVal DataSaida As String, ByVal controle As String) As Integer
Public Declare Function Daruma_FI_GeraCriptografia Lib "Daruma32.dll" (ByVal Str_Input As String, ByVal Str_Output As String) As Integer
'Metodos Prog e Config
Public Declare Function Daruma_FI_ProgramaAliquota Lib "Daruma32.dll" (ByVal Aliquota As String, ByVal ICMS_ou_ISS As Integer) As Integer
Public Declare Function Daruma_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "Daruma32.dll" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Daruma_FI_ProgramaFormasPagamento Lib "Daruma32.dll" (ByVal DescricaoFormasPgto As String) As Integer
Public Declare Function Daruma_FI_ProgramaOperador Lib "Daruma32.dll" (ByVal NomeOperador As String) As Integer
Public Declare Function Daruma_FI_ProgramaArredondamento Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ProgramaTruncamento Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_LinhasEntreCupons Lib "Daruma32.dll" (ByVal Linhas As Integer) As Integer
Public Declare Function Daruma_FI_EspacoEntreLinhas Lib "Daruma32.dll" (ByVal Dots As Integer) As Integer
Public Declare Function Daruma_FI_ProgramaHorarioVerao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_EqualizaFormasPgto Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ProgramaVinculados Lib "Daruma32.dll" (ByVal DESCRICAO As String) As Integer
Public Declare Function Daruma_FI_ProgFormasPagtoSemVincular Lib "Daruma32.dll" (ByVal DESCRICAO As String) As Integer

'Metodos de Configuracao do ECF
Public Declare Function Daruma_FI_CfgFechaAutomaticoCupom Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgRedZAutomatico Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgLeituraXAuto Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgImpEstGavVendas Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgCalcArredondamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgHorarioVerao Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgSensorAut Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgCupomAdicional Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgPermMensPromCNF Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgEspacamentoCupons Lib "Daruma32.dll" (ByVal DistanciaCupons As String) As Integer
Public Declare Function Daruma_FI_CfgHoraMinReducaoZ Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgLimiarNearEnd Lib "Daruma32.dll" (ByVal NumeroLinhas As String) As Integer
Public Declare Function Daruma_FI_CfgLegProdutos Lib "Daruma32.dll" (ByVal Flag As String) As Integer

'Metodos para Configuracao do Registry
Public Declare Function Daruma_Registry_CupomAdicionalDllConfig Lib "Daruma32.dll" (ByVal CupomAdicional As String) As Integer
Public Declare Function Daruma_Registry_CupomAdicionalDll Lib "Daruma32.dll" (ByVal CupomAdicional As String) As Integer
Public Declare Function Daruma_Registry_AplMensagem1 Lib "Daruma32.dll" (ByVal Str_AplMensagem_1 As String) As Integer
Public Declare Function Daruma_Registry_AplMensagem2 Lib "Daruma32.dll" (ByVal Str_AplMensagem_2 As String) As Integer
Public Declare Function Daruma_Registry_AlteraRegistry Lib "Daruma32.dll" (ByVal NomeChave As String, ByVal ValorChave As String) As Integer
Public Declare Function Daruma_Registry_Velocidade Lib "Daruma32.dll" (ByVal VelocidadeDaPortaSerial As String) As Integer
Public Declare Function Daruma_Registry_Porta Lib "Daruma32.dll" (ByVal NomePorta As String) As Integer
Public Declare Function Daruma_Registry_Path Lib "Daruma32.dll" (ByVal Path As String) As Integer
Public Declare Function Daruma_Registry_Status Lib "Daruma32.dll" (ByVal Status As String) As Integer
Public Declare Function Daruma_Registry_StatusFuncao Lib "Daruma32.dll" (ByVal StatusFuncao As String) As Integer
Public Declare Function Daruma_Registry_Retorno Lib "Daruma32.dll" (ByVal Retorno As String) As Integer
Public Declare Function Daruma_Registry_ControlePorta Lib "Daruma32.dll" (ByVal ControlePorta As String) As Integer
Public Declare Function Daruma_Registry_ModoGaveta Lib "Daruma32.dll" (ByVal ModoGaveta As String) As Integer
Public Declare Function Daruma_Registry_Log Lib "Daruma32.dll" (ByVal Log As String) As Integer
Public Declare Function Daruma_Registry_NomeLog Lib "Daruma32.dll" (ByVal NomeLog As String) As Integer
Public Declare Function Daruma_Registry_Separador Lib "Daruma32.dll" (ByVal Separador As String) As Integer
Public Declare Function Daruma_Registry_SeparaMsgPromo Lib "Daruma32.dll" (ByVal SeparaMsgPromo As String) As Integer
Public Declare Function Daruma_Registry_ZAutomatica Lib "Daruma32.dll" (ByVal ZAutomatica As String) As Integer
Public Declare Function Daruma_Registry_XAutomatica Lib "Daruma32.dll" (ByVal XAutomatica As String) As Integer
Public Declare Function Daruma_Registry_VendeItemUmaLinha Lib "Daruma32.dll" (ByVal VendeItem1Lin13Dig As String) As Integer
Public Declare Function Daruma_Registry_Default Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_Registry_RetornaValor Lib "Daruma32.dll" (ByVal Produto As String, ByVal chave As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_Registry_TerminalServer Lib "Daruma32.dll" (ByVal TerminalServer As String) As Integer
Public Declare Function Daruma_Registry_ErroExtendidoOk Lib "Daruma32.dll" (ByVal ValorErro As String) As Integer
Public Declare Function Daruma_Registry_AbrirDiaFiscal Lib "Daruma32.dll" (ByVal AbrirDiaFiscal As String) As Integer
Public Declare Function Daruma_Registry_VendaAutomatica Lib "Daruma32.dll" (ByVal VendaAutomatica As String) As Integer
Public Declare Function Daruma_Registry_IgnorarPoucoPapel Lib "Daruma32.dll" (ByVal IgnorarPoucoPapel As String) As Integer
Public Declare Function Daruma_Registry_ImprimeRegistry Lib "Daruma32.dll" (ByVal Produto As String) As Integer
Public Declare Function Daruma_Registry_PCExpanionLogin Lib "Daruma32.dll" (ByVal Flag_Login As String) As Integer
Public Declare Function Daruma_Registry_TEF_NumeroLinhasImpressao Lib "Daruma32.dll" (ByVal NumeroLinhasImpressao As String) As Integer
Public Declare Function Daruma_Registry_MFD_ArredondaValor Lib "Daruma32.dll" (ByVal ArredondaValor As String) As Integer
Public Declare Function Daruma_Registry_MFDValorFinal Lib "Daruma32.dll" (ByVal ValorFinal As String) As Integer
Public Declare Function Daruma_Registry_MFD_ArredondaQuantidade Lib "Daruma32.dll" (ByVal ArredondaQuantidade As String) As Integer
Public Declare Function Daruma_Registry_MFD_ProgramarSinalSonoro Lib "Daruma32.dll" (ByVal NomeChave As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_Registry_MFD_LeituraMFCompleta Lib "Daruma32.dll" (ByVal Valor As String) As Integer
Public Declare Function Daruma_Registry_NumeroSerieNaoFormatado Lib "Daruma32.dll" (ByVal Formatado As String) As Integer
Public Declare Function Daruma_Registry_LogTamMaxMB Lib "Daruma32.dll" (ByVal LogTamMaxMB As String) As Integer
Public Declare Function Daruma_Registry_SintegraPath Lib "Daruma32.dll" (ByVal Path As String) As Integer
Public Declare Function Daruma_Registry_SintegraSeparador Lib "Daruma32.dll" (ByVal Separador As String) As Integer
Public Declare Function Daruma_Registry_SintegraUF Lib "Daruma32.dll" (ByVal UF As String) As Integer

'Metodos de Status
Public Declare Function Daruma_FI_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StsCF As String) As Integer
Public Declare Function Daruma_FI_StatusRelatorioGerencial Lib "Daruma32.dll" (ByVal StsGerencial As String) As Integer
Public Declare Function Daruma_FI_StatusComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal StsCNFV As String) As Integer
Public Declare Function Daruma_FI_StatusComprovanteNaoFiscalNaoVinculado Lib "Daruma32.dll" (ByVal StsCNFV As String) As Integer
Public Declare Function Daruma_FI_VerificaImpressoraLigada Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_VerificaTotalizadoresParciais Lib "Daruma32.dll" (ByVal Totalizadores_Parciais As String) As Integer
Public Declare Function Daruma_FI_VerificaModoOperacao Lib "Daruma32.dll" (ByVal Modo As String) As Integer
Public Declare Function Daruma_FI_VerificaTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Daruma_FI_VerificaTotalizadoresNaoFiscaisEx Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Daruma_FI_VerificaTruncamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_VerificaAliquotasIss Lib "Daruma32.dll" (ByVal AliquotasIss As String) As Integer
Public Declare Function Daruma_FI_VerificaRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal Recebimentos As String) As Integer
Public Declare Function Daruma_FI_VerificaTipoImpressora Lib "Daruma32.dll" (ByRef tipoImpressora As Integer) As Integer
Public Declare Function Daruma_FI_VerificaIndiceAliquotasIss Lib "Daruma32.dll" (ByVal AliquotaIss As String) As Integer
Public Declare Function Daruma_FI_VerificaModeloECF Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_VerificaDescricaoFormasPagamento Lib "Daruma32.dll" (ByVal DESCRICAO As String) As Integer
Public Declare Function Daruma_FI_VerificaXPendente Lib "Daruma32.dll" (ByVal XPendente As String) As Integer
Public Declare Function Daruma_FI_VerificaZPendente Lib "Daruma32.dll" (ByVal ZPendente As String) As Integer
Public Declare Function Daruma_FI_VerificaDiaAberto Lib "Daruma32.dll" (ByVal DiaAberto As String) As Integer
Public Declare Function Daruma_FI_VerificaHorarioVerao Lib "Daruma32.dll" (ByVal HorarioVerao As String) As Integer
Public Declare Function Daruma_FI_VerificaFormasPagamento Lib "Daruma32.dll" (ByVal formasPagto As String) As Integer
Public Declare Function Daruma_FI_VerificaFormasPagamentoEx Lib "Daruma32.dll" (ByVal FormasPagtoEx As String) As Integer
Public Declare Function Daruma_FI_VerificaEpromConectada Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_VerificaEstadoImpressora Lib "Daruma32.dll" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer

'Metodos de Informacao do ECF e Contadores
Public Declare Function Daruma_FI_ClicheProprietario Lib "Daruma32.dll" (ByVal Cliche As String) As Integer
Public Declare Function Daruma_FI_ClicheProprietarioEx Lib "Daruma32.dll" (ByVal ClicheProprietarioEx As String) As Integer
Public Declare Function Daruma_FI_NumeroCaixa Lib "Daruma32.dll" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Daruma_FI_NumeroLoja Lib "Daruma32.dll" (ByVal NumeroLoja As String) As Integer
Public Declare Function Daruma_FI_NumeroSerie Lib "Daruma32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Daruma_FI_VersaoFirmware Lib "Daruma32.dll" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Daruma_FI_CGC_IE Lib "Daruma32.dll" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Daruma_FI_LerAliquotasComIndice Lib "Daruma32.dll" (ByVal AliquotasComIndice As String) As Integer
Public Declare Function Daruma_FI_NumeroCupom Lib "Daruma32.dll" (ByVal NUMEROCUPOM As String) As Integer
Public Declare Function Daruma_FI_COO Lib "Daruma32.dll" (ByVal Coo_Inicial As String, ByVal COO_Final As String) As Integer
Public Declare Function Daruma_FI_MinutosImprimindo Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Daruma_FI_MinutosLigada Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Daruma_FI_NumeroSubstituicoesProprietario Lib "Daruma32.dll" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Daruma_FI_NumeroIntervencoes Lib "Daruma32.dll" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Daruma_FI_NumeroReducoes Lib "Daruma32.dll" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Daruma_FI_NumeroCuponsCancelados Lib "Daruma32.dll" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Daruma_FI_NumeroOperacoesNaoFiscais Lib "Daruma32.dll" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Daruma_FI_DataHoraImpressora Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FI_DataHoraReducao Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FI_DataMovimento Lib "Daruma32.dll" (ByVal Data As String) As Integer
Public Declare Function Daruma_FI_ContadoresTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Contadores As String) As Integer

'Metodos Totalizadores Gerais
Public Declare Function Daruma_FI_VendaBruta Lib "Daruma32.dll" (ByVal VendaBruta As String) As Integer
Public Declare Function Daruma_FI_VendaBrutaAcumulada Lib "Daruma32.dll" (ByVal VendaBrutaAcumulada As String) As Integer
Public Declare Function Daruma_FI_GrandeTotal Lib "Daruma32.dll" (ByVal GrandeTotal As String) As Integer
Public Declare Function Daruma_FI_Descontos Lib "Daruma32.dll" (ByVal ValorDescontos As String) As Integer
Public Declare Function Daruma_FI_Acrescimos Lib "Daruma32.dll" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Daruma_FI_Cancelamentos Lib "Daruma32.dll" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Daruma_FI_DadosUltimaReducao Lib "Daruma32.dll" (ByVal DadosReducao As String) As Integer
Public Declare Function Daruma_FI_SubTotal Lib "Daruma32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Daruma_FI_Troco Lib "Daruma32.dll" (ByVal Troco As String) As Integer
Public Declare Function Daruma_FI_SaldoAPagar Lib "Daruma32.dll" (ByVal Saldo As String) As Integer
Public Declare Function Daruma_FI_RetornoAliquotas Lib "Daruma32.dll" (ByVal cAliquotas As String) As Integer
Public Declare Function Daruma_FI_ValorPagoUltimoCupom Lib "Daruma32.dll" (ByVal ValorCupom As String) As Integer
Public Declare Function Daruma_FI_ValorFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Daruma_FI_ValorTotalizadorNaoFiscal Lib "Daruma32.dll" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer
Public Declare Function Daruma_FI_UltimoItemVendido Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FI_UltimaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Daruma_FI_TipoUltimoDocumento Lib "Daruma32.dll" (ByVal TipoUltimoDoc As String) As Integer

'Metodos Relatorios Fiscais e Relatorios
Public Declare Function Daruma_FI_MapaResumo Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioTipo60Analitico Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioTipo60Mestre Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_FlagsFiscais Lib "Daruma32.dll" (ByRef Flag As Integer) As Integer
Public Declare Function Daruma_FI_FlagsFiscaisStr Lib "Daruma32.dll" (ByRef Flag As Integer) As Integer
Public Declare Function Daruma_FI_PalavraStatus Lib "Daruma32.dll" (ByVal PalavraStatus As String) As Integer
Public Declare Function Daruma_FI_PalavraStatusBinario Lib "Daruma32.dll" (ByVal PalavraStatusBinario As String) As Integer
Public Declare Function Daruma_FI_SimboloMoeda Lib "Daruma32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Daruma_FI_RetornoImpressora Lib "Daruma32.dll" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer
Public Declare Function Daruma_FI_RetornaErroExtendido Lib "Daruma32.dll" (ByVal ErroExtendido As String) As Integer
Public Declare Function Daruma_FI_RetornaAcrescimoNF Lib "Daruma32.dll" (ByVal AcrescimoNF As String) As Integer
Public Declare Function Daruma_FI_RetornaCFCancelados Lib "Daruma32.dll" (ByVal CNCancelados As String) As Integer
Public Declare Function Daruma_FI_RetornaCNFCancelados Lib "Daruma32.dll" (ByVal CNFCancelados As String) As Integer
Public Declare Function Daruma_FI_RetornaCLX Lib "Daruma32.dll" (ByVal RetornaCLX As String) As Integer
Public Declare Function Daruma_FI_RetornaCNFNV Lib "Daruma32.dll" (ByVal RetornaCNFNV As String) As Integer
Public Declare Function Daruma_FI_RetornaCNFV Lib "Daruma32.dll" (ByVal RetornaCNFV As String) As Integer
Public Declare Function Daruma_FI_RetornaDescricaoCNFV Lib "Daruma32.dll" (ByVal RetornaCNFV As String) As Integer
Public Declare Function Daruma_FI_RetornaCRO Lib "Daruma32.dll" (ByVal RetornaCRO As String) As Integer
Public Declare Function Daruma_FI_RetornaCRZ Lib "Daruma32.dll" (ByVal RetornaCRZ As String) As Integer
Public Declare Function Daruma_FI_RetornaCRZRestante Lib "Daruma32.dll" (ByVal RetornaCRZRestante As String) As Integer
Public Declare Function Daruma_FI_RetornaCancelamentoNF Lib "Daruma32.dll" (ByVal CancelamentoNF As String) As Integer
Public Declare Function Daruma_FI_RetornaDescontoNF Lib "Daruma32.dll" (ByVal DescontoNF As String) As Integer
Public Declare Function Daruma_FI_RetornaGNF Lib "Daruma32.dll" (ByVal RetornaGNF As String) As Integer
Public Declare Function Daruma_FI_RetornaTempoImprimindo Lib "Daruma32.dll" (ByVal TempoImprimindo As String) As Integer
Public Declare Function Daruma_FI_RetornaTempoLigado Lib "Daruma32.dll" (ByVal TempoLigado As String) As Integer
Public Declare Function Daruma_FI_RetornaTotalPagamentos Lib "Daruma32.dll" (ByVal TotalPagamentos As String) As Integer
Public Declare Function Daruma_FI_RetornaTroco Lib "Daruma32.dll" (ByVal Troco As String) As Integer
Public Declare Function Daruma_FI_RetornaValorComprovanteNaoFiscal Lib "Daruma32.dll" (ByVal IndiceRegistrCNF As String, ByVal Ref_Valor As String) As Integer
Public Declare Function Daruma_FI_RetornaIndiceComprovanteNaoFiscal Lib "Daruma32.dll" (ByVal DescricaoRegistrCNF As String, ByVal Ref_Indice As String) As Integer
Public Declare Function Daruma_FI_RetornaRegistradoresNaoFiscais Lib "Daruma32.dll" (ByVal RegistrNaoFiscais As String) As Integer
Public Declare Function Daruma_FI_RetornaRegistradoresFiscais Lib "Daruma32.dll" (ByVal RegistradoresFiscais As String) As Integer
Public Declare Function Daruma_FI_RetornarVersaoDLL Lib "Daruma32.dll" (ByVal RegistradoresFiscais As String) As Integer

'Metodos TEF
Public Declare Function Daruma_TEF_EsperarArquivo Lib "Daruma32.dll" (ByVal PathArquivo As String, ByVal Tempo As String, ByVal Travar As String) As Integer
Public Declare Function Daruma_TEF_ImprimirResposta Lib "Daruma32.dll" (ByVal PathArquivo As String, ByVal Forma As String, ByVal Travar As String) As Integer
Public Declare Function Daruma_TEF_ImprimirRespostaCartao Lib "Daruma32.dll" (ByVal PathArquivo As String, ByVal Forma As String, ByVal Travar As String, ByVal ValorPagamento As String) As Integer
Public Declare Function Daruma_TEF_FechaRelatorio Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TEF_SetFocus Lib "Daruma32.dll" (ByVal TituloJanela As String) As Integer
Public Declare Function Daruma_TEF_TravarTeclado Lib "Daruma32.dll" (ByVal Travar As String) As Integer

'Metodos FIB
Public Declare Function Daruma_FIB_AbreBilhetePassagem Lib "Daruma32.dll" (ByVal Origem As String, ByVal Destino As String, ByVal UF As String, ByVal Percurso As String, ByVal Prestadora As String, ByVal Plataforma As String, ByVal Poltrona As String, ByVal Modalidade As String, ByVal Categoria As String, ByVal DataEmbarque As String, ByVal PassRg As String, ByVal PassNome As String, ByVal PassEndereco As String) As Integer
Public Declare Function Daruma_FIB_VendeItem Lib "Daruma32.dll" (ByVal DESCRICAO As String, ByVal ST As String, ByVal Valor As String, ByVal DescontoAcrescimo As String, ByVal TipoDesconto As String, ByVal ValorDesconto As String) As Integer

'Metodos Sintegra
Public Declare Function Daruma_Sintegra_GerarRegistrosArq Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Municipio As String, ByVal Fax As String, ByVal Cod_Convenio As String, ByVal Cod_Natureza As String, ByVal Cod_Finalidade As String, ByVal Logradouro As String, ByVal Numero As String, ByVal Complemento As String, ByVal Bairro As String, ByVal CEP As String, ByVal Nome_Contato As String, ByVal Telefone As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro10 Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Municipio As String, ByVal Fax As String, ByVal Cod_Convenio As String, ByVal Cod_Finalidade As String, ByVal Cod_Natureza As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro11 Lib "Daruma32.dll" (ByVal Logradouro As String, ByVal Numero As String, ByVal Complemento As String, ByVal Bairro As String, ByVal CEP As String, ByVal Nome_Contato As String, ByVal Telefone As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro60A Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro60D Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro60I Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro60M Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro60R Lib "Daruma32.dll" (ByVal Data_Inicio_Movimento As String, ByVal Data_Fim_Movimento As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_Sintegra_GerarRegistro90 Lib "Daruma32.dll" (ByVal Retorno As String) As Integer
Public Declare Function Daruma_FIMFD_RetornarInfoDownloadMFD Lib "Daruma32.dll" (ByVal Str_Tipo_Download As String, ByVal Str_Data_ou_COO_Inicio As String, ByVal Str_Data_ou_COO_Fim As String, ByVal Str_Indice As String, ByVal Retorno As String) As Integer
Public Declare Function Daruma_FIMFD_RetornarInfoDownloadMFDArquivo Lib "Daruma32.dll" (ByVal Str_Tipo_Download As String, ByVal Str_Data_ou_COO_Inicio As String, ByVal Str_Data_ou_COO_Fim As String, ByVal Str_Indice As String) As Integer

'Metodos Ato Cotep 17
Public Declare Function Daruma_RFD_GerarArquivo Lib "Daruma32.dll" (ByVal Data_Inicial As String, ByVal Data_Final As String) As Integer

'Metodos RSA
Public Declare Function Daruma_RSA_CarregaChavePrivada_Arquivo Lib "Daruma32.dll" (ByVal Str_PathArquivo As String) As Integer
Public Declare Function Daruma_RSA_RetornaChavePublica Lib "Daruma32.dll" (ByVal Str_N As String, ByVal Str_E As String) As Integer
Public Declare Function Daruma_RSA_CriarAssinatura Lib "Daruma32.dll" (ByVal caminhoDoArquivo As String, ByVal sMD5 As String, ByVal sAssinaturaDigital As String) As Integer
  

        

'**********************************************************************************************************************'
'                                                                                                                      '
'                                                       FS2000                                                         '
'                                                                                                                      '
'**********************************************************************************************************************'

'Metodos exclusivos
Public Declare Function Daruma_Registry_FS2000_CupomAdicional Lib "Daruma32.dll" (ByVal CupomAdicional As String) As Integer
Public Declare Function Daruma_Registry_FS2000_TempoEsperaCheque Lib "Daruma32.dll" (ByVal TempodeEspera As String) As Integer
Public Declare Function Daruma_FI2000_DescontoSobreItemVendido Lib "Daruma32.dll" (ByVal NumeroItem As String, ByVal TipoDesconto As String, ByVal ValorDesconto As String) As Integer
Public Declare Function Daruma_FI2000_AcrescimosICMSISS Lib "Daruma32.dll" (ByVal AcrescICMS As String, ByVal AcrescISS As String) As Integer
Public Declare Function Daruma_FI2000_CancelamentosICMSISS Lib "Daruma32.dll" (ByVal CancelICMS As String, ByVal CancelISS As String) As Integer
Public Declare Function Daruma_FI2000_DescontosICMSISS Lib "Daruma32.dll" (ByVal DescICMS As String, ByVal DescISS As String) As Integer
Public Declare Function Daruma_FI2000_LeituraInformacaoUltimosCNF Lib "Daruma32.dll" (ByVal UltimosCNF As String) As Integer
Public Declare Function Daruma_FI2000_LeituraInformacaoUltimoDoc Lib "Daruma32.dll" (ByVal TipoUltimoDoc As String, ByVal ValorUltimoDoc As String) As Integer
Public Declare Function Daruma_FI2000_VerificaRelatorioGerencial Lib "Daruma32.dll" (ByVal Gerencial As String) As Integer
Public Declare Function Daruma_FI2000_CriaRelatorioGerencial Lib "Daruma32.dll" (ByVal NomeGerencial As String) As Integer
Public Declare Function Daruma_FI2000_AbreRelatorioGerencial Lib "Daruma32.dll" (ByVal Indice As String) As Integer
Public Declare Function Daruma_FI2000_CancelamentoCNFV Lib "Daruma32.dll" (ByVal COO_CNFV As String) As Integer
Public Declare Function Daruma_FI2000_SegundaViaCNFVinculado Lib "Daruma32.dll" () As Integer

'Metodos para cheques
Public Declare Function Daruma_FI2000_StatusCheque Lib "Daruma32.dll" (ByVal StatusCheque As String) As Integer
Public Declare Function Daruma_FI2000_ImprimirCheque Lib "Daruma32.dll" (ByVal Banco As String, ByVal Cidade As String, ByVal Data As String, ByVal Favorecido As String, ByVal Valor As String, ByVal PosicaoCheque As String) As Integer
Public Declare Function Daruma_FI2000_ImprimirVersoCheque Lib "Daruma32.dll" (ByVal VersoCheque As String) As Integer
Public Declare Function Daruma_FI2000_LiberarCheque Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI2000_LeituraCodigoMICR Lib "Daruma32.dll" (ByVal CodigoMICR As String) As Integer
Public Declare Function Daruma_FI2000_CancelarCheque Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI2000_LeituraTabelaCheque Lib "Daruma32.dll" (ByVal TabelaCheque As String) As Integer
Public Declare Function Daruma_FI2000_CarregarCheque Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI2000_CorrigirGeometriaCheque Lib "Daruma32.dll" (ByVal NumeroBanco As String, ByVal GeometriaCheque As String) As Integer

'********************************************************************************************************************* '
'                                                                                                                      '
'                                                     TA1000                                                           '
'                                                                                                                      '
'********************************************************************************************************************* '

'Metodos para Registry
Public Declare Function Daruma_Registry_TA1000_Porta Lib "Daruma32.dll" (ByVal Porta As String) As Integer
Public Declare Function Daruma_Registry_TA1000_PathProdutos Lib "Daruma32.dll" (ByVal PathProdutos As String) As Integer
Public Declare Function Daruma_Registry_TA1000_PathUsuarios Lib "Daruma32.dll" (ByVal PathUsuarios As String) As Integer
Public Declare Function Daruma_Registry_TA1000_NumeroItensEnviados Lib "Daruma32.dll" (ByVal NumeroItensEnviados As String) As Integer
Public Declare Function Daruma_Registry_TA1000_PathRelatorios Lib "Daruma32.dll" (ByVal PathRelatorios As String) As Integer


Public Declare Function Daruma_TA1000_CadastrarProdutos Lib "Daruma32.dll" (ByVal DESCRICAO As String, ByVal codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
Public Declare Function Daruma_TA1000_LerProdutos Lib "Daruma32.dll" (ByVal Indice As Integer, ByVal DESCRICAO As String, ByVal codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
Public Declare Function Daruma_TA1000_ConsultarProdutos Lib "Daruma32.dll" (ByVal DESCRICAO As String, ByVal codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
Public Declare Function Daruma_TA1000_AlterarProdutos Lib "Daruma32.dll" (ByVal Codigo_Consultar As String, ByVal DESCRICAO As String, ByVal codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
Public Declare Function Daruma_TA1000_EliminarProdutos Lib "Daruma32.dll" (ByVal codigo As String) As Integer
Public Declare Function Daruma_TA1000_EnviarBancoProdutos Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_ReceberBancoProdutos Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_ReceberProdutosVendidos Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_ZerarProdutos Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_ZerarProdutosVendidos Lib "Daruma32.dll" () As Integer

Public Declare Function Daruma_TA1000_EnviarBancoUsuarios Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_ReceberBancoUsuarios Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_ZerarUsuarios Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_TA1000_CadastrarUsuarios Lib "Daruma32.dll" (ByVal NOME As String, ByVal CPF As String, ByVal CodigoConvenio As String, ByVal CodigoUsuario As String, ByVal UsuarioAnterior As String, ByVal ProximoUsuario As String) As Integer
Public Declare Function Daruma_TA1000_ConsultarUsuarios Lib "Daruma32.dll" (ByVal Codigo_Consultar As String, ByVal NOME As String, ByVal CPF As String, ByVal CodigoConvenio As String, ByVal CodigoUsuario As String, ByVal UsuarioAnterior As String, ByVal ProximoUsuario As String) As Integer
Public Declare Function Daruma_TA1000_AlterarUsuarios Lib "Daruma32.dll" (ByVal Codigo_Consultar As String, ByVal NOME As String, ByVal CPF As String, ByVal CodigoConvenio As String, ByVal CodigoUsuario As String, ByVal UsuarioAnterior As String, ByVal ProximoUsuario As String) As Integer
Public Declare Function Daruma_TA1000_EliminarUsuarios Lib "Daruma32.dll" (ByVal codigo As String) As Integer

Public Declare Function Daruma_TA1000_LeStatusTransferencia Lib "Daruma32.dll" () As Integer
'********************************************************************************************************************* '
'                                                                                                                      '
'                                                       DUAL                                                           '
'                                                                                                                      '
'********************************************************************************************************************* '

'Metodos para Registry
Public Declare Function Daruma_Registry_DUAL_Enter Lib "Daruma32.dll" (ByVal Enter As String) As Integer
Public Declare Function Daruma_Registry_DUAL_Espera Lib "Daruma32.dll" (ByVal Espera As String) As Integer
Public Declare Function Daruma_Registry_DUAL_ModoEscrita Lib "Daruma32.dll" (ByVal ModoEscrita As String) As Integer
Public Declare Function Daruma_Registry_DUAL_Porta Lib "Daruma32.dll" (ByVal Porta As String) As Integer
Public Declare Function Daruma_Registry_DUAL_Tabulacao Lib "Daruma32.dll" (ByVal Tabulacao As String) As Integer
Public Declare Function Daruma_Registry_DUAL_Termica Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_Registry_DUAL_Velocidade Lib "Daruma32.dll" (ByVal Velocidade As String) As Integer

'Metodos para Status
Public Declare Function Daruma_DUAL_VerificaStatus Lib "Daruma32.dll" () As Integer  'Verificar Status
Public Declare Function Daruma_DUAL_VerificaDocumento Lib "Daruma32.dll" () As Integer 'Verifica Documento
Public Declare Function Daruma_DUAL_StatusGaveta Lib "Daruma32.dll" () As Integer  'Verificar Status Gaveta

'Metodos para Autenticacao a Impressao
Public Declare Function Daruma_DUAL_ImprimirArquivo Lib "Daruma32.dll" (ByVal Str_Path As String) As Integer 'Imprimir arquivo
Public Declare Function Daruma_DUAL_ImprimirTexto Lib "Daruma32.dll" (ByVal TextoLivre As String, ByVal TamanhoTexto As Integer) As Integer 'Imprimir Texto Livre
Public Declare Function Daruma_DUAL_Autenticar Lib "Daruma32.dll" (ByVal NumVias As String, ByVal texto As String, ByVal TempoAguardar As String) As Integer 'Autenticar
Public Declare Function Daruma_DUAL_AcionaGaveta Lib "Daruma32.dll" () As Integer  'AcionaGaveta
Public Declare Function Daruma_DUAL_EnviarBMP Lib "Daruma32.dll" (ByVal Path As String) As Integer  'Envia o logotipo
Public Declare Function Daruma_DUAL_VerificarGuilhotina Lib "Daruma32.dll" () As Integer  'Desolve se a impressora esta ou nao com a Guilhotina habilitada
Public Declare Function Daruma_DUAL_ConfigurarGuilhotina Lib "Daruma32.dll" (ByVal Int_Flag As Integer, ByVal Int_LinhasAcionamento As Integer) As Integer 'Configura a Guilhotina

'**********************************************************************************************************************'
'                                                                                                                      '
'                                                       FS318                                                          '
'                                                                                                                      '
'**********************************************************************************************************************'

Public Declare Function Daruma_FIR_ProgramaAliquota Lib "Daruma32.dll" (ByVal Valor_Aliquota As String, ByVal TIPOliquota As Integer) As Integer
Public Declare Function Daruma_FIR_NomeiaTotalizadorNaoSujeitoIcms Lib "Daruma32.dll" (ByVal Indice_do_Totalizador As Integer, ByVal Nome_do_Totalizador As String) As Integer
Public Declare Function Daruma_FIR_ProgramaFormasPagamento Lib "Daruma32.dll" (ByVal Descricao_das_Formas_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_ProgramaOperador Lib "Daruma32.dll" (ByVal Nome_do_Operador As String) As Integer
Public Declare Function Daruma_FIR_ProgramaArredondamento Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ProgramaTruncamento Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_LinhasEntreCupons Lib "Daruma32.dll" (ByVal Linhas_Entre_Cupons As Integer) As Integer
Public Declare Function Daruma_FIR_EspacoEntreLinhas Lib "Daruma32.dll" (ByVal Espaco_Entre_Linhas As Integer) As Integer
Public Declare Function Daruma_FIR_ProgramaHorarioVerao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_EqualizaFormasPgto Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ProgramaVinculados Lib "Daruma32.dll" (ByVal Vinculado As String) As Integer
Public Declare Function Daruma_FIR_ProgFormasPagtoSemVincular Lib "Daruma32.dll" (ByVal Descricao_da_Forma_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_ProgramaMsgTaxaServico Lib "Daruma32.dll" (ByVal Mensagem_da_Taxa_de_Servico As String) As Integer
Public Declare Function Daruma_FIR_AdicionaProdutoCardapio Lib "Daruma32.dll" (ByVal codigo As String, ByVal Valor_Unitario As String, ByVal Aliquota As String, ByVal DESCRICAO As String) As Integer
Public Declare Function Daruma_FIR_CfgEspacamentoCupons Lib "Daruma32.dll" (ByVal DistanciaCupons As String) As Integer
Public Declare Function Daruma_FIR_CfgHoraMinReducaoZ Lib "Daruma32.dll" (ByVal Hora_Min_para_ReducaoZ As String) As Integer
Public Declare Function Daruma_FIR_CfgLimiarNearEnd Lib "Daruma32.dll" (ByVal NumeroLinhas As String) As Integer
Public Declare Function Daruma_FIR_CfgHorarioVerao Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FIR_CfgLegProdutos Lib "Daruma32.dll" (ByVal Flag As String) As Integer

Public Declare Function Daruma_FIR_AbreCupom Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String) As Integer
Public Declare Function Daruma_FIR_AbreCupomRestaurante Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String) As Integer
Public Declare Function Daruma_FIR_AbreCupomBalcao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_VendeItem Lib "Daruma32.dll" (ByVal Mesa As String, ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal Valor_Unitario As String, ByVal Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer
Public Declare Function Daruma_FIR_VendeItemBalcao Lib "Daruma32.dll" (ByVal codigo As String, ByVal Quantidade As String, ByVal Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer
Public Declare Function Daruma_FIR_RegistrarVenda Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String, ByVal codigo As String, ByVal Quantidade As String, ByVal Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer
Public Declare Function Daruma_FIR_RegistroVendaSerial Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String) As Integer
Public Declare Function Daruma_FIR_FechaCupomRestauranteResumido Lib "Daruma32.dll" (ByVal Descricao_da_Forma_de_Pagamento As String, ByVal Mensagem_Promocional As String) As Integer
Public Declare Function Daruma_FIR_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal Acrescimo_ou_Desconto As String, ByVal Tipo_do_Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer
Public Declare Function Daruma_FIR_IniciaFechamentoCupomComServico Lib "Daruma32.dll" (ByVal Acrescimo_ou_Desconto As String, ByVal Tipo_do_Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String, ByVal Indicador_da_Operacao As String, ByVal Taxa_de_Servico As String) As Integer
Public Declare Function Daruma_FIR_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal Descricao_da_Forma_Pagamento As String, ByVal Valor_da_Forma_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal Descricao_da_Forma_Pagamento As String, ByVal Valor_da_Forma_Pagamento As String, ByVal Texto_Livre As String) As Integer
Public Declare Function Daruma_FIR_IdentificaConsumidor Lib "Daruma32.dll" (ByVal Nome_do_Consumidor As String, ByVal Endereco As String, ByVal CPF_ou_CNPJ As String) As Integer
Public Declare Function Daruma_FIR_FechaCupomResumido Lib "Daruma32.dll" (ByVal Descricao_da_Forma_de_Pagamento As String, ByVal Mensagem_Promocional As String) As Integer
Public Declare Function Daruma_FIR_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal Mensagem_Promocional As String) As Integer
Public Declare Function Daruma_FIR_TerminaFechamentoCupomID Lib "Daruma32.dll" (ByVal Mensagem_Promocional As String, ByVal Nome_do_Cliente As String, ByVal Endereco_do_Cliente As String, ByVal Documento_do_Cliente As String) As Integer
Public Declare Function Daruma_FIR_FechaCupomRestaurante Lib "Daruma32.dll" (ByVal Forma_de_Pagamento As String, ByVal Acrescimo_ou_Desconto As String, ByVal TIPOcrescimo_ou_Desconto As String, ByVal Valor_Acrescimo_ou_Desconto As String, ByVal Valor_Pago As String, ByVal Mensagem_Promocional As String) As Integer
Public Declare Function Daruma_FIR_CancelaItem Lib "Daruma32.dll" (ByVal Mesa As String, ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal Valor_Unitario As String, ByVal Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer
Public Declare Function Daruma_FIR_CancelaItemBalcao Lib "Daruma32.dll" (ByVal Codigo_do_Item As String) As Integer
Public Declare Function Daruma_FIR_CancelaCupom Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_CancelarVenda Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String, ByVal codigo As String, ByVal Quantidade As String, ByVal Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer

Public Declare Function Daruma_FIR_TranferirVenda Lib "Daruma32.dll" (ByVal Numero_da_Mesa_Origem As String, ByVal Numero_da_Mesa_Destino As String, ByVal codigo As String, ByVal Quantidade As String, ByVal Acrescimo_ou_Desconto As String, ByVal Valor_do_Acrescimo_ou_Desconto As String) As Integer
Public Declare Function Daruma_FIR_TransfereItem Lib "Daruma32.dll" (ByVal Mesa_Origem As String, ByVal Mesa_Destino As String, ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal Valor_Unitario As String, ByVal Acrescimo_ou_Desconto As String, ByVal Desconto_Percentual As String) As Integer
Public Declare Function Daruma_FIR_TranferirMesa Lib "Daruma32.dll" (ByVal Mesa_Origem As String, ByVal Mesa_Destino As String) As Integer
Public Declare Function Daruma_FIR_ConferenciaMesa Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String, ByVal Mensagem_Promocional As String) As Integer
Public Declare Function Daruma_FIR_LimparMesa Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String) As Integer
Public Declare Function Daruma_FIR_ImprimePrimeiroCupomDividido Lib "Daruma32.dll" (ByVal Numero_da_Mesa As String, ByVal Quantidade_Divisoria As String) As Integer
Public Declare Function Daruma_FIR_RestanteCupomDividido Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_AumentaDescricaoItem Lib "Daruma32.dll" (ByVal Descricao_Extendida As String) As Integer
Public Declare Function Daruma_FIR_UsaUnidadeMedida Lib "Daruma32.dll" (ByVal Unidade_Medida As String) As Integer
Public Declare Function Daruma_FIR_EmitirCupomAdicional Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_EstornoFormasPagamento Lib "Daruma32.dll" (ByVal Forma_de_Origem As String, ByVal Nova_Forma As String, ByVal Valor_Total_Pago As String) As Integer

Public Declare Function Daruma_FIR_AbreComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal Forma_de_Pagamento As String, ByVal Valor_Pago As String, ByVal Numero_do_Cupom As String) As Integer
Public Declare Function Daruma_FIR_UsaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal Texto_Livre As String) As Integer
Public Declare Function Daruma_FIR_FechaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RelatorioGerencial Lib "Daruma32.dll" (ByVal Texto_Livre As String) As Integer
Public Declare Function Daruma_FIR_AbreRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_EnviarTextoCNF Lib "Daruma32.dll" (ByVal Texto_Livre As String) As Integer
Public Declare Function Daruma_FIR_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal Descricao_do_Totalizador As String, ByVal Valor_do_Recebimento As String, ByVal Forma_de_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_AbreRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal Descricao_do_Totalizador As String, ByVal Acrescimo_ou_Desconto As String, ByVal TIPOcrescimo_ou_Desconto As String, ByVal Valor_Acrescimo_ou_Desconto As String, ByVal Valor_do_Recebimento As String, ByVal Texto_Livre As String) As Integer
Public Declare Function Daruma_FIR_EfetuaFormaPagamentoNaoFiscal Lib "Daruma32.dll" (ByVal Forma_de_Pagamento As String, ByVal Valor_da_Forma_Pagamento As String, ByVal Texto_Livre As String) As Integer
Public Declare Function Daruma_FIR_Sangria Lib "Daruma32.dll" (ByVal Valor_da_Sangria As String) As Integer
Public Declare Function Daruma_FIR_Suprimento Lib "Daruma32.dll" (ByVal Valor_do_Suprimento As String, ByVal Forma_de_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_FundoCaixa Lib "Daruma32.dll" (ByVal Valor_do_Fundo_Caixa As String, ByVal Forma_de_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_LeituraX Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ReducaoZ Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FIR_ReducaoZAjustaDataHora Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FIR_RelatorioMesasAbertas Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RelatorioMesasAbertasSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalData Lib "Daruma32.dll" (ByVal Data_Inicial As String, ByVal Data_Final As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalReducao Lib "Daruma32.dll" (ByVal Reducao_Inicial As String, ByVal Reducao_Final As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalSerialData Lib "Daruma32.dll" (ByVal Data_Inicial As String, ByVal Data_Final As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalSerialReducao Lib "Daruma32.dll" (ByVal Reducao_Inicial As String, ByVal Reducao_Final As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaTrabalho Lib "Daruma32.dll" () As Integer

Public Declare Function Daruma_FIR_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StatusCupomFiscal As String) As Integer
Public Declare Function Daruma_FIR_StatusRelatorioGerencial Lib "Daruma32.dll" (ByVal StatusRelGerencial As String) As Integer
Public Declare Function Daruma_FIR_StatusComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal StatusCNFV As String) As Integer
Public Declare Function Daruma_FIR_StatusComprovanteNaoFiscalNaoVinculado Lib "Daruma32.dll" (ByVal StatusCNFNV As String) As Integer
Public Declare Function Daruma_FIR_VerificaImpressoraLigada Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_VerificaTotalizadoresParciais Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Daruma_FIR_VerificaModoOperacao Lib "Daruma32.dll" (ByVal Modo As String) As Integer
Public Declare Function Daruma_FIR_VerificaTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Daruma_FIR_VerificaTotalizadoresNaoFiscaisEx Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Daruma_FIR_VerificaTruncamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FIR_VerificaAliquotasIss Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FIR_VerificaIndiceAliquotasIss Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FIR_VerificaRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal Recebimentos As String) As Integer
Public Declare Function Daruma_FIR_VerificaTipoImpressora Lib "Daruma32.dll" (ByRef tipoImpressora As Integer) As Integer
Public Declare Function Daruma_FIR_VerificaStatusCheque Lib "Daruma32.dll" (ByVal StatusCheque As Integer) As Integer
Public Declare Function Daruma_FIR_VerificaModeloECF Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_VerificaDescricaoFormasPagamento Lib "Daruma32.dll" (ByVal DESCRICAO As String) As Integer
Public Declare Function Daruma_FIR_VerificaXPendente Lib "Daruma32.dll" (ByVal XPendente As String) As Integer
Public Declare Function Daruma_FIR_VerificaZPendente Lib "Daruma32.dll" (ByVal ZPendente As String) As Integer
Public Declare Function Daruma_FIR_VerificaDiaAberto Lib "Daruma32.dll" (ByVal DiaAberto As String) As Integer
Public Declare Function Daruma_FIR_VerificaHorarioVerao Lib "Daruma32.dll" (ByVal HoraioVerao As String) As Integer
Public Declare Function Daruma_FIR_VerificaFormasPagamento Lib "Daruma32.dll" (ByVal Formas As String) As Integer
Public Declare Function Daruma_FIR_VerificaFormasPagamentoEx Lib "Daruma32.dll" (ByVal FormasEx As String) As Integer
Public Declare Function Daruma_FIR_VerificaEpromConectada Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FIR_VerificaEstadoImpressora Lib "Daruma32.dll" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer

Public Declare Function Daruma_FIR_ClicheProprietario Lib "Daruma32.dll" (ByVal Cliche As String) As Integer
Public Declare Function Daruma_FIR_ClicheProprietarioEx Lib "Daruma32.dll" (ByVal ClicheEx As String) As Integer
Public Declare Function Daruma_FIR_NumeroCaixa Lib "Daruma32.dll" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Daruma_FIR_NumeroLoja Lib "Daruma32.dll" (ByVal NumeroLoja As String) As Integer
Public Declare Function Daruma_FIR_NumeroSerie Lib "Daruma32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Daruma_FIR_VersaoFirmware Lib "Daruma32.dll" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Daruma_FIR_CGC_IE Lib "Daruma32.dll" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Daruma_FIR_LerAliquotasComIndice Lib "Daruma32.dll" (ByVal AliquotasComIndice As String) As Integer
Public Declare Function Daruma_FIR_NumeroCupom Lib "Daruma32.dll" (ByVal NUMEROCUPOM As String) As Integer
Public Declare Function Daruma_FIR_COO Lib "Daruma32.dll" (ByVal Inicial As String, ByVal Final As String) As Integer
Public Declare Function Daruma_FIR_MinutosLigada Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Daruma_FIR_NumeroSubstituicoesProprietario Lib "Daruma32.dll" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Daruma_FIR_NumeroIntervencoes Lib "Daruma32.dll" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Daruma_FIR_NumeroReducoes Lib "Daruma32.dll" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Daruma_FIR_NumeroCuponsCancelados Lib "Daruma32.dll" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Daruma_FIR_NumeroOperacoesNaoFiscais Lib "Daruma32.dll" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Daruma_FIR_DataHoraImpressora Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FIR_DataHoraReducao Lib "Daruma32.dll" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Daruma_FIR_DataMovimento Lib "Daruma32.dll" (ByVal Data As String) As Integer
Public Declare Function Daruma_FIR_ContadoresTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Contadores As String) As Integer

Public Declare Function Daruma_FIR_VendaBruta Lib "Daruma32.dll" (ByVal VendaBruta As String) As Integer
Public Declare Function Daruma_FIR_GrandeTotal Lib "Daruma32.dll" (ByVal GrandeTotal As String) As Integer
Public Declare Function Daruma_FIR_Descontos Lib "Daruma32.dll" (ByVal ValorDescontos As String) As Integer
Public Declare Function Daruma_FIR_Acrescimos Lib "Daruma32.dll" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Daruma_FIR_Cancelamentos Lib "Daruma32.dll" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Daruma_FIR_DadosUltimaReducao Lib "Daruma32.dll" (ByVal DadosReducao As String) As Integer
Public Declare Function Daruma_FIR_SubTotal Lib "Daruma32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Daruma_FIR_RetornoAliquotas Lib "Daruma32.dll" (ByVal Aliquotas As String) As Integer
Public Declare Function Daruma_FIR_ValorPagoUltimoCupom Lib "Daruma32.dll" (ByVal ValorCupom As String) As Integer
Public Declare Function Daruma_FIR_ValorPagoUltimoCupomFormatado Lib "Daruma32.dll" (ByVal ValorCupom As String) As Integer
Public Declare Function Daruma_FIR_ValorFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_FIR_ValorTotalizadorNaoFiscal Lib "Daruma32.dll" (ByVal Totalizador As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_FIR_UltimoItemVendido Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FIR_UltimoItemVendidoValor Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FIR_UltimaFormaPagamento Lib "Daruma32.dll" (ByVal Descricao_da_Forma As String, ByVal Valor_da_Forma As String) As Integer
Public Declare Function Daruma_FIR_TipoUltimoDocumento Lib "Daruma32.dll" (ByVal TipoUltimoDoc As String) As Integer

Public Declare Function Daruma_FIR_MapaResumo Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RelatorioTipo60Analitico Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RelatorioTipo60Mestre Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_FlagsFiscais Lib "Daruma32.dll" (ByRef Flag As Integer) As Integer
Public Declare Function Daruma_FIR_PalavraStatus Lib "Daruma32.dll" (ByVal PalavraStatus As String) As Integer
Public Declare Function Daruma_FIR_PalavraStatusBinario Lib "Daruma32.dll" (ByVal PalavraStatusBinario As String) As Integer
Public Declare Function Daruma_FIR_SimboloMoeda Lib "Daruma32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Daruma_FIR_RetornoImpressora Lib "Daruma32.dll" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer
Public Declare Function Daruma_FIR_RetornaErroExtendido Lib "Daruma32.dll" (ByVal ErroExtendido As String) As Integer
Public Declare Function Daruma_FIR_RetornaAcrescimoNF Lib "Daruma32.dll" (ByVal AcrescimoNF As String) As Integer
Public Declare Function Daruma_FIR_RetornaCFCancelados Lib "Daruma32.dll" (ByVal CFCancelados As String) As Integer
Public Declare Function Daruma_FIR_RetornaCNFCancelados Lib "Daruma32.dll" (ByVal CNFCancelados As String) As Integer
Public Declare Function Daruma_FIR_RetornaCLX Lib "Daruma32.dll" (ByVal CLX As String) As Integer
Public Declare Function Daruma_FIR_RetornaCNFNV Lib "Daruma32.dll" (ByVal CNFNV As String) As Integer
Public Declare Function Daruma_FIR_RetornaCNFV Lib "Daruma32.dll" (ByVal CNFV As String) As Integer
Public Declare Function Daruma_FIR_RetornaCRO Lib "Daruma32.dll" (ByVal CRO As String) As Integer
Public Declare Function Daruma_FIR_RetornaCRZ Lib "Daruma32.dll" (ByVal CRZ As String) As Integer
Public Declare Function Daruma_FIR_RetornaCRZRestante Lib "Daruma32.dll" (ByVal CRZRestante As String) As Integer
Public Declare Function Daruma_FIR_RetornaCancelamentoNF Lib "Daruma32.dll" (ByVal CancelamentoNF As String) As Integer
Public Declare Function Daruma_FIR_RetornaDescontoNF Lib "Daruma32.dll" (ByVal DescontoNF As String) As Integer
Public Declare Function Daruma_FIR_RetornaGNF Lib "Daruma32.dll" (ByVal GNF As String) As Integer
Public Declare Function Daruma_FIR_RetornaTempoImprimindo Lib "Daruma32.dll" (ByVal TempoImprimindo As String) As Integer
Public Declare Function Daruma_FIR_RetornaTempoLigado Lib "Daruma32.dll" (ByVal TempoLigado As String) As Integer
Public Declare Function Daruma_FIR_RetornaTotalPagamentos Lib "Daruma32.dll" (ByVal TotalPagamentos As String) As Integer
Public Declare Function Daruma_FIR_RetornaTroco Lib "Daruma32.dll" (ByVal Troco As String) As Integer
Public Declare Function Daruma_FIR_RetornaZeros Lib "Daruma32.dll" (ByVal Zeros As String) As Integer
Public Declare Function Daruma_FIR_RetornaValorComprovanteNaoFiscal Lib "Daruma32.dll" (ByVal Indice_CNF As String, ByVal Informacao As String) As Integer
Public Declare Function Daruma_FIR_RetornaIndiceComprovanteNaoFiscal Lib "Daruma32.dll" (ByVal DescricaoRegistrCNF As String, ByVal RefIndice As String) As Integer
Public Declare Function Daruma_FIR_RetornaRegistradoresNaoFiscais Lib "Daruma32.dll" (ByVal RegistrNaoFiscais As String) As Integer
Public Declare Function Daruma_FIR_RetornaRegistradoresFiscais Lib "Daruma32.dll" (ByVal RegistrFiscais As String) As Integer

Public Declare Function Daruma_FIR_VerificaDocAutenticacao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_Autenticacao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_AutenticacaoStr Lib "Daruma32.dll" (ByVal Autenticacao_Str As String) As Integer
Public Declare Function Daruma_FIR_VerificaEstadoGaveta Lib "Daruma32.dll" (ByRef Estado_Gaveta As Integer) As Integer
Public Declare Function Daruma_FIR_VerificaEstadoGavetaStr Lib "Daruma32.dll" (ByVal Estado_Gaveta As String) As Integer
Public Declare Function Daruma_FIR_AcionaGaveta Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_AbrePortaSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_FechaPortaSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_AberturaDoDia Lib "Daruma32.dll" (ByVal Valor_do_Suprimento As String, ByVal Forma_de_Pagamento As String) As Integer
Public Declare Function Daruma_FIR_FechamentoDoDia Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ImprimeConfiguracoesImpressora Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RegistraNumeroSerie Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_VerificaNumeroSerie Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RetornaSerialCriptografado Lib "Daruma32.dll" (ByVal SerialCriptografado As String, ByVal NumeroSerial As String) As Integer
Public Declare Function Daruma_FIR_ConfiguraHorarioVerao Lib "Daruma32.dll" (ByVal DataEntrada As String, ByVal DataSaida As String, ByVal controle As String) As Integer
Public Declare Function Daruma_FIR_ZeraCardapio Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ImprimeCardapio Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_CardapioSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIMFD_GTCodificado Lib "Daruma32.dll" (ByVal GTCodificado As String) As Integer
Public Declare Function Daruma_FIMFD_Verifica_GTCodificado Lib "Daruma32.dll" (ByVal GTCodificado As String) As Integer




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
    
    Select Case INTRETORNO

        Case 0
            'MsgBox "Erro de comunica��o com a impressora.", vbOKOnly + vbCritical, TituloJanela
            'RETORNO_ECF = Daruma_FI_AbrePortaSerial
            'RETORNO_ECF = Daruma_FI_FechaPortaSerial
            GoTo entra
        Case 1
entra:
            RETORNOSTATUS = Daruma_FI_RetornoImpressora(Ack, Int_St1, Int_St2)
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
    INTRETORNO = Daruma_TEF_TravarTeclado(1)
    cArquivoTemp = Dir(App.Path & "\IMPRIME.TXT")
    If cArquivoTemp <> "" Then
        INTRETORNO = Daruma_FI_AbreComprovanteNaoFiscalVinculado(cFormaPGTO, cValorPago, cCOO)
        VerificaRetornoImpressoraDaruma "", Trim(INTRETORNO), "Imprimindo Transa��o TEF"
    End If
    cConteudo = ""
    cLinha = ""
    
    Open App.Path & "\IMPRIME.TXT" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cLinha
        cConteudo = cConteudo + cLinha + Chr(13) + Chr(10)
        INTRETORNO = Daruma_FI_UsaComprovanteNaoFiscalVinculado(cLinha + Chr(13))
        VerificaRetornoImpressoraDaruma "", Trim(INTRETORNO), "Imprimindo Transa��o TEF"
        If EOF(1) Then
            cSaltaLinha = Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10)
            INTRETORNO = Daruma_FI_UsaComprovanteNaoFiscalVinculado(cSaltaLinha)
            VerificaRetornoImpressoraDaruma "", Trim(INTRETORNO), "Imprimindo Transa��o TEF"
            ' Est� sendo usado um form para a exibi��o desta mensagem
            frmMensagem.lblMensagem.Caption = "Por favor, destaque a 1� Via"
            frmMensagem.Show
            frmMensagem.Refresh
            Sleep (5000)
            Unload frmMensagem
            frmPrincipal.Refresh
            INTRETORNO = Daruma_FI_UsaComprovanteNaoFiscalVinculado(cConteudo)
            VerificaRetornoImpressoraDaruma "", Trim(INTRETORNO), "Imprimindo Transa��o TEF"
        End If
    Loop
    ' Desbloqeia o teclado e o mouse
    INTRETORNO = Daruma_TEF_TravarTeclado(0)
    Close #1
    Kill App.Path & "\IMPRIME.TXT"
    INTRETORNO = Daruma_FI_FechaComprovanteNaoFiscalVinculado()
    VerificaRetornoImpressoraDaruma "", Trim(INTRETORNO), "Imprimindo Transa��o TEF"
    
    Exit Function
ERRO_TRATA:
    'If Err.Number <> 0 Then f_TrataErro
End Function
