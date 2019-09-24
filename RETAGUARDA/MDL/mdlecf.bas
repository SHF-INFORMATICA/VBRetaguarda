Attribute VB_Name = "mdlECF"
'================== BEMATECH
'================== NORMAL
Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BEMAFI32.DLL" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BEMAFI32.DLL" (ByVal NUMEROCUPOM As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BEMAFI32.DLL" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BEMAFI32.DLL" (ByVal DESCRICAO As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BEMAFI32.DLL" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BEMAFI32.DLL" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BEMAFI32.DLL" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BEMAFI32.DLL" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NUMEROCUPOM As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_Sangria Lib "BEMAFI32.DLL" (ByVal Valor As String) As Integer
Public Declare Function Bematech_FI_Suprimento Lib "BEMAFI32.DLL" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmware Lib "BEMAFI32.DLL" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CGC_IE Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Bematech_FI_GrandeTotal Lib "BEMAFI32.DLL" (ByVal GrandeTotal As String) As Integer
Public Declare Function Bematech_FI_Cancelamentos Lib "BEMAFI32.DLL" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BEMAFI32.DLL" (ByVal ValorDescontos As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BEMAFI32.DLL" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Bematech_FI_NumeroCuponsCancelados Lib "BEMAFI32.DLL" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Bematech_FI_NumeroIntervencoes Lib "BEMAFI32.DLL" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Bematech_FI_NumeroReducoes Lib "BEMAFI32.DLL" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Bematech_FI_NumeroSubstituicoesProprietario Lib "BEMAFI32.DLL" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_ClicheProprietario Lib "BEMAFI32.DLL" (ByVal Cliche As String) As Integer
Public Declare Function Bematech_FI_NumeroCaixa Lib "BEMAFI32.DLL" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Bematech_FI_NumeroLoja Lib "BEMAFI32.DLL" (ByVal NumeroLoja As String) As Integer
Public Declare Function Bematech_FI_SimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_MinutosLigada Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_MinutosImprimindo Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_VerificaModoOperacao Lib "BEMAFI32.DLL" (ByVal Modo As String) As Integer
Public Declare Function Bematech_FI_VerificaEpromConectada Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BEMAFI32.DLL" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BEMAFI32.DLL" (ByVal ValorCupom As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BEMAFI32.DLL" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_Acrescimos Lib "BEMAFI32.DLL" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Bematech_FI_ContadorBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Bematech_FI_VerificaAliquotasIss Lib "BEMAFI32.DLL" (ByVal AliquotasIss As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamento Lib "BEMAFI32.DLL" (ByVal Formas As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaDepartamentos Lib "BEMAFI32.DLL" (ByVal Departamentos As String) As Integer
Public Declare Function Bematech_FI_VerificaTipoImpressora Lib "BEMAFI32.DLL" (ByRef tipoImpressora As Integer) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciais Lib "BEMAFI32.DLL" (ByVal cTotalizadores As String) As Integer
Public Declare Function Bematech_FI_RetornoAliquotas Lib "BEMAFI32.DLL" (ByVal cAliquotas As String) As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BEMAFI32.DLL" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BEMAFI32.DLL" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_MonitoramentoPapel Lib "BEMAFI32.DLL" (ByRef Linhas As Integer) As Integer
Public Declare Function Bematech_FI_Autenticacao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BEMAFI32.DLL" (ByVal Parametros As String) As Integer
Public Declare Function Bematech_FI_AcionaGaveta Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoGaveta Lib "BEMAFI32.DLL" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaSingular Lib "BEMAFI32.DLL" (ByVal MoedaSingular As String) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaPlural Lib "BEMAFI32.DLL" (ByVal MoedaPlural As String) As Integer
Public Declare Function Bematech_FI_CancelaImpressaoCheque Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaStatusCheque Lib "BEMAFI32.DLL" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Bematech_FI_ImprimeCheque Lib "BEMAFI32.DLL" (ByVal Banco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_ImprimeCopiaCheque Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_IncluiCidadeFavorecido Lib "BEMAFI32.DLL" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Public Declare Function Bematech_FI_EstornoFormasPagamento Lib "BEMAFI32.DLL" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
Public Declare Function Bematech_FI_ForcaImpactoAgulhas Lib "BEMAFI32.DLL" (ByVal ForcaImpacto As Integer) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BEMAFI32.DLL" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoDescricaoForma Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_AbreBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal PREFIXO As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
Public Declare Function Bematech_FI_MapaResumo Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Analitico Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Mestre Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeConfiguracoesImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AberturaDoDia Lib "BEMAFI32.DLL" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_FechamentoDoDia Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ValorFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Bematech_FI_ValorTotalizadorNaoFiscal Lib "BEMAFI32.DLL" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer
Public Declare Function Bematech_FI_DadosSintegra Lib "BEMAFI32.DLL" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Public Declare Function Bematech_FI_RegistrosTipo60 Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoIndice Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_VerificaReducaoZAutomatica Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer

'Funções para Impressora restaurante
Public Declare Function Bematech_FIR_RegistraVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_CancelaVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_ConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_AbreConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_FechaConferenciaMesa Lib "BEMAFI32.DLL" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaMesa Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_AbreCupomRestaurante Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_ContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaItem Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal codigo As String, ByVal DESCRICAO As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertas Lib "BEMAFI32.DLL" (ByVal TipoRelatorio As Integer) As Integer
Public Declare Function Bematech_FIR_ImprimeCardapio Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertasSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_CardapioPelaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_RegistroVendaSerial Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_VerificaMemoriaLivre Lib "BEMAFI32.DLL" (ByVal Bytes As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomRestaurante Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomResumidoRestaurante Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer

' Funções da Impressora Fiscal MFD
Public Declare Function Bematech_FI_SubTotalizaCupomMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_SubTotalizaRecebimentoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbreBilhetePassagemMFD Lib "BEMAFI32.DLL" (ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Agencia As String, ByVal Data As String, ByVal hora As String, ByVal Poltrona As String, ByVal Plataforma As String, ByVal TipoPassagem As String, ByVal RG As String, ByVal NOME As String, ByVal Endereco As String, ByVal UFDestino As String) As Integer
Public Declare Function Bematech_FI_AbreCupomMFD Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Bematech_FI_CancelaAcrescimoDescontoItemMFD Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal item As String) As Integer
Public Declare Function Bematech_FI_CancelaAcrescimoDescontoSubtotalMFD Lib "BEMAFI32.DLL" (ByVal cFlag As String) As Integer
Public Declare Function Bematech_FI_CancelaCupomMFD Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Bematech_FI_ProgramaFormaPagamentoMFD Lib "BEMAFI32.DLL" (ByVal FormaPagto As String, ByVal OperacaoTef As String) As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupomMFD Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimo As String, ByVal ValorDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoMFD Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal Parcelas As String, ByVal DescricaoFormaPagto As String) As Integer
Public Declare Function Bematech_FI_CupomAdicionalMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AcrescimoDescontoItemMFD Lib "BEMAFI32.DLL" (ByVal item As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_AcrescimoDescontoSubtotalMFD Lib "BEMAFI32.DLL" (ByVal cFlag As String, ByVal cTipo As String, ByVal cValor As String) As Integer
Public Declare Function Bematech_FI_NomeiaRelatorioGerencialMFD Lib "BEMAFI32.DLL" (ByVal Indice As String, ByVal DESCRICAO As String) As Integer
Public Declare Function Bematech_FI_AutenticacaoMFD Lib "BEMAFI32.DLL" (ByVal Linhas As String, ByVal texto As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculadoMFD Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NUMEROCUPOM As String, ByVal CGC As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Bematech_FI_ReimpressaoNaoFiscalVinculadoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbreRecebimentoNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Bematech_FI_EfetuaRecebimentoNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal IndiceTotalizador As String, ByVal ValorRecebimento As String) As Integer
Public Declare Function Bematech_FI_IniciaFechamentoRecebimentoNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimo As String, ByVal ValorDesconto As String) As Integer
Public Declare Function Bematech_FI_FechaRecebimentoNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_CancelaRecebimentoNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Bematech_FI_AbreRelatorioGerencialMFD Lib "BEMAFI32.DLL" (ByVal Indice As String) As Integer
Public Declare Function Bematech_FI_UsaRelatorioGerencialMFD Lib "BEMAFI32.DLL" (ByVal texto As String) As Integer
Public Declare Function Bematech_FI_SegundaViaNaoFiscalVinculadoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_EstornoNaoFiscalVinculadoMFD Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal NOME As String, ByVal Endereco As String) As Integer
Public Declare Function Bematech_FI_NumeroSerieMFD Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmwareMFD Lib "BEMAFI32.DLL" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CNPJMFD Lib "BEMAFI32.DLL" (ByVal CNPJ As String) As Integer
Public Declare Function Bematech_FI_InscricaoEstadualMFD Lib "BEMAFI32.DLL" (ByVal InscricaoEstadual As String) As Integer
Public Declare Function Bematech_FI_InscricaoMunicipalMFD Lib "BEMAFI32.DLL" (ByVal InscricaoMunicipal As String) As Integer
Public Declare Function Bematech_FI_TempoOperacionalMFD Lib "BEMAFI32.DLL" (ByVal TempoOperacional As String) As Integer
Public Declare Function Bematech_FI_MinutosEmitindoDocumentosFiscaisMFD Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscaisMFD Lib "BEMAFI32.DLL" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscaisMFD Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamentoMFD Lib "BEMAFI32.DLL" (ByVal FormasPagamento As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaRelatorioGerencialMFD Lib "BEMAFI32.DLL" (ByVal Relatorios As String) As Integer
Public Declare Function Bematech_FI_ContadorComprovantesCreditoMFD Lib "BEMAFI32.DLL" (ByVal Comprovantes As String) As Integer
Public Declare Function Bematech_FI_ContadorOperacoesNaoFiscaisCanceladasMFD Lib "BEMAFI32.DLL" (ByVal OperacoesCanceladas As String) As Integer
Public Declare Function Bematech_FI_ContadorRelatoriosGerenciaisMFD Lib "BEMAFI32.DLL" (ByVal Relatorios As String) As Integer
Public Declare Function Bematech_FI_ContadorCupomFiscalMFD Lib "BEMAFI32.DLL" (ByVal CuponsEmitidos As String) As Integer
Public Declare Function Bematech_FI_ContadorFitaDetalheMFD Lib "BEMAFI32.DLL" (ByVal ContadorFita As String) As Integer
Public Declare Function Bematech_FI_ComprovantesNaoFiscaisNaoEmitidosMFD Lib "BEMAFI32.DLL" (ByVal Comprovantes As String) As Integer
Public Declare Function Bematech_FI_NumeroSerieMemoriaMFD Lib "BEMAFI32.DLL" (ByVal NumeroSerieMFD As String) As Integer
Public Declare Function Bematech_FI_ReducoesRestantesMFD Lib "BEMAFI32.DLL" (ByVal Reducoes As String) As Integer
Public Declare Function Bematech_FI_MarcaModeloTipoImpressoraMFD Lib "BEMAFI32.DLL" (ByVal Marca As String, ByVal Modelo As String, ByVal TIPO As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciaisMFD Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducaoMFD Lib "BEMAFI32.DLL" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalDataMFD Lib "BEMAFI32.DLL" (ByVal DataInicial As String, ByVal DataFinal As String, ByVal FlagLeitura As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducaoMFD Lib "BEMAFI32.DLL" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String, ByVal FlagLeitura As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialDataMFD Lib "BEMAFI32.DLL" (ByVal DataInicial As String, ByVal DataFinal As String, ByVal FlagLeitura As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducaoMFD Lib "BEMAFI32.DLL" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String, ByVal FlagLeitura As String) As Integer
Public Declare Function Bematech_FI_LeituraChequeMFD Lib "BEMAFI32.DLL" (ByVal CodigoCMC7 As String) As Integer
Public Declare Function Bematech_FI_ImprimeChequeMFD Lib "BEMAFI32.DLL" (ByVal NumeroBanco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String, ByVal ImpressaoVerso As String, ByVal Linhas As String) As Integer
Public Declare Function Bematech_FI_HabilitaDesabilitaRetornoEstendidoMFD Lib "BEMAFI32.DLL" (ByVal FlagRetorno As String) As Integer
Public Declare Function Bematech_FI_RetornoImpressoraMFD Lib "BEMAFI32.DLL" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer, ByRef ST3 As Integer) As Integer
Public Declare Function Bematech_FI_TotalLivreMFD Lib "BEMAFI32.DLL" (ByVal cMemoriaLivre As String) As Integer
Public Declare Function Bematech_FI_TamanhoTotalMFD Lib "BEMAFI32.DLL" (ByVal cTamMFD As String) As Integer
Public Declare Function Bematech_FI_AcrescimoDescontoSubtotalRecebimentoMFD Lib "BEMAFI32.DLL" (ByVal cFlag As String, ByVal cTipo As String, ByVal cValor As String) As Integer
Public Declare Function Bematech_FI_CancelaAcrescimoDescontoSubtotalRecebimentoMFD Lib "BEMAFI32.DLL" (ByVal cFlag As String) As Integer
Public Declare Function Bematech_FI_TotalizaCupomMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_TotalizaRecebimentoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_PercentualLivreMFD Lib "BEMAFI32.DLL" (ByVal cMemoriaLivre As String) As Integer
Public Declare Function Bematech_FI_DataHoraUltimoDocumentoMFD Lib "BEMAFI32.DLL" (ByVal cDataHora As String) As Integer
Public Declare Function Bematech_FI_ValorFormaPagamentoMFD Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Bematech_FI_ValorTotalizadorNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressoraMFD Lib "BEMAFI32.DLL" (ByRef Ack As Integer, ByRef st1 As Integer, ByRef st2 As Integer, ByRef ST3 As Integer) As Integer
Public Declare Function Bematech_FI_MapaResumoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60AnaliticoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupomCodigoBarrasMFD Lib "BEMAFI32.DLL" (ByVal Mensagem As String, ByVal TipoCodigo As String, ByVal codigo As String, ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer, ByVal CorrecaoErros As Integer, ByVal Colunas As Integer) As Integer
Public Declare Function Bematech_FI_DownloadMF Lib "BEMAFI32.DLL" (ByVal Arquivo As String) As Integer
Public Declare Function Bematech_FI_DownloadMFD Lib "BEMAFI32.DLL" (ByVal Arquivo As String, ByVal TipoDownload As String, ByVal ParametroInicial As String, ByVal ParametroFinal As String, ByVal UsuarioECF As String) As Integer
Public Declare Function Bematech_FI_FormatoDadosMFD Lib "BEMAFI32.DLL" (ByVal ArquivoMFD As String, ByVal ArquivoDestino As String, ByVal FormatoDestino As String, ByVal TipoDownload As String, ByVal ParametroInicial As String, ByVal ParametroFinal As String, ByVal UsuarioECF As String) As Integer

' Funções disponíveis somente na impressora fiscal MP-2000 TH FI versão 01.01.01 ou 01.00.02
Public Declare Function Bematech_FI_AtivaDesativaVendaUmaLinhaMFD Lib "BEMAFI32.DLL" (ByVal iFlag As Integer) As Integer
Public Declare Function Bematech_FI_AtivaDesativaAlinhamentoEsquerdaMFD Lib "BEMAFI32.DLL" (ByVal iFlag As Integer) As Integer
Public Declare Function Bematech_FI_AtivaDesativaCorteProximoMFD Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AtivaDesativaTratamentoONOFFLineMFD Lib "BEMAFI32.DLL" (ByVal iFlag As Integer) As Integer
Public Declare Function Bematech_FI_StatusEstendidoMFD Lib "BEMAFI32.DLL" (ByRef iStatus As Integer) As Integer
Public Declare Function Bematech_FI_VerificaFlagCorteMFD Lib "BEMAFI32.DLL" (ByRef iFlag As Integer) As Integer
Public Declare Function Bematech_FI_TempoRestanteComprovanteMFD Lib "BEMAFI32.DLL" (ByVal cTempo As String) As Integer
Public Declare Function Bematech_FI_UFProprietarioMFD Lib "BEMAFI32.DLL" (ByVal cUF As String) As Integer
Public Declare Function Bematech_FI_GrandeTotalUltimaReducaoMFD Lib "BEMAFI32.DLL" (ByVal cGT As String) As Integer
Public Declare Function Bematech_FI_DataMovimentoUltimaReducaoMFD Lib "BEMAFI32.DLL" (ByVal cData As String) As Integer
Public Declare Function Bematech_FI_SubTotalComprovanteNaoFiscalMFD Lib "BEMAFI32.DLL" (ByVal cSubTotal As String) As Integer
Public Declare Function Bematech_FI_InicioFimCOOsMFD Lib "BEMAFI32.DLL" (ByVal cCOOIni As String, ByVal cCOOFim As String) As Integer
Public Declare Function Bematech_FI_InicioFimGTsMFD Lib "BEMAFI32.DLL" (ByVal cGTIni As String, ByVal cGTFim As String) As Integer

' utilidades
Public Declare Function Bematech_FI_InfoBalanca Lib "BEMAFI32.DLL" (ByVal port As String, ByVal model As Integer, ByVal weight As String, ByVal precoKilo As String, ByVal total As String) As Integer
Public Declare Function Bematech_FI_ImpressaoCarne Lib "BEMAFI32.DLL" (ByVal titulo As String, ByVal parcela As String, ByVal datas As String, ByVal Quantidade As Integer, ByVal texto As String, _
                                                ByVal Cliente As String, ByVal rgcpf As String, _
                                                ByVal cupom As String, ByVal vias As Integer, ByVal assina As Integer) As Integer

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpAplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public RETORNO_ECF As Integer
Public Funcao  As Integer

Public LocalRetorno As String
Public ArqRetorno As String
Public RetornoEstendidoHabilitado As Boolean
Public Desconto As String, Acrescimo As String, strIndiceDepartamento As String, strUnidadeMedida As String

'' utilidades
'Public Declare Function Bematech_FI_InfoBalanca Lib "BEMAFI32.DLL" (ByVal port As String, ByVal model As Integer, ByVal weight As String, ByVal precoKilo As String, ByVal total As String) As Integer
'Public Declare Function Bematech_FI_ImpressaoCarne Lib "BEMAFI32.DLL" (ByVal titulo As String, ByVal parcela As String, ByVal datas As String, ByVal quantidade As Integer, ByVal texto As String, _
'                                                ByVal cliente As String, ByVal rgcpf As String, _
'                                                ByVal cupom As String, ByVal vias As Integer, ByVal assina As Integer) As Integer
                                                  
Public Declare Function Bematech_FI_IniciaModoTEF Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FinalizaModoTEF Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AcionaGuilhotinaMFD Lib "BEMAFI32.DLL" (ByVal TipoCorte As Integer) As Integer
Public Declare Function BlockInput Lib "user32.dll" (ByVal Blk As Boolean) As Boolean

'================= NFc-e   BemaNFCe32
Public Declare Function Bematech_NFCe_AbreNota Lib "BemaNFCe32.DLL" (ByVal CPF_CNPJ_ID As String, ByVal cSerie As String, ByVal cNF As String) As Integer
Public Declare Function Bematech_NFCe_AcrescimoDescontoItem Lib "BemaNFCe32.DLL" (ByVal item As String, ByVal incrementTotalValue As String, ByVal discountTotalValue As String, ByVal newNetValue As String, ByVal newBasisCalculation As String, ByVal newTaxValue) As Integer
Public Declare Function Bematech_NFCe_CancelaFormaPagamento Lib "BemaNFCe32.DLL" (ByVal paymentSequence As String) As Integer
Public Declare Function Bematech_NFCe_CancelaItem Lib "BemaNFCe32.DLL" (ByVal item As String) As Integer
Public Declare Function Bematech_NFCe_CancelaNota Lib "BemaNFCe32.DLL" (ByVal serie As String, ByVal nf As String) As Integer
Public Declare Function Bematech_NFCe_DadosEmissaoNFCe Lib "BemaNFCe32.DLL" (ByVal UFCode As String, ByVal environmentType As String, ByVal emissionProcess As String, ByVal processVersion As String, ByVal CRT As String, ByVal timezone As String) As Integer
Public Declare Function Bematech_NFCe_DadosEmissor Lib "BemaNFCe32.DLL" (ByVal CNPJ As String, ByVal name As String, ByVal tradeName As String, ByVal address As String, ByVal number As String, ByValneighborhood As String, ByVal IBGECode As String, ByVal city As String, ByVal UF As String, ByVal CEP As String, ByVal countryCode As String, ByValcountry As String, ByVal phone As String, ByVal stateRegistration As String, ByVal stateRegistrationST As String, ByVal municipalRegistration As String) As Integer
Public Declare Function Bematech_NFCe_DadosConsumidor Lib "BemaNFCe32.DLL" (ByVal CPF_CNPJ_ID As String, ByVal name As String, ByVal address As String, ByVal complement As String, ByVal number As String, ByVal neighborhood As String, ByVal IBGECode As String, ByVal city As String, ByVal UF As String, ByVal CEP As String, ByVal countryCode As String, ByVal country As String, ByVal phone As String, ByVal stateRegistrationIndex As String, ByVal stateRegistration As String, ByVal SUFRAMACode As String, ByVal email As String) As Integer
Public Declare Function Bematech_NFCe_EfetuaFormaPagamento Lib "BemaNFCe32.DLL" (ByVal paymentFormIndex As String, ByVal Value As String) As Integer
Public Declare Function Bematech_NFCe_EfetuaFormaPagamentoCredenciadora Lib "BemaNFCe32.DLL" (ByVal paymentFormIndex As String, ByVal Value As String, ByVal licensingCNPJ As String, ByVal licensingCode As String, ByVal authorizationCode As String, ByVal integrationCode As String) As Integer
Public Declare Function Bematech_NFCe_FechaNota Lib "BemaNFCe32.DLL" (ByVal promotionalMessage As String, ByVal changeValue As String, ByVal taxValue As String, ByVal DANFELayout As String, ByVal DANFEOut As String, ByVal email As String) As Integer
Public Declare Function Bematech_NFCe_ImprimeTextoLivre Lib "BemaNFCe32.DLL" (ByVal filename As String) As Integer
Public Declare Function Bematech_NFCe_InsereTributacaoCOFINS Lib "BemaNFCe32.DLL" (ByVal item As String, ByVal CST_COFINS As String, ByVal COFINSBasisCalculation As String, ByVal COFINSTax As String, ByVal COFINSValue As String, ByVal COFINSQuantitySell As String, ByVal COFINSTaxValue As String, ByVal COFINSIncidentTaxValue As String) As Integer
Public Declare Function Bematech_NFCe_InsereTributacaoICMS Lib "BemaNFCe32.DLL" (ByVal item As String, ByVal CST_ICMS As String, ByVal basisCalculationMode As String, ByVal basisCalculationReductionPercentual As String, ByVal basisCalculationValue As String, ByVal tax As String, ByVal taxValue As String, ByVal ICMSSTBasisCalculationMode As String, ByVal ICMSSTValueAddedMarginPercentual As String, ByVal ICMSSTBasisCalculationReductionPercentual As String, ByVal ICMSSTBasisCalculationReductionValue As String, ByVal ICMSSTTax As String, ByVal ICMSSTValue As String, ByVal basisCalculationValueRetained As String, ByVal ICMSValueRetained As String, ByVal ICMSUnencumberedValue As String, ByVal ICMSUnburdeningMotive As String, ByVal incidentTaxTotalValue As String) As Integer
Public Declare Function Bematech_NFCe_InsereTributacaoSIMPLES Lib "BemaNFCe32.DLL" (ByVal item As String, ByVal CSOSN As String, ByVal basisCalculationMode As String, ByVal basisCalculationReductionPercentual As String, ByVal basisCalculationValue As String, ByVal tax As String, ByVal taxValue As String, ByVal ICMSSTBasisCalculationMode As String, ByVal ICMSSTValueAddedMarginPercentual As String, ByVal ICMSSTBasisCalculationReductionPercentual As String, ByVal ICMSSTBasisCalculationReductionValue As String, ByVal ICMSSTTax As String, ByVal ICMSSTValue As String, ByVal basisCalculationValueRetained As String, ByVal ICMSValueRetained As String, ByVal creditCalculationApplicableTax As String, ByVal ICMSSNCreditValue As String, ByVal incidentTaxTotalValue) As Integer
Public Declare Function Bematech_NFCe_InsereTributacaoPIS Lib "BemaNFCe32.DLL" (ByVal item As String, ByVal CST_PIS As String, ByVal PISBasisCalculation As String, ByVal PISTax As String, ByVal PISValue As String, ByVal PISQuantitySell As String, ByVal PISTaxValue As String, ByVal PISIncidentTaxValue) As Integer
Public Declare Function Bematech_NFCe_InutilizaNota Lib "BemaNFCe32.DLL" (ByVal serie As String, ByVal nf As String, ByVal reason As String) As Integer
Public Declare Function Bematech_NFCe_ReimprimeDANFEChave Lib "BemaNFCe32.DLL" (ByVal accessKey As String) As Integer
Public Declare Function Bematech_NFCe_ReimprimeDANFE Lib "BemaNFCe32.DLL" (ByVal serie As String, ByVal nf As String) As Integer
Public Declare Function Bematech_NFCe_StatusInutilizaNota Lib "BemaNFCe32.DLL" (ByVal serie As String, ByVal nf As String, ByVal SEFAZReturnCode As String, ByVal protocol As String, ByVal dateHourProtocol) As Integer
Public Declare Function Bematech_NFCe_StatusNFCe Lib "BemaNFCe32.DLL" (ByVal serie As String, ByVal nf As String, ByVal SEFAZReturnCode As String, ByVal keyAccess As String, ByVal protocol As String, ByVal dateHourProtocol As String) As Integer
Public Declare Function Bematech_NFCe_StatusUltimaNFCe Lib "BemaNFCe32.DLL" (ByVal serie As String, ByVal nf As String, ByVal SEFAZReturnCode As String, ByVal keyAccess As String, ByVal protocol As String, ByVal dateHourProtocol As String) As Integer
Public Declare Function Bematech_NFCe_VendeItem Lib "BemaNFCe32.DLL" (ByVal code As String, ByVal EAN13 As String, ByVal description As String, ByVal NCM As String, ByVal CFOP As String, ByVal unitOfMeasure As String, ByVal quantity As String, ByVal decimalsQuantity As String, ByVal unitaryValue As String, ByVal decimalsUnitaryValue As String, ByVal grossValue As String, ByVal incrementValue As String, ByVal discountValue As String, ByVal netValue As String, ByVal productOrigin As String, ByVal additionalInformation As String) As Integer
Public Declare Function Bematech_NFCe_VerificaNotaAberta Lib "BemaNFCe32.DLL" (ByVal Status As String) As Integer
'=======================
Public Declare Function Bematech_FI_AdicionaInformacoesCombustivel Lib "BEMAFI32.DLL" (ByVal itemIndex As String, ANPProductCode As String, ByVal percentMixGN As String, ByVal CODIF As String, ByVal quantity As String, ByVal consumeUF As String, ByVal BCProductCIDE As String, ByVal taxProductCIDE As String, ByVal valueCIDE As String, ByVal fuelNozzleNumber As String, ByVal fuelPumpNumber As String, ByVal fuelTankNumber As String, ByVal fuelGaugeInitial As String, ByVal fuelGaugeFinal As String) As Integer
Public Declare Function Bematech_FI_ChaveAcessoNFCe Lib "BEMAFI32.DLL" (ByVal Index As String, ByVal counter As String, ByRef accessKey As String) As Integer
Public Declare Function Bematech_FI_DadosConsumidorNFCe Lib "BEMAFI32.DLL" (ByVal CPF As String, ByVal name As String, ByVal address As String, ByVal complement As String, ByVal number As String, ByVal neighborhood As String, ByVal IBGECode As String, ByVal city As String, ByVal UF As String, ByVal CEP As String, ByVal countyCode As String, ByVal country As String, ByVal phone As String, ByVal stateRegistrationIndicator As String, ByVal stateRegistration As String, ByVal SUFRAMACode As String, ByVal email As String) As Integer
Public Declare Function Bematech_FI_DadosEnvioNFCe Lib "BEMAFI32.DLL" (ByVal TipoLayout As String, ByVal TipoEmissao As String, ByVal cEmail As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoNFCeEx Lib "BEMAFI32.DLL" (ByVal descBandeira As String, ByVal ValorForma As String, ByVal CNPJCrede As String, ByVal bandeira As String, ByVal CodAuto As String, ByVal CodIntegra As String) As Integer
Public Declare Function Bematech_FI_NumeroNotaNFCe Lib "BEMAFI32.DLL" (ByVal noteNumber As String) As Integer
Public Declare Function Bematech_FI_NumeroSerieNFCe Lib "BEMAFI32.DLL" (ByVal serialNumber As String) As Integer
Public Declare Function Bematech_FI_ProgramaContadorNFCe Lib "BEMAFI32.DLL" (ByVal Index As String, ByVal counter As String) As Integer
Public Declare Function Bematech_FI_ProtocoloUltimaNFCe Lib "BEMAFI32.DLL" (ByVal protocol As String, ByVal datehour As String) As Integer
Public Declare Function Bematech_FI_RetornaInformacoesNFCe Lib "BEMAFI32.DLL" (ByVal paramType As String, ByVal paramValue As String, ByVal retChaveAcesso As String, ByVal retSerie As String, ByVal retNumNFCe As String, ByVal retCancelled As String, ByVal retSendStatus As String, ByVal retSendProtocol As String, ByVal retSendProtocolDatetime As String, ByVal retCancellationStatus As String, ByVal retCancellationProtocol As String) As Integer
Public Declare Function Bematech_FI_StatusUltimaNFCe Lib "BEMAFI32.DLL" (ByVal Status As String) As Integer
Public Declare Function Bematech_FI_StatusUltimoCancelamentoNFCe Lib "BEMAFI32.DLL" (ByVal Status As String) As Integer
Public Declare Function Bematech_FI_UltimaChaveAcessoNFCe Lib "BEMAFI32.DLL" (ByVal accessKey As String) As Integer
Public Declare Function Bematech_FI_VendeItemCompleto Lib "BEMAFI32.DLL" (ByVal sParametros As String) As Integer
Public Declare Function Bematech_FI_VendeItemCompletoJSON Lib "BEMAFI32.DLL" (ByVal sParametros As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupomNFCe Lib "BEMAFI32.DLL" (ByVal Mensagem As String, ByVal Taxas As String) As Integer
'=======================
Public Sub TRATA_ECF()
'On Error GoTo ERRO_TRATA

   If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And USUARIO_ID_N = 144) Then
      Dim NumeroIntervencao   As String
      Dim LocalRetorno        As String
      Dim NumeroSerie         As String
      Dim NumeroCaixa         As String

      IMPRESSORA_ID_N = 0
      NUMERO_CAIXA_ECF = 0
      CONTA_REINICIO_N = 0
      NUMERO_SERIE_ECF = ""

'=====================================contador reinicio
      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NumeroIntervencao = Space(1)
         Else: NumeroIntervencao = Space(4)
      End If

      CONTA_REINICIO_N = 0
      If USA_NFC_E = False Then
         RETORNO_ECF = Bematech_FI_NumeroIntervencoes(NumeroIntervencao)

         If Trim(NumeroIntervencao) <> "" Then _
            CONTA_REINICIO_N = NumeroIntervencao

         If CONTA_REINICIO_N <= 0 Then _
            MsgBox "Erro ao ler contador reinicio da impressora fiscal, verificar !!!"
      End If
'=====================================SERIE IMPRESSORA FISCAL
      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NumeroSerie = Space(1)
         Else: NumeroSerie = Space(20)
      End If

      NUMERO_SERIE_ECF = ""
      If USA_NFC_E = False Then
         RETORNO_ECF = Bematech_FI_NumeroSerieMFD(NumeroSerie)
         NUMERO_SERIE_ECF = CStr(NumeroSerie)
         If Trim(NUMERO_SERIE_ECF) = "" Then _
            MsgBox "Erro ao ler número de série da impressora da impressora fiscal, verificar !!!"
         Else: NUMERO_SERIE_ECF = "NFC-e"
      End If
'=================================================numero caixa
      NumeroCaixa = Space(4)
      If USA_NFC_E = False Then
         RETORNO_ECF = Bematech_FI_NumeroCaixa(NumeroCaixa)
         NUMERO_CAIXA_ECF = 0 & RETORNO_ECF
      End If
      If Left(UCase(NUMERO_SERIE_ECF), 8) = UCase("emulador") Then _
         NUMERO_SERIE_ECF = "EMULADOR"

      NUMERO_SERIE_ECF = CaracteresValidos(NUMERO_SERIE_ECF)

      If USA_NFC_E = False Then
         If Len(NUMERO_SERIE_ECF) < 20 Then _
            NUMERO_SERIE_ECF = NUMERO_SERIE_ECF & NUMERO_CAIXA_CPU

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from IMPRESSORA "
         SQL = SQL & " where EMPRESA_ID = " & EMPRESA_ID_N
         'SQL = SQL & " and NUMR_CAIXA = " & NUMERO_CAIXA_ECF
         'SQL = SQL & " and CONTA_REINICIO = " & CONTA_REINICIO_N
         SQL = SQL & " and NUMR_SERIE_IMP = '" & Trim(NUMERO_SERIE_ECF) & "'"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            IMPRESSORA_ID_N = 0 & TabConsulta.Fields("IMPRESSORA_ID").Value

            If CONTA_REINICIO_N <> TabConsulta.Fields("CONTA_REINICIO").Value Then
               MsgBox "Numero contador diferente do cadastrado. Vá no menu (ECF/Parametros Impressora Fiscal) e atualize."
               CONTA_REINICIO_N = TabConsulta.Fields("CONTA_REINICIO").Value
            End If

            If Trim(TabConsulta.Fields("NUMR_SERIE_IMP").Value) <> "EMULADOR" Then _
               If Trim(TabConsulta.Fields("NUMR_SERIE_IMP").Value) <> Trim(NUMERO_SERIE_ECF) Then _
                  MsgBox "Número de Serial impressora fiscal diferente do registrado para essa base de dados." & NUMERO_SERIE_ECF
            Else: MsgBox "Impressora fiscal não cadastrada."
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, "mdlGeral", "TRATA_ECF"
End Sub

Public Function VerificaRetornoImpressora(Label As String, RetornoFuncao As String, TituloJanela As String) As Boolean
   Dim Ack As Integer
   Dim st1 As Integer
   Dim st2 As Integer
   Dim ST3 As Integer

   Dim RetornaMensagem As Integer
   Dim StringRetorno As String
   Dim ValorRetorno As String
   Dim RETORNOSTATUS As Integer
   Dim Mensagem As String

   VerificaRetornoImpressora = False
   Indr_Erro = False
   Indr_Cancela_Cupom = False
   INDR_CUPOM_ABERTO = False
   Msg = ""

   If RETORNO_ECF = 0 Then
      MsgBox "Erro de comunicação com a impressora. Atenção confira dados do cupom fiscal. " & CODG_PRODUTO_A, vbOKOnly + vbCritical, TituloJanela
CODG_PRODUTO_A = ""
      MOSTRA_RODAPE "Erro de comunicação com a impressora.", "", "", "", ""

      Indr_Erro = True

      Exit Function

   ElseIf RETORNO_ECF = 1 Or RETORNO_ECF = -27 Then
      If RetornoEstendidoHabilitado = True Then
         RETORNOSTATUS = Bematech_FI_RetornoImpressoraMFD(Ack, st1, st2, ST3)
         ValorRetorno = Str(Ack) & "," & Str(st1) & "," & Str(st2) & "," & Str(ST3)
         Else
            RETORNOSTATUS = Bematech_FI_RetornoImpressora(Ack, st1, st2)
            ValorRetorno = Str(Ack) & "," & Str(st1) & "," & Str(st2)
      End If

      If Label <> "" And RetornoFuncao <> "" Then _
         RetornaMensagem = 1

      If Ack = 21 Then

         If INDR_VENDA = False Then _
            MsgBox "Status da Impressora: 21 " & vbCr & vbLf & "Comando não executado", vbOKOnly + vbInformation, TituloJanela

         Indr_Erro = True

         Exit Function
      End If

      If (st1 <> 0 Or st2 <> 0) Then
         If (st1 >= 128) Then
            StringRetorno = "Fim de Papel" & vbCr
            MsgBox "Fim papel, colocar nova bobina."
            Indr_Erro = True
            Msg = Msg & " Fim de Papel"
            st1 = st1 - 128
         End If

         If (st1 >= 64) Then
            StringRetorno = StringRetorno & "Pouco Papel" & vbCr
            Msg = Msg & " Atenção, Pouco Papel"
            st1 = st1 - 64
         End If

         If (st1 >= 32) Then
            StringRetorno = StringRetorno & "Erro no relógio" & vbCr
            Indr_Erro = True
            Msg = Msg & " Erro no Relógio"
            st1 = st1 - 32
         End If

         If (st1 >= 16) Then
            StringRetorno = StringRetorno & "Impressora em erro" & vbCr
            Indr_Erro = True
            Msg = Msg & " Impressora em Erro"
            st1 = st1 - 16
            Indr_Cancela_Cupom = True
         End If

         If (st1 >= 8) Then
            StringRetorno = StringRetorno & "Primeiro dado do comando não foi Esc" & vbCr
            Indr_Erro = True
            Indr_Cancela_Cupom = True
            Indr_Cancela_Cupom = True
            st1 = st1 - 8
         End If

         If (st1 >= 4) Then
            StringRetorno = StringRetorno & "Comando inexistente" & vbCr
            Indr_Erro = True
            Msg = Msg & " Comando inexistente"
            st1 = st1 - 4
         End If

         If (st1 >= 2) Then
            StringRetorno = StringRetorno & "CUPOM FISCAL ABERTO." & vbCr
            Indr_Erro = True
            Msg = Msg & " Cupom Fiscal Aberto, verifique."
            st1 = st1 - 2
            Indr_Cancela_Cupom = True
            INDR_CUPOM_ABERTO = True
         End If

         If (st1 >= 1) Then
            StringRetorno = StringRetorno & "Número de parâmetros inválido no comando" & vbCr
            Indr_Erro = True
            Indr_Cancela_Cupom = True
            Msg = Msg & " Número de parâmetros inválido no comando"
            st1 = st1 - 1
         End If

         If (st2 >= 128) Then
            StringRetorno = "Tipo de Parâmetro de comando inválido" & vbCr
            Indr_Erro = True
            Indr_Cancela_Cupom = True
            Msg = Msg & " Tipo de parâmetro de comando inválido"
            st2 = st2 - 128
         End If

         If (st2 >= 64) Then
            StringRetorno = StringRetorno & "Memória fiscal lotada" & vbCr
            Indr_Erro = True
            Indr_Cancela_Cupom = True
            Msg = Msg & " Memória Fiscal Lotada"
            st2 = st2 - 64
         End If

         If (st2 >= 32) Then
            StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
            Indr_Erro = True
            Indr_Cancela_Cupom = True
            Msg = Msg & " Erro na CMOS"
            st2 = st2 - 32
         End If

         If (st2 >= 16) Then
            StringRetorno = StringRetorno & "Alíquota não programada" & vbCr
            Indr_Erro = True
            Indr_Cancela_Cupom = True
            Msg = Msg & " Alíquota não programada"
            st2 = st2 - 16
         End If

         If (st2 >= 8) Then
            StringRetorno = StringRetorno & "PEDIDOcidade de alíquota programáveis lotada" & vbCr
            Indr_Erro = True
            Msg = Msg & " PEDIDOcidade de alíquota programáveis lotada"
            st2 = st2 - 8
         End If

         If (st2 >= 4) Then
            StringRetorno = StringRetorno & "Cancelamento não permitido" & vbCr
            Indr_Erro = True
            Msg = Msg & " Cancelamento não permitido"
            st2 = st2 - 4
         End If

         If (st2 >= 2) Then
            StringRetorno = StringRetorno & "CGC/IE do proprietário não programados" & vbCr
            Indr_Erro = True
            Msg = Msg & " CGC/IE do proprietário não programados"
            st2 = st2 - 2
         End If

         If (st2 >= 1) Then
            StringRetorno = StringRetorno & "Comando não executado" & vbCr
            Indr_Erro = True
            Msg = Msg & " Comando não executado"
            st2 = st2 - 1
            'MsgBox Msg
         End If

         If RetornaMensagem Then
            Mensagem = "Status da Impressora: " & ValorRetorno & _
                       vbCr & vbLf & StringRetorno & vbCr & vbLf & _
                       Label & RetornoFuncao
            Else
               Mensagem = "Status da Impressora: " & ValorRetorno & _
               vbCr & vbLf & StringRetorno
         End If

         If INDR_VENDA = True Then
            If Trim(Msg) = "" Then _
               MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
            Else: MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
         End If

         Exit Function
      End If 'fim do ST1 <> 0 and ST2 <> 0

      If RetornaMensagem Then _
         Mensagem = Label & RetornoFuncao

      'If Mensagem <> "" Then _
         MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela

      VerificaRetornoImpressora = True

      Exit Function

   ElseIf RETORNO_ECF = -1 Then
       MsgBox "Erro de execução da função.", vbOKOnly + vbCritical, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Erro de execução da função"
       Exit Function

   ElseIf RETORNO_ECF = -2 Then
       MsgBox "Parâmetro inválido na função.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Parâmetro inválido na função"
       Exit Function

   ElseIf RETORNO_ECF = -3 Then
       MsgBox "Alíquota não programada.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Alíquota não programada"
       Exit Function

   ElseIf RETORNO_ECF = -4 Then
       MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default. " + vbCr + "Por favor, copie esse arquivo para o diretório de sistema do Windows." + vbCr + "Se for o Windows 95 ou 98 é o diretório 'System' se for o Windows NT é o diretório 'System32'.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default"
       Exit Function

   ElseIf RETORNO_ECF = -5 Then
       MsgBox "Erro ao abrir a porta de comunicação.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Erro ao abrir a porta de comunicação"
       Exit Function

   ElseIf RETORNO_ECF = -6 Then
       MsgBox "Impressora desligada ou cabo de comunicação desconectado.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Impressora desligada ou cabo de comunicação desconectado"
       Exit Function

   ElseIf RETORNO_ECF = -7 Then
       MsgBox "Banco não encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Banco não encontrado no arquivo BemaFI32.ini"
       Exit Function

   ElseIf RETORNO_ECF = -8 Then
       MsgBox "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Erro ao criar ou gravar no arquivo status.txt ou retorno.txt"
       Exit Function

   ElseIf RETORNO_ECF = -18 Then
       MsgBox "Não foi possível abrir arquivo INTPOS.001 !", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Não foi possível abrir arquivo INTPOS.001 !"
       Exit Function

   ElseIf RETORNO_ECF = -19 Then
       MsgBox "Parâmetro diferentes !", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Indr_Cancela_Cupom = True
       Msg = Msg & " Parâmetro diferentes !"
       Exit Function

   ElseIf RETORNO_ECF = -20 Then
       MsgBox "Transação cancelada pelo Operador !", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Transação cancelada pelo Operador !"
       Exit Function

   ElseIf RETORNO_ECF = -21 Then
       MsgBox "A Transação não foi aprovada !", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " A Transação não foi aprovada !"
       Exit Function

   ElseIf RETORNO_ECF = -22 Then
       MsgBox "Não foi possível terminal a Impressão !", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Não foi possível terminal a Impressão !"
       Exit Function

   ElseIf RETORNO_ECF = -23 Then
       MsgBox "Não foi possível terminal a Operação !", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Não foi possível terminal a Operação !"
       Exit Function

   ElseIf RETORNO_ECF = -24 Then
       MsgBox "Forma de pagamento não programada.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Forma de pagamento não programada"
       Exit Function

   ElseIf RETORNO_ECF = -25 Then
       MsgBox "Totalizador não fiscal não programado.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Totalizador não fiscal não programado"
       Exit Function

   ElseIf RETORNO_ECF = -26 Then
       MsgBox "Transação já realizada.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Transação já realizada"
       Exit Function

   ElseIf RETORNO_ECF = -28 Then
       MsgBox "Não há dados para serem impressos.", vbOKOnly + vbExclamation, TituloJanela
       Indr_Erro = True
       Msg = Msg & " Não há dados para serem impressos"
       Exit Function
   End If
   VerificaRetornoImpressora = True
End Function

Public Function AnalisaFlagsFiscais(FlagFiscal As Integer) As String
    Dim StringRetorno As String
    If (FlagFiscal >= 128) Then
        StringRetorno = "Memória fiscal lotada" & vbCr
        FlagFiscal = FlagFiscal - 128
    End If
    If (FlagFiscal >= 32) Then
        StringRetorno = StringRetorno & "Permite o cancelamento do cupom" & vbCr
        FlagFiscal = FlagFiscal - 32
    End If
    If (FlagFiscal >= 8) Then
        StringRetorno = StringRetorno & "Já houve redução 'Z' no dia" & vbCr
        FlagFiscal = FlagFiscal - 8
    End If
    If (FlagFiscal >= 4) Then
        StringRetorno = StringRetorno & "Horário de verão selecionado" & vbCr
        FlagFiscal = FlagFiscal - 4
    End If
    If (FlagFiscal >= 2) Then
        StringRetorno = StringRetorno & "Fechamento de formas de pagamento iniciado" & vbCr
        FlagFiscal = FlagFiscal - 2
    End If
    If (FlagFiscal >= 1) Then
        StringRetorno = StringRetorno & "Cupom fiscal Aberto" & vbCr
        FlagFiscal = FlagFiscal - 1
        INDR_CUPOM_ABERTO = True
    End If
    AnalisaFlagsFiscais = StringRetorno
End Function

Public Function AnalisaStatusCheque(StatusCheque As Integer) As String
    Dim StringRetorno As String

    If (StatusCheque = 1) Then
        StringRetorno = "Impressora ok." & vbCr
    ElseIf (StatusCheque = 2) Then
        StringRetorno = "Cheque em impressão." & vbCr
    ElseIf (StatusCheque = 3) Then
        StringRetorno = "Cheque posicionado." & vbCr
    ElseIf (StatusCheque = 4) Then
        StringRetorno = "Aguardando o posicionamento do cheque." & vbCr
    End If
    AnalisaStatusCheque = StringRetorno
End Function

Public Sub ExibeArquivoRetorno()
    If ArqRetorno <> Empty Then
      If Dir(ArqRetorno) <> "" Then
          Shell "notepad.exe" + " " + ArqRetorno, vbNormalFocus
      End If
    End If
End Sub

Public Sub VERIFICA_BEMATECH_LIGADA()
   INDR_DESLIGADA = False
   RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
   If RETORNO_ECF = -6 Then
     MsgBox "A Impressora se encontra DESLIGADA.", vbInformation + vbOKOnly, "Atenção"
     INDR_DESLIGADA = True
     Else       'MsgBox "A Impressora se encontra LIGADA.", vbInformation + vbOKOnly, "Atenção"
   End If
   Call VerificaRetornoImpressora("", "", "Impressora Fiscal")
End Sub

'===============TEF
Public Sub CHAMA_EASYTEF()
   'Msg = "Aguarde, iniciando aplicativo TEF ..."
   'frmPedidoVenda.MOSTRA_RODAPE_PEDIDO Msg, "", "", "", ""
   '===========================================================
   Dim i                   As Integer
   Dim Formas              As Variant
   Dim Valores             As Variant

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'Faz as verificações de TEF
   If frmINICIO.UsarTEF Then
      i = 0
      Formas = Array("")
      Valores = Array("")

      SQL = "select ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
      SQL = SQL & " from LANCAMENTO "
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
      SQL = SQL & " INNER JOIN FORMAPAGTO "
      SQL = SQL & " ON ITEMLANCAMENTO.formapagto_id = FORMAPAGTO.formapagto_id"

      SQL = SQL & " where LANCAMENTO.numr_doc = " & PEDIDO_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         ' se for uma forma de pagamento de cartão
         If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CARTAO") > 0 Then
             If i > 0 Then
                 ReDim Preserve Formas(UBound(Formas) + 1)
                 ReDim Preserve Valores(UBound(Valores) + 1)
             End If

             Formas(i) = Left(UCase(Trim(TabTemp.Fields("descricao").Value)), 6)
             Valores(i) = Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
             i = i + 1
         ElseIf InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CHEQUE") > 0 Then

             If MsgBox("Gostaria de consultar este cheque junto à SERASA (Somente Redecard)?", _
                 vbYesNo + vbQuestion, "Consulta de Cheque") = vbYes Then
         
               If Not frmINICIO.ConsultarCheque(TabTemp.Fields("valor_item").Value, _
                     Left(UCase(Trim(TabTemp.Fields("descricao").Value)), 15)) Then
                  Exit Sub
                 End If
             End If
         End If

         TabTemp.MoveNext
      Wend

      ' se encontrou alguma forma de pagamento de cartão
      If i > 0 Then
         If Not frmINICIO.tratarPagamentoComCartao(Valores, Formas) Then
             Exit Sub
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   '===========================================================
End Sub
