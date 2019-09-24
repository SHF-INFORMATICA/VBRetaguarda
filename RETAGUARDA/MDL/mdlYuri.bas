Attribute VB_Name = "mdlYuri"

   'Vetores/Matrizes
   Public vet(500) As String
   Public varaux(0 To 9)  As String
   Public matrizretorno() As Variant ' monta gride
   
   ' inicio yuri 01/05/2012 - Criado essas variaveis publica para os novos calculos de impostos
   Public pCodigoAliquota As Integer
   Public pcodigoaliquotaNC As Integer
   Public pPercentualICMS_NC As Currency
   Public pPercentualICMSSubst As Currency
   Public pcodigoaliquotaNCS As Integer
   Public pcodigoaliquotaisento As Integer
   Public pPercentualICMSIsento As Currency
   Public pPercentualICMS As Currency
   'informa se a empresa trabalha com regime do TARE
   Global g_trabalhacomtare_empresa As Byte
   'informa a observacao do tare para sair na nota fiscal
   Global g_observacaotare_empresa As String
   'informa a observacao do tare quando o cliente for TARE
   Global g_observacaotare_cliente As String
   'informa se o cliente é indústria / fabricante
   ' fim

'==============YURI
Public Sub BuscaAliquota(ByVal UF As String, ByVal CodigoCliente As Long)
'On Error GoTo Err_BuscaAliquota

   Dim curPorcICMS_subs_supersimples As Currency
   Dim BD_Record_SetII As New ADODB.Recordset
   Dim cCliente As New cCliente

    BD_Record_SetII.Source = "SELECT * FROM aliquota_uf WHERE uf = '" & UF & "' LIMIT 1"
    BD_Record_SetII.Open
    If BD_Record_SetII.RecordCount > 0 Then
        pCodigoAliquota = BD_Record_SetII!codigo
        pPercentualICMS = Format(BD_Record_SetII!Aliquota, "##,###,##0.00")
        
        pcodigoaliquotaNC = BD_Record_SetII!codigo_aliquota_nc
        pPercentualICMS_NC = Format(BD_Record_SetII!aliquota_nc, "##,###,##0.00")
                
        pcodigoaliquotaNCS = BD_Record_SetII!codigo_aliquota_substituicao
        
        ' Preenche o perccentual
        If cCliente.BuscaDadosdoClienteSuperSimples(CodigoCliente, curPorcICMS_subs_supersimples) Then
            pPercentualICMSSubst = Format(curPorcICMS_subs_supersimples, "##,###,##0.00") ' SOMENTE EMPRESA OPTANTE DO TAREE
        Else
            pPercentualICMSSubst = Format(BD_Record_SetII!aliquota_substituicao, "##,###,##0.00") ' PEGA EMPRESA NORMAL
        End If
    
        pcodigoaliquotaisento = BD_Record_SetII!codigo_aliquota_isento
        pPercentualICMSIsento = BD_Record_SetII!icms_isento
    
    Else
        pCodigoAliquota = 0
        pPercentualICMS = 0
        pcodigoaliquotaNC = 0
        pPercentualICMS_NC = 0
        pcodigoaliquotaNCS = 0
        pPercentualICMSSubst = 0
        pcodigoaliquotaisento = 0
        pPercentualICMSIsento = 0
        
        'Alerta "Alíquota Para este Estado não Encontrado!"
    End If
    BD_Record_SetII.Close

Exit Sub
'Err_BuscaAliquota: ValidaErros Err, Me.Caption & " - BuscaAliquota"
End Sub

Public Function ValidaProdNFE(ByVal strCodigo As String) ' yuri em 01/05/2012
'On Error GoTo ValidaProdNFE

Dim Sql_Record_Set As New ADODB.Recordset

ValidaProdNFE = False
'If g_nfe_empresa = 1 Then
    
    Sql_Record_Set.Source = "SELECT PD.codificacao_fiscal,PD.cst_pis,PD.pis,PD.pis_entrada," & _
                                   "PD.cst_cofins,PD.cofins,PD.cofins_entrada " & _
                                   "FROM produto PD " & _
                                   "WHERE PD.codigo = '" & strCodigo & "' AND PD.servico_produto = 'P' LIMIT 1"
    Sql_Record_Set.Open
    If Sql_Record_Set.RecordCount > 0 Then
       With Sql_Record_Set
           'VALIDA CODIFICAÇÃO FISCAL - POIS OBRIGATARIO NFE
           If Trim(!codificacao_fiscal) = "" Or Trim(!codificacao_fiscal) = "0" Then
             ' Alerta "Classificação fiscal do Produto " & strcodigo & " Não Informada ou Inválida!"
              Sql_Record_Set.Close
              Exit Function
           'VALIDA CST DO PIS
           ElseIf Trim(!cst_pis) = "0" Or Trim(!cst_pis) = "00" Then
              ' Alerta "CST do PIS do Produto " & strcodigo & " Inválido!"
               Sql_Record_Set.Close
               Exit Function
           'VALIDA CST DO COFINS
           ElseIf Trim(!cst_cofins) = "0" Or Trim(!cst_cofins) = "00" Then
              ' Alerta "CST do Cofins do Produto " & strcodigo & " Inválido!"
               Sql_Record_Set.Close
               Exit Function
           End If
           
           If Trim(!cst_pis) = "01" Or Trim(!cst_pis) = "02" Then
                If CCur(!Pis) <= 0 Then
                '   Alerta "Valor PIS Venda Não Pode ser Zero para CST " & !cst_pis & " Produto Cod. " & strcodigo & " !"
                   Sql_Record_Set.Close
                   Exit Function
                End If
           End If
           
           'VALIDA VALORES CST COFINS 01 - 02 - 09 OBRIGATORIO VALOR > 0
           If Trim(!cst_cofins) = "01" Or Trim(!cst_cofins) = "02" Then
                If CCur(!Cofins) <= 0 Then
                 '  Alerta "Valor Cofins Venda Não Pode ser Zero para CST " & !cst_cofins & " Produto Cod. " & strcodigo & " !"
                   Sql_Record_Set.Close
                   Exit Function
                End If
           End If
       End With
    End If
    Sql_Record_Set.Close
' End If
ValidaProdNFE = True

Exit Function
'ValidaProdNFE: ValidaErros Err, ValidaProdNFE & " - ValidaProdNFE"
End Function

Public Function BuscaIVAporEstadoporProduto(strCodProduto As String, _
                                            strUF As String, _
                                            bolEntrada As Boolean) As Double
'On Error GoTo Err_BuscaIVAporEstadoporProduto

Dim rsBuscaGeral As New ADODB.Recordset

    BuscaIVAporEstadoporProduto = 0
    Sql_Query = " SELECT CIEP.perc_iva_entrada,CIEP.perc_iva_saida" & _
                " FROM aliquota_uf AE" & _
                " LEFT JOIN cadastro_iva_uf_produto CIEP" & _
                " ON CIEP.aliquota_uf_id = AE.codigo" & _
                " WHERE AE.UF = '" & strUF & "'" & _
                " AND CIEP.produto_id = '" & strCodProduto & "' "
    'Set rsBuscaGeral = Conexao.GeraRecordset(Sql_Query, 0)
    rsBuscaGeral.Source = Sql_Query
    rsBuscaGeral.Open
    If rsBuscaGeral.RecordCount > 0 Then
        If bolEntrada = True Then
            BuscaIVAporEstadoporProduto = IIf(IsNull(rsBuscaGeral!perc_iva_entrada), 0, rsBuscaGeral!perc_iva_entrada)
        Else
            BuscaIVAporEstadoporProduto = IIf(IsNull(rsBuscaGeral!perc_iva_saida), 0, rsBuscaGeral!perc_iva_saida)
        End If
    End If
    rsBuscaGeral.Close

Exit Function
'Err_BuscaIVAporEstadoporProduto: ValidaErros Err, "mdl_ivaestadoproduto - BuscaIVAporEstadoporProduto"
End Function
' porfavornao retirar esta função será necessária
Private Function CodificacaoFiscal() As Boolean
   'On Error GoTo Err_CodificacaoFiscal
   Dim lbl_dentropais As String
   Dim lbl_servico_produto As String
   Dim pRevendedor As String

'    CodificacaoFiscal = True
'    ValidaEmissaoCupom = True
'    lNomeCodificacaoFiscal = ""
'    bolExportacao = False
'
'    If cbo_uf = g_uf_empresa Then
'        g_string = "D"
'    Else
'        g_string = "F"
'    End If
'
'    If cbo_uf = "EX" Then
'        lbl_dentropais = "F"
'        bolExportacao = True
'    Else
'        lbl_dentropais = "D"
'    End If
'
'    If grade1.TextMatrix(i, 19) = "S" Then
'        lbl_servico_produto = "S"
'    Else
'        lbl_servico_produto = "P"
'    End If
'
'    'ISENTO OU TRIBUTADOS
'    If lCodigoAliquotaNat = 1 Or lCodigoAliquotaNat = 3 Then
'        lnaturezatributacao = "('1','3')"
'    'SUBSTITUICAO
'    ElseIf lCodigoAliquotaNat = 2 Then
'        lnaturezatributacao = "('2')"
'    Else
'    'TRIBUTADOS
'        lnaturezatributacao = "('3')"
'    End If
'
'    'altera pra buscar certo o cfop
'    If xRevendedor = 1 Then
'        'busca revendedor
'        pRevendedor = 2
'    Else
'        'busca consumidor
'        pRevendedor = 1
'    End If
'
'
'    'CONSUMIDOR - A VISTA - A PRAZO
'    If Val(g_vista_prazo_outras) <= 2 Then
'            'primeira pesquisa observando principalmente a aliquota do produto e se e cliente contribuinte ou não e seus derivados
'            'segundo pesquisa observando principalmente a aliquota e seus derivados
'            'terceiro pesquisa observando apenas os derivados
'            Set gb_Recordset = Conexao.GeraRecordset("(SELECT codigo,descricao,item,codigo_csosn,('1') as ordem FROM natureza_operacao WHERE codigo_da_aliquota IN " & lnaturezatributacao & " and tipo_da_operacao = 'S' and estado = '" & g_string & "' and pais = '" & lbl_dentropais & "' and tipo_cliente = '" & pRevendedor & "' and tipo_venda IN ('1','2') and servico_produto = '" & lbl_servico_produto & "' and suframa = '" & x_suframa & "' and fora_estabelecimento = '" & xcfoptransito & "' and industria_revenda = '" & pIndustriaRevenda & "') UNION " & _
'                                                     " (SELECT codigo,descricao,item,codigo_csosn,('2') as ordem FROM natureza_operacao WHERE codigo_da_aliquota IN " & lnaturezatributacao & " and tipo_da_operacao = 'S' and estado = '" & g_string & "' and pais = '" & lbl_dentropais & "' and tipo_cliente = '3' and tipo_venda IN ('1','2') and servico_produto = '" & lbl_servico_produto & "' and suframa = '" & x_suframa & "' and fora_estabelecimento = '" & xcfoptransito & "' and industria_revenda = '" & pIndustriaRevenda & "') UNION " & _
'                                                     " (SELECT codigo,descricao,item,codigo_csosn,('3') as ordem FROM natureza_operacao WHERE tipo_da_operacao = 'S' and estado = '" & g_string & "' and pais = '" & lbl_dentropais & "' and tipo_venda IN ('1','2') and servico_produto = '" & lbl_servico_produto & "' and suframa = '" & x_suframa & "' and fora_estabelecimento = '" & xcfoptransito & "' and industria_revenda = '" & pIndustriaRevenda & "')ORDER BY ordem ASC", 1)
'            If gb_Recordset.RecordCount > 0 Then
'                lCodigoFiscal = gb_Recordset!Codigo
'                'lCodificacaoFiscal(lcodigoaliquota) = gb_Recordset!Codigo
'                lNomeCodificacaoFiscal = gb_Recordset!Descricao
'                lbl_itemcfop = gb_Recordset!Item
'                grade1.TextMatrix(i, 66) = gb_Recordset!codigo_csosn
'                gb_Recordset.Close
'            Else
'                gb_Recordset.Close
'                CodificacaoFiscal = False
'                Alerta "Codificação Fiscal Não Localizado Produto " & grade1.TextMatrix(i, 1) & "!"
'                Exit Function
'            End If
'
'            Call CalculaTributacao(i)
'        'OUTRAS SAIDAS
'    ElseIf Val(g_vista_prazo_outras) > 2 Then
'            Sql_Query = ""
'            Sql_Query = " (SELECT codigo, codigo_da_aliquota, descricao, item, ('1') as ordem, codigo_csosn" & _
'                        " FROM natureza_operacao" & _
'                        " WHERE codigo_da_aliquota IN " & lnaturezatributacao & " " & _
'                        " and tipo_da_operacao = 'S'" & _
'                        " and estado = '" & g_string & "'" & _
'                        " and pais = '" & lbl_dentropais & "'" & _
'                        " and tipo_venda = '" & g_vista_prazo_outras & "')" & _
'                        " UNION" & _
'                        " (SELECT codigo, codigo_da_aliquota, descricao, item, ('2') as ordem, codigo_csosn" & _
'                        " FROM natureza_operacao" & _
'                        " WHERE tipo_da_operacao = 'S'" & _
'                        " and estado = '" & g_string & "'" & _
'                        " and pais = '" & lbl_dentropais & "'" & _
'                        " and tipo_venda = '" & g_vista_prazo_outras & "'" & _
'                        " ) ORDER BY ordem ASC"
'            Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 1)
'            If gb_Recordset.RecordCount > 0 Then
'                lCodigoFiscal = gb_Recordset!Codigo
'                lbl_itemcfop = gb_Recordset!Item
'                lnaturezatributacao = "('" & gb_Recordset!codigo_da_aliquota & "')"
'                grade1.TextMatrix(i, 66) = gb_Recordset!codigo_csosn
'
'                'quando for isento
'                If gb_Recordset!codigo_da_aliquota = 1 Then
'                    lcodigoaliquota = 1
'                    grade1.TextMatrix(i, 12) = 1
'                End If
'                'lCodificacaoFiscal(lcodigoaliquota) = gb_Recordset!Codigo
'                lNomeCodificacaoFiscal = gb_Recordset!Descricao
'                gb_Recordset.Close
'            Else
'                gb_Recordset.Close
'                CodificacaoFiscal = False
'                Alerta "Codificação Fiscal Não Localizado!"
'                Exit Function
'            End If
'            ValidaEmissaoCupom = False
'
'            'verifica se sofre algum tipo de tributacao pela natureza
'            Call CalculaTributacao(i)
'    End If
'
'    '1- se for contribuinte so imprime nota
'    '2- se for servico e o municipio nao liberar servico
'    '   para cupom fiscal so imprime nota
'    '3 - FORA DO PAIS
'    '4 - se tiver outras despesas
'    '5 - Bloquear emissão de cupom fora do estado '*** Autor: Diego Martins Data:13/09/2011 OS:17304 ***
'
'    If (cmd_sair1.Visible = True) Or (lbl_servico_produto = "S" And g_libera_servico = 0) Or (CDbl(txt_valor_frete) + CDbl(txt_outras_despesas)) > 0 Or (g_imp_aut_juridica = 1 And lContribuinte = "S") Or (cbo_uf = "EX") Or (cbo_uf <> g_uf_empresa) Then
'        ValidaEmissaoCupom = False
'    End If
'
'    lbl_cfop = lCodigoFiscal

Exit Function
'Err_CodificacaoFiscal: ValidaErros Err, Me.Caption & " - CodificacaoFiscal"
End Function



'==============YURI
'==============YURI
Public Sub VERIFICA_BANCO_DADOS()
'On Error GoTo ERRO_TRATA

   Dim rs_Busca As New ADODB.Recordset

   ' INICIO SPED FISCAL - YURI
   If ExisteTabela("sped_info_conta_consumo") = False Then
      sSQL = "CREATE TABLE [dbo].[sped_info_conta_consumo] ( "
      sSQL = sSQL & "    [empresa] [char] (2) NOT NULL ,"
      sSQL = sSQL & "    [codigo] [tinyint] NOT NULL ,"
      sSQL = sSQL & "    [tipo_conta] [varchar] (200) COLLATE Latin1_General_CI_AS NULL ,"
      sSQL = sSQL & "    [CodigoFiscal] [varchar] (3) COLLATE Latin1_General_CI_AS NULL "
      sSQL = sSQL & ") ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute sSQL
   End If
   
   If ExisteTabela("ALTERACAO_PRODUTO") = False Then
      sSQL = "CREATE TABLE [dbo].[ALTERACAO_PRODUTO] ( "
      sSQL = sSQL & "    [CODG_PROD] [nvarchar](30) NOT NULL ,"
      sSQL = sSQL & "    [DESCRICAO] [varchar](200) NULL ,"
      sSQL = sSQL & "    [DATA_INICIAL] [date] NULL ,"
      sSQL = sSQL & "    [DATA_FINAL] [date] NULL "
      sSQL = sSQL & ") ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute sSQL
   End If
   
   If ExisteTabela("ALTERACAO_CLIENTE") = False Then
      sSQL = "CREATE TABLE [dbo].[ALTERACAO_CLIENTE] ( "
      sSQL = sSQL & "    [CODIGO_CAMPO] [bigint] NOT NULL,"
      sSQL = sSQL & "    [CODIGO_CLIENTE] [int] NULL,"
      sSQL = sSQL & "    [DATA_ALTERACAO] [date] NULL,"
      sSQL = sSQL & "    [VALOR_ANTERIOR] [varchar](100) NULL "
      sSQL = sSQL & ") ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute sSQL
   End If

   If ExisteTabela("SPED_TIPO_PRODUTO") = False Then
      sSQL = "CREATE TABLE [dbo].[SPED_TIPO_PRODUTO] ( "
      sSQL = sSQL & "    [codigo] [INT] NOT NULL ,"
      sSQL = sSQL & "    [TIPO_ITEM] [varchar] (200)  NULL "
      sSQL = sSQL & ") ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute sSQL

      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('00','Mercadoria para Revenda')"
      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('01','Matéria-Prima')"

      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('02','Embalagem')"
      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('03','Produto em Processo')"

      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('04','Produto Acabado')"
      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('05','Subproduto')"

      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('06','Produto Intermediário')"
      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('07','Material de Uso e Consumo')"

      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('08','Ativo Imobilizado')"
      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('09','Serviços')"

      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('10','Outros insumos')"
      CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('99','Outras')"
      Else
         rs_Busca.Open "SELECT ISNULL(COUNT(*), 0) AS QTD FROM SPED_TIPO_PRODUTO", CONECTA_RETAGUARDA, , , adCmdText
         If rs_Busca!qtd <> 12 Then
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('00','Mercadoria para Revenda')"
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('01','Matéria-Prima')"

              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('02','Embalagem')"
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('03','Produto em Processo')"

              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('04','Produto Acabado')"
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('05','Subproduto')"

              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('06','Produto Intermediário')"
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('07','Material de Uso e Consumo')"

              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('08','Ativo Imobilizado')"
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('09','Serviços')"

              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('10','Outros insumos')"
              CONECTA_RETAGUARDA.Execute "INSERT INTO SPED_TIPO_PRODUTO (Codigo, TIPO_ITEM) VALUES ('99','Outras')"
         End If
   End If
   If rs_Busca.State = 1 Then _
      rs_Busca.Close

'ESSAS ALTERAÇÕES SOMENTE APOS TESTE
'[PERCIVAENTRADA] [float] NULL,
'YURI - 01/03/2012 - CRIAR O PERCENTUAL DE ENTRADA DE IVA
' servico_produto
'If ExisteCampo("PERCIVAENTRADA", "PRODUTO") = False Then
'CONECTA_RETAGUARDA.execute  "UPDATE Parametro SET CodigoFilial = 1 "

'If ExisteCampo("CODLISTSERVICO", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CODLISTSERVICO VARCHAR(10) NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_CODLISTSERVICO]  DEFAULT ('  ') FOR [CODLISTSERVICO] "
'End If
'If ExisteCampo("PERCIVAENTRADA", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PERCIVAENTRADA float NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_PERCIVAENTRADA]  DEFAULT (0) FOR [PERCIVAENTRADA] "
'End If
'If ExisteCampo("SERVICO_PRODUTO", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD SERVICO_PRODUTO VARCHAR(10) NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_SERVICO_PRODUTO]  DEFAULT ('produto') FOR [SERVICO_PRODUTO] "
'End If
'If ExisteCampo("TIPO_PRODUTO_RELEVANTE", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD TIPO_PRODUTO_RELEVANTE VARCHAR(02) NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_TIPO_PRODUTO_RELEVANTE]  DEFAULT ('00') FOR [TIPO_PRODUTO_RELEVANTE] "
'End If

'If ExisteCampo("MODELO_NF", "NF") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.NF ADD MODELO_NF CHAR(02) NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
'End If

'If ExisteCampo("PERCIVAENTRADA", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PERCIVAENTRADA = 0 "
'End If

'If ExisteCampo("TIPO_PRODUTO_RELEVANTE", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET TIPO_PRODUTO_RELEVANTE = '00' "
'End If

'If ExisteCampo("SERVICO_PRODUTO", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET SERVICO_PRODUTO = 'produto' "
'End If

'If ExisteCampo("SERVICO_PRODUTO", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET SERVICO_PRODUTO = 'produto' "
'End If

   ' FIM SPED FISCAL

   ' depois voce vai dar uma olhada daqui para frente do que fiz pois e referente à nf
   ' INICIO PARA ATENDER NOTA ELETRONICA E SPED
   ' CADASTRO DE EMPRESA
   If ExisteCampo("Empresa_Regime_TARE", "EMPRESA") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.EMPRESA ADD Empresa_Regime_TARE INT NULL"
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.EMPRESA ADD CONSTRAINT [DF_EMPRESA_Empresa_Regime_TARE]  DEFAULT (0) FOR [Empresa_Regime_TARE] "
      CONECTA_RETAGUARDA.Execute "UPDATE EMPRESA SET Empresa_Regime_TARE = 0 "
   End If
   ' - Super_simples = é Empresa do Simples default = 1 , pois todos os nossos clientes são do super simples
   If ExisteCampo("super_simples", "EMPRESA") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.EMPRESA ADD super_simples INT NULL"
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.EMPRESA ADD CONSTRAINT [DF_EMPRESA_super_simples]  DEFAULT (1) FOR [super_simples] "
      CONECTA_RETAGUARDA.Execute "UPDATE EMPRESA SET super_simples = 1 "
   End If

'ESSAS ALTERAÇÕES SOMENTE DEPOIS DE TESTES
'CADASTRO DE CLIENTE
'Cliente_TARE = Aqui se o Cliente é do Tare = 0 para não
'If ExisteCampo("Cliente_TARE", "CLIENTE") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD Cliente_TARE INT NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD CONSTRAINT [DF_CLIENTE_Cliente_TARE]  DEFAULT (0) FOR [Cliente_TARE] "
'   CONECTA_RETAGUARDA.Execute "UPDATE CLIENTE SET Cliente_TARE = 0 "
'End If
'aliquota_subst_super_simples - colocar aliquota aqui se é cliente super_simpes e estiver enquadrado no tare]
'aqui estou definindo como zero
'If ExisteCampo("aliquota_subst_super_simples", "CLIENTE") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD aliquota_subst_super_simples float NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD CONSTRAINT [DF_CLIENTE_aliquota_subst_super_simples]  DEFAULT (0) FOR [aliquota_subst_super_simples] "
'   CONECTA_RETAGUARDA.Execute "UPDATE CLIENTE SET aliquota_subst_super_simples = 0 "
'End If

'super_simples - se este cliente é do super simples = default que nao e cliente do super simples
' este tratamento e a empresa que tem que fazer para cada cliente seu no momento do cadastro
'If ExisteCampo("Cliente_Simples", "CLIENTE") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD Cliente_Simples int NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD CONSTRAINT [DF_CLIENTE_Cliente_Simples]  DEFAULT (0) FOR [Cliente_Simples] "
'   CONECTA_RETAGUARDA.Execute "UPDATE CLIENTE SET Cliente_Simples = 0 "
'End If

   'CST_PIS DEPOIS DA UMA OLHADA  QUAL O CST_PIS QUE ESTAMOS USANDO PARA A NOTA ELETRONICA?????
   If ExisteCampo("CST_PIS", "CFOP") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CST_PIS VARCHAR(03) NULL"
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CONSTRAINT [DF_CFOP_CST_PIS]  DEFAULT ('49') FOR [CST_PIS] "
      CONECTA_RETAGUARDA.Execute "UPDATE CFOP SET CST_PIS = '49' "
   End If

   'CST_PIS DEPOIS DA UMA OLHADA  QUAL O CST_PIS QUE ESTAMOS USANDO PARA A NOTA ELETRONICA?????
   ' OS CSTDE PIS E COFINS OBRIGATORIAMENTE SÃO IGUAIS SEM EXCEÇÃO SE O PIS
   ' FOR 49 O COFINS TEM QUE SER 49
   If ExisteCampo("CST_COFINS", "CFOP") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CST_COFINS VARCHAR(03) NULL"
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CONSTRAINT [DF_CFOP_CST_COFINS]  DEFAULT ('49') FOR [CST_COFINS] "
      CONECTA_RETAGUARDA.Execute "UPDATE CFOP SET CST_COFINS = '49' "
   End If

'ESSAS AQUI SOMENTE DEPOIS DE TESTES
' NESTES DOIS CASOS  NO SITE DA RECEITA FEDERAL TEM OS VALORES, MAS NÃO SEPREOCUPE COM ISSO POR ENQUANDO.
'If ExisteCampo("PER_COFINS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PER_COFINS int NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_PER_COFINS]  DEFAULT (0) FOR [PER_COFINS] "
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PER_COFINS = 0 "
'End If

'If ExisteCampo("PER_PIS", "PRODUTO") = False Then
'  CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PER_PIS int NULL"
'  CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_PER_PIS]  DEFAULT (0) FOR [PER_PIS] "
'  CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PER_PIS= 0 "
'End If

    ' FIM PARA ATENDER NOTA ELETRONICA E SPED
    
    ' INICIO 01/05/2012 TABELAS E CAMPOS CRIADOS PARA O NOVO CALCULO FISCAL PARA ATENDER TODAS AS EMPRESAS
    
'If ExisteCampo("VAREJO_ATACADO", "CLIENTE") = False Then
'    CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD VAREJO_ATACADO INT  NULL"
'    CONECTA_RETAGUARDA.Execute "UPDATE dbo.CLIENTE SET VAREJO_ATACADO = 1 "
'    'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
'End If
    
   ' Estes códigos são fixos e somente podem ter 8 itens
   ' daqui  vou precisar mudar nao pode ser assim, tem que ser com esses campos,vou deletar de la
   ' retirei la embaixo  e ficara esse
   If ExisteTabela("ALIQUOTA") = False Then
      sSQL = "CREATE TABLE [dbo].[ALIQUOTA] ( "
      sSQL = sSQL & "    [codigo] [INT] NOT NULL ,"
      sSQL = sSQL & "    [NOME] [varchar] (50)  NULL ,"
      sSQL = sSQL & "    [aliquota_do_imposto] [real] NULL,"
      sSQL = sSQL & "    [EMPRESA] [INT] NOT NULL"
      sSQL = sSQL & ") ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute sSQL
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (1,'ISENTO',0.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (2,'Substituição Tributaria',0.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (3,'Nao Incidencia',0.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (4,'ICMS 12%',12.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (5,'ICMS 0%',0.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (6,'ISS 5%',5.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (7,'ICMS 15%',15.0000,1)"
           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (8,'ICMS 17%',17.0000,1)"
      Else
         rs_Busca.Open "SELECT ISNULL(COUNT(*), 0) AS QTD FROM [dbo].[ALIQUOTA]", CONECTA_RETAGUARDA, , , adCmdText
         If rs_Busca!qtd = 0 Then
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (1,'ISENTO',0.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (2,'Substituição Tributaria',0.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (3,'Nao Incidencia',0.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (4,'ICMS 12%',12.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (5,'ICMS 0%',0.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (6,'ISS 5%',5.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (7,'ICMS 15%',15.0000,1)"
            CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (8,'ICMS 17%',17.0000,1)"
         End If
   End If
   If rs_Busca.State = 1 Then _
     rs_Busca.Close

   ' Estes códigos são fixos e somente podem ter 27 itens QUE CORREPOMDEM A VINTE E SETE ESTADOS
   ' DEPOIS VC COLOCA ISTO AONDE VOCE CRIOU ESSA TABELA
   If ExisteTabela("ALIQUOTA_UF") = False Then
      SQL = " CREATE TABLE ALIQUOTA_UF("
      SQL = SQL & " CODIGO bigint not null"
      SQL = SQL & ", ESTADO varchar(2) NOT NULL"
      SQL = SQL & ", codigo_aliquota bigint NOT NULL"
      SQL = SQL & ", Aliquota decimal(4,2) NOT NULL "
      SQL = SQL & ", aliquota_nc decimal(4,2) NOT NULL "
      SQL = SQL & ", Descricao varchar(100) NOT NULL "
      SQL = SQL & ", codigo_aliquota_nc bigint NOT NULL "
      SQL = SQL & ", codigo_aliquota_substituicao bigint NOT NULL "
      SQL = SQL & ", aliquota_substituicao decimal(4,2) NOT NULL "
      SQL = SQL & ", empresa_id bigint NOT NULL "
      SQL = SQL & ", codigo_uf bigint NOT NULL "
      SQL = SQL & ", codigo_aliquota_isento bigint NOT NULL "
      SQL = SQL & ", icms_isento bigint NOT NULL "
      SQL = SQL & " ,CONSTRAINT [PK_ALIQUOTA_UF] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ("
      SQL = SQL & " [CODIGO] Asc"
      SQL = SQL & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      'CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (1,'AC',8,17.0000,17.0000,'ACRE',8,8,17.0000,'2',12,8,17.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (2,'AL',5,0.0000,0.0000,'ALAGOAS',5,5,0.0000,'2',27,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (3,'AM',5,0.0000,0.0000,'AMAZONAS',5,5,0.0000,'2',13,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (4,'AP',5,0.0000,0.0000,'AMAPA',5,5,0.0000,'2',16,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (5,'BA',5,0.0000,0.0000,'BAHIA',5,5,0.0000,'2',29,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (6,'CE',5,0.0000,0.0000,'CEARA',5,5,0.0000,'2',23,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (7,'DF',8,17.0000,17.0000,'DISTRITO FEDERAL',8,8,17.0000,'2',53,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (8,'ES',5,0.0000,0.0000,'ESPIRITO SANTO',5,5,0.0000,'2',32,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (9,'MA',5,0.0000,0.0000,'MARANHAO',5,5,0.0000,'2',21,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (10,'MG',5,0.0000,0.0000,'MINAS GERAIS',5,5,0.0000,'2',31,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (11,'MT',5,0.0000,0.0000,'MATO GROSSO',5,5,0.0000,'2',51,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (12,'MS',5,0.0000,0.0000,'MATO GROSSO DO SUL',5,5,0.0000,'2',50,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (13,'PA',5,0.0000,0.0000,'PARA',5,5,0.0000,'2',15,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (14,'PB',5,0.0000,0.0000,'PARAIBA',5,5,0.0000,'2',25,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (15,'PE',5,0.0000,0.0000,'PERNAMBUCO',5,5,0.0000,'2',26,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (16,'PI',5,0.0000,0.0000,'PIAUI',5,5,0.0000,'2',22,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (17,'PR',5,0.0000,0.0000,'PARANA',5,5,0.0000,'2',41,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (18,'RJ',5,0.0000,0.0000,'RIO DE JANEIRO',5,5,0.0000,'2',33,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (19,'RO',5,0.0000,0.0000,'RONDONIA',5,5,0.0000,'2',11,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (20,'RN',5,0.0000,0.0000,'RIO GRANDE DO NORTE',5,5,0.0000,'2',24,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (22,'RS',5,0.0000,0.0000,'RIO GRANDE DO SUL',5,5,0.0000,'2',43,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (23,'SC',5,0.0000,0.0000,'SANTA CATARINA',5,5,0.0000,'2',42,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (24,'SE',5,0.0000,0.0000,'SERGIPE',5,5,0.0000,'2',28,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (25,'SP',5,0.0000,0.0000,'SAO PAULO',5,5,0.0000,'2',35,1,0.0000)"
      
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (26,'TO',5,0.0000,0.0000,'TOCANTINS',5,5,0.0000,'2',17,1,0.0000)"
      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (27,'GO',4,12.0000,12.0000,'GOIAS',4,5,0.0000,'2',52,1,0.0000)"
      Else ' CASO EXISTA A TABELA E NAO EXISTA OS REGISTROS , CRIA OS 27 ESTADOS COM SUAS RESPECTIVAS ALIQUOTAS
         rs_Busca.Open "SELECT ISNULL(COUNT(*), 0) AS QTD FROM [dbo].[ALIQUOTA_UF]", CONECTA_RETAGUARDA, , , adCmdText
         If rs_Busca!qtd = 0 Then
              CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (1,'AC',8,17.0000,17.0000,'ACRE',8,8,17.0000,'2',12,8,17.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (2,'AL',5,0.0000,0.0000,'ALAGOAS',5,5,0.0000,'2',27,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (3,'AM',5,0.0000,0.0000,'AMAZONAS',5,5,0.0000,'2',13,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (4,'AP',5,0.0000,0.0000,'AMAPA',5,5,0.0000,'2',16,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (5,'BA',5,0.0000,0.0000,'BAHIA',5,5,0.0000,'2',29,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (6,'CE',5,0.0000,0.0000,'CEARA',5,5,0.0000,'2',23,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (7,'DF',8,17.0000,17.0000,'DISTRITO FEDERAL',8,8,17.0000,'2',53,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (8,'ES',5,0.0000,0.0000,'ESPIRITO SANTO',5,5,0.0000,'2',32,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (9,'MA',5,0.0000,0.0000,'MARANHAO',5,5,0.0000,'2',21,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (10,'MG',5,0.0000,0.0000,'MINAS GERAIS',5,5,0.0000,'2',31,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (11,'MT',5,0.0000,0.0000,'MATO GROSSO',5,5,0.0000,'2',51,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (12,'MS',5,0.0000,0.0000,'MATO GROSSO DO SUL',5,5,0.0000,'2',50,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (13,'PA',5,0.0000,0.0000,'PARA',5,5,0.0000,'2',15,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (14,'PB',5,0.0000,0.0000,'PARAIBA',5,5,0.0000,'2',25,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (15,'PE',5,0.0000,0.0000,'PERNAMBUCO',5,5,0.0000,'2',26,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (16,'PI',5,0.0000,0.0000,'PIAUI',5,5,0.0000,'2',22,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (17,'PR',5,0.0000,0.0000,'PARANA',5,5,0.0000,'2',41,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (18,'RJ',5,0.0000,0.0000,'RIO DE JANEIRO',5,5,0.0000,'2',33,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (19,'RO',5,0.0000,0.0000,'RONDONIA',5,5,0.0000,'2',11,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (20,'RN',5,0.0000,0.0000,'RIO GRANDE DO NORTE',5,5,0.0000,'2',24,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (22,'RS',5,0.0000,0.0000,'RIO GRANDE DO SUL',5,5,0.0000,'2',43,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (23,'SC',5,0.0000,0.0000,'SANTA CATARINA',5,5,0.0000,'2',42,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (24,'SE',5,0.0000,0.0000,'SERGIPE',5,5,0.0000,'2',28,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (25,'SP',5,0.0000,0.0000,'SAO PAULO',5,5,0.0000,'2',35,1,0.0000)"
         
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (26,'TO',5,0.0000,0.0000,'TOCANTINS',5,5,0.0000,'2',17,1,0.0000)"
         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (27,'GO',4,12.0000,12.0000,'GOIAS',4,5,0.0000,'2',52,1,0.0000)"
         End If
   End If
   If rs_Busca.State = 1 Then _
      rs_Busca.Close

   ' CRIAR NA TELA DE CADASTRO DE CLIENTES PARA
   ' O USUARIO PREENCHEER ESTE CAMPO VALORES VALIDOS 0(NAO) E 1(EMPRESA É OPTANTES DO TARE)
   If ExisteCampo("OPTANTE_TARE", "EMPRESA") = False Then
       CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.EMPRESA ADD OPTANTE_TARE BIT NULL"
       CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[EMPRESA] ADD  CONSTRAINT [DF_EMPRESA_OPTANTE_TARE]  DEFAULT (0) FOR [OPTANTE_TARE]"
       CONECTA_RETAGUARDA.Execute "UPDATE dbo.EMPRESA SET OPTANTE_TARE = 0 "
   End If
   
   If ExisteCampo("OPTANTE_TARE", "EMPRESA") = True Then ' POR ENQUANTO NÓS TRABALHAMOS COM A EMPRESAS QUE NAO É OPTANTE DO TARE
       'CONECTA_RETAGUARDA.Execute "UPDATE dbo.EMPRESA SET OPTANTE_TARE = 0 "
   End If
    
    ' é o mesmo de Empresa_Optante_Simples
    'não necessita super_simples, aliquota_subst_super_simples 01/05/2012
    ' CRIAR NA TELA DE CADASTRO DE CLIENTES PARA
    ' O USUARIO PREENCHEER ESTE CAMPO VALORES VALIDOS 0(NAO) E 1(CLIENTE É DO SUPER SIMPLES)
    'If ExisteCampo("SUPER_SIMPLES", "CLIENTE") = False Then
    '    CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD SUPER_SIMPLES INT  NULL"
    '    CONECTA_RETAGUARDA.Execute "UPDATE dbo.CLIENTE SET SUPER_SIMPLES = 1 "
    '    'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
    'End If
    
' CRIAR NA TELA DE CADASTRO DE CLIENTES PARA
' O USUARIO PREENCHEER ESTE CAMPO
'If ExisteCampo("ALIQUOTA_SUBST_SUPER_SIMPLES", "CLIENTE") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CLIENTE ADD ALIQUOTA_SUBST_SUPER_SIMPLES REAL NULL"
'   CONECTA_RETAGUARDA.Execute "UPDATE dbo.CLIENTE SET ALIQUOTA_SUBST_SUPER_SIMPLES = 1 "
   'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
'End If
    
    ' SERÁ OBRIGATÓRIO EM FUNÇÃO DO LUCRO PRESUMIDO E REAL NO SUPER SIMPLES
    ' CASO O CONTADOR SOLICITE ESTAS INFORMAÇÕES DEVERÃO SER IMPRESSAS NA NOTA FISCAL ELETRONICA POR ESTES CAMPOS CRIADOS.
    ' OU VOCE CRIA UMA TELA PAR4A ALIMENTAR AS CST_PIS OU CST_COFINS NO CADASTRO DE CFOP
    ' OU JA CRIA A TELA DE CADASTRO DE CLIENTES PARA
    ' O USUARIO PREENCHEER ESTE CAMPO CONFORME O ORIENTAÇÃO DO CONTADOR
    If ExisteCampo("CST_PIS", "CFOP") = False Then
        CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CST_PIS VARCHAR(03) NULL"
        CONECTA_RETAGUARDA.Execute "UPDATE dbo.CFOP SET CST_PIS = '07' "
        'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
    End If
    
    If ExisteCampo("CST_COFINS", "CFOP") = False Then
        CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CST_COFINS VARCHAR(03) NULL"
        CONECTA_RETAGUARDA.Execute "UPDATE dbo.CFOP SET CST_COFINS = '07' "
        'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
    End If
    
    If ExisteCampo("CODIGO_DA_ALIQUOTA", "CFOP") = False Then
        CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD CODIGO_DA_ALIQUOTA INT NULL"
        CONECTA_RETAGUARDA.Execute "UPDATE dbo.CFOP SET CODIGO_DA_ALIQUOTA = 1 "
        'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
    End If
    'tipo_da_operacao
    If ExisteCampo("TIPO_DA_OPERACAO", "CFOP") = False Then
        CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD TIPO_DA_OPERACAO VARCHAR(01) NULL"
        'CONECTA_RETAGUARDA.Execute "UPDATE dbo.CFOP SET TIPO_DA_OPERACAO = 1 "
        'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
    End If
    'fora_estabelecimento
    If ExisteCampo("FORA_ESTABELECIMENTO", "CFOP") = False Then
        CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.CFOP ADD FORA_ESTABELECIMENTO VARCHAR(01) NULL"
        CONECTA_RETAGUARDA.Execute "UPDATE dbo.CFOP SET FORA_ESTABELECIMENTO = 0 "
        'CONECTA_RETAGUARDA.Execute "ALTER TABLE [dbo].[NF] ADD  CONSTRAINT [DF_NF_MODELO_NF]  DEFAULT ('55') FOR [MODELO_NF]"
    End If

' PARA LUCRO PRESUMIDO E REAL PERCENTUAL DE ENTRADA E SAIDA
' NO SIMPLES SO SE O CONTADOR SO0LICITAR PARA IMPRESSÃO DA NOTA
' ESSES PERCENTUAIS SÃO PARA ALIMENTAR REFERENTES A CFOP DE SAIDA 5101,6102, TODAS REFERENTE A VENDAS E DEVOLUÇÕES
'If ExisteCampo("PER_SAIDA_CST_PIS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PER_SAIDA_CST_PIS REAL NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_PER_SAIDA_CST_PIS]  DEFAULT (0) FOR [PER_SAIDA_CST_PIS] "
'End If

'If ExisteCampo("PER_SAIDA_CST_COFINS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PER_SAIDA_CST_COFINS REAL NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_PER_SAIDA_CST_COFINS]  DEFAULT (0) FOR [PER_SAIDA_CST_COFINS] "
'End If

' ESSES PERCENTUAIS SÃO PARA ALIMENTAR REFERENTES A CFOP DE ENTRADA DIRETAS 1101,1102, TODAS AS ENTRADAS SOMENTE PARA REVENDA E DEVOLUÇÃO
'If ExisteCampo("PER_ENTRADA_CST_PIS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PER_ENTRADA_CST_PIS REAL NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_PER_ENTRADA_CST_PIS]  DEFAULT (0) FOR [PER_ENTRADA_CST_PIS] "
'End If

'If ExisteCampo("PER_ENTRADA_CST_COFINS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD PER_ENTRADA_CST_COFINS REAL NULL"
'   CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PRODUTO ADD CONSTRAINT [DF_PRODUTO_ENTRADA_CST_COFINS]  DEFAULT (0) FOR [PER_ENTRADA_CST_COFINS] "
'End If

'If ExisteCampo("PER_SAIDA_CST_PIS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PER_SAIDA_CST_PIS = 0 "
'End If
'If ExisteCampo("PER_SAIDA_CST_COFINS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PER_SAIDA_CST_COFINS = 0 "
'End If
'If ExisteCampo("PER_ENTRADA_CST_PIS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PER_ENTRADA_CST_PIS = 0 "
'End If
'If ExisteCampo("PER_ENTRADA_CST_COFINS", "PRODUTO") = False Then
'   CONECTA_RETAGUARDA.Execute "UPDATE PRODUTO SET PER_ENTRADA_CST_COFINS = 0 "
'End If

    ' IMPORTANTE NA TELA DE VENDAS VAI SER NECESSÁRIO CRIAR O CAMPO PARA ENTRADA DA CFO
    ' VAMOS ANALISAR COMO FICA ISSO, NOS CASOS DE SIMPLES REMESSA, E OUTROS 01/05/2012

    'FIM em 01/05/2012

'===================================
   TABELAS_RETAGUARDA
'===================================

Exit Sub
ERRO_TRATA:
   MsgBox Err.Description
End Sub
