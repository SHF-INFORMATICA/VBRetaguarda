USE [GLOBAL]
GO
/****** Object:  StoredProcedure [dbo].[spNFePesquisaTipos]    Script Date: 01/10/2017 08:50:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create  PROCEDURE [dbo].[spNFePesquisaTipos] (@Tipo int,
@Impresso int,
--@NaoImpresso int,
@Pesquisa varchar(06),
@TipoPes varchar(01),
@Sequencia INT=0,
@Carga varchar(06)=NULL,
@Nota varchar(06)=NULL,
@Data1 datetime,
@Data2 datetime,
@CodigoFilial varchar(02))

AS

if(@Tipo = 1)
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   --INNER JOIN dbo.MTSITTRIBU Natureza ON Natureza.MTSCODIGO = MFA.MFACODSITT
   --INNER JOIN TIM TIM ON TIM.CodigoEmpresa = MFA.CodigoEmpresa AND TIM.Codigo = MFA.CodigoTIM
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   --AND MFA.MFAFILIAL = Emitente.A1_FILIAL -- O MESMO CLIENTE PARA TODAS AS FILIAIS
   -- INNER JOIN OrdemCarga ON MFA.CodigoEmpresa = OrdemCarga.CodigoEmpresa AND MFA.OrdemCarga = OrdemCarga.Lancamento
   -- INNER JOIN Emitente EmitenteTransporte ON OrdemCarga.CodigoEmitenteTransportador = EmitenteTransporte.CodigoEmitente AND
   -- EmitenteTransporte.CodigoEmpresa = OrdemCarga.CodigoEmpresa
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   --AND EmitenteTransporte.MFTFILIAL = MFA.MFAFILIAL
   --INNER JOIN Veiculo ON OrdemCarga.CodigoVeiculo = Veiculo.Codigo AND Veiculo.CodigoEmpresa = OrdemCarga.CodigoEmpresa
   --INNER JOIN MTACIDADE CidadeTrans ON EmitenteTransporte.MFTCODCID = CidadeTrans.MTACODCID
   --AND EmitenteTransporte.MFTEST = CidadeTrans.MTAUF
   --INNER JOIN MTACIDADE CidadeCliente ON Emitente.A1_CODCIDENT  = CidadeCliente.MTACODCID
   --AND Emitente.A1_EST = CidadeCliente.MTAUF
   --INNER JOIN dbo.MTAUF UFCliente ON CidadeCliente.MTAUF = UFCliente.MTAUF
   --Left outer Join dbo.MFAOBS on dbo.MFAOBS.CodigoEmpresa=MFA.MFACODEMP AND
   --dbo.MFAOBS.CodigoFilial=MFA.MFAFILIAL AND MFAOBS.Sequencia=MFA.MFASEQUENCIA
   --left outer join dbo.ObservacaoNota on dbo.ObservacaoNota.Codigo = dbo.MFAOBS.CodigoObservacao
   where MFA.MFADELETE<>'*' and
   MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and isnull(MFA.MFACODSTAT, 0) = 0
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
Union All
SELECT DISTINCT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         'CCE'+ CASE WHEN LEN(CCE.CODG_CORRECAO)=1 THEN '0'+ CONVERT(CHAR(01),CCE.CODG_CORRECAO) ELSE CONVERT(CHAR(02),CCE.CODG_CORRECAO) END + ' - '+ EMITENTE.A1_NOME AS A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(CCE.COD_STATUS,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(CCE.NUMERO_PROTOCOLO,'') as CodigoProtocolo,
         '' as CodigoRecibo,
         isnull(CCE.CCEMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   inner join dbo.TB_CARTACORRECAO_NFE CCE on CONVERT(varchar(9),NUMR_NOTA)=MFA.MFADOC
   
   where MFA.MFADELETE<>'*'
   --and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and isnull(CCE.COD_STATUS, 0) = 0
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and CCE.REGISTRO_CORRECAO  BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial


--order by MFA.MFASEQUENCIA

End
if(@Tipo = 2) -- 'NF - Aguardando Resposta
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
   FROM dbo.MFA010 MFA
INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
where MFA.MFADELETE<>'*'
   and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and MFA.MFACODSTAT <> 100 AND MFA.MFACODSTAT <> 101 AND MFA.MFACODSTAT <> 102 AND MFA.MFACODSTAT > 0 AND MFA.MFACODSTAT < 200  --isnull(MFA.MFACODSTAT, 0) = 0
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
Union All
SELECT DISTINCT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         'CCE'+ CASE WHEN LEN(CCE.CODG_CORRECAO)=1 THEN '0'+ CONVERT(CHAR(01),CCE.CODG_CORRECAO) ELSE CONVERT(CHAR(02),CCE.CODG_CORRECAO) END + ' - '+ EMITENTE.A1_NOME AS A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         CCE.COD_STATUS AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(CCE.NUMERO_PROTOCOLO,'') as CodigoProtocolo,
         '' as CodigoRecibo,
         isnull(CCE.CCEMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   inner join dbo.TB_CARTACORRECAO_NFE CCE on CONVERT(varchar(9),NUMR_NOTA)=MFA.MFADOC
   
   where MFA.MFADELETE<>'*'
   --and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and CCE.COD_STATUS > 0 AND (CCE.COD_STATUS IN (108,109,129,136))
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and CCE.REGISTRO_CORRECAO  BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial

--order by MFA.MFASEQUENCIA
End

if(@Tipo = 3) --NF - Rejeitados
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   --INNER JOIN dbo.MTSITTRIBU Natureza ON Natureza.MTSCODIGO = MFA.MFACODSITT
   --INNER JOIN TIM TIM ON TIM.CodigoEmpresa = MFA.CodigoEmpresa AND TIM.Codigo = MFA.CodigoTIM
   --AND MFA.MFAFILIAL = Emitente.A1_FILIAL -- O MESMO CLIENTE PARA TODAS AS FILIAIS
   -- INNER JOIN OrdemCarga ON MFA.CodigoEmpresa = OrdemCarga.CodigoEmpresa AND MFA.OrdemCarga = OrdemCarga.Lancamento
   -- INNER JOIN Emitente EmitenteTransporte ON OrdemCarga.CodigoEmitenteTransportador = EmitenteTransporte.CodigoEmitente AND
   -- EmitenteTransporte.CodigoEmpresa = OrdemCarga.CodigoEmpresa
   
   --AND MFA.MFAFILIAL = Emitente.A1_FILIAL -- O MESMO CLIENTE PARA TODAS AS FILIAIS
   -- INNER JOIN OrdemCarga ON MFA.CodigoEmpresa = OrdemCarga.CodigoEmpresa AND MFA.OrdemCarga = OrdemCarga.Lancamento
   -- INNER JOIN Emitente EmitenteTransporte ON OrdemCarga.CodigoEmitenteTransportador = EmitenteTransporte.CodigoEmitente AND
   -- EmitenteTransporte.CodigoEmpresa = OrdemCarga.CodigoEmpresa
   --INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   --AND EmitenteTransporte.MFTFILIAL = MFA.MFAFILIAL
   --INNER JOIN Veiculo ON OrdemCarga.CodigoVeiculo = Veiculo.Codigo AND Veiculo.CodigoEmpresa = OrdemCarga.CodigoEmpresa
   --INNER JOIN MTACIDADE CidadeTrans ON EmitenteTransporte.MFTCODCID = CidadeTrans.MTACODCID
   --AND EmitenteTransporte.MFTEST = CidadeTrans.MTAUF
   --inner JOIN MTACIDADE CidadeCliente ON Emitente.A1_CODCIDADE   = CidadeCliente.MTACODIBGE
   --AND Emitente.A1_EST = CidadeCliente.MTAUF
   --INNER JOIN dbo.MTAUF UFCliente ON CidadeCliente.MTAUF = UFCliente.MTAUF
   --Left outer Join dbo.MFAOBS on dbo.MFAOBS.CodigoEmpresa=MFA.MFACODEMP AND
   --dbo.MFAOBS.CodigoFilial=MFA.MFAFILIAL AND MFAOBS.Sequencia=MFA.MFASEQUENCIA
   --left outer join dbo.ObservacaoNota on dbo.ObservacaoNota.Codigo = dbo.MFAOBS.CodigoObservacao
   where MFA.MFADELETE<>'*'
   and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and MFA.MFACODSTAT > 200  --isnull(MFA.MFACODSTAT, 0) = 0
and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
--and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
--and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
--and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
--and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
--and MFA.MFAFILIAL = @CodigoFilial
Union All
SELECT DISTINCT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         'CCE'+ CASE WHEN LEN(CCE.CODG_CORRECAO)=1 THEN '0'+ CONVERT(CHAR(01),CCE.CODG_CORRECAO) ELSE CONVERT(CHAR(02),CCE.CODG_CORRECAO) END + ' - '+ EMITENTE.A1_NOME AS A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         CCE.COD_STATUS AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(CCE.NUMERO_PROTOCOLO,'') as CodigoProtocolo,
         '' as CodigoRecibo,
         isnull(CCE.CCEMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   inner join dbo.TB_CARTACORRECAO_NFE CCE on CONVERT(varchar(9),NUMR_NOTA)=MFA.MFADOC
   
   where MFA.MFADELETE<>'*'
   --and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and CCE.COD_STATUS > 0 AND (CCE.COD_STATUS >= 203)
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and CCE.REGISTRO_CORRECAO  BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
--order by MFA.MFASEQUENCIA
End
if(@Tipo = 4) -- NF - Contingencia
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
   FROM dbo.MFA010 MFA
    INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   where MFA.MFADELETE<>'*'
   and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and MFA.MFACODSTAT = 200  --isnull(MFA.MFACODSTAT, 0) = 0
and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
order by MFA.MFASEQUENCIA
End
if(@Tipo = 5) -- NF - Aprovadas
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
    FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
      
   where MFA.MFADELETE<>'*'
   and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and MFA.MFACODSTAT = 100
   --and MFAFIMP=case when @Impresso=1 then 'S' else '' END
   --and MFAFIMP=case when @NaoImpresso=1 then '' else 'N' END
and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
--UNION ALL

End
if(@Tipo = 6)
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   where MFA.MFADELETE='*' and
   MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
order by MFA.MFASEQUENCIA

End



if(@Tipo = 7) -- NF - INUTILIZADAS
BEGIN
   SELECT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         EMITENTE.A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         isnull(MFA.MFACODSTAT,0) AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(MFA.MFACODPROT,'') as CodigoProtocolo,
         isnull(MFA.MFACODRECI,'') as CodigoRecibo,
         isnull(MFA.MFAMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   where MFA.MFADELETE<>'*'
   and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and MFA.MFACODSTAT = 102  --isnull(MFA.MFACODSTAT, 0) = 0
and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and MFA.MFADTDIGIT BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
order by MFA.MFASEQUENCIA
End
if(@Tipo = 8) -- CCE - APROVADAS
BEGIN
   SELECT DISTINCT   --MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
         MFA.MFASEQUENCIA,
         MFA.MFAFILIAL,
         MFA.MFADOC,
         MFA.MFADTDIGIT,
         EMITENTE.A1_COD ,
         EMITENTE.A1_LOJA,
         'CCE'+ CASE WHEN LEN(CCE.CODG_CORRECAO)=1 THEN '0'+ CONVERT(CHAR(01),CCE.CODG_CORRECAO) ELSE CONVERT(CHAR(02),CCE.CODG_CORRECAO) END + ' - '+ EMITENTE.A1_NOME AS A1_NOME,
         MFA.MFAVALLIQUI,
         MFA.MFACODPROT,
         MFA.MFACODMORE,
         MFA.MFACODRECI,
         MFA.MFACHAVENFE,
         CCE.COD_STATUS AS CodigoStatus ,
         isnull(MFA.MFACHAVENFE,'') as ChaveNFe,
         isnull(CCE.NUMERO_PROTOCOLO,'') as CodigoProtocolo,
         '' as CodigoRecibo,
         isnull(CCE.CCEMOTRESU,'') AS MotivoResultado,
         isnull(EMITENTE.A1_EMAIL,'') AS Email,
         isnull(MFA.MFAEMAILENVIADO,'') AS EmailEnviado,
         MFA.MFAPREFIXO AS Modelo -- Modelo do Documento Fiscal NFE OU NFC
         
   FROM dbo.MFA010 MFA
   INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
   INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE --AND Filial.ENT_ESTADO = cidade.MTAUF
   INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF
   
   INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
   
   INNER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
   inner join dbo.TB_CARTACORRECAO_NFE CCE on CONVERT(varchar(9),NUMR_NOTA)=MFA.MFADOC
   
   where MFA.MFADELETE<>'*'
   --and MFA.MFATIPO = @TipoPes --Tipos Pode ser N=Venda;D=Devolução de Venda;C=Complemento
   and CCE.COD_STATUS = 135
   and MFA.MFADOC = CASE WHEN ISNULL(@Nota,0)<>0 then @Nota ELSE MFA.MFADOC END
and MFA.MFASEQUENCIA = CASE WHEN ISNULL(@Sequencia,0)<>0 then @Sequencia  ELSE MFA.MFASEQUENCIA END
and MFA.MFACARGA = CASE WHEN ISNULL(@Carga,0)<>0 then @Carga ELSE MFA.MFACARGA END
and CCE.REGISTRO_CORRECAO  BETWEEN @Data1 AND @Data2
and MFA.MFAFILIAL = @CodigoFilial
order by MFA.MFASEQUENCIA
    --isnull(MFA.MFACODSTAT, 0) = 0
End
