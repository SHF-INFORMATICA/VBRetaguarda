USE [GLOBAL]
GO
/****** Object:  StoredProcedure [dbo].[spNFeCabeNota]    Script Date: 08/10/2017 17:44:50 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
ALTER  PROCEDURE [dbo].[spNFeCabeNota] (@CodigoEmpresa VARCHAR(02), @Sequencia int, @CCE int)
AS

if(@CCE=0)
begin
SELECT top 1 	MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
	  	MFA.MFASEQUENCIA as cNF,
		rtrim(dbo.funSubstituiCaracterEspe(MFACODSITT)) as natOp,
		(case when MFA.MFAINDPAG=1 or MFA.MFAINDPAG=4 then '0' when MFA.MFAINDPAG=5 then '1' else '2'  end) as  indPag,
		(case when MFA.MFAPREFIXO='NFE' then '55'  else '65'  end) as  mod,
		(case when isnull(MFA.MFASERIE,0)=0 then '1' else MFA.MFASERIE end)  as serie,
		MFA.MFADOC as nNF,
		convert(varchar(04),DATEPART(year,MFA.MFADTDIGIT))+'-'+case when len(DATEPART(month,MFA.MFADTDIGIT))=1 then '0'+Convert(varchar(02),DATEPART(month,MFA.MFADTDIGIT)) else Convert(varchar(02),DATEPART(month,MFA.MFADTDIGIT)) end +'-'+ case when len(DATEPART(day,MFA.MFADTDIGIT))=1 then '0'+Convert(varchar(02),DATEPART(day,MFA.MFADTDIGIT)) else Convert(varchar(02),DATEPART(day,MFA.MFADTDIGIT)) end as dEmi,
		--(case when MFA.MFAPREFIXO='NFE' then convert(varchar(04),DATEPART(year,MFA.MFADTDIGIT))+'-'+case when len(DATEPART(month,MFA.MFADTDIGIT ))=1 then '0'+Convert(varchar(02),DATEPART(month,MFA.MFADTDIGIT)) else Convert(varchar(02),DATEPART(month,MFA.MFADTDIGIT)) end +'-'+ case when len(DATEPART(day,MFA.MFADTDIGIT))=1 then '0'+Convert(varchar(02),DATEPART(day,MFA.MFADTDIGIT)) else Convert(varchar(02),DATEPART(day,MFA.MFADTDIGIT)) end else MFA.MFADTDIGIT end)  as dEmi,
		convert(varchar(04),DATEPART(year,MFA.MFADTENSAI))+'-'+case when len(DATEPART(month,MFA.MFADTENSAI))=1 then '0'+Convert(varchar(02),DATEPART(month,MFA.MFADTENSAI)) else Convert(varchar(02),DATEPART(month,MFA.MFADTENSAI)) end +'-'+ case when len(DATEPART(day,MFA.MFADTENSAI))=1 then '0'+Convert(varchar(02),DATEPART(day,MFA.MFADTENSAI)) else Convert(varchar(02),DATEPART(day,MFA.MFADTENSAI)) end as dSaiEnt,
		-- (Case when (MFA.MFATIPO='O' or MFA.MFATIPO='D') THEN '0' when (MFA.MFATIPO='N' or MFA.MFATIPO='P')  then '1' end) as tpNF,
		--(case when MFA.MFATIPO='N' then '1' when MFA.MFATIPO='D' then '0'   else '1'  end) as  tpNF, -- N= Venda Normal ; D = Devolucão de Venda; C = Devolução de Compra; versao nova 2017 tip0o movimento
		MFA.MFATIPO as tpNF,-- N= Venda Normal ; D = Devolucão de Venda; C = Devolução de Compra; versao nova 2017 tip0o movimento
		MFA.MFAIDDEST AS idDest,
		Cidade.MTACODIBGE as cMunFG,
		(case when MFA.MFAPREFIXO='NFE' then '1'  else '4'  end) as tpImp,
		--tpEmis INICIO
		--1=Emissão normal (não em contingência);
		--2=Contingência FS-IA, com impressão do DANFE em formulário de segurança;
		--3=Contingência SCAN (Sistema de Contingência do Ambiente Nacional);
		--4=Contingência DPEC (Declaração Prévia da Emissão em Contingência);
		--5=Contingência FS-DA, com impressão do DANFE em formulário de segurança;
		--6=Contingência SVC-AN (SEFAZ Virtual de Contingência do AN);
		--7=Contingência SVC-RS (SEFAZ Virtual de Contingência do RS);
		--9=Contingência off-line da NFC-e (as demais opções de contingência são válidas também para a NFC-e).
        --Para a NFC-e somente estão disponíveis e são válidas as opções de contingência 5 e 9.
		MFA.MFATIPOREM as tpEmis, -- vER COM SERGIO 
		-- FIM
		'1' as tpAmb,
		MFAFINNFE as finNFe,
		'0' as procEmi,
		'MEGASIM V2017' as verProc,
		--UF.CodigoIbge AS cUF, 
		dbo.funApenasNumeros(Filial.CNPJ) as CNPJ, 
		-- Filial.CPFDiretor, 
		rtrim(dbo.funSubstituiCaracterEspe(Filial.NOME_EMPRESA)) as xNome, 
		rtrim(dbo.funSubstituiCaracterEspe(Filial.NOME_REDUZ)) as xFant, 		
		rtrim(dbo.funSubstituiCaracterEspe(Filial.ENT_RUA_AV)) as xLgr, 
		'SEM' as nro,
		rtrim(dbo.funSubstituiCaracterEspe(Filial.ENT_BAIRRO)) as xBairro , 
		Cidade.MTACODIBGE as cMun, 
		rtrim(dbo.funSubstituiCaracterEspe(Cidade.MTADESC)) as xMun,
		UF.MTAUF  as UF ,
		'1058' as cPais ,
		'BRASIL' AS xPais,
		dbo.funApenasNumeros(Filial.INSC_ESTADUAL) as IE,
		dbo.funApenasNumeros(Emitente.A1_CGC) as destCNPJ, 
		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_NOME)) as destxNome, 
		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_END)) as enderDestxLgr , 
		CASE Emitente.A1_NUMERO  WHEN ''   THEN 'SEM' ELSE Emitente.A1_NUMERO END  as  enderDestnro, 
		--Emitente.A1_NUMERO  as  enderDestnro, 
		case Emitente.A1_BAIRRO  when ''   then 'SEM' else rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_BAIRRO)) end as enderDestxBairro,
        dbo.funApenasNumeros(Emitente.A1_CEP) as destCEP,
		dbo.funApenasNumeros(Emitente.A1_TEL) as destfone,
		CidadeCliente.MTACODIBGE as enderDestcMun, 
		rtrim(dbo.funSubstituiCaracterEspe(CidadeCliente.MTADESC)) as enderDestxMun,
		CidadeCliente.MTAUF as enderDestUF,
		'1058' as enderDestcPais,            
	        	'BRASIL' as enderDestxPais,
		(CASE UPPER(ltrim(rtrim(Emitente.A1_INSCR)))  WHEN 'ISENTA' THEN 'ISENTO' ELSE rtrim(dbo.funSubstituiCaracterEspeTira(Emitente.A1_INSCR)) end) as enderdestRGInscEst,
		--Emitente.A1_INSCR,
		dbo.funApenasNumeros(A1_ENENTCGC) as entregaCNPJ,
		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_ENDENT)) as entregaxLgr,
		Emitente.A1_ENDENTNR as entreganro,
        		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_BAIRROE)) as entregaxBairro,
		CidadeCliente.MTACODIBGE as entregacMun,
		rtrim(dbo.funSubstituiCaracterEspe(CidadeCliente.MTADESC)) as entregaxMun,
        		CidadeCliente.MTAUF as entregaUF,
		rtrim(dbo.funSubstituiCaracterEspe(left(MFA.MFAOBSNOTA,200))) as ObservacaoNota,
                	MFA.MFABASEICM AS BASEICMS,
                	MFA.MFAVALICM AS ValorICMS,
		MFA.MFABASEIPI AS BaseIPI,
		MFA.MFAVALIPI AS ValorIPI,
		MFA.MFATIFRETE as CIFFOB,
		isnull((SELECT COUNT(MFIREGISTRO) as Total FROM dbo.MFI010 WHERE dbo.MFI010.MFIFILIAL = MFA.MFAFILIAL  and dbo.MFI010.MFISEQUEN = MFA.MFASEQUENCIA),0) as Item,
		--ISNULL(MFA.MFAVALTOT,0) AS Item,
		isnull(MFA.MFABASICMST, 0) as VlBaseSubstituicao,
		isnull(MFA.MFAVALICMST, 0) as VlSubstituicao,
		isnull(MFA.MFAVALLIQUI, 0) as ValorLiquido,
		MFA.MFACODSTAT AS CodigoStatus,
		MFA.MFANFECNF as CodigoNumerico,
		Filial.CRT as CRT, -- criado em 24/03/2011
		CONVERT(VARCHAR(12),MFA.MFANFEEMISE,110) AS hSaiEnt, -- criado em 02/11/2011
		Filial.ENT_CEP  as CEP,  -- criado em 02/11/2011
        Emitente.A1_EMAIL AS email  ,
		left(MFA.MFACHAVEREFNFE,44) AS refNFe ,
		MFAINDFINAL AS indFinal,
		[MFAINDPRES] AS indPres,
		MFACODEMP as Empresa,
		MFAFILIAL AS Filial
FROM dbo.MFA010 MFA 
INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE 
--AND Filial.ENT_ESTADO = cidade.MTAUF 
INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF 
--INNER JOIN dbo.MTSITTRIBU Natureza ON Natureza.MTSCODIGO = MFA.MFACODSITT
INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
LEFT OUTER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
AND EmitenteTransporte.MFTFILIAL = MFA.MFAFILIAL
left outer JOIN MTACIDADE CidadeTrans ON EmitenteTransporte.MFTCODCID = CidadeTrans.MTACODIBGE
AND EmitenteTransporte.MFTEST = CidadeTrans.MTAUF 
inner JOIN MTACIDADE CidadeCliente ON Emitente.A1_CODMUN   = CidadeCliente.MTACODIBGE
--inner JOIN MTACIDADE CidadeCliente ON Emitente.A1_CODCIDADE  = CidadeCliente.MTACODIBGE
--AND Emitente.A1_EST = CidadeCliente.MTAUF  
INNER JOIN dbo.MTAUF UFCliente ON CidadeCliente.MTAUF = UFCliente.MTAUF 
where MFA.MFACODEMP =@CodigoEmpresa  and MFA.MFASEQUENCIA = @Sequencia
end
if(@CCE=1)
begin
SELECT top 1 	MFA.MFACLIENTE AS CodigoEmitente,UF.MTACODIIBGE AS cUF,
	  	MFA.MFASEQUENCIA as cNF,
		rtrim(dbo.funSubstituiCaracterEspe(MFACODSITT)) as natOp,
		(case when MFA.MFAINDPAG=1 or MFA.MFAINDPAG=4 then '0' when MFA.MFAINDPAG=5 then '1' else '2'  end) as  indPag,
		(case when MFA.MFAPREFIXO='NFE' then '55'  else '65'  end) as  mod,
		(case when isnull(MFA.MFASERIE,0)=0 then '1' else MFA.MFASERIE end)  as serie,
		MFA.MFADOC as nNF,
		convert(varchar(04),DATEPART(year,MFA.MFADTDIGIT))+'-'+case when len(DATEPART(month,MFA.MFADTDIGIT ))=1 then '0'+Convert(varchar(02),DATEPART(month,MFA.MFADTDIGIT)) else Convert(varchar(02),DATEPART(month,MFA.MFADTDIGIT)) end +'-'+ case when len(DATEPART(day,MFA.MFADTDIGIT))=1 then '0'+Convert(varchar(02),DATEPART(day,MFA.MFADTDIGIT)) else Convert(varchar(02),DATEPART(day,MFA.MFADTDIGIT)) end as dEmi,
		convert(varchar(04),DATEPART(year,MFA.MFADTENSAI))+'-'+case when len(DATEPART(month,MFA.MFADTENSAI))=1 then '0'+Convert(varchar(02),DATEPART(month,MFA.MFADTENSAI)) else Convert(varchar(02),DATEPART(month,MFA.MFADTENSAI)) end +'-'+ case when len(DATEPART(day,MFA.MFADTENSAI))=1 then '0'+Convert(varchar(02),DATEPART(day,MFA.MFADTENSAI)) else Convert(varchar(02),DATEPART(day,MFA.MFADTENSAI)) end as dSaiEnt,
		-- (Case when (MFA.MFATIPO='O' or MFA.MFATIPO='D') THEN '0' when (MFA.MFATIPO='N' or MFA.MFATIPO='P')  then '1' end) as tpNF,
		MFA.MFATIPO as tpNF,
		MFA.MFAIDDEST AS idDest,
		Cidade.MTACODIBGE as cMunFG,
		'1' as tpImp,
		'1' as tpEmis,
		'1' as tpAmb,
		MFAFINNFE as finNFe,
		'0' as procEmi,
		'MEGASIM V2017' as verProc,
		--UF.CodigoIbge AS cUF, 
		dbo.funApenasNumeros(Filial.CNPJ) as CNPJ, 
		-- Filial.CPFDiretor, 
		rtrim(dbo.funSubstituiCaracterEspe(Filial.NOME_EMPRESA)) as xNome, 
		rtrim(dbo.funSubstituiCaracterEspe(Filial.NOME_REDUZ)) as xFant, 		
		rtrim(dbo.funSubstituiCaracterEspe(Filial.ENT_RUA_AV)) as xLgr, 
		'SEM' as nro,
		rtrim(dbo.funSubstituiCaracterEspe(Filial.ENT_BAIRRO)) as xBairro , 
		Cidade.MTACODIBGE as cMun, 
		rtrim(dbo.funSubstituiCaracterEspe(Cidade.MTADESC)) as xMun,
		UF.MTAUF  as UF ,
		'1058' as cPais ,
		'BRASIL' AS xPais,
		dbo.funApenasNumeros(Filial.INSC_ESTADUAL) as IE,
		dbo.funApenasNumeros(Emitente.A1_CGC) as destCNPJ, 
		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_NOME)) as destxNome, 
		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_END)) as enderDestxLgr , 
		CASE Emitente.A1_NUMERO  WHEN ''   THEN 'SEM' ELSE Emitente.A1_NUMERO END  as  enderDestnro, 
		--Emitente.A1_NUMERO  as  enderDestnro, 
		case Emitente.A1_BAIRRO  when ''   then 'SEM' else rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_BAIRRO)) end as enderDestxBairro,
        dbo.funApenasNumeros(Emitente.A1_CEP) as destCEP,
		dbo.funApenasNumeros(Emitente.A1_TEL) as destfone,
		CidadeCliente.MTACODIBGE as enderDestcMun, 
		rtrim(dbo.funSubstituiCaracterEspe(CidadeCliente.MTADESC)) as enderDestxMun,
		CidadeCliente.MTAUF as enderDestUF,
		'1058' as enderDestcPais,            
	        	'BRASIL' as enderDestxPais,
		(CASE UPPER(ltrim(rtrim(Emitente.A1_INSCR)))  WHEN 'ISENTA' THEN 'ISENTO' ELSE rtrim(dbo.funSubstituiCaracterEspeTira(Emitente.A1_INSCR)) end) as enderdestRGInscEst,
		--Emitente.A1_INSCR,
		dbo.funApenasNumeros(A1_ENENTCGC) as entregaCNPJ,
		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_ENDENT)) as entregaxLgr,
		Emitente.A1_ENDENTNR as entreganro,
        		rtrim(dbo.funSubstituiCaracterEspe(Emitente.A1_BAIRROE)) as entregaxBairro,
		CidadeCliente.MTACODIBGE as entregacMun,
		rtrim(dbo.funSubstituiCaracterEspe(CidadeCliente.MTADESC)) as entregaxMun,
        		CidadeCliente.MTAUF as entregaUF,
		rtrim(dbo.funSubstituiCaracterEspe(left(MFA.MFAOBSNOTA,200))) as ObservacaoNota,
                	MFA.MFABASEICM AS BASEICMS,
                	MFA.MFAVALICM AS ValorICMS,
		MFA.MFABASEIPI AS BaseIPI,
		MFA.MFAVALIPI AS ValorIPI,
		MFA.MFATIFRETE as CIFFOB,
		isnull((SELECT COUNT(MFIREGISTRO) as Total FROM dbo.MFI010 WHERE dbo.MFI010.MFIFILIAL = MFA.MFAFILIAL  and dbo.MFI010.MFISEQUEN = MFA.MFASEQUENCIA),0) as Item,
		--ISNULL(MFA.MFAVALTOT,0) AS Item,
		isnull(MFA.MFABASICMST, 0) as VlBaseSubstituicao,
		isnull(MFA.MFAVALICMST, 0) as VlSubstituicao,
		isnull(MFA.MFAVALLIQUI, 0) as ValorLiquido,
		MFA.MFACODSTAT AS CodigoStatus,
		MFA.MFANFECNF as CodigoNumerico,
		Filial.CRT as CRT, -- criado em 24/03/2011
		CONVERT(VARCHAR(8),MFA.MFANFEEMISE,108) AS hSaiEnt, -- criado em 02/11/2011
		Filial.ENT_CEP  as CEP,  -- criado em 02/11/2011
		CCE.DADOS_CORRECAO AS xCorrecao,
		cce.CODG_CORRECAO as nSeqEvento,
		convert(varchar(04),DATEPART(year,CCE.REGISTRO_CORRECAO))+'-'+case when len(DATEPART(month,CCE.REGISTRO_CORRECAO ))=1 then '0'+Convert(varchar(02),DATEPART(month,CCE.REGISTRO_CORRECAO)) else Convert(varchar(02),DATEPART(month,CCE.REGISTRO_CORRECAO)) end +'-'+ case when len(DATEPART(day,CCE.REGISTRO_CORRECAO))=1 then '0'+Convert(varchar(02),DATEPART(day,CCE.REGISTRO_CORRECAO)) else Convert(varchar(02),DATEPART(day,CCE.REGISTRO_CORRECAO))+'T'+CONVERT(VARCHAR(8),CCE.REGISTRO_CORRECAO,108) end as dhEvento,
--		CCE.REGISTRO_CORRECAO as dhEvento,
		'110110' AS tpEvento,
		mfa.MFACHAVENFE as chNFe,
		Emitente.A1_EMAIL AS email ,
		left(MFA.MFACHAVEREFNFE,44) AS refNFe ,
		MFAINDFINAL AS indFinal,
		[MFAINDPRES] AS indPres,
		MFACODEMP as Empresa,
		MFAFILIAL AS Filial,
		MFAVALIRRF as TotalTributos

FROM dbo.MFA010 MFA 
INNER JOIN dbo.EMPRES Filial ON MFA.MFACODEMP = Filial.EMPRESA AND MFA.MFAFILIAL = Filial.FILIAL
INNER JOIN dbo.MTACIDADE cidade ON Filial.ENT_CODCID = cidade.MTACODIBGE 
--AND Filial.ENT_ESTADO = cidade.MTAUF 
INNER JOIN dbo.MTAUF UF ON cidade.MTAUF = UF.MTAUF 
--INNER JOIN dbo.MTSITTRIBU Natureza ON Natureza.MTSCODIGO = MFA.MFACODSITT
INNER JOIN dbo.SA1010 Emitente ON Emitente.A1_COD = MFA.MFACLIENTE
LEFT OUTER JOIN dbo.MFT010 EmitenteTransporte ON MFA.MFATRANSP = EmitenteTransporte.MFTCOD
AND EmitenteTransporte.MFTFILIAL = MFA.MFAFILIAL
left outer JOIN MTACIDADE CidadeTrans ON EmitenteTransporte.MFTCODCID = CidadeTrans.MTACODIBGE
AND EmitenteTransporte.MFTEST = CidadeTrans.MTAUF 
inner JOIN MTACIDADE CidadeCliente ON Emitente.A1_CODMUN   = CidadeCliente.MTACODIBGE
--inner JOIN MTACIDADE CidadeCliente ON Emitente.A1_CODCIDADE  = CidadeCliente.MTACODIBGE
--AND Emitente.A1_EST = CidadeCliente.MTAUF  
INNER JOIN dbo.MTAUF UFCliente ON CidadeCliente.MTAUF = UFCliente.MTAUF 
inner join dbo.TB_CARTACORRECAO_NFE CCE on CONVERT(varchar(9),NUMR_NOTA)=MFA.MFADOC
where MFA.MFACODEMP =@CodigoEmpresa  and MFA.MFASEQUENCIA = @Sequencia
end
