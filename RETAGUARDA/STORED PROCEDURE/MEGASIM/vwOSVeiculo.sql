--USE [MEGASIM]
GO

drop view vwOSVeiculo

/****** Object:  View [dbo].[vwOSVeiculo]    Script Date: 29/07/2019 11:17:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vwOSVeiculo] AS  SELECT OSVEICULO.VEICULO_ID, OSVEICULO.COMBUSTIVEL_ID, OSVEICULO.PLACA, OSVEICULO.DESCRICAO AS DESCRICAOVEICULO, OSVEICULO.MOTOR, OSVEICULO.CHASSI, OS.OS_ID, OS.ESTABELECIMENTO_ID, OS.PESSOA_ID, OS.CT_ID, OS.DT_OS, OS.DT_FECHA, OS.TIPO_OS, OS.SITUACAO_OS, OS.KM, OS.CLIENTE, OSPECA.OSPECA_ID, OSPECA.PRODUTO_ID, OSPECA.DT_CAD AS DTCADPECA, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, OSPECA.QTDE, OSPECA.DT_GARANTIA, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO AS DESCRICAOPRODUTO, PRODUTO.REFERENCIA, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, OSSERVICO.OSSERVICO_ID, OSSERVICO.OSTAREFA_ID, OSSERVICO.DT_CAD AS DTCADSERVICO, OSSERVICO.SITUACAO, OSSERVICO.RESPONSAVEL_ID, OSSERVICO.VALOR_SERVICO, OSSERVICO.DT_FIM, OSSERVICO.DT_INICIO, OSSERVICO.DT_FECHA AS DTFECHASERVICO, OSSERVICO.DESCONTO_SERVICO, OSTAREFA.DT_CAD AS DTCADTAREFA, OSSERVICO.DESCRICAO AS DESCRICAOSERVICO, OSAPONTAMENTO.APONTAMENTO_ID, OSAPONTAMENTO.DATAINICIAL, OSAPONTAMENTO.DATAFINAL, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS DESCRICAOPESSOA, PESSOA.RAZAO, PESSOA.SITUACAO AS SITUACAOPESSOA FROM OS WITH (NOLOCK) INNER JOIN PESSOA WITH (NOLOCK) ON OS.PESSOA_ID = PESSOA.PESSOA_ID INNER JOIN OSVEICEQP ON OS.OS_ID = OSVEICEQP.OS_ID INNER JOIN OSVEICULO WITH (NOLOCK) ON OSVEICEQP.VEICULO_ID = OSVEICULO.VEICULO_ID LEFT OUTER JOIN OSTAREFA WITH (NOLOCK) INNER JOIN OSSERVICO WITH (NOLOCK) ON OSTAREFA.OSTAREFA_ID = OSSERVICO.OSTAREFA_ID INNER JOIN OSAPONTAMENTO WITH (NOLOCK) ON OSTAREFA.OSTAREFA_ID = OSAPONTAMENTO.OSTAREFA_ID ON OS.OS_ID = OSSERVICO.OS_ID LEFT OUTER JOIN OSPECA WITH (NOLOCK) INNER JOIN PRODUTO WITH (NOLOCK) ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID ON OS.OS_ID = OSPECA.OS_ID
GO


