
----====================TABELA FONE
--ALTER TABLE FONE DROP CONSTRAINT PK_FONE
--ALTER TABLE FONE DROP COLUMN FONE_ID
IF (select isnull(count(sysobjects.name),0) AS qtde from sysobjects WHERE sysobjects.name = 'spInsertRegIDTabelaFONE') = 1 begin
   drop procedure spInsertRegIDTabelaFONE
end
go

if (select isnull(count(syscolumns.name),0) as qtde from sysobjects 
			  INNER JOIN syscolumns ON sysobjects.id = syscolumns.id WHERE (sysobjects.xtype = 'U') 
			  and sysobjects.name = 'FONE' and syscolumns.name = 'RAMAL') = 1
begin
	ALTER TABLE FONE DROP COLUMN RAMAL
END
GO

if (select isnull(count(syscolumns.name),0) as qtde from sysobjects 
			  INNER JOIN syscolumns ON sysobjects.id = syscolumns.id WHERE (sysobjects.xtype = 'U') 
			  and sysobjects.name = 'FONE' and syscolumns.name = 'FONE_ID') = 0 
begin
	ALTER TABLE FONE ADD FONE_ID BIGINT
END
GO

	DECLARE @Contador	AS smallint
	SET @Contador = 0

	update FONE SET @Contador = @Contador + 1, FONE_id = @Contador 
	select * from FONE order by FONE_id

	ALTER TABLE FONE ALTER COLUMN FONE_ID BIGINT NOT NULL

GO

ALTER TABLE FONE ADD CONSTRAINT pk_FONE PRIMARY KEY (FONE_ID)
GO
--================================================================================================
----====================TABELA EMAIL
--ALTER TABLE EMAIL DROP CONSTRAINT PK_EMAIL
--ALTER TABLE EMAIL DROP COLUMN EMAIL_ID
IF (select isnull(count(sysobjects.name),0) AS qtde from sysobjects WHERE sysobjects.name = 'spInsertRegIDTabelaEMAIL') = 1 begin
   drop procedure spInsertRegIDTabelaEMAIL
end
go

if (select isnull(count(syscolumns.name),0) as qtde from sysobjects 
			  INNER JOIN syscolumns ON sysobjects.id = syscolumns.id WHERE (sysobjects.xtype = 'U') 
			  and sysobjects.name = 'EMAIL' and syscolumns.name = 'EMAIL_ID') = 0 
begin
	ALTER TABLE EMAIL ADD EMAIL_ID BIGINT
END
GO

	DECLARE @Contador	AS smallint
	SET @Contador = 0

	update EMAIL SET @Contador = @Contador + 1, EMAIL_id = @Contador 
	select * from EMAIL order by EMAIL_id

	ALTER TABLE EMAIL ALTER COLUMN EMAIL_ID BIGINT NOT NULL

GO

ALTER TABLE EMAIL ADD CONSTRAINT pk_EMAIL PRIMARY KEY (EMAIL_ID)
GO
