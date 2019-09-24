EXEC sp_resetstatus 'loja1'
GO
ALTER DATABASE loja1 SET EMERGENCY
DBCC checkdb('loja1')
ALTER DATABASE loja1 SET SINGLE_USER WITH ROLLBACK IMMEDIATE
DBCC CheckDB ('loja1', REPAIR_ALLOW_DATA_LOSS)
ALTER DATABASE loja1 SET MULTI_USER

GO

-- Rebuild the index

ALTER INDEX ALL ON [ TableName] REBUILD