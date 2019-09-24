
CREATE PROCEDURE [dbo].[spInsertRegIDTabela] (@NomeTabela nVARCHAR(50), @NomeCampo nvarchar(50))
AS

if(@NomeTabela <> '' and @NomeCampo <> '')
begin

	DECLARE @Contador AS SMALLINT

	SET @Contador = 0

--WHILE @Contador <= 10

--  BEGIN

--    SELECT @Contador

--    SET @Contador = @Contador + 1

--  END	

	update @NomeTabela
	SET @Contador = @Contador + 1, @NomeCampo = @Contador 
end