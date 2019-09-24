DECLARE @Contador AS SMALLINT

SET @Contador = 0

--WHILE @Contador <= 10

--  BEGIN

--    SELECT @Contador

--    SET @Contador = @Contador + 1

--  END	


update PEDIDOITEM 
SET @Contador = @Contador + 1
, PEDIDOITEM_ID = @Contador 
select * from PEDIDOITEM order by PEDIDOITEM_id