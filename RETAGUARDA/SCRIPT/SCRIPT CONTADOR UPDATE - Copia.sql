DECLARE @Contador AS SMALLINT

SET @Contador = 0

--WHILE @Contador <= 10

--  BEGIN

--    SELECT @Contador

--    SET @Contador = @Contador + 1

--  END	


update pedidoitem 
SET @Contador = @Contador + 1
, SEQ_ID = @Contador 
select * from PEDIDOITEM order by pedido_id