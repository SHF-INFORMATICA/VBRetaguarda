use psloja

select * from FORMAPAGTO

UPDATE ITEMLANCAMENTO SET Status = 'B'
where FORMAPAGTO_ID = 13 and status = 'A'


update itemLANCAMENTO set dt_baixa = DT_VENCIMENTO
where dt_baixa Is Null 
and status = 'B' 