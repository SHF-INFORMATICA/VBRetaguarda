select numr_req from CABECAREQ
where status in (3,5)
and NUMR_REQ not in (select numr_doc from LANCAMENTO)
order by NUMR_REQ