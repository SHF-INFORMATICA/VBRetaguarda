use loja
SELECT     CABECAREQ.NUMR_REQ, CABECAREQ.VENDEDOR, CABECAREQ.CODG_USU, CABECAREQ.STATUS, CABECAREQ.DT_REQ, ITEMREQ.CODG_PROD, 
                      ITEMREQ.QTD_PEDIDA, ITEMREQ.STATUS AS Expr1
FROM         CABECAREQ INNER JOIN
                      ITEMREQ ON CABECAREQ.NUMR_REQ = ITEMREQ.NUMR_REQ
where CODG_PROD = 4652
and CABECAREQ.STATUS <> 9
order by DT_REQ desc

SELECT SUM(ITEMREQ.QTD_PEDIDA) as Venda
FROM         CABECAREQ INNER JOIN
                      ITEMREQ ON CABECAREQ.NUMR_REQ = ITEMREQ.NUMR_REQ
where CODG_PROD = 4652
and CABECAREQ.STATUS <> 9