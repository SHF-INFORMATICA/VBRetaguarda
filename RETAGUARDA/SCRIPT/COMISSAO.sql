use MEGASIM
/* ============================================================ */
/*   Database name:  MODEL_2                                    */
/*   DBMS name:      Microsoft SQL Server 6.x                   */
/*   Created on:     26/02/2012  12:45                          */
/* ============================================================ */

/* ============================================================ */
/*   Table: COMISSAOPARAM                                       */
/* ============================================================ */
create table COMISSAOPARAM
(
    COMISSAOPARAM_ID  int                   not null,
    TIPO              varchar(30)           not null,
    VALOR             float                 not null,
    constraint PK_COMISSAOPARAM primary key (COMISSAOPARAM_ID)
)
go

drop table COMISSAOITEM
drop table comissao
drop table CONSULTAGRIDCABECA
drop table CONSULTAGRIDITEM

/* ============================================================ */
/*   Table: COMISSAO                                            */
/* ============================================================ */
create table COMISSAO
(
    COMISSAO_ID       int                   not null,
    EMPRESA_ID        int                   not null,
    DTINI             datetime              not null,
    DTFIM             datetime              not null,
    SITUACAO          varchar(30)           not null,
    constraint PK_COMISSAO primary key (COMISSAO_ID)
)
go

/* ============================================================ */
/*   Table: COMISSAOITEM                                        */
/* ============================================================ */
create table COMISSAOITEM
(
    COMISSAO_ID       int                   not null,
    CODG_VEND         INT                   not null,
    CODG_PROD         varchar(30)           not null,
    PR_ITEM_VENDA     float                 null    ,
    PR_ITEM_VAREJO    float                 null    ,
    PR_ITEM_ATACADO   float                 null    ,
    VALR_COMIS_PROD   float                 null    ,
    VALR_COMIS_TOT    float                 null    ,
    QTDE_VENDIDA      float                 null    ,
    NUMR_REQ          INT                   not null,
    CNPJCPF           nvarchar(14)          null,
    NOME_CLI		  nvarchar(30)          null,
    NUMR_DOC          INT                   null,
    DESC_PROD		  nvarchar(80)          null,
    PERC_COMIS		  FLOAT					null,
    VALOR_FATURADO	  FLOAT					null,
    constraint PK_COMISSAOITEM primary key (COMISSAO_ID, NUMR_REQ, CODG_VEND, CODG_PROD)
)
go

alter table COMISSAOITEM
    add constraint FK_COMISSAO_REF_8_COMISSAO foreign key  (COMISSAO_ID)
       references COMISSAO (COMISSAO_ID)
go

