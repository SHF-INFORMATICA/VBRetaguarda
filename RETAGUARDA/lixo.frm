VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub INTEGRA_PEDIDO(NF_ID_N As Long, _
                   MFATRANSP As Long, _
                   MFAESPECIE As String, _
                   MFAPREFIXO As String, _
                   MFANFCUPOM As String, _
                   MFAOBSNOTA As String, _
                   MFAINDFINAL As String, _
                   MFAIDDEST As String, _
                   MFAINDPRES As String, _
                   MFACHAVEREFNFE As String, _
                   MFAFINNFE As String, _
                   MFAPLIQUI As String, _
                   MFAPBRUTO As String, _
                   MFATIFRETE As String, _
                   MFAEMAILENVIADO As String, _
                   MFACODSITT As String, _
                   MFATIPOREM As String, _
                   MFAVALIMP5 As Double, _
                   MFAVALIMP6 As Double, _
                   MFAVOLUME1 As Double)
'On Error GoTo ERRO_TRATA

   If NF_ID_N <= 0 Then
      MsgBox "Documento fiscal não encontrado !!!"
      Exit Sub
   End If
   If Trim(MFAINDFINAL) = "" Then
      MsgBox "Indica operação com Consumidor final da NF-e não informado !!!"
      Exit Sub
   End If
   If Trim(MFAIDDEST) = "" Then
      MsgBox "          da NF-e não informado !!!"
      Exit Sub
   End If
   If Trim(MFAINDPRES) = "" Then
      MsgBox "Indicador de presença do comprador da NF-e não informado !!!"
      Exit Sub
   End If
   'MFACHAVEREFNFE = ""
   If Trim(MFAFINNFE) = "" Then
      MsgBox "Finalidade de emissão da NF-e não informado !!!"
      Exit Sub
   End If

   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   'ABRE_BANCO_GLOBAL

   'If CONECTA_GLOBAL.State <> 1 Then
      'MsgBox "Banco GLOBAL não conectado."
   '   Exit Sub
   'End If
   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabPedidoIntegra As New ADODB.Recordset
   Dim TabCabecaIntegra As New ADODB.Recordset
   Dim strSQL           As String
   Dim MFADTREAJ        As String
   Dim MFAREAJUST       As String
   Dim MFAORDPAGO       As String
   Dim A1_CODMUN        As String
   Dim MFAFRETE         As Double
   Dim MFISITRIB        As String
   Dim ST_PRODUTO       As String
   Dim MFICLASFIS       As String
   Dim MFASEQUENCIA     As Long
   Dim MFANFECNF        As String
   Dim MFICF            As String
   Dim MFANOMECONSUMIDOR As String
   Dim MFACPFCONSUMIDOR As String
   Dim MFIITEM          As Long
   Dim MFAFILIAL        As String
   Dim MFASEGURO        As Double
   Dim MFAICMFRET       As Double
   Dim MFAVALBRUT       As Double
   Dim INSCRICAO_UF_A   As String

   MFIITEM = 0
   NUMR_NOTA_N = ""
   MFICF = ""
   MFACODSITT = Left(MFACODSITT, 59)

   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close

   strSQL = "select NF.NF_ID, NF.PESSOA_ID, NF.TRANSP_ID, NF.PEDIDO_ID, NF.NF_TIPO, NF.NUMR_NOTA, NF.SERIE_NOTA, NF.DT_EMISSAO, NF.DT_ENTRASAI, NF.STATUS, NF.DT_CANCELA, NF.QTD_VOLUME, "
   strSQL = strSQL & " NF.TIPO_ESPECIE, NF.PESO_BRUTO, NF.PESO_LIQUIDO, NF.NUMR_REQ_DEV, NF.ESTABELECIMENTO_ID, NF.indPres, NF.idDest, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID,NF.MODELO_DOC,"
   strSQL = strSQL & " NFITEM.VALOR, NFITEM.DESCONTO, NFITEM.QTDE, NFITEM.CFOP_ID, NFITEM.STRIBUTARIA, NFITEM.VLRBASEICMS, NFITEM.PERCICMS, NFITEM.VLRICMS, NFITEM.VLRBASEICMSSUBST,"
   strSQL = strSQL & " NFITEM.PERCICMSSUBST, NFITEM.VLRICMSSUBST, NFITEM.PERCREDUCAOICMS, NFITEM.PERCIVA, NFITEM.PERC_IPI, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.REFERENCIA,"
   strSQL = strSQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.UNIDADE_MEDIDA, PRODUTO.CODG_BARRA, PRODUTO.SITUACAO, PRODUTO.SITUACAO_TRIBUTARIA, PRODUTO.ALIQUOTA_ICMS, PRODUTO.PERC_DESCONTO,"
   strSQL = strSQL & " PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, PRODUTO.COMP_TRIBUTARIA, PRODUTO.FORNECEDOR_ID, PRODUTO.PRECO_CUSTO_ANTERIOR, PRODUTO.qtd_ped_anterior, PRODUTO.PRECO_CUSTO,"
   strSQL = strSQL & " PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, PRODUTO.PERCIVA AS ProdPercIVA, PRODUTO.DT_CADASTRO, PRODUTO.PERC_COMIS, PRODUTO.PATH_IMAGEM, PRODUTO.ORIGEM_MERCADO,"
   strSQL = strSQL & " PRODUTO.LOCACAO, PRODUTO.PRECO_VAREJO_ANTERIOR, PRODUTO.PRECO_ATACADO_ANTERIOR, PRODUTO.EMBALAGEM, PRODUTO.USUARIO_ID, PRODUTO.QTD_MINIMO, PRODUTO.QTD_MAXIMO,"
   strSQL = strSQL & " PRODUTO.DT_ULT_VENDA, PRODUTO.DT_ULT_COMPRA, PRODUTO.PESO_LIQUIDO AS PesoLiquiProduto, PRODUTO.PESO_BRUTO AS PesoBrutoProduto, PRODUTO.TAMANHO, PRODUTO.MARCA_ID, PRODUTO.PRODUTO_BALANCA,"
   strSQL = strSQL & " PRODUTO.PERMITE_DESCONTO, PRODUTO.CONCEDER_PRODUCAO, PRODUTO.PERC_COMPOE_VENDA, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFant, PESSOA.RAZAO, PESSOA.DATA_CAD,"
   strSQL = strSQL & " PESSOA.SITUACAO AS SitPessoa"
   strSQL = strSQL & " from NF WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN NFITEM WITH (NOLOCK)"
   strSQL = strSQL & " ON NF.NF_ID = NFITEM.NF_ID "
   strSQL = strSQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   strSQL = strSQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   strSQL = strSQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   strSQL = strSQL & " ON NF.PESSOA_ID = PESSOA.PESSOA_ID"

   strSQL = strSQL & " WHERE NF.nf_id = " & NF_ID_N

   TabPedidoIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoIntegra.EOF
      LIMPA_VARIAVEIS
      strSQL = ""
      PESSOA_ID_N = 0 & TabPedidoIntegra.Fields("pessoa_id").Value

      If Trim(NUMR_NOTA_N) <> Trim(TabPedidoIntegra.Fields("numr_nota").Value) Then
         A1_FILIAL = "0" & EMPRESA_ID_N
         NUMR_NOTA_N = "" & TabPedidoIntegra.Fields("numr_nota").Value
         MFASERIE = "1"
         MFACLIENTE = "" & TRAZ_ID_TABELA_GLOBAL("SA1010", "A1_COD", "A1_CGC", TabPedidoIntegra.Fields("CNPJCPF").Value)
         A1_LOJA = "0" & ESTABELECIMENTO_ID_N
         MFAFILIAL = "0" & ESTABELECIMENTO_ID_N

         MFANOMECONSUMIDOR = ""
         MFACPFCONSUMIDOR = ""
         'If Len(TabPedidoIntegra.Fields("CNPJCPF").Value) = 11 And _
            Trim(TabPedidoIntegra.Fields("CNPJCPF").Value) <> "99999999999" Then

         If Trim(TabPedidoIntegra.Fields("CNPJCPF").Value) <> "99999999999" Then
            MFACPFCONSUMIDOR = "" & TabPedidoIntegra.Fields("CNPJCPF").Value
            MFANOMECONSUMIDOR = "" & TabPedidoIntegra.Fields("NomeFant").Value
         End If

         MFASEQUENCIA = 1
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         strSQL = "select max(MFASEQUENCIA) from MFA010 WITH (NOLOCK)"
         TabCabecaIntegra.Open strSQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabCabecaIntegra.EOF Then _
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
               MFASEQUENCIA = TabCabecaIntegra.Fields(0).Value + 1
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

         MFACOND = "001"
         MFADUPL = ""
         MFAEMISSAO = "" & TabPedidoIntegra.Fields("DT_EMISSAO").Value
         MFAEST = "52"
         MFAFRETE = 0
         MFASEGURO = 0
         MFAICMFRET = 0
   
         If Len(Trim(TabPedidoIntegra.Fields("CNPJCPF").Value)) = 11 Then
            MFATIPOCLI = "F"
            Else: MFATIPOCLI = "J"
         End If
   
         MFAVALBRUT = 0
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         strSQL = "select sum((valor*qtde)-desconto) from NFITEM WITH (NOLOCK)"
         strSQL = strSQL & " where nf_id = " & NF_ID_N
         TabCabecaIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabecaIntegra.EOF Then _
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
               MFAVALBRUT = 0 & TabCabecaIntegra.Fields(0).Value
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
   
         MFAVALICM = ""
         MFABASEICM = ""
         MFAVALIPI = ""
         MFABASEIPI = ""
         MFAVALMERC = "" & MFAVALBRUT
         MFANFORI = ""
         MFADESCONT = "0"
         MFASERIORI = ""
         MFATIPO = "N"
         MFAESPECI2 = "PROPRIO"
         MFAESPECI3 = ""
         MFAESPECI4 = ""
         MFAVOLUME2 = ""
         MFAVOLUME3 = ""
         MFAVOLUME4 = ""
         MFAICMSRET = "0"
   
         If Trim(MFATRANSP) = "" Then _
            MFATRANSP = "1"
   
         MFAREDESP = ""
         MFAVEND1 = ""
         MFAVEND2 = ""
         MFAVEND3 = ""
         MFAVEND4 = ""
         MFAVEND5 = ""
         MFAOK = ""
         MFAFIMP = "0"
         MFADTREAJ = ""
         MFAREAJUST = ""
         MFAFATORB0 = "0"
         MFAFATORB1 = "0"
         MFAVARIAC = "0"
         MFABASEISS = "0"
         MFAVALISS = "0"
         MFAVALFAT = "" & MFAVALBRUT
         MFACONTSOC = "0"
         MFABRICMS = "0"
         MFAFRETAUT = "0"
         MFAICMAUTO = "0"
         MFADESPESA = "0"
         MFANEXTDOC = ""

         If Trim(MFAESPECIE) = "" Then _
            MFAESPECIE = "UN"

         MFAESPECI1 = "" & MFAESPECIE
         MFAESPECIE = ""
         MFAPDV = ""
         MFAMAPA = ""
         MFAECF = ""
   
         If Trim(MFAPREFIXO) = "" Then
            MFAPREFIXO = "NFE"
            Else
               If Trim(MFAPREFIXO) = "55" Then _
                  MFAPREFIXO = "NFE"
               If Trim(MFAPREFIXO) = "65" Then
                  MFAPREFIXO = "NFC"
                  'Tributos Totais Incidentes(Lei Federal 12.741/2012) :  R$ xxxx,xx
               End If
         End If
   
         MFABASIMP1 = "0"
         MFABASIMP2 = "0"
         MFABASIMP3 = "0"
         MFABASIMP4 = "0"
         MFABASIMP5 = "0"
         MFABASIMP6 = "0"
         MFAVALIMP1 = "0"
         MFAVALIMP2 = "0"
         MFAVALIMP3 = "0"
         MFAVALIMP4 = "0"
         MFAORDPAGO = ""
         MFAVALINSS = "0"
         MFAHORA = ""
         MFAMOEDA = "1"
         MFAREGIAO = ""
         MFAVALCSLL = "0"
         MFAVALCOFI = "0"
         MFAVALPIS = "0"
         MFALOTE = ""
         MFATXMOEDA = "0"
         MFAVALIRRF = "0"
         MFACARGA = ""
         MFASEQCAR = ""
         MFABASEINS = "0"
         MFANEXTSER = ""
         MFAPEDPEND = ""
         MFADESCCAB = "0"
         MFAFORMUL = ""
         MFATIPODOC = ""
         MFANFEACRS = ""
         MFASEQENT = ""
         MFADELETE = ""
   
         MFAREGISTRO = "1"
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         strSQL = "select max(MFAREGISTRO) from MFA010 WITH (NOLOCK)"
         TabCabecaIntegra.Open strSQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabCabecaIntegra.EOF Then _
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
               MFAREGISTRO = TabCabecaIntegra.Fields(0).Value + 1
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
   
         MFANFESERVI = ""
         MFANFEHRSE = ""
         MFANFECVS = ""
         MFACODPROT = ""
         MFACODSTAT = "0"
         MFACODMORE = "0"
         MFACHAVENFE = ""
         MFAMOTRESU = ""
         MFACODRECI = ""
         MFALOTENFE = ""
         MFAPLACA = ""
         MFAUFPLACA = ""
         MFAINDPAG = "0"
         MFAVALTOT = "" & MFAVALBRUT
         MFABASICMST = "0"
         MFAVALICMST = "0"
         MFAVALLIQUI = "" & MFAVALBRUT
         MFANFECNF = "" & TabPedidoIntegra.Fields("PEDIDO_ID").Value

         'CABEÇA DO PEDIDO
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

'   MODELO_DOCUMENTO = "" & MFAPREFIXO  'NF-e modelo 55 ou NF-e modelo 65
   
         strSQL = "select mfadoc from MFA010 WITH (NOLOCK)"
         strSQL = strSQL & " where mfadoc = '" & Trim(NUMR_NOTA_N) & "'"   'considera que a sequencia de nota fiscal é unica, por isso le dessa forma
         strSQL = strSQL & " and MFAPREFIXO = '" & Trim(MFAPREFIXO) & "'"
         TabCabecaIntegra.Open strSQL, CONECTA_GLOBAL, , , adCmdText
         If TabCabecaIntegra.EOF Then
            strSQL = "insert into MFA010 "
               strSQL = strSQL & " ("
               strSQL = strSQL & " MFACODEMP"
               strSQL = strSQL & ",MFADOC"
               strSQL = strSQL & ",MFASERIE"
               strSQL = strSQL & ",MFACLIENTE"
               strSQL = strSQL & ",MFALOJA"
               strSQL = strSQL & ",MFASEQUENCIA"
               strSQL = strSQL & ",MFACOND"
               strSQL = strSQL & ",MFADUPL"
               strSQL = strSQL & ",MFAEMISSAO"
               strSQL = strSQL & ",MFAEST"
               strSQL = strSQL & ",MFAFRETE"
               strSQL = strSQL & ",MFASEGURO"
               strSQL = strSQL & ",MFAICMFRET"
               strSQL = strSQL & ",MFATIPOCLI"
               strSQL = strSQL & ",MFAVALBRUT"
               strSQL = strSQL & ",MFAVALICM"
               strSQL = strSQL & ",MFABASEICM"
               strSQL = strSQL & ",MFAVALIPI"
               strSQL = strSQL & ",MFABASEIPI"
               strSQL = strSQL & ",MFAVALMERC"
               strSQL = strSQL & ",MFANFORI"
               strSQL = strSQL & ",MFADESCONT"
               strSQL = strSQL & ",MFASERIORI"
               strSQL = strSQL & ",MFATIPO"
               strSQL = strSQL & ",MFAESPECI1"
               strSQL = strSQL & ",MFAESPECI2"
               strSQL = strSQL & ",MFAESPECI3"
               strSQL = strSQL & ",MFAESPECI4"
               strSQL = strSQL & ",MFAVOLUME1"
               strSQL = strSQL & ",MFAVOLUME2"
               strSQL = strSQL & ",MFAVOLUME3"
               strSQL = strSQL & ",MFAVOLUME4"
               strSQL = strSQL & ",MFAICMSRET"
               strSQL = strSQL & ",MFAPLIQUI"
               strSQL = strSQL & ",MFAPBRUTO"
               strSQL = strSQL & ",MFATRANSP"
               strSQL = strSQL & ",MFAREDESP"
               strSQL = strSQL & ",MFAVEND1"
               strSQL = strSQL & ",MFAVEND2"
               strSQL = strSQL & ",MFAVEND3"
               strSQL = strSQL & ",MFAVEND4"
               strSQL = strSQL & ",MFAVEND5"
               strSQL = strSQL & ",MFAOK"
               strSQL = strSQL & ",MFAFIMP"
               strSQL = strSQL & ",MFADTLANC"
               strSQL = strSQL & ",MFADTREAJ"
               strSQL = strSQL & ",MFAREAJUST"
               strSQL = strSQL & ",MFADTBASE0"
               strSQL = strSQL & ",MFAFATORB0"
               strSQL = strSQL & ",MFADTBASE1"
               strSQL = strSQL & ",MFAFATORB1"
               strSQL = strSQL & ",MFAVARIAC"
               strSQL = strSQL & ",MFAFILIAL"
               strSQL = strSQL & ",MFABASEISS"
               strSQL = strSQL & ",MFAVALISS"
               strSQL = strSQL & ",MFAVALFAT"
               strSQL = strSQL & ",MFACONTSOC"
               strSQL = strSQL & ",MFABRICMS"
               strSQL = strSQL & ",MFAFRETAUT"
               strSQL = strSQL & ",MFAICMAUTO"
               strSQL = strSQL & ",MFADESPESA"
               strSQL = strSQL & ",MFANEXTDOC"
               strSQL = strSQL & ",MFAESPECIE"
               strSQL = strSQL & ",MFAPDV"
               strSQL = strSQL & ",MFAMAPA"
               strSQL = strSQL & ",MFAECF"
               strSQL = strSQL & ",MFAPREFIXO"
               strSQL = strSQL & ",MFABASIMP1"
               strSQL = strSQL & ",MFABASIMP2"
               strSQL = strSQL & ",MFABASIMP3"
               strSQL = strSQL & ",MFABASIMP4"
               strSQL = strSQL & ",MFABASIMP5"
               strSQL = strSQL & ",MFABASIMP6"
               strSQL = strSQL & ",MFAVALIMP1"
               strSQL = strSQL & ",MFAVALIMP2"
               strSQL = strSQL & ",MFAVALIMP3"
               strSQL = strSQL & ",MFAVALIMP4"
               strSQL = strSQL & ",MFAVALIMP5"
               strSQL = strSQL & ",MFAVALIMP6"
               strSQL = strSQL & ",MFAORDPAGO"
               strSQL = strSQL & ",MFANFCUPOM"
               strSQL = strSQL & ",MFAVALINSS"
               strSQL = strSQL & ",MFAHORA"
               strSQL = strSQL & ",MFAMOEDA"
               strSQL = strSQL & ",MFAREGIAO"
               strSQL = strSQL & ",MFAVALCSLL"
               strSQL = strSQL & ",MFAVALCOFI"
               strSQL = strSQL & ",MFAVALPIS"
               strSQL = strSQL & ",MFALOTE"
               strSQL = strSQL & ",MFATXMOEDA"
               strSQL = strSQL & ",MFAVALIRRF"
               strSQL = strSQL & ",MFACARGA"
               strSQL = strSQL & ",MFASEQCAR"
               strSQL = strSQL & ",MFABASEINS"
               strSQL = strSQL & ",MFANEXTSER"
               strSQL = strSQL & ",MFAPEDPEND"
               strSQL = strSQL & ",MFADESCCAB"
               strSQL = strSQL & ",MFADTENTR"
               strSQL = strSQL & ",MFAFORMUL"
               strSQL = strSQL & ",MFATIPODOC"
               strSQL = strSQL & ",MFANFEACRS"
               strSQL = strSQL & ",MFATIPOREM"
               strSQL = strSQL & ",MFASEQENT"
               strSQL = strSQL & ",MFADELETE"
               strSQL = strSQL & ",MFAREGISTRO"
               strSQL = strSQL & ",MFANFESERVI"
               strSQL = strSQL & ",MFANFEEMISE"
               strSQL = strSQL & ",MFANFEHRSE"
               strSQL = strSQL & ",MFANFECVS"
               strSQL = strSQL & ",MFACODPROT"
               strSQL = strSQL & ",MFACODSTAT"
               strSQL = strSQL & ",MFACODMORE"
               strSQL = strSQL & ",MFACHAVENFE"
               strSQL = strSQL & ",MFAMOTRESU"
               strSQL = strSQL & ",MFACODRECI"
               strSQL = strSQL & ",MFALOTENFE"
               strSQL = strSQL & ",MFAPLACA"
               strSQL = strSQL & ",MFAUFPLACA"
               strSQL = strSQL & ",MFACODSITT"
               strSQL = strSQL & ",MFAINDPAG"
               strSQL = strSQL & ",MFADTENSAI"
               strSQL = strSQL & ",MFATIFRETE"
               strSQL = strSQL & ",MFADTDIGIT"
               strSQL = strSQL & ",MFAVALTOT"
               strSQL = strSQL & ",MFABASICMST"
               strSQL = strSQL & ",MFAVALICMST"
               strSQL = strSQL & ",MFAVALLIQUI"
               strSQL = strSQL & ",MFAOBSNOTA"
               strSQL = strSQL & ",MFANFECNF"
               strSQL = strSQL & ",MFAEMAILENVIADO"
               strSQL = strSQL & ",MFAINDFINAL"
               strSQL = strSQL & ",MFAIDDEST"
               strSQL = strSQL & ",MFAINDPRES"
               strSQL = strSQL & ",MFACHAVEREFNFE"
               strSQL = strSQL & ",MFAFINNFE"
               strSQL = strSQL & ",MFANOMECONSUMIDOR"
               strSQL = strSQL & ",MFACPFCONSUMIDOR"
               strSQL = strSQL & " )"
            strSQL = strSQL & " VALUES ("
               strSQL = strSQL & "'" & A1_FILIAL & "'"
               strSQL = strSQL & ",'" & NUMR_NOTA_N & "'"
               strSQL = strSQL & ",'" & MFASERIE & "'"
               strSQL = strSQL & ",'" & MFACLIENTE & "'"
               strSQL = strSQL & ",'" & A1_LOJA & "'"
               strSQL = strSQL & "," & MFASEQUENCIA
               strSQL = strSQL & ",'" & MFACOND & "'"
               strSQL = strSQL & ",'" & MFADUPL & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFAEST & "'"
               strSQL = strSQL & "," & tpMOEDA(MFAFRETE)
               strSQL = strSQL & "," & tpMOEDA(MFASEGURO)
               strSQL = strSQL & "," & tpMOEDA(MFAICMFRET)
               strSQL = strSQL & ",'" & MFATIPOCLI & "'"
               strSQL = strSQL & "," & tpMOEDA(MFAVALBRUT)
               strSQL = strSQL & ",'" & MFAVALICM & "'"
               strSQL = strSQL & ",'" & MFABASEICM & "'"
               strSQL = strSQL & ",'" & MFAVALIPI & "'"
               strSQL = strSQL & ",'" & MFABASEIPI & "'"
               strSQL = strSQL & ",'" & tpMOEDA(MFAVALMERC) & "'"
               strSQL = strSQL & ",'" & MFANFORI & "'"
               strSQL = strSQL & ",'" & MFADESCONT & "'"
               strSQL = strSQL & ",'" & MFASERIORI & "'"
               strSQL = strSQL & ",'" & MFATIPO & "'"
               strSQL = strSQL & ",'" & MFAESPECI1 & "'"
               strSQL = strSQL & ",'" & MFAESPECI2 & "'"
               strSQL = strSQL & ",'" & MFAESPECI3 & "'"
               strSQL = strSQL & ",'" & MFAESPECI4 & "'"
               strSQL = strSQL & "," & tpMOEDA(MFAVOLUME1)
               'strSQL = strSQL & ",'" & Replace(MFAVOLUME1, ",", ".") & "'"
               strSQL = strSQL & ",'" & MFAVOLUME2 & "'"
               strSQL = strSQL & ",'" & MFAVOLUME3 & "'"
               strSQL = strSQL & ",'" & MFAVOLUME4 & "'"
               strSQL = strSQL & ",'" & MFAICMSRET & "'"
               strSQL = strSQL & ",'" & MFAPLIQUI & "'"
               strSQL = strSQL & ",'" & MFAPBRUTO & "'"
               strSQL = strSQL & ",'" & MFATRANSP & "'"
               strSQL = strSQL & ",'" & MFAREDESP & "'"
               strSQL = strSQL & ",'" & MFAVEND1 & "'"
               strSQL = strSQL & ",'" & MFAVEND2 & "'"
               strSQL = strSQL & ",'" & MFAVEND3 & "'"
               strSQL = strSQL & ",'" & MFAVEND4 & "'"
               strSQL = strSQL & ",'" & MFAVEND5 & "'"
               strSQL = strSQL & ",'" & MFAOK & "'"
               strSQL = strSQL & ",'" & MFAFIMP & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFADTREAJ & "'"
               strSQL = strSQL & ",'" & MFAREAJUST & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFAFATORB0 & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFAFATORB1 & "'"
               strSQL = strSQL & ",'" & MFAVARIAC & "'"

               strSQL = strSQL & ",'" & MFAFILIAL & "'"

               strSQL = strSQL & ",'" & MFABASEISS & "'"
               strSQL = strSQL & ",'" & MFAVALISS & "'"
               strSQL = strSQL & ",'" & tpMOEDA(MFAVALFAT) & "'"
               strSQL = strSQL & ",'" & MFACONTSOC & "'"
               strSQL = strSQL & ",'" & MFABRICMS & "'"
               strSQL = strSQL & ",'" & MFAFRETAUT & "'"
               strSQL = strSQL & ",'" & MFAICMAUTO & "'"
               strSQL = strSQL & ",'" & MFADESPESA & "'"
               strSQL = strSQL & ",'" & MFANEXTDOC & "'"
               strSQL = strSQL & ",'" & MFAESPECIE & "'"
               strSQL = strSQL & ",'" & MFAPDV & "'"
               strSQL = strSQL & ",'" & MFAMAPA & "'"
               strSQL = strSQL & ",'" & MFAECF & "'"
               strSQL = strSQL & ",'" & MFAPREFIXO & "'"
               strSQL = strSQL & ",'" & MFABASIMP1 & "'"
               strSQL = strSQL & ",'" & MFABASIMP2 & "'"
               strSQL = strSQL & ",'" & MFABASIMP3 & "'"
               strSQL = strSQL & ",'" & MFABASIMP4 & "'"
               strSQL = strSQL & ",'" & MFABASIMP5 & "'"
               strSQL = strSQL & ",'" & MFABASIMP6 & "'"
               strSQL = strSQL & ",'" & MFAVALIMP1 & "'"
               strSQL = strSQL & ",'" & MFAVALIMP2 & "'"
               strSQL = strSQL & ",'" & MFAVALIMP3 & "'"
               strSQL = strSQL & ",'" & MFAVALIMP4 & "'"
               strSQL = strSQL & "," & tpMOEDA(MFAVALIMP5)
               strSQL = strSQL & "," & tpMOEDA(MFAVALIMP6)

               'strSQL = strSQL & ",'" & Replace(MFAVALIMP5, ",", ".") & "'"
               'strSQL = strSQL & ",'" & Replace(MFAVALIMP6, ",", ".") & "'"

               strSQL = strSQL & ",'" & MFAORDPAGO & "'"
               strSQL = strSQL & ",'" & MFANFCUPOM & "'"
               strSQL = strSQL & ",'" & MFAVALINSS & "'"
               strSQL = strSQL & ",'" & MFAHORA & "'"
               strSQL = strSQL & ",'" & MFAMOEDA & "'"
               strSQL = strSQL & ",'" & MFAREGIAO & "'"
               strSQL = strSQL & ",'" & MFAVALCSLL & "'"
               strSQL = strSQL & ",'" & MFAVALCOFI & "'"
               strSQL = strSQL & ",'" & MFAVALPIS & "'"
               strSQL = strSQL & ",'" & MFALOTE & "'"
               strSQL = strSQL & ",'" & MFATXMOEDA & "'"
               strSQL = strSQL & ",'" & MFAVALIRRF & "'"
               strSQL = strSQL & ",'" & MFACARGA & "'"
               strSQL = strSQL & ",'" & MFASEQCAR & "'"
               strSQL = strSQL & ",'" & MFABASEINS & "'"
               strSQL = strSQL & ",'" & MFANEXTSER & "'"
               strSQL = strSQL & ",'" & MFAPEDPEND & "'"
               strSQL = strSQL & ",'" & MFADESCCAB & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFAFORMUL & "'"
               strSQL = strSQL & ",'" & MFATIPODOC & "'"
               strSQL = strSQL & ",'" & MFANFEACRS & "'"
               strSQL = strSQL & ",'" & MFATIPOREM & "'"
               strSQL = strSQL & ",'" & MFASEQENT & "'"
               strSQL = strSQL & ",'" & MFADELETE & "'"
               strSQL = strSQL & ",'" & MFAREGISTRO & "'"
               strSQL = strSQL & ",'" & MFANFESERVI & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFANFEHRSE & "'"
               strSQL = strSQL & ",'" & MFANFECVS & "'"
               strSQL = strSQL & ",'" & MFACODPROT & "'"
               strSQL = strSQL & ",'" & MFACODSTAT & "'"
               strSQL = strSQL & ",'" & MFACODMORE & "'"
               strSQL = strSQL & ",'" & MFACHAVENFE & "'"
               strSQL = strSQL & ",'" & MFAMOTRESU & "'"
               strSQL = strSQL & ",'" & MFACODRECI & "'"
               strSQL = strSQL & ",'" & MFALOTENFE & "'"
               strSQL = strSQL & ",'" & MFAPLACA & "'"
               strSQL = strSQL & ",'" & MFAUFPLACA & "'"
               strSQL = strSQL & ",'" & MFACODSITT & "'"
               strSQL = strSQL & ",'" & MFAINDPAG & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & MFATIFRETE & "'"
               strSQL = strSQL & ",'" & MFAEMISSAO & "'"
               strSQL = strSQL & ",'" & tpMOEDA(MFAVALTOT) & "'"
               strSQL = strSQL & ",'" & MFABASICMST & "'"
               strSQL = strSQL & ",'" & MFAVALICMST & "'"
               strSQL = strSQL & ",'" & tpMOEDA(MFAVALLIQUI) & "'"
               strSQL = strSQL & ",'" & MFAOBSNOTA & "'"
               strSQL = strSQL & ",'" & MFANFECNF & "'"
               strSQL = strSQL & ",'" & MFAEMAILENVIADO & "'"
               strSQL = strSQL & ",'" & MFAINDFINAL & "'"
               strSQL = strSQL & ",'" & MFAIDDEST & "'"
               strSQL = strSQL & ",'" & MFAINDPRES & "'"
               strSQL = strSQL & ",'" & MFACHAVEREFNFE & "'"
               strSQL = strSQL & ",'" & MFAFINNFE & "'"
               strSQL = strSQL & ",'" & Trim(Left(MFANOMECONSUMIDOR, 100)) & "'"
               strSQL = strSQL & ",'" & Trim(MFACPFCONSUMIDOR) & "'"
            strSQL = strSQL & " )"

            CONECTA_GLOBAL.Execute strSQL
            Else  'update
         End If
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
      End If 'fim cabeça pedido

'========================FIM CABEÇA MFA010
'========================FIM CABEÇA MFA010
'========================FIM CABEÇA MFA010
''''''''''''''''ITENS PRODUTO
      MFIUM = "" & Trim(TabPedidoIntegra.Fields("unidade_medida").Value)
      MFIQUANT = "" & Trim(TabPedidoIntegra.Fields("qtde").Value)
      MFIPRCVEN = "" & Trim(TabPedidoIntegra.Fields("valor").Value)
      MFITOTAL = "" & (TabPedidoIntegra.Fields("valor").Value * TabPedidoIntegra.Fields("qtde").Value)
      MFIVALIPI = "0"
      MFIVALICM = "0"
      MFITES = "001"
      MFICF = "" & TabPedidoIntegra.Fields("cfop_id").Value
      MFIDESC = "0"
      MFIIPI = "0"
      MFIPICM = "0"
      MFIPESO = "0"
      MFICONTA = ""
      MFIOP = ""
      MFIITEMPV = ""
      MFILOCAL = A1_LOJA
      MFIDOC = "" & NUMR_NOTA_N
      MFIGRUPO = "0001"
      MFITP = "MP"
      MFISERIE = "0"
      MFICUSTO1 = "0"
      MFICUSTO2 = "0"
      MFICUSTO3 = "0"
      MFICUSTO4 = "0"
      MFICUSTO5 = "0"
      MFIPRUNIT = "0"
      MFIQTSEGUM = "0"
      MFIEST = "52"
      MFIDESCON = "0"
      MFITIPO = "N"
      MFIQTDEDEV = "" & Trim(TabPedidoIntegra.Fields("qtde").Value)
      MFIVALDEV = "" & (TabPedidoIntegra.Fields("valor").Value * TabPedidoIntegra.Fields("qtde").Value)
      MFIORIGLAN = "NF" '& MFAPREFIXO
      MFIBRICMS = "0"
      MFIBASEORI = "0"
      MFIBASEICM = "0"
      MFIVALACRS = "0"
      MFIIDENTB6 = ""
      MFICODISS = ""
      MFIGRADE = ""
      'MFISEQCALC VERIFICAR O QUE É PRA ARMAZENAR NESSE CAMPO
      MFIICMSRET = "0"
      MFICOMIS1 = "0"
      MFICOMIS2 = "0"
      MFICOMIS3 = "0"
      MFICOMIS4 = "0"
      MFICOMIS5 = "0"
      MFILOTECTL = ""
      MFINUMLOTE = ""
      MFIDTVALID = ""
      MFIDESCZFR = "0"
      MFIPDV = ""
      MFINUMSERI = ""
      MFIDTLCTCT = "" & Trim(TabPedidoIntegra.Fields("DT_EMISSAO").Value)
      MFICUSFF1 = "0"
      MFICUSFF2 = "0"
      MFICUSFF3 = "0"
      MFICUSFF4 = "0"
      MFICUSFF5 = "0"

'========================
'========================
'========================
'Segue uma pequena orientação sobre o uso do CFOP e do CSOSN.
'- Quando efetuar a revenda dentro do Estado sem substituição tributária usar.
'CFOP: 5.102 podendo usar o CSOSN 0101 com permissão de crédito ou 0102 sem permissão de crédito quando efetuar a venda para pessoa física ou pessoa jurídica.
'- Quando efetuar uma revenda para fora do estado sem substituição tributária usar.
'CFOP: 6102 - podendo usar o CSOSN 0101 com permissão de crédito ou 0102 sem permissão de crédito quando efetuar a venda para pessoa física ou pessoa jurídica.
'- Quando efetuar uma revenda para dentro do estado com substituição tributária usar.
'CFOP: 5405 e o CSOSN 0500.
'Quando efetuar uma revenda para fora do estado com substituição tributária usar.
'CFOP: 6404 e o CSOSN 0500.

'12 Tributação do ICMS pelo Simples Nacional sem permissao  102
'15 Tributação do ICMS pelo Simples Nacional(500)           500
'16 Tributação do ICMS pelo Simples Nacional(900)           900
'17 Tributação do ICMS pelo Simples Nacional Nao tributado  400
INSCRICAO_UF_A = "" & TRAZ_IE(PESSOA_ID_N)
MFICLASFIS = "17" 'NFE
If MFAPREFIXO = "NFC" Then
   MFICLASFIS = "12"
   MFISITRIB = "42"
End If

If MFAPREFIXO = "NFC" Then
   If Trim(MFICF) <> "5102" Then
      MFICF = "5102"
      MFISITRIB = "42"
   End If
End If

ST_PRODUTO = "" & Trim(TabPedidoIntegra.Fields("SITUACAO_TRIBUTARIA").Value)
'SE O ITEM FOR SUBSTITUIÇÃO TRIBUTARIA PASSA 500
'If ST_PRODUTO = "10" Or ST_PRODUTO = "60" Then _
   MFICLASFIS = "15"

'If INSCRICAO_UF_A = "" Or INSCRICAO_UF_A = "ISENTO" Then
'   If MFAPREFIXO = "NFC" Then
      'MFICLASFIS = "12"
'      Else: MFICLASFIS = "17" 'NFE
'   End If
'End If
'========================
'========================
'========================

      MFIBASIMP1 = "0"
      MFIBASIMP2 = "0"
      MFIBASIMP3 = "0"
      MFIBASIMP4 = "0"
      MFIBASIMP5 = "0"
      MFIBASIMP6 = "0"
      MFIVALIMP1 = "0"
      MFIVALIMP2 = "0"
      MFIVALIMP3 = "0"
      MFIVALIMP4 = "0"
      MFIVALIMP5 = "0"
      MFIVALIMP6 = "0"
      MFIITEMORI = ""
      MFICODFAB = ""
      MFILOJAFA = ""
      MFICCUSTO = ""
      MFIITEMCC = ""
      MFILOCALIZ = ""
      MFIENVCNAB = ""
      MFIALIQINS = "0"
      MFIPREEMB = ""
      MFIALIQISS = "0"
      MFIBASEIPI = "0"
      MFIBASEISS = "0"
      MFIVALISS = "0"
      MFISEGURO = "0"
      MFIVALFRE = "0"
      MFIDESPESA = "0"
      MFICLVL = ""
      MFIBASEINS = "0"
      MFIICMFRET = "0"
      MFISERVIC = ""
      MFISTSERV = ""
      MFIVALINS = "0"
      MFIPROJPMS = ""
      MFITASKPMS = ""
      MFILICITA = ""
      MFIREMITO = ""
      MFISERIREM = ""
      MFIITEMREM = ""
      MFIALQIMP1 = "0"
      MFIALQIMP2 = "0"
      MFIALQIMP3 = "0"
      MFIALQIMP4 = "0"
      MFIALQIMP5 = "0"
      MFIALQIMP6 = "0"
      MFITPDCENV = ""
      MFIOK = ""
      MFIENDER = ""
      MFIEDTPMS = ""
      MFIVARPRUN = "0"
      MFIFORMUL = ""
      MFITIPODOC = ""
      MFIVAC = "0"
      MFITIPOREM = ""

      MFIQTDEFAT = "" & Trim(TabPedidoIntegra.Fields("qtde").Value)
      MFIQTDAFAT = "" & Trim(TabPedidoIntegra.Fields("qtde").Value)

      MFIPOTENCI = "0"
      MFIDELETE = ""
      MFIDESTOTIT = "0"
      MFIALIICMS = "0"
      MFIBASICMST = "0"
      MFIALIICMST = "0"
      MFIVALICMST = "0"
      MFIALIICMRED = "0"
      MFIVALBRUT = "0"
      MFIVALBONI = "0"
      MFIVALTROCA = "0"
      MFIQTDVOL = "0"
      MFIPESLIQ = "0"
      MFIPESBRU = "0"

      MFIVALLIQ = "0" & (TabPedidoIntegra.Fields("valor").Value * TabPedidoIntegra.Fields("qtde").Value)

'SET VERIFICAR SE PRECISA DE LER ANTES DE INSERIR
'SET VERIFICAR SE PRECISA DE LER ANTES DE INSERIR
'SET VERIFICAR SE PRECISA DE LER ANTES DE INSERIR
'SET VERIFICAR SE PRECISA DE LER ANTES DE INSERIR
'SET VERIFICAR SE PRECISA DE LER ANTES DE INSERIR
'SET VERIFICAR SE PRECISA DE LER ANTES DE INSERIR

      strSQL = "insert into MFi010 "
      strSQL = strSQL & "("
         strSQL = strSQL & "MFIFILIAL,MFIITEM,MFICOD,MFIUM,MFISEGUM,MFIQUANT,MFIPRCVEN,MFITOTAL,MFIVALIPI,MFIVALICM"
         strSQL = strSQL & ",MFITES,MFICF,MFIDESC,MFIIPI,MFIPICM,MFIPESO,MFICONTA,MFIOP,MFIITEMPV,MFICLIENTE"
         strSQL = strSQL & ",MFILOJA,MFILOCAL,MFIDOC,MFIEMISSAO,MFIGRUPO,MFITP,MFISERIE,MFICUSTO1,MFICUSTO2,MFICUSTO3"
         strSQL = strSQL & ",MFICUSTO4,MFICUSTO5,MFIPRUNIT,MFIQTSEGUM,MFINUMSEQ,MFIEST,MFIDESCON,MFITIPO,MFINFORI,MFISERIORI"
         strSQL = strSQL & ",MFIQTDEDEV,MFIVALDEV,MFIORIGLAN,MFIBRICMS,MFIBASEORI,MFIBASEICM,MFIVALACRS,MFIIDENTB6,MFICODISS"
         strSQL = strSQL & ",MFIGRADE,MFISEQCALC,MFIICMSRET,MFICOMIS1,MFICOMIS2,MFICOMIS3,MFICOMIS4,MFICOMIS5,MFILOTECTL"
         strSQL = strSQL & ",MFINUMLOTE,MFIDTVALID,MFIDESCZFR,MFIPDV,MFINUMSERI,MFIDTLCTCT,MFICUSFF1,MFICUSFF2,MFICUSFF3"
         strSQL = strSQL & ",MFICUSFF4,MFICUSFF5,MFICLASFIS,MFIBASIMP1,MFIBASIMP2,MFIBASIMP3,MFIBASIMP4,MFIBASIMP5,MFIBASIMP6"
         strSQL = strSQL & ",MFIVALIMP1,MFIVALIMP2,MFIVALIMP3,MFIVALIMP4,MFIVALIMP5,MFIVALIMP6,MFIITEMORI,MFICODFAB,MFILOJAFA"
         strSQL = strSQL & ",MFICCUSTO,MFIITEMCC,MFILOCALIZ,MFIENVCNAB,MFIALIQINS,MFIPREEMB,MFIALIQISS,MFIBASEIPI,MFIBASEISS"
         strSQL = strSQL & ",MFIVALISS,MFISEGURO,MFIVALFRE,MFIDESPESA,MFICLVL,MFIBASEINS,MFIICMFRET,MFISERVIC,MFISTSERV"
         strSQL = strSQL & ",MFIVALINS,MFIPROJPMS,MFITASKPMS,MFILICITA,MFIREMITO,MFISERIREM,MFIITEMREM,MFIALQIMP1,MFIALQIMP2"
         strSQL = strSQL & ",MFIALQIMP3,MFIALQIMP4,MFIALQIMP5,MFIALQIMP6,MFITPDCENV,MFIOK,MFIENDER,MFIEDTPMS,MFIVARPRUN,MFIFORMUL"
         strSQL = strSQL & ",MFITIPODOC,MFIVAC,MFITIPOREM,MFIQTDEFAT,MFIQTDAFAT,MFISEQUEN,MFIPOTENCI,MFIDELETE,MFIREGISTRO"
         strSQL = strSQL & ",MFISITRIB,MFIDESTOTIT,MFIALIICMS,MFIBASICMST,MFIALIICMST,MFIVALICMST,MFIALIICMRED,MFIVALBRUT"
         strSQL = strSQL & ",MFIVALBONI,MFIVALTROCA,MFIQTDVOL,MFIPESLIQ,MFIPESBRU,MFIVALLIQ"
      strSQL = strSQL & ")"
      strSQL = strSQL & " values("
         strSQL = strSQL & "'" & A1_FILIAL & "'"

'=============================
MFISEQUEN = "" & MFASEQUENCIA
         MFIITEM = 1
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select max(MFIITEM) from MFI010 WITH (NOLOCK)"
         SQL = SQL & " where MFISEQUEN  = " & MFISEQUEN
         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabCabecaIntegra.EOF Then
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then
               MFIITEM = 0 & TabCabecaIntegra.Fields(0).Value
               MFIITEM = MFIITEM + 1
            End If
         End If
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

         'strSQL = strSQL & ",'" & MFIITEM & "'"
         strSQL = strSQL & "," & MFIITEM

         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select b1_cod from SB1010 WITH (NOLOCK)"
         SQL = SQL & " where b1_codant = '" & Trim(TabPedidoIntegra.Fields("codg_produto").Value) & "'"
         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabCabecaIntegra.EOF Then _
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
               MFICOD = TabCabecaIntegra.Fields(0).Value
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

         strSQL = strSQL & ",'" & MFICOD & "'"
         strSQL = strSQL & ",'" & MFIUM & "'"
         strSQL = strSQL & ",'" & MFISEGUM & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIQUANT) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIPRCVEN) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFITOTAL) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIPI) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALICM) & "'"
         strSQL = strSQL & ",'" & MFITES & "'"
         strSQL = strSQL & ",'" & MFICF & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIDESC) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIIPI) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIPICM) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIPESO) & "'"
         strSQL = strSQL & ",'" & MFICONTA & "'"
         strSQL = strSQL & ",'" & MFIOP & "'"
         strSQL = strSQL & ",'" & MFIITEMPV & "'"
         strSQL = strSQL & ",'" & MFACLIENTE & "'"
         strSQL = strSQL & ",'" & A1_LOJA & "'"
         strSQL = strSQL & ",'" & MFILOCAL & "'"
         strSQL = strSQL & ",'" & MFIDOC & "'"
         strSQL = strSQL & ",'" & MFAEMISSAO & "'"
         strSQL = strSQL & ",'" & MFIGRUPO & "'"
         strSQL = strSQL & ",'" & MFITP & "'"
         strSQL = strSQL & ",'" & MFISERIE & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSTO1) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSTO2) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSTO3) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSTO4) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSTO5) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIPRUNIT) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIQTSEGUM) & "'"

         MFINUMSEQ = ""
         If MFIITEM = 1 Then
            MFINUMSEQ = 1
            If TabCabecaIntegra.State = 1 Then _
               TabCabecaIntegra.Close
            SQL = "select max(MFINUMSEQ) from MFI010 WITH (NOLOCK)"
            TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
            If Not TabCabecaIntegra.EOF Then _
               If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
                  MFINUMSEQ = TabCabecaIntegra.Fields(0).Value + 1
            If TabCabecaIntegra.State = 1 Then _
               TabCabecaIntegra.Close
         End If

         strSQL = strSQL & ",'" & MFINUMSEQ & "'"
         strSQL = strSQL & ",'" & MFIEST & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIDESCON) & "'"
         strSQL = strSQL & ",'" & MFITIPO & "'"
         strSQL = strSQL & ",'" & MFINFORI & "'"
         strSQL = strSQL & ",'" & MFISERIORI & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIQTDEDEV) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALDEV) & "'"
         strSQL = strSQL & ",'" & MFIORIGLAN & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBRICMS) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASEORI) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASEICM) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALACRS) & "'"
         strSQL = strSQL & ",'" & MFIIDENTB6 & "'"
         strSQL = strSQL & ",'" & MFICODISS & "'"
         strSQL = strSQL & ",'" & MFIGRADE & "'"
         strSQL = strSQL & ",'" & MFISEQCALC & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIICMSRET) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICOMIS1) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICOMIS2) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICOMIS3) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICOMIS4) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICOMIS5) & "'"
         strSQL = strSQL & ",'" & MFILOTECTL & "'"
         strSQL = strSQL & ",'" & MFINUMLOTE & "'"
         strSQL = strSQL & ",'" & MFIDTVALID & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIDESCZFR) & "'"
         strSQL = strSQL & ",'" & MFIPDV & "'"
         strSQL = strSQL & ",'" & MFINUMSERI & "'"
         strSQL = strSQL & ",'" & MFIDTLCTCT & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSFF1) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSFF2) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSFF3) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSFF4) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFICUSFF5) & "'"
         strSQL = strSQL & ",'" & MFICLASFIS & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASIMP1) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASIMP2) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASIMP3) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASIMP4) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASIMP5) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASIMP6) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIMP1) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIMP2) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIMP3) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIMP4) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIMP5) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALIMP6) & "'"
         strSQL = strSQL & ",'" & MFIITEMORI & "'"
         strSQL = strSQL & ",'" & MFICODFAB & "'"
         strSQL = strSQL & ",'" & MFILOJAFA & "'"
         strSQL = strSQL & ",'" & MFICCUSTO & "'"
         strSQL = strSQL & ",'" & MFIITEMCC & "'"
         strSQL = strSQL & ",'" & MFILOCALIZ & "'"
         strSQL = strSQL & ",'" & MFIENVCNAB & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALIQINS) & "'"
         strSQL = strSQL & ",'" & MFIPREEMB & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALIQISS) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASEIPI) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASEISS) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALISS) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFISEGURO) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALFRE) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIDESPESA) & "'"
         strSQL = strSQL & ",'" & MFICLVL & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASEINS) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIICMFRET) & "'"
         strSQL = strSQL & ",'" & MFISERVIC & "'"
         strSQL = strSQL & ",'" & MFISTSERV & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALINS) & "'"
         strSQL = strSQL & ",'" & MFIPROJPMS & "'"
         strSQL = strSQL & ",'" & MFITASKPMS & "'"
         strSQL = strSQL & ",'" & MFILICITA & "'"
         strSQL = strSQL & ",'" & MFIREMITO & "'"
         strSQL = strSQL & ",'" & MFISERIREM & "'"
         strSQL = strSQL & ",'" & MFIITEMREM & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALQIMP1) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALQIMP2) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALQIMP3) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALQIMP4) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALQIMP5) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALQIMP6) & "'"
         strSQL = strSQL & ",'" & MFITPDCENV & "'"
         strSQL = strSQL & ",'" & MFIOK & "'"
         strSQL = strSQL & ",'" & MFIENDER & "'"
         strSQL = strSQL & ",'" & MFIEDTPMS & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVARPRUN) & "'"
         strSQL = strSQL & ",'" & MFIFORMUL & "'"
         strSQL = strSQL & ",'" & MFITIPODOC & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVAC) & "'"
         strSQL = strSQL & ",'" & MFITIPOREM & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIQTDEFAT) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIQTDAFAT) & "'"

         MFISEQUEN = "" & MFASEQUENCIA
         strSQL = strSQL & ",'" & MFISEQUEN & "'"

         strSQL = strSQL & ",'" & tpMOEDA(MFIPOTENCI) & "'"
         strSQL = strSQL & ",'" & MFIDELETE & "'"

         MFIREGISTRO = "1"
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select max(MFIREGISTRO) from MFI010 WITH (NOLOCK)"
         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabCabecaIntegra.EOF Then _
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
               MFIREGISTRO = TabCabecaIntegra.Fields(0).Value + 1
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         strSQL = strSQL & ",'" & MFIREGISTRO & "'"

'MTSCODFIS DQUI PEGA O CFOP DO ITEM NO GLOBAL
         MFISITRIB = "42"
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select MTSCODIGO from MTSITTRIBU WITH (NOLOCK)"
         'SQL = SQL & " where MTSCODFIS = '" & Trim(TabPedidoIntegra.Fields("cfop_id").Value) & "'"
         SQL = SQL & " where MTSCODFIS = '" & Trim(MFICF) & "'"

         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabCabecaIntegra.EOF Then _
            If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
               MFISITRIB = TabCabecaIntegra.Fields(0).Value
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

ST_PRODUTO = "" & Trim(TabPedidoIntegra.Fields("SITUACAO_TRIBUTARIA").Value)

         If ST_PRODUTO = "00" Then
            If INDR_INDUSTRIA = False Then
               
            End If
         End If

         strSQL = strSQL & ",'" & MFISITRIB & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIDESTOTIT) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALIICMS) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIBASICMST) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALIICMST) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALICMST) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIALIICMRED) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALBRUT) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALBONI) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALTROCA) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIQTDVOL) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIPESLIQ) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIPESBRU) & "'"
         strSQL = strSQL & ",'" & tpMOEDA(MFIVALLIQ) & "'"
      strSQL = strSQL & ")"

'MFISEQCALC VERIFICAR O QUE É PRA ARMAZENAR NESSE CAMPO
CONECTA_GLOBAL.Execute strSQL

      TabPedidoIntegra.MoveNext
   Wend
   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_PEDIDO"
End Sub


