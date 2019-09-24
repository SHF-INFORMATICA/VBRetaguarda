VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLeProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'RECEBE
Private CODG_PROD_ENVIADO_A As Long

'ENVIA
Private Descricao_Produto_A As String

Public Property Let RecebeDadoID(valor As Long)
   CODG_PROD_ENVIADO_A = valor
End Property
========================
Public Property Set Descricao_Produto_A(valor As Long)
   Descricao_Produto_A = valor
End Property


Public Property Get EstoqueMedio() As Currency
    EstoqueMedio = curEstoqueMedio
End Property
Public Property Let EstoqueMedio(ByVal valor As Currency)
    curEstoqueMedio = valor
End Property

=============================
Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   Dim TabProduto As New ADODB.Recordset
   CRITERIO_A = ""
   CODG_PROD_ENVIADO_A = "" & Trim(CODG_PROD_ENVIADO_A)

'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select count(produto_id) from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODG_PROD_ENVIADO_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO_A = Trim(CODIGO_PRO_BARRA_A)

            frmPEDIDOBARRAS.Show 1
            If Trim(CRITERIO_A) <> "" Then
               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
               SQL = SQL & " where codg_produto = '" & Trim(CRITERIO_A) & "'"
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then _
                  CODG_PROD_ENVIADO_A = "" & TabProduto.Fields(0).Value
            End If
         End If
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "SELECT PRODUTO.PRODUTO_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.REFERENCIA, PRODUTO.FAMILIAPRODUTO_ID,"
   SQL = SQL & " PRODUTO.UNIDADE_MEDIDA, PRODUTO.SITUACAO, PRODUTO.SITUACAO_TRIBUTARIA, PRODUTO.ALIQUOTA_ICMS, PRODUTO.TIPO_PROD,"
   SQL = SQL & " PRODUTO.CODG_NCM, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, PRODUTO.LOCACAO, PRODUTO.QTD_MINIMO,"
   SQL = SQL & " PRODUTO.QTD_MAXIMO, PRODUTO.DT_ULT_VENDA, PRODUTO.DT_ULT_COMPRA, PRODUTO.PESO_LIQUIDO, PRODUTO.PESO_BRUTO, PRODUTO.MARCA_ID,"
   SQL = SQL & " PRODUTO.PRODUTO_BALANCA, PRODUTO.CONCEDER_PRODUCAO, PRODUTO.PERC_COMPOE_VENDA, FAMILIAPRODUTO.CODG_FAMILIA,"
   SQL = SQL & " FAMILIAPRODUTO.DESCRICAO AS DescricaoFamilia, FAMILIAPRODUTO.PRODUCAO,"
   SQL = SQL & " FAMILIAPRODUTO.PERC_COMPOE_VENDA AS PERC_COMPOE_VENDA_Familia, FAMILIAPRODUTO.UNIDADE_MEDIDA AS UN_FAMILIA,"
   SQL = SQL & " FAMILIAPRODUTO.DESC_UNIDADE_MEDIDA, PRODUTOFORNECEDOR.FORNECEDOR_ID, PRODUTOFORNECEDOR.CODG_PROD_FORNEC,"
   SQL = SQL & " PRODUTOFORNECEDOR.PRECO_CUSTO AS CUSTO_FORNECEDOR, PRODUTOFORNECEDOR.CODG_BARRA"
   SQL = SQL & " FROM PRODUTO "
   SQL = SQL & " LEFT OUTER JOIN FAMILIAPRODUTO "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " LEFT OUTER JOIN PRODUTOFORNECEDOR "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOFORNECEDOR.PRODUTO_ID"

   SQL = SQL & " where produto_id = " & Trim(CODG_PROD_ENVIADO_A)

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText


   'INDR_PRODUTO_PRODUCAO = False
   'INDR_PROD_BALANCA = False
   'PRODUTO_ID_N = TabProduto.Fields("produto_id").Value
   'txtProduto.Text = Trim(TabProduto.Fields("codg_produto").Value)
   'txtDescricao.Text = Trim(TabProduto.Fields("descricao").Value)
   'STATUS_PROD = Trim(TabProduto.Fields("SITUACAO").Value)
   'txtPesoItem.Text = Format(TabProduto.Fields("peso_liquido").Value, strFormatacao3Digitos)
   'txtAtacado.Text = Format(TabProduto.Fields("PRECO_ATACADO").Value, strFormatacao2Digitos)
   'txtVarejo.Text = Format(TabProduto.Fields("PRECO_VENDA").Value, strFormatacao2Digitos)
   'txtVlrUnit.Text = "" & Format(TabProduto.Fields("PRECO_VENDA").Value, strFormatacao2Digitos)
   'Valr_Venda_Produto_n = 0 & TabProduto.Fields("PRECO_VENDA").Value
   'txtValor_Unitario.Text = Format(Valr_Venda_Produto_n, strFormatacao2Digitos)
   'txtPreçoCusto.Text = "" & Format(TabProduto.Fields("preco_custo").Value, strFormatacao2Digitos)
   'VLR_ANTERIOR_N = TabProduto.Fields("PRECO_VENDA").Value
   'PRECO_PROD = 0 & txtAtacado.Text
   'CODG_PRODUTO_A = Trim(txtProduto.Text)
   'TabProduto.Fields("codg_ncm").Value)
   'CODIGO_BARRAS = ""

End Sub
