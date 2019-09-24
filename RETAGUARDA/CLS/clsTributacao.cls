VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTributacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Public Sub PREPARA_TRIBUTACAO_PRODUTO(CNPJ_CPF_A As String, VLR_UNIT_N As Double, QTDE_PEDIDO_N As Double)
'On Error GoTo ERRO_TRATA

'UF DO CLIENTE
'IVA
'MVA (Margem de Valor Agregado)
'PEGAR ALIQUOTA PARA O ESTADO TAL

   If VLR_UNIT_N <= 0 Then
      MsgBox "Valor não informado."
      Exit Sub
   End If
   If QTDE_PEDIDO_N <= 0 Then
      MsgBox "Quantidade não informado."
      Exit Sub
   End If
   CNPJ_CPF_A = Trim(CNPJ_CPF_A)
   If Trim(CNPJ_CPF_A) = "" Then
      MsgBox "CNPJ/CPF não informado."
      Exit Sub
   End If

   Dim CFOP_ID_N As Integer

   UF_CLIENTE_A = ""

CLIENTE_ID_N = 0 & TRAZ_ID_TABELA("CLIENTE", "cliente_id", "cgccpf", CNPJ_CPF_A)

   If CLIENTE_ID_N < 0 Then
      MsgBox "Cliente não informado, verifique !!!"
      Exit Sub
   End If

   If Trim(CNPJ_CPF_A) <> "99999999999" And Trim(CNPJ_CPF_A) <> "" Then
      TRATA_PESSOA Trim(CNPJCPF_A)
      If Trim(UF_CLIENTE_A) = "" Then
         If INDR_PEDIDO_VENDA = False Then
            MsgBox "Cliente com cadastro incompleto !!! UF_CLIENTE_A = " & UF_CLIENTE_A
            Exit Sub
         End If
      End If
   End If

   If Trim(UF_EMPRESA_A) = "" Then _
      PEGA_DADOS_EMPRESA

   'aqui é ajustado se for consumidor final tem que pegar o mesmo UF de destrino para aliquotas
   If Trim(CNPJ_CPF_A) = "99999999999" Then _
      UF_CLIENTE_A = "" & UF_EMPRESA_A

   If Trim(UF_CLIENTE_A) = "" Then _
      UF_CLIENTE_A = "" & UF_EMPRESA_A

   Dim rstProduto          As New ADODB.Recordset
   Dim ST_PRODUTO_A        As String
   Dim PERCIVA_A           As String
   Dim COMP_TRIBUTARIA_A   As String

   ST_PRODUTO_A = ""

   If rstProduto.State = 1 Then _
      rstProduto.Close

   SQL = "select situacao_tributaria,perciva,comp_tributaria from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   rstProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstProduto.EOF Then
      If rstProduto.State = 1 Then _
         rstProduto.Close

      MsgBox "Rotina Tributação, produto não cadastrado." & vbCrLf & "Verique"
      Exit Sub
      Else
         ST_PRODUTO_A = "" & Trim(rstProduto.Fields("SITUACAO_TRIBUTARIA").Value)
         PERCIVA_A = "" & rstProduto!PERCIVA
         COMP_TRIBUTARIA_A = "" & rstProduto!COMP_TRIBUTARIA
   End If
   If rstProduto.State = 1 Then _
      rstProduto.Close

   Dim TabTempCFOP                  As New ADODB.Recordset
   Dim VALOR_BASE_ICMS_N            As Double
   Dim PERC_ICMS_N                  As Double
   Dim VALOR_BASE_ICMS_SUBST_N      As Double
   Dim VALOR_ICMS_PRODUTO_SUBST_N   As Double
   Dim VALOR_PERC_ICMS_SUBST_N      As Double
   Dim strCFOP_ITEM                 As String
   Dim PERC_REDUCAO_ICMS_N          As Double
   Dim PERC_IVA_N                   As Double
   Dim VALOR_TOTAL_ITEM_N           As Double
   Dim Aliquota_N                   As Double
   Dim VALOR_ICMS_N                 As Double

   VALOR_BASE_ICMS_N = 0
   VALOR_ICMS_N = 0
   PERC_ICMS_N = 0
   VALOR_BASE_ICMS_SUBST_N = 0
   VALOR_ICMS_PRODUTO_SUBST_N = 0
   VALOR_PERC_ICMS_SUBST_N = 0
   PERC_REDUCAO_ICMS_N = 0
   PERC_IVA_N = 0
   Aliquota_N = 0
   VALOR_TOTAL_ITEM_N = 0

   strCFOP_ITEM = ""
   strCFOP_ITEM = "5102"   'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
   'strCFOP_ITEM = "5101"   'CFOP 5101 - Venda de produção do estabelecimento

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

   VALOR_ITEM_N = 0 & VLR_UNIT_N
   QTDE_N = 0 & QTDE_PEDIDO_N
   VALOR_TOTAL_ITEM_N = (QTDE_N * VALOR_ITEM_N)

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

'BUSCANDO ALIQUOTAS
   If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
      strCFOP_ITEM = "5102"
      Else: strCFOP_ITEM = "6102"
   End If

   'Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, UF_CLIENTE_A, 0)
   If IsNumeric(strCFOP_ITEM) Then _
      CFOP_ID_N = 0 & strCFOP_ITEM
   Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, "", CFOP_ID_N)

   '5405  Venda de mercadoria, adquirida ou recebida de terceiros,
   'sujeita ao regime de substituição tributária,
   'na condição de contribuinte-substituído

   'Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros
   'em operação com mercadorias sujeitas ao regime de substituição tributária,
   'na condição de contribuinte substituído.
   'strCFOP_ITEM = "5405"   'não é industria


'28/03/2017 VERIFICAR SE É ASSIM MESMO:
'QUANDO CLIENTE É CONSUMIDOR FINAL NÃO PASSA NO SEFAZ O PRODUTO COMO SUBSTITUIÇÃO TRIBUTÁRIA
'DAI MUDO AQUI MANUALMENTE A ST DO ITEM PARA 00-TRIBUTADO INTEGRALMENTE
'SIMPLES NACIONAL
If CTR_EMPRESA_N = 1 Then _
   If Trim(UCase(CCE_CLIENTE_A)) = "ISENTO" Or Trim(CCE_CLIENTE_A) = "" Then _
      ST_PRODUTO_A = "00"


'TEM QUE VER QUANDO NÃO FOR VENDA

   Select Case ST_PRODUTO_A
      Case "00"   'Tributada integralmente
         'If CTR_EMPRESA_N = 1 Then  'se é optante do simples nacional
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
               strCFOP_ITEM = "5102"
               Else: strCFOP_ITEM = "6102"
            End If
         'End If

         If INDR_INDUSTRIA_B = True Then  'se produto o proprio produto pra vender
            'Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros,
            'na condição de contribuinte substituto,
            'em operação com mercadorias sujeitas ao regime de substituição tributária.
            strCFOP_ITEM = "5403"   'CFOP 5403 - Venda de mercadoria adquirida ou recebida de terceiros
                                    'em operação com mercadoria sujeita ao regime de substituição tributária,
                                    'na condição de contribuinte substituto

            If CTR_EMPRESA_N = 1 Then  'se é optante do simples nacional
               If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
                  'INDR_PRODUTO_PRODUCAO_B = vem da tabela de familia informando que o produto é de produção
                  If INDR_PRODUTO_PRODUCAO_B = True Then _
                     strCFOP_ITEM = "5101"   'CFOP 5101 - Venda de produção do estabelecimento
                  Else
                     strCFOP_ITEM = "6102"
                     If INDR_PRODUTO_PRODUCAO_B = True Then _
                        strCFOP_ITEM = "6101"   'CFOP 6101 - Venda de produção do estabelecimento
               End If
            End If
         End If   'If INDR_INDUSTRIA_B = True Then

         'Desconto nao entra no valor do ICMS de acordo com informacoes da CONTABILIDADE
         VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N

         '==================EU
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then       'DENTRO DO ESTADO ICMS NORMAL
            PERC_ICMS_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF
            Else: PERC_ICMS_N = ALIQUOTA_ICMS_NORMAL_FORA_UF   'FORA DO ESTADO ICMS NORMAL
         End If
         VALOR_ICMS_N = ((VALOR_BASE_ICMS_N * PERC_ICMS_N) / 100)
      Case "10"   'Tributada  e com cobrança do ICMS por substituição tributária
         VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
   
         Aliquota_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF
         If Trim(UF_CLIENTE_A) <> Trim(UF_EMPRESA_A) Then _
            Aliquota_N = ALIQUOTA_ICMS_NORMAL_FORA_UF
   
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
            'Campo IVA nao existe nao tabela verificar se precisa, Índices de Valor Agregado
            If Not IsNull(PERCIVA_A) Then _
              VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * PERCIVA_A) / 100)  'Valor da Reducao da base
   
            'VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * 1) / 100)  'Valor da Reducao da base
            VALOR_ICMS_PRODUTO_SUBST_N = ((VALOR_BASE_ICMS_SUBST_N * Aliquota_N) / 100)  'é fixo o percentual, procurar saber se tem como parametrizar
            VALOR_PERC_ICMS_SUBST_N = Aliquota_N
         End If
      Case "20"   'Com redução de base de cálculo
         If COMP_TRIBUTARIA_A = 0 Then 'tipos de maquinas, normais, agricolas, industriais
            If CCE_CLIENTE_A <> "" Then    'Tem que ter inscricao estadual
               VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100)
               PERC_REDUCAO_ICMS_N = TP2_DE_CONTRIB
               Else  'Sem inscricao estadual
                  VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NCONTRIB) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_NCONTRIB
            End If
         End If
   
         'Maquinas agricolas
         If COMP_TRIBUTARIA_A = 1 Then
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
               If CCE_CLIENTE_A <> "" Then
                  VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CMAQ_IMP) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_CMAQ_IMP
                  Else
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NMAQ_IMP) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_DE_NMAQ_IMP
               End If
               Else 'Fora do Estado
                  If CCE_CLIENTE_A <> "" Then
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_CMAQ_IMP) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_FE_CMAQ_IMP
                     Else
                        VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_NMAQ_IMP) / 100)
                        PERC_REDUCAO_ICMS_N = TP2_FE_NMAQ_IMP
                  End If
            End If
         End If
   
         If COMP_TRIBUTARIA_A = 2 Then 'Maquinas industriais
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
               If CCE_CLIENTE_A <> "" Then
                  VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_CONTRIB
                  Else
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NCONTRIB) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_DE_NCONTRIB
               End If
               Else 'Fora do Estado
                  If CCE_CLIENTE_A <> "" Then
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_CAP_INDU) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_FE_CAP_INDU
                     Else
                        VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_NAP_INDU) / 100)
                        PERC_REDUCAO_ICMS_N = TP2_FE_NAP_INDU
                  End If
            End If
         End If
      Case "30"   'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
         VALOR_BASE_ICMS_N = 0
         VALOR_ICMS_N = 0
         PERC_ICMS_N = 0
   
         If UCase(UF_CLIENTE_A) <> UCase(UF_EMPRESA_A) Then
             '//Desconto nao entra no valor de ICMS de Acordo com as
             '//Informacoes Contabeis
             '//move (ITENS.TOTAL_ITEM - ITENS.VLR_DESC_RATEIO)  ;
             '//                                     To   ITENS.VLR_BASE_ICMS
             VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
             '??? nao grava o percentual do aliquota?
         End If
      Case "40"   '40 Isenta
         VALOR_BASE_ICMS_N = 0
         VALOR_ICMS_N = 0
         PERC_ICMS_N = 0
      Case "41"   'Não tributada
         VALOR_BASE_ICMS_N = 0
         VALOR_ICMS_N = 0
         PERC_ICMS_N = 0
      Case "50"   'Suspensão
      Case "51"   'Diferimento
      Case "60"   'ICMS cobrado anteriormente por substituição tributária
         VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
         If UCase(UF_CLIENTE_A) = UCase(UF_EMPRESA_A) Then
            If TIPO_CLIENTE_N = 2 Then 'Atacado
               '//Dentro do Estado e Cliente Contribuinte ele e Isento
               '/Emanoel Informacoes Contabilidade dia 30/05/2006
               VALOR_BASE_ICMS_N = 0
               VALOR_ICMS_N = 0
               PERC_ICMS_N = 0
            End If
            'Só é tratado o tipo de cliente 2, atacado, e os outros tipos de clientes (varejo),
            'nao precisa tratar?
            Else 'Fora do estado
               If TIPO_CLIENTE_N = 2 Then 'Atacado
                  VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
                  'nao grava o percentual? porque?
               End If
         End If

         'DENTRO DO ESTADO
         If UCase(UF_CLIENTE_A) = UCase(UF_EMPRESA_A) Then
            'If Trim(ST_PRODUTO_A) = 60 Then
               'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
               'CFOP 5405 - Venda de mercadoria adquirida/recebida de terceiros em operação _
                            com mercadoria sujeita ao regime de substituição tributária, na condição de _
                            contrib substituído
       
      'portanto o que vai diferenciar se será um codigo ou outro será a mercadoria em
      'si...se ela é substituiçao tributaria ou nao...se for varias mercadorias vc tem que
      'verificar uma por uma pra saber.
      
               strCFOP_ITEM = "5405"
               'Else: strCFOP_ITEM = CFOP_SAIDA_DENTRO_UF_N                     'cfop de venda dentro do estado
            'End If   'If Trim(ST_PRODUTO_A) = 60 Then
         End If
      
         'FORA DO ESTADO
         If UCase(UF_CLIENTE_A) <> UCase(UF_EMPRESA_A) Then
            'If Trim(ST_PRODUTO_A) = 60 Then
               strCFOP_ITEM = "6403"  'Fixo por enquanto
               '6403 Venda de mercadoria adquirida ou recebida de terceiros em operação _
                     com mercadoria sujeita ao regime de substituição tributária, _
                     na condição de contribuinte substituto _
                     Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros, _
                     na condição de contribuinte substituto, em operação com mercadorias sujeitas _
                     ao regime de substituição tributária.
      
               strCFOP_ITEM = "6404"
               '6404 Venda de mercadoria sujeita ao regime de substituição tributária, _
                     cujo imposto já tenha sido retido anteriormente _
                     Classificam-se neste código as vendas de mercadorias sujeitas ao regime de substituição tributária, _
                     na condição de substituto tributário, exclusivamente nas hipóteses em que o _
                     imposto já tenha sido retido anteriormente
      
            '   Else: strCFOP_ITEM = CFOP_SAIDA_FORA_UF_N                  'cfop de venda fora do estado do estado
            'End If
      
            SQL = "select * from CFOP WITH (NOLOCK)"
            SQL = SQL & " Where CFOP_ID = '" & Trim(strCFOP_ITEM) & "'"
            TabTempCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabTempCFOP.EOF Then
               If TabTempCFOP.State = 1 Then _
                  TabTempCFOP.Close
      
               MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP_ITEM & vbCrLf & "Não é possivel continuar a processar"
               'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
               Exit Sub
               Else
                  If Trim(Len(CNPJCPF_A)) > 11 Then ' Se for pessoa juridica
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TabTempCFOP!ALIQUOTA_ICMS_DENTRO) / 100) 'CFOP.P_ICMS_VND_F_UF - verificar se existe
                     PERC_ICMS_N = TabTempCFOP!ALIQUOTA_ICMS_DENTRO ' CFOP.P_ICMS_VND_F_UF'duas aliquotas para  o mesmo cfop
                     Else ' Pessoa fisica
                        VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TabTempCFOP!ALIQUOTA_ICMS_FORA) / 100)
                        PERC_ICMS_N = TabTempCFOP!ALIQUOTA_ICMS_FORA
                  End If
            End If
            If TabTempCFOP.State = 1 Then _
               TabTempCFOP.Close
         End If
      Case "70"   'Com redução de base de cálculo e cobrança de ICMS por substituição tributária
      Case "90"   'Outras
   End Select

   If VALOR_BASE_ICMS_N = 0 Then
      PERC_ICMS_N = 0
      VALOR_ICMS_N = 0
   End If

'ATUALIZAR PEDIDO
   SQL = "UPDATE PEDIDOITEM SET "

   SQL = SQL & " VlrBaseIcms = " & tpMOEDA(VALOR_BASE_ICMS_N)
   SQL = SQL & ", PERCICMS = " & tpMOEDA(PERC_ICMS_N)
   SQL = SQL & ", VlrIcms = " & tpMOEDA(VALOR_ICMS_N)

   SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(VALOR_BASE_ICMS_SUBST_N)
   SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(VALOR_PERC_ICMS_SUBST_N)
   SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(VALOR_ICMS_PRODUTO_SUBST_N)

   SQL = SQL & ", cfop_id = '" & Trim(strCFOP_ITEM) & "'"
   SQL = SQL & ", STRIBUTARIA = '" & ST_PRODUTO_A & "'"

   SQL = SQL & " Where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and produto_id = " & PRODUTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREPARA_TRIBUTACAO_PRODUTO"
End Sub
