VERSION 5.00
Begin VB.Form frmlixo 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Chaves e relações"
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Importação dados CASA DAS PEÇAS"
      Enabled         =   0   'False
      Height          =   735
      Left            =   10080
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdImpCliente 
      Caption         =   "Importa Cliente (Selaria)"
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Importa Produto (Selaria)"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Importa IBGE (Selaria)"
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Importa FAMILIA Produto"
      Height          =   735
      Left            =   7560
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmlixo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImpCliente_Click()
   Dim DT_NASC_D     As String
   Dim rua_a         As String
   Dim bairro_a      As String
   Dim complemento_a As String
   Dim Numero_A      As String
   Dim DDD_N         As Integer
   Dim CNPJCPF_selaria As String
   Dim INSC_ESTADUAL As String

   ABRE_BANCO_AUXILIAR "SELARIA", SERVIDOR_SHFSYS

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from Tbl_MaDadosPe "
   SQL = SQL & " where Cad_cpfcgccli  is not null"
   SQL = SQL & " ORDER BY CAD_nomecli"
   TabCliente.Open SQL, CONECTA_AUXILIAR, , , adCmdText
   While Not TabCliente.EOF

      CNPJCPF_selaria = "" & Replace(TabCliente.Fields("Cad_cpfcgccli").Value, ".", "")
      CNPJCPF_selaria = Replace(CNPJCPF_selaria, "/", "")
      CNPJCPF_selaria = Replace(CNPJCPF_selaria, "-", "")
      CNPJCPF_selaria = Trim(CNPJCPF_selaria)

      INSC_ESTADUAL = "" & Replace(TabCliente.Fields("Cad_nidentidadecli").Value, ".", "")
      INSC_ESTADUAL = Replace(INSC_ESTADUAL, "/", "")
      INSC_ESTADUAL = Replace(INSC_ESTADUAL, "-", "")
      INSC_ESTADUAL = Trim(INSC_ESTADUAL)

        If Trim(INSC_ESTADUAL) = "" Then _
            insc_estatual = "ISENTO"

      NOME_A = "" & Trim(TabCliente.Fields("Cad_nomecli").Value)
      RAZAO_A = "" & Trim(TabCliente.Fields("Cad_Fantasia").Value)

      DT_EXP_D = Date
      If Not IsNull(TabCliente.Fields("Cad_Data").Value) Then _
         If IsDate(TabCliente.Fields("Cad_Data").Value) Then _
            DT_EXP_D = Trim(TabCliente.Fields("Cad_Data").Value)

      STATUS_A = "A"

      If Trim(RAZAO_A) = "" Then _
         RAZAO_A = NOME_A

      DT_NASC_D = ""
      If Not IsNull(TabCliente.Fields("Cad_dtnasccli").Value) Then _
         If IsDate(TabCliente.Fields("Cad_dtnasccli").Value) Then _
            DT_NASC_D = Trim(TabCliente.Fields("Cad_Data").Value)

      SQL3 = ""
      If Not IsNull(TabCliente.Fields("Cad_ObsCli").Value) Then _
         If Trim(TabCliente.Fields("Cad_ObsCli").Value) <> "" Then _
            SQL3 = Trim(TabCliente.Fields("Cad_ObsCli").Value)

      frmATUALIZACAO.Caption = "CRIANDO PESSOA = " & Trim(NOME_A)

      If TabPessoa.State = 1 Then _
         TabPessoa.Close

      SQL = "select * from PESSOA"
      SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
      TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabPessoa.EOF Then
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         SP_PESSOA "I", _
                   0, _
                   Trim(CNPJCPF_selaria), _
                   Trim(NOME_A), _
                   Trim(RAZAO_A), _
                   DT_EXP_D, _
                   STATUS_A
      End If

      PESSOA_ID_N = 0

      If TabPessoa.State = 1 Then _
         TabPessoa.Close

      SQL = "select pessoa_id from PESSOA"
      SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
      TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPessoa.EOF Then _
         PESSOA_ID_N = 0 & TabPessoa.Fields(0).Value

      If TabPessoa.State = 1 Then _
         TabPessoa.Close

'''''''''''''
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from CLIENTE "
      SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_selaria) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         SQL = "insert into CLIENTE "
            SQL = SQL & "(CLIENTE_ID,PESSOA_ID,EMPRESA_ID,VENDEDOR_ID,CGCCPF,NOME,RAZAO_SOCIAL,"
            SQL = SQL & "DT_NASC,DT_CAD,PROFISSAO,STATUS,SEXO,CONTATO,REGIAO,ORIGEM,LIMITE_CREDITO,"
            SQL = SQL & "ESTRANGEIRO,TIPO_CLIENTE,IE,IM,PERC_DESC_CONVENIO,OBS,CODG_SUFRAMA)"
         SQL = SQL & " VALUES("
            SQL = SQL & PESSOA_ID_N                            'CLIENTE_ID
            SQL = SQL & "," & PESSOA_ID_N                      'PESSOA_ID
            SQL = SQL & "," & EMPRESA_ID_N                     'EMPRESA_ID
            SQL = SQL & "," & 0                                'VENDEDOR_ID
            SQL = SQL & ",'" & Trim(CNPJCPF_selaria) & "'"           'CGCCPF
            SQL = SQL & ",'" & Trim(NOME_A) & "'"              'NOME
            SQL = SQL & ",'" & Trim(RAZAO_A) & "'"             'RAZAO_SOCIAL
            SQL = SQL & ",'" & DMA(DT_NASC_D) & "'"            'DT_NASC
            SQL = SQL & ",'" & DMA(DT_EXP_D) & "'"             'DT_CAD
            SQL = SQL & "," & 0                                'PROFISSAO
            SQL = SQL & ",'" & STATUS_A & "'"                  'Status
            SQL = SQL & ",''"                                  'SEXO
            SQL = SQL & ",''"                                  'CONTATO
            SQL = SQL & "," & 0                                'REGIAO
            SQL = SQL & ",''"                                  'Origem
            SQL = SQL & "," & tpMOEDA(0)                       'LIMITE_CREDITO
            SQL = SQL & ",'FALSE'"                             'ESTRANGEIRO
            SQL = SQL & "," & 0                                'TIPO_CLIENTE
            SQL = SQL & ",'" & Trim(INSC_ESTADUAL) & "'"       'IE
            SQL = SQL & ",'ISENTO'"                            'IM
            SQL = SQL & "," & 0                                'PERC_DESC_CONVENIO
            SQL = SQL & ",'" & Trim(SQL3) & "'"                'OBS
            SQL = SQL & "," & 0                                'CODG_SUFRAMA
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
      End If

'''''''''''''

      '======================ENDERECO

      CEP_A = ""
      If Not IsNull(TabCliente.Fields("Cad_cbocepcli").Value) Then
         CEP_A = Replace(TabCliente.Fields("Cad_cbocepcli").Value, " ", "")
         CEP_A = Replace(CEP_A, ".", "")
         CEP_A = Replace(CEP_A, "-", "")
         CEP_A = Trim(CEP_A)

         CIDADE_A = Trim(TabCliente.Fields("Cad_cidadecli").Value)
         UF_A = Trim(TabCliente.Fields("Cad_cloestadocli").Value)

         If TabCEP.State = 1 Then _
            TabCEP.Close
   
         SQL = "select * from CEP "
         SQL = SQL & " where cep = '" & Trim(CEP_A) & "'"
         TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabCEP.EOF Then
            SQL = "insert into CEP "
               SQL = SQL & " (cep,cidade,uf)"
            SQL = SQL & " VALUES ("
               SQL = SQL & "'" & Trim(CEP_A) & "'"
               SQL = SQL & ",'" & CIDADE_A & "'"
               SQL = SQL & ",'" & Trim(UF_A) & "'"
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close

         rua_a = "" & Trim(TabCliente.Fields("Cad_endrescli").Value)
         bairro_a = "" & Trim(TabCliente.Fields("Cad_bairrocli").Value)
         complemento_a = ""

         If tabEndereco.State = 1 Then _
            tabEndereco.Close
   
         SQL = "select distinct(prop) from ENDERECO "
         SQL = SQL & " where prop = '" & Trim(CNPJCPF_selaria) & "'"
         SQL = SQL & " and tipo = 'C' "
         tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If tabEndereco.EOF Then
            ENDERECO_ID_N = MAX_ID("endereco_id", "endereco", "", "", "", "")
   
            SQL = "insert into ENDERECO "
               SQL = SQL & " (ENDERECO_ID,PESSOA_ID,PROP,CEP,RUA,BAIRRO,COMPLEMENTO,TIPO,IE_ID,NUMERO)"
            SQL = SQL & " VALUES ("
               SQL = SQL & ENDERECO_ID_N                       'ENDERECO_ID
               SQL = SQL & "," & PESSOA_ID_N                   'PESSOA_ID
               SQL = SQL & ",'" & Trim(CNPJCPF_selaria) & "'"        'PROP
               SQL = SQL & ",'" & Trim(CEP_A) & "'"            'CEP
               SQL = SQL & ",'" & Trim(Left(rua_a, 50)) & "'"           'RUA
               SQL = SQL & ",'" & Trim(bairro_a) & "'"         'BAIRRO
               SQL = SQL & ",'" & Trim(complemento_a) & "'"    'COMPLEMENTO
               SQL = SQL & ",'C'"                              'TIPO
               SQL = SQL & ",0"                                'IE_ID
               SQL = SQL & ",'0'"                                'NUMERO
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If tabEndereco.State = 1 Then _
            tabEndereco.Close
      
         If tabEndereco.State = 1 Then _
            tabEndereco.Close
   
         SQL = "select distinct(prop) from ENDERECO "
         SQL = SQL & " where prop = '" & Trim(CNPJCPF_selaria) & "'"
         SQL = SQL & " and tipo = 'R' "
         tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If tabEndereco.EOF Then
            ENDERECO_ID_N = MAX_ID("endereco_id", "endereco", "", "", "", "")
   
            SQL = "insert into ENDERECO "
               SQL = SQL & " (ENDERECO_ID,PESSOA_ID,PROP,CEP,RUA,BAIRRO,COMPLEMENTO,TIPO,IE_ID,NUMERO)"
            SQL = SQL & " VALUES ("
               SQL = SQL & ENDERECO_ID_N                       'ENDERECO_ID
               SQL = SQL & "," & PESSOA_ID_N                   'PESSOA_ID
               SQL = SQL & ",'" & Trim(CNPJCPF_selaria) & "'"        'PROP
               SQL = SQL & ",'" & Trim(CEP_A) & "'"            'CEP
               SQL = SQL & ",'" & Trim(Left(rua_a, 50)) & "'"           'RUA
               SQL = SQL & ",'" & Trim(bairro_a) & "'"         'BAIRRO
               SQL = SQL & ",'" & Trim(complemento_a) & "'"    'COMPLEMENTO
               SQL = SQL & ",'R'"                              'TIPO
               SQL = SQL & ",0"                                'IE_ID
               SQL = SQL & ",'0'"                                'NUMERO
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If tabEndereco.State = 1 Then _
            tabEndereco.Close
      End If

'======================FONE
      If Not IsNull(TabCliente.Fields("Cad_foneresicli").Value) Then
         If Trim(TabCliente.Fields("Cad_foneresicli").Value) <> "" Then
            Numero_A = "" & Trim(TabCliente.Fields("Cad_foneresicli").Value)
            Numero_A = "" & Replace(Numero_A, " ", "")
            Numero_A = "" & Replace(Numero_A, ".", "")
            Numero_A = "" & Replace(Numero_A, "-", "")

            If TabFone.State = 1 Then _
               TabFone.Close

            SQL = "select * from fone "
            SQL = SQL & " where numero = '" & Trim(Numero_A) & "'"
            SQL = SQL & " and prop = '" & Trim(CNPJCPF_selaria) & "'"
            TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabFone.EOF Then
               SQL = "insert into FONE "
                  SQL = SQL & "(pessoa_id,prop,numero,ddd,LOCAL)"
               SQL = SQL & " values("
                  SQL = SQL & PESSOA_ID_N                         'PESSOA_ID
                  SQL = SQL & ",'" & Trim(CNPJCPF_selaria) & "'"        'PROP
                  SQL = SQL & ",'" & Trim(Numero_A) & "'"         'numero
                  SQL = SQL & "," & DDD_N                         'ddd
                  SQL = SQL & ",'Residência'"                     'local
               SQL = SQL & ")"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabFone.State = 1 Then _
               TabFone.Close
         End If
      End If
'======================FONE
      If Not IsNull(TabCliente.Fields("Cad_FoneComercial").Value) Then
         If Trim(TabCliente.Fields("Cad_FoneComercial").Value) <> "" Then
            Numero_A = "" & Trim(TabCliente.Fields("Cad_FoneComercial").Value)
            Numero_A = "" & Replace(Numero_A, " ", "")
            Numero_A = "" & Replace(Numero_A, ".", "")
            Numero_A = "" & Replace(Numero_A, "-", "")
            Numero_A = "" & Trim(Numero_A)

            If TabFone.State = 1 Then _
               TabFone.Close

            SQL = "select * from fone "
            SQL = SQL & " where numero = '" & Trim(Numero_A) & "'"
            SQL = SQL & " and prop = '" & Trim(CNPJCPF_selaria) & "'"
            TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabFone.EOF Then
               SQL = "insert into FONE "
                  SQL = SQL & "(pessoa_id,prop,numero,ddd,LOCAL)"
               SQL = SQL & " values("
                  SQL = SQL & PESSOA_ID_N                         'PESSOA_ID
                  SQL = SQL & ",'" & Trim(CNPJCPF_selaria) & "'"        'PROP
                  SQL = SQL & ",'" & Trim(Numero_A) & "'"         'numero
                  SQL = SQL & "," & DDD_N                         'ddd
                  SQL = SQL & ",'Comercial'"                      'local
               SQL = SQL & ")"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabFone.State = 1 Then _
               TabFone.Close
         End If
      End If

      cmdImpCliente.Caption = PESSOA_ID_N

      DoEvents

      TabCliente.MoveNext
   Wend

   If TabCliente.State = 1 Then _
      TabCliente.Close

   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

   MsgBox "fim clinte"
End Sub

Private Sub Command10_Click()
   ABRE_BANCO_AUXILIAR "SELARIA", SERVIDOR_SHFSYS

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from Tbl_Produto  "
   TabConsulta.Open SQL, CONECTA_AUXILIAR, , , adCmdText
   While Not TabConsulta.EOF

      PRODUTO_ID_N = MAX_ID("produto_id", "PRODUTO", "", "", "", "")

      CODG_PRODUTO_A = Trim(TabConsulta.Fields("Pro_codmerc").Value)

TRAVEIS:

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select codg_produto from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         If TabEmpresa.State = 1 Then _
            TabEmpresa.Close

         SQL = "select seq_codg_prod from EMPRESA "
         SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
         TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabEmpresa.EOF Then
            SQL = "update EMPRESA set seq_codg_prod = seq_codg_prod + 1"
            SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            If TabEmpresa.State = 1 Then _
               TabEmpresa.Close

            SQL = "select seq_codg_prod from EMPRESA "
            SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
            TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabEmpresa.EOF Then _
               CODG_PRODUTO_A = Trim(TabEmpresa.Fields("tabempresa").Value)

            GoTo TRAVEIS
         End If

         If TabEmpresa.State = 1 Then _
            TabEmpresa.Close
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL3 = "" & Trim(TabConsulta.Fields("pro_similares").Value)

      SQL = "insert into PRODUTO "
         SQL = SQL & " ("

            SQL = SQL & "PRODUTO_ID,EMPRESA_ID,CODG_PRODUTO,FORNECEDOR_ID,DESCRICAO,QTDE,QTDE_RETIDO,REFERENCIA,"
            SQL = SQL & "FAMILIAPRODUTO_ID,UNIDADE_MEDIDA,SITUACAO,SITUACAO_TRIBUTARIA,"
            SQL = SQL & "ALIQUOTA_ICMS,TIPO_PROD,CODG_NCM,PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda,DT_CADASTRO,"
            SQL = SQL & "NACIONAL,DT_ULT_COMPRA "
         
         SQL = SQL & " )"
      SQL = SQL & " VALUES( "
         SQL = SQL & PRODUTO_ID_N                                                   'PRODUTO_ID
         SQL = SQL & "," & EMPRESA_ID_N                                             'EMPRESA_ID
         SQL = SQL & ",'" & Trim(CODG_PRODUTO_A) & "'"                              'CODG_PRODUTO
         SQL = SQL & ",0"                                                           'FORNECEDOR_ID
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("pro_descricao").Value) & "'"   'DESCRICAO
         SQL = SQL & ",0"                                                           'QTDE
         SQL = SQL & ",0"                                                           'QTDE_RETIDO
         SQL = SQL & ",'" & Trim(SQL3) & "'"                                        'REFERENCIA
         SQL = SQL & ",0" & Trim(TabConsulta.Fields("pro_codgrp").Value)            'FAMILIAPRODUTO_ID
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("pro_unidade").Value) & "'"     'UNIDADE_MEDIDA
         SQL = SQL & ",'A'"                                                         'SITUACAO
         SQL = SQL & ",'00'"                                                        'SITUACAO_TRIBUTARIA
         SQL = SQL & ",'17'"                                                        'ALIQUOTA_ICMS
         SQL = SQL & ",1"                                                           'TIPO_PROD
         SQL = SQL & ",'00'"                                                        'CODG_NCM
         SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("pro_valoratacado").Value)    'PRECO_CUSTO
         SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("pro_valoratacado").Value)    'PRECO_ATACADO
         SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("pro_valorvArejo").Value)     'PRECO_Venda
         SQL = SQL & ",'" & DMA(TabConsulta.Fields("pro_datacompra").Value) & "'"   'DT_CADASTRO
         SQL = SQL & ",0"                                                           'NACIONAL
         SQL = SQL & ",'" & DMA(TabConsulta.Fields("pro_datacompra").Value) & "'"   'DT_ULT_COMPRA
      SQL = SQL & " )"
      
      CONECTA_RETAGUARDA.Execute SQL

      If TabProduto.State = 1 Then _
         TabProduto.Close

Command10.Caption = Trim(TabConsulta.Fields("PRO_DESCRICAO").Value)
DoEvents

      TabConsulta.MoveNext
   Wend
   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

   MsgBox "fim PRODUTO"
End Sub

Private Sub Command11_Click()
   ABRE_BANCO_AUXILIAR "SELARIA", SERVIDOR_SHFSYS

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from CEP  "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      If TabCEP.State = 1 Then _
         TabCEP.Close

      SQL = "select Mun_CodigoIbgeMun from Tbl_Municipio "
      SQL = SQL & " where Mun_CodigoEstadoUF = '" & Trim(TabConsulta.Fields("uf").Value) & "'"
      SQL = SQL & " and Mun_CodigoEstadoUFDes = '" & Trim(TabConsulta.Fields("cidade").Value) & "'"
      TabCEP.Open SQL, CONECTA_AUXILIAR, , , adCmdText
      If Not TabCEP.EOF Then
         SQL = "update cep set codigo_ibge = " & TabCEP.Fields(0).Value
         SQL = SQL & " where uf = '" & Trim(TabConsulta.Fields("uf").Value) & "'"
         SQL = SQL & " and cidade = '" & Trim(TabConsulta.Fields("cidade").Value) & "'"
         SQL = SQL & " and cep = '" & Trim(TabConsulta.Fields("cep").Value) & "'"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabCEP.State = 1 Then _
         TabCEP.Close

Me.Caption = Trim(TabConsulta.Fields("cidade").Value)
DoEvents

      TabConsulta.MoveNext
   Wend
   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

   MsgBox "fim IBGE"
End Sub

Private Sub Command12_Click()
   ABRE_BANCO_AUXILIAR "SELARIA", SERVIDOR_SHFSYS

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from Tbl_Grupo "
   TabConsulta.Open SQL, CONECTA_AUXILIAR, , , adCmdText
   While Not TabConsulta.EOF

      SQL = "insert into FAMILIAPRODUTO "
         SQL = SQL & " (FAMILIAPRODUTO_ID,CODG_FAMILIA,DESCRICAO,UNIDADE_MEDIDA,DESC_UNIDADE_MEDIDA)"
      SQL = SQL & " VALUES("
         SQL = SQL & Trim(TabConsulta.Fields("Grp_Codigo").Value)                   'FAMILIAPRODUTO_ID
         SQL = SQL & "," & Trim(TabConsulta.Fields("Grp_Codigo").Value)             'CODG_FAMILIA
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("Grp_Descricao").Value) & "'"   'Descricao
         SQL = SQL & ",'UN'"                                                        'Unidade_Medida
         SQL = SQL & ",'UNIDADE'"                                                   'DESC_UNIDADE_MEDIDA
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

Me.Caption = Trim(TabConsulta.Fields("Grp_Descricao").Value)
DoEvents

      TabConsulta.MoveNext
   Wend
   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

   MsgBox "fim FAMILIA"
End Sub

Private Sub Command2_Click()
   CONT_N = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "SELECT * From IVPRODUT2"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      CONT_N = CONT_N + 1
      Command2.Caption = CONT_N

      CRITERIO = Replace(TabConsulta.Fields("DESCRICAO").Value, ",", ".")
      CRITERIO = Replace(CRITERIO, "'", "´")

      SQL = "insert into PRODUTO "
      SQL = SQL & "("
SQL = SQL & "EMPRESA_ID,PRODUTO_ID,CODG_PRODUTO,DESCRICAO,FAMILIAPRODUTO_ID,UNIDADE_MEDIDA,SITUACAO,CFOP,"
SQL = SQL & "SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,TIPO_PROD,CODG_NCM,PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda,DT_CADASTRO,STATUS,NACIONAL,REFERENCIA,QTDE"
      SQL = SQL & ")"
      SQL = SQL & "VALUES ("
      SQL = SQL & 1                                                                             'EMPRESA_ID
      SQL = SQL & "," & Trim(TabConsulta.Fields("CODIGO_PRD").Value)                            'PRODUTO_ID
      SQL = SQL & ",'" & Trim(Replace(TabConsulta.Fields("CODIGO_PRD").Value, ",", ".")) & "'"  'CODG_PRODUTO
      SQL = SQL & ",'" & Trim(CRITERIO) & "'"                                                   'DESCRICAO
      SQL = SQL & "," & Trim(TabConsulta.Fields("CODIGO_GRP").Value)             'FAMILIAPRODUTO_ID
      SQL = SQL & ",'" & Trim(TabConsulta.Fields("UNIDADE").Value) & "'"         'UNIDADE_MEDIDA
      SQL = SQL & ",'A'"                                                         'SITUACAO
      SQL = SQL & ",0" & Trim(TabConsulta.Fields("CFOP").Value)                  'CFOP
      SQL = SQL & ",'" & Trim(TabConsulta.Fields("C_TRB_ICMS").Value) & "'"      'SITUACAO_TRIBUTARIA
      SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("ICM_VENDA").Value)           'ALIQUOTA_ICMS
      SQL = SQL & ",0"                                                           'TIPO_PROD
      SQL = SQL & ",0"                                                           'CODG_NCM
      SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("PRECO_CUST").Value)          'PRECO_CUSTO
      SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("PRD_VR_VDA").Value)          'PRECO_ATACADO
      SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("PRD_VR_VDA").Value)          'PRECO_Venda
      SQL = SQL & ",'" & DMA(Date) & "'"                                         'DT_CADASTRO
      SQL = SQL & ",'A'"                                                         'STATUS
      SQL = SQL & ",0"                                                           'NACIONAL
      SQL = SQL & ",'" & Trim(TabConsulta.Fields("CD_RED_PRD").Value) & "'"      'REFERENCIA
      SQL = SQL & ",0"                                                           'QTDE

'      ,[REFERENCIA]
'      ,[TIPO_PRODU]
'      ,[SIT_TRIBUT]
'      ,[B_CALC_ICM]
'      ,[ICM_VENDA2]
'      ,[VC_ULT_MOV]

      SQL = SQL & ")"

      CONECTA_RETAGUARDA.Execute SQL

      TabConsulta.MoveNext

      DoEvents
   Wend

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   MsgBox "Importados = " & CONT_N
End Sub



Private Sub Command9_Click()
   Dim TAB_SHEMA  As New ADODB.Recordset

   CONT_N = 0
   NUMR_SEQ_N = 0

   If TAB_SHEMA.State = 1 Then _
      TAB_SHEMA.Close

   SQL = "SELECT distinct(table_name) FROM INFORMATION_SCHEMA.COLUMNS"
   SQL = SQL & " WHERE COLUMN_NAME = 'PROP'"
   SQL = SQL & " order by table_name"
   TAB_SHEMA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TAB_SHEMA.EOF
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("email") Then
         If ExisteCampo("PESSOA_ID", "email") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE email ADD PESSOA_ID BIGINT"

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CNPJCPF_selaria = Trim(TabTemp.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update EMAIL set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabTemp.MoveNext
            DoEvents
         Wend
   
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("Endereco") Then
         If ExisteCampo("PESSOA_ID", "endereco") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE endereco ADD PESSOA_ID BIGINT"

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CNPJCPF_selaria = Trim(TabTemp.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update endereco set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabTemp.MoveNext
            DoEvents
         Wend
   
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("fone") Then
         If ExisteCampo("PESSOA_ID", "FONE") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE FONE ADD PESSOA_ID BIGINT"

         If TabFone.State = 1 Then _
            TabFone.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabFone.EOF
            CNPJCPF_selaria = Trim(TabFone.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update FONE set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabFone.MoveNext
            DoEvents
         Wend
   
         If TabFone.State = 1 Then _
            TabFone.Close
      End If
      
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("lancamento") Then
         If ExisteCampo("PESSOA_ID", "LANCAMENTO") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO ADD PESSOA_ID BIGINT"

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabLancamento.EOF
            CNPJCPF_selaria = Trim(TabLancamento.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update LANCAMENTO set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabLancamento.MoveNext
            DoEvents
         Wend
   
         If TabLancamento.State = 1 Then _
            TabLancamento.Close
      End If
      
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("nf") Then
         If ExisteCampo("PESSOA_ID", "NF") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD PESSOA_ID BIGINT"

         If TabNF.State = 1 Then _
            TabNF.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabNF.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabNF.EOF
            CNPJCPF_selaria = Trim(TabNF.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update NF set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabNF.MoveNext
            DoEvents
         Wend
   
         If TabNF.State = 1 Then _
            TabNF.Close
      End If
      
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("obs") Then
         If ExisteCampo("PESSOA_ID", "OBS") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OBS ADD PESSOA_ID BIGINT"

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CNPJCPF_selaria = Trim(TabTemp.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update OBS set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabTemp.MoveNext
         Wend
   
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("RG") Then
         If ExisteCampo("PESSOA_ID", "RG") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE RG ADD PESSOA_ID BIGINT"

         If TabRG.State = 1 Then _
            TabRG.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabRG.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabRG.EOF
            CNPJCPF_selaria = Trim(TabRG.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update RG set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabRG.MoveNext
            DoEvents
         Wend
   
         If TabRG.State = 1 Then _
            TabRG.Close
      End If
      
      If Trim(UCase(TAB_SHEMA.Fields(0).Value)) = UCase("tipoempresacliente") Then
         If ExisteCampo("PESSOA_ID", "TIPOEMPRESACLIENTE") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOEMPRESACLIENTE ADD PESSOA_ID BIGINT"

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from " & Trim(TAB_SHEMA.Fields(0).Value)
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CNPJCPF_selaria = Trim(TabTemp.Fields("prop").Value)

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_selaria) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update TIPOEMPRESACLIENTE set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where PROP = '" & Trim(CNPJCPF_selaria) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               NUMR_SEQ_N = NUMR_SEQ_N + 1
               Command8.Caption = NUMR_SEQ_N
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
            TabTemp.MoveNext
            DoEvents
         Wend
   
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

CONT_N = CONT_N + 1
Command9.Caption = CONT_N
DoEvents

      TAB_SHEMA.MoveNext
   Wend

   If TAB_SHEMA.State = 1 Then _
      TAB_SHEMA.Close

End Sub

Private Sub Command8_Click()
   'VERIFICA_BANCO_DADOS

'=============TABELA PESSOA
   If ExisteTabela("pk_PESSOA", "") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PESSOA ADD CONSTRAINT pk_PESSOA PRIMARY KEY (PESSOA_ID)"

'=============TABELA EMPRESA
   If ExisteTabela("pk_EMPRESA", "") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD CONSTRAINT pk_EMPRESA PRIMARY KEY (EMPRESA_ID)"

   If ExisteTabela("FK_EMPRESA_PESSOA", "") = False Then
      SQL = "ALTER TABLE [dbo].[EMPRESA]  WITH CHECK ADD  CONSTRAINT [FK_EMPRESA_PESSOA] FOREIGN KEY([PESSOA_ID])"
      SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[EMPRESA] CHECK CONSTRAINT [FK_EMPRESA_PESSOA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

'=============
   SQL = "ALTER TABLE NF ADD CONSTRAINT pk_NF PRIMARY KEY (NF_ID)"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[NFITEM]  WITH CHECK ADD  CONSTRAINT [FK_NFITEM_NF] FOREIGN KEY([NF_ID]) "
   SQL = SQL & " References [dbo].[NF]([NF_ID])"
   CONECTA_RETAGUARDA.Execute SQL

'=============

   SQL = "ALTER TABLE TRANSPORTADORA ADD CONSTRAINT pk_TRANSP PRIMARY KEY (TRANSP_ID)"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[NF]  WITH CHECK ADD  CONSTRAINT [FK_NF_TRANSPORTADORA] FOREIGN KEY([TRANSP_ID])"
   SQL = SQL & " References [dbo].[TRANSPORTADORA]([TRANSP_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[NF] CHECK CONSTRAINT [FK_NF_TRANSPORTADORA]"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[TRANSPORTADORA]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORA_PESSOA] FOREIGN KEY([PESSOA_ID])"
   SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[TRANSPORTADORA] CHECK CONSTRAINT [FK_TRANSPORTADORA_PESSOA]"
   CONECTA_RETAGUARDA.Execute SQL


'=============
   SQL = "ALTER TABLE USUARIO ADD CONSTRAINT pk_USUARIO PRIMARY KEY (USUARIO_ID)"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[USUARIO]  WITH CHECK ADD  CONSTRAINT [FK_USUARIO_PESSOA] FOREIGN KEY([PESSOA_ID])"
   SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[USUARIO] CHECK CONSTRAINT [FK_USUARIO_PESSOA]"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[USUARIO]  WITH CHECK ADD  CONSTRAINT [FK_USUARIO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
   SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "ALTER TABLE [dbo].[USUARIO] CHECK CONSTRAINT [FK_USUARIO_EMPRESA]"
   CONECTA_RETAGUARDA.Execute SQL

'=============

'=============

   SQL = "ALTER TABLE VENDEDOR ADD CONSTRAINT pk_VENDEDOR PRIMARY KEY (VENDEDOR_ID)"
   CONECTA_RETAGUARDA.Execute SQL

'=============

'=============

   SQL = "ALTER TABLE CLIENTE ADD CONSTRAINT pk_CLIENTE PRIMARY KEY (CLIENTE_ID)"
   CONECTA_RETAGUARDA.Execute SQL

'=============

'=============

   SQL = "ALTER TABLE PRODUTO ADD CONSTRAINT pk_PRODUTO PRIMARY KEY (PRODUTO_ID)"
   CONECTA_RETAGUARDA.Execute SQL

'=============

'=============

   SQL = "ALTER TABLE PEDIDO ADD CONSTRAINT pk_PEDIDO PRIMARY KEY (PEDIDO_ID)"
   CONECTA_RETAGUARDA.Execute SQL

'=============

'=============
   SQL = "ALTER TABLE CUPOM ADD CONSTRAINT pk_cupom PRIMARY KEY (cupom_id)"
   CONECTA_RETAGUARDA.Execute SQL

SQL = "ALTER TABLE [dbo].[CUPOM] "
SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_CUPOM_PEDIDO] "
SQL = SQL & " FOREIGN KEY([PEDIDO_ID])"
SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = " ALTER TABLE [dbo].[CUPOM] CHECK CONSTRAINT [FK_CUPOM_PEDIDO]"
CONECTA_RETAGUARDA.Execute SQL
'=============

'=============

   SQL = "ALTER TABLE PEDIDOITEM ADD CONSTRAINT pk_PEDIDOITEM PRIMARY KEY (PEDIDO_ID,SEQ_ID)"
   CONECTA_RETAGUARDA.Execute SQL

'''''''''''
SQL = "ALTER TABLE [dbo].[PEDIDOITEM] "
SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_PEDIDOITEM_PEDIDO] "
SQL = SQL & " FOREIGN KEY([PEDIDO_ID])"
SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = " ALTER TABLE [dbo].[PEDIDOITEM] CHECK CONSTRAINT [FK_PEDIDOITEM_PEDIDO]"
CONECTA_RETAGUARDA.Execute SQL

'''''''''''
SQL = "ALTER TABLE [dbo].[PEDIDOITEM] "
SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_PEDIDOITEM_PRODUTO] "
SQL = SQL & " FOREIGN KEY([PRODUTO_ID])"
SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = " ALTER TABLE [dbo].[PEDIDOITEM] CHECK CONSTRAINT [FK_PEDIDOITEM_PRODUTO]"
CONECTA_RETAGUARDA.Execute SQL

'''''''''''
SQL = "ALTER TABLE [dbo].[CLIENTE] "
SQL = SQL & " WITH CHECK ADD CONSTRAINT [FK_CLIENTE_PESSOA] "
SQL = SQL & " FOREIGN KEY([PESSOA_ID])"
SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
CONECTA_RETAGUARDA.Execute SQL
'''''''''''
'''''''''''
SQL = "ALTER TABLE CAIXADIA ADD CONSTRAINT pk_CAIXADIA PRIMARY KEY (CAIXADIA_ID)"
CONECTA_RETAGUARDA.Execute SQL

SQL = "ALTER TABLE [dbo].[CAIXADIAITEM] "
SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_CAIXADIAITEM_CAIXADIA] "
SQL = SQL & " FOREIGN KEY([CAIXADIA_ID])"
SQL = SQL & " References [dbo].[CAIXADIA]([CAIXADIA_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = " ALTER TABLE [dbo].[CAIXADIAITEM] CHECK CONSTRAINT [FK_CAIXADIAITEM_CAIXADIA]"
CONECTA_RETAGUARDA.Execute SQL
'''''''''''

'''''''''''
SQL = "ALTER TABLE CAIXATESORARIA ADD CONSTRAINT pk_CAIXATESORARIA PRIMARY KEY (CAIXATESORARIA_ID)"
CONECTA_RETAGUARDA.Execute SQL

SQL = "ALTER TABLE [dbo].[CAIXATESORARIAITEM] "
SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_CAIXATESORARIAITEM_CAIXATESORARIA] "
SQL = SQL & " FOREIGN KEY([CAIXATESORARIA_ID])"
SQL = SQL & " References [dbo].[CAIXATESORARIA]([CAIXATESORARIA_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = " ALTER TABLE [dbo].[CAIXATESORARIAITEM] CHECK CONSTRAINT [FK_CAIXATESORARIAITEM_CAIXATESORARIA]"
CONECTA_RETAGUARDA.Execute SQL
'''''''''''
'''''''''''

SQL = "ALTER TABLE [dbo].[CAIXADIA]  WITH CHECK ADD  CONSTRAINT [FK_CAIXADIA_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = "ALTER TABLE [dbo].[CAIXADIA] CHECK CONSTRAINT [FK_CAIXADIA_EMPRESA]"
CONECTA_RETAGUARDA.Execute SQL


SQL = "ALTER TABLE [dbo].[CAIXATESORARIA]  WITH CHECK ADD  CONSTRAINT [FK_CAIXATESORARIA_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
CONECTA_RETAGUARDA.Execute SQL

SQL = "ALTER TABLE [dbo].[CAIXATESORARIA] CHECK CONSTRAINT [FK_CAIXATESORARIA_EMPRESA]"
CONECTA_RETAGUARDA.Execute SQL

'''''''''''
'SQL = "ALTER TABLE [dbo].[PEDIDO] "
'SQL = SQL & " WITH CHECK ADD CONSTRAINT [FK_PEDIDO_CLIENTE] "
'SQL = SQL & " FOREIGN KEY([CLIENTE_ID])"
'SQL = SQL & " References [dbo].[Cliente]([CLIENTE_ID])"
'CONECTA_RETAGUARDA.Execute SQL

'=============
MsgBox "FIM"
End Sub



Private Sub PREPARA_TRIBUTACAO_PRODUTO_OLD()
'On Error GoTo ERRO_TRATA

'Duvidas
'- 13/06/2006 Quando o item for subsituicao ou do tipo tributario = 60, ele terá dois valores de icms
' ou somente um valor. Exemplificando, se for 100,00 ele tera uma aliquota de 17% e outra de 10% por exemplo
'ou somente sera cobrado uma aliquota? Pergunto isto pois se houver dois valores para o mesmo item devera
'ser criado um outro registro no banco de dados.

   Dim rstProduto                As New ADODB.Recordset
   Dim RstTemp                   As New ADODB.Recordset
   Dim VALOR_BASE_ICMS_N         As Double
   Dim VALOR_PERC_ICMS_N         As Double
   Dim VALOR_ICMS_PRODUTO        As Double
   Dim VALOR_BASE_ICMS_SUBST_N   As Double
   Dim VALOR_ICMS_PRODUTOSubst   As Double
   Dim VALOR_PERC_ICMS_SUBST_N    As Double
   Dim dblPercReducICMS          As Double
   Dim dblPercIVA                As Double
   Dim dblTotalItem              As Double

   If Trim(CODG_PRODUTO_A) = "" Then
      MsgBox "Produto não informado, verifique !!!"
      Exit Sub
   End If
   If CLIENTE_ID_N <= 0 Then
      MsgBox "Cliente não informado, verifique !!!"
      Exit Sub
   End If

   VALOR_BASE_ICMS_N = 0
   VALOR_PERC_ICMS_N = 0
   VALOR_ICMS_PRODUTO = 0
   VALOR_BASE_ICMS_SUBST_N = 0
   VALOR_ICMS_PRODUTOSubst = 0
   VALOR_PERC_ICMS_SUBST_N = 0
   dblPercReducICMS = 0
   dblPercIVA = 0
   dblTotalItem = 0
   strCFOP = ""
   SITUAÇÃO_TRIBUTARIA_PRODUTO = ""

   If Trim(UF_CLIENTE) = "" Then _
      TRATA_CLIENTE

   If USA_NFe = True Then
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
         If Trim(UF_CLIENTE) = "" Then
            MsgBox "Cliente com cadastro incompleto !!!"
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      End If
   End If

   If UF_EMPRESA_A = "" Then _
      PEGA_DADOS_EMPRESA

   dblTotalItem = (txtQTDE.Text * txtValor_Unitario.Text)

   If rstProduto.State = 1 Then _
      rstProduto.Close

   SQL = "Select situacao_tributaria,perciva,comp_tributaria from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " Where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   rstProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstProduto.EOF Then
      If rstProduto.State = 1 Then _
         rstProduto.Close

      MsgBox "O sistema nao localizou nenhum produto com o seguinte codigo: " & CODG_PRODUTO_A & vbCrLf & "Verique"
      Exit Sub
   End If

   'Inicio yuri 01/05/2012
     ' Aqui será colocado a rotina para calcular os tributos em substituição a toda essa regra que esta
     ' nesta instrução
     ' busca aliquota do Unidade federativa do Cliente
     ' aqui nao retirar aqui vamos dar o inicio a toda carga tributaria
     ' comentei aqui para nao atraplhar se codigo
   'Call BuscaAliquota(strUFCliente, CLng(ClienteId))

   ' fim yuri 01/05/2012

   'Tentando fazer igual o dataflex faz
   '//Impostos  Tributos
   '// ---- Calculo das Reducoes de ICMS e Substituicao Tributaria -------- //
    '  //0 = Tributado integralmente
    '  //1 = Tributado e com cobranca do ICMS por Substituicao Tributaria
    '  //2 = Com Reducao de Base de Calculo
    '  //3 = Isenta ou nao tributada e com cobranca do ICMS por Sub. Tributaria
    '  //4 = Isenta ou nao Tributado
    '  //5 = Com Suspensao ou diferimento
    '  //6 = ICMS cobrado anteriormente por subst. Tributaria
    '  //7 = Com reducao de base de Calculo e Cobranca do icms por Subst. Tributaria
    '  //9 = Outras
    '  //Compensacao Tribuaria
    '  //0 = Mercadorias Normais
    '  //1 = Maquinas e Implementos Agricolas
    '  //2 = Maquinas Aparelhos Equipamentos Industriais

'==========banco de dados
'CODIGO  DESCRICAO
'00      Tributada integralmente
'10      Tributada  e com cobrança do ICMS por substituição tributária
'20      Com redução de base de cálculo
'30      Isenta ou não tributada e com cobrança do ICMS por substituição tributária
'40      Isenta
'41      Não tributada
'50      Suspensão
'51      Diferimento
'60      ICMS cobrado anteriormente por substituição tributária
'70      Com redução de base de cálculo e cobrança de ICMS por substituição tributária
'90      Outras
'==========banco de dados

   'Tributada integralmente
   If rstProduto!SITUACAO_TRIBUTARIA = "00" Then
      'Desconto nao entra no valor do ICMS de acordo com informacoes
      'da CONTABILIDADE
      VALOR_BASE_ICMS_N = dblTotalItem

      'Criar campo de TIPO DE CLIENTE NO CADASTRO DE CLIENTE
      If dblTipoCliente = 2 Then
         If Trim(UF_CLIENTE) = Trim(UF_EMPRESA_A) Then
            VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_CONTRIB) / 100)  'Valor da Reducao da base
            VALOR_PERC_ICMS_N = TP2_DE_CONTRIB                ' Percentual da reducao
         End If
      End If
   End If

   'Tributada e com cobrança do ICMS por substituição tributária
   If rstProduto!SITUACAO_TRIBUTARIA = 10 Then 'Substituicao Tributaria
      VALOR_BASE_ICMS_N = dblTotalItem

      If Trim(UF_CLIENTE) = Trim(UF_EMPRESA_A) Then
         'Campo IVA nao existe nao tabela verificar se precisa
         If Not IsNull(rstProduto!PERCIVA) Then _
           VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * rstProduto!PERCIVA) / 100)  'Valor da Reducao da base

         'VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * 1) / 100)  'Valor da Reducao da base
         VALOR_ICMS_PRODUTOSubst = ((VALOR_BASE_ICMS_SUBST_N * 17) / 100) 'é fixo o percentual, procurar saber se tem como parametrizar
         VALOR_PERC_ICMS_SUBST_N = 17
      End If
   End If

   'Com redução de base de cálculo
   If rstProduto!SITUACAO_TRIBUTARIA = 20 Then 'Reducao da base de calculo
      If rstProduto!COMP_TRIBUTARIA = 0 Then 'tipos de maquinas, normais, agricolas, industriais
         If strInscEstadual <> "" Then   'Tem que ter inscricao estadual
            VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_CONTRIB) / 100)
            dblPercReducICMS = TP2_DE_CONTRIB
            Else  'Sem inscricao estadual
               VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_NCONTRIB) / 100)
               dblPercReducICMS = TP2_DE_NCONTRIB
         End If
      End If

      'Maquinas agricolas
      If rstProduto!COMP_TRIBUTARIA = 1 Then
         If Trim(UF_CLIENTE) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
            If strInscEstadual <> "" Then
               VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_CMAQ_IMP) / 100)
               dblPercReducICMS = TP2_DE_CMAQ_IMP
               Else
                  VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_NMAQ_IMP) / 100)
                  dblPercReducICMS = TP2_DE_NMAQ_IMP
            End If
            Else 'Fora do Estado
               If strInscEstadual <> "" Then
                  VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_FE_CMAQ_IMP) / 100)
                  dblPercReducICMS = TP2_FE_CMAQ_IMP
                  Else
                     VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_FE_NMAQ_IMP) / 100)
                     dblPercReducICMS = TP2_FE_NMAQ_IMP
               End If
         End If
      End If

      If rstProduto!COMP_TRIBUTARIA = 2 Then 'Maquinas industriais
         If Trim(UF_CLIENTE) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
            If strInscEstadual <> "" Then
               VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_CONTRIB) / 100)
               dblPercReducICMS = TP2_DE_CONTRIB
               Else
                  VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_DE_NCONTRIB) / 100)
                  dblPercReducICMS = TP2_DE_NCONTRIB
            End If
            Else 'Fora do Estado
               If strInscEstadual <> "" Then
                  VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_FE_CAP_INDU) / 100)
                  dblPercReducICMS = TP2_FE_CAP_INDU
                  Else
                     VALOR_BASE_ICMS_N = ((dblTotalItem * TP2_FE_NAP_INDU) / 100)
                     dblPercReducICMS = TP2_FE_NAP_INDU
               End If
         End If
      End If
   End If

   'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
   If rstProduto!SITUACAO_TRIBUTARIA = 30 Then '//Isenta ou nao Tributada Com ICMS por Subs. Trib
      VALOR_BASE_ICMS_N = 0
      VALOR_PERC_ICMS_N = 0

      If UCase(UF_CLIENTE) <> UCase(UF_EMPRESA_A) Then
          '//Desconto nao entra no valor de ICMS de Acordo com as
          '//Informacoes Contabeis
          '//move (ITENS.TOTAL_ITEM - ITENS.VLR_DESC_RATEIO)  ;
          '//                                     To   ITENS.VLR_BASE_ICMS
          VALOR_BASE_ICMS_N = dblTotalItem
          '??? nao grava o percentual do aliquota?
      End If
   End If

   'Isenta ou Não tributada
   If rstProduto!SITUACAO_TRIBUTARIA = 40 Or rstProduto!SITUACAO_TRIBUTARIA = 41 Then '//Isento ou nao Tributado
      VALOR_BASE_ICMS_N = 0
      VALOR_PERC_ICMS_N = 0
   End If

'50      Suspensão
'51      Diferimento

   'ICMS cobrado anteriormente por substituição tributária
   If rstProduto!SITUACAO_TRIBUTARIA = 60 Then '//Situacao Tributaria com Substitui‡ao Tributaria
      '//Desconto nao entra no valor de ICMS de Acordo com as
      '//Informacoes Contabeis

      VALOR_BASE_ICMS_N = dblTotalItem
      If UCase(UF_CLIENTE) = UCase(UF_EMPRESA_A) Then
         If dblTipoCliente = 2 Then 'Atacado
            '//Dentro do Estado e Cliente Contribuinte ele e Isento
            '/Emanoel Informacoes Contabilidade dia 30/05/2006
            VALOR_BASE_ICMS_N = 0
            VALOR_PERC_ICMS_N = 0
         End If

         'Só é tratado o tipo de cliente 2, atacado, e os outros tipos de clientes (varejo),
         'nao precisa tratar?
         Else 'Fora do estado
            If dblTipoCliente = 2 Then 'Atacado
               VALOR_BASE_ICMS_N = dblTotalItem
               'nao grava o percentual? porque?
            End If
      End If
   End If

'70      Com redução de base de cálculo e cobrança de ICMS por substituição tributária
'90      Outras

'========================================================================
'========================================================================
'========================================================================

   'If Not IsNull(rstProduto.Fields("CFOP_id").Value) Then
      
   'End If

   'DENTRO DO ESTADO
   If UCase(UF_CLIENTE) = UCase(UF_EMPRESA_A) Then
      If rstProduto!SITUACAO_TRIBUTARIA = 60 Then
         'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
         'CFOP 5405 - Venda de mercadoria adquirida/recebida de terceiros em operação _
                      com mercadoria sujeita ao regime de substituição tributária, na condição de _
                      contrib substituído
 
'portanto o que vai diferenciar se será um codigo ou outro será a mercadoria em
'si...se ela é substituiçao tributaria ou nao...se for varias mercadorias vc tem que
'verificar uma por uma pra saber.

         strCFOP = "5405"
         Else: strCFOP = CFOP_SAIDA_DE                     'cfop de venda dentro do estado
      End If

      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "Select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " Where CFOP_ID = '" & Trim(strCFOP) & "'"
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If RstTemp.EOF Then
         If RstTemp.State = 1 Then _
            RstTemp.Close

         If rstProduto.State = 1 Then _
            rstProduto.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      'if rstTEMP!Tipo = 0 then 'Dentro do Estado
      VALOR_ICMS_PRODUTO = ((dblTotalItem * RstTemp!PERC_ICMS) / 100)
      VALOR_PERC_ICMS_N = RstTemp!PERC_ICMS

      If RstTemp.State = 1 Then _
         RstTemp.Close
   End If

   'FORA DO ESTADO
   If UCase(UF_CLIENTE) <> UCase(UF_EMPRESA_A) Then
      If rstProduto!SITUACAO_TRIBUTARIA = 60 Then
         strCFOP = "6403"  'Fixo por enquanto
         '6403 Venda de mercadoria adquirida ou recebida de terceiros em operação _
               com mercadoria sujeita ao regime de substituição tributária, _
               na condição de contribuinte substituto _
               Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros, _
               na condição de contribuinte substituto, em operação com mercadorias sujeitas _
               ao regime de substituição tributária.

         strCFOP = "6404"
         '6404 Venda de mercadoria sujeita ao regime de substituição tributária, _
               cujo imposto já tenha sido retido anteriormente _
               Classificam-se neste código as vendas de mercadorias sujeitas ao regime de substituição tributária, _
               na condição de substituto tributário, exclusivamente nas hipóteses em que o _
               imposto já tenha sido retido anteriormente

         Else: strCFOP = CFOP_SAIDA_FE                  'cfop de venda fora do estado do estado
      End If

      SQL = "Select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " Where CFOP_ID = '" & Trim(strCFOP) & "'"
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If RstTemp.EOF Then
         If RstTemp.State = 1 Then _
            RstTemp.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      If Trim(Len(strCPFCNPJ)) > 11 Then ' Se for pessoa juridica
         VALOR_ICMS_PRODUTO = ((dblTotalItem * RstTemp!PERC_ICMS) / 100)  'CFOP.P_ICMS_VND_F_UF - verificar se existe
         VALOR_PERC_ICMS_N = RstTemp!PERC_ICMS ' CFOP.P_ICMS_VND_F_UF'duas aliquotas para  o mesmo cfop
         Else ' Pessoa fisica
            VALOR_ICMS_PRODUTO = ((dblTotalItem * RstTemp!ICMS_PJ_F_UF) / 100)
            VALOR_PERC_ICMS_N = RstTemp!ICMS_PJ_F_UF
      End If

      If RstTemp.State = 1 Then _
         RstTemp.Close
   End If

   'HOJE 12/06/2006 22:00
   'FALTA VERIFICAR SE EXISTE DUAS ALIQUOTAS PARA O MESMO CFOP
   'FALTA GRAVAR OS DADOS CORRETAMENTE NA TABELA
   'FALTA VER O LANCE ABAIXO
   
   'Ver depois com o emanoel para que estes campos
   'se for necessarario mesmo, acho que criarei um campo asc de tamanho x
   ' vou appendando os CFOPS que existir separando-os com com um ';"
   'farei uma funcao para tratar os cfops appendando depois
   '   //Testa Cfop para Cabeca!
   '   if PRODUTOS.COD_TRIBUTACAO eq 60 begin
   '      if CIDADE.UF eq DOCUMENT.UF begin
   '         move 5405                               To   CFOP1_D
   '      End
   '      if CIDADE.UF ne DOCUMENT.UF move 6403      To   CFOP1_F
   '   End
   '   if PRODUTOS.COD_TRIBUTACAO ne 60 begin
   '      if CIDADE.UF eq DOCUMENT.UF begin
   '          move CFOP.VND_MERC_D_UF                To   CFOP_D
   '      End
   '      if CIDADE.UF ne DOCUMENT.UF move CFOP.VND_MERC_F_UF;
   '                                                 To   CFOP_F
   '   End

   SITUAÇÃO_TRIBUTARIA_PRODUTO = "" & rstProduto!SITUACAO_TRIBUTARIA

   'If Not isnull(rstProduto!PERCIVA) Then dblPercIVA = rstProduto!PERCIVA

   If VALOR_BASE_ICMS_N = 0 Then _
      VALOR_PERC_ICMS_N = 0
   
   If rstProduto.State = 1 Then _
      rstProduto.Close

   If RstTemp.State = 1 Then _
      RstTemp.Close

   SQL = "Select PEDIDO.pedido_id,pedidoitem.produto_id from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"

   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " Where PEDIDO.PEDIDO_ID = " & txtPedido.Text
   SQL = SQL & " And CODG_PRODuto = '" & Trim(txtProduto.Text) & "'"
   SQL = SQL & " and pedidoitem.status <> 'C' "

   RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not RstTemp.EOF Then
      PEDIDO_ID_N = RstTemp.Fields(0).Value
      PRODUTO_ID_N = RstTemp.Fields(1).Value

      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "UPDATE PEDIDOITEM SET "
      SQL = SQL & " VlrBaseIcms = " & tpMOEDA(VALOR_BASE_ICMS_N)
      SQL = SQL & ", PERCICMS = " & tpMOEDA(VALOR_PERC_ICMS_N)
      SQL = SQL & ", VlrIcms = " & tpMOEDA(VALOR_ICMS_PRODUTO)
      SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(VALOR_BASE_ICMS_NSubst)
      SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(VALOR_PERC_ICMS_SUBST_N)
      SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(VALOR_ICMS_PRODUTOSubst)
      SQL = SQL & ", cfop_id = '" & strCFOP & "'"
      SQL = SQL & ", STRIBUTARIA = '" & SITUAÇÃO_TRIBUTARIA_PRODUTO & "'"

      SQL = SQL & " Where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and produto_id = " & PRODUTO_ID_N

      CONECTA_RETAGUARDA.Execute SQL
   End If
   If RstTemp.State = 1 Then _
      RstTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREPARA_TRIBUTACAO_PRODUTO"
End Sub



