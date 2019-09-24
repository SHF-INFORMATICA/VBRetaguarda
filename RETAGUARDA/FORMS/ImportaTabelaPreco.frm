VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportaTabelaPreco 
   Caption         =   "Atualização de Preço Produtos"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ImportaTabelaPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMata 
      Caption         =   "Excluir Produtos não Utilizados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox txtPerc 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   3480
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "16,00"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdINFOMAIS 
      Caption         =   "Importa tabela preço INFORMAIS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton cmdTC 
      Caption         =   "Importa tabela preço TC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaTabelaPreco.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaTabelaPreco.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaTabelaPreco.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaTabelaPreco.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaTabelaPreco.frx":A010
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "%Lucro"
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label lblDesc 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   4440
   End
   Begin VB.Label lblAtualizados 
      AutoSize        =   -1  'True
      Caption         =   "Atualizados:"
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label lblNovos 
      AutoSize        =   -1  'True
      Caption         =   "Novos:"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label lblProc 
      AutoSize        =   -1  'True
      Caption         =   "Processados:"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   4560
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fornecedor"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "frmImportaTabelaPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim FAMILIA_PRODUTO_ID_N   As Integer

   Dim Proc_n  As Long
   Dim Novos_n As Long
   Dim At_n    As Long
   Dim oConn   As ADODB.Connection
Dim PERC_N As Double
   Dim xl As New Excel.Application
   Dim xlw As Excel.Workbook

   Dim DESCRICAO, CODIGOBARRA, REFERENCIA, NCM, FAMILIA, SITUACAOTRIBUTARIA
   Dim FORNECEDOR, PESOLIQUIDO, VENDACUSTO, UNIT, QTDCX, PERC, CUSTOCX, ST_A
   Dim strRegistro, Marca

Private Sub cmdMata_Click()
   Msg = "Confirma exclusão de produtos não utilizados ?"
   PERGUNTA Msg, vbYesNo + 32, "Importa Tabela Preço", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SQL = "delete estoque"
      SQL = SQL & " where PRODUTO_ID not in (select produto_id from INVENTARIO)"
      SQL = SQL & " and PRODUTO_ID not in (select produto_id from pedidoitem)"
      SQL = SQL & " and PRODUTO_ID not in (select produto_id from notaentradaitem)"
      SQL = SQL & " and PRODUTO_ID not in (select produto_id from OSPECA)"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " Delete Produto"
      SQL = SQL & " where PRODUTO_ID not in (select produto_id from INVENTARIO)"
      SQL = SQL & " and PRODUTO_ID not in (select produto_id from pedidoitem)"
      SQL = SQL & " and PRODUTO_ID not in (select produto_id from notaentradaitem)"
      SQL = SQL & " and PRODUTO_ID not in (select produto_id from OSPECA)"
      CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtPerc_LostFocus()
   If Trim(txtPerc.Text) = "" Then _
      txtPerc.Text = 0
End Sub

Private Sub cmdTC_Click()
'On Error GoTo ERRO_TRATA

   Proc_n = 0
   Novos_n = 0
   At_n = 0
   lblProc.Caption = ""
   lblNovos.Caption = ""
   lblAtualizados.Caption = ""
   lblDesc.Caption = ""

   'Abrir o arquivo do Excel
   Set xlw = xl.Workbooks.Open("c:\megasim\txt\tabelapreco\tc.xls")

   ' definir qual a planilha de trabalho
   xlw.Sheets("TABELA").Select

   If TabTemp.State = 1 Then _
      TabTemp.Close

   FORNEC_ID_N = 0

   SQL = "select * from vwFornecedor "
   SQL = SQL & " where cnpjcpf = '03346837000185'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      MsgBox "Fornecedor não cadastrado. 03346837000185"
      Exit Sub
      Else: FORNEC_ID_N = TabTemp.Fields(0).Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
'=================================================
   DESCRICAO = "a"
   Proc_n = 0
   CONT_N = 0
   REFERENCIA = ""

   While Trim(DESCRICAO) <> ""
      Proc_n = Proc_n + 1
      lblProc.Caption = "Processados = " & Proc_n
      DoEvents

      CODG_PRODUTO_A = Trim(xlw.Application.Cells(Proc_n, 2).Value)

      If Trim(CODG_PRODUTO_A) = "ESTOQUE" Then _
         CONT_N = Proc_n

      If CONT_N > 0 And Trim(CODG_PRODUTO_A) <> "ESTOQUE" Then
         If Trim(CODG_PRODUTO_A) <> "D" Then
            If Trim(CODG_PRODUTO_A) <> "ND" Then
               FAMILIA = xlw.Application.Cells(Proc_n, 2).Value
               TRAZ_ID_FAMILIA_PRODUTO Trim(FAMILIA)
            End If
         End If
         If Trim(CODG_PRODUTO_A) = "D" Or Trim(CODG_PRODUTO_A) = "ND" Then
            DESCRICAO = xlw.Application.Cells(Proc_n, 3).Value
            VENDACUSTO = xlw.Application.Cells(Proc_n, 4).Value

            If IsNumeric(VENDACUSTO) Then _
               GRAVA_PRODUTO_TC
         End If
         If Trim(CODG_PRODUTO_A) = "" Then _
            DESCRICAO = ""
      End If
   Wend
   ' Fechar a planilha sem salvar alterações
   ' Para salvar mude False para True
   xlw.Close False

   ' Liberamos a memória
   Set xlw = Nothing
   Set xl = Nothing
 
MsgBox "ok"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdTC_Click"
End Sub

Private Sub cmdINFOMAIS_Click()
'On Error GoTo ERRO_TRATA

   Proc_n = 0
   Novos_n = 0
   At_n = 0
   lblProc.Caption = ""
   lblNovos.Caption = ""
   lblAtualizados.Caption = ""
   lblDesc.Caption = ""

   'Abrir o arquivo do Excel
   Set xlw = xl.Workbooks.Open("c:\megasim\txt\tabelapreco\informais.xls")

   ' definir qual a planilha de trabalho
   xlw.Sheets("TABELA").Select

   If TabTemp.State = 1 Then _
      TabTemp.Close

   FORNEC_ID_N = 0

   SQL = "select * from vwFornecedor "
   SQL = SQL & " where cnpjcpf = '10610854000143'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      MsgBox "Fornecedor não cadastrado. 10610854000143"
      Exit Sub
      Else: FORNEC_ID_N = TabTemp.Fields(0).Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
'=================================================
   DESCRICAO = "a"
   Proc_n = 0

   While Trim(DESCRICAO) <> ""
      Proc_n = Proc_n + 1
      lblProc.Caption = "Processados = " & Proc_n
      DoEvents

      CODG_PRODUTO_A = Trim(xlw.Application.Cells(Proc_n, 1).Value)
      FAMILIA = xlw.Application.Cells(Proc_n, 3).Value
      VENDACUSTO = xlw.Application.Cells(Proc_n, 6).Value
      REFERENCIA = xlw.Application.Cells(Proc_n, 1).Value
      Marca = xlw.Application.Cells(Proc_n, 5).Value

      If Trim(UCase(CODG_PRODUTO_A)) = UCase("CÓDIGO") Or Trim(UCase(CODG_PRODUTO_A)) = UCase("CODIGO") Then
         DESCRICAO = xlw.Application.Cells(Proc_n, 4).Value
         Else
            If Proc_n >= 18 Then _
               DESCRICAO = xlw.Application.Cells(Proc_n, 4).Value
      End If

      If IsNumeric(CODG_PRODUTO_A) And IsNumeric(VENDACUSTO) Then
         TRAZ_ID_FAMILIA_PRODUTO Trim(FAMILIA)
         GRAVA_PRODUTO
      End If
   Wend
   ' Fechar a planilha sem salvar alterações
   ' Para salvar mude False para True
   xlw.Close False

   ' Liberamos a memória
   Set xlw = Nothing
   Set xl = Nothing
'=================================================
 
MsgBox "ok"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdINFOMAIS_Click"
End Sub

Sub TRAZ_ID_FAMILIA_PRODUTO(DESCRICAO_FAMILIA As String)
'On Error GoTo ERRO_TRATA

   DESCRICAO_FAMILIA = Left(Trim(Replace(DESCRICAO_FAMILIA, "'", "´")), 60)

   FAMILIA_PRODUTO_ID_N = 0
   If Trim(DESCRICAO_FAMILIA) <> "" Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select familiaproduto_id from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where descricao = '" & Trim(DESCRICAO_FAMILIA) & "'"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProduto.EOF Then
         FAMILIA_PRODUTO_ID_N = MAX_ID("familiaproduto_id", "familiaproduto", "", "", "", "")

         SQL = "insert into FAMILIAPRODUTO "
            SQL = SQL & "(FAMILIAPRODUTO_ID,CODG_FAMILIA,DESCRICAO,UNIDADE_MEDIDA,DESC_UNIDADE_MEDIDA,PRODUCAO)"
         SQL = SQL & " values("
            SQL = SQL & FAMILIA_PRODUTO_ID_N                   'FAMILIAPRODUTO_ID
            SQL = SQL & ",'" & FAMILIA_PRODUTO_ID_N & "'"      'CODG_FAMILIA
            SQL = SQL & ",'" & Trim(DESCRICAO_FAMILIA) & "'"   'DESCRICAO
            SQL = SQL & ",'" & Trim("UN") & "'"                'UNIDADE_MEDIDA
            SQL = SQL & ",'" & Trim("UNIDADE") & "'"           'DESC_UNIDADE_MEDIDA
            SQL = SQL & ",0"                                   'PRODUCAO
         SQL = SQL & " )"
         CONECTA_RETAGUARDA.Execute SQL
         Else: FAMILIA_PRODUTO_ID_N = TabProduto.Fields(0).Value
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRAZ_ID_FAMILIA_PRODUTO"
End Sub

Sub GRAVA_PRODUTO()
'On Error GoTo ERRO_TRATA

   DESCRICAO = "" & Left(Trim(Replace(DESCRICAO, "'", "´")), 200)
   PERC_N = 0 & txtPerc.Text

   If TabTemp.State = 1 Then _
      TabTemp.Close
      
   SQL = "select produto_id,preco_custo,preco_atacado,preco_venda from PRODUTO "
   SQL = SQL & " where descricao = '" & Trim(DESCRICAO) & "'"
   'SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      PRODUTO_ID_N = MAX_ID("produto_id", "produto", "", "", "", "")

      SQL = "insert into PRODUTO "
      SQL = SQL & " (PRODUTO_ID,EMPRESA_ID,CODG_PRODUTO,FORNECEDOR_ID,Descricao,"
      SQL = SQL & " FAMILIAPRODUTO_ID,Unidade_Medida,SITUACAO,SITUACAO_TRIBUTARIA,"
      SQL = SQL & " Aliquota_Icms,Tipo_Prod,CODG_NCM,PRECO_CUSTO,PRECO_ATACADO,PRECO_VENDA,referencia)"
      SQL = SQL & " values("
         SQL = SQL & PRODUTO_ID_N                        'PRODUTO_ID
         SQL = SQL & "," & EMPRESA_ID_N                  'EMPRESA_ID
         SQL = SQL & ",'" & PRODUTO_ID_N & "'"           'CODG_PRODUTO
         SQL = SQL & "," & FORNEC_ID_N                   'FORNECEDOR_ID
         SQL = SQL & ",'" & Trim(DESCRICAO) & "'"   'Descricao
         SQL = SQL & "," & FAMILIA_PRODUTO_ID_N          'FAMILIAPRODUTO_ID
         SQL = SQL & ",'UN'"                             'Unidade_Medida
         SQL = SQL & ",'A'"                              'SITUACAO
         SQL = SQL & ",'00'"                             'SITUACAO_TRIBUTARIA
         SQL = SQL & ",17"                               'ALIQUOTA_ICMS_NORMAL_DENTRO_UF
         SQL = SQL & ",1"                                'Tipo_Prod
         SQL = SQL & ",00"                               'CODG_NCM
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)           'PRECO_CUSTO
         SQL = SQL & "," & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)       'PRECO_atacado
         SQL = SQL & "," & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)       'PRECO_venda
         SQL = SQL & ",'" & Trim(REFERENCIA) & "'"   'referencia
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      Novos_n = Novos_n + 1
      lblNovos.Caption = "Novos = " & Novos_n
      Else
         PRODUTO_ID_N = TabTemp.Fields("produto_id").Value

         SQL = "update PRODUTO set "
         SQL = SQL & " preco_custo_anterior = preco_custo"
         SQL = SQL & ",preco_custo = " & tpMOEDA(VENDACUSTO)
         SQL = SQL & ",preco_atacado = " & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)
         SQL = SQL & ",preco_venda = " & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)
         SQL = SQL & ", referencia = '" & Trim(REFERENCIA) & "'"

         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         'SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         At_n = At_n + 1
         lblAtualizados.Caption = "Atualizados = " & At_n
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO"
End Sub

Sub GRAVA_PRODUTO_TC()
'On Error GoTo ERRO_TRATA

   DESCRICAO = "" & Left(Trim(Replace(DESCRICAO, "'", "´")), 200)
   PERC_N = 0 & txtPerc.Text

   If TabTemp.State = 1 Then _
      TabTemp.Close
      
   SQL = "select produto_id,preco_custo,preco_atacado,preco_venda,fornecedor_id from PRODUTO "
   SQL = SQL & " where descricao = '" & Trim(DESCRICAO) & "'"
   'SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      PRODUTO_ID_N = MAX_ID("produto_id", "produto", "", "", "", "")

      SQL = "insert into PRODUTO "
      SQL = SQL & " (PRODUTO_ID,EMPRESA_ID,CODG_PRODUTO,FORNECEDOR_ID,Descricao,"
      SQL = SQL & " FAMILIAPRODUTO_ID,Unidade_Medida,SITUACAO,SITUACAO_TRIBUTARIA,"
      SQL = SQL & " Aliquota_Icms,Tipo_Prod,CODG_NCM,PRECO_CUSTO,PRECO_ATACADO,PRECO_VENDA,referencia)"
      SQL = SQL & " values("
         SQL = SQL & PRODUTO_ID_N                        'PRODUTO_ID
         SQL = SQL & "," & EMPRESA_ID_N                  'EMPRESA_ID
         SQL = SQL & ",'" & PRODUTO_ID_N & "'"           'CODG_PRODUTO
         SQL = SQL & "," & FORNEC_ID_N                   'FORNECEDOR_ID
         SQL = SQL & ",'" & Trim(DESCRICAO) & "'"   'Descricao
         SQL = SQL & "," & FAMILIA_PRODUTO_ID_N          'FAMILIAPRODUTO_ID
         SQL = SQL & ",'UN'"                             'Unidade_Medida
         SQL = SQL & ",'A'"                              'SITUACAO
         SQL = SQL & ",'00'"                             'SITUACAO_TRIBUTARIA
         SQL = SQL & ",17"                               'ALIQUOTA_ICMS_NORMAL_DENTRO_UF
         SQL = SQL & ",1"                                'Tipo_Prod
         SQL = SQL & ",00"                               'CODG_NCM
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)           'PRECO_CUSTO
         SQL = SQL & "," & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)       'PRECO_atacado
         SQL = SQL & "," & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)       'PRECO_venda
         SQL = SQL & ",'" & Trim(REFERENCIA) & "'"   'referencia
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      Novos_n = Novos_n + 1
      lblNovos.Caption = "Novos = " & Novos_n
      Else
         PRODUTO_ID_N = TabTemp.Fields("produto_id").Value

         SQL = "update PRODUTO set "
         SQL = SQL & " preco_custo_anterior = preco_custo"
         SQL = SQL & ",preco_custo = " & tpMOEDA(VENDACUSTO)
         SQL = SQL & ",preco_atacado = " & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)
         SQL = SQL & ",preco_venda = " & tpMOEDA(VENDACUSTO * PERC_N / 100 + VENDACUSTO)
         SQL = SQL & ",referencia = '" & Trim(REFERENCIA) & "'"

         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         'SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         At_n = At_n + 1
         lblAtualizados.Caption = "Atualizados = " & At_n
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO_TC"
End Sub
