VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportaToscana 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Importação/Atualização Produtos"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ImportaToscana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Revenda"
      TabPicture(0)   =   "ImportaToscana.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblNovos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAtualizados"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPath"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCons"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Produção"
      TabPicture(1)   =   "ImportaToscana.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1(2)"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "lblProc2"
      Tab(1).Control(3)=   "lblNovos2"
      Tab(1).Control(4)=   "lblAtualizados2"
      Tab(1).Control(5)=   "Line1(3)"
      Tab(1).Control(6)=   "txtPath2"
      Tab(1).Control(7)=   "cmdCons2"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdCons 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         Picture         =   "ImportaToscana.frx":5C4A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdCons2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -70680
         Picture         =   "ImportaToscana.frx":664C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtPath2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Text            =   "c:\megasim\txt\TXITENS.txt"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtPath 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Text            =   "c:\megasim\txt\ATPROD.xlsx"
         Top             =   960
         Width           =   3975
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   3
         X1              =   -74880
         X2              =   -70320
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblAtualizados2 
         AutoSize        =   -1  'True
         Caption         =   "Atualizados:"
         Height          =   285
         Left            =   -74760
         TabIndex        =   11
         Top             =   2355
         Width           =   1440
      End
      Begin VB.Label lblNovos2 
         AutoSize        =   -1  'True
         Caption         =   "Novos:"
         Height          =   285
         Left            =   -74760
         TabIndex        =   10
         Top             =   1995
         Width           =   840
      End
      Begin VB.Label lblProc2 
         AutoSize        =   -1  'True
         Caption         =   "Processados:"
         Height          =   285
         Left            =   -74760
         TabIndex        =   9
         Top             =   1635
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   " Arquivo:"
         Height          =   240
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   2
         X1              =   -74880
         X2              =   -70320
         Y1              =   2715
         Y2              =   2715
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   4680
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblAtualizados 
         AutoSize        =   -1  'True
         Caption         =   "Atualizados:"
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   2355
         Width           =   1440
      End
      Begin VB.Label lblNovos 
         AutoSize        =   -1  'True
         Caption         =   "Novos:"
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1995
         Width           =   840
      End
      Begin VB.Label lblProc 
         AutoSize        =   -1  'True
         Caption         =   "Processados:"
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1635
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   " Arquivo:"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   4680
         Y1              =   2715
         Y2              =   2715
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   2355
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            Key             =   "importar"
            ImageIndex      =   5
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
               Picture         =   "ImportaToscana.frx":704E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaToscana.frx":81E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaToscana.frx":9277
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaToscana.frx":A22C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ImportaToscana.frx":B44C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmImportaToscana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Proc_n  As Long
   Dim Novos_n As Long
   Dim At_n    As Long

   Dim xl As New Excel.Application
   Dim xlw As Excel.Workbook

   Dim DESCRICAO, CODIGOBARRA, REFERENCIA, NCM, FAMILIA, SITUACAOTRIBUTARIA
   Dim FORNECEDOR, PESOLIQUIDO, VENDACUSTO, UNIT, QTDCX, PERC, CUSTOCX, ST_A
   Dim FAMILIA_PRODUTO_ID_N, strRegistro

Private Sub cmdCons_Click()
   frmINICIO.Dialogo.DialogTitle = "Selecionar Caminho Arquivo!"
   frmINICIO.Dialogo.Filter = "*.xlsx;*.xls"
   frmINICIO.Dialogo.ShowOpen
   If Trim(frmINICIO.Dialogo.FileName) <> "" Then
      txtPath.Text = Trim(frmINICIO.Dialogo.FileName)
      Else: Exit Sub
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "importar"
         VAI_FERA
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdCons2_Click()
   frmINICIO.Dialogo.DialogTitle = "Selecionar Caminho Arquivo!"
   frmINICIO.Dialogo.Filter = "TXITENS.txt"
   frmINICIO.Dialogo.ShowOpen
   If Trim(frmINICIO.Dialogo.FileName) <> "" Then
      txtPath2.Text = Trim(frmINICIO.Dialogo.FileName)
      Else: Exit Sub
   End If
End Sub

Sub VAI_FERA()
'On Error GoTo ERRO_TRATA

   Proc_n = 0
   Novos_n = 0
   At_n = 0
   lblProc.Caption = ""
   lblNovos.Caption = ""
   lblAtualizados.Caption = ""
   lblProc2.Caption = ""
   lblNovos2.Caption = ""
   lblAtualizados2.Caption = ""
   strRegistro = ""

   If SSTab1.Tab = 0 Then
      'Abrir o arquivo do Excel
      Set xlw = xl.Workbooks.Open(txtPath.Text)
   
      ' definir qual a planilha de trabalho
      xlw.Sheets("Plan1").Select
   
      DESCRICAO = "a"
      Proc_n = 0
   
      While Trim(DESCRICAO) <> ""
         Proc_n = Proc_n + 1
         lblProc.Caption = "Processados = " & Proc_n
         DoEvents

         DESCRICAO = ""
         CODIGOBARRA = ""
         REFERENCIA = ""
         NCM = ""
         FAMILIA = ""
         SITUACAOTRIBUTARIA = ""
         FORNECEDOR = ""
         PESOLIQUIDO = ""
         VENDACUSTO = ""
         UNIT = ""
         QTDCX = ""
         PERC = ""
         CUSTOCX = ""
   
         DESCRICAO = xlw.Application.Cells(Proc_n, 1).Value
         CODIGOBARRA = xlw.Application.Cells(Proc_n, 2).Value
         REFERENCIA = xlw.Application.Cells(Proc_n, 3).Value
         NCM = xlw.Application.Cells(Proc_n, 4).Value
         FAMILIA = xlw.Application.Cells(Proc_n, 5).Value
         SITUACAOTRIBUTARIA = xlw.Application.Cells(Proc_n, 6).Value
         FORNECEDOR = xlw.Application.Cells(Proc_n, 7).Value
         PESOLIQUIDO = xlw.Application.Cells(Proc_n, 8).Value
         VENDACUSTO = xlw.Application.Cells(Proc_n, 9).Value
         UNIT = xlw.Application.Cells(Proc_n, 10).Value
         QTDCX = xlw.Application.Cells(Proc_n, 11).Value
         PERC = xlw.Application.Cells(Proc_n, 12).Value
         CUSTOCX = xlw.Application.Cells(Proc_n, 13).Value
   
         If Trim(DESCRICAO) <> "DESCRICAO" Then
            'gravar produto
            If Trim(VENDACUSTO) <> "" Then
               If Trim(FAMILIA) <> "" Then
                  If Trim(SITUACAOTRIBUTARIA) = "TRIBUTADA INTEGRALMENTE" Then _
                     SITUACAOTRIBUTARIA = "00"
                  If Trim(SITUACAOTRIBUTARIA) = "SUBSTITUICAO" Then _
                     SITUACAOTRIBUTARIA = "01"
                  If Trim(NCM) = "" Then _
                     NCM = "00"
                  If Trim(PESOLIQUIDO) = "" Then _
                     PESOLIQUIDO = "0"

                  TRAZ_ID_FAMILIA_PRODUTO Trim(FAMILIA)
                  GRAVA_PRODUTO
               End If
            End If
         End If
      Wend

      ' Fechar a planilha sem salvar alterações
      ' Para salvar mude False para True
      xlw.Close False

      ' Liberamos a memória
      Set xlw = Nothing
      Set xl = Nothing
   End If
   If SSTab1.Tab = 1 Then
      If Not FSO.FileExists(txtPath2.Text) Then
         MsgBox "Arquivo não encontrado, verifique."
         End
      End If

      Open txtPath2.Text For Input As #1

      Do While Not EOF(1)
         Proc_n = Proc_n + 1
         lblProc2.Caption = "Processados = " & Proc_n
         DoEvents

         DESCRICAO = ""
         CODIGOBARRA = ""
         REFERENCIA = ""
         NCM = ""
         FAMILIA = ""
         SITUACAOTRIBUTARIA = ""
         FORNECEDOR = ""
         PESOLIQUIDO = ""
         VENDACUSTO = ""
         UNIT = ""
         QTDCX = ""
         PERC = ""
         CUSTOCX = ""
         CODG_PRODUTO_A = ""
         VALOR_ITEM_N = 0

         Line Input #1, strRegistro

         CODG_PRODUTO_A = Int(Trim(Mid$(strRegistro, 6, 6)))   'codg_produto

         VENDACUSTO = Trim(Mid$(strRegistro, 12, 6))      'valor
         VALOR_ITEM_N = VENDACUSTO
         VALOR_ITEM_N = VALOR_ITEM_N / 100
         VENDACUSTO = VALOR_ITEM_N

         DESCRICAO = Trim(Mid$(strRegistro, 21, 25))      'descrição

         If Trim(DESCRICAO) <> "" Then
            SITUACAOTRIBUTARIA = "00"
            NCM = "00"
            PESOLIQUIDO = "0"

            TRAZ_ID_FAMILIA_PRODUTO Trim("NovoProdutos")
            GRAVA_PRODUTO2 CODG_PRODUTO_A
         End If
      Loop
      Close #1
   End If

   MsgBox "Ok"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VAI_FERA"
End Sub

Sub TRAZ_ID_FAMILIA_PRODUTO(DESCRICAO_FAMILIA As String)
'On Error GoTo ERRO_TRATA

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

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where descricao = '" & Trim(DESCRICAO) & "'"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then
      PRODUTO_ID_N = MAX_ID("produto_ID", "produto", "", "", "", "")
      Novos_n = Novos_n + 1
      lblNovos.Caption = "Novos: " & Novos_n

TRAVEIS:
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where codg_produto = '" & PRODUTO_ID_N & "'"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         PRODUTO_ID_N = PRODUTO_ID_N + 1
         GoTo TRAVEIS
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "insert into PRODUTO "
      SQL = SQL & "("
         SQL = SQL & " produto_id,codg_produto,descricao,familiaproduto_id,"
         SQL = SQL & " unidade_medida,situacao,tipo_prod,preco_custo_anterior,"
         SQL = SQL & " preco_custo,preco_atacado,preco_venda,dt_cadastro,"
         SQL = SQL & " preco_varejo_anterior,preco_atacado_anterior,empresa_id,"
         SQL = SQL & " situacao_tributaria,aliquota_icms,"
         SQL = SQL & " codg_barra,referencia,codg_ncm,peso_liquido,peso_bruto"
      SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & PRODUTO_ID_N                     'produto_id
         SQL = SQL & ",'" & PRODUTO_ID_N & "'"        'codg_produto
         SQL = SQL & ",'" & Trim(DESCRICAO) & "'"     'descricao
         SQL = SQL & "," & FAMILIA_PRODUTO_ID_N       'familiaproduto_id
         SQL = SQL & ",'" & Trim("UN") & "'"          'Unidade_Medida
         SQL = SQL & ",'A' "                          'SITUACAO
         SQL = SQL & ",1"                             'Tipo_Prod
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'PRECO_CUSTO_ANTERIOR,
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_custo
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_atacado
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_venda
         SQL = SQL & "," & Now                  'dt_cadastro
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_varejo_anterior
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_atacado_anterior
         SQL = SQL & ",1"                             'empresa_id
         SQL = SQL & ",'" & SITUACAOTRIBUTARIA & "'"  'st
         SQL = SQL & ",17"                            'aliq_icms

         SQL = SQL & ",'" & Trim(CODIGOBARRA) & "'"   'Codg_Barra
         SQL = SQL & ",'" & Trim(REFERENCIA) & "'"    'REFERENCIA
         SQL = SQL & ",'" & Trim(Left(NCM, 8)) & "'"  'codg_ncm
         SQL = SQL & "," & tpMOEDA(PESOLIQUIDO)       'peso_liquido
         SQL = SQL & "," & tpMOEDA(PESOLIQUIDO)       'peso_bruto

      SQL = SQL & " )"

      CONECTA_RETAGUARDA.Execute SQL
      RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N
      Else
         At_n = At_n + 1
         lblAtualizados.Caption = "Atualizados: " & At_n
         PRODUTO_ID_N = TabProduto.Fields(0).Value

         SQL = "update produto set "
            SQL = SQL & " situacao_tributaria = '" & SITUACAOTRIBUTARIA & "'"
            SQL = SQL & ",familiaproduto_id = " & FAMILIA_PRODUTO_ID_N        'familiaproduto_id
            SQL = SQL & ",Unidade_Medida = '" & Trim("UN") & "'"              'Unidade_Medida
            SQL = SQL & ",SITUACAO = 'A' "                                    'SITUACAO
            SQL = SQL & ",Tipo_Prod = 1"                                      'Tipo_Prod
            SQL = SQL & ",PRECO_CUSTO_ANTERIOR = " & tpMOEDA(VENDACUSTO)      'PRECO_CUSTO_ANTERIOR
            SQL = SQL & ",preco_custo = " & tpMOEDA(VENDACUSTO)               'preco_custo
            SQL = SQL & ",preco_atacado = " & tpMOEDA(VENDACUSTO)             'preco_atacado
            SQL = SQL & ",preco_venda = " & tpMOEDA(VENDACUSTO)               'preco_venda
            SQL = SQL & ",preco_varejo_anterior = " & tpMOEDA(VENDACUSTO)     'preco_varejo_anterior
            SQL = SQL & ",preco_atacado_anterior = " & tpMOEDA(VENDACUSTO)    'preco_atacado_anterior
            SQL = SQL & " ,aliquota_icms = 17 "
            SQL = SQL & ",Codg_Barra = '" & Trim(CODIGOBARRA) & "'"           'Codg_Barra
            SQL = SQL & ",REFERENCIA = '" & Trim(REFERENCIA) & "'"            'REFERENCIA
            SQL = SQL & ",codg_ncm = '" & Trim(Left(NCM, 8)) & "'"                    'codg_ncm
            SQL = SQL & ",peso_liquido = " & tpMOEDA(PESOLIQUIDO)        'peso_liquido
            SQL = SQL & ",peso_bruto = " & tpMOEDA(PESOLIQUIDO)         'peso_bruto

         SQL = SQL & " where descricao = '" & Trim(DESCRICAO) & "'"
         CONECTA_RETAGUARDA.Execute SQL
         RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO"
End Sub

Sub GRAVA_PRODUTO2(Codg_Produto As String)
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_produto = '" & Trim(Codg_Produto) & "'"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then
      PRODUTO_ID_N = MAX_ID("produto_ID", "produto", "", "", "", "")
      Novos_n = Novos_n + 1
      lblNovos2.Caption = "Novos: " & Novos_n

      SQL = "insert into PRODUTO "
      SQL = SQL & "("
         SQL = SQL & " produto_id,codg_produto,descricao,familiaproduto_id,"
         SQL = SQL & " unidade_medida,situacao,tipo_prod,preco_custo_anterior,"
         SQL = SQL & " preco_custo,preco_atacado,preco_venda,dt_cadastro,"
         SQL = SQL & " preco_varejo_anterior,preco_atacado_anterior,empresa_id,"
         SQL = SQL & " situacao_tributaria,aliquota_icms,produto_balanca,"
         SQL = SQL & " codg_barra,referencia,codg_ncm,peso_liquido,peso_bruto"
      SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & PRODUTO_ID_N                     'produto_id
         SQL = SQL & ",'" & Trim(Codg_Produto) & "'"     'codg_produto
         SQL = SQL & ",'" & Trim(DESCRICAO) & "'"     'descricao
         SQL = SQL & "," & FAMILIA_PRODUTO_ID_N       'familiaproduto_id
         SQL = SQL & ",'" & Trim("UN") & "'"          'Unidade_Medida
         SQL = SQL & ",'A' "                          'SITUACAO
         SQL = SQL & ",1"                             'Tipo_Prod
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'PRECO_CUSTO_ANTERIOR,
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_custo
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_atacado
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_venda
         SQL = SQL & "," & Now                  'dt_cadastro
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_varejo_anterior
         SQL = SQL & "," & tpMOEDA(VENDACUSTO)        'preco_atacado_anterior
         SQL = SQL & ",1"                             'empresa_id
         SQL = SQL & ",'" & SITUACAOTRIBUTARIA & "'"  'st
         SQL = SQL & ",17"                            'aliq_icms
SQL = SQL & ", 'TRUE'"                       'produto_balanca
         SQL = SQL & ",'" & Trim(CODIGOBARRA) & "'"   'Codg_Barra
         SQL = SQL & ",'" & Trim(REFERENCIA) & "'"    'REFERENCIA
         SQL = SQL & ",'" & Trim(Left(NCM, 8)) & "'"  'codg_ncm
         SQL = SQL & "," & tpMOEDA(PESOLIQUIDO)       'peso_liquido
         SQL = SQL & "," & tpMOEDA(PESOLIQUIDO)       'peso_bruto

      SQL = SQL & " )"

      CONECTA_RETAGUARDA.Execute SQL
      RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N
      Else
         At_n = At_n + 1
         lblAtualizados2.Caption = "Atualizados: " & At_n
         PRODUTO_ID_N = TabProduto.Fields(0).Value

         SQL = "update produto set "
            SQL = SQL & " preco_atacado = " & tpMOEDA(VENDACUSTO)       'preco_atacado
            SQL = SQL & ",preco_venda = " & tpMOEDA(VENDACUSTO)        'preco_venda
            SQL = SQL & ",produto_balanca = 'TRUE'"
         SQL = SQL & " where codg_produto = '" & Trim(Codg_Produto) & "'"
         CONECTA_RETAGUARDA.Execute SQL
         RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO2"
End Sub
