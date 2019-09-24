VERSION 5.00
Begin VB.Form frmATUALIZACAO2 
   Caption         =   "Manutenção Base Dados"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   Icon            =   "ATUALIZACAO2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrimoCliente 
      Caption         =   "PRIMO Cliente"
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
      Left            =   4800
      TabIndex        =   42
      Top             =   8280
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrimoProduto 
      Caption         =   "PRIMO Produto"
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
      Left            =   4800
      TabIndex        =   41
      Top             =   7680
      Width           =   2415
   End
   Begin VB.CommandButton cmdENTREGA 
      Caption         =   "Endereço Entrega"
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
      Left            =   2640
      TabIndex        =   40
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmd400 
      Caption         =   "Excluir _400 \NFE\nfe\wsdl\Homologacao\GO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   39
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdEmpresa 
      Caption         =   "Tabela Empresa"
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
      TabIndex        =   38
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "MENU/PERMISSÃO"
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
      Left            =   7680
      TabIndex        =   37
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "TURNO/PRODUÇÃO"
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
      Left            =   7680
      TabIndex        =   36
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "IMPORTA NCM"
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
      Left            =   7680
      TabIndex        =   35
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCupom 
      Caption         =   "TABELA CUPOM"
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
      Left            =   7680
      TabIndex        =   34
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TABELAS/CAMPOS/GLOBAL"
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
      Left            =   1800
      TabIndex        =   33
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MFA010 e MFT010 CODIGO"
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
      Left            =   1800
      TabIndex        =   32
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tabela Cartões GLOBAL"
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
      Left            =   7680
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EstoqueFornecedor"
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
      Left            =   2640
      TabIndex        =   30
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCompra 
      Caption         =   "Pedido de Compra"
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
      Left            =   2640
      TabIndex        =   29
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCEP 
      Caption         =   "TABELA CEP"
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
      TabIndex        =   28
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TABPREÇO PEDIDO"
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
      Left            =   7680
      TabIndex        =   27
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdTributacao 
      Caption         =   "TRIBUTAÇÃO"
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
      Left            =   2640
      TabIndex        =   26
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdCliente 
      Caption         =   "Tabela Cliente"
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
      Left            =   2640
      TabIndex        =   25
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdPessoa 
      Caption         =   "Tabela Pessoa"
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
      TabIndex        =   24
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdOS 
      Caption         =   "Ordem de Serviço"
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
      Left            =   7680
      TabIndex        =   23
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Caption         =   "ESTOQUE"
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
      Left            =   5160
      TabIndex        =   22
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdInventario 
      Caption         =   "INVENTÁRIO"
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
      Left            =   7680
      TabIndex        =   21
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdTabelaPreco 
      Caption         =   "Tabela Preço"
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
      Left            =   7680
      TabIndex        =   20
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdUsuario 
      Caption         =   "Tabela Usuario"
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
      TabIndex        =   19
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdProduto 
      Caption         =   "Tabela Produto"
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
      Left            =   5160
      TabIndex        =   18
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdFamilia 
      Caption         =   "FAMILIA PRODUTO"
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
      Left            =   5160
      TabIndex        =   17
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdTabela 
      Caption         =   "Excluir Tabelas"
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
      TabIndex        =   16
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdVendedor 
      Caption         =   "Tabela Equipe/Vendedor"
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
      TabIndex        =   15
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdFatura 
      Caption         =   "TIPOVENDA/PAGTO"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdDescritor 
      Caption         =   "DESCRITORES"
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
      Left            =   7680
      TabIndex        =   13
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdNF 
      Caption         =   "Tabela NF"
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
      Left            =   5160
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdRG 
      Caption         =   "Tabela RG"
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
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdOBS 
      Caption         =   "Tabela OBS"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdFinanc 
      Caption         =   "FINANCEIRO"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdPedido 
      Caption         =   "PEDIDO E PEDIDOITEM"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdTransp 
      Caption         =   "Tabela Transportadora"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdCFOP 
      Caption         =   "CFOP"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdNFEntrada 
      Caption         =   "NotaEntrada"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdSP 
      Caption         =   "stored procedure e VW"
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
      Left            =   7680
      TabIndex        =   4
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdCaixa 
      Caption         =   "Tabela Caixa"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdEndereco 
      Caption         =   "Tabela Endereço"
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
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdIE_IM 
      Caption         =   "Tabelas IE EM"
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
      Left            =   5160
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton cmdFornecedor 
      Caption         =   "Tabela Fornecedor"
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
      TabIndex        =   1
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10080
      Y1              =   5520
      Y2              =   5520
   End
End
Attribute VB_Name = "frmATUALIZACAO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd400_Click()
'On Error GoTo ERRO_TRATA

   'EXCLUIR_400

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmd400_Click"
End Sub

Private Sub Command3_Click()

   Dim strRegistro
   Dim CODG_NCM_A As String
   Dim DESC_NCM_A As String

   CONT_N = 0

   Open "C:\MEGASIM\RETAGUARDA\txt\NCMIMPORT.TXT" For Input As #1
    Do While Not EOF(1)
       DoEvents
       Line Input #1, strRegistro

'MsgBox strRegistro
'MsgBox Mid(strRegistro, Len(strRegistro) - 9, 9)
'MsgBox Mid(strRegistro, 1, Len(strRegistro) - 10)

      CODG_NCM_A = "" & Mid(strRegistro, Len(strRegistro) - 9, 9)
      DESC_NCM_A = "" & Mid(strRegistro, 1, Len(strRegistro) - 10)

      DESC_NCM_A = Replace(DESC_NCM_A, ",", ".")
      DESC_NCM_A = Replace(DESC_NCM_A, "'", " ")

      If Right(CODG_NCM_A, 1) = ";" Then _
         CODG_NCM_A = Left(CODG_NCM_A, 8)

      Command3.Caption = "" & CODG_NCM_A

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TABNCM "
      SQL = SQL & " where descricao = '" & Trim(DESC_NCM_A) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         spNFePesquisaTipos.Text = "" & Mid(strRegistro, 1, Len(strRegistro) - 9)

         SQL = "insert into TABNCM "
            SQL = SQL & "(CODG_NCM,DESCRICAO)"
         SQL = SQL & " values("
            SQL = SQL & "'" & Trim(CODG_NCM_A) & "'"
            SQL = SQL & ",'" & Trim(DESC_NCM_A) & "'"
         SQL = SQL & ")"

         CONECTA_RETAGUARDA.Execute SQL
         CONT_N = CONT_N + 1
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      DoEvents
    Loop
    Close #1    ' Close file.
    
    MsgBox "Foram Importador = " & CONT_N

End Sub

Private Sub cmdCupom_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CUPOM", "U") = True Then
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CUPOM_IMPRESSORA", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM DROP CONSTRAINT FK_CUPOM_IMPRESSORA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_ENDERECO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO DROP CONSTRAINT IX_ENDERECO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MODELO_DOC", "CUPOM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM ADD MODELO_DOC NVARCHAR(3)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VALOR_CUPOM", "CUPOM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM DROP COLUMN VALOR_CUPOM"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IMPRESSORA_ID", "CUPOM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM DROP COLUMN IMPRESSORA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTA_REINICIO", "CUPOM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM DROP COLUMN CONTA_REINICIO"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_CUPOM", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM ADD CONSTRAINT PK_CUPOM PRIMARY KEY (CUPOM_ID)"
      Else  'NÃO EXISTE A TABELA
    End If
MsgBox "ok    "
End Sub

Private Sub cmdPessoa_Click()
   
   ATUALIZA_TABELA_PESSOA
   
   MsgBox "Ok, aTENÇÃO CRIAR INDICE DO CAMPO CNPJCPF  "
End Sub

Private Sub cmdCLIENTE_Click()
   VERIFICA_TABELA_CLIENTE
   ATUALIZA_CLIENTE_GLOBAL
   MsgBox "Ok, " & CONT_N
End Sub

Private Sub cmdIE_IM_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "IE", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENDERECO_ID", "IE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IE ADD ENDERECO_ID BIGINT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_IE", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IE ADD CONSTRAINT PK_IE PRIMARY KEY (IE_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IE_ENDERECO", "") = False Then
         SQL = "ALTER TABLE [dbo].[IE] WITH CHECK ADD CONSTRAINT [FK_IE_ENDERECO] FOREIGN KEY([ENDERECO_ID])"
         SQL = SQL & " References [dbo].[ENDERECO]([ENDERECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[IE] CHECK CONSTRAINT [FK_IE_ENDERECO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_IE", "") = False Then
         SQL = "ALTER TABLE IE ADD CONSTRAINT IX_IE UNIQUE NONCLUSTERED "
         SQL = SQL & " ([NUMR_IE] ASC,[ENDERECO_ID] ASC,[PESSOA_ID] Asc)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IE_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[IE]  WITH CHECK ADD  CONSTRAINT [FK_IE_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[IE] CHECK CONSTRAINT [FK_IE_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "IM", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENDERECO_ID", "IM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IM ADD ENDERECO_ID BIGINT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_IM", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IM ADD CONSTRAINT PK_IM PRIMARY KEY (IM_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IM_ENDERECO", "") = False Then
         SQL = "ALTER TABLE [dbo].[IM] WITH CHECK ADD CONSTRAINT [FK_IM_ENDERECO] FOREIGN KEY([ENDERECO_ID])"
         SQL = SQL & " References [dbo].[ENDERECO]([ENDERECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[IM] CHECK CONSTRAINT [FK_IM_ENDERECO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_IM", "") = False Then
         SQL = "ALTER TABLE IM ADD CONSTRAINT IX_IM UNIQUE NONCLUSTERED "
         SQL = SQL & " ([NUMR_IM] ASC,[ENDERECO_ID] ASC,[PESSOA_ID] Asc)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IM_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[IM]  WITH CHECK ADD  CONSTRAINT [FK_IM_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[IM] CHECK CONSTRAINT [FK_IM_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

MsgBox "Ok, TABELEA ie im"
End Sub

Private Sub cmdFornecedor_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FORNECEDOR", "U") = True Then
'===========alter=====================================================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "FORNECEDOR") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD PESSOA_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT not null", "FORNECEDOR", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORNECEDOR_ID", "FORNECEDOR") = True Then _
         Alteração_Definição_Campo_Tabela "FORNECEDOR_ID", "BIGINT not null", "FORNECEDOR", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_FORNECEDOR", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD CONSTRAINT pk_FORNECEDOR PRIMARY KEY (FORNECEDOR_ID)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTATO", "FORNECEDOR") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD CONTATO NVARCHAR(30) "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "FORNECEDOR") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD ESTABELECIMENTO_ID INT"
         SQL = "update FORNECEDOR set estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
'===========drop=====================================================
      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_FORNECEDOR_CNPJCPF", "") = True Then
         SQL = "alter table FORNECEDOR drop CONSTRAINT IX_FORNECEDOR_CNPJCPF"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_FORNECEDOR_CNPJ", "") = True Then
         SQL = "alter table FORNECEDOR "
         SQL = SQL & " drop CONSTRAINT IX_FORNECEDOR_CNPJ"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MARKUP_ATACADO", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN MARKUP_ATACADO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MARKUP_VAREJO", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN MARKUP_VAREJO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Aliquotafornec", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN Aliquotafornec"

      ENDERECO_ID_N = 0
      PESSOA_ID_N = 0
'=========================== IE
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IE", "FORNECEDOR") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from fornecedor WITH (NOLOCK)"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value
            ENDERECO_ID_N = 0 & TRAZ_ID_ENDERECO("C")

            If Not IsNull(TabTemp.Fields("ie").Value) Then _
               If IsNumeric(TabTemp.Fields("ie").Value) Then _
                  GRAVA_IE TabTemp.Fields("ie").Value

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN IE"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORNECEDOR_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[FORNECEDOR]  WITH CHECK ADD  CONSTRAINT [FK_FORNECEDOR_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[FORNECEDOR] CHECK CONSTRAINT [FK_FORNECEDOR_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORNECEDOR_EMPRESA", "") = True Then
         SQL = "alter table FORNECEDOR "
         SQL = SQL & " drop CONSTRAINT FK_FORNECEDOR_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN EMPRESA_ID"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORNECEDOR_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[FORNECEDOR]  WITH CHECK ADD  CONSTRAINT [FK_FORNECEDOR_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[FORNECEDOR] CHECK CONSTRAINT [FK_FORNECEDOR_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cgccpf", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN cgccpf"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "nome", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN nome"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "razao_social", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN razao_social"
'===================================

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FORNECEDORCOMPRADOR", "U") = False Then
         SQL = "CREATE TABLE [dbo].[FORNECEDORCOMPRADOR]("
         SQL = SQL & " [USUARIO_ID] [BIGint] NOT NULL,"
         SQL = SQL & " [FORNECEDOR_ID] [bigint] NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_FORNECEDORCOMPRADOR] PRIMARY KEY CLUSTERED"
         SQL = SQL & " ([USUARIO_ID] ASC,[FORNECEDOR_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
         SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = " ALTER TABLE [dbo].[FORNECEDORCOMPRADOR]  WITH CHECK ADD  CONSTRAINT [FK_FORNECEDORCOMPRADOR_FORNECEDOR] FOREIGN KEY([FORNECEDOR_ID])"
         SQL = SQL & " References [dbo].[FORNECEDOR]([FORNECEDOR_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = " ALTER TABLE [dbo].[FORNECEDORCOMPRADOR] CHECK CONSTRAINT [FK_FORNECEDORCOMPRADOR_FORNECEDOR]"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = " ALTER TABLE [dbo].[FORNECEDORCOMPRADOR]  WITH CHECK ADD  CONSTRAINT [FK_FORNECEDORCOMPRADOR_USUARIO] FOREIGN KEY([USUARIO_ID])"
         SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = " ALTER TABLE [dbo].[FORNECEDORCOMPRADOR] CHECK CONSTRAINT [FK_FORNECEDORCOMPRADOR_USUARIO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   MsgBox "Ok, TABELEA FORNECEDOR "
End Sub

Private Sub cmdCep_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CEP", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cep", "cep") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'cep.Cep'" & "," & "'CEP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CIDADE", "cep") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'cep.CIDADE'" & "," & "'CIDADE'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO_IBGE", "cep") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'cep.CODIGO_IBGE'" & "," & "'IBGE_ID'" & "," & "'COLUMN'"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_PROCURA_CEP", "") = True Then _
         CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_PROCURA_CEP"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_PROCURA_CEP", "") = True Then
         SQL = "CREATE PROCEDURE [dbo].[SP_PROCURA_CEP] (@CEP_ID nvarchar(8)) as SET NOCOUNT on"
         SQL = SQL & " select * from CEP where cep_id  = @CEP_ID"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CEP", "") = False Then

         If TabTemp.State = 1 Then _
            TabTemp.Close
'set
         SQL = "update CEP set "
         SQL = SQL & " ibge_id = "

         If TabTemp.State = 1 Then _
            TabTemp.Close

         'CONECTA_RETAGUARDA.Execute "ALTER TABLE CEP ADD CONSTRAINT pk_CEP PRIMARY KEY (CEP_ID)"
      End If
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "IBGE", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO_IBGE", "IBGE") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'IBGE.CODIGO_IBGE'" & "," & "'IBGE_ID'" & "," & "'COLUMN'"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_IBGE", "") = False Then
         SQL = "delete from IBGE where ibge_id = 5300108"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "insert into IBGE values("
            SQL = SQL & 5300108
            SQL = SQL & ",'Brasília'"
            SQL = SQL & ",'DF'"
         SQL = SQL & " )"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE IBGE ADD CONSTRAINT pk_IBGE PRIMARY KEY (IBGE_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   MsgBox "OK"
End Sub

Private Sub cmdEndereco_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ENDERECO", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "ENDERECO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO DROP COLUMN EMPRESA_ID"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_ENDERECO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO DROP CONSTRAINT IX_ENDERECO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ie_ID", "ENDERECO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO DROP COLUMN ie_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "ENDERECO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO ADD PESSOA_ID BIGINT "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENDERECO_ID", "ENDERECO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO ADD ENDERECO_ID BIGINT "
         Else: Alteração_Definição_Campo_Tabela "ENDERECO_ID", "BIGINT NOT NULL", "ENDERECO", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMERO", "ENDERECO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO ADD NUMERO nvarchar(50) "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cep", "ENDERECO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'ENDERECO.CEP'" & "," & "'CEP_ID'" & "," & "'COLUMN'"

      'CRIA PROCEDURE ALIMENTAR SEQUENCIA CAMPO BANCO_ID
      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_ENDERECO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_UPDATE_ID_ENDERECO"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_ENDERECO", "") = False Then
         SQL = "CREATE PROCEDURE SP_UPDATE_ID_ENDERECO "
         SQL = SQL & " as "
         SQL = SQL & " DECLARE @Contador AS SMALLINT"
         SQL = SQL & " SET @Contador = 0"
         SQL = SQL & " Update ENDERECO "
         SQL = SQL & " SET @Contador = @Contador + 1"
         SQL = SQL & " , ENDERECO_ID = @Contador"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      'CONECTA_RETAGUARDA.Execute "EXEC SP_UPDATE_ID_ENDERECO "

      Alteração_Definição_Campo_Tabela "ENDERECO_ID", "BIGINT NOT NULL", "ENDERECO", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_ENDERECO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO ADD CONSTRAINT pk_ENDERECO PRIMARY KEY (ENDERECO_ID)"

'==================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "ENDERECO") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CRITERIO_A = cmdEndereco.Caption

         SQL = "select distinct(prop) from ENDERECO where pessoa_id is null "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(TabTemp.Fields("prop").Value) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update ENDERECO set pessoa_id = " & TabPessoa.Fields(0).Value
               'SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " where prop = '" & Trim(TabTemp.Fields("prop").Value) & "'"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            cmdEndereco.Caption = Trim(TabTemp.Fields("prop").Value)
            DoEvents
   
            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         cmdEndereco.Caption = CRITERIO_A
      End If

      'correção no endereço duplicado
      SQL = "delete endereco where tipo = '0'"
      CONECTA_RETAGUARDA.Execute SQL

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select pessoa_id from PESSOA "
      SQL = SQL & " order by pessoa_id "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select count(pessoa_id) from ENDERECO "
         SQL = SQL & " where pessoa_id = " & TabTemp.Fields("pessoa_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If TabConsulta.Fields(0).Value > 3 Then
               MsgBox "checando registro de endereço tabela endereço, Pessoaid na tabela endereço = " & TabTemp.Fields(0).Value
            End If
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         DoEvents
         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ENDERECO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_UPDATE_ENDERECO"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ENDERECO", "") = False Then
         SQL = "CREATE PROCEDURE [dbo].[SP_UPDATE_ENDERECO] "
         SQL = SQL & " (@pessoa_id Varchar(14),@cep_id nvarchar(8),@rua varchar(50),@bairro Varchar(50),@comp Varchar(50),@tipo char(1),@id int,@numero nvarchar(10)) "
         SQL = SQL & " as SET NOCOUNT on UPDATE ENDERECO  "
         SQL = SQL & " SET pessoa_id = @pessoa_id, cep_id = @cep_id, rua = @rua, bairro = @bairro, complemento = @comp, tipo = @tipo, numero = @numero "
         SQL = SQL & " WHERE pessoa_id = @pessoa_id"
         SQL = SQL & " and tipo = @tipo "
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_ENDERECO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_ENDERECO"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_ENDERECO", "") = False Then
         SQL = "CREATE PROCEDURE [dbo].[SP_INSERT_ENDERECO]  @pessoa_id as numeric , @cep_id NVARCHAR(8), "
         SQL = SQL & " @rua Varchar(50), @bairro Varchar(50), @complemento Varchar(50), @TIPO char(1), "
         SQL = SQL & " @id as numeric, @NUMERO AS NVARCHAR(50)"
         SQL = SQL & " as SET NOCOUNT on INSERT INTO ENDERECO "
         SQL = SQL & " (PESSOA_ID,cep_id, rua, bairro, complemento, TIPO, ENDERECO_ID,NUMERO)"
         SQL = SQL & " VALUES (@pessoa_id, @cep_id, @rua, @bairro, @complemento, @TIPO, @id, @NUMERO)"
         SQL = SQL & " select PESSOA_ID,cep_id, rua, bairro, complemento, TIPO, ENDERECO_ID, NUMERO from ENDERECO"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_ENDERECO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO ADD CONSTRAINT pk_ENDERECO PRIMARY KEY (ENDERECO_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ENDERECO_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENDERECO]  WITH CHECK ADD  CONSTRAINT [FK_ENDERECO_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[ENDERECO] CHECK CONSTRAINT [FK_ENDERECO_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

'so pode excluir quando rodar todas rotinas
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "ENDERECO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENDERECO DROP COLUMN PROP"
   End If

   MsgBox "Ok, TABELEA endereço"
End Sub

Sub cmdCaixa_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CAIXADIA", "U") = True Then
      Alteração_Definição_Campo_Tabela "CAIXADIA_ID", "BIGINT NOT NULL", "CAIXADIA", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CAIXADIA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA ADD CONSTRAINT pk_CAIXADIA PRIMARY KEY (CAIXADIA_ID)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_usu_fecha", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN codg_usu_fecha"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "tipo", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN tipo"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "status", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN status"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "numr_reducao_z", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN numr_reducao_z"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "valor_reducao", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN valor_reducao"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "valor_dolar", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN valor_dolar"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "descricao", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN descricao"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXADIA_EMPRESA", "") = True Then
         SQL = "alter table CAIXADIA "
         SQL = SQL & " drop CONSTRAINT FK_CAIXADIA_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMERO_CAIXA_CPU", "CAIXADIA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA ADD NUMERO_CAIXA_CPU INT"
   
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "CAIXADIA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIA ADD ESTABELECIMENTO_ID INT"
   
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXADIA_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CAIXADIA]  WITH CHECK ADD  CONSTRAINT [FK_CAIXADIA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CAIXADIA] CHECK CONSTRAINT [FK_CAIXADIA_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_USU_ABRE", "CAIXADIA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXADIA.CODG_USU_ABRE'" & "," & "'USUARIO_ID'" & "," & "'COLUMN'"
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXADIA_USUARIO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CAIXADIA]  WITH CHECK ADD  CONSTRAINT [FK_CAIXADIA_USUARIO] FOREIGN KEY([USUARIO_ID])"
         SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CAIXADIA] CHECK CONSTRAINT [FK_CAIXADIA_USUARIO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
'=================CAIXADIAITEM
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CAIXADIAITEM", "U") = True Then
      Alteração_Definição_Campo_Tabela "CAIXADIA_ID", "BIGINT NOT NULL", "CAIXADIAITEM", "RETAGUARDA"
      Alteração_Definição_Campo_Tabela "CAIXADIAITEM_ID", "BIGINT NOT NULL", "CAIXADIAITEM", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIAITEM DROP COLUMN SEQ"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VALOR_INICIAL", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXADIAITEM.VALOR_INICIAL'" & "," & "'VALOR'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_DC", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXADIAITEM.TIPO_DC'" & "," & "'TIPO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "STATUS", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIAITEM DROP COLUMN STATUS"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_DOC", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIAITEM DROP COLUMN NUMR_DOC"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CAIXADIAITEM", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXADIAITEM ADD CONSTRAINT pk_CAIXADIAITEM PRIMARY KEY (CAIXADIA_ID,CAIXADIAITEM_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXADIAITEM_CAIXADIA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CAIXADIAITEM]  WITH CHECK ADD  CONSTRAINT [FK_CAIXADIAITEM_CAIXADIA] FOREIGN KEY([CAIXADIA_ID])"
         SQL = SQL & " References [dbo].[CAIXADIA]([CAIXADIA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CAIXADIAITEM] CHECK CONSTRAINT [FK_CAIXADIAITEM_CAIXADIA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
'=================CAIXATESORARIA
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CAIXATESORARIA", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CAIXA_ID", "CAIXATESORARIA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXATESORARIA.CAIXA_ID'" & "," & "'CAIXATESORARIA_ID'" & "," & "'COLUMN'"

      Alteração_Definição_Campo_Tabela "CAIXATESORARIA_ID", "BIGINT NOT NULL", "CAIXATESORARIA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_usu_fecha", "CAIXATESORARIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIA DROP COLUMN codg_usu_fecha"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "CAIXATESORARIA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIA ADD ESTABELECIMENTO_ID INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMERO_CAIXA_CPU", "CAIXATESORARIA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIA ADD NUMERO_CAIXA_CPU INT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CAIXATESORARIA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIA ADD CONSTRAINT pk_CAIXATESORARIA PRIMARY KEY (CAIXATESORARIA_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXATESORARIA_EMPRESA", "") = True Then
         SQL = "alter table CAIXATESORARIA"
         SQL = SQL & " drop CONSTRAINT FK_CAIXATESORARIA_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "CAIXATESORARIA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIA DROP COLUMN EMPRESA_ID"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXATESORARIA_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CAIXATESORARIA]  WITH CHECK ADD  CONSTRAINT [FK_CAIXATESORARIA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CAIXATESORARIA] CHECK CONSTRAINT [FK_CAIXATESORARIA_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_USU_ABRE", "CAIXATESORARIA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXATESORARIA.CODG_USU_ABRE'" & "," & "'USUARIO_ID'" & "," & "'COLUMN'"
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXATESORARIA_USUARIO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CAIXATESORARIA]  WITH CHECK ADD  CONSTRAINT [FK_CAIXATESORARIA_USUARIO] FOREIGN KEY([USUARIO_ID])"
         SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CAIXATESORARIA] CHECK CONSTRAINT [FK_CAIXATESORARIA_USUARIO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

'=================CAIXATESORARIAITEM
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CAIXATESORARIAITEM", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "empresa_id", "CAIXATESORARIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIAITEM DROP COLUMN empresa_id"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CAIXA_ID", "CAIXATESORARIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXATESORARIAITEM.CAIXA_ID'" & "," & "'CAIXATESORARIA_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "CAIXATESORARIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXATESORARIAITEM.SEQ'" & "," & "'CAIXATESORARIAITEM_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CC_ID", "CAIXATESORARIAITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIAITEM ADD CC_ID INT"

      Alteração_Definição_Campo_Tabela "CAIXATESORARIA_ID", "BIGINT NOT NULL", "CAIXATESORARIAITEM", "RETAGUARDA"
      Alteração_Definição_Campo_Tabela "CAIXATESORARIAITEM_ID", "BIGINT NOT NULL", "CAIXATESORARIAITEM", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CAIXATESORARIAITEM", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CAIXATESORARIAITEM ADD CONSTRAINT pk_CAIXATESORARIAITEM PRIMARY KEY (CAIXATESORARIA_ID,CAIXATESORARIAITEM_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CAIXATESORARIAITEM_CAIXATESORARIA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CAIXATESORARIAITEM]  WITH CHECK ADD  CONSTRAINT [FK_CAIXATESORARIAITEM_CAIXATESORARIA] FOREIGN KEY([CAIXATESORARIA_ID])"
         SQL = SQL & " References [dbo].[CAIXATESORARIA]([CAIXATESORARIA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CAIXATESORARIAITEM] CHECK CONSTRAINT [FK_CAIXATESORARIAITEM_CAIXATESORARIA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   
   MsgBox "Ok"
End Sub

Private Sub cmdSP_Click()
'On Error Resume Next

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRegistroVenda", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRegistroVenda]"

   SQL = " CREATE VIEW [dbo].[vwRegistroVenda] AS"

   SQL = SQL & " select ESTABELECIMENTO.EMPRESA_ID, ESTABELECIMENTO.ESTABELECIMENTO_ID, ESTABELECIMENTO.DESCRICAO, PEDIDO.DT_REQ AS DT_REGISTRO, PEDIDO.STATUS, "
   SQL = SQL & " PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA AS QTDE, PEDIDOITEM.VALOR_ITEM AS VALOR, PRODUTO.CODG_PRODUTO,"
   SQL = SQL & " PRODUTO.DESCRICAO AS DescProduto, PRODUTO.PESO_LIQUIDO, PRODUTO.UNIDADE_MEDIDA AS UN, FAMILIAPRODUTO.PRODUCAO, PRODUTO.PRECO_VENDA, PEDIDO.PEDIDO_ID"
   SQL = SQL & " from PEDIDO INNER JOIN"
   SQL = SQL & " PEDIDOITEM ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID INNER JOIN"
   SQL = SQL & " PRODUTO ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID INNER JOIN"
   SQL = SQL & " ESTABELECIMENTO ON PEDIDO.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID INNER JOIN"
   SQL = SQL & " FAMILIAPRODUTO ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwProducaoPerda", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwProducaoPerda]"

   SQL = " CREATE VIEW [dbo].[vwProducaoPerda] AS"

   SQL = SQL & " select PRODUCAOPERDA.*, ESTABELECIMENTO.EMPRESA_ID, ESTABELECIMENTO.DESCRICAO, "
   SQL = SQL & " Turno.HoraIni, Turno.HoraFim, PRODUTO.CODG_PRODUTO,"
   SQL = SQL & " PRODUTO.DESCRICAO AS DescProduto, PRODUTO.PESO_LIQUIDO, PRODUTO.UNIDADE_MEDIDA, PRODUTO.PRECO_VENDA "
   SQL = SQL & " from PRODUCAOPERDA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON PRODUCAOPERDA.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN Turno "
   SQL = SQL & " ON PRODUCAOPERDA.TURNO_ID = Turno.TURNO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PRODUCAOPERDA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRegistroPerda", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRegistroPerda]"

   SQL = " CREATE VIEW [dbo].[vwRegistroPerda] AS"

   SQL = SQL & " select CONTROLEPERDA.*, CONTROLEPERDAITEM.SEQ_ID, CONTROLEPERDAITEM.PRODUTO_ID, "
   SQL = SQL & " CONTROLEPERDAITEM.QTDE, CONTROLEPERDAITEM.VALOR, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO  AS ProdDesc, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, ESTABELECIMENTO.EMPRESA_ID, ESTABELECIMENTO.DESCRICAO as EstabDesc,"
   SQL = SQL & " PRODUTO.PESO_LIQUIDO, PRODUTO.UNIDADE_MEDIDA AS UN, PRODUTO.PRECO_VENDA "
   SQL = SQL & " from CONTROLEPERDA "
   SQL = SQL & " INNER JOIN CONTROLEPERDAITEM "
   SQL = SQL & " ON CONTROLEPERDA.CONTROLEPERDA_ID = CONTROLEPERDAITEM.CONTROLEPERDA_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON CONTROLEPERDA.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON CONTROLEPERDAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRegistroProducao", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRegistroProducao]"

   SQL = " CREATE VIEW [dbo].[vwRegistroProducao] AS"

   SQL = SQL & " select REGISTROPRODUCAO.*, ESTABELECIMENTO.EMPRESA_ID, ESTABELECIMENTO.DESCRICAO as EstabDesc, "
   SQL = SQL & " REGISTROPRODUCAOITEM.SEQ_ID, REGISTROPRODUCAOITEM.PRODUTO_ID, REGISTROPRODUCAOITEM.QTDE, "
   SQL = SQL & " REGISTROPRODUCAOITEM.VALOR, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO AS ProdDesc,"
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID,PRODUTO.PESO_LIQUIDO, PRODUTO.UNIDADE_MEDIDA AS UN, PRODUTO.PRECO_VENDA "
   SQL = SQL & " from REGISTROPRODUCAO "
   SQL = SQL & " INNER JOIN REGISTROPRODUCAOITEM "
   SQL = SQL & " ON REGISTROPRODUCAO.REGISTROPRODUCAO_ID = REGISTROPRODUCAOITEM.REGISTROPRODUCAO_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON REGISTROPRODUCAO.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON REGISTROPRODUCAOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwProduto", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwProduto]"

   SQL = " CREATE VIEW [dbo].[vwProduto] AS"

   SQL = SQL & " select PRODUTO.PRODUTO_ID, PRODUTO.EMPRESA_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " PRODUTO.UNIDADE_MEDIDA, PRODUTO.CODG_BARRA, PRODUTO.SITUACAO, PRODUTO.SITUACAO_TRIBUTARIA, PRODUTO.ALIQUOTA_ICMS,"
   SQL = SQL & " PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, PRODUTO.COMP_TRIBUTARIA, PRODUTO.FORNECEDOR_ID, PRODUTO.PRECO_CUSTO_ANTERIOR,"
   SQL = SQL & " PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, PRODUTO.DT_CADASTRO, PRODUTO.QTD_MINIMO, PRODUTO.QTD_MAXIMO,"
   SQL = SQL & " PRODUTO.DT_ULT_VENDA, PRODUTO.DT_ULT_COMPRA, PRODUTO.PESO_LIQUIDO, PRODUTO.PESO_BRUTO, PRODUTO.MARCA_ID,PRODUTO.ORIGEM_MERCADO,"
   SQL = SQL & " PRODUTO.PRODUTO_BALANCA, PRODUTO.PERMITE_DESCONTO, PRODUTO.CONCEDER_PRODUCAO, FAMILIAPRODUTO.CODG_FAMILIA,"
   SQL = SQL & " FAMILIAPRODUTO.DESCRICAO AS DescFamilia, FAMILIAPRODUTO.PRODUCAO, PRODUTOFORNECEDOR.CODG_PROD_FORNEC,"
   SQL = SQL & " PRODUTOFORNECEDOR.PRECO_CUSTO AS PrCustoFornec, PRODUTOFORNECEDOR.CODG_BARRA AS BarraFornec, PRODUTO.REFERENCIA"
   SQL = SQL & " from PRODUTO WITH (NOLOCK) "
   SQL = SQL & " LEFT OUTER JOIN FAMILIAPRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " LEFT OUTER JOIN PRODUTOFORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOFORNECEDOR.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL
'==================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwConsTransf", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwConsTransf]"

   SQL = " CREATE VIEW [dbo].[vwConsTransf] AS"

   SQL = SQL & " select PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, ESTOQUETRANSF.ESTAB_ORIGEM_ID, "
   SQL = SQL & " ESTOQUETRANSF.ESTAB_ORIGEM_DESC,ESTOQUETRANSF.ESTAB_DESTINO_DESC,"
   SQL = SQL & " ESTOQUETRANSF.ESTAB_DESTINO_ID, ESTOQUETRANSF.QTDE_TRANSF, "
   SQL = SQL & " Estoque.QTDE_ESTOQUE,ESTOQUETRANSF.SITUACAO , ESTOQUETRANSF.DT_TRANSF,"
   SQL = SQL & " ESTOQUETRANSF.transf_id as Lote, ESTOQUETRANSF.dt_entrada, estoquetransf.produto_id"
   SQL = SQL & ",estabelecimento_id, transf_id"
   SQL = SQL & " from ESTOQUETRANSF WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ESTOQUE WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUETRANSF.PRODUTO_ID = ESTOQUE.PRODUTO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUE.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

'===============vwLeCaixaTesoraria
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwLeCaixaTesoraria", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwLeCaixaTesoraria]"

   SQL = " CREATE VIEW [dbo].[vwLeCaixaTesoraria] AS"

   SQL = SQL & " select CAIXATESORARIA.CAIXATESORARIA_ID, CAIXATESORARIA.DT_ABERTURA, CAIXATESORARIA.DT_FECHAMENTO, "
   SQL = SQL & " CAIXATESORARIA.USUARIO_ID, CAIXATESORARIA.ESTABELECIMENTO_ID, "
   SQL = SQL & " CAIXATESORARIA.NUMERO_CAIXA_CPU, CAIXATESORARIAITEM.CAIXATESORARIAITEM_ID,"
   SQL = SQL & " CAIXATESORARIAITEM.NUMR_DOC, CAIXATESORARIAITEM.FORMAPAGTO_ID, CAIXATESORARIAITEM.VALOR, CAIXATESORARIAITEM.STATUS,"
   SQL = SQL & " CAIXATESORARIAITEM.Origem , CAIXATESORARIAITEM.TIPO, CAIXATESORARIAITEM.HISTORICO, CAIXATESORARIAITEM.CC_ID"
   SQL = SQL & " from CAIXATESORARIA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " ON CAIXATESORARIA.CAIXATESORARIA_ID = CAIXATESORARIAITEM.CAIXATESORARIA_ID"
   CONECTA_RETAGUARDA.Execute SQL

'===============vwVendedor
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwVendedor", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwVendedor]"

   SQL = " CREATE VIEW [dbo].[vwVendedor] AS"
   SQL = SQL & " select VENDEDOR.VENDEDOR_ID, VENDEDOR.PESSOA_ID, VENDEDOR.EQUIPE_ID, "
   SQL = SQL & " VENDEDOR.TABELAPRECO_ID, VENDEDOR.STATUS, VENDEDOR.DT_NASCIMENTO, "
   SQL = SQL & " VENDEDOR.DT_BAIXA, VENDEDOR.TIPO_COMIS, VENDEDOR.CATEGORIA, "
   SQL = SQL & " VENDEDOR.PERC_COMISSAO, PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO, "
   SQL = SQL & " PESSOA.DATA_CAD, PESSOA.SITUACAO, ESTABVENDEDOR.ESTABELECIMENTO_ID, "
   SQL = SQL & " EQUIPE.DESCRICAO AS DescEquipe, EQUIPE.RESPONSAVEL"
   SQL = SQL & " from VENDEDOR WITH (NOLOCK) "
   SQL = SQL & " LEFT OUTER JOIN PESSOA "
   SQL = SQL & " ON VENDEDOR.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN ESTABVENDEDOR "
   SQL = SQL & " ON VENDEDOR.VENDEDOR_ID = ESTABVENDEDOR.VENDEDOR_ID "
   SQL = SQL & " LEFT OUTER JOIN EQUIPE "
   SQL = SQL & " ON VENDEDOR.EQUIPE_ID = EQUIPE.EQUIPE_ID"
   CONECTA_RETAGUARDA.Execute SQL
'===============vwNotaEntrada
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwNotaEntrada", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwNotaEntrada]"

   SQL = " CREATE VIEW [dbo].[vwNotaEntrada] AS"
   SQL = SQL & " select NOTAENTRADA.*, PRODUTO.EMPRESA_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.SITUACAO,"
   SQL = SQL & " PRODUTO.TIPO_PROD, PRODUTO.PRECO_CUSTO AS PrecoCustoProduto, PRODUTO.PRECO_Venda, PRODUTO.PRODUTO_BALANCA, NOTAENTRADAITEM.SEQ_ID,"
   SQL = SQL & " NOTAENTRADAITEM.PRODUTO_ID, NOTAENTRADAITEM.PRECO_CUSTO, NOTAENTRADAITEM.QTDE_ENTRADA, NOTAENTRADAITEM.STATUS AS StatusItem, "
   SQL = SQL & " NOTAENTRADAITEM.CFOP_ID, NOTAENTRADAITEM.PERC_IPI, NOTAENTRADAITEM.PERC_ICMS AS PercIcmsI, NOTAENTRADAITEM.VALOR_DESCONTO AS DescontoItem,"
   SQL = SQL & " NOTAENTRADAITEM.PERC_ICMS_SUB, NOTAENTRADAITEM.PERC_FRETE, NOTAENTRADAITEM.NCM, NOTAENTRADAITEM.CST, NOTAENTRADAITEM.UN,"
   SQL = SQL & " FORNECEDOR.PESSOA_ID, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFornecedor, PESSOA.RAZAO"
   SQL = SQL & " from NOTAENTRADA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
   SQL = SQL & " AND NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"
   CONECTA_RETAGUARDA.Execute SQL
'===============vwFornecedor
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwFornecedor", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwFornecedor]"

   SQL = " CREATE VIEW [dbo].[vwFornecedor] AS"
   SQL = SQL & " select FORNECEDOR.FORNECEDOR_ID, FORNECEDOR.DT_CAD, FORNECEDOR.STATUS, "
   SQL = SQL & " FORNECEDOR.CONTATO, FORNECEDOR.ESTABELECIMENTO_ID, PESSOA.PESSOA_ID, "
   SQL = SQL & " PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO, PESSOA.DATA_CAD, PESSOA.SITUACAO, "
   SQL = SQL & " ENDERECO.ENDERECO_ID, ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
   SQL = SQL & " ENDERECO.TIPO, ENDERECO.NUMERO AS NUMERO_ENDERECO,"
   SQL = SQL & " FONE.NUMERO AS NUMERO_FONE, FONE.DDD, FONE.LOCAL, EMAIL.EMAIL, IE.IE_ID, "
   SQL = SQL & " IE.ENDERECO_ID AS IE_ENDERECO_ID, IM.IM_ID, IM.NUMR_IM, IM.ENDERECO_ID AS IM_ENDERECO_ID, IE.NUMR_IE"
   SQL = SQL & " from IE WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK) "
   SQL = SQL & " ON IE.ENDERECO_ID = ENDERECO.ENDERECO_ID "
   SQL = SQL & " INNER JOIN IM WITH (NOLOCK) "
   SQL = SQL & " ON ENDERECO.ENDERECO_ID = IM.ENDERECO_ID "
   SQL = SQL & " RIGHT OUTER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN FONE WITH (NOLOCK) "
   SQL = SQL & " ON PESSOA.PESSOA_ID = FONE.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN EMAIL WITH (NOLOCK) "
   SQL = SQL & " ON PESSOA.PESSOA_ID = EMAIL.PESSOA_ID "
   SQL = SQL & " ON ENDERECO.PESSOA_ID = PESSOA.PESSOA_ID"
   CONECTA_RETAGUARDA.Execute SQL
'===============vwTRANSPORTADORA
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwTRANSPORTADORA", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwTRANSPORTADORA]"

   SQL = " CREATE VIEW [dbo].[vwTRANSPORTADORA] AS"
   SQL = SQL & " select TRANSPORTADORA.TRANSP_ID, TRANSPORTADORA.DT_CAD, TRANSPORTADORA.STATUS, "
   SQL = SQL & " TRANSPORTADORA.CONTATO, TRANSPORTADORA.ESTABELECIMENTO_ID, PESSOA.PESSOA_ID, "
   SQL = SQL & " PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO, PESSOA.DATA_CAD, PESSOA.SITUACAO, "
   SQL = SQL & " ENDERECO.ENDERECO_ID, ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
   SQL = SQL & " ENDERECO.TIPO, ENDERECO.NUMERO AS NUMERO_ENDERECO,"
   SQL = SQL & " FONE.NUMERO AS NUMERO_FONE, FONE.DDD, FONE.LOCAL, EMAIL.EMAIL, IE.IE_ID, "
   SQL = SQL & " IE.ENDERECO_ID AS IE_ENDERECO_ID, IM.IM_ID, IM.NUMR_IM, IM.ENDERECO_ID AS IM_ENDERECO_ID, IE.NUMR_IE"
   SQL = SQL & " from IE WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK) "
   SQL = SQL & " ON IE.ENDERECO_ID = ENDERECO.ENDERECO_ID "
   SQL = SQL & " INNER JOIN IM WITH (NOLOCK) "
   SQL = SQL & " ON ENDERECO.ENDERECO_ID = IM.ENDERECO_ID "
   SQL = SQL & " RIGHT OUTER JOIN TRANSPORTADORA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON TRANSPORTADORA.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN FONE WITH (NOLOCK) "
   SQL = SQL & " ON PESSOA.PESSOA_ID = FONE.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN EMAIL WITH (NOLOCK) "
   SQL = SQL & " ON PESSOA.PESSOA_ID = EMAIL.PESSOA_ID "
   SQL = SQL & " ON ENDERECO.PESSOA_ID = PESSOA.PESSOA_ID"
   CONECTA_RETAGUARDA.Execute SQL
'===============spPessoa
   If EXISTE_OBJ_BANCO("RETAGUARDA", "spPessoa", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE spPessoa"

   SQL = "create procedure spPessoa "
   SQL = SQL & " @Acao int,@PESSOA_ID bigint,@CNPJCPF nvarchar(14),@DESCRICAO nvarchar(MAX),@RAZAO nvarchar(MAX),@SITUACAO nvarchar(1) "
   SQL = SQL & " as "
   SQL = SQL & " begin "
   SQL = SQL & "    if (@Acao = 1) begin "
   SQL = SQL & "       SET @PESSOA_ID = (select max(pessoa_id) from pessoa) + 1 "
   SQL = SQL & "         INSERT INTO PESSOA "
   SQL = SQL & "         (PESSOA_ID, CnpjCpf, Descricao, RAZAO, DATA_CAD, SITUACAO) "
   SQL = SQL & "         Values "
   SQL = SQL & "       (@PESSOA_ID,@CNPJCPF,@DESCRICAO,@RAZAO,getdate(),@SITUACAO) "
   SQL = SQL & "    End "
   SQL = SQL & "    else if (@Acao = 2) begin "
   SQL = SQL & "       update PESSOA SET "
   SQL = SQL & "          Descricao = @Descricao, "
   SQL = SQL & "          RAZAO = @RAZAO, "
   SQL = SQL & "          SITUACAO = @SITUACAO "
   SQL = SQL & "       where pessoa_id = @PESSOA_ID "
   SQL = SQL & "    End "
   SQL = SQL & "    else if (@Acao = 3) begin "
   SQL = SQL & "       delete from PESSOA where pessoa_id = @PESSOA_ID "
   SQL = SQL & "    End "
   SQL = SQL & "    begin "
   SQL = SQL & "       raiserror('Erro, não executado, spPessoa',14,1) "
   SQL = SQL & "    End "
   SQL = SQL & " End "
'MsgBox SQL
   CONECTA_RETAGUARDA.Execute SQL
'===============vwPedidoVendaItens
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwPedidoVendaItens", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwPedidoVendaItens]"

   SQL = " CREATE VIEW [dbo].[vwPedidoVendaItens] AS"
   SQL = SQL & " select produto.CODG_PRODuto AS Código, "
   SQL = SQL & " PRODUTO.REFERENCIA AS Ref, "
   SQL = SQL & " PRODUTO.DESCRICAO AS Produto, "
   SQL = SQL & " PEDIDOITEM.QTD_PEDIDA AS Qtde, "
   SQL = SQL & " PEDIDOITEM.VALOR_ITEM AS ValorItem, "
   SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO AS Desconto,"
   SQL = SQL & " (PEDIDOITEM.VALOR_ITEM - PEDIDOITEM.VALOR_DESCONTO) * PEDIDOITEM.QTD_PEDIDA AS TotItem, "
   SQL = SQL & " PRODUTO.SITUACAO_TRIBUTARIA AS ST,"
   SQL = SQL & " PRODUTO.ALIQUOTA_ICMS AS ICMS, "
   SQL = SQL & " PRODUTO.CODG_NCM AS NCM, "
   SQL = SQL & " PEDIDOITEM.PEDIDO_ID, "
   SQL = SQL & " PEDIDOITEM.SEQ_ID,"
   SQL = SQL & " PEDIDOITEM.PRODUTO_ID, "
   SQL = SQL & " PEDIDOITEM.STATUS AS StatusItem, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " FAMILIAPRODUTO.PRODUCAO"

   SQL = SQL & " from PEDIDOITEM WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

'===============vwCHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwCHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwCHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO]"

   SQL = "CREATE VIEW [dbo].[vwCHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO]"
   SQL = SQL & " AS "
   SQL = SQL & " select ESTABELECIMENTO.VLR_DIA_COMPRA_PROD, (PEDIDOITEM.QTD_PEDIDA * PEDIDOITEM.VALOR_ITEM) as TotalCompra"
   SQL = SQL & " from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " INNER JOIN USUARIO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.CGCCPF = USUARIO.CPF"
   SQL = SQL & " Where producao = 1"
   SQL = SQL & " and vlr_dia_compra_prod Is Not Null"
   SQL = SQL & " and vlr_dia_compra_prod > 0"
   SQL = SQL & " and PEDIDO.STATUS <> 9"
'   CONECTA_RETAGUARDA.Execute SQL

'===============vwRel_EQUIPAMENTO
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRel_EQUIPAMENTO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRel_EQUIPAMENTO]"

   SQL = "CREATE VIEW [dbo].[vwRel_EQUIPAMENTO]"
   SQL = SQL & " AS "
   SQL = SQL & " select OSEQUIPAMENTO.*, PESSOA.CNPJCPF, PESSOA.situacao "
   SQL = SQL & " from OSEQUIPAMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON OSEQUIPAMENTO.PESSOA_ID = PESSOA.PESSOA_ID"
   CONECTA_RETAGUARDA.Execute SQL

'===============vwRelCheque
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRelCheque", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRelCheque]"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRelCheque", "") = False Then
      SQL = "CREATE VIEW [dbo].[vwRelCheque]"
      SQL = SQL & " AS"
      SQL = SQL & " select BANCO.BANCO_ID, BANCO.CODG_BANCO, BANCO.NOME_BANCO, "
      SQL = SQL & " AGENCIA.AGENCIA_ID, AGENCIA.NUMR_AGENCIA, AGENCIA.NOME_AGENCIA, CONTA.CONTA_ID, "
      SQL = SQL & " CONTA.PESSOA_ID, CONTA.NUMR_CONTA, CONTA.DESC_CONTA, CONTA.DT_Cadastro, "
      SQL = SQL & " CHEQUE.CHEQUE_ID, CHEQUE.NUMR_CHEQUE, CHEQUE.SERIE_CHEQUE,"
      SQL = SQL & " CHEQUE.VALOR, CHEQUE.DT_EMISSAO, CHEQUE.DT_DEPOSITO, CHEQUE.DT_COMPENSA, "
      SQL = SQL & " CHEQUE.STATUS, CHEQUE.RESP_ID, CHEQUE.NUMR_DOC, CHEQUE.ESTABELECIMENTO_ID,"

      SQL = SQL & " CHEQUE.cmc7, CHEQUE.praça, CHEQUE.responsavel, CHEQUE.REPASSE_ID,CHEQUE.REPASSE,"

      SQL = SQL & " PESSOA.CNPJCPF AS CNPJCPF_TERC, PESSOA.DESCRICAO AS NOME_TERC, "
      SQL = SQL & " PESSOA.RAZAO AS RAZAO_TERC, PESSOA.SITUACAO AS ST_TERC, PESSOA_1.CNPJCPF AS CNPJCPF_PROP, PESSOA_1.DESCRICAO AS NOME_PROP,"
      SQL = SQL & " PESSOA_1.RAZAO AS RAZAO_PROP, PESSOA_1.SITUACAO AS ST_PROP"

      SQL = SQL & " from PESSOA AS PESSOA_1 WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN AGENCIA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN BANCO WITH (NOLOCK)"
      SQL = SQL & " ON AGENCIA.BANCO_ID = BANCO.BANCO_ID "
      SQL = SQL & " INNER JOIN CONTA WITH (NOLOCK)"
      SQL = SQL & " ON AGENCIA.AGENCIA_ID = CONTA.AGENCIA_ID "
      SQL = SQL & " ON PESSOA_1.PESSOA_ID = CONTA.PESSOA_ID "
      SQL = SQL & " RIGHT OUTER JOIN PESSOA WITH (NOLOCK)"
      SQL = SQL & " RIGHT OUTER JOIN CHEQUE WITH (NOLOCK)"
      SQL = SQL & " ON PESSOA.PESSOA_ID = CHEQUE.RESP_ID "
      SQL = SQL & " ON CONTA.CONTA_ID = CHEQUE.CONTA_ID"
      CONECTA_RETAGUARDA.Execute SQL
   End If
'===============vwFATURAMENTO
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwFATURAMENTO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwFATURAMENTO]"

   SQL = "CREATE VIEW [dbo].[vwFATURAMENTO] AS "
   SQL = SQL & "select ITEMLANCAMENTO.NUMR_DOC, ITEMLANCAMENTO.SEQ, ITEMLANCAMENTO.FORMAPAGTO_ID, ITEMLANCAMENTO.VALOR_ITEM, "
   SQL = SQL & " ITEMLANCAMENTO.STATUS, ITEMLANCAMENTO.DT_VENCIMENTO, ITEMLANCAMENTO.DT_BAIXA, ITEMLANCAMENTO.DT_CANCELA,"
   SQL = SQL & " ITEMLANCAMENTO.VALOR_DESCONTO, ITEMLANCAMENTO.NUMR_DP, LANCAMENTO.TIPO_LANCAMENTO, ITEMLANCAMENTO.LANCAMENTO_ID,"
   SQL = SQL & " LANCAMENTO.dt_cad, FORMAPAGTO.DESCRICAO as FormaPagto, PESSOA.DESCRICAO AS NomePessoa, PESSOA.RAZAO as RazaoPessoa,"
   SQL = SQL & " PESSOA.SITUACAO as SituacaoPessoa,LANCAMENTO.ESTABELECIMENTO_ID, PESSOA.CNPJCPF, LANCAMENTO.PESSOA_ID, ESTABELECIMENTO.EMPRESA_ID,"
   SQL = SQL & " ESTABELECIMENTO.DESCRICAO AS NomeEstabelecimento, ITEMLANCAMENTO.HISTORICO, itemlancamento.perc_desconto "
   SQL = SQL & " from LANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ESTABELECIMENTO  WITH (NOLOCK)"
   SQL = SQL & " ON LANCAMENTO.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " FULL OUTER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " FULL OUTER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " LEFT OUTER JOIN FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID "
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

'===============vwRel_Nf_Entrada
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRel_Nf_Entrada", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRel_Nf_Entrada]"

   SQL = "CREATE VIEW [dbo].[vwRel_Nf_Entrada]"
   SQL = SQL & " AS "
   SQL = SQL & " select NOTAENTRADA.ESTABELECIMENTO_ID, NOTAENTRADA.FORNECEDOR_ID, NOTAENTRADA.TRANSP_ID, NOTAENTRADA.TIPOENTRADA_ID, NOTAENTRADA.PEDIDOCOMPRA_ID, "
   SQL = SQL & " NOTAENTRADA.USUARIO_ID, NOTAENTRADA.NUMR_NOTA, NOTAENTRADA.SERIE_NOTA, NOTAENTRADA.DT_ENTRADA, NOTAENTRADA.DT_EMISSAO,"
   SQL = SQL & " NOTAENTRADA.STATUS AS Status_Nota, NOTAENTRADA.VALOR_FRETE, NOTAENTRADAITEM.ENTRADA_ID, NOTAENTRADAITEM.SEQ_ID, NOTAENTRADAITEM.PRODUTO_ID,"
   SQL = SQL & " NOTAENTRADAITEM.PRECO_CUSTO, NOTAENTRADAITEM.QTDE_ENTRADA, NOTAENTRADAITEM.STATUS AS STATUS_ITEM, NOTAENTRADAITEM.CFOP_ID AS CFOP_NOTA_ITEM,"
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO AS DescProduto, PRODUTO.SITUACAO, PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.LOCACAO, PRODUTO.DT_ULT_COMPRA,"
   SQL = SQL & " PRODUTO.DT_ULT_VENDA, PRODUTO.REFERENCIA, ESTABELECIMENTO.EMPRESA_ID, ESTABELECIMENTO.DESCRICAO AS NomeEstabelecimento, PRODUTO.UNIDADE_MEDIDA,"
   SQL = SQL & " FAMILIAPRODUTO.CODG_FAMILIA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFornecedor, FORNECEDOR.PESSOA_ID"
   SQL = SQL & " from NOTAENTRADA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK) "
   SQL = SQL & " ON NOTAENTRADA.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"
   CONECTA_RETAGUARDA.Execute SQL

'===============vwRel_Nf_Saida
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRel_Nf_Saida", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwRel_Nf_Saida]"

   SQL = "CREATE VIEW [dbo].[vwRel_Nf_Saida]"
   SQL = SQL & " AS "
   SQL = SQL & " select NF.NF_ID, NF.PEDIDO_ID, NF.PESSOA_ID, NF.estabelecimento_ID, NF.NF_TIPO, NF.TRANSP_ID, NF.NUMR_NOTA, NF.SERIE_NOTA, "
   SQL = SQL & " PESSOA.CNPJCPF AS CNPJCPF_A, NF.DT_EMISSAO, NF.STATUS AS STATUS_NF, NF.DT_CANCELA, NF.QTD_VOLUME,"
   SQL = SQL & " NF.PESO_BRUTO, NF.PESO_LIQUIDO, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID, NFITEM.VALOR,"
   SQL = SQL & " NFITEM.DESCONTO, NFITEM.QTDE, nfitem.cfop_id AS CFOP_ITEM, NFITEM.STRIBUTARIA, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO,"
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.UNIDADE_MEDIDA, PRODUTO.SITUACAO, "
   SQL = SQL & " PRODUTO.SITUACAO_TRIBUTARIA, PRODUTO.ALIQUOTA_ICMS, PRODUTO.CODG_NCM, PRODUTO.PRECO_CUSTO,"
   SQL = SQL & " PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, "
   SQL = SQL & " TRANSPORTADORA.STATUS AS STATUS_TRANSP, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS Nome,"
   SQL = SQL & " PESSOA.RAZAO, PESSOA.SITUACAO AS STATUS_PESSOA"
   SQL = SQL & " from NF WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN NFITEM WITH (NOLOCK) "
   SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN TRANSPORTADORA WITH (NOLOCK) "
   SQL = SQL & " ON NF.TRANSP_ID = TRANSPORTADORA.TRANSP_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON NF.PESSOA_ID = PESSOA.PESSOA_ID"
   CONECTA_RETAGUARDA.Execute SQL

' ============================================================ */
'   Table: BOLETO                                              */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "BOLETO", "") = False Then
      SQL = "create table BOLETO "
      SQL = SQL & " ("
      SQL = SQL & " BOLETO_ID       BIGINT        not null,"
      SQL = SQL & " LANCAMENTO_ID   BIGINT        not null,"
      SQL = SQL & " estabelecimento_ID      INT           not null,"
      SQL = SQL & " DESCRICAO       NVARCHAR(200) not null,"
      SQL = SQL & " VALOR           float         not null,"
      SQL = SQL & " PERC_COMISSAO   float         not null,"
      SQL = SQL & " DT_CAD          datetime      not null,"
      SQL = SQL & " DT_VENC         datetime      not null,"
      SQL = SQL & " constraINT PK_BOLETO primary key (BOLETO_ID))"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vw_BOLETO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vw_BOLETO]"

   SQL = "CREATE VIEW [dbo].[vw_BOLETO] AS "

   SQL = SQL & " select PESSOA.CNPJCPF, PESSOA.DESCRICAO, LANCAMENTO.TIPO_LANCAMENTO, "
   SQL = SQL & " ITEMLANCAMENTO.NUMR_DOC, ITEMLANCAMENTO.SEQ, ITEMLANCAMENTO.FORMAPAGTO_ID, "
   SQL = SQL & " ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.STATUS, ITEMLANCAMENTO.DT_VENCIMENTO,"
   SQL = SQL & " ITEMLANCAMENTO.DT_BAIXA , ITEMLANCAMENTO.DT_CANCELA, ITEMLANCAMENTO.Valor_Desconto, "
   SQL = SQL & " ITEMLANCAMENTO.NUMR_DP, LANCAMENTO.estabelecimento_ID"
   SQL = SQL & " from LANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
   CONECTA_RETAGUARDA.Execute SQL

' ============================================================ */
'   Table: vwLerComanda                                              */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwLerComanda", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwLerComanda]"

   SQL = "CREATE VIEW [dbo].[vwLerComanda] AS "
   SQL = SQL & " SELECT COMANDA.*, COMANDAITEM.SEQ_ID, COMANDAITEM.PRODUTO_ID, COMANDAITEM.QTDE, COMANDAITEM.VALOR_ITEM, COMANDAITEM.USUARIO_ID AS USUCABECA,"
   SQL = SQL & " PEDIDOCOMANDA.PEDIDO_ID, PEDIDOCOMANDA.SEQ_COMANDA_ID, PEDIDOCOMANDA.SEQ_PEDIDO_ID, COMANDAITEM.SITUACAO AS STITEM"
   SQL = SQL & " FROM COMANDA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN COMANDAITEM WITH (NOLOCK) "
   SQL = SQL & " ON COMANDA.COMANDA_ID = COMANDAITEM.COMANDA_ID "
   SQL = SQL & " INNER JOIN PEDIDOCOMANDA WITH (NOLOCK) "
   SQL = SQL & " ON COMANDA.CARTAOBARRA_ID = PEDIDOCOMANDA.CARTAOBARRA_ID"

   CONECTA_RETAGUARDA.Execute SQL

MsgBox "ok, Rotinas VW atualizadas "
End Sub
'===================================
Private Sub cmdNFEntrada_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "NOTAENTRADA", "U") = True Then
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADA_EMPRESA", "") = True Then
         SQL = "alter table NOTAENTRADA "
         SQL = SQL & " drop CONSTRAINT FK_NOTAENTRADA_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENTRADA_ID", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "ENTRADA_ID", "BIGINT NOT NULL", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_PEDIDO_COMPRA", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_PEDIDO_COMPRA", "BIGINT", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPOENTRADA_ID", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "TIPOENTRADA_ID", "BIGINT", "NOTAENTRADA", "RETAGUARDA"
      
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_NOTA", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_NOTA", "BIGINT  NOT NULL", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SERIE_NOTA", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "SERIE_NOTA", "NVARCHAR(10)", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_FORNEC", "NOTAENTRADA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADA.CODG_FORNEC'" & "," & "'FORNECEDOR_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORNECEDOR_ID", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "FORNECEDOR_ID", "BIGINT NOT NULL", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_TRANSP", "NOTAENTRADA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADA.CODG_TRANSP'" & "," & "'TRANSP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TRANSP_ID", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "TRANSP_ID", "BIGINT", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "numr_pedido_compra", "NOTAENTRADA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADA.numr_pedido_compra'" & "," & "'PEDIDOCOMPRA_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "NOTAENTRADA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADA.EMPRESA_ID'" & "," & "'ESTABELECIMENTO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_USU", "NOTAENTRADA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADA.CODG_USU'" & "," & "'USUARIO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USUARIO_ID", "NOTAENTRADA") = True Then _
         Alteração_Definição_Campo_Tabela "USUARIO_ID", "INT", "NOTAENTRADA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cfop", "NOTAENTRADA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADA DROP COLUMN cfop"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CHAVENFEIMPORTADO", "NOTAENTRADA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADA ADD CHAVENFEIMPORTADO NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VERSAOLAYOUTNFEIMPOR", "NOTAENTRADA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADA ADD VERSAOLAYOUTNFEIMPOR NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DTIMPORTACAO", "NOTAENTRADA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADA ADD DTIMPORTACAO DATETIME"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_NOTAENTRADA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADA ADD CONSTRAINT pk_NOTAENTRADA PRIMARY KEY (ENTRADA_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADA_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[NOTAENTRADA]  WITH CHECK ADD  CONSTRAINT [FK_NOTAENTRADA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NOTAENTRADA] CHECK CONSTRAINT [FK_NOTAENTRADA_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADA_FORNECEDOR", "") = False Then
         SQL = "ALTER TABLE [dbo].[NOTAENTRADA]  WITH CHECK ADD  CONSTRAINT [FK_NOTAENTRADA_FORNECEDOR] FOREIGN KEY([FORNECEDOR_ID])"
         SQL = SQL & " References [dbo].[FORNECEDOR]([FORNECEDOR_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NOTAENTRADA] CHECK CONSTRAINT [FK_NOTAENTRADA_FORNECEDOR]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADA_TRANSP", "") = False Then
         SQL = "ALTER TABLE [dbo].[NOTAENTRADA]  WITH CHECK ADD  CONSTRAINT [FK_NOTAENTRADA_TRANSP] FOREIGN KEY([TRANSP_ID])"
         SQL = SQL & " References [dbo].[TRANSPORTADORA]([TRANSP_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NOTAENTRADA] CHECK CONSTRAINT [FK_NOTAENTRADA_TRANSP]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'=================NOTAENTRADAITEM
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_prod", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM DROP COLUMN codg_prod"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRECO_VENDA", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM DROP COLUMN PRECO_VENDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_PEDIDO_COMPRA", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM DROP COLUMN NUMR_PEDIDO_COMPRA"
   
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUTO_ID", "NOTAENTRADAITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM ADD PRODUTO_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENTRADA_ID", "NOTAENTRADAITEM") = True Then _
         Alteração_Definição_Campo_Tabela "ENTRADA_ID", "BIGINT NOT NULL", "NOTAENTRADAITEM", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "NOTAENTRADAITEM") = True Then _
         Alteração_Definição_Campo_Tabela "SEQ", "BIGINT NOT NULL", "NOTAENTRADAITEM", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADAITEM.SEQ'" & "," & "'SEQ_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NCM", "NOTAENTRADAITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM ADD NCM BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CST", "NOTAENTRADAITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM ADD CST nvarchar(4)"
         Else: Alteração_Definição_Campo_Tabela "CST", "nvarchar(4)", "NOTAENTRADAITEM", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "UN", "NOTAENTRADAITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM ADD UN NVARCHAR(2)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "qtd_entrada", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADAITEM.qtd_entrada'" & "," & "'QTDE_ENTRADA'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "qtde_entrada", "NOTAENTRADAITEM") = True Then _
         Alteração_Definição_Campo_Tabela "qtde_entrada", "NUMERIC(18,3) NOT NULL", "NOTAENTRADAITEM", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQNOTAITEMIMPORTADO", "NOTAENTRADAITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM ADD SEQNOTAITEMIMPORTADO BIGINT"

MsgBox "MUDAR NOME DA CHAVE DA TABELA NOTAENTRADAITEM"
      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_NOTAENTRADAITEM", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NOTAENTRADAITEM ADD CONSTRAINT pk_NOTAENTRADAITEM PRIMARY KEY (ENTRADA_ID,SEQ)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADAITEM_NOTAENTRADA", "") = False Then
         SQL = "ALTER TABLE [dbo].[NOTAENTRADAITEM]  WITH CHECK ADD  CONSTRAINT [FK_NOTAENTRADAITEM_NOTAENTRADA] FOREIGN KEY([ENTRADA_ID])"
         SQL = SQL & " References [dbo].[NOTAENTRADA]([ENTRADA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NOTAENTRADAITEM] CHECK CONSTRAINT [FK_NOTAENTRADAITEM_NOTAENTRADA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

'========================================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP", "NOTAENTRADAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'NOTAENTRADAITEM.CFOP'" & "," & "'CFOP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP_ID", "NOTAENTRADAITEM") = True Then _
         Alteração_Definição_Campo_Tabela "CFOP_ID", "nvarchar (10)", "NOTAENTRADAITEM", "RETAGUARDA"

Call cmdCFOP_Click

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADAITEM_CFOP", "") = False Then
         SQL = "update NOTAENTRADAITEM set "
         SQL = SQL & " cfop_id = '1102'"
         SQL = SQL & " where cfop_id is null"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[NOTAENTRADAITEM]  WITH CHECK ADD  CONSTRAINT [FK_NOTAENTRADAITEM_CFOP] FOREIGN KEY([CFOP_ID])"
         SQL = SQL & " References [dbo].[CFOP]([CFOP_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NOTAENTRADAITEM] CHECK CONSTRAINT [FK_NOTAENTRADAITEM_CFOP]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NOTAENTRADAITEM_PRODUTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[NOTAENTRADAITEM]  WITH CHECK ADD  CONSTRAINT [FK_NOTAENTRADAITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
         SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NOTAENTRADAITEM] CHECK CONSTRAINT [FK_NOTAENTRADAITEM_PRODUTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Private Sub cmdCFOP_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CFOP", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "CFOP") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP ADD ESTABELECIMENTO_ID INT"

         SQL = "update CFOP set estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " where estabelecimento_id is null"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CFOP.CODIGO'" & "," & "'CFOP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_DA_OPERACAO", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN TIPO_DA_OPERACAO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORA_ESTABELECIMENTO", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN FORA_ESTABELECIMENTO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO_DA_ALIQUOTA", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN CODIGO_DA_ALIQUOTA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN TIPO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESCRICAO", "CFOP") = True Then _
         Alteração_Definição_Campo_Tabela "DESCRICAO", "NVARCHAR(MAX)", "CFOP", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MSGFISCO", "CFOP") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP ADD MSGFISCO NVARCHAR(MAX)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "OBS", "CFOP") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP ADD OBS NVARCHAR(MAX)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CFOP", "") = False Then
         SQL = "ALTER TABLE CFOP ADD CONSTRAINT pk_CFOP PRIMARY KEY (CFOP_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_ICMS", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CFOP.PERC_ICMS'" & "," & "'ALIQUOTA_ICMS_DENTRO'" & "," & "'COLUMN'"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "ICMS_PJ_F_UF", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CFOP.ICMS_PJ_F_UF'" & "," & "'ALIQUOTA_ICMS_FORA'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_ICMS", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN PERC_ICMS"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ICMS_PJ_F_UF", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN ICMS_PJ_F_UF"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALIQUOTA_ICMS_dentro", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN ALIQUOTA_ICMS_dentro"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALIQUOTA_ICMS_FORA", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN ALIQUOTA_ICMS_FORA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CST_PIS", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN CST_PIS"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CST_COFINS", "CFOP") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CFOP DROP COLUMN CST_COFINS"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CFOP_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CFOP]  WITH CHECK ADD  CONSTRAINT [FK_CFOP_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CFOP] CHECK CONSTRAINT [FK_CFOP_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      'RODA_CFOP_PLANILHA

   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CFOPUF", "U") = False Then
      SQL = "CREATE TABLE [dbo].[CFOPUF]("
      SQL = SQL & " [CFOPUF_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CFOP_ID] [nvarchar](10) NOT NULL,"
      SQL = SQL & " [UF_ORIGEM] [nvarchar](2) NOT NULL,"
      SQL = SQL & " [UF_DESTINO] [nvarchar](2) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_CFOPUF] PRIMARY KEY CLUSTERED([CFOPUF_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[CFOPUF]  WITH CHECK ADD  CONSTRAINT [FK_CFOPUF_CFOP] FOREIGN KEY([CFOP_ID]) References [dbo].[CFOP]([CFOP_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[CFOPUF] CHECK CONSTRAINT [FK_CFOPUF_CFOP]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

MsgBox "OK"
End Sub
'''''======================================
Sub RODA_CFOP_PLANILHA()
'On Error GoTo ERRO_TRATA

   Msg = "Deseja importar cadastro de CFOP "
   PERGUNTA Msg, vbYesNo + 32, "Atualização", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      Dim TabCFOP As New ADODB.Recordset

      'frmINICIO.Dialogo.FileName = ""
      'frmINICIO.Dialogo.InitDir = App.Path
      'frmINICIO.Dialogo.DialogTitle = "Importação arquivo"
      'frmINICIO.Dialogo.Filter = "*.csv;*.txt"
      'frmINICIO.Dialogo.ShowOpen
      'If frmINICIO.Dialogo.FileName <> "" Then _
         SQL3 = frmINICIO.Dialogo.FileName

      Set oConn = New ADODB.Connection
      oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                         "FIL=excel 8.0;" & _
                         "DefaultDir=" & App.Path & "\TXT\" & ";" & _
                         "MaxBufferSize=2048;" & _
                         "PageTimeout=5;" & _
                         "DBQ=" & App.Path & "\TXT\CFOP_AT.xls" & ";"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'aabre o recordset pelo nome da planilha
      TabConsulta.Open "[TABELACFOP$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

      TabConsulta.MoveFirst

      If TabConsulta.EOF Then
         MsgBox "Planilha incorreta !!!"
         Exit Sub
      End If

      While Not TabConsulta.EOF

         If Not IsNull(TabConsulta.Fields(0).Value) Then
            If Trim(TabConsulta.Fields(0).Value) <> "" Then
               If IsNumeric(TabConsulta.Fields(0).Value) Then
                  If TabConsulta.Fields(0).Value > 0 Then
                     If Not IsNull(TabConsulta.Fields(1).Value) Then
                        If TabCFOP.State = 1 Then _
                           TabCFOP.Close

                        SQL = "select * from CFOP "
                        SQL = SQL & " where cfop_id = '" & Trim(TabConsulta.Fields(0).Value) & "'"
                        TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If TabCFOP.EOF Then
                           If TabCFOP.State = 1 Then _
                              TabCFOP.Close

                           SQL = "INSERT INTO CFOP "
                              SQL = SQL & " (cfop_id,descricao,estabelecimento_id)"
                           SQL = SQL & " VALUES ("
                              SQL = SQL & "'" & Trim(TabConsulta.Fields(0).Value) & "'"
                              SQL = SQL & ",'" & Trim(TabConsulta.Fields(1).Value) & "'"
                              SQL = SQL & "," & ESTABELECIMENTO_ID_N
                           SQL = SQL & ")"

                           Me.Caption = TabConsulta.Fields(0).Value
                           Else
                              SQL = "update CFOP set "
                              SQL = SQL & " descricao = '" & Trim(TabConsulta.Fields(1).Value) & "'"
                              SQL = SQL & " where cfop_id = '" & Trim(TabConsulta.Fields(0).Value) & "'"

                              cmdCFOP.Caption = TabConsulta.Fields(0).Value
                        End If
                        If TabCFOP.State = 1 Then _
                           TabCFOP.Close

                        CONECTA_RETAGUARDA.Execute SQL
                     End If
                  End If
               End If
            End If
         End If
         DoEvents
         On Error Resume Next
         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      Command25.Caption = SQL3
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RODA_BANCO"
End Sub

Private Sub cmdTRANSP_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "TRANSPORTADORA", "U") = True Then
'===========alter=====================================================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "TRANSPORTADORA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA ADD PESSOA_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT not null", "TRANSPORTADORA", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TRANSP_ID", "TRANSPORTADORA") = True Then _
         Alteração_Definição_Campo_Tabela "TRANSP_ID", "BIGINT not null", "TRANSPORTADORA", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_TRANSPORTADORA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA ADD CONSTRAINT pk_TRANSPORTADORA PRIMARY KEY (TRANSP_ID)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTATO", "TRANSPORTADORA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA ADD CONTATO NVARCHAR(30) "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "TRANSPORTADORA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA ADD ESTABELECIMENTO_ID INT"
         SQL = "update TRANSPORTADORA set estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If

'===========drop=====================================================
      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_TRANSPORTADORA_CNPJCPF", "") = True Then
         SQL = "alter table TRANSPORTADORA "
         SQL = SQL & " drop CONSTRAINT IX_TRANSPORTADORA_CNPJCPF"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_TRANSPORTADORA_CNPJ", "") = True Then
         SQL = "alter table TRANSPORTADORA "
         SQL = SQL & " drop CONSTRAINT IX_TRANSPORTADORA_CNPJ"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'==================
   
'======================
      
'===========================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cgccpf", "TRANSPORTADORA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA DROP COLUMN cgccpf"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "nome", "TRANSPORTADORA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA DROP COLUMN nome"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "razao_social", "TRANSPORTADORA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA DROP COLUMN razao_social"

      '=========================== IE
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IE", "TRANSPORTADORA") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from TRANSPORTADORA WITH (NOLOCK)"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value
            ENDERECO_ID_N = 0 & TRAZ_ID_ENDERECO("C")

            If Not IsNull(TabTemp.Fields("ie").Value) Then _
               If IsNumeric(TabTemp.Fields("ie").Value) Then _
                  GRAVA_IE TabTemp.Fields("ie").Value

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA DROP COLUMN IE"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_TRANSPORTADORA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA] CHECK CONSTRAINT [FK_TRANSPORTADORA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_TRANSPORTADORA_EMPRESA", "") = True Then
         SQL = "alter table TRANSPORTADORA "
         SQL = SQL & " drop CONSTRAINT FK_TRANSPORTADORA_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "TRANSPORTADORA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA DROP COLUMN EMPRESA_ID"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_TRANSPORTADORA_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA] CHECK CONSTRAINT [FK_TRANSPORTADORA_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   ENDERECO_ID_N = 0
   PESSOA_ID_N = 0
   If EXISTE_OBJ_BANCO("RETAGUARDA", "TRANSPORTADORACOMPRADOR", "U") = False Then
      SQL = "CREATE TABLE [dbo].[TRANSPORTADORACOMPRADOR]("
      SQL = SQL & " [USUARIO_ID] [BIGint] NOT NULL,"
      SQL = SQL & " [TRANSP_ID] [bigint] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_TRANSPORTADORACOMPRADOR] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([USUARIO_ID] ASC,[TRANSP_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TRANSPORTADORACOMPRADOR]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORACOMPRADOR_TRANSPORTADORA] FOREIGN KEY([TRANSP_ID])"
      SQL = SQL & " References [dbo].[TRANSPORTADORA]([TRANSP_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TRANSPORTADORACOMPRADOR] CHECK CONSTRAINT [FK_TRANSPORTADORACOMPRADOR_TRANSPORTADORA]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TRANSPORTADORACOMPRADOR]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORACOMPRADOR_USUARIO] FOREIGN KEY([USUARIO_ID])"
      SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TRANSPORTADORACOMPRADOR] CHECK CONSTRAINT [FK_TRANSPORTADORACOMPRADOR_USUARIO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   MsgBox "Ok, TABELEA TRANSPORTADORA "
End Sub

Private Sub cmdPedido_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDO", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "PEDIDO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD PEDIDO_ID BIGINT"
         Else
            If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "PEDIDO") = True Then _
               Alteração_Definição_Campo_Tabela "PEDIDO_ID", "BIGINT", "PEDIDO", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_REGISTRO", "PEDIDO") = True Then _
         Alteração_Definição_Campo_Tabela "TIPO_REGISTRO", "NVARCHAR(2)", "PEDIDO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VENDEDOR", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'PEDIDO.VENDEDOR'" & "," & "'VENDEDOR_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VENDEDOR_ID", "PEDIDO") = True Then _
         Alteração_Definição_Campo_Tabela "VENDEDOR_ID", "BIGINT", "PEDIDO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODGFUNC", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN CODGFUNC"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_DESC_CONVENIO", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN PERC_DESC_CONVENIO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VLR_FRETE", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN VLR_FRETE"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_CLIENTE", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'PEDIDO.CODG_CLIENTE'" & "," & "'CLIENTE_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CLIENTE_ID", "PEDIDO") = True Then _
         Alteração_Definição_Campo_Tabela "CLIENTE_ID", "BIGINT", "PEDIDO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMERO_CAIXA_CPU", "PEDIDO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD NUMERO_CAIXA_CPU INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "PEDIDO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD ESTABELECIMENTO_ID INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP_ID", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN cfop_id"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN cfop"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BAIXADO", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN BAIXADO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_CUPOM", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN NUMR_CUPOM"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_DOC", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN NUMR_DOC"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERCCOMISSAO", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN PERCCOMISSAO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "RESP_VENDA", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN RESP_VENDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONT_ENTRADA", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN CONT_ENTRADA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_DOC", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN TIPO_DOC"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_USU", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'PEDIDO.CODG_USU'" & "," & "'USUARIO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USUARIO_ID", "PEDIDO") = True Then _
         Alteração_Definição_Campo_Tabela "USUARIO_ID", "INT", "PEDIDO", "RETAGUARDA"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "TABELAPRECO_ID", "PEDIDO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD TABELAPRECO_ID INT "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_REQ", "PEDIDO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN NUMR_REQ"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PREFIXO", "PEDIDO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD PREFIXO nvarchar(3)"

'=========================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "tabelapreco_id", "PEDIDO") = True Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
         SQL = "select tabelapreco_id,pedido_id,tipovenda_id from PEDIDO WITH (NOLOCK)"
         'SQL = SQL & " where tabelapreco_id > 0 "
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabCabeca.EOF
            PEDIDO_ID_N = 0 & TabCabeca.Fields("pedido_id").Value
            TABELAPRECO_ID_N = 0 & TabCabeca.Fields("tabelapreco_id").Value
            FORMAPAGTO_ID_N = 0 & TabCabeca.Fields("tipovenda_id").Value
            TIPOVENDA_ID_N = 0 & TabCabeca.Fields("tipovenda_id").Value

            If TabTemp.State = 1 Then _
               TabTemp.Close
            SQL = "select pedido_id from PEDIDOFATURA WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               Acao_N = 2
               Else: Acao_N = 1
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close

            spPEDIDOFATURA Acao_N, 0, PEDIDO_ID_N, TABELAPRECO_ID_N, FORMAPAGTO_ID_N, TIPOVENDA_ID_N

cmdPedido.Caption = PEDIDO_ID_N
DoEvents

            TabCabeca.MoveNext
         Wend
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN tabelapreco_id"

         Else 'ACERTO NA TABELA PEDIDOFATURA
            CONT_N = 0
            TABELAPRECO_ID_N = 0
            FORMAPAGTO_ID_N = 1
            TIPOVENDA_ID_N = 9999

         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select pedido_id from PEDIDO "
         SQL = SQL & " where pedido_id not in (select pedido_id from PEDIDOFATURA)"
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabCabeca.EOF
            PEDIDO_ID_N = 0 & TabCabeca.Fields("pedido_id").Value
            TABELAPRECO_ID_N = 0
            TIPOVENDA_ID_N = 9999
            FORMAPAGTO_ID_N = 1

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select tipovenda_id from LANCAMENTO"
            SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               If Not TabTemp.Fields(0).Value Then _
                  TIPOVENDA_ID_N = 0 & TabTemp.Fields(0).Value
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select formapagto_id from itemLANCAMENTO"
            SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               If Not TabTemp.Fields(0).Value Then _
                  FORMAPAGTO_ID_N = 0 & TabTemp.Fields(0).Value

            If TabTemp.State = 1 Then _
               TabTemp.Close
            SQL = "select pedido_id from PEDIDOFATURA WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               Acao_N = 2
               Else: Acao_N = 1
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close

            spPEDIDOFATURA Acao_N, 0, PEDIDO_ID_N, TABELAPRECO_ID_N, FORMAPAGTO_ID_N, TIPOVENDA_ID_N

CONT_N = CONT_N + 1
cmdPedido.Caption = CONT_N
DoEvents

            TabCabeca.MoveNext
         Wend
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
'=========
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "tipovenda_id", "PEDIDO") = True Then
         'If EXISTE_CAMPO_TABELA("RETAGUARDA", "tipovenda_id", "PEDIDO") = True Then
         'End If
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN tipovenda_id"
      End If
'=========================

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_PEDIDO", "") = False Then
         SQL = "ALTER TABLE PEDIDO ADD CONSTRAINT pk_PEDIDO PRIMARY KEY (PEDIDO_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDO_USUARIO", "") = False Then
         SQL = "ALTER TABLE [dbo].[PEDIDO]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDO_USUARIO] FOREIGN KEY([USUARIO_ID])"
         SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[PEDIDO] CHECK CONSTRAINT [FK_PEDIDO_USUARIO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDO_VENDEDOR", "") = False Then
         SQL = "ALTER TABLE [dbo].[PEDIDO]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDO_VENDEDOR] FOREIGN KEY([VENDEDOR_ID])"
         SQL = SQL & " References [dbo].[VENDEDOR]([VENDEDOR_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[PEDIDO] CHECK CONSTRAINT [FK_PEDIDO_VENDEDOR]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDO_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[PEDIDO]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[PEDIDO] CHECK CONSTRAINT [FK_PEDIDO_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDO_CLIENTE", "") = False Then
         If EXISTE_OBJ_BANCO("RETAGUARDA", "cgccpf", "") = True Then
            If TabTemp.State = 1 Then _
               TabTemp.Close
   
            SQL = "select cgccpf from PEDIDO "
            SQL = SQL & " where CLIENTE_ID not in (select CLIENTE_ID from CLIENTE)"
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            While Not TabTemp.EOF
               If Not IsNull(TabTemp.Fields(0).Value) Then
                  If Trim(TabTemp.Fields(0).Value) <> "" Then
   
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
   
                     SQL = "select cliente_id from CLIENTE "
                     SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields(0).Value) & "'"
                     TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabConsulta.EOF Then
                        SQL = "update PEDIDO set "
                        SQL = SQL & " cliente_id = " & TabConsulta.Fields(0).Value
                        SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields(0).Value) & "'"
   
                        CONECTA_RETAGUARDA.Execute SQL
                        Else
                           If TabConsulta.State = 1 Then _
                              TabConsulta.Close
   
                           SQL = "select cliente_id from CLIENTE "
                           SQL = SQL & " where cgccpf = '99999999999' "
                           TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabConsulta.EOF Then
                              SQL = "update PEDIDO set cliente_id = " & TabConsulta.Fields(0).Value
                              SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields(0).Value) & "'"
                              CONECTA_RETAGUARDA.Execute SQL
                           End If
   
                           If TabConsulta.State = 1 Then _
                              TabConsulta.Close
                     End If
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
                  End If
               End If
               TabTemp.MoveNext
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
         SQL = "ALTER TABLE [dbo].[PEDIDO]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDO_CLIENTE] FOREIGN KEY([CLIENTE_ID])"
         SQL = SQL & " References [dbo].[CLIENTE]([CLIENTE_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[PEDIDO] CHECK CONSTRAINT [FK_PEDIDO_CLIENTE]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

   End If

If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOTIME", "U") = False Then
   SQL = "CREATE TABLE [dbo].[PEDIDOTIME]("
   SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
   SQL = SQL & " [DT_IN] [datetime] NOT NULL,"
   SQL = SQL & " [DT_FIM] [datetime] NOT NULL,"
   SQL = SQL & " [TIPO_DOC] [nvarchar](3) NOT NULL,"
   SQL = SQL & " [NUMR_DOC] [bigint]) ON [PRIMARY]"
   CONECTA_RETAGUARDA.Execute SQL
End If
'------------PEDIDOITEM
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOITEM", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "PEDIDOITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD PEDIDO_ID BIGINT"
         Else
            If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "PEDIDOITEM") = True Then _
               Alteração_Definição_Campo_Tabela "PEDIDO_ID", "BIGINT", "PEDIDOITEM", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRECO_CUSTO", "PEDIDOITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD PRECO_CUSTO FLOAT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESO_ITEM", "PEDIDOITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD PESO_ITEM FLOAT"
         Else: Alteração_Definição_Campo_Tabela "PESO_ITEM", "FLOAT", "PEDIDOITEM", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_REG", "PEDIDOITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD TIPO_REG CHAR(2)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_PROD", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN codg_prod"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTDE_BALANCA", "PEDIDOITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD QTDE_BALANCA FLOAT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP", "pedidoitem") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'pedidoitem.cfop'" & "," & "'CFOP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP_ID", "pedidoitem") = True Then _
         Alteração_Definição_Campo_Tabela "CFOP_ID", "nvarchar (10)", "pedidoitem", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_REQ", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN NUMR_REQ"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USU_ATENDE", "PEDIDOITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM add USU_ATENDE INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALTURA", "PEDIDOITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD ALTURA FLOAT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LARGURA", "PEDIDOITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD LARGURA FLOAT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_ID", "PEDIDOITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM add SEQ_ID BIGINT"
         CONT_N = 0
      
         SQL = "UPDATE PEDIDOITEM set SEQ_ID = 0"
         'SQL = SQL & " where SEQ_ID Is Null "
         CONECTA_RETAGUARDA.Execute SQL

         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         SQL = "select * from PEDIDOITEM "
         'SQL = SQL & " WHERE tipo_reg = 'PC' "
         SQL = SQL & " order by pedido_id"
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabPedidoItem.EOF

            If TabPedidoItem.Fields("pedido_id").Value <> Numero_Pedido_N Then
               CONT_N = 1
               Numero_Pedido_N = TabPedidoItem.Fields("pedido_id").Value
            End If

            SQL = "UPDATE PEDIDOITEM set "
            SQL = SQL & " SEQ_ID = " & CONT_N
            SQL = SQL & " , tipo_reg = 'PC' "
            SQL = SQL & " where pedido_id = " & Numero_Pedido_N
            SQL = SQL & " and produto_id = '" & TabPedidoItem.Fields("produto_id").Value & "'"
            SQL = SQL & " and SEQ_ID = 0 "
            CONECTA_RETAGUARDA.Execute SQL

            CONT_N = CONT_N + 1
            frmATUALIZACAO2.Caption = TabPedidoItem.Fields("pedido_id").Value & " / " & CONT_N

            TabPedidoItem.MoveNext
            DoEvents
         Wend

         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN SEQ"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN EMPRESA_ID"

      '''''''''''relação itens com produto
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUTO_ID", "PEDIDOITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM add PRODUTO_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PRODUTO_ID", "BIGINT", "PEDIDOITEM", "RETAGUARDA"
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDOITEM_PRODUTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[PEDIDOITEM] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_PEDIDOITEM_PRODUTO] "
         SQL = SQL & " FOREIGN KEY([PRODUTO_ID])"
         SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[PEDIDOITEM] CHECK CONSTRAINT [FK_PEDIDOITEM_PRODUTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDO_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[PEDIDO]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[PEDIDO] CHECK CONSTRAINT [FK_PEDIDO_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOENCOMENDA", "U") = False Then
      SQL = " CREATE TABLE [dbo].[PEDIDOENCOMENDA]("
      SQL = SQL & " [PEDIDOENCOMENDA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DT_RECEBE] [datetime] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [VLR_TX_ENTREGA] [float] "
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOENCOMENDA]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOENCOMENDA_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
      SQL = SQL & " References [dbo].[Pedido]([PEDIDO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOENCOMENDA] CHECK CONSTRAINT [FK_PEDIDOENCOMENDA_PEDIDO]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "VLR_TX_ENTREGA", "PEDIDOENCOMENDA") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOENCOMENDA ADD VLR_TX_ENTREGA float"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOTEMPITEM", "U") = False Then
      SQL = "select * into PEDIDOTEMPITEM FROM PEDIDOITEM WHERE PEDIDO_ID < 0"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOCOMANDA", "U") = False Then
      SQL = "CREATE TABLE [dbo].[PEDIDOCOMANDA]("
      SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CARTAOBARRA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_COMANDA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_PEDIDO_ID] [bigint] NOT NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdFinanc_Click()
'set pensar em criar campo pra gravar nome cliente na tabela lancamento quando não esta cadastrado no sistema
   
   If EXISTE_OBJ_BANCO("RETAGUARDA", "LANCAMENTO", "U") = True Then

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NOME_PESSOA", "LANCAMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO ADD NOME_PESSOA NVARCHAR(50) "

         SQL = "UPDATE LANCAMENTO SET NOME_PESSOA = left(PESSOA.DESCRICAO,30)"
         SQL = SQL & " FROM LANCAMENTO INNER JOIN PESSOA"
         SQL = SQL & " ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_LANC", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'lancamento.DT_LANC'" & "," & "'DT_CAD'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_PAGTO", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'lancamento.TIPO_PAGTO'" & "," & "'TIPOVENDA_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LANCAMENTO_ID", "LANCAMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ITEMLANCAMENTO ADD LANCAMENTO_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "LANCAMENTO_ID", "BIGINT NOT NULL", "LANCAMENTO", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "LANCAMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO ADD PESSOA_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_DOC", "LANCAMENTO") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_DOC", "BIGINT", "LANCAMENTO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "prop", "LANCAMENTO") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close
   CONT_N = 0
         SQL = "select distinct(prop) from LANCAMENTO  "
         SQL = SQL & " where pessoa_id is null "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CONT_N = CONT_N + 1
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
   
            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(TabTemp.Fields("prop").Value) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update LANCAMENTO set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
   
            frmATUALIZACAO2.Caption = "LANCAMENTO = " & Trim(TabTemp.Fields("prop").Value)
            DoEvents
   
            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LANCAMENTO_ID", "LANCAMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO ADD LANCAMENTO_ID BIGINT NOT NULL"
         Else
            If EXISTE_CAMPO_TABELA("RETAGUARDA", "LANCAMENTO_ID", "LANCAMENTO") = True Then _
               Alteração_Definição_Campo_Tabela "LANCAMENTO_ID", "BIGINT NOT NULL", "LANCAMENTO", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_DOC", "ITEMLANCAMENTO") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_DOC", "BIGINT", "ITEMLANCAMENTO", "RETAGUARDA"

      SQL = "UPDATE ITEMLANCAMENTO SET Status = 'B'"
      SQL = SQL & " where FORMAPAGTO_ID = 1 and status = 'A'"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "update itemLANCAMENTO set "
      SQL = SQL & " dt_baixa = '" & Now & "'"
      SQL = SQL & " where dt_baixa Is Null "
      SQL = SQL & " and status = 'B' "
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "update LANCAMENTO set LANCAMENTO_ID = NUMR_DOC"
      SQL = SQL & " where LANCAMENTO_ID Is Null Or LANCAMENTO_ID <= 0"
      CONECTA_RETAGUARDA.Execute SQL

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "LANCAMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO ADD ESTABELECIMENTO_ID INT"

         SQL = "update LANCAMENTO set "
         SQL = SQL & " LANCAMENTO.estabelecimento_id = pedido.estabelecimento_id "

         SQL = SQL & " from PEDIDO "
         SQL = SQL & " where pedido.pedido_id = lancamento.numr_doc "
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_LANCAMENTO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO ADD CONSTRAINT pk_LANCAMENTO PRIMARY KEY (LANCAMENTO_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_LANCAMENTO_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[LANCAMENTO]  WITH CHECK ADD  CONSTRAINT [FK_LANCAMENTO_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[LANCAMENTO] CHECK CONSTRAINT [FK_LANCAMENTO_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

'so pode excluir quando rodar todas rotinas
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "lancamento") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN prop"

      '============ITEMLANCAMENTO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORMA_ID", "ITEMLANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'ITEMLANCAMENTO.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LANCAMENTO_ID", "ITEMLANCAMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ITEMLANCAMENTO ADD LANCAMENTO_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "LANCAMENTO_ID", "BIGINT NOT NULL", "ITEMLANCAMENTO", "RETAGUARDA"
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CC_ID", "ITEMLANCAMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ITEMLANCAMENTO ADD CC_ID INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "HISTORICO", "ITEMLANCAMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ITEMLANCAMENTO ADD HISTORICO NVARCHAR(MAX)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_ITEMLANCAMENTO", "") = False Then
         SQL = "ALTER TABLE ITEMLANCAMENTO ADD CONSTRAINT pk_ITEMLANCAMENTO PRIMARY KEY (LANCAMENTO_ID,SEQ)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "GerouRemessa", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN GerouRemessa"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CodigoContaCorrente", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN CodigoContaCorrente"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VALOR_LANC", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN VALOR_LANC"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TOTAL_DESCONTO", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN TOTAL_DESCONTO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "doc_antigo", "LANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE LANCAMENTO DROP COLUMN doc_antigo"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ITEMLANC_LANCAMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[ITEMLANCAMENTO]  WITH CHECK ADD  CONSTRAINT [FK_ITEMLANC_LANCAMENTO] FOREIGN KEY([LANCAMENTO_ID])"
         SQL = SQL & " References [dbo].[LANCAMENTO]([LANCAMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ITEMLANCAMENTO] CHECK CONSTRAINT [FK_ITEMLANC_LANCAMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdOBS_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OBS", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "OBS") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE OBS DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "OBS") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE OBS ADD PESSOA_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "pedido_ID", "OBS") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE OBS ADD pedido_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "OBS") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONT_N = 0

         SQL = "select distinct(prop) from OBS  where pessoa_id is null "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CONT_N = CONT_N + 1
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
   
            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(TabTemp.Fields("prop").Value) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update OBS set pessoa_id = " & TabPessoa.Fields(0).Value
               'SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " where prop = '" & Trim(TabTemp.Fields("prop").Value) & "'"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close
   
            frmATUALIZACAO2.Caption = "OBS = " & Trim(TabTemp.Fields("prop").Value)
            DoEvents
   
            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OBS_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[OBS]  WITH CHECK ADD  CONSTRAINT [FK_OBS_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[OBS] CHECK CONSTRAINT [FK_OBS_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OBS_PEDIDO", "") = False Then
         SQL = " ALTER TABLE [dbo].[OBS]  WITH CHECK ADD  CONSTRAINT [FK_OBS_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
         SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[OBS] CHECK CONSTRAINT [FK_OBS_PEDIDO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'so pode excluir quando rodar todas rotinas
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "obs") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE OBS DROP COLUMN PROP"
      Else
         SQL = "CREATE TABLE [dbo].[OBS]("
         SQL = SQL & " [PESSOA_ID] [bigint] NULL,"
         SQL = SQL & " [PEDIDO_ID] [bigint] NULL,"
         SQL = SQL & " [PROP] [nvarchar](max) NULL,"
         SQL = SQL & " [SEQ] [int] NOT NULL,"
         SQL = SQL & " [OBS] [nvarchar](max) NOT NULL,"
         SQL = SQL & " [TIPO_REGISTRO] [nvarchar](max) NULL"
         SQL = SQL & " ) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[OBS]  WITH CHECK ADD  CONSTRAINT [FK_OBS_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
         SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[OBS] CHECK CONSTRAINT [FK_OBS_PEDIDO]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[OBS]  WITH CHECK ADD  CONSTRAINT [FK_OBS_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[OBS] CHECK CONSTRAINT [FK_OBS_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdRG_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "RG", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "RG") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE RG DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "RG") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE RG ADD PESSOA_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "RG") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close
         CONT_N = 0
         SQL = "select distinct(prop) from RG where pessoa_id is null "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            CONT_N = CONT_N + 1
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(TabTemp.Fields("prop").Value) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPessoa.EOF Then
               SQL = "update RG set pessoa_id = " & TabPessoa.Fields(0).Value
               SQL = SQL & " where prop = '" & Trim(TabTemp.Fields("prop").Value) & "'"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            cmdRg.Caption = "RG = " & Trim(TabTemp.Fields("prop").Value)
            DoEvents

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
'so pode excluir quando rodar todas rotinas
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "rg") = True Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE RG DROP COLUMN PROP"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_RG_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[RG]  WITH CHECK ADD  CONSTRAINT [FK_RG_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[RG] CHECK CONSTRAINT [FK_RG_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdNF_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "NF", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cfop", "nf") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE nf DROP COLUMN cfop"

      'não servia pra nada
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_ESPECIE", "nf") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE nf DROP COLUMN TIPO_ESPECIE"

'so pode excluir quando rodar todas rotinas
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROP", "nf") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE nf DROP COLUMN PROP"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "NF") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD PESSOA_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_VOLUME", "NF") = True Then _
         Alteração_Definição_Campo_Tabela "QTD_VOLUME", "FLOAT", "NF", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESO_BRUTO", "NF") = True Then _
         Alteração_Definição_Campo_Tabela "PESO_BRUTO", "FLOAT", "NF", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESO_LIQUIDO", "NF") = True Then _
         Alteração_Definição_Campo_Tabela "PESO_LIQUIDO", "FLOAT", "NF", "RETAGUARDA"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "NF") = False Then
      '   CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD PEDIDO_ID BIGINT"
      '   Else: Alteração_Definição_Campo_Tabela "PEDIDO_ID", "BIGINT", "NF", "RETAGUARDA"
      'End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NF_ID", "NF") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD NF_ID BIGINT NOT NULL"
         Else: Alteração_Definição_Campo_Tabela "NF_ID", "BIGINT NOT NULL", "NF", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TRANSP_ID", "NF") = True Then _
         Alteração_Definição_Campo_Tabela "TRANSP_ID", "BIGINT", "NF", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_REQ", "NF") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF DROP COLUMN NUMR_REQ"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "NF") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD ESTABELECIMENTO_ID INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "indPres", "NF") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD indPres INT"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "idDest", "NF") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD idDest INT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_NF", "") = False Then
         SQL = "ALTER TABLE NF ADD CONSTRAINT pk_NF PRIMARY KEY (NF_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MODELO_DOC", "NF") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD MODELO_DOC nvarchar(3)"
         Else: Alteração_Definição_Campo_Tabela "MODELO_DOC", "nvarchar(3)", "NF", "RETAGUARDA"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NF_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[NF]  WITH CHECK ADD  CONSTRAINT [FK_NF_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[NF] CHECK CONSTRAINT [FK_NF_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "empresa_id", "nf") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE nf DROP COLUMN empresa_id"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NF_TRANSPORTADORA", "") = False Then
         SQL = "ALTER TABLE [dbo].[NF]  WITH CHECK ADD  CONSTRAINT [FK_NF_TRANSPORTADORA] FOREIGN KEY([TRANSP_ID])"
         SQL = SQL & " References [dbo].[TRANSPORTADORA]([TRANSP_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[NF] CHECK CONSTRAINT [FK_NF_TRANSPORTADORA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NF_ESTABELECIMENTO", "") = False Then
         SQL = "update NF set estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[NF]  WITH CHECK ADD  CONSTRAINT [FK_NF_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NF] CHECK CONSTRAINT [FK_NF_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "NF") = True Then
         If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDONF", "U") = True Then
            Dim TabTempLocal  As New ADODB.Recordset
            
            If TabTempLocal.State = 1 Then _
               TabTempLocal.Close

            SQL = "select pedido_id from PEDIDO "
            SQL = SQL & " order by pedido_id"
            TabTempLocal.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            While Not TabTempLocal.EOF
               NF_ID_N = 0
               PEDIDO_ID_N = 0 & TabTempLocal.Fields(0).Value

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select nf_id from NF "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then
                  NF_ID_N = 0 & TabConsulta.Fields(0).Value
DoEvents
cmdNF.Caption = NF_ID_N

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

                  If NF_ID_N > 0 Then
                     SQL = "insert into PEDIDONF "
                        SQL = SQL & " (PEDIDONF_ID,PEDIDO_ID,NF_ID)"
                     SQL = SQL & " values("
                        SQL = SQL & MAX_ID("PEDIDONF_id", "PEDIDONF", "", "", "", "")
                        SQL = SQL & "," & PEDIDO_ID_N
                        SQL = SQL & "," & NF_ID_N
                     SQL = SQL & " )"

                     CONECTA_RETAGUARDA.Execute SQL
                  End If
               End If
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               TabTempLocal.MoveNext
            Wend
            If TabTempLocal.State = 1 Then _
               TabTempLocal.Close

            CONECTA_RETAGUARDA.Execute "ALTER TABLE nf DROP COLUMN pedido_id"

            Else: MsgBox "Criar tabela PEDIDONF"
         End If
      End If
   End If
'''''''''''''''''''''
   If EXISTE_OBJ_BANCO("RETAGUARDA", "NFITEM", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_ID", "NFITEM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NFITEM add SEQ_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_prod", "NFitem") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NFitem DROP COLUMN codg_prod"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NF_ID", "NF") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NFITEM ADD NF_ID BIGINT NOT NULL"
         Else: Alteração_Definição_Campo_Tabela "NF_ID", "BIGINT NOT NULL", "NFITEM", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP", "nfitem") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'nfitem.cfop'" & "," & "'CFOP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP_ID", "nfitem") = True Then _
         Alteração_Definição_Campo_Tabela "CFOP_ID", "nvarchar (10)", "nfitem", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_NFITEM", "") = False Then
         SQL = "ALTER TABLE NFITEM ADD CONSTRAINT pk_NFITEM PRIMARY KEY (NF_ID,SEQ_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NFITEM_NF", "") = True Then
         SQL = "ALTER TABLE NFITEM drop CONSTRAINT FK_NFITEM_NF"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NF_ID", "NF") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NFITEM ADD NF_ID BIGINT NOT NULL"
         Else: Alteração_Definição_Campo_Tabela "NF_ID", "BIGINT NOT NULL", "NFITEM", "RETAGUARDA"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NFITEM_NF", "") = False Then
         SQL = "ALTER TABLE [dbo].[NFITEM] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_NFITEM_NF] "
         SQL = SQL & " FOREIGN KEY([NF_ID])"
         SQL = SQL & " References [dbo].[NF]([NF_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NFITEM] CHECK CONSTRAINT [FK_NFITEM_NF]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      '''''''''''relação itens com produto
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUTO_ID", "NFITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NFITEM add PRODUTO_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PRODUTO_ID", "BIGINT", "NFITEM", "RETAGUARDA"
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_NFITEM_PRODUTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[NFITEM] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_NFITEM_PRODUTO] "
         SQL = SQL & " FOREIGN KEY([PRODUTO_ID])"
         SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[NFITEM] CHECK CONSTRAINT [FK_NFITEM_PRODUTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdDescritor_Click()
On Error Resume Next

   If EXISTE_OBJ_BANCO("RETAGUARDA", "DESCR", "") = True Then
      MsgBox "ALTERAR NOME DO CAMPO CODIGO PARAR DESCR_ID TABELA DESCR E TIPO DE DADOS BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO", "DESCR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'DESCR.CODIGO'" & "," & "'CODIGO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_a", "DESCR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'DESCR.TIPO_a'" & "," & "'TIPO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO", "DESCR") = True Then _
         Alteração_Definição_Campo_Tabela "TIPO", "nvarchar(2) NOT NULL", "DESCR", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESC_A", "DESCR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'DESCR.DESC_A'" & "," & "'DESCRICAO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESCRICAO", "DESCR") = True Then _
         Alteração_Definição_Campo_Tabela "DESCRICAO", "nvarchar(max) NOT NULL", "DESCRICAO", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_DESCR", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE DESCR DROP CONSTRAINT IX_DESCR "

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_DESCR", "") = False Then
         SQL = "ALTER TABLE DESCR ADD CONSTRAINT pk_DESCR PRIMARY KEY (CODIGO,TIPO)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      SQL = "delete from DESCR where "
      SQL = SQL & " TIPO = 'K' "
      SQL = SQL & " or TIPO = 'Y' "
      SQL = SQL & " or TIPO = 'D' "
      SQL = SQL & " or TIPO = 'F' "
      SQL = SQL & " or TIPO = 'P'"
      SQL = SQL & " or TIPO = 'L'"
      SQL = SQL & " or TIPO = 'E'"
      SQL = SQL & " or TIPO = 'Q'"
      SQL = SQL & " or TIPO = 'G'"
      SQL = SQL & " or TIPO = 'Q'"
      SQL = SQL & " or TIPO = 'L'"
      SQL = SQL & " or TIPO = 'C'"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'E'
      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'E'"
         SQL = SQL & ",'Simples Nacional'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 2
         SQL = SQL & ",'E'"
         SQL = SQL & ",'Simples Nacional-Excesso de sublimite receita bruta'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 3
         SQL = SQL & ",'E'"
         SQL = SQL & ",'Regime Normal - RPA'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'Q'
      SQL = "insert into DESCR values("
         SQL = SQL & 0
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Nacional'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Estrangeira - Importação Direta'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 2
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Estrangeira - Adquirida no mercado interno'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 3
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Nacional, mercadoria ou bem com Conteúdo de Importação superior a 40%'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 4
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Nacional, cuja produção tenha sido feita em conformidade com os processos'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 5
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Nacional, mercadoria ou bem com Conteúdo de Importação inferior ou igual a 40%'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 6
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Estrangeira - Importação direta, sem similar nacionaL'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 7
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Estrangeira - Adquirida no mercado interno, sem similar nacional'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 8
         SQL = SQL & ",'Q'"
         SQL = SQL & ",'Produto nacional com conteúdo importado acima de 70%'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'K'
      SQL = "insert into DESCR values("
         SQL = SQL & 0
         SQL = SQL & ",'K'"
         SQL = SQL & ",'Não se aplica'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'K'"
         SQL = SQL & ",'OPERAÇÃO PRESENCIAL'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 2
         SQL = SQL & ",'K'"
         SQL = SQL & ",'OPERAÇÃO NÃO PRESENCIAL, PELA INTERNET'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 3
         SQL = SQL & ",'K'"
         SQL = SQL & ",'Operação não presencial, Teleatendimento'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 4
         SQL = SQL & ",'K'"
         SQL = SQL & ",'NFC-e em operação com entrega em domicílio'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 5
         SQL = SQL & ",'K'"
         SQL = SQL & ",'Operação não presencial, outros'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'Y'
      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'Y'"
         SQL = SQL & ",'Operação interna'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 2
         SQL = SQL & ",'Y'"
         SQL = SQL & ",'Operação interestadual'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 3
         SQL = SQL & ",'Y'"
         SQL = SQL & ",'Operação com exterior'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 4
         SQL = SQL & ",'Y'"
         SQL = SQL & ",'Operação interna'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'P'
      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'P'"
         SQL = SQL & ",'Ativo'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 2
         SQL = SQL & ",'P'"
         SQL = SQL & ",'Cancelado'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 3
         SQL = SQL & ",'P'"
         SQL = SQL & ",'Promoção'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'F'
      SQL = "insert into DESCR values("
         SQL = SQL & 0
         SQL = SQL & ",'F'"
         SQL = SQL & ",'Matéria Prima'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'F'"
         SQL = SQL & ",'Produto Acabado'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'D'
      'Finalidade de emissão da NF-e
      SQL = "insert into DESCR values("
         SQL = SQL & 1
         SQL = SQL & ",'D'"
         SQL = SQL & ",'NF-e normal'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 2
         SQL = SQL & ",'D'"
         SQL = SQL & ",'NF-e complementar'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 3
         SQL = SQL & ",'D'"
         SQL = SQL & ",'NF-e de ajuste'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & 4
         SQL = SQL & ",'D'"
         SQL = SQL & ",'Devolução/Retorno'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
   
'===TIPO 'G'   BANDEIRA CARTÃO DE CREDITO E DEBITO
      SQL = "insert into DESCR values("
         SQL = SQL & "01"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'VISA'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "02"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Mastercard'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "03"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'American Express'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "04"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Sorocred'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "05"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Diners Club'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "06"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Elo'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "07"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Hipercard'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "08"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Aura'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "09"
         SQL = SQL & ",'G'"
         SQL = SQL & ",'Cabal'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

'===TIPO 'L'   tipo frete
      SQL = "insert into DESCR values("
         SQL = SQL & "0"
         SQL = SQL & ",'L'"
         SQL = SQL & ",'Contratação do Frete por conta do Remetente (CIF)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "1"
         SQL = SQL & ",'L'"
         SQL = SQL & ",'Contratação do Frete por conta do Destinatário (FOB)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "2"
         SQL = SQL & ",'L'"
         SQL = SQL & ",'Contratação do Frete por conta de Terceiros'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "3"
         SQL = SQL & ",'L'"
         SQL = SQL & ",'Transporte Próprio por conta do Remetente'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "4"
         SQL = SQL & ",'L'"
         SQL = SQL & ",'Transporte Próprio por conta do Destinatário'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "9"
         SQL = SQL & ",'L'"
         SQL = SQL & ",'Sem Ocorrência de Transporte'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
'===TIPO 'A1'   TIPO EMISSÃO NFE/NFC-E
      SQL = "insert into DESCR values("
         SQL = SQL & "1"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Emissão normal (não em contingência)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "2"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência FS-IA, com impressão do DANFE em formulário de segurança'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "3"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência SCAN (Sistema de Contingência do Ambiente Nacional)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "4"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência DPEC (Declaração Prévia da Emissão em Contingência)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "5"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência FS-DA, com impressão do DANFE em formulário de segurança'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "6"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência SVC-AN (SEFAZ Virtual de Contingência do AN)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
            
      SQL = "insert into DESCR values("
         SQL = SQL & "7"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência SVC-RS (SEFAZ Virtual de Contingência do RS)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into DESCR values("
         SQL = SQL & "9"
         SQL = SQL & ",'A1'"
         SQL = SQL & ",'Contingência off-line da NFC-e (as demais opções de contingência são válidas também para a NFC-e)'"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdFatura_Click()
   ATUALIZA_TABELA_FORMAPAGTO
   MsgBox "Ok   "
End Sub

Private Sub cmdVENDEDOR_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "equipe", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "EQUIPE") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EQUIPE ADD PESSOA_ID BIGINT"

         If TabEQUIPE.State = 1 Then _
            TabEQUIPE.Close

         SQL = "select PESSOA_ID from PESSOA WITH (NOLOCK)"
         SQL = SQL & " where cnpjcpf = '" & Trim(CNPJ_EMPRESA_N) & "'"
         TabEQUIPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabEQUIPE.EOF Then _
            If Not IsNull(TabEQUIPE.Fields(0).Value) Then _
               PESSOA_ID_N = 0 & TabEQUIPE.Fields(0).Value
         If TabEQUIPE.State = 1 Then _
            TabEQUIPE.Close

         SQL = "update EQUIPE set pessoa_id = " & PESSOA_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_EQUIPE_PESSOA", "") = False Then
            SQL = "ALTER TABLE [dbo].[EQUIPE]  WITH CHECK ADD  CONSTRAINT [FK_EQUIPE_PESSOA] FOREIGN KEY([PESSOA_ID])"
            SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
            CONECTA_RETAGUARDA.Execute SQL

            SQL = " ALTER TABLE [dbo].[EQUIPE] CHECK CONSTRAINT [FK_EQUIPE_PESSOA]"
            CONECTA_RETAGUARDA.Execute SQL
         End If
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "EQUIPE") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EQUIPE ADD ESTABELECIMENTO_ID INT"
         SQL = "update EQUIPE set estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_EQUIPE_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[EQUIPE]  WITH CHECK ADD  CONSTRAINT [FK_EQUIPE_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[EQUIPE] CHECK CONSTRAINT [FK_EQUIPE_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cgccpf", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE equipe DROP COLUMN cgccpf"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE equipe DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_eq", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'equipe.codg_eq'" & "," & "'EQUIPE_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CATEGORIA", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE equipe DROP COLUMN CATEGORIA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_COMIS", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE equipe DROP COLUMN TIPO_COMIS"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CGC_EQUIPE", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE equipe DROP COLUMN CGC_EQUIPE"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_COMISSAO", "equipe") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE equipe DROP COLUMN PERC_COMISSAO"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_EQUIPE", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EQUIPE ADD CONSTRAINT pk_EQUIPE PRIMARY KEY (EQUIPE_ID)"
      Else
         SQL = "CREATE TABLE [dbo].[EQUIPE]("
         SQL = SQL & " [EQUIPE_ID] [int] NOT NULL,"
         SQL = SQL & " [PESSOA_ID] [bigint] NULL,"
         SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
         SQL = SQL & " [DESCRICAO] [nvarchar](50) NOT NULL,"
         SQL = SQL & " [RESPONSAVEL] [nvarchar](50) NOT NULL,"
         SQL = SQL & " [STATUS] [nvarchar](1) NOT NULL,"
         SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
         SQL = SQL & " [DT_BAIXA] [datetime] NULL,"
         SQL = SQL & " CONSTRAINT [pk_EQUIPE] PRIMARY KEY CLUSTERED ([EQUIPE_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
         SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[EQUIPE]  WITH CHECK ADD  CONSTRAINT [FK_EQUIPE_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[EQUIPE] CHECK CONSTRAINT [FK_EQUIPE_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[EQUIPE]  WITH CHECK ADD  CONSTRAINT [FK_EQUIPE_PESSOA] FOREIGN KEY([PESSOA_ID]) References [dbo].[PESSOA]([Pessoa_Id])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[EQUIPE] CHECK CONSTRAINT [FK_EQUIPE_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "VENDEDOR", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "VENDEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "dt_cad", "VENDEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR DROP COLUMN dt_cad"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_EQ", "VENDEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'VENDEDOR.CODG_EQ'" & "," & "'EQUIPE_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "VENDEDOR") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR ADD PESSOA_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT", "VENDEDOR", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_vend", "VENDEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'VENDEDOR.codg_vend'" & "," & "'VENDEDOR_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VENDEDOR_ID", "VENDEDOR") = True Then _
         Alteração_Definição_Campo_Tabela "VENDEDOR_ID", "BIGINT", "VENDEDOR", "RETAGUARDA"
   
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_VENDEDOR_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[VENDEDOR]  WITH CHECK ADD  CONSTRAINT [FK_VENDEDOR_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[VENDEDOR] CHECK CONSTRAINT [FK_VENDEDOR_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'===================================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CPF", "VENDEDOR") = True Then
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
CONT_N = 0
         SQL = "select * from VENDEDOR "
         'SQL = SQL & " where pessoa_id is null"
         SQL = SQL & " ORDER BY PESSOA_ID"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabConsulta.EOF
         CONT_N = CONT_N + 1
            If Not IsNull(TabConsulta.Fields("nome_vend").Value) Then
               cmdVendedor.Caption = Trim(TabConsulta.Fields("nome_vend").Value)
               CNPJCPF_A = "" & TabConsulta.Fields("CPF").Value
               NOME_A = Trim(TabConsulta.Fields("nome_vend").Value)
               RAZAO_A = "" & Trim(NOME_A)
               DT_EXP_D = Date
               STATUS_A = "" & Trim(TabConsulta.Fields("status").Value)

               If Trim(RAZAO_A) = "" Then _
                  RAZAO_A = NOME_A

               If Trim(STATUS_A) = "" Then _
                  STATUS_A = "C"

               frmATUALIZACAO2.Caption = "CRIANDO PESSOA = " & Trim(NOME_A)
               DoEvents

               'se é nulo buscar por cnpj
               If IsNull(TabConsulta.Fields("pessoa_id").Value) Then
                  INDR_PRI = False

                  If TabPessoa.State = 1 Then _
                     TabPessoa.Close

                  SQL = "select * from PESSOA"
                  SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
                  TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabPessoa.EOF Then _
                     INDR_PRI = True

                  If TabPessoa.State = 1 Then _
                     TabPessoa.Close

                  If INDR_PRI = True Then
                     frmATUALIZACAO2.Caption = "CRIANDO PESSOA = " & Trim(CNPJCPF_A)
                     DoEvents

                     spPessoa 1, _
                               0, _
                               Trim(CNPJCPF_A), _
                               Trim(NOME_A), _
                               Trim(RAZAO_A), _
                               STATUS_A
                     INDR_PRI = False
                  End If
                  Else                 'se NÃO é nulo verifica se já está vinculado id corretamente
                     If TabPessoa.State = 1 Then _
                        TabPessoa.Close

                     SQL = "select pessoa_id from PESSOA"
                     SQL = SQL & " where pessoa_id = " & TabConsulta.Fields("pessoa_id").Value
                     TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If TabPessoa.EOF Then
                        frmATUALIZACAO2.Caption = "CRIANDO PESSOA = " & Trim(CNPJCPF_A)
                        DoEvents

                        spPessoa 1, _
                                  0, _
                                  Trim(CNPJCPF_A), _
                                  Trim(NOME_A), _
                                  Trim(RAZAO_A), _
                                  Trim(TabConsulta.Fields("status").Value)
                        CONT_N = CONT_N + 1
                     End If

                     If TabPessoa.State = 1 Then _
                        TabPessoa.Close
               End If
            End If
            CONT_N = CONT_N + 1

            If TabPessoa.State = 1 Then _
               TabPessoa.Close
                  SQL = "select pessoa_id from PESSOA"
                  SQL = SQL & " where cnpjcpf = '" & Trim(TabConsulta.Fields("CPF").Value) & "'"
                  TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabPessoa.EOF Then
                     SQL = "update VENDEDOR set pessoa_id = " & TabPessoa.Fields(0).Value
                     SQL = SQL & " where CPF = '" & Trim(TabConsulta.Fields("CPF").Value) & "'"
                     CONECTA_RETAGUARDA.Execute SQL
                  End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            TabConsulta.MoveNext

            frmATUALIZACAO2.Caption = "VENDEDOR = " & CONT_N
            DoEvents
         Wend
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NOME_VEND", "VENDEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR DROP COLUMN NOME_VEND"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CPF", "VENDEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR DROP COLUMN cpf"


      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_VENDEDOR", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR ADD CONSTRAINT pk_VENDEDOR PRIMARY KEY (VENDEDOR_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTABVENDEDOR", "") = False Then
         SQL = "CREATE TABLE [dbo].[ESTABVENDEDOR]("
         SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
         SQL = SQL & " [VENDEDOR_ID] [BIGInt] NOT NULL"
         SQL = SQL & " ) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ESTABVENDEDOR]  WITH CHECK ADD  CONSTRAINT [FK_ESTABVENDEDOR_ESTAB] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ESTABVENDEDOR] CHECK CONSTRAINT [FK_ESTABVENDEDOR_ESTAB]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ESTABVENDEDOR]  WITH CHECK ADD  CONSTRAINT [FK_ESTABVENDEDOR_VENDEDOR] FOREIGN KEY([VENDEDOR_ID])"
         SQL = SQL & " References [dbo].[VENDEDOR]([VENDEDOR_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ESTABVENDEDOR] CHECK CONSTRAINT [FK_ESTABVENDEDOR_VENDEDOR]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'===================================
   End If
   MsgBox "Ok"
End Sub

Private Sub cmdTabela_Click()
   
   MATA_TABELAS
   
   MsgBox "Ok  =  " & CONT_N
End Sub



Private Sub cmdFamilia_Click()

   ATUALIZA_TABELA_FAMILIAPRODUTO

MsgBox "OK"
End Sub

Private Sub cmdProduto_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUTO", "U") = True Then
      CHECA_TABELA_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select familiaproduto_id,producao from FAMILIAPRODUTO "
      SQL = SQL & " where producao = 1 "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabProduto.EOF

         If Not IsNull(TabProduto.Fields(1).Value) Then
            SQL = "update produto set "

            If TabProduto.Fields(1).Value = True Then
               SQL = SQL & " conceder_producao = 1"
               Else: SQL = SQL & " conceder_producao = 0"
            End If
            SQL = SQL & " where familiaproduto_id = " & TabProduto.Fields(0).Value
            CONECTA_RETAGUARDA.Execute SQL
         End If

         TabProduto.MoveNext
      Wend
      If TabProduto.State = 1 Then _
         TabProduto.Close

      Msg = " Confirma rodar produto_id ?"
      PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then

         CONT_N = 0

         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select codg_produto from PRODUTO "
         SQL = SQL & " where PRODUTO_id Is Null Or PRODUTO_id <= 0"
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabProduto.EOF
            SQL = "update PRODUTO set "
            SQL = SQL & " PRODUTO_ID = " & CONT_N
            SQL = SQL & " where codg_produto = '" & Trim(TabProduto.Fields(0).Value) & "'"
            CONECTA_RETAGUARDA.Execute SQL

            CONT_N = CONT_N + 1
            frmATUALIZACAO2.Caption = Trim(TabProduto.Fields(0).Value) & " / " & CONT_N

            TabProduto.MoveNext
            DoEvents
         Wend

         If TabProduto.State = 1 Then _
            TabProduto.Close
      End If
   End If

INDR_PRI = False
If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUTOFORNECEDOR", "U") = True Then
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTOFORNECEDOR"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then _
      INDR_PRI = True
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If INDR_PRI = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE PRODUTOFORNECEDOR"
End If
INDR_PRI = False

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUTOFORNECEDOR", "U") = False Then
      SQL = "CREATE TABLE [dbo].[PRODUTOFORNECEDOR]("
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [FORNECEDOR_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CODG_PROD_FORNEC] [nvarchar](50) NOT NULL,"
      SQL = SQL & " [PRECO_CUSTO] [float] NOT NULL,"
      SQL = SQL & " [CODG_BARRA] [nvarchar](50) NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTOFORNECEDOR]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTOFORNECEDOR_FORNECEDOR] FOREIGN KEY([FORNECEDOR_ID])"
      SQL = SQL & " References [dbo].[FORNECEDOR]([FORNECEDOR_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTOFORNECEDOR] CHECK CONSTRAINT [FK_PRODUTOFORNECEDOR_FORNECEDOR]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTOFORNECEDOR]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTOFORNECEDOR_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTOFORNECEDOR] CHECK CONSTRAINT [FK_PRODUTOFORNECEDOR_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   MsgBox "Ok, ATENÇÃO CRIAR INDICE CODG_PRODUTO  =  " & CONT_N
End Sub

Private Sub cmdUsuario_Click()

   ATUALIZA_TABELA_USUARIO

   If EXISTE_OBJ_BANCO("RETAGUARDA", "USUARIO", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CPF", "USUARIO") = True Then
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
CONT_N = 0
         SQL = "select * from USUARIO "
         'SQL = SQL & " where pessoa_id is null"
         SQL = SQL & " ORDER BY PESSOA_ID"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabConsulta.EOF
         CONT_N = CONT_N + 1
            If Not IsNull(TabConsulta.Fields("NOME").Value) Then
               cmdUsuario.Caption = Trim(TabConsulta.Fields("NOME").Value)
               CNPJCPF_A = "" & TabConsulta.Fields("CPF").Value
               NOME_A = Trim(TabConsulta.Fields("NOME").Value)
               RAZAO_A = "" & Trim(NOME_A)
               DT_EXP_D = Date
               STATUS_A = "" & Trim(TabConsulta.Fields("status").Value)

               If Trim(RAZAO_A) = "" Then _
                  RAZAO_A = NOME_A

               If Trim(STATUS_A) = "" Then _
                  STATUS_A = "C"

               frmATUALIZACAO2.Caption = "CRIANDO PESSOA = " & Trim(NOME_A)
               DoEvents

               'se é nulo buscar por cnpj
               If IsNull(TabConsulta.Fields("pessoa_id").Value) Then
                  INDR_PRI = False

                  If TabPessoa.State = 1 Then _
                     TabPessoa.Close

                  SQL = "select * from PESSOA"
                  SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
                  TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabPessoa.EOF Then _
                     INDR_PRI = True

                  If TabPessoa.State = 1 Then _
                     TabPessoa.Close

                  If INDR_PRI = True Then
                     frmATUALIZACAO2.Caption = "CRIANDO PESSOA = " & Trim(CNPJCPF_A)
                     DoEvents

                     spPessoa 1, _
                               0, _
                               Trim(CNPJCPF_A), _
                               Trim(NOME_A), _
                               Trim(RAZAO_A), _
                               STATUS_A
                     INDR_PRI = False
                  End If
                  Else                 'se NÃO é nulo verifica se já está vinculado id corretamente
                     If TabPessoa.State = 1 Then _
                        TabPessoa.Close

                     SQL = "select pessoa_id from PESSOA"
                     SQL = SQL & " where pessoa_id = " & TabConsulta.Fields("pessoa_id").Value
                     TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If TabPessoa.EOF Then
                        frmATUALIZACAO2.Caption = "CRIANDO PESSOA = " & Trim(CNPJCPF_A)
                        DoEvents

                        spPessoa 1, _
                                  0, _
                                  Trim(CNPJCPF_A), _
                                  Trim(NOME_A), _
                                  Trim(RAZAO_A), _
                                  Trim(TabConsulta.Fields("status").Value)
                        CONT_N = CONT_N + 1
                     End If

                     If TabPessoa.State = 1 Then _
                        TabPessoa.Close
               End If
            End If
            CONT_N = CONT_N + 1

            If TabPessoa.State = 1 Then _
               TabPessoa.Close
                  SQL = "select pessoa_id from PESSOA"
                  SQL = SQL & " where cnpjcpf = '" & Trim(TabConsulta.Fields("CPF").Value) & "'"
                  TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabPessoa.EOF Then
                     SQL = "update USUARIO set pessoa_id = " & TabPessoa.Fields(0).Value
                     SQL = SQL & " where CPF = '" & Trim(TabConsulta.Fields("CPF").Value) & "'"
                     CONECTA_RETAGUARDA.Execute SQL
                  End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            TabConsulta.MoveNext

            frmATUALIZACAO2.Caption = "USUARIO = " & CONT_N
            DoEvents
         Wend
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "obs", "USUARIO") = True Then
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_REGISTRO", "OBS") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OBS ADD TIPO_REGISTRO NVARCHAR(MAX)"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_USUARIO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO ADD CONSTRAINT pk_USUARIO PRIMARY KEY (USUARIO_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_USUARIO_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[USUARIO]  WITH CHECK ADD  CONSTRAINT [FK_USUARIO_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[USUARIO] CHECK CONSTRAINT [FK_USUARIO_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   MsgBox "Ok, =  " & CONT_N
End Sub

Private Sub cmdTabelaPreco_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "TABELAPRECO", "U") = True Then
      Else
         SQL = "CREATE TABLE [dbo].[TABELAPRECO]("
         SQL = SQL & " [TABELAPRECO_ID] [int] NOT NULL,"
         SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
         SQL = SQL & " [CODG_TABELA] [nvarchar](30) NOT NULL,"
         SQL = SQL & " [DESCRICAO] [nvarchar](60) NOT NULL,"
         SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
         SQL = SQL & " [DT_VALIDADE] [datetime] NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_TABELAPRECO] PRIMARY KEY CLUSTERED("
         SQL = SQL & " [TABELAPRECO_ID] Asc )"
         SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
         SQL = SQL & " ON [PRIMARY] ) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECO]  WITH CHECK ADD  "
         SQL = SQL & " CONSTRAINT [FK_TABELAPRECO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECO] CHECK CONSTRAINT [FK_TABELAPRECO_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "TABELAPRECOITEM", "U") = True Then
      Else
         SQL = "CREATE TABLE [dbo].[TABELAPRECOITEM]("
         SQL = SQL & " [TABELAPRECO_ID] [int] NOT NULL,"
         SQL = SQL & " [TABELAPRECOITEM_ID] [int] NOT NULL,"
         SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
         SQL = SQL & " [FORMAPAGTO_ID] [Int] NOT NULL,"
         SQL = SQL & " [VALOR_VENDA] [float] NOT NULL,"
         SQL = SQL & " [VALOR_CUSTO] [float] NOT NULL,"
         SQL = SQL & " [PERC_COMISSAO] [float] NULL,"
         SQL = SQL & " CONSTRAINT [PK_TABELAPRECOITEM] PRIMARY KEY CLUSTERED("
         
         SQL = SQL & " [TABELAPRECO_ID] ASC,[PRODUTO_ID] ASC,[FORMAPAGTO_ID] ASC)"
         SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
         SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECOITEM]  "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_TABELAPRECOITEM_TABELAPRECO] "
         SQL = SQL & " FOREIGN KEY([TABELAPRECO_ID])"
         SQL = SQL & " References [dbo].[TABELAPRECO]([TABELAPRECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECOITEM] CHECK CONSTRAINT [FK_TABELAPRECOITEM_TABELAPRECO]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECOITEM]  WITH CHECK ADD  CONSTRAINT [FK_TABELAPRECOITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
         SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECOITEM] CHECK CONSTRAINT [FK_TABELAPRECOITEM_PRODUTO]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECOITEM]  WITH CHECK ADD  CONSTRAINT [FK_TABELAPRECOITEM_FORMAPAGTO] FOREIGN KEY([FORMAPAGTO_ID])"
         SQL = SQL & " References [dbo].[FORMAPAGTO]([FORMAPAGTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TABELAPRECOITEM] CHECK CONSTRAINT [FK_TABELAPRECOITEM_FORMAPAGTO]"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "VENDEDOR", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TABELAPRECO_ID", "VENDEDOR") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE VENDEDOR ADD TABELAPRECO_ID INT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_VENDEDOR_TABELAPRECO", "") = False Then
         SQL = " ALTER TABLE [dbo].[VENDEDOR]  "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_VENDEDOR_TABELAPRECO] "
         SQL = SQL & " FOREIGN KEY([TABELAPRECO_ID])"
         SQL = SQL & " References [dbo].[TABELAPRECO]([TABELAPRECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[VENDEDOR] CHECK CONSTRAINT [FK_VENDEDOR_TABELAPRECO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDO", "U") = True Then _
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TABELAPRECO_ID", "PEDIDO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD TABELAPRECO_ID INT "

   MsgBox "ok"
End Sub

Private Sub cmdInventario_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "INVENTARIO", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "empresa_id", "INVENTARIO") = True Then
         MsgBox "excluir campo empresa_id tabela inventário"
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_ANTERIOR", "INVENTARIO") = True Then _
         Alteração_Definição_Campo_Tabela "QTD_ANTERIOR", "NUMERIC(18,3)", "INVENTARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_PRIMEIRA", "INVENTARIO") = True Then _
         Alteração_Definição_Campo_Tabela "QTD_PRIMEIRA", "NUMERIC(18,3)", "INVENTARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_SEGUNDA", "INVENTARIO") = True Then _
         Alteração_Definição_Campo_Tabela "QTD_SEGUNDA", "NUMERIC(18,3)", "INVENTARIO", "RETAGUARDA"
      
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_ATUAL", "INVENTARIO") = True Then _
         Alteração_Definição_Campo_Tabela "QTD_ATUAL", "NUMERIC(18,3)", "INVENTARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_LOTE", "INVENTARIO") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_LOTE", "BIGINT", "INVENTARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "INVENTARIO") = True Then _
         Alteração_Definição_Campo_Tabela "SEQ", "BIGINT", "INVENTARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_MOV", "INVENTARIO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INVENTARIO ADD TIPO_MOV CHAR(2)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_PROD", "INVENTARIO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE inventario DROP COLUMN codg_prod"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUTO_ID", "INVENTARIO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INVENTARIO ADD PRODUTO_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "INVENTARIO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INVENTARIO ADD ESTABELECIMENTO_ID INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "REGISTRO", "INVENTARIO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INVENTARIO ADD REGISTRO CHAR(1)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_INVENTARIO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INVENTARIO ADD CONSTRAINT pk_INVENTARIO PRIMARY KEY (NUMR_LOTE,SEQ)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_INVENTARIO_PRODUTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[INVENTARIO]  WITH CHECK ADD  CONSTRAINT [FK_INVENTARIO_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
         SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[INVENTARIO] CHECK CONSTRAINT [FK_INVENTARIO_PRODUTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_INVENTARIO_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[INVENTARIO]  WITH CHECK ADD  CONSTRAINT [FK_INVENTARIO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[INVENTARIO] CHECK CONSTRAINT [FK_INVENTARIO_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   MsgBox "Ok"
End Sub

Private Sub Command1_Click()
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from pedido where TABELAPRECO_ID <= 0 "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TABELAPRECO"
      SQL = SQL & " where estabelecimento_id = " & TabConsulta.Fields("estabelecimento_id").Value
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         SQL = "update PEDIDO set "
         SQL = SQL & " tabelapreco_id = " & TabTemp.Fields("tabelapreco_id").Value
         SQL = SQL & " where pedido_id = " & TabConsulta.Fields("pedido_id").Value
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      TabConsulta.MoveNext
      DoEvents
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
MsgBox "ok"
End Sub

Private Sub Command14_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTOQUE", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ESTOQUE]("
      SQL = SQL & " [ESTOQUE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE_ESTOQUE] [numeric(18, 3)] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_ESTOQUE] PRIMARY KEY CLUSTERED("
      SQL = SQL & " [ESTOQUE_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],"
      SQL = SQL & " CONSTRAINT [IX_ESTOQUE_PRODUTO] UNIQUE NONCLUSTERED("
      SQL = SQL & " [ESTABELECIMENTO_ID] ASC,[PRODUTO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE]  WITH CHECK ADD  CONSTRAINT [FK_ESTOQUE_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE] CHECK CONSTRAINT [FK_ESTOQUE_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE]  WITH CHECK ADD  CONSTRAINT [FK_ESTOQUE_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE] CHECK CONSTRAINT [FK_ESTOQUE_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   Msg = " Confirma rodar cadastro TABELA ESTOQUE ?"
   PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      RODA_AT_ESTOQUE 0, ESTABELECIMENTO_ID_N
      MsgBox "ok"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTOQUETRANSF", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ESTOQUETRANSF]("
      SQL = SQL & " [TRANSF_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_ID] [bigint] NOT NULL,"
      SQL = SQL & " [EMPRESA_ID] [int] NOT NULL,"
      SQL = SQL & " [ESTAB_ORIGEM_ID] [int] NOT NULL,"
      SQL = SQL & " [ESTAB_ORIGEM_DESC] [NVARCHAR](30) NOT NULL,"
      SQL = SQL & " [ESTAB_DESTINO_ID] [int] NOT NULL,"
      SQL = SQL & " [ESTAB_DESTINO_DESC] [NVARCHAR](30) NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE_TRANSF] [numeric](18, 3) NOT NULL,"
      SQL = SQL & " [DT_TRANSF] [datetime] NOT NULL,"
      SQL = SQL & " [DT_ENTRADA] [datetime] ,"
      SQL = SQL & " [SITUACAO] [nchar](1) NOT NULL,"

      SQL = SQL & " CONSTRAINT [PK_ESTOQUETRANSF] PRIMARY KEY CLUSTERED("
      SQL = SQL & " [TRANSF_ID] Asc,[SEQ_ID] ASC)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY] ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_ENTRADA", "ESTOQUETRANSF") = False Then
            CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTOQUETRANSF ADD DT_ENTRADA datetime "
         End If
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ESTOQUETRANSF_PRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[ESTOQUETRANSF]  WITH CHECK ADD  CONSTRAINT [FK_ESTOQUETRANSF_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUETRANSF] CHECK CONSTRAINT [FK_ESTOQUETRANSF_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If


'==============================
   'If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'SQL = "select * from ESTOQUETRANSF where transf_id = 1065 "

   'TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'While Not TabConsulta.EOF

'      SQL = "update estoque set "
 '     SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(TabConsulta.Fields("qtde_transf").Value)
  '    SQL = SQL & " where estabelecimento_id = " & TabConsulta.Fields("estab_origem_id").Value
   '   SQL = SQL & " and produto_id = " & TabConsulta.Fields("produto_id").Value
'MsgBox SQL
'      CONECTA_RETAGUARDA.Execute SQL

   '   TabConsulta.MoveNext
   'Wend
   'If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'SQL = "update estoquetransf set "
   'SQL = SQL & " situacao = 'C'"
   'SQL = SQL & " where transf_id = 1065"
   'CONECTA_RETAGUARDA.Execute SQL

'==============================
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

SQL = "select ESTAB_ORIGEM_id,ESTAB_destino_id,* from ESTOQUETRANSF WITH (NOLOCK)"
SQL = SQL & " Where ESTAB_ORIGEM_DESC Is Null Or ESTAB_destino_DESC Is Null"
TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
While Not TabConsulta.EOF

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select descricao from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & TabConsulta.Fields("estab_origem_id").Value
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update ESTOQUETRANSF set "
      SQL = SQL & " ESTAB_ORIGEM_DESC = '" & Trim(TabTemp.Fields(0).Value) & "'"
      SQL = SQL & " Where ESTAB_ORIGEM_DESC Is Null "
      SQL = SQL & " and   estab_origem_id = " & TabConsulta.Fields("estab_origem_id").Value
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select descricao from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & TabConsulta.Fields("estab_destino_id").Value
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update ESTOQUETRANSF set "
      SQL = SQL & " ESTAB_destino_DESC = '" & Trim(TabTemp.Fields(0).Value) & "'"
      SQL = SQL & " Where ESTAB_destino_DESC Is Null "
      SQL = SQL & " and   estab_destino_id = " & TabConsulta.Fields("estab_destino_id").Value
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabConsulta.MoveNext
Wend
End Sub

Private Sub cmdOS_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ospedido", "") = True Then
      SQL = "drop table ospedido"
      CONECTA_RETAGUARDA.Execute SQL
   End If

' ============================================================ */
'   Table: OSEQUIPAMENTO                                       */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_EQUIPAMENTO_PESSOA", "") = True Then
      SQL = "ALTER TABLE equipamento DROP CONSTRAINT FK_EQUIPAMENTO_PESSOA"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OS_EQUIPAME", "") = True Then
      SQL = "ALTER TABLE OS DROP CONSTRAINT FK_OS_EQUIPAME"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_VEICULO_EQUIPAME", "") = True Then
      SQL = "ALTER TABLE VEICULO DROP CONSTRAINT FK_VEICULO_EQUIPAME"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_EQUIPAMENTO", "") = True Then
      SQL = "ALTER TABLE EQUIPAMENTO DROP CONSTRAINT PK_EQUIPAMENTO"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "EQUIPAMENTO", "") = True Then
      SQL = "EXEC sp_rename " & "'EQUIPAMENTO'" & "," & "'OSEQUIPAMENTO'"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSEQUIPAMENTO", "") = False Then
      SQL = " create table OSEQUIPAMENTO "
      SQL = SQL & " ("
      SQL = SQL & " [EQUIPAMENTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PESSOA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [MARCA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [COR_ID] [bigint] NOT NULL,"
      SQL = SQL & " [TIPO_EQP] [bigint] NOT NULL,"
      SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](200) NOT NULL,"
      SQL = SQL & " [IDENTIFICACAO] [nvarchar](100) NOT NULL,"
      SQL = SQL & " [ANO] [bigint] NULL,"
      SQL = SQL & " [MODELO] [bigint] NULL,"
      SQL = SQL & " [NOME_CLIENTE] [nvarchar](100) NULL,"
      SQL = SQL & " constraINT PK_OSEQUIPAMENTO primary key (EQUIPAMENTO_ID))"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_OSEQUIPAMENTO", "") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OSEQUIPAMENTO ADD CONSTRAINT PK_OSEQUIPAMENTO PRIMARY KEY (EQUIPAMENTO_ID)"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSEQUIPAMENTO_PESSOA", "") = False Then
      SQL = " alter table OSEQUIPAMENTO "
      SQL = SQL & " add constraINT FK_OSEQUIPAMENTO_PESSOA foreign key (PESSOA_ID)"
      SQL = SQL & " References PESSOA(PESSOA_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   Table: VEICULO                                             */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_VEICULO", "") = True Then
      SQL = "ALTER TABLE VEICULO DROP CONSTRAINT PK_VEICULO"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "VEICULO", "") = True Then
      SQL = "EXEC sp_rename " & "'VEICULO'" & "," & "'OSVEICULO'"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSVEICULO", "") = False Then
      SQL = " create table OSVEICULO "
      SQL = SQL & " ("
      SQL = SQL & " VEICULO_ID      BIGINT         not null,"
      SQL = SQL & " COMBUSTIVEL_ID  BIGINT         not null,"
      SQL = SQL & " PLACA           NVARCHAR(10)   not null,"
      SQL = SQL & " DESCRICAO       NVARCHAR(100)  not null,"
      SQL = SQL & " MOTOR           NVARCHAR(100)  null    ,"
      SQL = SQL & " CHASSI          NVARCHAR(100)  null    ,"
      SQL = SQL & " NUMR_FROTA      NVARCHAR(10)   not null,"
      SQL = SQL & " constraINT PK_OSVEICULO primary key (VEICULO_ID))"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_OSVEICULO", "") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OSVEICULO ADD CONSTRAINT PK_OSVEICULO PRIMARY KEY (VEICULO_ID)"
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_FROTA", "OSVEICULO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE OSVEICULO ADD NUMR_FROTA NVARCHAR(10)"
' ============================================================ */
'   Table: OS                                                  */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OS_EQUIPAME", "") = True Then
      SQL = "ALTER TABLE OS DROP CONSTRAINT FK_OS_EQUIPAME"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OS_EMPRESA", "") = True Then
      SQL = "ALTER TABLE OS DROP CONSTRAINT FK_OS_EMPRESA"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "OS") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'OS.EMPRESA_ID'" & "," & "'ESTABELECIMENTO_ID'" & "," & "'COLUMN'"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "OS", "") = False Then
      SQL = " create table OS "
      SQL = SQL & " ("
      SQL = SQL & " OS_ID              BIGINT         not null,"
      SQL = SQL & " ESTABELECIMENTO_ID INT            not null,"
      SQL = SQL & " PESSOA_ID          BIGINT         not null,"
      SQL = SQL & " OSEQUIPAMENTO_ID   BIGINT         not null,"
      SQL = SQL & " CT_ID              BIGINT         not null,"
      SQL = SQL & " DT_OS              datetime       not null,"
      SQL = SQL & " DT_FECHA           datetime       null,"
      SQL = SQL & " TIPO_OS            BIGINT         not null,"
      SQL = SQL & " SITUACAO_OS        int            not null,"
      SQL = SQL & " KM                 numeric        not null,"
      SQL = SQL & " CLIENTE            NVARCHAR(50)   not null,"
      SQL = SQL & " constraINT PK_OS primary key (OS_ID))"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "OS") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE OS ADD PESSOA_ID BIGINT"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CLIENTE", "OS") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE OS ADD CLIENTE NVARCHAR(50)"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "SITUACAO_OS", "OS") = True Then _
      Alteração_Definição_Campo_Tabela "SITUACAO_OS", "INT", "OS", "RETAGUARDA"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OS_PESSOA", "") = False Then
      SQL = " alter table OS "
      SQL = SQL & " add constraINT FK_OS_PESSOA foreign key (PESSOA_ID)"
      SQL = SQL & " References PESSOA(PESSOA_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OS_ESTABELECIMENTO", "") = False Then
      SQL = " alter table OS "
      SQL = SQL & " add constraINT FK_OS_ESTABELECIMENTO foreign key (ESTABELECIMENTO_ID)"
      SQL = SQL & " References ESTABELECIMENTO(ESTABELECIMENTO_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   Table: OSAPONTAMENTO                                       */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSAPONTAMENTO", "") = False Then
      SQL = "CREATE TABLE [dbo].[OSAPONTAMENTO]("
      SQL = SQL & " [APONTAMENTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [OSTAREFA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DATAINICIAL] [datetime] NOT NULL,"
      SQL = SQL & " [DATAFINAL] [datetime] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_OSAPONTAMENTO] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([APONTAMENTO_ID] ASC)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSAPONTAMENTO]  WITH CHECK ADD  CONSTRAINT "
      SQL = SQL & " [FK_OSAPONTAMENTO_OSTAREFA] FOREIGN KEY([OSTAREFA_ID]) "
      SQL = SQL & " References [dbo].[OSTAREFA]([OSTAREFA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSAPONTAMENTO] CHECK CONSTRAINT [FK_OSAPONTAMENTO_OSTAREFA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   Table: OSTAREFA                                              */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSTAREFA", "") = False Then
      SQL = "create table OSTAREFA "
      SQL = SQL & " ("
      SQL = SQL & " OSTAREFA_ID     BIGINT        not null,"
      SQL = SQL & " DESCRICAO       NVARCHAR(200) not null,"
      SQL = SQL & " VALOR           float         not null,"
      SQL = SQL & " PERC_COMISSAO   float         not null,"
      SQL = SQL & " DT_CAD          datetime      not null,"
      SQL = SQL & " constraINT PK_OSTAREFA primary key (OSTAREFA_ID))"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   Table: OSPECA                                              */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSPECA", "") = False Then
      SQL = " create table OSPECA ("
      SQL = SQL & " OSPECA_ID          BIGINT         not null,"
      SQL = SQL & " OS_ID              BIGINT         not null,"
      SQL = SQL & " PRODUTO_ID         BIGINT         not null,"
      SQL = SQL & " DT_CAD             datetime       not null,"
      SQL = SQL & " SOLICITANTE_ID     BIGINT         not null,"
      SQL = SQL & " VALOR_ITEM         float          not null,"
      SQL = SQL & " DESCONTO_PRODUTO   float          not null,"
      SQL = SQL & " QTDE               numeric(18, 3) not null,"
      SQL = SQL & " DT_GARANTIA        datetime       null,"
      SQL = SQL & " constraINT PK_OSPECA primary key (OSPECA_ID))"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_GARANTIA", "OSPECA") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OSPECA ADD DT_GARANTIA DATETIME"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSPECA_OS", "") = False Then
      SQL = " alter table OSPECA "
      SQL = SQL & " add constraINT FK_OSPECA_OS foreign key  (OS_ID)"
      SQL = SQL & " References OS(OS_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSPECA_PRODUTO", "") = False Then
      SQL = " alter table OSPECA "
      SQL = SQL & " add constraINT FK_OSPECA_PRODUTO foreign key  (PRODUTO_ID)"
      SQL = SQL & " References PRODUTO(PRODUTO_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   Table: OSSERVICO                                           */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSSERVICO", "") = False Then
      SQL = " create table OSSERVICO ("
      SQL = SQL & " OS_ID              BIGINT         not null,"
      SQL = SQL & " OSSERVICO_ID       BIGINT         not null,"
      SQL = SQL & " OSTAREFA_ID        BIGINT         not null,"
      SQL = SQL & " DT_CAD             datetime       not null,"
      SQL = SQL & " DT_FIM           datetime       not null,"
      SQL = SQL & " DT_INICIO           datetime       not null,"
      SQL = SQL & " RESPONSAVEL_ID     BIGINT         not null,"
      SQL = SQL & " VALOR_SERVICO      float          not null,"
      SQL = SQL & " DESCRICAO          NVARCHAR(200)  not null,"
      SQL = SQL & " DESCONTO_SERVICO   float          not null,"
      SQL = SQL & " SITUACAO           CHAR(1)        not null,"
      SQL = SQL & " CONSTRAINT [PK_OSSERVICO] PRIMARY KEY CLUSTERED "
      SQL = SQL & " ([OS_ID] ASC, [OSSERVICO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY] ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "SITUACAO", "OSSERVICO") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OSSERVICO ADD SITUACAO CHAR(1)"

         If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_FIM", "OSSERVICO") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OSSERVICO ADD DT_FIM DATETIME"

         If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_INICIO", "OSSERVICO") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE OSSERVICO ADD DT_INICIO DATETIME"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSSERVICO_OS", "") = False Then
      SQL = " alter table OSSERVICO "
      SQL = SQL & " add constraINT FK_OSSERVICO_OS foreign key  (OS_ID)"
      SQL = SQL & " References OS(OS_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSSERVICO_OSTAREFA", "") = False Then
      SQL = " alter table OSSERVICO "
      SQL = SQL & " add constraINT FK_OSSERVICO_OSTAREFA foreign key  (OSTAREFA_ID)"
      SQL = SQL & " References OSTAREFA(OSTAREFA_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_FECHA", "OSSERVICO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE OSSERVICO ADD DT_FECHA DATETIME "

' ============================================================ */
' Table: OSTERMO                                               */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSTERMO", "") = False Then
      SQL = "CREATE TABLE [dbo].[OSTERMO]("
      SQL = SQL & " [OSTERMO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [OS_ID] [bigint] NOT NULL,"
      SQL = SQL & " [OSTERMOOBS] [nvarchar](250) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_OSTERMO] PRIMARY KEY CLUSTERED([OSTERMO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSTERMO]  WITH CHECK ADD  CONSTRAINT [FK_OSTERMO_OS] FOREIGN KEY([OS_ID])"
      SQL = SQL & " References [dbo].[OS]([OS_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSTERMO] CHECK CONSTRAINT [FK_OSTERMO_OS]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   VIEW: vwOSServico                                          */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwOS_Servico", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwOS_Servico]"
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwOSServico", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwOSServico]"

   SQL = "CREATE VIEW [dbo].[vwOSServico]"
   SQL = SQL & " AS "
   SQL = SQL & " SELECT OS.OS_ID, OS.ESTABELECIMENTO_ID, OS.PESSOA_ID, OS.CT_ID, OS.DT_OS, OS.DT_FECHA, OS.TIPO_OS, OS.SITUACAO_OS, "
   SQL = SQL & " OS.KM, OS.CLIENTE, OSSERVICO.OSSERVICO_ID, OSSERVICO.OSTAREFA_ID, OSSERVICO.DT_CAD AS DTCADSERVICO,"
   SQL = SQL & " OSSERVICO.SITUACAO, OSSERVICO.RESPONSAVEL_ID, OSSERVICO.VALOR_SERVICO, OSSERVICO.DT_FIM, OSSERVICO.DT_INICIO,"
   SQL = SQL & " OSSERVICO.DT_FECHA AS DTFECHASERVICO, OSSERVICO.DESCONTO_SERVICO, OSTAREFA.DT_CAD AS DTCADTAREFA,"
   SQL = SQL & " OSSERVICO.DESCRICAO AS DESCRICAOSERVICO, OSAPONTAMENTO.APONTAMENTO_ID, OSAPONTAMENTO.DATAINICIAL,"
   SQL = SQL & " OSAPONTAMENTO.DATAFINAL, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS DescPessoa, PESSOA.RAZAO,"
   SQL = SQL & " PESSOA.SITUACAO AS SITUACAOPESSOA"
   SQL = SQL & " FROM OS WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN OSSERVICO WITH (NOLOCK) "
   SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID "
   SQL = SQL & " INNER JOIN OSTAREFA WITH (NOLOCK) "
   SQL = SQL & " ON OSSERVICO.OSTAREFA_ID = OSTAREFA.OSTAREFA_ID "
   SQL = SQL & " INNER JOIN OSAPONTAMENTO WITH (NOLOCK) "
   SQL = SQL & " ON OSTAREFA.OSTAREFA_ID = OSAPONTAMENTO.OSTAREFA_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON OS.PESSOA_ID = PESSOA.PESSOA_ID"
CONECTA_RETAGUARDA.Execute SQL

' ============================================================ */
'   VIEW: vwOSpeca                                          */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwOSPECA", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwOSPECA]"

   SQL = "CREATE VIEW [dbo].[vwOSPECA]"
   SQL = SQL & " AS "
   SQL = SQL & " SELECT OS.OS_ID, OS.ESTABELECIMENTO_ID, OS.PESSOA_ID, OS.CT_ID, OS.DT_OS, OS.DT_FECHA, OS.TIPO_OS, OS.SITUACAO_OS, "
   SQL = SQL & " OS.KM, OS.CLIENTE, OSPECA.OSPECA_ID, OSPECA.PRODUTO_ID, OSPECA.DT_CAD AS DTCADPECA, OSPECA.SOLICITANTE_ID,"
   SQL = SQL & " OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, OSPECA.QTDE, OSPECA.DT_GARANTIA, PRODUTO.CODG_PRODUTO,"
   SQL = SQL & " PRODUTO.DESCRICAO AS DESCRICAOPRODUTO, PRODUTO.REFERENCIA, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_ATACADO,"
   SQL = SQL & " PRODUTO.PRECO_Venda, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS DescPessoa, PESSOA.RAZAO,"
   SQL = SQL & " PESSOA.SITUACAO AS SITUACAOPESSOA"
   SQL = SQL & " FROM PESSOA WITH (NOLOCK) "
   SQL = SQL & " RIGHT OUTER JOIN OS WITH (NOLOCK) "
   SQL = SQL & " ON PESSOA.PESSOA_ID = OS.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN OSPECA WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = OSPECA.PRODUTO_ID "
   SQL = SQL & " ON OS.OS_ID = OSPECA.OS_ID"
CONECTA_RETAGUARDA.Execute SQL

' ============================================================ */
'   VIEW: vwEquipamento                                            */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwEquipamento", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW vwEquipamento"

   SQL = "CREATE VIEW vwEquipamento"
   SQL = SQL & " AS "
   SQL = SQL & " SELECT OSEQUIPAMENTO.EQUIPAMENTO_ID, OSEQUIPAMENTO.PESSOA_ID, OSEQUIPAMENTO.COR_ID, OSEQUIPAMENTO.MARCA_ID, "
   SQL = SQL & " OSEQUIPAMENTO.DT_CAD, OSEQUIPAMENTO.DESCRICAO, OSEQUIPAMENTO.IDENTIFICACAO, OSEQUIPAMENTO.TIPO_EQP,"
   SQL = SQL & " OSEQUIPAMENTO.ANO, OSEQUIPAMENTO.MODELO, OSEQUIPAMENTO.NOME_CLIENTE, PESSOA.CNPJCPF,"
   SQL = SQL & " PESSOA.DESCRICAO AS DescPessoa, PESSOA.RAZAO, PESSOA.SITUACAO"
   SQL = SQL & " FROM OSEQUIPAMENTO "
   SQL = SQL & " INNER JOIN PESSOA "
   SQL = SQL & " ON OSEQUIPAMENTO.PESSOA_ID = PESSOA.PESSOA_ID"
CONECTA_RETAGUARDA.Execute SQL
' ============================================================ */
' Table: OSOBS                                               */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSOBS", "") = False Then
      SQL = "CREATE TABLE [dbo].[OSOBS]("
      SQL = SQL & " [OSOBS_ID] [bigint] NOT NULL,"
      SQL = SQL & " [OS_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [OBS] [nvarchar](250) NOT NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSOBS]  WITH CHECK ADD  CONSTRAINT [FK_OSOBS_OS] FOREIGN KEY([OS_ID]) References [dbo].[OS]([OS_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSOBS] CHECK CONSTRAINT [FK_OSOBS_OS]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
' VIEW: vwOSEQUIPAMENTO                                              */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwOSEQUIPAMENTO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW vwOSEQUIPAMENTO"

   SQL = "CREATE VIEW vwOSEQUIPAMENTO"
   SQL = SQL & " AS "

   SQL = SQL & " SELECT OS.OS_ID, OS.ESTABELECIMENTO_ID, OS.CT_ID, OS.DT_OS, OS.DT_FECHA, OS.TIPO_OS, OS.SITUACAO_OS, OS.KM, OS.CLIENTE, OSPECA.OSPECA_ID, "
   SQL = SQL & " OSPECA.PRODUTO_ID, OSPECA.DT_CAD AS DTCADPECA, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, OSPECA.QTDE,"
   SQL = SQL & " OSPECA.DT_GARANTIA, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO AS DESCRICAOPRODUTO, PRODUTO.REFERENCIA, PRODUTO.PRECO_CUSTO,"
   SQL = SQL & " PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, OSSERVICO.OSSERVICO_ID, OSSERVICO.OSTAREFA_ID, OSSERVICO.DT_CAD AS DTCADSERVICO, OSSERVICO.SITUACAO,"
   SQL = SQL & " OSSERVICO.RESPONSAVEL_ID, OSSERVICO.VALOR_SERVICO, OSSERVICO.DT_FIM, OSSERVICO.DT_INICIO, OSSERVICO.DT_FECHA AS DTFECHASERVICO,"
   SQL = SQL & " OSSERVICO.DESCONTO_SERVICO, OSTAREFA.DT_CAD AS DTCADTAREFA, OSSERVICO.DESCRICAO AS DESCRICAOSERVICO, OSAPONTAMENTO.APONTAMENTO_ID,"
   SQL = SQL & " OSAPONTAMENTO.DATAINICIAL, OSAPONTAMENTO.DATAFINAL, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS DESCRICAOPESSOA, PESSOA.RAZAO,"
   SQL = SQL & " PESSOA.SITUACAO AS SITUACAOPESSOA, OSEQUIPAMENTO.EQUIPAMENTO_ID, OSEQUIPAMENTO.COR_ID, OSEQUIPAMENTO.MARCA_ID, OSEQUIPAMENTO.DESCRICAO,"
   SQL = SQL & " OSEQUIPAMENTO.identificacao , OSEQUIPAMENTO.TIPO_EQP, OSEQUIPAMENTO.Ano, OSEQUIPAMENTO.MODELO, OSEQUIPAMENTO.NOME_CLIENTE, os.Pessoa_id"
   SQL = SQL & " FROM   OS WITH (NOLOCK) INNER JOIN"
   SQL = SQL & " PESSOA WITH (NOLOCK) ON OS.PESSOA_ID = PESSOA.PESSOA_ID INNER JOIN"
   SQL = SQL & " OSVEICEQP ON OS.OS_ID = OSVEICEQP.OS_ID INNER JOIN"
   SQL = SQL & " OSEQUIPAMENTO ON OSVEICEQP.EQUIPAMENTO_ID = OSEQUIPAMENTO.EQUIPAMENTO_ID LEFT OUTER JOIN"
   SQL = SQL & " OSTAREFA WITH (NOLOCK) INNER JOIN"
   SQL = SQL & " OSSERVICO WITH (NOLOCK) ON OSTAREFA.OSTAREFA_ID = OSSERVICO.OSTAREFA_ID INNER JOIN"
   SQL = SQL & " OSAPONTAMENTO WITH (NOLOCK) ON OSTAREFA.OSTAREFA_ID = OSAPONTAMENTO.OSTAREFA_ID ON OS.OS_ID = OSSERVICO.OS_ID LEFT OUTER JOIN"
   SQL = SQL & " OSPECA WITH (NOLOCK) INNER JOIN"
   SQL = SQL & " PRODUTO WITH (NOLOCK) ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID ON OS.OS_ID = OSPECA.OS_ID"
   CONECTA_RETAGUARDA.Execute SQL

' ============================================================ */
' Table: OSOBS                                               */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSVEICEQP", "") = False Then
      SQL = "CREATE TABLE [dbo].[OSVEICEQP]("
      SQL = SQL & " [OS_ID] [bigint] NOT NULL,"
      SQL = SQL & " [VEICULO_ID] [bigint] NULL,"
      SQL = SQL & " [EQUIPAMENTO_ID] [bigint] NULL"
      SQL = SQL & " ) ON [PRIMARY]"
CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSVEICEQP]  WITH CHECK ADD  CONSTRAINT [FK_OSVEICEQP_OS] FOREIGN KEY([OS_ID]) References [dbo].[OS]([OS_ID])"
CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSVEICEQP] CHECK CONSTRAINT [FK_OSVEICEQP_OS]"
CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSVEICEQP]  WITH CHECK ADD  CONSTRAINT [FK_OSVEICEQP_OSEQUIPAMENTO] FOREIGN KEY([EQUIPAMENTO_ID])"
      SQL = SQL & " References [dbo].[OSEQUIPAMENTO]([EQUIPAMENTO_ID])"
CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSVEICEQP] CHECK CONSTRAINT [FK_OSVEICEQP_OSEQUIPAMENTO]"
CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSVEICEQP]  WITH CHECK ADD  CONSTRAINT [FK_OSVEICEQP_OSVEICULO] FOREIGN KEY([VEICULO_ID])"
      SQL = SQL & " References [dbo].[OSVEICULO]([VEICULO_ID])"
CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OSVEICEQP] CHECK CONSTRAINT [FK_OSVEICEQP_OSVEICULO]"
CONECTA_RETAGUARDA.Execute SQL
   End If

   CRIA_TABELA_REL_OS

MsgBox "ok"
End Sub

Public Sub CRIA_TABELA_REL_OS()
' ============================================================ */
'   Table: OSREL                                            */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSRELITEM", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE OSRELITEM"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSREL", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE OSREL"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSREL", "") = False Then
      SQL = " CREATE TABLE [dbo].[OSREL]("
      SQL = SQL & " [OS_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PESSOA_ID_CLIENTE] [BIGint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_OS] [datetime] NOT NULL,"
      SQL = SQL & " [TIPO_OS] [nvarchar](20) NOT NULL,"
      SQL = SQL & " [SITUACAO_OS] [nvarchar](20) NOT NULL,"
      SQL = SQL & " [CONSULTOR_OS] [nvarchar](20) NOT NULL,"
      SQL = SQL & " [KM_OS] [bigint] NULL,"
      SQL = SQL & " [PLACA_OS] [nvarchar](8) NULL,"
      SQL = SQL & " [DT_OS_FEHCA] [datetime] NULL,"
      SQL = SQL & " [NUMR_FROTA_OS] [bigint] NULL,"
      SQL = SQL & " [NOME_EMP] [nvarchar](100) NULL,"
      SQL = SQL & " [CNPJ_EMP] [nvarchar](30) NULL,"
      SQL = SQL & " [ENDERECO_EMP] [nvarchar](100) NULL,"
      SQL = SQL & " [NUMERO_EMP] [bigint] NULL,"
      SQL = SQL & " [COMPLEM_EMP] [nvarchar](50) NULL,"
      SQL = SQL & " [CEP_EMP] [nvarchar](15) NULL,"
      SQL = SQL & " [BAIRRO_EMP] [nvarchar](50) NULL,"
      SQL = SQL & " [CIDADE_EMP] [nvarchar](30) NULL,"
      SQL = SQL & " [UF_EMP] [nvarchar](2) NULL,"
      SQL = SQL & " [FONE_EMP] [nvarchar](20) NULL,"
      SQL = SQL & " [NOME_CLI] [nvarchar](100) NULL,"
      SQL = SQL & " [CNPJCPF_CLI] [nvarchar](30) NULL,"
      SQL = SQL & " [FONE_CLI] [nvarchar](20) NULL,"
      SQL = SQL & " [DESC_VEICULO] [nvarchar](100) NULL,"
      SQL = SQL & " [COR_VEICULO] [nvarchar](20) NULL,"
      SQL = SQL & " [MARCA_VEICULO] [nvarchar](20) NULL,"
      SQL = SQL & " [TIPO_VEICULO] [nvarchar](20) NULL,"
      SQL = SQL & " [ANO_VEICULO] [nvarchar](4) NULL,"
      SQL = SQL & " [MODELO_VEICULO] [nvarchar](4) NULL,"
      SQL = SQL & " [COMB_VEICULO] [nvarchar](20) NULL,"
      SQL = SQL & " [CHASSI_VEICULO] [nvarchar](80) NULL,"
      SQL = SQL & " [MOTOR_VEICULO] [nvarchar](20) NULL,"
      SQL = SQL & " [FONE_RESP] [nvarchar](40) NULL,"
      SQL = SQL & " CONSTRAINT [PK_OSREL] PRIMARY KEY CLUSTERED([OS_ID] Asc )"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"

      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSREL_OS", "") = False Then
      SQL = " alter table OSREL "
      SQL = SQL & " add constraINT FK_OSREL_OS foreign key  (OS_ID)"
      SQL = SQL & " References OS(OS_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If
' ============================================================ */
'   Table: OSRELITEM                                            */
' ============================================================ */
   If EXISTE_OBJ_BANCO("RETAGUARDA", "OSRELITEM", "") = False Then
      SQL = " create table OSRELITEM ("
      SQL = SQL & " OS_ID              BIGINT         not null,"
      SQL = SQL & " OSRELITEM_ID       BIGINT         not null,"
      SQL = SQL & " TIPO_ITEM          NVARCHAR(1)    not null,"
      SQL = SQL & " USU_ID             BIGINT         not null,"
      SQL = SQL & " PROSERV_ID         BIGINT         not null,"
      SQL = SQL & " DT_CAD             datetime       not null,"
      SQL = SQL & " DESCRICAO          NVARCHAR(250)  not null,"
      SQL = SQL & " VALR_ITEM          float          not null,"
      SQL = SQL & " VALR_DESCONTO      float          not null,"
      SQL = SQL & " QTDE               float          not null,"
      SQL = SQL & " RESPONSAVEL        NVARCHAR(20)   not null,"
      SQL = SQL & " CODG_PRODUTO       NVARCHAR(100)  not null,"
      SQL = SQL & " DT_GARANTIA        datetime       null,"

      SQL = SQL & " constraINT PK_OSRELITEM primary key (OS_ID,OSRELITEM_ID,TIPO_ITEM))"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_OSRELITEM_OS", "") = False Then
      SQL = " alter table OSRELITEM "
      SQL = SQL & " add constraINT FK_OSRELITEM_OS foreign key  (OS_ID)"
      SQL = SQL & " References OSREL(OS_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

End Sub

Sub RODA_CFOP_PLANILHA_2()
'On Error GoTo ERRO_TRATA

   Msg = "Deseja importar cadastro de CFOP "
   PERGUNTA Msg, vbYesNo + 32, "Atualização", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      Dim TabCFOP As New ADODB.Recordset
      Dim Linha_Atual_A As String

      Set oConn = New ADODB.Connection
      oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                         "FIL=excel 8.0;" & _
                         "DefaultDir=" & App.Path & "\TXT\" & ";" & _
                         "MaxBufferSize=2048;" & _
                         "PageTimeout=5;" & _
                         "DBQ=" & App.Path & "\TXT\CFOP.xls" & ";"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'aabre o recordset pelo nome da planilha
      TabConsulta.Open "[Plan1$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

      TabConsulta.MoveFirst

      If TabConsulta.EOF Then
         MsgBox "Planilha incorreta !!!"
         Exit Sub
      End If

      While Not TabConsulta.EOF

         If Not IsNull(TabConsulta.Fields(0).Value) Then
            If Not IsNull(TabConsulta.Fields(1).Value) Then
               Linha_Atual_A = "" & Trim(TabConsulta.Fields(0).Value)

               If TabCFOP.State = 1 Then _
                  TabCFOP.Close

               SQL = "select * from CFOP "
               SQL = SQL & " where cfop_id = '" & Trim(Linha_Atual_A) & "'"
               TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabCFOP.EOF Then
                  If TabCFOP.State = 1 Then _
                     TabCFOP.Close
   
                  SQL = "INSERT INTO CFOP "
                     SQL = SQL & " (cfop_id,descricao,estabelecimento_id,obs)"
                  SQL = SQL & " VALUES ("
                     SQL = SQL & "'" & Trim(Linha_Atual_A) & "'"
                     SQL = SQL & ",'" & Trim(TabConsulta.Fields(1).Value) & "'"
                     SQL = SQL & "," & ESTABELECIMENTO_ID_N
                     SQL = SQL & ",'" & Trim(TabConsulta.Fields(2).Value) & "'"
                  SQL = SQL & ")"
   
                  Me.Caption = TabConsulta.Fields(0).Value
                  Else
                     SQL = "update CFOP set "
                     SQL = SQL & " obs = '" & Trim(TabConsulta.Fields(2).Value) & "'"
                     SQL = SQL & " where cfop_id = '" & Trim(TabConsulta.Fields(0).Value) & "'"
   
                     cmdCFOP.Caption = TabConsulta.Fields(0).Value
               End If
               If TabCFOP.State = 1 Then _
                  TabCFOP.Close
            End If
            Else 'linha de baixo
            
         End If

         DoEvents
         On Error Resume Next
         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      Command25.Caption = SQL3
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RODA_CFOP_PLANILHA_2"
End Sub

Private Sub cmdCompra_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOCOMPRA", "U") = False Then
      SQL = "CREATE TABLE [dbo].[PEDIDOCOMPRA]("
      SQL = SQL & " [PEDIDOCOMPRA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [FORNECEDOR_ID] [bigint] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_CADASTRO] [datetime] NOT NULL,"
      SQL = SQL & " [SITUACAO] [nvarchar](1) NOT NULL,"
      SQL = SQL & " [DT_BAIXA] [datetime] NULL,"
      SQL = SQL & " CONSTRAINT [pk_PEDIDOCOMPRA] PRIMARY KEY CLUSTERED("
      SQL = SQL & " [PEDIDOCOMPRA_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRA]  WITH CHECK ADD  "
      SQL = SQL & " CONSTRAINT [FK_PEDIDOCOMPRA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRA] CHECK CONSTRAINT [FK_PEDIDOCOMPRA_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRA]  WITH CHECK "
      SQL = SQL & " ADD  CONSTRAINT [FK_PEDIDOCOMPRA_FORNECEDOR] FOREIGN KEY([FORNECEDOR_ID])"
      SQL = SQL & " References [dbo].[FORNECEDOR]([FORNECEDOR_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRA] CHECK CONSTRAINT [FK_PEDIDOCOMPRA_FORNECEDOR]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOCOMPRAITEM", "U") = False Then
      SQL = " CREATE TABLE [dbo].[PEDIDOCOMPRAITEM]("
      SQL = SQL & " [PEDIDOCOMPRA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PEDIDOCOMPRAITEM_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRECO] [float] NOT NULL,"
      SQL = SQL & " [QTDE] [numeric](18, 3) NOT NULL,"
      SQL = SQL & " CONSTRAINT [pk_PEDIDOCOMPRAITEM] "
      SQL = SQL & " PRIMARY KEY CLUSTERED([PEDIDOCOMPRA_ID] ASC,[PEDIDOCOMPRAITEM_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRAITEM]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOCOMPRAITEM_PEDIDOCOMPRA] FOREIGN KEY([PEDIDOCOMPRA_ID])"
      SQL = SQL & " References [dbo].[PEDIDOCOMPRA]([PEDIDOCOMPRA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRAITEM] CHECK CONSTRAINT [FK_PEDIDOCOMPRAITEM_PEDIDOCOMPRA]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRAITEM]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOCOMPRAITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOCOMPRAITEM] CHECK CONSTRAINT [FK_PEDIDOCOMPRAITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_ALTERA", "PEDIDOCOMPRAITEM") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOCOMPRAITEM ADD DT_ALTERA DATETIME "
   End If
'=========================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ENTRADAESTOQUE", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ENTRADAESTOQUE]("
      SQL = SQL & " [ENTRADAESTOQUE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [FORNECEDOR_ID] [bigint] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_CADASTRO] [datetime] NOT NULL,"
      SQL = SQL & " [SITUACAO] [nvarchar](1) NOT NULL,"
      SQL = SQL & " [DT_BAIXA] [datetime] NULL,"
      SQL = SQL & " CONSTRAINT [pk_ENTRADAESTOQUE] PRIMARY KEY CLUSTERED("
      SQL = SQL & " [ENTRADAESTOQUE_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUE]  WITH CHECK ADD  "
      SQL = SQL & " CONSTRAINT [FK_ENTRADAESTOQUE_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUE] CHECK CONSTRAINT [FK_ENTRADAESTOQUE_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUE]  WITH CHECK "
      SQL = SQL & " ADD  CONSTRAINT [FK_ENTRADAESTOQUE_FORNECEDOR] FOREIGN KEY([FORNECEDOR_ID])"
      SQL = SQL & " References [dbo].[FORNECEDOR]([FORNECEDOR_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUE] CHECK CONSTRAINT [FK_ENTRADAESTOQUE_FORNECEDOR]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ENTRADAESTOQUEITEM", "U") = False Then
      SQL = " CREATE TABLE [dbo].[ENTRADAESTOQUEITEM]("
      SQL = SQL & " [ENTRADAESTOQUE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ENTRADAESTOQUEITEM_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRECO] [float] NOT NULL,"
      SQL = SQL & " [QTDE] [numeric](18, 3) NOT NULL,"
      SQL = SQL & " CONSTRAINT [pk_ENTRADAESTOQUEITEM] "
      SQL = SQL & " PRIMARY KEY CLUSTERED([ENTRADAESTOQUE_ID] ASC,[ENTRADAESTOQUEITEM_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUEITEM]  WITH CHECK ADD  CONSTRAINT [FK_ENTRADAESTOQUEITEM_ENTRADAESTOQUE] FOREIGN KEY([ENTRADAESTOQUE_ID])"
      SQL = SQL & " References [dbo].[ENTRADAESTOQUE]([ENTRADAESTOQUE_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUEITEM] CHECK CONSTRAINT [FK_ENTRADAESTOQUEITEM_ENTRADAESTOQUE]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUEITEM]  WITH CHECK ADD  CONSTRAINT [FK_ENTRADAESTOQUEITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTRADAESTOQUEITEM] CHECK CONSTRAINT [FK_ENTRADAESTOQUEITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub Command4_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTOQUEFORNEC", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ESTOQUEFORNEC]("
      SQL = SQL & " [ESTOQUEFORNEC_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTAB_ORIGEM_ID] [int] NOT NULL,"
      SQL = SQL & " [FORNECEDOR_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE_ENVIO] [numeric](18, 3) NOT NULL,"
      SQL = SQL & " [QTDE_RETORNO] [numeric](18, 3),"
      SQL = SQL & " [DT_MOVIMENTO] [datetime] NOT NULL,"
      SQL = SQL & " [DT_RETORNO] [datetime] ,"
      SQL = SQL & " [SITUACAO] [nchar](1) NOT NULL"
      SQL = SQL & " CONSTRAINT [PK_ESTOQUEFORNEC] PRIMARY KEY CLUSTERED([ESTOQUEFORNEC_ID] ASC,[SEQ_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUEFORNEC]  WITH CHECK ADD  CONSTRAINT [FK_ESTOQUEFORNEC_PRODUTO] "
      SQL = SQL & " FOREIGN KEY([PRODUTO_ID]) References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUEFORNEC] CHECK CONSTRAINT [FK_ESTOQUEFORNEC_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUEFORNEC]  WITH CHECK ADD  CONSTRAINT [FK_ESTOQUEFORNEC_FORNECEDOR] "
      SQL = SQL & " FOREIGN KEY([FORNECEDOR_ID]) References [dbo].[FORNECEDOR]([FORNECEDOR_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUEFORNEC] CHECK CONSTRAINT [FK_ESTOQUEFORNEC_FORNECEDOR]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUEFORNEC]  WITH CHECK ADD  CONSTRAINT [FK_ESTOQUEFORNEC_ESTABELECIMENTO] "
      SQL = SQL & " FOREIGN KEY([ESTAB_ORIGEM_ID]) References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[ESTOQUEFORNEC] CHECK CONSTRAINT [FK_ESTOQUEFORNEC_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "OK"
End Sub

Private Sub Command5_Click()

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then
      'MsgBox "Banco GLOBAL não conectado."
      Exit Sub
   End If

   SQL = "CREATE TABLE [dbo].[CARTOES]("
   SQL = SQL & "    [IDGrupoCredenciada] [int] identity (1,1) primary key NOT NULL,"
   SQL = SQL & "    [FormaPagID] [int] NOT NULL,"
   SQL = SQL & "    [CNPJcREDENCIADA] [nvarchar](14) NULL,"
   SQL = SQL & "    [CredenciadatBand] [nvarchar](2) NULL,"
   SQL = SQL & "    [NumeroAurorizacao] [nvarchar](40) NULL"
   SQL = SQL & " ) ON [PRIMARY]"
   CONECTA_GLOBAL.Execute SQL

   SQL = "insert into CARTOES "
   SQL = SQL & " values("
   SQL = SQL & "01"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ")"
   CONECTA_GLOBAL.Execute SQL

   SQL = "insert into CARTOES "
   SQL = SQL & " values("
   SQL = SQL & "02"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ")"
   CONECTA_GLOBAL.Execute SQL

   SQL = "insert into CARTOES "
   SQL = SQL & " values("
   SQL = SQL & "03"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ")"
   CONECTA_GLOBAL.Execute SQL

   SQL = "insert into CARTOES "
   SQL = SQL & " values("
   SQL = SQL & "04"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ")"
   CONECTA_GLOBAL.Execute SQL

   SQL = "insert into CARTOES "
   SQL = SQL & " values("
   SQL = SQL & "99"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ",NULL"
   SQL = SQL & ")"
   CONECTA_GLOBAL.Execute SQL

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close
End
End Sub

Private Sub Command6_Click()

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then
      'MsgBox "Banco GLOBAL não conectado."
      Exit Sub
   End If

   Dim TabCliIntegra    As New ADODB.Recordset
   Dim TabTempIntegra   As New ADODB.Recordset
   Dim CODG_CLIENTE     As Long
   Dim CODG_TRANSP      As Long

   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close

   SQL = "select mfatransp,mfACLIENTE,mfadoc from MFA010 "

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabCliIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
   While Not TabCliIntegra.EOF
      CODG_CLIENTE = 0 & TabCliIntegra.Fields("mfacliente").Value
      CODG_TRANSP = 0 & TabCliIntegra.Fields("mfatransp").Value

      SQL = "update MFA010 set "
      SQL = SQL & " mfatransp = '" & CODG_TRANSP & "'"
      SQL = SQL & ",mfacliente = '" & CODG_CLIENTE & "'"
      SQL = SQL & " where mfadoc = '" & Trim(TabCliIntegra.Fields("mfadoc").Value) & "'"
      CONECTA_GLOBAL.Execute SQL
Command6.Caption = TabCliIntegra.Fields("mfadoc").Value
DoEvents
      TabCliIntegra.MoveNext
   Wend
   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close

   SQL = "select mftcod,mftcgc from MFt010 "
   TabCliIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
   While Not TabCliIntegra.EOF
      CODG_TRANSP = 0 & TabCliIntegra.Fields("mftcod").Value

      SQL = "update MFt010 set "
      SQL = SQL & " mftcod = '" & CODG_TRANSP & "'"
      SQL = SQL & " where mftcgc = '" & TabCliIntegra.Fields("mftcgc").Value & "'"
      CONECTA_GLOBAL.Execute SQL
Command6.Caption = TabCliIntegra.Fields("mftcgc").Value
DoEvents
      TabCliIntegra.MoveNext
   Wend
   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close

   SQL = "select a1_cod,a1_cgc from sa1010 "
   TabCliIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
   While Not TabCliIntegra.EOF
      CODG_CLIENTE = 0 & TabCliIntegra.Fields("a1_cod").Value

      SQL = "update sa1010 set "
      SQL = SQL & " a1_cod = '" & CODG_CLIENTE & "'"
      SQL = SQL & " where a1_cgc = '" & TabCliIntegra.Fields("a1_cgc").Value & "'"
      CONECTA_GLOBAL.Execute SQL
Command6.Caption = TabCliIntegra.Fields("a1_cgc").Value
DoEvents
      TabCliIntegra.MoveNext
   Wend
   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

MsgBox "OK"
End Sub

Private Sub Command2_Click()
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then _
      Exit Sub

'MFA010
   If EXISTE_CAMPO_TABELA("GLOBAL", "MFANOMECONSUMIDOR", "MFA010") = False Then
      CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFANOMECONSUMIDOR nvarchar(60)"
      Else: Alteração_Definição_Campo_Tabela "MFANOMECONSUMIDOR", "nvarchar(60)", "MFa010", "GLOBAL"
   End If

   If EXISTE_CAMPO_TABELA("GLOBAL", "MFACPFCONSUMIDOR", "MFA010") = False Then
      CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFACPFCONSUMIDOR varchar(14)"
      Else: Alteração_Definição_Campo_Tabela "MFACPFCONSUMIDOR", "varchar(14)", "MFa010", "GLOBAL"
   End If

   If EXISTE_CAMPO_TABELA("GLOBAL", "vFCPST", "MFA010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD vFCPST float NULL CONSTRAINT DF_MFA010_vFCPST DEFAULT 0"

   If EXISTE_CAMPO_TABELA("GLOBAL", "vFCPSTRet", "MFA010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD vFCPSTRet float NULL CONSTRAINT DF_MFA010_vFCPSTRet DEFAULT 0"


'MFI010
   If EXISTE_CAMPO_TABELA("GLOBAL", "MFIPEDIDO", "MFi010") = True Then _
      Alteração_Definição_Campo_Tabela "MFIPEDIDO", "bigint", "MFi010", "GLOBAL"

   If EXISTE_OBJ_BANCO("GLOBAL", "MFI010", "U") = True Then _
      If EXISTE_CAMPO_TABELA("GLOBAL", "MFIITEM", "MFI010") = True Then _
         Alteração_Definição_Campo_Tabela "MFIITEM", "INT", "MFI010", "GLOBAL"

   If EXISTE_OBJ_BANCO("GLOBAL", "MFI010", "U") = True Then _
      If EXISTE_CAMPO_TABELA("GLOBAL", "MFIBASICMST", "MFI010") = True Then _
         Alteração_Definição_Campo_Tabela "MFIBASICMST", "float NULL CONSTRAINT DF_MFI010_MFIALIICMS DEFAULT 0", "MFI010", "GLOBAL"
   
   If EXISTE_CAMPO_TABELA("GLOBAL", "vBCFCPSTRet", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD vBCFCPSTRet float NULL CONSTRAINT DF_MFI010_vBCFCPSTRet DEFAULT 0"

   If EXISTE_CAMPO_TABELA("GLOBAL", "pFCPSTRet", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD pFCPSTRet float NULL CONSTRAINT DF_MFI010_pFCPSTRet DEFAULT 0"

   If EXISTE_CAMPO_TABELA("GLOBAL", "vFCPSTRet", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD vFCPSTRet float NULL CONSTRAINT DF_MFI010_vFCPSTRet DEFAULT 0"

   If EXISTE_CAMPO_TABELA("GLOBAL", "MFICEAN", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD MFICEAN varchar(50) NULL CONSTRAINT DF_MFI010_MFICEAN DEFAULT 'SEM GTIN'"

   If EXISTE_CAMPO_TABELA("GLOBAL", "MFICEANTRIB", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD MFICEANTRIB varchar(50) NULL CONSTRAINT DF_MFI010_MFICEANTRIB DEFAULT 'SEM GTIN'"

'=======
'---na tabela MFI010  CRIAR OS SEGUINTES CAMPOS :

'PARA O PIS E SEGUE O PADRAO DO ICMS TIPO NUMERICO 12,2 PARA VALOR E PERCENTAUL AI NAO SEI QUAL E
'1 - PISVBC VALOR BSASE DE CALUILO DO PIS
'2 - PISPPIS = PERCENTUAL DO MPIS PADRAO 0,65%
'3 - PISVPIS = VALOR DO PIS APLICAR OM PERCENTAUL SOBRE A BASE PEDE A CONTADORA PARA SABER SABER A BASE DE CACULO

   If EXISTE_CAMPO_TABELA("GLOBAL", "PISCST", "MFi010") = False Then
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD PISCST nvarchar(3)"
      Else: Alteração_Definição_Campo_Tabela "PISCST", "nvarchar(3)", "MFI010", "GLOBAL"
   End If

   If EXISTE_CAMPO_TABELA("GLOBAL", "COFINSCST", "MFi010") = False Then
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD COFINSCST nvarchar(3)"
      Else: Alteração_Definição_Campo_Tabela "COFINSCST", "nvarchar(3)", "MFI010", "GLOBAL"
   End If

   If EXISTE_CAMPO_TABELA("GLOBAL", "PISVBC", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD PISVBC FLOAT"

   If EXISTE_CAMPO_TABELA("GLOBAL", "PISPPIS", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD PISPPIS FLOAT"

   If EXISTE_CAMPO_TABELA("GLOBAL", "PISVPIS", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD PISVPIS FLOAT"


'---PARA COFINS :   MESMO CRITERIO DE CONCEITOS
'1 - COFINSVBC ,
'2 - COFINSPCOFINS, = 3% PADRAO
'3 - COFINSVCOFINS

   If EXISTE_CAMPO_TABELA("GLOBAL", "COFINSVBC", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD COFINSVBC FLOAT"

   If EXISTE_CAMPO_TABELA("GLOBAL", "COFINSPCOFINS", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD COFINSPCOFINS FLOAT"

   If EXISTE_CAMPO_TABELA("GLOBAL", "COFINSVCOFINS", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD COFINSVCOFINS FLOAT"


'---CRIAR MAIS ESTES CAMPOS NO MFI010 :

'PisqBCProd , n, 16, 4
'PisvAliqProd , n, 15, 4

   If EXISTE_CAMPO_TABELA("GLOBAL", "PisqBCProd", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD PisqBCProd FLOAT"

   If EXISTE_CAMPO_TABELA("GLOBAL", "PisvAliqProd", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD PisvAliqProd numeric(15, 4)"


'COFINSqBCProd , n, 16, 4
'COFINSvAliqProd , n, 15, 4
   If EXISTE_CAMPO_TABELA("GLOBAL", "COFINSqBCProd", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD COFINSqBCProd FLOAT"

   If EXISTE_CAMPO_TABELA("GLOBAL", "COFINSvAliqProd", "MFi010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE MFi010 ADD COFINSvAliqProd numeric(15, 4)"
'=================

'SB1010
   If EXISTE_CAMPO_TABELA("GLOBAL", "B1_CEAN", "SB1010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE SB1010 ADD B1_CEAN varchar(50) NULL "

   If EXISTE_CAMPO_TABELA("GLOBAL", "B1_CEANTRIB", "SB1010") = False Then _
      CONECTA_GLOBAL.Execute "ALTER TABLE SB1010 ADD B1_CEANTRIB varchar(50) NULL "


   'If EXISTE_OBJ_BANCO("GLOBAL", "SE1010", "U") = True Then _
      If EXISTE_CAMPO_TABELA("GLOBAL", "E1_NUM", "SE1010") = True Then _
         Alteração_Definição_Campo_Tabela "E1_NUM", "NVARCHAR(20)", "SE1010", "GLOBAL"

   'If EXISTE_OBJ_BANCO("GLOBAL", "SE1010", "U") = True Then _
      If EXISTE_CAMPO_TABELA("GLOBAL", "E1_NUMnota", "SE1010") = True Then _
         Alteração_Definição_Campo_Tabela "E1_NUMnota", "NVARCHAR(20)", "SE1010", "GLOBAL"

   'If EXISTE_OBJ_BANCO("GLOBAL", "MFA010", "U") = True Then _
      If EXISTE_CAMPO_TABELA("GLOBAL", "MFADOC", "MFA010") = True Then _
         Alteração_Definição_Campo_Tabela "MFADOC", "NVARCHAR(20)", "MFA010", "GLOBAL"

   'If EXISTE_OBJ_BANCO("GLOBAL", "MFI010", "U") = True Then _
      If EXISTE_CAMPO_TABELA("GLOBAL", "MFIDOC", "MFI010") = True Then _
         Alteração_Definição_Campo_Tabela "MFIDOC", "NVARCHAR(20)", "MFI010", "GLOBAL"

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

MsgBox "ok"
End Sub

Private Sub Command7_Click()

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUCAOPERDAVENDA", "U") = False Then
      SQL = " CREATE TABLE [dbo].[PRODUCAOPERDAVENDA]("
      SQL = SQL & " [PRODUTO_ID] [nchar](10) NOT NULL,"
      SQL = SQL & " [QtdeProducao] [float] NULL,"
      SQL = SQL & " [QtdePerda] [float] NULL,"
      SQL = SQL & " [QtdeVenda] [float] NULL,"
      SQL = SQL & " [QtdeVendaEstimada] [float] NULL,"
      SQL = SQL & " [QtdeVendaSistema] [float] NULL,"
      SQL = SQL & " [TotalVenda] [float] NULL,"
      SQL = SQL & " [PercVenda] [float] NULL,"
      SQL = SQL & " [PercProducao] [float] NULL,"
      SQL = SQL & " CONSTRAINT [PK_PRODUCAOPERDAVENDA] PRIMARY KEY CLUSTERED([PRODUTO_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "Turno", "U") = False Then
      SQL = "CREATE TABLE [dbo].[Turno]("
      SQL = SQL & " [TURNO_ID] [int] NOT NULL,"
      SQL = SQL & " [HoraIni] [time](0) NOT NULL,"
      SQL = SQL & " [HoraFim] [time](0) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_Turno] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([TURNO_ID] Asc )"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "idTurno", "TURNO") = True Then _
            CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'TURNO.idTurno'" & "," & "'TURNO_ID'" & "," & "'COLUMN'"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUCAOPERDA", "U") = False Then
      SQL = " CREATE TABLE [dbo].[PRODUCAOPERDA]("
      SQL = SQL & " [PRODUCAOPERDA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [DtRegistro] [datetime] NOT NULL,"
      SQL = SQL & " [TURNO_ID] [int] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [Qtde] [float] NOT NULL,"
      SQL = SQL & " [Valor] [float] NOT NULL,"
      SQL = SQL & " [ValorKG] [float] ,"
      SQL = SQL & " [PesoLiquido] [float] NOT NULL,"
      SQL = SQL & " [TipoRegistro] [nvarchar](5) NOT NULL,"
      SQL = SQL & " [Un] [nvarchar](10) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_PRODUCAOPERDA] PRIMARY KEY CLUSTERED([PRODUCAOPERDA_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOPERDA]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAOPERDA_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOPERDA] CHECK CONSTRAINT [FK_PRODUCAOPERDA_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOPERDA]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAOPERDA_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOPERDA] CHECK CONSTRAINT [FK_PRODUCAOPERDA_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOPERDA]  WITH CHECK ADD  CONSTRAINT [FK_PRODUCAOPERDA_Turno] FOREIGN KEY([TURNO_ID])"
      SQL = SQL & " References [dbo].[Turno]([TURNO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUCAOPERDA] CHECK CONSTRAINT [FK_PRODUCAOPERDA_Turno]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "OK"
End Sub

Private Sub Command8_Click()
   SQL = "delete PERMISSAO where Menuid = 'barINI.Buttons.Item(9)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'barINI.Buttons.Item(10)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'barINI.Buttons.Item(11)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'barINI.Buttons.Item(12)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuBoleto'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuECF'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuECFImp'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuECFOpera'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuECFVenda'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuFinanceiro'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuFinanceiroACcli'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuFinanceiroACFUNC'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuFinanceiroSTA'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuImportaToscana'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'MNULISTA'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuManut'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuManutBanco'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete PERMISSAO where Menuid = 'mnuRemessa'"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete MENU where Menuid = 'barINI.Buttons.Item(9)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'barINI.Buttons.Item(10)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'barINI.Buttons.Item(11)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'barINI.Buttons.Item(12)'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuBoleto'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuECF'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuECFImp'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuECFOpera'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuECFVenda'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuFinanceiro'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuFinanceiroACcli'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuFinanceiroACFUNC'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuFinanceiroSTA'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuImportaToscana'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'MNULISTA'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuManut'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuManutBanco'"
   CONECTA_RETAGUARDA.Execute SQL
   SQL = "delete MENU where Menuid = 'mnuRemessa'"
   CONECTA_RETAGUARDA.Execute SQL

   MsgBox "OK"
End Sub

Sub ATUALIZA_CLIENTE_GLOBAL()

'MsgBox "OK"
End Sub

Private Sub cmdEmpresa_Click()
   ATUALIZA_TABELA_EMPRESA
   ATUALIZA_ESTABELECIMENTO
   'CREATE INDEX IX_CNPJCPF ON PESSOA(CNPJCPF) WITH (ONLINE=ON, SORT_IN_TEMPDB=ON)

   If EXISTE_OBJ_BANCO("RETAGUARDA", "EMPRESAPARAMETRO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[EMPRESAPARAMETRO]("
      SQL = SQL & " [EMPRESAPARAMETRO_ID] [int] NOT NULL,"
      SQL = SQL & " [SEQ_PEDIDO] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_LOTE] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_PEDCOMPRA] [bigint] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_EMPRESAPARAMETRO] PRIMARY KEY CLUSTERED([EMPRESAPARAMETRO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into EMPRESAPARAMETRO "
         SQL = SQL & " (EMPRESAPARAMETRO_ID,SEQ_PEDIDO,SEQ_LOTE,SEQ_PEDCOMPRA)"
      SQL = SQL & " values("
         SQL = SQL & 1
         SQL = SQL & "," & 0
         SQL = SQL & "," & 0
         SQL = SQL & "," & 0
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   MsgBox "Ok, ATENÇÃO CRIAR INDICE CGC  "
End Sub

Private Sub cmdENTREGA_Click()
   ATUALIZA_TABELA_ENTREGA
   MsgBox "Ok"
End Sub


Private Sub cmdPrimoProduto_Click()
'FAMILIA PRODUTO

   Dim TabProdPrimo        As New ADODB.Recordset
   Dim TabProdMegasim      As New ADODB.Recordset
   Dim FAMILIAPRODUTO_ID_N As Long
   Dim CODG_NCM_N          As String

   SQL = "delete DESCR where tipo = 'W'"
   CONECTA_RETAGUARDA.Execute SQL

   If TabProdPrimo.State = 1 Then _
      TabProdPrimo.Close

   SQL = "select * from primoproduto"
   TabProdPrimo.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProdPrimo.EOF
'=====FAMILIAPRODUTO
      If TabProdMegasim.State = 1 Then _
         TabProdMegasim.Close

      SQL = "select * from FAMILIAPRODUTO "
      SQL = SQL & " where descricao = '" & Trim(TabProdPrimo.Fields("nome_grupo").Value) & "'"
      TabProdMegasim.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProdMegasim.EOF Then
         If TabProdMegasim.State = 1 Then _
            TabProdMegasim.Close
         FAMILIAPRODUTO_ID_N = 0 & MAX_ID("familiaproduto_id", "familiaproduto", "", "", "", "")
         SQL = "spFAMILIAPRODUTO " & 1 & "," & FAMILIAPRODUTO_ID_N & "," & FAMILIAPRODUTO_ID_N & ",'" & Trim(TabProdPrimo.Fields("nome_grupo").Value) & "','UN','UNIDADE',0,0"
         CONECTA_RETAGUARDA.Execute "EXEC " & SQL

         cmdPrimo.Caption = FAMILIAPRODUTO_ID_N
      End If
      If TabProdMegasim.State = 1 Then _
         TabProdMegasim.Close

'=====MARCA
      If TabProdMegasim.State = 1 Then _
         TabProdMegasim.Close

      SQL = "select * from DESCR "
      SQL = SQL & " where descricao = '" & Trim(TabProdPrimo.Fields("nome_marca").Value) & "'"
      SQL = SQL & " and tipo = 'W'"
      TabProdMegasim.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProdMegasim.EOF Then
         If TabProdMegasim.State = 1 Then _
            TabProdMegasim.Close
         SQL = "spDESCR " & 1 & _
                        "," & Trim(TabProdPrimo.Fields("CODGMARCA").Value) & _
                        ",'W'" & _
                        ",'" & Trim(TabProdPrimo.Fields("nome_marca").Value) & "'"

         CONECTA_RETAGUARDA.Execute "EXEC " & SQL

         cmdPrimo.Caption = FAMILIAPRODUTO_ID_N
      End If
      If TabProdMegasim.State = 1 Then _
         TabProdMegasim.Close

'=====PRODUTO
      If TabProdMegasim.State = 1 Then _
         TabProdMegasim.Close

      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(TabProdPrimo.Fields("codigo_prod").Value) & "'"
      TabProdMegasim.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProdMegasim.EOF Then
         PRODUTO_ID_N = 0 & Trim(TabProdPrimo.Fields("codigo_prod").Value)
         FORNEC_ID_N = 0
         FAMILIAPRODUTO_ID_N = 1
         MARCA_ID_N = 0
         CODG_NCM_N = "" & Trim(TabProdPrimo.Fields("codigo_ncm").Value)
         If Trim(CODG_NCM_N) = "" Then
            CODG_NCM_N = "00"
         End If

         If TabProdMegasim.State = 1 Then _
            TabProdMegasim.Close
         SQL = "select FAMILIAPRODUTO_id from FAMILIAPRODUTO "
         SQL = SQL & " where descricao = '" & Trim(TabProdPrimo.Fields("nome_grupo").Value) & "'"
         TabProdMegasim.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProdMegasim.EOF Then _
            FAMILIAPRODUTO_ID_N = 0 & TabProdMegasim.Fields(0).Value
         If TabProdMegasim.State = 1 Then _
            TabProdMegasim.Close

         SQL = "select codigo from DESCR "
         SQL = SQL & " where descricao = '" & Trim(TabProdPrimo.Fields("nome_marca").Value) & "'"
         SQL = SQL & " and tipo = 'W'"
         TabProdMegasim.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProdMegasim.EOF Then _
            MARCA_ID_N = 0 & TabProdMegasim.Fields(0).Value
         If TabProdMegasim.State = 1 Then _
            TabProdMegasim.Close

         SQL = "spPRODUTO "
         SQL = SQL & 1                                                   '@Acao int
         SQL = SQL & "," & PRODUTO_ID_N                                       '@PRODUTO_ID bigint
         SQL = SQL & "," & EMPRESA_ID_N                                       '@EMPRESA_ID int
         SQL = SQL & ",'" & PRODUTO_ID_N & "'"                                '@CODG_PRODUTO nvarchar(100)
         SQL = SQL & "," & FORNEC_ID_N                                        '@FORNECEDOR_ID int
         SQL = SQL & ",'" & Trim(TabProdPrimo.Fields("nome_produto").Value) & "'"  '@DESCRICAO varchar(200)
         SQL = SQL & ",'" & Trim(TabProdPrimo.Fields("referencia").Value) & "'"    '@REFERENCIA nvarchar(200)
         SQL = SQL & "," & FAMILIAPRODUTO_ID_N                               '@FAMILIAPRODUTO_ID bigint
         SQL = SQL & ",'" & Trim(TabProdPrimo.Fields("unidade").Value) & "'"       '@UNIDADE_MEDIDA nvarchar(10)
         SQL = SQL & ",''"                                                    '@CODG_BARRA nvarchar(50)
         SQL = SQL & ",'A'"                                                   '@SITUACAO nvarchar(1)
         SQL = SQL & ",'00'"                                                  '@SITUACAO_TRIBUTARIA nvarchar(80)
         SQL = SQL & ",'" & Trim(TabProdPrimo.Fields("aliquota").Value) & "'"      '@ALIQUOTA_ICMS float
         SQL = SQL & ",0"                                                     '@PERC_DESCONTO int
         SQL = SQL & ",1"                                                     '@TIPO_PROD nvarchar(50)
         SQL = SQL & ",'" & CODG_NCM_N & "'"                                  '@CODG_NCM nvarchar(8)
         SQL = SQL & ",0"                                                     '@COMP_TRIBUTARIA int
         SQL = SQL & ",0"                                                     '@PRECO_CUSTO_ANTERIOR float
         SQL = SQL & ",0"                                                     '@qtd_ped_anterior float
         SQL = SQL & ",'" & tpMOEDA(TabProdPrimo.Fields("preco_cust").Value) & "'" '@PRECO_CUSTO float
         SQL = SQL & ",'" & tpMOEDA(TabProdPrimo.Fields("preco_atac").Value) & "'" '@PRECO_ATACADO float
         SQL = SQL & ",'" & tpMOEDA(TabProdPrimo.Fields("preco_vend").Value) & "'" '@PRECO_Venda float
         SQL = SQL & ",0"                                                     '@PERCIVA float
         SQL = SQL & ",'" & DMA(TabProdPrimo.Fields("data_cadastro").Value) & "'"  '@DT_CADASTRO datetime
         SQL = SQL & ",0"                                                        '@PERC_COMIS float
         SQL = SQL & ",''"                                                       '@PATH_IMAGEM varchar(MAX)
         SQL = SQL & ",0"                                                        '@ORIGEM_MERCADO int
         SQL = SQL & ",''"                                                       '@LOCACAO varchar(MAX)
         SQL = SQL & ",0"                                                        '@PRECO_VAREJO_ANTERIOR float
         SQL = SQL & ",0"                                                        '@PRECO_ATACADO_ANTERIOR float
         SQL = SQL & ",0"                                                        '@EMBALAGEM int
         SQL = SQL & ",144"                                                      '@USUARIO_ID int
         SQL = SQL & ",0"                                                        '@QTD_MINIMO float
         SQL = SQL & ",0"                                                        '@QTD_MAXIMO float
         SQL = SQL & ",'" & DMA(TabProdPrimo.Fields("ultm_compra").Value) & "'"      '@DT_ULT_VENDA datetime
         SQL = SQL & ",'" & DMA(TabProdPrimo.Fields("ultm_compra").Value) & "'"      '@DT_ULT_COMPRA datetime
         SQL = SQL & ",0"                                                        '@PESO_LIQUIDO float
         SQL = SQL & ",0"                                                        '@PESO_BRUTO float
         SQL = SQL & ",0"                                                        '@TAMANHO bigint
         SQL = SQL & "," & MARCA_ID_N                                            '@MARCA_ID bigint
         SQL = SQL & ",1"                                                        '@PRODUTO_BALANCA bit
         SQL = SQL & ",0"                                                        '@PERMITE_DESCONTO bit
         SQL = SQL & ",0"                                                        '@CONCEDER_PRODUCAO bit
         SQL = SQL & ",0"                                                        '@PERC_COMPOE_VENDA float

         CONECTA_RETAGUARDA.Execute "EXEC " & SQL
      End If
      If TabProdMegasim.State = 1 Then _
         TabProdMegasim.Close
'===========
      cmdComanda.Caption = "" & Trim(TabProdPrimo.Fields("nome_produto").Value)
      DoEvents
      TabProdPrimo.MoveNext
   Wend
   If TabProdPrimo.State = 1 Then _
      TabProdPrimo.Close

MsgBox "ok PRODUTO PRIMO"
End Sub

Private Sub cmdPrimoCliente_Click()

   Dim TabPrimoPessoa      As New ADODB.Recordset
   Dim TabMegaPessoa       As New ADODB.Recordset
   Dim TabMegaTemp         As New ADODB.Recordset
   Dim NUMR_FONE_A         As String
   Dim DDD_A               As String
   Dim CEP_A               As String
   Dim CIDADE_A            As String
   Dim UF_A                As String
   Dim IBGE_A              As String
   Dim EMAIL_A             As String
   Dim IE_A                As String
   Dim IM_A                As String
   Dim RUA_A               As String
   Dim BAIRRO_A            As String
   Dim COMPLEMENTO_A       As String
   Dim TIPO_A              As String
   Dim Numero_A            As String
   Dim DT_NASC_A           As String
   Dim DT_CAD_A            As String
   Dim STATUS_A            As String
   Dim CONTATO_A           As String
   Dim OBS_A               As String
   Dim CNPJ_CPF_A          As String
   Dim TAMANHO_FONE_N      As Integer

'=====CLIENTE
   If TabPrimoPessoa.State = 1 Then _
      TabPrimoPessoa.Close

   SQL = "select * from PRIMOVAI"
   SQL = SQL & " order by codigo"
   TabPrimoPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPrimoPessoa.EOF
      CNPJ_CPF_A = "" & Trim(TabPrimoPessoa.Fields("cnpj").Value)

      CNPJ_CPF_A = Replace(CNPJ_CPF_A, "/", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, ".", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, "-", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, ";", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, " ", "")
      CNPJ_CPF_A = Trim(CNPJ_CPF_A)

      If Trim(CNPJ_CPF_A) = "" Then _
         CNPJ_CPF_A = "" & Trim(TabPrimoPessoa.Fields("cpf").Value)

      CNPJ_CPF_A = Replace(CNPJ_CPF_A, "/", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, ".", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, "-", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, ";", "")
      CNPJ_CPF_A = Replace(CNPJ_CPF_A, " ", "")
      CNPJ_CPF_A = Trim(CNPJ_CPF_A)

'Debug.Print CNPJ_CPF_A

      If Trim(CNPJ_CPF_A) <> "" Then
         If VALIDA_CNPJCPF(Trim(CNPJ_CPF_A)) = True Then
''''''''''''PESSOA
            NOME_A = "" & Trim(TabPrimoPessoa.Fields("NOME").Value)
            NOME_A = Replace(NOME_A, "'", "´")

            If TabMegaPessoa.State = 1 Then _
               TabMegaPessoa.Close

            SQL = "select * from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJ_CPF_A) & "'"
            TabMegaPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabMegaPessoa.EOF Then
               spPessoa 1, 0, Trim(CNPJ_CPF_A), Trim(NOME_A), Trim(NOME_A), "A"
               PESSOA_ID_N = 0 & TRAZ_ID_TABELA("PESSOA", "pessoa_id", "cnpjcpf", Trim(CNPJ_CPF_A))
               Else: PESSOA_ID_N = 0 & TabMegaPessoa.Fields("pessoa_id").Value
            End If

''''''''''''FONE
            If Trim(TabPrimoPessoa.Fields("FONE").Value) <> "" Then
               DDD_A = "" & Mid(TabPrimoPessoa.Fields("FONE").Value, 2, 2)
               NUMR_FONE_A = "" & Trim(TabPrimoPessoa.Fields("FONE").Value)
               Seq_N = 1
               CRITERIO_A = ""

TAMANHO_FONE_N = 0 & Len(NUMR_FONE_A) + 1

While TAMANHO_FONE_N > Seq_N

'MsgBox Len(Mid(NUMR_FONE_A, 1, Len(NUMR_FONE_A) - Seq_N)) + 1
'MsgBox "No Laço >>> " & Len(NUMR_FONE_A) & "   " & Seq_N & "   " & CRITERIO_A

   If Mid(NUMR_FONE_A, Seq_N, 1) <> " " Then
      CRITERIO_A = CRITERIO_A & Mid(NUMR_FONE_A, Seq_N, 1)
   End If

   Seq_N = Seq_N + 1
Wend
'Debug.Print CRITERIO_A
               NUMR_FONE_A = CRITERIO_A
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", " ")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", "")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "/", "")
               NUMR_FONE_A = Trim(NUMR_FONE_A)
               NUMR_FONE_A = Right(NUMR_FONE_A, 9)
'Debug.Print NUMR_FONE_A
'=========================

If IsNumeric(Left(NUMR_FONE_A, 4)) Then
   If Trim(DDD_A) = "" Then
      DDD_A = "62"
   End If
               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from FONE "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and numero = '" & Trim(NUMR_FONE_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then _
                  spFONE 1, 0, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_A), ""
               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
            End If
End If
''''''''''''Fax
            If Trim(TabPrimoPessoa.Fields("fax").Value) <> "" Then
               DDD_A = "" & Mid(TabPrimoPessoa.Fields("FAX").Value, 2, 2)
               NUMR_FONE_A = "" & Trim(TabPrimoPessoa.Fields("FAX").Value)
               Seq_N = 1
               CRITERIO_A = ""

TAMANHO_FONE_N = 0 & Len(NUMR_FONE_A) + 1

While TAMANHO_FONE_N > Seq_N

'MsgBox Len(Mid(NUMR_FONE_A, 1, Len(NUMR_FONE_A) - Seq_N)) + 1
'MsgBox "No Laço >>> " & Len(NUMR_FONE_A) & "   " & Seq_N & "   " & CRITERIO_A

   If Mid(NUMR_FONE_A, Seq_N, 1) <> " " Then
      CRITERIO_A = CRITERIO_A & Mid(NUMR_FONE_A, Seq_N, 1)
   End If

   Seq_N = Seq_N + 1
Wend
'Debug.Print CRITERIO_A
               NUMR_FONE_A = CRITERIO_A
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", " ")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", "")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "/", "")
               NUMR_FONE_A = Trim(NUMR_FONE_A)
               NUMR_FONE_A = Right(NUMR_FONE_A, 9)
'Debug.Print NUMR_FONE_A
'=========================
If IsNumeric(Left(NUMR_FONE_A, 4)) Then
   If Trim(DDD_A) = "" Then
      DDD_A = "62"
   End If
               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from FONE "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and numero = '" & Trim(NUMR_FONE_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then _
                  spFONE 1, 0, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_A), "FAX"
            End If
End If

''''''''''''FONE_EMPR
            If Trim(TabPrimoPessoa.Fields("FONE_EMPR").Value) <> "" Then
               DDD_A = "" & Mid(TabPrimoPessoa.Fields("FONE_EMPR").Value, 2, 2)
               NUMR_FONE_A = "" & Trim(TabPrimoPessoa.Fields("FONE_EMPR").Value)
               Seq_N = 1
               CRITERIO_A = ""

TAMANHO_FONE_N = 0 & Len(NUMR_FONE_A) + 1

While TAMANHO_FONE_N > Seq_N

'MsgBox Len(Mid(NUMR_FONE_A, 1, Len(NUMR_FONE_A) - Seq_N)) + 1
'MsgBox "No Laço >>> " & Len(NUMR_FONE_A) & "   " & Seq_N & "   " & CRITERIO_A

   If Mid(NUMR_FONE_A, Seq_N, 1) <> " " Then
      CRITERIO_A = CRITERIO_A & Mid(NUMR_FONE_A, Seq_N, 1)
   End If

   Seq_N = Seq_N + 1
Wend
'Debug.Print CRITERIO_A
               NUMR_FONE_A = CRITERIO_A
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", " ")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", "")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "/", "")
               NUMR_FONE_A = Trim(NUMR_FONE_A)
               NUMR_FONE_A = Right(NUMR_FONE_A, 9)
'Debug.Print NUMR_FONE_A
'=========================
If IsNumeric(Left(NUMR_FONE_A, 4)) Then
   If Trim(DDD_A) = "" Then
      DDD_A = "62"
   End If
               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from FONE "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and numero = '" & Trim(NUMR_FONE_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then
                  spFONE 1, 0, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_A), "FONE_EMPR"
                  Else: spFONE 2, TabMegaTemp.Fields("fone_id").Value, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_A), "Empresa"
               End If
            End If
End If

''''''''''''CELULAR_CONTATO
            If Trim(TabPrimoPessoa.Fields("CELULAR_CONTATO").Value) <> "" Then
               DDD_A = "" & Mid(TabPrimoPessoa.Fields("CELULAR_CONTATO").Value, 2, 2)
               NUMR_FONE_A = "" & Trim(TabPrimoPessoa.Fields("CELULAR_CONTATO").Value)
               Seq_N = 1
               CRITERIO_A = ""

TAMANHO_FONE_N = 0 & Len(NUMR_FONE_A) + 1

While TAMANHO_FONE_N > Seq_N

'MsgBox Len(Mid(NUMR_FONE_A, 1, Len(NUMR_FONE_A) - Seq_N)) + 1
'MsgBox "No Laço >>> " & Len(NUMR_FONE_A) & "   " & Seq_N & "   " & CRITERIO_A

   If Mid(NUMR_FONE_A, Seq_N, 1) <> " " Then
      CRITERIO_A = CRITERIO_A & Mid(NUMR_FONE_A, Seq_N, 1)
   End If

   Seq_N = Seq_N + 1
Wend
'Debug.Print CRITERIO_A
               NUMR_FONE_A = CRITERIO_A
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", " ")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "'", "")
               NUMR_FONE_A = Replace(NUMR_FONE_A, "/", "")
               NUMR_FONE_A = Trim(NUMR_FONE_A)
               NUMR_FONE_A = Right(NUMR_FONE_A, 9)
'Debug.Print NUMR_FONE_A
'=========================
If IsNumeric(Left(NUMR_FONE_A, 4)) Then
               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from FONE "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and numero = '" & Trim(NUMR_FONE_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then _
                  spFONE 1, 0, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_A), "CELULAR"
            End If
End If

''''''''''''EMAIL
            If Trim(TabPrimoPessoa.Fields("email").Value) <> "" Then
               EMAIL_A = "" & Trim(TabPrimoPessoa.Fields("email").Value)

               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from EMAIL "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and email = '" & Trim(EMAIL_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then _
                  spEmail 1, 0, Trim(EMAIL_A), PESSOA_ID_N
            End If

''''''''''''EMAIL
            If Trim(TabPrimoPessoa.Fields("email_NFE").Value) <> "" Then
               EMAIL_A = "" & Trim(TabPrimoPessoa.Fields("email_NFE").Value)

               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from EMAIL "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and email = '" & Trim(EMAIL_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then _
                  spEmail 1, 0, Trim(EMAIL_A), PESSOA_ID_N
            End If

''''''''''''CEP
            If Trim(TabPrimoPessoa.Fields("cep").Value) <> "" Then
               CEP_A = "" & Trim(TabPrimoPessoa.Fields("cep").Value)
               CEP_A = Replace(CEP_A, "/", "")
               CEP_A = Replace(CEP_A, ".", "")
               CEP_A = Replace(CEP_A, "-", "")
               CEP_A = Replace(CEP_A, ";", "")
               CEP_A = Replace(CEP_A, " ", "")
               CEP_A = Trim(CEP_A)

               If Len(Trim(CEP_A)) = 8 Then
                  If TabMegaTemp.State = 1 Then _
                     TabMegaTemp.Close
                  SQL = "select * from CEP "
                  SQL = SQL & " where cep_id = '" & Trim(CEP_A) & "'"
                  TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabMegaTemp.EOF Then
                     If CONSULTA_CEP_WEB(Trim(CEP_A)) = True Then
                        CIDADE_A = "" & Trim(Xcidade_A)
                        UF_A = "" & Trim(Xuf_A)

                        IBGE_A = ""
                        If Trim(UF_A) <> "" And Trim(CIDADE_A) <> "" Then
                           If TabCEP.State = 1 Then _
                              TabCEP.Close
                           SQL = "select ibge_id from IBGE "
                           SQL = SQL & " WHERE estado = '" & Trim(UF_A) & "'"
                           SQL = SQL & " and municipio = '" & Trim(CIDADE_A) & "'"
                           TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabCEP.EOF Then _
                              IBGE_A = "" & TabCEP.Fields(0).Value
                           If TabCEP.State = 1 Then _
                              TabCEP.Close
                        End If

                        spCEP 1, Trim(CEP_A), Trim(CIDADE_A), Trim(UF_A), Trim(IBGE_A)
                     End If
                     Else: UF_A = "" & Trim(TabMegaTemp.Fields("uf").Value)
                  End If
               End If
            End If

''''''''''''ENDEREÇO
            If Trim(TabPrimoPessoa.Fields("endereco").Value) <> "" Then
               ENDERECO_A = "" '& Trim(TabPrimoPessoa.Fields("endereco").Value)
               RUA_A = "" & Trim(TabPrimoPessoa.Fields("endereco").Value)
               BAIRRO_A = "" & Trim(TabPrimoPessoa.Fields("bairro").Value)
               COMPLEMENTO_A = "" & Trim(TabPrimoPessoa.Fields("endereco_complemento").Value)
               TIPO_A = "C"
               Numero_A = "" & Trim(TabPrimoPessoa.Fields("endereco_numero").Value)

               If Trim(RUA_A) <> "" Then
                  If Trim(CEP_A) <> "" Then
                     If TabMegaTemp.State = 1 Then _
                        TabMegaTemp.Close
                     SQL = "select * from ENDERECO "
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     SQL = SQL & " and cep_id = '" & Trim(CEP_A) & "'"
                     SQL = SQL & " and tipo = '" & TIPO_A & "'"
                     TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If TabMegaTemp.EOF Then
                        spENDERECO 1, ENDERECO_ID_N, PESSOA_ID_N, CEP_A, RUA_A, BAIRRO_A, COMPLEMENTO_A, TIPO_A, Numero_A
                        ENDERECO_ID_N = 0 & TRAZ_ID_ENDERECO("C")
                        Else: ENDERECO_ID_N = 0 & TabMegaTemp.Fields("endereco_id").Value
                     End If
                  End If
               End If
            End If

''''''''''''IE
            If Trim(TabPrimoPessoa.Fields("INSCRICAO_ESTADUAL").Value) <> "" Then
               IE_A = "" & Trim(TabPrimoPessoa.Fields("INSCRICAO_ESTADUAL").Value)
               If Trim(IE_A) <> "ISENTO" Then
                  If Trim(IE_A) <> "" Then
                     If Valida_Inscricao_Estadual(IE_A, UF_A) <> 1 Then
                        If TabMegaTemp.State = 1 Then _
                           TabMegaTemp.Close
                        SQL = "select * from IE "
                        SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                        SQL = SQL & " and numr_iE = '" & Trim(IE_A) & "'"
                        TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If TabMegaTemp.EOF Then _
                           spIE 1, 0, Trim(IE_A), ENDERECO_ID_N
                        'Else: MsgBox "INSCRIÇÃO ERRADA = " & IE_A
                     End If
                  End If
               End If
            End If

''''''''''''IM
            If Trim(TabPrimoPessoa.Fields("INSCRICAO_MUNICIPAL").Value) <> "" Then
               IM_A = "" & Trim(TabPrimoPessoa.Fields("INSCRICAO_MUNICIPAL").Value)

               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from IM "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and numr_IM = '" & Trim(IM_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then _
                  spIM 1, 0, Trim(IM_A), ENDERECO_ID_N
            End If

''''''''''''CLIENTE
            If PESSOA_ID_N > 0 Then
               CLIENTE_ID_N = 0 & Trim(TabPrimoPessoa.Fields("CODIGO").Value)
               VENDEDOR_ID_N = 1
               DT_NASC_A = "" & Trim(TabPrimoPessoa.Fields("DATA_NASC").Value)
               DT_NASC_A = "NULL"
               DT_CAD_A = "" & DMA(TabPrimoPessoa.Fields("DATA_FICHA").Value)
               CONTATO_A = "" & Trim(TabPrimoPessoa.Fields("contato").Value)
               STATUS_A = "A"

               If TabMegaTemp.State = 1 Then _
                  TabMegaTemp.Close
               SQL = "select * from CLIENTE "
               'SQL = SQL & " where cliente_id = " & CLIENTE_ID_N
               SQL = SQL & " where cgccpf = '" & Trim(CNPJ_CPF_A) & "'"
               TabMegaTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabMegaTemp.EOF Then
                  spCliente 1, 1, CNPJ_CPF_A, NOME_A, NOME_A, DT_NASC_A, _
                               DT_CAD_A, STATUS_A, "", CONTATO_A, 0, 0, 0 _
                             , 0, "", 0, OBS_A, ""
                  Else
                     CLIENTE_ID_N = 0 & TabMegaTemp.Fields("cliente_id").Value
                     spCliente 2, 1, CNPJ_CPF_A, NOME_A, NOME_A, DT_NASC_A, _
                                  DT_CAD_A, STATUS_A, "", CONTATO_A, 0, 0, 0 _
                                , 0, "", 0, OBS_A, ""
               End If
            End If

''''''''''''OBS

         End If   'If VALIDA_CNPJCPF(Trim(CNPJ_CPF_A)) = True Then
      End If   'If Trim(CNPJ_CPF_A) <> "" Then

      cmdPrimoProduto.Caption = "" & TabPrimoPessoa.Fields("codigo").Value
      'cmdComanda.Caption = "" & TabPrimoPessoa.Fields("codigo").Value
      Me.Caption = "" & NOME_A
      DoEvents

      TabPrimoPessoa.MoveNext
   Wend
   If TabPrimoPessoa.State = 1 Then _
      TabPrimoPessoa.Close

MsgBox "Ok Clientes"
End Sub
'=======YURI
Private Sub cmdYuri_ClickOLD()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ALIQUOTA_UF", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ALIQUOTA_UF]("
      SQL = SQL & " [UF_ORIGEM] [char](2) NOT NULL,"
      SQL = SQL & " [ALIQUOTA_ICMS_DENTRO] [int] NOT NULL,"
      SQL = SQL & " [UF_DESTINO] [char](2) NOT NULL,"
      SQL = SQL & " [ALIQUOTA_ICMS_FORA] [Int] NOT NULL) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'GO',0"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'AC',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'AL',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'AM',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'AP',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'BA',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'CE',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'DF',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'ES',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'MA',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'MT',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'MS',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'MG',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'PA',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'PB',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'PR',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'PE',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'PI',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'RN',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'RS',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'RJ',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'RO',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'RR',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'SC',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'SP',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'SE',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into ALIQUOTA_UF "
         SQL = SQL & " (uf_origem,ALIQUOTA_ICMS_DENTRO,UF_DESTINO,ALIQUOTA_ICMS_FORA)"
      SQL = SQL & " values("
         SQL = SQL & "'GO',17,'TO',12"
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "UF", "U") = False Then
      SQL = "CREATE TABLE [dbo].[UF]("
      SQL = SQL & " [UF_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](100) NOT NULL,"
      SQL = SQL & " [ESTADO] [nvarchar](2) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_UF] PRIMARY KEY CLUSTERED([UF_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) "
      SQL = SQL & " ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "11,'Rondônia','RO'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "12,'Acre','AC'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "13,'Amazonas','AM'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "14,'Roraima','RR'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "15,'Pará','PA'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "16,'Amapá','AP'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "17,'Tocantins','TO'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "21,'Maranhão','MA'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "22,'Piauí','Pi'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "23,'Ceará','CE'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "24,'Rio Grande do Norte','RN'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "25,'Paraíba','PB'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "26,'Pernambuco','PE'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "27,'Alagoas','AL'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "28,'Sergipe','SE'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "29,'Bahia','BA'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "31,'Minas Gerais','MG'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "32,'Espírito Santo','ES'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "33,'Rio de Janeiro','RJ'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "35,'São Paulo','SP'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "41,'Paraná','PR'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "42,'Santa Catarina','SC'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "43,'Rio Grande do Sul','RS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "50,'Mato Grosso do Sul','MS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "51,'Mato Grosso','MT'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
         
      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "52,'Goiás','GO'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (uf_id,descricao,estado)"
      SQL = SQL & " values("
         SQL = SQL & "53,'Distrito Federal','DF'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else: MsgBox "Alterar campo id da tabela (UF) pra uf_id"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CSOSN", "") = False Then
      SQL = " CREATE TABLE CSOSN("
      SQL = SQL & " CODIGO NVARCHAR(3) NOT NULL,"
      SQL = SQL & " DESCRICAO NVARCHAR(max) NOT NULL,"
      SQL = SQL & " OBS NVARCHAR(max) NULL,"
      SQL = SQL & " CONSTRAINT PK_TRIBUTACAO_CSOSN PRIMARY KEY CLUSTERED"
      SQL = SQL & " (Codigo Asc )"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 101
      SQL = SQL & ",'Tributada pelo Simples Nacional com permissão de crédito'"
      SQL = SQL & ",'classificam-se neste código as operações que permitem a indicação da alíquota de ICMS devido no Simples Nacional e o valor do crédito correspondente'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 102
      SQL = SQL & ",'Tributada pelo Simples Nacional sem permissão de crédito'"
      SQL = SQL & ",'classificam-se código as operações que não permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor do crédito, e não estejam abrangidas nas hipóteses dos códigos 103, 203, 300, 400, 500 e 900'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 103
      SQL = SQL & ",'Isenção do ICMS no Simples Nacional para faixa de receita bruta'"
      SQL = SQL & ",'classificam-se neste código as operações praticadas por optantes do Simples Nacional contempladas com isenção concedida para faixa de receita bruta nos termos da Lei Complementar n. 123 de 2006'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 201
      SQL = SQL & ",'Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações  que permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor crédito e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 202
      SQL = SQL & ",'Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações  que não permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor crédito, e não estejam abrangidas nas hipóteses dos códigos 103, 203, 300, 400, 500 e 900 e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 203
      SQL = SQL & ",'Isenção do ICMS no Simples Nacional para a faixa de receita bruta e com cobrança de ICMS por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações que praticadas por optantes do Simples Nacional contemplados com isenção para a faixa de receita bruta, mas com ICMS cobrado por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 300
      SQL = SQL & ",'Imune'"
      SQL = SQL & ",'classificam-se neste código as operações que praticadas por optantes do Simples Nacional contempladas com imunidade do ICMS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 400
      SQL = SQL & ",'Não tributada pelo Simples Nacional'"
      SQL = SQL & ",'classificam-se neste código as operações que praticadas por optantes do Simples NacionaL não sujeitas à tributação pelo ICMS dentro do Simples Nacional'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 500
      SQL = SQL & ",'ICMS cobrado anteriormente por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações sujeitas exclusivamente ao regime de substituição tributária na condição de substituído tributário ou no caso de antecipações'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 900
      SQL = SQL & ",'Outros'"
      SQL = SQL & ",'classificam-se neste código as operações que não se enquadrem nos códigos 101, 102, 103, 201, 202, 203, 300, 400 e 500'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   ' Estes códigos são fixos e somente podem ter 8 itens
   ' daqui  vou precisar mudar nao pode ser assim, tem que ser com esses campos,vou deletar de la
   ' retirei la embaixo  e ficara esse
   'If EXISTE_OBJ_BANCO("RETAGUARDA", "ALIQUOTA", "U") = False Then
   '   sSQL = "CREATE TABLE [dbo].[ALIQUOTA] ( "
   '   sSQL = sSQL & "    [CODIGO] [INT] NOT NULL ,"
   '   sSQL = sSQL & "    [NOME] [varchar] (50)  NULL ,"
   '   sSQL = sSQL & "    [aliquota_do_imposto] [real] NULL,"
   '   sSQL = sSQL & "    [EMPRESA] [INT] NOT NULL"
   '   sSQL = sSQL & ") ON [PRIMARY]"
   '   CONECTA_RETAGUARDA.Execute sSQL
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (1,'ISENTO',0.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (2,'Substituição Tributaria',0.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (3,'Nao Incidencia',0.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (4,'ICMS 12%',12.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (5,'ICMS 0%',0.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (6,'ISS 5%',5.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (7,'ICMS 15%',15.0000,1)"
   '        CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (8,'ICMS 17%',17.0000,1)"
   '   Else
   '      TabTemp.Open "select ISNULL(COUNT(*), 0) AS QTD from [dbo].[ALIQUOTA]", CONECTA_RETAGUARDA, , , adCmdText
   '      If TabTemp!QTD = 0 Then
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (1,'ISENTO',0.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (2,'Substituição Tributaria',0.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (3,'Nao Incidencia',0.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (4,'ICMS 12%',12.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (5,'ICMS 0%',0.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (6,'ISS 5%',5.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (7,'ICMS 15%',15.0000,1)"
   '         CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA] (Codigo,NOME,aliquota_do_imposto,EMPRESA) VALUES (8,'ICMS 17%',17.0000,1)"
   '      End If
   'End If
   'If TabTemp.State = 1 Then _
     TabTemp.Close

   ' Estes códigos são fixos e somente podem ter 27 itens QUE CORREPOMDEM A VINTE E SETE ESTADOS
   ' DEPOIS VC COLOCA ISTO AONDE VOCE CRIOU ESSA TABELA
   'If EXISTE_OBJ_BANCO("RETAGUARDA", "ALIQUOTA_UF", "U") = False Then
   '   SQL = " CREATE TABLE ALIQUOTA_UF("
   '   SQL = SQL & " CODIGO bigint not null"
   '   SQL = SQL & ", ESTADO varchar(2) NOT NULL"
   '   SQL = SQL & ", codigo_aliquota bigint NOT NULL"
   '   SQL = SQL & ", Aliquota decimal(4,2) NOT NULL "
   '   SQL = SQL & ", aliquota_nc decimal(4,2) NOT NULL "
   '   SQL = SQL & ", Descricao varchar(100) NOT NULL "
   '   SQL = SQL & ", codigo_aliquota_nc bigint NOT NULL "
   '   SQL = SQL & ", codigo_aliquota_substituicao bigint NOT NULL "
   '   SQL = SQL & ", aliquota_substituicao decimal(4,2) NOT NULL "
   '   SQL = SQL & ", empresa_id bigint NOT NULL "
   '   SQL = SQL & ", codigo_uf bigint NOT NULL "
   '   SQL = SQL & ", codigo_aliquota_isento bigint NOT NULL "
   '   SQL = SQL & ", icms_isento bigint NOT NULL "
   '   SQL = SQL & " ,CONSTRAINT [PK_ALIQUOTA_UF] PRIMARY KEY CLUSTERED"
   '   SQL = SQL & " ("
   '   SQL = SQL & " [CODIGO] Asc"
   '   SQL = SQL & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
   '   SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
   '   CONECTA_RETAGUARDA.Execute SQL

   '   'CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (1,'AC',8,17.0000,17.0000,'ACRE',8,8,17.0000,'2',12,8,17.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (2,'AL',5,0.0000,0.0000,'ALAGOAS',5,5,0.0000,'2',27,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (3,'AM',5,0.0000,0.0000,'AMAZONAS',5,5,0.0000,'2',13,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (4,'AP',5,0.0000,0.0000,'AMAPA',5,5,0.0000,'2',16,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (5,'BA',5,0.0000,0.0000,'BAHIA',5,5,0.0000,'2',29,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (6,'CE',5,0.0000,0.0000,'CEARA',5,5,0.0000,'2',23,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (7,'DF',8,17.0000,17.0000,'DISTRITO FEDERAL',8,8,17.0000,'2',53,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (8,'ES',5,0.0000,0.0000,'ESPIRITO SANTO',5,5,0.0000,'2',32,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (9,'MA',5,0.0000,0.0000,'MARANHAO',5,5,0.0000,'2',21,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (10,'MG',5,0.0000,0.0000,'MINAS GERAIS',5,5,0.0000,'2',31,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (11,'MT',5,0.0000,0.0000,'MATO GROSSO',5,5,0.0000,'2',51,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (12,'MS',5,0.0000,0.0000,'MATO GROSSO DO SUL',5,5,0.0000,'2',50,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (13,'PA',5,0.0000,0.0000,'PARA',5,5,0.0000,'2',15,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (14,'PB',5,0.0000,0.0000,'PARAIBA',5,5,0.0000,'2',25,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (15,'PE',5,0.0000,0.0000,'PERNAMBUCO',5,5,0.0000,'2',26,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (16,'PI',5,0.0000,0.0000,'PIAUI',5,5,0.0000,'2',22,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (17,'PR',5,0.0000,0.0000,'PARANA',5,5,0.0000,'2',41,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (18,'RJ',5,0.0000,0.0000,'RIO DE JANEIRO',5,5,0.0000,'2',33,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (19,'RO',5,0.0000,0.0000,'RONDONIA',5,5,0.0000,'2',11,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (20,'RN',5,0.0000,0.0000,'RIO GRANDE DO NORTE',5,5,0.0000,'2',24,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (22,'RS',5,0.0000,0.0000,'RIO GRANDE DO SUL',5,5,0.0000,'2',43,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (23,'SC',5,0.0000,0.0000,'SANTA CATARINA',5,5,0.0000,'2',42,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (24,'SE',5,0.0000,0.0000,'SERGIPE',5,5,0.0000,'2',28,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (25,'SP',5,0.0000,0.0000,'SAO PAULO',5,5,0.0000,'2',35,1,0.0000)"
      
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (26,'TO',5,0.0000,0.0000,'TOCANTINS',5,5,0.0000,'2',17,1,0.0000)"
   '   CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (27,'GO',4,12.0000,12.0000,'GOIAS',4,5,0.0000,'2',52,1,0.0000)"
   '   Else ' CASO EXISTA A TABELA E NAO EXISTA OS REGISTROS , CRIA OS 27 ESTADOS COM SUAS RESPECTIVAS ALIQUOTAS
   '      TabTemp.Open "select ISNULL(COUNT(*), 0) AS QTD from [dbo].[ALIQUOTA_UF]", CONECTA_RETAGUARDA, , , adCmdText
   '      If TabTemp!QTD = 0 Then
   '           CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (1,'AC',8,17.0000,17.0000,'ACRE',8,8,17.0000,'2',12,8,17.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (2,'AL',5,0.0000,0.0000,'ALAGOAS',5,5,0.0000,'2',27,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (3,'AM',5,0.0000,0.0000,'AMAZONAS',5,5,0.0000,'2',13,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (4,'AP',5,0.0000,0.0000,'AMAPA',5,5,0.0000,'2',16,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (5,'BA',5,0.0000,0.0000,'BAHIA',5,5,0.0000,'2',29,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (6,'CE',5,0.0000,0.0000,'CEARA',5,5,0.0000,'2',23,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (7,'DF',8,17.0000,17.0000,'DISTRITO FEDERAL',8,8,17.0000,'2',53,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (8,'ES',5,0.0000,0.0000,'ESPIRITO SANTO',5,5,0.0000,'2',32,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (9,'MA',5,0.0000,0.0000,'MARANHAO',5,5,0.0000,'2',21,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (10,'MG',5,0.0000,0.0000,'MINAS GERAIS',5,5,0.0000,'2',31,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (11,'MT',5,0.0000,0.0000,'MATO GROSSO',5,5,0.0000,'2',51,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (12,'MS',5,0.0000,0.0000,'MATO GROSSO DO SUL',5,5,0.0000,'2',50,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (13,'PA',5,0.0000,0.0000,'PARA',5,5,0.0000,'2',15,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (14,'PB',5,0.0000,0.0000,'PARAIBA',5,5,0.0000,'2',25,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (15,'PE',5,0.0000,0.0000,'PERNAMBUCO',5,5,0.0000,'2',26,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (16,'PI',5,0.0000,0.0000,'PIAUI',5,5,0.0000,'2',22,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (17,'PR',5,0.0000,0.0000,'PARANA',5,5,0.0000,'2',41,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (18,'RJ',5,0.0000,0.0000,'RIO DE JANEIRO',5,5,0.0000,'2',33,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (19,'RO',5,0.0000,0.0000,'RONDONIA',5,5,0.0000,'2',11,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (20,'RN',5,0.0000,0.0000,'RIO GRANDE DO NORTE',5,5,0.0000,'2',24,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (21,'RR',1,0.0000,0.0000,'RORAIMA',1,1,0.0000,'2',14,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (22,'RS',5,0.0000,0.0000,'RIO GRANDE DO SUL',5,5,0.0000,'2',43,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (23,'SC',5,0.0000,0.0000,'SANTA CATARINA',5,5,0.0000,'2',42,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (24,'SE',5,0.0000,0.0000,'SERGIPE',5,5,0.0000,'2',28,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (25,'SP',5,0.0000,0.0000,'SAO PAULO',5,5,0.0000,'2',35,1,0.0000)"
         
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (26,'TO',5,0.0000,0.0000,'TOCANTINS',5,5,0.0000,'2',17,1,0.0000)"
   '      CONECTA_RETAGUARDA.Execute "INSERT INTO [dbo].[ALIQUOTA_UF] (codigo, estado, codigo_aliquota, aliquota, aliquota_nc, descricao, codigo_aliquota_nc, codigo_aliquota_substituicao, aliquota_substituicao, empresa_id,codigo_uf, codigo_aliquota_isento, icms_isento) VALUES (27,'GO',4,12.0000,12.0000,'GOIAS',4,5,0.0000,'2',52,1,0.0000)"
   '      End If
   'End If
   'If TabTemp.State = 1 Then _
      TabTemp.Close

MsgBox "ok"
End Sub

Private Sub cmdTributacao_Click()
   MsgBox "RODAR SCRIPT"
End Sub
