VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmINICIO 
   BackColor       =   &H80000002&
   Caption         =   "Sistema de Automação Comercial"
   ClientHeight    =   5880
   ClientLeft      =   3375
   ClientTop       =   1785
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "INICIO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2.94851e8
   ScaleMode       =   0  'User
   ScaleWidth      =   5.82555e8
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   240
      Top             =   4680
   End
   Begin MSComctlLib.Toolbar barINI 
      Align           =   1  'Align Top
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1693
      ButtonWidth     =   2646
      ButtonHeight    =   1535
      Appearance      =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido Venda"
            Key             =   "req"
            Object.ToolTipText     =   "Pedido Venda"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cadastro Cliente"
            Key             =   "cliente"
            Object.ToolTipText     =   "Cadastro Cliente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cadastro Produto"
            Key             =   "produto"
            Object.ToolTipText     =   "Cadastro Produto"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Recebimento"
            Key             =   "recebimento"
            Object.ToolTipText     =   "Recebimento"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Entrega"
            Key             =   "entrega"
            Object.ToolTipText     =   "Entrega"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "O.S."
            Key             =   "os"
            Object.ToolTipText     =   "Ordem Serviço"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar Sistema"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1080
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":71E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":8393
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":93E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":AAB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":BE65
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":D28D
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":D6DF
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":F880
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":11862
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INICIO.frx":130FE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   840
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   128
      ImageHeight     =   128
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INICIO.frx":144FB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm Sweda 
      Left            =   120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar BARI 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5505
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "INICIO.frx":15B7D
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCNPJ 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lblCNPJ 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblEstabNome 
      AutoSize        =   -1  'True
      Caption         =   "Estabelecimento"
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblEstabelecimento 
      AutoSize        =   -1  'True
      Caption         =   "Estabelecimento"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Menu mnucad2 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnucad2Emp 
         Caption         =   "Cadastra &Empresa"
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu mnuParametro 
         Caption         =   "Cadastra &Parametros Empresa"
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnucadUsu 
         Caption         =   "Cadastro de &Usuário"
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Cadastro de &Clientes"
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFornec 
         Caption         =   "Cadastro de &Fornecedores"
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEquipe 
         Caption         =   "Cadastro Equipe Venda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendedor 
         Caption         =   "Cadastro &Vendedor"
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu mnucad2Cartao 
         Caption         =   "Cadastro Cartão Barras"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "Cadastro &Transportadora"
         Shortcut        =   ^H
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIBGE 
         Caption         =   "Cadastro Código &IBGE"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCEP 
         Caption         =   "Cadastro C&EP"
         Shortcut        =   ^J
         Visible         =   0   'False
      End
      Begin VB.Menu mnupRG 
         Caption         =   "Programa &Menu"
         Shortcut        =   ^K
         Visible         =   0   'False
      End
      Begin VB.Menu mnuACESSO 
         Caption         =   "Controle de &Acesso"
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTurno 
         Caption         =   "Cadastro Turno Trabalho"
         Visible         =   0   'False
      End
      Begin VB.Menu mnucadSenha 
         Caption         =   "Troca &Senha"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnucadTr 
         Caption         =   "-"
      End
      Begin VB.Menu mnucadFim 
         Caption         =   "F&im"
      End
   End
   Begin VB.Menu mnuCons 
      Caption         =   "C&onsultas"
      Begin VB.Menu mnuConsPedido 
         Caption         =   "&Pedido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsVendaCusto 
         Caption         =   "&Venda/Custo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsultaProduto 
         Caption         =   "P&roduto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsProdSimp 
         Caption         =   "Produto &Simplificado"
         Shortcut        =   ^Z
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsCliente 
         Caption         =   "&Cliente/Venda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsCliSimp 
         Caption         =   "C&liente Simplificado"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsCliRel 
         Caption         =   "Cliente Por Estabelecimento"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsCliRelVenda 
         Caption         =   "Cliente/Estabelecimento/Venda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsPedidoLiberacao 
         Caption         =   "Consulta Liberação Pedido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsPedidoAtendente 
         Caption         =   "Consulta Pedido Atendente"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsPedidoHora 
         Caption         =   "Consutla Pedido Hora"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsPedidoFatura 
         Caption         =   "Pedido Faturamento"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsAgendaServico 
         Caption         =   "Agenda Serviços"
      End
      Begin VB.Menu mnuComissao 
         Caption         =   "Com&issão"
         Visible         =   0   'False
         Begin VB.Menu mnuComissaoRel 
            Caption         =   "&Relatório"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuComissaonada 
            Caption         =   ""
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuConsnada 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuCAIXA 
      Caption         =   "Cai&xa"
      Begin VB.Menu mnuCAIXAAbreBalcao 
         Caption         =   "Abre Caixa Balcão"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCancelaPedido 
         Caption         =   "Cancela &Pedido"
         Visible         =   0   'False
         Begin VB.Menu mnuCAIXAbalcaonada 
            Caption         =   ""
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuTesoura 
         Caption         =   "Caixa Tesouraria"
         Begin VB.Menu mnuAbCxTesoura 
            Caption         =   "Abertura/Fechamento Tesouraria"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLancCxTeosura 
            Caption         =   "Lancamento Caixa Tesouraria"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAbCxTesouraReabre 
            Caption         =   "Reabertura Tesouraria"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTesouranada 
            Caption         =   ""
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuCAIXAbalcaoRELdiario 
         Caption         =   "Relatório Caixa"
      End
      Begin VB.Menu mnuCAIXAnada 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuESTOQUE 
      Caption         =   "&Estoque"
      Begin VB.Menu mnuEstoqueCAD 
         Caption         =   "Cadastros"
         Begin VB.Menu mnuEstoqueCADPRODUTO 
            Caption         =   "&Produto"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuESTOQUEtabelapreco 
            Caption         =   "&Tabela Preço"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueCADMarkup 
            Caption         =   "Atualiza Preço Produto"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueCADIVA 
            Caption         =   "IVA UF Produto"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueCADProdPromo 
            Caption         =   "Cadastro Promoção Produto"
         End
         Begin VB.Menu mnuEstoqueestnada 
            Caption         =   ""
         End
      End
      Begin VB.Menu mnuEstoqueProc 
         Caption         =   "&Processo"
         Begin VB.Menu mnuEstoqueProcEntra 
            Caption         =   "Nota Fiscal Entrada"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueProcCancela 
            Caption         =   "Cancelar Nota Fiscal Entrada"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueProcEntradaAvulsa 
            Caption         =   "Entrada Produto Estoque"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueProcnada 
            Caption         =   ""
         End
      End
      Begin VB.Menu mnuEstoqueCons 
         Caption         =   "Consultas"
         Begin VB.Menu mnuEstoqueConsCurvaABC 
            Caption         =   "Curva ABC"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueConsEntrada 
            Caption         =   "Consulta Nf Entrada"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueConsEst 
            Caption         =   "Posição Estoque"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueConsProdPromo 
            Caption         =   "Produtos Promoção"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueConsItens 
            Caption         =   "Consulta Faturamento Produtos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueConsnada 
            Caption         =   ""
         End
      End
      Begin VB.Menu mnuEstoqueRel 
         Caption         =   "Relatórios"
         Begin VB.Menu mnuEstoqueRelEntrada 
            Caption         =   "Rel de Entrada"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueRellISTAR 
            Caption         =   "Lista de Preços"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueRelFamilia 
            Caption         =   "Produto/Familia"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueRelContador 
            Caption         =   "Inventário/Contador"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEstoqueRelnada 
            Caption         =   ""
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuESTOQUEnada 
            Caption         =   ""
         End
      End
   End
   Begin VB.Menu mnuInventario 
      Caption         =   "&Inventário"
      Begin VB.Menu mnuContagem 
         Caption         =   "Acerto/Contagem Estoque"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInventarioCons 
         Caption         =   "Consulta Inventário"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInventarioTransf 
         Caption         =   "Transferência Estoque"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInventarioProcMP 
         Caption         =   "MP Fornecedor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInventarionada 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFinanceiro 
      Caption         =   "&Financeiro"
      Begin VB.Menu mnuLancReceber 
         Caption         =   "Contas a Receber"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConsultaReceber 
         Caption         =   "Consultar Contas a Receber"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCONTASPAGAR 
         Caption         =   "Contas a Pagar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConPagar 
         Caption         =   "Consultar Contas a Pagar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCancelaVenda 
         Caption         =   "Cancela Venda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCENTROCUSTO 
         Caption         =   "Cadastro Centro de Custo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBANCO 
         Caption         =   "Banco"
         Visible         =   0   'False
         Begin VB.Menu mnuCADBANCO 
            Caption         =   "Cadastro Banco"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAGENCIA 
            Caption         =   "Cadastro Agencia"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuConta 
            Caption         =   "Cadastro Conta Corrente"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCheque 
            Caption         =   "Cadastra Cheque"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCONCHEQUE 
            Caption         =   "Consulta Cheque"
            Shortcut        =   ^X
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBANCOnada 
            Caption         =   ""
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuFinanceiroACFUNC 
         Caption         =   "Acerto de Funcionário"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFinanceiroACcli 
         Caption         =   "Acerto de Clientes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFinanceironada 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuNFe 
      Caption         =   "&NFE"
      Begin VB.Menu mnuNfePedido 
         Caption         =   "Pedido NFe"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNfeEmissao 
         Caption         =   "Consulta NFe"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNfeManut 
         Caption         =   "Manutenção NFe"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNfeDevolucao 
         Caption         =   "NFe Diversas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNFeCarta 
         Caption         =   "Carta de Correção"
         Visible         =   0   'False
      End
      Begin VB.Menu mnNfeMegasimNfe 
         Caption         =   "Megasim NFe"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNFeContador 
         Caption         =   "Enviar XML Contador"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNFenada 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuCOMPRAS 
      Caption         =   "Com&pras"
      Begin VB.Menu mnuPEDIDOCOMPRAS 
         Caption         =   "Pedido de Compras"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCOMPRASnada 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuImporta 
      Caption         =   "Integração"
      Begin VB.Menu mnuImpCEP 
         Caption         =   "Importa CEP e IBGE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImporta12741 
         Caption         =   "Imporatação Lei 12.741"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImportaToscana 
         Caption         =   "Importação/Atualização Produto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImportaMGV 
         Caption         =   "Integração MGV"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImportaU 
         Caption         =   ""
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuProd 
      Caption         =   "Produção"
      Begin VB.Menu mnuProdPerda 
         Caption         =   "Registro Perda Produção"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdProducao 
         Caption         =   "Registro de Produção"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdRPPV 
         Caption         =   "RESUMO PRODUÇAO/PERDA/VENDAS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdnada 
         Caption         =   ""
      End
      Begin VB.Menu mnuProdVersao 
         Caption         =   "Sobre"
      End
      Begin VB.Menu mnuProdU 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "frmINICIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
   Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
   Private Declare Function BlockInput Lib "user32" (ByVal Block As Boolean) As Boolean

   Private Const MF_BYPOSITION = &H400&

   'Public WithEvents EasyTEF           As EasyTEF.EasyTEFDiscado
   Public NumeroCupomFiscal            As String
   Public UsuarioNaoQuerOutraFormaPgto As Boolean
   Public TEF                          As Boolean
   Public FuncaoAdministrativa         As Boolean

   Dim Rede, Nsu, Finalizacao          As String
   Dim VALOR                           As Double

   Dim logKeypress                     As Boolean

   Const ECF_RETORNO_OK                As Integer = 1
   Const FORMA_PGTO_CARTAO             As String = "Cartao"
   Const FORMA_PGTO_CHEQUE             As String = "Cheque"

   Public strSQL

Private Sub Timer1_Timer()
   VERIFICA_SISTEMA
End Sub

Private Sub Form_Load()
   FadeForm Me, False, True, True

   '// desabilita o botão fechar
   REMOVE_MENU

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   SqL2 = App.Major & App.Revision

   If UCase(VERSAO_APLICATIVO) <> SqL2 Then
      SQL = "update ESTABELECIMENTO set versao_aplicativo = '" & SqL2 & "'"
      SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      SqL2 = ""
      SQL = ""
      CRITERIO_A = ""
      SQL3 = ""

      ATUALIZA_TABELA_EMPRESA
      CRIA_IMPREL
      CHECA_TABELA_PRODUTO
      CHECA_TABELA_ESTOQUE
      ATUALIZA_TABELA_PESSOA
      ATUALIZA_ESTABELECIMENTO
      ATUALIZA_TABELA_FORMAPAGTO
      If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDO", "U") = True Then _
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "CARTAOBARRA_ID", "PEDIDO") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO ADD CARTAOBARRA_ID BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESC_A", "DESCR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'DESCR.DESC_A'" & "," & "'DESCRICAO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO", "DESCR") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'DESCR.TIPO'" & "," & "'TIPO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO", "DESCR") = True Then _
         Alteração_Definição_Campo_Tabela "TIPO", "nvarchar(2) NOT NULL", "DESCR", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USUARIO_ID", "ERRO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ERRO ADD USUARIO_ID INT "
   End If

   CHECA_CADASTRO_MENU
End Sub

Private Sub Form_Activate()
'On Error GoTo ERRO_TRATA

   lblEmpresa.Caption = "Empresa: " & EMPRESA_ID_N
   lblEmpresa.Visible = True

   txtCNPJ.Mask = "##.###.###/####-##"
   txtCNPJ.PromptInclude = False
      txtCNPJ.Text = CNPJ_EMPRESA_N
   txtCNPJ.PromptInclude = True

   lblCNPJ.Caption = "CNPJ: " & txtCNPJ.Text
   lblCNPJ.Visible = True

   lblEstabelecimento.Caption = "Estabelecimento: " & ESTABELECIMENTO_ID_N
   lblEstabelecimento.Visible = True

   lblUsuario.Caption = "" & USU_LOGADO
   lblUsuario.Visible = True

   If INDR_CAIXA = True Then
      If (App.PrevInstance) Then
          Dim nome_tela As String
          nome_tela = App.Title
          App.Title = "Já estou em execução, frmINICIO !!!"
          'AppActivate  nome_tela
          SendKeys "%R", True
          MsgBox "Já em execução !!!"
          End
          Exit Sub
      End If
   End If

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   logKeypress = False

   If USUARIO_ID_N > 0 Then _
      CONTROLE_ACESSO

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select descricao,localizacao from ESTABELECIMENTO"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         SQL = Trim(TabTemp.Fields(0).Value)

      lblEstabNome.Caption = Trim(TabTemp.Fields("localizacao").Value)
      lblEstabNome.Visible = True
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   Me.Caption = "Sistema de Gestão Comercial - " & NOME_EMPRESA_A & _
                " | CPU: " & NUMERO_CAIXA_CPU & _
                " | DB: " & NOME_BANCO_DADOS & _
                " | " & SQL & _
                " | Versão: " & VERSAO_APLICATIVO

   If MULT_EMPRESA_B = True Then
      MOSTRA_RODAPE "F2-Cliente", "F4-Produto", "F5-Pedido", "F6-Caixa", ""
      Else: MOSTRA_RODAPE "F1-Logon", "F2-Cliente", "F4-Produto", "F5-Pedido", "F6-Caixa"
   End If

   If TRAZ_TIPO_USUARIO = 1 Then _
      VENDA_PRODUTO

If Left(NOME_EMPRESA_A, 2) = "PS" Then
   Timer1.Enabled = False
End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF1
         If MULT_EMPRESA_B = False Or USUARIO_ID_N = 144 Then
            INDR_FIM = True
            frmLOGON.Show
         End If
      Case vbKeyF2
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1
      Case vbKeyF4
         frmCADASTROPRODUTO.Show 1
      Case vbKeyF5
         VENDA_PRODUTO
      Case vbKeyF6
         frmDISPLAYEMISSOR.Show 1
      Case vbKeyF7
         'CRITERIO_A = InputBox("Entre com a senha", "Controle de Acesso")
         'If UCase(CRITERIO_A) = UCase("vacaveia") Then _
            frmImportaTabelaPreco.Show 1
      Case vbKeyF8
         If USUARIO_ID_N = 144 Then _
            frmSenha.txtSenha.Text = "vacaveia2"

         frmSenha.Show 1

         If UCase(CRITERIO_A) = UCase("vacaveia") Then _
            frmATUALIZACAO.Show 1

         If UCase(CRITERIO_A) = UCase("vacaveia2") Then _
            frmATUALIZACAO2.Show 1

         If UCase(CRITERIO_A) = UCase("PROTECshf") Then _
            frmCRIPTO.Show 1
      Case vbKeyF10
         If USUARIO_ID_N = 144 Then
            CRITERIO_A = InputBox("Entre com a senha", "Controle de Acesso")
            If UCase(CRITERIO_A) = UCase("fdp") Then _
               frmINTEGRA.Show 1
               'frmMigra.Show
         End If
      Case vbKeyF11
         CRITERIO_A = InputBox("Entre com a senha", "Controle de Acesso")
         If UCase(CRITERIO_A) = UCase("fdp") Then _
            FrmControleAcesso.Show 1
         End
      Case vbKeyF12
         If INDR_CAIXA = False And INDR_REMOTO = True Then _
            Call ExitWindowsEx(0, 0)
         End
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'Fecha conexão com banco de dados
   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

   End

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub mnuCAIXAAbreBalcao_Click()
   frmCaixa.Show 1
End Sub

Private Sub mnuConsAgendaServico_Click()
   frmAGENDASERVICO.Show 1
End Sub

Private Sub mnuConsCliRel_Click()
   TIPO_PESSOA_CADASTRO = "ESTABELECIMENTO"
   frmClienteVendedor.Show 1
End Sub

Private Sub mnuConsCliRelVenda_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmClienteVendedor.Show 1
End Sub

Private Sub mnuConsPedidoFatura_Click()
   frmPEDIDOFATURA.Show 1
End Sub

Private Sub mnuConsPedidoHora_Click()
   frmCONSULTAPEDIDOHORA.Show 1
End Sub

Private Sub mnuEstoqueConsProdPromo_Click()
   frmDisplayPromocao.Show 1
End Sub

Private Sub mnuEstoqueCADProdPromo_Click()
   frmPRODUTOPROMOCAO.Show 1
End Sub

Private Sub mnuEstoqueConsItens_Click()
    frmPedidoItemConsulta.Show 1
End Sub

Private Sub mnuInventarioProcMP_Click()
    frmEstoqueMP.Show 1
End Sub

Private Sub mnuAbCxTesouraReabre_Click()
   SQL = "CAIXATESORARIA"
   frmCAIXATESORARIAAFREABRE.Show 1
End Sub

Private Sub mnuAbCxBalcalReabre_Click()
   SQL = "CAIXADIA"
   frmCAIXATESORARIAAFREABRE.Show 1
End Sub

Private Sub mnuConsultaPagar_Click()
   INDR_RECEITA = 2
   frmFINCONSULTAFATURA.Show 1
End Sub

Private Sub mnNfeMegasimNfe_Click()
   LoadEXE ("C:\MEGASIM\MEGASIMNFe\2010.exe")
End Sub

Private Sub mnucad2Cartao_Click()
   frmCADASTROCARTAOBARRA.Show 1
End Sub

Private Sub mnuConsPedidoAtendente_Click()
   frmPedidoConsultaAtendente.Show 1
End Sub

Private Sub mnuConsPedidoLiberacao_Click()
   frmPedidoConsultaLibera.Show 1
End Sub

Private Sub mnuConsProdSimp_Click()
   CHAMA_PRODUTO_SIMPLIFICADO
End Sub

Private Sub mnuConsultaReceber_Click()
   INDR_RECEITA = 1
   frmFINCONSULTAFATURA.Show 1
End Sub

Private Sub mnuLancPagar_Click()
   INDR_RECEITA = 2
   frmFINGERALANC.Show 1
End Sub

Private Sub mnuEstoqueConsEst_Click()
   frmESTOQUEPOSICAO.Show 1
End Sub

Private Sub mnuEstoqueProcEntradaAvulsa_Click()
   frmEntradaEstoque.Show 1
End Sub

Private Sub mnuEstoqueRelContador_Click()
   frmInventarioProduto.Show 1
End Sub

Private Sub mnuEstoqueRelFamilia_Click()
   frmProdutoFamilia.Show 1
End Sub

Private Sub MNUESTOQUEtabelapreco_Click()
   frmTabelaPreco.Show 1
End Sub

Private Sub mnuFinanceiroACcli_Click()
   frmAcertoCliente.Show 1
End Sub

Private Sub mnuFinanceiroACFUNC_Click()
   frmAcertoFunc.Show 1
End Sub

Private Sub mnuImporta12741_Click()
   frmIMPORTA12741.Show 1
End Sub

Private Sub mnuImportaMGV_Click()
   frmMGV.Show 1
End Sub

Private Sub mnuInventarioTransf_Click()
   frmEstoqueTransfFilial.Show 1
End Sub

Private Sub mnuLancReceber_Click()
   INDR_RECEITA = 1
   frmFINGERALANC.Show 1
End Sub

Private Sub mnucad2Emp_Click()
   If Libera_Acesso("frmcadastroempresa") Then
      frmCADASTROEMPRESA.Show 1
      Else: MsgBox "Acesso não permitido."
   End If
End Sub

Private Sub mnuAbCxTesoura_Click()
    frmCAIXATESORARIAAF.Show 1
End Sub

Private Sub MNUACESSO_Click()
    FrmControleAcesso.Show 1
End Sub

Private Sub mnuProdRPPV_Click()
'On Error GoTo ERRO_TRATA

   If Libera_Acesso("frmProducaoPerdaVenda") Then
      frmProducaoPerdaVenda.Show 1
      Else: MsgBox "Acesso não permitido."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "mnuProdRPPV_Click"
End Sub

Private Sub mnuTurno_Click()
'On Error GoTo ERRO_TRATA

   If Libera_Acesso("frmturnocadastro") Then
      frmTurnoCadastro.Show 1
      Else: MsgBox "Acesso não permitido."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "mnuTurno_Click"
End Sub

Private Sub mnucadSenha_Click()
'On Error GoTo ERRO_TRATA

   If Libera_Acesso("frmlojtrocasenha") Then
      frmTROCASENHA.Show 1
      Else: MsgBox "Acesso não permitido."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "mnucadSenha_Click"
End Sub

Private Sub mnuCAIXAbalcaoRELdiario_Click()
    frmCAIXAREL.Show 1
End Sub

Private Sub mnuCancelaPedido_Click()
   If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
      frmPedidoCancela.Show 1
      Else: MsgBox "Não permitido."
   End If
End Sub

Private Sub mnuCancelaVenda_Click()
   If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
      frmPedidoCancela.Show 1
      Else: MsgBox "Não permitido."
   End If
End Sub

Private Sub mnuEstoqueProcCancela_Click()
    frmNFECANCELA.Show 1
End Sub

Private Sub MNUCENTROCUSTO_Click()
    frmCADASTROCENTROCUSTO.Show
End Sub

Private Sub mnuCEP_Click()
    frmCADASTROCEP.Show 1
End Sub

Private Sub mnuCheque_Click()
    frmCHEQUECADASTRO.Show 1
End Sub

Private Sub mnuClientes_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaCadastro.Show 1
End Sub

Private Sub mnuComissaoRel_Click()
   frmVENDACOMISSAO.Show 1
End Sub

Private Sub MNUCONCHEQUE_Click()
    frmCHEQUECONSULTA.Show 1
End Sub

Private Sub mnuConPagar_Click()
   INDR_RECEITA = 2
   frmFINCONSULTAFATURA.Show 1
End Sub

Private Sub mnuConsCliente_Click()
   frmCLIENTECONSULTA.Show 1
End Sub

Private Sub mnuConsCliSimp_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
End Sub

Private Sub mnuFornec_Click()
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaCadastro.Show 1
End Sub

Private Sub mnuTrans_Click()
    TIPO_PESSOA_CADASTRO = "TRANSPORTADORA"
    frmPessoaCadastro.Show 1
End Sub

Private Sub mnuEquipe_Click()
    TIPO_PESSOA_CADASTRO = "EQUIPE"
    frmPessoaCadastro.Show 1
End Sub

Private Sub mnuConsPedido_Click()
   CRITERIO_A = ""
   CNPJCPF_A = ""
   frmPedidoConsulta.Show 1
End Sub

Private Sub mnuConsultaProduto_Click()
    frmProdutoConsulta.Show 1
End Sub

Private Sub mnuConsVendaCusto_Click()
   frmVENDACUSTO.Show 1
End Sub

Private Sub mnuContagem_Click()
    frmINVENTARIO.Show 1
End Sub

Private Sub MNUCONTASPAGAR_Click()
   INDR_RECEITA = 2
   frmFINGERALANC.Show 1
End Sub

Private Sub mnuEstoqueProcEntra_Click()
   frmNOTAENTRADA.Show 1
End Sub

Private Sub mnuEstoqueConsEntrada_Click()
   frmNOTACONSULTA.Show 1
End Sub

Private Sub mnuEstoqueCADIVA_Click()
   'frmCADASTROIVAUFPRODUTO.Show 1
End Sub

Private Sub mnuEstoqueCADPRODUTO_Click()
   frmCADASTROPRODUTO.Show 1
End Sub

Private Sub mnuIBGE_Click()
    frmIBGECADASTRO.Show 1
End Sub

Private Sub mnuImpCEP_Click()
    frmImporta.Show 1
End Sub

Private Sub mnuInventarioCons_Click()
   frmInventarioConsulta.Show 1
End Sub

Private Sub mnuLancCxTeosura_Click()
    frmCAIXATESORARIA.Show 1
End Sub

Private Sub mnuEstoqueRellISTAR_Click()
    frmLstPreco.Show 1
End Sub

Private Sub mnuManutBanco_Click()
   frmBACKUP.TimerBar.Enabled = False
   frmBACKUP.Show 1
End Sub

Private Sub mnuEstoqueCADMarkup_Click()
    frmCADASTROTAXAMARC.Show 1
End Sub

Private Sub mnuNFeCarta_Click()
   frmCCe.Show 1
End Sub

Private Sub mnuNFeContador_Click()
   frmNFeContador.Show 1
End Sub

Private Sub mnuNfeDevolucao_Click()
   frmNFEDIVERSAS.Show 1
End Sub

Private Sub mnuNfeEmissao_Click()
   frmNFECONSULTA.Show 1
End Sub

Private Sub mnuNfePedido_Click()
   VENDA_PRODUTO
End Sub

Private Sub mnuNfeManut_Click()
    frmNFeMANUT.Show 1
End Sub

Private Sub mnuParametro_Click()
   frmCADASTROPARAMETRO.Show 1
End Sub

Private Sub MNUpRG_Click()
   frmCADASTROPROGRAMA.Show 1
End Sub

Private Sub mnuEstoqueRelEntrada_Click()
    frmENTRADAREL.Show 1
End Sub

Private Sub mnuProdPerda_Click()
   frmProducaoRegistroPerdaCadastro.Show 1
End Sub

Private Sub mnuProdProducao_Click()
   frmProducaoRegistroProducaoCadastro.Show 1
End Sub

Private Sub mnuutilVendaProduto_Click()
   VENDA_PRODUTO
End Sub

Private Sub mnuVendedor_Click()
   frmCADASTROVENDEDOR.Show 1
End Sub

Private Sub mnuPEDIDOCOMPRAS_Click()
    frmPedidoCompraCadastro.Show 1
End Sub

Private Sub BARINI_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "os"
         frmOSInicio.Show 1
      Case "entrega"
         frmProdutoEntrega.Show 1
      Case "cliente"
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1
      Case "produto"
         frmCADASTROPRODUTO.Show 1
      Case "recebimento"
         frmDISPLAYEMISSOR.Show 1
      Case "req"
         VENDA_PRODUTO
      Case "sair"
         On Error Resume Next
         If INDR_CAIXA = False And INDR_REMOTO = True Then _
            Call ExitWindowsEx(0, 0)

         End
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BARINI_ButtonClick"
End Sub

Private Sub mnucadFim_Click()
'On Error GoTo ERRO_TRATA

   If INDR_CAIXA = False And INDR_REMOTO = True Then _
      Call ExitWindowsEx(0, 0)

   End

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "mnucadFim_Click"
End Sub

Private Sub mnucadUsu_Click()
'On Error GoTo ERRO_TRATA

   If Libera_Acesso("frmLojCadUsu") Then
      frmCADASTROUSUARIO.Show 1
      Else: MsgBox "Acesso não permitido."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "mnucadUsu_Click"
End Sub

Sub VENDA_PRODUTO()
   PEDIDO_ID_N = 0
   INDR_PEDIDO_VENDA = True
   If TRAZ_TIPO_USUARIO = 7 Then
      If CHECAR_CAIXA(USUARIO_ID_N, Date) = True Then
         If USA_TAB_PRECO_B = False Then
            frmPedidoVenda.Show 1
            Else: frmPEDIDOBALCAO.Show 1
         End If
      End If
      Else
         If USA_TAB_PRECO_B = False Then
            If Trim(UCase(USU_LOGADO)) = "BALCAO" Then
               frmPedidoComanda.Show 1
               Else: frmPedidoVenda.Show 1
            End If
            Else: frmPEDIDOBALCAO.Show 1
         End If
   End If
   INDR_PEDIDO_VENDA = False
End Sub

Private Sub CONTROLE_ACESSO()
   Dim ctlMenu As Control
   Dim TabMenu As New ADODB.Recordset

   'ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

'====================DESABILITA TUDO
   barINI.Buttons.item(1).Visible = False 'Pedido Venda
   barINI.Buttons.item(2).Visible = False 'Cadastro Cliente
   barINI.Buttons.item(3).Visible = False 'Cadastro Produto
   barINI.Buttons.item(4).Visible = False 'Recebimento
   barINI.Buttons.item(5).Visible = False 'Entrega
   barINI.Buttons.item(6).Visible = False 'O.S.
   barINI.Buttons.item(7).Visible = False 'Sair

   Me.mnuEstoqueCADPRODUTO.Visible = False
   mnuCancelaPedido.Visible = False
   mnuCAIXAbalcaoRELdiario.Visible = False

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MENUID, DescMenu from MENU"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   TabTemp.MoveFirst
   Do Until TabTemp.EOF
      For Each ctlMenu In Me

         If UCase(Trim(ctlMenu.Name)) = UCase(Trim(TabTemp!menuid)) Then
            On Error Resume Next

         If Left(UCase(ctlMenu.Name), 3) <> UCase("mnu") Then _
            ctlMenu.Visible = False

            Err.Clear
            Exit For
            DoEvents
         End If
      Next
      TabTemp.MoveNext
   Loop
   If TabTemp.State = 1 Then _
      TabTemp.Close

'====================HABILITA

   Err.Clear
   barINI.Visible = True

   If TabTemp.State = 1 Then _
      TabTemp.Close

   strSQL = "select MENUID, USUID, ACESSO from PERMISSAO"
   strSQL = strSQL & " Where PERMISSAO.USUID = " & USUARIO_ID_N
   strSQL = strSQL & " Order by PERMISSAO.MENUID"
   TabMenu.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabMenu.EOF Then
      TabMenu.Close
      MsgBox "Solicite ao Administrador do Sistema, para liberar acesso aos itens de Menu do Sistema !!!"
      Screen.MousePointer = vbDefault
      Exit Sub
      'End
   End If

   Screen.MousePointer = vbHourglass
   'Menus
   TabMenu.MoveFirst
   CRITERIO_A = Trim(UCase(TabMenu!menuid))
   Do Until TabMenu.EOF
      For Each ctlMenu In Me

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(1)") Then
            barINI.Buttons.item(1).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(2)") Then
            barINI.Buttons.item(2).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(3)") Then
            barINI.Buttons.item(3).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(4)") Then
            barINI.Buttons.item(4).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(5)") Then
            barINI.Buttons.item(5).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(6)") Then
            barINI.Buttons.item(6).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(7)") Then
            barINI.Buttons.item(7).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(8)") Then
            barINI.Buttons.item(8).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = UCase("barINI.Buttons.Item(9)") Then
            barINI.Buttons.item(9).Visible = True
            Exit For
         End If

         If Trim(UCase(TabMenu!menuid)) = "Form_KeyDown" Then
            logKeypress = True
            Exit For
         End If

         'ctlMenu.Visible = False
'MsgBox "LENDO OBJETOS MENU   =   " & UCase(Trim(ctlMenu.Name)) & "   /////TABELA = " & UCase(Trim(ucase(TabMenu!menuid)))
         If UCase(Trim(ctlMenu.Name)) = UCase(Trim(UCase(TabMenu!menuid))) Then

'MsgBox "LENDO OBJETOS MENU   =   " & UCase(Trim(ctlMenu.Name)) & "   /////TABELA = " & UCase(Trim(ucase(TabMenu!menuid)))

            If Not IsNull(TabMenu.Fields("acesso").Value) Then
               ctlMenu.Visible = TabMenu.Fields("acesso").Value
               'ctlMenu.Refresh
            End If

            Exit For
         End If

      Next
      TabMenu.MoveNext
   Loop
   If TabMenu.State = 1 Then _
      TabMenu.Close

   Screen.MousePointer = vbDefault

   barINI.Buttons(1).Caption = "Pedido Venda"
   barINI.Buttons(2).Caption = "Cadastro Cliente"
   barINI.Buttons(3).Caption = "Cadastro Produto"
   barINI.Buttons(4).Caption = "Recebimento"
   barINI.Buttons(5).Caption = "Entrega"
   barINI.Buttons(6).Caption = "O.S."
   barINI.Buttons(7).Caption = "Sair"

   If USA_NFe = True Then
      mnuNFe.Visible = True
      'barINI.Buttons.Item(8).Visible = True
      Else
         mnuNFe.Visible = False
         'barINI.Buttons.Item(8).Visible = False
   End If

   'Me.mnuFErros.Visible = False
   If USUARIO_ID_N = 144 Then
      'Me.mnuFErros.Visible = True
      barINI.Buttons.item(8).Visible = True
   End If
End Sub

Private Sub REMOVE_MENU()
   Dim hMenu As Long
   hMenu = GetSystemMenu(hwnd, False)
   DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Sub CHECA_CADASTRO_MENU()
'On Error GoTo ERRO_TRATA

   SQL = "delete permissao where left(menuid,3) = 'frm'"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete menu where left(menuid,3) = 'frm'"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from menu where menuid not in (select menuid from permissao)"
   CONECTA_RETAGUARDA.Execute SQL

   Dim i As Integer

   If FSO.FileExists(App.Path & "\lstmenu.txt") Then
      Dim f, sLine
      f = FreeFile

      Open App.Path & "\lstmenu.txt" For Input As f

      Do While Not EOF(f)
         Line Input #f, sLine

         CRITERIO_A = ""
         SQL3 = ""
         NUMR_SEQ_N = 0

         For i = 1 To Len(sLine)
            If Mid(sLine, i, 1) <> " " Then
               CRITERIO_A = CRITERIO_A & Mid(sLine, i, 1)
               NUMR_SEQ_N = Len(CRITERIO_A)
               Else
                  Exit For
            End If
         Next

         SQL3 = Mid(sLine, NUMR_SEQ_N + 2, Len(sLine))
         If TabTemp.State = 1 Then _
            TabTemp.Close
         SQL = "select * from MENU "
         SQL = SQL & " where menuid = '" & Trim(CRITERIO_A) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            SQL = "insert into MENU "
            SQL = SQL & "(menuid,descmenu) "
            SQL = SQL & " values("
               SQL = SQL & "'" & Trim(CRITERIO_A) & "'" 'menuid
               SQL = SQL & ",'" & Trim(SQL3) & "'"    'descmenu
            SQL = SQL & " )"
            CONECTA_RETAGUARDA.Execute SQL
            Else
               SQL = "update MENU set "
               SQL = SQL & " descmenu = '" & Trim(SQL3) & "'"
               SQL = SQL & " where menuid = '" & Trim(CRITERIO_A) & "'"
               CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CRITERIO_A = ""
         SQL3 = ""

         DoEvents

      Loop
      Close #f
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_CADASTRO_MENU"
End Sub

