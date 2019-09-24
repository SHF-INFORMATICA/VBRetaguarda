VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROPRODUTO 
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   2025
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROPRODUTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9765
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   30
      Top             =   720
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "CADASTROPRODUTO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tributação"
      TabPicture(1)   =   "CADASTROPRODUTO.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPercIVA"
      Tab(1).Control(1)=   "cmbSTAUX"
      Tab(1).Control(2)=   "cmbALIQUOTA"
      Tab(1).Control(3)=   "cmbOrigemMercadoriaAUX"
      Tab(1).Control(4)=   "cmbOrigemMercadoria"
      Tab(1).Control(5)=   "cmbST"
      Tab(1).Control(6)=   "lblST"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "lblCST"
      Tab(1).Control(9)=   "Label5"
      Tab(1).Control(10)=   "Label9"
      Tab(1).Control(11)=   "Label3"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Produto Fornecedor"
      TabPicture(2)   =   "CADASTROPRODUTO.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label16"
      Tab(2).Control(2)=   "lblProdFornec"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "Label32"
      Tab(2).Control(5)=   "Line1"
      Tab(2).Control(6)=   "txtCNPJCPF"
      Tab(2).Control(7)=   "lstProdFornec"
      Tab(2).Control(8)=   "txtFORNEC"
      Tab(2).Control(9)=   "txtCodgFornec"
      Tab(2).Control(10)=   "cmdFornec"
      Tab(2).Control(11)=   "txtCodgBarraFornec"
      Tab(2).Control(12)=   "txtCustoProdFornec"
      Tab(2).Control(13)=   "cmdCadFor"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Consulta"
      TabPicture(3)   =   "CADASTROPRODUTO.frx":5C66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtDesc2"
      Tab(3).Control(1)=   "cmdCons"
      Tab(3).Control(2)=   "chkTodos"
      Tab(3).Control(3)=   "cmdMata"
      Tab(3).Control(4)=   "lstProduto"
      Tab(3).Control(5)=   "Label6"
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtPercIVA 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -74040
         TabIndex        =   90
         ToolTipText     =   "Caso o Item For Substituicao Tributaria, Informe o Percentual de IVA , Consultar Contador! "
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox cmbSTAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72720
         TabIndex        =   89
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbALIQUOTA 
         Height          =   360
         Left            =   -66360
         TabIndex        =   86
         ToolTipText     =   "Aliquota de ICMS do produto."
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbOrigemMercadoriaAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72720
         TabIndex        =   85
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbOrigemMercadoria 
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
         Left            =   -72720
         TabIndex        =   83
         ToolTipText     =   "Selecione a nacionalidade do produto."
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cmbST 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72720
         TabIndex        =   81
         Text            =   "-- Selecione --"
         ToolTipText     =   "Selecione o CFOP para este produto."
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtDesc2 
         Height          =   360
         Left            =   -72780
         MaxLength       =   100
         TabIndex        =   78
         ToolTipText     =   "Informe "
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton cmdCons 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   -68460
         Picture         =   "CADASTROPRODUTO.frx":5C82
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   600
         Width           =   405
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   -67620
         TabIndex        =   76
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdMata 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   -66300
         TabIndex        =   75
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdCadFor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -70320
         Picture         =   "CADASTROPRODUTO.frx":6684
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Cadastro Produto"
         Top             =   1380
         Width           =   405
      End
      Begin VB.TextBox txtCustoProdFornec 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   -69120
         MaxLength       =   12
         TabIndex        =   24
         ToolTipText     =   "Valor unitario de custo do produto"
         Top             =   1860
         Width           =   1455
      End
      Begin VB.TextBox txtCodgBarraFornec 
         Height          =   375
         Left            =   -73080
         TabIndex        =   23
         ToolTipText     =   "Referencia do produto"
         Top             =   1860
         Width           =   2175
      End
      Begin VB.CommandButton cmdFornec 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   -70889
         Picture         =   "CADASTROPRODUTO.frx":BC86
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1380
         Width           =   405
      End
      Begin VB.TextBox txtCodgFornec 
         Height          =   375
         Left            =   -73080
         TabIndex        =   25
         ToolTipText     =   "Referencia do produto"
         Top             =   2340
         Width           =   2175
      End
      Begin VB.TextBox txtFORNEC 
         Height          =   375
         Left            =   -69840
         MaxLength       =   100
         TabIndex        =   27
         ToolTipText     =   "Fornecedor deste produto"
         Top             =   1380
         Width           =   4455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Código Produto                                       Descrição "
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   50
         TabIndex        =   40
         Top             =   420
         Width           =   9615
         Begin VB.CheckBox chkBarras 
            Caption         =   "Lê Barras"
            Height          =   240
            Left            =   8040
            TabIndex        =   60
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox txtProduto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            MaxLength       =   30
            TabIndex        =   0
            ToolTipText     =   "Informe o código do produto."
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtDesc 
            Height          =   360
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "Informe "
            Top             =   240
            Width           =   5535
         End
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00FFFFFF&
            Height          =   350
            Left            =   3520
            Picture         =   "CADASTROPRODUTO.frx":C688
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   50
         TabIndex        =   31
         Top             =   1080
         Width           =   9615
         Begin VB.CommandButton cmdCadPlaca 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4080
            Picture         =   "CADASTROPRODUTO.frx":D08A
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Cadastro Familia"
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox txtPercVenda 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4920
            MaxLength       =   12
            TabIndex        =   14
            ToolTipText     =   "$ Compoe Preço Venda Produto"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CheckBox chkConceder 
            Caption         =   "Conceder Produção?"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   7080
            TabIndex        =   65
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkDesconto 
            Caption         =   "Conceder Desconto Venda?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   8160
            TabIndex        =   64
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox chkBalanca 
            Caption         =   "Balança?"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   8160
            TabIndex        =   63
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox cmbMarcaAUX 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   62
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cmbMarca 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5400
            TabIndex        =   11
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox cmbTamanhoAUX 
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1680
            TabIndex        =   59
            ToolTipText     =   "Aliquota de ICMS do produto."
            Top             =   3840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cmbTamanho 
            Height          =   360
            Left            =   1710
            TabIndex        =   19
            ToolTipText     =   "Aliquota de ICMS do produto."
            Top             =   3840
            Width           =   735
         End
         Begin VB.TextBox txtPesoB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000002&
            Height          =   375
            Left            =   8040
            TabIndex        =   18
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtPesoL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000002&
            Height          =   360
            Left            =   4920
            TabIndex        =   17
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtDtUltEntrada 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   8040
            TabIndex        =   28
            ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtEmbalagem 
            Height          =   360
            Left            =   8040
            TabIndex        =   21
            ToolTipText     =   "Informe a quantidade miníma para este produto."
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox txtCustoAnterior 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtPrecoAtacado 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   16
            ToolTipText     =   "Valor unitario de atacado do produto"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtPrecoCusto 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H80000012&
            Height          =   375
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   13
            ToolTipText     =   "Valor unitario de custo do produto"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox txtPrecoVenda 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   8040
            MaxLength       =   12
            TabIndex        =   15
            ToolTipText     =   "Valor unitario de venda(varejo) do produto."
            Top             =   2880
            Width           =   1455
         End
         Begin VB.ComboBox cmbFamiliaAux 
            BackColor       =   &H80000001&
            ForeColor       =   &H80000004&
            Height          =   360
            Left            =   1080
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtUN 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   8880
            MaxLength       =   2
            TabIndex        =   6
            ToolTipText     =   "Informe a unidade (KG, UN, GR etc...) para este produto."
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtPerc 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   7920
            MaxLength       =   6
            TabIndex        =   5
            ToolTipText     =   "Percentual de comissão para este produto."
            Top             =   480
            Width           =   855
         End
         Begin VB.ComboBox cmbFamilia 
            Height          =   360
            Left            =   1080
            TabIndex        =   2
            ToolTipText     =   "Selecione o grupo do produto."
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtQtde 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   360
            Left            =   4560
            TabIndex        =   33
            ToolTipText     =   "Quantidade existente atuamente no estoque"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtEstoqueMaximo 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   6000
            TabIndex        =   3
            ToolTipText     =   "Informe a quantidade máxima para este produto."
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtEstoqueMinimo 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   6960
            TabIndex        =   4
            ToolTipText     =   "Informe a quantidade miníma para este produto."
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtRef 
            Height          =   360
            Left            =   1680
            TabIndex        =   10
            ToolTipText     =   "Referencia do produto"
            Top             =   1935
            Width           =   1455
         End
         Begin VB.ComboBox cmbSituacao 
            Height          =   360
            Left            =   1080
            TabIndex        =   7
            ToolTipText     =   "Selecione a situação para este produto"
            Top             =   960
            Width           =   1455
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            Picture         =   "CADASTROPRODUTO.frx":F1B4
            ScaleHeight     =   495
            ScaleWidth      =   600
            TabIndex        =   32
            Top             =   2760
            Width           =   600
         End
         Begin VB.ComboBox cmbTipoProd 
            Height          =   360
            Left            =   4920
            TabIndex        =   20
            ToolTipText     =   "Tipo de produto."
            Top             =   3855
            Width           =   1455
         End
         Begin VB.TextBox txtLocacao 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4800
            MaxLength       =   6
            TabIndex        =   12
            ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
            Top             =   1935
            Width           =   1455
         End
         Begin VB.TextBox txtCodgNCM 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4560
            MaxLength       =   8
            TabIndex        =   8
            ToolTipText     =   "F1-Consultar NCM"
            Top             =   945
            Width           =   2295
         End
         Begin VB.TextBox txtBarra 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1680
            TabIndex        =   9
            ToolTipText     =   "Digite Codigo de Barra do Produto"
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "%CompoeVenda"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   3255
            TabIndex        =   73
            Top             =   2880
            Width           =   1560
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Marca:"
            Height          =   255
            Left            =   4560
            TabIndex        =   61
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Medida Peça:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   330
            TabIndex        =   58
            Top             =   3840
            Width           =   1305
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Peso Bruto:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6840
            TabIndex        =   56
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Peso Liquido:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3480
            TabIndex        =   55
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚltimaEntrada:"
            Height          =   255
            Left            =   6480
            TabIndex        =   54
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Embalagem:"
            Height          =   240
            Left            =   6750
            TabIndex        =   53
            Top             =   3840
            Width           =   1200
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "CustoAnterior:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Preço Atacado:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Preço Custo:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   360
            TabIndex        =   50
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Preço Venda:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   6600
            TabIndex        =   49
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "UN"
            Height          =   240
            Left            =   9000
            TabIndex        =   48
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "%Comis."
            Height          =   240
            Left            =   7920
            TabIndex        =   47
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Ext.Mín."
            Height          =   240
            Left            =   6960
            TabIndex        =   46
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Ext.Máx."
            Height          =   240
            Left            =   6000
            TabIndex        =   45
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Familia:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Qtde.Estoque"
            Height          =   240
            Left            =   4560
            TabIndex        =   43
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Referência:"
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Situação:"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   75
            TabIndex        =   38
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Produto:"
            Height          =   255
            Left            =   3480
            TabIndex        =   37
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Locação:"
            Height          =   255
            Left            =   3720
            TabIndex        =   36
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "NCM:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3960
            TabIndex        =   35
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Codg.Barras:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lstProdFornec 
         Height          =   2415
         Left            =   -74835
         TabIndex        =   68
         Top             =   2940
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   4260
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CodgProdFornec"
            Object.Width           =   3919
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CodgBarra"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Pr.Custo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fornecedor"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fornec_id"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   -73080
         TabIndex        =   22
         Top             =   1380
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstProduto 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   79
         Top             =   960
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   7858
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qtde"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "QTD.BC."
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qtde.Dep."
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Pr.Venda"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Pr.Atacado"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Pr.Custo"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Fornecedor"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "+ Est."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "- Est."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "%"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Referência"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Grupo"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "ST"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "produto_id"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label lblST 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   92
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "IVA:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   91
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label lblCST 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CST = "
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -68295
         TabIndex        =   88
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ICMS:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -67080
         TabIndex        =   87
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Origem Mercdoria:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74640
         TabIndex        =   84
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Situação Tributária:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   82
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição Produto:"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -74640
         TabIndex        =   80
         Top             =   600
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -65280
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Preço Custo:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -70440
         TabIndex        =   72
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "CodgBarra:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   71
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label lblProdFornec 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Referência Produto/Fornecedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   -74950
         TabIndex        =   70
         Top             =   660
         Width           =   9650
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "CodgProdFornec.:"
         Height          =   240
         Left            =   -74910
         TabIndex        =   67
         Top             =   2340
         Width           =   1725
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Fornecedor:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   66
         Top             =   1380
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":102D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":10729
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":10A45
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":10E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":112ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":11741
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":11A5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":11EB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":12311
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":12BEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":13208
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROPRODUTO.frx":13889
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   7
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Cadastro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imagem"
            Key             =   "imagem"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   9960
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   0
         Top             =   400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":1405B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":15483
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":16512
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":1777A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":18E77
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":19E2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":1AF37
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":1C0D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":1D303
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROPRODUTO.frx":1D755
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   9765
      DesignHeight    =   6285
   End
End
Attribute VB_Name = "frmCADASTROPRODUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim INDR_ACHOU                   As Boolean
   Dim PRECO_CUSTO_ANTERIOR         As Double
   Dim PRECO_CUSTO_ANTERIOR_MATERIA As Double
   Dim PRECO_CUSTO_MATERIA          As Double
   Dim VALOR_PRECO_CUSTO            As Double

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
      
   chkTodos.Visible = False
   cmdMata.Visible = False

   If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
'1  OPERADOR
'2  VENDEDOR (a)
'3  ADMINISTRATIVO
'4  GERENTE
'5  DIRETOR
'6  FINANCEIRO
'7  CAIXA
      chkTodos.Visible = True
      cmdMata.Visible = True
   End If

   lstProduto.Visible = True

   CARREGA_COMB_FAMILIA_PRODUTO

   '1=PA;0=MP
   cmbTipoProd.Clear
   cmbTamanho.Clear
   cmbTamanhoAUX.Clear
   'cmbTipoProd.AddItem "1 - Produto Acabado"
   'cmbTipoProd.AddItem "0 - Matéria Prima"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'F'"
   SQL = SQL & " order by codigo desc "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbTipoProd.AddItem TabTemp.Fields("codigo").Value & " - " & TabTemp.Fields("DESCRICAO").Value

      TabTemp.MoveNext
   Wend
   cmbTipoProd.Text = "1" 'Default

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'N'"
   SQL = SQL & " order by codigo "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbTamanho.AddItem TabTemp.Fields("codigo").Value & " - " & TabTemp.Fields("DESCRICAO").Value
      cmbTamanhoAUX.AddItem TabTemp.Fields("codigo").Value

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'BUSCA ALIQUOTA ICMSss
   cmbALIQUOTA.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'J' "
   SQL = SQL & "order by codigo"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbALIQUOTA.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

   PreencheComboSituacaoTributaria
   PreencheComboNacionalidade

   cmbSituacao.Clear

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'P'"
   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabAUX.EOF
      cmbSituacao.AddItem Trim(TabAUX.Fields("DESCRICAO").Value)
      TabAUX.MoveNext
   Wend
   If TabAUX.State = 1 Then _
      TabAUX.Close

   cmbMarcaAUX.Clear
   cmbMarca.Clear

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'W' "
   SQL = SQL & "order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbMarca.AddItem Trim(TabDESCR!DESCRICAO)
      cmbMarcaAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         LIMPA_PECA
         txtProduto.SetFocus
      Case vbKeyF10
         GRAVA_REGISTRO
      Case vbKeyF3
         Msg = "Confirma Exclusão?"
         Style = vbYesNo + 32
         Title = "Atenção !!!"
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If Not RESPOSTA = vbYes Then _
            Exit Sub

         SQL = "delete from INVENTARIO "
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from PEDIDOITEM"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from NOTAENTRADAITEM "
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from CONTROLEPERDAITEM "
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from NFITEM"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from OSPECA"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from PEDIDOCOMPRAITEM"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from PEDIDOITEMOBS"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from PRODUTOFORNECEDOR"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from PRODUTOLOTE"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from TABELAPRECOITEM"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from TAXAMARKUP"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from TROCAPRODUTO"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from Estoque"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from COMISSAOITEM "
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "delete from PRODUTO "
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         LIMPA_PECA
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
   MOSTRA_RODAPE "", "", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub cmdCadPlaca_Click()
   frmCADASTROPARAMETRO.SSTab1.Tab = 5
   Call frmCADASTROPARAMETRO.SETA_GRID_GRUPO_PRODUTOS
   frmCADASTROPARAMETRO.Show 1
End Sub

Private Sub lstProdFornec_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(lstProdFornec.SelectedItem.Text) And Not IsNull(lstProdFornec.SelectedItem.ListSubItems.item(7).Text) Then
            If IsNumeric(lstProdFornec.SelectedItem.ListSubItems.item(7).Text) Then
               If Trim(lstProdFornec.SelectedItem.Text) <> "" Then
                  Msg = "Confirma exclusão ?"
                  PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     SQL = "delete PRODUTOFORNECEDOR "
                     SQL = SQL & " where produto_id = " & lstProdFornec.SelectedItem.Text
                     SQL = SQL & " and fornecedor_id = " & lstProdFornec.SelectedItem.ListSubItems.item(7).Text
                     CONECTA_RETAGUARDA.Execute SQL
                     SETA_GRID_PRODUTO_FORNEC
                  End If
               End If
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProdFornec_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imagem"
         If Trim(txtProduto.Text) <> "" Then _
            frmIMAGEM.Show 1
      Case "gravar"
         GRAVA_REGISTRO
      Case "limpar"
         LIMPA_PECA
         txtProduto.SetFocus
      Case "matar"
         MATA_PRODUTO
         txtProduto.SetFocus
      Case "voltar"
         Unload Me
      Case "print"
         MONTA_REL
      Case "consultar"
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then _
      txtProduto.SetFocus

   If SSTab1.Tab = 3 Then _
      txtDesc2.SetFocus

   If SSTab1.Tab = 2 Then
      lblProdFornec.Caption = "" & Trim(txtDesc.Text)
      SETA_GRID_PRODUTO_FORNEC
      txtCustoProdFornec.Text = txtPrecoCusto.Text
      txtCodgBarraFornec.Text = txtBarra.Text
      txtCodgFornec.Text = txtRef.Text
      txtCNPJCPF.SetFocus
   End If
End Sub

Private Sub cmbTamanho_GotFocus()
   cmbTamanho.SelStart = 0
   cmbTamanho.SelLength = Len(cmbTamanho)
   cmbTamanho.BackColor = &HC0FFFF
End Sub

Private Sub cmbALIQUOTA_LostFocus()
   cmbALIQUOTA.BackColor = &HFFFFFF
End Sub

Private Sub cmbFamilia_LostFocus()
   cmbFamilia.BackColor = &HFFFFFF
End Sub

Private Sub cmbMarca_LostFocus()
   cmbMarca.BackColor = &HFFFFFF
End Sub

Private Sub cmbOrigemMercadoria_LostFocus()
   cmbOrigemMercadoria.BackColor = &HFFFFFF
End Sub

Private Sub cmbSituacao_LostFocus()
   cmbSituacao.BackColor = &HFFFFFF
End Sub

Private Sub cmbST_LostFocus()
   cmbSt.BackColor = &HFFFFFF
End Sub

Private Sub cmbTamanho_LostFocus()
   cmbTamanho.BackColor = &HFFFFFF
End Sub

Private Sub cmbTipoProd_LostFocus()
   cmbTipoProd.BackColor = &HFFFFFF
End Sub

Private Sub cmdFornec_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   FORNEC_ID_N = 0
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A

      If TabFornecedor.State = 1 Then _
          TabFornecedor.Close

      SQL = "select descricao,razao, fornecedor_id from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & CNPJCPF_A & "'"
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFornecedor.EOF Then
         If Trim(TabFornecedor.Fields("descricao").Value) = "" Then
            txtFornec.Text = Trim(TabFornecedor!RAZAO)
            Else: txtFornec.Text = Trim(TabFornecedor!DESCRICAO)
         End If
         FORNEC_ID_N = 0 & TabFornecedor!FORNECEDOR_ID
      End If
      If TabFornecedor.State = 1 Then _
          TabFornecedor.Close
   End If
   CNPJCPF_A = ""
   txtFornec.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdFornec_Click"
End Sub

Private Sub cmdMata_Click()
'On Error Resume Next

   If TRAZ_TIPO_USUARIO = 5 Or TRAZ_TIPO_USUARIO = 4 Then
      Dim i As Integer

      If lstProduto.ListItems.Count > 0 Then
         For i = lstProduto.ListItems.Count To 1 Step -1
            If lstProduto.ListItems(i).Checked = True Then
               If Trim(lstProduto.ListItems(i).SubItems(15)) <> "" Then
                  Dim INDR_VENDIDO As Boolean
                  Dim INDR_ENTRADA As Boolean

                  INDR_VENDIDO = False
                  INDR_ENTRADA = False
'==========================verificando pedidoitem
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
   
                  SQL = "select produto_id from PEDIDOITEM WITH (NOLOCK)"
                  SQL = SQL & " where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     Msg = "Esse produto já foi vendido, deseja continuar essa operação de EXCLUSÃO ? "
                     PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
                     If RESPOSTA = vbYes Then _
                        INDR_VENDIDO = True
                     Else: INDR_VENDIDO = True 'não achou, esta liberado
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
'==========================verificando notaentradaitem
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
   
                  SQL = "select produto_id from NOTAENTRADAITEM WITH (NOLOCK)"
                  SQL = SQL & " where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     Msg = "Esse produto já foi realizado operação de entrada de mercadoria, deseja continuar essa operação de EXCLUSÃO ? "
                     PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
                     If RESPOSTA = vbYes Then _
                        INDR_ENTRADA = True
                     Else: INDR_ENTRADA = True 'não achou, esta liberado
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
'==================

                  If INDR_ENTRADA = True And INDR_VENDIDO = True Then
                     Msg = "Confirma exclusão produto ? " & Trim(lstProduto.ListItems(i).SubItems(1)) & " : " & Trim(lstProduto.ListItems(i).Text)
                     PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
                     If RESPOSTA = vbYes Then
                        SQL = "delete from ESTOQUE where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from PEDIDOITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from NOTAENTRADAITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from ENTRADAESTOQUEITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from ESTOQUETRANSF where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from NFITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from OSPECA where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from PEDIDOCOMPRAITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from POSICAOESTOQUE where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from PRODUTOEMPRESA where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from PRODUTOFORNECEDOR where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from PRODUTOLOTE where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from REGISTROPRODUCAOITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from TABELAPRECOITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from TAXAMARKUP where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from TROCAPRODUTO where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from CONTROLEPERDAITEM where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "delete from PRODUTO where produto_id = " & Trim(lstProduto.ListItems(i).SubItems(15))
                        CONECTA_RETAGUARDA.Execute SQL

                     End If
                  End If
               End If
            End If
         Next i
      End If
         CONSULTA_TUDO
      If INDR_PRI = True Then
         MsgBox "Processo realizado com sucesso."
         'LIMPA_TUDO
         CONSULTA_TUDO
      End If
   End If
End Sub

Private Sub chkTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstProduto.ListItems.Count > 0 Then
      For i = lstProduto.ListItems.Count To 1 Step -1
         If chkTodos.Value = 1 Then
            lstProduto.ListItems(i).Checked = True
            Else: lstProduto.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "chkTodos_Click"
End Sub

Private Sub lstProduto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstProduto, ColumnHeader
End Sub

Private Sub chkBarras_Click()
   txtProduto.SetFocus
End Sub

Private Sub lstProduto_DblClick()
'On Error GoTo ERRO_TRATA

   If Trim(lstProduto.SelectedItem.Text) <> "" Then
      txtProduto.Text = Trim(lstProduto.SelectedItem.Text)
      SSTab1.Tab = 0
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_DblClick"
End Sub

Private Sub lstProduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(lstProduto.SelectedItem.Text) Then
            Msg = "Confirma exclusão produto? ? "
            PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "delete from ESTOQUE where produto_id = " & Trim(lstProduto.SelectedItem.ListSubItems(15).Text)
               CONECTA_RETAGUARDA.Execute SQL

               SQL = "delete from PRODUTO where codg_produto = '" & Trim(lstProduto.SelectedItem.Text) & "'"
               CONECTA_RETAGUARDA.Execute SQL
               lstProduto.SelectedItem.ForeColor = vbRed
               lstProduto.SelectedItem.ListSubItems(1).ForeColor = vbRed
               DoEvents
               CONSULTA_TUDO
               lstProduto.Refresh
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstProduto_KeyDown"
End Sub

Private Sub cmbMarca_GotFocus()
   MOSTRA_RODAPE "Informe marca do Equipamento", "", "", "", ""
   cmbMarca.SelStart = 0
   cmbMarca.SelLength = Len(cmbMarca)
   cmbMarca.BackColor = &HC0FFFF
End Sub

Private Sub cmbmarca_Click()
'On Error GoTo ERRO_TRATA

   cmbMarcaAUX.ListIndex = cmbMarca.ListIndex
   txtLocacao.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmarca_Click"
End Sub

Private Sub cmbmarca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtLocacao.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmarca_KeyPress"
End Sub

Private Sub cmbALIQUOTA_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe ICMS do produto", "", "", ""
   cmbALIQUOTA.SelStart = 0
   cmbALIQUOTA.SelLength = Len(cmbALIQUOTA)
   cmbALIQUOTA.BackColor = &HC0FFFF
End Sub

Private Sub cmbFamilia_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Selecione grupo do produto", "", "", ""
   cmbFamilia.SelStart = 0
   cmbFamilia.SelLength = Len(cmbFamilia)
   cmbFamilia.BackColor = &HC0FFFF
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub cmbTAMANHO_Click()
'On Error GoTo ERRO_TRATA

   cmbTamanhoAUX.ListIndex = cmbTamanho.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTAMANHO_Click"
End Sub

Private Sub cmbTAMANHO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTipoProd.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPesoL_KeyPress"
End Sub

Private Sub cmbOrigemMercadoria_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a nacionalidade do produto", "", "", ""
   cmbOrigemMercadoria.SelStart = 0
   cmbOrigemMercadoria.SelLength = Len(cmbOrigemMercadoria)
   cmbOrigemMercadoria.BackColor = &HC0FFFF
End Sub

Private Sub cmbOrigemMercadoria_Click()
'On Error GoTo ERRO_TRATA

   cmbOrigemMercadoriaAUX.ListIndex = cmbOrigemMercadoria.ListIndex
   lblCST.Caption = "CST = " & Trim(cmbOrigemMercadoriaAUX.Text) & Trim(cmbSTAUX.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbOrigemMercadoria_Click"
End Sub

Private Sub cmbST_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a situação tributária do produto", "", "", ""
   cmbSt.SelStart = 0
   cmbSt.SelLength = Len(cmbSt)
   cmbSt.BackColor = &HC0FFFF
End Sub

Private Sub cmbST_Click()
On Error Resume Next

   cmbSTAUX.ListIndex = cmbSt.ListIndex
   lblCST.Caption = "CST = " & Trim(cmbOrigemMercadoriaAUX.Text) & Trim(cmbSTAUX.Text)
   lblST.Caption = "" & Trim(cmbSt.Text)

End Sub

Private Sub cmbSituacao_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a situação do produto", "", "", ""
   cmbSituacao.SelStart = 0
   cmbSituacao.SelLength = Len(cmbSituacao)
   cmbSituacao.BackColor = &HC0FFFF
End Sub

Private Sub cmbSituacao_Click()
   txtCodgNCM.SetFocus
End Sub

Private Sub cmbTipoProd_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Selecione tipo de produto", "", "", ""
   cmbTipoProd.SelStart = 0
   cmbTipoProd.SelLength = Len(cmbTipoProd)
   cmbTipoProd.BackColor = &HC0FFFF
End Sub

Private Sub cmbALIQUOTA_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtUN.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbALIQUOTA_KeyUp"
End Sub

Private Sub txtbarra_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe código de barras do produto", "", "", ""
   txtBarra.SelStart = 0
   txtBarra.SelLength = Len(txtBarra)
   txtBarra.BackColor = &HC0FFFF
End Sub

Private Sub txtBarra_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtRef.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbarra_KeyPress"
End Sub

Private Sub txtbarra_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtBarra.Text) <> "" Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select codg_produto,descricao from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_barra = '" & Trim(txtBarra.Text) & "'"
      'SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         If Not IsNull(TabProduto.Fields(0).Value) Then
            MsgBox "Codigo de barras já cadastrado para o produto : " & Trim(TabProduto.Fields(0).Value) & "-" & Trim(TabProduto.Fields(1).Value)
         End If
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   txtBarra.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbarra_LostFocus"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)
   txtCNPJCPF.BackColor = &HC0FFFF

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
       
      'If Trim(txtCNPJCPF.Text) <> "" Then
         txtCodgBarraFornec.SetFocus

   ElseIf KeyAscii = vbKeyDelete Then
      If Not IsNumeric(txtCNPJCPF.Text) Then
         txtCNPJCPF.Mask = "##############"
      End If
   ElseIf KeyAscii = vbKeyBack Then
      If Not IsNumeric(txtCNPJCPF.Text) Then
         txtCNPJCPF.Mask = "##############"
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If TabFornecedor.State = 1 Then _
          TabFornecedor.Close

      SQL = "select descricao,razao, fornecedor_id from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFornecedor.EOF Then
         If Trim(TabFornecedor.Fields("descricao").Value) = "" Then
            txtFornec.Text = Trim(TabFornecedor!RAZAO)
            Else: txtFornec.Text = Trim(TabFornecedor!DESCRICAO)
         End If
         FORNEC_ID_N = 0 & TabFornecedor!FORNECEDOR_ID
      End If
      If TabFornecedor.State = 1 Then _
          TabFornecedor.Close
   End If
   CNPJCPF_A = ""

   txtCNPJCPF.PromptInclude = False
   If Len(txtCNPJCPF.Text) > 0 Then
      If CInt(Len(txtCNPJCPF.Text)) = 11 Then
         Label31.Visible = False
         If Not ValidaCPF(txtCNPJCPF.Text) Then
            MsgBox "CPF com DV incorreto !!!"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF = ""
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      ElseIf CInt(Len(txtCNPJCPF.Text)) = 14 Then
         Label31.Visible = True
         If Not VALIDACNPJ(txtCNPJCPF.Text) Then
            MsgBox "CNPJ com DV incorreto !!! "
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "CNPJ/CPF com DV incorreto !!! "
         txtCNPJCPF = ""
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   ElseIf Len(txtCNPJCPF.Text) <> 0 Then
       MsgBox "CNPJ/CPF com DV incorreto !!! "
       txtCNPJCPF = ""
       TXTCNPJCPF_GotFocus
       txtCNPJCPF.SetFocus
       Exit Sub
   End If
   
   txtCNPJCPF.PromptInclude = False
   CRITERIO_A = txtCNPJCPF.Text
   txtCNPJCPF.PromptInclude = False
   
   If txtCNPJCPF.Text <> "" Then
      CRITERIO_A = txtCNPJCPF.Text

      If Not IsNull(txtCNPJCPF.Text) Then
          If Len(txtCNPJCPF.Text) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(txtCNPJCPF.Text) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.Text = CRITERIO_A
   End If

   txtCNPJCPF.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCodgBarraFornec_GotFocus()
   txtCodgBarraFornec.SelStart = 0
   txtCodgBarraFornec.SelLength = Len(txtCodgBarraFornec)
   txtCodgBarraFornec.BackColor = &HC0FFFF
End Sub

Private Sub txtCodgBarraFornec_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCustoProdFornec.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgBarraFornec_KeyPress"
End Sub

Private Sub txtCodgBarraFornec_LostFocus()
   txtCodgBarraFornec.BackColor = &HFFFFFF
End Sub

Private Sub txtCodgFornec_GotFocus()
   txtCodgFornec.SelStart = 0
   txtCodgFornec.SelLength = Len(txtCodgFornec)
   txtCodgFornec.BackColor = &HC0FFFF
End Sub

Private Sub txtCodgFornec_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If PRODUTO_ID_N > 0 And FORNEC_ID_N > 0 Then _
         GRAVA_PRODUTO_FORNEC

      SETA_GRID_PRODUTO_FORNEC

      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgFornec_KeyPress"
End Sub

Private Sub txtCodgFornec_LostFocus()
   txtCodgFornec.BackColor = &HFFFFFF
End Sub

Private Sub txtCodgNCM_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe NCM do produto", "", "", ""
   txtCodgNCM.SelStart = 0
   txtCodgNCM.SelLength = Len(txtCodgNCM)
   txtCodgNCM.BackColor = &HC0FFFF
End Sub

Private Sub txtCodgNCM_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF1
         CRITERIO_A = ""
         frmNCMConsulta.Show 1
         If Trim(CRITERIO_A) <> "" Then _
            txtCodgNCM.Text = CRITERIO_A
         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgNCM_KeyDown"
End Sub

Private Sub txtCodgNCM_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBarra.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgNCM_KeyPress"
End Sub

Private Sub txtCodgNCM_LostFocus()
   txtCodgNCM.BackColor = &HFFFFFF
   If Trim(txtCodgNCM.Text) = "" Then _
      txtCodgNCM.Text = "00"

End Sub

Private Sub txtCustoProdFornec_GotFocus()
   txtCustoProdFornec.SelStart = 0
   txtCustoProdFornec.SelLength = Len(txtCustoProdFornec)
   txtCustoProdFornec.BackColor = &HC0FFFF
End Sub

Private Sub txtCustoProdFornec_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCodgFornec.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgBarraFornec_KeyPress"
End Sub

Private Sub txtCustoProdFornec_LostFocus()
   txtCustoProdFornec.BackColor = &HFFFFFF
   If Trim(txtCustoProdFornec.Text) = "" Then _
      txtCustoProdFornec.Text = 0
   txtCustoProdFornec.Text = "" & Format(txtCustoProdFornec.Text, strFormatacao2Digitos)
End Sub

Private Sub txtDesc_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a descrição do produto", "", "", ""
   txtDesc.SelStart = 0
   txtDesc.SelLength = Len(txtDesc)
   txtDesc.BackColor = &HC0FFFF
End Sub

Private Sub cmdCons_Click()
   CONSULTA_TUDO
End Sub

Private Sub txtDesc2_GotFocus()
   txtDesc2.SelStart = 0
   txtDesc2.SelLength = Len(txtDesc2)
   txtDesc2.BackColor = &HC0FFFF
End Sub

Private Sub txtDesc2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      CONSULTA_TUDO
   End If
End Sub

Private Sub txtDesc2_LostFocus()
   txtDesc2.BackColor = &HFFFFFF
End Sub

Private Sub txtDtUltEntrada_GotFocus()
   txtDtUltEntrada.SelStart = 0
   txtDtUltEntrada.SelLength = Len(txtDtUltEntrada)
   txtDtUltEntrada.BackColor = &HC0FFFF
End Sub

Private Sub txtDtUltEntrada_LostFocus()
   txtDtUltEntrada.BackColor = &HFFFFFF
End Sub

Private Sub txtEmbalagem_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe embalagem do produto", "", "", ""
   txtEmbalagem.SelStart = 0
   txtEmbalagem.SelLength = Len(txtEmbalagem)
   txtEmbalagem.BackColor = &HC0FFFF
End Sub

Private Sub txtEmbalagem_LostFocus()
   txtEmbalagem.BackColor = &HFFFFFF
End Sub

Private Sub txtEstoqueMaximo_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a quantidade máxima em estoque do produto", "", "", ""
   txtEstoqueMaximo.SelStart = 0
   txtEstoqueMaximo.SelLength = Len(txtEstoqueMaximo)
   txtEstoqueMaximo.BackColor = &HC0FFFF
End Sub

Private Sub txtEstoqueMaximo_LostFocus()
   txtEstoqueMaximo.BackColor = &HFFFFFF
End Sub

Private Sub txtEstoqueMinimo_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a quantidade minima em estoque do produto", "", "", ""
   txtEstoqueMinimo.SelStart = 0
   txtEstoqueMinimo.SelLength = Len(txtEstoqueMinimo)
   txtEstoqueMinimo.BackColor = &HC0FFFF
End Sub

Private Sub txtEstoqueMinimo_LostFocus()
   txtEstoqueMinimo.BackColor = &HFFFFFF
End Sub

Private Sub txtFORNEC_GotFocus()
   txtCodgFornec.SetFocus
End Sub

Private Sub txtlocacao_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a locação do produto", "", "", ""
   txtLocacao.SelStart = 0
   txtLocacao.SelLength = Len(txtLocacao)
   txtLocacao.BackColor = &HC0FFFF
End Sub

Private Sub txtLocacao_LostFocus()
   txtLocacao.BackColor = &HFFFFFF
End Sub

Private Sub txtPerc_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe % comissão de venda do produto", "", "", ""
   txtPerc.SelStart = 0
   txtPerc.SelLength = Len(txtPerc)
   txtPerc.BackColor = &HC0FFFF
End Sub

Private Sub txtPercIVA_LostFocus()
   txtPercIVA.BackColor = &HFFFFFF
End Sub

Private Sub txtPrecoAtacado_LostFocus()
   txtPrecoAtacado.BackColor = &HFFFFFF
   txtPrecoAtacado.Text = "" & Format(txtPrecoAtacado.Text, strFormatacao3Digitos)
End Sub

Private Sub txtPrecoVenda_LostFocus()
   txtPrecoVenda.BackColor = &HFFFFFF
   txtPrecoVenda.Text = "" & Format(txtPrecoVenda.Text, strFormatacao3Digitos)
End Sub

Private Sub txtref_LostFocus()
   txtRef.BackColor = &HFFFFFF
End Sub

Private Sub txtUN_LostFocus()
   txtUN.BackColor = &HFFFFFF
End Sub

Private Sub txtPerc_LostFocus()
   txtPerc.BackColor = &HFFFFFF
End Sub

Private Sub txtPercIVA_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe IVA do produto", "", "", ""
   txtPercIVA.SelStart = 0
   txtPercIVA.SelLength = Len(txtPercIVA)
   txtPercIVA.BackColor = &HC0FFFF
End Sub

Private Sub txtPercIVA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPrecoCusto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercIVA_KeyPress"
End Sub

Private Sub txtDesc_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDesc.Text = UCase(txtDesc.Text)
   txtDesc.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_LostFocus"
End Sub

Private Sub txtlocacao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPrecoCusto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtlocacao_KeyPress"
End Sub

Private Sub cmbTipoProd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEmbalagem.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtPrecoAtacado_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe preço de atacado do produto", "", "", ""
   txtPrecoAtacado.SelStart = 0
   txtPrecoAtacado.SelLength = Len(txtPrecoAtacado)
   txtPrecoAtacado.BackColor = &HC0FFFF
End Sub

Private Sub txtPrecoVenda_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe preço de venda do produto", "", "", ""
   txtPrecoVenda.SelStart = 0
   txtPrecoVenda.SelLength = Len(txtPrecoVenda)
   txtPrecoVenda.BackColor = &HC0FFFF
End Sub

Private Sub txtRef_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a referência do produto", "", "", ""
   txtRef.SelStart = 0
   txtRef.SelLength = Len(txtRef)
   txtRef.BackColor = &HC0FFFF
End Sub

Private Sub txtPesoL_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe peso Liquido do produto", "", "", ""
   txtPesoL.SelStart = 0
   txtPesoL.SelLength = Len(txtPesoL)
   txtPesoL.BackColor = &HC0FFFF
End Sub

Private Sub txtPesoL_LostFocus()
   txtPesoL.Text = "" & Format(txtPesoL.Text, strFormatacao3Digitos)
   txtPesoL.BackColor = &HFFFFFF
End Sub

Private Sub txtPesoB_LostFocus()
   txtPesoB.Text = "" & Format(txtPesoB.Text, strFormatacao3Digitos)
   txtPesoB.BackColor = &HFFFFFF
End Sub

Private Sub txtPesoB_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe peso Bruto do produto", "", "", ""
   txtPesoB.SelStart = 0
   txtPesoB.SelLength = Len(txtPesoB)
   txtPesoB.BackColor = &HC0FFFF
End Sub

Private Sub txtPesoL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPesoB.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPesoL_KeyPress"
End Sub

Private Sub txtPesoB_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTamanho.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPesoL_KeyPress"
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtLocacao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRef_KeyPress"
End Sub

Private Sub txtUN_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe a unidade de medida do produto", "", "", ""
   txtUN.SelStart = 0
   txtUN.SelLength = Len(txtUN)
   txtUN.BackColor = &HC0FFFF
End Sub

Private Sub txtun_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      cmbSt.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtun_KeyUp"
End Sub

Private Sub cmbST_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtPerc.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbST_KeyUp"
End Sub

Private Sub txtperc_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtEstoqueMinimo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtperc_KeyUp"
End Sub

Private Sub txtestoqueminimo_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtEstoqueMaximo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtestoqueminimo_KeyUp"
End Sub

Private Sub txtestoquemaximo_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      cmbFamilia.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtestoquemaximo_KeyUp"
End Sub

Private Sub txtprecovenda_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtDesc.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecovenda_KeyUp"
End Sub

Private Sub txtDesc_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_KeyUp"
End Sub

Private Sub cmbICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbICMS_KeyPress"
End Sub

Private Sub cmbFamilia_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEstoqueMaximo.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_KeyPress"
End Sub

Private Sub txtestoquemaximo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEstoqueMinimo.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtestoquemaximo_KeyPress"
End Sub

Private Sub txtestoqueminimo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPerc.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtestoqueminimo_KeyPress"
End Sub
Private Sub txtprecoatacado_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPesoL.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecoatacado_KeyPress"
End Sub

Private Sub txtcustoanterior_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEmbalagem.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecoatacado_KeyPress"
End Sub

Private Sub txtembalagem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      GRAVA_REGISTRO
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtembalagem_KeyPress"
End Sub


Private Sub txtPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUN.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPerc_KeyPress"
End Sub

Private Sub cmbSituacao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCodgNCM.SetFocus
      Else: KeyAscii = 0
      'Else
      '   If KeyAscii = 8 Or KeyAscii = 44 Then
      '      Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
      '   End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbSituacao_KeyPress"
End Sub


Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "F6 - Excluir", "F7 - Consultar Produtos", "F10 - GRAVAR ", ""

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPRODUTO_GotFocus"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            Msg = "Confirma Exclusão?"
            Style = vbYesNo + 32
            Title = "Atenção !!!"
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If txtProduto.Text <> "" Then
               If RESPOSTA = vbYes Then
                  SQL = "delete from PRODUTO "
                  SQL = SQL & " where produto_id = " & TabConsulta.Fields("produto_id").Value
                  CONECTA_RETAGUARDA.Execute SQL
                  LIMPA_PECA
               End If
            End If
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         txtProduto.SetFocus
      Case vbKeyF7
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then _
            txtProduto.Text = SQL3

         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPRODUTO_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtProduto.Text) = "" Then
         LIMPA_PECA_QUASE_TUDO

         GERA_CODIGO_PRODUTO
         'Opcao de Zerar Estoque no caso de aproveitar Codigo!
         ZERA_ESTOQUE
         
         txtProduto.Text = NUMR_PROD_N
      End If

      If (LE_PRODUTO(Trim(txtProduto.Text), "CADASTRO")) = False Then
         txtProduto.Enabled = True
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)
         Exit Sub
      End If

      PROCESSA_DADOS_PRODUTOS
      txtDesc.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPRODUTO_KeyPress"
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbFamilia.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_KeyPress"
End Sub

Private Sub cmbST_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbALIQUOTA.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbST_KeyPress"
End Sub

Private Sub txtun_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSituacao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtun_KeyPress"
End Sub

Private Sub cmbaliquota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbOrigemMercadoria.SetFocus
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbaliquota_KeyPress"
End Sub

Private Sub cmbOrigemMercadoria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSituacao.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtprecovenda_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPrecoAtacado.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecovenda_KeyPress"
End Sub
'=========================SUBROTINAS
Private Sub GRAVA_REGISTRO()
'On Error GoTo ERRO_TRATA

   Dim PRECO_CUSTO_ANTERIOR         As Double
   Dim PRECO_CUSTO_ANTERIOR_MATERIA As Double
   Dim PRECO_CUSTO_MATERIA          As Double
   Dim VALOR_PRECO_CUSTO            As Double
   Dim PESO_LIQUIDO_N               As Double
   Dim PESO_BRUTO_N                 As Double
   Dim INDR_TIPO_PROD               As Integer

   PRECO_CUSTO_ANTERIOR = 0
   PRECO_CUSTO_ANTERIOR_MATERIA = 0
   PRECO_CUSTO_MATERIA = 0
   PESO_LIQUIDO_N = 0
   PESO_BRUTO_N = 0
   INDR_TIPO_PROD = 1

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Código do produto deve ser informado."
      txtProduto.SetFocus
      Exit Sub
   End If
   If Trim(txtDesc.Text) = "" Then
      MsgBox "Descrição do produto deve ser informado."
      txtDesc.SetFocus
      Exit Sub
   End If
   If Trim(txtPrecoVenda.Text) = "" Then
      MsgBox "Preço de venda do produto deve ser informado."
      txtPrecoVenda.SetFocus
      Exit Sub
   End If
   If Trim(txtProduto.Text) = "" Then
      MsgBox "Código do produto deve ser informado."
      txtProduto.SetFocus
      Exit Sub
   End If

   If Trim(cmbOrigemMercadoriaAUX.Text) = "" Then _
      cmbOrigemMercadoriaAUX.Text = 0

   If Trim(txtPrecoAtacado.Text) = "" Then _
      txtPrecoAtacado.Text = 0

   If Trim(txtEmbalagem.Text) = "" Then _
      txtEmbalagem.Text = 0

   If Trim(txtQTDE.Text) = "" Then _
      txtQTDE.Text = 0

   If Trim(txtEstoqueMaximo.Text) = "" Then _
      txtEstoqueMaximo.Text = 0

   If Trim(txtEstoqueMinimo.Text) = "" Then _
      txtEstoqueMinimo.Text = 0

   If PRECO_CUSTO_ANTERIOR_MATERIA > 0 Then _
      PRECO_CUSTO_ANTERIOR_MATERIA = 0

   If PRECO_CUSTO_MATERIA > 0 Then _
      PRECO_CUSTO_MATERIA = 0

   If Trim(txtPerc.Text) = "" Then _
      txtPerc.Text = 0

   If Trim(txtLocacao.Text) = "" Then _
      txtLocacao.Text = 0

   If Trim(txtPesoL.Text) = "" Then _
      txtPesoL.Text = 0

   If Trim(txtPesoB.Text) = "" Then _
      txtPesoB.Text = 0

   If Trim(cmbTipoProd.Text) <> "" Then _
      INDR_TIPO_PROD = Left(cmbTipoProd.Text, 1)
   
   If Trim(txtCodgNCM.Text) = "" Then
      MsgBox "Favor Cadastrar Codigo NCM da Mercadoria!"
      txtCodgNCM.SetFocus
      Exit Sub
   End If

   If Trim(cmbALIQUOTA.Text) = "" Then _
      cmbALIQUOTA.Text = 0

   If Trim(cmbTamanho.Text) = "" Then _
      cmbTamanhoAUX.Text = 0

   If Trim(cmbFamiliaAUX.Text) = "" Then _
      cmbFamiliaAUX.Text = 0

   If Trim(txtUN.Text) = "" Then _
      txtUN.Text = "UN"

   If Trim(cmbSTAUX.Text) = "" Then _
      cmbSTAUX.Text = "00"

   If Trim(cmbSituacao.Text) = "" Then _
      cmbSituacao.Text = "A"

   If Trim(txtPercIVA.Text) = "" Then _
      txtPercIVA.Text = 0

   MARCA_ID_N = 0 & cmbMarcaAUX.Text

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select Tipo_Prod,preco_custo,produto_id from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Trim(txtPrecoCusto.Text) <> "" Then
         If Not TabProduto!Tipo_Prod = 1 Then 'Produto acabado
            'Quardar Dados Para salvar Dados anteriores materia prima
            PRECO_CUSTO_ANTERIOR = txtPrecoCusto.Text 'Passando para variavel

            If TabProduto!PRECO_CUSTO <> PRECO_CUSTO_ANTERIOR Then _
               PRECO_CUSTO_ANTERIOR_MATERIA = TabProduto!PRECO_CUSTO

            PRECO_CUSTO_MATERIA = txtPrecoCusto.Text
         End If
      End If

      NUMR_ID_N = TabProduto!PRODUTO_ID

      SQL = "UPDATE PRODUTO SET "
         SQL = SQL & " codg_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & ", Codg_Barra = '" & txtBarra.Text & "'"
         SQL = SQL & ", Descricao = '" & txtDesc.Text & "'"
         SQL = SQL & ", SITUACAO = '" & Left(cmbSituacao.Text, 1) & "'"
         SQL = SQL & ", familiaproduto_id = " & cmbFamiliaAUX.Text
         SQL = SQL & ", Preco_venda = " & tpMOEDA(txtPrecoVenda.Text)

         SQL = SQL & ", perc_compoe_venda = " & tpMOEDA(txtPercVenda.Text)

         SQL = SQL & ", unidade_medida = '" & txtUN.Text & "'"
         SQL = SQL & ", codg_ncm = '" & Trim(txtCodgNCM.Text) & "'"
         SQL = SQL & ", Aliquota_Icms = " & cmbALIQUOTA.Text
         SQL = SQL & ", TAMANHO = " & cmbTamanhoAUX.Text
         SQL = SQL & ", PERC_DESCONTO = " & 0
         SQL = SQL & ", Situacao_Tributaria = '" & cmbSTAUX.Text & "'"
         SQL = SQL & ", Tipo_Prod = " & INDR_TIPO_PROD
         SQL = SQL & ", PATH_IMAGEM = '" & Trim(LOCAL_IMAGEM) & "'"
         SQL = SQL & ", ORIGEM_MERCADO = " & cmbOrigemMercadoriaAUX.Text
         SQL = SQL & ", LOCACAO = '" & txtLocacao.Text & "'"
         SQL = SQL & ", EMBALAGEM = " & tpMOEDA(txtEmbalagem.Text)
         SQL = SQL & ", QTD_MINIMO = " & tpMOEDA(txtEstoqueMinimo.Text)
         SQL = SQL & ", QTD_MAXIMO = " & tpMOEDA(txtEstoqueMaximo.Text)
         SQL = SQL & ", PRECO_ATACADO = " & tpMOEDA(txtPrecoAtacado.Text)
         SQL = SQL & ", PRECO_CUSTO = " & tpMOEDA(txtPrecoCusto.Text)
         SQL = SQL & ", Preco_Custo_anterior = " & tpMOEDA(PRECO_CUSTO_ANTERIOR)
         SQL = SQL & ", fornecedor_id = 0" & FORNEC_ID_N
         SQL = SQL & ", perciva = 0" & txtPercIVA.Text
         SQL = SQL & ", referencia = '" & Trim(txtRef.Text) & "'"
         SQL = SQL & ", peso_liquido = " & tpMOEDA(txtPesoL.Text)
         SQL = SQL & ", peso_bruto = " & tpMOEDA(txtPesoB.Text)
         SQL = SQL & ", MARCA_ID = " & MARCA_ID_N                          'MARCA_ID
         
         If chkBalanca.Value = 0 Then
            SQL = SQL & ", produto_balanca = 'FALSE'"
            Else: SQL = SQL & ", produto_balanca = 'TRUE'"
         End If
         If chkDesconto.Value = 0 Then
            SQL = SQL & ", PERMITE_DESCONTO = 'FALSE'"
            Else: SQL = SQL & ", PERMITE_DESCONTO = 'TRUE'"
         End If
         If chkConceder.Value = 0 Then
            SQL = SQL & ", CONCEDER_PRODUCAO = 'FALSE'"
            Else: SQL = SQL & ", CONCEDER_PRODUCAO = 'TRUE'"
         End If

      SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
      Else
         NUMR_ID_N = MAX_ID("produto_id", "produto", "", "", "", "")

         SQL = "insert into PRODUTO "
         SQL = SQL & "("
            SQL = SQL & "EMPRESA_ID,PRODUTO_ID,CODG_PRODUTO,DESCRICAO,FAMILIAPRODUTO_ID,"
            SQL = SQL & "UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,SITUACAO_TRIBUTARIA,"
            SQL = SQL & "ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,REFERENCIA,CODG_NCM,"
            SQL = SQL & "COMP_TRIBUTARIA,fornecedor_id,PRECO_CUSTO_ANTERIOR,qtd_ped_anterior,"
            SQL = SQL & "PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda,PERCIVA,DT_CADASTRO,PERC_COMIS,PATH_IMAGEM,ORIGEM_MERCADO,"
            SQL = SQL & "LOCACAO,PRECO_VAREJO_ANTERIOR,PRECO_ATACADO_ANTERIOR,EMBALAGEM,USUARIO_ID,"
            SQL = SQL & "QTD_MINIMO,QTD_MAXIMO,DT_ULT_VENDA,DT_ULT_COMPRA,PESO_LIQUIDO,PESO_BRUTO,"
            SQL = SQL & " TAMANHO, marca_id,produto_balanca, PERMITE_DESCONTO,CONCEDER_PRODUCAO,perc_compoe_venda "
         SQL = SQL & ")"
         SQL = SQL & " values ("
            SQL = SQL & EMPRESA_ID_N                                 'EMPRESA_ID
            SQL = SQL & ",0" & NUMR_ID_N                             'PRODUTO_ID
            SQL = SQL & ",'" & Trim(txtProduto.Text) & "'"            'CODG_PRODUTO
            SQL = SQL & ",'" & Trim(txtDesc.Text) & "'"              'DESCRICAO
            SQL = SQL & ",0" & cmbFamiliaAUX.Text                      'FAMILIAPRODUTO_ID
            SQL = SQL & ",'" & Trim(txtUN.Text) & "'"                'UNIDADE_MEDIDA
            SQL = SQL & ",'" & Trim(txtBarra.Text) & "'"             'CODG_BARRA
            SQL = SQL & ",'" & Left(cmbSituacao.Text, 1) & "'"         'SITUACAO
            SQL = SQL & ",'" & cmbSTAUX.Text & "'"            'SITUACAO_TRIBUTARIA
            SQL = SQL & ",0" & cmbALIQUOTA.Text                      'ALIQUOTA_ICMS_NORMAL_DENTRO_UF
            SQL = SQL & ",0"                                         'PERC_DESCONTO
            SQL = SQL & ",0" & Left(cmbTipoProd.Text, 1)             'TIPO_PROD
            SQL = SQL & ",'" & Trim(txtRef.Text) & "'"               'REFERENCIA
            SQL = SQL & ",'" & Trim(txtCodgNCM.Text) & "'"           'CODG_NCM
            SQL = SQL & ",0"                                         'COMP_TRIBUTARIA
            SQL = SQL & ",0" & FORNEC_ID_N                           'fornecedor_id
            SQL = SQL & "," & tpMOEDA(PRECO_CUSTO_ANTERIOR)          'PRECO_CUSTO_ANTERIOR
            SQL = SQL & ",0"                                         'qtd_ped_anterior
            SQL = SQL & "," & tpMOEDA(txtPrecoCusto.Text)            'PRECO_CUSTO
            SQL = SQL & "," & tpMOEDA(txtPrecoAtacado.Text)          'PRECO_ATACADO
            SQL = SQL & "," & tpMOEDA(txtPrecoVenda.Text)            'PRECO_Venda
            SQL = SQL & ",0" & Replace(txtPercIVA.Text, ",", ".")                     'PERCIVA
            SQL = SQL & ",'" & Now & "'"                       'DT_CADASTRO
            SQL = SQL & "," & tpMOEDA(txtPerc.Text)                  'PERC_COMIS
            SQL = SQL & ",'" & Trim(LOCAL_IMAGEM) & "'"         'PATH_IMAGEM
            SQL = SQL & ",0" & cmbOrigemMercadoriaAUX.Text        'ORIGEM_MERCADO
            SQL = SQL & ",'" & Trim(txtLocacao.Text) & "'"           'LOCACAO
            SQL = SQL & ",0"                                         'PRECO_VAREJO_ANTERIOR
            SQL = SQL & ",0"                                         'PRECO_ATACADO_ANTERIOR
            SQL = SQL & ",0" & txtEmbalagem.Text                     'EMBALAGEM
            SQL = SQL & ",0" & USUARIO_ID_N                            'USUARIO_ID
            SQL = SQL & "," & tpMOEDA(txtEstoqueMinimo.Text)         'QTD_MINIMO
            SQL = SQL & "," & tpMOEDA(txtEstoqueMaximo.Text)         'QTD_MAXIMO
            SQL = SQL & ",'" & DMA(0) & "'"                          'DT_ULT_VENDA
            SQL = SQL & ",'" & DMA(0) & "'"                          'DT_ULT_COMPRA
            SQL = SQL & "," & tpMOEDA(txtPesoL.Text)                 'peso_liquido
            SQL = SQL & "," & tpMOEDA(txtPesoB.Text)                 'peso_bruto
            SQL = SQL & ",0" & cmbTamanhoAUX.Text                    'TAMANHO
            SQL = SQL & "," & MARCA_ID_N                             'MARCA_ID

            If chkBalanca.Value = 0 Then
               SQL = SQL & ",'false'"
               Else: SQL = SQL & ",'true'"
            End If

            If chkDesconto.Value = 0 Then
               SQL = SQL & ",'false'"
               Else: SQL = SQL & ",'true'"
            End If

            If chkConceder.Value = 0 Then
               SQL = SQL & ",'false'"
               Else: SQL = SQL & ",'true'"
            End If

            SQL = SQL & "," & tpMOEDA(txtPercVenda.Text)

         SQL = SQL & ")"
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close
   CONECTA_RETAGUARDA.Execute SQL

   RODA_AT_ESTOQUE NUMR_ID_N, ESTABELECIMENTO_ID_N

   LIMPA_PECA

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_REGISTRO"
   Err.Clear
End Sub

Private Sub LIMPA_PECA()
'On Error GoTo ERRO_TRATA

   PreencheComboNacionalidade
   PRODUTO_ID_N = 0
   chkDesconto.Value = 0
   chkBalanca.Value = 0
   chkConceder.Value = 0
   txtDtUltEntrada.Text = ""
   MARCA_ID_N = 0
   cmbMarca.Text = ""
   cmbMarcaAUX.Text = ""
   chkBarras.Value = 0
   txtPesoL.Text = 0
   txtPesoB.Text = 0
   FORNEC_ID_N = 0
   txtCodgNCM.Text = ""
   txtBarra.Text = ""
   LOCAL_IMAGEM = ""
   cmbTipoProd.Text = ""
   txtCNPJCPF.Text = ""
   txtFornec.Text = ""
   txtRef.Text = ""
   cmbSt.Text = ""
   cmbSTAUX.Text = ""
   txtUN.Text = ""
   lstProduto.ListItems.Clear
   txtProduto.Text = ""
   txtDesc.Text = ""
   txtPrecoVenda.Text = ""
   txtPercVenda.Text = "0"
   txtPrecoCusto.Text = ""
   txtPrecoAtacado.Text = ""
   txtEmbalagem.Text = ""
   txtQTDE.Text = 0
   cmbFamilia.Text = ""
   cmbFamiliaAUX.Text = ""
   txtEstoqueMaximo.Text = ""
   txtEstoqueMinimo.Text = ""
   txtCustoAnterior.Text = ""
   txtPerc.Text = ""
   txtLocacao = ""
   cmbSituacao.Text = ""
   cmbALIQUOTA.Text = ""
   cmbTamanho.Text = ""
   cmbTamanhoAUX.Text = ""
   lstProduto.Visible = True
   txtPercIVA.Text = ""
   lstProdFornec.ListItems.Clear
   SSTab1.Tab = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_PECA"
End Sub

Private Sub LIMPA_PECA_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   PreencheComboNacionalidade

   PRODUTO_ID_N = 0
   chkDesconto.Value = 0
   chkBalanca.Value = 0
   chkConceder.Value = 0
   txtDtUltEntrada.Text = ""
   MARCA_ID_N = 0
   txtCustoAnterior.Text = ""
   txtBarra.Text = ""
   txtDtUltEntrada.Text = ""
   LOCAL_IMAGEM = ""
   cmbTipoProd.Text = ""
   txtCNPJCPF.Text = ""
   txtFornec.Text = ""
   txtRef.Text = ""
   cmbSt.Text = ""
   cmbSTAUX.Text = ""
   txtUN.Text = ""
   lstProduto.ListItems.Clear
   txtDesc.Text = ""
   txtPrecoCusto.Text = ""
   txtPrecoAtacado.Text = ""
   txtEmbalagem.Text = ""
   txtQTDE.Text = 0
   cmbFamilia.Text = ""
   cmbFamiliaAUX.Text = ""
   txtEstoqueMaximo.Text = ""
   txtEstoqueMinimo.Text = ""
   txtPerc.Text = ""
   txtLocacao.Text = ""
   cmbSituacao.Text = ""
   cmbALIQUOTA.Text = ""
   cmbTamanho.Text = ""
   cmbTamanhoAUX.Text = ""
   txtDtUltEntrada.Text = ""
   lstProduto.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_PECA_QUASE_TUDO"
End Sub

Private Sub ZERA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   txtQTDE.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ZERA_ESTOQUE"
End Sub

Public Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Sub CARREGA_COMB_FAMILIA_PRODUTO()
'On Error GoTo ERRO_TRATA

   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & "order by descricao "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFamilia.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("codg_familia").Value)
      cmbFamiliaAUX.AddItem Trim(TabDESCR.Fields("FAMILIAPRODUTO_ID").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMB_FAMILIA_PRODUTO"
End Sub

Private Sub PreencheComboSituacaoTributaria()
'On Error GoTo ERRO_TRATA

   cmbSt.Clear
   cmbSTAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select codigo, descricao from CST WITH (NOLOCK)"
   SQL = SQL & " order by codigo"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      DoEvents
      Do Until TabDESCR.EOF
         cmbSt.AddItem Trim(TabDESCR!Codigo) & "-" & Trim(TabDESCR!DESCRICAO)
         cmbSTAUX.AddItem Trim(TabDESCR!Codigo)
         TabDESCR.MoveNext
      Loop
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PreencheComboSituacaoTributaria"
End Sub

Private Sub PreencheComboNacionalidade()
'On Error GoTo ERRO_TRATA

   cmbOrigemMercadoria.Clear
   cmbOrigemMercadoriaAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'Q'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbOrigemMercadoria.AddItem TabTemp.Fields("codigo").Value & " - " & TabTemp.Fields("DESCRICAO").Value
      cmbOrigemMercadoriaAUX.AddItem TabTemp.Fields("codigo").Value

      TabTemp.MoveNext
   Wend
   cmbOrigemMercadoriaAUX.Text = "0"         'Default
   cmbOrigemMercadoria.Text = "0 - Nacional" 'Default

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PreencheComboNacionalidade"
End Sub

Sub CONSULTA_TUDO()
   CONT_N = 0
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select count(produto_id) from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_produto is not null"

   If Trim(txtDesc2.Text) <> "" Then _
      SQL = SQL & " and descricao like '" & UCase(Trim(txtDesc2.Text)) & "%" & "'"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then _
      If Not IsNull(TabProduto.Fields(0).Value) Then _
         CONT_N = 0 & TabProduto.Fields(0).Value
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If CONT_N > 500 Then
      Msg = "Esta operação irá processar todos produtos cadastrado, deseja continuar ? " & CONT_N & " registros"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbNo Then _
         Exit Sub
   End If

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where produto_id is not null"
   If Trim(txtDesc2.Text) <> "" Then _
      SQL = SQL & " and descricao like '" & UCase(Trim(txtDesc2.Text)) & "%" & "'"

   SQL = SQL & " order by descricao"
   
   SETA_GRID
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim dblContador   As Double
   Dim VALOR_CUSTO_N As Double

   lstProduto.Visible = False
   lstProduto.ListItems.Clear
   dblContador = 0

   If TabProduto.State = 1 Then _
      TabProduto.Close

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      MsgBox "Não encontrei nenhum produto com o CRITERIO_A de procura especificado", vbExclamation
      txtDesc2.SetFocus
      Exit Sub
      Else: TabProduto.MoveFirst
   End If
   Me.Enabled = True
   While Not TabProduto.EOF
      DoEvents
      dblContador = dblContador + 1

      Me.Caption = "Aguarde, Processando ...  " & dblContador

      Set item = lstProduto.ListItems.Add(, "seq." & dblContador, Trim(TabProduto.Fields("codg_produto").Value))
      
      item.SubItems(1) = "" & Trim(TabProduto!DESCRICAO)
      item.SubItems(2) = "" & Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value), strFormatacao3Digitos)

      item.SubItems(4) = "" & Format(0, strFormatacao3Digitos)
      item.SubItems(4) = "-"

      If CONECTA_AUXILIAR.State = 1 Then _
         item.SubItems(4) = "" & Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("PRODUTO_ID").Value), strFormatacao3Digitos)

      item.SubItems(5) = "" & Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
      item.SubItems(6) = "" & Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)

      VALOR_CUSTO_N = 0 & TabProduto!PRECO_CUSTO

      item.SubItems(7) = "" & Format(0, strFormatacao2Digitos)
      If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
         item.SubItems(7) = "" & Format(VALOR_CUSTO_N, strFormatacao2Digitos)

      If Not IsNull(TabProduto.Fields("produto_id").Value) Then
         If TabProduto.Fields("produto_id").Value > 0 Then
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

            If Not IsNull(TabProduto.Fields("fornecedor_id").Value) Then
               SQL = "select Descricao from vwFornecedor WITH (NOLOCK)"
               SQL = SQL & " where fornecedor_id = " & TabProduto.Fields("fornecedor_id").Value
               TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabFornecedor.EOF Then _
                  item.SubItems(8) = "" & Trim(TabFornecedor.Fields(0).Value)
            End If
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close
         End If
      End If

      item.SubItems(9) = "" & Format(TabProduto!Qtd_minimo, strFormatacao3Digitos)
      item.SubItems(10) = "" & Format(TabProduto!qtd_maximo, strFormatacao3Digitos)
      item.SubItems(11) = "0"
      item.SubItems(12) = "" & Trim(TabProduto!REFERENCIA)
      item.SubItems(14) = "" & TabProduto.Fields("situacao_tributaria").Value
      item.SubItems(15) = "" & TabProduto.Fields("produto_id").Value

      NUMR_ID_N = 0 & TabProduto!FAMILIAPRODUTO_ID

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where familiaproduto_id = " & NUMR_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp!DESCRICAO) Then
            item.SubItems(13) = TabTemp!DESCRICAO
            Else: item.SubItems(13) = "SEM GRUPO"
         End If
         Else: item.SubItems(13) = "SEM GRUPO"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If TabProduto.Fields("situacao").Value = "A" Then
         item.ForeColor = vbBlue
         item.ListSubItems(1).ForeColor = vbBlue
         item.ListSubItems(2).ForeColor = vbBlue
         item.ListSubItems(3).ForeColor = vbBlue
         item.ListSubItems(4).ForeColor = vbBlue
         item.ListSubItems(5).ForeColor = vbBlue
         item.ListSubItems(6).ForeColor = vbBlue
         item.ListSubItems(7).ForeColor = vbBlue
         item.ListSubItems(8).ForeColor = vbBlue
         item.ListSubItems(9).ForeColor = vbBlue
         item.ListSubItems(10).ForeColor = vbBlue
         item.ListSubItems(11).ForeColor = vbBlue
         item.ListSubItems(12).ForeColor = vbBlue
         item.ListSubItems(13).ForeColor = vbBlue
      End If
      If TabProduto.Fields("situacao").Value = "P" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(3).ForeColor = vbRed
         item.ListSubItems(4).ForeColor = vbRed
         item.ListSubItems(5).ForeColor = vbRed
         item.ListSubItems(6).ForeColor = vbRed
         item.ListSubItems(7).ForeColor = vbRed
         item.ListSubItems(8).ForeColor = vbRed
         item.ListSubItems(9).ForeColor = vbRed
         item.ListSubItems(10).ForeColor = vbRed
         item.ListSubItems(11).ForeColor = vbRed
         item.ListSubItems(12).ForeColor = vbRed
         item.ListSubItems(13).ForeColor = vbRed
      End If

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   Me.Enabled = True
   lstProduto.Visible = True

   If CONECTA_AUXILIAR.State = 1 Then _
      CONECTA_AUXILIAR.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub MATA_PRODUTO()
'On Error GoTo ERRO_TRATA

   INDR_PRI = True
   If Trim(txtProduto.Text) <> "" Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         PRODUTO_ID_N = TabProduto.Fields(0).Value

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select produto_id from NOTAENTRADAITEM WITH (NOLOCK)"
         SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            INDR_PRI = False
            MsgBox "Exclusão não permitida, item com movimento de entrada."
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select produto_id from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         SQL = SQL & " and tipo_reg = 'PC' "
         SQL = SQL & " and pedidoitem.status <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            INDR_PRI = False
            MsgBox "Exclusão não permitida, item já vendido."
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         If INDR_PRI = True Then
            If PRODUTO_ID_N > 0 Then
               Msg = "Confirma Exclusão?"
               Style = vbYesNo + 32
               Title = "Atenção !!!"
               Help = "DEMO.HLP"
               Ctxt = 1000
               RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
               If txtProduto.Text <> "" Then
                  If RESPOSTA = vbYes Then
                     SQL = "delete  from ESTOQUE "
                     SQL = SQL & " where produto_id = " & PRODUTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     SQL = "delete  from PRODUTO "
                     SQL = SQL & " where produto_id = " & PRODUTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
                     LIMPA_PECA
                  End If
               End If
            End If
         End If
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_PRODUTO"
End Sub

Sub MONTA_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{PRODUTO.empresa_id} = " & EMPRESA_ID_N
   FORMULA_REL = FORMULA_REL & " and {ESTOQUE.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = FORMULA_REL & " and {PRODUTO.produto_ID} = {ESTOQUE.produto_ID}"

   If txtProduto.Text <> "" Then
      SqL2 = Chr$(39) & txtProduto.Text & "%" & Chr(39)
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.codg_produto} = '" & Trim(txtProduto.Text) & "'"
      Else
         If txtDesc.Text <> "" Then
            SqL2 = Chr$(39) & txtDesc.Text & "%" & Chr(39)
            FORMULA_REL = FORMULA_REL & " and {PRODUTO.descricao} = " & SqL2
         End If
   End If

   If cmbFamiliaAUX.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.familiaproduto_id} = " & cmbFamiliaAUX.Text

   If cmbSituacao.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.situacao} = '" & Trim(Left(cmbSituacao.Text, 1)) & "'"

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Msg = "Agrupar por codigo de produto? ? "
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000

   If RESPOSTA = vbYes Then
      Nome_Relatorio = "rel_Estoque2.rpt"
      Else: Nome_Relatorio = "rel_Estoque.rpt"
   End If

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL"
End Sub

Sub GRAVA_PRODUTO_FORNEC()
'On Error GoTo ERRO_TRATA

   Dim TabProdFornec As New ADODB.Recordset

   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

   SQL = "select * from PRODUTOFORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProdFornec.EOF Then
      SQL = "insert into PRODUTOFORNECEDOR "
         SQL = SQL & "(PRODUTO_ID,FORNECEDOR_ID,CODG_PROD_FORNEC,PRECO_CUSTO,CODG_BARRA)"
      SQL = SQL & " values("
         SQL = SQL & PRODUTO_ID_N
         SQL = SQL & "," & FORNEC_ID_N
         SQL = SQL & ",'" & Trim(txtCodgFornec.Text) & "'"
         SQL = SQL & "," & tpMOEDA(txtCustoProdFornec.Text)
         SQL = SQL & ",'" & Trim(txtCodgBarraFornec.Text) & "'"
      SQL = SQL & ")"
      Else
         SQL = "update PRODUTOFORNECEDOR set "

            SQL = SQL & " preco_custo = " & tpMOEDA(txtCustoProdFornec.Text)
            SQL = SQL & " ,codg_barra = '" & Trim(txtCodgBarraFornec.Text) & "'"
            SQL = SQL & " ,codg_prod_fornec = '" & Trim(txtCodgFornec.Text) & "'"

         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   End If
   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

   CONECTA_RETAGUARDA.Execute SQL

   txtCNPJCPF.Text = ""
   txtFornec.Text = ""
   txtCodgBarraFornec.Text = ""
   txtCustoProdFornec.Text = ""
   txtCodgFornec.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO_FORNEC"
End Sub

Sub SETA_GRID_PRODUTO_FORNEC()
'On Error GoTo ERRO_TRATA

   Dim TabProdFornec As New ADODB.Recordset
   CONT_N = 0
   lstProdFornec.ListItems.Clear

   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

   SQL = "select PRODUTO.PRODUTO_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
   SQL = SQL & " PRODUTOFORNECEDOR.FORNECEDOR_ID, PRODUTOFORNECEDOR.CODG_PROD_FORNEC, "
   SQL = SQL & " PRODUTOFORNECEDOR.PRECO_CUSTO, PRODUTOFORNECEDOR.CODG_BARRA, "
   SQL = SQL & " FORNECEDOR.PESSOA_ID, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFornec "
   SQL = SQL & " from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTOFORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOFORNECEDOR.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTOFORNECEDOR.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where PRODUTOFORNECEDOR.produto_id = " & PRODUTO_ID_N

   TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProdFornec.EOF
      Set item = lstProdFornec.ListItems.Add(, "seq." & CONT_N, Trim(TabProdFornec.Fields("produto_id").Value))
      item.SubItems(1) = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
      item.SubItems(2) = "" & Trim(TabProdFornec.Fields("descricao").Value)
      item.SubItems(3) = "" & Trim(TabProdFornec.Fields("CODG_PROD_FORNEC").Value)
      item.SubItems(4) = "" & Trim(TabProdFornec.Fields("CODG_BARRA").Value)
      item.SubItems(5) = "" & Format(TabProdFornec.Fields("PRECO_CUSTO").Value, strFormatacao2Digitos)
      item.SubItems(6) = "" & Trim(TabProdFornec.Fields("NomeFornec").Value)
      item.SubItems(7) = "" & Trim(TabProdFornec.Fields("fornecedor_id").Value)
      CONT_N = CONT_N + 1
      TabProdFornec.MoveNext
   Wend
   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PRODUTO_FORNEC"
End Sub

Private Sub txtprecocusto_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtPrecoVenda.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecocusto_KeyUp"
End Sub

Private Sub txtPrecoCusto_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe preço de custo do produto", "", "", ""
   txtPrecoCusto.SelStart = 0
   txtPrecoCusto.SelLength = Len(txtPrecoCusto)
   txtPrecoCusto.BackColor = &HC0FFFF
End Sub

Private Sub txtprecocusto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercVenda.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecocusto_KeyPress"
End Sub

Private Sub txtPrecoCusto_LostFocus()
   txtPrecoCusto.BackColor = &HFFFFFF
   txtPrecoCusto.Text = "" & Format(txtPrecoCusto.Text, strFormatacao3Digitos)
   CALCULA_PRECO_VENDA
End Sub

Private Sub txtPercVenda_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtPrecoVenda.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercVenda_KeyUp"
End Sub

Private Sub txtPercVenda_GotFocus()
   MOSTRA_TOP "ESC - SAIR", "Informe preço de custo do produto", "", "", ""
   txtPercVenda.SelStart = 0
   txtPercVenda.SelLength = Len(txtPercVenda)
   txtPercVenda.BackColor = &HC0FFFF
End Sub

Private Sub txtPercVenda_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPrecoVenda.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercVenda_KeyPress"
End Sub

Private Sub txtPercVenda_LostFocus()
   txtPercVenda.BackColor = &HFFFFFF
   txtPercVenda.Text = "" & Format(txtPercVenda.Text, strFormatacao2Digitos)
   CALCULA_PRECO_VENDA
End Sub

Private Sub cmdCadFor_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   PESSOA_ID_N = 0
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaCadastro.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCadFor_Click"
End Sub

Sub CALCULA_PRECO_VENDA()
'On Error GoTo ERRO_TRATA

   If Trim(txtPercVenda.Text) <> "" And Trim(txtPrecoCusto.Text) <> "" Then
      If IsNumeric(txtPercVenda.Text) And IsNumeric(txtPrecoCusto.Text) Then
         Dim Perc_Compoe_Venda_N As Double
         Dim VALOR_CUSTO_N       As Double
         Perc_Compoe_Venda_N = 0 & txtPercVenda.Text
         VALOR_CUSTO_N = 0 & txtPrecoCusto.Text

         If Perc_Compoe_Venda_N > 0 And VALOR_CUSTO_N > 0 Then
            'Msg = "Deseja que o sistema calcule preço de venda baseado no percentual cadastrado na familia de produto ?"
            'PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            'If RESPOSTA = vbYes Then
               txtPrecoVenda.Text = "" & Format((VALOR_CUSTO_N * Perc_Compoe_Venda_N / 100) + VALOR_CUSTO_N, strFormatacao2Digitos)
               txtPrecoAtacado.Text = "" & Format((VALOR_CUSTO_N * Perc_Compoe_Venda_N / 100) + VALOR_CUSTO_N, strFormatacao2Digitos)
            'End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CALCULA_PRECO_VENDA"
End Sub

Sub GERA_CODIGO_PRODUTO()
'On Error GoTo ERRO_TRATA

NUMR_SEQ_N = 1

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select CODG_PROD_RESERVA from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         NUMR_SEQ_N = 0 & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_codg_prod = " & NUMR_SEQ_N

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

RODA_PRODUTO:

   NUMR_PROD_N = 1

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select seq_codg_prod from EMPRESA WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         If IsNumeric(TabTemp.Fields(0).Value) Then _
            NUMR_PROD_N = 1 + TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_codg_prod = " & NUMR_PROD_N

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "select codg_produto from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_produto = '" & NUMR_PROD_N & "'"
   'SQL = SQL & " and situacao <> 'C' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      GoTo RODA_PRODUTO

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & NUMR_PROD_N
   'SQL = SQL & " and situacao <> 'C' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      GoTo RODA_PRODUTO

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_CODIGO_PRODUTO"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   'MOSTRA_DADOS_PRODUTO

   MOSTRA_PRODUTO

   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub

Private Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If PRODUTO_ID_N <= 0 Then _
      PRODUTO_ID_N = MAX_ID("produto_id", "produto", "", "", "", "")

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select PRODUTO.*, FAMILIAPRODUTO.CODG_FAMILIA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia, "
   SQL = SQL & " FAMILIAPRODUTO.UNIDADE_MEDIDA AS UN_FAMILIA, FAMILIAPRODUTO.DESC_UNIDADE_MEDIDA, "
   SQL = SQL & " FAMILIAPRODUTO.PRODUCAO, FAMILIAPRODUTO.PERC_COMPOE_VENDA AS F_COMPOE_VENDA"
   SQL = SQL & " from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabProduto.EOF Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select PRODUTO.*, FAMILIAPRODUTO.CODG_FAMILIA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia, "
      SQL = SQL & " FAMILIAPRODUTO.UNIDADE_MEDIDA AS UN_FAMILIA, FAMILIAPRODUTO.DESC_UNIDADE_MEDIDA, "
      SQL = SQL & " FAMILIAPRODUTO.PRODUCAO, FAMILIAPRODUTO.PERC_COMPOE_VENDA AS F_COMPOE_VENDA"
      SQL = SQL & " from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"
      SQL = SQL & " where codg_produto = '" & CODG_PRODUTO_A & "'"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProduto.EOF Then
         If TabProduto.State = 1 Then _
            TabProduto.Close
         Exit Sub
      End If
   End If

   FORNEC_ID_N = 0 & TabProduto.Fields("fornecedor_id").Value
   txtPercVenda.Text = Format(0, strFormatacao2Digitos)
   txtPrecoCusto.Text = Format(TabProduto.Fields("preco_custo").Value, strFormatacao2Digitos)

   If Not IsNull(TabProduto.Fields("fornecedor_id").Value) Then
      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      SQL = "select Descricao,cnpjcpf from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where fornecedor_id = " & TabProduto.Fields("fornecedor_id").Value
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFornecedor.EOF Then
         txtFornec.Text = Trim(TabFornecedor.Fields("descricao").Value)
         txtCNPJCPF.Text = Trim(TabFornecedor.Fields("cnpjcpf").Value)
      End If
      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close
   End If

   If Not IsNull(TabProduto!ORIGEM_MERCADO) Then
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from DESCR WITH (NOLOCK)"
      SQL = SQL & " where TIPO = 'Q'"
      SQL = SQL & " and codigo = '" & Trim(TabProduto!ORIGEM_MERCADO) & "'"
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then
         cmbOrigemMercadoria.Text = TabAUX.Fields("codigo").Value & " - " & TabAUX.Fields("DESCRICAO").Value
         cmbOrigemMercadoriaAUX.Text = TabAUX.Fields("codigo").Value
      End If
      If TabAUX.State = 1 Then _
         TabAUX.Close
   End If

   If Not IsNull(TabProduto.Fields("CONCEDER_PRODUCAO").Value) Then
      If TabProduto.Fields("CONCEDER_PRODUCAO").Value = False Then
         chkConceder.Value = 0
         Else: chkConceder.Value = 1
      End If
   End If

   If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then
      If TabProduto.Fields("produto_balanca").Value = False Then
         chkBalanca.Value = 0
         Else: chkBalanca.Value = 1
      End If
   End If
   If Not IsNull(TabProduto.Fields("PERMITE_DESCONTO").Value) Then
      If TabProduto.Fields("PERMITE_DESCONTO").Value = False Then
         chkDesconto.Value = 0
         Else: chkDesconto.Value = 1
      End If
   End If
   If Not IsNull(TabProduto!DT_ULT_COMPRA) Then _
      txtDtUltEntrada.Text = TabProduto!DT_ULT_COMPRA

   If Not IsNull(TabProduto!Tipo_Prod) Then
      If TabProduto!Tipo_Prod = 1 Then _
         cmbTipoProd.Text = "1" & "-" & "Produto Acabado"
      If TabProduto!Tipo_Prod = 0 Then _
         cmbTipoProd.Text = "0" & "-" & "Matéria Prima"
   End If

   If Not IsNull(TabProduto!path_imagem) Then _
      LOCAL_IMAGEM = "" & Trim(TabProduto!path_imagem)

   If Not IsNull(TabProduto!REFERENCIA) Then _
      txtRef.Text = TabProduto!REFERENCIA

   If Not IsNull(TabProduto!SITUACAO_TRIBUTARIA) Then
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select codigo, descricao from CST WITH (NOLOCK)"
      SQL = SQL & " where codigo = '" & Trim(TabProduto!SITUACAO_TRIBUTARIA) & "'"
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then
         cmbSt.Text = Trim(TabDESCR!Codigo) & "-" & Trim(TabDESCR!DESCRICAO)
         cmbSTAUX.Text = Trim(TabDESCR!Codigo)
      End If
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

   If Not IsNull(TabProduto!UNIDADE_MEDIDA) Then _
      txtUN.Text = TabProduto!UNIDADE_MEDIDA

   If Not IsNull(TabProduto!perc_comis) Then _
      txtPerc.Text = TabProduto!perc_comis

   If Not IsNull(TabProduto!DESCRICAO) Then _
      txtDesc.Text = Trim(TabProduto!DESCRICAO)

   If Not IsNull(TabProduto!FAMILIAPRODUTO_ID) Then
      If IsNumeric(TabProduto!FAMILIAPRODUTO_ID) Then
         cmbFamilia.Text = Trim(TabProduto.Fields("DescFamilia").Value) & "-" & Trim(TabProduto.Fields("familiaproduto_id").Value)
         cmbFamiliaAUX.Text = Trim(TabProduto.Fields("familiaproduto_id").Value)
         txtPercVenda.Text = "" & TabProduto.Fields("perc_compoe_venda").Value
      End If
   End If

   If Not IsNull(TabProduto.Fields("perc_compoe_venda").Value) Then
      If TabProduto.Fields("perc_compoe_venda").Value > 0 Then
         txtPercVenda.Text = "" & Format(TabProduto.Fields("perc_compoe_venda").Value, strFormatacao2Digitos)
      End If
   End If

   txtPrecoVenda.Text = "" & Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
   txtEmbalagem.Text = "" & Format(TabProduto!embalagem, strFormatacao2Digitos)
   txtPesoL.Text = "" & Format(TabProduto.Fields("peso_liquido").Value, strFormatacao2Digitos)
   txtPesoB.Text = "" & Format(TabProduto.Fields("peso_bruto").Value, strFormatacao2Digitos)
   txtCustoAnterior.Text = "" & Format(TabProduto!PRECO_CUSTO_ANTERIOR, strFormatacao2Digitos)

   If Len(txtPrecoAtacado.Text) = 0 Then
      txtPrecoAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      Else: txtPrecoAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
   End If

   If Not IsNull(TabProduto!qtd_maximo) Then _
      txtEstoqueMaximo.Text = TabProduto!qtd_maximo
   If Not IsNull(TabProduto!Qtd_minimo) Then _
      txtEstoqueMinimo.Text = TabProduto!Qtd_minimo

   If Not IsNull(TabProduto!perc_comis) Then _
      txtPerc.Text = TabProduto!perc_comis
   
   If Not IsNull(TabProduto!LOCACAO) Then _
      txtLocacao.Text = TabProduto!LOCACAO

   If Not IsNull(TabProduto.Fields("situacao").Value) Then
      If TabProduto.Fields("situacao").Value = "A" Then _
         cmbSituacao.Text = "ATIVO"
      If TabProduto.Fields("situacao").Value = "C" Then _
         cmbSituacao.Text = "CANCELADO"
      If TabProduto.Fields("situacao").Value = "R" Then _
         cmbSituacao.Text = "REATIVADO"
      If TabProduto.Fields("situacao").Value = "V" Then _
         cmbSituacao.Text = "VENCIDO"
      If TabProduto.Fields("situacao").Value = "P" Then _
         cmbSituacao.Text = "PROMOÇAO"
   End If

   If Not IsNull(TabProduto!Aliquota_Icms) Then _
      cmbALIQUOTA.Text = TabProduto!Aliquota_Icms

   If Not IsNull(TabProduto!tamanho) Then
      cmbTamanhoAUX.Text = TabProduto!tamanho
      cmbTamanho.Text = TabProduto!tamanho
   End If

   If Not IsNull(TabProduto!CODG_NCM) Then _
      txtCodgNCM.Text = TabProduto!CODG_NCM
   If Not IsNull(TabProduto!Codg_Barra) Then _
      txtBarra.Text = TabProduto!Codg_Barra
   If Not IsNull(TabProduto!PERCIVA) Then _
      txtPercIVA.Text = TabProduto!PERCIVA

   If Not IsNull(TabProduto.Fields("marca_id").Value) Then
      If IsNumeric(TabProduto.Fields("marca_id").Value) Then
         cmbMarcaAUX.Text = "" & TabProduto.Fields("marca_id").Value
         cmbMarca.Text = "" & TRAZ_DESCRITOR("W", cmbMarcaAUX.Text)
      End If
   End If

   txtProduto.Text = TabProduto.Fields("codg_produto").Value
   INDR_ACHOU = True

'===================ESTOQUE
   txtQTDE.Text = "" & Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value), strFormatacao3Digitos)
'==========================

   lblCST.Caption = "CST = " & Trim(cmbOrigemMercadoriaAUX.Text) & Trim(cmbSTAUX.Text)

   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub
