VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCOMPRAS 
   Caption         =   "Departamento de Compras"
   ClientHeight    =   8490
   ClientLeft      =   1485
   ClientTop       =   1920
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "COMPRAS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   12240
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pedido de Compras"
      TabPicture(0)   =   "COMPRAS.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "stBarpedido"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "listaprod"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Definir"
      TabPicture(1)   =   "COMPRAS.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame5 
         Caption         =   "Dados Complementares"
         ForeColor       =   &H00400000&
         Height          =   3255
         Left            =   -74880
         TabIndex        =   35
         Top             =   4200
         Width           =   11295
         Begin VB.TextBox txtcusto 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9720
            TabIndex        =   58
            Top             =   900
            Width           =   1455
         End
         Begin VB.TextBox txtvarejo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   57
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox txtatacado 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   56
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox txtprecodigitado 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9720
            TabIndex        =   55
            Top             =   400
            Width           =   1455
         End
         Begin VB.Frame Frame7 
            Caption         =   "Penúltimas Compras:"
            ForeColor       =   &H00400000&
            Height          =   855
            Left            =   0
            TabIndex        =   50
            Top             =   2400
            Width           =   11295
            Begin VB.CommandButton cmdanterior 
               Caption         =   "&Anterior"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   9600
               TabIndex        =   64
               Top             =   290
               Width           =   735
            End
            Begin VB.CommandButton cmdproximo 
               Caption         =   "&Próximo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10440
               TabIndex        =   63
               Top             =   290
               Width           =   735
            End
            Begin VB.TextBox txtvlrcomprapen 
               Enabled         =   0   'False
               Height          =   375
               Left            =   4920
               TabIndex        =   61
               Top             =   350
               Width           =   1335
            End
            Begin VB.TextBox txtqtdpen 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1680
               TabIndex        =   60
               Top             =   350
               Width           =   1335
            End
            Begin MSMask.MaskEdBox txtdtpen 
               Height          =   375
               Left            =   7920
               TabIndex        =   52
               ToolTipText     =   "Data do lote"
               Top             =   345
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label18 
               Caption         =   "Data Compra:"
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   6720
               TabIndex        =   54
               Top             =   405
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Valor Compra:"
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   3600
               TabIndex        =   53
               Top             =   405
               Width           =   1215
            End
            Begin VB.Label Label16 
               Caption         =   "Quantidade:"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   480
               TabIndex        =   51
               Top             =   405
               Width           =   1095
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000013&
            Caption         =   "Últimas Compras"
            ForeColor       =   &H00400000&
            Height          =   855
            Left            =   0
            TabIndex        =   45
            Top             =   1440
            Width           =   11295
            Begin VB.TextBox txtqtdult 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   62
               Top             =   320
               Width           =   1335
            End
            Begin VB.TextBox txtvlrcomprault 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4920
               TabIndex        =   59
               Top             =   320
               Width           =   1335
            End
            Begin MSMask.MaskEdBox txtdatault 
               Height          =   375
               Left            =   7920
               TabIndex        =   49
               ToolTipText     =   "Data do lote"
               Top             =   320
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label15 
               Caption         =   "Data Compra:"
               ForeColor       =   &H8000000D&
               Height          =   375
               Left            =   6720
               TabIndex        =   48
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label8 
               Caption         =   "Valor Compra:"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   3600
               TabIndex        =   47
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000004&
               Caption         =   "Quantidade:"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   480
               TabIndex        =   46
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.TextBox txtreservado 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   40
            Top             =   400
            Width           =   1335
         End
         Begin VB.TextBox txtestoque 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   38
            Top             =   400
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Preço Custo:"
            Height          =   375
            Left            =   8640
            TabIndex        =   44
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Preço Varejo:"
            Height          =   375
            Left            =   5640
            TabIndex        =   43
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Preco Atacado:"
            Height          =   375
            Left            =   1200
            TabIndex        =   42
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Preço Digitado:"
            Height          =   375
            Left            =   8400
            TabIndex        =   41
            Top             =   460
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Quantidade Estoque Reservado:"
            Height          =   255
            Left            =   4080
            TabIndex        =   39
            Top             =   460
            Width           =   2655
         End
         Begin VB.Label Label9 
            Caption         =   "Quantidade Estoque Atual:"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   460
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Histórico dos Últimos Seis Meses de Movimentação do Estoque do Produto Selecionado"
         ForeColor       =   &H00400000&
         Height          =   2295
         Left            =   -71280
         TabIndex        =   31
         Top             =   1920
         Width           =   7695
         Begin MSComctlLib.ListView listahist_ent 
            Height          =   975
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   1720
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483639
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Ent. Mes"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Ult. Compra"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Entrada Ultimos 6 Meses."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView listahist_sai 
            Height          =   975
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   1720
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483639
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Saida. Mes"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Ult. Venda"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Saida  Ultimos 6 Meses."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estoque das Filiais"
         ForeColor       =   &H00400000&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   30
         Top             =   1920
         Width           =   3495
         Begin MSComctlLib.ListView listafiliais 
            Height          =   1935
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   3413
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Filial"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Estoque "
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Reservado"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Consulta por Código Produto e Descrição do Produto ou pelo Fornecedor"
         ForeColor       =   &H00400000&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   11295
         Begin VB.TextBox txtprodutoconsulta 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1440
            MaxLength       =   12
            TabIndex        =   28
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   860
            Width           =   1935
         End
         Begin VB.TextBox txtnomeconsulta 
            DataField       =   "Nome"
            Height          =   345
            Left            =   3600
            MaxLength       =   80
            TabIndex        =   27
            Top             =   860
            Width           =   7095
         End
         Begin VB.TextBox txtdescricaoconsulta 
            DataField       =   "Nome"
            Height          =   345
            Left            =   3600
            MaxLength       =   80
            TabIndex        =   25
            Top             =   400
            Width           =   7095
         End
         Begin MSMask.MaskEdBox txtcgccpfconsulta 
            Height          =   345
            Left            =   1440
            TabIndex        =   26
            Top             =   400
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
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
         Begin VB.Label Label1 
            Caption         =   "Produto:"
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   900
            Width           =   735
         End
         Begin VB.Label lblfornec 
            Caption         =   "Fornecedor:"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pedido"
         ForeColor       =   &H00400000&
         Height          =   2415
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   12015
         Begin VB.CommandButton cmdConsProd 
            BackColor       =   &H00FFFFFF&
            Height          =   350
            Left            =   4080
            Picture         =   "COMPRAS.frx":5C4A
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   1440
            Width           =   405
         End
         Begin VB.TextBox txt_qtd_anterior 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   67
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1880
            Width           =   1095
         End
         Begin VB.TextBox txt_custo_anterior 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   10080
            MaxLength       =   6
            TabIndex        =   65
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1880
            Width           =   1335
         End
         Begin VB.TextBox txtPreco 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   6720
            MaxLength       =   6
            TabIndex        =   8
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1880
            Width           =   1335
         End
         Begin VB.TextBox txtpedido 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2115
            MaxLength       =   6
            TabIndex        =   1
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtNome 
            DataField       =   "Nome"
            Height          =   345
            Left            =   4200
            MaxLength       =   80
            TabIndex        =   14
            Top             =   640
            Width           =   7095
         End
         Begin VB.TextBox txtPrazo 
            Height          =   330
            Left            =   6120
            TabIndex        =   4
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtFinanceiro 
            Height          =   330
            Left            =   10200
            TabIndex        =   5
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtseq 
            Height          =   330
            Left            =   4200
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtDescricao 
            DataField       =   "Nome"
            Height          =   360
            Left            =   4680
            MaxLength       =   80
            TabIndex        =   12
            Top             =   1490
            Width           =   6615
         End
         Begin VB.TextBox txtquantidade 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   7
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1880
            Width           =   1095
         End
         Begin VB.TextBox txtproduto 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   2115
            MaxLength       =   15
            TabIndex        =   6
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1490
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txtCgcCpf 
            Height          =   345
            Left            =   2115
            TabIndex        =   2
            Top             =   645
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
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
         Begin MSMask.MaskEdBox txtdatapedido 
            Height          =   375
            Left            =   2115
            TabIndex        =   3
            ToolTipText     =   "Data do lote"
            Top             =   1035
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qtd. Anterior:"
            Height          =   240
            Left            =   2640
            TabIndex        =   68
            Top             =   1905
            Width           =   1170
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Preco Custo Anterior:"
            Height          =   240
            Left            =   8160
            TabIndex        =   66
            Top             =   1905
            Width           =   1875
         End
         Begin VB.Label lblpreco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Preço Mercadoria:"
            Height          =   240
            Left            =   5115
            TabIndex        =   22
            Top             =   1905
            Width           =   1590
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pedido de Compras:"
            Height          =   240
            Left            =   285
            TabIndex        =   21
            Top             =   255
            Width           =   1770
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
            Height          =   240
            Left            =   1020
            TabIndex        =   20
            Top             =   675
            Width           =   1035
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data Pedido:"
            Height          =   240
            Left            =   930
            TabIndex        =   19
            Top             =   1095
            Width           =   1125
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
            Height          =   240
            Left            =   5520
            TabIndex        =   18
            Top             =   1095
            Width           =   570
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Financeiro:"
            Height          =   240
            Left            =   9000
            TabIndex        =   17
            Top             =   1095
            Width           =   960
         End
         Begin VB.Label lblproduto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   240
            Left            =   1320
            TabIndex        =   16
            Top             =   1500
            Width           =   735
         End
         Begin VB.Label lblquantidade 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   240
            Left            =   195
            TabIndex        =   15
            Top             =   1905
            Width           =   1050
         End
      End
      Begin MSComctlLib.ListView listaprod 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7011
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
            Text            =   "Seq."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Produto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descricao"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qtd."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Qtd P. Anterior"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Valor Custo Atual"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Valor Total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Custo Anterior"
            Object.Width           =   1764
         EndProperty
      End
      Begin ComctlLib.StatusBar stBarpedido 
         Height          =   375
         Left            =   8400
         TabIndex        =   36
         Top             =   6840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   2
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   2646
               MinWidth        =   2646
               Text            =   "Valor Total:"
               TextSave        =   "Valor Total:"
               Key             =   "disponivel"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   1270
      ButtonWidth     =   3069
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "exclui"
            Object.ToolTipText     =   "Exclui itens"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Rel. Pedido"
            Key             =   "print"
            Object.ToolTipText     =   "Imprime Contagem"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consulta"
            Object.ToolTipText     =   "Consultar Produto"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "Gravar"
            Object.ToolTipText     =   "Gravar Pedido"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "clear"
            Object.ToolTipText     =   "Limpa a Tela"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10440
         TabIndex        =   69
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9600
         Top             =   120
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
               Picture         =   "COMPRAS.frx":664C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":7A74
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":9171
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":AFE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":C0F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":D35B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":E3EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":F9B9
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":10B53
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":11D85
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10920
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   23
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":13021
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":1369B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":13D15
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":1438F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":14529
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":14E03
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "COMPRAS.frx":156DD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCOMPRAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Indr_Lote_Existe As Boolean

   Dim qtd_mes_02 As Integer, qtd_mes_03 As Integer, qtd_mes_04 As Integer, qtd_mes_05 As Integer, qtd_mes_06 As Integer
   Dim qtd_mes_07 As Integer, qtd_mes_08 As Integer, qtd_mes_09 As Integer, qtd_mes_10 As Integer, qtd_mes_11 As Integer
   Dim qtd_mes_12 As Integer, qtd_mes As Integer

Private Sub cmdanterior_Click()
'On Error GoTo ERRO_TRATA

   If txtcgccpfconsulta.Text <> "" Then
      CONSULTA_ANTERIOR_REGISTRO
      Else:
          If txtnomeconsulta.Text <> "" Then
             CONSULTA_ANTERIOR_REGISTRO
             Else:
                 MsgBox "Selecione Fornecedor Ou o Produto Desejado Para Consulta"
                 Exit Sub
                 txtcgccpfconsulta.SetFocus
          End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdanterior_Click"
End Sub

Private Sub cmdproximo_Click()
'On Error GoTo ERRO_TRATA

   If txtcgccpfconsulta.Text <> "" Then
      CONSULTA_PROXIMO_REGISTRO
      Else:
          If txtnomeconsulta.Text <> "" Then
             CONSULTA_PROXIMO_REGISTRO
             Else:
                 MsgBox "Selecione Fornecedor Ou o Produto Desejado Para Consulta"
                 Exit Sub
                 txtcgccpfconsulta.SetFocus
          End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdproximo_Click"
End Sub

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmCOMPRAS)
   Me.Caption = Me.Caption & " - " & Me.Name
   txtdatapedido = Format(Date, "dd/mm/yyyy")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
      Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub listaprod_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView listaprod, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

    Select Case Button.key
      Case "sair"
         Unload Me
      Case "print"
         If txtPedido.Text <> "" Then
            NUMR_COMPRA_N = txtPedido.Text
            FORMULA_REL = "{PEDIDOCOMPRA.PEDIDOCOMPRA_id} = " & NUMR_COMPRA_N
            Else
               FORMULA_REL = InputBox(SQL3, "Informe número do Pedido a ser impressa ")
               If IsNumeric(FORMULA_REL) Then
                  NUMR_COMPRA_N = FORMULA_REL
               End If
               If NUMR_COMPRA_N = 0 Then
                  Exit Sub
               End If
               FORMULA_REL = "{PEDIDOCOMPRA.PEDIDOCOMPRA_id} = " & NUMR_COMPRA_N
         End If

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Pedido_Compras.rpt"
         frmRELATORIO10.Show 1
      Case "salvar"
         If txtPedido.Text = "" Then
            MsgBox "Precisa Gerar um pedido para salvar!", vbExclamation, ""
            Exit Sub
         End If
         INDR_GRAVA = False
         NUMR_COMPRA_N = txtPedido.Text
         txtPedido.SetFocus
         GRAVA_CABECA
         LIMPA_BODY
         LIMPA_TUDO
      Case "clear"
         LIMPA_TUDO
      Case "exclui"
         EXCLUI_ITEM
      Case "consulta"
          If txtPedido.Text <> "" Then
             SQL3 = ""
             frmProdutoConsulta.optPA.Value = True
             frmProdutoConsulta.Show 1
             If SQL3 <> "" Then
                txtProduto.Text = SQL3
                txtProduto.SetFocus
             End If
             SQL3 = ""
          End If
          If SSTab1.Tab = 1 Then
             SQL3 = ""
             frmProdutoConsulta.optPA.Value = True
             frmProdutoConsulta.Show 1
             If SQL3 <> "" Then
                txtprodutoconsulta.Text = SQL3
                txtprodutoconsulta.SetFocus
             End If
             SQL3 = ""
          End If
      Case "exclui_lote"
            MATA_PEDIDO_CABECA
    End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtPEDIDO_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear

   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Tecle <ENTER> para gerar novo Pedido ou informe uma já existente"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpedido_GotFocus"
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPedido.Text = "" Then
         GERA_NUMR_PEDIDO_COMPRA
         txtPedido.Text = NUMR_COMPRA_N
         txtCgcCpf.SetFocus
         Else
            NUMR_COMPRA_N = Int(txtPedido.Text)
            SQL = "select * from PEDIDOCOMPRA "
            SQL = SQL & " where PEDIDOCOMPRA_id = " & NUMR_COMPRA_N
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TABCOMPRA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TABCOMPRA.EOF Then
               TABCOMPRA.Close
               Indr_Lote_Existe = False
               Beep
               MsgBox "Pedido não encontrado.", vbOKOnly, "Erro."
               Exit Sub
               Else
                  Indr_Lote_Existe = True
                  MOSTRA_DADOS_PEDIDO
                  End If
            End If
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyPress"
End Sub

Private Sub TXTCGCCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear

   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Fornecedores"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Inform Fornecedor"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents
   
   txtCgcCpf.PromptInclude = False
   txtCgcCpf.Mask = "###############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_GotFocus"
End Sub

Private Sub TXTCGCCPFCONSULTA_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear

   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Fornecedores"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Inform Fornecedor"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents
   
   txtcgccpfconsulta.PromptInclude = False
   txtcgccpfconsulta.Mask = "###############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPFCONSULTA_GotFocus"
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCgcCpf.PromptInclude = False
            txtCgcCpf.Text = CNPJCPF_A
         End If
         CNPJCPF_A = ""
         
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_KeyDown"
End Sub

Private Sub TXTCGCCPFCONSULTA_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtcgccpfconsulta.PromptInclude = False
            txtcgccpfconsulta.Text = CNPJCPF_A
         End If
         CNPJCPF_A = ""
         
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPFCONSULTA_KeyDown"
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      ENDERECO_A = ""
      txtCgcCpf.PromptInclude = False
      If txtCgcCpf.Text = "" Then
         txtCgcCpf.Text = "99999999999"
         Else
            If Len(txtCgcCpf.Text) > 0 Then
               Select Case Len(txtCgcCpf.Text)
                  Case Is = 11
                    If Not CALCULACPF(txtCgcCpf.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtCgcCpf.PromptInclude = False
                       txtCgcCpf = ""
                       txtCgcCpf.SetFocus
                       Exit Sub
                    End If
                  Case Is = 14
                    If Not VALIDACGC(txtCgcCpf.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtCgcCpf.PromptInclude = False
                       txtCgcCpf = ""
                       txtCgcCpf.SetFocus
                       Exit Sub
                    End If
                  Case Is > 14
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCgcCpf = ""
                     txtCgcCpf.SetFocus
                     Exit Sub
                  Case Is < 11
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCgcCpf = ""
                     txtCgcCpf.SetFocus
                     Exit Sub
               End Select
               Else
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCgcCpf = ""
                  txtCgcCpf.SetFocus
                  Exit Sub
            End If
            txtCgcCpf.PromptInclude = False
            CRITERIO = txtCgcCpf.Text
      End If
      txtCgcCpf.PromptInclude = False
      If txtCgcCpf.Text <> "" Then
         CRITERIO = txtCgcCpf.Text
         If Not IsNull(txtCgcCpf.Text) Then
            If Len(txtCgcCpf.Text) <= 11 Then
               txtCgcCpf.Mask = "###.###.###-##"
               Else: txtCgcCpf.Mask = "##.###.###/####-##"
            End If
         End If
         txtCgcCpf.Text = CRITERIO
      End If
      txtCgcCpf.PromptInclude = False
      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where CGCCPF = '" & txtCgcCpf.Text & "'"
      If TabFORNECEDOR.State = 1 Then TabFORNECEDOR.Close
      TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabFORNECEDOR.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção."
         txtCgcCpf.SetFocus
         Exit Sub
         Else
            If TabFORNECEDOR!NOME <> "" Then
               txtNome.Text = TabFORNECEDOR!NOME
            End If
      End If
         Beep
         Msg = "Deseja Carregar os Itens Cadastrados Para Este Fornecedor?"
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbNo Then
            txtPrazo.SetFocus
            Exit Sub
         End If
         If RESPOSTA = vbYes Then
            CARREGA_ITENS_FORNECEDOR
            Exit Sub
         End If
         txtPrazo.SetFocus
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_KeyPress"
End Sub

Private Sub txtCGCCPFCONSULTA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      ENDERECO_A = ""
      txtcgccpfconsulta.PromptInclude = False
      If txtcgccpfconsulta.Text = "" Then
         txtcgccpfconsulta.Text = "99999999999"
         Else
            If Len(txtcgccpfconsulta.Text) > 0 Then
               Select Case Len(txtcgccpfconsulta.Text)
                  Case Is = 11
                    If Not CALCULACPF(txtcgccpfconsulta.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtcgccpfconsulta.PromptInclude = False
                       txtcgccpfconsulta = ""
                       txtcgccpfconsulta.SetFocus
                       Exit Sub
                    End If
                  Case Is = 14
                    If Not VALIDACGC(txtcgccpfconsulta.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtcgccpfconsulta.PromptInclude = False
                       txtcgccpfconsulta = ""
                       txtcgccpfconsulta.SetFocus
                       Exit Sub
                    End If
                  Case Is > 14
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtcgccpfconsulta = ""
                     txtcgccpfconsulta.SetFocus
                     Exit Sub
                  Case Is < 11
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtcgccpfconsulta = ""
                     txtcgccpfconsulta.SetFocus
                     Exit Sub
               End Select
               Else
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtcgccpfconsulta = ""
                  txtcgccpfconsulta.SetFocus
                  Exit Sub
            End If
            txtcgccpfconsulta.PromptInclude = False
            CRITERIO = txtcgccpfconsulta.Text
      End If
      txtcgccpfconsulta.PromptInclude = False
      If txtcgccpfconsulta.Text <> "" Then
         CRITERIO = txtcgccpfconsulta.Text
         If Not IsNull(txtcgccpfconsulta.Text) Then
            If Len(txtcgccpfconsulta.Text) <= 11 Then
               txtcgccpfconsulta.Mask = "###.###.###-##"
               Else: txtcgccpfconsulta.Mask = "##.###.###/####-##"
            End If
         End If
         txtcgccpfconsulta.Text = CRITERIO
      End If
      txtcgccpfconsulta.PromptInclude = False
      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where CGCCPF = '" & txtcgccpfconsulta.Text & "'"
      TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabFORNECEDOR.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção."
         txtcgccpfconsulta.SetFocus
         Exit Sub
         Else
            If TabFORNECEDOR!NOME <> "" Then
               txtdescricaoconsulta.Text = TabFORNECEDOR!NOME
            End If
      End If
         CARREGA_ITENS_FORNECEDOR_CONSULTA
         cmdproximo.SetFocus
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPFCONSULTA_KeyPress"
End Sub

Private Sub txtnomeconsulta_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      UCase (txtnomeconsulta.Text)
      KeyAscii = 0

      CRITERIO = UCase(txtnomeconsulta.Text) & "*"

      SQL = "select * from PRODUTO "
      SQL = SQL & " where descricao like '" & CRITERIO & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      SQL = SQL & " order by descricao"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProduto.EOF Then
         MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
         txtprodutoconsulta.SelStart = 0
         txtprodutoconsulta.SelLength = Len(txtProduto)
         txtprodutoconsulta.SetFocus
         Exit Sub
         Else
            txtprodutoconsulta.Text = Trim(TabProduto!Codg_Prod)
            txtnomeconsulta.Text = Trim(TabProduto!Descricao)
      End If
      SETA_GRID_ITENS_FORNEC_CONSULTA
      MOSTRA_DADOS_CONSULTA
      MOSTRA_ULTIMAS_COMPRAS
      txtnomeconsulta.SetFocus
      If KeyAscii = 8 Then
         Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnomeconsulta_KeyPress"
End Sub

Private Sub txtprazo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtFinanceiro.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprazo_KeyPress"
End Sub

Private Sub txtfinanceiro_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfinanceiro_KeyPress"
End Sub

Private Sub txtprodutoconsulta_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      UCase (txtprodutoconsulta.Text)
      SP_PROCURA_PRODUTO EMPRESA_ID_N, Trim(txtprodutoconsulta.Text), 0, "", "", "", 1
      If TabProduto.State = 1 Then TabProduto.Close
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProduto.EOF Then
         MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
         txtprodutoconsulta.SelStart = 0
         txtprodutoconsulta.SelLength = Len(txtProduto)
         txtprodutoconsulta.SetFocus
         Exit Sub
         Else
            txtprodutoconsulta.Text = Trim(TabProduto!CODG_PRODUTO)
            txtnomeconsulta.Text = Trim(TabProduto!Descricao)
       End If
       MOSTRA_DADOS_CONSULTA
       SETA_GRID_ITENS_FORNEC_CONSULTA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprodutoconsulta_KeyPress"
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text = "" Then _
      txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If SSTab1.Tab = 0 Then
         If txtseq.Text = "" Then
            NUMR_SEQ_N = 1
            SQL = "select max(SEQUENCIA) as ultimo_reg from PEDIDOCOMPRAITEM "
            SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
            TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TaBPedidoCompraItem.EOF Then _
               If Not IsNull(TaBPedidoCompraItem!ultimo_reg) Then _
                 NUMR_SEQ_N = TaBPedidoCompraItem!ultimo_reg + 1
            TaBPedidoCompraItem.Close
            txtseq.Text = NUMR_SEQ_N
            Else
                SQL = "select * from PEDIDOCOMPRAITEM "
                SQL = SQL & " where sequencia = " & txtseq.Text
                SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
                TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                If Not TaBPedidoCompraItem.EOF Then
                   txtProduto.Text = Trim(TaBPedidoCompraItem!Produto)
                   SP_PROCURA_PRODUTO EMPRESA_ID_N, Trim(TaBPedidoCompraItem!Produto), 0, "", "", "", 1
                   If Not TabProduto.EOF Then
                   
                       txtDescricao.Text = TabProduto!Descricao
                       If Not IsNull(TabProduto!PRECO_CUSTO_ANTERIOR) Then
                          txt_custo_anterior.Text = TabProduto!PRECO_CUSTO_ANTERIOR
                       End If
                       If Not IsNull(TabProduto!qtd_ped_anterior) Then
                          txt_qtd_anterior.Text = TabProduto!qtd_ped_anterior
                       End If
                       
                       Beep
                       Msg = "Seqüência Já Existente Nesse Pedido, Deseja Alterar?"
                       Style = vbYesNo + 32
                       Title = "Atenção."
                       Help = "DEMO.HLP"
                       Ctxt = 1000
                       RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                       If RESPOSTA = vbNo Then
                          TabProduto.Close
                          TaBPedidoCompraItem.Close
                          LIMPA_BODY
                          txtProduto.SetFocus
                          Exit Sub
                       End If
                       If RESPOSTA = vbYes Then
                          TaBPedidoCompraItem.Close
                          TabProduto.Close
                          txtProduto.SetFocus
                          Exit Sub
                       End If
                   End If
                   TabProduto.Close
                End If
         End If
         txtProduto.SetFocus
      End If
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyPress"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtseq.Text = "" And txtPedido.Text <> "" Then
      NUMR_SEQ_N = 1
      SQL = "select max(sequencia) as ultimo_reg from PEDIDOCOMPRAITEM "
      SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
      TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TaBPedidoCompraItem.EOF Then _
         If Not IsNull(TaBPedidoCompraItem!ultimo_reg) Then _
            NUMR_SEQ_N = TaBPedidoCompraItem!ultimo_reg + 1
      TaBPedidoCompraItem.Close
      txtseq.Text = NUMR_SEQ_N
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If txtPedido.Text <> "" Then
        If KeyAscii = 13 Then
           KeyAscii = 0
           UCase (txtProduto.Text)
           'SP_PROCURA_PRODUTO empresa_id_n, Trim(txtProduto.Text), 0, "", "", "", 1

           SQL = "select * from PRODUTO "
           SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
           SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
           SQL = SQL & " and situacao <> 'C' "
           If TabProduto.State = 1 Then TabProduto.Close
           TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
           If TabProduto.EOF Then
              MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
              txtProduto.SelStart = 0
              txtProduto.SelLength = Len(txtProduto)
              txtProduto.SetFocus
              Exit Sub
              Else
                 'txtProduto.Text = Trim(TABPRODUTO!Codg_Prod)
                 txtDescricao.Text = Trim(TabProduto!Descricao)
                 If Not IsNull(TabProduto!PRECO_CUSTO_ANTERIOR) Then
                    txt_custo_anterior.Text = TabProduto!PRECO_CUSTO_ANTERIOR
                 End If
                 If Not IsNull(TabProduto!qtd_ped_anterior) Then
                    txt_qtd_anterior.Text = TabProduto!qtd_ped_anterior
                 End If
                 'txtquantidade.SetFocus
                 
                 
                 SQL = "select * from PEDIDOCOMPRAITEM "
                 SQL = SQL & " where produto = '" & Trim(txtProduto.Text) & "'"
                 SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
                 TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                 If Not TaBPedidoCompraItem.EOF Then
                    txtseq.Text = TaBPedidoCompraItem!sequencia
                    txtquantidade.Text = TaBPedidoCompraItem!Qtd
                    txtPreco.Text = TaBPedidoCompraItem!Preco
                    MsgBox "Produto já consta nesse Lote seqüência = " & TaBPedidoCompraItem!sequencia
                 End If
                 txtquantidade.SetFocus
                 TabProduto.Close
                 TaBPedidoCompraItem.Close
           End If
           'txtquantidade.SetFocus
        End If
        Else:
           MsgBox "Para Digitar os Produtos gere o numero do Pedido!"
           txtPedido.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtquantidade_GotFocus()
'On Error GoTo ERRO_TRATA

   SQL = "select * from PEDIDOCOMPRAITEM "
   SQL = SQL & " where produto = '" & Trim(txtProduto.Text) & "'"
   SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
   TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TaBPedidoCompraItem.EOF Then
      If TaBPedidoCompraItem!Qtd <> "" Then
         txtquantidade.Text = TaBPedidoCompraItem!Qtd
         'txtPreco.SetFocus
         Exit Sub
      End If
   End If
   TaBPedidoCompraItem.Close
   
   If txtseq.Text = Empty Then
      MsgBox "Numero de seqüência inválido.", vbOKOnly, "Erro."
      txtseq.Text = 1
      txtProduto.SetFocus
      Exit Sub
   End If
   If Trim(txtProduto.Text) = Empty Then
      MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
      txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   If txtquantidade.Text <> "" Then
      txtquantidade.SelStart = 0
      txtquantidade.SelLength = Len(txtquantidade)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtquantidade_GotFocus"
End Sub

Private Sub txtquantidade_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'If txtquantidade.Text <> "" Then
         txtPreco.SetFocus
       '  Exit Sub
        ' Else
       '  txtPreco.SetFocus
      'End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtquantidade_KeyPress"
End Sub

Private Sub txtpreco_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtPreco.Text = "" Then
      txtPreco.Text = 0
      'txtPreco.SelStart = 0
      txtPreco.SelLength = Len(txtPreco.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpreco_GotFocus"
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPreco.Text <> "" Or txtPreco.Text = 0 Then
         GRAVA_PEDIDOCOMPRAITEM
         LIMPA_BODY
         txtProduto.SetFocus
         Exit Sub
         Else
         txtProduto.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpreco_KeyPress"
End Sub

Private Sub GRAVA_CABECA()
'On Error GoTo ERRO_TRATA

   SQL = "select * from PEDIDOCOMPRA "
   SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TABCOMPRA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TABCOMPRA.EOF Then
      
      SQL = "insert into PEDIDOCOMPRA values ("
      SQL = SQL & EMPRESA_ID_N
      SQL = SQL & "," & txtPedido.Text & ","
      SQL = SQL & "'" & txtCgcCpf.Text & "'"
      SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)
      SQL = SQL & "," & 0
      If txtFinanceiro.Text <> "" Then
         SQL = SQL & "," & tpMOEDA(txtFinanceiro.Text)
         Else: SQL = SQL & "," & 0
      End If
      If txtPrazo.Text <> "" Then
         SQL = SQL & "," & txtPrazo.Text
         Else: SQL = SQL & "," & 0
      End If
      
      CRITERIO = DMA(txtdatapedido.Text)
      SQL = SQL & ",'" & CRITERIO & "'"
      
      If Not IsNull(CODG_USU_N) Then _
         SQL = SQL & "," & CODG_USU_N
      
      SQL = SQL & ",'" & "S" & "'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      
   End If
   TABCOMPRA.Close
   
   Msg = "Deseja Imprimir Pedido de Compras?"
   Style = vbYesNo + 32
   Title = "Atenção."
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then
      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_Pedido_Compras.rpt"
      frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA"
End Sub

Private Sub GRAVA_PEDIDOCOMPRAITEM()
'On Error GoTo ERRO_TRATA

    If SSTab1.Tab = 0 Then
        SQL = "select * from PEDIDOCOMPRAITEM "
        SQL = SQL & " where produto = '" & Trim(txtProduto.Text) & "'"
        SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
        TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
        If TaBPedidoCompraItem.EOF Then
            SQL = "insert into PEDIDOCOMPRAITEM ( PEDIDOCOMPRA_id, PRODUTO, SEQUENCIA, PRECO, QTD) "
            SQL = SQL & " values ("
            SQL = SQL & txtPedido.Text & ","
            SQL = SQL & "'" & txtProduto.Text & "'"
            SQL = SQL & "," & txtseq.Text
            SQL = SQL & "," & tpMOEDA(txtPreco.Text)
            SQL = SQL & "," & tpMOEDA(txtquantidade.Text)
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
            Else 'Existe registro so altera !
               SQL = "update PEDIDOCOMPRAITEM set "
               SQL = SQL & "qtd = " & Replace(txtquantidade.Text, ",", ".")
               SQL = SQL & ",preco = " & Replace(txtPreco.Text, ",", ".")
               SQL = SQL & " where produto = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
               CONECTA_RETAGUARDA.Execute SQL
        End If
        TaBPedidoCompraItem.Close
        
        'Registrando Qtd para proxima compra, a nivel de controle
        SQL = "select * from PRODUTO "
        SQL = SQL & " where CODG_Produto = '" & Trim(txtProduto.Text) & "'"
        SQL = SQL & " and situacao <> 'C' "
        If TabProduto.State = 1 Then TabProduto.Close
        TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
        If Not TabProduto.EOF Then
           SQL = "update PRODUTO set "
           SQL = SQL & "qtd_ped_anterior = " & tpMOEDA(txtquantidade.Text)
           SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
           CONECTA_RETAGUARDA.Execute SQL
        End If
        TabProduto.Close
        
    End If
    SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PEDIDOCOMPRAITEM"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then
      listaprod.ListItems.Clear
      NUMR_SEQ_N = 0
      CONT_N = 0
      SQL = "select * from PEDIDOCOMPRAITEM "
      SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
      SQL = SQL & " order by sequencia "
      TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TaBPedidoCompraItem.EOF
         CONT_N = CONT_N + 1
         Set Item = listaprod.ListItems.Add(, "seq." & CONT_N, TaBPedidoCompraItem!sequencia)
         Item.SubItems(1) = Trim(TaBPedidoCompraItem!Produto)
         SP_PROCURA_PRODUTO EMPRESA_ID_N, Trim(TaBPedidoCompraItem!Produto), 0, "", "", "", 1
         If Not TabProduto.EOF Then _
            Item.SubItems(2) = Trim(TabProduto!Descricao)
            Item.SubItems(3) = TaBPedidoCompraItem!Qtd
            If Not IsNull(TabProduto!qtd_ped_anterior) Then
               Item.SubItems(4) = TabProduto!qtd_ped_anterior
            End If
            Item.SubItems(5) = Format(TaBPedidoCompraItem!Preco, strFormatacao2Digitos)
            Item.SubItems(6) = Format((TaBPedidoCompraItem!Qtd * TaBPedidoCompraItem!Preco), strFormatacao2Digitos)
            Item.SubItems(7) = Format(TabProduto!PRECO_CUSTO_ANTERIOR, strFormatacao2Digitos)
            TabProduto.Close
            TaBPedidoCompraItem.MoveNext
            'Valor Total do pedido
            VALOR_TOTAL_N = 0
            SqL2 = "select sum(i.preco * i.qtd) from PEDIDOCOMPRAITEM i"
            SqL2 = SqL2 & " where i.PEDIDOCOMPRA_id = " & NUMR_COMPRA_N
            TabTemp.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               VALOR_TOTAL_N = TabTemp.Fields(0).Value
            TabTemp.Close
            stBarpedido.Panels(2).Text = Format(VALOR_TOTAL_N, "##,##0.00")
         Wend
      TaBPedidoCompraItem.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

  txtPedido.Text = ""
  txtseq.Text = ""
  txtdatapedido.Mask = "##/##/####"
  txtProduto.Text = ""
  txtquantidade.Text = ""
  txtPreco.Text = ""
  txt_custo_anterior.Text = ""
  txt_qtd_anterior.Text = ""
  txtDescricao.Text = ""
  txtCgcCpf.Text = ""
  txtNome.Text = ""
  txtFinanceiro.Text = ""
  txtPrazo.Text = ""
  listaprod.ListItems.Clear
  NUMR_COMPRA_N = 0
  NUMR_SEQ_N = 0
  txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   txtseq.Text = ""
   txtDescricao.Text = ""
   txtquantidade.Text = ""
   txtPreco.Text = ""
   txt_custo_anterior.Text = ""
   txt_qtd_anterior.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub MOSTRA_DADOS_PEDIDO()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then
      SQL = "select * from PEDIDOCOMPRA "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
      TABCOMPRA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TABCOMPRA.EOF Then
          txtdatapedido.Text = TABCOMPRA!DT_CRIACAO
          STATUS_A = TABCOMPRA!PROCESSADO
          txtFinanceiro.Text = TABCOMPRA!PERC_FIN
          txtPrazo.Text = TABCOMPRA!DIAS_PZ
          
          SqL2 = "select * from FORNECEDOR "
          SqL2 = SqL2 & " where CGCCPF = '" & TABCOMPRA!CGCCPF & "'"
          TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
          
          txtCgcCpf.Text = TABCOMPRA!CGCCPF
          txtNome.Text = TabFORNECEDOR!NOME
          
      End If
      TABCOMPRA.Close
   End If
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PEDIDO"
End Sub

Private Sub EXCLUI_ITEM()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then
      If txtProduto.Text <> "" Then
         SQL = "select * from PEDIDOCOMPRAITEM "
         SQL = SQL & " where produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            Msg = "Deseja Excluir Esse Item?"
            Style = vbYesNo + 32
            Title = "Atenção."
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then
               TabTemp.Delete
               TabTemp.Close
               LIMPA_BODY
               SETA_GRID
               Else: TabTemp.Close
            End If
            Else: MsgBox "Produto não encontrado."
         End If
         txtProduto.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUI_ITEM"
End Sub

Private Sub MATA_PEDIDO_CABECA()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" Then
      SQL = "select * from PEDIDOCOMPRA "
      SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Deseja Excluir Esse Pedido?"
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            CONECTA_RETAGUARDA.Execute "Delete From PEDIDOCOMPRA Where PEDIDOCOMPRA_id = " & txtPedido.Text
            TabTemp.Close
            LIMPA_TUDO
            MATA_PEDIDO_ITENS
            Else: TabTemp.Close
         End If
         Else: MsgBox "Pedido nao Localizado."
      End If
      txtPedido.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_PEDIDO_CABECA"
End Sub

Private Sub MATA_PEDIDO_ITENS()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" Then
      SQL = "select * from PEDIDOCOMPRAITEM "
      SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         CONECTA_RETAGUARDA.Execute "Delete From PEDIDOCOMPRAITEM Where PEDIDOCOMPRA_id = " & txtPedido.Text
      End If
      TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_PEDIDO_ITENS"
End Sub

Private Sub CARREGA_ITENS_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then
      SETA_GRID_ITENS_FORNEC
      txtPrazo.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_ITENS_FORNECEDOR"
End Sub

Private Sub CARREGA_ITENS_FORNECEDOR_CONSULTA()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 1 Then
      SETA_GRID_ITENS_FORNEC_CONSULTA
      cmdproximo.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_ITENS_FORNECEDOR_CONSULTA"
End Sub

Private Sub SETA_GRID_ITENS_FORNEC()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then
      listaprod.ListItems.Clear
      NUMR_SEQ_N = 0
      NUMR_SEQ_N = 1
      CONT_N = 0
      SQL = "select * from PRODUTO "
      SQL = SQL & " where fornecedor_id = " & FORNEC_ID_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      SQL = SQL & " order by DESCRICAO "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabProduto.EOF
         
         CONT_N = CONT_N + 1
         SQL = "select max(SEQUENCIA) as ultimo_reg from PEDIDOCOMPRAITEM "
         SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
         TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TaBPedidoCompraItem.EOF Then _
            If Not IsNull(TaBPedidoCompraItem!ultimo_reg) Then _
               NUMR_SEQ_N = TaBPedidoCompraItem!ultimo_reg + 1
            TaBPedidoCompraItem.Close
         txtseq.Text = NUMR_SEQ_N
         
         Set Item = listaprod.ListItems.Add(, "seq." & CONT_N, NUMR_SEQ_N)
         Item.SubItems(1) = Trim(TabProduto!Codg_Prod)
         Item.SubItems(2) = Trim(TabProduto!Descricao)
         Item.SubItems(3) = TabProduto!Qtd
         
         If Not IsNull(TabProduto!qtd_ped_anterior) Then
            Item.SubItems(4) = Format(TabProduto!qtd_ped_anterior, strFormatacao2Digitos)
            Item.SubItems(6) = 0 'Format((TABPRODUTO!qtd_ped_anterior * TABPRODUTO!preco_custo), strFormatacao2Digitos)
            Else
               Item.SubItems(4) = 0
               Item.SubItems(6) = 0
         End If
         Item.SubItems(5) = Format(TabProduto!preco_custo, strFormatacao2Digitos)
         
         If Not IsNull(TabProduto!PRECO_CUSTO_ANTERIOR) Then
            Item.SubItems(7) = Format(TabProduto!PRECO_CUSTO_ANTERIOR, strFormatacao2Digitos)
            Else
               Item.SubItems(7) = 0
         End If
         stBarpedido.Panels(2).Text = Format(VALOR_TOTAL_N, "##,##0.00")
      Wend
      TabProduto.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ITENS_FORNEC"
End Sub

Private Sub SETA_GRID_ITENS_FORNEC_CONSULTA()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 1 Then
      listahist_ent.ListItems.Clear
      listahist_sai.ListItems.Clear
      Dim qtd_mes_01 As Integer
      Dim data_controle As Data
           
      SQL = "Select * From QryFinalKardex "
      SQL = SQL & " Where QryFinalKardex.codg_prod ='" & txtprodutoconsulta.Text & "'"
      SQL = SQL & " order by DT_ENTRADA "
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCABECA.EOF = False Then
         TabCABECA.MoveFirst
         Do While Not TabCABECA.EOF
            qtd_mes_01 = 0
            
            
            DoEvents
            If TabCABECA!tipo = "ENTRADA" Then
               Do While Not Format("/01/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_01 = (TabCABECA!QTD_ENTRADA + qtd_mes_01)
            Loop
            Do While Not Format("/02/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_02 = (TabCABECA!QTD_ENTRADA + qtd_mes_02)
            Loop
            Do While Not Format("/03/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_03 = (TabCABECA!QTD_ENTRADA + qtd_mes_03)
            Loop
            Do While Not Format("/04/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_04 = (TabCABECA!QTD_ENTRADA + qtd_mes_04)
            Loop
            Do While Not Format("/05/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_05 = (TabCABECA!QTD_ENTRADA + qtd_mes_05)
            Loop
            Do While Not Format("/06/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_06 = (TabCABECA!QTD_ENTRADA + qtd_mes_06)
            Loop
            Do While Not Format("/07/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_07 = (TabCABECA!QTD_ENTRADA + qtd_mes_07)
            Loop
            Do While Not Format("/08/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_08 = (TabCABECA!QTD_ENTRADA + qtd_mes_08)
            Loop
            Do While Not Format("/09/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_09 = (TabCABECA!QTD_ENTRADA + qtd_mes_09)
            Loop
            Do While Not Format("/10/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_10 = (TabCABECA!QTD_ENTRADA + qtd_mes_10)
            Loop
            Do While Not Format("/11/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_11 = (TabCABECA!QTD_ENTRADA + qtd_mes_11)
            Loop
            Do While Not Format("/12/") = Format(TabCABECA!Data, "/mm/")
               qtd_mes_12 = (TabCABECA!QTD_ENTRADA + qtd_mes_12)
            Loop
               
               Set Item = listahist_ent.ListItems.Add(, , TabCABECA!QTD_ENTRADA)
               qtd_mes = (TabCABECA!QTD_ENTRADA + qtd_mes)
               Item.SubItems(1) = TabCABECA!Data
               Item.SubItems(2) = qtd_mes_01
               Item.SubItems(3) = qtd_mes_02
               Item.SubItems(4) = qtd_mes_03
               Item.SubItems(5) = qtd_mes_04
               Item.SubItems(6) = qtd_mes_05
               Item.SubItems(7) = qtd_mes_06
               Item.SubItems(8) = qtd_mes_07
               Item.SubItems(9) = qtd_mes_08
               Item.SubItems(10) = qtd_mes_09
               Item.SubItems(11) = qtd_mes_10
               Item.SubItems(12) = qtd_mes_11
               Item.SubItems(13) = qtd_mes_12
               
            End If
            If TabCABECA!tipo = "SAIDA" Then
               Set Item = listahist_sai.ListItems.Add(, , TabCABECA!QTD_ENTRADA)
               Item.SubItems(1) = TabCABECA!Data
               Item.SubItems(2) = TabCABECA!QTD_ENTRADA
               'lst.SubItems(3) = TABCABECA!QTD_ENTRADA
               'lst.SubItems(4) = TABCABECA!QTD_ENTRADA
               'lst.SubItems(5) = TABCABECA!QTD_ENTRADA
               'lst.SubItems(6) = TABCABECA!QTD_ENTRADA
               'lst.SubItems(7) = TABCABECA!QTD_ENTRADA
            End If
            TabCABECA.MoveNext
            Loop
        'Loop
        
        Else
        MsgBox "Nao há nenhum registro para a consulta realizada. Verifique os parametros e tente novamente"
      End If
      TabCABECA.Close
      
   listahist_sai.Refresh
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ITENS_FORNEC_CONSULTA"
End Sub

Private Sub MOSTRA_DADOS_CONSULTA()
'On Error GoTo ERRO_TRATA

   SQL = "select * from PRODUTO "
   If txtcgccpfconsulta.Text <> "" Then
      SQL = SQL & " where CGCCPF = '" & txtcgccpfconsulta.Text & "'"
      Else: SQL = SQL & " where codg_produto ='" & txtprodutoconsulta.Text & "'"
   End If
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   SQL = SQL & " order by DESCRICAO desc"
   If TabProduto.State = 1 Then TabProduto.Close
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      TabProduto.MoveFirst
      txtestoque.Text = TabProduto!Qtde
      txtcusto.Text = TabProduto!preco_custo
      txtatacado.Text = TabProduto!PRECO_ATACADO
      txtvarejo.Text = TabProduto!PRECO_VENDA
      If txtprodutoconsulta.Text = "" Then _
         txtprodutoconsulta.Text = TabProduto!CODG_PRODUTO
      If txtnomeconsulta.Text = "" Then _
         txtnomeconsulta.Text = TabProduto!Descricao
      MOSTRA_ULTIMAS_COMPRAS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_CONSULTA"
End Sub

Private Sub MOSTRA_ULTIMAS_COMPRAS()
'On Error GoTo ERRO_TRATA

   SQL = "Select * From QryFinalKardex "
   SQL = SQL & " Where QryFinalKardex.codg_prod = '" & txtprodutoconsulta.Text & "'"
   'SQL = SQL & " and QryFinalKardex.tipo = '" & "ENTRADA" & "'"
   'SQL = SQL & " and QryFinalKardex.data >= " & Format(Date, "dd/mm/yyyy")
   SQL = SQL & " order by DT_ENTRADA desc"
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      'Pegando Ultimo Registro para Informacao da Ultima Compra
      TabCABECA.MoveFirst
      txtqtdult.Text = TabCABECA!QTD_ENTRADA
      txtvlrcomprault.Text = TabCABECA!preco_custo
      txtdatault = TabCABECA!Data
      
      'Pegando penultimo Registro para Informacao da Ultima Compra
      TabCABECA.MoveNext
      txtqtdpen.Text = TabCABECA!QTD_ENTRADA
      txtvlrcomprapen.Text = TabCABECA!preco_custo
      txtdtpen = TabCABECA!Data
   End If
   TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ULTIMAS_COMPRAS"
End Sub

Private Sub CONSULTA_PROXIMO_REGISTRO()
'On Error GoTo ERRO_TRATA

   SQL = "select * from PRODUTO "
   
   CRITERIO = UCase(txtnomeconsulta.Text) & "*"
   If txtcgccpfconsulta.Text <> "" Then
      SQL = SQL & " where CGCCPF = '" & txtcgccpfconsulta.Text & "'"
      Else: SQL = SQL & " where descricao like '" & CRITERIO & "'"
   End If
   SQL = SQL & " and situacao <> 'C' "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO desc"
   If TabProduto.State = 1 Then TabProduto.Close
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      TabProduto.MoveNext
      DoEvents
      txtestoque.Text = TabProduto!Qtd
      txtcusto.Text = TabProduto!preco_custo
      txtatacado.Text = TabProduto!PRECO_ATACADO
      txtvarejo.Text = TabProduto!PRECO_VENDA
      txtprodutoconsulta.Text = TabProduto!Codg_Prod
      txtnomeconsulta.Text = TabProduto!Descricao
   End If
   MOSTRA_ULTIMAS_COMPRAS
   TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PROXIMO_REGISTRO"
End Sub

Private Sub CONSULTA_ANTERIOR_REGISTRO()
'On Error GoTo ERRO_TRATA

   CRITERIO = UCase(txtnomeconsulta.Text) & "*"
   SQL = "select * from PRODUTO "
   If txtcgccpfconsulta.Text <> "" Then
      SQL = SQL & " where CGCCPF = '" & txtCgcCpf.Text & "'"
      Else: SQL = SQL & " where descricao like '" & CRITERIO & "'"
   End If
   SQL = SQL & " and situacao <> 'C' "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO desc "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      DoEvents
      TabProduto.MoveFirst
      txtestoque.Text = TabProduto!Qtd
      txtcusto.Text = TabProduto!preco_custo
      txtatacado.Text = TabProduto!PRECO_ATACADO
      txtvarejo.Text = TabProduto!PRECO_VENDA
      txtprodutoconsulta.Text = TabProduto!Codg_Prod
      txtnomeconsulta.Text = TabProduto!Descricao
      TabProduto.MoveNext
   End If
   MOSTRA_ULTIMAS_COMPRAS
   TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_ANTERIOR_REGISTRO"
End Sub
