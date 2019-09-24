VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Movimento_Nota_Fiscal_Entrada 
   Caption         =   "Entrada de Mercadorias"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   12240
   Icon            =   "Movimento_Nota_Fiscal_Entrada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Movimento_Nota_Fiscal_Entrada"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":08BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":60AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":63CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":66E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":72BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":CEDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":D33E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":DD52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":13546
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1916A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1EECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1F02E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1F18E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1F2EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1F44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1F8A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1FA02
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":1FE56
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":202AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":20CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":21110
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":21DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":235D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   4
      Text            =   "Não Apagar"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8040
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   14182
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados Principais"
      TabPicture(0)   =   "Movimento_Nota_Fiscal_Entrada.frx":2688D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmdados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grade1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "buscaxml"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Transportadora"
      TabPicture(1)   =   "Movimento_Nota_Fiscal_Entrada.frx":268A9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm_dados1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Parâmetros/Cálculo"
      TabPicture(2)   =   "Movimento_Nota_Fiscal_Entrada.frx":268C5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frm_dados2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Conferência dos Produtos"
      TabPicture(3)   =   "Movimento_Nota_Fiscal_Entrada.frx":268E1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmd_salvar_conferencia"
      Tab(3).Control(1)=   "cmd_pesquisar_conferencia"
      Tab(3).Control(2)=   "cmd_imprimir"
      Tab(3).Control(3)=   "grade3"
      Tab(3).Control(4)=   "frm_conf"
      Tab(3).ControlCount=   5
      Begin MSComDlg.CommonDialog buscaxml 
         Left            =   11520
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade1 
         Height          =   3405
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Pressione F4 para excluir item da grade"
         Top             =   3720
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6006
         _Version        =   393216
         RowHeightMin    =   300
         BackColorBkg    =   -2147483634
         GridColor       =   -2147483633
         FocusRect       =   0
         AllowUserResizing=   2
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton cmd_salvar_conferencia 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67680
         TabIndex        =   128
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmd_pesquisar_conferencia 
         Caption         =   "Pesquisar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -66120
         TabIndex        =   127
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -64560
         TabIndex        =   126
         Top             =   540
         Width           =   1380
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade3 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   125
         Top             =   1560
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   7858
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorBkg    =   16777215
         GridColor       =   -2147483637
         WordWrap        =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame frm_conf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   7665
         Left            =   -74970
         TabIndex        =   109
         Top             =   310
         Width           =   12165
         Begin VB.ComboBox cbo_unidade 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8070
            TabIndex        =   141
            Top             =   885
            Width           =   855
         End
         Begin VB.TextBox txt_totalqtde 
            Alignment       =   1  'Right Justify
            DataField       =   "Preco Compra"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10250
            TabIndex        =   123
            Top             =   5760
            Width           =   1755
         End
         Begin VB.TextBox txt_senha 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   10250
            PasswordChar    =   "*"
            TabIndex        =   122
            Top             =   6270
            Width           =   1755
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   4
            Left            =   2160
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":268FD
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   120
            Top             =   6308
            Width           =   255
         End
         Begin VB.TextBox txt_nome_usuario 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2445
            TabIndex        =   119
            Top             =   6270
            Width           =   4620
         End
         Begin VB.TextBox txt_codigo_usuario 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1425
            MaxLength       =   4
            TabIndex        =   118
            Top             =   6270
            Width           =   735
         End
         Begin VB.CheckBox chk_somaautomatica 
            Caption         =   "Soma Automática"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3045
            TabIndex        =   117
            ToolTipText     =   "O sistema soma as quantidades um a um de acordo que vai passando a leitora de código de barras"
            Top             =   180
            Width           =   1740
         End
         Begin VB.CheckBox chk_codigofornecedor 
            Caption         =   "Busca pelo Código Fornecedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   116
            ToolTipText     =   "Busca pelo código interno do fornecedor de acordo com o cadastro de produtos"
            Top             =   180
            Width           =   3060
         End
         Begin VB.TextBox txt_qtdeconf 
            Alignment       =   1  'Right Justify
            DataField       =   "Preco Compra"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   115
            Top             =   885
            Width           =   1395
         End
         Begin VB.CommandButton cmd_inserir 
            Caption         =   "Inserir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   893
            Width           =   1380
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   3
            Left            =   1680
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":26C3F
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   112
            Top             =   920
            Width           =   255
         End
         Begin VB.TextBox txt_descricao 
            DataField       =   "Preco Compra"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1980
            TabIndex        =   111
            Top             =   885
            Width           =   5940
         End
         Begin VB.TextBox txt_codigo_produto 
            Alignment       =   1  'Right Justify
            DataField       =   "Preco Compra"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   75
            TabIndex        =   110
            Top             =   885
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total de Quantidade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   8505
            TabIndex        =   124
            Top             =   5820
            Width           =   1635
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":26F81
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   121
            Top             =   6330
            Width           =   9720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":2704A
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   90
            TabIndex        =   113
            Top             =   675
            Width           =   9855
         End
      End
      Begin VB.Frame frm_dados2 
         Height          =   7665
         Left            =   -74970
         TabIndex        =   72
         Top             =   310
         Width           =   12135
         Begin VB.Frame Frame5 
            Caption         =   "Incidência no Cálculo do IPI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1700
            Left            =   90
            TabIndex        =   149
            Top             =   2880
            Width           =   2655
            Begin VB.CheckBox chk_frete_ipi 
               Caption         =   "Frete da Nota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   151
               Top             =   240
               Width           =   1635
            End
            Begin VB.CheckBox chk_outras_ipi 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   150
               Top             =   480
               Width           =   2115
            End
         End
         Begin VB.CheckBox chk_ContaConsumo 
            Caption         =   "Lançamento Conta de consumo Água/Luz/Gás e Comunicação"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2880
            TabIndex        =   130
            ToolTipText     =   "Utilizar quando for lançamento de Agua/Luz/Gás/Telefone"
            Top             =   2760
            Width           =   5535
         End
         Begin VB.Frame frm_LancamentoConsumo 
            Height          =   1695
            Left            =   2760
            TabIndex        =   131
            Top             =   2880
            Width           =   8865
            Begin VB.ComboBox cbo_CodigoConsumo 
               Height          =   315
               Left            =   2880
               Style           =   2  'Dropdown List
               TabIndex        =   140
               Top             =   360
               Width           =   5415
            End
            Begin VB.ComboBox cbo_TipoContaConsumo 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   360
               Width           =   2535
            End
            Begin VB.Frame frm_DetalhesContaEnergia 
               Caption         =   "Conta de energia"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   132
               Top             =   720
               Width           =   8655
               Begin VB.ComboBox cbo_GrupoDeTensao 
                  Height          =   315
                  Left            =   4200
                  Style           =   2  'Dropdown List
                  TabIndex        =   134
                  Top             =   480
                  Width           =   3975
               End
               Begin VB.ComboBox cbo_TipodeLigacao 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   133
                  Top             =   480
                  Width           =   3855
               End
               Begin VB.Label lbl_TituloEnergia 
                  Caption         =   "Grupo de tensão"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   4200
                  TabIndex        =   138
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.Label lbl_TituloEnergia 
                  Caption         =   "Tipo de ligação                                             "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   135
                  Top             =   240
                  Width           =   2895
               End
            End
            Begin VB.Label lbl_TituloCC 
               Caption         =   "Código do consumo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   139
               Top             =   120
               Width           =   1815
            End
            Begin VB.Label lbl_TituloCC 
               Caption         =   "Tipo de conta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   137
               Top             =   120
               Width           =   2970
            End
         End
         Begin VB.CheckBox chk_calcularpesos 
            Caption         =   "Calcular Peso Bruto/Liquido Aut."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8625
            TabIndex        =   129
            ToolTipText     =   "(Peso x Quantidade) baseado no cadastro do produto"
            Top             =   150
            Value           =   1  'Checked
            Width           =   2970
         End
         Begin VB.CheckBox chk_atualiza_carteira 
            Caption         =   "Carteira"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   106
            Top             =   480
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chk_impressao_nf 
            Caption         =   "Impressão N.F."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4305
            TabIndex        =   105
            Top             =   465
            Width           =   1575
         End
         Begin VB.CheckBox chk_retido 
            Caption         =   "Imp. Retido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4305
            TabIndex        =   104
            ToolTipText     =   "Quando se tratar do cálculo substituição tributária "
            Top             =   150
            Width           =   1335
         End
         Begin VB.CheckBox chk_calculo_nota 
            Caption         =   "Cálculo da Nota "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2190
            TabIndex        =   103
            ToolTipText     =   "Força uma entrada da nota pelo cálculo da nota e não pelo cálculo do sistema (Apenas em caráter de Urgencia)"
            Top             =   465
            Width           =   1575
         End
         Begin VB.CheckBox chk_atualiza_caixa 
            Caption         =   "Caixa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   102
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chk_lancamento_venda 
            Caption         =   "Lanc.Venda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   101
            Top             =   480
            Width           =   1290
         End
         Begin VB.CheckBox chk_complementar 
            Caption         =   "Nota Complementar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6240
            TabIndex        =   100
            ToolTipText     =   "Irá zerar os valores dos produtos aproveitando apenas os calculos do ICMS"
            Top             =   165
            Width           =   2145
         End
         Begin VB.ComboBox cbo_atualizacusto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   99
            ToolTipText     =   "Forma de Atualização no Cadastro de Produtos"
            Top             =   405
            Width           =   1860
         End
         Begin VB.Frame Frame1 
            Caption         =   "Incidência na Base Cálc.ICMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1785
            Left            =   90
            TabIndex        =   93
            Top             =   900
            Width           =   2655
            Begin VB.CheckBox chk_seguro 
               Caption         =   "Seguro"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   98
               Top             =   1200
               Width           =   1575
            End
            Begin VB.CheckBox chk_desc_bc 
               Caption         =   "Desconto"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   97
               Top             =   960
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox chk_soma_ipi 
               Caption         =   "I.P.I."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   720
               Width           =   1695
            End
            Begin VB.CheckBox chk_soma_outras 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   480
               Value           =   1  'Checked
               Width           =   2115
            End
            Begin VB.CheckBox chk_frete_nota 
               Caption         =   "Frete da Nota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Value           =   1  'Checked
               Width           =   1635
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Incidência na Base Cálc.Reduzida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1785
            Left            =   2760
            TabIndex        =   87
            Top             =   900
            Width           =   3015
            Begin VB.CheckBox chk_frete_red 
               Caption         =   "Frete da Nota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   240
               Value           =   1  'Checked
               Width           =   1635
            End
            Begin VB.CheckBox chk_outras_red 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   480
               Value           =   1  'Checked
               Width           =   2115
            End
            Begin VB.CheckBox chk_ipi_red 
               Caption         =   "I.P.I."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   720
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox chk_desc_red 
               Caption         =   "Desconto"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   89
               Top             =   960
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox chk_seguro_red 
               Caption         =   "Seguro"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   88
               Top             =   1200
               Width           =   1575
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Incidência na Base Cálc.Subst."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1785
            Left            =   5880
            TabIndex        =   81
            Top             =   900
            Width           =   2775
            Begin VB.CheckBox chk_fretenota_subst 
               Caption         =   "Frete da Nota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   86
               Top             =   1200
               Width           =   2400
            End
            Begin VB.CheckBox chk_outras_despesas 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   720
               Width           =   2400
            End
            Begin VB.CheckBox chk_red_subst 
               Caption         =   "Redução na Substituição"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   960
               Width           =   2400
            End
            Begin VB.CheckBox chk_ipi_subs 
               Caption         =   "I.P.I."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   480
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox chk_frete_subst 
               Caption         =   "Frete Conhecimento"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Value           =   1  'Checked
               Width           =   2235
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Incidência no Valor Total da Nota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1785
            Left            =   8760
            TabIndex        =   74
            Top             =   900
            Width           =   2895
            Begin VB.CheckBox chk_soma_subs 
               Caption         =   "Soma ICMS Sub"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   80
               Top             =   1440
               Width           =   1695
            End
            Begin VB.CheckBox chk_frete_total 
               Caption         =   "Frete da Nota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Value           =   1  'Checked
               Width           =   1635
            End
            Begin VB.CheckBox chk_outras_total 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   480
               Value           =   1  'Checked
               Width           =   2115
            End
            Begin VB.CheckBox chk_ipi_total 
               Caption         =   "I.P.I."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   720
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox chk_desconto_total 
               Caption         =   "Desconto"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   76
               Top             =   960
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox chk_seguro_total 
               Caption         =   "Seguro"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   75
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkImportacao 
            Caption         =   "Nota de Importação"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2190
            TabIndex        =   73
            Top             =   165
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Atualiza Custo Por:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   90
            TabIndex        =   107
            Top             =   165
            Width           =   1560
         End
      End
      Begin VB.Frame frm_dados1 
         Height          =   7695
         Left            =   -74970
         TabIndex        =   51
         Top             =   310
         Width           =   12195
         Begin VB.TextBox txt_transportadora_cnpj 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   1080
            TabIndex        =   67
            Top             =   1680
            Width           =   2205
         End
         Begin VB.TextBox txt_transportadora_uf 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   120
            TabIndex        =   66
            Top             =   1680
            Width           =   885
         End
         Begin VB.TextBox txt_transportadora_placa_uf 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   7455
            TabIndex        =   65
            Top             =   1680
            Width           =   1245
         End
         Begin VB.TextBox txt_transportadora_placa 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   5475
            TabIndex        =   64
            Top             =   1680
            Width           =   1845
         End
         Begin VB.TextBox txt_transportadora_inscricao_estadual 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   63
            Text            =   "12345678901234567890"
            Top             =   1680
            Width           =   1980
         End
         Begin VB.TextBox txt_transportadora_endereco 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   120
            TabIndex        =   62
            Top             =   1050
            Width           =   4770
         End
         Begin VB.TextBox txt_transportadora_bairro 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   5025
            TabIndex        =   61
            Top             =   1050
            Width           =   3690
         End
         Begin VB.TextBox txt_transportadora_cidade 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   8835
            TabIndex        =   60
            Top             =   1035
            Width           =   2865
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   2
            Left            =   1440
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":2710F
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   59
            Top             =   525
            Width           =   255
         End
         Begin VB.TextBox txt_codigo_transportadora 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   1245
         End
         Begin VB.TextBox txt_transportadora_nome 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   1710
            TabIndex        =   57
            Top             =   480
            Width           =   10005
         End
         Begin VB.ComboBox cbo_frete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8835
            TabIndex        =   56
            Text            =   "cbo_uf"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txt_peso_bruto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   3360
            TabIndex        =   55
            Top             =   2280
            Width           =   2025
         End
         Begin VB.TextBox txt_peso_liquido 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   5520
            TabIndex        =   54
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox txt_especie 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   2040
            TabIndex        =   53
            Top             =   2280
            Width           =   1245
         End
         Begin VB.TextBox txt_volume 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   120
            TabIndex        =   52
            Top             =   2280
            Width           =   1845
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código                      Nome da Transportadora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   3630
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27451
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   1440
            Width           =   9705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27500
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   11340
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Volume                            Espécie               Peso Bruto                            Peso Líquido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   2040
            Width           =   6480
         End
      End
      Begin VB.Frame frmdados 
         Height          =   7575
         Left            =   120
         TabIndex        =   6
         Top             =   375
         Width           =   12015
         Begin VB.CheckBox chk_ImportarProdCodigoBarras 
            Caption         =   "Importar produtos do XML pelo código de barras."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9720
            TabIndex        =   154
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txt_PctIPI 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            TabIndex        =   153
            Text            =   "txt_IPI"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txt_PctICMS 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9720
            TabIndex        =   152
            Text            =   "txt_PctICMS"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txt_redicmssubst 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10635
            TabIndex        =   147
            Text            =   " "
            Top             =   2955
            Width           =   1245
         End
         Begin VB.CommandButton cmd_ler_xml 
            Caption         =   "Preencher dados"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   10440
            TabIndex        =   144
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox txt_caminho_xml 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   145
            Top             =   250
            Width           =   5775
         End
         Begin VB.CommandButton cmd_explorer 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1365
            TabIndex        =   143
            Top             =   250
            Width           =   500
         End
         Begin VB.TextBox txt_fornecedor 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2820
            TabIndex        =   25
            Text            =   " "
            Top             =   765
            Width           =   4860
         End
         Begin VB.TextBox txt_chave_acesso 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8400
            MaxLength       =   44
            TabIndex        =   142
            Top             =   1095
            Width           =   3495
         End
         Begin VB.TextBox txt_porc_red_icms 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   10635
            TabIndex        =   108
            Top             =   2550
            Width           =   1245
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   5295
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":275EB
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   27
            Top             =   1485
            Width           =   225
         End
         Begin VB.TextBox txt_codigo_fornecedor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1365
            TabIndex        =   26
            Text            =   " "
            Top             =   750
            Width           =   1125
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   2550
            Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":2792D
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   24
            Top             =   795
            Width           =   255
         End
         Begin VB.TextBox txt_numero_nf 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1365
            TabIndex        =   23
            Text            =   " "
            Top             =   1095
            Width           =   1125
         End
         Begin VB.TextBox txt_serie_nf 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3075
            MaxLength       =   3
            TabIndex        =   22
            Text            =   "999"
            Top             =   1095
            Width           =   405
         End
         Begin VB.TextBox txt_modelo_nf 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4395
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "99"
            Top             =   1095
            Width           =   405
         End
         Begin VB.TextBox txt_total 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10635
            TabIndex        =   20
            Text            =   " "
            Top             =   1455
            Width           =   1245
         End
         Begin VB.TextBox txt_bc_icms 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1365
            TabIndex        =   19
            Text            =   " "
            Top             =   1815
            Width           =   1125
         End
         Begin VB.TextBox txt_icms 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4155
            TabIndex        =   18
            Text            =   " "
            Top             =   1815
            Width           =   1125
         End
         Begin VB.TextBox txt_bc_substituicao 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7635
            TabIndex        =   17
            Top             =   1815
            Width           =   1245
         End
         Begin VB.TextBox txt_substituicao 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10635
            TabIndex        =   16
            Text            =   " "
            Top             =   1815
            Width           =   1245
         End
         Begin VB.TextBox txt_desconto 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4155
            TabIndex        =   15
            Text            =   " "
            Top             =   2535
            Width           =   1125
         End
         Begin VB.TextBox txt_outras 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1365
            TabIndex        =   14
            Text            =   " "
            Top             =   2535
            Width           =   1125
         End
         Begin VB.TextBox txt_seguro 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10635
            TabIndex        =   13
            Text            =   " "
            Top             =   2175
            Width           =   1245
         End
         Begin VB.TextBox txt_frete 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4155
            TabIndex        =   12
            Text            =   " "
            Top             =   2175
            Width           =   1125
         End
         Begin VB.TextBox txt_ipi 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1365
            TabIndex        =   11
            Text            =   " "
            Top             =   2175
            Width           =   1125
         End
         Begin VB.TextBox txt_frete_conhecimento 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7635
            TabIndex        =   10
            Text            =   " "
            Top             =   2175
            Width           =   1245
         End
         Begin VB.TextBox txt_observacoes 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1365
            TabIndex        =   9
            Text            =   " "
            Top             =   2955
            Width           =   7530
         End
         Begin VB.TextBox txt_codigo_forma 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4155
            MaxLength       =   4
            TabIndex        =   8
            Top             =   1455
            Width           =   1125
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   4560
            Top             =   4320
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   1
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox msk_entrada 
            Height          =   330
            Left            =   1365
            TabIndex        =   28
            Top             =   1455
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_emissao 
            Height          =   330
            Left            =   5760
            TabIndex        =   29
            Top             =   1095
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_chegada 
            Height          =   330
            Left            =   7635
            TabIndex        =   30
            Top             =   2535
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Red. ICMS Subst."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   9000
            TabIndex        =   148
            Top             =   3045
            Width           =   1575
         End
         Begin VB.Label lbl_caminho_xml_demo 
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27C6F
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   315
            Width           =   10335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27D03
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   6840
            Width           =   11160
         End
         Begin VB.Label lbl_total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10680
            TabIndex        =   49
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_ipi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9480
            TabIndex        =   48
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_outras_despesas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8340
            TabIndex        =   47
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_seguro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7185
            TabIndex        =   46
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_frete 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6030
            TabIndex        =   45
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_total_produtos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4860
            TabIndex        =   44
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_icms_substituicao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3675
            TabIndex        =   43
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_bc_substituicao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2490
            TabIndex        =   42
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_valor_icms 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1305
            TabIndex        =   41
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label lbl_bc_icms 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   40
            Top             =   7080
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27D99
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   39
            Top             =   810
            Width           =   7830
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27E3A
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   38
            Top             =   1155
            Width           =   8265
         End
         Begin VB.Label lbl_cfop 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8400
            TabIndex        =   37
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27EC5
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   15
            TabIndex        =   35
            Top             =   1875
            Width           =   10530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":27F77
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   30
            TabIndex        =   34
            Top             =   2235
            Width           =   10500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":2802C
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   45
            TabIndex        =   33
            Top             =   2595
            Width           =   10455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   225
            TabIndex        =   32
            Top             =   3015
            Width           =   1095
         End
         Begin VB.Label lbl_forma 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5535
            TabIndex        =   31
            Top             =   1455
            Width           =   3345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   $"Movimento_Nota_Fiscal_Entrada.frx":280E6
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   75
            TabIndex        =   36
            Top             =   1515
            Width           =   10485
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   1005
      ButtonWidth     =   1429
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            ImageIndex      =   2
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alterar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesquisar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Frete"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imp.NF"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ajuda"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "V.Custos"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Canc.NF"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pd.Transf"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox pct_icms 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8400
         Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":28199
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   45
         Width           =   255
      End
      Begin VB.PictureBox pct_orçamento 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8415
         Picture         =   "Movimento_Nota_Fiscal_Entrada.frx":284DB
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   45
         Width           =   255
      End
      Begin Threed.SSCommand cmd_confirma 
         Height          =   330
         Left            =   10440
         TabIndex        =   1
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   262144
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Confirma Produto"
         ButtonStyle     =   2
         Outline         =   0   'False
      End
   End
End
Attribute VB_Name = "Movimento_Nota_Fiscal_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSenha                    As String
Private strLocacao                  As String
Private lngSequenciaControle        As Long
'Private xcodigoforma As String
Private xid_deposito                As String
'Private lobservacao As String
'Private xUltimoRow As Byte
'Private xLanca_Grade As Boolean
'Private lContrapartida As Byte
Private lSequencia                  As Long
Private industria_revenda           As Byte
Private x_naotributados             As Currency
Private xtotalnotaprodutos          As Currency
Private xCodigoProduto              As String
Private xSomaRateioIsencao          As Boolean
Private SomaPercentual              As Currency
Private lTipoCaixa                  As String
Private lDataCaixa                  As Date
Private lbl_serie_nf                As String
Private g_reducao_invertido         As Byte
Private x_substituicaoant           As Currency
Private x_perc_icms_subst           As Currency
Private x_base_calculo_conhecimento As Currency
Private x_calculo                   As Currency
'Private x_inf_estoque  As String
Private xplanoconta As String
Private xvaloricms_conhecimento     As Currency
Private x_aliquota_conhecimento     As Currency
'Private x_influencia_estoque As Boolean
'Private x_digitacao_grade As Boolean
Private lbl_inscricao               As String
Private lbl_endereco                As String
Private lbl_cidade                  As String
Private lbl_uf                      As String
Private lbl_pais                    As String
Private lbl_pessoa                  As Byte
Private lContribuinte               As String
Private lbl_cgc                     As String
'Private Imp_Aliquota As String
'Private Imp_lei As String
Private x_preco_varejo              As String
Private x_preco_atacado             As String
Private x_desc_atacado              As String
Private x_desc_varejo               As String
Private lporcentagemreducao         As Currency
Private lporcredicmssubst           As Currency
'Private lDentroForaPais As String
Private FimdeProdutos               As Integer
Private lNomeCodificacaoFiscal      As String
'Private ValidaLei As Boolean
'Private ValidaLei_1 As Boolean
'Private lreducaobasecalculo As Byte
'Private lQuantidadeTotalProdutos As Currency
Private lPagina                     As Integer
Private lBaseCalculoIcms_red        As Currency
Private lValorIcms_red              As Currency
Private lLancamento                 As Long
Private xbasecalculo                As String
Private xValorIcms                  As String
'Private x_bc_icms As Currency
'Private x_valor_icms As Currency
Private xbcsubstituicao             As Currency
Private xValorSubstituicao          As Currency
Private xValorSubstituicaocusto     As Currency
'Private z, i                        As Long
Private zGrupo                      As Long
Private zProduto                    As String
'Private zCodigoFornecedor As String
Private zvalorbruto                 As String
Private zporc_desconto              As String
Private zUnidade                    As String
Private zNome                       As String
Private zQuantidade                 As String
Private zValorUnitario              As String
Private zCustoAnterior              As String
Private zValorTotal                 As String
Private zPorcentagemIPI             As String
Private zValorIPI                   As String
Private zValorICMS                  As String
Private lbl_cep                     As String
Private lbl_telefone                As String
Private lbl_bairro                  As String
Private lContaContabilVista         As Long
Private lcontareduzida              As Long
Private l_opcao                     As Byte
Private lPessoa                     As String
Private lCGCFornecedor              As String
'Private lSubstituto As Integer
Private lLinha                      As Integer
Private lCol                        As Byte
Private lRow                        As Integer
Private lgrupo                      As Long
Private lUnidade                    As String
'PRIVATE Qtd_Aliquotas As Integer
Private lcodigoaliquota             As Integer
'PRIVATE lAliquota As Currency
'Private lCodificacaoFiscal(10) As Long
Private lBaseCalculoIcms(10)        As Currency
Private lValorIcms(10)              As Currency
Private lBaseCalculoSubstituicao    As Currency
Private lValorSubstituicao          As Currency
Private lValorIPI                   As Currency
Private lDesconto(10)               As Currency
Private lFrete(10)                  As Currency
Private lSeguro(10)                 As Currency
Private lOutras(10)                 As Currency
Private lFreteConhecimento(10)      As Currency
Private lFreteBCConhecimento(10)    As Currency
Private lpis(10)                    As Currency
Private lcofins(10)                 As Currency
Private lTotalDesconto              As Currency
Private lTotalFrete                 As Currency
Private lTotalSeguro                As Currency
Private lTotalOutras                As Currency
Private lTotalProdutos              As Currency
Private ltotalcusto                 As String
Private lTotalIPI                   As Currency
Private lSubTotal                   As Currency
Private lTotal                      As Currency
Private lTotalPIS                   As Currency
Private lTotalCofins                As Currency
Private lTotalII                    As Currency
Private flag_tela_entrada_mercadoria As Byte
'Private lDentroFora As String
Private f                           As Integer
Private col                         As Byte
Private lcustoprodutos(1000)
Private lbl_nota                    As String
Private bolConferencia              As Boolean
Private bolEntradaTMP               As Boolean
Private lngCodigoDocImportacao      As Long
Private lngCodigoContaConsumo       As Long
Private lngIDGrupoDeTensao(20)      As Long
Private lngIDConsumo(99)            As Long
Public blnRealizouTransacao         As Boolean 'lote
Private strMotivoCancelamento       As String
Private blnExisteNota               As Boolean
'Grade produtos
Private lngIDGrade                  As Long
Public bolCentralNFe                As Boolean
Public bolNFePendente               As Boolean
Private bolPedidoTransf             As Boolean 'informa se e uma nota originado de pedido de transferencia

'---------declarações do sistema-------------------------------------
Const NovaLinha                     As String = ">*"      ' Indica uma nova linha
Private LastRow                     As Long               ' Ultima linha em que se editou
Private LastCol                     As Long               ' ultima coluna em que se editou
Private xdigitado                   As Byte
Private MovimentoCaixaGeral         As New SabreRG.cMovimentoCaixaGeral
Private ContasAPagar                As New SabreRG.cContasAPagar
Private MovCabNotaFiscalEntrada     As New SabreRG.cMovCabNotaFiscalEntrada
Private MovNotaFiscalEntrada        As New SabreRG.cMovNotaFiscalEntrada
Private MovimentoCaixa              As New SabreRG.cMovimentoCaixa
Private cMapLote                    As New cMapLotes
Private cDPEmpresa                  As New cDPEmpresas
Private cDPNotasSaida               As New cDPNotasSaidas
Private cCtrlEntradaSaida           As New AutCont.cCtrlEntradaSaida
Private colTributacao               As New Collection   'coleção do cadastro de propriedade tributação
Private colFormaFaturamento         As New Collection
Private cDPFFaturamento             As New cDPFormaFaturamento 'objeto cadastro forma de faturamento
Private cUtGeral                    As New Utilitarios.cUtlGeral
Private cDPParamSistema             As New cDPParametrosSistema
Private strIDLote                   As String '*** Armazena Lote da NFe para efetuar cancelamento ***
'Para validação alteração da nota 'RONALDO
Private colProdutos                 As New Collection 'coleção de produtos
Private bolAlteraEstoquenaAlteracaoNF As Boolean 'informa se vai alterar os estoques na alteração da nota
Private bolLiberaAlteracaoCustoAlteracaoNF As Boolean 'informa se pode também atualizar o custo na alteracao da nota
Private bolLiberaAlteracaoFinanceiro    As Boolean  'informa se pode efetuar alteração no financeiro
Private bolExisteLancFinanceiro         As Boolean  'identifica que houve lançamento no financeiro
Private bolExisteLancCaixa              As Boolean  'identifica que houve lançamento no caixa
Private lngCodFormaFatuAnterior         As Long 'armazena o cod forma faturamento ao carregar tela


'*****************************************************************************
'Criação: Ronaldo Robledo                                     Data: 03/09/2010
'Propósito:
'Alteração: João Batista                                      Data: 24/11/2012
'           inserido chk_complementar.Value = 0 no metodo limpatela correção ticket TT1918.
'*****************************************************************************
Private Sub LimpaTela()
On Error Resume Next

    grade1.Rows = 0
    If grade1.Rows = 0 Then
        grade1.Rows = 2
        grade1.FixedRows = 1
    End If
        
    With grade1
        .TextMatrix(.Rows - 1, 0) = NovaLinha
         LastRow = .Rows - 1
         LastCol = 1
        .col = LastCol
        .Row = LastRow
        .RowSel = LastRow
        .ColSel = LastCol
    End With
    
    ZeraGrade

    FormaGrid
    cbo_atualizacusto.Text = cbo_atualizacusto.List(1)
    chk_atualiza_carteira.Value = 1
    chk_atualiza_caixa.Value = 0
    cDPFFaturamento.Contrapartida = 0
    lSequencia = 0
    'xcodigoforma = ""
    txt_serie_nf.Enabled = True
    Call LimpaTelaFornecedor
    txt_codigo_fornecedor = "0"
    txt_codigo_forma = ""
    lngCodFormaFatuAnterior = 0
    lbl_forma = ""
    lbl_cgc = ""
    lbl_inscricao = ""
    lbl_endereco = ""
    lbl_cidade = ""
    lbl_uf = ""
    lbl_pais = ""
    chk_impressao_nf.Value = 0
    chk_retido.Value = 0
    chk_complementar.Value = 0
    chk_calculo_nota.Value = 0
    txt_numero_nf = ""
    txt_serie_nf = ""
    txt_modelo_nf = ""
    msk_emissao = "__/__/____"
    msk_entrada = "__/__/____"
    msk_chegada = "__/__/____"
    'x_inf_estoque = "S"
    
    lbl_cfop = ""
    txt_bc_icms = Format(0, "##,###,##0.00")
    txt_ipi = Format(0, "##,###,##0.00")
    txt_icms = Format(0, "##,###,##0.00")
    txt_bc_substituicao = Format(0, "##,###,##0.00")
    txt_substituicao = Format(0, "##,###,##0.00")
    txt_frete = Format(0, "##,###,##0.00")
    txt_seguro = Format(0, "##,###,##0.00")
    txt_outras = Format(0, "##,###,##0.00")
    txt_desconto = Format(0, "##,###,##0.00")
    txt_total = Format(0, "##,###,##0.00")
    txt_frete_conhecimento = Format(0, "##,###,##0.00")
    txt_observacoes = ""
    
    txt_PctIPI = Format(fValidaValorNovo(0), "##,###,##0.00")
    txt_PctICMS = Format(fValidaValorNovo(0), "##,###,##0.00")
    x_aliquota_conhecimento = 0
    x_base_calculo_conhecimento = 0
    lbl_bc_icms = Format(0, "##,###,##0.00")
    lbl_valor_icms = Format(0, "##,###,##0.00")
    lbl_bc_substituicao = Format(0, "##,###,##0.00")
    lbl_icms_substituicao = Format(0, "##,###,##0.00")
    lbl_total_produtos = Format(0, "##,###,##0.00")
    lbl_frete = Format(0, "##,###,##0.00")
    lbl_seguro = Format(0, "##,###,##0.00")
    lbl_outras_despesas = Format(0, "##,###,##0.00")
    lbl_ipi = Format(0, "##,###,##0.00")
    lbl_total = Format(0, "##,###,##0.00")
    txt_totalqtde = Format(0, "##,###,##0.00")
    
    txt_codigo_transportadora = 0
    txt_transportadora_nome = ""
    txt_transportadora_endereco = ""
    txt_transportadora_bairro = ""
    txt_transportadora_cidade = ""
    txt_transportadora_uf = ""
    txt_transportadora_cnpj = ""
    txt_transportadora_inscricao_estadual = ""
    txt_transportadora_placa = ""
    txt_transportadora_placa_uf = ""
    cbo_frete.Text = cbo_frete.List(0)
    x_perc_icms_subst = 0
    xCodigoProduto = ""
    xid_deposito = 0
    txt_especie = ""
    txt_volume = 0
    txt_peso_bruto = 0
    txt_peso_liquido = 0
    lngSequenciaControle = 0
    lngIDGrade = 0
    lLancamento = 0
    lDataCaixa = "00:00:00"
    lTipoCaixa = ""
    bolExisteLancFinanceiro = False
    bolExisteLancCaixa = False
    bolLiberaAlteracaoFinanceiro = False
    
    grade3.Clear
    FormaGridConf
    LimpaConferencia
    txt_codigo_usuario = ""
    txt_nome_usuario = ""
    txt_senha = ""

    lLinha = 1
    lCol = 0
    lRow = 0
    
    lgrupo = 0
    lUnidade = ""
'    lobservacao = ""
    txt_porc_red_icms = 0
    If g_entrada_amarrado_pedido = 0 Or g_entrada_amarrado_pedido = 1 Then
       bolConferencia = False
       bolEntradaTMP = False
       SSTab1.TabVisible(3) = False
    ElseIf g_entrada_amarrado_pedido = 2 Or g_entrada_amarrado_pedido = 3 Then
           bolConferencia = True
           bolEntradaTMP = True
            SSTab1.TabVisible(3) = True
    End If
    
    lngCodigoDocImportacao = 0
    
    '*** Implementações SPED ***
    chk_ContaConsumo.Value = 0
    cbo_TipoContaConsumo.Text = cbo_TipoContaConsumo.List(0)
    cbo_TipodeLigacao.Text = cbo_TipodeLigacao.List(0)
    cbo_GrupoDeTensao.Text = cbo_GrupoDeTensao.List(0)
    HabilitaControlesConsumo (False)
    lngCodigoContaConsumo = 0
    txt_caminho_xml.Text = ""
    txt_chave_acesso.Text = ""
    chk_ImportarProdCodigoBarras.Value = 0
    
End Sub

Private Sub HabilitaControlesConsumo(bolValor As Boolean)

    frm_LancamentoConsumo.Enabled = bolValor
    frm_LancamentoConsumo.Visible = bolValor
    cbo_CodigoConsumo.Enabled = bolValor
    cbo_CodigoConsumo.Visible = bolValor
    
End Sub

Private Sub AtivaBotoes()
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(11).Enabled = True
    Toolbar1.Buttons(12).Enabled = True
    Toolbar1.Buttons(13).Enabled = False
    frm_dados1.Enabled = False
    frm_dados2.Enabled = False
    frmdados.Enabled = False
    cmd_confirma.Visible = False
End Sub

Private Sub DesativaBotoes()
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(11).Enabled = False
    Toolbar1.Buttons(12).Enabled = False
    Toolbar1.Buttons(13).Enabled = True
    frmdados.Enabled = True
    frm_dados1.Enabled = True
    frm_dados2.Enabled = True
    frm_conf.Enabled = False
    cmd_confirma.Visible = False
End Sub

'*****************************************************************************
'Alteração: Fernando Silva                                          15/02/2011
'           Inserida a SUB AtualTelaFornecedor para carregar as informações do
'           fornecedor corretamente quando vem da tela de Conferência- O. S.:15631
'*****************************************************************************
Private Sub AtualTela()
On Error GoTo AtualTela

    blnExisteNota = True
    
    With gb_Recordset
        
        l_opcao = 2
        lSequencia = !sequencia
        lngSequenciaControle = IIf(IsNull(!seq_controle), 0, !seq_controle)
        txt_codigo_fornecedor = !codigo_do_fornecedor
        txt_numero_nf = !Numero
        txt_serie_nf = !Serie
        txt_modelo_nf = !modelo_nf
        msk_emissao = !data_de_emissao
        txt_chave_acesso = !chave_acesso_nfe
        msk_entrada = !data_de_entrada
        lbl_cfop = !codificacao_fiscal
        txt_observacoes = !Observacoes
        txt_codigo_forma = !tipo_do_documento
        lngCodFormaFatuAnterior = !tipo_do_documento
        txt_bc_icms = Format(!base_calculo_do_icms, "##,###,##0.00")
        txt_ipi = Format(!Valor_IPI, "##,###,##0.00")
        txt_icms = Format(!valor_do_icms, "##,###,##0.00")
        txt_bc_substituicao = Format(!base_calculo_substituicao_icms, "##,###,##0.00")
        txt_substituicao = Format(!valor_icms_substituicao, "##,###,##0.00")
        txt_frete = Format(!Frete, "##,###,##0.00")
        txt_frete_conhecimento = Format(!frete_conhecimento, "##,###,##0.00")
        txt_seguro = Format(!Seguro, "##,###,##0.00")
        txt_outras = Format(!outras_despesas, "##,###,##0.00")
        txt_desconto = Format(!valor_do_desconto, "##,###,##0.00")
        txt_total = Format(!total_da_nota, "##,###,##0.00")
        
        lbl_bc_icms = Format(!base_calculo_do_icms, "##,###,##0.00")
        lbl_ipi = Format(!Valor_IPI, "##,###,##0.00")
        lbl_valor_icms = Format(!valor_do_icms, "##,###,##0.00")
        lbl_bc_substituicao = Format(!base_calculo_substituicao_icms, "##,###,##0.00")
        lbl_total_produtos = Format(!valor_total_produtos, "##,###,##0.00")
        lbl_icms_substituicao = Format(!valor_icms_substituicao, "##,###,##0.00")
        lbl_frete = Format(!Frete, "##,###,##0.00")
        lbl_seguro = Format(!Seguro, "##,###,##0.00")
        lbl_outras_despesas = Format(!outras_despesas, "##,###,##0.00")
        lbl_total = Format(!total_da_nota, "##,###,##0.00")
        lLancamento = !numero_movimento_caixa
        
        chk_impressao_nf.Value = !impresso_nf
        cbo_atualizacusto.Text = cbo_atualizacusto.List(!atualiza_custo)
        chk_retido.Value = !imposto_retido
        chk_calculo_nota.Value = !calculo_nota
        chk_atualiza_carteira.Value = !atualiza_carteira
        chk_atualiza_caixa.Value = !atualiza_caixa
        lDataCaixa = IIf(IsNull(!Data_Caixa), Date, !Data_Caixa)
        lTipoCaixa = !Tipo_Caixa
        chk_complementar.Value = !nf_complementar
        bolExisteLancFinanceiro = IIf((chk_atualiza_carteira.Value = 1), True, False)
        bolExisteLancCaixa = IIf((chk_atualiza_caixa.Value = 1), True, False)
        
        txt_volume = Format(!Quantidade, "##,###,##0.000")
        txt_especie = !Especie
        txt_peso_bruto = Format(IIf((IsNull(!peso_bruto) Or Trim(!peso_bruto) = ""), 0, !peso_bruto), "##,###,##0.000")
        txt_peso_liquido = Format(IIf((IsNull(!peso_liquido) Or Trim(!peso_liquido) = ""), 0, !peso_liquido), "##,###,##0.000")
        
        txt_codigo_transportadora = !codigo_transportadora
        txt_transportadora_placa = !Placa
        txt_transportadora_placa_uf = !uf_placa
            
        If cDPEmpresa.NotaFiscalEletronica = 1 And chk_impressao_nf.Value = 1 Then
           If Not IsNull(!id_nfe) Then
                BD_Record_SetII.Source = "SELECT MN.num_nfe FROM movimento_nfe MN WHERE MN.empresa = " & !Empresa & " and MN.id_nfe = '" & !id_nfe & "'"
                BD_Record_SetII.Open
                If BD_Record_SetII.RecordCount > 0 Then
                   strNumNFe = BD_Record_SetII!num_nfe
                Else
                   strNumNFe = ""
                End If
                BD_Record_SetII.Close
            End If
        End If
           
        BD_Record_SetII.Source = "SELECT * FROM transportadora WHERE codigo = '" & txt_codigo_transportadora & "'"
        BD_Record_SetII.Open
        If BD_Record_SetII.RecordCount > 0 Then
            txt_transportadora_nome = BD_Record_SetII!NOME
            txt_transportadora_endereco = BD_Record_SetII!Endereco
            txt_transportadora_bairro = BD_Record_SetII!Bairro
            txt_transportadora_cidade = BD_Record_SetII!Cidade
            txt_transportadora_uf = BD_Record_SetII!UF
            txt_transportadora_cnpj = BD_Record_SetII!CGC
            txt_transportadora_inscricao_estadual = BD_Record_SetII!inscricao_estadual
        End If
        BD_Record_SetII.Close
        cbo_frete.Text = cbo_frete.List(!tipo_frete - 1)
        
        chk_ContaConsumo.Value = !conta_consumo
        If chk_ContaConsumo.Value = 1 Then Call PreencheCamposContaConsumo
        
    End With
    gb_Recordset.Close
        
    Set gb_Recordset = Conexao.GeraRecordset("SELECT descricao FROM forma_faturamento WHERE codigo = " & txt_codigo_forma, 1)
    If gb_Recordset.RecordCount > 0 Then
        lbl_forma = Mid(gb_Recordset!Descricao, 1, 30)
    End If
    gb_Recordset.Close
    
    '****Alterado para preencher corretamente os dados do fornecedor
    'quando vem da tela de conferência - Fernando Silva - O. S.:15631****
    Call BuscaFornecedor(txt_codigo_fornecedor)
    
    Call AtualTelaConferencia

Exit Sub
AtualTela: If Err.Number = 94 Then Resume Next Else ValidaErros Err, Me.Caption & " - AtualTela"
End Sub

Private Sub PreencheCamposContaConsumo()
On Error GoTo Err_PreencheCamposContaConsumo
Dim intIndex As Integer
    
    lngCodigoContaConsumo = gb_Recordset!PkCodigoConsumo
    
    '*** Preenche qual tipo de conta Agua/Luz/Gás ***
    Select Case gb_Recordset!TipoConta
           Case 1
                intIndex = 0
           Case 2
                intIndex = 1
           Case 3
                intIndex = 2
    End Select
    cbo_TipoContaConsumo.Text = cbo_TipoContaConsumo.List(intIndex)
    
    '*** Tipo de ligação ***
    Select Case gb_Recordset!TipoLigacao
           Case 1
                intIndex = 0
           Case 2
                intIndex = 1
           Case 3
                intIndex = 2
    End Select
    cbo_TipodeLigacao.Text = cbo_TipodeLigacao.List(intIndex)
            
    cbo_CodigoConsumo.Text = cbo_CodigoConsumo.List(lngIDConsumo(gb_Recordset!CodigoConsumo))
    cbo_GrupoDeTensao.Text = cbo_GrupoDeTensao.List(lngIDGrupoDeTensao(gb_Recordset!GrupoTensao))
    
Exit Sub
Err_PreencheCamposContaConsumo: If Err.Number = 13 Or Err.Number = 94 Then Resume Next Else ValidaErros Err, Me.Caption & " - PreencheCamposContaConsumo"
End Sub

'*****************************************************************************
'Criação: Ronaldo Robledo Mendes Souza                        Data: 09/03/2011
'Propósito: Calcular Nota Fiscal
'Alteração: Ronaldo Robledo                                 data: 09/03/2011
'           Retirado a variavel lIva e inserido o campo na grade grade1.TextMatrix(f, 42)
'           para o usuário poder alterar o percentual do iva]
'Alteração  Ronaldo Robledo                                 data: 31/05/2011
'           Inserido a função BuscaIVAporEstadoporProduto para atender a nova legislação
'           substituição tributaria por estado e produto
'
'Alteração: Nayden Luiz dos Santos Cruz                    Data:19/09/2011
'         : OS 17506
'         : Foi adicionado uma validação nos Grid onde caso esteja prenchido
'         : valor de ICMS no grid o mesmo ira verificar se foi preenchido a
'         : base de calculo e vice-versa porque antes estava passando o valor
'         : de Icms ou base de calculo com valor 0 caso o usuario nao preenchesse
'
'Alteração: Nayden Luiz dos Santos Cruz                    Data:31/10/2011
'         : OS17821-Corrigida validação se foi digitado base de calculo icms
'         : e se algum produto da nota tem icms senao tiver ele nao da a entrada da
'         : nota e informa ao usuario.
'
'Alteração: Diego Martins                                  Data:08/11/2011
'         : OS17890-Alterado o tipo da variavél dblFator antes era long mudado para double
'Alteração: Ronaldo Robledo                                     01/08/2012
'           Inserido ROUND os valores totais para trabalhar o arrendondamento correto
'*****************************************************************************
Private Function CalculaNotaFiscal() As Boolean
On Error GoTo CalculaNotaFiscal
'Dim lcodigoforma As String
'Dim xservicoproduto As String
Dim dblFator As Double
Dim blnIcms As Boolean

    g_string4 = 0
    dblFator = 1
    CalculaNotaFiscal = False

    If chk_lancamento_venda.Value = 0 And Val(txt_numero_nf) > 0 Then
        Call Conexao.DeleteSintetico("calculo_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and numero_nf = '" & txt_numero_nf & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & "' and outros = '" & txt_serie_nf & "'", cDPEmpresa.codigo)
    End If

    ZeraVariaveisTotais
    blnIcms = False
    
    For f = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(f, 0)) Then
            Call CalculaValorPesoPorUnidadeProduto(grade1.TextMatrix(f, 2))
            
            If grade1.TextMatrix(f, 10) > 0 Then    ' Alteraçao efetuada perante OS 17506
               If txt_bc_icms.Text = "0,00" Then
                   Alerta ("O calculo de base de ICMS não pode ter valor 0,00 se o mesmo tem ICMS")
                   CalculaNotaFiscal = False
                   Exit Function
               End If
            End If
        
           'Alteraçao efetuada perante OS 17506
            If grade1.TextMatrix(f, 10) <> 0 Then blnIcms = True
            
'            Set gb_Recordset = Conexao.GeraRecordset("SELECT aliquota_estadual,industria_revenda,servico_produto " & _
'                                                    "FROM produto WHERE codigo = '" & grade1.TextMatrix(f, 2) & "'", 1)
'            If gb_Recordset.RecordCount > 0 Then
'                lcodigoaliquota = gb_Recordset!aliquota_estadual
'                grade1.TextMatrix(f, 25) = gb_Recordset!industria_revenda
'                xservicoproduto = Mid(gb_Recordset!servico_produto, 1, 1)
'            Else
                'verificar pelo subcodigo se veio produto do cadastro de produto
                If grade1.TextMatrix(f, 46) = 0 Or grade1.TextMatrix(f, 46) = "" Then
                    grade1.TextMatrix(f, 42) = 0
                    lcodigoaliquota = 4
                    grade1.TextMatrix(f, 25) = "R"
                    'xservicoproduto = "P"
                Else
                    lcodigoaliquota = grade1.TextMatrix(f, 16)
                    'xservicoproduto = "P"
                End If
'            End If
'            gb_Recordset.Close
            
            
            'If Val(grade1.TextMatrix(f, 42)) = 0 Then
            '    grade1.TextMatrix(f, 42) = BuscaIVAporEstadoporProduto(grade1.TextMatrix(f, 2), lbl_uf, True)
            'End If
            
            lValorIPI = CDbl(grade1.TextMatrix(f, 13))
            zValorICMS = CDbl(grade1.TextMatrix(f, 10))
            lSubTotal = CDbl(grade1.TextMatrix(f, 11))
            xbcsubstituicao = 0
            xValorSubstituicao = 0
            
            'If lbl_uf = g_uf_empresa Then lDentroFora = "D" Else lDentroFora = "F"
            
            'If lbl_pais = "BRASIL" Then lDentroForaPais = "D" Else lDentroForaPais = "F"
            
'            If lcodigoaliquota = 1 Or lcodigoaliquota = 3 Then
'                g_string2 = 1
'            ElseIf lcodigoaliquota = 2 Then
'                g_string2 = 2
'            Else
'                g_string2 = 3
'            End If
'
'            If Trim(grade1.TextMatrix(f, 28)) = "" Then
'                If CLng(txt_codigo_forma) >= 1 And CLng(txt_codigo_forma) <= 2 Then
'                    lcodigoforma = "('1','2')"
'                Else
'                    lcodigoforma = "(" & txt_codigo_forma & ")"
'                End If
'                Set gb_Recordset = Conexao.GeraRecordset("(SELECT codigo,descricao,conta_vista,codigo_csosn,('1') as ordem FROM natureza_operacao WHERE servico_produto = '" & xservicoproduto & "' AND codigo_da_aliquota = '" & g_string2 & "' and estado = '" & lDentroFora & "' and pais = '" & lDentroForaPais & "' and tipo_da_operacao = 'E' and tipo_venda IN " & lcodigoforma & ") UNION " & _
'                                                         "(SELECT codigo,descricao,conta_vista,codigo_csosn,('2') as ordem FROM natureza_operacao WHERE servico_produto = '" & xservicoproduto & "' AND codigo_da_aliquota IN ('1','3') and estado = '" & lDentroFora & "' and pais = '" & lDentroForaPais & "' and tipo_da_operacao = 'E' and tipo_venda IN " & lcodigoforma & ") UNION " & _
'                                                         "(SELECT codigo,descricao,conta_vista,codigo_csosn,('3') as ordem FROM natureza_operacao WHERE servico_produto = '" & xservicoproduto & "' AND estado = '" & lDentroFora & "' and pais = '" & lDentroForaPais & "' and tipo_da_operacao = 'E' and tipo_venda = '" & CLng(txt_codigo_forma) & "') ORDER BY ordem ASC", 1)
'                If gb_Recordset.RecordCount > 0 Then
'                    grade1.TextMatrix(f, 28) = gb_Recordset!Codigo
'                    grade1.TextMatrix(f, 41) = gb_Recordset!codigo_csosn
'                    lbl_cfop = gb_Recordset!Codigo
'                    lNomeCodificacaoFiscal = gb_Recordset!descricao
'                    lContaContabilVista = gb_Recordset!conta_vista
'                    gb_Recordset.Close
'                Else
'                    gb_Recordset.Close
'                    Alerta "Codificação Fiscal não Encontrada Verifique a Aliquota do seu Produto - " & grade1.TextMatrix(f, 2) & "!", vbCritical
'                    Exit Function
'                End If
'            End If
            
            If chk_soma_subs.Value = 1 Then g_string2 = CDbl(x_substituicaoant) Else g_string2 = 0
            
            Call SomaCalculoPercentual

            'Achar valor e percentual por produto
            If CDbl(txt_frete) > 0 Then
                lFrete(lcodigoaliquota) = CDbl(txt_frete) / SomaPercentual * CDbl(lSubTotal)
            End If
        
            If CDbl(txt_frete_conhecimento) > 0 Then
                lFreteConhecimento(lcodigoaliquota) = CDbl(txt_frete_conhecimento) / SomaPercentual * CDbl(lSubTotal)
                lFreteBCConhecimento(lcodigoaliquota) = CDbl(x_base_calculo_conhecimento) / SomaPercentual * CDbl(lSubTotal)
            End If
        
            If CDbl(txt_seguro) > 0 Then
                'lSeguro(lcodigoaliquota) = cdbl(txt_seguro) / cdbl((cdbl(txt_total) - cdbl(txt_ipi)) - cdbl(txt_seguro) - cdbl(txt_frete) - cdbl(txt_outras) - cdbl(g_string2)) * cdbl(lSubTotal)
                lSeguro(lcodigoaliquota) = CDbl(txt_seguro) / SomaPercentual * CDbl(lSubTotal)
            End If
        
            If CDbl(txt_outras) > 0 Then
                lOutras(lcodigoaliquota) = CDbl(txt_outras) / SomaPercentual * CDbl(lSubTotal)
            End If
        
            If CDbl(txt_desconto) > 0 Then
                lDesconto(lcodigoaliquota) = CDbl(txt_desconto) / SomaPercentual * CDbl(lSubTotal)
            End If
        
            If pct_orçamento.Visible = True Then
                grade1.TextMatrix(f, 17) = "CX"
                lbl_nota = "CX"
            ElseIf pct_icms.Visible = True Then
                grade1.TextMatrix(f, 17) = "CI"
                lbl_nota = "CI"
            Else
                grade1.TextMatrix(f, 17) = "NF"
                lbl_nota = "NF"
            End If
            
            If Not ObtemDadosTributacao(f) Then Exit Function
            
            
             If chk_calculo_nota.Value = 0 Then
                Call Somatoria
                xSomaRateioIsencao = False
                If lcodigoaliquota = 1 Or lcodigoaliquota = 3 Then
                     CalculaIsencao
                ElseIf lcodigoaliquota = 2 Then
                     CalculaSubstituicao
                ElseIf lcodigoaliquota > 3 Then
                     CalculaTributada
                End If
            Else
                Call CalculoNota
            End If
            
'            '*** Se é do Super Simples então preenche o CST com o código do CSOSN ***
'            If g_simples_empresa = 1 Then
'               grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 41)
'            End If
            
            If xSomaRateioIsencao = False Then
                grade1.TextMatrix(f, 14) = CDbl(grade1.TextMatrix(f, 14)) + xbasecalculo
                grade1.TextMatrix(f, 15) = CDbl(grade1.TextMatrix(f, 15)) + xValorIcms
            End If
            
            'nota complementar aproveita apenas credito do icms
            If chk_complementar.Value = 1 Then
                grade1.TextMatrix(f, 5) = 0
                grade1.TextMatrix(f, 6) = 0
                grade1.TextMatrix(f, 7) = 0
                grade1.TextMatrix(f, 8) = 0
                grade1.TextMatrix(f, 9) = 0
                grade1.TextMatrix(f, 11) = 0
                grade1.TextMatrix(f, 14) = 0
                lSubTotal = 0
                txt_total = 0
                cbo_atualizacusto.Text = cbo_atualizacusto.List(0)
            End If
            
            grade1.TextMatrix(f, 16) = lcodigoaliquota
            grade1.TextMatrix(f, 19) = xbcsubstituicao
            grade1.TextMatrix(f, 20) = xValorSubstituicao
            grade1.TextMatrix(f, 27) = lDesconto(lcodigoaliquota)
            
            grade1.TextMatrix(f, 51) = lSubTotal    'base calculo pis
            grade1.TextMatrix(f, 52) = lSubTotal    'base calculo cofins
            lpis(lcodigoaliquota) = (lSubTotal * CDbl(grade1.TextMatrix(f, 21)) / 100)
            lcofins(lcodigoaliquota) = (lSubTotal * CDbl(grade1.TextMatrix(f, 22)) / 100)
            
            grade1.TextMatrix(f, 31) = (lSubTotal * CDbl(grade1.TextMatrix(f, 21)) / 100)   'valor do pis
            grade1.TextMatrix(f, 32) = (lSubTotal * CDbl(grade1.TextMatrix(f, 22)) / 100)   'valor do cofins
                                                            
            lTotalIPI = lTotalIPI + lValorIPI
            lTotalDesconto = lTotalDesconto + lDesconto(lcodigoaliquota)
            lTotalFrete = lTotalFrete + lFrete(lcodigoaliquota)
            lTotalSeguro = lTotalSeguro + lSeguro(lcodigoaliquota)
            lTotalOutras = lTotalOutras + lOutras(lcodigoaliquota)
            lTotalProdutos = lTotalProdutos + lSubTotal
            lTotalPIS = lTotalPIS + grade1.TextMatrix(f, 31)
            lTotalCofins = lTotalCofins + grade1.TextMatrix(f, 32)
            lTotalII = lTotalII + grade1.TextMatrix(f, 36)
            
            Call CalculoOutras
             
            'If grade1.TextMatrix(f, 2) = "164" Then Stop
             
            'SE TEM QUANTIDADE ENTRA E FAZ O CALCULO
            grade1.TextMatrix(f, 18) = 0
            If CCur(grade1.TextMatrix(f, 5)) > 0 Then
                'PREENCHE A VARIAVEL COM O VALOR DO FATOR DA UNIDADE SECUNDARIA
                If IsNumeric(grade1.TextMatrix(f, 45)) Then dblFator = grade1.TextMatrix(f, 45)
                'SOMATORIA DO CUSTO DO PRODUTO NORMAL
                If g_custo_sem_imposto_entrada = 0 Then
                    lcustoprodutos(f) = Format((lValorIPI + lSubTotal + lOutras(lcodigoaliquota) + lSeguro(lcodigoaliquota) + lFrete(lcodigoaliquota) + lFreteConhecimento(lcodigoaliquota) - lpis(lcodigoaliquota) - lcofins(lcodigoaliquota) + xValorSubstituicaocusto - xValorIcms - xvaloricms_conhecimento - lDesconto(lcodigoaliquota)) / grade1.TextMatrix(f, 5), "##,###,##0.0000") / dblFator
                    ltotalcusto = CDbl(ltotalcusto) + lcustoprodutos(f)
                'SOMATORIA DO CUSTO DO PRODUTO APENAS VALOR DO PRODUTO
                ElseIf g_custo_sem_imposto_entrada = 1 Then
                    lcustoprodutos(f) = Format((lValorIPI + lSubTotal - lDesconto(lcodigoaliquota)) / grade1.TextMatrix(f, 5), "##,###,##0.0000") / dblFator
                    ltotalcusto = CDbl(ltotalcusto) + lcustoprodutos(f)
                'SOMATORIA DO CUSTO DO PRODUTO SEM CREDITO DO ICMS
                Else
                    lcustoprodutos(f) = Format((lValorIPI + lSubTotal + lOutras(lcodigoaliquota) + lSeguro(lcodigoaliquota) + lFrete(lcodigoaliquota) + lFreteConhecimento(lcodigoaliquota) + xValorSubstituicaocusto - lDesconto(lcodigoaliquota)) / grade1.TextMatrix(f, 5), "##,###,##0.0000") / dblFator
                    ltotalcusto = CDbl(ltotalcusto) + lcustoprodutos(f)
                End If
                grade1.TextMatrix(f, 18) = lcustoprodutos(f)
                
                Call AtualizaTabelaTMPVerificacaoCustos(f)
                
            End If
        End If
    Next

    lTotal = lTotal + lTotalProdutos
    
    If chk_ipi_total.Value = 1 Then
        lTotal = lTotal + lTotalIPI
    End If
    
    If chk_frete_total.Value = 1 Then
        lTotal = lTotal + lTotalFrete
    End If
    
    If chk_seguro_total.Value = 1 Then
        lTotal = lTotal + lTotalSeguro
    End If
    
    If chk_outras_total.Value = 1 Then
        lTotal = lTotal + lTotalOutras
    End If
    
    If chk_desconto_total.Value = 1 Then
        lTotal = lTotal - lTotalDesconto
    End If
    
    For f = 1 To 4
        lbl_bc_icms = CDbl(lbl_bc_icms) + lBaseCalculoIcms(f)
        lbl_bc_icms = Format(Round(lbl_bc_icms, 2), "##,###,##0.00")
    
        lbl_valor_icms = CDbl(lbl_valor_icms) + lValorIcms(f)
        lbl_valor_icms = Format(Round(lbl_valor_icms, 2), "##,###,##0.00")
    Next

    lbl_bc_substituicao = Format(Round(lBaseCalculoSubstituicao, 2), "##,###,##0.00")
    lbl_icms_substituicao = Format(Round(lValorSubstituicao, 2), "##,###,##0.00")
    lbl_total_produtos = Format(Round(lTotalProdutos, 2), "##,###,##0.00")
    lbl_frete = Format(Round(lTotalFrete, 2), "##,###,##0.00")
    lbl_outras_despesas = Format(Round(lTotalOutras, 2), "##,###,##0.00")
    lbl_seguro = Format(Round(lTotalSeguro, 2), "##,###,##0.00")
    lbl_ipi = Format(Round(lTotalIPI, 2), "##,###,##0.00")
    lbl_total = Format(Round(lTotal, 2), "##,###,##0.00")
    
    If blnIcms = False And txt_bc_icms.Text <> "0,00" Then ''OS17821 Realizada verificaçao se tem base de calculo e se tem ICMS se nao tiver um dos dois ele cancela a entrada da nota
        Alerta ("O ICMS não pode ter valor 0,00 se o mesmo tem Base de Calculo ICMS")
        CalculaNotaFiscal = False
        Exit Function
    End If
    
    CalculaNotaFiscal = True

Exit Function
CalculaNotaFiscal: ValidaErros Err, Me.Caption & " - CalculaNotaFiscal"
End Function

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 07/06/2013
'==========================================================================
Private Function AtualizaTabelaTMPVerificacaoCustos(ByVal f As Long) As Boolean
On Error GoTo Err_AtualizaTabelaTMPVerificacaoCustos

    If chk_lancamento_venda.Value = 1 And xCodigoProduto = grade1.TextMatrix(f, 2) Or chk_lancamento_venda.Value = 0 Then
        If Trim(grade1.TextMatrix(f, 0)) <> ">*" Then
            Dim strCampos As String
            Dim strValores As String
            
            strCampos = ""
            strCampos = strCampos & "empresa,codigo_usuario,ordem_item,data_emissao,codigo_fornecedor,"
            strCampos = strCampos & "numero_nf,codigo_produto,porc_reduzida,iva,icms,"
            strCampos = strCampos & "porc_substituicao,sub_total,base_calc_icms,valor_icms,"
            strCampos = strCampos & "base_calc_substituicao,valor_icms_subst,valor_ipi,valor_outras,"
            strCampos = strCampos & "valor_seguro,valor_frete,valor_icms_frete,valor_desconto,"
            strCampos = strCampos & "custo_atual,outros,pis,cofins"
            
            strValores = strValores & "'" & cDPEmpresa.codigo & "','" & g_usuario & "','"
            strValores = strValores & grade1.TextMatrix(f, 0) & "'," & FormataData(msk_entrada)
            strValores = strValores & ",'" & txt_codigo_fornecedor & "','" & txt_numero_nf & "','"
            strValores = strValores & grade1.TextMatrix(f, 2) & "'," & fValidaValor2(lporcentagemreducao)
            strValores = strValores & "," & fValidaValor2(grade1.TextMatrix(f, 42)) & "," & fValidaValor2(grade1.TextMatrix(f, 10))
            strValores = strValores & "," & fValidaValor2(x_perc_icms_subst) & "," & fValidaValor2(lSubTotal)
            strValores = strValores & "," & fValidaValor2(grade1.TextMatrix(f, 14)) & ","
            strValores = strValores & fValidaValor2(grade1.TextMatrix(f, 15)) & ","
            strValores = strValores & fValidaValor2(xbcsubstituicao) & ","
            strValores = strValores & fValidaValor2(xValorSubstituicao) & ","
            strValores = strValores & fValidaValor2(grade1.TextMatrix(f, 13)) & ","
            strValores = strValores & fValidaValor2(lOutras(lcodigoaliquota)) & ","
            strValores = strValores & fValidaValor2(lSeguro(lcodigoaliquota)) & ","
            strValores = strValores & fValidaValor2(lFrete(lcodigoaliquota) + lFreteConhecimento(lcodigoaliquota))
            strValores = strValores & "," & fValidaValor2(xvaloricms_conhecimento)
            strValores = strValores & "," & fValidaValor2(lDesconto(lcodigoaliquota)) & ","
            strValores = strValores & fValidaValor2(lcustoprodutos(f)) & ",'" & lbl_nota & "',"
            strValores = strValores & fValidaValor2(lpis(lcodigoaliquota)) & ","
            strValores = strValores & fValidaValor2(lcofins(lcodigoaliquota))
            
            Call Conexao.InserirRecordset("calculo_entrada_tmp", strCampos, strValores, cDPEmpresa.codigo)
        End If
    End If

Exit Function
Err_AtualizaTabelaTMPVerificacaoCustos: ValidaErros Err, Me.Caption & " - AtualizaTabelaTMPVerificacaoCustos"
End Function

'*****************************************************************************
'Criação:                                                     Data: 19/08/2011
'Propósito:
'Alteração: Paulo Senhorini                                   Data: 19/08/2011
'           Adicionei chamada a função preenchecolunacst.
'*****************************************************************************
Private Sub CalculaTributada()
On Error GoTo Err_CalculaTributada

    If UCase(grade1.TextMatrix(f, 12)) <> "S" Then
        lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + xbasecalculo
        lValorIcms(4) = lValorIcms(4) + ((xbasecalculo * CDbl(zValorICMS)) / 100)
        xValorIcms = (xbasecalculo * CDbl(zValorICMS)) / 100
        xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
        
        'cst tributado integralmente
        'Call PreencheColunaCst("00") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "00"
    Else
        lBaseCalculoIcms_red = xbasecalculo
        lValorIcms_red = 0
        If g_reducao_invertido = 0 Then
            lBaseCalculoIcms_red = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
            lValorIcms_red = lBaseCalculoIcms_red * CDbl(zValorICMS) / 100
        
        Else
            x_calculo = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
            lBaseCalculoIcms_red = lBaseCalculoIcms_red - x_calculo
            lValorIcms_red = lBaseCalculoIcms_red * CDbl(zValorICMS) / 100
            
        End If
        
        'cst tributado com reducao
        'Call PreencheColunaCst("20") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "20"
        
        lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + lBaseCalculoIcms_red
        lValorIcms(4) = lValorIcms(4) + lValorIcms_red
        
        xbasecalculo = lBaseCalculoIcms_red
        xValorIcms = lValorIcms_red
        
        xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
    End If

Exit Sub
Err_CalculaTributada: ValidaErros Err, Me.Caption & " - CalculaTributada"
End Sub

'*****************************************************************************
'Criação:                                                  Data: 19/08/2011
'
'Propósito:
'
'Alteração: Paulo Senhorini                                Data: 19/08/2011
'           Adicionei chamada a função preenchecolunacst.
'Alteração: Ronaldo Robledo                                       22/06/2012
'           Acrescentado variavel lporcredicmssubst para ter reducao icms subst diferente do
'           icms normal
'           Alterado código do cálculo da substituição reduzido pois para reduzir a BCSUBST'
'           primeiro se aplica a redução depois se calcula o IVA ticket 618
'Alteração: Ronaldo Robledo                                 Data: 23/10/2012
'           Validaçào para entrar no cálculo da substituicao antes pelo IVA agora se tiver
'           Perc.ICMS Sust. pois o produto sempre pode ter Percentual mais nem sempre IVA
'*****************************************************************************
Private Sub CalculaSubstituicao()
On Error GoTo Err_CalculaSubstituicao
Dim curValorRedSubstituicao As Currency

    If grade1.TextMatrix(f, 12) <> "S" Then
        If zValorICMS = 0 And xbasecalculo > 0 Then
            Call LancaCalculo
        Else
            lBaseCalculoIcms(2) = lBaseCalculoIcms(2) + xbasecalculo
            lValorIcms(2) = lValorIcms(2) + ((xbasecalculo * CDbl(zValorICMS)) / 100)
            xValorIcms = (xbasecalculo * CDbl(zValorICMS)) / 100
        End If
        
'        If xValorIcms > 0 Then
'            'TRIBUTACAO NORMAL
'            Call PreencheColunaCst("00") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "00"
'        Else
'            'CST SUBSTITUICAO
'            Call PreencheColunaCst("60") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "60"
'        End If
    Else
        lBaseCalculoIcms_red = xbasecalculo
        
        lValorIcms_red = 0
        
        If g_reducao_invertido = 0 Then
            lBaseCalculoIcms_red = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
            lValorIcms_red = lBaseCalculoIcms_red * CDbl(zValorICMS) / 100
        Else
            x_calculo = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
            lBaseCalculoIcms_red = lBaseCalculoIcms_red - x_calculo
            lValorIcms_red = lBaseCalculoIcms_red * CDbl(zValorICMS) / 100
        End If
        
        'CST TRIBUTACAO COM REDUCAO
        'Call PreencheColunaCst("20") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "20"
        
        lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + lBaseCalculoIcms_red
        lValorIcms(4) = lValorIcms(4) + lValorIcms_red
        
        xbasecalculo = lBaseCalculoIcms_red
        xValorIcms = lValorIcms_red
    End If
    
     'reter imposto calculo substituicao com o iva
    If chk_retido.Value = 1 Then
        'If Val(grade1.TextMatrix(f, 42)) > 0 Then
        If x_perc_icms_subst > 0 Then
            SomatoriaSubst

            If grade1.TextMatrix(f, 12) = "S" And chk_red_subst.Value = 1 Then
                                                  
                If g_reducao_invertido = 0 Then
                    curValorRedSubstituicao = (xbcsubstituicao + (xbcsubstituicao * lporcredicmssubst) / 100)
                    xbcsubstituicao = (curValorRedSubstituicao * CDbl(grade1.TextMatrix(f, 42)) / 100)
                Else
                    x_calculo = (lporcredicmssubst * xbcsubstituicao) / 100
                    xbcsubstituicao = xbcsubstituicao - x_calculo
                    xbcsubstituicao = (xbcsubstituicao * CDbl(grade1.TextMatrix(f, 42)) / 100)
                End If
                                                
                'CST SUBSTITUICAO COM REDUCAO
                'Call PreencheColunaCst("70") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "70"
            Else
                xbcsubstituicao = xbcsubstituicao + ((xbcsubstituicao) * CDbl(grade1.TextMatrix(f, 42)) / 100)
                
                'CST SUBSTITUICAO SEM REDUCAO
                'Call PreencheColunaCst("10") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "10"
            End If
            
            lBaseCalculoSubstituicao = lBaseCalculoSubstituicao + xbcsubstituicao
          
            If xbcsubstituicao > 0 Then
                If chk_frete_subst.Value = 1 Then
                    xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
                Else
                    xvaloricms_conhecimento = 0
                End If
                xValorSubstituicao = ((xbcsubstituicao * x_perc_icms_subst) / 100) - (xValorIcms + xvaloricms_conhecimento)
                xValorSubstituicaocusto = ((xbcsubstituicao * x_perc_icms_subst) / 100)
                lValorSubstituicao = lValorSubstituicao + xValorSubstituicao
            End If
            
            If chk_soma_subs.Value = 1 Then
                lTotal = lTotal + xValorSubstituicao
            End If
        End If
    Else
        txt_bc_substituicao = Format(0, "##,###,##0.00")
        txt_substituicao = Format(0, "##,###,##0.00")
    End If

Exit Sub
Err_CalculaSubstituicao: ValidaErros Err, Me.Caption & " - CalculaSubstituicao"
End Sub

'*****************************************************************************
'Criação:                                                     Data: 19/08/2011
'
'Propósito:
'
'Alteração: Paulo Senhorini                                   Data: 19/08/2011
'           adicionei chamada a função preenchecolunacst.
'*****************************************************************************
Private Sub CalculaIsencao()
On Error GoTo Err_CalculaIsencao

    'Call PreencheColunaCst("40") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "40"

    If grade1.TextMatrix(f, 12) <> "S" Then
        If zValorICMS = 0 And xbasecalculo > 0 Then
            Call LancaCalculo
        Else
            lBaseCalculoIcms(lcodigoaliquota) = lBaseCalculoIcms(lcodigoaliquota) + xbasecalculo
            lValorIcms(lcodigoaliquota) = lValorIcms(lcodigoaliquota) + ((xbasecalculo * CDbl(zValorICMS)) / 100)
            'lBaseCalculoIcms(3) = lBaseCalculoIcms(3) + xbasecalculo
            'lValorIcms(3) = lValorIcms(3) + ((xbasecalculo * cdbl(zValorICMS)) / 100)
            xValorIcms = (xbasecalculo * CDbl(zValorICMS)) / 100
            xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
                  
            'CST TRIBUTADO
            'If xValorIcms > 0 Then Call PreencheColunaCst("00")  'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "00"
        End If
    Else
        lBaseCalculoIcms_red = xbasecalculo
        lValorIcms_red = 0
        If g_reducao_invertido = 0 Then
            lBaseCalculoIcms_red = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
            lValorIcms_red = lBaseCalculoIcms_red * CDbl(zValorICMS) / 100
        Else
            x_calculo = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
            lBaseCalculoIcms_red = lBaseCalculoIcms_red - x_calculo
            lValorIcms_red = lBaseCalculoIcms_red * CDbl(zValorICMS) / 100
        End If
                
        lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + lBaseCalculoIcms_red
        lValorIcms(4) = lValorIcms(4) + lValorIcms_red
        
        xbasecalculo = lBaseCalculoIcms_red
        xValorIcms = lValorIcms_red
        
        xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
        
        'CST TRIBUTADO COM REDUCAO
        'If xValorIcms > 0 Then Call PreencheColunaCst("20")  'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "20"
    End If

Exit Sub
Err_CalculaIsencao: ValidaErros Err, Me.Caption & " - CalculaIsencao"
End Sub

''*****************************************************************************
''Criação: Paulo Henrique de Aguiar Senhorini                   Data: 19/08/2011
''
''Propósito: se o campo 43 estiver preenchido o cst que
''           será lancado na nota será o que foi digitado, se não continuará dá
''           mesma forma.
''*****************************************************************************
'Private Sub PreencheColunaCst(ByVal strcodigo As String)
'On Error GoTo Err_PreencheColunaCst
'
'    If grade1.TextMatrix(f, 43) = "" Then
'        grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & strcodigo
'    End If
'
'Exit Sub
'Err_PreencheColunaCst: ValidaErros Err, Me.Caption & " - PreencheColunaCst"
'End Sub

Private Sub ZeraVariaveisTotais()
Dim j As Integer
    'Qtd_Aliquotas = 0
    If Not IsNumeric(txt_bc_substituicao) Then
        txt_bc_substituicao = 0
        x_substituicaoant = txt_substituicao
        txt_substituicao = 0
    Else
        x_substituicaoant = txt_substituicao
    End If
    
    lContaContabilVista = 0
    xtotalnotaprodutos = 0
    xvaloricms_conhecimento = 0
    xValorSubstituicao = 0
    xbcsubstituicao = 0
    xValorSubstituicaocusto = 0
    lValorSubstituicao = 0
    'txt_bc_substituicao = Format(0, "##,###,##0.00")
    'txt_substituicao = Format(0, "##,###,##0.00")
    lbl_bc_icms = 0
    lbl_valor_icms = 0
    lbl_ipi = 0
   ' lLancamento = 0
    ltotalcusto = 0
    lBaseCalculoSubstituicao = 0
    lBaseCalculoIcms_red = 0
    lValorIcms_red = 0
    lValorSubstituicao = 0
    lValorIPI = 0
    lTotalDesconto = 0
    lTotalFrete = 0
    lTotalSeguro = 0
    lTotalOutras = 0
    lTotalIPI = 0
    lTotalProdutos = 0
    lSubTotal = 0
    lTotal = 0
    lTotalPIS = 0
    lTotalCofins = 0
    lTotalII = 0
    
    If chk_calcularpesos.Value = 1 Then
        txt_peso_bruto = 0
        txt_peso_liquido = 0
    End If
    
    For j = 1 To 10
        lBaseCalculoIcms(j) = 0
        lValorIcms(j) = 0
        lDesconto(j) = 0
        lFrete(j) = 0
        lSeguro(j) = 0
        lOutras(j) = 0
        lFreteConhecimento(j) = 0
        lFreteBCConhecimento(j) = 0
        lpis(j) = 0
        lcofins(j) = 0
    Next
    lporcentagemreducao = txt_porc_red_icms
    lporcredicmssubst = txt_redicmssubst
    x_naotributados = 0
    
     For j = 0 To grade1.Rows - 1
         If IsNumeric(grade1.TextMatrix(j, 0)) Then
            If Trim(grade1.TextMatrix(j, 1)) = "" Then
                grade1.TextMatrix(j, 1) = 0
            End If
            grade1.TextMatrix(j, 14) = 0
            grade1.TextMatrix(j, 15) = 0
            grade1.TextMatrix(j, 19) = 0
            grade1.TextMatrix(j, 20) = 0
            'loop para o caso de calculo da nota
            If CDbl(grade1.TextMatrix(j, 10)) = 0 Then
                x_naotributados = x_naotributados + CDbl(grade1.TextMatrix(j, 11))
            End If
        End If
    Next
    
    ReDim VetorTributacaoProduto(0)
    
End Sub

Function ExisteProduto()
On Error GoTo Err_ExisteProduto
Dim z As Long

    ExisteProduto = False
    If g_duplicidade_produto = 0 Then
        For z = 0 To grade1.Rows - 1
            If z <> LastRow Then
                If grade1.TextMatrix(z, 2) = g_string And g_string <> "0" And grade1.TextMatrix(z, 1) = g_string2 Then
                   ExisteProduto = True
                End If
            End If
        Next
    End If

Exit Function
Err_ExisteProduto: ValidaErros Err, Me.Caption & " - ExisteProduto"
End Function

Private Function Verificacodigo()
On Error GoTo Err_Verificacodigo

    Verificacodigo = True
    For f = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(f, 0)) Then
            If grade1.TextMatrix(f, 2) <> grade1.TextMatrix(f, 29) And grade1.TextMatrix(f, 2) <> "0" Then
                Alerta "Código do Produto " & grade1.TextMatrix(f, 2) & " Não Confere com Descrição!"
                Verificacodigo = False
                Exit Function
            ElseIf Trim(grade1.TextMatrix(f, 3)) = "" Then
                Alerta "Informe a Descrição do Produto " & grade1.TextMatrix(f, 2) & "!"
                Verificacodigo = False
                Exit Function
            End If
            
            Call Verificafornecedor(grade1.TextMatrix(f, 2), grade1.TextMatrix(f, 3), txt_codigo_fornecedor, txt_fornecedor)
        End If
    Next

Exit Function
Err_Verificacodigo: ValidaErros Err, Me.Caption & " - Verificacodigo"
End Function
'*****************************************************************************
'Criação: Diego Martins dos Santos                      Data: 21/07/2010
'
'Propósito:Não permitir que o usuário exclua ou cancele notas de entrada
'com Devolução de compras,pois isto prejudica o rastreio das informações para estoque
'Alteraçao: Ronaldo Robledo                                     26/06/2013
'           Inserido validação de bloqueio para alteração da nota que já existirem devolucao de compra
'*****************************************************************************
Private Function VerificaDevolucaoCompras(ByVal bolApresentaMensagemAlteracao As Boolean) As Boolean
On Error GoTo Err_VerificaDevolucaoCompras

    VerificaDevolucaoCompras = False
    
    Sql_Query = "SELECT sequencia " & _
                "FROM movimento_nota_fiscal_saida MNFS " & _
                "WHERE numero_nota_referente = '" & txt_numero_nf & _
                "' AND outros = '" & lbl_nota & _
                "' and codigo_do_cliente = '" & txt_codigo_fornecedor & "'"
    Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 0)
    If gb_Recordset.RecordCount > 0 Then
        If Not bolApresentaMensagemAlteracao Then
            Alerta "Existe Devolução de Compras para esta NF - Verifique!!! ", 48
            VerificaDevolucaoCompras = True
        Else
            If bolAlteraEstoquenaAlteracaoNF Then
                Alerta "Quantidade dos produtos não pode ser alterada por ja existir devolução de compra para esta nota!"
                VerificaDevolucaoCompras = True
            End If
        End If
    End If
    gb_Recordset.Close

Exit Function
Err_VerificaDevolucaoCompras: ValidaErros Err, Me.Caption & " - VerificaDevolucaoCompras"
End Function
Function ExisteProdutoII()
On Error GoTo Err_ExisteProdutoII
Dim z As Long

    ExisteProdutoII = False
    If g_duplicidade_produto = 0 Then
        For f = 0 To grade1.Rows - 1
            If IsNumeric(grade1.TextMatrix(f, 0)) Then
                g_string = grade1.TextMatrix(f, 2)
                g_string2 = grade1.TextMatrix(f, 1)
                For z = 0 To grade1.Rows - 1
                    If z <> f Then
                        If grade1.TextMatrix(z, 2) = g_string And g_string <> "0" And grade1.TextMatrix(z, 1) = g_string2 Then
                           ExisteProdutoII = True
                        End If
                    End If
                Next
            End If
        Next
    End If

Exit Function
Err_ExisteProdutoII: ValidaErros Err, Me.Caption & " - ExisteProdutoII"
End Function

Private Sub EfetuaCancelamentoNF()
On Error GoTo Err_EfetuaCancelamentoNF

        If chk_impressao_nf.Value = 1 Then
            If Confirma("Confirma Cancelamento da Nota Fiscal Entrada ?") = vbYes Then
                If buscasenhaCanc Then
                    If cDPEmpresa.NotaFiscalEletronica = 1 Then
                        strMotivoCancelamento = InputBox("Informe o motivo do cancelamento da NF-e!", "Motivo Cancelamento!")
                    End If
                    Conexao.BeginTrans
                    If Cancelamento Then
                        Conexao.CommitTrans
                        Call Conexao.InserirRecordset("log_senhas", "data,hora,codigo_usuario,nome_usuario,historico,tela,observacoes,outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "','Cancelamento Nota Entrada " & txt_numero_nf & " " & txt_fornecedor & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)
                        g_string = ""
                        If bolCentralNFe Then
                            g_string = "NF-e Cancelada": Unload Me
                        Else
                            Alerta "Cancelamento Executado com Sucesso!"
                        End If
                        flag_tela_entrada_mercadoria = 0
                        Form_Activate
                    Else
                        Conexao.RollbackTrans
                    End If
                End If
            End If
        Else
            Alerta "Cancelamento apenas para notas impressas pela empresa!"
        End If
        
Exit Sub
Err_EfetuaCancelamentoNF: ValidaErros Err, Me.Caption & " - EfetuaCancelamentoNF"
End Sub

Private Sub cbo_frete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_placa_uf.SetFocus
End Sub

Private Sub cbo_frete_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_volume.SetFocus
End Sub

Private Sub cbo_TipoContaConsumo_Click()
    Call TrataRecursosConsumo
End Sub

Private Sub cbo_TipoContaConsumo_LostFocus()
    If Trim(txt_observacoes) = "" Then
        txt_observacoes = "Conta de " & Mid(cbo_TipoContaConsumo.Text, 4, 100)
    End If
End Sub

Private Sub cbo_unidade_GotFocus()
    SendKeys ("%{Down}")
End Sub

Private Sub cbo_unidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_qtdeconf.SetFocus
End Sub

'*****************************************************************************
'Alteração: Nayden Luiz                                       Data: 21/11/2011
'         : OS17972-Adicionada verificaçao se o cbo_Tipocontaconsumo esta
'         : habilitado ou nao para assim fazer a chamada do focus a ele pois
'         : quando ele nao esta habilitado e chama o focus ele da erro e para
'         : o sistema.
'
'Propósito:
'*****************************************************************************
Private Sub chk_ContaConsumo_Click()
    If chk_ContaConsumo.Value = 0 Then
        HabilitaControlesConsumo (False)
    Else
        HabilitaControlesConsumo (True)
        If cbo_TipoContaConsumo.Visible And cbo_TipoContaConsumo.Enabled Then 'OS17972-Se ele estiver visivel e habilitado
           cbo_TipoContaConsumo.SetFocus                                     'ele chama o foco para ele
        End If
    End If
End Sub

'*****************************************************************************
'Criação:                                                  Data: 20/04/2011
'
'Propósito:
'
'Alteração: Foi necessário incluir no txt_numero_nf 0 e travar o txt para que
'           quando o usuario marcar impressão de NF o sistema gerar o numero
'           sozinho.
'*****************************************************************************
Private Sub chk_impressao_nf_Click()
    If chk_impressao_nf.Value = 1 And blnExisteNota = False Then
        txt_numero_nf.Text = 0
        txt_numero_nf.Locked = True
    Else
        txt_numero_nf.Locked = False
    End If
End Sub

Private Sub chkImportacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_desconto.SetFocus
End Sub

Private Sub chkImportacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_observacoes.SetFocus
End Sub

Private Sub cmd_explorer_Click()
    buscaxml.InitDir = "C:\"
    buscaxml.ShowOpen
    txt_caminho_xml.Text = buscaxml.FileName
    buscaxml.FileName = ""
End Sub

Private Sub cmd_imprimir_Click()
    Call ImprimirConf
End Sub

'*****************************************************************************
'Criação: Diego Martins dos Santos                            Data: 19/06/2010
'Propósito:Pesquisa notas pendentes para conferência
'*****************************************************************************
Private Sub cmd_pesquisar_conferencia_Click()
On Error GoTo Err_cmd_pesquisar_conferencia_Click
    Call PesquisaConferencia
Exit Sub
Err_cmd_pesquisar_conferencia_Click: ValidaErros Err, Me.Caption & " - cmd_pesquisar_conferencia_Click"
End Sub

'*****************************************************************************
'Criação: Diego Martins dos Santos                            Data: 19/06/2010
'Propósito:Salva Conferência
'*****************************************************************************
Private Sub cmd_salvar_conferencia_Click()
On Error GoTo Err_cmd_salvar_conferencia_Click
    If grade3.TextMatrix(1, 0) <> "" Then
        If validacamposii Then
            Conexao.BeginTrans
            If AtualizaMovimentoConferencia Then
               Conexao.CommitTrans
               Alerta "Conferência Salva com Sucesso!", 64
               LimpaTela
            Else
               Conexao.RollbackTrans
            End If
        End If
    Else
        Alerta "Informe os Itens para Conferência!", 48
    End If
Exit Sub
Err_cmd_salvar_conferencia_Click: ValidaErros Err, Me.Caption & " - cmd_salvar_conferencia_Click"
End Sub

Private Sub Form_Deactivate()
    flag_tela_entrada_mercadoria = 1
End Sub

Private Sub grade1_Click()
    If frmdados.Enabled = False Then SSTab1.SetFocus
End Sub

Private Sub grade1_DblClick()
On Error GoTo Err_grade1_DblClick

    'abre tela de II
    If chkImportacao = 1 Then cadastro_II.Show vbModal
    
    If Val(grade1.TextMatrix(grade1.RowSel, 39)) = 1 Then
        Call ChamaGradeTamanho(grade1.TextMatrix(grade1.RowSel, 2), grade1.RowSel)
    End If
    
    '***Trata informações do Cod.Barras ***
    If Trim(grade1.TextMatrix(grade1.RowSel, 2)) <> "" And chk_impressao_nf.Value = 1 Then
        grade1.TextMatrix(grade1.RowSel, 38) = ObtenhaCodigoBarras(grade1.TextMatrix(grade1.RowSel, 2))
    End If

Exit Sub
Err_grade1_DblClick: ValidaErros Err, Me.Caption & " - grade1_DblClick"
End Sub

Private Sub grade3_KeyDown(KeyCode As Integer, Shift As Integer)
    If frm_conf.Enabled = True Then
        If KeyCode = &H2E Then
            ' Excluir linhas selecionadas
            If grade3.Rows - 1 > 1 Then
                grade3.RemoveItem grade3.RowSel
            Else
                grade3.Clear
                FormaGridConf
            End If
            Somaprodutosconf
            txt_codigo_produto.SetFocus
        End If
    End If
End Sub


Private Sub grade3_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txt_chave_acesso_GotFocus()
    txt_chave_acesso.BackColor = 12648447
    txt_chave_acesso.SelStart = 0
    txt_chave_acesso.SelLength = Len(txt_chave_acesso)
End Sub

Private Sub txt_chave_acesso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then msk_emissao.SetFocus
End Sub

Private Sub txt_chave_acesso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then msk_entrada.SetFocus
End Sub

Private Sub txt_chave_acesso_LostFocus()
    txt_chave_acesso.BackColor = &H8000000E
    '*** Autor: Diego Martins Ticket: TT438 Data: 26/06/2012 ***
    If Len(txt_chave_acesso) < 44 And Trim(txt_modelo_nf) = "55" Then
        Alerta "Informe o número da Chave da NF-e corretamente!", 48
        txt_chave_acesso.SetFocus
    End If
End Sub

Private Sub txt_especie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cbo_frete.SetFocus
End Sub

Private Sub txt_especie_GotFocus()
    txt_especie.BackColor = 12648447
    txt_especie.SelStart = 0
    txt_especie.SelLength = Len(txt_especie)
End Sub

Private Sub txt_especie_LostFocus()
    txt_especie.BackColor = &H8000000E
    txt_especie = UCase(Trim(txt_especie))
End Sub

Private Sub txt_especie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_peso_bruto.SetFocus
End Sub

Private Sub txt_fornecedor_GotFocus()
    txt_fornecedor.BackColor = 12648447
    txt_fornecedor.SelStart = 0
    txt_fornecedor.SelLength = Len(txt_fornecedor)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = g_formato_carac Then
        If Liberacao Then
            pct_orçamento.Visible = True
            pct_icms.Visible = False
            If frmdados.Enabled = True Then
                txt_serie_nf = "CX"
                txt_modelo_nf = "1"
                txt_serie_nf.Enabled = False
            End If
        End If
    
    ElseIf KeyAscii = 14 Then
        pct_orçamento.Visible = False
        pct_icms.Visible = False
        txt_serie_nf.Enabled = True
        If frmdados.Enabled = True Then
            If txt_serie_nf = "CX" Then
                txt_serie_nf = ""
                txt_modelo_nf = ""
            End If
        End If
    ElseIf KeyAscii = 9 Then
        If Liberacao Then
            pct_orçamento.Visible = False
            pct_icms.Visible = True
            txt_serie_nf.Enabled = True
            If frmdados.Enabled = True Then
                If txt_serie_nf = "CX" Then
                    txt_serie_nf = ""
                    txt_modelo_nf = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cDPParamSistema = Nothing
    Set Movimento_Nota_Fiscal_Entrada = Nothing
    Set MovimentoCaixaGeral = Nothing
    Set ContasAPagar = Nothing
    Set MovCabNotaFiscalEntrada = Nothing
    Set MovNotaFiscalEntrada = Nothing
    Set MovimentoCaixa = Nothing
    Set cMapLote = Nothing
    Set cDPEmpresa = Nothing
    Set cDPNotasSaida = Nothing
    Set cCtrlEntradaSaida = Nothing
    Set colTributacao = Nothing
    Set colFormaFaturamento = Nothing
    Set cDPFFaturamento = Nothing
    Set cUtGeral = Nothing
End Sub

Private Function BaixaEstoque() As Boolean
On Error GoTo file

    BaixaEstoque = False
    
    Set gb_Recordset = Conexao.GeraRecordset("SELECT codigo_do_produto,outros,quantidade,industria_revenda FROM movimento_nota_fiscal_entrada WHERE empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, 0)
    If gb_Recordset.RecordCount > 0 Then
        gb_Recordset.MoveFirst
        Do While Not gb_Recordset.EOF
            If gb_Recordset!codigo_do_produto <> "0" Then
                If Trim(gb_Recordset!Outros) = "CX" Then
                    If Conexao.AlterarRecordset("estoque", "quantidade_cx = quantidade_cx - " & fValidaValor2(gb_Recordset!Quantidade), "empresa = '" & cDPEmpresa.codigo & "' and codigo_do_produto = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo) Then
                        Call analizaEstoque(gb_Recordset!codigo_do_produto, gb_Recordset!Quantidade, txt_numero_nf, "S", "Entrada Nota CX", "Estoque")
                    Else
                        Alerta "Erro na atualização do estoque " & Chr(10) & Err.Description & "!", vbCritical
                        gb_Recordset.Close
                        Exit Function
                    End If
                ElseIf Trim(gb_Recordset!Outros) = "NF" Then
                    If Conexao.AlterarRecordset("estoque", "quantidade = quantidade - " & fValidaValor2(gb_Recordset!Quantidade), "empresa = '" & cDPEmpresa.codigo & "' and  codigo_do_produto = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo) Then
                        Call analizaEstoque(gb_Recordset!codigo_do_produto, gb_Recordset!Quantidade, txt_numero_nf, "S", "Entrada Nota NF", "Estoque")
                    Else
                        Alerta "Erro na atualização do estoque " & Chr(10) & Err.Description & "!", vbCritical
                        gb_Recordset.Close
                        Exit Function
                    End If
                    If Conexao.AlterarRecordset("estoque", "quantidade_cx = quantidade_cx - " & fValidaValor2(gb_Recordset!Quantidade), " empresa = '" & cDPEmpresa.codigo & "' and codigo_do_produto = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo) Then
                        Call analizaEstoque(gb_Recordset!codigo_do_produto, gb_Recordset!Quantidade, txt_numero_nf, "S", "Entrada Nota CX", "Estoque")
                    Else
                        Alerta "Erro na atualização do estoque " & Chr(10) & Err.Description & "!", vbCritical
                        gb_Recordset.Close
                        Exit Function
                    End If
                ElseIf Trim(gb_Recordset!Outros) = "CI" Then
                    If Conexao.AlterarRecordset("estoque", "quantidade = quantidade - " & fValidaValor2(gb_Recordset!Quantidade), "empresa = '" & cDPEmpresa.codigo & "' and  codigo_do_produto = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo) Then
                        Call analizaEstoque(gb_Recordset!codigo_do_produto, gb_Recordset!Quantidade, txt_numero_nf, "S", "Entrada Nota CI", "Estoque")
                    Else
                        Alerta "Erro na atualização do estoque " & Chr(10) & Err.Description & "!", vbCritical
                        gb_Recordset.Close
                        Exit Function
                    End If
                End If
            End If
            gb_Recordset.MoveNext
        Loop
        BaixaEstoque = True
    Else
        Alerta "Produtos da Nota Não Localizado p/ Saida do Estoque!"
    End If
    gb_Recordset.Close

Exit Function
file: ValidaErros Err, Me.Caption & " - BaixaEstoque"
End Function


'*****************************************************************************
'Criação: Ronaldo Robledo Mendes Souza                      Data: 30/05/2011
'
'Propósito:
'Alteração:Ronaldo Robledo                                  Data: 30/05/2011
'          Inserido gravar na tabela produto campo ultimo_precocompra para atender legislação
'          da substituição tributaria junto ao inventário
'Alteração:Clejunior                                        Data: 23/04/2013
'         : If IsNumeric(grade1.TextMatrix(z, 45)) Then colocado validação para quando
'         : se utiliza unidade secundaria. TT3166
'Alteração: Ronaldo Robledo                                     06/06/2013
'           Inserido chamada de método para exclusão da nota quando for alteração
'*****************************************************************************
Private Function AtualizaTabelas() As Boolean
On Error GoTo Err_AtualizaTabelas

    AtualizaTabelas = False

    'Alteração de Nota fiscal
    If l_opcao = 2 Then If Not ExclusaoNota(True) Then Exit Function
    
    If Not AtualizaDocumentoImportacao Then Exit Function
    
    If Not AtualizaInfoContaConsumo Then Exit Function
    
    If Not AtualizaContasAPagar Then Exit Function
    
    If Not AtualizaMovimentoCabecalhoNotaFiscalEntrada Then Exit Function
    
    If Not AtualizaConferencia Then Exit Function
    
    If Not AtualizaItensdaNota Then Exit Function
    
    If bolAlteraEstoquenaAlteracaoNF Then
        If cDPFFaturamento.InfluenciaEstoque = 1 And g_utiliza_estoque_deposito = 1 And pct_icms.Visible = False Then
            If Not EntradaSaidaEstoqueDeposito(xid_deposito, lSequencia, "E", 1, "+", False) Then Exit Function
        End If
        
        If Not AtualizaMovimentoGrade Then Exit Function
    End If

    Call Conexao.InserirRecordset("log_senhas", "data, hora, codigo_usuario, nome_usuario, historico, tela, observacoes, outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "','Entrada " & msk_entrada & " Nº da Nota " & txt_numero_nf & " Forn. " & txt_fornecedor & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)
    
    g_NFEConfirmada = False
    'verifica NOTA FISCAL ELETRONICA
    If cDPEmpresa.NotaFiscalEletronica = 1 And chk_impressao_nf.Value = 1 Then
        'Call BuscaCFOP
        
        frm_Processo_NFe.lbl_texto = "Iniciando processo de emissão de NF-e..."
        frm_Processo_NFe.lSequencia = lSequencia
        frm_Processo_NFe.strObservacaoLei = cDPFFaturamento.MensagemPadrao
        Set frm_Processo_NFe.frm = Me
        frm_Processo_NFe.Show 1
        
        If g_NFEConfirmada = False Then Exit Function
    End If
    
    AtualizaTabelas = True

Exit Function
Err_AtualizaTabelas: ValidaErros Err, Me.Caption & " - AtualizaTabelas"
End Function
'==========================================================================
' Purpose:  Persistir os dados referente aos itens da nota
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
' Alteração: Ronaldo Robledo                                    21/06/2013
'           Inserido validação de alteração do movimento_lote quando se tratar de alteração de nota
'==========================================================================
Private Function AtualizaItensdaNota() As Boolean
On Error GoTo Err_AtualizaItensdaNota
Dim z               As Long
    
    AtualizaItensdaNota = False
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 0)) Then
            
            If bolLiberaAlteracaoCustoAlteracaoNF Then If Not PersisteCustoProdutos(z) Then Exit Function
                
            If Not AtualizaMovimentoNotaFiscalEntrada(z) Then Exit Function
                        
            If bolAlteraEstoquenaAlteracaoNF Then
                If Not AtualizaPedidoMercadoria(z) Then Exit Function
    
                If cDPFFaturamento.InfluenciaEstoque = 1 Then If Not AtualizaEstoque(z) Then Exit Function
            
                'atualiza tabela de movimento_transito_compra
                If cDPFFaturamento.Contrapartida = 1 Then If Not AtualizaContra(z) Then Exit Function
            
                'Atualizar sequencia do lote.
                If grade1.TextMatrix(z, 24) = "1" Then
                    If l_opcao = 2 Then
                        If Not Conexao.AlterarRecordset("movimento_lote", "sequencia = " & lSequencia, "empresa = '" & cDPEmpresa.codigo & "' and seq_controle = '" & lngSequenciaControle & "'", cDPEmpresa.codigo) Then Exit Function
                    Else
                        If Not Conexao.AlterarRecordset("movimento_lote", "sequencia = " & lSequencia, "empresa = '" & cDPEmpresa.codigo & "' and codigo_produto = '" & grade1.TextMatrix(z, 2) & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & "' and numero_pedido = '" & txt_numero_nf & "' and data_emissao = " & FormataData(msk_emissao) & " and sequencia = '1' and status = 'E'", cDPEmpresa.codigo) Then Exit Function
                    End If
                End If
            End If
        End If
    Next
    AtualizaItensdaNota = True

Exit Function
Err_AtualizaItensdaNota: ValidaErros Err, Me.Caption & " - AtualizaItensdaNota"
End Function

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Function PersisteCustoProdutos(ByVal z As Long) As Boolean
On Error GoTo Err_PersisteCustoProdutos
Dim x_preco_varejo  As String
Dim x_preco_atacado As String
Dim x_desc_atacado  As String
Dim x_desc_varejo   As String
Dim curVerificaFatorConversao As Currency
    
    PersisteCustoProdutos = False
    x_preco_atacado = "0"
    x_preco_varejo = "0"
    g_string2 = 0
    g_string3 = 0
    g_string4 = 0
    zValorTotal = 0
    zQuantidade = 0
    zCustoAnterior = 0
    'se libera atualizar custo
    'If grade1.TextMatrix(z, 2) <> "0" Then
    If pct_icms.Visible = False Then
        'Calculo do Custo Medio
        Set gb_Recordset = Conexao.GeraRecordset(" SELECT estoque.quantidade_cx,produto.custo_medio " & _
                                                 " FROM estoque, produto " & _
                                                 " WHERE estoque.codigo_do_produto = '" & grade1.TextMatrix(z, 2) & "' " & _
                                                 " and produto.codigo = '" & grade1.TextMatrix(z, 2) & "'", 1)
        If gb_Recordset.RecordCount > 0 Then
            If gb_Recordset!quantidade_cx > 0 Then
                zQuantidade = gb_Recordset!quantidade_cx
            Else
                zQuantidade = 0
            End If
            zCustoAnterior = gb_Recordset!custo_medio
        End If
        gb_Recordset.Close
        
        If CDbl(grade1.TextMatrix(z, 5)) > 0 Then
            'zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + (grade1.TextMatrix(z, 5) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5))), "##,###,##0.0000")
            'zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + ((grade1.TextMatrix(z, 5) * (grade1.TextMatrix(z, 45))) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5) * (grade1.TextMatrix(z, 45)))), "##,###,##0.0000")
            If CDbl(grade1.TextMatrix(z, 5)) > 0 Then
                If IsNumeric(grade1.TextMatrix(z, 45)) Then
                    zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + ((grade1.TextMatrix(z, 5) * (grade1.TextMatrix(z, 45))) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5) * (grade1.TextMatrix(z, 45)))), "##,###,##0.0000")
                Else
                    zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + (grade1.TextMatrix(z, 5) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5))), "##,###,##0.0000")
                End If
            End If
            g_string2 = lcustoprodutos(z)
            g_string3 = zValorTotal
            grade1.TextMatrix(z, 55) = Format(zValorTotal, g_decimal_custo)
        End If
                       
        If Left(cbo_atualizacusto, 1) > 0 Then
            Call AtualizaCustoMedioAnterior(z)
            
            Set gb_Recordset = Conexao.GeraRecordset("SELECT marckup_atacado,desconto_atacado,marckup_varejo,desconto_varejo FROM produto WHERE codigo = '" & grade1.TextMatrix(z, 2) & "'", 1)
            If gb_Recordset.RecordCount > 0 Then
                x_preco_atacado = gb_Recordset!marckup_atacado
                x_preco_varejo = gb_Recordset!marckup_varejo
                x_desc_atacado = gb_Recordset!desconto_atacado
                x_desc_varejo = gb_Recordset!desconto_varejo
            End If
            gb_Recordset.Close
            
            'grava precos que serão alterados
            'If Left(cbo_atualizacusto, 1) = 1 Then
            '    If CDbl(x_preco_atacado) > 0 Or CDbl(x_preco_varejo) > 0 Then
                    Call GravaPrecos(grade1.TextMatrix(z, 2), "E")
            '    End If
            'End If
            
            'atualiza precos e custos
            If Left(cbo_atualizacusto, 1) = 1 Then
                'alteração de preco manualmente
                If chk_lancamento_venda.Value = 0 Then
                    If CDbl(x_preco_atacado) > 0 Then
                         'calculo do sem coeficiente
                         x_preco_atacado = (lcustoprodutos(z) * CDbl(x_preco_atacado)) / 100
                         x_preco_atacado = Format(lcustoprodutos(z) + CDbl(x_preco_atacado), g_decimal_venda)
                        
                         g_string = 0
                         g_string = Format((x_desc_atacado * x_preco_atacado) / 100, g_decimal_venda)
                         x_desc_atacado = Format(CDbl(x_preco_atacado) - CDbl(g_string), g_decimal_venda)
                         x_desc_atacado = Format(x_desc_atacado, g_decimal_venda)
                                                               
                         Call Conexao.AlterarRecordset("produto", "data_alteracao = " & FormataData(Date) & ",preco_atacado = " & fValidaValor2(x_preco_atacado) & ",minimo_atacado = " & fValidaValor2(x_desc_atacado) & ",data_atacado = " & FormataData(Date) & ",alterar = 'S'", "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                         
                         Call Conexao.AlterarRecordsetI("codigo_barras", "descricao = 'ATACADO       VAREJO" & Mid(x_preco_atacado, 1, 6) & "' WHERE codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                    End If
                                    
                    If CDbl(x_preco_varejo) > 0 Then
                        x_preco_varejo = (lcustoprodutos(z) * CDbl(x_preco_varejo)) / 100
                        x_preco_varejo = Format(lcustoprodutos(z) + CDbl(x_preco_varejo), g_decimal_venda)
                    
                        g_string = 0
                        g_string = Format((x_desc_varejo * x_preco_varejo) / 100, g_decimal_venda)
                        x_desc_varejo = Format(CDbl(x_preco_varejo) - CDbl(g_string), g_decimal_venda)
                        x_desc_varejo = Format(x_desc_varejo, g_decimal_venda)
                        
                        Call Conexao.AlterarRecordset("produto", "data_alteracao = " & FormataData(Date) & ",preco_varejo = " & fValidaValor2(x_preco_varejo) & ",minimo_varejo = " & fValidaValor2(x_desc_varejo) & ",data_varejo = " & FormataData(Date) & ",alterar = 'S'", "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                        
                        Call Conexao.AlterarRecordsetI("codigo_barras", "preco_venda = " & fValidaValor2(x_preco_varejo) & " WHERE codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                    End If
                Else
                    If Not SalvarPrecoManual(z) Then Exit Function
                End If
            End If
            
            'OS 11059 - ATUALIZAR CADASTRO PRODUTO QUANDO ALTERA CUSTO NA ENTRADA
            '  TT3168 "ultimo_precocompra = " & fValidaValor2(grade1.TextMatrix(z, 8)), _
            '  feito validação da grade de fator de conversao e numerica ou não
            
            If IsNumeric(grade1.TextMatrix(z, 45)) Then
               curVerificaFatorConversao = grade1.TextMatrix(z, 45)
            Else
               curVerificaFatorConversao = 1
            End If
            
            If Not Conexao.AlterarRecordset("produto", "data_alteracao = " & FormataData(Date) & "," & _
                                            "custo = " & fValidaValor2(g_string2) & "," & _
                                            "custo_medio = " & fValidaValor2(zValorTotal) & "," & _
                                            "data_custo = " & FormataData(Date) & "," & _
                                            "data_custo_medio = " & FormataData(Date) & "," & _
                                            "alterar = 'S',altera_custo = 'S'," & _
                                            "ultimo_precocompra = " & fValidaValor2(grade1.TextMatrix(z, 8) / curVerificaFatorConversao), _
                                            "codigo = '" & grade1.TextMatrix(z, 2) & "'", _
                                            cDPEmpresa.codigo) Then Exit Function
            
        End If
                    
        If pct_orçamento.Visible = False Then If Not CalculoMeiaNota(z) Then Exit Function

    Else
        'calculo para entrada credito icms e meia nota
        If Not CalculoMeiaNota(z) Then Exit Function
    End If

    PersisteCustoProdutos = True
    
Exit Function
Err_PersisteCustoProdutos: ValidaErros Err, Me.Caption & " - PersisteCustoProdutos"
End Function


'*****************************************************************************
'Criação: Diego Martins dos Santos                            Data: 23/03/2011
'Propósito:
'*****************************************************************************
Private Function AtualizaDocumentoImportacao() As Boolean
On Error GoTo Err_AtualizaDocumentoImportacao

    AtualizaDocumentoImportacao = False
        
        lngCodigoDocImportacao = 0
        
        If chkImportacao.Value = 1 Then
            movimento_documento_importacao.txt_codigo_exportador = txt_codigo_fornecedor
            movimento_documento_importacao.txt_nome_exportador = txt_fornecedor
            movimento_documento_importacao.Show 1
            
            If Val(g_string3) > 0 Then
                lngCodigoDocImportacao = CLng(g_string3)
            Else
                Alerta "Obrigatório o lançamento do documento de Importação!", 48
                Exit Function
            End If
        End If
        
    AtualizaDocumentoImportacao = True

Exit Function
Err_AtualizaDocumentoImportacao: ValidaErros Err, Me.Caption & " - AtualizaDocumentoImportacao"
End Function

'*****************************************************************************
'Criação: Diego Martins dos Santos                      Data: 23/03/2011
'
'Propósito:
'*****************************************************************************
Private Function AtualizaInfoContaConsumo() As Boolean
On Error GoTo Err_AtualizaInfoContaConsumo
Dim strCampos, strValores As String

    AtualizaInfoContaConsumo = False
    
    lngCodigoContaConsumo = 0
    
    If chk_ContaConsumo.Value = 1 Then
        lngCodigoContaConsumo = BuscaProximoCodigoMovimentacao("CONSUMO")
         
         strCampos = "empresa,codigo,tipo_conta,codigo_consumo,tipo_ligacao,grupo_tensao"
         
         strValores = "'" & cDPEmpresa.codigo & "','" & lngCodigoContaConsumo & _
                      "','" & Left(cbo_TipoContaConsumo.Text, 1) & "',"
 
         If Left(cbo_TipoContaConsumo.Text, 1) = 3 Then
            strValores = strValores & "'" & Left(cbo_CodigoConsumo.Text, 1) & "',"
         Else
             strValores = strValores & "'" & Left(cbo_CodigoConsumo.Text, 2) & "',"
         End If
         
         If Left(cbo_TipoContaConsumo.Text, 1) = 1 Then
            strValores = strValores & "'" & Left(cbo_TipodeLigacao.Text, 1) & _
                                      "','" & Left(cbo_GrupoDeTensao.Text, 2) & "'"
         Else
             strValores = strValores & "null,null"
         End If
        
        If Not Conexao.InserirRecordset("info_conta_consumo", strCampos, strValores, cDPEmpresa.codigo) Then Exit Function
    End If
    
AtualizaInfoContaConsumo = True

Exit Function
Err_AtualizaInfoContaConsumo: ValidaErros Err, Me.Caption & " - AtualizaInfoContaConsumo"
End Function

Private Function AtualizaTabelasConferencia() As Boolean
On Error GoTo AtualizaTabelasConferencia
Dim z As Long

    AtualizaTabelasConferencia = False
    
    If Not AtualizaMovimentoCabecalhoNotaFiscalEntrada Then Exit Function
    
    If Not AtualizaConferencia Then Exit Function
    
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 0)) And Trim(grade1.TextMatrix(z, 2)) <> "" Then
            grade1.TextMatrix(z, 18) = "0" ' Variavel para Custo
            grade1.TextMatrix(z, 56) = "0" ' Variacel para Custo Contabil
            grade1.TextMatrix(z, 55) = "0" ' Variavel para Custo Médio
            If Not AtualizaMovimentoNotaFiscalEntrada(z) Then Exit Function
        End If
    Next
    
    AtualizaTabelasConferencia = True

Exit Function
AtualizaTabelasConferencia: ValidaErros Err, Me.Caption & " - AtualizaTabelasConferencia"
End Function

Private Sub AtualizaCustoMedioAnterior(ByVal z As Long)
    If l_opcao = 1 Then
        Call Conexao.AlterarRecordset("produto", "custo_anterior = produto.custo,custo_medio_anterior = produto.custo_medio", "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
    End If
End Sub

Private Function AtualizaEstoque(ByVal z As Long) As Boolean
On Error GoTo Err_AtualizaEstoque
Dim CodigoProduto   As String
Dim FatorConversao  As Double 'Fator de conversão caso o produto tem unidade secundária


    AtualizaEstoque = False

    'Caso não esteja usando a unid. secundária, o fator receberá o valor 1 (um)
    If IsNumeric(grade1.TextMatrix(z, 45)) Then FatorConversao = grade1.TextMatrix(z, 45) Else FatorConversao = 1

    If Trim(grade1.TextMatrix(z, 2)) <> "" And Trim(grade1.TextMatrix(z, 2)) <> "0" Then
        CodigoProduto = grade1.TextMatrix(z, 2)
        If pct_orçamento.Visible = False And pct_icms.Visible = False Then
            If Not Conexao.AlterarRecordset("estoque", "quantidade = quantidade + " & fValidaValor2(grade1.TextMatrix(z, 5) * FatorConversao), "codigo_do_produto = '" & CodigoProduto & "' and empresa = '" & cDPEmpresa.codigo & "'", cDPEmpresa.codigo) Then
                Exit Function
            Else
                Call analizaEstoque(CodigoProduto, grade1.TextMatrix(z, 5) * FatorConversao, txt_numero_nf, "E", "Entrada Nota NF", "Estoque")
            End If
            If Not Conexao.AlterarRecordset("estoque", "quantidade_cx = quantidade_cx + " & fValidaValor2(grade1.TextMatrix(z, 5) * FatorConversao), "codigo_do_produto = '" & CodigoProduto & "' and empresa = '" & cDPEmpresa.codigo & "'", cDPEmpresa.codigo) Then
                Exit Function
            Else
                Call analizaEstoque(CodigoProduto, grade1.TextMatrix(z, 5) * FatorConversao, txt_numero_nf, "E", "Entrada Nota CX", "Estoque")
            End If
        ElseIf pct_icms.Visible = True Then
            If Not Conexao.AlterarRecordset("estoque", "quantidade = quantidade + " & fValidaValor2(grade1.TextMatrix(z, 5) * FatorConversao), "codigo_do_produto = '" & CodigoProduto & "' and empresa = '" & cDPEmpresa.codigo & "'", cDPEmpresa.codigo) Then
                Exit Function
            Else
                Call analizaEstoque(CodigoProduto, grade1.TextMatrix(z, 5) * FatorConversao, txt_numero_nf, "E", "Entrada Nota CI", "Estoque")
            End If
        Else
            If Not Conexao.AlterarRecordset("estoque", "quantidade_cx = quantidade_cx + " & fValidaValor2(grade1.TextMatrix(z, 5) * FatorConversao), "codigo_do_produto = '" & CodigoProduto & "' and empresa = '" & cDPEmpresa.codigo & "'", cDPEmpresa.codigo) Then
                Exit Function
            Else
                Call analizaEstoque(CodigoProduto, grade1.TextMatrix(z, 5) * FatorConversao, txt_numero_nf, "E", "Entrada Nota CX", "Estoque")
            End If
        End If
    
        Call Conexao.AlterarRecordset("produto", "codigo_fornecedor = '" & txt_codigo_fornecedor & "',nome_fornecedor = '" & txt_fornecedor & "'", "codigo = '" & CodigoProduto & "'", cDPEmpresa.codigo)
    End If
    AtualizaEstoque = True

Exit Function
Err_AtualizaEstoque: ValidaErros Err, Me.Caption & " - AtualizaEstoque"
End Function

Private Function AtualizaMovimentoCabecalhoNotaFiscalEntrada() As Boolean
On Error GoTo Err_AtualizaMovimentoCabecalhoNotaFiscalEntrada

    AtualizaMovimentoCabecalhoNotaFiscalEntrada = False
        
    '*** Alterado para contemplar,conceito de notas através da conferência
    ' e não afetar o funcionamento normal do sistema ***
    If lSequencia = 0 Then lSequencia = MovCabNotaFiscalEntrada.ProximaSequencia
    
    With MovCabNotaFiscalEntrada
        .Empresa = cDPEmpresa.codigo
        .DataEntrada = EnviaData(msk_entrada)
        .DataEmissao = EnviaData(msk_emissao)
        .Numero = txt_numero_nf
        .Serie = txt_serie_nf
        .ModeloNf = txt_modelo_nf
        .TipoDocumento = txt_codigo_forma
        .CodificacaoFiscal = grade1.TextMatrix(1, 28)
        .CodigoFornecedor = txt_codigo_fornecedor
        .NomeFornecedor = txt_fornecedor.Text
        .UF = lbl_uf
        .CGC = lbl_cgc
        .InscricaoEstadual = lbl_inscricao
        .ValorTotalProdutos = lbl_total_produtos
        .BaseCalculoIcms = lbl_bc_icms
        .ValorIcms = lbl_valor_icms
        .BaseCalculoSubstituicaoIcms = lbl_bc_substituicao
        .ValorICMSSubstituicao = lbl_icms_substituicao
        .ValorIPI = lbl_ipi
        .Frete = lbl_frete
        .FreteConhecimento = txt_frete_conhecimento
        .Seguro = lbl_seguro
        .OutrasDespesas = lbl_outras_despesas
        .ValorDesconto = txt_desconto
        .TotalNota = lbl_total
        .TotalCusto = ltotalcusto
        .NumeroMovimentoCaixa = lLancamento
        .ImpressoNF = chk_impressao_nf.Value
        .ImpostoRetido = chk_retido.Value
        .AtualizaCusto = Left(cbo_atualizacusto.Text, 1)
        .CodigoTransportadora = txt_codigo_transportadora
        .Placa = txt_transportadora_placa
        .UfPlaca = txt_transportadora_placa_uf
        .tipoFrete = Left(cbo_frete, 1)
        .Observacoes = txt_observacoes
        .Usuario = g_usuario
        .CalculoNota = chk_calculo_nota.Value
        .AtualizaCarteira = chk_atualiza_carteira.Value
        .AtualizaCaixa = chk_atualiza_caixa.Value
        .TipoCaixa = lTipoCaixa
        .DataCaixa = EnviaData(lDataCaixa)
        .DataChegada = EnviaData(msk_chegada)
        .sequencia = lSequencia
        .TotalPis = lTotalPIS
        .TotalCofins = lTotalCofins
        .TotalImportacao = lTotalII
        .Volume = txt_volume
        .Especie = txt_especie
        .PesoBruto = txt_peso_bruto
        .PesoLiquido = txt_peso_liquido
        .CodigoDI = lngCodigoDocImportacao
        .ContaConsumo = chk_ContaConsumo.Value
        .CodigoContaConsumo = cUtGeral.EnviaBanco(lngCodigoContaConsumo)
        .ChaveAcesso = txt_chave_acesso
        .NotaComplementar = chk_complementar.Value
    End With
    
    If Not MovCabNotaFiscalEntrada.Incluir(ValidaBooleanoConferencia) Then Exit Function
    
    AtualizaMovimentoCabecalhoNotaFiscalEntrada = True

Exit Function
Err_AtualizaMovimentoCabecalhoNotaFiscalEntrada: ValidaErros Err, Me.Caption & " - AtualizaMovimentoCabecalhoNotaFiscalEntrada"
End Function

'*****************************************************************************
'Criação: Ronaldo Robledo Mendes de Souza                      Data:
'Alteraçao: Ronaldo Robledo                                          29/01/2012
'           Inserido função para atualizar entidade movimento_impostos para atender SPED
'*****************************************************************************
Private Function AtualizaMovimentoNotaFiscalEntrada(ByVal z As Long) As Boolean
On Error GoTo Err_AtualizaMovimentoNotaFiscalEntrada
Dim lngPkCodigo As Long

    AtualizaMovimentoNotaFiscalEntrada = False
    With grade1
        zGrupo = .TextMatrix(z, 1)
        zProduto = Trim(.TextMatrix(z, 2))
        zNome = .TextMatrix(z, 3)
        zUnidade = .TextMatrix(z, 4)
        If IsNumeric(.TextMatrix(z, 45)) Then
            zQuantidade = .TextMatrix(z, 5) * .TextMatrix(z, 45)
        Else
            zQuantidade = .TextMatrix(z, 5)
        End If
        zvalorbruto = .TextMatrix(z, 6)
        zporc_desconto = .TextMatrix(z, 7)
        zValorUnitario = .TextMatrix(z, 8)
        zValorIPI = .TextMatrix(z, 13)
        zValorICMS = .TextMatrix(z, 10)
        zValorTotal = .TextMatrix(z, 11)
        xbasecalculo = .TextMatrix(z, 14)
        xValorIcms = .TextMatrix(z, 15)
        lcodigoaliquota = .TextMatrix(z, 16)
        lbl_nota = .TextMatrix(z, 17)
        zPorcentagemIPI = .TextMatrix(z, 9)
    End With
    
    With MovNotaFiscalEntrada
        lngPkCodigo = .ProximoPkCodigo(cDPEmpresa.codigo)
    
        .Empresa = cDPEmpresa.codigo
        .DataEntrada = EnviaData(msk_entrada)
        .CodigoDoFornecedor = txt_codigo_fornecedor
        .Numero = txt_numero_nf
        .Serie = txt_serie_nf
        .TipoDocumento = txt_codigo_forma
        .codigogrupo = zGrupo
        .CodigoProduto = zProduto
        .CodigoFornecedor = txt_codigo_fornecedor
        .Unidade = zUnidade
        .NomeProduto = zNome
        .Quantidade = zQuantidade
        .ValorBruto = zvalorbruto
        .PorcDesconto = zporc_desconto
        .ValorUnitario = zValorUnitario
        .ValorTotal = zValorTotal
        .PorcentagemIpi = zPorcentagemIPI
        .ValorIPI = zValorIPI
        .PercIcms = zValorICMS
        .BaseCalculoIcms = xbasecalculo
        .ValorIcms = xValorIcms
'        .PrecoCusto = g_string2
'        .PrecoCustoMedio = g_string3
'        .PrecoCustoMedioCont = g_string4
        .PrecoCusto = grade1.TextMatrix(z, 18)
        .PrecoCustoMedio = grade1.TextMatrix(z, 55)
        .PrecoCustoMedioCont = grade1.TextMatrix(z, 56)
        .CodigoAliquota = lcodigoaliquota
        .CodificacaoFiscal = grade1.TextMatrix(z, 28)
        .ReducaoIcms = grade1.TextMatrix(z, 12)
        .Outros = lbl_nota
        .InfluenciaEstoque = IIf((cDPFFaturamento.InfluenciaEstoque = 1), "S", "N")
        .GarantiaProduto = 0
        .BaseCalcSubstituicao = grade1.TextMatrix(z, 19)
        .IcmsSubstituicao = grade1.TextMatrix(z, 20)
        .OutrasDespesas = grade1.TextMatrix(z, 23)
        .ChecaLote = grade1.TextMatrix(z, 24)
        .IndustriaRevenda = grade1.TextMatrix(z, 25)
        .DescontoPe = grade1.TextMatrix(z, 27)
        .sequencia = lSequencia
        .AliqICMSSubst = x_perc_icms_subst
        .RedICMSSubst = lporcentagemreducao
        .CST = grade1.TextMatrix(z, 30)
        .RedICMS = lporcentagemreducao
        .PrecoPis = grade1.TextMatrix(z, 31)
        .PrecoCofins = grade1.TextMatrix(z, 32)
        .BCImportacao = grade1.TextMatrix(z, 34)
        .DespAduaneiras = grade1.TextMatrix(z, 35)
        .ImpostoImportacao = grade1.TextMatrix(z, 36)
        .ValorIOF = grade1.TextMatrix(z, 37)
        .codigoBarras = grade1.TextMatrix(z, 38)
        If IsNumeric(grade1.TextMatrix(z, 45)) Then
            .FatorConversao = grade1.TextMatrix(z, 45)
        Else
            .FatorConversao = 1
        End If
        .PkCodigo = lngPkCodigo
    End With

    If MovNotaFiscalEntrada.Incluir(ValidaBooleanoConferencia) Then AtualizaMovimentoNotaFiscalEntrada = True

    If Not ValidaBooleanoConferencia Then
        'CST IPI
        If CCur(grade1.TextMatrix(z, 13)) > 0 Then grade1.TextMatrix(z, 50) = "00" Else grade1.TextMatrix(z, 50) = "01"
        If Not AtualizaMovimentoImpostos("ENTRADA", lngPkCodigo, 0, 0, grade1.TextMatrix(z, 51), grade1.TextMatrix(z, 21), _
                                         grade1.TextMatrix(z, 52), grade1.TextMatrix(z, 22), grade1.TextMatrix(z, 48), _
                                         grade1.TextMatrix(z, 49), grade1.TextMatrix(z, 50), grade1.TextMatrix(z, 53)) Then Exit Function
    End If
    
Exit Function
Err_AtualizaMovimentoNotaFiscalEntrada: ValidaErros Err, Me.Caption & " - AtualizaMovimentoNotaFiscalEntrada"
End Function

'*****************************************************************************
'Criação: Diego Martins dos Santos                      Data: 19/06/2010
'
'Propósito:Atualiza movimento conferência
'*****************************************************************************
Private Function AtualizaMovimentoConferencia() As Boolean
On Error GoTo Err_AtualizaMovimentoConferencia
Dim z As Long

    AtualizaMovimentoConferencia = False
    
    If Not Conexao.DeleteSintetico("conferencia_mercadoria_entrada", "empresa = '" & cDPEmpresa.codigo & "' and sequencia_nf = " & lSequencia, cDPEmpresa.codigo) Then Exit Function
    
    For z = 1 To grade3.Rows - 1
        If Trim(grade3.TextMatrix(z, 0)) <> "" Then
            
            If Not Conexao.InserirRecordset("conferencia_mercadoria_entrada", "empresa,codigo_produto," & _
                                            "quantidade,codigo_usuario,sequencia_nf,ordem", _
                                            "'" & cDPEmpresa.codigo & "','" & grade3.TextMatrix(z, 0) & "'," & _
                                            fValidaValor2(grade3.TextMatrix(z, 4)) & ",'" & txt_codigo_usuario & "','" & _
                                            lSequencia & "','" & z & "'", cDPEmpresa.codigo) Then Exit Function
            
        End If
    Next
    
    Call Conexao.InserirRecordset("log_senhas", "data, hora, codigo_usuario, nome_usuario, historico, tela, observacoes, outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "','Entrada P/Conferência : " & msk_entrada & " Nº da Nota " & txt_numero_nf & " Forn. " & txt_fornecedor & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)
    
    AtualizaMovimentoConferencia = True

Exit Function
Err_AtualizaMovimentoConferencia: ValidaErros Err, Me.Caption & " - AtualizaMovimentoConferencia"
End Function

'==========================================================================
' Purpose:  Efetuar validação para verificar se os calculos do sistema estão
'           batendo com os lançados da nota
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start:
' Alteração: Ronaldo Robledo                                    01/08/2012
'           Inserido cálculo para ratear a diferença entre os itens da nota
'           para bater o valor final
'==========================================================================
Private Function Verificatotais()
On Error GoTo Verificatotais

Dim xBCIcms As Currency
Dim xICMS As Currency
Dim xIPI As Currency
Dim xBCIcmsSubstituicao As Currency
Dim xIcmsSubstituicao As Currency
Dim xFrete As Currency
Dim xSeguro As Currency
Dim xoutras As Currency
Dim xTotal As Currency

    Verificatotais = True
    xBCIcms = CDbl(txt_bc_icms) - CDbl(lbl_bc_icms)
    xICMS = CDbl(txt_icms) - CDbl(lbl_valor_icms)
    xIPI = CDbl(txt_ipi) - CDbl(lbl_ipi)
    xIcmsSubstituicao = CDbl(txt_substituicao) - CDbl(lbl_icms_substituicao)
    xBCIcmsSubstituicao = CDbl(lbl_bc_substituicao) - CDbl(txt_bc_substituicao)
    xFrete = CDbl(txt_frete) - CDbl(lbl_frete)
    xSeguro = CDbl(txt_seguro) - CDbl(lbl_seguro)
    xoutras = CDbl(txt_outras) - CDbl(lbl_outras_despesas)
    xTotal = CDbl(txt_total) - CDbl(lbl_total)
   
   If CDbl(lbl_bc_icms) < 0 Then
        Verificatotais = False
        Alerta "O Valor da Base de Cálculo do Icms está incorreto!"
   End If
   
   If CDbl(lbl_valor_icms) < 0 Then
        Verificatotais = False
        Alerta "O Valor do Icms está incorreto!"
   End If

   If CDbl(lbl_ipi) < 0 Then
        Verificatotais = False
        Alerta "O Valor do IPI está incorreto!"
   End If

   If Val(lbl_icms_substituicao) < 0 Then
        Verificatotais = False
        Alerta "O Valor do Icms Substituição está incorreto!"
   End If

   If Val(lbl_bc_substituicao) < 0 Then
        Verificatotais = False
        Alerta "O Valor da Base de Cálculo Substituição está incorreto!"
   End If

   If Val(lbl_frete) < 0 Then
        Verificatotais = False
        Alerta "O Valor do Frete está incorreto!"
   End If

   If Val(lbl_seguro) < 0 Then
        Verificatotais = False
        Alerta "O Valor do Seguro está incorreto!"
   End If

   If Val(lbl_outras_despesas) < 0 Then
        Verificatotais = False
        Alerta "O Valor de Outras Despesas está incorreto!"
   End If

   If Val(lbl_total) < 0 Then
        Verificatotais = False
        Alerta "O Valor Total está incorreto!"
   End If

    If xTotal < -0.09 Then
        Verificatotais = False
        Alerta "O Valor Total está incorreto!"
    ElseIf xTotal > 0.09 Then
        Verificatotais = False
        Alerta "O Valor Total está incorreto!"
    Else
       If xTotal <> 0 Then Call Rateio(xTotal, 11, txt_total)
       lbl_total = Format(txt_total, "##,###,##0.00")
    End If
  
    If CDbl(txt_bc_icms) = 0 Then
        lbl_bc_icms = Format(txt_bc_icms, "##,###,##0.00")
    ElseIf xBCIcms < -0.15 Then
        Verificatotais = False
        Alerta "O Valor da Base de Cálculo do Icms está incorreto!"
    ElseIf xBCIcms > 0.15 Then
        Verificatotais = False
        Alerta "O Valor da Base de Cálculo do Icms está incorreto!"
    Else
        If xBCIcms <> 0 Then Call Rateio(xBCIcms, 14, txt_bc_icms)
        lbl_bc_icms = Format(txt_bc_icms, "##,###,##0.00")
    End If
    
    If xICMS < -0.09 Then
        Verificatotais = False
        Alerta "O Valor do Icms está incorreto!"
    ElseIf xICMS > 0.09 Then
        Verificatotais = False
        Alerta "O Valor do Icms está incorreto!"
    Else
        If xICMS <> 0 Then Call Rateio(xICMS, 15, txt_icms)
        lbl_valor_icms = Format(txt_icms, "##,###,##0.00")
    End If
   
    If xIPI < -0.09 Then
        Verificatotais = False
        Alerta "O Valor do IPI está incorreto!"
    ElseIf xIPI > 0.09 Then
        Verificatotais = False
        Alerta "O Valor do IPI está incorreto!"
    Else
        If xIPI <> 0 Then Call Rateio(xIPI, 13, txt_ipi)
        lbl_ipi = Format(txt_ipi, "##,###,##0.00")
    End If
   
    If CDbl(txt_bc_substituicao) <> 0 Then
        If xBCIcmsSubstituicao < -0.09 Then
            Verificatotais = False
            Alerta "O Valor da Base de Calculo Substituicao Icms está incorreto!"
        ElseIf xBCIcmsSubstituicao > 0.09 Then
            Verificatotais = False
            Alerta "O Valor da Base de Calculo Substituicao Icms está incorreto!"
        Else
            If xBCIcmsSubstituicao <> 0 Then Call Rateio(xBCIcmsSubstituicao, 19, txt_bc_substituicao)
            lbl_bc_substituicao = Format(txt_bc_substituicao, "##,###,##0.00")
        End If
    
        If xIcmsSubstituicao < -0.09 Then
            Verificatotais = False
            Alerta "O Valor do Substituicao Icms está incorreto!"
        ElseIf xIcmsSubstituicao > 0.09 Then
            Verificatotais = False
            Alerta "O Valor do Substituicao Icms está incorreto!"
        Else
            If xIcmsSubstituicao <> 0 Then Call Rateio(xIcmsSubstituicao, 20, txt_substituicao)
            lbl_icms_substituicao = Format(txt_substituicao, "##,###,##0.00")
        End If
    End If
    
    If xFrete < -0.09 Then
        Verificatotais = False
        Alerta "O Valor do Frete está incorreto!"
    ElseIf xFrete > 0.09 Then
        Verificatotais = False
        Alerta "O Valor do Frete está incorreto!"
    Else
        lbl_frete = Format(txt_frete, "##,###,##0.00")
    End If
        
    If xSeguro < -0.09 Then
        Verificatotais = False
        Alerta "O Valor do Seguro está incorreto!"
    ElseIf xSeguro > 0.09 Then
        Verificatotais = False
        Alerta "O Valor do Seguro está incorreto!"
    Else
        lbl_seguro = Format(txt_seguro, "##,###,##0.00")
    End If
   
    If xoutras < -0.09 Then
        Verificatotais = False
        Alerta "O Valor de Outras Despesas está incorreto!"
    ElseIf xoutras > 0.09 Then
        Verificatotais = False
        Alerta "O Valor de Outras Despesas está incorreto!"
    Else
        lbl_outras_despesas = Format(txt_outras, "##,###,##0.00")
    End If
   
    If Verificatotais = False Then
        Sql_Record_Set.Source = "SELECT * FROM menu WHERE nome = 'Entrada (Sem Cálculo)' and usuario = '" & g_usuario & "'"
        Sql_Record_Set.Open
        If Sql_Record_Set.RecordCount > 0 Then
            Sql_Record_Set.Close
            If (Confirma("Valores Totais Não Confere com o Cálculo do Sistema, Deseja Continuar ?")) = 6 Then
                Alerta "ATENÇÃO NOTAS NÃO TERÃO CÁLCULO CORRETO PARA O SINTEGRA NEM NOTA FISCAL ELETRONICA!"
                Senha.lbl_serie = "NF"
                Senha.lbl_titulo = "Liberação para Entrada de Nota Sem Cálculo"
                Senha.lbl_liberacao = "Entrada (Sem Cálculo)"
                Senha.lbl_tela = "Entrada de Mercadorias"
                Senha.lbl_historico = "Entrada de Nota Sem Cálculo N.: " & txt_numero_nf & " Forn. " & txt_fornecedor
                Senha.Show (1)
                If g_string = "OK" Then
                    lbl_bc_icms = txt_bc_icms
                    lbl_valor_icms = txt_icms
                    lbl_ipi = txt_ipi
                    lbl_icms_substituicao = txt_substituicao
                    lbl_bc_substituicao = txt_bc_substituicao
                    lbl_frete = txt_frete
                    lbl_seguro = txt_seguro
                    lbl_outras_despesas = txt_outras
                    lbl_total = txt_total
                    Verificatotais = True
                End If
            End If
        Else
            Sql_Record_Set.Close
        End If
   End If

Exit Function
Verificatotais: ValidaErros Err, Me.Caption & " - Verificatotais"
End Function
'*****************************************************************************
'Criação: Diego Martins dos Santos                            Data: 21/06/2010
'Propósito:Verifica Qtde total lançado na conferência
'*****************************************************************************
Private Sub VerificaQtdeTotalConferencia()
On Error GoTo Err_VerificaQtdeTotalConferencia

    Sql_Query = "SELECT sum(MNFE.quantidade) as QtdeNota,sum(CFE.quantidade) as QtdeConferencia " & _
                "FROM movimento_nota_fiscal_entrada_tmp MNFE " & _
                "INNER JOIN conferencia_mercadoria_entrada CFE " & _
                "ON CFE.empresa = MNFE.empresa and CFE.sequencia_nf = MNFE.sequencia " & _
                "WHERE MNFE.empresa = '" & cDPEmpresa.codigo & "' and MNFE.sequencia = '" & lSequencia & "' "
    Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 0)
    If gb_Recordset.RecordCount > 0 Then
       If IsNumeric(gb_Recordset!QtdeConferencia) Then
            If CCur(gb_Recordset!QtdeNota) <> CCur(gb_Recordset!QtdeConferencia) Then
               gb_Recordset.Close
               Exit Sub
            Else
                Toolbar1.Buttons(6).Enabled = True
            End If
       End If
    End If
    gb_Recordset.Close
Exit Sub
Err_VerificaQtdeTotalConferencia: ValidaErros Err, Me.Caption & " - VerificaQtdeTotalConferencia"
End Sub
'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: João Batista                                Start:
' Alteração: João Batista: Solução TICKET TT3052      Data:10/04/2013
' Correção: Inserido modelo 99 poque não existe regra de negocio que verifique modelso nf validos. ou seja não existe modelos
'     fixos tanto ao SPED Quanto a notas de serviço.
'==========================================================================
Private Function ValidaCampos()
On Error GoTo ValidaCampo
Dim lngQtdDigitos   As Long
Dim z               As Long
ValidaCampos = False

       lngQtdDigitos = Len(txt_chave_acesso.Text)
     
    If grade1.TextMatrix(grade1.Rows - 1, 0) = NovaLinha And Trim(grade1.TextMatrix(grade1.Rows - 1, 3)) <> "" Then
         Call ChamaCelula
    End If
        
    If Not IsNumeric(txt_codigo_fornecedor) Then
        SSTab1.Tab = 0
        Alerta "Digite o Código do Fornecedor!"
        txt_codigo_fornecedor.SetFocus
    ElseIf Not IsNumeric(txt_codigo_transportadora) Then
        Alerta "Digite o Código da Transportadora!"
    ElseIf Not ExisteFornecedor Then
        SSTab1.Tab = 0
        Alerta "Fornecedor não Cadastrado ou Inativado no Plano de Contas!"
        txt_codigo_fornecedor.SetFocus
    ElseIf txt_modelo_nf.Text = "55" And lngQtdDigitos < 44 And chk_impressao_nf.Value = 0 Then
         Alerta "Por favor, preencha a chave de acesso da NF-e com 44 digitos!"
         txt_chave_acesso.SetFocus
         Exit Function
    ElseIf Not IsNumeric(txt_numero_nf) And chk_impressao_nf.Value = 0 Then
        SSTab1.Tab = 0
        Alerta "Digite o numero da Nota Fiscal!"
        txt_numero_nf.SetFocus
        Exit Function
    ElseIf Trim(txt_serie_nf) = "" Then
        SSTab1.Tab = 0
        Alerta "Digite a Série da Nota Fiscal!"
        txt_serie_nf.SetFocus
    ElseIf Not IsNumeric(txt_modelo_nf) Then
        SSTab1.Tab = 0
        Alerta "Digite o Modelo da Nota Fiscal!"
        txt_modelo_nf.SetFocus
    '*** Alteração Diego : Solução OS - 12743,12879 ENTRADA NF MODELO 2,4 ***
    ElseIf Trim(txt_modelo_nf) <> "1" And Trim(txt_modelo_nf) <> "2" And Trim(txt_modelo_nf) <> "3" And Trim(txt_modelo_nf) <> "4" And Trim(txt_modelo_nf) <> "6" And Trim(txt_modelo_nf) <> "21" And Trim(txt_modelo_nf) <> "22" And Trim(txt_modelo_nf) <> "55" And Trim(txt_modelo_nf) <> "99" Then
        SSTab1.Tab = 0
        
        Alerta "Modelo da Nota Fiscal Inválida!"
        txt_modelo_nf.SetFocus
    ElseIf Not IsDate(msk_emissao) Then
        SSTab1.Tab = 0
        
        Alerta "Data de Emissao Inválida!"
        msk_emissao.SetFocus
    ElseIf CDate(msk_emissao) > Date Then
        SSTab1.Tab = 0
        
        Alerta "Data de Emissao deve ser menor ou igual a " & Date & "!"
        msk_emissao.SetFocus
    ElseIf Not IsDate(msk_entrada) Then
        SSTab1.Tab = 0
        
        Alerta "Data de Entrada Inválida!"
        msk_entrada.SetFocus
    ElseIf Not IsDate(msk_entrada) <> Date Then
        SSTab1.Tab = 0
        
        Alerta "Data de Entrada deve ser igual a " & Date & "!"
        msk_entrada.SetFocus
    ElseIf Not IsNumeric(txt_bc_icms) Then
        SSTab1.Tab = 0
        
        Alerta "Digite a Base de Calculo do ICMS!"
        txt_bc_icms.SetFocus
    ElseIf Not IsNumeric(txt_icms) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor do ICMS! "
        txt_icms.SetFocus
    ElseIf Not IsNumeric(txt_bc_substituicao) Then
        SSTab1.Tab = 0
        Alerta "Digite a Base de Calculo Substituicao ICMS! "
        txt_bc_substituicao.SetFocus
    ElseIf Not IsNumeric(txt_substituicao) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor da Substituição do ICMS! "
        txt_substituicao.SetFocus
    ElseIf CDbl(txt_substituicao) < 0 Then
        SSTab1.Tab = 0
        Alerta "Valor da Substituição do ICMS Inválido! "
        txt_substituicao.SetFocus
    ElseIf Not IsNumeric(txt_ipi) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor do IPI ! "
        txt_ipi.SetFocus
    ElseIf Not IsNumeric(txt_frete) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor do Frete ! "
        txt_frete.SetFocus
    ElseIf Not IsNumeric(txt_seguro) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor do Seguro ! "
        txt_seguro.SetFocus
    ElseIf Not IsNumeric(txt_outras) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor de Outras Despesas ! "
        txt_outras.SetFocus
    ElseIf Not IsNumeric(txt_desconto) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor do Desconto! "
        txt_desconto.SetFocus
    ElseIf Not IsNumeric(txt_porc_red_icms) Then
        SSTab1.Tab = 0
        Alerta "Porcentagem do ICMS Redução inválido! "
     
    ElseIf Trim(txt_observacoes) = "" And chk_ContaConsumo.Value = 1 Then
        SSTab1.Tab = 0
        Alerta "SPED: Obrigatório o preenchimento das observações em notas de consumo!", 48
     
    ElseIf Not IsNumeric(txt_volume) Then
        SSTab1.Tab = 0
        Alerta "Quantidade do Volume inválida! "
    ElseIf Not IsNumeric(txt_peso_bruto) Then
        SSTab1.Tab = 0
        Alerta "Peso Bruto Inválido! "
    ElseIf Not IsNumeric(txt_peso_liquido) Then
        SSTab1.Tab = 0
        Alerta "Peso Líquido inválido! "
    ElseIf Not IsNumeric(txt_total) Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor Total! "
        txt_total.SetFocus
    ElseIf Not IsNumeric(txt_frete_conhecimento) Then
        SSTab1.Tab = 0
        Alerta "Valor do Frete de Conhecimento Inválido! "
        txt_frete_conhecimento.SetFocus
    ElseIf CDbl(txt_total) <= 0 And chk_calculo_nota.Value = 0 Then
        SSTab1.Tab = 0
        Alerta "Digite o Valor Total! "
        txt_total.SetFocus
    ElseIf Not IsNumeric(txt_codigo_forma) Then
        Alerta "Forma de Pagamento Inválida!"
        txt_codigo_forma.SetFocus
    ElseIf Not Verificacodigo Then
        SSTab1.Tab = 0
        grade1.SetFocus
    ElseIf Not ExisteTitulo(txt_codigo_forma, "('A','E')") Then
        SSTab1.Tab = 0
        txt_codigo_forma.SetFocus
        Exit Function
    ElseIf Calculaporcentagem Then
        SSTab1.Tab = 0
        Alerta "Porcentagem do IPI ou ICMS Inválido!"
        grade1.SetFocus
    ElseIf VerificaPedido Then
        SSTab1.Tab = 0
        grade1.SetFocus
    ElseIf ExisteProdutoII Then
        SSTab1.Tab = 0
        Alerta "Existe Produto Duplicado na Grade!"
        grade1.SetFocus
    ElseIf ValidaCodigoBarras(z) Then
        Alerta "Informe o Código de Barras do Produto Cod. " & grade1.TextMatrix(z, 2)
        grade1.TextMatrix(z, 38) = ObtenhaCodigoBarras(grade1.TextMatrix(z, 2))
        SSTab1.Tab = 0
        grade1.SetFocus
    ElseIf Not validagrade Then
        SSTab1.Tab = 0
        grade1.SetFocus
    ElseIf Not validacamposii Then
        Exit Function
    ElseIf ValidaBloqueioMovimentacao Then
        Exit Function
    End If

    If chk_impressao_nf.Value = 1 And pct_orçamento.Visible = False And cDPEmpresa.NotaFiscalEletronica > 0 Then
        txt_numero_nf = 0
        If Not VerificaSeExisteNFePendentes Then ValidaCampos = True Else Exit Function
    Else
        ValidaCampos = True
    End If
      
Exit Function
ValidaCampo: ValidaErros Err, Me.Caption & " - ValidaCampo"
End Function

Function ExisteNF()
On Error GoTo ExisteNota

    ExisteNF = False
  
    If chk_impressao_nf.Value = 0 And Val(txt_numero_nf) > 0 Then
        Set gb_Recordset = Conexao.GeraRecordset("SELECT empresa FROM movimento_cabecalho_nota_fiscal_entrada WHERE data_de_emissao = " & FormataData(msk_emissao) & " and numero = '" & txt_numero_nf & "' and serie = '" & txt_serie_nf & "' and modelo_nf = '" & txt_modelo_nf & "' and codigo_do_fornecedor = " & txt_codigo_fornecedor & " and codificacao_fiscal = '" & lbl_cfop & "'", 1)
        If gb_Recordset.RecordCount > 0 Then
           ExisteNF = True
           Alerta "Nota Fiscal já Cadastrada!", vbInformation
        End If
        gb_Recordset.Close
    End If

Exit Function
ExisteNota: ValidaErros Err, Me.Caption & " - ExisteNota"
End Function

Function ExisteFornecedor()
ExisteFornecedor = False
         
    Set gb_Recordset = Conexao.GeraRecordset("SELECT FO.*,PL.inativo FROM fornecedor FO,plano_conta PL WHERE FO.codigo = " & txt_codigo_fornecedor & " and conta = FO.conta_contabil", 1)
    If gb_Recordset.RecordCount > 0 Then
        If gb_Recordset!Inativo = 0 Then ExisteFornecedor = True
    End If
    gb_Recordset.Close
    
End Function

Private Sub txt_fornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_codigo_fornecedor.SetFocus
End Sub

Private Sub txt_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_numero_nf.SetFocus
End Sub

Private Sub txt_fornecedor_LostFocus()
    txt_fornecedor.BackColor = &H8000000E
End Sub

Private Sub Form_Activate()
On Error GoTo AtivandoForm
Dim strFiltro As String

    Movimento_Nota_Fiscal_Entrada.Caption = "Movimento Nota Fiscal de Entrada "
    
    If flag_tela_entrada_mercadoria = 0 Then
        SSTab1.Tab = 0
        
        pct_orçamento.Visible = False
        pct_icms.Visible = False
        LimpaTela
             
            strFiltro = ""
            If bolCentralNFe Then strFiltro = " AND sequencia = '" & g_string2 & "' "
            
            Sql_Query = "SELECT MCNFE.*,(INCC.codigo) as PkCodigoConsumo, " & _
                               "(INCC.tipo_conta) as TipoConta,(INCC.codigo_consumo) as CodigoConsumo," & _
                               "(INCC.tipo_ligacao) as TipoLigacao,(INCC.grupo_tensao) as GrupoTensao,ML.seq_controle " & _
                        "FROM movimento_cabecalho_nota_fiscal_entrada MCNFE " & _
                        "LEFT JOIN movimento_lote ML ON ML.empresa = MCNFE.empresa and ML.sequencia = MCNFE.sequencia and ML.status = 'E'" & _
                        "LEFT JOIN info_conta_consumo INCC ON INCC.empresa = MCNFE.empresa AND " & _
                        "INCC.codigo = MCNFE.fk_codigo_conta_consumo " & _
                        "WHERE MCNFE.serie <> 'CX' " & strFiltro & _
                        " ORDER BY MCNFE.data_de_entrada DESC, MCNFE.numero  DESC"
            Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 1)
            If gb_Recordset.RecordCount > 0 Then
                AtivaBotoes
                AtualTela
                bolEntradaTMP = False
                AtualizaGrid (0)
                Movimento_Nota_Fiscal_Entrada.SetFocus
                If bolCentralNFe Then
                   Toolbar1.Enabled = False
                   Call EfetuaCancelamentoNF
                End If
            Else
                gb_Recordset.Close
                DesativaBotoes
                Toolbar1.Buttons(1).Enabled = True
                Toolbar1.Buttons(7).Enabled = False
                Toolbar1.Buttons(8).Enabled = False
                Toolbar1.Buttons(13).Enabled = True
                Toolbar1.Buttons(12).Enabled = True
                Toolbar1.Buttons(4).Enabled = True
                frmdados.Enabled = False
            End If
           flag_tela_entrada_mercadoria = 1
           
     End If

Exit Sub
AtivandoForm: ValidaErros Err, Me.Caption & " - AtivandoForm"
End Sub

'*****************************************************************************
'Criação: Diego Martins dos Santos                            Data: 19/06/2010
'Propósito:Atualiza Informações na grid
'Alteração: Ronaldo Robledo                                         06/06/2013
'           Inserido chamada de método PreencheColecaoProdutos
'*****************************************************************************
Private Sub AtualizaGrid(ByVal bytOpcao As Byte)
On Error GoTo Err_AtualizaGrid
Dim strTabela As String
    
    '*** Consulta de Nota Normal ***
    If bytOpcao = 0 Then
        strTabela = "movimento_nota_fiscal_entrada"
    '*** Consulta de Nota Pela conferência ***
    'Adicionado PD.utiliza_grade solução 13817 Autor:Diego Martins
    ElseIf bytOpcao = 1 Then
        strTabela = "movimento_nota_fiscal_entrada_tmp"
    End If
    
    Sql_Query = "SELECT MNFE.codigo_do_grupo,MNFE.codigo_do_produto,MNFE.nome_do_produto,MNFE.unidade, " & _
                "(MNFE.quantidade / MNFE.Fator_conversao) as quantidade,MNFE.valor_bruto, MNFE.porc_desconto, " & _
                "MNFE.valor_unitario, MNFE.porcentagem_do_ipi, " & _
                "MNFE.perc_icms,MNFE.valor_total,MNFE.reducao_icms,MNFE.valor_do_ipi,MNFE.base_calculo_icms," & _
                "MNFE.valor_icms,MNFE.codigo_da_aliquota,MNFE.outros,MNFE.preco_custo,MNFE.base_calculo_subst," & _
                "MNFE.icms_substituicao,MI.percentual_pis,MI.percentual_cofins,MNFE.outras_despesas,MNFE.checa_lote," & _
                "MNFE.industria_revenda,PD.fracionado,MNFE.desconto_pe,MNFE.codificacao_fiscal,MNFE.codigo_do_produto," & _
                "MNFE.cst,MNFE.preco_pis,MNFE.preco_cofins,0,MNFE.base_calculo_importacao," & _
                "MNFE.despesas_aduaneiras,MNFE.imposto_importacao,MNFE.valor_iof,MNFE.codigo_barras,PD.utiliza_grade," & _
                "PD.codificacao_fiscal,0,0,'',0,'',PD.subcodigo,PD.fk_cadastrotributacao,MI.cst_pis,MI.cst_cofins,0,0,0,0, " & _
                "TIM.pk_item_movimentacao,MNFE.preco_custo_medio,MNFE.preco_custo_medio_cont " & _
                "FROM " & strTabela & " MNFE " & _
                "INNER JOIN produto PD ON PD.codigo = MNFE.codigo_do_produto " & _
                "LEFT JOIN movimento_impostos MI ON MI.pk_empresa = MNFE.empresa and MI.fk_movnotafiscalentrada = MNFE.pk_codigo " & _
                "LEFT JOIN tb_pedido_compra TPC ON TPC.fk_empresa = MNFE.empresa and TPC.numeropedido = MNFE.codigo_do_grupo " & _
                "LEFT JOIN tb_item_movimentacao TIM ON TIM.fk_empresa = TPC.fk_empresa and TIM.fk_movimentacao = TPC.fk_movimentacao " & _
                "and TIM.fk_produto = MNFE.codigo_do_produto " & _
                "WHERE empresa = '" & cDPEmpresa.codigo & "' and sequencia = '" & lSequencia & "'"
    
    
    Adodc1.RecordSource = Sql_Query
    Adodc1.Refresh
    
    Set grade1.DataSource = Adodc1
    FormaGrid
    
    PreencheColecaoProdutos

Exit Sub
Err_AtualizaGrid: ValidaErros Err, Me.Caption & " - AtualizaGrid"
End Sub

Private Sub FormaGrid()
On Error GoTo Err_FormaGrid
Dim z As Long

    grade1.GridLines = flexGridFlat
    grade1.FixedCols = 1
    grade1.AllowUserResizing = flexResizeColumns
    
    grade1.Cols = 57
    grade1.TextMatrix(0, 0) = "Item"
    grade1.ColWidth(0) = 20
    'grade1.ColAlignmentFixed(0) = 4
    grade1.ColAlignment(0) = flexAlignLeftCenter
    
    grade1.TextMatrix(0, 1) = "Pedido"
    grade1.ColWidth(1) = 630
    'grade1.ColAlignmentFixed(1) = 4
    grade1.ColAlignment(1) = flexAlignLeftCenter
    
    grade1.TextMatrix(0, 2) = "Codigo"
    grade1.ColWidth(2) = 855
    'grade1.ColAlignmentFixed(2) = 4
    grade1.ColAlignment(2) = flexAlignCenterCenter
    
    grade1.TextMatrix(0, 3) = "Descrição"
    grade1.ColWidth(3) = 2900
    'grade1.ColAlignmentFixed(3) = 4
    grade1.ColAlignment(3) = flexAlignLeftCenter
    
    grade1.TextMatrix(0, 4) = "UND"
    grade1.ColWidth(4) = 435
    'grade1.ColAlignmentFixed(4) = 4
    grade1.ColAlignment(4) = flexAlignRightCenter
    
    grade1.TextMatrix(0, 5) = "Qtde."
    grade1.ColWidth(5) = 915
    'grade1.ColAlignmentFixed(5) = 4
    grade1.ColAlignment(5) = flexAlignRightCenter
    
    grade1.TextMatrix(0, 6) = "Vlr. Bruto"
    grade1.ColWidth(6) = 1140
    'grade1.ColAlignmentFixed(6) = 4
    grade1.ColAlignment(6) = flexAlignRightCenter
    
    grade1.TextMatrix(0, 7) = "Desc"
    grade1.ColWidth(7) = 525
    'grade1.ColAlignmentFixed(7) = 4
    grade1.ColAlignment(7) = flexAlignRightCenter
    
    grade1.TextMatrix(0, 8) = "Vlr. Unitario"
    grade1.ColWidth(8) = 1230
    'grade1.ColAlignmentFixed(8) = 4
    grade1.ColAlignment(8) = flexAlignRightCenter
    
    grade1.TextMatrix(0, 9) = "% IPI"
    grade1.ColWidth(9) = 510
    'grade1.ColAlignmentFixed(9) = 4
    grade1.ColAlignment(9) = flexAlignRightCenter
    
    grade1.TextMatrix(0, 10) = "% ICMS"
    grade1.ColWidth(10) = 720
    'grade1.ColAlignmentFixed(10) = 4
    grade1.ColAlignment(10) = flexAlignRightCenter
    
    'SUBTOTAL
    grade1.TextMatrix(0, 11) = "Sub Total"
    grade1.ColWidth(11) = 1365
    'grade1.ColAlignmentFixed(11) = 4
    grade1.ColAlignment(11) = flexAlignRightCenter
    
    'REDUCAO ICMS
    grade1.TextMatrix(0, 12) = "R.I"
    grade1.ColWidth(12) = 470
    'grade1.ColAlignmentFixed(12) = 4
    grade1.ColAlignment(12) = flexAlignLeftCenter
    
    'valor ipi
    grade1.TextMatrix(0, 13) = "Vlr. IPI"
    grade1.ColWidth(13) = 1200
    'grade1.ColAlignmentFixed(13) = 4
    grade1.ColAlignment(13) = flexAlignRightCenter
    
    'base calc. icms
    grade1.TextMatrix(0, 14) = "B.C.ICMS"
    grade1.ColWidth(14) = 1200
    'grade1.ColAlignmentFixed(14) = 4
    grade1.ColAlignment(14) = flexAlignRightCenter
    
    'valor do icms
    grade1.TextMatrix(0, 15) = "Vlr.ICMS"
    grade1.ColWidth(15) = 1200
    'grade1.ColAlignmentFixed(15) = 4
    grade1.ColAlignment(15) = flexAlignRightCenter
    
    'aliquota
    grade1.TextMatrix(0, 16) = "Aliq."
    grade1.ColWidth(16) = 800
    'grade1.ColAlignmentFixed(16) = 4
    grade1.ColAlignment(16) = flexAlignRightCenter
    
    'outros
    grade1.TextMatrix(0, 17) = "Outros"
    grade1.ColWidth(17) = 800
    'grade1.ColAlignmentFixed(17) = 4
    grade1.ColAlignment(17) = flexAlignCenterCenter
    
    'preco_custo
    grade1.TextMatrix(0, 18) = "Custo"
    grade1.ColWidth(18) = 1200
    grade1.ColAlignment(18) = flexAlignRightCenter
    
    'bcsubstituicao
    grade1.TextMatrix(0, 19) = "Bc.Subst."
    grade1.ColWidth(19) = 1200
    grade1.ColAlignment(19) = flexAlignRightCenter
    
    'xvlrsubstituicao
    grade1.TextMatrix(0, 20) = "Vlr.Subst."
    grade1.ColWidth(20) = 1200
    grade1.ColAlignment(20) = flexAlignRightCenter
    
    'pis %
    grade1.TextMatrix(0, 21) = "Pis"
    grade1.ColWidth(21) = 1
    grade1.ColAlignment(21) = flexAlignCenterCenter
    
    'cofins %
    grade1.TextMatrix(0, 22) = "Cofins"
    grade1.ColWidth(22) = 1
    grade1.ColAlignment(22) = flexAlignRightCenter
    
    'outras
    grade1.TextMatrix(0, 23) = "Vlr.Outras"
    grade1.ColWidth(23) = 1200
    grade1.ColAlignment(23) = flexAlignRightCenter
    
    'checa lote
    grade1.ColWidth(24) = 500
    grade1.ColAlignment(24) = flexAlignCenterCenter
    
    'industria_revenda
    grade1.TextMatrix(0, 25) = "I./R."
    grade1.ColWidth(25) = 500
    grade1.ColAlignment(25) = flexAlignCenterCenter
    
    'fracionado
    grade1.ColWidth(26) = 1
    grade1.ColAlignment(26) = flexAlignCenterCenter
    
    'desconto pe
    grade1.ColWidth(27) = 1
    grade1.ColAlignment(27) = flexAlignCenterCenter
    
    'CFOP
    grade1.TextMatrix(0, 28) = "CFOP"
    grade1.ColWidth(28) = 1200
    grade1.ColAlignment(28) = flexAlignCenterCenter
    
    'CODIGO NOVAMENTE
    grade1.TextMatrix(0, 29) = "Código"
    grade1.ColWidth(29) = 1
    grade1.ColAlignment(29) = flexAlignLeftCenter
    
    'CST
    grade1.TextMatrix(0, 30) = "CST"
    grade1.ColWidth(30) = 600
    grade1.ColAlignment(30) = flexAlignRightCenter
    
    'Preço PIS
    grade1.TextMatrix(0, 31) = "VLR.PIS"
    grade1.ColWidth(31) = 1200
    grade1.ColAlignment(31) = flexAlignRightCenter
    
    'Preço Cofins
    grade1.TextMatrix(0, 32) = "VLR COFINS"
    grade1.ColWidth(32) = 1200
    grade1.ColAlignment(32) = flexAlignRightCenter
    
    'Tipo Importacao 'perdeu sua utilidade depois da propriedade tributacao
    grade1.TextMatrix(0, 33) = "T.Imp"
    grade1.ColWidth(33) = 500
    grade1.ColAlignment(33) = flexAlignRightCenter
    
    'Base de calculo ii
    grade1.TextMatrix(0, 34) = "BC II"
    grade1.ColWidth(34) = 700
    grade1.ColAlignment(34) = flexAlignRightCenter
    
    'valor DespAduneiras
    grade1.TextMatrix(0, 35) = "Desp. Adu. II"
    grade1.ColWidth(35) = 700
    grade1.ColAlignment(35) = flexAlignRightCenter
    
    'valor II
    grade1.TextMatrix(0, 36) = "II"
    grade1.ColWidth(36) = 700
    grade1.ColAlignment(36) = flexAlignRightCenter
    
    'valor iof
    grade1.TextMatrix(0, 37) = "IOF II"
    grade1.ColWidth(37) = 700
    grade1.ColAlignment(37) = flexAlignRightCenter
    
    'Código Barras
    grade1.TextMatrix(0, 38) = "Código Barras"
    grade1.ColWidth(38) = 700
    grade1.ColAlignment(38) = flexAlignRightCenter
    
    'UtilizaGrade
    grade1.TextMatrix(0, 39) = "UtilizaGrade"
    grade1.ColWidth(39) = 700
    grade1.ColAlignment(39) = flexAlignRightCenter
    
    'Classificação Fiscal
    grade1.TextMatrix(0, 40) = "Clas.Fiscal"
    grade1.ColWidth(40) = 1000
    grade1.ColAlignment(40) = flexAlignRightCenter
    
    'CSOSN sem utilidade depois do cad prop tributacao
    grade1.TextMatrix(0, 41) = "CSOSN"
    grade1.ColWidth(41) = 1
    grade1.ColAlignment(41) = flexAlignRightCenter
    
    '% IVA
    grade1.TextMatrix(0, 42) = "% IVA"
    grade1.ColWidth(42) = 800
    grade1.ColAlignment(42) = flexAlignRightCenter
    
    'Verificar se foi digitado o cst
    grade1.TextMatrix(0, 43) = ""
    grade1.ColWidth(43) = 0
    grade1.ColAlignment(43) = flexAlignRightCenter
    
    'Estoque novamente caso de unidade secundaria
    grade1.ColWidth(44) = 1
    grade1.ColAlignment(44) = flexAlignCenterCenter
    
    'fator de operacao unidade secundaria
    grade1.ColWidth(45) = 1
    grade1.ColAlignment(45) = flexAlignCenterCenter
    
    'subcodigo
    grade1.ColWidth(46) = 1
    grade1.ColAlignment(46) = flexAlignCenterCenter
    
    'IDCadastro Tributacao
    grade1.ColWidth(47) = 1
    grade1.ColAlignment(47) = flexAlignCenterCenter

    'cst pis
    grade1.ColWidth(48) = 1
    grade1.ColAlignment(48) = flexAlignCenterCenter
    
    'cst cofins
    grade1.ColWidth(49) = 1
    grade1.ColAlignment(49) = flexAlignCenterCenter

    'cst IPI
    grade1.ColWidth(50) = 1
    grade1.ColAlignment(50) = flexAlignCenterCenter

    'base calc. PIS
    grade1.ColWidth(51) = 1
    grade1.ColAlignment(51) = flexAlignCenterCenter
    
    'base calc cofins
    grade1.ColWidth(52) = 1
    grade1.ColAlignment(52) = flexAlignCenterCenter

    'base calc IPI
    grade1.ColWidth(53) = 1
    grade1.ColAlignment(53) = flexAlignCenterCenter

    'ID da entidade tb_item_movimentacao '
    'para poder trabalhar com os pedidos de mercadoria
    grade1.ColWidth(54) = 1
    grade1.ColAlignment(54) = flexAlignCenterCenter
    
    'custo médio
    grade1.ColWidth(55) = 1
    grade1.ColAlignment(55) = flexAlignCenterCenter

    'custo médio cont
    grade1.ColWidth(56) = 1
    grade1.ColAlignment(56) = flexAlignCenterCenter

    
    For z = grade1.FixedRows To grade1.Rows - 1
        grade1.TextMatrix(z, 0) = z
        '*** Autor: João Batista Data: 01/06/2012 Motivo: Sistema não formatava corretamente coluna qtde. ***
        grade1.TextMatrix(z, 5) = Format(fValidaValorNovo(grade1.TextMatrix(z, 5)), g_decimal_estoque)
        grade1.TextMatrix(z, 6) = Format(fValidaValorNovo(grade1.TextMatrix(z, 6)), g_decimal_compra)
        grade1.TextMatrix(z, 7) = Format(fValidaValorNovo(grade1.TextMatrix(z, 7)), "##,###,##0.00")
        grade1.TextMatrix(z, 8) = Format(fValidaValorNovo(grade1.TextMatrix(z, 8)), g_decimal_compra)
        grade1.TextMatrix(z, 9) = Format(fValidaValorNovo(grade1.TextMatrix(z, 9)), "##,###,##0.00")
        grade1.TextMatrix(z, 10) = Format(fValidaValorNovo(grade1.TextMatrix(z, 10)), "##,###,##0.00")
        grade1.TextMatrix(z, 11) = Format(fValidaValorNovo(grade1.TextMatrix(z, 11)), "##,###,##0.00")
        grade1.TextMatrix(z, 15) = Format(fValidaValorNovo(grade1.TextMatrix(z, 15)), "##,###,##0.00")
        grade1.TextMatrix(z, 18) = Format(fValidaValorNovo(grade1.TextMatrix(z, 18)), g_decimal_custo)
        grade1.TextMatrix(z, 19) = Format(fValidaValorNovo(grade1.TextMatrix(z, 19)), "##,###,##0.00")
        grade1.TextMatrix(z, 20) = Format(fValidaValorNovo(grade1.TextMatrix(z, 20)), "##,###,##0.00")
        grade1.TextMatrix(z, 21) = Format(fValidaValorNovo(grade1.TextMatrix(z, 21)), "##,###,##0.00")
        grade1.TextMatrix(z, 22) = Format(fValidaValorNovo(grade1.TextMatrix(z, 22)), "##,###,##0.00")
        grade1.TextMatrix(z, 23) = Format(fValidaValorNovo(grade1.TextMatrix(z, 23)), "##,###,##0.00")
        grade1.TextMatrix(z, 27) = Format(fValidaValorNovo(grade1.TextMatrix(z, 27)), "##,###,##0.0000")
        grade1.TextMatrix(z, 31) = Format(fValidaValorNovo(grade1.TextMatrix(z, 31)), "##,###,##0.00")
        grade1.TextMatrix(z, 32) = Format(fValidaValorNovo(grade1.TextMatrix(z, 32)), "##,###,##0.00")
        grade1.TextMatrix(z, 34) = Format(fValidaValorNovo(grade1.TextMatrix(z, 34)), "##,###,##0.00")
        grade1.TextMatrix(z, 35) = Format(fValidaValorNovo(grade1.TextMatrix(z, 35)), "##,###,##0.00")
        grade1.TextMatrix(z, 36) = Format(fValidaValorNovo(grade1.TextMatrix(z, 36)), "##,###,##0.00")
        grade1.TextMatrix(z, 37) = Format(fValidaValorNovo(grade1.TextMatrix(z, 37)), "##,###,##0.00")
        grade1.TextMatrix(z, 55) = Format(fValidaValorNovo(grade1.TextMatrix(z, 55)), g_decimal_custo)
        grade1.TextMatrix(z, 56) = Format(fValidaValorNovo(grade1.TextMatrix(z, 56)), g_decimal_custo)
        
        If grade1.TextMatrix(z, 17) = "NIE" Then
            lbl_nota = "NF"
        Else
            lbl_nota = grade1.TextMatrix(z, 17)
        End If
    Next
    
    With grade1
        If l_opcao = 2 Then grade1.Rows = grade1.Rows + 1
        ' Mostrar os números nas colunas
        For z = 1 To .Rows - 1
            .TextMatrix(z, 0) = z
        Next
        
        .TextMatrix(.Rows - 1, 0) = NovaLinha
    End With

Exit Sub
Err_FormaGrid: ValidaErros Err, Me.Caption & " - FormaGrid"
End Sub

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Clejunior Freitas                               Start:14/03/2013
' Alteração:Clejunior Freitas                             Start:14/03/2013
'           PreencheColecaocomFormaFaturamento colFormaFaturamento, False
'           foi alterado para true para buscar as formas de faturamento de entrada
'==========================================================================
Private Sub Form_Load()
On Error GoTo Err_Form_Load

    Adodc1.ConnectionString = g_StrCon
    Adodc1.CursorLocation = adUseClient
                
    flag_tela_entrada_mercadoria = 0
    cbo_frete.Clear
    cbo_frete.AddItem "1- CIF"
    cbo_frete.AddItem "2- FOB"
    
    FormaGrid
    FormaGridConf
    
    cbo_atualizacusto.Clear
    cbo_atualizacusto.AddItem "0- Não Atualiza"
    cbo_atualizacusto.AddItem "1- Custo/Preços"
    cbo_atualizacusto.AddItem "2- Apenas Custos"
    
    cDPFFaturamento.InfluenciaEstoque = 1
    
    cbo_TipoContaConsumo.Clear
    cbo_TipoContaConsumo.AddItem "1 - Energia Elétrica/Gás"
    cbo_TipoContaConsumo.AddItem "2 - Água"
    cbo_TipoContaConsumo.AddItem "3 - Telefone/Comunicação"
    
    cbo_TipodeLigacao.Clear
    cbo_TipodeLigacao.AddItem "1 - Monofásico"
    cbo_TipodeLigacao.AddItem "2 - Bifásico"
    cbo_TipodeLigacao.AddItem "3 - Trifásico"
    
    cbo_GrupoDeTensao.Clear
    cbo_GrupoDeTensao.AddItem "01 - A1 - Alta Tensão (230kV ou mais)"
    lngIDGrupoDeTensao(1) = 0
    cbo_GrupoDeTensao.AddItem "02 - A2 - Alta Tensão (88 a 138kV)"
    lngIDGrupoDeTensao(2) = 1
    cbo_GrupoDeTensao.AddItem "03 - A3 - Alta Tensão (69kV)"
    lngIDGrupoDeTensao(3) = 2
    cbo_GrupoDeTensao.AddItem "04 - A3a - Alta Tensão (30kV a 44kV)"
    lngIDGrupoDeTensao(4) = 3
    cbo_GrupoDeTensao.AddItem "05 - A4 - Alta Tensão (2,3kV a 25kV)"
    lngIDGrupoDeTensao(5) = 4
    cbo_GrupoDeTensao.AddItem "06 - AS - Alta Tensão Subterrâneo 06"
    lngIDGrupoDeTensao(6) = 5
    cbo_GrupoDeTensao.AddItem "07 - B1 - Residencial 07"
    lngIDGrupoDeTensao(7) = 6
    cbo_GrupoDeTensao.AddItem "08 - B1 - Residencial Baixa Renda 08"
    lngIDGrupoDeTensao(8) = 7
    cbo_GrupoDeTensao.AddItem "09 - B2 - Rural 09"
    lngIDGrupoDeTensao(9) = 8
    cbo_GrupoDeTensao.AddItem "10 - B2 - Cooperativa de Eletrificação Rural"
    lngIDGrupoDeTensao(10) = 9
    cbo_GrupoDeTensao.AddItem "11 - B2 - Serviço Público de Irrigação"
    lngIDGrupoDeTensao(11) = 10
    cbo_GrupoDeTensao.AddItem "12 - B3 - Demais Classes"
    lngIDGrupoDeTensao(12) = 11
    cbo_GrupoDeTensao.AddItem "13 - B4a - Iluminação Pública - rede de distribuição"
    lngIDGrupoDeTensao(13) = 12
    cbo_GrupoDeTensao.AddItem "14 - B4b - Iluminação Pública - bulbo de lâmpada"
    lngIDGrupoDeTensao(14) = 13

    Dim cCrdCriador As New AutCria.cCrdCadastros
    Set cDPEmpresa = cCrdCriador.CrieUtilitariosCadastros.CarregaDominioProblemaEmpresa
    Set cDPParamSistema = cCrdCriador.CrieUtilitariosCadastros.CarregaDominioProblemaParametrosSistema(cDPEmpresa)
    cCtrlEntradaSaida.PreencheColecaocomPropriedadesTributacao cDPEmpresa.codigo, colTributacao
    cCtrlEntradaSaida.PreencheColecaocomFormaFaturamento colFormaFaturamento, True
    Set cCrdCriador = Nothing
        

Exit Sub
Err_Form_Load: ValidaErros Err, Me.Caption & " - Form_Load"
End Sub

Private Sub msk_emissao_GotFocus()
    msk_emissao.BackColor = 12648447
    msk_emissao.SelStart = 0
    msk_emissao.SelLength = 10
End Sub

Private Sub msk_emissao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_modelo_nf.SetFocus
End Sub

Private Sub msk_emissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '*** Autor: Diego Martins Ticket: TT438 Data: 26/06/2012 ***
        If Trim(txt_modelo_nf) = "55" Then
            txt_chave_acesso.SetFocus
        Else
            msk_entrada.SetFocus
        End If
    End If
End Sub

Private Sub msk_emissao_LostFocus()
On Error GoTo Emissao

    msk_emissao = MascaraData(msk_emissao)
    
    If IsNumeric(txt_numero_nf) And IsNumeric(txt_codigo_fornecedor) Then
        Set gb_Recordset = Conexao.GeraRecordset("SELECT numero FROM movimento_cabecalho_nota_fiscal_entrada WHERE data_de_emissao = " & FormataData(msk_emissao) & " and tipo_do_documento <> 5 and numero = " & txt_numero_nf & " and serie = '" & txt_serie_nf & "' and modelo_nf = '" & txt_modelo_nf & "' and codigo_do_fornecedor = " & txt_codigo_fornecedor, 1)
        If gb_Recordset.RecordCount > 0 Then
            Alerta "Nota Fiscal já Cadastrada!"
            txt_numero_nf.SetFocus
        End If
        gb_Recordset.Close
    End If
    msk_emissao.BackColor = &H8000000E

Exit Sub
Emissao: ValidaErros Err, Me.Caption & " - Emissao"
End Sub

Private Sub msk_entrada_GotFocus()
    msk_entrada.BackColor = 12648447
    msk_entrada.SelStart = 0
    msk_entrada.SelLength = 10
End Sub

Private Sub msk_entrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        '*** Autor: Diego Martins Ticket: TT438 Data: 26/06/2012 ***
        If Trim(txt_modelo_nf) = "55" Then
            txt_chave_acesso.SetFocus
        Else
            msk_emissao.SetFocus
        End If
    End If
End Sub

Private Sub msk_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_codigo_forma.SetFocus
End Sub

Private Sub msk_entrada_LostFocus()
    msk_entrada = MascaraData(msk_entrada)
    msk_entrada.BackColor = &H8000000E
End Sub

Private Sub txt_bc_substituicao_GotFocus()
    txt_bc_substituicao.BackColor = 12648447
    txt_bc_substituicao.SelStart = 0
    txt_bc_substituicao.SelLength = Len(txt_bc_substituicao)
End Sub

Private Sub txt_bc_substituicao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_icms.SetFocus
End Sub

Private Sub txt_bc_substituicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_substituicao.SetFocus
End Sub

Private Sub txt_bc_substituicao_LostFocus()
    txt_bc_substituicao.BackColor = &H8000000E
            
    If IsNumeric(txt_bc_substituicao) Then
        txt_bc_substituicao = Format(fValidaValorNovo(txt_bc_substituicao), "##,###,##0.00")
    End If
End Sub

Private Sub txt_desconto_GotFocus()
    txt_desconto.BackColor = 12648447
    txt_desconto.SelStart = 0
    txt_desconto.SelLength = Len(txt_desconto)
End Sub

Private Sub txt_desconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_outras.SetFocus
End Sub

Private Sub txt_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then msk_chegada.SetFocus
End Sub

Private Sub txt_desconto_LostFocus()
    txt_desconto.BackColor = &H8000000E
            
    If IsNumeric(txt_desconto) Then
        txt_desconto = Format(fValidaValorNovo(txt_desconto), "##,###,##0.00")
    End If
End Sub

Private Sub txt_observacoes_GotFocus()
    txt_observacoes.BackColor = 12648447
    txt_observacoes.SelStart = 0
    txt_observacoes.SelLength = Len(txt_observacoes)
End Sub

Private Sub txt_observacoes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then msk_chegada.SetFocus
End Sub

Private Sub txt_observacoes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grade1.SetFocus
End Sub

Private Sub txt_observacoes_LostFocus()
    txt_observacoes.BackColor = &H8000000E
End Sub

Private Sub txt_codigo_fornecedor_GotFocus()
    txt_codigo_fornecedor.BackColor = 12648447
    txt_codigo_fornecedor.SelStart = 0
    txt_codigo_fornecedor.SelLength = Len(txt_codigo_fornecedor)
End Sub

Private Sub txt_codigo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsNumeric(txt_codigo_fornecedor) Then
            g_string = ""
            consulta_fornecedor.Show (1)
            If Len(g_string) > 0 Then
                Call BuscaFornecedor(g_string)
            End If
        Else
            Call BuscaFornecedor(txt_codigo_fornecedor)
        End If
    End If
    Call ValidaInteiro(KeyAscii)
End Sub

Private Sub BuscaFornecedor(xcodigo As String, Optional strCNPJ As String)
On Error GoTo BuscaFornecedor
Dim strCondicao As String

    If strCNPJ <> "" Then
        strCondicao = " CGC = '" & strCNPJ & "'"
    Else
        strCondicao = " codigo = '" & xcodigo & "'"
    End If
    
    Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM fornecedor WHERE " & strCondicao, 1)
    If gb_Recordset.RecordCount > 0 Then
        LimpaTelaFornecedor
        AtualTelaFornecedor
        '*** Autor:Fernando Silva Data:16/02/2011
        'Efetuado teste para evitar nos casos de consulta ou conferência de produtos
        'gerar algum erro por causa do campo desabilitado ***
        If frmdados.Enabled = True Then txt_fornecedor.SetFocus
        
    Else
        LimpaTelaFornecedor
    End If
    gb_Recordset.Close

Exit Sub
BuscaFornecedor: ValidaErros Err, Me.Caption & " - buscafornecedor"
End Sub

Private Sub BuscaTransportadora(strcodigo As String, Optional strCNPJ As String)
On Error GoTo BuscaTransportadora
Dim strCondicao As String

    If strCNPJ <> "" Then
        strCondicao = " CGC = '" & strCNPJ & "'"
    Else
        strCondicao = " codigo = '" & strcodigo & "'"
    End If
    
    Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM transportadora WHERE " & strCondicao, 1)
    If gb_Recordset.RecordCount > 0 Then
        Call PreencheTransportadora
        txt_transportadora_nome.SetFocus
    Else
        Alerta "Transportadora não cadastrada!"
    End If
    gb_Recordset.Close

Exit Sub
BuscaTransportadora: ValidaErros Err, Me.Caption & " - BuscaTransportadora"
End Sub

'*****************************************************************************
'Criação: Thiago Leão                                         Data: 09/02/2012
'Propósito: Preenche dados da transportadora
'*****************************************************************************
Private Sub PreencheTransportadora()
On Error GoTo Err_PreencheTransportadora

    With gb_Recordset
        txt_codigo_transportadora = !codigo
        txt_transportadora_nome = !NOME
        txt_transportadora_endereco = !Endereco
        txt_transportadora_bairro = !Bairro
        txt_transportadora_cidade = !Cidade
        txt_transportadora_uf = !UF
        txt_transportadora_cnpj = !CGC
        txt_transportadora_inscricao_estadual = !inscricao_estadual
    End With

Exit Sub
Err_PreencheTransportadora: ValidaErros Err, Me.Caption & " - PreencheTransportadora"
End Sub

Private Sub txt_codigo_fornecedor_LostFocus()
    txt_codigo_fornecedor.BackColor = &H8000000E
End Sub

Private Sub AtualTelaFornecedor()
On Error GoTo file

    With gb_Recordset
       If !Pessoa = 1 Then
           lbl_cgc = !CGC
           lbl_inscricao = !Identidade
           lCGCFornecedor = ""
       ElseIf !Pessoa = 3 Then
           lbl_cgc = !CGC
           lbl_inscricao = !inscricao_estadual
           lCGCFornecedor = ""
       Else
           lbl_cgc = !CGC
           lCGCFornecedor = !CGC
           lbl_inscricao = !inscricao_estadual
       End If
       lPessoa = !Pessoa
       txt_codigo_fornecedor = !codigo
       txt_fornecedor.Text = !razao_social
       lbl_endereco = !Endereco & " - " & !Bairro
       lbl_cidade = !Cidade
       If IsNull(!Pais) Or Trim(!Pais) = "" Then
           lbl_pais = "BRASIL"
       Else
           lbl_pais = !Pais
       End If
    
       lbl_uf = !UF
       lbl_telefone = !Telefone
       lbl_bairro = !Bairro
       lbl_cep = !CEP
    
       lporcentagemreducao = !red_icms
       lporcredicmssubst = !red_icms
       x_perc_icms_subst = !icms_subst
       xplanoconta = !conta_contabil
       g_reducao_invertido = !reducao_invertido
       txt_porc_red_icms = !red_icms
       txt_redicmssubst = !red_icms
    End With
             
Exit Sub
file: If Err.Number = 94 Then Resume Next
End Sub

Private Sub LimpaTelaFornecedor()
    txt_codigo_fornecedor = ""
    txt_fornecedor = ""
    lbl_cgc = ""
    lbl_inscricao = ""
    lbl_endereco = ""
    lbl_cidade = ""
    lbl_uf = ""
    lbl_pais = ""
    x_perc_icms_subst = 0
    txt_porc_red_icms = 0
End Sub
        
Private Sub txt_frete_GotFocus()
    txt_frete.BackColor = 12648447
    txt_frete.SelStart = 0
    txt_frete.SelLength = Len(txt_frete)
End Sub

Private Sub txt_Frete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_ipi.SetFocus
End Sub

Private Sub txt_frete_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_frete_conhecimento.SetFocus
End Sub

Private Sub txt_frete_LostFocus()
    txt_frete.BackColor = &H8000000E
    
    If IsNumeric(txt_frete) Then
        txt_frete = Format(fValidaValorNovo(txt_frete), "##,###,##0.00")
    End If
End Sub

Private Sub txt_frete_conhecimento_GotFocus()
    txt_frete_conhecimento.BackColor = 12648447
    txt_frete_conhecimento.SelStart = 0
    txt_frete_conhecimento.SelLength = Len(txt_frete_conhecimento)
End Sub

Private Sub txt_frete_conhecimento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_frete.SetFocus
End Sub

Private Sub txt_frete_conhecimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_seguro.SetFocus
End Sub

Private Sub txt_frete_conhecimento_LostFocus()
    txt_frete_conhecimento.BackColor = &H8000000E
    
    If IsNumeric(txt_frete_conhecimento) Then
        txt_frete_conhecimento = Format(fValidaValorNovo(txt_frete_conhecimento), "##,###,##0.00")
    End If
End Sub

Private Sub txt_icms_GotFocus()
    txt_icms.BackColor = 12648447
    txt_icms.SelStart = 0
    txt_icms.SelLength = Len(txt_icms)
End Sub

Private Sub txt_icms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_bc_icms.SetFocus
End Sub

Private Sub txt_icms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_bc_substituicao.SetFocus
End Sub

Private Sub txt_icms_LostFocus()
    txt_icms.BackColor = &H8000000E
            
    If IsNumeric(txt_icms) Then
        txt_icms = Format(fValidaValorNovo(txt_icms), "##,###,##0.00")
    End If
End Sub

Private Sub txt_ipi_GotFocus()
    txt_ipi.BackColor = 12648447
    txt_ipi.SelStart = 0
    txt_ipi.SelLength = Len(txt_ipi)
End Sub

Private Sub txt_ipi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_substituicao.SetFocus
End Sub

Private Sub txt_ipi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_frete.SetFocus
End Sub

Private Sub txt_ipi_LostFocus()
    txt_ipi.BackColor = &H8000000E
        
    If IsNumeric(txt_ipi) Then
        txt_ipi = Format(fValidaValorNovo(txt_ipi), "##,###,##0.00")
    End If
End Sub

Private Sub txt_modelo_nf_GotFocus()
    txt_modelo_nf.BackColor = 12648447
    txt_modelo_nf.SelStart = 0
End Sub

Private Sub txt_modelo_nf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If txt_serie_nf.Enabled = True Then
            txt_serie_nf.SetFocus
        Else
            txt_numero_nf.SetFocus
        End If
    End If
End Sub

Private Sub txt_modelo_nf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then msk_emissao.SetFocus
End Sub

Private Sub txt_modelo_nf_LostFocus()
    txt_modelo_nf.BackColor = &H8000000E
End Sub

Private Sub txt_numero_nf_GotFocus()
    txt_numero_nf.BackColor = 12648447
    txt_numero_nf.SelStart = 0
    txt_numero_nf.SelLength = Len(txt_numero_nf)
End Sub

Private Sub txt_numero_nf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_codigo_fornecedor.SetFocus
End Sub

Private Sub txt_numero_nf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt_serie_nf.Enabled = True Then
            txt_serie_nf.SetFocus
        Else
            txt_modelo_nf.SetFocus
        End If
    End If
    Call ValidaInteiro(KeyAscii)
End Sub

Private Sub txt_numero_nf_LostFocus()
    txt_numero_nf.BackColor = &H8000000E
End Sub

Private Sub txt_outras_GotFocus()
    txt_outras.BackColor = 12648447
    txt_outras.SelStart = 0
    txt_outras.SelLength = Len(txt_outras)
End Sub

Private Sub txt_outras_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_seguro.SetFocus
End Sub

Private Sub txt_outras_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_desconto.SetFocus
End Sub

Private Sub txt_outras_LostFocus()
    txt_outras.BackColor = &H8000000E
    If IsNumeric(txt_outras) Then
        txt_outras = Format(fValidaValorNovo(txt_outras), "##,###,##0.00")
    End If
End Sub

Private Sub txt_bc_icms_GotFocus()
    txt_bc_icms.BackColor = 12648447
    txt_bc_icms.SelStart = 0
    txt_bc_icms.SelLength = Len(txt_bc_icms)
End Sub

Private Sub txt_bc_icms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_total.SetFocus
End Sub

Private Sub txt_bc_icms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_icms.SetFocus
End Sub

Private Sub txt_bc_icms_LostFocus()
    txt_bc_icms.BackColor = &H8000000E
            
    If IsNumeric(txt_bc_icms) Then txt_bc_icms = Format(fValidaValorNovo(txt_bc_icms), "##,###,##0.00")
End Sub

Private Sub txt_PctIPI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then cmd_explorer.SetFocus
End Sub

Private Sub txt_PctIPI_GotFocus()
    txt_PctIPI.BackColor = 12648447
    txt_PctIPI.SelStart = 0
    txt_PctIPI.SelLength = Len(txt_PctIPI)
End Sub

Private Sub txt_PctIPI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_PctICMS.SetFocus
End Sub

Private Sub txt_PctIPI_LostFocus()
    txt_PctIPI.BackColor = &H8000000E
            
    If IsNumeric(txt_PctIPI) Then txt_PctIPI = Format(fValidaValorNovo(txt_PctIPI), "##,###,##0.00")
End Sub

Private Sub txt_PctICMS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txt_PctIPI.SetFocus
End Sub

Private Sub txt_PctICMS_GotFocus()
    txt_PctICMS.BackColor = 12648447
    txt_PctICMS.SelStart = 0
    txt_PctICMS.SelLength = Len(txt_PctICMS)
End Sub

Private Sub txt_PctICMS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmd_ler_xml.SetFocus
End Sub

Private Sub txt_PctICMS_LostFocus()
    txt_PctICMS.BackColor = &H8000000E
            
    If IsNumeric(txt_PctICMS) Then txt_PctICMS = Format(fValidaValorNovo(txt_PctICMS), "##,###,##0.00")
End Sub

Private Sub txt_peso_bruto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_especie.SetFocus
End Sub
Private Sub txt_peso_bruto_GotFocus()
    txt_peso_bruto.BackColor = 12648447
    txt_peso_bruto.SelStart = 0
    txt_peso_bruto.SelLength = Len(txt_peso_bruto)
End Sub
Private Sub txt_peso_bruto_LostFocus()
    txt_peso_bruto.BackColor = &H8000000E
    txt_peso_bruto = Format(fValidaValor(txt_peso_bruto), "##,###,##0.000")
End Sub

Private Sub txt_peso_bruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_peso_liquido.SetFocus
End Sub

Private Sub txt_peso_liquido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_peso_bruto.SetFocus
End Sub

Private Sub txt_peso_liquido_GotFocus()
    txt_peso_liquido.BackColor = 12648447
    txt_peso_liquido.SelStart = 0
    txt_peso_liquido.SelLength = Len(txt_peso_liquido)
End Sub
Private Sub txt_peso_liquido_LostFocus()
    txt_peso_liquido.BackColor = &H8000000E
    txt_peso_liquido = Format(fValidaValor(txt_peso_liquido), "##,###,##0.000")
End Sub

Private Sub txt_porc_red_icms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cbo_frete.SetFocus
End Sub

Private Sub txt_porc_red_icms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_observacoes.SetFocus
End Sub

Private Sub txt_seguro_GotFocus()
    txt_seguro.BackColor = 12648447
    txt_seguro.SelStart = 0
    txt_seguro.SelLength = Len(txt_seguro)
End Sub

Private Sub txt_seguro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_frete.SetFocus
End Sub

Private Sub txt_seguro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_outras.SetFocus
End Sub

Private Sub txt_seguro_LostFocus()
    txt_seguro.BackColor = &H8000000E
        
    If IsNumeric(txt_seguro) Then txt_seguro = Format(fValidaValorNovo(txt_seguro), "##,###,##0.00")
End Sub

Private Sub txt_serie_nf_GotFocus()
    txt_serie_nf.BackColor = 12648447
    txt_serie_nf.SelStart = 0
    txt_serie_nf.SelLength = Len(txt_serie_nf)
End Sub

Private Sub txt_serie_nf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_numero_nf.SetFocus
End Sub

Private Sub txt_serie_nf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_modelo_nf.SetFocus
End Sub

Private Sub txt_serie_nf_LostFocus()
    txt_serie_nf.BackColor = &H8000000E
    txt_serie_nf = UCase(Trim(txt_serie_nf))
End Sub

Private Sub txt_substituicao_GotFocus()
    txt_substituicao.BackColor = 12648447
    txt_substituicao.SelStart = 0
    txt_substituicao.SelLength = Len(txt_substituicao)
End Sub

Private Sub txt_substituicao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_bc_substituicao.SetFocus
End Sub

Private Sub txt_substituicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_ipi.SetFocus
End Sub

Private Sub txt_substituicao_LostFocus()
    txt_substituicao.BackColor = &H8000000E
        
    If IsNumeric(txt_substituicao) Then txt_substituicao = Format(fValidaValorNovo(txt_substituicao), "##,###,##0.00")
End Sub

Private Sub txt_total_GotFocus()
    txt_total.BackColor = 12648447
    txt_total.SelStart = 0
    txt_total.SelLength = Len(txt_total)
End Sub

Private Sub txt_total_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_codigo_forma.SetFocus
End Sub

Private Sub txt_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(txt_total) Then
            txt_total = Format(fValidaValorNovo(txt_total), "##,###,##0.00")
            txt_bc_icms.SetFocus
        Else
            Alerta "Informe o valor total da nota!"
            txt_total.SetFocus
        End If
    End If
End Sub

Private Sub txt_total_LostFocus()
    txt_total.BackColor = &H8000000E
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err_Toolbar1_ButtonClick

    Select Case Button
        Case Is = "Novo"
            Call Novo
            
        Case Is = "Alterar"
            Alterar
            
        Case Is = "Excluir"
            Call Excluir
        
        Case Is = "Pesquisar"
            Pesquisar
        
        Case Is = "Salvar"
            SalvarDocumento
    
        Case Is = "Cancelar"
            Call StatusImportarXML(False)
            flag_tela_entrada_mercadoria = 0
            Form_Activate
                
        Case Is = "Imprimir"
            ImpressaoEspelho
    
        Case Is = "Ajuda"
            frm_help.Show (1)
        
        Case Is = "Sair"
            Unload Me
            
        Case Is = "Imp.NF"
            If chk_impressao_nf.Value = 1 Then
                If Confirma("Deseja Reimprimir a Nota Fiscal de Entrada?") = vbYes Then
                    'If BuscaCFOP Then
                       If cDPEmpresa.NotaFiscalEletronica = 1 Then
                          '*** Autor: Diego Martins Data:04/04/2011 Objetivo:Desvincular Emissão Danfe do Processo de emissão de NF-e ***
                          Call ImprimirNFE(TipoNFe.entrada, lSequencia)
                       Else
                          If g_novo_configNF = 0 Then
                             Imprime_NF
                          Else
                             Imprime_NFNovo
                          End If
                       End If
                    'End If
                End If
            Else
                Alerta "Tela não Liberada Para Reimpressão de Nota!"
            End If
            
        Case Is = "Frete"
            g_string = "E"
            frm_conhecimento.Show
        
        Case Is = "Pedidos"
            Call Pedidos
        
        Case Is = "V.Custos"
            Call VerificarCustos
        
        Case Is = "Canc.NF"
            Call EfetuaCancelamentoNF
            
        Case Is = "Pd.Transf"
            PesquisaConferencia
    End Select

Exit Sub
Err_Toolbar1_ButtonClick: ValidaErros Err, Me.Caption & " - Toolbar1_ButtonClick"
End Sub

Private Sub VerificarCustos()
On Error GoTo VerificaCusto

    If ValidaCampos Then
        BuscaConhecimento
        
        If chk_lancamento_venda.Value = 1 Then
            If Trim(xCodigoProduto) <> "" Then
                Call Conexao.DeleteSintetico("calculo_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and codigo_produto = '" & xCodigoProduto & "' and numero_nf = '" & txt_numero_nf & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & "' and outros = '" & lbl_nota & "'", 0)
            End If
        Else
            Call Conexao.DeleteSintetico("calculo_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and numero_nf = '" & txt_numero_nf & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & "' and outros = '" & lbl_nota & "'", 0)
        End If
    
        If CalculaNotaFiscal Then
            g_string4 = xCodigoProduto
            g_string = lbl_nota
            frm_calculo_entrada_produto.Show 1
            xCodigoProduto = ""
        End If
    End If

Exit Sub
VerificaCusto: ValidaErros Err, Me.Caption & " - VerificaCusto"
End Sub

Private Sub Excluir()
On Error GoTo file
    If chk_impressao_nf.Value = 0 Or g_nivel_acesso = 1 Then
        If VerificaDevolucaoCompras(False) Then Exit Sub
        If Confirma("Confirma Exclusão da Nota Fiscal Entrada?") = vbYes Then
            Conexao.BeginTrans
            If buscasenha Then
                If ExclusaoNota(False) Then
                    Conexao.CommitTrans
                    flag_tela_entrada_mercadoria = 0
                    Form_Activate
                    Exit Sub
                End If
            End If
            Conexao.RollbackTrans
        End If
    Else
        Alerta "Notas fiscais liberada apenas para cancelamento!"
    End If
Exit Sub
file: ValidaErros Err, Me.Caption & " - Excluir"
End Sub

Private Sub Novo()
On Error GoTo Err_Novo

    Call StatusImportarXML(True)
    l_opcao = 1
    pct_orçamento.Visible = False
    pct_icms.Visible = False
    txt_serie_nf.Enabled = True
    If frmdados.Enabled = True Then
    If txt_serie_nf = "CX" Then
            txt_serie_nf = ""
            txt_modelo_nf = ""
        End If
    End If
    Movimento_Nota_Fiscal_Entrada.Caption = "Movimento Nota Fiscal de Entrada - Inclusão"
    DesativaBotoes
    Toolbar1.Buttons(2).Enabled = False
    LimpaTela
    SSTab1.Tab = 0
    msk_entrada = Date
    cmd_confirma.Visible = True
    txt_codigo_fornecedor.SetFocus

Exit Sub
Err_Novo: ValidaErros Err, Me.Caption & " - Novo"
End Sub
'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Sub Alterar()
On Error GoTo Err_Alterar
    
    If chk_impressao_nf.Value = 1 Then
        Alerta "Nota fiscal impressa não disponível para alteração!" & Chr(13) & "Efetue cancelamento ou devolução!"
        Exit Sub
    End If
    Call StatusImportarXML(True)
    Movimento_Nota_Fiscal_Entrada.Caption = "Movimento Nota Fiscal de Entrada - Alteração"
    DesativaBotoes
    SSTab1.Tab = 0
    cmd_confirma.Visible = True
    txt_observacoes.SetFocus
    
Exit Sub
Err_Alterar: ValidaErros Err, Me.Caption & " - Alterar"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Form_KeyDown

    If KeyCode = vbKeyF2 And Toolbar1.Buttons(1).Enabled = True Then
        Call Novo
    ElseIf KeyCode = vbKeyF3 And Toolbar1.Buttons(2).Enabled = True Then
        Call Alterar
    ElseIf KeyCode = vbKeyF11 And Toolbar1.Buttons(6).Enabled = True Then
        Call VerificarCustos
    ElseIf KeyCode = vbKeyF6 And Toolbar1.Buttons(6).Enabled = True Then
        SalvarDocumento
    ElseIf KeyCode = vbKeyF7 And Toolbar1.Buttons(7).Enabled = True Then
        flag_tela_entrada_mercadoria = 0
        Form_Activate
    ElseIf KeyCode = vbKeyF10 And Toolbar1.Buttons(12).Enabled = True Then
        Unload Me
    ElseIf KeyCode = vbKeyF1 And Toolbar1.Buttons(11).Enabled = True Then
        frm_help.Show (1)
    ElseIf KeyCode = vbKeyF4 And Toolbar1.Buttons(3).Enabled = True Then
        Call Excluir
    ElseIf KeyCode = vbKeyF5 And Toolbar1.Buttons(4).Enabled = True Then
        Call Pesquisar
    ElseIf KeyCode = vbKeyF8 And Toolbar1.Buttons(5).Enabled = True Then
        ImpressaoEspelho
    End If

Exit Sub
Err_Form_KeyDown: ValidaErros Err, Me.Caption & " - Form_KeyDown"
End Sub

Private Sub Pesquisar()
On Error GoTo Pesquisa

    consulta_nota_fiscal_entrada.bolConferencia = False
    consulta_nota_fiscal_entrada.Show 1
    If g_string <> "" Then
        Sql_Query = "SELECT MCNFE.*,(INCC.codigo) as PkCodigoConsumo, " & _
                           "(INCC.tipo_conta) as TipoConta,(INCC.codigo_consumo) as CodigoConsumo," & _
                           "(INCC.tipo_ligacao) as TipoLigacao,(INCC.grupo_tensao) as GrupoTensao,ML.seq_controle " & _
                    "FROM movimento_cabecalho_nota_fiscal_entrada MCNFE " & _
                    "LEFT JOIN movimento_lote ML ON ML.empresa = MCNFE.empresa and ML.sequencia = MCNFE.sequencia and ML.status = 'E'" & _
                    "LEFT JOIN info_conta_consumo INCC " & _
                    "ON INCC.empresa = MCNFE.empresa AND INCC.codigo = MCNFE.fk_codigo_conta_consumo " & _
                    "WHERE MCNFE.sequencia = " & g_string
        Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 1)
        If gb_Recordset.RecordCount > 0 Then
            LimpaTela
            AtualTela
            AtualizaGrid (0)
            AtivaBotoes
            bolEntradaTMP = False
        Else
            gb_Recordset.Close
        End If
    End If

Exit Sub
Pesquisa: ValidaErros Err, Me.Caption & " - Pesquisa"
End Sub

'*****************************************************************************
'Criação: Diego Martins dos Santos                            Data: 18/06/2010
'
'Propósito:Pesquisa notas pendentes na conferência
'Alteração: Fernando Silva                                          15/02/2011
'           Adicionada a linha: grade1.row = grade1.Rows - 1 para mover o foco
'           Para a ultima linha da grade principal de produtos   -   OS: 15631
'Alteração: João Batista Medeiros                             Data: 23/07/2012
'Propósito: Atualizado select da pesquisa conferencia para pegar informações
'do lote, do tipo de consumo e demais alterações para atender o SPED Fiscal. TICKET :TT904
'Alteração: Ronaldo Robledo                                         05/03/2013
'           Inserido variavel bolPedidoTransf para detectar que o pedido originado
'           de pedido transferencia
'*****************************************************************************
Private Sub PesquisaConferencia()
On Error GoTo Err_PesquisaConferencia

    consulta_nota_fiscal_entrada.bolConferencia = True
    consulta_nota_fiscal_entrada.Show 1
    If g_string <> "" Then
        Set gb_Recordset = Conexao.GeraRecordset("SELECT MCNFE.*,(INCC.codigo) as PkCodigoConsumo, " & _
                                                 "(INCC.tipo_conta) as TipoConta,(INCC.codigo_consumo) as CodigoConsumo," & _
                                                 "(INCC.tipo_ligacao) as TipoLigacao,(INCC.grupo_tensao) as GrupoTensao,ML.seq_controle " & _
                                                 "FROM movimento_cabecalho_nota_fiscal_entrada_tmp MCNFE " & _
                                                 "LEFT JOIN movimento_lote ML ON ML.empresa = MCNFE.empresa and ML.sequencia = MCNFE.sequencia and ML.status = 'E' " & _
                                                 "LEFT JOIN info_conta_consumo INCC " & _
                                                 "ON INCC.empresa = MCNFE.empresa AND INCC.codigo = MCNFE.fk_codigo_conta_consumo " & _
                                                 "WHERE MCNFE.sequencia = " & g_string, 1 & "")
        If gb_Recordset.RecordCount > 0 Then
            bolPedidoTransf = IIf((gb_Recordset!pedido_transferencia = 1), True, False)
            LimpaTela
            AtualTela
            AtualizaGrid (1)
            bolEntradaTMP = True
            AtivaBotoes
            frm_conf.Enabled = True
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(7).Enabled = True
            l_opcao = 1
            bolConferencia = False
            frm_dados1.Enabled = True
            frm_dados2.Enabled = True
            If bolPedidoTransf Then frmdados.Enabled = True
            Call VerificaQtdeTotalConferencia
            grade1.Row = grade1.Rows - 1
        Else
            gb_Recordset.Close
        End If
    End If

Exit Sub
Err_PesquisaConferencia: ValidaErros Err, Me.Caption & " - PesquisaConferencia"
Resume
End Sub

Private Sub txt_codigo_transportadora_GotFocus()
    txt_codigo_transportadora.BackColor = 12648447
    txt_codigo_transportadora.SelStart = 0
    txt_codigo_transportadora.SelLength = Len(txt_codigo_transportadora)
End Sub

Private Sub txt_codigo_transportadora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '*** Autor: Diego Martins Data:31/03/2012 Motivo: Não estava buscando corretamente dados da transportadora ***
        If IsNumeric(txt_codigo_transportadora) Then
            Call BuscaTransportadora(txt_codigo_transportadora)
        Else
            g_string = ""
            consulta_transportadora.Show 1
            If Len(g_string) > 0 Then
                Call BuscaTransportadora(g_string)
            End If
        End If
    End If
End Sub

Private Sub txt_codigo_transportadora_LostFocus()
On Error GoTo file

    txt_codigo_transportadora.BackColor = &H8000000E
    
Exit Sub
file: ValidaErros Err, Me.Caption & " - CodigoTransportadora"
End Sub

Private Sub txt_transportadora_cidade_GotFocus()
    txt_transportadora_cidade.BackColor = 12648447
    txt_transportadora_cidade.SelStart = 0
    txt_transportadora_cidade.SelLength = Len(txt_transportadora_cidade)
End Sub

Private Sub txt_transportadora_cidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_endereco.SetFocus
End Sub

Private Sub txt_transportadora_cidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_uf.SetFocus
End Sub

Private Sub txt_transportadora_cidade_LostFocus()
    txt_transportadora_cidade = UCase(Trim(txt_transportadora_cidade))
    txt_transportadora_cidade.BackColor = &H8000000E
End Sub

Private Sub txt_transportadora_cnpj_GotFocus()
    txt_transportadora_cnpj.BackColor = 12648447
    txt_transportadora_cnpj.SelStart = 0
    txt_transportadora_cnpj.SelLength = Len(txt_transportadora_cnpj)
End Sub

Private Sub txt_transportadora_cnpj_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_uf.SetFocus
End Sub

Private Sub txt_transportadora_cnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_inscricao_estadual.SetFocus
End Sub

Private Sub txt_transportadora_cnpj_LostFocus()
    txt_transportadora_cnpj.BackColor = &H8000000E
    txt_transportadora_cnpj = Trim(txt_transportadora_cnpj)
End Sub

Private Sub txt_transportadora_endereco_GotFocus()
    txt_transportadora_endereco.BackColor = 12648447
    txt_transportadora_endereco.SelStart = 0
    txt_transportadora_endereco.SelLength = Len(txt_transportadora_endereco)
End Sub

Private Sub txt_transportadora_endereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_nome.SetFocus
End Sub

Private Sub txt_transportadora_endereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_cidade.SetFocus
End Sub

Private Sub txt_transportadora_endereco_LostFocus()
    txt_transportadora_endereco = UCase(Trim(txt_transportadora_endereco))
    txt_transportadora_endereco.BackColor = &H8000000E
End Sub

Private Sub txt_transportadora_inscricao_estadual_GotFocus()
    txt_transportadora_inscricao_estadual.BackColor = 12648447
    txt_transportadora_inscricao_estadual.SelStart = 0
    txt_transportadora_inscricao_estadual.SelLength = Len(txt_transportadora_inscricao_estadual)
End Sub

Private Sub txt_transportadora_inscricao_estadual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_cnpj.SetFocus
End Sub

Private Sub txt_transportadora_inscricao_estadual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_placa.SetFocus
End Sub

Private Sub txt_transportadora_inscricao_estadual_LostFocus()
    txt_transportadora_inscricao_estadual = Trim(txt_transportadora_inscricao_estadual)
    txt_transportadora_inscricao_estadual.BackColor = &H8000000E
End Sub

Private Sub txt_transportadora_nome_GotFocus()
    txt_transportadora_nome.BackColor = 12648447
    txt_transportadora_nome.SelStart = 0
    txt_transportadora_nome.SelLength = Len(txt_transportadora_nome)
End Sub

Private Sub txt_transportadora_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_endereco.SetFocus
End Sub

Private Sub txt_transportadora_nome_LostFocus()
    txt_transportadora_nome = UCase(Trim(txt_transportadora_nome))
    txt_transportadora_nome.BackColor = &H8000000E
End Sub

Private Sub txt_transportadora_placa_GotFocus()
    txt_transportadora_placa.BackColor = 12648447
    txt_transportadora_placa.SelStart = 0
    txt_transportadora_placa.SelLength = Len(txt_transportadora_placa)
End Sub

Private Sub txt_transportadora_placa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_inscricao_estadual.SetFocus
End Sub

Private Sub txt_transportadora_placa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_placa_uf.SetFocus
End Sub

Private Sub txt_transportadora_placa_LostFocus()
    txt_transportadora_placa.BackColor = &H8000000E
    txt_transportadora_placa = UCase(Trim(txt_transportadora_placa))
End Sub

Private Sub txt_transportadora_placa_uf_GotFocus()
    txt_transportadora_placa_uf.BackColor = 12648447
    txt_transportadora_placa_uf.SelStart = 0
    txt_transportadora_placa_uf.SelLength = Len(txt_transportadora_placa_uf)
End Sub

Private Sub txt_transportadora_placa_uf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_placa.SetFocus
End Sub

Private Sub txt_transportadora_placa_uf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_porc_red_icms.SetFocus
End Sub

Private Sub txt_transportadora_placa_uf_LostFocus()
    txt_transportadora_placa_uf.BackColor = &H8000000E
    txt_transportadora_placa_uf = UCase(Trim(txt_transportadora_placa_uf))
End Sub

Private Sub txt_transportadora_uf_GotFocus()
    txt_transportadora_uf.BackColor = 12648447
    txt_transportadora_uf.SelStart = 0
    txt_transportadora_uf.SelLength = Len(txt_transportadora_uf)
End Sub

Private Sub txt_transportadora_uf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_transportadora_cidade.SetFocus
End Sub

Private Sub txt_transportadora_uf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_transportadora_cnpj.SetFocus
End Sub

Private Sub txt_transportadora_uf_LostFocus()
    txt_transportadora_uf = UCase(Trim(txt_transportadora_uf))
    txt_transportadora_uf.BackColor = &H8000000E
End Sub

Private Sub BuscanumeroNotaFiscal()
On Error GoTo Err_BuscanumeroNotaFiscal
Dim strPedidoLote As String

    strPedidoLote = txt_numero_nf
    
    Call Conexao.AlterarRecordset("empresas", "loja = '" & g_loja & "'", "codigo = '" & cDPEmpresa.codigo & "'", cDPEmpresa.codigo)
    Set gb_Recordset = Conexao.GeraRecordset("SELECT nota_fiscal FROM empresas WHERE codigo = '" & cDPEmpresa.codigo & "'", 1)
    If gb_Recordset.RecordCount > 0 Then
        txt_numero_nf = gb_Recordset!nota_fiscal + 1
        txt_numero_nf = Format(txt_numero_nf, "00000#")
    Else
        txt_numero_nf = "1"
        txt_numero_nf = Format(txt_numero_nf, "00000#")
    End If
    gb_Recordset.Close
    
    txt_modelo_nf = cCtrlEntradaSaida.ObtenhaModeloNF(1, cDPEmpresa) '*** Inserido 24/12/2010 contemplar SPED Fiscal ***
    
    Call Conexao.AlterarRecordset("empresas", "nota_fiscal  = '" & txt_numero_nf & "'", "codigo = '" & cDPEmpresa.codigo & "'", cDPEmpresa.codigo)
    
    '*** Autor: Diego Martins Data: 24/02/2011 Objetivo: Atualizar o numero do pedido anterior com o numero da nota gerado ***
    Call Conexao.AlterarRecordset("movimento_lote", "numero_pedido = " & txt_numero_nf, "empresa = '" & cDPEmpresa.codigo & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & "' and numero_pedido = '" & strPedidoLote & "' and data_emissao = " & FormataData(msk_emissao) & " and sequencia = '1' and status = 'E'", cDPEmpresa.codigo)
    
    Call Conexao.InserirRecordset("log_senhas", "data,hora,codigo_usuario,nome_usuario,historico,tela,observacoes,outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "','Busca Numero Nota Saida " & txt_numero_nf & " " & txt_fornecedor & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)

Exit Sub
Err_BuscanumeroNotaFiscal: ValidaErros Err, Me.Caption & " - BuscanumeroNotaFiscal"
End Sub

'Private Function BuscaCFOP() As Boolean
'On Error GoTo Err_BuscaCFOP
'Dim lcodigoforma As String
'
'    BuscaCFOP = True
'
'    lbl_cfop = ""
'    For f = 0 To grade1.Rows - 1
'        If IsNumeric(grade1.TextMatrix(f, 28)) Then
'            If InStr(1, lbl_cfop, grade1.TextMatrix(f, 28), vbTextCompare) = 0 Then
'                Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM cadastro_observacoes WHERE cfop = '" & grade1.TextMatrix(f, 28) & "' and status = 0", 1)
'                If gb_Recordset.RecordCount > 0 Then
'                    lobservacao = lobservacao & gb_Recordset!descricao & Chr(10)
'                End If
'                gb_Recordset.Close
'                If lbl_cfop = "" Then
'                    lbl_cfop = grade1.TextMatrix(f, 28)
'                Else
'                    lbl_cfop = lbl_cfop & "/" & grade1.TextMatrix(f, 28)
'                End If
'            End If
'        End If
'    Next
'
'    If IsNumeric(grade1.TextMatrix(1, 28)) Then
'        Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM natureza_operacao WHERE codigo = '" & grade1.TextMatrix(1, 28) & "'", 0)
'        If gb_Recordset.RecordCount > 0 Then lNomeCodificacaoFiscal = gb_Recordset!descricao
'        gb_Recordset.Close
'    End If
'
'Exit Function
'Err_BuscaCFOP: ValidaErros Err, Me.Caption & " - BuscaCFOP"
'End Function

'*****************************************************************************
'Criação: Ronaldo Robledo Mendes Souza                      Data:
'
'Propósito: Efetuar Calculo dos precos/custos contábeis
'Alteração: Ronaldo Robledo                                 Data: 30/05/2011
'           Alterado a linha de gravar apenas custos colocado ela pra fora
'           da condição.
'           Inserido para gravar campo ultimo_precocompracont
'Alteração:Clejunior                                        Data: 23/04/2013
'         :TT3168 "ultimo_precocompra = " & fValidaValor2(grade1.TextMatrix(z, 8)), divido
'         :pelo valor de operação da unidade secundaria
'Alteração:Clejunior                                        Data: 23/04/2013
'         : If IsNumeric(grade1.TextMatrix(z, 45)) Then colocado validação para quando
'         : se utiliza unidade secundaria criado variavel curVerificaFatorConversao
'*****************************************************************************
Private Function CalculoMeiaNota(ByVal z As Long) As Boolean
On Error GoTo Err_CalculoMeiaNota
Dim curVerificaFatorConversao As Currency
    
    CalculoMeiaNota = False
    
     Set gb_Recordset = Conexao.GeraRecordset("SELECT estoque.quantidade,produto.custo_medio_cont FROM estoque,produto WHERE estoque.codigo_do_produto = '" & grade1.TextMatrix(z, 2) & "' and produto.codigo = '" & grade1.TextMatrix(z, 2) & "'", 1)
     If gb_Recordset.RecordCount > 0 Then
         If gb_Recordset!Quantidade > 0 Then
             zQuantidade = gb_Recordset!Quantidade
         Else
             zQuantidade = 0
         End If
         zCustoAnterior = gb_Recordset!Custo_medio_cont
     End If
     gb_Recordset.Close
     
     grade1.TextMatrix(z, 56) = 0
     If CDbl(grade1.TextMatrix(z, 5)) > 0 Then
       'zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + (grade1.TextMatrix(z, 5) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5))), "##,###,##0.0000")
       If CDbl(grade1.TextMatrix(z, 5)) > 0 Then
           If IsNumeric(grade1.TextMatrix(z, 45)) Then
               zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + ((grade1.TextMatrix(z, 5) * (grade1.TextMatrix(z, 45))) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5) * (grade1.TextMatrix(z, 45)))), "##,###,##0.0000")
           Else
               zValorTotal = Format(((CDbl(zQuantidade) * CDbl(zCustoAnterior)) + (grade1.TextMatrix(z, 5) * lcustoprodutos(z))) / (CDbl(zQuantidade) + CDbl(grade1.TextMatrix(z, 5))), "##,###,##0.0000")
           End If
       End If
       g_string4 = zValorTotal
       g_string2 = lcustoprodutos(z)
       grade1.TextMatrix(z, 56) = Format(zValorTotal, g_decimal_custo)
     End If
     
    If Left(cbo_atualizacusto, 1) > 0 Then
             Call Conexao.AlterarRecordset("produto", "custo_anterior_cont = produto.custo_cont,custo_medio_ant_cont = produto.custo_medio_cont", "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
             
             Set gb_Recordset = Conexao.GeraRecordset("SELECT marckup_atacado_cont,desconto_atacado_cont,marckup_varejo_cont,desconto_varejo_cont FROM produto WHERE codigo = '" & grade1.TextMatrix(z, 2) & "'", 1)
             If gb_Recordset.RecordCount > 0 Then
                 x_preco_atacado = gb_Recordset!marckup_atacado_cont
                 x_preco_varejo = gb_Recordset!marckup_varejo_cont
                 x_desc_atacado = gb_Recordset!desconto_atacado_cont
                 x_desc_varejo = gb_Recordset!desconto_varejo_cont
             Else
                 x_preco_atacado = 0
                 x_preco_varejo = 0
                 x_desc_atacado = 0
                 x_desc_varejo = 0
             End If
             gb_Recordset.Close
    
           'grava precos que serão alterados
           'If Left(cbo_atualizacusto, 1) = 1 Then
               'grava precos que serão alterados
            '   If CDbl(x_preco_atacado) > 0 Or CDbl(x_preco_varejo) > 0 Then
                     Call GravaPrecos(grade1.TextMatrix(z, 2), "E")
             '  End If
           'End If
             
           If Left(cbo_atualizacusto, 1) = 1 Then
             If CDbl(x_preco_atacado) > 0 Then
                   'calculo do sem coeficiente
                   x_preco_atacado = (lcustoprodutos(z) * CDbl(x_preco_atacado)) / 100
                   x_preco_atacado = Format(lcustoprodutos(z) + CDbl(x_preco_atacado), g_decimal_venda)
                   
                   g_string = 0
                   g_string = Format((x_desc_atacado * x_preco_atacado) / 100, g_decimal_venda)
                   x_desc_atacado = Format(CDbl(x_preco_atacado) - CDbl(g_string), g_decimal_venda)
                   x_desc_atacado = Format(x_desc_atacado, g_decimal_venda)
       
                   Call Conexao.AlterarRecordset("produto", "preco_atacado_cont = " & fValidaValor2(x_preco_atacado) & ",minimo_atacado_cont = " & fValidaValor2(x_desc_atacado) & ",data_atacado_cont = " & FormataData(Date) & ",alterar = 'S'", "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
             End If
             
             If CDbl(x_preco_varejo) > 0 Then
                   x_preco_varejo = (lcustoprodutos(z) * CDbl(x_preco_varejo)) / 100
                   x_preco_varejo = Format(lcustoprodutos(z) + CDbl(x_preco_varejo), g_decimal_venda)
                   
                   g_string = 0
                   g_string = Format((x_desc_varejo * x_preco_varejo) / 100, g_decimal_venda)
                   x_desc_varejo = Format(CDbl(x_preco_varejo) - CDbl(g_string), g_decimal_venda)
                   x_desc_varejo = Format(x_desc_varejo, g_decimal_venda)
                 
                   Call Conexao.AlterarRecordset("produto", "preco_varejo_cont = " & fValidaValor2(x_preco_varejo) & ",minimo_varejo_cont = " & fValidaValor2(x_desc_varejo) & ",data_varejo_cont = " & FormataData(Date) & ",alterar = 'S'", "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
             End If
           End If
           
           
           '  TT3168 "ultimo_precocompracont = " & fValidaValor2(grade1.TextMatrix(z, 8)),
           ' colocado verificação se a grade e numeria ou não
           
           If IsNumeric(grade1.TextMatrix(z, 45)) Then
              curVerificaFatorConversao = grade1.TextMatrix(z, 45)
           Else
              curVerificaFatorConversao = 1
           End If
    
           Call Conexao.AlterarRecordset("produto", "custo_cont = " & fValidaValor2(g_string2) & "," & _
                                        "custo_medio_cont = " & fValidaValor2(zValorTotal) & "," & _
                                        "data_custo_cont = " & FormataData(Date) & "," & _
                                        "data_custo_medio_cont = " & FormataData(Date) & "," & _
                                        "alterar = 'S'," & _
                                        "ultimo_precocompracont = " & fValidaValor2(grade1.TextMatrix(z, 8) / curVerificaFatorConversao), _
                                        "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
    End If
    
    CalculoMeiaNota = True

Exit Function
Err_CalculoMeiaNota: ValidaErros Err, Me.Caption & " - CalculoMeiaNota"
End Function

Sub grade1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo file
    Select Case KeyCode
        Case &H8 'BACKSPACE
            If Len(grade1.Text) > 0 Then
                grade1.Text = Left(grade1.Text, (Len(grade1.Text) - 1))
            End If
                
        Case &H2E 'DEL
            grade1.Text = ""
            
            If grade1.col = 30 Then
                grade1.TextMatrix(grade1.RowSel, 43) = ""
            End If
            
        Case 27
            celulaanterior
            
        Case vbKeyF4
            ' Excluir linhas selecionadas
            ExcluirLinhas
            'Call SomaProduto
            
        Case 38
            grade1.Row = grade1.Row + 1
            grade1.col = grade1.col
            
        Case 39
            grade1.col = grade1.col - 1
            
        Case 37
            grade1.col = grade1.col + 1
            
        Case 40
            grade1.Row = grade1.Row - 1
    End Select
Exit Sub
file: Resume Next
End Sub

Private Sub grade1_KeyPress(KeyAscii As Integer)
On Error GoTo Err_grade1_KeyPress

    Select Case KeyAscii
        Case Is = 13
            LastRow = grade1.Row
            LastCol = grade1.col

            ProximaCelula
            
            xdigitado = 0
            Text2.Move grade1.CellLeft - Screen.TwipsPerPixelX, grade1.CellTop + 4595 - Screen.TwipsPerPixelY, grade1.CellWidth + Screen.TwipsPerPixelX * 2, grade1.CellHeight + Screen.TwipsPerPixelY * 2
                  
        Case Is <> 8 And 9 And 10 And 13 'não imprimíveis
            If grade1.col = 3 And cDPFFaturamento.DigitacaoGrade = 1 Then
                If xdigitado = 0 Then
                    grade1.Text = ""
                End If
            
                grade1.Text = grade1.Text + Chr(KeyAscii)
                xdigitado = 1
            ElseIf grade1.col < 13 Or grade1.col = 28 Or grade1.col = 42 Or grade1.col = 30 Then
                If xdigitado = 0 Then grade1.Text = ""
                grade1.Text = grade1.Text + Chr(KeyAscii)
                xdigitado = 1
            End If
    End Select

Exit Sub
Err_grade1_KeyPress: ValidaErros Err, Me.Caption & " - grade1_KeyPress"
    
End Sub

Private Sub celulaanterior()
On Error GoTo Err_celulaanterior

    Select Case grade1.col
        Case 1
            If grade1.Row > 0 Then
                grade1.Row = grade1.Row - 1
                grade1.col = 8
            End If
        Case 4
            grade1.col = 1
        
        Case 6
            grade1.col = 4
            
        Case 8
            grade1.col = 7
            
        Case Else
            grade1.col = grade1.col - 1
    End Select

Exit Sub
Err_celulaanterior: ValidaErros Err, Me.Caption & " - celulaanterior"
End Sub

Private Sub DestacarLinha(strCor As String, lngLinha As Long)

    'marca itens da substituicao
        'If grade1.TextMatrix(lngLinha, 12) = "2" Then
           grade1.Row = lngLinha
           grade1.FillStyle = flexFillRepeat
           grade1.col = 0
           grade1.ColSel = grade1.Cols - 1
           grade1.CellBackColor = strCor
        'Else
        If strCor = "&HFFFFFF" Then grade1.FillStyle = flexFillSingle
        'End If
        
End Sub


Private Sub ExcluirLinhas()
On Error GoTo Err_ExcluirLinhas
'Excluir linhas selecionadas
Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Long

    'Não excluir se for a ultima linha
    If grade1.RowSel = grade1.Rows - 1 Then
        grade1.TextMatrix(LastRow, 1) = ""
        grade1.TextMatrix(LastRow, 2) = ""
        grade1.TextMatrix(LastRow, 3) = ""
        grade1.TextMatrix(LastRow, 4) = ""
        ZeraGrade
        Exit Sub
    End If
    If grade1.Row = grade1.Rows - 1 Then
        grade1.TextMatrix(LastRow, 1) = ""
        grade1.TextMatrix(LastRow, 2) = ""
        grade1.TextMatrix(LastRow, 3) = ""
        grade1.TextMatrix(LastRow, 4) = ""
        ZeraGrade
        Exit Sub
    End If

    'Exclui sempre da linha maior par menor
    i = grade1.Row
    j = grade1.RowSel
    
    If i < j Then
        k = i
        i = j
        j = k
    End If
    
    For n = i To j Step -1
        grade1.RemoveItem n
    Next
    
    LastRow = grade1.Rows - 1
    LastCol = 1
    grade1.col = LastCol
    grade1.Row = LastRow
    grade1.RowSel = LastRow
    grade1.ColSel = LastCol

Exit Sub
Err_ExcluirLinhas: ValidaErros Err, Me.Caption & " - ExcluirLinhas"
End Sub

Private Sub ZeraGrade()
On Error GoTo Err_ZeraGrade

    With grade1
        .TextMatrix(LastRow, 2) = 0
        .TextMatrix(LastRow, 5) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 6) = Format(0, "##,###,##0.00000")
        .TextMatrix(LastRow, 7) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 8) = Format(0, "##,###,##0.00000")
        .TextMatrix(LastRow, 9) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 10) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 11) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 13) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 14) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 15) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 16) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 18) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 19) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 20) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 21) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 22) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 23) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 24) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 25) = Format(0, "##,###,##0.00")
        .TextMatrix(LastRow, 26) = 0
        .TextMatrix(LastRow, 27) = 0
        .TextMatrix(LastRow, 29) = 0
        .TextMatrix(LastRow, 33) = 0
        .TextMatrix(LastRow, 34) = 0
        .TextMatrix(LastRow, 35) = 0
        .TextMatrix(LastRow, 36) = 0
        .TextMatrix(LastRow, 37) = 0
        .TextMatrix(LastRow, 39) = 0
        .TextMatrix(LastRow, 42) = 0
        .TextMatrix(LastRow, 51) = 0
        .TextMatrix(LastRow, 52) = 0
        .TextMatrix(LastRow, 53) = 0
    End With

Exit Sub
Err_ZeraGrade: ValidaErros Err, Me.Caption & " - ZeraGrade"
End Sub

'*****************************************************************************
'Alteração: Nayden Luiz                                     Data: 26/01/2012
'         : OS18842-Corrigida validação para quando o produto é Inativo
'Alteração: Clejunior                                       Data: 15/03/2013
'         : mudado Sql_Query para quando for digitado o codigo de barras ele verificar
'         : se o produto esta inativo ou não, estava deichando passar o codigo de barras
'*****************************************************************************
Private Sub BuscaProdutos(strcodigo As String, bytCondicao As Byte)
On Error GoTo Err_BuscaProdutos
Dim strCondicao As String

    '*** Consulta padrão - Código do produto ou código de barras ***
    If bytCondicao = 0 Then
    '*** codigo antigo
    'strCondicao = "(codigo = (SELECT codigo " & _
                ' "FROM codigo_barras WHERE codigo_barras = '" & strcodigo & "' and inativo = 0 LIMIT 1) ) or (codigo = '" & strcodigo & "') and inativo = 0 "
    
    strCondicao = "(codigo = (SELECT CB.codigo " & _
                  "FROM codigo_barras CB INNER JOIN produto PD2 ON PD2.codigo = CB.codigo AND PD2.inativo = 0 " & _
                  "WHERE CB.codigo_barras = '" & strcodigo & "' LIMIT 1) ) or (PD.codigo = '" & strcodigo & "') and PD.inativo = 0 "
                      
    '*** Consulta pelo Código do produto - Fornecedor ***
    ElseIf bytCondicao = 1 Then
            strCondicao = "(codigo = (SELECT codigo FROM (" & _
                          "(SELECT cast(codigo as varchar),(1) as ordem " & _
                          "FROM codigo_barras WHERE codigo_barras = '" & strcodigo & "' LIMIT 1) " & _
                          "Union " & _
                          "(SELECT codigo_produto,(2) as ordem " & _
                          "FROM fornecedor_produtos WHERE codfornecedor = '" & strcodigo & "' LIMIT 1)) " & _
                          " AS VerifcaCod ORDER BY ordem ASC LIMIT 1)) and inativo = 0"
    End If
    '*** Codigo Antigo
'    Sql_Query = "SELECT codigo_grupo,codigo,descricao,unidade," & _
'                "controle_lote,fracionado,utiliza_grade,codificacao_fiscal, " & _
'                "subcodigo,fk_cadastrotributacao,aliquota_estadual,industria_revenda,servico_produto " & _
'                "FROM produto " & _
'                "WHERE " & strCondicao
    
    
    Sql_Query = "SELECT PD.codigo_grupo,PD.codigo,PD.descricao,PD.unidade," & _
                "PD.controle_lote,PD.fracionado,PD.utiliza_grade,PD.codificacao_fiscal, " & _
                "PD.subcodigo,PD.fk_cadastrotributacao,PD.aliquota_estadual,PD.industria_revenda,PD.servico_produto " & _
                "FROM produto PD " & _
                "WHERE " & strCondicao
                
    Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 1)
    If gb_Recordset.RecordCount > 0 Then
         grade1.TextMatrix(LastRow, 38) = ObtenhaCodigoBarras(strcodigo)
         Call DestacarLinha("&HFFFFFF", LastRow)
         TelaProduto
        'VAI PARA PROXIMA COLUNA
         grade1.col = grade1.col + 3
         gb_Recordset.Close
         
        'verifica se o produto tem grade
        If Val(grade1.TextMatrix(LastRow, 39)) = 1 Then
            Call ChamaGradeTamanho(grade1.TextMatrix(LastRow, 2), LastRow)
        End If
    Else
        gb_Recordset.Close
        
        Alerta "O Produto é inexistente ou esta inativo!" & vbCrLf _
             & "Consulte o cadastro do Produto."
        Call ZeraGrade
    End If
    
Exit Sub
Err_BuscaProdutos: ValidaErros Err, Me.Caption & " - BuscaProdutos"
End Sub

'*****************************************************************************
'Criação: Ronaldo Robledo Mendes de Souza                      Data: 28/01/2012
'Alteração: Inserido novo campo subcodigo na consulta
'Alteração: ronaldo Robledo                                          23/03/2013
'           alterado código de busca de pedidos para o método buscaitenspedido
'           buscando da nova tabela de pedido de mercadoria
'*****************************************************************************
Private Sub ProximaCelula()
On Error GoTo file

    
    Select Case grade1.col
        Case 1
            If Trim(grade1.TextMatrix(LastRow, 1)) = "" Then
                If g_entrada_amarrado_pedido = 1 Or g_entrada_amarrado_pedido = 2 Then
                    Call Pedidos
                Else
                    grade1.TextMatrix(LastRow, 1) = 0
                End If
            End If
            grade1.col = grade1.col + 1
        
        Case 2
            If grade1.TextMatrix(LastRow, 1) = "" Or Not IsNumeric(grade1.TextMatrix(LastRow, 1)) Then
                grade1.TextMatrix(LastRow, 1) = 0
            End If
            
            If Trim(grade1.Text) <> "" Then
                If grade1.TextMatrix(LastRow, 1) > 0 Or (g_entrada_amarrado_pedido = 1 Or g_entrada_amarrado_pedido = 2) And pct_icms.Visible = False Then
                    Call BuscaItensPedido(grade1.TextMatrix(LastRow, 1), grade1.TextMatrix(LastRow, 2), LastRow)
                Else
                    Call BuscaProdutos(UCase(grade1.Text), 0)
                End If
            Else
                If pct_icms.Visible = True Or pct_orçamento.Visible = True Then
                    g_string3 = "CONSULTA"
                Else
                    g_string3 = ""
                End If
                g_string = ""
                frm_consulta_produtos.Show 1
                grade1.Text = g_string
                If Len(g_string) > 0 Then Call BuscaProdutos(g_string, 0)
            End If
            
        Case 4
            'If g_unidade_secundaria = 1 Then
                If Trim(grade1.Text) <> "" Then
                    Call CalculaUnidadeSecundaria(CInt(LastRow), True)
                Else
                    Alerta "Unidade informada é inválida!"
                End If
            'End If
        
        'QUANTIDADE
        Case 5
            If IsNumeric(grade1.Text) And IsNumeric(grade1.TextMatrix(LastRow, 8)) Then
                If CDbl(grade1.Text) > 0 Then
                    'se o produto não e fracionado
                    If grade1.TextMatrix(LastRow, 26) <= 0 Then
                        grade1.TextMatrix(LastRow, 5) = Format(Arredonda(grade1.TextMatrix(LastRow, 5)), "##,###,##0.00")
                    End If
                    grade1.TextMatrix(LastRow, 11) = Format(grade1.TextMatrix(LastRow, 8) * grade1.TextMatrix(LastRow, 5), "##,###,##0.00")
                    
                    grade1.col = grade1.col + 1
                Else
                    Alerta "A Quantidade não pode ser Zero!"
                End If
            Else
                Alerta "Quantidade ou Preço Unitário Inválido!"
            End If
        
        'VALOR BRUTO
        Case 6
            If IsNumeric(grade1.Text) And IsNumeric(grade1.TextMatrix(LastRow, 5)) Then
                'se o produto não e fracionado
                If grade1.TextMatrix(LastRow, 26) <= 0 Then
                    grade1.TextMatrix(LastRow, 5) = Format(Arredonda(grade1.TextMatrix(LastRow, 5)), "##,###,##0.00")
                End If
                
                g_string = 0
                grade1.TextMatrix(LastRow, 6) = Format(grade1.TextMatrix(LastRow, 6), g_decimal_compra)
                If CDbl(grade1.TextMatrix(LastRow, 7)) > 0 Then
                    g_string = Format(CDbl(grade1.TextMatrix(LastRow, 6)) * CDbl(grade1.TextMatrix(LastRow, 7)) / 100, g_decimal_compra)
                    grade1.TextMatrix(LastRow, 8) = Format(CDbl(grade1.TextMatrix(LastRow, 6)) - CDbl(g_string), g_decimal_compra)
                Else
                    grade1.TextMatrix(LastRow, 8) = Format(grade1.TextMatrix(LastRow, 6), g_decimal_compra)
                End If
                grade1.TextMatrix(LastRow, 11) = Format(grade1.TextMatrix(LastRow, 5) * grade1.TextMatrix(LastRow, 8), "##,###,##0.00")
                grade1.col = grade1.col + 1
            Else
                Alerta "Valor Bruto Inválido!"
            End If
    
        'Porc.DEsconto
        Case 7
            If IsNumeric(grade1.Text) And IsNumeric(grade1.TextMatrix(LastRow, 6)) Then
                'se o produto não e fracionado
                If grade1.TextMatrix(LastRow, 26) <= 0 Then
                    grade1.TextMatrix(LastRow, 5) = Format(Arredonda(grade1.TextMatrix(LastRow, 5)), "##,###,##0.00")
                End If
                
                g_string = 0
                grade1.TextMatrix(LastRow, 7) = Format(grade1.TextMatrix(LastRow, 7), "##,###,##0.00")
                If CDbl(grade1.TextMatrix(LastRow, 7)) > 0 Then
                    g_string = Format(CDbl(grade1.TextMatrix(LastRow, 6)) * CDbl(grade1.TextMatrix(LastRow, 7)) / 100, g_decimal_compra)
                    grade1.TextMatrix(LastRow, 8) = Format(CDbl(grade1.TextMatrix(LastRow, 6)) - CDbl(g_string), g_decimal_compra)
                Else
                    grade1.TextMatrix(LastRow, 8) = Format(grade1.TextMatrix(LastRow, 6), g_decimal_compra)
                End If
                grade1.TextMatrix(LastRow, 11) = Format(grade1.TextMatrix(LastRow, 5) * grade1.TextMatrix(LastRow, 8), "##,###,##0.00")
                grade1.col = grade1.col + 1
            Else
                Alerta "Porcentagem Desconto Inválido!"
            End If
    
        'PRECO UNITARIO
        Case 8
            If IsNumeric(grade1.Text) And IsNumeric(grade1.TextMatrix(LastRow, 5)) Then
                'se o produto não e fracionado
                If grade1.TextMatrix(LastRow, 26) <= 0 Then
                    grade1.TextMatrix(LastRow, 5) = Format(Arredonda(grade1.TextMatrix(LastRow, 5)), "##,###,##0.00")
                End If
                
                g_string = 0
                grade1.TextMatrix(LastRow, 8) = Format(grade1.TextMatrix(LastRow, 8), g_decimal_compra)
                grade1.TextMatrix(LastRow, 11) = Format(grade1.TextMatrix(LastRow, 5) * grade1.TextMatrix(LastRow, 8), "##,###,##0.00")
                
                grade1.col = grade1.col + 1
            Else
                Alerta "Preço Unitário Inválido!"
            End If
        
        'PORC. IPI
        Case 9
            If IsNumeric(grade1.Text) Then
                'IPI
                grade1.TextMatrix(LastRow, 9) = Format(fValidaValorNovo(grade1.TextMatrix(LastRow, 9)), "##,###,##0.00")
                grade1.col = grade1.col + 1
            Else
                Alerta "Porcentagem IPI Inválido!"
            End If
        
        'PORC. icms
        Case 10
            If IsNumeric(grade1.Text) Then
                'IPI
                grade1.TextMatrix(LastRow, 10) = Format(fValidaValorNovo(grade1.TextMatrix(LastRow, 10)), "##,###,##0.00")
                grade1.col = grade1.col + 2
            Else
                Alerta "Porcentagem ICMS Inválido!"
            End If
                     
        Case 12
            'SE NAO EXISTE PRODUTO NA GRADE INSERI
            grade1.Text = UCase(grade1.Text)
            g_string = grade1.TextMatrix(LastRow, 2)
            g_string2 = grade1.TextMatrix(LastRow, 1)
            
            If Not ExisteProduto Then
                If chkImportacao = 1 Then
                    'abre tela de II
                    cadastro_II.Show vbModal
                End If
                            
                ChamaCelula
                'para alteração do preco de venda manual
                If chk_lancamento_venda.Value = 1 Then
                    xCodigoProduto = g_string
                    Call VerificarCustos
                End If
            Else
                Alerta "Produto já Lançado na Grade!"
            End If
            
        Case 30
            If IsNumeric(grade1.TextMatrix(grade1.RowSel, 30)) Then
                grade1.TextMatrix(grade1.RowSel, 43) = "S"
            End If
            
        Case 42
            grade1.col = 1
            If grade1.Row < grade1.Rows - 1 Then
                grade1.Row = grade1.Row + 1
            End If
        
        Case Else
            If grade1.col < grade1.Cols - 2 Then
                grade1.col = grade1.col + 1
            Else
                grade1.col = 1
                If grade1.Row < grade1.Rows - 1 Then
                    grade1.Row = grade1.Row + 1
                    Call ZeraGrade
                End If
            End If
    End Select
    grade1.SetFocus
    
Exit Sub
file: Resume Next
End Sub

Private Sub TelaProduto()
On Error GoTo Err_TelaProduto

g_string2 = 0

    With gb_Recordset
         'If ValidaProdNFE(!Codigo) Then
            grade1.TextMatrix(LastRow, 2) = !codigo
            grade1.TextMatrix(LastRow, 3) = !Descricao
            grade1.TextMatrix(LastRow, 4) = !Unidade
            grade1.TextMatrix(LastRow, 16) = !aliquota_estadual
            grade1.TextMatrix(LastRow, 21) = 0 'Format(!pis_entrada, "##,###,##0.00")
            grade1.TextMatrix(LastRow, 22) = 0 'Format(!cofins_entrada, "##,###,##0.00")
            grade1.TextMatrix(LastRow, 24) = !controle_lote
            grade1.TextMatrix(LastRow, 25) = !industria_revenda
            grade1.TextMatrix(LastRow, 26) = !Fracionado
            grade1.TextMatrix(LastRow, 29) = !codigo
            grade1.TextMatrix(LastRow, 30) = "000"
            grade1.TextMatrix(LastRow, 33) = 0 'tipo importacao
            grade1.TextMatrix(LastRow, 39) = !utiliza_grade
            grade1.TextMatrix(LastRow, 40) = !codificacao_fiscal
            grade1.TextMatrix(LastRow, 46) = !Subcodigo
            grade1.TextMatrix(LastRow, 47) = VerificaNulo(!fk_cadastrotributacao)
            grade1.TextMatrix(LastRow, 48) = 0 '!cst_pis
            grade1.TextMatrix(LastRow, 49) = 0 '!cst_cofins
         'End If
    End With

Exit Sub
Err_TelaProduto: ValidaErros Err, Me.Caption & " - TelaProduto"
End Sub

Private Sub TrataRecursosConsumo()
On Error GoTo Err_TrataRecursosConsumo
Dim intHeight As Integer
        
    cbo_CodigoConsumo.Clear
    intHeight = 855
    frm_DetalhesContaEnergia.Visible = False
    lbl_TituloCC(1).Caption = "Código do consumo"
    
    Select Case Left(cbo_TipoContaConsumo.Text, 1)
           
           Case 1
                Call CarregaCboCodigoConsumoPadrao
                frm_DetalhesContaEnergia.Visible = True
                intHeight = 1695
           Case 2
                Call CarregaCboCodigoConsumoAgua
           Case 3
                lbl_TituloCC(1).Caption = "Tipo de Assinante"
                Call CarregaCboCodigoConsumoTelefone
           
    End Select
    cbo_CodigoConsumo.Text = cbo_CodigoConsumo.List(0)
    frm_LancamentoConsumo.Height = intHeight

Exit Sub
Err_TrataRecursosConsumo: ValidaErros Err, Me.Caption & " - TrataRecursosConsumo"
End Sub

Private Sub CarregaCboCodigoConsumoPadrao()
On Error GoTo Err_CarregaCboCodigoConsumoPadrao

    cbo_CodigoConsumo.AddItem "01  - Comercial"
    lngIDConsumo(1) = 0
    cbo_CodigoConsumo.AddItem "02  - Consumo Próprio"
    lngIDConsumo(2) = 1
    cbo_CodigoConsumo.AddItem "03  - Iluminação Pública"
    lngIDConsumo(3) = 2
    cbo_CodigoConsumo.AddItem "04  - Industrial"
    lngIDConsumo(4) = 3
    cbo_CodigoConsumo.AddItem "05  - Poder Público"
    lngIDConsumo(5) = 4
    cbo_CodigoConsumo.AddItem "06  - Residencial"
    lngIDConsumo(6) = 5
    cbo_CodigoConsumo.AddItem "07  - Rural"
    lngIDConsumo(7) = 6
    cbo_CodigoConsumo.AddItem "08  - Serviço Público"
    lngIDConsumo(8) = 7
                    
Exit Sub
Err_CarregaCboCodigoConsumoPadrao: ValidaErros Err, Me.Caption & " - CarregaCboCodigoConsumoPadrao"
End Sub

Private Sub CarregaCboCodigoConsumoTelefone()
On Error GoTo Err_CarregaCboCodigoConsumoTelefone

    cbo_CodigoConsumo.AddItem "1 - Comercial/Industrial"
    lngIDConsumo(1) = 0
    cbo_CodigoConsumo.AddItem "2 - Poder Público"
    lngIDConsumo(2) = 1
    cbo_CodigoConsumo.AddItem "3 - Residencial/Pessoa física"
    lngIDConsumo(3) = 2
    cbo_CodigoConsumo.AddItem "4 - Público"
    lngIDConsumo(4) = 3
    cbo_CodigoConsumo.AddItem "5 - Semi-Público"
    lngIDConsumo(5) = 4
    cbo_CodigoConsumo.AddItem "6 - Outros"
    lngIDConsumo(6) = 5
                    
Exit Sub
Err_CarregaCboCodigoConsumoTelefone: ValidaErros Err, Me.Caption & " - CarregaCboCodigoConsumoTelefone"
End Sub

Private Sub CarregaCboCodigoConsumoAgua()
On Error GoTo Err_CarregaCboCodigoConsumoAgua

    cbo_CodigoConsumo.AddItem "00 - Consumo residencial até R$ 50,00"
    lngIDConsumo(0) = 0
    cbo_CodigoConsumo.AddItem "01 - Consumo residencial de R$ 50,01 a R$ 100,00"
    lngIDConsumo(1) = 1
    cbo_CodigoConsumo.AddItem "02 - Consumo residencial de R$ 100,01 a R$ 200,00"
    lngIDConsumo(2) = 2
    cbo_CodigoConsumo.AddItem "03 - Consumo residencial de R$ 200,01 a R$ 300,00"
    lngIDConsumo(3) = 3
    cbo_CodigoConsumo.AddItem "04 - Consumo residencial de R$ 300,01 a R$ 400,00"
    lngIDConsumo(4) = 4
    cbo_CodigoConsumo.AddItem "05 - Consumo residencial de R$ 400,01 a R$ 500,00"
    lngIDConsumo(5) = 5
    cbo_CodigoConsumo.AddItem "06 - Consumo residencial de R$ 500,01 a R$ 1000,00"
    lngIDConsumo(6) = 6
    cbo_CodigoConsumo.AddItem "07 - Consumo residencial acima de R$ 1.000,01"
    lngIDConsumo(7) = 7
    cbo_CodigoConsumo.AddItem "20 - consumo comercial/industrial até R$ 50,00"
    lngIDConsumo(20) = 8
    cbo_CodigoConsumo.AddItem "21 - Consumo comercial/industrial de R$ 50,01 a R$ 100,00"
    lngIDConsumo(21) = 9
    cbo_CodigoConsumo.AddItem "22 - Consumo comercial/industrial de R$ 100,01 a R$ 200,00"
    lngIDConsumo(22) = 10
    cbo_CodigoConsumo.AddItem "23 - Consumo comercial/industrial de R$ 200,01 a R$ 300,00"
    lngIDConsumo(23) = 11
    cbo_CodigoConsumo.AddItem "24 - Consumo comercial/industrial de R$ 300,01 a R$ 400,00"
    lngIDConsumo(24) = 12
    cbo_CodigoConsumo.AddItem "25 - consumo comercial/industrial de R$ 400,01 a R$ 500,00"
    lngIDConsumo(25) = 13
    cbo_CodigoConsumo.AddItem "26 - Consumo comercial/industrial de R$ 500,01 a R$ 1.000,00"
    lngIDConsumo(26) = 14
    cbo_CodigoConsumo.AddItem "27 - Consumo comercial/industrial acima de R$ 1.000,01"
    lngIDConsumo(27) = 15
    cbo_CodigoConsumo.AddItem "80 - Consumo de órgão público"
    lngIDConsumo(80) = 16
    cbo_CodigoConsumo.AddItem "90 - Outros tipos de consumo até R$ 50,00"
    lngIDConsumo(90) = 17
    cbo_CodigoConsumo.AddItem "91 - Outros tipos de consumo de R$ 50,01 a R$ 100,00"
    lngIDConsumo(91) = 18
    cbo_CodigoConsumo.AddItem "92 - Outros tipos de consumo de R$ 100,01 a R$ 200,00"
    lngIDConsumo(92) = 19
    cbo_CodigoConsumo.AddItem "93 - Outros tipos de consumo de R$ 200,01 a R$ 300,00"
    lngIDConsumo(93) = 20
    cbo_CodigoConsumo.AddItem "94 - Outros tipos de consumo de R$ 300,01 a R$ 400,00"
    lngIDConsumo(94) = 21
    cbo_CodigoConsumo.AddItem "95 - Outros tipos de consumo de R$ 400,01 a R$ 500,00"
    lngIDConsumo(95) = 22
    cbo_CodigoConsumo.AddItem "96 - Outros tipos de consumo de R$ 500,01 a R$ 1.000,00"
    lngIDConsumo(96) = 23
    cbo_CodigoConsumo.AddItem "97 - Outros tipos de consumo acima de R$ 1.000,01"
    lngIDConsumo(97) = 24
    cbo_CodigoConsumo.AddItem "99 - Documento fiscal emitido"
    lngIDConsumo(99) = 25
Exit Sub
Err_CarregaCboCodigoConsumoAgua: ValidaErros Err, Me.Caption & " - CarregaCboCodigoConsumoAgua"
End Sub
Private Function Calculaporcentagem()
On Error GoTo Err_Calculaporcentagem
Dim curRateio   As Currency
Dim z           As Long

    Calculaporcentagem = False
    
    Call SomaCalculoPercentual
    
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 0)) Then
            Call Calculo(z)
            If IsNumeric(grade1.TextMatrix(z, 9)) Or IsNumeric(grade1.TextMatrix(z, 10)) Then
                                
                'IPI
                grade1.TextMatrix(z, 53) = grade1.TextMatrix(z, 11) 'base calculo ipi
                
                '*** EMBUTIR FRETE NO IPI ***
                If chk_frete_ipi.Value = 1 Then
                   curRateio = grade1.TextMatrix(z, 11) / SomaPercentual
                   grade1.TextMatrix(z, 53) = grade1.TextMatrix(z, 53) + (curRateio * txt_frete)
                End If
                
                '*** EMBUTIR OUTRAS NO IPI ***
                If chk_outras_ipi.Value = 1 Then
                   curRateio = grade1.TextMatrix(z, 11) / SomaPercentual
                   grade1.TextMatrix(z, 53) = grade1.TextMatrix(z, 53) + (curRateio * txt_outras)
                End If
                
                grade1.TextMatrix(z, 13) = (CDbl(grade1.TextMatrix(z, 9)) * CDbl(grade1.TextMatrix(z, 53))) / 100   'valor do ipi
                grade1.TextMatrix(z, 13) = Format(fValidaValorNovo(grade1.TextMatrix(z, 13)), "##,###,##0.00")
                            
                'ICMS
                grade1.TextMatrix(z, 15) = (CDbl(grade1.TextMatrix(z, 10)) * CDbl(grade1.TextMatrix(z, 11))) / 100
                grade1.TextMatrix(z, 15) = Format(fValidaValorNovo(grade1.TextMatrix(z, 15)), "##,###,##0.00")      'valor do icms
            Else
                Calculaporcentagem = True
                Exit Function
            End If
        End If
    Next

Exit Function
Err_Calculaporcentagem: ValidaErros Err, Me.Caption & " - Calculaporcentagem"
End Function
'*****************************************************************************
'Criação: Ronaldo Robledo Mendes Souza                      Data: 24/11/2010
'
'Propósito: Calcular o peso dos produtos automatico baseado no cadastro de produto
'Alteração: Ronaldo Robledo                                     24/11/2010
'           Inserido checkbox se calcular automatico ou não
'*****************************************************************************
Private Sub CalculaValorPesoPorUnidadeProduto(strcodigo As String)
On Error GoTo Err_CalculaValorPesoPorUnidadeProduto
Dim FatorPesoBruto As Double
Dim FatorPesoLiquido As Double
    
    If chk_calcularpesos.Value = 1 Then
        Sql_Query = ""
        Sql_Query = "SELECT P.Codigo, P.Descricao, P.Unidade AS UnidadePrimaria, P.Peso_Bruto, P.Peso_Liquido, "
        Sql_Query = Sql_Query & "US.Unidade AS UnidadeSecundaria, US.Fator_Operacao "
        Sql_Query = Sql_Query & "FROM Produto P LEFT JOIN Unidade_Secundaria US "
        Sql_Query = Sql_Query & "ON P.Codigo = US.Codigo WHERE P.Codigo = "
        Sql_Query = Sql_Query & "'" & strcodigo & "'"
        
        Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 0)
        If gb_Recordset.RecordCount > 0 Then
            FatorPesoBruto = gb_Recordset!peso_bruto
            FatorPesoLiquido = gb_Recordset!peso_liquido
            Call AjustaFatorOperacaoUnidadeProduto(gb_Recordset, FatorPesoBruto, FatorPesoLiquido)
            
            txt_peso_bruto = Format(CCur(txt_peso_bruto) + (FatorPesoBruto * CCur(grade1.TextMatrix(f, 5))), "##,###,##0.000")
            txt_peso_liquido = Format(CCur(txt_peso_liquido) + (FatorPesoLiquido * CCur(grade1.TextMatrix(f, 5))), "##,###,##0.000")
        End If
        gb_Recordset.Close
    End If

Exit Sub
Err_CalculaValorPesoPorUnidadeProduto: ValidaErros Err, Me.Caption & " - CalculaValorPesoPorUnidadeProduto"
End Sub
Private Sub AjustaFatorOperacaoUnidadeProduto(Registro As ADODB.Recordset, ByRef PesoBruto As Double, ByRef PesoLiquido As Double)
On Error GoTo Err_AjustaFatorOperacaoUnidadeProduto

    PesoBruto = 1
    PesoLiquido = 1
    Do While Not Registro.EOF
        If Not IsNull(Registro!UnidadeSecundaria) Then
            If (Registro!UnidadeSecundaria = Trim(grade1.TextMatrix(f, 4))) Then
                PesoBruto = Registro!fator_operacao
                PesoLiquido = Registro!fator_operacao
                Exit Sub
            End If
        End If
        Registro.MoveNext
    Loop

Exit Sub
Err_AjustaFatorOperacaoUnidadeProduto: ValidaErros Err, Me.Caption & " - AjustaFatorOperacaoUnidadeProduto"
End Sub


Private Sub BuscaConhecimento()
On Error GoTo Err_BuscaConhecimento

    Set gb_Recordset = Conexao.GeraRecordset("SELECT movimento_conhecimento.base_calculo_divisao,movimento_conhecimento.valor_divisao,movimento_cabecalho_conhecimento.aliquota_icms FROM movimento_conhecimento,movimento_cabecalho_conhecimento WHERE movimento_conhecimento.codigo_fornecedor = '" & txt_codigo_fornecedor & "' and movimento_conhecimento.nota_fiscal = '" & txt_numero_nf & "' and movimento_cabecalho_conhecimento.numero_conhecimento = movimento_conhecimento.numero_conhecimento and movimento_cabecalho_conhecimento.condicao = 'EN'", 1)
    If gb_Recordset.RecordCount > 0 Then
        txt_frete_conhecimento = Format(gb_Recordset!valor_divisao, "##,###,##0.00")
        x_aliquota_conhecimento = gb_Recordset!Aliquota_Icms
        x_base_calculo_conhecimento = gb_Recordset!base_calculo_divisao
    End If
    gb_Recordset.Close

Exit Sub
Err_BuscaConhecimento: ValidaErros Err, Me.Caption & " - BuscaConhecimento"
End Sub

Private Function buscasenha()
On Error GoTo Err_buscasenha

    buscasenha = False
    Senha.lbl_serie = "NF"
    Senha.lbl_titulo = "Liberação para Exclusão de Nota Entrada"
    Senha.lbl_liberacao = "Entrada (Exclusão)"
    Senha.lbl_tela = "Entrada de Mercadorias"
    Senha.lbl_historico = "Liberação para Exclusão da Nota N.: " & txt_numero_nf & " Forn. " & txt_fornecedor
    Senha.Show (1)
    If g_string = "OK" Then buscasenha = True

Exit Function
Err_buscasenha: ValidaErros Err, Me.Caption & " - buscasenha"
End Function

Private Function AtualizaContasAPagar() As Boolean
On Error GoTo Err_AtualizaContasAPagar

    If pct_orçamento.Visible = True Then
        lbl_serie_nf = "CX"
    Else
        lbl_serie_nf = "NF"
    End If
    AtualizaContasAPagar = False
    Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM nota_fiscal_parcelamento_entrada_tmp WHERE empresa = '" & cDPEmpresa.codigo & "' and usuario = '" & g_usuario & "'", 0)
    If gb_Recordset.RecordCount > 0 Then
        With gb_Recordset
            Do Until .EOF
                'variaveis para atualizar no cabeçalho da entrada
                lDataCaixa = !Data_Caixa
                lTipoCaixa = Left(!Tipo_Caixa, 1)
                'atualizacaixa dinheiro - cheque a vista
                If !tipo_documento = 100 Or !tipo_documento = 101 Then
                   If chk_atualiza_caixa.Value = 1 Then
                        If !Tipo_Caixa = "Geral" Then
                             If Not AtualizaCaixaGeral Then
                                 .Close
                                 Exit Function
                             End If
                        Else
                             If Not AtualizaCaixa Then
                                 .Close
                                 Exit Function
                             End If
                         End If
                    End If
                Else
                    If chk_atualiza_carteira.Value = 1 Then
                            BD_Record_SetII.Source = "SELECT descricao FROM tipo_documento_pagar WHERE codigo = '" & !tipo_documento & "'"
                            BD_Record_SetII.Open
                                    If BD_Record_SetII.RecordCount > 0 Then
                                        g_string = Mid(BD_Record_SetII!Descricao, 5, 30)
                                    End If
                            BD_Record_SetII.Close
                            
                            With ContasAPagar
                                .Empresa = cDPEmpresa.codigo
                                .NumeroDocumento = gb_Recordset!numero_da_duplicata
                                .codigo = txt_codigo_fornecedor.Text
                                .NOME = txt_fornecedor.Text
                                .TipoDocumento = gb_Recordset!tipo_documento
                                .NomeDocumento = g_string
                                .DataEntrada = EnviaData(msk_entrada.Text)
                                .DataEmissao = EnviaData(msk_emissao.Text)
                                .DataVencimento = EnviaData(gb_Recordset!data_do_vencimento)
                                .DataProrrogacao = EnviaData(gb_Recordset!data_do_vencimento)
                                .DataPrevisao = "00:00:00"
                                .Valor = gb_Recordset!valor_do_vencimento
                                .Desconto = 0
                                .ValorJuros = 0
                                .ValorRestante = gb_Recordset!valor_do_vencimento
                                .CodigoPortador = gb_Recordset!codigo_portador
                                .nomeportador = gb_Recordset!nome_portador
                                .NumeroConta = xplanoconta
                                .nomeconta = txt_fornecedor.Text
                                .NumeroNf = txt_numero_nf.Text
                                .SerieNf = lbl_serie_nf
                                .Perdidos = ""
                                .Observacoes = gb_Recordset!Observacoes
                                .CodigoUsuario = g_usuario
                                .NomeUsuario = g_nome_usuario
                                .Data = Date
                                .Time = Time
                                .CodCusto = gb_Recordset!centro_custo
                            End With
                            
                            If Not ContasAPagar.Incluir(0) Then
                                gb_Recordset.Close
                                Exit Function
                            End If
                    End If
                End If
                            
                .MoveNext
            Loop
        End With
    End If
    gb_Recordset.Close
    
    AtualizaContasAPagar = True

    Call Conexao.DeleteSintetico("nota_fiscal_parcelamento_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and usuario = '" & g_usuario & "'", 0)

Exit Function
Err_AtualizaContasAPagar: ValidaErros Err, Me.Caption & " - AtualizaContasAPagar"
End Function

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start:
' Alteração: Ronaldo Robledo                                    05/03/2013
'         Inserido variavel bolpedidotransf para nao efetuar checagem de pedido
'        mercadoria quando originado de pedido de transferencia
'==========================================================================
Private Function VerificaPedido()
On Error GoTo Err_VerificaPedido
Dim strSQL As String

    VerificaPedido = False
    If bolPedidoTransf Then Exit Function
    If (g_entrada_amarrado_pedido = 1 Or g_entrada_amarrado_pedido = 2) And pct_icms.Visible = False Then
        For f = 0 To grade1.Rows - 1
            If IsNumeric(grade1.TextMatrix(f, 0)) Then
'                strSql = ""
'                strSql = strSql & "SELECT MP.* "
'                strSql = strSql & "FROM movimento_pedido_mercadoria MP,movimento_cabecalho_pedido MC "
'                strSql = strSql & "WHERE MC.numero_pedido = '" & grade1.TextMatrix(f, 1) & "' and"
'                strSql = strSql & "MC.codigo_fornecedor = '" & txt_codigo_Fornecedor & "' and"
'                strSql = strSql & "MP.numero_pedido = MC.numero_pedido and  MP.codigo_do_produto = '" & grade1.TextMatrix(f, 2) & "' "
'                strSql = strSql & "and MP.status = 'A'"
'                Set gb_Recordset = Conexao.GeraRecordset(, 1)
'                If gb_Recordset.RecordCount = 0 Then
                If Val(grade1.TextMatrix(f, 54)) = 0 Then
                    Alerta "Produto N. " & grade1.TextMatrix(f, 2) & " Não Localizado no Pedido Mercadoria!"
                    VerificaPedido = True
                    gb_Recordset.Close
                    Exit Function
                End If
'                End If
'                gb_Recordset.Close
            End If
        Next
    End If

Exit Function
Err_VerificaPedido: ValidaErros Err, Me.Caption & " - VerificaPedido"
End Function

'==========================================================================
' Purpose:  Atualize quantidade recebida e situacao do item do pedio de compra
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 23/03/2013
'==========================================================================
Private Function AtualizaPedidoMercadoria(ByVal z As Long) As Boolean
On Error GoTo Err_AtualizaPedidoMercadoria
Dim strSQL As String
    
    AtualizaPedidoMercadoria = False
    'se tem pedido e encontrou o id do item do pedido
    If grade1.TextMatrix(z, 1) > 0 And Val(grade1.TextMatrix(z, 54)) > 0 Then
    
        Call Conexao.AlterarRecordset("tb_item_pedidocompra", _
                                    "quantidade_recebida = quantidade_recebida + " & fValidaValor2(grade1.TextMatrix(z, 5)), _
                                    "fk_empresa = '" & cDPEmpresa.codigo & "' and " & _
                                    "fk_item_movimentacao = '" & grade1.TextMatrix(z, 54) & "' ", _
                                    cDPEmpresa.codigo)
            
        strSQL = ""
        strSQL = strSQL & "SELECT quantidade "
        strSQL = strSQL & "FROM tb_item_movimentacao "
        strSQL = strSQL & "WHERE fk_empresa = '" & cDPEmpresa.codigo & "' and "
        strSQL = strSQL & "pk_item_movimentacao = '" & grade1.TextMatrix(z, 54) & "' and "
        strSQL = strSQL & "fk_produto = '" & grade1.TextMatrix(z, 2) & "' "
        Set gb_Recordset = Conexao.GeraRecordset(strSQL, 1)
        If gb_Recordset.RecordCount > 0 Then
            Call Conexao.AlterarRecordset("tb_item_pedidocompra", "situacao_item = '" & eTipoSituacaoItemPedido.Baixado & " '", _
                                        "fk_empresa = '" & cDPEmpresa.codigo & "' and " & _
                                        "quantidade_recebida >= " & cUtGeral.fValidaValor2(gb_Recordset!Quantidade) & " " & _
                                        "and fk_item_movimentacao = '" & grade1.TextMatrix(z, 54) & "'", _
                                        cDPEmpresa.codigo)
            gb_Recordset.Close
        Else
            Alerta "Item " & grade1.TextMatrix(z, 2) & " do Pedido de Compra " & grade1.TextMatrix(z, 1) & _
            "Mercadoria Não Localizada para atualização"
            gb_Recordset.Close
            Exit Function
        End If
        
        'verifica se todos os itens do pedido de compra estão fechado
        'para poder fechar o cabecalho do pedido
        strSQL = ""
        strSQL = strSQL & "SELECT situacao_item "
        strSQL = strSQL & "FROM tb_pedido_compra TPC "
        strSQL = strSQL & "LEFT JOIN tb_item_pedidocompra TIPC ON TIPC.fk_empresa = TPC.fk_empresa and "
        strSQL = strSQL & "TIPC.fk_movimentacao = TPC.fk_movimentacao "
        strSQL = strSQL & "WHERE TPC.fk_empresa = '" & cDPEmpresa.codigo & "' and "
        strSQL = strSQL & "TPC.numeropedido = '" & grade1.TextMatrix(z, 1) & "' "
        strSQL = strSQL & "and TIPC.situacao_item =  '" & eTipoSituacaoItemPedido.Aberto & "'"
        Set gb_Recordset = Conexao.GeraRecordset(strSQL, 1)
        If gb_Recordset.RecordCount = 0 Then
            If Not Conexao.AlterarRecordset("tb_pedido_compra", "situacao = '" & eTipoSituacaoPedido.Fechado & " '", _
                                        "fk_empresa = '" & cDPEmpresa.codigo & "' and " & _
                                        " numeropedido = '" & grade1.TextMatrix(z, 1) & "'", _
                                        cDPEmpresa.codigo) Then
                gb_Recordset.Close
                Alerta "Inconsistência no fechamento do pedido de mercadoria!"
                Exit Function
            End If
        End If
        gb_Recordset.Close

    End If
    AtualizaPedidoMercadoria = True

Exit Function
Err_AtualizaPedidoMercadoria: ValidaErros Err, Me.Caption & " - AtualizaPedidoMercadoria"
End Function

Private Sub ImpressaoEspelho()
On Error GoTo Err_ImpressaoEspelho

    frm_seleciona_impressora.Show 1
    Set Printer = Printers(Impressora_default)
    
    Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM dados_impressora", 1)
    If gb_Recordset.RecordCount > 0 Then
            'para tamanho 10 matricial
            If g_tipo_impressora = 0 Then
                g_tamanho_pagina = gb_Recordset!linha_matricial
                g_tamanho_pagina_final = gb_Recordset!final_matricial
            'para tamanho 9 jato tinta
            Else
                g_tamanho_pagina = gb_Recordset!linha_tinta4
                g_tamanho_pagina_final = gb_Recordset!final_tinta4
            End If
    End If
    gb_Recordset.Close
    
    lPagina = 0
    ImpCabEsp
    ImpDadosEsp
    'ImpRodape
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    BioFechaImprime1
    frm_preview.txt_preview.FileName = "c:\sabre\Arquivo" & g_usuario & ".txt"
    frm_preview.Show 1

Exit Sub
Err_ImpressaoEspelho: ValidaErros Err, Me.Caption & " - ImpressaoEspelho"
End Sub

Private Sub Somatoria()
On Error GoTo Err_Somatoria

    If CDbl(grade1.TextMatrix(f, 10)) > 0 Then
        xbasecalculo = lSubTotal
    Else
        xbasecalculo = 0
        xValorIcms = 0
    End If
    
    If grade1.TextMatrix(f, 12) <> "S" Then
        If chk_frete_nota.Value = 1 Then
            xbasecalculo = xbasecalculo + lFrete(lcodigoaliquota)
        End If
        
        If chk_soma_outras.Value = 1 Then
            xbasecalculo = xbasecalculo + lOutras(lcodigoaliquota)
        End If
        
        If chk_soma_ipi.Value = 1 Then
            xbasecalculo = xbasecalculo + lValorIPI
        End If
        
        If chk_seguro.Value = 1 Then
            xbasecalculo = xbasecalculo + lSeguro(lcodigoaliquota)
        End If
        
        If chk_desc_bc.Value = 1 And xbasecalculo > 0 Then
            xbasecalculo = xbasecalculo - lDesconto(lcodigoaliquota)
        End If
    Else
        If chk_frete_red.Value = 1 Then
            xbasecalculo = xbasecalculo + lFrete(lcodigoaliquota)
        End If
        
        If chk_outras_red.Value = 1 Then
            xbasecalculo = xbasecalculo + lOutras(lcodigoaliquota)
        End If
        
        If chk_ipi_red.Value = 1 Then
            xbasecalculo = xbasecalculo + lValorIPI
        End If
        
        If chk_desc_red.Value = 1 Then
            xbasecalculo = xbasecalculo - lDesconto(lcodigoaliquota)
        End If
        
        If chk_seguro_red.Value = 1 Then
            xbasecalculo = xbasecalculo + lSeguro(lcodigoaliquota)
        End If
    End If

Exit Sub
Err_Somatoria: ValidaErros Err, Me.Caption & " - Somatoria"
End Sub

Private Sub SomatoriaSubst()
On Error GoTo Err_SomatoriaSubst

    xbcsubstituicao = lSubTotal
    
    If chk_ipi_subs.Value = 1 Then xbcsubstituicao = xbcsubstituicao + lValorIPI
    
    If chk_frete_subst.Value = 1 Then xbcsubstituicao = xbcsubstituicao + lFreteConhecimento(lcodigoaliquota)
    
    If chk_outras_despesas.Value = 1 Then xbcsubstituicao = xbcsubstituicao + lOutras(lcodigoaliquota)
    
    If chk_fretenota_subst.Value = 1 Then xbcsubstituicao = xbcsubstituicao + lFrete(lcodigoaliquota)

Exit Sub
Err_SomatoriaSubst: ValidaErros Err, Me.Caption & " - SomatoriaSubst"
End Sub

Private Sub txt_porc_red_icms_GotFocus()
    txt_porc_red_icms.BackColor = 12648447
    txt_porc_red_icms.SelStart = 0
    txt_porc_red_icms.SelLength = Len(txt_porc_red_icms)
End Sub

Private Sub txt_porc_red_icms_LostFocus()
    txt_porc_red_icms = UCase(Trim(txt_porc_red_icms))
    txt_porc_red_icms.BackColor = &H8000000E
End Sub

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 04/06/2013
'==========================================================================
Private Function RetornaEstoque() As Boolean
On Error GoTo Err_RetornaEstoque
    
    RetornaEstoque = False
    Set gb_Recordset = Conexao.GeraRecordset("SELECT influencia_estoque FROM movimento_nota_fiscal_entrada Where empresa = '" & cDPEmpresa.codigo & "' and sequencia = '" & lSequencia & "'", 1)
    If gb_Recordset.RecordCount > 0 Then
        If UCase(Trim(gb_Recordset!influencia_estoque)) = "S" Then
            cDPFFaturamento.InfluenciaEstoque = 1
        Else
            cDPFFaturamento.InfluenciaEstoque = 0
        End If
        gb_Recordset.Close
    Else
        Alerta "Produtos da Nota Não Localizado!"
        gb_Recordset.Close
        Exit Function
    End If
    
    If cDPFFaturamento.InfluenciaEstoque = 1 Then
        If Not BaixaEstoque Then Exit Function
        
        If g_utiliza_estoque_deposito = 1 And pct_icms.Visible = False Then
            If Not EntradaSaidaEstoqueDeposito(0, lSequencia, "E", 1, "-", True) Then Exit Function
        End If
        
        'retorno de estoque grade de produtos
        If Not RetornoEstoqueGrade(lSequencia, "E") Then Exit Function
        If Not ExclusaoGrade(lSequencia, "E") Then Exit Function
        
    End If
    
    Dim z As Long
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 1)) And Trim(grade1.TextMatrix(z, 2)) <> "" And Val(grade1.TextMatrix(z, 54)) > 0 Then
            'Call Conexao.AlterarRecordset("movimento_pedido_mercadoria", "quantidade_recebida = quantidade_recebida - " & fValidaValor2(grade1.TextMatrix(z, 5)) & ",status = 'A'", "empresa = '" & cDPEmpresa.Codigo & "' and codigo_do_produto = '" & grade1.TextMatrix(z, 2) & "' and numero_pedido = '" & grade1.TextMatrix(z, 1) & "'", cDPEmpresa.Codigo)
            If Not AbateQuantidadeItemPedidoCompra(grade1.TextMatrix(z, 5), grade1.TextMatrix(z, 54), grade1.TextMatrix(z, 1)) Then Exit Function
        End If
    Next
    
    RetornaEstoque = True

Exit Function
Err_RetornaEstoque: ValidaErros Err, Me.Caption & " - RetornaEstoque"
End Function

'==========================================================================
' Purpose:  Retorna o custo do produto procurando qual foi a ultima nota a entrar no sistema
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 04/06/2013
'==========================================================================
Private Function RetornaCustoProduto() As Boolean
On Error GoTo Err_RetornaCustoProduto
Dim z As Long

    RetornaCustoProduto = False
    
    'Linha de codigo para atualiza produto para calculo de custo medio
    If Left(cbo_atualizacusto, 1) > 0 Then
        For z = 0 To grade1.Rows - 1
            If IsNumeric(grade1.TextMatrix(z, 1)) Then
                'so atualiza em caso de compra
                'If Left(Trim(cbo_tipo_documento), 2) >= 1 And Left(Trim(cbo_tipo_documento), 2) <= 2 Then
                        Set gb_Recordset = Conexao.GeraRecordset("SELECT MN.numero,MN.preco_custo,MN.preco_custo_medio,MN.preco_custo_medio_cont,MN.outros,MN.codigo_do_produto,MN.data_de_entrada,MN.valor_unitario FROM movimento_nota_fiscal_entrada MN,movimento_cabecalho_nota_fiscal_entrada MCN " & _
                                                                 "WHERE MCN.empresa = '" & cDPEmpresa.codigo & "' and MCN.atualiza_custo > 0 and MN.empresa = MCN.empresa and MN.sequencia = MCN.sequencia and MN.codigo_do_produto = '" & grade1.TextMatrix(z, 2) & "' and MN.outros = '" & lbl_nota & "' ORDER BY MN.data_de_entrada DESC,MN.preco_custo DESC", 1)
                        If gb_Recordset.RecordCount > 0 Then
                            If Trim(gb_Recordset!Outros) = "NF" Then
                                Call Conexao.AlterarRecordset("produto", "custo = " & fValidaValor2(gb_Recordset!preco_custo) & ",custo_medio = " & fValidaValor2(gb_Recordset!preco_custo_medio) & ",custo_cont = " & fValidaValor2(gb_Recordset!preco_custo) & ",custo_medio_cont = " & fValidaValor2(gb_Recordset!preco_custo_medio_cont) & ",data_custo = " & FormataData(gb_Recordset!data_de_entrada) & ",data_custo_cont = " & _
                                                              FormataData(gb_Recordset!data_de_entrada) & ",data_custo_medio = " & FormataData(gb_Recordset!data_de_entrada) & ",data_custo_medio_cont = " & FormataData(gb_Recordset!data_de_entrada) & _
                                                              ",ultimo_precocompra = " & fValidaValor2(gb_Recordset!Valor_Unitario) & _
                                                              ",ultimo_precocompracont = " & fValidaValor2(gb_Recordset!Valor_Unitario), _
                                                              "codigo = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo)
                            ElseIf Trim(gb_Recordset!Outros) = "CX" Then
                                Call Conexao.AlterarRecordset("produto", "custo = " & fValidaValor2(gb_Recordset!preco_custo) & _
                                                               ",custo_medio = " & fValidaValor2(gb_Recordset!preco_custo_medio) & _
                                                               ",data_custo = " & FormataData(gb_Recordset!data_de_entrada) & _
                                                               ",data_custo_medio = " & FormataData(gb_Recordset!data_de_entrada) & _
                                                               ",ultimo_precocompra = " & fValidaValor2(gb_Recordset!Valor_Unitario), _
                                                               "codigo = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo)
                            Else
                                Call Conexao.AlterarRecordset("produto", "custo_cont = " & fValidaValor2(gb_Recordset!preco_custo) & _
                                                              ",custo_medio_cont = " & fValidaValor2(gb_Recordset!preco_custo_medio_cont) & _
                                                              ",data_custo_cont = " & FormataData(gb_Recordset!data_de_entrada) & _
                                                              ",data_custo_medio_cont = " & FormataData(gb_Recordset!data_de_entrada) & _
                                                              ",ultimo_precocompracont = " & fValidaValor2(gb_Recordset!Valor_Unitario), _
                                                              "codigo = '" & gb_Recordset!codigo_do_produto & "'", cDPEmpresa.codigo)
                            End If
                            gb_Recordset.Close
                        Else
                            gb_Recordset.Close
                            Set gb_Recordset = Conexao.GeraRecordset("SELECT custo,custo_anterior,custo_medio_anterior,custo_medio,custo_cont," & _
                                                                    "custo_anterior_cont,custo_medio_ant_cont,custo_medio_cont, " & _
                                                                    "ultimo_precocompra,ultimo_precocompracont " & _
                                                                    "FROM produto WHERE codigo = '" & grade1.TextMatrix(z, 2) & "'", 1)
                            If gb_Recordset.RecordCount > 0 Then
                                If Trim(lbl_nota) = "NF" Then
                                    Call Conexao.AlterarRecordset("produto", "custo = " & fValidaValor2(gb_Recordset!custo_anterior) & ",custo_medio = " & fValidaValor2(gb_Recordset!custo_medio_anterior) & ",custo_cont = " & fValidaValor2(gb_Recordset!custo_anterior_cont) & ",custo_medio_cont = " & fValidaValor2(gb_Recordset!custo_medio_ant_cont) & ",data_custo = " & FormataData(Date) & ",data_custo_cont = " & _
                                                                            FormataData(Date) & ",data_custo_medio = " & FormataData(Date) & ",data_custo_medio_cont = " & FormataData(Date) & _
                                                                            ",ultimo_precocompra = " & fValidaValor2(0) & _
                                                                            ",ultimo_precocompracont = " & fValidaValor2(0), _
                                                                            "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                                ElseIf Trim(lbl_nota) = "CX" Then
                                    Call Conexao.AlterarRecordset("produto", "custo = " & fValidaValor2(gb_Recordset!custo_anterior) & _
                                                                 ",custo_medio = " & fValidaValor2(gb_Recordset!custo_medio_anterior) & _
                                                                 ",data_custo = " & FormataData(Date) & _
                                                                 ",data_custo_medio = " & FormataData(Date) & _
                                                                 ",ultimo_precocompra = " & fValidaValor2(0), _
                                                                 "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                                Else
                                    Call Conexao.AlterarRecordset("produto", "custo_cont = " & fValidaValor2(gb_Recordset!custo_anterior_cont) & _
                                                                  ",custo_medio_cont = " & fValidaValor2(gb_Recordset!custo_medio_ant_cont) & _
                                                                  ",data_custo_cont = " & FormataData(Date) & _
                                                                  ",data_custo_medio_cont = " & FormataData(Date) & _
                                                                  ",ultimo_precocompracont = " & fValidaValor2(0), _
                                                                  "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
                                End If
                            End If
                            gb_Recordset.Close
                        End If
            End If
        Next
    End If

    RetornaCustoProduto = True

Exit Function
Err_RetornaCustoProduto: ValidaErros Err, Me.Caption & " - RetornaCustoProduto"
End Function
'==========================================================================
' Purpose: verificar e retornar lançamentos no caixa e financeiro
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 04/06/2013
'==========================================================================
Private Function RetornaFinanceiro(ByVal bolExclusaodeAlteracao As Boolean) As Boolean
On Error GoTo Err_RetornaFinanceiro
Dim strNota         As String
Dim strInformacao   As String

    RetornaFinanceiro = False
    
    If bolExisteLancFinanceiro Then
        'se for diferente de CX vai tudo como NF pois no contas a pagar esta tudo como NF ou CX
        If lbl_nota <> "CX" Then strNota = "NF" Else strNota = "CX"
        Call ContasAPagar.InsertExclusão("empresa = '" & cDPEmpresa.codigo & "' and codigo = '" & txt_codigo_fornecedor & "' and numero_nf = '" & txt_numero_nf & "' and serie_nf = '" & strNota & "'", cDPEmpresa.codigo)
        Call Conexao.DeleteSintetico("contas_apagar", "empresa = '" & cDPEmpresa.codigo & "' and codigo = '" & txt_codigo_fornecedor & "' and numero_nf = '" & txt_numero_nf & "' and serie_nf = '" & strNota & "'", cDPEmpresa.codigo)
        
        If Not bolExclusaodeAlteracao Then strInformacao = "'Exclusão " Else strInformacao = "'Alteração "
        
        Call Conexao.InserirRecordset("log_senhas", "data,hora,codigo_usuario,nome_usuario,historico,tela,observacoes,outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "'," & strInformacao & " do Contas a Pagar Nº da Nota " & txt_numero_nf & " Forn. " & txt_fornecedor & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)
    End If

    'efetua a alteração apenas se o lançamento for o do mesmo dia
    'para não influenciar em caixas já fechados.
    If bolExisteLancCaixa And msk_entrada = g_data_geral Then
        If lTipoCaixa = "D" Then
            Call Conexao.DeleteSintetico("movimento_caixa", "empresa = '" & cDPEmpresa.codigo & "' and numero_do_movimento = " & lLancamento & " and Data = " & FormataData(str(lDataCaixa)), cDPEmpresa.codigo)
        Else
            Call Conexao.DeleteSintetico("movimento_caixa_geral", "codigo_empresa = '" & cDPEmpresa.codigo & "' and  numero_do_movimento = " & lLancamento & " and Data = " & FormataData(str(lDataCaixa)), cDPEmpresa.codigo)
        End If
        lLancamento = 0
    End If
    
    RetornaFinanceiro = True
    
Exit Function
Err_RetornaFinanceiro: ValidaErros Err, Me.Caption & " - RetornaFinanceiro"
End Function
'==========================================================================
' Purpose:  Efetua validações no financeiro para verificar se o financeiro lançado
'           pela nota pode sofrer alterações tanto na exclusão como alteração da nota
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 04/06/2013
'==========================================================================
Private Function ValidaRetornoFinanceiro(ByVal strNota As String) As Boolean
On Error GoTo Err_ValidaRetornoFinanceiro
Dim strSQL  As String
    
    ValidaRetornoFinanceiro = True
    If l_opcao = 1 Then Exit Function
    
    strSQL = ""
    strSQL = strSQL & "SELECT valor,valor_restante, ('A') as status "
    strSQL = strSQL & "FROM contas_apagar CAP "
    strSQL = strSQL & "WHERE empresa = '" & cDPEmpresa.codigo & "' and "
    strSQL = strSQL & "     codigo = '" & txt_codigo_fornecedor & "' and "
    strSQL = strSQL & "     numero_nf = '" & txt_numero_nf & "' and serie_nf = '" & strNota & "' UNION ALL "
    strSQL = strSQL & "SELECT valor_a_pagar,valor_pago,('B') as status "
    strSQL = strSQL & "FROM contas_pagas CPS "
    strSQL = strSQL & "WHERE empresa = '" & cDPEmpresa.codigo & "' and "
    strSQL = strSQL & "     codigo = '" & txt_codigo_fornecedor & "' and "
    strSQL = strSQL & "     numero_nf = '" & txt_numero_nf & "' and serie_nf = '" & strNota & "'"
    Set gb_Recordset = Conexao.GeraRecordset(strSQL, 0)
    If gb_Recordset.RecordCount > 0 Then
        Do Until gb_Recordset.EOF
            If gb_Recordset!Status = "B" Then
                Alerta "Baixa já efetuada no financeiro - alteração do contas a pagar terá que ser feita manualmente"
                ValidaRetornoFinanceiro = False
                Exit Do
            ElseIf gb_Recordset!Valor <> gb_Recordset!valor_restante Then
                Alerta "Baixa parcial já efetuada no financeiro - alteração do contas a pagar terá que ser feita manualmente"
                ValidaRetornoFinanceiro = False
                Exit Do
            End If
            gb_Recordset.MoveNext
        Loop
    End If
    gb_Recordset.Close
    
    
Exit Function
Err_ValidaRetornoFinanceiro: ValidaErros Err, Me.Caption & " - ValidaRetornoFinanceiro"
End Function

'*****************************************************************************
'Criação: Ronaldo Robledo Mendes Souza                      Data:
'
'Propósito: Estorno das tabelas que movimentam na entrada do produto
'Alteração: Ronaldo Robledo                                         30/05/2011
'           Inserido atualizar campos dos produtos ultimo_precocompra
'Alteeração: Ronaldo Robledo                                        06/06/2013
'           Desmembrado método para atender alteração da nota
'           Inserido variavel bolExclusaodeAlteracao que informa se a exclusão e referente
'           alteracao da nota
'*****************************************************************************
Private Function ExclusaoNota(ByVal bolExclusaodeAlteracao As Boolean) As Boolean
On Error GoTo Err_ExclusaoNota
Dim z   As Long

    ExclusaoNota = False
    
    '*** Se é entrada para conferência ,efetua exclusão ***
    If bolEntradaTMP = True Then
       Call Conexao.DeleteSintetico("movimento_cabecalho_nota_fiscal_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, cDPEmpresa.codigo)
       Call Conexao.DeleteSintetico("movimento_nota_fiscal_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, cDPEmpresa.codigo)
       Call Conexao.DeleteSintetico("conferencia_mercadoria_entrada", "empresa = '" & cDPEmpresa.codigo & "' and sequencia_nf = " & lSequencia, cDPEmpresa.codigo)
       Call Conexao.InserirRecordset("log_senhas", "data,hora,codigo_usuario,nome_usuario,historico,tela,observacoes,outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "','Exclusão Nota P/Conferência Nº " & txt_numero_nf & " Forn. " & txt_fornecedor & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)
       ExclusaoNota = True
       Exit Function
    End If
        
    'verifica se exclusão de alteracao ou apenas exclusão geral
    If (bolExclusaodeAlteracao And bolAlteraEstoquenaAlteracaoNF) Or Not bolExclusaodeAlteracao Then
        If Not RetornaEstoque Then Exit Function
        Call Conexao.DeleteSintetico("movimento_transito_compra", "empresa = '" & cDPEmpresa.codigo & "' and sequencia_nf = " & lSequencia, cDPEmpresa.codigo)
    End If
    
    If Not bolExclusaodeAlteracao Then         '*** Deleta Movimento do Lote ***
        If lngSequenciaControle > 0 Then Call cMapLote.RetornaLote(lngSequenciaControle)
    End If
    
    Call Conexao.DeleteSintetico("movimento_cabecalho_nota_fiscal_entrada", "empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, cDPEmpresa.codigo)
    Call Conexao.DeleteSintetico("movimento_impostos", "fk_movnotafiscalentrada IN (SELECT pk_codigo FROM movimento_nota_fiscal_entrada WHERE empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia & ")", cDPEmpresa.codigo)
    Call Conexao.DeleteSintetico("movimento_nota_fiscal_entrada", "empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, cDPEmpresa.codigo)
    Call Conexao.DeleteSintetico("info_conta_consumo", "empresa = '" & cDPEmpresa.codigo & "' AND codigo = " & lngCodigoContaConsumo, cDPEmpresa.codigo)
    
    'so efetua alteração para exclusão normal
    If Not bolExclusaodeAlteracao Then If Not RetornaCustoProduto Then Exit Function
    
    If Not bolExclusaodeAlteracao Then
        If ValidaRetornoFinanceiro(IIf((lbl_nota <> "CX"), "NF", "CX")) Then If Not RetornaFinanceiro(bolExclusaodeAlteracao) Then Exit Function
    ElseIf bolLiberaAlteracaoFinanceiro Then
        If Not RetornaFinanceiro(bolExclusaodeAlteracao) Then Exit Function
    End If
             
    ExclusaoNota = True

Exit Function
Err_ExclusaoNota: ValidaErros Err, Me.Caption & " - ExclusaoNota"
End Function
'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author:
' Alteração: Clejunior Freitas                               Start:19/03/2013
'            validação do controle de lote verifica se esta marcado verifica
'            lote nos parametros do sistema!
' Alteração: Ronaldo Robledo                                        06/06/2013
'            Inserido RetornaLote para garantir limpeza dos lotes incluidos no checalote
'==========================================================================
Private Function IniciaChecagemLote() As Boolean
On Error GoTo Err_IniciaChecagemLote
Dim f As Integer

    If g_controle_lote = 1 Then
       IniciaChecagemLote = False
       Conexao.BeginTrans
       If ChecaLote(f) Then
           Conexao.CommitTrans
           If Not ChecaLoteManual(f) Then
                If l_opcao = 1 Then cMapLote.RetornaLote (lngSequenciaControle): Exit Function Else Exit Function
           End If
           IniciaChecagemLote = True
       Else
           Conexao.RollbackTrans
       End If
    Else
        IniciaChecagemLote = True
    End If
    
Exit Function
Err_IniciaChecagemLote: ValidaErros Err, Me.Caption & " - IniciaChecagemLote"
End Function

'Private Sub InconsistenciaProdutoCadastrado()
'
'    If Trim(grade1.TextMatrix(grade1.RowSel, 2)) = "" Or Trim(grade1.TextMatrix(grade1.RowSel, 2)) = "0" Then
'        Alerta "O Produto é inexistente ou esta inativo!" & vbCrLf _
'             & "Consulte o cadastro do Produto."
'    End If
'End Sub

'==========================================================================
' Purpose:  Iniciar persistencia da nota
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start:
' Alterado: Ronaldo Robledo                                     18/06/2012
'           Inserido bloco de transação no checalote
' Alteração: Ronaldo Robledo                                    06/06/2013
'            Reestruturado método para atender alteração de nota
'==========================================================================
Private Sub SalvarDocumento()
On Error GoTo file
    
    If ValidaCampos Then
        BuscaConhecimento
        If CalculaNotaFiscal Then
            If Not Verificatotais Then VoltanaTela: Exit Sub
            If l_opcao = 1 Then
                
                '*** Alteração para tratamento do controle de confêrência
                'na entrada de mercadoria , se for conferência executa inclusão
                'na tabela temporaria. ***
                If bolConferencia Then SalvarConferencia: Exit Sub

                If ExisteNF Then Exit Sub
                If Confirma("Deseja Incluir a Nota Fiscal de Entrada?") = vbNo Then VoltanaTela: Exit Sub
                bolLiberaAlteracaoCustoAlteracaoNF = True
                bolAlteraEstoquenaAlteracaoNF = True
                IniciaPersistenciaDados
                
            ElseIf l_opcao = 2 Then
                If Confirma("Deseja Alterar a Nota Fiscal de Entrada?") = vbNo Then VoltanaTela: Exit Sub
                If Not ValidaAlteraNF Then Exit Sub
                If Not VerificaAlteracaoQuantidadeProdutos Then Exit Sub
                If Not ValidaLiberacaoAlteracaoCusto Then Exit Sub
                If VerificaDevolucaoCompras(True) Then Exit Sub
                IniciaPersistenciaDados
            End If
        End If
    End If
Exit Sub
file: ValidaErros Err, Me.Caption & " - Salvardocumento"
End Sub
Private Sub VoltanaTela()
On Error GoTo Err_VoltanaTela

    If frmdados.Enabled = True Then
        SSTab1.Tab = 0
        grade1.SetFocus
    End If

Exit Sub
Err_VoltanaTela: ValidaErros Err, Me.Caption & " - VoltanaTela"
End Sub
'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 18/06/2013
'==========================================================================
Private Function ValidaLancamentoFinanceiroCaixa() As Boolean
On Error GoTo Err_ValidaLancamentoFinanceiroCaixa
    
    ValidaLancamentoFinanceiroCaixa = True
    bolLiberaAlteracaoFinanceiro = True
    If l_opcao = 1 Then Exit Function
    
    bolLiberaAlteracaoFinanceiro = False
    ValidaLancamentoFinanceiroCaixa = False

    If bolExisteLancCaixa And lLancamento > 0 Then
        If Confirma("Existe lançamento caixa para esta nota - Desejar Alterar?") = vbYes Then
            If (bolExisteLancCaixa Or chk_atualiza_caixa.Value = 1) And msk_entrada <> g_data_geral Then
                Alerta "Lançamento de alteração no caixa não permitido para datas retroativas!" & Chr(13) & "Efetue o lançamento manualmente no caixa"
                bolLiberaAlteracaoFinanceiro = False
                ValidaLancamentoFinanceiroCaixa = False
                Exit Function
            Else
                bolLiberaAlteracaoFinanceiro = True
                ValidaLancamentoFinanceiroCaixa = True
            End If
        End If
    End If
    
    If bolExisteLancFinanceiro Then
        If Confirma("Existe lançamento financeiro para esta nota - Desejar Alterar?") = vbYes Then
            If ValidaRetornoFinanceiro(IIf((lbl_nota <> "CX"), "NF", "CX")) Then
                bolLiberaAlteracaoFinanceiro = True
                ValidaLancamentoFinanceiroCaixa = True
            End If
        End If
    End If
    
    
Exit Function
Err_ValidaLancamentoFinanceiroCaixa: ValidaErros Err, Me.Caption & " - ValidaLancamentoFinanceiroCaixa"
End Function
'==========================================================================
' Purpose:  Iniciar a persistencia na base de dados depois das validações
'            e atender também alteração da nota
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Sub IniciaPersistenciaDados()
On Error GoTo Err_IniciaPersistenciaDados
    
    g_string = ""
    If ValidaLancamentoFinanceiroCaixa Then
        If chk_atualiza_carteira.Value = 1 Or chk_atualiza_caixa.Value = 1 Then
            If chk_atualiza_caixa.Value = 0 Then movimento_parcelamento_entrada.frm_caixa.Enabled = False
            movimento_parcelamento_entrada.txt_reduzida = xplanoconta
            movimento_parcelamento_entrada.Show 1
        End If
    Else
        If Confirma("Deseja continuar o lançamento?") = vbNo Then Exit Sub
    End If
    
    If g_string <> "Parcelamento Cancelado" Then
        
        If bolAlteraEstoquenaAlteracaoNF Then
            If Not IniciaChecagemLote Then Exit Sub
            If Not PreencheDeposito Then Exit Sub
        End If
        
        Conexao.BeginTrans
        If chk_impressao_nf.Value = 1 Then Call BuscanumeroNotaFiscal
        
        If AtualizaTabelas Then
            '*** Autor: Diego Martins Data:30/04/2011
            'Com a reformulação no modulo NF-e onde o commit podera ser feito dentro do
            'do modulo,esta validação garante que o comeando não sera executado novamente.
            If Not g_NFEConfirmada Then Conexao.CommitTrans
            
            If chk_impressao_nf.Value = 1 And cDPEmpresa.NotaFiscalEletronica = 0 Then
                'Alerta "Prepare a impressora para impressão da N.F.!"
                If g_novo_configNF = 0 Then Imprime_NF Else Imprime_NFNovo

            ElseIf chk_impressao_nf.Value = 1 And cDPEmpresa.NotaFiscalEletronica > 0 Then
                   '*** Autor: Diego Martins Ticket: TT721 Data: 03/07/2012 - Gravar chave de acesso NFe ***
                   Call Conexao.AlterarRecordset("movimento_cabecalho_nota_fiscal_entrada", "chave_acesso_nfe = '" & txt_chave_acesso & "'", " sequencia = '" & lSequencia & "'", cDPEmpresa.codigo)
                   
                   '*** Autor : Diego Martins Data:04/04/2011
                   'Alteração :Desvincular Emissão Danfe do Processo de emissão de NF-e ***
                   Call ImprimirNFE(TipoNFe.entrada, lSequencia)
                   
            End If
        
            If Confirma("Deseja Imprimir Espelho da Entrada?") = vbYes Then
                ImpressaoEspelho
            End If
            
            ImpSeparacao
            
            flag_tela_entrada_mercadoria = 0
            Form_Activate
        Else
            Conexao.RollbackTrans
            If l_opcao <> 2 Then Call cMapLote.RetornaLote(lngSequenciaControle)
            Call Conexao.DeleteSintetico("nota_fiscal_parcelamento_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and usuario = '" & g_usuario & "'", cDPEmpresa.codigo)
            Exit Sub
        End If
    Else
        If l_opcao <> 2 Then Call cMapLote.RetornaLote(lngSequenciaControle)
        Call Conexao.DeleteSintetico("nota_fiscal_parcelamento_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and usuario = '" & g_usuario & "'", cDPEmpresa.codigo)
    End If

Exit Sub
Err_IniciaPersistenciaDados: ValidaErros Err, Me.Caption & " - IniciaPersistenciaDados"
End Sub

'*****************************************************************************
'Criação: Diego Martins dos Santos                      Data: 19/06/2010
'
'Propósito:Salvar tratamento conferência
'*****************************************************************************
Private Sub SalvarConferencia()
On Error GoTo Err_SalvarConferencia

    Conexao.BeginTrans
    If AtualizaTabelasConferencia Then
       Conexao.CommitTrans
       Alerta "Entrada efetuada com Sucesso! Aguardando Conferência!!!", 64
       LimpaTela
       flag_tela_entrada_mercadoria = 0
       Form_Activate
    Else
        Conexao.RollbackTrans
        Alerta "Erro na inclusão da Nota para Conferência!", 48
    End If

Exit Sub
Err_SalvarConferencia: ValidaErros Err, Me.Caption & " - SalvarConferencia"
End Sub

Private Sub ChamaCelula()
    
    LastRow = grade1.Row
    LastCol = grade1.col
    
    'Nova Celula
    With grade1
        If .TextMatrix(LastRow, 0) = NovaLinha Then
            .Rows = .Rows + 1
            .TextMatrix(LastRow, 0) = LastRow
            .TextMatrix(.Rows - 1, 0) = NovaLinha
            LastRow = LastRow + 1
            ZeraGrade
       End If
    End With

grade1.col = 1
grade1.Row = grade1.Row + 1
End Sub

'*****************************************************************************
'Criação: Ronaldo Robledo Mendes de Souza                      Data: 28/01/2012
'Alteração: Inserido busca do campo subcodigo
'Alteração: Ronaldo Robledo                                         21/03/2013
'           Implementado para buscar os pedidos pelas novas entidades utilizando
'           mapeador
'*****************************************************************************
Private Sub Pedidos()
On Error GoTo file

    g_string = ""
    g_string4 = ""
    
    consulta_pedidos_entrada.Show (1)
    If g_string <> "" Then
        
        Call BuscaItensPedido(Val(g_string), g_string4, 0)
        
        If IsNumeric(txt_codigo_fornecedor) Then
            Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM fornecedor WHERE codigo = '" & txt_codigo_fornecedor & "'", 1)
            If gb_Recordset.EOF = False Then
                LimpaTelaFornecedor
                AtualTelaFornecedor
                txt_fornecedor.SetFocus
            End If
            gb_Recordset.Close
        End If
    End If
    
Exit Sub
file: ValidaErros Err, Me.Caption & " - Pedidos"
End Sub
'==========================================================================
' Purpose:  Obter itens do pedido
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 23/03/2013
'==========================================================================
Private Sub BuscaItensPedido(ByVal lngNumeroPedido As Long, ByVal strCodProduto As String, ByVal lngLinha As Long)
On Error GoTo Err_BuscaItensPedido
Dim z   As Long
Dim Criador         As New AutCria.cCrdMovimento
Dim MapPedidoCompra As New cIMapPedidoCompra
Dim DPPedidoCompra  As New cDPPedidoCompra
Dim DPItemPedido    As New cDPItemMovimentacao
Dim i               As Integer

        DPPedidoCompra.NumeroPedido = lngNumeroPedido
        Set DPPedidoCompra.movimentacao.Empresa = cDPEmpresa
        Set MapPedidoCompra = Criador.CrieMapPedidoCompra
        If Not MapPedidoCompra.ObtemPedidoCompra(DPPedidoCompra) Then Alerta "Produto não Localizado ou Pedido já Baixado!": Exit Sub
        
        
        z = grade1.Rows - 1
        grade1.Rows = grade1.Rows
        txt_codigo_fornecedor = DPPedidoCompra.DPFornecedor.codigo
        txt_fornecedor = DPPedidoCompra.DPFornecedor.DPPessoa.RazaoSocial
        For i = 1 To DPPedidoCompra.movimentacao.ItemMovimentacao.Count
            Set DPItemPedido = DPPedidoCompra.movimentacao.ItemMovimentacao(i)
            If DPItemPedido.DPItemPedidoCompra.SituacaoItem = eTipoSituacaoItemPedido.Aberto Then
                'quando busca apenas por item especifico
                If (DPItemPedido.cDPProduto.codigo = strCodProduto) Or Trim(strCodProduto) = "" Then
                    If lngLinha = 0 Then
                        grade1.Rows = grade1.Rows + 1
                        grade1.Row = z
                    Else
                        z = lngLinha
                    End If
                    
                    grade1.TextMatrix(z, 0) = z
                    grade1.TextMatrix(z, 1) = DPPedidoCompra.NumeroPedido
                    grade1.TextMatrix(z, 2) = DPItemPedido.cDPProduto.codigo
                    grade1.TextMatrix(z, 3) = DPItemPedido.cDPProduto.Descricao
                    grade1.TextMatrix(z, 4) = DPItemPedido.cDPProduto.DPUnidade.Unidade
                    grade1.TextMatrix(z, 5) = Format(DPItemPedido.Quantidade - DPItemPedido.DPItemPedidoCompra.QuantidadeRecebida, "##,###,##0.00")
                    grade1.TextMatrix(z, 6) = Format(DPItemPedido.ValorBruto, g_decimal_compra)
                    grade1.TextMatrix(z, 7) = Format(DPItemPedido.ValorDesconto, "##,###,##0.00000")
                    grade1.TextMatrix(z, 8) = Format(DPItemPedido.ValorLiquido, g_decimal_compra)
                    grade1.TextMatrix(z, 9) = Format(DPItemPedido.ObtenhaValorTributo(eTipoTributo.Ipi, Percentual), "##,###,##0.00")
                    grade1.TextMatrix(z, 10) = Format(DPItemPedido.ObtenhaValorTributo(eTipoTributo.Icms, Percentual), "##,###,##0.00")
                    grade1.TextMatrix(z, 11) = Format((DPItemPedido.Quantidade - DPItemPedido.DPItemPedidoCompra.QuantidadeRecebida) * DPItemPedido.ValorLiquido, "##,###,##0.00")
                    grade1.TextMatrix(z, 21) = 0 'perc pis
                    grade1.TextMatrix(z, 22) = 0 'perc cofins
                    grade1.TextMatrix(z, 24) = DPItemPedido.cDPProduto.ControleLote
                    grade1.TextMatrix(z, 26) = DPItemPedido.cDPProduto.Fracionado
                    grade1.TextMatrix(z, 29) = DPItemPedido.cDPProduto.codigo
                    grade1.TextMatrix(z, 13) = 0
                    grade1.TextMatrix(z, 14) = 0
                    grade1.TextMatrix(z, 15) = 0
                    grade1.TextMatrix(z, 16) = DPItemPedido.cDPProduto.CodTributacao
                    grade1.TextMatrix(z, 18) = 0
                    grade1.TextMatrix(z, 19) = 0
                    grade1.TextMatrix(z, 20) = 0
                    grade1.TextMatrix(z, 23) = 0
                    grade1.TextMatrix(z, 25) = 0
                    grade1.TextMatrix(z, 27) = 0
                    grade1.TextMatrix(z, 33) = 0 'tipp importacao
                    grade1.TextMatrix(z, 34) = 0
                    grade1.TextMatrix(z, 35) = 0
                    grade1.TextMatrix(z, 36) = 0
                    grade1.TextMatrix(z, 37) = 0
                    grade1.TextMatrix(z, 38) = DPItemPedido.cDPProduto.codigoBarras
                    grade1.TextMatrix(z, 39) = DPItemPedido.cDPProduto.UtilizaGrade
                    grade1.TextMatrix(z, 40) = DPItemPedido.cDPProduto.CodigoNCM
                    grade1.TextMatrix(z, 46) = DPItemPedido.cDPProduto.Subcodigo
                    grade1.TextMatrix(z, 47) = DPItemPedido.cDPProduto.CodGrupoTributacao
                    grade1.TextMatrix(z, 48) = 0 'cst pis
                    grade1.TextMatrix(z, 49) = 0 'cst cofins
                    grade1.TextMatrix(z, 54) = DPItemPedido.Identificador
        
                    z = z + 1
                End If
            End If
        Next
        If lngLinha = 0 Then
            grade1.TextMatrix(grade1.Rows - 1, 0) = NovaLinha
            Call ChamaCelula
        Else
            grade1.TextMatrix(LastRow, 0) = NovaLinha
            grade1.col = grade1.col + 3
        End If
        
        Set DPItemPedido = Nothing
        Set DPPedidoCompra = Nothing
        Set Criador = Nothing
        Set MapPedidoCompra = Nothing


Exit Sub
Err_BuscaItensPedido: ValidaErros Err, Me.Caption & " - BuscaItensPedido"
End Sub
Private Function AtualizaCaixa() As Boolean
'Atualiza Caixa Produtos a Prazo e a Vista
On Error GoTo Err_AtualizaCaixa

    AtualizaCaixa = False

    BD_Record_Set.Source = "SELECT conta FROM plano_conta WHERE conta_reduzida = '" & gb_Recordset!conta_reduzida & "'"
    BD_Record_Set.Open
    If BD_Record_Set.RecordCount > 0 Then
        lContaContabilVista = BD_Record_Set!conta
    End If
    BD_Record_Set.Close
    
    If pct_orçamento.Visible = False Then
        g_string = "PGTO. NF. " & txt_numero_nf & " - " & txt_fornecedor.Text
    ElseIf pct_orçamento.Visible = True Then
        g_string = "PGTO. OR. " & txt_numero_nf & " - " & txt_fornecedor.Text
    End If
    
    BD_Record_Set.Source = "SELECT caixa FROM abertura_caixa WHERE empresa = '" & cDPEmpresa.codigo & "' and codigo = '" & gb_Recordset!codigo_caixa & "' and data_abertura = " & FormataData(gb_Recordset!Data_Caixa)
    BD_Record_Set.Open
    If BD_Record_Set.RecordCount > 0 Then
        g_string3 = BD_Record_Set!Caixa
    Else
        g_string3 = 1
    End If
    BD_Record_Set.Close
    
    With MovimentoCaixa
        .Empresa = cDPEmpresa.codigo
        .Data = EnviaData(gb_Recordset!Data_Caixa)
        .Caixa = g_string3
        .CodigoResponsavel = gb_Recordset!codigo_caixa
        .NumeroMovimento = lLancamento
        .NumeroConta = lContaContabilVista
        .NumeroContaReduzida = lcontareduzida
        .DigitoConta = 0
        .CreditoOuDebito = "D"
        .Valor = gb_Recordset!valor_do_vencimento
        If gb_Recordset!tipo_documento = 100 Then
            .Especie = "1"
        Else
            .Especie = "2"
        End If
        .CodigoHistorico = "4"
        .ComplementoHistorico = g_string
        .DataEmissao = "00:00:00"
        .DataVencimento = "00:00:00"
        .DataMovimentacao = EnviaData(gb_Recordset!Data_Caixa)
        .CodigoCliente = 0
        .NomeCliente = ""
        .CodigoEmitido = 0
        .NomeEmitido = ""
        .TipoDocumento = 0
        .FormaContas = 2
        .numeroduplicata = ""
        .ValorVencimento = 0
        .Saldo = 0
        .TipoLancamento = 1
        .TipoConta = "D"
        .SerieNota = lbl_serie_nf
        .CaixaFechado = "A"
        .Observacoes = gb_Recordset!Observacoes
        .RelacaoPagarReceber = 0
        If .Incluir Then
            AtualizaCaixa = True
            If lLancamento = 0 Then
                lLancamento = .NumeroIdentificador
            End If
        End If
    End With

Exit Function
Err_AtualizaCaixa: ValidaErros Err, Me.Caption & " - AtualizaCaixa"
End Function

Private Function AtualizaCaixaGeral() As Boolean
    'Atualiza Caixa Produtos a Prazo e a Vista
On Error GoTo Err_AtualizaCaixaGeral

    AtualizaCaixaGeral = False
    
    With gb_Recordset
        BD_Record_Set.Source = "SELECT * FROM plano_conta WHERE conta_reduzida = '" & !conta_reduzida & "'"
        BD_Record_Set.Open
        If BD_Record_Set.RecordCount > 0 Then
            lContaContabilVista = BD_Record_Set!conta
        End If
        BD_Record_Set.Close
        
        If pct_orçamento.Visible = False Then
            g_string = "PGTO. NF. " & txt_numero_nf & " - " & txt_fornecedor.Text
        ElseIf pct_orçamento.Visible = True Then
            g_string = "PGTO. OR. " & txt_numero_nf & " - " & txt_fornecedor.Text
        End If
    
        MovimentoCaixaGeral.CodigoEmpresa = cDPEmpresa.codigo
        MovimentoCaixaGeral.NomeEmpresa = cDPEmpresa.NOME
        MovimentoCaixaGeral.Data = EnviaData(!Data_Caixa)
        MovimentoCaixaGeral.Caixa = !codigo_caixa
        MovimentoCaixaGeral.CodigoBanco = !codigo_portador
        MovimentoCaixaGeral.NomeBanco = !nome_portador
        MovimentoCaixaGeral.NumeroMovimento = lLancamento
        MovimentoCaixaGeral.NumeroConta = lContaContabilVista
        MovimentoCaixaGeral.NumeroContaReduzida = !conta_reduzida
        MovimentoCaixaGeral.DigitoConta = 0
        MovimentoCaixaGeral.CreditoOuDebito = "D"
        MovimentoCaixaGeral.Valor = !valor_do_vencimento
        
        If !tipo_documento = 100 Then
            MovimentoCaixaGeral.Especie = 1
        Else
            MovimentoCaixaGeral.Especie = 2
        End If
        
        MovimentoCaixaGeral.CodigoHistorico = "4"
        MovimentoCaixaGeral.ComplementoHistorico = g_string
        MovimentoCaixaGeral.DataEmissao = "00:00:00"
        MovimentoCaixaGeral.DataVencimento = "00:00:00"
        MovimentoCaixaGeral.DataMovimentacao = EnviaData(!Data_Caixa)
        MovimentoCaixaGeral.CodigoCliente = 0
        MovimentoCaixaGeral.NomeCliente = ""
        MovimentoCaixaGeral.CodigoEmitido = 0
        MovimentoCaixaGeral.NomeEmitido = ""
        MovimentoCaixaGeral.TipoDocumento = 0
        MovimentoCaixaGeral.FormaContas = 2
        MovimentoCaixaGeral.numeroduplicata = ""
        MovimentoCaixaGeral.ValorVencimento = 0
        MovimentoCaixaGeral.Saldo = 0
        MovimentoCaixaGeral.TipoLancamento = 1
        MovimentoCaixaGeral.TipoConta = "D"
        MovimentoCaixaGeral.SerieNota = lbl_serie_nf
        MovimentoCaixaGeral.Observacoes = !Observacoes
        MovimentoCaixaGeral.SeqCarteira = 0
        MovimentoCaixaGeral.ContaPartida = g_codigoportadorpadrao
        MovimentoCaixaGeral.NomeContaPartida = g_nomeportadorpadrao
        MovimentoCaixaGeral.Conciliado = "N"
        MovimentoCaixaGeral.CentroCusto = !centro_custo
        
        If MovimentoCaixaGeral.Incluir Then
            AtualizaCaixaGeral = True
            If lLancamento = 0 Then
                lLancamento = MovimentoCaixaGeral.NumeroIdentificador
            End If
        End If
    End With

Exit Function
Err_AtualizaCaixaGeral: ValidaErros Err, Me.Caption & " - AtualizaCaixaGeral"
End Function

'Private Sub CalculoCustoMedio()
'
'Set gb_Recordset = Conexao.GeraRecordset("SELECT estoque.quantidade_cx,produto.custo_medio_anterior FROM estoque,produto WHERE estoque.codigo_do_produto = '" & grade1.TextMatrix(z, 2) & "' and produto.codigo = '" & grade1.TextMatrix(z, 2) & "'", 1)
'If gb_Recordset.RecordCount > 0 Then
'  If gb_Recordset!quantidade_cx > 0 Then
'      zQuantidade = gb_Recordset!quantidade_cx
'  Else
'      zQuantidade = 0
'  End If
'  zCustoAnterior = gb_Recordset!custo_medio_anterior
'End If
'gb_Recordset.Close
'
'zValorTotal = Format(((cdbl(zQuantidade) * cdbl(zCustoAnterior)) + (grade1.TextMatrix(z, 5) * lcustoprodutos(z))) / (cdbl(zQuantidade) + cdbl(grade1.TextMatrix(z, 5))), "##,###,##0.0000")
'g_string2 = lcustoprodutos(z)
'g_string3 = zValorTotal


'End Sub
'Private Sub GravaTabLista(ByVal pCodigo As String)
'
'Call Conexao.InserirRecordsetII("variacao_preco", "codigo,subcodigo,banho,tipo,modelo,codigo_do_grupo,descricao,data_alteracao,preco_varejo,preco_atacado,custo,custo_medio,preco_varejo_cont,preco_atacado_cont,custo_cont,custo_medio_cont,usuario,horario,status", "" & _
'                                "produto.codigo,produto.subcodigo,produto.banho,produto.tipo,produto.modelo,produto.codigo_grupo,produto.descricao," & FormataData(Date) & "," & fvalidavalornovo(lcustoprodutos(z)) & "," & fvalidavalornovo(lcustoprodutos(z)) & "," & fvalidavalornovo(lcustoprodutos(z)) & "," & _
'                                fvalidavalornovo(xcustomedio) & "," & fvalidavalornovo(lcustoprodutos(z)) & "," & fvalidavalornovo(lcustoprodutos(z)) & "," & fvalidavalornovo(lcustoprodutos(z)) & "," & fvalidavalornovo(x_customediocont) & ",'" & g_usuario & "','" & Time & "','EP'", "produto.codigo = '" & pCodigo & "'", "produto")
'
'End Sub

Private Sub SomaCalculoPercentual()

On Error GoTo Err_SomaCalculoPercentual

    SomaPercentual = txt_total
    
    If chk_ipi_total.Value = 1 Then SomaPercentual = CDbl(SomaPercentual) - CDbl(txt_ipi)
    
    If chk_frete_total.Value = 1 Then SomaPercentual = CDbl(SomaPercentual) - CDbl(txt_frete)
    
    If chk_seguro_total.Value = 1 Then SomaPercentual = CDbl(SomaPercentual) - CDbl(txt_seguro)
    
    If chk_outras_total.Value = 1 Then SomaPercentual = CDbl(SomaPercentual) - CDbl(txt_outras)
    
    If chk_soma_subs.Value = 1 Then SomaPercentual = CDbl(SomaPercentual) - CDbl(g_string2)
    
    If chk_desconto_total.Value = 1 Then SomaPercentual = CDbl(SomaPercentual) + CDbl(txt_desconto)

Exit Sub
Err_SomaCalculoPercentual: ValidaErros Err, Me.Caption & " - SomaCalculoPercentual"
End Sub

Private Function SalvarPrecoManual(ByVal z As Long) As Boolean
On Error GoTo Err_SalvarPrecoManual
    
    SalvarPrecoManual = False
    Set gb_Recordset = Conexao.GeraRecordset("SELECT preco_varejo,preco_atacado,marckup_varejo,marckup_atacado FROM calculo_entrada_tmp WHERE empresa = '" & cDPEmpresa.codigo & "' and codigo_usuario = '" & g_usuario & "' and numero_nf = '" & txt_numero_nf & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & "' and data_emissao = " & FormataData(msk_entrada) & " and codigo_produto = '" & grade1.TextMatrix(z, 2) & "' and outros = '" & lbl_nota & "'", 1)
    If gb_Recordset.RecordCount > 0 Then
        Call Conexao.AlterarRecordset("produto", "preco_varejo = " & fValidaValor2(gb_Recordset!preco_varejo) & ",preco_atacado = " & fValidaValor2(gb_Recordset!preco_atacado) & ",marckup_varejo = " & fValidaValor2(gb_Recordset!marckup_varejo) & ",marckup_atacado = " & fValidaValor2(gb_Recordset!marckup_atacado) & ",data_atacado = " & FormataData(Date) & ",data_varejo = " & FormataData(Date), "codigo = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
        Call Conexao.InserirRecordset("log_senhas", "data,hora,codigo_usuario,nome_usuario,historico,tela,observacoes,outros", FormataData(Date) & ",'" & Time & "','" & g_usuario & "','" & g_nome_usuario & "','Alteração Preço Manual Nota " & txt_numero_nf & " Pd." & grade1.TextMatrix(z, 2) & "','Entradas de Mercadoria','Tela Principal','NF'", cDPEmpresa.codigo)
    End If
    gb_Recordset.Close
    SalvarPrecoManual = True

Exit Function
Err_SalvarPrecoManual: ValidaErros Err, Me.Caption & " - SalvarPrecoManual"
End Function

Private Sub LancaCalculo()
On Error GoTo Err_LancaCalculo
Dim g As Integer
    For g = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(g, 0)) Then
            'se tiver porcentagem do icms calcular
            If CLng(grade1.TextMatrix(g, 10)) > 0 Then
                    If grade1.TextMatrix(g, 12) <> "S" Then
                        lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + xbasecalculo
                        lValorIcms(4) = lValorIcms(4) + ((xbasecalculo * CDbl(grade1.TextMatrix(g, 10))) / 100)
                        xValorIcms = (xbasecalculo * CDbl(grade1.TextMatrix(g, 10))) / 100
                        xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
                        grade1.TextMatrix(g, 14) = CDbl(grade1.TextMatrix(g, 14)) + xbasecalculo
                        grade1.TextMatrix(g, 15) = CDbl(grade1.TextMatrix(g, 15)) + xValorIcms
                        xSomaRateioIsencao = True
                        Exit Sub
                    End If
            Else
                xbasecalculo = 0
            End If
        End If
    Next
    
    For g = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(g, 0)) Then
            If CLng(grade1.TextMatrix(g, 10)) > 0 Then
                If grade1.TextMatrix(g, 12) = "S" Then
                    lBaseCalculoIcms_red = xbasecalculo
                    lValorIcms_red = 0
                    If g_reducao_invertido = 0 Then
                        lBaseCalculoIcms_red = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
                        lValorIcms_red = lBaseCalculoIcms_red * CDbl(grade1.TextMatrix(g, 10)) / 100
                    Else
                        x_calculo = (lporcentagemreducao * lBaseCalculoIcms_red) / 100
                        lBaseCalculoIcms_red = lBaseCalculoIcms_red - x_calculo
                        lValorIcms_red = lBaseCalculoIcms_red * CDbl(grade1.TextMatrix(g, 10)) / 100
                    End If
        
                    lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + lBaseCalculoIcms_red
                    lValorIcms(4) = lValorIcms(4) + lValorIcms_red
        
                    xbasecalculo = lBaseCalculoIcms_red
                    xValorIcms = lValorIcms_red
        
                    xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
                End If
                grade1.TextMatrix(g, 14) = CDbl(grade1.TextMatrix(g, 14)) + xbasecalculo
                grade1.TextMatrix(g, 15) = CDbl(grade1.TextMatrix(g, 15)) + xValorIcms
                xSomaRateioIsencao = True
                Exit For
            End If
        End If
    Next

Exit Sub
Err_LancaCalculo: ValidaErros Err, Me.Caption & " - LancaCalculo"
End Sub

Private Sub msk_chegada_GotFocus()
    msk_chegada.BackColor = 12648447
    msk_chegada.SelStart = 0
    msk_chegada.SelLength = 10
End Sub

Private Sub msk_chegada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_desconto.SetFocus
End Sub

Private Sub msk_chegada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkImportacao.Visible = True Then
            chkImportacao.SetFocus
        Else
            txt_observacoes.SetFocus
        End If
    End If
End Sub

Private Sub msk_chegada_LostFocus()
    msk_chegada = MascaraData(msk_chegada)
    msk_chegada.BackColor = &H8000000E
End Sub

'==========================================================================
' Purpose:Calcular a diferença entre base de calculo e total dos produtos que será utilizado
'         pela contabilidade nos relatorios fiscais na situação outras despesas
'        diferente de outras despesas que vem na nota este se encontra na tabela MOV.CAB.N.ENTRADA
'        e caso se precise saber o seu valor por produto terá que se efetuado o rateio manualmente
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 19/05/2009
'==========================================================================
Private Sub CalculoOutras()
On Error GoTo Err_CalculoOutras

    xtotalnotaprodutos = lSubTotal
    If chk_ipi_total.Value = 1 Then xtotalnotaprodutos = CDbl(xtotalnotaprodutos) + CDbl(lValorIPI)
    
    If chk_frete_total.Value = 1 Then xtotalnotaprodutos = CDbl(xtotalnotaprodutos) + CDbl(lFrete(lcodigoaliquota))
    
    If chk_seguro_total.Value = 1 Then xtotalnotaprodutos = CDbl(xtotalnotaprodutos) + CDbl(lSeguro(lcodigoaliquota))
    
    If chk_outras_total.Value = 1 Then xtotalnotaprodutos = CDbl(xtotalnotaprodutos) + CDbl(lOutras(lcodigoaliquota))
    
    If chk_desconto_total.Value = 1 Then xtotalnotaprodutos = CDbl(xtotalnotaprodutos) - CDbl(lDesconto(lcodigoaliquota))
    
    If chk_soma_subs.Value = 1 Then xtotalnotaprodutos = CDbl(xtotalnotaprodutos) + xValorSubstituicao

    'retornado o código antigo pois foi alterado indevidamente  ticket 642
    grade1.TextMatrix(f, 23) = Abs(xtotalnotaprodutos - xbasecalculo)

'===================================================================
    'If cdbl(grade1.TextMatrix(f, 10)) > 0 Then
        '*** Autor: Diego Martins OS: 17648 Data: 04/11/2011
        'grade1.TextMatrix(f, 23) = Abs(xtotalnotaprodutos - xbasecalculo)
        'g_string4 = cdbl(g_string4) + cdbl(grade1.TextMatrix(f, 23))
    'Else
    '    If xSomaRateioIsencao = False Then
    '        grade1.TextMatrix(f, 23) = Abs(xtotalnotaprodutos - xbasecalculo)
        
        '*** Autor: Diego Martins OS: 17648 Data: 04/11/2011
    '     grade1.TextMatrix(f, 23) = Abs(Format(lSubTotal - (lValorIPI + lSubTotal + lOutras(lcodigoaliquota) + lSeguro(lcodigoaliquota) + lFrete(lcodigoaliquota) - lDesconto(lcodigoaliquota)), "##,###,##0.0000"))
    '    Else
    '        grade1.TextMatrix(f, 23) = Abs(xtotalnotaprodutos - xbasecalculo)
    '    End If
    'End If

Exit Sub
Err_CalculoOutras: ValidaErros Err, Me.Caption & " - CalculoOutras"
End Sub

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start:
' Alteração: Ronaldo Robledo                                    26/04/2013
'            Inserido bolPegouSequencia para buscar a sequencia uma unica vez
'            e pegar apenas quando produto for controlado por lote
' Alteração: Ronaldo Robledo                                    07/06/2013
'           Inserido cálculo da quantidade secundária
'           Inserido validação quando se tratar de alteeracao de nota
'==========================================================================
Private Function ChecaLote(ByRef f As Integer) As Boolean
On Error GoTo Err_ChecaLote
Dim z                   As Long
'Dim bolPegouSequencia As Boolean

    ChecaLote = False
    
    'If l_opcao = 2 Then f = 1: ChecaLote = True: Exit Function
    
    If pct_icms.Visible = False Then
        f = 0
        For z = 0 To grade1.Rows - 1
            If IsNumeric(grade1.TextMatrix(z, 0)) Then
                'veerifica se o produto e controlado por lote
                If CLng(grade1.TextMatrix(z, 24)) > 0 Then
                    'se for inclusao não precisa checar se ja lancou
                    If l_opcao = 1 Then
                        If Not LancaLoteManual(z) Then Exit Function
                    Else    'se for alteraçao verifica se ja existe lançamento daquele lote
                        If Not cMapLote.VerificaLoteParaNotaEntrada(grade1.TextMatrix(z, 2), lSequencia) Then
                            If Not LancaLoteManual(z) Then Exit Function
                        End If
                    End If
                    f = f + 1
                    
                End If
            End If
        Next
    End If
    
    ChecaLote = True
    
Exit Function
Err_ChecaLote: ValidaErros Err, Me.Caption & " - ChecaLote"
End Function
'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 21/06/2013
'==========================================================================
Private Function LancaLoteManual(ByVal z As Long) As Boolean
On Error GoTo Err_LancaLoteManual
Dim curFatorConversao   As Currency
    
    LancaLoteManual = False
    'verifica fator de conversao
    If IsNumeric(grade1.TextMatrix(z, 45)) Then curFatorConversao = grade1.TextMatrix(z, 45) Else curFatorConversao = 1
    
    'If Not bolPegouSequencia Then lngSequenciaControle = cMapLote.SequenciaControleLote: bolPegouSequencia = True
    If lngSequenciaControle = 0 Then lngSequenciaControle = cMapLote.SequenciaControleLote
    
    Call Conexao.DeleteSintetico("movimento_lote", "empresa = '" & cDPEmpresa.codigo & "' and codigo_fornecedor = '" & txt_codigo_fornecedor & _
                                "' and numero_pedido = '" & txt_numero_nf & "' and data_emissao = " & FormataData(msk_emissao) & _
                                " and status = 'E' and codigo_produto = '" & grade1.TextMatrix(z, 2) & "'", cDPEmpresa.codigo)
    
    If Not Conexao.InserirRecordset("movimento_lote", "empresa, codigo_fornecedor, codigo_cliente, numero_pedido, data_emissao, sequencia, " & _
                                 "codigo_produto, numero_lote, data_vencimento, qtde_entrada, qtde_saida, qtde_atual, codigo_usuario, data, " & _
                                 "hora, status, qtde_da_nota,lote_seq, seq_controle", _
                                 "'" & cDPEmpresa.codigo & "','" & CLng(txt_codigo_fornecedor) & "','0','" & txt_numero_nf & "'," & FormataData(msk_emissao) & _
                                 ",'1','" & grade1.TextMatrix(z, 2) & "',''," & FormataData(0) & ",0,0,0,'" & g_usuario & _
                                 "'," & FormataData(Date) & ",'" & Time & "','E'," & fValidaValor2(grade1.TextMatrix(z, 5) * curFatorConversao) & _
                                 ",'" & IIf((l_opcao = 2), cMapLote.SequenciaMovimentoLote, 0) & "','" & lngSequenciaControle & "'", cDPEmpresa.codigo) Then Exit Function

    LancaLoteManual = True
    
Exit Function
Err_LancaLoteManual: ValidaErros Err, Me.Caption & " - LancaLoteManual"
End Function

'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 18/06/2012
'==========================================================================
Private Function ChecaLoteManual(f As Integer) As Boolean
On Error GoTo Err_ChecaLoteManual
            
    ChecaLoteManual = True
    If f > 0 Then
        If l_opcao = 2 Then frm_confirma_lote.bolAlteraNotaEntrada = True
        frm_confirma_lote.bytOpcao = 1
        frm_confirma_lote.lbl_pedido = txt_numero_nf
        frm_confirma_lote.txt_codigo_fornecedor = txt_codigo_fornecedor
        frm_confirma_lote.txt_nome_fornecedor = txt_fornecedor
        frm_confirma_lote.lngSequencia = 1
        frm_confirma_lote.lbl_data = msk_emissao
        frm_confirma_lote.lbl_tipo = "E"
        frm_confirma_lote.txt_codigocliente = "0"
        frm_confirma_lote.lngSeqControle = lngSequenciaControle
        frm_confirma_lote.Show 1
        If g_string <> "OK" Then ChecaLoteManual = False
    End If
        
Exit Function
Err_ChecaLoteManual: ValidaErros Err, Me.Caption & " - ChecaLoteManual"
End Function

'*****************************************************************************
'Criação:                                                  Data: 19/08/2011
'
'Propósito:
'
'Alteração: Paulo Senhorini                                Data: 19/08/2011
'           Adicionei chamada a função preenchecolunacst.
'*****************************************************************************
Private Sub CalculoNota()
On Error GoTo Err_CalculoNota
'Dim i As Integer
    
    Call SomaCalculoPercentual
    SomaPercentual = SomaPercentual - x_naotributados
    
    xbasecalculo = 0
    xValorIcms = 0
    xbcsubstituicao = 0
    xValorSubstituicaocusto = 0
    xValorSubstituicao = 0
    
    'cst isento
    'Call PreencheColunaCst("40") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "40"
    
    If CDbl(grade1.TextMatrix(f, 10)) > 0 Then
        xbasecalculo = CDbl(txt_bc_icms) / SomaPercentual * CDbl(lSubTotal)
        xValorIcms = CDbl(txt_icms) / SomaPercentual * CDbl(lSubTotal)
        
        lBaseCalculoIcms(4) = lBaseCalculoIcms(4) + xbasecalculo
        lValorIcms(4) = lValorIcms(4) + xValorIcms
        'CST
        'Call PreencheColunaCst("00") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "00"
    End If
    
    If lcodigoaliquota = 2 Then
        If SomaPercentual > 0 Or CDbl(txt_bc_substituicao) > 0 Then
            xbcsubstituicao = CDbl(txt_bc_substituicao) / SomaPercentual * CDbl(lSubTotal)
            lBaseCalculoSubstituicao = lBaseCalculoSubstituicao + xbcsubstituicao
            xValorSubstituicao = CDbl(txt_substituicao) / SomaPercentual * CDbl(lSubTotal)
            xValorSubstituicaocusto = xValorSubstituicao
            lValorSubstituicao = lValorSubstituicao + xValorSubstituicao
            xvaloricms_conhecimento = (lFreteBCConhecimento(lcodigoaliquota) * x_aliquota_conhecimento) / 100
        End If
        If chk_soma_subs.Value = 1 Then
            lTotal = lTotal + xValorSubstituicao
        End If
        
'        If xValorSubstituicao > 0 Then
'            'CST SUBSTITUICAO
'            Call PreencheColunaCst("10") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "10"
'        Else
'            'CST SUBSTITUICAO
'            Call PreencheColunaCst("60") 'grade1.TextMatrix(f, 30) = grade1.TextMatrix(f, 33) & "60"
'        End If
    End If

Exit Sub
Err_CalculoNota: ValidaErros Err, Me.Caption & " - CalculoNota"
End Sub

Private Sub Calculo(ByVal z As Long)
On Error GoTo Err_Calculo
'se o produto não e fracionado

    If grade1.TextMatrix(z, 26) <= 0 Then
        grade1.TextMatrix(z, 5) = Format(Arredonda(grade1.TextMatrix(z, 5)), "##,###,##0.00")
        g_string = 0
        grade1.TextMatrix(z, 7) = Format(grade1.TextMatrix(z, 7), "##,###,##0.00")
        If CDbl(grade1.TextMatrix(z, 7)) > 0 Then
            g_string = Format(CDbl(grade1.TextMatrix(z, 6)) * CDbl(grade1.TextMatrix(z, 7)) / 100, g_decimal_compra)
            grade1.TextMatrix(z, 8) = Format(CDbl(grade1.TextMatrix(z, 6)) - CDbl(g_string), g_decimal_compra)
        Else
            grade1.TextMatrix(z, 8) = Format(grade1.TextMatrix(z, 6), g_decimal_compra)
        End If
        grade1.TextMatrix(z, 11) = Format(grade1.TextMatrix(z, 5) * grade1.TextMatrix(z, 8), "##,###,##0.00")
    End If

Exit Sub
Err_Calculo: ValidaErros Err, Me.Caption & " - Calculo"
End Sub

Private Sub txt_codigo_forma_GotFocus()
    txt_codigo_forma.SelStart = 0
    txt_codigo_forma.SelLength = Len(txt_codigo_forma)
    txt_codigo_forma.BackColor = 12648447
End Sub

Private Sub txt_codigo_forma_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then msk_emissao.SetFocus
End Sub

Private Sub txt_codigo_forma_LostFocus()
    'Call BuscaCFOP
    txt_codigo_forma.BackColor = &H80000005
End Sub

Private Sub txt_codigo_forma_KeyPress(KeyAscii As Integer)
On Error GoTo file

    If KeyAscii = 13 Then
        Dim cCrdDisplay     As New AutCria.cCrdDisplayTelas
        Dim cDPFormaFat     As New cDPFormaFaturamento
        cDPFormaFat.codigo = Val(txt_codigo_forma)
        cDPFormaFat.EntradaSaida = "E"
        If cCrdDisplay.CrieDisplayTelasCadastro.ConsultaFormaFat(False, cDPFormaFat) Then
            cDPFFaturamento.ChaveColecao = Val(cDPFormaFat.codigo)
            CarregaCamposFaturamento
            txt_total.SetFocus
        End If
        Set cCrdDisplay = Nothing
        Set cDPFormaFat = Nothing
    End If
    Call ValidaInteiro(KeyAscii)

Exit Sub
file: ValidaErros Err, Me.Caption & " - CodigoForma"
End Sub


Private Function AtualizaContra(ByVal z As Long) As Boolean
On Error GoTo file

    AtualizaContra = False
    If Not Conexao.InserirRecordset("movimento_transito_compra", "empresa,codigo_forma,data_de_emissao,numero_da_nota,serie_da_nota,codigo_do_produto,unidade,nome_do_produto,quantidade,valor_unitario,valor_total,codigo_do_fornecedor,Outros,data_liquidada,sequencia_nf", _
                                    cDPEmpresa.codigo & ",'" & txt_codigo_forma & "'," & FormataData(msk_entrada) & ",'" & txt_numero_nf & "','" & txt_serie_nf & "','" & grade1.TextMatrix(z, 2) & "','" & grade1.TextMatrix(z, 4) & "','" & grade1.TextMatrix(z, 3) & "'," & fValidaValor2(grade1.TextMatrix(z, 5)) & "," & fValidaValor2(grade1.TextMatrix(z, 8)) & _
                                    "," & fValidaValor2(grade1.TextMatrix(z, 11)) & ",'" & txt_codigo_fornecedor & "','" & lbl_nota & "'," & FormataData(0) & ",'" & lSequencia & "'", cDPEmpresa.codigo) = True Then
                                    Alerta "Erro na atualização movimento em trânsito!"
        Exit Function
    End If
    AtualizaContra = True

Exit Function
file: ValidaErros Err, Me.Caption & " - AtualizaConta"
End Function

Private Function Cancelamento() As Boolean
On Error GoTo file
Dim MovimentoCabecalhoSaida As New SabreRG.cMovimentoCabecalhoSaida
Dim lSequenciaSaida         As Long
Dim blnRetornaPedido        As Boolean
Dim z                       As Long

    Cancelamento = False

    'verifica se existe pedido de mercadoria na grade
    blnRetornaPedido = False
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 0)) And Trim(grade1.TextMatrix(z, 1)) > 0 Then
            If Confirma("Deseja reabrir pedido de mercadoria desta nota?") = vbYes Then
                blnRetornaPedido = True
                Exit For
            End If
        End If
    Next


    If bolNFePendente = False Then
        Set gb_Recordset = Conexao.GeraRecordset(" SELECT MN.id_nfe, MN.num_nfe, MN.num_protocolo, MN.num_lote, MCNE.*, MNE.*, FO.codigo_cidade, FO.codigo_grupo,TIM.pk_item_movimentacao " & _
                                                 " FROM movimento_cabecalho_nota_fiscal_entrada MCNE  " & _
                                                 " INNER JOIN movimento_nota_fiscal_entrada MNE ON MNE.sequencia = MCNE.sequencia " & _
                                                 " INNER JOIN fornecedor FO ON FO.codigo = MCNE.codigo_do_fornecedor " & _
                                                 " LEFT JOIN Movimento_nfe MN ON MN.id_nfe = MCNE.id_nfe AND MN.empresa = MCNE.empresa " & _
                                                 " LEFT JOIN tb_pedido_compra TPC ON TPC.fk_empresa = MNE.empresa and TPC.numeropedido = MNE.codigo_do_grupo " & _
                                                 " LEFT JOIN tb_item_movimentacao TIM ON TIM.fk_empresa = TPC.fk_empresa and TIM.fk_movimentacao = TPC.fk_movimentacao " & _
                                                 " and TIM.fk_produto = MNE.codigo_do_produto " & _
                                                 " WHERE MCNE.sequencia = " & lSequencia & " ", 0)
        If gb_Recordset.RecordCount > 0 Then
            strNumNFe = IIf(IsNull(gb_Recordset!num_nfe), 0, gb_Recordset!num_nfe)
            strNumProtocolo = IIf(IsNull(gb_Recordset!num_protocolo), 0, gb_Recordset!num_protocolo)
            strIDLote = IIf(IsNull(gb_Recordset!num_lote), 0, gb_Recordset!num_lote)
            
            With gb_Recordset
                lSequenciaSaida = MovimentoCabecalhoSaida.ProximaSequencia
                If lSequenciaSaida = 0 Then Exit Function
                        
                If Not Conexao.InserirRecordsetIII(" movimento_cabecalho_nota_fiscal_cancelada ", _
                                                   " empresa, data_de_emissao, numero_nf, serie_nf, numero_do_cupom,serie_do_cupom," & _
                                                   " codificacao_fiscal, tipo_cliente, venda, tipo_do_documento, forma_pagamento," & _
                                                   " codigo_do_cliente, nome_do_cliente, pessoa, cgc, inscricao_estadual, endereco," & _
                                                   " bairro, cidade, uf, cep, telefone, codigo_do_vendedor, nome_do_vendedor," & _
                                                   " valor_total_produtos, valor_total_servico, base_calculo_do_icms, valor_do_icms," & _
                                                   " base_calculo_substituicao_icms, valor_icms_substituicao, bc_iss, valor_do_iss, " & _
                                                   " valor_ipi, tipo_frete, frete, outras_despesas, valor_do_desconto_produtos," & _
                                                   " valor_do_desconto_servico, arredondamento, total_da_nota, total_custo, total_custo_cont," & _
                                                   " contribuinte, numero_movimento_caixa_venda, numero_movimento_caixa_servico, codigo_transportadora," & _
                                                   " placa, uf_placa, quantidade, especie, peso_bruto, peso_liquido, liberacao_venda, outros, numero_pedido," & _
                                                   " desconto_gerente, observacoes, codigo_cidade, codigo_regiao, codigo_grupo_cliente, codigo_rota, " & _
                                                   " codigo_conveniado, tmk, total_pis, total_cofins, custo_medio, custo_medio_cont, entrega, data_cancelamento, " & _
                                                   " desconto_pe, numero_orcamento, transito, prazo_medio, saida_dev, caixa, sequencia, id_nfe, modelo_nf", _
                                                   !Empresa & "," & FormataData(!data_de_emissao) & ",'" & !Numero & "','" & !Serie & "','0','0'," & CLng(!codificacao_fiscal) & ",'2','A'," & _
                                                   "'1','" & CLng(txt_codigo_forma) & "','" & CLng(txt_codigo_fornecedor) & "','" & txt_fornecedor & "','" & lbl_pessoa & "','" & lbl_cgc & "','" & lbl_inscricao & "','" & _
                                                   lbl_endereco & "','" & lbl_bairro & "','" & lbl_cidade & "','" & lbl_uf & "','0','0'," & _
                                                   CLng(g_usuario) & ",'" & g_nome_usuario & "'," & fValidaValor2(lbl_total_produtos) & "," & fValidaValor2(0) & "," & fValidaValor2(lbl_bc_icms) & "," & fValidaValor2(lbl_valor_icms) & "," & fValidaValor2(lbl_bc_substituicao) & "," & _
                                                   fValidaValor2(lbl_icms_substituicao) & "," & fValidaValor2(0) & "," & fValidaValor2(0) & "," & fValidaValor2(lbl_ipi) & ",'1'," & fValidaValor2(lbl_frete) & "," & fValidaValor2(lbl_outras_despesas) & "," & _
                                                   fValidaValor2(txt_desconto) & "," & fValidaValor2(0) & ",0," & fValidaValor2(!total_da_nota) & "," & fValidaValor2(!total_custo) & "," & fValidaValor2(!total_custo) & ",'" & lContribuinte & "',0,0,'" & CLng(!codigo_transportadora) & _
                                                   "','" & !Placa & "','" & !uf_placa & "'," & fValidaValor2(0) & ",'','0','0','0','" & lbl_nota & "','0','0','" & txt_observacoes & "','" & _
                                                   CLng(!codigo_cidade) & "','" & CLng(1) & "','" & CLng(!codigo_grupo) & "','" & CLng(1) & "','" & CLng(0) & "','" & CLng(0) & "'," & fValidaValor2(0) & "," & fValidaValor2(0) & "," & fValidaValor2(0) & "," & fValidaValor2(0) & _
                                                   ",'N'," & FormataData(Date) & "," & fValidaValor2(0) & ",'" & CLng(0) & "','0','0','E','1','" & lSequenciaSaida & "', '" & IIf(IsNull(!id_nfe), 0, !id_nfe) & "', '" & !modelo_nf & "'", cDPEmpresa.codigo) Then
                                                
                                                    Alerta "Erro na atualização ou venda ja emitida!"
                                                    Exit Function
                End If
            
                Do Until gb_Recordset.EOF
                    If Not Conexao.InserirRecordset(" movimento_nota_fiscal_cancelada", "empresa,data_de_emissao,numero_da_nota,serie_da_nota,numero_do_cupom,serie_do_cupom,codigo_do_grupo,codigo_do_produto,servico_produto,unidade,nome_do_produto,quantidade,valor_unitario,codigo_da_aliquota,valor_total,porcentagem_do_icms,valor_do_icms,desconto_produto,codigo_do_vendedor,codigo_do_mecanico,codigo_do_cliente,outros,base_calculo_icms,codificacao_fiscal,preco_custo,preco_custo_cont,preco_ipi,porcentagem_ipi,numero_nota_referente,influencia_estoque,codigo_fiscal_nacional,codigo_aliquota_cf,porcentagem_icms_cf,preco_pis,preco_cofins,preco_custo_medio,preco_custo_medio_cont,locacao,desconto_pe,codigo_situacao_tributaria,base_calc_substituicao,valor_icms_substituicao,caixa,fator_conversao,sequencia,preco_promocao", _
                                                    !Empresa & "," & FormataData(!data_de_emissao) & "," & !Numero & ",'" & !Serie & "','0','CF','" & !codigo_do_grupo & "','" & !codigo_do_produto & "','P','" & !Unidade & "','" & !nome_do_produto & "'," & fValidaValor2(!Quantidade) & "," & fValidaValor2(!Valor_Unitario) & "," & CInt(!codigo_da_aliquota) & "," & fValidaValor2(!valor_total) & "," & fValidaValor2(!perc_icms) & "," & fValidaValor2(!valor_do_icms) & "," & fValidaValor2(!valor_bruto - !Valor_Unitario) & "," & _
                                                    "'0','0','" & !codigo_do_fornecedor & "','" & lbl_nota & "'," & fValidaValor2(!base_calculo_icms) & ",'" & !codificacao_fiscal & "'," & fValidaValor2(!preco_custo) & "," & fValidaValor2(!preco_custo) & "," & fValidaValor2(!Valor_do_IPI) & "," & fValidaValor2(!porcentagem_do_ipi) & ",'0','" & !influencia_estoque & "','','" & CInt(!codigo_da_aliquota) & "',0,0,0," & fValidaValor2(!preco_custo_medio) & _
                                                    "," & fValidaValor2(!preco_custo_medio_cont) & ",'',0,'0'," & fValidaValor2(!base_calculo_subst) & "," & fValidaValor2(!icms_substituicao) & ",'1',1,'" & lSequenciaSaida & "','0'", cDPEmpresa.codigo) = True Then
                        Exit Function
                    End If
                    
                    If Not IsNull(!pk_item_movimentacao) Then
                        If Val(!codigo_do_grupo) > 0 And blnRetornaPedido = True And Val(!pk_item_movimentacao) > 0 Then
                            If Not AbateQuantidadeItemPedidoCompra(!Quantidade, !pk_item_movimentacao, !codigo_do_grupo) Then Exit Function
                            'Call Conexao.AlterarRecordset("movimento_pedido_mercadoria", "quantidade_recebida = quantidade_recebida - " & fValidaValor2(!Quantidade) & ",status = 'A'", "empresa = '" & cDPEmpresa.Codigo & "' and codigo_do_produto = '" & !codigo_do_produto & "' and numero_pedido = '" & !codigo_do_grupo & "'", cDPEmpresa.Codigo)
                        End If
                    End If
        
                    gb_Recordset.MoveNext
                Loop
            End With
            gb_Recordset.Close
        Else
            Alerta "Nota Não Localizada para Cancelamento!"
            gb_Recordset.Close
            Exit Function
        End If
    End If

    If Not ExclusaoNota(False) Then Exit Function
    
    If cDPEmpresa.NotaFiscalEletronica = 1 Then
        If bolNFePendente Then
            If Not ReaproveitaNFePendente(txt_numero_nf, msk_emissao) Then Exit Function
        
        ElseIf strNumProtocolo <> "" And strNumNFe <> "" And strMotivoCancelamento <> "" Then
            If Not CancelaNFE(strIDLote, strNumNFe, strNumProtocolo, strMotivoCancelamento) Then Exit Function
        
        ElseIf strMotivoCancelamento = "" Then
            Alerta "Informe o motivo do cancelamento!"
            Exit Function
        End If
    End If
    
    Cancelamento = True


Exit Function
file: ValidaErros Err, Me.Caption & " - Cancelamento"
End Function

Private Function buscasenhaCanc()
On Error GoTo Err_buscasenhaCanc

    buscasenhaCanc = False
    Senha.lbl_titulo = "Senha Para Cancelamento de Entrada"
    Senha.lbl_liberacao = "Liberação (Cancelamento)"
    Senha.lbl_tela = "Entrada de NF."
    Senha.lbl_serie = txt_serie_nf
    Senha.lbl_historico = "Liberação Cancelamento Nota " & txt_numero_nf
    Senha.Show (1)
    If g_string = "OK" Then
        buscasenhaCanc = True
    End If

Exit Function
Err_buscasenhaCanc: ValidaErros Err, Me.Caption & " - buscasenhaCanc"
End Function

Private Function PreencheDeposito() As Boolean
On Error GoTo Err_PreencheDeposito

    PreencheDeposito = False
    'se alterou quantidade e do tipo vendas ou reserva de estoque então
    If cDPFFaturamento.InfluenciaEstoque = 1 And g_utiliza_estoque_deposito = 1 And pct_icms.Visible = False Then
        g_string = xid_deposito
        g_string2 = 0
        g_string3 = "E"
        g_string4 = False
        frm_grade_deposito.Show 1
        PreencheDeposito = g_string4
        xid_deposito = g_string
    Else
        PreencheDeposito = True
    End If

Exit Function
Err_PreencheDeposito: ValidaErros Err, Me.Caption & " - PreencheDeposito"
End Function

Private Sub ImpSeparacao()
On Error GoTo file

    If cDPFFaturamento.InfluenciaEstoque = 1 And g_utiliza_estoque_deposito = 1 And pct_icms.Visible = False Then
            Call ImpDetSeparacao("Entrada", Date, lSequencia, "", "", txt_codigo_fornecedor, txt_fornecedor, "", lbl_endereco, lbl_bairro, lbl_cidade, lbl_uf, lbl_telefone, lbl_cep, "", "", "", "", "E", txt_numero_nf)
    End If

Exit Sub
file: ValidaErros Err, Me.Caption & " - ImpSeparacao"
End Sub

'Private Sub LimpaGradeCfop()
'    For z = 0 To grade1.Rows - 1
'        If IsNumeric(grade1.TextMatrix(z, 0)) Then
'            grade1.TextMatrix(z, 28) = ""
'            'grade1.TextMatrix(z, 41) = "000"
'        End If
'    Next
'End Sub

Private Sub txt_volume_GotFocus()
    txt_volume.BackColor = 12648447
    txt_volume.SelStart = 0
    txt_volume.SelLength = Len(txt_volume)
End Sub

Private Sub txt_volume_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cbo_frete.SetFocus
End Sub

Private Sub txt_volume_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_especie.SetFocus
End Sub

Private Sub txt_volume_LostFocus()
    txt_volume.BackColor = &H8000000E
    txt_volume = UCase(Trim(txt_volume))
End Sub

Private Sub txt_codigo_produto_GotFocus()
    txt_codigo_produto.SelStart = 0
    txt_codigo_produto.SelLength = Len(txt_codigo_produto)
    txt_codigo_produto.BackColor = &HFFFF&
End Sub

Private Sub txt_codigo_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
    End If
End Sub

Private Sub txt_codigo_produto_KeyPress(KeyAscii As Integer)
On Error GoTo Err_txt_codigo_produto_KeyPress

    If KeyAscii = 13 Then
        If Trim(txt_codigo_produto) = "" Then
            CONSULTA_PRODUTO.Show 1
            If Len(g_string) > 0 Then
                Call BuscaProduto(g_string) 'Criado esta função OS: 11904
            End If
        Else
            Call BuscaProduto(txt_codigo_produto) 'Criado esta função OS: 11904
        End If
    End If
    Call ValidaInteiro(KeyAscii)

Exit Sub
Err_txt_codigo_produto_KeyPress: ValidaErros Err, Me.Caption & " - txt_codigo_produto_KeyPress"
End Sub

Private Sub txt_codigo_produto_LostFocus()
    txt_codigo_produto.BackColor = &H80000005
End Sub

'*****************************************************************************
'Criação: Ronaldo Robledo                                     Data: 04/06/2010
'Propósito:busca produtos por codigo de barras ou codigo do produto
'*****************************************************************************
Private Sub BuscaProduto(strCodProduto As String) 'Alteração OS:11904
On Error GoTo Err_BuscaProduto

    Sql_Query = " SELECT produto_fornecedor, codigo, descricao, " & _
                " unidade, marca, locacao_1, embalagem " & _
                " FROM produto PR " & _
                " WHERE (PR.codigo = (" & _
                " SELECT codigo " & _
                " FROM codigo_barras " & _
                " WHERE codigo_barras = '" & txt_codigo_produto & "' LIMIT 1)) " & _
                " OR (PR.codigo = '" & txt_codigo_produto & "')"
                
    BD_Record_Set.Source = Sql_Query
    BD_Record_Set.Open
    If BD_Record_Set.EOF = False Then
        txt_codigo_produto = BD_Record_Set!codigo
        txt_descricao = BD_Record_Set!Descricao
        cbo_unidade.Clear
        cbo_unidade.AddItem (BD_Record_Set!Unidade)
        strLocacao = BD_Record_Set!locacao_1
        Call PreencheComboBox
        If chk_somaautomatica.Value = 0 Then
            cbo_unidade.SetFocus
        Else
            cmd_inserir_Click
        End If
    Else
        Alerta "Produto não localizado", 64
        LimpaConferencia
    End If
    BD_Record_Set.Close

Exit Sub
Err_BuscaProduto: ValidaErros Err, Me.Caption & " - BuscaProduto"
End Sub

Private Sub cmd_inserir_Click()
On Error GoTo Err_cmd_inserir_Click

    If Trim(txt_descricao) <> "" Then
    
        If Not ExisteProdutoConf Then
           
            grade3.TextMatrix(grade3.Rows - 1, 0) = txt_codigo_produto
            grade3.TextMatrix(grade3.Rows - 1, 1) = txt_descricao
            grade3.TextMatrix(grade3.Rows - 1, 2) = strLocacao
            grade3.TextMatrix(grade3.Rows - 1, 3) = cbo_unidade.Text
            grade3.TextMatrix(grade3.Rows - 1, 4) = Format(txt_qtdeconf, "##,###,##0.00")
            grade3.Rows = grade3.Rows + 1
            
        End If
    End If
    
    LimpaConferencia
    Somaprodutosconf
    txt_codigo_produto.SetFocus

Exit Sub
Err_cmd_inserir_Click: ValidaErros Err, Me.Caption & " - cmd_inserir_Click"
End Sub

Private Sub LimpaConferencia()
    txt_codigo_produto = ""
    txt_descricao = ""
    strLocacao = ""
    cbo_unidade = ""
    txt_qtdeconf = Format(1, g_decimal_estoque)
End Sub

Private Function ExisteProdutoConf() As Boolean
On Error GoTo Err_ExisteProdutoConf
Dim i As Long

    ExisteProdutoConf = False
    For i = 1 To grade3.Rows - 1
        If grade3.TextMatrix(i, 0) = txt_codigo_produto Then
           grade3.TextMatrix(i, 4) = Format(CDbl(grade3.TextMatrix(i, 4)) + CDbl(txt_qtdeconf), "##,###,##0.00")
           ExisteProdutoConf = True
        End If
    Next

Exit Function
Err_ExisteProdutoConf: ValidaErros Err, Me.Caption & " - ExisteProdutoConf"
End Function

Private Sub FormaGridConf()
On Error GoTo Err_FormaGridConf

    grade3.Clear
    grade3.GridLines = flexGridFlat
    grade3.AllowUserResizing = flexResizeColumns
    
    grade3.Rows = 2
    grade3.FixedRows = 1
               
    grade3.Cols = 5
    grade3.TextMatrix(0, 0) = "Código"
    grade3.ColWidth(0) = 1300
    grade3.ColAlignmentFixed(0) = 5
    grade3.ColAlignment(0) = flexAlignLeftCenter
    
    grade3.TextMatrix(0, 1) = "Descricao"
    grade3.ColWidth(1) = 6000
    grade3.ColAlignmentFixed(1) = 5
    grade3.ColAlignment(1) = flexAlignLeftCenter
    
    grade3.TextMatrix(0, 2) = "Locacao"
    grade3.ColWidth(2) = 2000
    grade3.ColAlignmentFixed(2) = 5
    grade3.ColAlignment(2) = flexAlignRightCenter
    
    grade3.TextMatrix(0, 3) = "Unid."
    grade3.ColWidth(3) = 500
    grade3.ColAlignmentFixed(3) = 5
    grade3.ColAlignment(3) = flexAlignRightCenter
    
    grade3.TextMatrix(0, 4) = "Quantidade"
    grade3.ColWidth(4) = 1600
    grade3.ColAlignmentFixed(4) = 5
    grade3.ColAlignment(4) = flexAlignRightCenter
    
    Dim i As Long
    For i = grade3.FixedRows To grade3.Rows - 1
        grade3.TextMatrix(i, 4) = Format(fValidaValor(grade3.TextMatrix(i, 4)), "##,###,##0.00")
    Next

Exit Sub
Err_FormaGridConf: ValidaErros Err, Me.Caption & " - FormaGridConf"
End Sub

Private Sub txt_qtdeconf_GotFocus()
    txt_qtdeconf.SelStart = 0
    txt_qtdeconf.SelLength = Len(txt_qtdeconf)
    txt_qtdeconf.BackColor = &HFFFF&
End Sub

Private Sub txt_qtdeconf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_codigo_produto.SetFocus
End Sub

Private Sub txt_qtdeconf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmd_inserir.SetFocus
End Sub

Private Sub txt_qtdeconf_LostFocus()
On Error GoTo file

    If Not IsNumeric(txt_qtdeconf) Then
        txt_qtdeconf.SetFocus
    Else
        txt_qtdeconf = Format(txt_qtdeconf, "##,###,##0.00")
    End If
    
    txt_qtdeconf.BackColor = &H80000005

Exit Sub
file:    If Err.Number = 5 Then Exit Sub
End Sub
Private Sub cmd_inserir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_qtdeconf.SetFocus
End Sub

Private Sub txt_codigo_usuario_GotFocus()
    txt_codigo_usuario.SelStart = 0
    txt_codigo_usuario.SelLength = Len(txt_codigo_usuario)
    txt_codigo_usuario.BackColor = &HFFFF&
End Sub

Private Sub txt_codigo_usuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(txt_codigo_usuario) Then
            frm_consulta_usuario.Show (1)
            If Len(g_string) > 0 Then
                Call BuscaUsuario(g_string)
            End If
        Else
            If Val(txt_codigo_usuario) > 0 Then
                Call BuscaUsuario(txt_codigo_usuario)
            Else
                txt_codigo_usuario.SetFocus
            End If
        End If
    End If

End Sub
Private Sub txt_codigo_usuario_LostFocus()
    txt_codigo_usuario.BackColor = &HFFFFFF
End Sub

Private Sub BuscaUsuario(strcodigo As String)
On Error GoTo Err_BuscaUusario

    Sql_Record_Set.Source = "SELECT codigo,nome,senha FROM usuario WHERE codigo = '" & strcodigo & "' and inativo = 0 LIMIT 1"
    Sql_Record_Set.Open
    If Sql_Record_Set.RecordCount > 0 Then
        txt_codigo_usuario = Sql_Record_Set!codigo
        txt_nome_usuario = Sql_Record_Set!NOME
        strSenha = Sql_Record_Set!Senha
        txt_senha.SetFocus
    Else
        Alerta "Usuário não encontrado", 64 '
        txt_codigo_usuario.SetFocus
        txt_nome_usuario = ""
        strSenha = ""
        txt_senha = ""
    End If
    Sql_Record_Set.Close

Exit Sub
Err_BuscaUusario: ValidaErros Err, Me.Caption & " - BuscaUusario"
End Sub
'==========================================================================
' Purpose:
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 25/06/2013
'==========================================================================
Public Function ValidaBloqueioMovimentacao() As Boolean
On Error GoTo Err_ValidaBloqueioMovimentacao
Dim z               As Integer
Dim strCodProdutos  As String

    For z = 1 To grade1.Rows - 1
        If Trim(grade1.TextMatrix(z, 2)) <> "" Then
            If Trim(strCodProdutos) <> "" Then strCodProdutos = strCodProdutos & ","
            strCodProdutos = strCodProdutos & grade1.TextMatrix(z, 2)
        End If
    Next
    
    ValidaBloqueioMovimentacao = cCtrlEntradaSaida.ValidaBloqueioMovimentacao(strCodProdutos)

Exit Function
Err_ValidaBloqueioMovimentacao: ValidaErros Err, Me.Caption & " - ValidaBloqueioMovimentacao"
End Function

Private Function validacamposii() As Boolean
On Error GoTo Err_validacamposii
Dim z   As Long
Dim i   As Long

    validacamposii = True
    If bolConferencia = True Then validacamposii = True: Exit Function
    'se não esta amarrado a conferência salta fora
    If (g_entrada_amarrado_pedido < 2) Then validacamposii = True: Exit Function
        
    If Trim(grade3.TextMatrix(1, 0)) = "" Then
        Alerta "Itens para conferência não foram lançados"
        SSTab1.Tab = 3
        validacamposii = False
        Exit Function
    End If
    
    For z = 1 To grade1.Rows - 1
        If Trim(grade1.TextMatrix(z, 2)) <> "" Then
            For i = 1 To grade3.Rows - 1
               If Trim(grade3.TextMatrix(i, 0)) <> "" Then
                   If grade1.TextMatrix(z, 2) = grade3.TextMatrix(i, 0) Then
                       If IsNumeric(grade3.TextMatrix(i, 4)) Then
                           If CDbl(grade1.TextMatrix(z, 5)) = CDbl(grade3.TextMatrix(i, 4)) Then
                                Exit For
                           Else
                               validacamposii = False
                               g_string = grade1.TextMatrix(z, 2)
                           End If
                       End If
                   End If
               End If
            Next
        End If
    Next
    
    If validacamposii = False Then
        Alerta "Verifique os Seguintes Dados do Código " & g_string & Chr(13) & "1- Não Confere com a Nota de Entrada " & Chr(13) & "2- Não Lançado na Grade " & Chr(13) & "3- Quantidade Inválida!"
        Exit Function
    End If
    
    For z = 1 To grade3.Rows - 1
        If Trim(grade3.TextMatrix(z, 0)) <> "" Then
            validacamposii = False
            For i = 1 To grade1.Rows - 1
                If grade3.TextMatrix(z, 0) = grade1.TextMatrix(i, 2) Then
                    validacamposii = True
                    Exit For
                End If
            Next
            
            If validacamposii = False Then
                Alerta "Produto Código N. " & grade3.TextMatrix(z, 0) & " Não Lançado na Nota de Entrada!"
                Exit Function
            End If
        End If
    Next

    
    If Not IsNumeric(txt_codigo_usuario) Then
        SSTab1.Tab = 3
        txt_codigo_usuario.SetFocus
        Alerta "Informe o responsável pela conferência"
        validacamposii = False
        Exit Function
    ElseIf Trim(txt_senha) = "" And Toolbar1.Buttons(6).Enabled = False Then
        SSTab1.Tab = 3
        txt_senha.SetFocus
        Alerta "Informe a senha do responsável pela conferência"
        validacamposii = False
        Exit Function
    ElseIf strSenha <> Kriptografa(txt_senha) And Toolbar1.Buttons(6).Enabled = False Then
        SSTab1.Tab = 3
        txt_senha.SetFocus
        Alerta "Senha do responsável pela conferência não confere"
        validacamposii = False
        Exit Function
    End If

Exit Function
Err_validacamposii: ValidaErros Err, Me.Caption & " - validacamposii"
End Function

'*****************************************************************************
'Criação: Diego Martins dos Santos                      Data: 09/06/2010
'
'Propósito:
'*****************************************************************************
Private Function ValidaCodigoBarras(ByRef z As Long) As Boolean
On Error GoTo Err_ValidaCodigoBarras

    ValidaCodigoBarras = False
    If g_impressao_codbarras = 0 Then Exit Function
    If g_impressao_codbarras = 1 And chk_impressao_nf.Value = 0 Then Exit Function
    
    For z = 1 To grade1.Rows - 1
        
        If (Trim(grade1.TextMatrix(z, 2)) <> "") And (Trim(grade1.TextMatrix(z, 38)) = "") Then
           ValidaCodigoBarras = True
           Exit Function
        End If
    
    Next

Exit Function
Err_ValidaCodigoBarras: ValidaErros Err, Me.Caption & " - ValidaCodigoBarras"
End Function
Private Sub Somaprodutosconf()
On Error GoTo Err_Somaprodutosconf
Dim z As Long

    txt_totalqtde = 0
    For z = 1 To grade3.Rows - 1
        If Trim(grade3.TextMatrix(z, 0)) <> "" Then
            txt_totalqtde = CDbl(txt_totalqtde) + CDbl(grade3.TextMatrix(z, 4))
        End If
    Next
    
    txt_totalqtde = Format(txt_totalqtde, "##,###,##0.00")

Exit Sub
Err_Somaprodutosconf: ValidaErros Err, Me.Caption & " - Somaprodutosconf"
End Sub

Private Function AtualizaConferencia() As Boolean
On Error GoTo Err_AtualizaConferencia

    AtualizaConferencia = False

    If bolConferencia = True Then
        AtualizaConferencia = True
        Exit Function
    End If

    If AtualizaMovimentoConferencia = False Then Exit Function

    '*** Exclui Dados da Tabela Temporária ***
    If Not Conexao.DeleteSintetico("movimento_cabecalho_nota_fiscal_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, cDPEmpresa.codigo) Then Exit Function
    If Not Conexao.DeleteSintetico("movimento_nota_fiscal_entrada_tmp", "empresa = '" & cDPEmpresa.codigo & "' and sequencia = " & lSequencia, cDPEmpresa.codigo) Then Exit Function

    AtualizaConferencia = True

Exit Function
Err_AtualizaConferencia: ValidaErros Err, Me.Caption & " - AtualizaConferencia"
End Function

Private Sub AtualTelaConferencia()
On Error GoTo Err_AtualTelaConferencia
Dim i As Long

    Sql_Query = "SELECT CME.*,PR.descricao,PR.locacao_1,PR.unidade,US.nome " & _
                "FROM conferencia_mercadoria_entrada CME " & _
                "INNER JOIN produto PR ON PR.codigo = CME.codigo_produto " & _
                "INNER JOIN usuario US ON US.codigo = CME.codigo_usuario " & _
                "WHERE sequencia_nf = '" & lSequencia & "' "
    Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 0)
    i = 1
    If gb_Recordset.RecordCount > 0 Then
        txt_codigo_usuario = gb_Recordset!codigo_usuario
        txt_nome_usuario = gb_Recordset!NOME
        Do Until gb_Recordset.EOF
            grade3.TextMatrix(i, 0) = gb_Recordset!codigo_produto
            grade3.TextMatrix(i, 1) = gb_Recordset!Descricao
            grade3.TextMatrix(i, 2) = gb_Recordset!locacao_1
            grade3.TextMatrix(i, 3) = gb_Recordset!Unidade
            grade3.TextMatrix(i, 4) = Format(gb_Recordset!Quantidade, "##,###,##0.00")
            
            grade3.Rows = grade3.Rows + 1
            i = i + 1
            
            gb_Recordset.MoveNext
        Loop
    End If
    gb_Recordset.Close
    
    Somaprodutosconf

Exit Sub
Err_AtualTelaConferencia: ValidaErros Err, Me.Caption & " - AtualTelaConferencia"
End Sub

Private Sub ImprimirConf()
On Error GoTo Err_ImprimirConf

    frm_seleciona_impressora.Show 1
    Set Printer = Printers(Impressora_default)
            
    BD_Record_Set.Source = "SELECT * FROM dados_impressora"
    BD_Record_Set.Open
    If BD_Record_Set.RecordCount > 0 Then
        'para tamanho 12 matricial
        If g_tipo_impressora = 0 Then
            g_tamanho_pagina = BD_Record_Set!linha_matricial1
            g_tamanho_pagina_final = BD_Record_Set!final_matricial1
        'para tamanho 8 jato tinta
        Else
            g_tamanho_pagina = BD_Record_Set!linha_tinta3
            g_tamanho_pagina_final = BD_Record_Set!final_tinta3
        End If
    End If
    BD_Record_Set.Close
    
    lPagina = 0
    lLinha = 0
    ImpCabConf
    ImpDetConf
    
    If lPagina > 0 Then
        ImpTotalConf
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        BioFechaImprime1
        frm_preview.txt_preview.FileName = "c:\sabre\Arquivo" & g_usuario & ".txt"
        frm_preview.lbl_senha = ""
        frm_preview.lbl_nome = ""
        frm_preview.Show 1
    End If

Exit Sub
Err_ImprimirConf: ValidaErros Err, Me.Caption & " - ImprimirConf"
End Sub



Private Sub ImpDetConf()
On Error GoTo Err_ImpDetConf
Dim x_linha     As String
Dim i           As Long
Dim z           As Long


    For i = 1 To grade1.Rows - 1
        If lLinha >= g_tamanho_pagina Then
            x_linha = "------------------------------------------------------------------------------------------------"
            BioImprime "@Printer.Print " & x_linha
            BioImprime1 x_linha
            BioImprime "@@Printer.NewPage"
            ImpCabConf
        End If
                
        '                   1         2         3         4         5         6         7         8         9        10        11        12        13        14        15      15
        '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678
        x_linha = "Código        Descrição                                       Lote                    Quantidade"
        x_linha = "                                                                                                "
        If IsNumeric(grade3.TextMatrix(i, 0)) Then
            z = Len(Format(grade3.TextMatrix(i, 0), "000"))
            Mid(x_linha, 1 + 13 - z, z) = Format(grade3.TextMatrix(i, 0), "000")
        Else
            Mid(x_linha, 1, 13) = grade3.TextMatrix(i, 0)
        End If
        
        Mid(x_linha, 15, 43) = grade3.TextMatrix(i, 1)
        Mid(x_linha, 63, 21) = grade3.TextMatrix(i, 2)
        
        z = Len(Format(grade3.TextMatrix(i, 3), g_decimal_estoque))
        Mid(x_linha, 85 + 12 - z, z) = Format(grade3.TextMatrix(i, 3), g_decimal_estoque)
            
        BioImprime "@Printer.Print " & x_linha
        BioImprime1 x_linha
        lLinha = lLinha + 1
    Next

Exit Sub
Err_ImpDetConf: ValidaErros Err, Me.Caption & " - ImpDetConf"
End Sub

Private Sub ImpTotalConf()
On Error GoTo Err_ImpTotalConf
Dim x_linha     As String
Dim z           As Long

    With BD_Record_Set
        If lLinha >= g_tamanho_pagina Then
            x_linha = "------------------------------------------------------------------------------------------------"
            BioImprime "@Printer.Print " & x_linha
            BioImprime1 x_linha
            BioImprime "@@Printer.NewPage"
            ImpCabConf
        End If
                
        '                   1         2         3         4         5         6         7         8         9        10        11        12        13        14        15      15
        '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678
        x_linha = "------------------------------------------------------------------------------------------------"
        BioImprime "@Printer.Print " & x_linha
        BioImprime1 x_linha
        
        x_linha = ""
        BioImprime "@Printer.Print " & x_linha
        BioImprime1 x_linha
        x_linha = "                                                           Totalizador:                        "
        z = Len(txt_totalqtde)
        Mid(x_linha, 85 + 12 - z, z) = txt_totalqtde
        BioImprime "@Printer.Print " & x_linha
        BioImprime1 x_linha
        lLinha = lLinha + 3
        
    End With

Exit Sub
Err_ImpTotalConf: ValidaErros Err, Me.Caption & " - ImpTotalConf"
End Sub

Private Sub ImpCabConf()
On Error GoTo Err_ImpCabConf
Dim x_linha As String
Dim i       As Long

    If lPagina = 0 Then
        BioCriaImprime
        BioCriaImprime1
        If g_tipo_impressora = 0 Then
            BioImprime "@@Printer.ScaleMode = 7"
            BioImprime "@@Printer.PaperSize = 1"
        'impressora jato tinta
        ElseIf g_tipo_impressora = 1 Then
            BioImprime "@@Printer.ScaleMode = 7"
            BioImprime "@@Printer.PaperSize = 9"
        End If

    End If
    lPagina = lPagina + 1
    lLinha = 0
    
    If g_tipo_impressora = 0 Then
        BioImprime "@@Printer.FontName = Draft 12cpi"
    ElseIf g_tipo_impressora = 1 Then
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontSize = 8"
    End If
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.FontBold = True"

    x_linha = "                                                                                   PAGINA: ___  "
    Mid(x_linha, 3, 40) = cDPEmpresa.NOME
    Mid(x_linha, 92, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "  RELATÓRIO REPOSIÇÃO DE ESTOQUE                                            CIDADE, __/__/____  "
    i = Len(Trim(cDPEmpresa.NomeCidade))
    Mid(x_linha, 53 + 30 - i, i) = Trim(cDPEmpresa.NomeCidade)
    Mid(x_linha, 85, 10) = Date
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "================================================================================================"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "Código        Descrição                                       Lote                    Quantidade"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "================================================================================================"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    lLinha = lLinha + 6
    

Exit Sub
Err_ImpCabConf: ValidaErros Err, Me.Caption & " - ImpCabConf"
End Sub

Private Sub ChamaGradeTamanho(strcodigo As String, strRow As Long)
On Error GoTo Err_ChamaGradeTamanho

    'não chama grade quando não influencia estoque ou venda CI
    If cDPFFaturamento.InfluenciaEstoque = 0 Or pct_icms.Visible = True Then Exit Sub
    
    frm_saidaentradagrade.txt_codproduto = strcodigo
    frm_saidaentradagrade.lngCodigoTMP = lngIDGrade
    frm_saidaentradagrade.strStatus = "T"
    frm_saidaentradagrade.lngSequencia = 0 'Variavel Só Preenchida Qdo Se Tratar de Devolução;
    frm_saidaentradagrade.strTipoMovimento = "E" 'Identifica qual Operação o Usuário esta fazendo;
    frm_saidaentradagrade.bolDevolucao = False 'Variavel Só Preenchida Qdo Se Tratar de Devolução.
    frm_saidaentradagrade.Show 1
    lngIDGrade = g_string
    grade1.TextMatrix(strRow, 5) = Format(Arredonda(BuscaQuantidadeGrade(str(lngIDGrade), strcodigo)), g_decimal_estoque)

Exit Sub
Err_ChamaGradeTamanho: ValidaErros Err, Me.Caption & " - ChamaGradeTamanho"
End Sub

Private Function AtualizaMovimentoGrade() As Boolean
On Error GoTo Err_AtualizaMovimentoGrade

    AtualizaMovimentoGrade = False
    
    'efetua relacionamento com a grade e dar baixa no estoque
    'quando houver codigo de movimentacao da grade
    'se não veio do VENDAS/OR pois se veio de la já tem que ter baixado o estoque da grade
    'se não e venda CI
    Call VerificaConsistenciaProdutos(lngIDGrade, lSequencia, "movimento_nota_fiscal_entrada", "codigo_do_produto")
    If lngIDGrade > 0 And pct_icms.Visible = False Then
        If Not AtualizaSequenciaGrade(lngIDGrade, lSequencia, "E", "E") Then Exit Function
        If Not AlteraEstoqueGrade(lngIDGrade) Then Exit Function
    End If
    
    AtualizaMovimentoGrade = True

Exit Function
Err_AtualizaMovimentoGrade: ValidaErros Err, Me.Caption & " - AtualizaMovimentoGrade"
End Function


'*****************************************************************************
'Criação: Ronaldo Robledo                               Data: 21/08/2010
'
'Propósito: Validar os campos quando se tratar de utilização da grade
'*****************************************************************************
Private Function validagrade() As Boolean
On Error GoTo Err_validagrade
Dim i As Long

    validagrade = False
    
    For i = 1 To grade1.Rows - 1
        
        If Trim(grade1.TextMatrix(i, 2)) <> "" Then
            'funcões de verificação da grade
            If lngIDGrade = 0 And Val(grade1.TextMatrix(i, 39)) = 1 Then
                Alerta "Efetue o lançamento da quantidade da grade "
                Exit Function
            '*** Removido pois na entrada não precisa validar OS 13628 item (C)
            'Verificar com Sr.Ronaldo
            'ElseIf Not ValidaEstoqueGrade(str(lngIDGrade), grade1.TextMatrix(i, 2), grade1.TextMatrix(i, 39)) Then
             '   Exit Function
            ElseIf Not VerificaQtdGradeComQtdeTela(lngIDGrade, grade1.TextMatrix(i, 2), grade1.TextMatrix(i, 5), grade1.TextMatrix(i, 39)) Then
                grade1.SetFocus
                Exit Function
            End If
        End If
    Next
    
    validagrade = True

Exit Function
Err_validagrade: ValidaErros Err, Me.Caption & " - ValidaGrade"
End Function

Private Sub CalculaUnidadeSecundaria(lngRow As Integer, bolSalto As Boolean)
On Error GoTo Err_CalculaUnidadeSecundaria

    grade1.TextMatrix(lngRow, 4) = UCase(grade1.TextMatrix(lngRow, 4))
    Sql_Query = "SELECT US.codigo, US.unidade, US.fator_operacao,(ES.quantidade_cx - ES.qtde_transito_cx) as quantidade_cx, PR.indexado_dolar,PR.custo,PR.preco_varejo," & _
                "PR.preco_atacado,PR.data_inicial,PR.data_final,PR.preco_promocao,PR.desconto_atacado,PR.desconto_varejo " & _
                "FROM unidade_secundaria US, estoque ES, produto PR " & _
                "WHERE (PR.codigo = (SELECT codigo FROM codigo_barras WHERE codigo_barras = '" & grade1.TextMatrix(lngRow, 2) & _
                "' LIMIT 1) and  US.empresa = '" & cDPEmpresa.codigo & "' and US.codigo = PR.codigo and " & _
                "US.unidade = '" & grade1.TextMatrix(lngRow, 4) & "' and ES.codigo_do_produto = PR.codigo) or " & _
                "(PR.codigo = '" & grade1.TextMatrix(lngRow, 2) & "' and US.empresa = '" & cDPEmpresa.codigo & "' and " & _
                "US.codigo = PR.codigo and US.unidade = '" & grade1.TextMatrix(lngRow, 4) & "'  and  ES.codigo_do_produto = PR.codigo)"
    Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 1)
    If gb_Recordset.RecordCount > 0 Then
        LastRow = lngRow
        grade1.TextMatrix(lngRow, 44) = Format(CDbl(grade1.TextMatrix(lngRow, 5)) * CDbl(gb_Recordset!fator_operacao), g_decimal_estoque)
        grade1.TextMatrix(lngRow, 45) = Format(gb_Recordset!fator_operacao, g_decimal_estoque)
        If bolSalto = True Then grade1.col = grade1.col + 1
    Else
        gb_Recordset.Close
        Sql_Query = "SELECT ES.quantidade_cx,PD.codigo,PD.unidade,PD.indexado_dolar,PD.custo,PD.preco_varejo,PD.preco_atacado," & _
                    "PD.data_inicial,PD.data_final,PD.preco_promocao,PD.desconto_atacado,PD.desconto_varejo " & _
                    "FROM estoque ES,produto PD " & _
                    "WHERE (PD.codigo = (SELECT codigo FROM codigo_barras WHERE codigo_barras = '" & grade1.TextMatrix(lngRow, 2) & _
                    "' LIMIT 1) and ES.codigo_do_produto = PD.codigo) or " & _
                    "(PD.codigo = '" & grade1.TextMatrix(lngRow, 2) & "' and ES.codigo_do_produto = PD.codigo)"
        Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 1)
        If gb_Recordset.RecordCount > 0 Then
            grade1.TextMatrix(lngRow, 4) = gb_Recordset!Unidade
            LastRow = lngRow
            grade1.TextMatrix(lngRow, 44) = Format(grade1.TextMatrix(lngRow, 4), g_decimal_estoque)
            grade1.TextMatrix(lngRow, 45) = 1
            If bolSalto = True Then grade1.col = grade1.col + 1
        Else
            MsgBox "Unidade de Medida Não Localizada ", vbCritical, "Atenção!"
            grade1.TextMatrix(lngRow, 4) = ""
        End If
    End If
    gb_Recordset.Close

Exit Sub
Err_CalculaUnidadeSecundaria: ValidaErros Err, Me.Caption & " - CalculaUnidadeSecundaria"
End Sub

'*****************************************************************************
'Criação: Thiago Leão                                         Data: 13/09/2011
'Propósito: Preenche a combobox com as unidades disponiveis do produto!
'*****************************************************************************

Private Sub PreencheComboBox()
On Error GoTo Err_PreencheComboBox
Dim strResultado() As String
Dim str         As String
Dim i           As Integer

    i = 0
    Sql_Query = "SELECT unidade FROM unidade_secundaria WHERE empresa = " & cDPEmpresa.codigo & _
                " and codigo = " & txt_codigo_produto
    Set gb_Recordset = Conexao.GeraRecordset(Sql_Query, 0)
    
    If gb_Recordset.RecordCount > 0 Then
        ReDim strResultado(gb_Recordset.RecordCount)
        With gb_Recordset
            Do Until .EOF
                str = !Unidade
                strResultado(i) = str
                i = i + 1
                cbo_unidade.AddItem (str)
                .MoveNext
            Loop
        End With
        cbo_unidade.Text = cbo_unidade.List(0)
    Else
        cbo_unidade.Text = cbo_unidade.List(0)
    End If
    If gb_Recordset.State = 1 Then gb_Recordset.Close 'Fecha o RecordSet se aberto

Exit Sub
Err_PreencheComboBox: ValidaErros Err, Me.Caption & " - PreencheComboBox"
End Sub

'==========================================================================
' Purpose:  Efetuar a busca dos dados tributários referente a cada produto no
'           cadastro de tributação conjunto de chave de procura (UF - codigo grupo tributacao - codigo forma faturamento)
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 30/01/2012
' Alteeração: Ronaldo Robledo                                    17/08/2012
'             Reimplementado para utilização definitiva
'==========================================================================
Private Function ObtemDadosTributacao(ByVal i As Long) As Boolean
On Error GoTo Err_ObtemDadosTributacao
Dim DPPropTrib     As New cDPPropTributacao
    
    ObtemDadosTributacao = False
    
    If cDPNotasSaida.UFPessoa = "" Then PreencheDominioProblema
    Set DPPropTrib = colTributacao(CStr(cDPNotasSaida.UFPessoa & "." & grade1.TextMatrix(i, 47) & "." & cDPNotasSaida.CodFormaFaturamento))
                                                   
    grade1.TextMatrix(i, 21) = DPPropTrib.AliquotaPis
    grade1.TextMatrix(i, 22) = DPPropTrib.AliquotaCofins
    If (Trim(grade1.TextMatrix(i, 28)) = "") Or (lngCodFormaFatuAnterior <> Val(txt_codigo_forma)) Then
        grade1.TextMatrix(i, 28) = DPPropTrib.CodigoCFOP
        lbl_cfop = DPPropTrib.CodigoCFOP
    End If
    
    If grade1.TextMatrix(i, 43) = "" Then
        If cDPEmpresa.SuperSimples = 3 Then 'regime normal
            grade1.TextMatrix(i, 30) = DPPropTrib.OrigemMercadoria & DPPropTrib.CST
        Else
            grade1.TextMatrix(i, 30) = DPPropTrib.CST 'simples
        End If
    End If
    If Val(grade1.TextMatrix(i, 42)) = 0 Then
        grade1.TextMatrix(i, 42) = DPPropTrib.Iva
    End If
    grade1.TextMatrix(i, 48) = DPPropTrib.CstPis
    grade1.TextMatrix(i, 49) = DPPropTrib.CstCofins
    grade1.TextMatrix(i, 50) = DPPropTrib.CstIpi

    ObtemDadosTributacao = True
    Set DPPropTrib = Nothing
 
Exit Function
Err_ObtemDadosTributacao:
    If Err.Number = 5 Then
        Alerta "Tributação não encontrada para o produto " & grade1.TextMatrix(i, 2) & Chr(13) & "Verifique seu cadastro de tributação associado a este produto"
    Else
        ValidaErros Err, Me.Caption & " - ObtemDadosTributacao"
    End If
End Function
'==========================================================================
' Purpose:  Preenche objeto DP
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 11/08/2012
'==========================================================================
Private Sub PreencheDominioProblema()
On Error GoTo Err_PreencheDominioProblema
    
    cDPNotasSaida.CodPessoa = Val(txt_codigo_fornecedor)
    cDPNotasSaida.NomePessoa = txt_fornecedor
    cDPNotasSaida.UFPessoa = lbl_uf
    cDPNotasSaida.CodCondicaoPagamento = Val(txt_codigo_forma)
    cDPNotasSaida.NomeCondPagamento = lbl_forma
    cDPNotasSaida.CodFormaFaturamento = Val(txt_codigo_forma)
    cDPNotasSaida.NomeFormaFaturamento = lbl_forma
    cDPNotasSaida.Empresa = cDPEmpresa.codigo
    cDPNotasSaida.VarejoAtacadoOutras = "O"
    
Exit Sub
Err_PreencheDominioProblema: ValidaErros Err, Me.Caption & " - PreencheDominioProblema"
End Sub

'==========================================================================
' Purpose:  Busca dados para preencher campos da forma de faturamento
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 11/08/2012
'==========================================================================
Private Sub CarregaCamposFaturamento()
On Error GoTo Err_CarregaCamposFaturamento

    Call PreencheDominioProblema
    Call cCtrlEntradaSaida.CarregaCamposFaturamento(cDPNotasSaida, cDPFFaturamento, colFormaFaturamento, 0, cDPFFaturamento.ChaveColecao)
    txt_codigo_forma = cDPNotasSaida.CodFormaFaturamento
    lbl_forma = cDPNotasSaida.NomeFormaFaturamento
    
Exit Sub
Err_CarregaCamposFaturamento: ValidaErros Err, Me.Caption & " - CarregaCamposFaturamento"
End Sub

Private Sub cmd_ler_xml_Click()
    If Trim(txt_caminho_xml) <> "" Then
        Call PreencheCabecalho
        Call PreencheTransporte
        Call PreencheProduto(txt_caminho_xml)
    Else
        Alerta "Selecione um arquivo XML válido, para efetuar a importação!", 48
    End If
End Sub

'*****************************************************************************
'Criação: Thiago Leão                                         Data: 07/02/2012
'Propósito: Preenche todos os campos do cabeçalho da nota.
'           A busca do fornecedor é feita pelo CNPJ da mesma. O fornecedor tem
'           que estar cadastrado no sistema com o CNPJ da NF-e!
'*****************************************************************************
Private Sub PreencheCabecalho()
On Error GoTo Err_PreencheCabecalho

    Call BuscaFornecedor(0, cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "emit//CNPJ"))                                  'Fornecedor
    txt_numero_nf.Text = Val(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "ide//nNF"))                                   'Número
    txt_serie_nf.Text = Val(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "ide//serie"))                                  'Série
    txt_modelo_nf.Text = Val(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "ide//mod"))                                   'Modelo
    lbl_forma.Caption = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "ide//natOp")                                       'Natureza
    msk_emissao.Text = Format(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "ide//dEmi"), "dd/mm/yyyy")                   'Data da emissão
    txt_chave_acesso.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "infProt//chNFe")                               'Chave de acesso
    txt_total.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vNF"), ".", ",")              'Total da nota
    txt_bc_icms.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vBC"), ".", ",")            'Base Calculo ICMS
    txt_icms.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vICMS"), ".", ",")             'Valor ICMS
    txt_bc_substituicao.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vBCST"), ".", ",")  'Base Subs. Trib.
    txt_substituicao.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vST"), ".", ",")       'Valor ST
    txt_ipi.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vIPI"), ".", ",")               'Valor IPI
    txt_frete.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vFrete"), ".", ",")           'Valor Frete
    txt_seguro.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vSeg"), ".", ",")            'Valor Seguro
    txt_desconto.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vDesc"), ".", ",")         'Valor Desconto
    txt_outras.Text = Replace(cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vOutro"), ".", ",")          'Valor Outro

    
Exit Sub
Err_PreencheCabecalho: ValidaErros Err, Me.Caption & " - PreencheCabecalho"
End Sub

'*****************************************************************************
'Criação: Thiago Leão                                         Data: 09/02/2012
'Propósito: Preenche os campos da transportadora com os dados no XML.
'           O sistema busca os dados pelo CNPJ da transp., onde o mesmo tem
'           que estar cadastrado.
'*****************************************************************************
Private Sub PreencheTransporte()
On Error GoTo Err_PreencheTransporte
Dim strCNPJTranportadora As String

    strCNPJTranportadora = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//transporta//CNPJ")
    If Trim(strCNPJTranportadora) <> "" Then
        Call BuscaTransportadora(0, strCNPJTranportadora)
        txt_volume.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//qVol")
        txt_especie.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//esp")
        txt_peso_liquido.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//pesoL")
        txt_peso_bruto.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//pesoB")
        txt_transportadora_placa.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//veicTransp//placa")
        txt_transportadora_placa_uf.Text = cUtGeral.RetornaTagXML((Trim(txt_caminho_xml)), "transp//veicTransp//UF")
    End If
Exit Sub
Err_PreencheTransporte: ValidaErros Err, Me.Caption & " - PreencheTransporte"
End Sub

Public Sub PreencheProduto(strCaminhoXML As String)
On Error Resume Next
Dim lngItem         As Long
Dim lngLinha        As Long
Dim XML             As DOMDocument
Dim xmlElem         As IXMLDOMNode
Dim bolFimProdutos  As Boolean
Dim StrCodBarras    As String
Set XML = New DOMDocument
    
    lngItem = 0
    lngLinha = 1
    XML.async = False
    bolFimProdutos = False
    If XML.Load(strCaminhoXML) Then
        
        Do Until bolFimProdutos
            Set xmlElem = XML.selectNodes("/nfeProc/NFe/infNFe/det").Item(lngItem).firstChild
                
                '*** Verfica se tem algum valor no código do produto, senão tiver finaliza o loop ***
                If xmlElem.selectSingleNode("cProd").Text = "" Then bolFimProdutos = True: Exit Do
                
                StrCodBarras = xmlElem.selectSingleNode("cEAN").Text
                
                If chk_ImportarProdCodigoBarras.Value = 1 Then
                   If Trim(StrCodBarras) = "" Then StrCodBarras = xmlElem.selectSingleNode("cProd").Text
                   LastRow = lngLinha
                   Call BuscaProdutos(StrCodBarras, 1)
                End If
                    
                If Trim(grade1.TextMatrix(lngLinha, 2)) = "" Or Trim(grade1.TextMatrix(lngLinha, 2)) = "0" Then
                    grade1.TextMatrix(lngLinha, 2) = xmlElem.selectSingleNode("cProd").Text                                          'Codigo produto
                    grade1.TextMatrix(lngLinha, 3) = xmlElem.selectSingleNode("xProd").Text                                          'Descrição produto
                    Call DestacarLinha("&HFFC0C0", lngLinha)
                End If
                
                grade1.TextMatrix(lngLinha, 4) = xmlElem.selectSingleNode("uCom").Text                                               'Unidade produto
                grade1.TextMatrix(lngLinha, 5) = Format(Replace(xmlElem.selectSingleNode("qCom").Text, ".", ","), g_decimal_estoque) 'Quantidade Produto
                grade1.TextMatrix(lngLinha, 6) = Replace(xmlElem.selectSingleNode("vUnCom").Text, ".", ",")                          'Valor unitário Produto
                grade1.TextMatrix(lngLinha, 7) = Replace(xmlElem.selectSingleNode("vDesc").Text, ".", ",")                           'Valor Desconto Produto
                grade1.TextMatrix(lngLinha, 8) = Replace(xmlElem.selectSingleNode("vUnCom").Text, ".", ",")                          'Valor unitário Produto
                grade1.TextMatrix(lngLinha, 9) = Format(fValidaValorNovo(txt_PctIPI), "##,###,##0.00")                               '% IPI
                grade1.TextMatrix(lngLinha, 10) = Format(fValidaValorNovo(txt_PctICMS), "##,###,##0.00")                             '% ICMS
                grade1.TextMatrix(lngLinha, 11) = Replace(xmlElem.selectSingleNode("vProd").Text, ".", ",")                          'Valor Total Produto
                grade1.TextMatrix(lngLinha, 17) = "NF"
                'grade1.TextMatrix(lngLinha, 28) = xmlElem.SelectSingleNode("CFOP").Text                                              'CFOP Produto
                grade1.TextMatrix(lngLinha, 38) = StrCodBarras                                                                       'Cód. barras produto
                grade1.TextMatrix(lngLinha, 40) = xmlElem.selectSingleNode("NCM").Text                                               'Código NCM Produto
                
                xmlElem.selectSingleNode("cProd").Text = "" 'Limpa o objeto para setar um novo.

                lngItem = lngItem + 1
                lngLinha = lngLinha + 1
                Call ChamaCelula
        Loop
    Else
        MsgBox "Não foi possível abrir o arquivo XML da NFe especificada para Leitura.", vbCritical, "Erro."
    End If
End Sub

'*****************************************************************************
'Criação: Thiago Leão                                         Data: 11/02/2012
'Propósito: Ativar/Desativar componentes da importação do XML.
'*****************************************************************************
Private Sub StatusImportarXML(bolStatus As Boolean)
On Error GoTo Err_StatusImportarXML
    
    cmd_explorer.Enabled = bolStatus
    cmd_ler_xml.Enabled = bolStatus

Exit Sub
Err_StatusImportarXML: ValidaErros Err, Me.Caption & " - StatusImportarXML"
End Sub


Private Sub txt_redicmssubst_GotFocus()
    txt_redicmssubst.BackColor = 12648447
    txt_redicmssubst.SelStart = 0
    txt_redicmssubst.SelLength = Len(txt_redicmssubst)
End Sub

Private Sub txt_redicmssubst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txt_observacoes.SetFocus
End Sub

Private Sub txt_redicmssubst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grade1.SetFocus
End Sub

Private Sub txt_redicmssubst_LostFocus()
    txt_redicmssubst.BackColor = &H8000000E
    txt_redicmssubst = Format(fValidaValor(txt_redicmssubst), "##,###,##0.00")
End Sub



'==========================================================================
' Purpose:  Efetuar o rateio da diferença entre os itens da nota
' Input:
'   ByVal dblDiferenca As Double - valor da diferença
'   ByVal intCol As Integer - coluna que referencia o cálculo que será lançado o rateio
'   ByVal dblValorNota As Double - representa o valor do campo referenciado que influenciara no calculo do rateio
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 01/08/2012
'==========================================================================
Private Sub Rateio(ByVal dblDiferenca As Double, ByVal intCol As Integer, ByVal dblValorNota As Double)
On Error GoTo Err_Rateio
Dim dblRateio   As Double
Dim intDec      As Integer
Dim i           As Long

    'arredonda para 2 quando subtotal
    If intCol = 11 Then intDec = 2 Else intDec = 4
    
    For i = 1 To grade1.Rows - 1
        If Trim(grade1.TextMatrix(i, 3)) <> "" Then
            'se a diferença estiver dentro dos 0,05 lança direto
            If dblDiferenca >= -0.05 And dblDiferenca <= 0.05 Then
                If CDbl(grade1.TextMatrix(i, intCol)) >= Abs(dblDiferenca) Then
                    grade1.TextMatrix(i, intCol) = Format(Round(CDbl(grade1.TextMatrix(i, intCol)) + dblDiferenca, intDec), "##,###,##0.0000")
                    Exit For
                End If
            Else
                'se for maior efetua o rateio
                dblRateio = Format(Round(dblDiferenca / dblValorNota * CDbl(grade1.TextMatrix(i, intCol)), intDec), "##,###,##0.0000")
                
                If CDbl(grade1.TextMatrix(i, intCol)) >= Abs(dblRateio) Then
                    grade1.TextMatrix(i, intCol) = Format(Round(CDbl(grade1.TextMatrix(i, intCol)) + dblRateio, intDec), "##,###,##0.0000")
                End If
            End If
        End If
    Next
Exit Sub
Err_Rateio: ValidaErros Err, Me.Caption & " - Rateio"
End Sub

'
''==========================================================================
'' Purpose:  Recalcular a soma dos 4totais que necessitarem efetuar o rateio da diferença
'' Input:
''   ByRef lblValor As Label - valor que será retornado do recalculo
''   ByVal intCol As Integer - coluna que se referencia o recalculo
'' Output:
'' Remarks:
'' Author: Ronaldo Robledo                                Start: 01/08/2012
''==========================================================================
'Private Sub RecalculaSoma(ByRef lblValor As Label, ByVal intCol As Integer)
'On Error GoTo Err_RecalculaSoma
'    lblValor = 0
'    For i = 1 To grade1.Rows - 1
'        If Trim(grade1.TextMatrix(i, 3)) <> "" Then
'            lblValor = CDbl(grade1.TextMatrix(i, intCol)) + CDbl(lblValor)
'        End If
'    Next
'    lblValor = Format(Round(lblValor, 2), "##,###,##0.00")
'
'Exit Sub
'Err_RecalculaSoma: ValidaErros Err, Me.Caption & " - RecalculaSoma"
'End Sub
Private Sub Imprime_NF()
On Error GoTo Err_Imprime_NF
Dim z As Long

    Sql_Record_Set.Source = "Select * From configuracao_nf"
    Sql_Record_Set.Open
        With Sql_Record_Set
            If .RecordCount > 0 Then
                'Call BuscaCFOP
                
                Set Printer = Printers(g_impressora_default_notas)
                Call Imprime_Cabecalho_NF(txt_numero_nf, "E", lNomeCodificacaoFiscal, lbl_cfop, txt_fornecedor, lbl_cgc, lbl_endereco, lbl_bairro, lbl_cep, lbl_cidade, lbl_telefone, lbl_uf, lbl_inscricao, lPessoa, msk_entrada)
                 
                FimdeProdutos = 0
                For z = 0 To grade1.Rows - 1
                    If IsNumeric(grade1.TextMatrix(z, 0)) Then
                        If FimdeProdutos <= !qtde_itens Then
                            Call Imprime_Produtos_NF(TrataCodigoProdutos(grade1.TextMatrix(z, 2), grade1.TextMatrix(z, 38)), grade1.TextMatrix(z, 3), grade1.TextMatrix(z, 10), grade1.TextMatrix(z, 16), "", grade1.TextMatrix(z, 4), grade1.TextMatrix(z, 5), grade1.TextMatrix(z, 8), 0, grade1.TextMatrix(z, 11), grade1.TextMatrix(z, 10), FimdeProdutos)
                            FimdeProdutos = FimdeProdutos + 1
                        Else
                            Call Imprime_Calculo_Imposto("Continua Fl.Seguinte", "XXXXXXXXXX", "XXXXXXXXXX", "", "", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX")
                            Printer.NewPage
                            Call Imprime_Cabecalho_NF(txt_numero_nf, "E", lNomeCodificacaoFiscal, lbl_cfop, txt_fornecedor, lbl_cgc, lbl_endereco, lbl_bairro, lbl_cep, lbl_cidade, lbl_telefone, lbl_uf, lbl_inscricao, lPessoa, msk_entrada)
                            FimdeProdutos = 0
                            z = z - 1
                        End If
                    End If
                Next
                            
                Call Imprime_Calculo_Imposto(0, 0, 0, "", "", 0, lbl_bc_icms, lbl_valor_icms, lbl_bc_substituicao, lbl_icms_substituicao, lbl_total_produtos, lbl_frete, lbl_outras_despesas, lbl_ipi, lbl_total, lbl_seguro)
                Call Imprime_Transportadora_NF(txt_transportadora_nome, cbo_frete.Text, txt_transportadora_placa, txt_transportadora_placa_uf, txt_transportadora_cnpj, txt_transportadora_endereco, txt_transportadora_cidade, txt_transportadora_uf, txt_transportadora_inscricao_estadual, txt_volume, txt_especie, txt_peso_bruto, txt_peso_liquido)
                Call Imprime_Dados_Adicionais(cDPFFaturamento.MensagemPadrao, "", lbl_forma, txt_observacoes, txt_numero_nf, "", "E", 0, txt_fornecedor, lbl_total)
            Else
                Alerta "Não existe dados da Configuração da NF!"
                Sql_Record_Set.Close
                Exit Sub
            End If
        End With
    Sql_Record_Set.Close

Exit Sub
Err_Imprime_NF: ValidaErros Err, Me.Caption & " - Imprime_NF"
End Sub


Private Sub Imprime_NFNovo()
On Error GoTo file
Dim z As Long

    Sql_Record_Set.Source = "SELECT * FROM configuracao_nf WHERE tipo_nota = 1"
    Sql_Record_Set.Open
    With Sql_Record_Set
        If .RecordCount > 0 Then
            'Call BuscaCFOP
            
            Call Imprime_Cabecalho_NFNovo(txt_numero_nf, "E", lNomeCodificacaoFiscal, lbl_cfop, txt_fornecedor, lbl_cgc, lbl_endereco, lbl_bairro, lbl_cep, lbl_cidade, lbl_telefone, lbl_uf, lbl_inscricao, lPessoa, msk_entrada, g_impressora_default_notas)
             
            FimdeProdutos = 0
            For z = 0 To grade1.Rows - 1
                If IsNumeric(grade1.TextMatrix(z, 0)) Then
                    If FimdeProdutos <= !qtde_itens Then
                        Call Imprime_Produtos_NFNovo(TrataCodigoProdutos(grade1.TextMatrix(z, 2), grade1.TextMatrix(z, 38)), grade1.TextMatrix(z, 3), grade1.TextMatrix(z, 30), grade1.TextMatrix(z, 16), grade1.TextMatrix(z, 40), grade1.TextMatrix(z, 4), grade1.TextMatrix(z, 5), grade1.TextMatrix(z, 8), 0, grade1.TextMatrix(z, 11), grade1.TextMatrix(z, 10), FimdeProdutos)
                        FimdeProdutos = FimdeProdutos + 1
                    Else
                        Call Imprime_Calculo_ImpostoNovo("Continua Fl.Seguinte", "XXXXXXXXXX", "XXXXXXXXXX", "", "", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX", "XXXXXXXXXX")
                        Call MandaImpressoraNF
                        Call Imprime_Cabecalho_NFNovo(txt_numero_nf, "E", lNomeCodificacaoFiscal, lbl_cfop, txt_fornecedor, lbl_cgc, lbl_endereco, lbl_bairro, lbl_cep, lbl_cidade, lbl_telefone, lbl_uf, lbl_inscricao, lPessoa, msk_entrada, g_impressora_default_notas)
                        FimdeProdutos = 0
                        z = z - 1
                    End If
                End If
            Next
            Call Imprime_Calculo_ImpostoNovo(0, 0, 0, "", "", 0, lbl_bc_icms, lbl_valor_icms, lbl_bc_substituicao, lbl_icms_substituicao, lbl_total_produtos, lbl_frete, lbl_outras_despesas, lbl_ipi, lbl_total, lbl_seguro)
            Call Imprime_Transportadora_NFNovo(txt_transportadora_nome, cbo_frete.Text, txt_transportadora_placa, txt_transportadora_placa_uf, txt_transportadora_cnpj, txt_transportadora_endereco, txt_transportadora_cidade, txt_transportadora_uf, txt_transportadora_inscricao_estadual, txt_volume, txt_especie, txt_peso_bruto, txt_peso_liquido)
            Call Imprime_Dados_AdicionaisNovo("", "", lbl_forma, txt_observacoes, txt_numero_nf, "", "E", 0, txt_fornecedor, lbl_total)
            Call MandaImpressoraNF
        Else
            Alerta "Não existe dados da Configuração da NF!"
            Sql_Record_Set.Close
            Exit Sub
        End If
    End With
    Sql_Record_Set.Close

Exit Sub
file: ValidaErros Err, Me.Caption & " - Imprime_NFNovo"
End Sub
'*****************************************************************************
'Alteração: Nayden Luiz                                     Data: 29/12/2011
'         : OS18259-Corrigido Relatorio que nao estava mostrando corretamente
'         : os dados do codigo de produto.
'Propósito:
'*****************************************************************************
Private Sub ImpDadosEsp()
On Error GoTo ImpDados
Dim x_linha     As String
Dim xtotalqtde  As Currency
Dim z           As Long
    
    If g_tipo_impressora = 0 Then
        BioImprime "@@Printer.FontName = Draft 12cpi"
    ElseIf g_tipo_impressora = 1 Then
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontSize = 8"
    End If
    x_linha = "                                                                                                "
    Mid(x_linha, 3, 12) = "Fornecedor : "
    Mid(x_linha, 17, 60) = txt_codigo_fornecedor & " - " & txt_fornecedor.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                                                                                                "
    Mid(x_linha, 3, 12) = "Endereço   : "
    Mid(x_linha, 17, 110) = lbl_endereco & " - " & lbl_cidade & " - " & lbl_uf
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                                                                                                "
    Mid(x_linha, 3, 12) = "C.G.C.     : "
    Mid(x_linha, 17, 18) = Format(lbl_cgc, "@@.@@@.@@@/@@@@-@@")
    
    Mid(x_linha, 37, 25) = "Insc.Estadual : "
    Mid(x_linha, 59, 20) = lbl_inscricao
    
    Mid(x_linha, 79, 10) = "I.C.M.S : "
    Mid(x_linha, 89, 13) = Format(lbl_valor_icms, "##,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                                                                                                "
    Mid(x_linha, 3, 12) = "Numero N.F.: "
    Mid(x_linha, 17, 20) = txt_numero_nf
    
    Mid(x_linha, 38, 12) = "Série N.F.: "
    Mid(x_linha, 50, 2) = txt_serie_nf
    
    Mid(x_linha, 62, 12) = "Modelo NF : "
    Mid(x_linha, 76, 3) = txt_modelo_nf
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                                                                                                "
    Mid(x_linha, 3, 12) = "Emissão    : "
    Mid(x_linha, 17, 12) = msk_emissao
    
    Mid(x_linha, 38, 12) = "Entrada   : "
    Mid(x_linha, 50, 12) = msk_entrada
    
    Mid(x_linha, 62, 12) = "Tipo      : "
    Mid(x_linha, 76, 20) = txt_codigo_forma
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "  Observação :                                                                                  "
    Mid(x_linha, 15, 79) = Mid(txt_observacoes, 1, 78)
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "                                                                                                "
    Mid(x_linha, 15, 79) = Mid(txt_observacoes, 79, 78)
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    If lbl_nota <> "CX" Then
        lbl_nota = "NF"
    End If
                             
    Set gb_Recordset = Conexao.GeraRecordset("Select * From contas_apagar WHERE numero_nf = " & CLng(txt_numero_nf) & " and serie_nf = '" & lbl_nota & "' and codigo = '" & txt_codigo_fornecedor & "' Order By data_do_vencimento asc", 0)
    If gb_Recordset.RecordCount > 0 Then
        gb_Recordset.MoveFirst
            x_linha = "------------------------------------------------------------------------------------------------"
            BioImprime "@Printer.Print " & x_linha
            BioImprime1 x_linha
            x_linha = " DOC.        Vencimento:       Valor:            DOC.            Vencimento:         Valor:     "
            BioImprime "@Printer.Print " & x_linha
            BioImprime1 x_linha
            lLinha = lLinha + 2
            
            Do While Not gb_Recordset.EOF
                    x_linha = "                                                                                                "
                    Mid(x_linha, 2, 11) = gb_Recordset!numero_do_documento
                    Mid(x_linha, 13, 12) = gb_Recordset!data_do_vencimento
                    Mid(x_linha, 31, 13) = Format(gb_Recordset!Valor, "##,###,##0.00")
                    If Not gb_Recordset.EOF Then
                        gb_Recordset.MoveNext
                    End If
                    
                    If Not gb_Recordset.EOF Then
                        Mid(x_linha, 49, 12) = gb_Recordset!numero_do_documento
                        Mid(x_linha, 65, 12) = gb_Recordset!data_do_vencimento
                        Mid(x_linha, 85, 13) = Format(gb_Recordset!Valor, "##,###,##0.00")
                    End If
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime1 x_linha
                    lLinha = lLinha + 1
                    
                    If Not gb_Recordset.EOF Then
                        gb_Recordset.MoveNext
                    End If
            Loop
    End If
    gb_Recordset.Close
    
    x_linha = "                                                                                                                                                                  "
    If g_tipo_impressora = 0 Then
        BioImprime "@@Printer.FontName = Draft 17cpi"
    ElseIf g_tipo_impressora = 1 Then
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontSize = 7"
    End If
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                                                                                                                                        "
                           '         1         2         3         4         5         6         7         8         9        10        11        12        13         14
                           '123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
    
    Mid(x_linha, 1, 180) = "PED.   PRODUTO       DESCRIÇÃO                      CFOP UN.  QUANTIDADE   VALOR UNITARIO     SUBTOTAL R$  %ICMS   VLR.ICMS      P.CUSTO"
    'Mid(x_linha, 1, 180) = "PED.   PRODUTO    DESCRIÇÃO                      CFOP UN. QUANTIDADE  VL.BRUTO  %DESC  VLR. UNIT  SUBTOTAL R$  %ICMS  VLR.ICMS  P.CUSTO "
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
            
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
     For z = 1 To (grade1.Rows - 1)
        If Trim(grade1.TextMatrix(z, 2)) <> "" Then
                x_linha = "                                                                                                                                        "
                'NºPEDIDO
                f = Len(Mid(Format(grade1.TextMatrix(z, 1), "0000"), 1, 7))
                Mid(x_linha, 1 + 4 - f, f) = Mid(Format(grade1.TextMatrix(z, 1), "0000"), 1, 7)
                'CODPRODUTO
                If IsNumeric(grade1.TextMatrix(z, 2)) = True And g_sub_codigos = 0 Then
                    f = Len(Format(grade1.TextMatrix(z, 2), "000"))
                    Mid(x_linha, 8 + 10 - f, f) = Format(grade1.TextMatrix(z, 2), "000")
                Else
                    Mid(x_linha, 9, 13) = grade1.TextMatrix(z, 2)
                End If
            
                'CFOP
                Mid(x_linha, 53, 4) = grade1.TextMatrix(z, 28)
                
                'DESCRICAO
                Mid(x_linha, 22, 30) = grade1.TextMatrix(z, 3)
                
                'UN
                Mid(x_linha, 58, 2) = grade1.TextMatrix(z, 4)
                
                'QUANTIDADE
                f = Len(grade1.TextMatrix(z, 5))
                Mid(x_linha, 59 + 14 - f, f) = grade1.TextMatrix(z, 5)
                xtotalqtde = xtotalqtde + CCur(grade1.TextMatrix(z, 5))
                
                'VLR.BRUTO
                'f = Len(grade1.TextMatrix(z, 6))
                'Mid(x_linha, 71 + 9 - f, f) = grade1.TextMatrix(z, 6)
                
                'VLR.DESC
                'f = Len(grade1.TextMatrix(z, 7))
                'Mid(x_linha, 80 + 5 - f, f) = Format(grade1.TextMatrix(z, 7), "##,###,##0.00")
                
                'VLR.UNITARIO
                f = Len(grade1.TextMatrix(z, 8))
                Mid(x_linha, 74 + 16 - f, f) = grade1.TextMatrix(z, 8)
                
                'SUBTOTAL
                f = Len(grade1.TextMatrix(z, 11))
                Mid(x_linha, 91 + 15 - f, f) = grade1.TextMatrix(z, 11)
                   
                'ICMS
                f = Len(Format(grade1.TextMatrix(z, 10), "##,###,##0.00"))
                Mid(x_linha, 107 + 6 - f, f) = Format(grade1.TextMatrix(z, 10), "##,###,##0.00")
                
                'VLR.ICMS
                f = Len(Format(grade1.TextMatrix(z, 15), "##,###,##0.00"))
                Mid(x_linha, 114 + 10 - f, f) = Format(grade1.TextMatrix(z, 15), "##,###,##0.00")
                
                'preco custo
                f = Len(Format(fValidaValorNovo(grade1.TextMatrix(z, 18)), g_decimal_custo))
                Mid(x_linha, 125 + 12 - f, f) = Format(fValidaValorNovo(grade1.TextMatrix(z, 18)), g_decimal_custo)
                
                BioImprime "@Printer.Print " & x_linha
                BioImprime1 x_linha
        End If
     Next
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "  BASE CALCULO ICMS R$  VALOR DO I.C.M.S. R$    B.C.SUBSTIT. ICMS R$   VALOR SUBST. ICMS R$     VALOR DO IPI R$      TOTAL PRODUTOS     "
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                       |                       |                      |                      |                   |                      "
    Mid(x_linha, 23 - Len(lbl_bc_icms), 15) = lbl_bc_icms
    Mid(x_linha, 46 - Len(lbl_valor_icms), 15) = lbl_valor_icms
    Mid(x_linha, 70 - Len(lbl_bc_substituicao), 15) = lbl_bc_substituicao
    Mid(x_linha, 93 - Len(lbl_icms_substituicao), 15) = lbl_icms_substituicao
    Mid(x_linha, 113 - Len(lbl_ipi), 15) = lbl_ipi
    Mid(x_linha, 136 - Len(lbl_total_produtos), 15) = lbl_total_produtos
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
              '         1         2         3         4         5         6         7         8         9        10        11        12        13         14
              '123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
    x_linha = "  VALOR DO FRETE R$    |  VALOR DO SEGURO R$   |  OUTRAS DESPESAS  R$ | VALOR DO DESCONTO R$ |     QUANTIDADE    |  TOTAL DA N.F. R$    "
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "                       |                       |                      |                      |                   |                      "
    Mid(x_linha, 23 - Len(lbl_frete), 15) = lbl_frete
    Mid(x_linha, 46 - Len(lbl_seguro), 15) = lbl_seguro
    Mid(x_linha, 70 - Len(lbl_outras_despesas), 15) = lbl_outras_despesas
    Mid(x_linha, 93 - Len(txt_desconto), 15) = txt_desconto
    Mid(x_linha, 113 - Len(Format(xtotalqtde, g_decimal_estoque)), 15) = Format(xtotalqtde, g_decimal_estoque)
    Mid(x_linha, 136 - Len(lbl_total), 15) = lbl_total
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    x_linha = "----------------------------------------------------------------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha

Exit Sub
ImpDados: ValidaErros Err, Me.Caption & " - ImpDados"
End Sub

Private Sub ImpCabEsp()
On Error GoTo Err_ImpCabEsp
Dim x_linha As String
Dim i As Integer
    
    If lPagina = 0 Then
        BioCriaImprime
        BioCriaImprime1
        'seleciona medidas para centímetros
        If g_tipo_impressora = 0 Then
            BioImprime "@@Printer.ScaleMode = 7"
            BioImprime "@@Printer.PaperSize = 1"
            BioImprime "@@Printer.FontName = Draft 10cpi"
            'teste para imprimir letra correta
            BioImprime "@@Printer.FontBold = False"
            BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
        'impressora jato tinta
        ElseIf g_tipo_impressora = 1 Then
            BioImprime "@@Printer.ScaleMode = 7"
            BioImprime "@@Printer.PaperSize = 9"
        End If
    End If
    lPagina = lPagina + 1
    lLinha = 0
    If g_tipo_impressora = 0 Then
        BioImprime "@@Printer.FontName = Draft 10cpi"
    ElseIf g_tipo_impressora = 1 Then
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontSize = 8"
    End If

    BioImprime "@@Printer.CurrentY = 0"
    
    BioImprime "@@Printer.FontBold = True"
    x_linha = "                                                                   PAGINA: ___  "
    Mid(x_linha, 3, 40) = cDPEmpresa.NOME
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    
    BioImprime "@@Printer.FontBold = False"
    x_linha = "  ENTRADA DE MERCARCADORIA - N.F. DE ENTRADA                CIDADE, __/__/____  "
    i = Len(Trim(cDPEmpresa.NomeCidade))
    Mid(x_linha, 37 + 30 - i, i) = Trim(cDPEmpresa.NomeCidade)
    Mid(x_linha, 69, 10) = Date
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha
    x_linha = "--------------------------------------------------------------------------------"
    BioImprime "@Printer.Print " & x_linha
    BioImprime1 x_linha

Exit Sub
Err_ImpCabEsp: ValidaErros Err, Me.Caption & " - ImpCabEsp"
End Sub

'==========================================================================
' Purpose:  Abate a quantidade do pedido de mercadoria
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 23/03/2013
'==========================================================================
Private Function AbateQuantidadeItemPedidoCompra(ByVal curQuantidade As Currency, _
                                                ByVal lngItemMov As Long, _
                                                ByVal lngNumeroPedido As Long) As Boolean
On Error GoTo Err_AbateQuantidadeItemPedidoCompra
    
    AbateQuantidadeItemPedidoCompra = False
    'reabre o item do pedido
    If Not Conexao.AlterarRecordset("tb_item_pedidocompra", _
                                "quantidade_recebida = quantidade_recebida - " & fValidaValor2(curQuantidade) & ",situacao_item = '" & eTipoSituacaoItemPedido.Aberto & "'", _
                                "fk_empresa = '" & cDPEmpresa.codigo & "' and " & _
                                "fk_item_movimentacao = '" & lngItemMov & "' ", _
                                cDPEmpresa.codigo) Then Exit Function
    
    'reabre o cabecalho do pedido
    If Not Conexao.AlterarRecordset("tb_pedido_compra", _
                                "situacao = '" & eTipoSituacaoPedido.Aberto & "'", _
                                "fk_empresa = '" & cDPEmpresa.codigo & "' and " & _
                                "numeropedido = '" & lngNumeroPedido & "' ", _
                                cDPEmpresa.codigo) Then Exit Function
    
    AbateQuantidadeItemPedidoCompra = True
                                    
Exit Function
Err_AbateQuantidadeItemPedidoCompra: ValidaErros Err, Me.Caption & " - AbateQuantidadeItemPedidoCompra"
End Function


'==========================================================================
' Purpose:  Validar a alteração da nota fiscal
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Function ValidaAlteraNF() As Boolean
On Error GoTo Err_ValidaAlteraNF

    ValidaAlteraNF = False
    Senha.lbl_serie = "NF"
    Senha.lbl_titulo = "Liberação para Alteração de Nota Entrada"
    Senha.lbl_liberacao = "Entrada (Alteração)"
    Senha.lbl_tela = "Entrada de Mercadorias"
    Senha.lbl_historico = "Liberação para Alteração da Nota N.: " & txt_numero_nf & " Forn. " & txt_fornecedor
    Senha.Show (1)
    If g_string = "OK" Then ValidaAlteraNF = True

Exit Function
Err_ValidaAlteraNF: ValidaErros Err, Me.Caption & " - ValidaAlteraNF"
End Function



'==========================================================================
' Purpose:  Preenche a coleção de produtos para efetuar a validação na alteração da nota
'           veerificando se a quantidade foi alterada
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Function PreencheColecaoProdutos() As Boolean
On Error GoTo Err_PreencheColecaoProdutos
Dim z As Long

    Set colProdutos = Nothing
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 0)) Then
            Dim DPItemMov As New cDPItemMovimentacao
            
            DPItemMov.cDPProduto.codigo = grade1.TextMatrix(z, 2)
            DPItemMov.cDPProduto.Descricao = grade1.TextMatrix(z, 3)
            DPItemMov.Quantidade = grade1.TextMatrix(z, 5)
            
            Call colProdutos.Add(DPItemMov, DPItemMov.cDPProduto.codigo)
            Set DPItemMov = Nothing
            
        End If
    Next

Exit Function
Err_PreencheColecaoProdutos: ValidaErros Err, Me.Caption & " - PreencheColecaoProdutos"
End Function


'==========================================================================
' Purpose: Verifica se houve alteração na quantidade dos produtos na alteração da nota
'           pois se houve altera os estoques caso contrário não entra nos métdos de alteração de
'           estoques, lotes, depositos etc..
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Function VerificaAlteracaoQuantidadeProdutos() As Boolean
On Error GoTo Err_VerificaAlteracaoQuantidadeProdutos
Dim z           As Long
Dim DPItemMov   As New cDPItemMovimentacao
    
    VerificaAlteracaoQuantidadeProdutos = False
    
    bolAlteraEstoquenaAlteracaoNF = False
    For z = 0 To grade1.Rows - 1
        If IsNumeric(grade1.TextMatrix(z, 0)) Then
            'localiza o produto na coleção e verifica se houve alteracao na quantidade
            If cUtGeral.KeyExistis(colProdutos, grade1.TextMatrix(z, 2)) Then
                Set DPItemMov = colProdutos(grade1.TextMatrix(z, 2))
                If CDbl(DPItemMov.Quantidade) <> CDbl(grade1.TextMatrix(z, 5)) Then
                    bolAlteraEstoquenaAlteracaoNF = True
                    Exit For
                End If
            Else
                'se não encontrou e porque pode estar adicionando um novo produto
                'então altera o estoque tambem
                bolAlteraEstoquenaAlteracaoNF = True
                Exit For
            End If
        End If
    Next
    Set DPItemMov = Nothing
    
    'não permite alteracao de quantidade para produtos conferidos
    If bolConferencia And bolAlteraEstoquenaAlteracaoNF Then
        Alerta "Nota não pode sofrer alteração na quantidade por ter passado por " & _
               "procedimento de conferência dos produtos" & Chr(13) & "Efetue a exclusão e refaça o lançamento"
               VerificaAlteracaoQuantidadeProdutos = False
    Else
        VerificaAlteracaoQuantidadeProdutos = True
    End If
    
Exit Function
Err_VerificaAlteracaoQuantidadeProdutos: ValidaErros Err, Me.Caption & " - VerificaAlteracaoQuantidadeProdutos"
End Function


'==========================================================================
' Purpose:  Verificar se esta e a ultima nota que foi dada entrada, sendo positivo
'           poderá alterar o custo
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 06/06/2013
'==========================================================================
Private Function ValidaLiberacaoAlteracaoCusto() As Boolean
On Error GoTo Err_ValidaLiberacaoAlteracaoCusto
Dim strSQL As String
    
    ValidaLiberacaoAlteracaoCusto = False
    bolLiberaAlteracaoCustoAlteracaoNF = False
    strSQL = ""
    strSQL = strSQL & " SELECT MCN.data_de_entrada,MCN.sequencia"
    strSQL = strSQL & " FROM movimento_nota_fiscal_entrada MN,movimento_cabecalho_nota_fiscal_entrada MCN "
    strSQL = strSQL & " WHERE MCN.empresa = '" & cDPEmpresa.codigo & "' and MCN.atualiza_custo > 0 and "
    strSQL = strSQL & " MN.empresa = MCN.empresa and MN.sequencia = MCN.sequencia and MN.outros = '" & lbl_nota & "' "
    strSQL = strSQL & " ORDER BY data_de_entrada DESC,sequencia DESC"
    Set gb_Recordset = Conexao.GeraRecordset(strSQL, 1)
    If gb_Recordset.RecordCount > 0 Then
        If lSequencia = gb_Recordset!sequencia Then
            bolLiberaAlteracaoCustoAlteracaoNF = True
        Else
            Alerta "Esta não e a última nota lançada custo/preço não será atualizado!"
        End If
    Else
        bolLiberaAlteracaoCustoAlteracaoNF = True
    End If
    gb_Recordset.Close
    
    ValidaLiberacaoAlteracaoCusto = True
    
Exit Function
Err_ValidaLiberacaoAlteracaoCusto: ValidaErros Err, Me.Caption & " - ValidaLiberacaoAlteracaoCusto"
End Function


'==========================================================================
' Purpose:  Validar a função bolconferencia principalmente quando se trata de alteracao NF
' Input:
' Output:
' Remarks:
' Author: Ronaldo Robledo                                Start: 07/06/2013
'==========================================================================
Private Function ValidaBooleanoConferencia() As Boolean
On Error GoTo Err_ValidaBooleanoConferencia
    If l_opcao = 1 Then
        ValidaBooleanoConferencia = bolConferencia
    Else
        ValidaBooleanoConferencia = False
    End If
Exit Function
Err_ValidaBooleanoConferencia: ValidaErros Err, Me.Caption & " - ValidaBooleanoConferencia"
End Function
