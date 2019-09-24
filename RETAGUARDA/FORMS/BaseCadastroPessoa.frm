VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBaseCadastroPessoa 
   Caption         =   "Cadastro"
   ClientHeight    =   6600
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BaseCadastroPessoa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10200
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":90A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":A7A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":B8AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":C860
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5895
      Left            =   0
      TabIndex        =   43
      Top             =   720
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1- Dados Pessoais"
      TabPicture(0)   =   "BaseCadastroPessoa.frx":DA92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraPessoa"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtOBS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFone"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2- Endereço"
      TabPicture(1)   =   "BaseCadastroPessoa.frx":DAAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraComercial"
      Tab(1).Control(1)=   "fraResidencial"
      Tab(1).Control(2)=   "fraCobranca"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&3- Parâmetros"
      TabPicture(2)   =   "BaseCadastroPessoa.frx":DACA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtInscSuframa"
      Tab(2).Control(1)=   "chkSuframa"
      Tab(2).Control(2)=   "cmdAtGlobal"
      Tab(2).Control(3)=   "txtCredito"
      Tab(2).Control(4)=   "txtIM"
      Tab(2).Control(5)=   "chkImp"
      Tab(2).Control(6)=   "cmbVendedorAUX"
      Tab(2).Control(7)=   "chkESTRANGEIRO"
      Tab(2).Control(8)=   "cmbTipoCli"
      Tab(2).Control(9)=   "txtContato"
      Tab(2).Control(10)=   "cmbVendedor"
      Tab(2).Control(11)=   "txtPercConv"
      Tab(2).Control(12)=   "cmbAuxRegiao"
      Tab(2).Control(13)=   "cmbRegiao"
      Tab(2).Control(14)=   "lblSuframa"
      Tab(2).Control(15)=   "lblCredito"
      Tab(2).Control(16)=   "lblIM"
      Tab(2).Control(17)=   "lblTipoCli"
      Tab(2).Control(18)=   "lblContato"
      Tab(2).Control(19)=   "lblVendedor"
      Tab(2).Control(20)=   "lblConvenio"
      Tab(2).Control(21)=   "lblRegiao"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "&4 - Histórico"
      TabPicture(3)   =   "BaseCadastroPessoa.frx":DAE6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line1"
      Tab(3).Control(1)=   "lstCompras"
      Tab(3).Control(2)=   "staCompras"
      Tab(3).Control(3)=   "staAberto"
      Tab(3).Control(4)=   "lstAberto"
      Tab(3).Control(5)=   "txtTotalVendas"
      Tab(3).Control(6)=   "txtSaldoDevedor"
      Tab(3).ControlCount=   7
      Begin VB.TextBox txtInscSuframa 
         DataField       =   "Nome"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -71160
         MaxLength       =   100
         TabIndex        =   38
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chkSuframa 
         Caption         =   "&Suframa:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Frame fraFone 
         Caption         =   "Telefone(s)"
         Height          =   3255
         Left            =   120
         TabIndex        =   105
         Top             =   2400
         Width           =   6255
         Begin VB.CommandButton cmdExcluirFone 
            Height          =   375
            Left            =   5640
            Picture         =   "BaseCadastroPessoa.frx":DB02
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   240
            Width           =   450
         End
         Begin VB.TextBox txtL 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3360
            MaxLength       =   30
            TabIndex        =   8
            Top             =   240
            Width           =   2265
         End
         Begin VB.TextBox txtN 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtDDD 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   600
            MaxLength       =   2
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin MSDataGridLib.DataGrid Grid 
            Bindings        =   "BaseCadastroPessoa.frx":E943
            Height          =   2415
            Left            =   45
            TabIndex        =   106
            Top             =   720
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   18
            WrapCellPointer =   -1  'True
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "ddd"
               Caption         =   "DDD"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "numero"
               Caption         =   "Telefone"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "local"
               Caption         =   "Local"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Alignment       =   2
                  ColumnWidth     =   929,764
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnWidth     =   1844,787
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   5820,095
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc adoFone 
            Height          =   330
            Left            =   0
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
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
            Caption         =   "Grid Cabeça"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Local:"
            Height          =   240
            Index           =   13
            Left            =   2730
            TabIndex        =   109
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "DDD:"
            Height          =   240
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.TextBox txtOBS 
         DataField       =   "Endereco_Res"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3390
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   104
         Top             =   2400
         Width           =   4545
      End
      Begin VB.CommandButton cmdAtGlobal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Atualizar cadastro do cliente no banco NFe?"
         Height          =   615
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
         Top             =   3360
         Width           =   3105
      End
      Begin VB.TextBox txtCredito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -65160
         MaxLength       =   50
         TabIndex        =   36
         Text            =   "00,00"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtIM 
         Height          =   360
         Left            =   -69360
         MaxLength       =   25
         TabIndex        =   31
         Top             =   600
         Width           =   2085
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   -65400
         TabIndex        =   97
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbVendedorAUX 
         BackColor       =   &H80000000&
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
         Left            =   -73200
         TabIndex        =   96
         Top             =   1125
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkESTRANGEIRO 
         Caption         =   "Estrangeiro"
         Height          =   240
         Left            =   -65400
         TabIndex        =   91
         Top             =   1125
         Width           =   1455
      End
      Begin VB.ComboBox cmbTipoCli 
         Height          =   360
         Left            =   -69360
         TabIndex        =   35
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtContato 
         DataField       =   "Nome"
         Height          =   375
         Left            =   -73320
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1560
         Width           =   2055
      End
      Begin VB.ComboBox cmbVendedor 
         Height          =   360
         Left            =   -73320
         TabIndex        =   32
         Top             =   1125
         Width           =   2055
      End
      Begin VB.TextBox txtPercConv 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   -69360
         MaxLength       =   5
         TabIndex        =   33
         Text            =   "00,00"
         Top             =   1125
         Width           =   1095
      End
      Begin VB.ComboBox cmbAuxRegiao 
         BackColor       =   &H80000000&
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
         Left            =   -73200
         TabIndex        =   90
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbRegiao 
         Height          =   360
         Left            =   -73320
         TabIndex        =   30
         Top             =   600
         Width           =   2055
      End
      Begin VB.Frame fraComercial 
         Caption         =   " Endereço Comercial "
         ForeColor       =   &H000000FF&
         Height          =   1755
         Left            =   -74895
         TabIndex        =   78
         Top             =   2280
         Width           =   10995
         Begin VB.TextBox txtIBGEc 
            DataField       =   "Bairro_Res"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9000
            LinkTimeout     =   7
            MaxLength       =   50
            TabIndex        =   40
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtNumeroC 
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
            Left            =   7920
            MaxLength       =   50
            TabIndex        =   18
            Top             =   570
            Width           =   855
         End
         Begin VB.TextBox txtEndC 
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
            Left            =   9000
            MaxLength       =   50
            TabIndex        =   19
            Top             =   570
            Width           =   1905
         End
         Begin VB.TextBox txtUFC 
            Alignment       =   2  'Center
            DataField       =   "Estado_Com"
            DataSource      =   "Data1"
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
            Left            =   7920
            MaxLength       =   2
            TabIndex        =   22
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtCidadeC 
            DataField       =   "Cidade_Com"
            DataSource      =   "Data1"
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
            Left            =   2940
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1200
            Width           =   4875
         End
         Begin VB.TextBox txtBairroC 
            DataField       =   "Bairro_Com"
            DataSource      =   "Data1"
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   20
            Top             =   1200
            Width           =   2745
         End
         Begin VB.TextBox txtRuaC 
            DataField       =   "Endereco_Com"
            DataSource      =   "Data1"
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
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   17
            Top             =   570
            Width           =   5955
         End
         Begin VB.CommandButton CmdCopiaEnderecoPessoal2 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10440
            Picture         =   "BaseCadastroPessoa.frx":E959
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Clique aqui para copiar o endereço"
            Top             =   120
            Width           =   465
         End
         Begin MSMask.MaskEdBox txtCepC 
            Height          =   345
            Left            =   120
            TabIndex        =   16
            Top             =   570
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####-###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "*IBGE:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   9000
            TabIndex        =   87
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7920
            TabIndex        =   86
            Top             =   330
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "*Comp."
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9030
            TabIndex        =   85
            Top             =   330
            Width           =   690
         End
         Begin VB.Label lblEstadoCom 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7920
            TabIndex        =   84
            Top             =   960
            Width           =   315
         End
         Begin VB.Label lblCepCom 
            AutoSize        =   -1  'True
            Caption         =   "*Cep:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   83
            Top             =   330
            Width           =   510
         End
         Begin VB.Label lblCidadeCom 
            AutoSize        =   -1  'True
            Caption         =   "*Cidade:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2940
            TabIndex        =   82
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblBairroCom 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   81
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lblRuaCom 
            AutoSize        =   -1  'True
            Caption         =   "*Rua:"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1890
            TabIndex        =   80
            Top             =   330
            Width           =   510
         End
      End
      Begin VB.Frame fraResidencial 
         Caption         =   " Endereço Residencial"
         ForeColor       =   &H000000FF&
         Height          =   1665
         Left            =   -74895
         TabIndex        =   69
         Top             =   480
         Width           =   10995
         Begin VB.TextBox txtNumeroR 
            Height          =   375
            Left            =   7920
            MaxLength       =   50
            TabIndex        =   11
            Top             =   570
            Width           =   855
         End
         Begin VB.TextBox txtEndR 
            Height          =   375
            Left            =   9000
            MaxLength       =   50
            TabIndex        =   12
            Top             =   570
            Width           =   1935
         End
         Begin VB.TextBox txtUFR 
            Alignment       =   2  'Center
            DataField       =   "Estado"
            Height          =   375
            Left            =   7920
            MaxLength       =   2
            TabIndex        =   15
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtCidadeR 
            DataField       =   "Cidade"
            Height          =   360
            Left            =   2940
            MaxLength       =   50
            TabIndex        =   14
            Top             =   1200
            Width           =   4875
         End
         Begin VB.TextBox txtBairroR 
            DataField       =   "Bairro_Res"
            DataSource      =   "Data1"
            Height          =   360
            Left            =   90
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1200
            Width           =   2745
         End
         Begin VB.TextBox txtRuaR 
            DataField       =   "Endereco_Res"
            DataSource      =   "Data1"
            Height          =   360
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   10
            Top             =   570
            Width           =   5955
         End
         Begin VB.TextBox txtIBGE 
            DataField       =   "Bairro_Res"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9000
            LinkTimeout     =   7
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1200
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txtCepR 
            Height          =   360
            Left            =   90
            TabIndex        =   9
            Top             =   570
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####-###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7920
            TabIndex        =   77
            Top             =   330
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "*UF:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7920
            TabIndex        =   76
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "*Complemento:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9000
            TabIndex        =   75
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label lblCep 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "*Cep:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   30
            TabIndex        =   74
            Top             =   330
            Width           =   525
         End
         Begin VB.Label lblCidade 
            AutoSize        =   -1  'True
            Caption         =   "*Cidade:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2970
            TabIndex        =   73
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblBairro 
            AutoSize        =   -1  'True
            Caption         =   "*Bairro:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   720
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            Caption         =   "*Rua:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1890
            TabIndex        =   71
            Top             =   330
            Width           =   510
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "*IBGE:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   9000
            TabIndex        =   70
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame FraPessoa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1995
         Left            =   105
         TabIndex        =   55
         Top             =   360
         Width           =   10995
         Begin VB.CommandButton cmdFoto 
            Caption         =   "&Foto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7200
            TabIndex        =   112
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtUFValida 
            Alignment       =   2  'Center
            DataField       =   "Estado"
            Height          =   375
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   111
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdEmail 
            Caption         =   "&Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   103
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdRg 
            Caption         =   "RG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5520
            Picture         =   "BaseCadastroPessoa.frx":EA2B
            TabIndex        =   102
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtIE 
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
            Left            =   2520
            MaxLength       =   25
            TabIndex        =   4
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtNome 
            DataField       =   "Nome"
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
            Left            =   2520
            MaxLength       =   100
            TabIndex        =   1
            Top             =   480
            Width           =   5535
         End
         Begin VB.ComboBox cmbStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9480
            TabIndex        =   5
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00FFFFFF&
            Height          =   350
            Left            =   2080
            Picture         =   "BaseCadastroPessoa.frx":FB71
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox txtRazao 
            DataField       =   "Nome"
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
            Left            =   2520
            MaxLength       =   100
            TabIndex        =   2
            Top             =   960
            Width           =   5535
         End
         Begin MSMask.MaskEdBox txtCNPJCPF 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   480
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
         Begin MSMask.MaskEdBox txtDtNasc 
            Height          =   375
            Left            =   9480
            TabIndex        =   3
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDtCad 
            Height          =   375
            Left            =   9480
            TabIndex        =   57
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPessoaID 
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   101
            Top             =   960
            Width           =   540
         End
         Begin VB.Label lblIE 
            Alignment       =   1  'Right Justify
            Caption         =   "Inscrição Estadual:"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   630
            TabIndex        =   64
            Top             =   1440
            Width           =   1785
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Dt.Cadastro:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   8220
            TabIndex        =   63
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label lblDtNasc 
            Alignment       =   1  'Right Justify
            Caption         =   "Dt.Nascim.:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   8310
            TabIndex        =   62
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Situação:"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   8520
            TabIndex        =   61
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label lblNomeCli 
            Caption         =   "Nome:"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2640
            TabIndex        =   60
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblCNPJCPF 
            Alignment       =   1  'Right Justify
            Caption         =   "CNPJ/CPF:"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   195
            TabIndex        =   59
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label lblRazao 
            Alignment       =   1  'Right Justify
            Caption         =   "Razão Social:"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1095
            TabIndex        =   58
            Top             =   960
            Width           =   1320
         End
      End
      Begin VB.Frame fraCobranca 
         Caption         =   " Endereço Cobrança "
         ForeColor       =   &H00400000&
         Height          =   1695
         Left            =   -74895
         TabIndex        =   46
         Top             =   4080
         Width           =   10995
         Begin VB.TextBox txtIBGEb 
            DataField       =   "Bairro_Res"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9000
            LinkTimeout     =   7
            MaxLength       =   50
            TabIndex        =   41
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton CmdCopiaEnderecoPessoal1 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10440
            Picture         =   "BaseCadastroPessoa.frx":10573
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Clique aqui para copiar o endereço"
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtEndB 
            Height          =   360
            Left            =   9000
            MaxLength       =   50
            TabIndex        =   26
            Top             =   570
            Width           =   1905
         End
         Begin VB.TextBox txtUFB 
            Alignment       =   2  'Center
            DataField       =   "Estado"
            Height          =   360
            Left            =   7920
            MaxLength       =   2
            TabIndex        =   29
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtCidadeB 
            DataField       =   "Cidade"
            Height          =   360
            Left            =   2940
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1200
            Width           =   4875
         End
         Begin VB.TextBox txtBaIrroB 
            DataField       =   "Bairro_Res"
            DataSource      =   "Data1"
            Height          =   360
            Left            =   120
            MaxLength       =   50
            TabIndex        =   27
            Top             =   1200
            Width           =   2745
         End
         Begin VB.TextBox txtRuaB 
            DataField       =   "Endereco_Res"
            DataSource      =   "Data1"
            Height          =   360
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   24
            Top             =   570
            Width           =   5955
         End
         Begin VB.TextBox txtNumeroB 
            Height          =   360
            Left            =   7920
            MaxLength       =   50
            TabIndex        =   25
            Top             =   570
            Width           =   855
         End
         Begin MSMask.MaskEdBox txtCepB 
            Height          =   345
            Left            =   120
            TabIndex        =   23
            Top             =   570
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####-###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "*IBGE:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   9000
            TabIndex        =   88
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   240
            Left            =   7920
            TabIndex        =   54
            Top             =   960
            Width           =   315
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "*Comp."
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9030
            TabIndex        =   53
            Top             =   330
            Width           =   690
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "*Cep:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   60
            TabIndex        =   52
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "*Cidade:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2940
            TabIndex        =   51
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "*Rua:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1860
            TabIndex        =   49
            Top             =   330
            Width           =   510
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7920
            TabIndex        =   48
            Top             =   330
            Width           =   810
         End
      End
      Begin VB.TextBox txtSaldoDevedor 
         Alignment       =   1  'Right Justify
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   -66030
         MaxLength       =   50
         TabIndex        =   45
         Top             =   2445
         Width           =   2145
      End
      Begin VB.TextBox txtTotalVendas 
         Alignment       =   1  'Right Justify
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   -66060
         MaxLength       =   50
         TabIndex        =   44
         Top             =   5220
         Width           =   2175
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   7920
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":10645
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":10A99
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":10DB5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":11209
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":1165D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":1197D
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":11DD1
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":120F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BaseCadastroPessoa.frx":12545
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstAberto 
         Height          =   1905
         Left            =   -74895
         TabIndex        =   65
         Top             =   480
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   3360
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Modalidade"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Dt.Venc."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Dt.Baixa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OBS."
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.StatusBar staAberto 
         Height          =   405
         Left            =   -74880
         TabIndex        =   66
         Top             =   2400
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   714
         Style           =   1
         SimpleText      =   "                                                 Títulos em Aberto                                           Saldo Devedor = "
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
      Begin MSComctlLib.StatusBar staCompras 
         Height          =   405
         Left            =   -74880
         TabIndex        =   67
         Top             =   5160
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   714
         Style           =   1
         SimpleText      =   "                                                 Últimas Compras                                               Total Vendas = "
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
      Begin MSComctlLib.ListView lstCompras 
         Height          =   1905
         Left            =   -74895
         TabIndex        =   68
         Top             =   3120
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Req./Orç."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor Venda"
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor Desc."
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Dt.Emis."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Tipo Venda"
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vendedor"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblSuframa 
         Alignment       =   1  'Right Justify
         Caption         =   "Inscrição Suframa:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -73050
         TabIndex        =   110
         Top             =   2160
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -63840
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label lblCredito 
         AutoSize        =   -1  'True
         Caption         =   "Limite Crédito:"
         Height          =   240
         Left            =   -66600
         TabIndex        =   99
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label lblIM 
         Alignment       =   1  'Right Justify
         Caption         =   "Insc.Municipal:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -70890
         TabIndex        =   98
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label lblTipoCli 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Cliente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -70710
         TabIndex        =   95
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblContato 
         Alignment       =   1  'Right Justify
         Caption         =   "Contato:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74220
         TabIndex        =   94
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblVendedor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74415
         TabIndex        =   93
         Top             =   1125
         Width           =   990
      End
      Begin VB.Label lblConvenio 
         Alignment       =   1  'Right Justify
         Caption         =   "% Convenio:"
         Height          =   240
         Left            =   -70620
         TabIndex        =   92
         Top             =   1125
         Width           =   1170
      End
      Begin VB.Label lblRegiao 
         Alignment       =   1  'Right Justify
         Caption         =   "Região:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74160
         TabIndex        =   89
         Top             =   600
         Width           =   735
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
      DesignWidth     =   11250
      DesignHeight    =   6600
   End
End
Attribute VB_Name = "frmBaseCadastroPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim ID_ENTIDADE_N As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
   PREPARA_TELA_CADASTRO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imprimir"
      Case "voltar"
         Unload Me
      Case "gravar"
         If Trim(cmbStatus.Text) = "" Then _
            cmbStatus.Text = "Ativo"
         txtCNPJCPF.PromptInclude = False
         If GRAVA_DADOS_PESSOA(Trim(txtCNPJCPF.Text), Trim(txtNome.Text), Trim(txtRazao.Text), Trim(Left(cmbStatus.Text, 1))) = True Then
            GRAVA_OBS

            'ENDEREÇO RESIDENCIAL
               txtCepR.PromptInclude = False
               If Not IsNumeric(txtIBGE.Text) Then _
                  txtIBGE.Text = "5201211"

               If Trim(txtCepR.Text) <> "" Or Trim(txtRuaR.Text) <> "" Or Trim(txtBairroR.Text) <> "" Or Trim(txtEndR.Text) <> "" Then
                  sp_Grava_Endereco Trim(txtCepR.Text), Trim(txtRuaR.Text), Trim(txtBairroR.Text), Trim(txtEndR.Text), "R", Trim(txtNumeroR.Text)
                  Else: SP_MATA_ENDEREÇO "R"
               End If

            'ENDEREÇO COMERCIAL
               If Not IsNumeric(txtIBGEc.Text) Then _
                  txtIBGEc.Text = "5201211"

               txtCepC.PromptInclude = False
               If Trim(txtCepC.Text) <> "" Or Trim(txtRuaC.Text) <> "" Or Trim(txtBairroC.Text) <> "" Or Trim(txtEndC.Text) <> "" Then
                  sp_Grava_Endereco Trim(txtCepC.Text), Trim(txtRuaC.Text), Trim(txtBairroC.Text), Trim(txtEndC.Text), "C", Trim(txtNumeroC.Text)
                  Else: SP_MATA_ENDEREÇO "C"
               End If

            'ENDEREÇO COBRANÇA
               If Not IsNumeric(txtIBGEb.Text) Then _
                  txtIBGEb.Text = "5201211"

               txtCepB.PromptInclude = False
               If Trim(txtCepB.Text) <> "" Or Trim(txtRuaB.Text) <> "" Or Trim(txtBaIrroB.Text) <> "" Or Trim(txtEndB.Text) <> "" Then
                  sp_Grava_Endereco Trim(txtCepB.Text), Trim(txtRuaB.Text), Trim(txtBaIrroB.Text), Trim(txtEndB.Text), "B", Trim(txtNumeroB.Text)
                  Else: SP_MATA_ENDEREÇO "B"
               End If

            If Trim(txtIE.Text) <> "ISENTO" Then
               If IsNumeric(txtIE.Text) Then
                 If Valida_Inscricao_Estadual(txtIE.Text, txtUFValida.Text) <> 0 Then
                    SSTab.Tab = 0
                    txtIE.SetFocus
                    Exit Sub
                 End If
               End If
               GRAVA_IE Trim(txtIE.Text)
            End If

            If Trim(TIPO_PESSOA_CADASTRO) = "FORNECEDOR" Then _
               If GRAVA_DADOS_FORNECEDOR = True Then _
                  LIMPA_TUDO
            If Trim(TIPO_PESSOA_CADASTRO) = "TRANSPORTADORA" Then _
               If GRAVA_DADOS_TRANSPORTADORA = True Then _
                  LIMPA_TUDO
         End If
         txtCNPJCPF.SetFocus
      Case "matar"
      Case "limpar"
         LIMPA_TUDO
         txtCNPJCPF.SetFocus
      Case "consultar"
         CNPJCPF_A = ""
         PESSOA_ID_N = 0
         'TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmConsultaPessoa.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
            txtCNPJCPF.PromptInclude = True
            Call TXTCNPJCPF_LostFocus
            txtCNPJCPF.SetFocus
         End If
         CNPJCPF_A = ""
         PESSOA_ID_N = 0
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   PESSOA_ID_N = 0
   'TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmConsultaPessoa.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      txtCNPJCPF.PromptInclude = True
      Call TXTCNPJCPF_LostFocus
      txtCNPJCPF.SetFocus
   End If
   CNPJCPF_A = ""
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdConsulta_Click"
End Sub

Private Sub cmdRG_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      CNPJCPF_A = Trim(txtCNPJCPF.Text)
      frmCADASTRORG.Show 1
   End If
   txtCNPJCPF.PromptInclude = True
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdEmail_Click"
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   If SSTab.Tab = 0 Then _
      txtCNPJCPF.SetFocus

   If SSTab.Tab = 1 Then
      If txtCepR.Visible = True Then
         txtCepR.SetFocus
         Else: txtCepC.SetFocus
      End If
   End If

   'If ssTab.Tab = 2 Then _
      txtCepC.SetFocus

   'If ssTab.Tab = 3 Then
   '   fraTel.Visible = False
   '   FlexTel.Visible = False

   '   txtCNPJCPF.PromptInclude = False
   '   If txtCNPJCPF.Text <> "" Then
   '      CONSULTA_VENDAS_CLIENTE
   '      CONSULTA_LANÇAMENTOS
   '      MOSTRA_CONTAS_CORRENTE
   '   End If
   '   txtCNPJCPF.PromptInclude = True

   '   VALOR_TOTAL_N = 0
   '   Else
   '      fraTel.Visible = True
   '      FlexTel.Visible = True
   'End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcidade_GotFocus"
End Sub

Private Sub cmdEmail_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      CNPJCPF_A = Trim(txtCNPJCPF.Text)
      frmEmail.Show 1
   End If
   txtCNPJCPF.PromptInclude = True
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdEmail_Click"
End Sub

Private Sub cmdFoto_Click()
   'If Trim(lstINTEGRA.selectedItem.Text) <> "" Then
   '   LOCAL_IMAGEM = "" & Trim(lstSeleciona.selectedItem.ListSubItems.Item(2).Text)

      frmIMAGEM.Show 1

   '   Item.SubItems(2) = "" & Trim(LOCAL_IMAGEM)
   'End If
End Sub

Private Sub cmbRegiao_Click()
On Error Resume Next

   cmbAuxRegiao.ListIndex = cmbRegiao.ListIndex
End Sub

Private Sub chkSuframa_Click()
   If chkSuframa.Value = 0 Then
      txtInscSuframa.Enabled = False
      Else
         If chkSuframa.Value = 1 Then
            txtInscSuframa.Enabled = True
            txtInscSuframa.SetFocus
         End If
   End If
End Sub

Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         EXCLUIR_REGISTRO_PESSOA
      Case vbKeyF7
         CNPJCPF_A = ""
         frmConsultaPessoa.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
      Case vbKeyDelete
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            EXCLUIR_REGISTRO_PESSOA
      Case vbKeyBack
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            txtCNPJCPF.Mask = "##############"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
       
      If txtCNPJCPF.Text <> "" Then _
         txtNome.SetFocus

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
   TRATA_ERROS Err.description, Me.name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub TXTCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_DADOS_PESSOA

   txtCNPJCPF.PromptInclude = False
   If Len(txtCNPJCPF.Text) > 0 Then
      If CInt(Len(txtCNPJCPF.Text)) = 11 Then
         If Not ValidaCPF(txtCNPJCPF.Text) Then
            MsgBox "CPF com DV incorreto !!!"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF = ""
            SSTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      ElseIf CInt(Len(txtCNPJCPF.Text)) = 14 Then
         If Not VALIDACNPJ(txtCNPJCPF.Text) Then
            MsgBox "CNPJ com DV incorreto !!! "
            txtCNPJCPF.PromptInclude = False
            SSTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "CNPJ/CPF com DV incorreto !!! "
         txtCNPJCPF = ""
         SSTab.Tab = 0
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   ElseIf Len(txtCNPJCPF.Text) <> 0 Then
       MsgBox "CNPJ/CPF com DV incorreto !!! "
       txtCNPJCPF = ""
       SSTab.Tab = 0
       txtCNPJCPF_GotFocus
       txtCNPJCPF.SetFocus
       Exit Sub
   End If
   
   txtCNPJCPF.PromptInclude = False
   CRITERIO = txtCNPJCPF.Text
   txtCNPJCPF.PromptInclude = False
   
   If txtCNPJCPF.Text <> "" Then
      CRITERIO = txtCNPJCPF.Text

      If Not IsNull(txtCNPJCPF.Text) Then
          If Len(txtCNPJCPF.Text) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(txtCNPJCPF.Text) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.Text = CRITERIO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "TXTCNPJCPF_LostFocus"
End Sub

Private Sub txtIE_LostFocus()
   If Trim(txtIE.Text) = "" Then _
      Exit Sub
   If Trim(txtIE.Text) <> "ISENTO" Then
     If Valida_Inscricao_Estadual(txtIE.Text, txtUFValida.Text) <> 0 Then
        'ssTab.Tab = 0
        'txtIE.SetFocus
        'Exit Sub
     End If
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      txtRazao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtNome_KeyPress"
End Sub

Private Sub TXTRAZAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      txtIE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtRazao_KeyPress"
End Sub

Private Sub txtDTNasc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDTNasc_KeyPress"
End Sub

Private Sub txtie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtie_KeyPress"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmbStatus_KeyPress"
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDDD_KeyPress"
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtL.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtN_KeyPress"
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If GRAVA_DADOS_PESSOA(Trim(txtCNPJCPF.Text), Trim(txtNome.Text), Trim(txtRazao.Text), Trim(Left(cmbStatus.Text, 1))) = True Then
         If GRAVA_FONE_PESSOA(Trim(txtN.Text), Trim(txtDDD.Text), Trim(txtL.Text), "0") = True Then
            SETA_FONE
            txtDDD.Text = ""
            txtL.Text = ""
            txtN.Text = ""
         End If
      End If
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtL_KeyPress"
End Sub

Private Sub cmdExcluirFone_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtN.Text) <> "" And PESSOA_ID_N > 0 Then
      EXCLUIR_REGISTRO_FONE Trim(txtN.Text)

      txtN.Text = ""
      txtDDD.Text = ""
      txtL.Text = ""
      SETA_FONE
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdExcluirFone_Click"
End Sub

Private Sub CmdCopiaEnderecoPessoal1_Click()
'On Error GoTo ERRO_TRATA

   txtCepR.PromptInclude = False
   txtCepB.PromptInclude = False
   txtCepC.PromptInclude = False
   If Trim(txtCepR.Text) <> "" Then
      txtCepB.Text = Replace(txtCepR.Text, "-", "")
      txtRuaB.Text = txtRuaR.Text
      txtEndB.Text = txtEndR.Text
      txtBaIrroB.Text = txtBairroR.Text
      txtCidadeB.Text = txtCidadeR.Text
      txtUFB.Text = txtUFR.Text
      txtNumeroB.Text = txtNumeroR.Text
      txtIBGEb.Text = "" & txtIBGE.Text
      Else
         If Trim(txtCepC.Text) <> "" Then
            txtCepB.Text = Replace(txtCepC.Text, "-", "")
            txtRuaB.Text = txtRuaC.Text
            txtEndB.Text = txtEndC.Text
            txtBaIrroB.Text = txtBairroC.Text
            txtCidadeB.Text = txtCidadeC.Text
            txtUFB.Text = txtUFC.Text
            txtNumeroB.Text = txtNumeroC.Text
            txtIBGEb.Text = "" & txtIBGEc.Text
         End If
   End If
   txtCepR.PromptInclude = True
   txtCepB.PromptInclude = True
   txtCepC.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "CmdCopiaEnderecoPessoal1_Click"
End Sub

Private Sub CmdCopiaEnderecoPessoal2_Click()
'On Error GoTo ERRO_TRATA

   txtCepC.PromptInclude = False
   txtCepC.Text = "" & Replace(txtCepR.Text, "-", "")
   txtRuaC.Text = "" & txtRuaR.Text
   txtEndC.Text = "" & txtEndR.Text
   txtBairroC.Text = "" & txtBairroR.Text
   txtCidadeC.Text = "" & txtCidadeR.Text
   txtUFC.Text = "" & txtUFR.Text
   txtUFValida.Text = "" & txtUFR.Text
   txtNumeroC.Text = "" & txtNumeroR.Text
   txtIBGEc.Text = "" & txtIBGE.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "CmdCopiaEnderecoPessoal2_Click"
End Sub
'====================residencial
Private Sub txtCepR_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO = ""
         frmCADASTROCEP.Show 1
         txtCepR.PromptInclude = False
         txtCepR.Text = CRITERIO
         txtCepR.PromptInclude = True
         CRITERIO = ""
      Case vbKeyF7
         CRITERIO = ""
         frmCONSULTACEP.Show 1
         txtCepR.PromptInclude = False
         txtCepR.Text = CRITERIO
         txtCepR.PromptInclude = True
         CRITERIO = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtCepR_KeyDown"
End Sub

Private Sub txtcepr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepR.PromptInclude = False
      If txtCepR.Text <> "" Then
         SP_PROCURA_CEP txtCepR.Text
         If TabCEP.EOF Then
            If TabCEP.State = 1 Then _
               TabCEP.Close

            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepR.SetFocus
            Exit Sub
            Else
               txtCidadeR.Text = TabCEP!Cidade
               txtUFR.Text = TabCEP!UF
               If Not IsNull(TabCEP!IBGE_ID) Then _
                  txtIBGE.Text = TabCEP!IBGE_ID
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
      txtRuaR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcepr_KeyPress"
End Sub

Private Sub txtruar_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNumeroR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtruar_KeyPress"
End Sub

Private Sub txtNumeroR_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtNumeroR_KeyPress"
End Sub

Private Sub txtendr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairroR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtendr_KeyPress"
End Sub

Private Sub txtbairror_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidadeR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtbairror_KeyPress"
End Sub

Private Sub txtcidader_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUFR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcidader_KeyPress"
End Sub

Private Sub txtufr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtufr_KeyPress"
End Sub
'================cobran
Private Sub txtCepb_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO = ""
         frmCADASTROCEP.Show 1
         txtCepB.PromptInclude = False
         txtCepB.Text = CRITERIO
         txtCepB.PromptInclude = True
         CRITERIO = ""
      Case vbKeyF7
         CRITERIO = ""
         frmCONSULTACEP.Show 1
         txtCepB.PromptInclude = False
         txtCepB.Text = CRITERIO
         txtCepB.PromptInclude = True
         CRITERIO = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtCepb_KeyDown"
End Sub

Private Sub txtcepb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepB.PromptInclude = False
      If txtCepB.Text <> "" Then
         SP_PROCURA_CEP txtCepB.Text
         If TabCEP.EOF Then
            If TabCEP.State = 1 Then _
               TabCEP.Close
            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepB.SetFocus
            Exit Sub
            Else
               txtCidadeB.Text = TabCEP!Cidade
               txtUFB.Text = TabCEP!UF
               txtIBGE.Text = TabCEP!IBGE_ID
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
      txtRuaB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcepb_KeyPress"
End Sub

Private Sub txtruab_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtruab_KeyPress"
End Sub

Private Sub txtnumerob_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtnumerob_KeyPress"
End Sub

Private Sub txtendb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBaIrroB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtendb_KeyPress"
End Sub

Private Sub txtbairrob_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidadeB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtbairrob_KeyPress"
End Sub

Private Sub txtcidadeb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUFB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcidadeb_KeyPress"
End Sub

Private Sub txtufb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtufb_KeyPress"
End Sub
'============================comercial
Private Sub txtCepc_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO = ""
         frmCADASTROCEP.Show 1
         txtCepC.PromptInclude = False
         txtCepC.Text = CRITERIO
         txtCepC.PromptInclude = True
         CRITERIO = ""
      Case vbKeyF7
         CRITERIO = ""
         frmCONSULTACEP.Show 1
         txtCepC.PromptInclude = False
         txtCepC.Text = CRITERIO
         txtCepC.PromptInclude = True
         CRITERIO = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtCepc_KeyDown"
End Sub

Private Sub txtcepc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepC.PromptInclude = False
      If txtCepC.Text <> "" Then
         SP_PROCURA_CEP txtCepC.Text
         If TabCEP.EOF Then
            If TabCEP.State = 1 Then _
               TabCEP.Close

            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepC.SetFocus
            Exit Sub
            Else
               txtCidadeC.Text = TabCEP!Cidade
               txtUFC.Text = TabCEP!UF
               txtUFValida.Text = TabCEP!UF
               txtIBGEc.Text = TabCEP!IBGE_ID
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
      txtRuaC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcepc_KeyPress"
End Sub

Private Sub txtruac_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNumeroC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtruac_KeyPress"
End Sub

Private Sub txtnumeroc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtnumeroc_KeyPress"
End Sub

Private Sub txtendc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairroC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtendc_KeyPress"
End Sub

Private Sub txtbairroc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidadeC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtbairroc_KeyPress"
End Sub

Private Sub txtcidadec_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUFC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtcidadec_KeyPress"
End Sub

Private Sub txtufc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtufc_KeyPress"
End Sub

'================
Sub PREPARA_TELA_CADASTRO()
'On Error GoTo ERRO_TRATA

   cmbStatus.Clear
   cmbStatus.AddItem "Ativo"
   cmbStatus.AddItem "Cancelado"
   cmbStatus.AddItem "Especial"

   cmbStatus.Text = "Ativo"

   txtDtCad.PromptInclude = False
   txtDtCad.Text = Date
   txtDtCad.PromptInclude = True
   txtRazao.Visible = False
   lblRazao.Visible = False
   lblDtNasc.Visible = False
   txtDtNasc.Visible = False
   lblIE.Visible = False
   txtIE.Visible = False
   fraResidencial.Visible = False
   fraComercial.Visible = False
   fraCobranca.Visible = False
   SSTab.TabVisible(3) = False

   lblTipoCli.Visible = False
   cmbTipoCli.Visible = False
   lblRegiao.Visible = False
   cmbRegiao.Visible = False
   lblIM.Visible = False
   txtIM.Visible = False
   lblVendedor.Visible = False
   cmbVendedor.Visible = False
   chkESTRANGEIRO.Visible = False
   lblContato.Visible = False
   txtContato.Visible = False
   lblConvenio.Visible = False
   txtPercConv.Visible = False
   lblCredito.Visible = False
   txtCredito.Visible = False
   lstAberto.Visible = False
   staAberto.Visible = False
   txtSaldoDevedor.Visible = False
   lstCompras.Visible = False
   staCompras.Visible = False
   txtTotalVendas.Visible = False
   lblSuframa.Visible = False
   chkSuframa.Visible = False
   txtInscSuframa.Visible = False
   cmdAtGlobal.Visible = False

   If Trim(TIPO_PESSOA_CADASTRO) = "FORNECEDOR" Then _
      MONTA_TELA_FORNECEDOR

   If Trim(TIPO_PESSOA_CADASTRO) = "TRANSPORTADORA" Then _
      MONTA_TELA_TRANSPORTADORA
'SET
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "PREPARA_TELA_CADASTRO"
End Sub

Sub MONTA_TELA_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   Me.Caption = "Cadastro de Fornecedor"
   SSTab.Caption = "Dados Fornecedor"
   txtRazao.Visible = True
   lblRazao.Visible = True
   lblIE.Visible = True
   txtIE.Visible = True
   fraComercial.Visible = True
   fraCobranca.Visible = True
   lblRegiao.Visible = True
   cmbRegiao.Visible = True
   lblIM.Visible = True
   txtIM.Visible = True
   lblVendedor.Visible = True
   lblVendedor.Caption = "Comprador"
   cmbVendedor.Visible = True
   chkESTRANGEIRO.Visible = True
   lblContato.Visible = True
   txtContato.Visible = True
   lstAberto.Visible = True
   staAberto.Visible = True
   txtSaldoDevedor.Visible = True
   lstCompras.Visible = True
   staCompras.Visible = True
   txtTotalVendas.Visible = True
   staAberto.SimpleText = "                                                 Títulos em Aberto                                           à Pagar = "
   staCompras.SimpleText = "                                                 Últimas Compras                                               Total Compras = "

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MONTA_TELA_FORNECEDOR"
End Sub


Sub MONTA_TELA_TRANSPORTADORA()
'On Error GoTo ERRO_TRATA

   Me.Caption = "Cadastro de Transportadora"
   SSTab.Caption = "Dados Transportadora"
   txtRazao.Visible = True
   lblRazao.Visible = True
   lblIE.Visible = True
   txtIE.Visible = True
   fraComercial.Visible = True
   fraCobranca.Visible = True
   lblRegiao.Visible = True
   cmbRegiao.Visible = True
   lblIM.Visible = True
   txtIM.Visible = True
   lblVendedor.Visible = True
   lblVendedor.Caption = "Comprador"
   cmbVendedor.Visible = True
   chkESTRANGEIRO.Visible = True
   lblContato.Visible = True
   txtContato.Visible = True
   lstAberto.Visible = True
   staAberto.Visible = True
   txtSaldoDevedor.Visible = True
   lstCompras.Visible = True
   staCompras.Visible = True
   txtTotalVendas.Visible = True
   staAberto.SimpleText = "                                                 Títulos em Aberto                                           à Pagar = "
   staCompras.SimpleText = "                                                 Últimas Compras                                               Total Compras = "

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MONTA_TELA_TRANSPORTADORA"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   ID_ENTIDADE_N = 0
   txtIM.Text = ""
   txtInscSuframa.Text = ""
   
   txtRazao.Text = ""
   cmbTipoCli.Text = ""
   txtPercConv.Text = "00,00"
   chkESTRANGEIRO.Value = 0
   chkESTRANGEIRO.ForeColor = vbBlack
   txtOBS.Text = ""
   VALOR_TOTAL_N = 0
   lstAberto.ListItems.Clear
   lstCompras.ListItems.Clear
   txtSaldoDevedor.Text = ""
   cmbRegiao.Text = ""
   cmbAuxRegiao.Text = ""
   txtContato.Text = ""
   txtDtCad.PromptInclude = False
   txtDtCad.Text = ""
   txtDtNasc.PromptInclude = False
   txtDtNasc.Text = ""
   cmbVendedor.Text = ""
   cmbVendedorAUX.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCNPJCPF.Mask = "##############"
   txtNome.Text = ""
   cmbStatus.Text = "Ativo"
   txtNumeroR.Text = ""
   txtNumeroC.Text = ""
   txtNumeroB.Text = ""
   txtIE.Text = ""
   txtCepR.PromptInclude = False
   txtCepR.Text = ""
   txtRuaR.Text = ""
   txtEndR.Text = ""
   txtBairroR.Text = ""
   txtCidadeR.Text = ""
   txtUFR.Text = ""
   txtCepC.PromptInclude = False
   txtCepC.Text = ""
   txtRuaC.Text = ""
   txtEndC.Text = ""
   txtBairroC.Text = ""
   txtCidadeC.Text = ""
   txtUFC.Text = ""
   txtUFValida.Text = ""
   txtCepB.PromptInclude = False
   txtCepB.Text = ""
   txtRuaB.Text = ""
   txtEndB.Text = ""
   txtBaIrroB.Text = ""
   txtCidadeB.Text = ""
   txtUFB.Text = ""
   CRITERIO = 0
   txtCredito.Text = ""
   SSTab.Tab = 0
   PESSOA_ID_N = 0
   lblPessoaID.Caption = PESSOA_ID_N
   SETA_FONE
   MOSTRA_OBS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "LIMPA_TUDO"
End Sub

Sub MOSTRA_DADOS_PESSOA()
'On Error GoTo ERRO_TRATA

   ID_ENTIDADE_N = 0
   PESSOA_ID_N = 0
   lblPessoaID.Caption = PESSOA_ID_N
   txtNome.Text = ""
   txtRazao.Text = ""
   txtDtCad.PromptInclude = False
   txtDtCad.Text = ""
   txtDtNasc.PromptInclude = False
   txtDtNasc.Text = ""
   txtIE.Text = ""
   cmbStatus.Text = ""

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If IsNumeric(txtCNPJCPF.Text) Then
         Dim TabPessoa     As New ADODB.Recordset
         Dim TabEntidade   As New ADODB.Recordset

         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         SQL = "select * from PESSOA WITH (NOLOCK)"
         SQL = SQL & " where CNPJCPF = '" & Trim(txtCNPJCPF.Text) & "'"
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then
            PESSOA_ID_N = 0 & TabPessoa.Fields("pessoa_id").Value
            lblPessoaID.Caption = PESSOA_ID_N
            txtNome.Text = "" & Trim(TabPessoa.Fields("descricao").Value)
            txtRazao.Text = "" & Trim(TabPessoa.Fields("razao").Value)
            txtDtCad.Text = "" & Trim(TabPessoa.Fields("data_cad").Value)
            If Not IsNull(TabPessoa.Fields("situacao").Value) Then
               If Trim(TabPessoa.Fields("situacao").Value) = "A" Then
                  cmbStatus.Text = "Ativo"
                  Else
                     If Trim(TabPessoa.Fields("situacao").Value) = "E" Then
                        cmbStatus.Text = "Especial"
                        Else
                           If Trim(TabPessoa.Fields("situacao").Value) = "C" Then _
                              cmbStatus.Text = "Cancelado"
                     End If
               End If
            End If
         End If
         If TabPessoa.State = 1 Then _
            TabPessoa.Close
         Else: Exit Sub
      End If
      Else: Exit Sub
   End If
   If PESSOA_ID_N <= 0 Then _
      Exit Sub

   If TabEntidade.State = 1 Then _
      TabEntidade.Close
         If Trim(TIPO_PESSOA_CADASTRO) = "FORNECEDOR" Then
            SQL = "select * from vwFornecedor WITH (NOLOCK)"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            TabEntidade.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabEntidade.EOF Then
               ID_ENTIDADE_N = 0 & TabEntidade.Fields("fornecedor_id").Value
               txtDtCad.Text = "" & TabEntidade.Fields("dt_cad").Value
               txtContato.Text = "" & TabEntidade.Fields("contato").Value
               If Not IsNull(TabEntidade.Fields("status").Value) Then
                  If Trim(TabEntidade.Fields("status").Value) = "A" Then
                     cmbStatus.Text = "Ativo"
                     Else
                        If Trim(TabEntidade.Fields("status").Value) = "E" Then
                           cmbStatus.Text = "Especial"
                           Else
                              If Trim(TabEntidade.Fields("status").Value) = "C" Then _
                                 cmbStatus.Text = "Cancelado"
                        End If
                  End If
               End If
            End If
         End If
         If Trim(TIPO_PESSOA_CADASTRO) = "TRANSPORTADORA" Then
            SQL = "select * from vwTRANSPORTADORA WITH (NOLOCK)"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            TabEntidade.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabEntidade.EOF Then
               ID_ENTIDADE_N = 0 & TabEntidade.Fields("TRANSP_ID").Value
               txtDtCad.Text = "" & TabEntidade.Fields("dt_cad").Value
               txtContato.Text = "" & TabEntidade.Fields("contato").Value
               If Not IsNull(TabEntidade.Fields("status").Value) Then
                  If Trim(TabEntidade.Fields("status").Value) = "A" Then
                     cmbStatus.Text = "Ativo"
                     Else
                        If Trim(TabEntidade.Fields("status").Value) = "E" Then
                           cmbStatus.Text = "Especial"
                           Else
                              If Trim(TabEntidade.Fields("status").Value) = "C" Then _
                                 cmbStatus.Text = "Cancelado"
                        End If
                  End If
               End If
            End If
         End If
   If TabEntidade.State = 1 Then _
      TabEntidade.Close
   txtCNPJCPF.PromptInclude = True

   SETA_FONE
   MOSTRA_OBS
   MOSTRA_ENDERECO
   txtIE.Text = Trim(TRAZ_IE(PESSOA_ID_N))
   txtIM.Text = Trim(TRAZ_IM)
   TRAZ_FORNECEDORCOMPRADOR

   If Trim(txtIE.Text) = "" Then _
      txtIE.Text = "ISENTO"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MOSTRA_DADOS_PESSOA"
End Sub

Function GRAVA_DADOS_PESSOA(CNPJCPF_A As String, NOME_A As String, RAZAO_A As String, SITUACAO_A As String) As Boolean
'On Error GoTo ERRO_TRATA

   GRAVA_DADOS_PESSOA = False
   If Trim(CNPJCPF_A) <> "" And Trim(NOME_A) <> "" And Trim(SITUACAO_A) <> "" Then
      Dim TabPessoa     As New ADODB.Recordset

      PESSOA_ID_N = 0 & TRAZ_ID_TABELA("PESSOA", "pessoa_id", "cnpjcpf", Trim(txtCNPJCPF.Text))

      CONT_N = 1
      If PESSOA_ID_N > 0 Then _
         CONT_N = 2

      'executa stored procedure sp_pessoa
      SP_PESSOA CONT_N, PESSOA_ID_N, Trim(CNPJCPF_A), Trim(NOME_A), Trim(RAZAO_A), Trim(SITUACAO_A)

      PESSOA_ID_N = 0 & TRAZ_ID_TABELA("PESSOA", "pessoa_id", "cnpjcpf", Trim(txtCNPJCPF.Text))
      lblPessoaID.Caption = PESSOA_ID_N

      GRAVA_DADOS_PESSOA = True
      Else: MsgBox "Informar dados corretamente."
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "GRAVA_DADOS_PESSOA"
End Function

Function GRAVA_DADOS_FORNECEDOR() As Boolean
'On Error GoTo ERRO_TRATA

   GRAVA_DADOS_FORNECEDOR = False
   ID_ENTIDADE_N = 0

   If PESSOA_ID_N > 0 Then
      Dim TabFor  As New ADODB.Recordset

      If TabFor.State = 1 Then _
         TabFor.Close

      SQL = "select fornecedor_id from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabFor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFor.EOF Then
         ID_ENTIDADE_N = 0 & TabFor.Fields(0).Value
         SQL = "update FORNECEDOR set "
            SQL = SQL & " status = '" & Trim(Left(cmbStatus.Text, 1)) & "'"
            SQL = SQL & ", contato = '" & Trim(txtContato.Text) & "'"
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         Else
            SQL = "insert into FORNECEDOR "
               SQL = SQL & "(FORNECEDOR_ID,PESSOA_ID,ESTABELECIMENTO_ID,DT_CAD,STATUS,contato)"
            SQL = SQL & " values( "
               SQL = SQL & MAX_ID("fornecedor_id", "FORNECEDOR", "", "", "", "")
               SQL = SQL & "," & PESSOA_ID_N
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
               SQL = SQL & ",'" & Now & "'"
               SQL = SQL & ",'" & Trim(Left(cmbStatus.Text, 1)) & "'"
               SQL = SQL & ",'" & Trim(txtContato.Text) & "'"
            SQL = SQL & ")"
      End If
      If TabFor.State = 1 Then _
         TabFor.Close

      CONECTA_RETAGUARDA.Execute SQL
      GRAVA_DADOS_FORNECEDOR = True
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "GRAVA_DADOS_FORNECEDOR"
End Function

Function GRAVA_DADOS_TRANSPORTADORA() As Boolean
'On Error GoTo ERRO_TRATA

   GRAVA_DADOS_TRANSPORTADORA = False
   ID_ENTIDADE_N = 0

   If PESSOA_ID_N > 0 Then
      Dim TabFor  As New ADODB.Recordset

      If TabFor.State = 1 Then _
         TabFor.Close

      SQL = "select TRANSP_ID from vwTRANSPORTADORA WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabFor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFor.EOF Then
         ID_ENTIDADE_N = 0 & TabFor.Fields(0).Value
         SQL = "update TRANSPORTADORA set "
            SQL = SQL & " status = '" & Trim(Left(cmbStatus.Text, 1)) & "'"
            SQL = SQL & ", contato = '" & Trim(txtContato.Text) & "'"
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         Else
            SQL = "insert into TRANSPORTADORA "
               SQL = SQL & "(TRANSP_ID,PESSOA_ID,ESTABELECIMENTO_ID,DT_CAD,STATUS,contato)"
            SQL = SQL & " values( "
               SQL = SQL & MAX_ID("TRANSP_ID", "TRANSPORTADORA", "", "", "", "")
               SQL = SQL & "," & PESSOA_ID_N
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
               SQL = SQL & ",'" & Now & "'"
               SQL = SQL & ",'" & Trim(Left(cmbStatus.Text, 1)) & "'"
               SQL = SQL & ",'" & Trim(txtContato.Text) & "'"
            SQL = SQL & ")"
      End If
      If TabFor.State = 1 Then _
         TabFor.Close

      CONECTA_RETAGUARDA.Execute SQL
      GRAVA_DADOS_TRANSPORTADORA = True
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "GRAVA_DADOS_TRANSPORTADORA"
End Function

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   adoFone.Enabled = True
   adoFone.ConnectionString = AUTENTICA_GRID

   SQL = "select ddd,numero,local from FONE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " order by NUMERO"

   adoFone.RecordSource = SQL
   adoFone.Enabled = True
   adoFone.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "SETA_FONE"
End Sub

Sub MOSTRA_OBS()
'On Error GoTo ERRO_TRATA

   Dim TabOBS  As New ADODB.Recordset

   txtOBS.Text = ""
   If TabOBS.State = 1 Then _
      TabOBS.Close

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and seq = 0 "
   TabOBS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabOBS.EOF Then _
      If Not IsNull(TabOBS!obs) Then _
         txtOBS.Text = "" & TabOBS!obs
   If TabOBS.State = 1 Then _
      TabOBS.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MOSTRA_OBS"
End Sub

Sub MOSTRA_ENDERECO()
'On Error GoTo ERRO_TRATA

   BUSCA_ENDERECO_PESSOA "", ""
   While Not tabEndereco.EOF
      If Trim(tabEndereco.Fields("tipo").Value) = "R" Then
         txtCepR.Text = "" & tabEndereco.Fields("cep_id").Value
         txtRuaR.Text = "" & tabEndereco.Fields("rua").Value
         txtNumeroR.Text = "" & tabEndereco.Fields("numero").Value
         txtEndR.Text = "" & tabEndereco.Fields("complemento").Value
         txtBairroR.Text = "" & tabEndereco.Fields("bairro").Value
         txtCidadeR.Text = "" & tabEndereco.Fields("cidade").Value
         txtUFR.Text = "" & tabEndereco.Fields("uf").Value
         txtIBGE.Text = "" & tabEndereco.Fields("IBGE_ID").Value
      End If
      If Trim(tabEndereco.Fields("tipo").Value) = "C" Then
         txtCepC.Text = "" & tabEndereco.Fields("cep_id").Value
         txtRuaC.Text = "" & tabEndereco.Fields("rua").Value
         txtNumeroC.Text = "" & tabEndereco.Fields("numero").Value
         txtEndC.Text = "" & tabEndereco.Fields("complemento").Value
         txtBairroC.Text = "" & tabEndereco.Fields("bairro").Value
         txtCidadeC.Text = "" & tabEndereco.Fields("cidade").Value
         txtUFC.Text = "" & tabEndereco.Fields("uf").Value
         txtUFValida.Text = "" & tabEndereco.Fields("uf").Value
         txtIBGEc.Text = "" & tabEndereco.Fields("IBGE_ID").Value
      End If
      If Trim(tabEndereco.Fields("tipo").Value) = "B" Then
         txtCepB.Text = "" & tabEndereco.Fields("cep_id").Value
         txtRuaB.Text = "" & tabEndereco.Fields("rua").Value
         txtNumeroB.Text = "" & tabEndereco.Fields("numero").Value
         txtEndC.Text = "" & tabEndereco.Fields("complemento").Value
         txtBaIrroB.Text = "" & tabEndereco.Fields("bairro").Value
         txtCidadeB.Text = "" & tabEndereco.Fields("cidade").Value
         txtUFB.Text = "" & tabEndereco.Fields("uf").Value
         txtIBGEb.Text = "" & tabEndereco.Fields("IBGE_ID").Value
      End If
      tabEndereco.MoveNext
   Wend
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MOSTRA_ENDERECO"
End Sub

Sub GRAVA_OBS()
'On Error GoTo ERRO_TRATA

   If PESSOA_ID_N <= 0 Then _
      Exit Sub

   Dim TabOBS  As New ADODB.Recordset

   If TabOBS.State = 1 Then _
      TabOBS.Close

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and seq = 0 "
   TabOBS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabOBS.EOF Then
      SQL = "insert into OBS (PESSOA_ID,SEQ,OBS) "
      SQL = SQL & " values("
         SQL = SQL & PESSOA_ID_N
         SQL = SQL & ",0"                                      '[SEQ]
         SQL = SQL & ",'" & Trim(txtOBS.Text) & "'"   '[OBS]
      SQL = SQL & ")"
      Else
         SQL = "update OBS set"
         SQL = SQL & " obs = '" & Trim(txtOBS.Text) & "'"  '[OBS]
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and seq = 0"
   End If
   If TabOBS.State = 1 Then _
      TabOBS.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "GRAVA_OBS"
End Sub

Sub TRAZ_FORNECEDORCOMPRADOR()
'On Error GoTo ERRO_TRATA

   If PESSOA_ID_N > 0 Then
      Dim TabFor  As New ADODB.Recordset

      If TabFor.State = 1 Then _
         TabFor.Close

      SQL = "select usuario.USUARIO_ID,usuario.nome "
      SQL = SQL & " FROM FORNECEDOR "
      SQL = SQL & " INNER JOIN FORNECEDORCOMPRADOR "
      SQL = SQL & " ON FORNECEDOR.FORNECEDOR_ID = FORNECEDORCOMPRADOR.FORNECEDOR_ID "
      SQL = SQL & " INNER JOIN USUARIO "
      SQL = SQL & " ON FORNECEDORCOMPRADOR.USUARIO_ID = USUARIO.USUARIO_ID"
      SQL = SQL & " where FORNECEDOR.pessoa_id = " & PESSOA_ID_N
      TabFor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFor.EOF Then
      End If
      If TabFor.State = 1 Then _
         TabFor.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "TRAZ_FORNECEDORCOMPRADOR"
End Sub
'=============
Sub EXCLUIR_PESSOA()
         txtCNPJCPF.PromptInclude = False

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select cliente_id,pessoa_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then
            PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
            CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value

            If TabTemp.State = 1 Then _
               TabTemp.Close
            'procura venda
            SQL = "select * from PEDIDO WITH (NOLOCK)"
            SQL = SQL & " where cliente_id = " & CLIENTE_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabTemp.EOF Then
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               'procura faturamento
               SQL = "select * from LANCAMENTO WITH (NOLOCK)"
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabConsulta.EOF Then
                  Msg = "Confirma exclusão do registro?"
                  PERGUNTA Msg, vbYesNo + 32, "Cadastro Cliente NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     SQL = "delete from OBS "
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     SQL = "delete from EMAIL"
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     SQL = "delete FONE"
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     SQL = "delete from ENDERECO"
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     SQL = "delete from RG"
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     SQL = "delete from CLIENTE "
                     SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL

                     'executa stored procedure sp_pessoa
                     SP_PESSOA 3, PESSOA_ID_N, "", "", "", ""

                     LIMPA_TUDO
                  End If
                  Else: MsgBox "Operação não permitida, cliente possue movimentação no financeiro."
               End If
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               Else: MsgBox "Operação não permitida, cliente possue movimentação de venda."
            End If
            txtCNPJCPF.SetFocus

            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         If TabCliente.State = 1 Then _
            TabCliente.Close

End Sub
Sub EXCLUIR_REGISTRO_PESSOA()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
         If txtCNPJCPF.Text <> "" Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select * from PEDIDO WITH (NOLOCK)"
            SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then
               If TabCliente.State = 1 Then _
                  TabCliente.Close

               MsgBox "Impossível excluir, cliente possue movimento de vendas."
               Exit Sub
               Else
                  If TabCliente.State = 1 Then _
                     TabCliente.Close

                  SQL = "select * from LANCAMENTO WITH (NOLOCK)"
                  SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                  TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabCliente.EOF Then
                     If TabCliente.State = 1 Then _
                        TabCliente.Close

                     MsgBox "Impossível excluir, cliente possue movimento financeiro."
                     Exit Sub
                  End If
            End If
            If TabCliente.State = 1 Then _
               TabCliente.Close

            Msg = "Confirma exclusão total de cliente ?"
            PERGUNTA Msg, vbYesNo + 32, "Cadastro Cliente NFE", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "delete from CLIENTE "
               SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
               CONECTA_RETAGUARDA.Execute SQL
               MsgBox "Cliente excluido definitivamente do banco de dados."
               LIMPA_TUDO
               SSTab.Tab = 0
               txtCNPJCPF.SetFocus
            End If
         End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "EXCLUIR_REGISTRO_PESSOA"
End Sub
