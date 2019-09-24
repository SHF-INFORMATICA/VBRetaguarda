VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROCLIENTE 
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   930
   ClientWidth     =   11220
   Icon            =   "CADASTROCLIENTE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   50
      TabIndex        =   48
      Top             =   6480
      Width           =   11145
      Begin VB.CommandButton cmdExcluirFone 
         Height          =   375
         Left            =   10680
         Picture         =   "CADASTROCLIENTE.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   375
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
         Height          =   315
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   23
         Top             =   150
         Width           =   6945
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
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   22
         Top             =   150
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
         Height          =   315
         Left            =   720
         MaxLength       =   2
         TabIndex        =   21
         Top             =   150
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Local:"
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
         Index           =   13
         Left            =   2850
         TabIndex        =   51
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DDD:"
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
         Left            =   240
         TabIndex        =   50
         Top             =   210
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexTel 
      Height          =   1305
      Left            =   45
      TabIndex        =   47
      Top             =   7020
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   2302
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   16711680
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   -150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":6A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":6EA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":71C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":7617
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":7A6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":7D8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":81DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":84FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":8953
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":A0E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":AAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":B50B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":BF1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":C92F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROCLIENTE.frx":D341
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5775
      Left            =   50
      TabIndex        =   52
      Top             =   720
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   10186
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
      TabPicture(0)   =   "CADASTROCLIENTE.frx":DD53
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPessoa_id"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraPessoa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraRes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2- Endereço Cobrança"
      TabPicture(1)   =   "CADASTROCLIENTE.frx":DD6F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label25"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "txtOBS"
      Tab(1).Control(3)=   "cmdObs"
      Tab(1).Control(4)=   "txtObs2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&3- Dados Comerciais  "
      TabPicture(2)   =   "CADASTROCLIENTE.frx":DD8B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "fraCom"
      Tab(2).Control(2)=   "chkSuframa"
      Tab(2).Control(3)=   "txtInscSuframa"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&4 - Histórico"
      TabPicture(3)   =   "CADASTROCLIENTE.frx":DDA7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label21"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label23"
      Tab(3).Control(3)=   "Label30"
      Tab(3).Control(4)=   "LISTAASS"
      Tab(3).Control(5)=   "StatusBar1"
      Tab(3).Control(6)=   "StatusBar2"
      Tab(3).Control(7)=   "LISTAITEM"
      Tab(3).Control(8)=   "GRIDCONTA"
      Tab(3).Control(9)=   "cmbAgencia"
      Tab(3).Control(10)=   "cmbBanco"
      Tab(3).Control(11)=   "cmbNumr_Conta"
      Tab(3).Control(12)=   "cmbAuxConta"
      Tab(3).Control(13)=   "cmbAUXB"
      Tab(3).Control(14)=   "cmbAUXA"
      Tab(3).Control(15)=   "DATACONTA"
      Tab(3).Control(16)=   "txtSaldoDevedor"
      Tab(3).Control(17)=   "txtTotalVendas"
      Tab(3).Control(18)=   "txtLIMITE"
      Tab(3).ControlCount=   19
      Begin VB.TextBox txtObs2 
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
         Height          =   2670
         Left            =   -73080
         MultiLine       =   -1  'True
         TabIndex        =   116
         Text            =   "CADASTROCLIENTE.frx":DDC3
         Top             =   2520
         Visible         =   0   'False
         Width           =   11025
      End
      Begin VB.CommandButton cmdObs 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   -73560
         Picture         =   "CADASTROCLIENTE.frx":DF6D
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1845
         Width           =   350
      End
      Begin VB.TextBox txtInscSuframa 
         DataField       =   "Nome"
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
         Left            =   -71760
         MaxLength       =   100
         TabIndex        =   32
         Top             =   3360
         Width           =   3435
      End
      Begin VB.CheckBox chkSuframa 
         Caption         =   "&Suframa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   112
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtLIMITE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -65100
         MaxLength       =   50
         TabIndex        =   44
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTotalVendas 
         Alignment       =   1  'Right Justify
         DataField       =   "Bairro_Res"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   -66060
         MaxLength       =   50
         TabIndex        =   95
         Top             =   5340
         Width           =   2175
      End
      Begin VB.TextBox txtSaldoDevedor 
         Alignment       =   1  'Right Justify
         DataField       =   "Bairro_Res"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   -66030
         MaxLength       =   50
         TabIndex        =   94
         Top             =   3650
         Width           =   2145
      End
      Begin VB.Data DATACONTA 
         Caption         =   "CONTA"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -70500
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1140
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.ComboBox cmbAUXA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74100
         TabIndex        =   93
         Top             =   780
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbAUXB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74100
         TabIndex        =   92
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbAuxConta 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73260
         TabIndex        =   91
         Top             =   1140
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbNumr_Conta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73260
         TabIndex        =   43
         Top             =   1140
         Width           =   2535
      End
      Begin VB.ComboBox cmbBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74100
         TabIndex        =   41
         Top             =   420
         Width           =   3375
      End
      Begin VB.ComboBox cmbAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74100
         TabIndex        =   42
         Top             =   780
         Width           =   3375
      End
      Begin VB.Frame fraCom 
         Caption         =   " Endereço Comercial "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2715
         Left            =   -74880
         TabIndex        =   83
         Top             =   480
         Width           =   10995
         Begin VB.CommandButton cmdAtGlobal 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Atualizar cadastro do cliente no banco NFe?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
            Top             =   1920
            Width           =   3105
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
            Left            =   6720
            MaxLength       =   50
            TabIndex        =   35
            Top             =   660
            Width           =   825
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
            Height          =   345
            Left            =   7680
            MaxLength       =   50
            TabIndex        =   36
            Top             =   660
            Width           =   3135
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
            Height          =   345
            Left            =   7680
            MaxLength       =   2
            TabIndex        =   39
            Top             =   1380
            Width           =   3135
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
            Height          =   345
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   38
            Top             =   1380
            Width           =   4695
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
            Height          =   345
            Left            =   120
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1380
            Width           =   2655
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
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   34
            Top             =   660
            Width           =   3735
         End
         Begin VB.CommandButton CmdCopiaEnderecoPessoal2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Copiar Endereço Pessoal."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7710
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
            Top             =   2070
            Width           =   3105
         End
         Begin MSMask.MaskEdBox txtCepC 
            Height          =   345
            Left            =   120
            TabIndex        =   33
            Top             =   660
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####-###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6720
            TabIndex        =   108
            Top             =   420
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7680
            TabIndex        =   90
            Top             =   420
            Width           =   960
         End
         Begin VB.Label lblEstadoCom 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7680
            TabIndex        =   89
            Top             =   1140
            Width           =   315
         End
         Begin VB.Label lblCepCom 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   88
            Top             =   420
            Width           =   435
         End
         Begin VB.Label lblCidadeCom 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2940
            TabIndex        =   87
            Top             =   1140
            Width           =   735
         End
         Begin VB.Label lblBairroCom 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   86
            Top             =   1140
            Width           =   645
         End
         Begin VB.Label lblRuaCom 
            AutoSize        =   -1  'True
            Caption         =   "Rua:"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2880
            TabIndex        =   85
            Top             =   420
            Width           =   435
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
         Height          =   3510
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "CADASTROCLIENTE.frx":E0B7
         Top             =   2160
         Width           =   11025
      End
      Begin VB.Frame Frame1 
         Caption         =   " Endereço Cobrança "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1455
         Left            =   -74850
         TabIndex        =   74
         Top             =   360
         Width           =   10995
         Begin VB.TextBox txtNumeroB 
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
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   26
            Top             =   360
            Width           =   705
         End
         Begin VB.TextBox txtRuaB 
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
            Height          =   360
            Left            =   2340
            MaxLength       =   50
            TabIndex        =   25
            Top             =   360
            Width           =   2115
         End
         Begin VB.TextBox txtBaIrroB 
            DataField       =   "Bairro_Res"
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
            Height          =   345
            Left            =   840
            MaxLength       =   50
            TabIndex        =   28
            Top             =   840
            Width           =   2595
         End
         Begin VB.TextBox txtCidadeB 
            DataField       =   "Cidade"
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
            Left            =   4380
            MaxLength       =   50
            TabIndex        =   29
            Top             =   840
            Width           =   3195
         End
         Begin VB.TextBox txtUFB 
            Alignment       =   2  'Center
            DataField       =   "Estado"
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
            Left            =   8040
            MaxLength       =   2
            TabIndex        =   30
            Top             =   840
            Width           =   645
         End
         Begin VB.TextBox txtEndB 
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
            Left            =   7680
            MaxLength       =   50
            TabIndex        =   27
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton CmdCopiaEnderecoPessoal1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Copiar endereço pessoal"
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
            Left            =   8760
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço cobrança."
            Top             =   840
            Width           =   2175
         End
         Begin MSMask.MaskEdBox txtCepB 
            Height          =   345
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####-###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4800
            TabIndex        =   107
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Rua:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1860
            TabIndex        =   81
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   80
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3510
            TabIndex        =   79
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   90
            TabIndex        =   78
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6600
            TabIndex        =   77
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
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
            Left            =   7680
            TabIndex        =   76
            Top             =   840
            Width           =   315
         End
      End
      Begin VB.Frame fraRes 
         Caption         =   " Endereço Residencial / Comercial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1665
         Left            =   100
         TabIndex        =   66
         Top             =   3480
         Width           =   10995
         Begin VB.TextBox txtNumeroR 
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
            TabIndex        =   15
            Top             =   570
            Width           =   825
         End
         Begin VB.TextBox txtEndR 
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
            TabIndex        =   16
            Top             =   570
            Width           =   1905
         End
         Begin VB.TextBox txtUFR 
            Alignment       =   2  'Center
            DataField       =   "Estado"
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
            TabIndex        =   19
            Top             =   1200
            Width           =   555
         End
         Begin VB.TextBox txtCidadeR 
            DataField       =   "Cidade"
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
            TabIndex        =   18
            Top             =   1200
            Width           =   4755
         End
         Begin VB.TextBox txtBairroR 
            DataField       =   "Bairro_Res"
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
            Left            =   90
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1200
            Width           =   2745
         End
         Begin VB.TextBox txtRuaR 
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
            Height          =   360
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   14
            Top             =   570
            Width           =   5835
         End
         Begin VB.TextBox txtIBGE 
            DataField       =   "Bairro_Res"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9540
            LinkTimeout     =   7
            MaxLength       =   50
            TabIndex        =   20
            Top             =   1200
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtCepR 
            Height          =   360
            Left            =   90
            TabIndex        =   13
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
               Weight          =   400
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
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7920
            TabIndex        =   106
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "*UF:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7950
            TabIndex        =   73
            Top             =   990
            Width           =   390
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            Caption         =   "*Complemento:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9030
            TabIndex        =   72
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label lblCep 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "*Cep:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   30
            TabIndex        =   71
            Top             =   330
            Width           =   525
         End
         Begin VB.Label lblCidade 
            AutoSize        =   -1  'True
            Caption         =   "*Cidade:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2970
            TabIndex        =   70
            Top             =   990
            Width           =   810
         End
         Begin VB.Label lblBairro 
            AutoSize        =   -1  'True
            Caption         =   "*Bairro:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   69
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            Caption         =   "*Rua:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1890
            TabIndex        =   68
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "*Codigo IBGE:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9540
            TabIndex        =   67
            Top             =   990
            Width           =   1335
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
         Height          =   2955
         Left            =   105
         TabIndex        =   53
         Top             =   360
         Width           =   10995
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
            Left            =   9480
            TabIndex        =   120
            Top             =   2640
            Width           =   1455
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
            Height          =   360
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   4
            Top             =   960
            Width           =   5235
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
            Left            =   8640
            TabIndex        =   117
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00FFFFFF&
            Height          =   350
            Left            =   1980
            Picture         =   "CADASTROCLIENTE.frx":E261
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   480
            Width           =   405
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
            Left            =   7800
            TabIndex        =   111
            Top             =   960
            Width           =   855
         End
         Begin VB.ComboBox cmbAuxVendedor 
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
            Left            =   2040
            TabIndex        =   109
            Top             =   1875
            Visible         =   0   'False
            Width           =   735
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
            Left            =   7680
            TabIndex        =   105
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtPercConv 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   5640
            MaxLength       =   5
            TabIndex        =   9
            Text            =   "00,00"
            Top             =   1920
            Width           =   1095
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
            Left            =   9810
            TabIndex        =   3
            Top             =   480
            Width           =   1095
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
            Height          =   360
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   1
            Top             =   480
            Width           =   5235
         End
         Begin VB.ComboBox cmbVENDEDOR 
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
            Left            =   2040
            TabIndex        =   8
            Top             =   1875
            Width           =   1935
         End
         Begin VB.ComboBox cmbRegiao 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7740
            TabIndex        =   2
            Top             =   480
            Width           =   1965
         End
         Begin VB.TextBox txtContato 
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
            Height          =   360
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2400
            Width           =   1935
         End
         Begin VB.ComboBox cmbTipoCli 
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
            Left            =   5640
            TabIndex        =   12
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CheckBox chkESTRANGEIRO 
            Caption         =   "Estrangeiro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   9480
            TabIndex        =   40
            Top             =   2400
            Width           =   1215
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
            Height          =   360
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   5
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtIM 
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
            Left            =   5640
            MaxLength       =   25
            TabIndex        =   6
            Top             =   1440
            Width           =   2085
         End
         Begin MSMask.MaskEdBox txtCNPJCPF 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   635
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
            Height          =   360
            Left            =   9480
            TabIndex        =   10
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Height          =   360
            Left            =   9480
            TabIndex        =   7
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
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
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   960
            TabIndex        =   118
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "% Convenio:"
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
            Left            =   4380
            TabIndex        =   65
            Top             =   1920
            Width           =   1170
         End
         Begin VB.Label lblTipoCli 
            AutoSize        =   -1  'True
            Caption         =   "*CNPJ/CPF:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   90
            TabIndex        =   64
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label lblNomeCli 
            AutoSize        =   -1  'True
            Caption         =   "*Nome:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2460
            TabIndex        =   63
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "*Status:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   9810
            TabIndex        =   62
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   960
            TabIndex        =   61
            Top             =   1875
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "*Região:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7770
            TabIndex        =   60
            Top             =   270
            Width           =   675
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dt.Nascim.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8280
            TabIndex        =   59
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data Cadastro:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7920
            TabIndex        =   58
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Quem Liberou:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   57
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insc.Municipal:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4080
            TabIndex        =   56
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblInsc 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "*Inscrição Estadual:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cliente:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4290
            TabIndex        =   54
            Top             =   2400
            Width           =   1215
         End
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
               Picture         =   "CADASTROCLIENTE.frx":EC63
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":F0B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":F3D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":F827
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":FC7B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":FF9B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":103EF
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":1070F
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":10B63
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSDBGrid.DBGrid GRIDCONTA 
         Bindings        =   "CADASTROCLIENTE.frx":122F7
         Height          =   1155
         Left            =   -70620
         OleObjectBlob   =   "CADASTROCLIENTE.frx":1230F
         TabIndex        =   96
         Top             =   360
         Width           =   5415
      End
      Begin MSComctlLib.ListView LISTAITEM 
         Height          =   1665
         Left            =   -74900
         TabIndex        =   45
         Top             =   1560
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   2937
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
      Begin MSComctlLib.StatusBar StatusBar2 
         Height          =   405
         Left            =   -74880
         TabIndex        =   101
         Top             =   3600
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   714
         Style           =   1
         SimpleText      =   "                                                 Títulos em Aberto                                              Saldo Devedor = "
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
         OLEDropMode     =   1
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   405
         Left            =   -74880
         TabIndex        =   102
         Top             =   5280
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   714
         Style           =   1
         SimpleText      =   "                                                 Últimas Compras                                               Total Compras = "
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
         OLEDropMode     =   1
      End
      Begin MSComctlLib.ListView LISTAASS 
         Height          =   1185
         Left            =   -74900
         TabIndex        =   46
         Top             =   4080
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   2090
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
      Begin VB.Label lblPessoa_id 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9915
         TabIndex        =   119
         Top             =   5280
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Suframa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -73560
         TabIndex        =   113
         Top             =   3360
         Width           =   1785
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Limite Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -65100
         TabIndex        =   100
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74820
         TabIndex        =   99
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74730
         TabIndex        =   98
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74880
         TabIndex        =   97
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74880
         TabIndex        =   82
         Top             =   1920
         Width           =   1185
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   1270
      ButtonWidth     =   2858
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
            Key             =   "voltar"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "mata"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   10250
         TabIndex        =   104
         Top             =   0
         Width           =   945
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5280
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":12EA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":14042
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":150D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":16339
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":17444
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROCLIENTE.frx":183F9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11220
      DesignHeight    =   8400
   End
End
Attribute VB_Name = "frmCADASTROCLIENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim intCodigo  As Long
   Dim strRegiao  As Long
   Dim strDtNac   As String

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   VERIFICA_TABELA_CLIENTE

   SSTab.Tab = 0

   MONTA_DESCRITORES
   MONTA_VENDEDORES

   CRITERIO_A = 0

   cmbStatus.Clear
   cmbStatus.AddItem "Ativo"
   cmbStatus.AddItem "Cancelado"
   cmbStatus.AddItem "Especial"

   PreencheTipoCliente
   refresh_GRID
   SETA_FONE
   GeraCodigo

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub listaass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTAASS, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "mata"
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
                  Msg = ""
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

                     'executa stored procedure spPessoa
                     spPessoa 3, PESSOA_ID_N, "", "", "", ""

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
      Case "voltar"
         Unload Me
      Case "gravar"
         GRAVA_TUDO
         SSTab.Tab = 0
         txtCNPJCPF.SetFocus
      Case "print"
         MONTA_REL_CLI
      Case "limpar"
         LIMPA_TUDO
         SSTab.Tab = 0
         txtCNPJCPF.SetFocus
      Case "consultar"
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A

            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select * from CLIENTE WITH (NOLOCK)"
            SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
            If TabCliente.State = 1 Then _
               TabCliente.Close
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               MOSTRA_DADOS
            If TabCliente.State = 1 Then _
               TabCliente.Close
         End If
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A

      If TabCliente.State = 1 Then _
         TabCliente.Close
      SQL = "select * from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then _
         MOSTRA_DADOS
      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub cmbRegiao_LostFocus()
cmbRegiao.BackColor = &HFFFFFF
End Sub

Private Sub cmbStatus_LostFocus()
cmbStatus.BackColor = &HFFFFFF
End Sub

Private Sub cmbTipoCli_GotFocus()
cmbTipoCli.BackColor = &HC0FFFF
End Sub

Private Sub cmbTipoCli_LostFocus()
cmbTipoCli.BackColor = &HFFFFFF
End Sub

Private Sub cmbVendedor_LostFocus()
   cmbVendedor.BackColor = &HFFFFFF
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
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
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
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
End Sub

Private Sub cmdAtGlobal_Click()
'On Error GoTo ERRO_TRATA

   Msg = "Atenção, essa operação irá atualizar os dados cadastrais desse cliente no banco de dados NFe, confira se estão corretos os dados."
   PERGUNTA Msg, vbYesNo + 32, "Cadastro Cliente NFE", "DEMO.HLP", 1000
   Msg = ""
   If RESPOSTA = vbYes Then
      txtCNPJCPF.PromptInclude = False
      Call frmINTEGRA.INTEGRA_CLIENTE(txtCNPJCPF.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdAtGlobal_Click"
End Sub

Private Sub cmdExcluirFone_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Dim i As Integer

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" And txtN.Text <> "" Then
      Dim Achou As Boolean
      Achou = False
      Select Case Button.key
         Case "matar"
            FlexTel.Col = 1
            For i = 1 To FlexTel.Rows - 1
               If Replace(txtN.Text, "-", "") = FlexTel.TextMatrix(i, 1) Then
                  Achou = True
                  Exit For
               End If
            Next
            If Achou = True Then
               If FlexTel.Rows > 2 Then
                  FlexTel.RemoveItem (i)
                  Else
                     FlexTel.AddItem ""
                     FlexTel.RemoveItem (i)
                     'refresh_GRID
               End If
            End If
            txtCNPJCPF.PromptInclude = True
      End Select
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub PreencheTipoCliente()
'On Error GoTo ERRO_TRATA

   Dim rstTipoCliente As New ADODB.Recordset
   Dim strSQL As String
   
   If rstTipoCliente.State = 1 Then _
      rstTipoCliente.Close
   
   strSQL = "select * from TIPOCLIENTE WITH (NOLOCK) Order By TIPOCLIENTEID"
   rstTipoCliente.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   cmbTipoCli.Clear
   If Not rstTipoCliente.EOF Then
      rstTipoCliente.MoveFirst
      Do Until rstTipoCliente.EOF
         cmbTipoCli.AddItem rstTipoCliente!TIPOCLIENTEID & " - " & rstTipoCliente!Tipocliente
         rstTipoCliente.MoveNext
      Loop
   End If
   If rstTipoCliente.State = 1 Then _
      rstTipoCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub CmdCopiaEnderecoPessoal1_Click()
'On Error GoTo ERRO_TRATA

   txtCepB.PromptInclude = False
   txtCepB.Text = Replace(txtCepR.Text, "-", "")
   txtRuaB.Text = txtRuaR.Text
   txtEndB.Text = txtEndR.Text
   txtBaIrroB.Text = txtBairroR.Text
   txtCidadeB.Text = txtCidadeR.Text
   txtUFB.Text = txtUFR.Text
   txtNumeroB.Text = txtNumeroR.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub CmdCopiaEnderecoPessoal2_Click()
'On Error GoTo ERRO_TRATA

   txtCepC.PromptInclude = False
   txtCepC.Text = Replace(txtCepR.Text, "-", "")
   txtRuaC.Text = txtRuaR.Text
   txtEndC.Text = txtEndR.Text
   txtBairroC.Text = txtBairroR.Text
   txtCidadeC.Text = txtCidadeR.Text
   txtUFC.Text = txtUFR.Text
   txtNumeroC.Text = txtNumeroR.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmbRegiao_Click()
On Error Resume Next

   cmbAuxRegiao.ListIndex = cmbRegiao.ListIndex
End Sub

Private Sub cmbRegiao_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe a regiao que Cliente pertence", "", "", ""
   cmbRegiao.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmbStatus_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe situação cliente", "", "", ""
   cmbStatus.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmbRegiao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbStatus.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If CInt(Len(txtCNPJCPF.Text)) = 14 Then
         txtRazao.SetFocus
         Else: txtIE.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtBaIrroB_LostFocus()
txtBaIrroB.BackColor = &HFFFFFF
End Sub

Private Sub txtBairroC_LostFocus()
txtBairroC.BackColor = &HFFFFFF
End Sub

Private Sub txtBairroR_LostFocus()
txtBairroR.BackColor = &HFFFFFF
End Sub

Private Sub txtCepB_LostFocus()
txtCepB.BackColor = &HFFFFFF
End Sub

Private Sub txtCepC_LostFocus()
txtCepC.BackColor = &HFFFFFF
End Sub

Private Sub txtCepR_LostFocus()
txtCepR.BackColor = &HFFFFFF
End Sub

Private Sub txtCidadeR_LostFocus()
txtCidadeR.BackColor = &HFFFFFF
End Sub

Private Sub txtContato_GotFocus()
txtContato.BackColor = &HC0FFFF
End Sub

Private Sub txtContato_LostFocus()
txtContato.BackColor = &HFFFFFF
End Sub

Private Sub txtDDD_LostFocus()
txtDDD.BackColor = &HFFFFFF
End Sub

Private Sub txtDtNasc_LostFocus()
txtDtNasc.BackColor = &HFFFFFF
End Sub

Private Sub txtEndB_LostFocus()
txtEndB.BackColor = &HFFFFFF
End Sub

Private Sub txtEndC_LostFocus()
txtEndC.BackColor = &HFFFFFF
End Sub

Private Sub txtEndR_LostFocus()
txtEndR.BackColor = &HFFFFFF
End Sub

Private Sub txtibge_LostFocus()
txtIBGE.BackColor = &HFFFFFF
End Sub

Private Sub txtIE_LostFocus()
txtIE.BackColor = &HFFFFFF

   If Trim(txtIE.Text) <> "" Then
      txtIE.Text = Replace(txtIE.Text, "-", "")
      txtIE.Text = Replace(txtIE.Text, ",", "")
      txtIE.Text = Replace(txtIE.Text, ".", "")
      txtIE.Text = Replace(txtIE.Text, "/", "")
      txtIE.Text = Replace(txtIE.Text, "\", "")
   End If

End Sub

Private Sub txtIM_LostFocus()
txtIM.BackColor = &HFFFFFF

   If Trim(txtIM.Text) <> "" Then
      txtIM.Text = Replace(txtIM.Text, "-", "")
      txtIM.Text = Replace(txtIM.Text, ",", "")
      txtIM.Text = Replace(txtIM.Text, ".", "")
      txtIM.Text = Replace(txtIM.Text, "/", "")
      txtIM.Text = Replace(txtIM.Text, "\", "")
   End If

End Sub

Private Sub txtL_LostFocus()
txtL.BackColor = &HFFFFFF
End Sub

Private Sub txtN_LostFocus()
txtN.BackColor = &HFFFFFF
End Sub

Private Sub txtNome_LostFocus()
txtNome.BackColor = &HFFFFFF
End Sub

Private Sub txtNumeroB_LostFocus()
txtNumeroB.BackColor = &HFFFFFF
End Sub

Private Sub txtNumeroC_LostFocus()
txtNumeroC.BackColor = &HFFFFFF
End Sub

Private Sub txtNumeroR_LostFocus()
txtNumeroR.BackColor = &HFFFFFF
End Sub

Private Sub txtPercConv_LostFocus()
txtPercConv.BackColor = &HFFFFFF
End Sub

Private Sub txtRazao_GotFocus()
   txtRazao.BackColor = &HC0FFFF
End Sub

Private Sub TXTRAZAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtIE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRazao_KeyPress"
End Sub

Private Sub cmbVENDEDOR_Click()
   On Error Resume Next
   cmbVendedor.BackColor = &HC0FFFF

   cmbAuxVendedor.ListIndex = cmbVendedor.ListIndex
End Sub

Private Sub cmbVENDEDOR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Selecione Vendedor", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmbvendedor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercConv.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmbTipoCli_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtIE_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe nº da Inscrição Estadual", "", "", ""
   txtIE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtPercConv_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe Percentual de Desconto Convenio", "", "", ""
   txtPercConv.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercConv_GotFocus"
End Sub

Private Sub txtPercConv_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtContato.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtIM_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe nº da Inscrição Municipal", "", "", ""
   txtIM.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub LISTAITEM_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Duplo Click selecionar itens", "Click mostra itens", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub chkESTRANGEIRO_Click()
'On Error GoTo ERRO_TRATA
    
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   If SSTab.Tab = 0 Then
      txtCNPJCPF.Enabled = True
      txtCNPJCPF.SetFocus
   End If

   If SSTab.Tab = 1 Then _
      txtCepB.SetFocus

   If SSTab.Tab = 2 Then _
      txtCepC.SetFocus

   If SSTab.Tab = 3 Then
      fraTel.Visible = False
      FlexTel.Visible = False

      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         CONSULTA_VENDAS_CLIENTE
         CONSULTA_LANÇAMENTOS
         MOSTRA_CONTAS_CORRENTE
      End If
      txtCNPJCPF.PromptInclude = True

      VALOR_TOTAL_N = 0
      Else
         fraTel.Visible = True
         FlexTel.Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub ssTab_Validate(Cancel As Boolean)
'On Error GoTo ERRO_TRATA

   If SSTab.Tab = 3 Then
      fraTel.Visible = False
      FlexTel.Visible = False
      Else
         fraTel.Visible = True
         FlexTel.Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtBaIrroB_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o bairro", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtBairroC_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o bairro", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtBairroR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o bairro", "", "", ""
   txtBairroR.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCidadeB_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe a cidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCidadeC_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe a cidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCidadeR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe a cidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtContato_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTipoCli.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDDD_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o DDD", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDtCad_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbVendedor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDTNasc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtContato.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtIM.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbVendedor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtEndB_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o endereço", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtEndC_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o endereço", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtEndR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o endereço", "", "", ""
   txtEndR.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtL_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe descrição", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "cmdExcluirFone_Click"
End Sub

Private Sub txtLIMITE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLIMITE_KeyPress"
End Sub

Private Sub txtLIMITE_LostFocus()
    txtLIMITE.Text = Format(txtLIMITE.Text)
End Sub

Private Sub txtN_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o número de telefone", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtL.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe nome do Cliente", "", "", ""
   txtNome.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtInscSuframa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyPress"
End Sub

Private Sub txtRazao_LostFocus()
txtRazao.BackColor = &HFFFFFF
End Sub

Private Sub txtRuaB_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o nome da rua ou logradouro", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtRuaB_LostFocus()
txtRuaB.BackColor = &HFFFFFF
End Sub

Private Sub txtRuaC_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o nome da rua ou logradouro", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub
Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)
   txtCNPJCPF.BackColor = &HC0FFFF

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

   MOSTRA_TOP "ESC - Sair", "F6 - Excluir Cliente", "F7 - Consultar Cliente", "Informe CPF do cliente", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         txtCNPJCPF.PromptInclude = False
         If txtCNPJCPF.Text <> "" Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select * from PEDIDO WITH (NOLOCK)"
            SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
            SQL = SQL & " and cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then
               If TabCliente.State = 1 Then _
                  TabCliente.Close

               MsgBox "Impossível excluir, cliente possue movimento de vendas."
               Exit Sub
               Else
                  PESSOA_ID_N = 0 & TabCliente.Fields("pessoa_id").Value
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
            Msg = ""
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
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
      Case vbKeyDelete
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            txtCNPJCPF.Mask = "##############"
      Case vbKeyBack
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            txtCNPJCPF.Mask = "##############"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.BackColor = &HFFFFFF

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" Then
      PESSOA_ID_N = 0
      lblPessoa_id.Caption = PESSOA_ID_N
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         LIMPA_QUASE_TUDO

         MOSTRA_DADOS
      End If
   End If

   txtCNPJCPF.PromptInclude = False
   If Len(txtCNPJCPF.Text) > 0 Then
      If CInt(Len(txtCNPJCPF.Text)) = 11 Then
         Label31.Visible = False
         txtRazao.Visible = False
         If Not ValidaCPF(txtCNPJCPF.Text) Then
            MsgBox "CPF com DV incorreto !!!"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF = ""
            SSTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      ElseIf CInt(Len(txtCNPJCPF.Text)) = 14 Then
         Label31.Visible = True
         txtRazao.Visible = True
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDtCad_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtCad.PromptInclude = False
   
   If txtDtCad.Text = "" Then _
      txtDtCad.Text = Date
   txtDtCad.PromptInclude = True
   

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtDtNasc_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe data de nascimento", "", "", ""
   txtDtNasc.BackColor = &HC0FFFF

   txtDtNasc.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      If txtN.Text <> "" And txtCNPJCPF.Text <> "" Then _
         GRAVA_FONE_TEMP
      txtCNPJCPF.PromptInclude = True
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbRegiao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtRuaC_LostFocus()
txtRuaC.BackColor = &HFFFFFF
End Sub

Private Sub txtRuaR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o nome da rua ou logradouro", "", "", ""
   txtRuaR.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtNumeroR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o número", "", "", ""
   txtNumeroR.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtRuaR_LostFocus()
txtRuaR.BackColor = &HFFFFFF
End Sub

Private Sub txtUFB_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o estado", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtUFC_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o estado", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtUFB_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = vbKeyReturn Then _
      txtDDD.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtUFC_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = vbKeyReturn Then _
      txtDDD.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtUFR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Informe o estado", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtUFR_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = vbKeyReturn Then _
      txtDDD.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub
'====================residencial
Private Sub txtCepR_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCepR.PromptInclude = True
txtCepR.BackColor = &HC0FFFF
   MOSTRA_TOP "ESC - Sair", "F4 - Cadastra Cep", "F7 - Consulta Cep", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtCepR_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO_A = ""
         frmCADASTROCEP.Show 1
         txtCepR.PromptInclude = False
         txtCepR.Text = CRITERIO_A
         txtCepR.PromptInclude = True
         CRITERIO_A = ""
      Case vbKeyF7
         CRITERIO_A = ""
         frmCONSULTACEP.Show 1
         txtCepR.PromptInclude = False
         txtCepR.Text = CRITERIO_A
         txtCepR.PromptInclude = True
         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcepr_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcepr_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtruar_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtruar_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub txtufr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub
'================cobran
Private Sub txtCepB_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "F4 - Cadastra Cep", "F7 - Consulta Cep", "", ""
   
   txtCepB.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepb_GotFocus"
End Sub

Private Sub txtCepb_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO_A = ""
         frmCADASTROCEP.Show 1
         txtCepB.PromptInclude = False
         txtCepB.Text = CRITERIO_A
         txtCepB.PromptInclude = True
         CRITERIO_A = ""
      Case vbKeyF7
         CRITERIO_A = ""
         frmCONSULTACEP.Show 1
         txtCepB.PromptInclude = False
         txtCepB.Text = CRITERIO_A
         txtCepB.PromptInclude = True
         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepb_KeyDown"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcepb_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtruab_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtendb_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtbairrob_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcidadeb_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtufb_KeyPress"
End Sub
'============================comercial
Private Sub txtCepC_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "F4 - Cadastra Cep", "F7 - Consulta Cep", "", ""
   
   txtCepC.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepc_GotFocus"
End Sub

Private Sub txtCepc_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO_A = ""
         frmCADASTROCEP.Show 1
         txtCepC.PromptInclude = False
         txtCepC.Text = CRITERIO_A
         txtCepC.PromptInclude = True
         CRITERIO_A = ""
      Case vbKeyF7
         CRITERIO_A = ""
         frmCONSULTACEP.Show 1
         txtCepC.PromptInclude = False
         txtCepC.Text = CRITERIO_A
         txtCepC.PromptInclude = True
         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepc_KeyDown"
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
               txtIBGE.Text = TabCEP!IBGE_ID
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
      txtRuaC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcepc_KeyPress"
End Sub

Private Sub txtruac_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtruac_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtendc_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtbairroc_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtcidadec_KeyPress"
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
   TRATA_ERROS Err.Description, Me.Name, "txtufc_KeyPress"
End Sub

Private Sub cmbBanco_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Selecione um banco", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_GotFocus"
End Sub

Private Sub cmbbanco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbbanco_KeyPress"
End Sub

Private Sub cmbBanco_LostFocus()
'On Error GoTo ERRO_TRATA

   cmbAUXB.ListIndex = cmbBanco.ListIndex
   If cmbAUXB.Text <> "" Then
      If cmbAUXA.Text <> "" Then
         CRITERIO_A = cmbAUXA.Text
         NOME_A = cmbAgencia.Text
      End If
      cmbAgencia.Clear
      cmbAUXA.Clear
      MONTA_AGENCIA
      If CRITERIO_A <> "" Then
         cmbAUXA.Text = CRITERIO_A
         cmbAgencia.Text = NOME_A
      End If
   End If
   NOME_A = ""
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_lostfocus"
End Sub

Private Sub cmbAgencia_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Selecione uma agencia", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAgencia_GotFocus"
End Sub

Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbagencia_KeyPress"
End Sub

Private Sub cmbAgencia_LostFocus()
'On Error GoTo ERRO_TRATA

   cmbAUXA.ListIndex = cmbAgencia.ListIndex
   If cmbAUXA.Text <> "" Then
      If cmbAuxConta <> "" Then
         CRITERIO_A = cmbAuxConta.Text
         NOME_A = cmbNumr_Conta.Text
      End If
      cmbNumr_Conta.Clear
      cmbAuxConta.Clear
      MONTA_CONTA
      If NOME_A <> "" Then
         cmbAuxConta.Text = CRITERIO_A
         cmbNumr_Conta.Text = NOME_A
      End If
   End If
   NOME_A = ""
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbagencia_LostFocus"
End Sub

Private Sub cmbNumr_Conta_Click()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "F2 - Visualizar Posição Atual Conta Corrente", "", "", ""

   cmbAuxConta.ListIndex = cmbNumr_Conta.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbNumr_Conta_Click"
End Sub

Private Sub cmbNumr_Conta_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - Sair", "Selecione uma Conta Corrente", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbNumr_Conta_GotFocus"
End Sub

Private Sub cmbnumr_conta_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbnumr_conta_KeyPress"
End Sub

Private Sub listaitem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTAITEM, ColumnHeader
End Sub

Private Sub listaitem_Click()
'On Error GoTo ERRO_TRATA

   On Error Resume Next
   If Not IsNull(LISTAITEM.SelectedItem.Text) Then
      If Indr_Consulta = False Then
         LISTAITEM.Visible = True
         SETA_GRID2
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "listaitem_Click"
End Sub

Private Sub FlexTel_Click()
'On Error GoTo ERRO_TRATA

   FlexTel.Col = 0
   txtDDD.Text = "" & FlexTel.Text

   FlexTel.Col = 1
   txtN.Text = "" & FlexTel.Text

   FlexTel.Col = 2
   txtL.Text = "" & FlexTel.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FlexTel_Click"
End Sub

Private Sub cmdOBS_Click()
   txtOBS.Text = txtObs2.Text
End Sub
'==========================subrotinas
Private Sub refresh_GRID()
'On Error GoTo ERRO_TRATA

   FlexTel.Clear
   FlexTel.Row = 0
   FlexTel.Col = 0: FlexTel.ColWidth(0) = (FlexTel.Width / 8) - 100: FlexTel.Text = "DDD": FlexTel.ColAlignment(0) = 3
   FlexTel.Col = 1: FlexTel.ColWidth(1) = FlexTel.Width / 4: FlexTel.Text = "NÚMERO": FlexTel.ColAlignment(1) = 1
   FlexTel.Col = 2: FlexTel.ColWidth(2) = FlexTel.Width / 1.65: FlexTel.Text = "LOCAL"
   FlexTel.Col = 3: FlexTel.ColWidth(3) = 0: FlexTel.Text = "CNPJCPF"
   FlexTel.Rows = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REFRESH_GRID"
End Sub

Private Sub LIMPA_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   txtNome.Text = ""
   txtDtNasc.PromptInclude = False
   txtDtCad.PromptInclude = False
   txtDtNasc.Text = ""
   txtDtCad.Text = ""
   cmbStatus.Text = ""
   txtIE.Text = ""
   txtIM.Text = ""
   txtIBGE.Text = ""
   txtCepR.PromptInclude = False
   txtCepR.Text = ""
   txtRuaR.Text = ""
   txtNumeroR.Text = ""
   txtEndR.Text = ""
   txtBairroR.Text = ""
   txtCidadeR.Text = ""
   txtUFR.Text = ""
   txtCepC.PromptInclude = False
   txtCepC.Text = ""
   txtRuaC.Text = ""
   txtNumeroC.Text = ""
   txtEndC.Text = ""
   txtBairroC.Text = ""
   txtCidadeC.Text = ""
   txtUFC.Text = ""
   txtCepB.PromptInclude = False
   txtCepB.Text = ""
   txtRuaB.Text = ""
   txtNumeroB.Text = ""
   txtEndB.Text = ""
   txtBaIrroB.Text = ""
   txtCidadeB.Text = ""
   txtUFB.Text = ""
   LIMPA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtIM.Text = ""
   txtInscSuframa.Text = ""
   txtRazao.Text = ""
   cmbTipoCli.Text = ""
   txtPercConv.Text = "00,00"
   chkESTRANGEIRO.Value = 0
   chkESTRANGEIRO.ForeColor = vbBlack
   txtOBS.Text = ""
   VALOR_TOTAL_N = 0
   LISTAITEM.ListItems.Clear
   LISTAASS.ListItems.Clear
   txtSaldoDevedor.Text = ""
   cmbRegiao.Text = ""
   cmbAuxRegiao.Text = ""
   txtContato.Text = ""
   txtDtCad.PromptInclude = False
   txtDtCad.Text = ""
   txtDtNasc.PromptInclude = False
   txtDtNasc.Text = ""
   cmbVendedor.Text = ""
   cmbAuxVendedor.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCNPJCPF.Mask = "##############"
   txtNome.Text = ""
   cmbStatus.Text = ""
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
   txtCepB.PromptInclude = False
   txtCepB.Text = ""
   txtRuaB.Text = ""
   txtEndB.Text = ""
   txtBaIrroB.Text = ""
   txtCidadeB.Text = ""
   txtUFB.Text = ""
   LIMPA_FONE
   CRITERIO_A = 0
   txtLIMITE.Text = ""
   GeraCodigo
   SSTab.Tab = 0
   PESSOA_ID_N = 0
   lblPessoa_id.Caption = PESSOA_ID_N
   SETA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub MOSTRA_DADOS()
'On Error GoTo ERRO_TRATA

   Dim i As Integer

   PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
   lblPessoa_id.Caption = PESSOA_ID_N

   chkSuframa.Value = 0
   txtInscSuframa.Text = "" & TabCliente.Fields("codg_suframa").Value
   If Not IsNull(txtInscSuframa.Text) Then
      If Trim(txtInscSuframa.Text) <> "" Then
         If IsNumeric(txtInscSuframa.Text) Then
            If Len(txtInscSuframa.Text) > 4 Then
               chkSuframa.Value = 1
            End If
         End If
      End If
   End If

   txtCNPJCPF.PromptInclude = False
   Label4.Visible = False
   txtDtNasc.Visible = False
   chkESTRANGEIRO.Visible = False
   cmdRg.Visible = False

   Label31.Visible = False
   txtRazao.Visible = False

   If Len(txtCNPJCPF.Text) <= 11 Then
      Label4.Visible = True
      txtDtNasc.Visible = True
      chkESTRANGEIRO.Visible = True
      cmdRg.Visible = True
      Else
         Label31.Visible = True
         txtRazao.Visible = True
   End If

   If Not IsNull(TabCliente.Fields("estrangeiro").Value) Then
      chkESTRANGEIRO.ForeColor = vbBlack
      If TabCliente!ESTRANGEIRO = True Then
         chkESTRANGEIRO.Value = 1
         chkESTRANGEIRO.ForeColor = vbRed
         Else
            chkESTRANGEIRO.Value = 0
            chkESTRANGEIRO.ForeColor = vbBlue
      End If
   End If

   chkESTRANGEIRO.ForeColor = vbBlack
   If TabCliente!ESTRANGEIRO = True Then
      chkESTRANGEIRO.Value = 1
      chkESTRANGEIRO.ForeColor = vbRed
      Else
         chkESTRANGEIRO.Value = 0
         chkESTRANGEIRO.ForeColor = vbBlue
   End If

   txtCodigo.Text = TabCliente!CLIENTE_ID
   txtNome.Text = "" & Trim(TabCliente!NOME)
   txtRazao.Text = "" & Trim(TabCliente.Fields("razao_social").Value)
   If IsNumeric(TabCliente!PERC_DESC_CONVENIO) Then
      txtPercConv.Text = TabCliente!PERC_DESC_CONVENIO
      Else: txtPercConv.Text = "00.00"
   End If

'PESSOA
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select razao from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtRazao.Text = Trim(TabTemp.Fields(0).Value)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Not IsNull(TabCliente!DT_CAD) Then
      If IsDate(TabCliente!DT_CAD) Then
         txtDtCad.PromptInclude = False
            txtDtCad.Text = TabCliente!DT_CAD
         txtDtCad.PromptInclude = True
      End If
   End If

   If Not IsNull(TabCliente!LIMITE_CREDITO) Then _
      txtLIMITE.Text = Format(TabCliente!LIMITE_CREDITO, strFormatacao2Digitos)

   txtDtNasc.PromptInclude = False

   If IsDate(TabCliente!DT_NASC) Then _
      txtDtNasc.Text = TabCliente!DT_NASC

   If Not IsNull(TabCliente!CONTATO) Then _
      txtContato.Text = TabCliente!CONTATO

   If Not IsNull(TabCliente!REGIAO) Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from DESCR WITH (NOLOCK)"
      SQL = SQL & " where TIPO = 'R' "
      SQL = SQL & " and codigo = '" & Trim(TabCliente!REGIAO) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp!DESCRICAO) Then
            cmbRegiao.Text = Trim(TabTemp!DESCRICAO)
            cmbAuxRegiao.Text = Trim(TabTemp!codigo)
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   If Not IsNull(TabCliente!VENDEDOR_ID) Then
      If IsNumeric(TabCliente!VENDEDOR_ID) Then
         If TabEQUIPE.State = 1 Then _
            TabEQUIPE.Close

         SQL = "select vendedor_id,descricao from vwVendedor WITH (NOLOCK)"
         SQL = SQL & " where vendedor_id = " & TabCliente!VENDEDOR_ID
         TabEQUIPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabEQUIPE.EOF Then
            cmbAuxVendedor.Text = TabEQUIPE!VENDEDOR_ID
            cmbVendedor.Text = TabEQUIPE!DESCRICAO
         End If
         If TabEQUIPE.State = 1 Then _
            TabEQUIPE.Close
      End If
   End If

   If Not IsNull(TabCliente!TIPO_CLIENTE) Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from TIPOCLIENTE WITH (NOLOCK) Where TIPOCLIENTEID=" & TabCliente!TIPO_CLIENTE & " Order by TIPOCLIENTEID"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         cmbTipoCli.Text = TabCliente!TIPO_CLIENTE & " - " & TabConsulta!Tipocliente
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

   txtIE.Text = "" & TRAZ_IE(PESSOA_ID_N)
   txtIM.Text = "" & TRAZ_IM(PESSOA_ID_N)

'FONE
   txtCNPJCPF.PromptInclude = False
   CRITERIO_A = txtCNPJCPF.Text
   SETA_FONE

'ENDEREÇO RESIDENCIAL
   txtCNPJCPF.PromptInclude = False
'ok
   BUSCA_ENDERECO_PESSOA "R", ""
   If Not tabEndereco.EOF Then
      txtNumeroR.Text = tabEndereco!Numero & ""
      txtRuaR.Text = tabEndereco!Rua & ""
      txtBairroR.Text = tabEndereco!Bairro & ""
      txtEndR.Text = tabEndereco!Complemento & ""
      If Not IsNull(tabEndereco!CEP_id) Then
         If tabEndereco!CEP_id <> "" Then
            txtCepR.Text = tabEndereco!CEP_id
            SP_PROCURA_CEP tabEndereco!CEP_id
            If Not TabCEP.EOF Then
               txtCepR.Text = tabEndereco!CEP_id & ""
               txtCidadeR.Text = TabCEP!Cidade & ""
               txtUFR.Text = TabCEP!UF & ""
               txtIBGE.Text = TabCEP!IBGE_ID & ""
            End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

'ENDEREÇO COMERCIAL
'ok
   BUSCA_ENDERECO_PESSOA "C", ""
   If Not tabEndereco.EOF Then
      If Not IsNull(tabEndereco!Rua) Then _
         txtNumeroC.Text = tabEndereco!Numero & ""

         txtRuaC.Text = tabEndereco!Rua & ""
         txtBairroC.Text = tabEndereco!Bairro & ""
         txtEndC.Text = tabEndereco!Complemento & ""
      If Not IsNull(tabEndereco!CEP_id) Then
         If tabEndereco!CEP_id <> "" Then
            txtCepC.PromptInclude = False
            txtCepC.Text = tabEndereco!CEP_id

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from CEP WITH (NOLOCK)"
            SQL = SQL & " where cep_ID = '" & tabEndereco!CEP_id & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCidadeC.Text = TabConsulta!Cidade
               txtUFC.Text = TabConsulta!UF
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

 'ENDEREÇO COBRANÇA
 'ok
   BUSCA_ENDERECO_PESSOA "B", ""
   If Not tabEndereco.EOF Then
      If Not IsNull(tabEndereco!Rua) Then
         txtRuaB.Text = tabEndereco!Rua
      End If
      If Not IsNull(tabEndereco!Bairro) Then
         txtBaIrroB.Text = tabEndereco!Bairro
      End If
      If Not IsNull(tabEndereco!Complemento) Then
         txtEndB.Text = tabEndereco!Complemento
      End If
      txtNumeroB.Text = tabEndereco!Numero & ""
      If Not IsNull(tabEndereco!CEP_id) Then
         If tabEndereco!CEP_id <> "" Then
            txtCepB.PromptInclude = False
            txtCepB.Text = tabEndereco!CEP_id
            
            SP_PROCURA_CEP tabEndereco!CEP_id
            If Not TabCEP.EOF Then
               txtCidadeB.Text = TabCEP!Cidade
               txtUFB.Text = TabCEP!UF
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and seq = 0 "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta!Obs) Then _
         txtOBS.Text = TabConsulta!Obs
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   NUMR_SEQ_N = 0

   txtOBS.Text = "" & Trim(TabCliente.Fields("obs").Value)

   If Not IsNull(TabCliente!Status) Then
      If TabCliente!Status = "A" Then _
         cmbStatus.Text = "Ativo"
         
      If TabCliente!Status = "E" Then _
         cmbStatus.Text = "Especial"

      txtNome.ForeColor = vbBlack
      txtNome.FontBold = True

      If TabCliente!Status = "C" Then
         cmbStatus.Text = "Cancelado"
         txtNome.ForeColor = vbRed
         txtNome.FontBold = True
         'MsgBox "Cliente foi cancelado."
      End If
      Else: cmbStatus.Text = "Ativo"
   End If

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from FONE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   refresh_GRID
   Do While Not TabAUX.EOF
      FlexTel.AddItem ""
      FlexTel.Row = FlexTel.Rows - 1
      FlexTel.Col = 0
      FlexTel.Text = TabAUX!DDD & ""
      FlexTel.Col = 1
      FlexTel.Text = TabAUX!Numero
      FlexTel.Col = 2
      FlexTel.Text = TabAUX!local & ""
      FlexTel.Col = 3
      txtCNPJCPF.PromptInclude = False
      FlexTel.Text = txtCNPJCPF.Text
      txtCNPJCPF.PromptInclude = True
      TabAUX.MoveNext
   Loop
   If TabAUX.State = 1 Then _
      TabAUX.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_FONE"
End Sub

Private Sub LIMPA_FONE()
'On Error GoTo ERRO_TRATA

   txtN.Text = ""
   txtDDD.Text = ""
   txtL.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_FONE"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   Dim strTipoCliente As Integer
   Dim LimiteCredito As Double
   Dim strStatus As String
   Dim booEstrangeiro As Integer

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe CPF do cliente."
      SSTab.Tab = 0
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   If Trim(txtNome.Text) = "" Then
      MsgBox "Informe Nome do cliente."
      SSTab.Tab = 0
      txtNome.SetFocus
      Exit Sub
   End If

   If txtUFR.Text = "" Then
      MsgBox "Informe a UF do cliente."
      SSTab.Tab = 0
      'txtUFR.SetFocus
      'Exit Sub
   End If

   If txtCidadeR.Text = "" Then
      MsgBox "Informe a Cidade do cliente."
      'ssTab.Tab = 0
      'txtCidadeR.SetFocus
      'Exit Sub
   End If

   If txtRuaR.Text = "" Then
      MsgBox "Informe a Cidade do cliente."
      'ssTab.Tab = 0
      'txtRuaR.SetFocus
      'Exit Sub
   End If

   If txtBairroR.Text = "" Then
      MsgBox "Informe o Bairro do cliente."
      'ssTab.Tab = 0
      'txtBairroR.SetFocus
      'Exit Sub
   End If

   If txtIBGE.Text = "" Then
      MsgBox "Informe o Codigo do IBGE da Cidade do cliente."
      'ssTab.Tab = 0
      'txtIbge.SetFocus
      'Exit Sub
   End If

   If txtN.Text <> "" Then
      If Len(txtN.Text) < 8 Then
         MsgBox "o Numero de Telefone tem que ser maior que 8 digitos !!! Exemplo : (32873267)"
         SSTab.Tab = 0
         txtIBGE.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtIE.Text) = "" Then
      txtIE.Text = "ISENTO"
      Else
         If Trim(txtIE.Text) <> "ISENTO" Then
           If Valida_Inscricao_Estadual(txtIE.Text, txtUFC.Text) <> 0 Then
              SSTab.Tab = 0
              txtIE.SetFocus
              Exit Sub
           End If
         End If
   End If

   If cmbTipoCli.Text <> "" Then _
      strTipoCliente = Mid(cmbTipoCli.Text, 1, 1)

   If Not IsNull(txtLIMITE.Text) Then _
      If txtLIMITE.Text <> "" Then _
         LimiteCredito = txtLIMITE.Text

   If cmbStatus.Text <> "" Then
      strStatus = Left(cmbStatus.Text, 1)
      Else: strStatus = "A"
   End If

   txtDtNasc.PromptInclude = True
   If IsDate(txtDtNasc.Text) Then
      strDtNac = txtDtNasc.Text
      Else: strDtNac = 0
   End If

   If chkESTRANGEIRO.Value = 1 Then
      booEstrangeiro = 1
      Else: booEstrangeiro = 0
   End If

   If txtContato.Text = "" Then _
      txtContato.Text = "CLIENTE"

   If Trim(cmbAuxRegiao.Text) = "" Then
      strRegiao = 0
      Else: strRegiao = cmbAuxRegiao.Text
   End If

'=========================
   PESSOA_ID_N = 0
   lblPessoa_id.Caption = PESSOA_ID_N

'PESSOA
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close
   lblPessoa_id.Caption = PESSOA_ID_N

'=========================
   CONT_N = 2
   If PESSOA_ID_N <= 0 Then _
      CONT_N = 1

   'executa stored procedure spPessoa
   spPessoa CONT_N, PESSOA_ID_N, Trim(txtCNPJCPF.Text), Trim(txtNome.Text), Trim(txtRazao.Text), Trim(strStatus)

   PESSOA_ID_N = 0
   lblPessoa_id.Caption = PESSOA_ID_N

'PESSOA
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCliente.EOF Then
      txtContato.Text = "CLIENTE"
      SQL = "INSERT INTO CLIENTE "
      SQL = SQL & " ("
         SQL = SQL & " cliente_id,DT_CAD,CGCCPF,NOME,razao_social,TIPO_CLIENTE,LIMITE_CREDITO,DT_NASC,CONTATO,REGIAO, "
         SQL = SQL & " VENDEDOR_id,estabelecimento_ID,ESTRANGEIRO,PERC_DESC_CONVENIO,STATUS,obs,PESSOA_ID,codg_suframa"
      SQL = SQL & " )"
      SQL = SQL & " VALUES ("
         SQL = SQL & MAX_ID("cliente_id", "cliente", "", "", "", "")
         SQL = SQL & ",'" & Now & "'"
         SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRazao.Text) & "'"
         SQL = SQL & "," & strTipoCliente
         SQL = SQL & "," & Str(LimiteCredito)
         SQL = SQL & ",'" & DMA(strDtNac) & "'"
         SQL = SQL & ",'" & txtContato.Text & "'"
         SQL = SQL & "," & strRegiao
         SQL = SQL & ",0" & cmbAuxVendedor.Text
         SQL = SQL & "," & ESTABELECIMENTO_ID_N
         SQL = SQL & "," & booEstrangeiro
         SQL = SQL & "," & Str(txtPercConv.Text)
         SQL = SQL & ",'" & strStatus & "'"
         SQL = SQL & ",'" & Trim(txtOBS.Text) & "'"
         SQL = SQL & "," & PESSOA_ID_N
         SQL = SQL & ",0" & txtInscSuframa.Text
      SQL = SQL & ")"
      Else
         SQL = "UPDATE CLIENTE SET "
         SQL = SQL & " NOME = '" & Trim(txtNome.Text) & "'"
         SQL = SQL & ", razao_social = '" & Trim(txtRazao.Text) & "'"
         SQL = SQL & ", TIPO_CLIENTE = " & strTipoCliente
         SQL = SQL & ", LIMITE_CREDITO = " & tpMOEDA(LimiteCredito)
         SQL = SQL & ", DT_NASC = '" & DMA(strDtNac) & "'"
         SQL = SQL & ", CONTATO = '" & txtContato.Text & "'"
         SQL = SQL & ", REGIAO = " & strRegiao
         SQL = SQL & ", VENDEDOR_id = 0" & cmbAuxVendedor
         SQL = SQL & ", ESTRANGEIRO = " & booEstrangeiro
         SQL = SQL & ", PERC_DESC_CONVENIO = " & Str(txtPercConv.Text)
         SQL = SQL & ", Status = '" & strStatus & "'"
         SQL = SQL & ", obs = '" & Trim(txtOBS.Text) & "'"
         SQL = SQL & ", pessoa_id = " & PESSOA_ID_N
         SQL = SQL & ", codg_suframa = 0" & txtInscSuframa.Text
         SQL = SQL & ", estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

   CONECTA_RETAGUARDA.Execute SQL

   SQL = "Delete IE "
   SQL = SQL & " where PESSOA_ID = " & PESSOA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "Delete IM "
   SQL = SQL & " where PESSOA_ID = " & PESSOA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

'ENDEREÇO RESIDENCIAL
   txtCepR.PromptInclude = False
   If Not IsNumeric(txtIBGE.Text) Then _
      txtIBGE.Text = "5201211"

SP_MATA_ENDEREÇO "R"
   If Trim(txtCepR.Text) <> "" Or Trim(txtRuaR.Text) <> "" Or Trim(txtBairroR.Text) <> "" Or Trim(txtEndR.Text) <> "" Then
      sp_Grava_Endereco Trim(txtCepR.Text), Trim(txtRuaR.Text), Trim(txtBairroR.Text), Trim(txtEndR.Text), "R", Trim(txtNumeroR.Text)
   End If

'ENDEREÇO COMERCIAL
SP_MATA_ENDEREÇO "C"
   txtCepC.PromptInclude = False
   If Trim(txtCepC.Text) <> "" Or Trim(txtRuaC.Text) <> "" Or Trim(txtBairroC.Text) <> "" Or Trim(txtEndC.Text) <> "" Then
      sp_Grava_Endereco Trim(txtCepC.Text), Trim(txtRuaC.Text), Trim(txtBairroC.Text), Trim(txtEndC.Text), "C", Trim(txtNumeroC.Text)
   End If

'ENDEREÇO COBRANÇA
SP_MATA_ENDEREÇO "B"
   txtCepB.PromptInclude = False
   If txtCepB.Text <> "" Or txtRuaB.Text <> "" Or txtBaIrroB.Text <> "" Or txtEndB.Text <> "" Then
      sp_Grava_Endereco txtCepB.Text, txtRuaB.Text, txtBaIrroB.Text, txtEndB.Text, "B", txtNumeroB.Text
   End If

'GRAVA INSCRIÇÃO ESTADUAL
   'If Valida_Inscricao_Estadual(txtIE.Text, txtUFC.Text) <> 0 Then
      ENDERECO_ID_N = 0
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select ENDERECO_ID from ENDERECO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
      SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

      SQL = SQL & " where ENDERECO.pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and tipo = 'C'"

      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabEndereco.EOF Then _
         ENDERECO_ID_N = 0 & tabEndereco.Fields(0).Value
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      GRAVA_IE Trim(txtIE.Text)
   'End If

'FONE
   SQL = "Delete FONE "
   SQL = SQL & " where PESSOA_ID = " & PESSOA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   Dim i As Integer
   For i = 1 To FlexTel.Rows - 1
      FlexTel.Row = i
      SQL = "insert into FONE (pessoa_id,DDD,NUMERO,LOCAL) values ("
      SQL = SQL & PESSOA_ID_N                                                 'pessoa_id
      SQL = SQL & ",'" & FlexTel.TextMatrix(i, 0) & "'"                       'DDD
      SQL = SQL & ",'" & FlexTel.TextMatrix(i, 1) & "'"                       'NUMERO
      SQL = SQL & ",'" & Replace(FlexTel.TextMatrix(i, 2), "|", "/") & "')"   'LOCAL
      CONECTA_RETAGUARDA.Execute SQL
   Next

   txtCNPJCPF.PromptInclude = False
   Call frmINTEGRA.INTEGRA_CLIENTE(txtCNPJCPF.Text)

   CRITERIO_A = 0
   PEDIDO_ID_N = 0
   
   LIMPA_TUDO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub MONTA_AGENCIA()
'On Error GoTo ERRO_TRATA

   If TabAGENCIA.State = 1 Then _
      TabAGENCIA.Close

   SQL = "select * from AGENCIA a, BANCO b WITH (NOLOCK)"
   SQL = SQL & " where a.banco_id = b.banco_id "
   SQL = SQL & "and a.codg_banco = '" & cmbAUXB.Text & "' "
   SQL = SQL & "order by a.nome_agencia"
   TabAGENCIA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabAGENCIA.EOF
      cmbAgencia.AddItem TabAGENCIA!NUMR_AGENCIA & " - " & TabAGENCIA!nome_agencia
      cmbAUXA.AddItem TabAGENCIA!NUMR_AGENCIA
      TabAGENCIA.MoveNext
   Wend
   If TabAGENCIA.State = 1 Then _
      TabAGENCIA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_AGENCIA"
End Sub

Private Sub MONTA_CONTA()
'On Error GoTo ERRO_TRATA
    
   If TabCONTA.State = 1 Then _
      TabCONTA.Close
    
   SQL = "select * from CONTA c, AGENCIA a WITH (NOLOCK)"
   SQL = SQL & " where c.numr_agencia=a.numr_agencia "
   SQL = SQL & "and c.numr_agencia='" & cmbAUXA.Text & "' "
   SQL = SQL & "order by c.desc_conta"
   TabCONTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCONTA.EOF
      cmbNumr_Conta.AddItem TabCONTA!NUMR_CONTA & " - " & TabCONTA!DESC_CONTA
      cmbAuxConta.AddItem TabCONTA!NUMR_CONTA
      TabCONTA.MoveNext
   Wend
   If TabCONTA.State = 1 Then _
      TabCONTA.Close

Exit Sub
ERRO_TRATA:
    If Err.Number = -2147217865 Then Exit Sub
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONTA"
End Sub

Private Sub CONSULTA_VENDAS_CLIENTE()
'On Error GoTo ERRO_TRATA

   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id > 0 "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   'SQL = SQL & "and dt_req >= '" & now & "'"
   SQL = SQL & " and cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
   'SQL = SQL & " and status > 2 "
   SQL = SQL & " and status < 9 "

   SETA_GRID_VENDAS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_VENDAS_CLIENTE"
End Sub

Private Sub SETA_GRID_VENDAS()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_N = 0
   LISTAASS.ListItems.Clear
   VALR_SALDO_ATUAL_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = LISTAASS.ListItems.Add(, "seq." & TabTemp!PEDIDO_ID, TabTemp!PEDIDO_ID)

      VALOR_DESCONTO_N = 0

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabTemp!PEDIDO_ID
      SQL = SQL & " and tipo_reg = 'PC' "
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      'BUSCA VALOR TOTAL VENDA
      VALOR_ITEM_N = 0

      SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabTemp!PEDIDO_ID
      SQL = SQL & " and tipo_reg = 'PC' "
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      item.SubItems(1) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      item.SubItems(2) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
      VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N

      VALR_SALDO_ATUAL_N = VALR_SALDO_ATUAL_N + VALOR_TOTAL_N
      txtTotalVendas.Text = Format(VALR_SALDO_ATUAL_N, strFormatacao2Digitos)
      txtTotalVendas.Refresh

      item.SubItems(3) = TabTemp!dt_req

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " where tipovenda_id = " & TabTemp!TipoVenda_ID
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         item.SubItems(4) = TabDESCR!DESCRICAO
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      If TabUSU.State = 1 Then _
         TabUSU.Close

      SQL = "select descricao from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where vendedor_id = " & TabTemp!VENDEDOR_ID
      TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabUSU.EOF Then _
         item.SubItems(5) = TabUSU!DESCRICAO
      If TabUSU.State = 1 Then _
         TabUSU.Close

      If Not IsNull(TabTemp!Status) Then
         If TabTemp!Status = 1 Then _
            item.SubItems(6) = "ORÇAMENTO"
         If TabTemp!Status = 2 Then _
            item.SubItems(6) = "Pedido"
         If TabTemp!Status = 3 Then _
            item.SubItems(6) = "Pedido c/ Nota"
         If TabTemp!Status = 4 Then _
            item.SubItems(6) = "Pedido c/ Cupom"
         If TabTemp!Status = 9 Then _
            item.SubItems(6) = "Pedido Cancelada"
      End If
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_VENDAS"
End Sub

Private Sub CONSULTA_LANÇAMENTOS()
'On Error GoTo ERRO_TRATA

   SQL = "select distinct(i.numr_doc),* from ITEMLANCAMENTO i, LANCAMENTO l WITH (NOLOCK)"
   SQL = SQL & " where i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and l.pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and i.status = 'A' "

   SETA_GRID_LANCAMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_LANÇAMENTOS"
End Sub

Private Sub SETA_GRID_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "Duplo Click selecionar itens", "Click mostra itens", "", ""

   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 1
   LISTAITEM.ListItems.Clear

   'cmbaux.Clear
   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NUMR_SEQ_N = 1 + NUMR_SEQ_N
      Set item = LISTAITEM.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp!numr_doc)
      'cmbaux.AddItem TABTEMP!numr_doc

      item.SubItems(1) = Format(TabTemp!Valor_Item, strFormatacao2Digitos)
      VALOR_TOTAL_N = TabTemp!Valor_Item + VALOR_TOTAL_N
      txtSaldoDevedor.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtSaldoDevedor.Refresh

      If TabTemp!Status = "A" Then _
         item.SubItems(5) = "Aberto"
      If TabTemp!Status = "B" Then _
         item.SubItems(5) = "Baixado"
      If TabTemp!Status = "C" Then _
         item.SubItems(5) = "Cancelado"
      If Not IsNull(TabTemp!DT_VENCIMENTO) Then _
         item.SubItems(2) = TabTemp!DT_VENCIMENTO
      If Not IsNull(TabTemp!DT_BAIXA) Then
         If TabTemp!DT_BAIXA > 0 Then _
            item.SubItems(3) = TabTemp!DT_BAIXA
      End If
      item.SubItems(4) = TabTemp!DT_CAD
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_LANCAMENTO"
End Sub

Private Sub SETA_GRID2()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Duplo Click ocultar itens", "", "", ""

   LISTAITEM.ListItems.Clear

   If LISTAITEM.ListItems.Count = 0 Then _
      Exit Sub

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where numr_doc = " & LISTAITEM.SelectedItem.Text
   SQL = SQL & " and status = 'A' "
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLancamento.EOF
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where formapagto_id = " & TabLancamento!FORMAPAGTO_ID
      SQL = SQL & " and status = 'true' "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then
         Set ITEM2 = LISTAITEM.ListItems.Add(, "seq." & TabLancamento!seq, TabDESCR!DESCRICAO)
         Else: Set ITEM2 = LISTAITEM.ListItems.Add(, "seq." & TabLancamento!seq, "Nao Tem Forma Pgto")
      End If
      ITEM2.SubItems(1) = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
      
      ITEM2.SubItems(2) = TabLancamento!DT_VENCIMENTO
      
      If Not IsNull(TabLancamento!DT_BAIXA) Then
         ITEM2.SubItems(3) = TabLancamento!DT_BAIXA
         Else: ITEM2.SubItems(3) = ""
      End If
      
      If Not IsNull(TabLancamento!Status) Then
         If TabLancamento!Status = "A" Then _
            ITEM2.SubItems(4) = "Aberto"
         If TabLancamento!Status = "B" Then _
            ITEM2.SubItems(4) = "Baixado"
         If TabLancamento!Status = "C" Then _
            ITEM2.SubItems(4) = "Cancelado"
      End If

      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from OBS WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and seq = " & TabLancamento!seq
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then _
         If Not IsNull(TabAUX!Obs) Then _
            ITEM2.SubItems(5) = TabAUX!Obs
      If TabAUX.State = 1 Then _
         TabAUX.Close

      TabLancamento.MoveNext
   Wend
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID2"
End Sub

Private Sub MOSTRA_CONTAS_CORRENTE()
'On Error GoTo ERRO_TRATA

   SQL = "select * from CONTA c, AGENCIA a, BANCO b WITH (NOLOCK)"
   SQL = SQL & " where C.pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and c.numr_agencia=a.numr_agencia "
   SQL = SQL & "and a.banco_id = b.banco_id "

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CONTAS_CORRENTE"
End Sub

Private Sub GRAVA_CONTAS_CORRENTE()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CONTA WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and numr_conta = " & cmbAuxConta.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SqL2 = "INSERT INTO CONTA (dt_cadastro) "
      SqL2 = SqL2 & " VALUES ('" & Now & "')"
      CONECTA_RETAGUARDA.Execute SqL2
      Else
         SQL = "UPDATE CONTA SET dt_cadastro = '" & Now & "'"
         SQL = SQL & " where cgccpf = '" & txtCNPJCPF.Text & "'"
         SQL = SQL & " and numr_conta = " & cmbAuxConta.Text
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MOSTRA_CONTAS_CORRENTE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CONTAS_CORRENTE"
End Sub

Private Sub GeraCodigo()
'On Error GoTo ERRO_TRATA

   txtCodigo.Text = MAX_ID("cliente_id", "cliente", "", "", "", "")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GeraCodigo"
End Sub

Sub MONTA_DESCRITORES()
   cmbRegiao.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'R'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbRegiao.AddItem Trim(TabTemp!DESCRICAO)
      cmbAuxRegiao.AddItem Trim(TabTemp!codigo)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
End Sub

Sub MONTA_VENDEDORES()
   'EQUIPE
   cmbVendedor.Clear

   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close

   SQL = "select vendedor_id,descricao from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & "order by descricao"
   TabEQUIPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabEQUIPE.EOF
      cmbVendedor.AddItem Trim(TabEQUIPE!DESCRICAO)
      cmbAuxVendedor.AddItem Trim(TabEQUIPE!VENDEDOR_ID)
      TabEQUIPE.MoveNext
   Wend
   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Private Sub GRAVA_FONE_TEMP()
'On Error GoTo ERRO_TRATA

   Dim strAux As String * 2
   Dim Achou As Boolean
   Dim i As Integer
   Achou = False
   If SSTab.Tab = 0 Then
      strAux = "DP"
      Else
         If SSTab.Tab = 2 Then
            strAux = "DC"
            Else: strAux = "EC"
         End If
    End If
   txtCNPJCPF.PromptInclude = False
   FlexTel.Col = 1
   
   For i = 1 To FlexTel.Rows - 1
      FlexTel.Row = i
      If Replace(txtN.Text, "-", "") = FlexTel.Text Then
         Achou = True
         Exit For
      End If
   Next

   If Not Achou Then
      FlexTel.AddItem ""
      FlexTel.Row = FlexTel.Rows - 1
      FlexTel.Col = 0
      FlexTel.Text = txtDDD.Text
      FlexTel.Col = 1
      FlexTel.Text = Replace(txtN.Text, "-", "")
      FlexTel.Col = 2
      FlexTel.Text = strAux & " / " & txtL.Text
      FlexTel.Col = 3
      FlexTel.Text = txtCNPJCPF.Text
      Else
         FlexTel.Col = 0
         FlexTel.Text = txtDDD.Text
         FlexTel.Col = 1
         FlexTel.Text = Replace(txtN.Text, "-", "")
         FlexTel.Col = 2
         FlexTel.Text = strAux & " / " & txtL.Text
         FlexTel.Col = 3
         FlexTel.Text = txtCNPJCPF.Text
   End If
   FlexTel.Refresh
   LIMPA_FONE
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_FONE_TEMP"
End Sub

Sub MONTA_REL_CLI()
'On Error GoTo ERRO_TRATA

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Msg = "Relatório completo? "
   PERGUNTA Msg, vbYesNo + 32, "Cadastro Cliente", "DEMO.HLP", 1000
   Msg = ""
   If RESPOSTA = vbYes Then
      If EXISTE_OBJ_BANCO("RETAGUARDA", "REL_CLI_COMPLETO", "U") = True Then
         SQL = "drop table rel_cli_completo"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      SQL = "CREATE TABLE [dbo].[REL_CLI_COMPLETO]"
      SQL = SQL & "("
         SQL = SQL & " [PESSOA_ID] [bigint] NOT NULL,"
         SQL = SQL & " [ESTABELECIMENTO_ID] [int] NULL,"
         SQL = SQL & " [CNPJCPF] [nvarchar](50) NULL,"
         SQL = SQL & " [NOME] [nvarchar](50) NULL,"
         SQL = SQL & " [RAZAO] [nvarchar](50) NULL,"
         SQL = SQL & " [DT_NASC] [datetime] NULL,"
         SQL = SQL & " [SITUACAO] [nvarchar](10) NULL,"
         SQL = SQL & " [REGIAO] [nvarchar](20) NULL,"
         SQL = SQL & " [RG] [nvarchar](50) NULL,"
         SQL = SQL & " [ORGAO_RG] [nvarchar](50) NULL,"
         SQL = SQL & " [DT_EMIS_RG] [datetime] NULL,"
         SQL = SQL & " [IE] [nvarchar](50) NULL,"
         SQL = SQL & " [IM] [nvarchar](50) NULL,"
         SQL = SQL & " [CONVENIO] [nvarchar](50) NULL,"
         SQL = SQL & " [VENDEDOR_ID] [bigint] NULL,"
         SQL = SQL & " [NOME_VENDEDOR] [nvarchar](50) NULL,"
         SQL = SQL & " [QUEM_LIBEROU] [nvarchar](50) NULL,"
         SQL = SQL & " [TIPOCLIENTE] [nvarchar](50) NULL,"
         SQL = SQL & " [DT_CAD] [datetime] NULL,"
         SQL = SQL & " [cep_res] [nvarchar](20) NULL,"
         SQL = SQL & " [rua_res] [nvarchar](50) NULL,"
         SQL = SQL & " [numero_res] [nvarchar](50) NULL,"
         SQL = SQL & " [complemento_res] [nvarchar](50) NULL,"
         SQL = SQL & " [bairro_res] [nvarchar](50) NULL,"
         SQL = SQL & " [cidade_res] [nvarchar](50) NULL,"
         SQL = SQL & " [uf_res] [nvarchar](2) NULL,"
         SQL = SQL & " [ibge_res] [nvarchar](50) NULL,"
         SQL = SQL & " [cep_cob] [nvarchar](20) NULL,"
         SQL = SQL & " [rua_cob] [nvarchar](50) NULL,"
         SQL = SQL & " [numero_cob] [nvarchar](50) NULL,"
         SQL = SQL & " [complemento_cob] [nvarchar](50) NULL,"
         SQL = SQL & " [bairro_cob] [nvarchar](50) NULL,"
         SQL = SQL & " [cidade_cob] [nvarchar](50) NULL,"
         SQL = SQL & " [uf_cob] [nvarchar](2) NULL,"
         SQL = SQL & " [ibge_cob] [nvarchar](50) NULL,"
         SQL = SQL & " [cep_comer] [nvarchar](20) NULL,"
         SQL = SQL & " [rua_comer] [nvarchar](50) NULL,"
         SQL = SQL & " [numero_comer] [nvarchar](50) NULL,"
         SQL = SQL & " [complemento_comer] [nvarchar](50) NULL,"
         SQL = SQL & " [bairro_comer] [nvarchar](50) NULL,"
         SQL = SQL & " [cidade_comer] [nvarchar](50) NULL,"
         SQL = SQL & " [uf_comer] [nvarchar](2) NULL,"
         SQL = SQL & " [ibge_comer] [nvarchar](50) NULL,"
         SQL = SQL & " [OBS] [nvarchar](MAX) NULL,"
         SQL = SQL & " CONSTRAINT [PK_REL_CLI_COMPLETO] PRIMARY KEY CLUSTERED"
         SQL = SQL & " ([PESSOA_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & ") "
      SQL = SQL & "ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "delete from REL_CLI_COMPLETO"
      CONECTA_RETAGUARDA.Execute SQL

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE WITH (NOLOCK)"
      If PESSOA_ID_N > 0 Then _
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         txtDtNasc.PromptInclude = True
         If IsDate(txtDtNasc.Text) Then
            strDtNac = txtDtNasc.Text
            Else: strDtNac = 0
         End If

         SQL = "insert into REL_CLI_COMPLETO "
         SQL = SQL & " values("
         SQL = SQL & PESSOA_ID_N
         SQL = SQL & "," & ESTABELECIMENTO_ID_N
         SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRazao.Text) & "'"
         SQL = SQL & ",'" & DMA(strDtNac) & "'"
         SQL = SQL & ",'" & Trim(cmbStatus.Text) & "'"
         SQL = SQL & ",'" & Trim(cmbRegiao.Text) & "'"
         'SQL = SQL & ",'" & Trim(txtRg.Text) & "'"
         'SQL = SQL & ",'" & Trim(txtOrigem.Text) & "'"
         'SQL = SQL & ",'" & DMA(txtDTEXP.Text) & "'"
         SQL = SQL & ",'" & Trim(txtIE.Text) & "'"
         SQL = SQL & ",'" & Trim(txtIM.Text) & "'"
         SQL = SQL & "," & tpMOEDA(txtPercConv.Text)
         SQL = SQL & "," & cmbAuxVendedor.Text
         SQL = SQL & ",'" & Trim(cmbVendedor.Text) & "'"
         SQL = SQL & ",'" & Trim(txtContato.Text) & "'"
         SQL = SQL & ",'" & Trim(cmbTipoCli.Text) & "'"
         SQL = SQL & ",'" & DMA(txtDtCad.Text) & "'"

         SQL = SQL & ",'" & Trim(txtCepR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRuaR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtNumeroR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtEndR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtBairroR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtCidadeR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtUFR.Text) & "'"
         SQL = SQL & ",'" & Trim(txtIBGE.Text) & "'"

         SQL = SQL & ",'" & Trim(txtCepB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRuaB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtNumeroB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtEndB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtBaIrroB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtCidadeB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtUFB.Text) & "'"
         SQL = SQL & ",'" & Trim(txtIBGE.Text) & "'"

         SQL = SQL & ",'" & Trim(txtCepC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRuaC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtNumeroC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtEndC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtBairroC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtCidadeC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtUFC.Text) & "'"
         SQL = SQL & ",'" & Trim(txtIBGE.Text) & "'"

         SQL = SQL & ",'" & Trim(txtOBS.Text) & "'"
         SQL = SQL & " )"

         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      FORMULA_REL = "{REL_CLI_COMPLETO.cnpjcpf} = '" & Trim(txtCNPJCPF.Text) & "'"

      Nome_Relatorio = "rel_Cli_completo.rpt"
      Else
         FORMULA_REL = "{CLIENTE.empresa_id} = " & EMPRESA_ID_N

         If txtCNPJCPF.Text <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {CLIENTE.cgccpf} = '" & Trim(txtCNPJCPF.Text) & "'"

         'If txtCepR.Text <> "" Then _
            formula_rel = formula_rel & " and {CEP.Cep_ID} = '" & txtCepR.Text & "'"

         'If txtCidadeR.Text <> "" Then _
            formula_rel = formula_rel & " and {CEP.Cidade} = '" & txtCidadeR.Text & "'"

         'If txtUFR.Text <> "" Then _
            formula_rel = formula_rel & " and {CEP.uf} = '" & txtUFR.Text & "'"

         'If Left(cmbStatus.Text, 1) <> "" Then _
            formula_rel = formula_rel & " and {CLIENTE.Status} = '" & Left(cmbStatus.Text, 1) & "'"

         Nome_Relatorio = "rel_Cliente.rpt"
   End If
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL_CLI"
End Sub

Private Sub txtUFR_LostFocus()
txtUFR.BackColor = &HFFFFFF
End Sub
