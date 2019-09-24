VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROEMPRESA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Empresa"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   915
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROEMPRESA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab stbEmpresa 
      Height          =   7095
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "CADASTROEMPRESA.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Estabelecimento(s)"
      TabPicture(1)   =   "CADASTROEMPRESA.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCNPJCPF"
      Tab(1).Control(1)=   "txtNFCe"
      Tab(1).Control(2)=   "txtCSC"
      Tab(1).Control(3)=   "cmbCredAUX"
      Tab(1).Control(4)=   "cmbCred"
      Tab(1).Control(5)=   "txtDiasAtrazo"
      Tab(1).Control(6)=   "cmbIdEstabAUX"
      Tab(1).Control(7)=   "cmbIdEstab"
      Tab(1).Control(8)=   "txtNfe"
      Tab(1).Control(9)=   "txtInstruçãoBoleto"
      Tab(1).Control(10)=   "txtLoc"
      Tab(1).Control(11)=   "txtNomeEstab"
      Tab(1).Control(12)=   "txtDias"
      Tab(1).Control(13)=   "txtJuros"
      Tab(1).Control(14)=   "txtMSG"
      Tab(1).Control(15)=   "txtDescDesconto"
      Tab(1).Control(16)=   "txtDadosAdicionais"
      Tab(1).Control(17)=   "chkIndustria"
      Tab(1).Control(18)=   "chkDesconto"
      Tab(1).Control(19)=   "chkNFE"
      Tab(1).Control(20)=   "chkControleEstoque"
      Tab(1).Control(21)=   "chkEstoque"
      Tab(1).Control(22)=   "chkFatPedido"
      Tab(1).Control(23)=   "chkLEI_12741"
      Tab(1).Control(24)=   "chkBaixaEstPedido"
      Tab(1).Control(25)=   "chkMarkap"
      Tab(1).Control(26)=   "chkCli"
      Tab(1).Control(27)=   "chkFunc"
      Tab(1).Control(28)=   "chkPercDesconto"
      Tab(1).Control(29)=   "txtCNPJCRED"
      Tab(1).Control(30)=   "chkLimpaPedido"
      Tab(1).Control(31)=   "chkDoc_Fiscal"
      Tab(1).Control(32)=   "chkBloqFat"
      Tab(1).Control(33)=   "chkTabPreco"
      Tab(1).Control(34)=   "Label1(2)"
      Tab(1).Control(35)=   "Label26"
      Tab(1).Control(36)=   "lblCred"
      Tab(1).Control(37)=   "Label25"
      Tab(1).Control(38)=   "Label24"
      Tab(1).Control(39)=   "Label1(4)"
      Tab(1).Control(40)=   "Label1(3)"
      Tab(1).Control(41)=   "Label20"
      Tab(1).Control(42)=   "Label1(1)"
      Tab(1).Control(43)=   "Label19"
      Tab(1).Control(44)=   "Label7"
      Tab(1).Control(45)=   "Label17"
      Tab(1).Control(46)=   "Label12"
      Tab(1).Control(47)=   "Label11"
      Tab(1).ControlCount=   48
      TabCaption(2)   =   "Fone/Endereço"
      TabPicture(2)   =   "CADASTROEMPRESA.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtIM"
      Tab(2).Control(1)=   "txtIE"
      Tab(2).Control(2)=   "txtComp"
      Tab(2).Control(3)=   "txtL"
      Tab(2).Control(4)=   "txtN"
      Tab(2).Control(5)=   "txtDDD"
      Tab(2).Control(6)=   "txtBairro"
      Tab(2).Control(7)=   "txtUF"
      Tab(2).Control(8)=   "txtCidade"
      Tab(2).Control(9)=   "txtRua"
      Tab(2).Control(10)=   "txtNumero"
      Tab(2).Control(11)=   "txtCep"
      Tab(2).Control(12)=   "Toolbar_Fone"
      Tab(2).Control(13)=   "adoFone"
      Tab(2).Control(14)=   "grdFone"
      Tab(2).Control(15)=   "cmdEmail"
      Tab(2).Control(16)=   "cmdMataIE"
      Tab(2).Control(17)=   "Toolbar2"
      Tab(2).Control(18)=   "Label21"
      Tab(2).Control(19)=   "Label6"
      Tab(2).Control(20)=   "Label13"
      Tab(2).Control(21)=   "Label8"
      Tab(2).Control(22)=   "Label2"
      Tab(2).Control(23)=   "Label16"
      Tab(2).Control(24)=   "lblLabels(13)"
      Tab(2).Control(25)=   "Label15"
      Tab(2).Control(26)=   "Label14"
      Tab(2).Control(27)=   "Label10"
      Tab(2).Control(28)=   "Label9"
      Tab(2).Control(29)=   "Label1(0)"
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "GlobalEmpres"
      TabPicture(3)   =   "CADASTROEMPRESA.frx":5C66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtENVIONFE"
      Tab(3).Control(1)=   "txtAmbiente"
      Tab(3).Control(2)=   "txtFilial"
      Tab(3).Control(3)=   "txtEmpresa"
      Tab(3).Control(4)=   "txtVersaoNFe"
      Tab(3).Control(5)=   "cmdGravaNFe"
      Tab(3).Control(6)=   "Label1(8)"
      Tab(3).Control(7)=   "Label1(7)"
      Tab(3).Control(8)=   "Label1(6)"
      Tab(3).Control(9)=   "Label1(5)"
      Tab(3).Control(10)=   "Label18"
      Tab(3).ControlCount=   11
      Begin PVMaskEditLib.PVMaskEdit txtCNPJCPF 
         Height          =   375
         Left            =   -73560
         TabIndex        =   64
         Top             =   1140
         Width           =   2415
         _Version        =   524288
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         Text            =   ""
         Mask            =   "##.###.###/####-##"
      End
      Begin VB.TextBox txtENVIONFE 
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
         Left            =   -69720
         TabIndex        =   96
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtAmbiente 
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
         Left            =   -71400
         MaxLength       =   1
         TabIndex        =   93
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtFilial 
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
         Left            =   -72840
         MaxLength       =   2
         TabIndex        =   91
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtEmpresa 
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
         Left            =   -73920
         MaxLength       =   2
         TabIndex        =   89
         Text            =   "01"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtVersaoNFe 
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
         Left            =   -73680
         MaxLength       =   6
         TabIndex        =   87
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtNFCe 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   -71970
         TabIndex        =   85
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox txtCSC 
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
         Left            =   -68415
         TabIndex        =   81
         Top             =   6060
         Width           =   1575
      End
      Begin VB.ComboBox cmbCredAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   -69000
         TabIndex        =   79
         Top             =   6660
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbCred 
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
         Left            =   -71205
         TabIndex        =   77
         Top             =   6660
         Width           =   4455
      End
      Begin VB.TextBox txtDiasAtrazo 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   -69000
         MaxLength       =   4
         TabIndex        =   75
         Top             =   5580
         Width           =   735
      End
      Begin VB.ComboBox cmbIdEstabAUX 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -74880
         TabIndex        =   73
         Top             =   660
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbIdEstab 
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
         Left            =   -74880
         TabIndex        =   72
         Top             =   660
         Width           =   1215
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
         Height          =   375
         Left            =   -69165
         MaxLength       =   20
         TabIndex        =   65
         Top             =   1860
         Width           =   1815
      End
      Begin VB.TextBox txtNfe 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   -73560
         TabIndex        =   61
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox txtInstruçãoBoleto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   46
         Tag             =   " "
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   4740
         Width           =   3165
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
         Left            =   -73125
         MaxLength       =   20
         TabIndex        =   42
         Top             =   1860
         Width           =   1815
      End
      Begin VB.TextBox txtComp 
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
         Left            =   -70245
         MaxLength       =   50
         TabIndex        =   47
         Top             =   2820
         Width           =   3375
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
         Left            =   -72255
         MaxLength       =   30
         TabIndex        =   57
         Top             =   5100
         Width           =   4215
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
         Left            =   -73965
         MaxLength       =   30
         TabIndex        =   56
         Top             =   5100
         Width           =   1575
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
         Left            =   -74565
         MaxLength       =   2
         TabIndex        =   55
         Top             =   5100
         Width           =   495
      End
      Begin VB.TextBox txtBairro 
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
         Left            =   -73125
         MaxLength       =   50
         TabIndex        =   49
         Top             =   3300
         Width           =   6255
      End
      Begin VB.TextBox txtUF 
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
         Left            =   -73125
         MaxLength       =   2
         TabIndex        =   54
         Top             =   4260
         Width           =   615
      End
      Begin VB.TextBox txtCidade 
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
         Left            =   -70845
         MaxLength       =   50
         TabIndex        =   53
         Top             =   3780
         Width           =   3975
      End
      Begin VB.TextBox txtRua 
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
         Left            =   -73125
         MaxLength       =   50
         TabIndex        =   44
         Top             =   2340
         Width           =   6255
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   -73125
         MaxLength       =   50
         TabIndex        =   45
         Top             =   2820
         Width           =   975
      End
      Begin VB.TextBox txtLoc 
         Height          =   360
         Left            =   -69645
         TabIndex        =   20
         Top             =   1140
         Width           =   2775
      End
      Begin VB.TextBox txtNomeEstab 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -73605
         TabIndex        =   19
         Top             =   660
         Width           =   6735
      End
      Begin VB.TextBox txtDias 
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
         Left            =   -71040
         TabIndex        =   18
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox txtJuros 
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
         Left            =   -73560
         TabIndex        =   16
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox txtMSG 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "CADASTROEMPRESA.frx":5C82
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   2820
         Width           =   3165
      End
      Begin VB.TextBox txtDescDesconto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1980
         Width           =   3165
      End
      Begin VB.TextBox txtDadosAdicionais 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   " "
         Text            =   "CADASTROEMPRESA.frx":5CE6
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   3660
         Width           =   3165
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   50
         TabIndex        =   8
         Top             =   360
         Width           =   8175
         Begin PVMaskEditLib.PVMaskEdit txtCNPJ 
            Height          =   375
            Left            =   1920
            TabIndex        =   0
            Top             =   360
            Width           =   2655
            _Version        =   524288
            _ExtentX        =   4683
            _ExtentY        =   661
            _StockProps     =   253
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Text            =   ""
            Mask            =   "##.###.###/####-##"
            CopyTextAndMask =   1
         End
         Begin VB.TextBox txtCOFINS 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   100
            Top             =   3000
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtPIS 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   98
            Top             =   2520
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbCRTAUX 
            BackColor       =   &H80000000&
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
            Left            =   5400
            TabIndex        =   71
            Top             =   1920
            Width           =   615
         End
         Begin VB.ComboBox cmbCRT 
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
            Left            =   3120
            TabIndex        =   69
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtRazao 
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
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   1
            Top             =   840
            Width           =   5895
         End
         Begin VB.TextBox txtFant 
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
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   2
            Top             =   1320
            Width           =   5895
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Index           =   1
            Left            =   2760
            TabIndex        =   103
            Top             =   3000
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Index           =   0
            Left            =   2760
            TabIndex        =   102
            Top             =   2520
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Alq.COFINS:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   10
            Left            =   600
            TabIndex        =   101
            Top             =   3000
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Alq.PIS:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   9
            Left            =   960
            TabIndex        =   99
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Código Regime Tributário:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   240
            TabIndex        =   70
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Razao Social:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   9
            Top             =   480
            Width           =   615
         End
      End
      Begin Threed.SSCheck chkIndustria 
         Height          =   255
         Left            =   -72120
         TabIndex        =   22
         Top             =   3900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Industria ?"
      End
      Begin Threed.SSCheck chkDesconto 
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   4980
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Libera Desconto ?"
      End
      Begin Threed.SSCheck chkNFE 
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   4140
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "NFE ?"
      End
      Begin Threed.SSCheck chkControleEstoque 
         Height          =   255
         Left            =   -72120
         TabIndex        =   25
         Top             =   4140
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Controle Estoque ?"
      End
      Begin Threed.SSCheck chkEstoque 
         Height          =   255
         Left            =   -72120
         TabIndex        =   26
         Top             =   4380
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Trabalha Negativo"
      End
      Begin Threed.SSCheck chkFatPedido 
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   6300
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Recebimento Tela Pedido Venda ?"
      End
      Begin Threed.SSCheck chkLEI_12741 
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   4620
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Imprime Msg Lei 12.741 ?"
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   375
         Left            =   -73125
         TabIndex        =   51
         Top             =   3780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
      Begin MSComctlLib.Toolbar Toolbar_Fone 
         Height          =   390
         Left            =   -68055
         TabIndex        =   29
         Top             =   5100
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "matar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoFone 
         Height          =   330
         Left            =   -69000
         Top             =   6300
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
      Begin MSDataGridLib.DataGrid grdFone 
         Bindings        =   "CADASTROEMPRESA.frx":5D6B
         Height          =   1335
         Left            =   -74565
         TabIndex        =   30
         Top             =   5580
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2355
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
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3344,882
            EndProperty
         EndProperty
      End
      Begin Threed.SSCommand cmdEmail 
         Height          =   495
         Left            =   -70815
         TabIndex        =   31
         Top             =   4260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         _Version        =   262144
         ForeColor       =   -2147483635
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "CADASTROEMPRESA.frx":5D81
         Caption         =   "e-mail"
         Alignment       =   8
         PictureAlignment=   2
      End
      Begin Threed.SSCheck chkBaixaEstPedido 
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   3900
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Baixa Est.Pedido"
      End
      Begin MSComctlLib.Toolbar cmdMataIE 
         Height          =   330
         Left            =   -71280
         TabIndex        =   52
         Top             =   1860
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "matar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin Threed.SSCheck chkMarkap 
         Height          =   270
         Left            =   -74880
         TabIndex        =   58
         Top             =   6540
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   476
         _Version        =   262144
         ForeColor       =   255
         BackStyle       =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Atualiza Preço Venda Markap ?"
      End
      Begin Threed.SSCheck chkCli 
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   5220
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Libera Desconto Cliente ?"
      End
      Begin Threed.SSCheck chkFunc 
         Height          =   255
         Left            =   -74880
         TabIndex        =   60
         Top             =   5460
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Libera Desconto Funcionário ?"
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   -67320
         TabIndex        =   66
         Top             =   1860
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "matar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin Threed.SSCheck chkPercDesconto 
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   5700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Libera % Desconto ?"
      End
      Begin MSMask.MaskEdBox txtCNPJCRED 
         Height          =   360
         Left            =   -70920
         TabIndex        =   80
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
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
      Begin Threed.SSCheck chkLimpaPedido 
         Height          =   255
         Left            =   -74880
         TabIndex        =   83
         Top             =   6060
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Limpa Tela Pedido Venda ?"
      End
      Begin Threed.SSCheck chkDoc_Fiscal 
         Height          =   270
         Left            =   -74880
         TabIndex        =   84
         Top             =   4380
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         _Version        =   262144
         ForeColor       =   255
         BackStyle       =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "DocumentoFiscal?"
      End
      Begin MSComctlLib.Toolbar cmdGravaNFe 
         Height          =   390
         Left            =   -73200
         TabIndex        =   95
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "gravar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin Threed.SSCheck chkBloqFat 
         Height          =   255
         Left            =   -74880
         TabIndex        =   104
         Top             =   3480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Permite Alterar Recebimento Venda?"
      End
      Begin Threed.SSCheck chkTabPreco 
         Height          =   255
         Left            =   -72120
         TabIndex        =   105
         Top             =   4680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Usa Tabela Preço"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ENVIONFE:"
         Height          =   255
         Index           =   8
         Left            =   -70800
         TabIndex        =   97
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ambiente:"
         Height          =   255
         Index           =   7
         Left            =   -72360
         TabIndex        =   94
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   255
         Index           =   6
         Left            =   -73440
         TabIndex        =   92
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   90
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versão NFe:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   88
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NFCe:"
         Height          =   240
         Index           =   2
         Left            =   -72645
         TabIndex        =   86
         Top             =   1740
         Width           =   570
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "CódigoSegurançaContribuinte:"
         Height          =   240
         Left            =   -71460
         TabIndex        =   82
         Top             =   6060
         Width           =   2940
      End
      Begin VB.Label lblCred 
         AutoSize        =   -1  'True
         Caption         =   "Credenciadora Cartão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -71130
         TabIndex        =   78
         Top             =   6420
         Width           =   1875
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Dias de atrazo."
         Height          =   240
         Left            =   -68160
         TabIndex        =   76
         Top             =   5580
         Width           =   1425
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "BloquearVendaApós: "
         Height          =   240
         Left            =   -71040
         TabIndex        =   74
         Top             =   5580
         Width           =   2070
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Insc. Municipal:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -70680
         TabIndex        =   67
         Top             =   1935
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Localização:"
         Height          =   240
         Index           =   4
         Left            =   -71040
         TabIndex        =   63
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NFe:"
         Height          =   240
         Index           =   3
         Left            =   -74100
         TabIndex        =   62
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Instrução Boleto Cobrança"
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
         Height          =   240
         Left            =   -69960
         TabIndex        =   48
         Top             =   4500
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Insc. Estadual:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -74520
         TabIndex        =   43
         Top             =   1935
         Width           =   1365
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -71640
         TabIndex        =   41
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Telefone:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -73965
         TabIndex        =   40
         Top             =   4860
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DDD"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -74565
         TabIndex        =   39
         Top             =   4860
         Width           =   405
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -71745
         TabIndex        =   38
         Top             =   2865
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Local:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   13
         Left            =   -72285
         TabIndex        =   37
         Top             =   4860
         Width           =   585
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -73875
         TabIndex        =   36
         Top             =   3300
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -73485
         TabIndex        =   35
         Top             =   4305
         Width           =   315
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rua/Av./Praça:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -74640
         TabIndex        =   34
         Top             =   2415
         Width           =   1410
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -73680
         TabIndex        =   33
         Top             =   3825
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Número"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   -73950
         TabIndex        =   32
         Top             =   2865
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ/CPF:"
         Height          =   240
         Index           =   1
         Left            =   -74730
         TabIndex        =   21
         Top             =   1140
         Width           =   1020
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Dias Atrazo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72225
         TabIndex        =   17
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Juros Atrazo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74820
         TabIndex        =   15
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem Rodapé (NFe)"
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
         Height          =   240
         Left            =   -69960
         TabIndex        =   14
         Top             =   2580
         Width           =   2220
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem Desconto (NFe)"
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
         Height          =   240
         Left            =   -69960
         TabIndex        =   13
         Top             =   1740
         Width           =   2385
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Dados Adicionais (NFe)"
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
         Height          =   240
         Left            =   -69960
         TabIndex        =   12
         Top             =   3420
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9850
      _ExtentX        =   17383
      _ExtentY        =   1270
      ButtonWidth     =   2858
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar dados"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Excluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir empresa"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar empresas cadastradas"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7680
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":67CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":761C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7080
         Top             =   0
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
               Picture         =   "CADASTROEMPRESA.frx":27E0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":29232
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":2A2C1
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":2B529
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":2C634
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROEMPRESA.frx":2D73F
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
      DesignWidth     =   8385
      DesignHeight    =   7860
   End
End
Attribute VB_Name = "frmCADASTROEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim ID_EMPRESA_N        As Long
   Dim TabEstabelecimento  As New ADODB.Recordset
   Dim TabEmpres           As New ADODB.Recordset
   Dim Id_Empresa_Cadastro As Long
   Dim IE_ID               As Integer

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   LIMPA_TUDO
   CARREGA_CRT
   cmbIdEstabAUX.Text = ESTABELECIMENTO_ID_N
   PROCURA_DADOS_EMPRESA
   MOSTRA_ESTABELECIMENTO ESTABELECIMENTO_ID_N
   CARREGA_ENDEREÇO
   CARREGA_IE
   CARREGA_IM
   MOSTRA_EMPRES

   txtVersaoNFe.Text = "" & MOSTRA_VERSAO_NFe(txtCNPJ.Text)

   cmbIdEstab.Clear

   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

   SQL = "select * from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   TabEstabelecimento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabEstabelecimento.EOF
      cmbIdEstab.AddItem Trim(TabEstabelecimento.Fields("DESCRICAO").Value)
      cmbIdEstabAUX.AddItem Trim(TabEstabelecimento.Fields("estabelecimento_id").Value)
      TabEstabelecimento.MoveNext
   Wend
   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

   SQL = "select * from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabEstabelecimento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEstabelecimento.EOF Then
      cmbIdEstab.Text = "" & Trim(TabEstabelecimento.Fields("DESCRICAO").Value)
      cmbIdEstabAUX.Text = "" & Trim(TabEstabelecimento.Fields("estabelecimento_id").Value)
   End If
   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

   SQL = "select * from CARTAOADM WITH (NOLOCK)"
   TabEstabelecimento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabEstabelecimento.EOF
      txtCNPJCRED.PromptInclude = False
      txtCNPJCRED.Text = "" & TRAZ_TEXTO_TABELA("CARTAOADM", "cnpj", "CARTAOADM_ID", TabEstabelecimento.Fields("CARTAOADM_ID").Value)
      If Len(txtCNPJCRED.Text) <= 11 Then
         txtCNPJCRED.Mask = "###.###.###-##"
         Else: txtCNPJCRED.Mask = "##.###.###/####-##"
      End If
      txtCNPJCRED.PromptInclude = True

      cmbCred.AddItem Trim(TabEstabelecimento.Fields("fantasia").Value) & " - " & txtCNPJCRED.Text
      cmbCredAUX.AddItem Trim(TabEstabelecimento.Fields("cartaoadm_id").Value)
      TabEstabelecimento.MoveNext
   Wend
   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub cmbcrt_Click()
On Error Resume Next

   cmbCRTAUX.ListIndex = cmbCRT.ListIndex

   If Trim(cmbCRTAUX.Text) = "3" Then
      txtPIS.Enabled = True
      txtCOFINS.Enabled = True
      Else
         txtPIS.Enabled = False
         txtCOFINS.Enabled = False
   End If
End Sub

Private Sub stbEmpresa_Click(PreviousTab As Integer)
   If stbEmpresa.Tab = 1 Then
      cmbIdEstabAUX.Text = ESTABELECIMENTO_ID_N

      Call cmbIdEstab_Click
   End If
   If stbEmpresa.Tab = 2 Then
      CARREGA_ENDEREÇO
      CARREGA_IE
      SETA_FONE
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         CRITERIO_A = ""
         frmEmpresaConsulta.Show 1

         txtCNPJ.Text = CRITERIO_A

         CRITERIO_A = ""
         txtCNPJ.SetFocus
      Case "limpar"
         LIMPA_TUDO
         txtCNPJ.SetFocus
      Case "gravar"
         If Trim(txtCNPJ.Text) <> "" Then
            If Trim(txtRazao.Text) <> "" Then
               If Trim(txtRazao.Text) <> "" Then
                  TIPO_PESSOA = "E"
                  GRAVA_TUDO
                  LIMPA_TUDO
                  txtCNPJ.SetFocus
                  Call Form_Load
               End If
            End If
         End If
      Case "sair"
         Unload Me
   End Select
   stbEmpresa.Tab = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub


Private Sub cmbIdEstab_Click()
On Error Resume Next

   cmbIdEstabAUX.ListIndex = cmbIdEstab.ListIndex
   MOSTRA_ESTABELECIMENTO cmbIdEstabAUX.Text
End Sub

Private Sub chkDesconto_Click(Value As Integer)
'On Error GoTo ERRO_TRATA

   chkCli.Enabled = False
   chkFunc.Enabled = False
   chkPercDesconto.Enabled = False
   If chkDesconto.Value <> 0 Then
      chkCli.Enabled = True
      chkFunc.Enabled = True
      chkPercDesconto.Enabled = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
End Sub

Private Sub cmdEmail_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""

   If Trim(txtCNPJ.Text) <> "" Then
      CNPJCPF_A = Trim(txtCNPJ.Text)
      frmEmail.Show 1
   End If

   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
End Sub

Private Sub TXTCNPJ_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear

   MOSTRA_RODAPE "Informe CNPJ da empresa", "F7-Consulta Cadastro Empresa", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJ_GotFocus"
End Sub

Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CRITERIO_A = ""
         frmEmpresaConsulta.Show 1

         txtCNPJ.Text = CRITERIO_A

         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJ_KeyDown"
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCNPJ.Text) = "" Then
         Else
            If TIPO_PESSOA = "E" Then
               If Len(txtCNPJ.Text) > 0 Then
                  Select Case Len(txtCNPJ.Text)
                     Case Is <> 14
                        MsgBox "CNPJ/CNPJ com DV incorreto !!! "
                        txtCNPJ = ""
                        txtCNPJ.SetFocus
                        Exit Sub
                     Case Is = 14
                       If Not VALIDACNPJ(txtCNPJ.Text) Then _
                          MsgBox "CNPJ com DV incorreto !!! "
                  End Select
                  Else
                     MsgBox "CNPJ/CNPJ com DV incorreto !!! "
                     txtCNPJ = ""
                     txtCNPJ.SetFocus
                     Exit Sub
               End If
            End If

            CRITERIO_A = txtCNPJ.Text
      End If

      PROCURA_DADOS_EMPRESA

      'SendKeys "{tab}"
      txtRazao.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJ_KeyPress"
End Sub

Private Sub txtIDFormaPagto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "F6 - Excluir item", "", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIDFormaPagto_GotFocus"
End Sub

Private Sub cmdMataIE_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   If Trim(txtIE.Text) <> "" Then
      SQL = "delete from IE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numr_ie = '" & Trim(txtIE.Text) & "'"
      SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      txtIE.Text = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMataIE_ButtonClick"
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      'SendKeys "{tab}"
      txtJuros.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDias_KeyPress"
End Sub

Private Sub txtDiasAtrazo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      'SendKeys "{tab}"
      'txtJuros.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDiasAtrazo_KeyPress"
End Sub

Private Sub txtIE_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe Inscrição Estadual", "", "", ""

   If txtCNPJ.Text <> "" And txtRazao.Text <> "" Then
      frmINICIO.BARI.Panels.Add (3)
      frmINICIO.BARI.Panels(3).Text = "F10 - Gravar"
      frmINICIO.BARI.Panels(3).AutoSize = sbrContents
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIE_GotFocus"
End Sub

Private Sub txtie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtRua.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtie_KeyPress"
End Sub

Private Sub txtIE_LostFocus()
   If Trim(txtIE.Text) <> "" Then
      txtIE.Text = Replace(txtIE.Text, ".", "")
      txtIE.Text = Replace(txtIE.Text, ",", "")
      txtIE.Text = Replace(txtIE.Text, "-", "")
      txtIE.Text = Replace(txtIE.Text, "/", "")
      txtIE.Text = Replace(txtIE.Text, "\", "")
   End If
End Sub

Private Sub txtIM_LostFocus()
   If Trim(txtIM.Text) <> "" Then
      txtIM.Text = Replace(txtIM.Text, ".", "")
      txtIM.Text = Replace(txtIM.Text, ",", "")
      txtIM.Text = Replace(txtIM.Text, "-", "")
      txtIM.Text = Replace(txtIM.Text, "/", "")
      txtIM.Text = Replace(txtIM.Text, "\", "")
   End If
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      'SendKeys "{tab}"
      txtDias.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtJuros_KeyPress"
End Sub

Private Sub TXTRAZAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTRAZAO_KeyPress"
End Sub

Private Sub txtFant_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFant_KeyPress"
End Sub

Private Sub txtrua_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNumero.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRua_KeyPress"
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtComp.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnumero_KeyPress"
End Sub

Private Sub txtcomp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairro.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtComp_KeyPress"
End Sub

Private Sub txtbairro_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCep.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBairro_KeyPress"
End Sub

Private Sub txtcep_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCep.PromptInclude = False
      If txtCep.Text <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from CEP WITH (NOLOCK)"
         SQL = SQL & " where cep_ID = '" & txtCep.Text & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCep.SetFocus
            Exit Sub
            Else
               txtCidade.Text = TabTemp!CIDADE
               txtUF.Text = TabTemp!UF
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      txtCidade.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_KeyPress"
End Sub

Private Sub txtCEP_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCEP.Show 1
         txtCep.PromptInclude = False
         txtCep.Text = CRITERIO_A
         txtCep.PromptInclude = True
      Case vbKeyF7
         frmCONSULTACEP.Show 1
         txtCep.PromptInclude = False
         txtCep.Text = CRITERIO_A
         txtCep.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_KeyDown"
End Sub

Private Sub txtcidade_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_KeyPress"
End Sub

Private Sub txtuf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtuf_KeyPress"
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDDD_KeyPress"
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtL.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtN_KeyPress"
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtN.Text <> "" And txtCNPJ.Text <> "" Then
         TIPO_PESSOA = "E"
         GRAVA_TUDO
         SETA_FONE
         LIMPA_FONE
      End If

      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtL_KeyPress"
End Sub

Private Sub Toolbar_Fone_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   If Trim(txtCNPJ.Text) <> "" And Trim(txtN.Text) <> "" Then
      SQL = "delete from FONE "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero = '" & Trim(txtN.Text) & "'"
      CONECTA_RETAGUARDA.Execute SQL

      LIMPA_FONE
      SETA_FONE
      txtN.Text = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar_Fone_ButtonClick"
End Sub

Private Sub cmbCRED_Click()
'On Error GoTo ERRO_TRATA

   cmbCredAUX.ListIndex = cmbCred.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCRED_Click"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   adoFone.Enabled = True
   adoFone.ConnectionString = AUTENTICA_GRID

   SQL = "select * from FONE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " order by NUMERO"

   adoFone.RecordSource = SQL
   adoFone.Enabled = True
   adoFone.Refresh

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

Private Sub PROCURA_DADOS_EMPRESA()
'On Error GoTo ERRO_TRATA

   LIMPA_QUASE_TUDO
   PESSOA_ID_N = 0
   ID_EMPRESA_N = 0

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select EMPRESA.pessoa_id,empresa.empresa_id,"
   SQL = SQL & " CRT,seq_nota_saida,SEQ_CUPOM,"
   SQL = SQL & " pessoa.descricao as NomeFant,razao,pessoa.cnpjcpf "

   SQL = SQL & " from EMPRESA WITH (NOLOCK)"

   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"

   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabPessoa.EOF Then _
      MsgBox "Empresa não cadastrado."
   If Not TabPessoa.EOF Then
      txtNFe.Text = "" & TabPessoa.Fields("seq_nota_saida").Value
      txtNFCe = "" & TabPessoa.Fields("SEQ_CUPOM").Value
      PESSOA_ID_N = TabPessoa.Fields("pessoa_id").Value
      ID_EMPRESA_N = TabPessoa.Fields("empresa_id").Value
      txtCNPJ.Text = Trim(TabPessoa.Fields("cnpjcpf").Value)
      cmbCRTAUX.Text = "" & TabPessoa.Fields("crt").Value

      If Trim(cmbCRTAUX.Text) = "3" Then
         txtPIS.Enabled = True
         txtCOFINS.Enabled = True

         txtPIS.Text = "" & TabPessoa.Fields("pis").Value
         txtCOFINS.Text = "" & TabPessoa.Fields("cofins").Value
         Else
            txtPIS.Enabled = False
            txtCOFINS.Enabled = False
      End If


      txtFant.Text = "" & Trim(TabPessoa.Fields("NomeFant").Value)
      txtRazao.Text = "" & Trim(TabPessoa.Fields("razao").Value)

      If TabPessoa.State = 1 Then _
         TabPessoa.Close

      SQL = "select * from DESCR WITH (NOLOCK)"
      SQL = SQL & " where TIPO = 'E' "
      SQL = SQL & " and codigo = " & cmbCRTAUX.Text
      TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPessoa.EOF Then _
         cmbCRT.Text = "" & Trim(TabPessoa!DESCRICAO)
   End If
   If TabPessoa.State = 1 Then _
      TabPessoa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_DADOS_EMPRESA"
End Sub

Private Sub LIMPA_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   txtPIS.Text = ""
   txtCOFINS.Text = ""
   txtNFe.Text = ""
   txtNFCe.Text = ""
   txtInstruçãoBoleto.Text = ""
   PESSOA_ID_N = 0
   txtRazao.Text = ""
   txtIE.Text = ""
   txtIM.Text = ""
   txtCep.PromptInclude = False
   txtCep.Text = ""
   txtRua.Text = ""
   txtComp.Text = ""
   txtBairro.Text = ""
   txtCidade.Text = ""
   txtUF.Text = ""
   txtNumero.Text = ""
   txtFant.Text = ""
   txtDescDesconto.Text = ""
   txtMSG.Text = ""
   txtDadosAdicionais.Text = ""
   chkControleEstoque.Value = 0
   chkIndustria.Value = 0
   chkDesconto.Value = 0
   chkCli.Value = 0
   chkFunc.Value = 0
   chkPercDesconto.Value = 0
   chkMarkap.Value = 0
   chkBaixaEstPedido.Value = 0
   chkBaixaEstPedido.Refresh

   chkNFE.Value = 0
   chkDoc_Fiscal.Value = 0

   chkLEI_12741.Value = 0
   chkFatPedido.Value = 0
   chkBloqFat.Value = 0
   chkTabPreco.Value = 0
   chkLimpaPedido.Value = 0
   chkEstoque.Value = 0
   cmbIdEstab.Text = ""
   cmbIdEstabAUX.Text = ""
   txtNomeEstab.Text = ""
   txtCNPJCPF.Text = ""
   txtCSC.Text = ""
   cmbCred.Text = ""
   cmbCredAUX.Text = ""
   CNPJ_CRED_CARTAO_ESTAB = ""
   txtLoc.Text = ""
   txtJuros.Text = 0
   txtDias.Text = 0
   txtDiasAtrazo.Text = 0

   CRITERIO_A = 0
   LIMPA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_QUASE_TUDO"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtCNPJ.Text = ""
   cmbCRT.Text = ""
   cmbCRTAUX.Text = ""
   LIMPA_QUASE_TUDO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   Dim strTipoEmpresa      As String

   PESSOA_ID_N = 0
   ENDERECO_ID_N = 0
   CEP_ID_A = ""

   If txtCNPJ.Text = "" Then
      MsgBox "Informe CNPJ."
      txtCNPJ.SetFocus
      Exit Sub
   End If
   If txtRazao.Text = "" Then
      MsgBox "Informe Nome."
      txtRazao.SetFocus
      Exit Sub
   End If
   If TIPO_PESSOA = "F" Then
      strTipoEmpresa = "FOR"
      Else: strTipoEmpresa = "EMP"
   End If


   If Trim(txtIE.Text) = "" Then
      txtIE.Text = "ISENTO"
      Else
         If Trim(txtIE.Text) <> "ISENTO" Then
            If Trim(txtUF.Text) = "" Then _
               MsgBox "Estado (UF) não informado para inscrição : " & txtIE.Text
            If Valida_Inscricao_Estadual(txtIE.Text, txtUF.Text) <> 0 Then
               txtIE.SetFocus
               Exit Sub
            End If
         End If
   End If

GRAVA_PESSOA
GRAVA_EMPRESA

If Trim(cmbIdEstabAUX.Text) <> "" Then _
   If IsNumeric(cmbIdEstabAUX.Text) Then _
      GRAVA_ESTABELECIMENTO cmbIdEstabAUX.Text, EMPRESA_ID_N

GRAVA_FONE
BUSCA_CEP
GRAVA_ENDEREÇO

   If Trim(txtIE.Text) <> "" Then _
      If Trim(txtIE.Text) <> "ISENTO" Then _
         GRAVA_IE Trim(txtIE.Text)

   If Trim(txtIM.Text) <> "" Then _
      GRAVA_IM Trim(txtIM.Text)

GRAVA_OBS
'GLOBAL_EMPRES

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Sub CARREGA_OBS()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select obs,seq from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      If Not IsNull(TabConsulta.Fields("obs").Value) Then
         If Not IsNull(TabConsulta.Fields("seq").Value) Then

            If TabConsulta.Fields("seq").Value = 1 Then _
               txtDadosAdicionais.Text = "" & Trim(TabConsulta.Fields("obs").Value)

            If TabConsulta.Fields("seq").Value = 2 Then _
               txtMSG.Text = "" & Trim(TabConsulta.Fields("obs").Value)

            If TabConsulta.Fields("seq").Value = 3 Then _
               txtDescDesconto.Text = "" & Trim(TabConsulta.Fields("obs").Value)
         End If
      End If

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_OBS"
End Sub

Sub CARREGA_ENDEREÇO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from ENDERECO WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      ENDERECO_ID_N = 0 & TabConsulta.Fields("endereco_id").Value

      If Not IsNull(TabConsulta!Rua) Then _
         txtRua.Text = TabConsulta!Rua
      If Not IsNull(TabConsulta!Bairro) Then _
         txtBairro.Text = TabConsulta!Bairro
      If Not IsNull(TabConsulta!Complemento) Then _
         txtComp.Text = TabConsulta!Complemento

      txtNumero.Text = "" & TabConsulta.Fields("numero").Value

      If TabUSU.State = 1 Then _
         TabUSU.Close

      SQL = "select * from CEP WITH (NOLOCK)"
      SQL = SQL & " where cep_ID = '" & Trim(TabConsulta.Fields("cep_id").Value) & "'"
      TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabUSU.EOF Then
         If Not IsNull(TabUSU.Fields("Cep_id").Value) Then _
            txtCep.Text = TabUSU.Fields("Cep_id").Value

         If Not IsNull(TabUSU!CIDADE) Then _
            txtCidade.Text = TabUSU!CIDADE

         If Not IsNull(TabUSU!UF) Then _
            txtUF.Text = TabUSU!UF
      End If
      If TabUSU.State = 1 Then _
         TabUSU.Close
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_ENDEREÇO"
End Sub

Sub MOSTRA_ESTABELECIMENTO(Estab_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

   SQL = "select * from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & Estab_ID_N
   TabEstabelecimento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEstabelecimento.EOF Then
      txtNomeEstab.Text = "" & Trim(TabEstabelecimento.Fields("DESCRICAO").Value)
      txtLoc.Text = "" & Trim(TabEstabelecimento.Fields("DESCRICAO").Value)
      txtLoc.Text = "" & Trim(TabEstabelecimento.Fields("localizacao").Value)
      txtCNPJCPF.Text = "" & Trim(TabEstabelecimento.Fields("CNPJCPF").Value)

      If Not IsNull(TabEstabelecimento.Fields("CARTAOADM_ID").Value) Then
         txtCNPJCRED.PromptInclude = False
         txtCNPJCRED.Text = "" & TRAZ_TEXTO_TABELA("CARTAOADM", "cnpj", "CARTAOADM_ID", TabEstabelecimento.Fields("CARTAOADM_ID").Value)
         If Len(txtCNPJCRED.Text) <= 11 Then
            txtCNPJCRED.Mask = "###.###.###-##"
            Else: txtCNPJCRED.Mask = "##.###.###/####-##"
         End If
         txtCNPJCRED.PromptInclude = True

         cmbCred.Text = "" & TRAZ_TEXTO_TABELA("CARTAOADM", "FANTASIA", "CARTAOADM_ID", TabEstabelecimento.Fields("CARTAOADM_ID").Value) _
                           & " - " & txtCNPJCRED.Text
         cmbCredAUX.Text = "" & TabEstabelecimento.Fields("CARTAOADM_ID").Value
      End If

      txtDias.Text = "" & Trim(TabEstabelecimento.Fields("QTD_DIAS_ATRAZO").Value)
      txtDiasAtrazo.Text = "" & Trim(TabEstabelecimento.Fields("DiasAtrazoCliente").Value)
      txtJuros.Text = "" & Trim(TabEstabelecimento.Fields("PERC_JUROS_ATRAZO").Value)
      txtInstruçãoBoleto.Text = "" & Trim(TabEstabelecimento.Fields("instrucao_boleto").Value)

      If Not IsNull(TabEstabelecimento.Fields("USA_NFE").Value) Then
         If TabEstabelecimento.Fields("USA_NFE").Value = False Then
            chkNFE.Value = 0
            Else: chkNFE.Value = 1
         End If
      End If
      If Not IsNull(TabEstabelecimento.Fields("DOC_FISCAL").Value) Then
         If TabEstabelecimento.Fields("DOC_FISCAL").Value = False Then
            chkDoc_Fiscal.Value = 0
            Else: chkDoc_Fiscal.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("LIBERA_DESCONTO").Value) Then
         If TabEstabelecimento.Fields("LIBERA_DESCONTO").Value = False Then
            chkDesconto.Value = 0
            Else: chkDesconto.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("DESCONTO_CLIENTE").Value) Then
         If TabEstabelecimento.Fields("DESCONTO_CLIENTE").Value = False Then
            chkCli.Value = 0
            Else: chkCli.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("DESCONTO_FUNCIONARIO").Value) Then
         If TabEstabelecimento.Fields("DESCONTO_FUNCIONARIO").Value = False Then
            chkFunc.Value = 0
            Else: chkFunc.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("LiberaPercDesconto").Value) Then
         If TabEstabelecimento.Fields("LiberaPercDesconto").Value = False Then
            chkPercDesconto.Value = 0
            Else: chkPercDesconto.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("ATUALIZA_ESTOQUE_req").Value) Then
         If TabEstabelecimento.Fields("ATUALIZA_ESTOQUE_req").Value = False Then
            chkBaixaEstPedido.Value = 0
            Else: chkBaixaEstPedido.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("AT_VENDA_MKP").Value) Then
         If TabEstabelecimento.Fields("AT_VENDA_MKP").Value = False Then
            chkMarkap.Value = 0
            Else: chkMarkap.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("RECEBE_PEDIDO_VENDA").Value) Then
         If TabEstabelecimento.Fields("RECEBE_PEDIDO_VENDA").Value = False Then
            chkFatPedido.Value = 0
            Else: chkFatPedido.Value = 1
         End If
      End If

      'libera ou bloqueia mudanda de tipo de venda na tela de recebimento
      If Not IsNull(TabEstabelecimento.Fields("ALTERA_FATURA").Value) Then
         If TabEstabelecimento.Fields("ALTERA_FATURA").Value = False Then
            chkBloqFat.Value = 0
            Else: chkBloqFat.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("LEI_12741").Value) Then
         If TabEstabelecimento.Fields("LEI_12741").Value = False Then
            chkLEI_12741.Value = 0
            Else: chkLEI_12741.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("indr_industria").Value) Then
         If TabEstabelecimento.Fields("indr_industria").Value = False Then
            chkIndustria.Value = 0
            Else: chkIndustria.Value = 1
         End If
      End If

      INDR_CONTROLA_ESTOQUE = False
      If Not IsNull(TabEstabelecimento.Fields("CONTROLE_ESTOQUE").Value) Then
         If TabEstabelecimento.Fields("CONTROLE_ESTOQUE").Value = False Then
            chkControleEstoque.Value = 0
            Else
               chkControleEstoque.Value = 1
               INDR_CONTROLA_ESTOQUE = TabEstabelecimento.Fields("CONTROLE_ESTOQUE").Value
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("ESTOQUE_NEGATIVO").Value) Then
         If TabEstabelecimento.Fields("ESTOQUE_NEGATIVO").Value = False Then
            chkEstoque.Value = 0
            Else: chkEstoque.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("csc").Value) Then _
         txtCSC.Text = "" & TabEstabelecimento.Fields("csc").Value

      If Not IsNull(TabEstabelecimento.Fields("limpa_pedido").Value) Then
         If TabEstabelecimento.Fields("limpa_pedido").Value = False Then
            chkLimpaPedido.Value = 0
            Else: chkLimpaPedido.Value = 1
         End If
      End If

      If Not IsNull(TabEstabelecimento.Fields("USA_TAB_PRECO").Value) Then
         If TabEstabelecimento.Fields("USA_TAB_PRECO").Value = False Then
            chkTabPreco.Value = 0
            Else: chkTabPreco.Value = 1
         End If
      End If

      CARREGA_OBS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ESTABELECIMENTO"
End Sub

Sub CARREGA_IE()
'On Error GoTo ERRO_TRATA

   txtIE.Text = ""

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from IE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      txtIE.Text = "" & Trim(TabConsulta.Fields("numr_ie").Value)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_IE"
End Sub

Sub CARREGA_IM()
'On Error GoTo ERRO_TRATA

   txtIM.Text = ""

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from IM WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      txtIM.Text = "" & Trim(TabConsulta.Fields("numr_IM").Value)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_IM"
End Sub

Sub GRAVA_PESSOA()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJ.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close

   CONT_N = 2
   If PESSOA_ID_N <= 0 Then _
      CONT_N = 1

   spPessoa CONT_N, PESSOA_ID_N, Trim(txtCNPJ.Text), Trim(txtFant.Text), Trim(txtRazao.Text), "A"

   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJ.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PESSOA"
End Sub

Sub GRAVA_EMPRESA()
'On Error GoTo ERRO_TRATA

   Dim CRT_N As Integer

   CRT_N = 1
   If cmbCRTAUX.Text <> "" Then _
      If IsNumeric(cmbCRTAUX.Text) Then _
         CRT_N = cmbCRTAUX.Text

   If Trim(txtNFe.Text) = "" Then _
      txtNFe.Text = 0
   If Trim(txtNFCe.Text) = "" Then _
      txtNFCe.Text = 0
   If Not IsNumeric(txtNFe.Text) Then _
      txtNFe.Text = 0
   If Not IsNumeric(txtNFCe.Text) Then _
      txtNFCe.Text = 0

   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

   SQL = "select * from EMPRESA WITH (NOLOCK)"

   'SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   'SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"

   'SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   'SQL = SQL & " and cgc = '" & Trim(txtCNPJ.Text) & "'"

SQL = SQL & " where cgc = '" & Trim(txtCNPJ.Text) & "'"

   tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If tabEmpresa.EOF Then
      Id_Empresa_Cadastro = MAX_ID("empresa_id", "empresa", "", "", "", "")

      SQL = "INSERT INTO EMPRESA "
      SQL = SQL & " ("
         SQL = SQL & " EMPRESA_ID,PESSOA_ID,CGC,RAZAO_SOCIAL,nome_fant,CRT"
      SQL = SQL & " ) "
      SQL = SQL & " VALUES ("
         SQL = SQL & Id_Empresa_Cadastro
         SQL = SQL & "," & PESSOA_ID_N
         SQL = SQL & ",'" & Trim(txtCNPJ.Text) & "'"
         SQL = SQL & ",'" & Trim(txtRazao.Text) & "'"
         SQL = SQL & ",'" & Trim(txtFant.Text) & "'"
         SQL = SQL & "," & CRT_N
      SQL = SQL & ")"
      Else
         Id_Empresa_Cadastro = tabEmpresa!EMPRESA_ID

         SQL = "update EMPRESA SET "

            SQL = SQL & " RAZAO_social = '" & Trim(txtRazao.Text) & "'"
            SQL = SQL & ",nome_fant = '" & Trim(txtFant.Text) & "'"
            SQL = SQL & ",CRT = " & CRT_N

         SQL = SQL & " where cgc = '" & Trim(txtCNPJ.Text) & "'"
   End If
   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_EMPRESA"
End Sub

Sub GRAVA_ESTABELECIMENTO(Estab_ID_N As Long, ID_EMPRESA_N As Long)
'On Error GoTo ERRO_TRATA

   If Estab_ID_N <= 0 Then
      MsgBox "Estabelecimento não informado."
      Exit Sub
   End If

   If ID_EMPRESA_N <= 0 Then
      MsgBox "Empresa com identificação incorreta, verificar. " & ID_EMPRESA_N
      Exit Sub
   End If

   If Trim(txtJuros.Text) = "" Then _
      txtJuros.Text = 0
   If Trim(txtDias.Text) = "" Then _
      txtDias.Text = 0
   If Trim(txtDiasAtrazo.Text) = "" Then _
      txtDiasAtrazo.Text = 0

   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

   SQL = "select * from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & ID_EMPRESA_N
   SQL = SQL & " and estabelecimento_id = " & Estab_ID_N
   TabEstabelecimento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabEstabelecimento.EOF Then
      SQL = "insert into ESTABELECIMENTO "
      SQL = SQL & " VALUES("
         SQL = SQL & Estab_ID_N
         SQL = SQL & "," & EMPRESA_ID_N                     'EMPRESA_ID
         SQL = SQL & ",'" & Trim(txtNomeEstab.Text) & "'"   'DESCRICAO
         SQL = SQL & ",'" & Trim(txtLoc.Text) & "'"         'LOCALIZACAO

         'CONTROLE_ESTOQUE
         If chkControleEstoque.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'INDR_INDUSTRIA_B
         If chkIndustria.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'LIBERA_DESCONTO
         If chkDesconto.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'USA_NFe
         If chkNFE.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'RECEBE_PEDIDO_VENDA
         If chkFatPedido.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'ESTOQUE_NEGATIVO
         If chkEstoque.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'LEI_12741
         If chkLEI_12741.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'INSTRUCAO_BOLETO
         SQL = SQL & ",'" & Trim(txtInstruçãoBoleto.Text) & "'"                         'INSTRUCAO_BOLETO

         'ATUALIZA_ESTOQUE_REQ
         If chkBaixaEstPedido.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         SQL = SQL & "," & tpMOEDA(txtJuros.Text)                                     'PERC_JUROS_ATRAZO
         SQL = SQL & "," & Trim(txtDias.Text)                                         'QTD_DIAS_ATRAZO

         'AT_VENDA_MKP
         If chkMarkap.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         If chkCli.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         If chkFunc.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         If chkPercDesconto.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         SQL = SQL & "," & Trim(txtDiasAtrazo.Text)

         If Trim(cmbCredAUX.Text) = "" Then
            SQL = SQL & ", Null"
            Else: SQL = SQL & "," & Trim(cmbCredAUX.Text)
         End If

         If Trim(cmbCredAUX.Text) = "" Then
            SQL = SQL & ", Null"
            Else: SQL = SQL & "," & Trim(txtCSC.Text)
         End If

         'RECEBE_PEDIDO_VENDA
         If chkLimpaPedido.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         If chkDoc_Fiscal.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         'libera ou bloqueia mudanda de tipo de venda na tela de recebimento
         If chkBloqFat.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         If chkTabPreco.Value = 0 Then
            SQL = SQL & ",'false'"
            Else: SQL = SQL & ",'true'"
         End If

         SQL = SQL & ")"
      Else  'UPDATE
         SQL = "update ESTABELECIMENTO set "

SQL = SQL & " seq_nota_saida = " & Trim(txtNFe.Text)
SQL = SQL & ",seq_CUPOM = " & Trim(txtNFCe.Text)

            SQL = SQL & ", DESCRICAO = '" & Trim(txtNomeEstab.Text) & "'"     'DESCRICAO
            SQL = SQL & ", LOCALIZACAO = '" & Trim(txtLoc.Text) & "'"         'LOCALIZACAO

            'CONTROLE_ESTOQUE
            If chkControleEstoque.Value = 0 Then
               SQL = SQL & ", CONTROLE_ESTOQUE = 'false'"
               Else: SQL = SQL & ", CONTROLE_ESTOQUE = 'true'"
            End If

            'INDR_INDUSTRIA_B
            If chkIndustria.Value = 0 Then
               SQL = SQL & ", INDR_INDUSTRIA = 'false'"
               Else: SQL = SQL & ", INDR_INDUSTRIA = 'true'"
            End If

            'LIBERA_DESCONTO
            If chkDesconto.Value = 0 Then
               SQL = SQL & ", LIBERA_DESCONTO = 'false'"
               Else: SQL = SQL & ", LIBERA_DESCONTO = 'true'"
            End If

            If chkCli.Value = 0 Then
               SQL = SQL & ", DESCONTO_CLIENTE = 'false'"
               Else: SQL = SQL & ", DESCONTO_CLIENTE = 'true'"
            End If

            If chkFunc.Value = 0 Then
               SQL = SQL & ", DESCONTO_FUNCIONARIO = 'false'"
               Else: SQL = SQL & ", DESCONTO_FUNCIONARIO = 'true'"
            End If

            If chkPercDesconto.Value = 0 Then
               SQL = SQL & ", LiberaPercDesconto = 'false'"
               Else: SQL = SQL & ", LiberaPercDesconto = 'true'"
            End If

            'AT_VENDA_MKP
            If chkMarkap.Value = 0 Then
               SQL = SQL & ", AT_VENDA_MKP = 'false'"
               Else: SQL = SQL & ", AT_VENDA_MKP = 'true'"
            End If

            'USA_NFe
            If chkNFE.Value = 0 Then
               SQL = SQL & ", USA_NFe = 'false'"
               Else: SQL = SQL & ", USA_NFe = 'true'"
            End If

            'RECEBE_PEDIDO_VENDA
            If chkFatPedido.Value = 0 Then
               SQL = SQL & ", RECEBE_PEDIDO_VENDA = 'false'"
               Else: SQL = SQL & ", RECEBE_PEDIDO_VENDA = 'true'"
            End If

            'libera ou bloqueia mudanda de tipo de venda na tela de recebimento
            If chkBloqFat.Value = 0 Then
               SQL = SQL & ", ALTERA_FATURA = 'false'"
               Else: SQL = SQL & ", ALTERA_FATURA = 'true'"
            End If

            'ESTOQUE_NEGATIVO
            If chkEstoque.Value = 0 Then
               SQL = SQL & ", ESTOQUE_NEGATIVO = 'false'"
               Else: SQL = SQL & ",  ESTOQUE_NEGATIVO = 'true'"
            End If

            'LEI_12741
            If chkLEI_12741.Value = 0 Then
               SQL = SQL & ", LEI_12741 = 'false'"
               Else: SQL = SQL & ", LEI_12741 = 'true'"
            End If

            If chkTabPreco.Value = 0 Then
               SQL = SQL & ", USA_TAB_PRECO = 'false'"
               Else: SQL = SQL & ", USA_TAB_PRECO = 'true'"
            End If

            SQL = SQL & ", INSTRUCAO_BOLETO = '" & Trim(txtInstruçãoBoleto.Text) & "'" 'INSTRUCAO_BOLETO
   
            'ATUALIZA_ESTOQUE_REQ
            If chkBaixaEstPedido.Value = 0 Then
               SQL = SQL & ", ATUALIZA_ESTOQUE_REQ = 'false'"
               Else: SQL = SQL & ", ATUALIZA_ESTOQUE_REQ = 'true'"
            End If

            SQL = SQL & ", PERC_JUROS_ATRAZO = " & tpMOEDA(txtJuros.Text)             'PERC_JUROS_ATRAZO
            SQL = SQL & ", QTD_DIAS_ATRAZO = " & txtDias.Text                         'QTD_DIAS_ATRAZO
            SQL = SQL & ", DiasAtrazoCliente = " & txtDiasAtrazo.Text                 'DiasAtrazoCliente

            If Trim(cmbCredAUX.Text) = "" Then
               SQL = SQL & ", CARTAOADM_ID = Null"
               Else: SQL = SQL & ", CARTAOADM_ID = " & Trim(cmbCredAUX.Text)
            End If

            SQL = SQL & ", csc = '" & txtCSC.Text & "'"                                'CódigoSegurançaContribuinte

            If chkLimpaPedido.Value = 0 Then
               SQL = SQL & ", limpa_pedido = 'false'"
               Else: SQL = SQL & ", limpa_pedido = 'true'"
            End If

            If chkDoc_Fiscal.Value = 0 Then
               SQL = SQL & ", DOC_FISCAL = 'false'"
               Else: SQL = SQL & ", DOC_FISCAL = 'true'"
            End If

         SQL = SQL & " where empresa_id = " & ID_EMPRESA_N
         SQL = SQL & " and estabelecimento_id = " & Estab_ID_N
   End If
   If TabEstabelecimento.State = 1 Then _
      TabEstabelecimento.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ESTABELECIMENTO"
End Sub

Sub GRAVA_FONE()
'On Error GoTo ERRO_TRATA

   If Trim(txtN.Text) <> "" Then
      Dim FONE_ID As Integer

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FONE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero = '" & Trim(txtN.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         spFONE 1, 0, Trim(txtN.Text), PESSOA_ID_N, Trim(txtDDD.Text), Trim(txtL.Text)
         Else: spFONE 2, TabTemp.Fields("fone_id").Value, Trim(txtN.Text), PESSOA_ID_N, Trim(txtDDD.Text), Trim(txtL.Text)
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_FONE"
End Sub

Sub BUSCA_CEP()
'On Error GoTo ERRO_TRATA

   CEP_ID_A = "74000000"
   If Trim(txtCep.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CEP WITH (NOLOCK)"
      SQL = SQL & " where cep_ID = '" & Trim(txtCep.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         CEP_ID_A = TabTemp!CEP_ID

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_CEP"
End Sub

Sub GRAVA_ENDEREÇO()
'On Error GoTo ERRO_TRATA

   ENDERECO_ID_N = 0
   If Trim(txtRua.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from ENDERECO WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         ENDERECO_ID_N = TabTemp.Fields("ENDERECO_ID").Value

         SQL = "UPDATE ENDERECO SET "
         SQL = SQL & " Rua = '" & txtRua.Text & "'"
         SQL = SQL & ", Bairro = '" & txtBairro.Text & "'"
         SQL = SQL & ", Complemento = '" & txtComp.Text & "'"
         SQL = SQL & ", Tipo = 'C'"
         SQL = SQL & ", numero = '" & Trim(txtNumero.Text) & "'"
         SQL = SQL & ", cep_id = '" & Trim(txtCep.Text) & "'"
         SQL = SQL & " where endereco_id = " & ENDERECO_ID_N
         Else
            ENDERECO_ID_N = MAX_ID("endereco_id", "endereco", "", "", "", "")

            SQL = "INSERT INTO ENDERECO "
            SQL = SQL & "(endereco_id, CEP_id, Rua, Bairro, Complemento, Tipo, numero,pessoa_id) "
            SQL = SQL & " VALUES ("
            SQL = SQL & ENDERECO_ID_N
            SQL = SQL & ",'" & Trim(txtCep.Text) & "'"
            SQL = SQL & ",'" & txtRua.Text & "'"
            SQL = SQL & ",'" & txtBairro.Text & "'"
            SQL = SQL & ",'" & txtComp.Text & "'"
            SQL = SQL & ",'" & "C" & "'"
            SQL = SQL & ",'" & Trim(txtNumero.Text) & "'"
            SQL = SQL & "," & PESSOA_ID_N
            SQL = SQL & ")"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ENDEREÇO"
End Sub

Sub GRAVA_OBS()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and seq = 1"
   SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into OBS (PESSOA_ID,SEQ,OBS,TIPO_REGISTRO)  values("
         SQL = SQL & PESSOA_ID_N
         SQL = SQL & ",1"                             '[SEQ]
         SQL = SQL & ",'" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
         SQL = SQL & ",'" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
      SQL = SQL & ")"
      Else
         SQL = "update OBS set"
         SQL = SQL & " obs = '" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and seq = 1"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and seq = 2"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into OBS (PESSOA_ID,SEQ,OBS,TIPO_REGISTRO)  values("
         SQL = SQL & PESSOA_ID_N
         SQL = SQL & ",2"                             '[SEQ]
         SQL = SQL & ",'" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
         SQL = SQL & ",'" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
      SQL = SQL & ")"
      Else
         SQL = "update OBS set"
         SQL = SQL & " obs = '" & Trim(txtMSG.Text) & "'"  '[OBS]
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and seq = 2"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and seq = 3"
   SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into OBS (PESSOA_ID,SEQ,OBS,TIPO_REGISTRO)  values("
         SQL = SQL & PESSOA_ID_N
         SQL = SQL & ",3"                             '[SEQ]
         SQL = SQL & ",'" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
         SQL = SQL & ",'" & Trim(txtDescDesconto.Text) & "'"  '[OBS]
      SQL = SQL & ")"
      Else
         SQL = "update OBS set"
         SQL = SQL & " obs = '" & Trim(txtDadosAdicionais.Text) & "'"  '[OBS]
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and seq = 3"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_OBS"
End Sub

Sub CARREGA_CRT()
'On Error GoTo ERRO_TRATA

   cmbCRT.Clear
   cmbCRTAUX.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'E' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      cmbCRT.AddItem Trim(TabConsulta!DESCRICAO)
      cmbCRTAUX.AddItem Trim(TabConsulta!Codigo)
      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_CRT"
End Sub

Private Sub cmdGravaNFe_ButtonClick(ByVal Button As MSComctlLib.Button)
   GRAVA_VERSAO_NFe
End Sub

Sub GRAVA_VERSAO_NFe()
'On Error GoTo ERRO_TRATA

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then _
      Exit Sub

   If Trim(txtVersaoNFe.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtVersaoNFe.Text) Then _
      Exit Sub

   SQL = "update EMPRES set versaonfe = '" & Trim(txtVersaoNFe.Text) & "'"
   SQL = SQL & " where cnpj = '" & Trim(txtCNPJ.Text) & "'"
   CONECTA_GLOBAL.Execute SQL
   
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   MsgBox "Alterado Versão Nota Fiscal Eletrônica."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_VERSAO_NFe"
End Sub

Sub GLOBAL_EMPRES()
'On Error GoTo ERRO_TRATA

   If Trim(txtCNPJ.Text) <> "" And Trim(txtRazao.Text) <> "" And Trim(txtFant.Text) <> "" And PESSOA_ID_N > 0 Then
   
      ABRE_BANCO_GLOBAL
   
      If CONECTA_GLOBAL.State <> 1 Then
         MsgBox "Banco GLOBAL não conectado."
         Else
            Dim R_E_C_N_O_ As Long
            Dim EMPRESA_A  As String
            Dim NOME_REDUZ As String
            Dim FILIAL_REDUZ As String
            Dim NOME_EMPRESA As String
            Dim TELEFONE As String
            Dim FAX As String
            Dim CNPJ As String
            Dim INSC_ESTADUAL As String
            Dim ENT_RUA_AV As String
            Dim ENT_COMPLEMENTO As String
            Dim ENT_BAIRRO As String
            Dim ENT_CIDADE As String
            Dim ENT_ESTADO As String
            Dim ENT_CEP As String
            Dim COB_RUA_AV As String
            Dim COB_COMPLEMENTO As String
            Dim COB_BAIRRO As String
            Dim COB_CIDADE As String
            Dim COB_ESTADO As String
            Dim COB_CEP As String
            Dim TIPO_INSCRICAO As String
            Dim INSC_SUFRAMA As String
            Dim NFESERIEPA As String
            Dim NFETPAMBI As String
            Dim NFECODMODP As String
            Dim NFEMIELETRO As String
            Dim ENT_CODCID As String
            Dim ENT_RUA As String
            Dim UFNFe As String
            Dim GRAVATEXTO As String
            Dim ENVIONFE As String
            Dim CRIADIRETORIO As String
            Dim CERTITEXTO As String
            Dim CRT As String
            Dim VersaoNfe As String

            NOME_REDUZ = "" & Trim(Left(txtFant.Text, 60))
            FILIAL_REDUZ = "" & Trim(Left(txtFant.Text, 60))
            NOME_EMPRESA = "" & Trim(Left(txtRazao.Text, 60))
            CNPJ = "" & Trim(txtCNPJ.Text)

            INSC_ESTADUAL = "" & Trim(TRAZ_IE(PESSOA_ID_N))

            TELEFONE = ""
            FAX = ""

            If TabEmpres.State = 1 Then _
               TabEmpres.Close

            SQL = "select * from FONE WITH (NOLOCK)"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            TabEmpres.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabEmpres.EOF Then
               TELEFONE = "" & TabEmpres.Fields("numero").Value
               FAX = "" & TELEFONE
            End If
            If TabEmpres.State = 1 Then _
               TabEmpres.Close

            ENT_RUA_AV = ""
            ENT_COMPLEMENTO = ""
            ENT_BAIRRO = ""
            ENT_CIDADE = ""
            ENT_ESTADO = ""
            ENT_CEP = ""
            COB_RUA_AV = ""
            COB_COMPLEMENTO = ""
            COB_BAIRRO = ""
            COB_CIDADE = ""
            COB_ESTADO = ""
            COB_CEP = ""
            ENT_CODCID = ""
            ENT_RUA = ""
            R_E_C_N_O_ = 0

            If TabEmpres.State = 1 Then _
               TabEmpres.Close

            SQL = "SELECT ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, ENDERECO.NUMERO, CEP.CEP_ID, CEP.CIDADE, CEP.UF, CEP.IBGE_ID"
            SQL = SQL & " FROM ENDERECO WITH (NOLOCK)"
            SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
            SQL = SQL & " ON ENDERECO.CEP_ID = CEP.CEP_ID "
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            TabEmpres.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabEmpres.EOF Then
               ENT_RUA_AV = "" & TabEmpres.Fields("rua").Value
               ENT_COMPLEMENTO = "" & Trim(Left(TabEmpres.Fields("complemento").Value, 70))
               ENT_BAIRRO = "" & TabEmpres.Fields("bairro").Value
               ENT_CIDADE = "" & TabEmpres.Fields("cidade").Value
               ENT_ESTADO = "" & TabEmpres.Fields("uf").Value
               ENT_CEP = "" & TabEmpres.Fields("cep_id").Value
               COB_RUA_AV = "" & TabEmpres.Fields("rua").Value
               COB_COMPLEMENTO = "" & TabEmpres.Fields("complemento").Value
               COB_BAIRRO = "" & TabEmpres.Fields("bairro").Value
               COB_CIDADE = "" & TabEmpres.Fields("cidade").Value
               COB_ESTADO = "" & TabEmpres.Fields("uf").Value
               COB_CEP = "" & TabEmpres.Fields("cep_id").Value
               ENT_CODCID = "" & TabEmpres.Fields("ibge_id").Value
               ENT_RUA = "" & TabEmpres.Fields("rua").Value
            End If
            If TabEmpres.State = 1 Then _
               TabEmpres.Close

            UFNFe = ""
   
            TIPO_INSCRICAO = "2"
            INSC_SUFRAMA = ""
            NFESERIEPA = "1"
            NFECODMODP = "55"
            NFEMIELETRO = "Normal"
            GRAVATEXTO = ""
            CRIADIRETORIO = "0"
            CERTITEXTO = ""
            CRT = "" & cmbCRTAUX.Text
            VersaoNfe = "" & Trim(txtVersaoNFe.Text)

            EMPRESA_A = "" & txtEmpresa.Text
            txtFilial.Text = "0" & ESTABELECIMENTO_ID_N
            NFETPAMBI = "" & txtAmbiente.Text   'AMBIENTE: P=PRODUÇÃO / H=HOMOLOGAÇÃO
            ENVIONFE = "" & txtENVIONFE.Text

            If TabEmpres.State = 1 Then _
               TabEmpres.Close
   
            SQL = "select R_E_C_N_O_ from EMPRES WITH (NOLOCK) "
               SQL = SQL & " where cnpj = '" & Trim(txtCNPJ.Text) & "'"
            TabEmpres.Open SQL, CONECTA_GLOBAL, , , adCmdText
            If TabEmpres.EOF Then
               If TabEmpres.State = 1 Then _
                  TabEmpres.Close
               SQL = "select max(R_E_C_N_O_) from EMPRES WITH (NOLOCK)"
               TabEmpres.Open SQL, CONECTA_GLOBAL, , , adCmdText
               If Not TabEmpres.EOF Then _
                  If Not IsNull(TabEmpres.Fields(0).Value) Then _
                     R_E_C_N_O_ = TabEmpres.Fields(0).Value
               If TabEmpres.State = 1 Then _
                  TabEmpres.Close

               'INSERT
               SQL = "insert into EMPRES "
               SQL = SQL & " ("
                  SQL = SQL & " Empresa,FILIAL,NOME_REDUZ,FILIAL_REDUZ,NOME_EMPRESA,TELEFONE,"
                  SQL = SQL & " FAX,CNPJ,INSC_ESTADUAL,ENT_RUA_AV,ENT_COMPLEMENTO,ENT_BAIRRO,"
                  SQL = SQL & " ENT_CIDADE,ENT_ESTADO,ENT_CEP,COB_RUA_AV,COB_COMPLEMENTO,COB_BAIRRO,"
                  SQL = SQL & " COB_CIDADE,COB_ESTADO,COB_CEP,TIPO_INSCRICAO,INSC_SUFRAMA,"
                  SQL = SQL & " NFESERIEPA,NFETPAMBI,NFECODMODP,NFEMIELETRO,ENT_CODCID,ENT_RUA,UFNFe,"
                  SQL = SQL & " GRAVATEXTO,ENVIONFE,CRT,VersaoNfe"
               SQL = SQL & " )"
               SQL = SQL & " values("
                  SQL = SQL & "'" & EMPRESA_A & "'"               'empresa_id_n  sempre vai ser 01
                  SQL = SQL & ",'" & Trim(txtFilial.Text) & "'"
                  SQL = SQL & ",'" & NOME_REDUZ & "'"
                  SQL = SQL & ",'" & FILIAL_REDUZ & "'"
                  SQL = SQL & ",'" & NOME_EMPRESA & "'"
                  SQL = SQL & ",'" & TELEFONE & "'"
                  SQL = SQL & ",'" & FAX & "'"
                  SQL = SQL & ",'" & CNPJ & "'"
                  SQL = SQL & ",'" & INSC_ESTADUAL & "'"
                  SQL = SQL & ",'" & ENT_RUA_AV & "'"
                  SQL = SQL & ",'" & ENT_COMPLEMENTO & "'"
                  SQL = SQL & ",'" & ENT_BAIRRO & "'"
                  SQL = SQL & ",'" & ENT_CIDADE & "'"
                  SQL = SQL & ",'" & ENT_ESTADO & "'"
                  SQL = SQL & ",'" & ENT_CEP & "'"
                  SQL = SQL & ",'" & COB_RUA_AV & "'"
                  SQL = SQL & ",'" & COB_COMPLEMENTO & "'"
                  SQL = SQL & ",'" & COB_BAIRRO & "'"
                  SQL = SQL & ",'" & COB_CIDADE & "'"
                  SQL = SQL & ",'" & COB_ESTADO & "'"
                  SQL = SQL & ",'" & COB_CEP & "'"
                  SQL = SQL & ",'" & TIPO_INSCRICAO & "'"
                  SQL = SQL & ",'" & INSC_SUFRAMA & "'"
                  SQL = SQL & ",'" & NFESERIEPA & "'"
                  SQL = SQL & ",'" & NFETPAMBI & "'"
                  SQL = SQL & ",'" & NFECODMODP & "'"
                  SQL = SQL & ",'" & NFEMIELETRO & "'"
                  SQL = SQL & ",'" & ENT_CODCID & "'"
                  SQL = SQL & ",'" & ENT_RUA & "'"
                  SQL = SQL & ",'" & UFNFe & "'"
                  SQL = SQL & ",'" & GRAVATEXTO & "'"
                  SQL = SQL & ",'" & ENVIONFE & "'"
                     'SQL = SQL & ",'" & CRIADIRETORIO & "'"
                     'SQL = SQL & ",'" & CERTITEXTO & "'"
                  SQL = SQL & ",'" & CRT & "'"
                  SQL = SQL & ",'" & VersaoNfe & "'"
               SQL = SQL & " )"
               Else  'UPDATE
                  SQL = "update EMPRES set "
                     SQL = SQL & "empresa = '" & EMPRESA_A & "'"
                     SQL = SQL & ",filial = '" & txtFilial.Text & "'"
                     SQL = SQL & ",NOME_REDUZ = '" & NOME_REDUZ & "'"
                     SQL = SQL & ",FILIAL_REDUZ = '" & FILIAL_REDUZ & "'"
                     SQL = SQL & ",NOME_EMPRESA = '" & NOME_EMPRESA & "'"
                     SQL = SQL & ",TELEFONE = '" & TELEFONE & "'"
                     SQL = SQL & ",FAX = '" & FAX & "'"
                     SQL = SQL & ",CNPJ = '" & CNPJ & "'"
                     SQL = SQL & ",INSC_ESTADUAL = '" & INSC_ESTADUAL & "'"
                     SQL = SQL & ",ENT_RUA_AV = '" & ENT_RUA_AV & "'"
                     SQL = SQL & ",ENT_COMPLEMENTO = '" & ENT_COMPLEMENTO & "'"
                     SQL = SQL & ",ENT_BAIRRO = '" & ENT_BAIRRO & "'"
                     SQL = SQL & ",ENT_CIDADE = '" & ENT_CIDADE & "'"
                     SQL = SQL & ",ENT_ESTADO = '" & ENT_ESTADO & "'"
                     SQL = SQL & ",ENT_CEP = '" & ENT_CEP & "'"
                     SQL = SQL & ",COB_RUA_AV = '" & COB_RUA_AV & "'"
                     SQL = SQL & ",COB_COMPLEMENTO = '" & COB_COMPLEMENTO & "'"
                     SQL = SQL & ",COB_BAIRRO = '" & COB_BAIRRO & "'"
                     SQL = SQL & ",COB_CIDADE = '" & COB_CIDADE & "'"
                     SQL = SQL & ",COB_ESTADO = '" & COB_ESTADO & "'"
                     SQL = SQL & ",COB_CEP = '" & COB_CEP & "'"
                     SQL = SQL & ",TIPO_INSCRICAO = '" & TIPO_INSCRICAO & "'"
                     SQL = SQL & ",INSC_SUFRAMA = '" & INSC_SUFRAMA & "'"
                     SQL = SQL & ",NFESERIEPA = '" & NFESERIEPA & "'"
                     SQL = SQL & ",NFETPAMBI = '" & NFETPAMBI & "'"
                     SQL = SQL & ",NFECODMODP = '" & NFECODMODP & "'"
                     SQL = SQL & ",NFEMIELETRO = '" & NFEMIELETRO & "'"
                     SQL = SQL & ",ENT_CODCID = '" & ENT_CODCID & "'"
                     SQL = SQL & ",ENT_RUA = 'AV'" '& ENT_RUA & "'"
                     SQL = SQL & ",UFNFe = '" & UFNFe & "'"
                     SQL = SQL & ",GRAVATEXTO = '" & GRAVATEXTO & "'"
                     SQL = SQL & ",ENVIONFE = '" & ENVIONFE & "'"
                        'SQL = SQL & ",CRIADIRETORIO = '" & CRIADIRETORIO & "'"
                        'SQL = SQL & ",CERTITEXTO = '" & CERTITEXTO & "'"
                     SQL = SQL & ",CRT = '" & CRT & "'"
                     SQL = SQL & ",VersaoNfe = '" & VersaoNfe & "'"
                  SQL = SQL & " where cnpj = '" & Trim(txtCNPJ.Text) & "'"
            End If
            If TabEmpres.State = 1 Then _
               TabEmpres.Close

'Debug.Print SQL

         CONECTA_GLOBAL.Execute SQL
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GLOBAL_EMPRES"
End Sub

Sub MOSTRA_EMPRES()
'On Error GoTo ERRO_TRATA

   If Trim(txtCNPJ.Text) <> "" And Trim(txtRazao.Text) <> "" And Trim(txtFant.Text) <> "" And PESSOA_ID_N > 0 Then
      txtEmpresa.Text = ""
      txtFilial.Text = ""
      txtAmbiente.Text = ""
      txtENVIONFE.Text = ""

      ABRE_BANCO_GLOBAL

      If CONECTA_GLOBAL.State <> 1 Then
         MsgBox "Banco GLOBAL não conectado."
         Else
            If TabEmpres.State = 1 Then _
               TabEmpres.Close

            SQL = "select ENVIONFE,EMPRESA,FILIAL,NFETPAMBI from EMPRES WITH (NOLOCK) "
               SQL = SQL & " where cnpj = '" & Trim(txtCNPJ.Text) & "'"
            TabEmpres.Open SQL, CONECTA_GLOBAL, , , adCmdText
            If Not TabEmpres.EOF Then
               txtEmpresa.Text = "" & Trim(TabEmpres.Fields("empresa").Value)
               txtFilial.Text = "" & Trim(TabEmpres.Fields("filial").Value)
               txtAmbiente.Text = "" & Trim(TabEmpres.Fields("NFETPAMBI").Value)
               txtENVIONFE.Text = "" & Trim(TabEmpres.Fields("ENVIONFE").Value)
            End If
            If TabEmpres.State = 1 Then _
               TabEmpres.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_EMPRES"
End Sub
