VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSVeiculoCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Veículo"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "OSVEICULOCADASTRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "OSVEICULOCADASTRO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCpf"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCNPJCPF"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDtIni"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbComb"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbCor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMotor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDescricao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbTIPO"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMODELO"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtANO"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtNome"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCHASSI"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbCombAUX"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbCorAUX"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbTipoAUX"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmbMarca"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmbMarcaAUX"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdCadCli"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdConsCli"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdConsVeiculo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtPlaca"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "&Histórico"
      TabPicture(1)   =   "OSVEICULOCADASTRO.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstOS"
      Tab(1).Control(1)=   "lstProduto"
      Tab(1).Control(2)=   "lstServico"
      Tab(1).Control(3)=   "lstOBs"
      Tab(1).ControlCount=   4
      Begin PVMaskEditLib.PVMaskEdit txtPlaca 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1695
         _Version        =   524288
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "@@@-####"
      End
      Begin VB.CommandButton cmdConsVeiculo 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3360
         Picture         =   "OSVEICULOCADASTRO.frx":5C4A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Pesquisa Equipamento"
         Top             =   480
         Width           =   405
      End
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3050
         Picture         =   "OSVEICULOCADASTRO.frx":664C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Pesquisa Veículo"
         Top             =   3480
         Width           =   405
      End
      Begin VB.CommandButton cmdCadCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3465
         Picture         =   "OSVEICULOCADASTRO.frx":704E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Consulta Cadastro Veículo"
         Top             =   3480
         Width           =   405
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
         Left            =   5760
         TabIndex        =   31
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbMarca 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5400
         TabIndex        =   9
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox cmbTipoAUX 
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
         Left            =   1800
         TabIndex        =   29
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbCorAUX 
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
         Left            =   2040
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbCombAUX 
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
         Left            =   5760
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCHASSI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   13
         Top             =   3480
         Width           =   3735
      End
      Begin VB.TextBox txtANO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtMODELO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cmbTIPO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   8
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   6135
      End
      Begin VB.TextBox txtMotor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1920
         Width           =   3015
      End
      Begin VB.ComboBox cmbCor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   6
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox cmbComb 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5400
         TabIndex        =   7
         Top             =   2400
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   405
         Left            =   6360
         TabIndex        =   14
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   405
         Left            =   960
         TabIndex        =   10
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstOS 
         Height          =   3465
         Left            =   -74955
         TabIndex        =   26
         Top             =   360
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   6112
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "O.S."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstProduto 
         Height          =   1185
         Left            =   -72360
         TabIndex        =   32
         Tag             =   "Produtos O.S."
         Top             =   360
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   2090
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qtde."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DtGarantia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Marca"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lstServico 
         Height          =   1185
         Left            =   -72360
         TabIndex        =   33
         Tag             =   "Serviços O.S."
         Top             =   1560
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   2090
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serviço"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Responsável"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView lstOBs 
         Height          =   1065
         Left            =   -72360
         TabIndex        =   34
         Tag             =   "Serviços O.S."
         Top             =   2760
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   1879
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Obs"
            Object.Width           =   195987
         EndProperty
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
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
         Left            =   4095
         TabIndex        =   30
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chassi:"
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
         Left            =   720
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
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
         Left            =   4680
         TabIndex        =   23
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Modelo:"
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
         Left            =   6000
         TabIndex        =   22
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Cadastro:"
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
         Left            =   4560
         TabIndex        =   21
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Veículo:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
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
         Left            =   840
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc/Modelo:"
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
         Left            =   180
         TabIndex        =   18
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motor:"
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
         Left            =   840
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cor:"
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
         Left            =   1080
         TabIndex        =   16
         Top             =   2400
         Width           =   390
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Combustível:"
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
         Left            =   3600
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":9178
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":95CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":98E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":9D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":A190
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":A4B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":A904
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":AC24
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":C92E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":D608
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":1322A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Termo"
            Key             =   "termo"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   7845
      DesignHeight    =   4575
   End
End
Attribute VB_Name = "frmOSVeiculoCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim COR_ID_N         As Long
   Dim MARCA_ID_N       As Long
   Dim TIPO_EQP_ID_N    As Long
   Dim ANO_N            As Long
   Dim MODELO_ID_N      As Long
   Dim VEICULO_ID_N     As Long
   Dim OSEQUIPAMENTO_ID_N As Long
   Dim COMBUSTIVEL_ID_N As Long
   Dim ANO_ID_N         As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   'SSTab1.TabVisible(1) = False

   CARREGA_DESCRITORES

   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
   Toolbar1.Buttons(7).Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If conecta_retaguarda.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(7).Visible = False
   If SSTab1.Tab = 0 Then _
      txtPlaca.SetFocus

   If SSTab1.Tab = 1 Then
      MOSTRA_HISTORICO
      Toolbar1.Buttons(7).Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSTab1_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "termo"
         MOSTRA_TERMO
      Case "consultar"
         INDR_OS_VEICULO = True
         txtPlaca.Text = "" & CONSULTA_EQP_VEICULO
      Case "voltar"
         Unload Me
      Case "matar"
         If Trim(txtPlaca.Text) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from OSVEICULO WITH (NOLOCK)"
            SQL = SQL & " where placa = '" & Trim(txtPlaca.Text) & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If TabAUX.State = 1 Then _
                  TabAUX.Close

               SQL = "select * from OS WITH (NOLOCK)"
               SQL = SQL & " where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
               TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabAUX.EOF Then
                  If TabAUX.State = 1 Then _
                     TabAUX.Close

                  If TabTemp.State = 1 Then _
                     TabTemp.Close
                  
                  MsgBox "Impossível excluir, veículo possue movimentação na oficina."
                  Exit Sub
               End If
               If TabAUX.State = 1 Then _
                  TabAUX.Close

               Msg = "Confirma Exclusão do VEICULO ?"
               PERGUNTA Msg, vbYesNo + 32, "", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  SQL = "delete from OSVEICULO "
                  SQL = SQL & " where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
                  CONECTA_RETAGUARDA.Execute SQL

                  SQL = "delete from OSEQUIPAMENTO "
                  SQL = SQL & " where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
                  CONECTA_RETAGUARDA.Execute SQL

                  LIMPA_VEICULO
                  txtPlaca.SetFocus
               End If
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      Case "gravar"
         GRAVA_VEICULO
         txtPlaca.SetFocus
      Case "limpar"
         LIMPA_VEICULO
      Case "imprimir"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsVeiculo_Click()
'On Error GoTo ERRO_TRATA

   INDR_OS_VEICULO = True
   txtPlaca.Text = "" & CONSULTA_EQP_VEICULO
   If Trim(txtPlaca.Text) <> "" Then _
      Call txtPLACA_LostFocus
   txtPlaca.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsVeiculo_Click"
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      'Call TXTCNPJCPF_LostFocus

      'Exit Sub
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsCli_Click"
End Sub

Private Sub cmdCadCli_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaCadastro.Show 1
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmbComb_LostFocus()
   cmbComb.BackColor = &HFFFFFF
End Sub

Private Sub cmbCor_LostFocus()
   cmbCor.BackColor = &HFFFFFF
End Sub

Private Sub cmbMarca_LostFocus()
   cmbMarca.BackColor = &HFFFFFF
End Sub

Private Sub cmbTIPO_LostFocus()
   cmbTIPO.BackColor = &HFFFFFF
End Sub

Private Sub txtANO_LostFocus()
   txtANO.BackColor = &HFFFFFF
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF)
   txtCNPJCPF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      ENDERECO_A = ""
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         'MsgBox "Informe CNPJ/CPF corretamente"
         txtCNPJCPF.Text = "99999999999"
         Else
            If Len(txtCNPJCPF.Text) > 0 Then
               Select Case Len(txtCNPJCPF.Text)
                  Case Is = 11
                    If Not CALCULACPF(txtCNPJCPF.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtCNPJCPF.PromptInclude = False
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    End If
                  Case Is = 14
                    If Not VALIDACGC(txtCNPJCPF.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtCNPJCPF.PromptInclude = False
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    End If
                  Case Is > 14
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCNPJCPF = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
                  Case Is < 11
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCNPJCPF = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
               End Select
               Else
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCNPJCPF = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End If
            txtCNPJCPF.PromptInclude = False
            CRITERIO_A = "" & txtCNPJCPF.Text
      End If
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         CRITERIO_A = "" & txtCNPJCPF.Text
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
         End If
         txtCNPJCPF.Text = CRITERIO_A
      End If
      txtCNPJCPF.PromptInclude = False

      txtPlaca.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
   txtCNPJCPF.PromptInclude = False
   PESSOA_ID_N = 0

   If Trim(txtCNPJCPF.Text) <> "" Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         'txtCNPJCPF.SetFocus
         Exit Sub
         Else
            If TabCliente!NOME <> "" Then
               txtNome.Text = TabCliente!NOME
               PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
            End If
      End If

      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If
   txtCNPJCPF.BackColor = &HFFFFFF
End Sub

Private Sub txtCHASSI_GotFocus()
   txtCHASSI.SelStart = 0
   txtCHASSI.SelLength = Len(txtCHASSI)
   txtCHASSI.BackColor = &HC0FFFF
End Sub

Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtMotor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCHASSI_KeyPress"
End Sub

Private Sub txtDescricao_LostFocus()
   txtDescricao.BackColor = &HFFFFFF
End Sub

Private Sub txtMODELO_LostFocus()
   txtMODELO.BackColor = &HFFFFFF
End Sub

Private Sub txtMotor_LostFocus()
   txtMotor.BackColor = &HFFFFFF
End Sub

Private Sub txtNome_GotFocus()
   txtNome.SelStart = 0
   txtNome.SelLength = Len(txtNome)
   txtNome.BackColor = &HC0FFFF
End Sub

Private Sub txtNome_LostFocus()
   txtNome.BackColor = &HFFFFFF
End Sub

Private Sub txtPlaca_GotFocus()
   txtPlaca.BackColor = &HC0FFFF
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtPlaca.Text = "" & CONSULTA_EQP_VEICULO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPLACA_KeyDown"
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescricao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtplaca_KeyPress"
End Sub

Private Sub txtPLACA_LostFocus()
   If Trim(txtPlaca.Text) <> "" Then _
      MOSTRA_VEICULO
   txtPlaca.BackColor = &HFFFFFF
End Sub

Private Sub txtDescricao_GotFocus()
   txtDescricao.SelStart = 0
   txtDescricao.SelLength = Len(txtDescricao)
   txtDescricao.BackColor = &HC0FFFF
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCHASSI.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescricao_KeyPress"
End Sub

Private Sub txtMotor_GotFocus()
   txtMotor.SelStart = 0
   txtMotor.SelLength = Len(txtMotor)
   txtMotor.BackColor = &HC0FFFF
End Sub

Private Sub txtmotor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtANO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtmotor_KeyPress"
End Sub

Private Sub txtCHASSI_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtCHASSI.Text = "" Then
   '   txtCHASSI.Text = txtCHASSI.Text
   '   MsgBox "Chassi inválido."
   '   txtCHASSI.SetFocus
   '   Exit Sub
   End If
   txtCHASSI.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCHASSI_LostFocus"
End Sub

Private Sub txtAno_GotFocus()
   txtANO.SelStart = 0
   txtANO.SelLength = Len(txtANO)
   txtANO.BackColor = &HC0FFFF
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtMODELO.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtANO_KeyPress"
End Sub

Private Sub txtmodelo_GotFocus()
   txtMODELO.SelStart = 0
   txtMODELO.SelLength = Len(txtMODELO)
   txtMODELO.BackColor = &HC0FFFF
End Sub

Private Sub txtMODELO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCor.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMODELO_KeyPress"
End Sub

Private Sub cmbcor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbComb.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_KeyPress"
End Sub

Private Sub cmbComb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTIPO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbComb_KeyPress"
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbMarca.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPO_KeyPress"
End Sub

Private Sub cmbTipo_GotFocus()
   cmbTIPO.SelStart = 0
   cmbTIPO.SelLength = Len(cmbTIPO)
   cmbTIPO.BackColor = &HC0FFFF
End Sub

Private Sub cmbTipo_Click()
'On Error GoTo ERRO_TRATA

   cmbTipoAUX.ListIndex = cmbTIPO.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipo_Click"
End Sub

Private Sub cmbMarca_GotFocus()
   cmbMarca.SelStart = 0
   cmbMarca.SelLength = Len(cmbMarca)
   cmbMarca.BackColor = &HC0FFFF
End Sub

Private Sub cmbmarca_Click()
'On Error GoTo ERRO_TRATA

   cmbMarcaAUX.ListIndex = cmbMarca.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmarca_Click"
End Sub

Private Sub cmbmarca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmarca_KeyPress"
End Sub

Private Sub cmbCor_GotFocus()
   cmbCor.SelStart = 0
   cmbCor.SelLength = Len(cmbCor)
   cmbCor.BackColor = &HC0FFFF
End Sub

Private Sub cmbcor_Click()
'On Error GoTo ERRO_TRATA

   cmbCorAUX.ListIndex = cmbCor.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_Click"
End Sub

Private Sub cmbComb_GotFocus()
   cmbComb.SelStart = 0
   cmbComb.SelLength = Len(cmbComb)
   cmbComb.BackColor = &HC0FFFF
End Sub

Private Sub cmbComb_Click()
'On Error GoTo ERRO_TRATA

   cmbCombAUX.ListIndex = cmbComb.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbComb_Click"
End Sub

Private Sub LSTos_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text
         MOSTRA_PRODUTO
         MOSTRA_SERVICO
         MOSTRA_OBS
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LSTos_Click"
End Sub

Private Sub MOSTRA_VEICULO()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   If Trim(txtPlaca.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwVeiculo WITH (NOLOCK)"
      SQL = SQL & " where placa = '" & Trim(txtPlaca.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         txtNome.Text = "" & Trim(TabTemp.Fields("DescPessoa").Value)
         If Trim(txtNome.Text) = "" Then _
            txtNome.Text = "" & Trim(TabTemp.Fields("DescPessoa").Value)
         PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value

         txtCHASSI.Text = "" & Trim(TabTemp!CHASSI)
         txtDescricao.Text = "" & TabTemp!DESCRICAO
         txtMotor.Text = "" & TabTemp!MOTOR
         txtANO.Text = "" & TabTemp!Ano
         txtMODELO.Text = "" & TabTemp!MODELO

         If Not IsNull(TabTemp.Fields("cor_id").Value) Then
            If IsNumeric(TabTemp.Fields("cor_id").Value) Then
               cmbCor.Text = "" & TRAZ_DESCRITOR("S", TabTemp.Fields("cor_id").Value)
               cmbCorAUX.Text = "" & TabTemp.Fields("cor_id").Value
            End If
         End If

         If Not IsNull(TabTemp.Fields("marca_id").Value) Then
            If IsNumeric(TabTemp.Fields("marca_id").Value) Then
               cmbMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)
               cmbMarcaAUX.Text = "" & TabTemp.Fields("marca_id").Value
            End If
         End If

         If Not IsNull(TabTemp!COMBUSTIVEL_ID) Then
            If IsNumeric(TabTemp.Fields("combustivel_id").Value) Then
               cmbComb.Text = "" & TRAZ_DESCRITOR("U", TabTemp.Fields("combustivel_id").Value)
               cmbCombAUX.Text = "" & TabTemp.Fields("combustivel_id").Value
            End If
         End If

         If Not IsNull(TabTemp!tipo_veiculo_id) Then
            If IsNumeric(TabTemp!tipo_veiculo_id) Then
               cmbTIPO.Text = "" & TRAZ_DESCRITOR("A", TabTemp!tipo_veiculo_id)
               cmbTipoAUX.Text = "" & TabTemp!tipo_veiculo_id
            End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VEICULO"
End Sub

Private Sub LIMPA_VEICULO()
'On Error GoTo ERRO_TRATA

   COR_ID_N = 0
   MARCA_ID_N = 0
   TIPO_EQP_ID_N = 0
   ANO_N = 0
   MODELO_ID_N = 0
   VEICULO_ID_N = 0
   OSEQUIPAMENTO_ID_N = 0
   COMBUSTIVEL_ID_N = 0
   PESSOA_ID_N = 0
   txtPlaca.Text = ""
   txtDescricao.Text = ""
   txtMotor.Text = ""
   txtCHASSI.Text = ""
   cmbCor.Text = ""
   cmbCorAUX.Text = ""
   cmbCombAUX.Text = ""
   cmbComb.Text = ""
   cmbMarcaAUX.Text = ""
   cmbMarca.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   cmbTIPO.Text = ""
   cmbTipoAUX.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   'MOSTRA_HISTORICO
   txtPlaca.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_VEICULO"
End Sub

Private Sub GRAVA_VEICULO()
'On Error GoTo ERRO_TRATA

   If Trim(txtPlaca.Text) = "" Then
      MsgBox "Número de Placa deve ser informado."
      txtPlaca.SetFocus
      Exit Sub
   End If
   'If Trim(txtCHASSI.Text) = "" Then
   '   MsgBox "Número de Chassi deve ser informado."
   '   txtCHASSI.SetFocus
   '   Exit Sub
   'End If
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Cliente deve ser informado."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Descrição do Veículo deve ser informada."
      txtDescricao.SetFocus
      Exit Sub
   End If

   COR_ID_N = 0 & cmbCorAUX.Text
   MARCA_ID_N = 0 & cmbMarcaAUX.Text
   TIPO_EQP_ID_N = 0 & cmbTipoAUX.Text
   ANO_ID_N = 0 & txtANO.Text
   MODELO_ID_N = 0 & txtMODELO.Text
   COMBUSTIVEL_ID_N = 0 & cmbCombAUX.Text

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSVEICULO WITH (NOLOCK)"
   SQL = SQL & " where placa = '" & Trim(txtPlaca.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      VEICULO_ID_N = MAX_ID("veiculo_id", "OSVEICULO", "", "", "", "")
   
      SQL = "insert into OSVEICULO "
      SQL = SQL & "("
         SQL = SQL & "VEICULO_ID,PESSOA_ID,PLACA,DESCRICAO,MOTOR,CHASSI,numr_frota,"
         SQL = SQL & "Ano,MODELO,COMBUSTIVEL_ID,COR_ID,TIPO_VEICULO_ID,MARCA_ID"
      SQL = SQL & ")"

      SQL = SQL & " values("
         SQL = SQL & VEICULO_ID_N                           'VEICULO_ID
         SQL = SQL & "," & PESSOA_ID_N                      'PESSOA_ID
         SQL = SQL & ",'" & Trim(txtPlaca.Text) & "'"       'PLACA
         SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'DESCRICAO
         SQL = SQL & ",'" & Trim(txtMotor.Text) & "'"       'MOTOR
         SQL = SQL & ",'" & Trim(txtCHASSI.Text) & "'"      'CHASSI
         SQL = SQL & ",'0'"                                 'numr_frota
         SQL = SQL & "," & ANO_ID_N                         'ANO
         SQL = SQL & "," & MODELO_ID_N                      'modelo
         SQL = SQL & "," & COMBUSTIVEL_ID_N                 'COMBUSTIVEL_ID
         SQL = SQL & "," & COR_ID_N                         'COR_ID
         SQL = SQL & "," & TIPO_EQP_ID_N                    'tipo_veiculo_id
         SQL = SQL & "," & MARCA_ID_N                       'MARCA_ID
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "update OSVEICULO "
         SQL = SQL & "set "
            SQL = SQL & " PESSOA_ID = " & PESSOA_ID_N                      'PESSOA_ID
            SQL = SQL & ", PLACA = '" & Trim(txtPlaca.Text) & "'"          'PLACA
            SQL = SQL & ", DESCRICAO = '" & Trim(txtDescricao.Text) & "'"  'DESCRICAO
            SQL = SQL & ", MOTOR = '" & Trim(txtMotor.Text) & "'"          'MOTOR
            SQL = SQL & ", CHASSI = '" & Trim(txtCHASSI.Text) & "'"        'CHASSI
            SQL = SQL & ", numr_frota = '0'"                               'numr_frota
            SQL = SQL & ", ANO = " & ANO_ID_N                              'ANO
            SQL = SQL & ", modelo = " & MODELO_ID_N                        'modelo
            SQL = SQL & ", COMBUSTIVEL_ID = " & COMBUSTIVEL_ID_N            'COMBUSTIVEL_ID
            SQL = SQL & ", COR_ID = " & COR_ID_N                           'COR_ID
            SQL = SQL & ", tipo_veiculo_id = " & TIPO_EQP_ID_N                    'TIPO_EQP
            SQL = SQL & ", MARCA_ID = " & MARCA_ID_N                       'MARCA_ID
         SQL = SQL & " where placa = '" & Trim(txtPlaca.Text) & "'"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   LIMPA_VEICULO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_VEICULO"
End Sub

Sub CARREGA_DESCRITORES()
'On Error GoTo ERRO_TRATA

'Tipo Função
' A   Tipo Veículo
' S   Cor
' U   Combustivel
' W   Marca

   cmbTipoAUX.Clear
   cmbTIPO.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where tipo = 'A' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTIPO.AddItem Trim(TabDESCR!DESCRICAO)
      cmbTipoAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbCorAUX.Clear
   cmbCor.Clear

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where tipo = 'S' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbCor.AddItem Trim(TabDESCR!DESCRICAO)
      cmbCorAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close
   
   cmbCombAUX.Clear
   cmbComb.Clear

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where tipo = 'U' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbComb.AddItem Trim(TabDESCR!DESCRICAO)
      cmbCombAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbMarcaAUX.Clear
   cmbMarca.Clear

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where tipo = 'W' "
   SQL = SQL & "order by descricao"
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
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_DESCRITORES"
End Sub

Private Sub MOSTRA_HISTORICO()
'On Error GoTo ERRO_TRATA

   lstOS.ListItems.Clear

   If Trim(txtPlaca.Text) <> "" Then
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from vwOSVeiculo WITH (NOLOCK)"
      SQL = SQL & " where placa = '" & Trim(txtPlaca.Text) & "'"
      SQL = SQL & " order by dt_OS desc"

      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabAUX.EOF
         Set item = lstOS.ListItems.Add(, "seq." & TabAUX.Fields("os_id").Value, TabAUX.Fields("os_id").Value)
         item.SubItems(1) = TabAUX.Fields("dt_os").Value
         TabAUX.MoveNext
      Wend
      If TabAUX.State = 1 Then _
         TabAUX.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_HISTORICO"
End Sub

Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

   lstProduto.Visible = False
   lstProduto.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select OSPECA.OSPECA_ID, OSPECA.OS_ID, OSPECA.PRODUTO_ID, OSPECA.DT_CAD, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, OSPECA.QTDE, OSPECA.DT_GARANTIA, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.UNIDADE_MEDIDA, PRODUTO.SITUACAO, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_Venda,"
   SQL = SQL & " Produto.PRECO_ATACADO , Produto.MARCA_ID"
   SQL = SQL & " FROM OSPECA  WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      QTDE_N = 0 & TabCabeca.Fields("qtde").Value
      VALOR_ITEM_N = 0 & TabCabeca.Fields("valor_item").Value
      Set item = lstProduto.ListItems.Add(, "seq." & CONT_N, Trim(TabCabeca.Fields("codg_produto").Value))

      item.SubItems(1) = "" & Trim(TabCabeca.Fields("descricao").Value)
      item.SubItems(2) = "" & Format(QTDE_N, strFormatacao3Digitos)
      item.SubItems(3) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      item.SubItems(4) = "" & Trim(TabCabeca.Fields("dt_garantia").Value)
      If Not IsNull(TabCabeca.Fields("marca_id").Value) Then _
         item.SubItems(5) = "" & TRAZ_DESCRITOR("W", TabCabeca.Fields("marca_id").Value)

      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstProduto.Refresh
   lstProduto.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub MOSTRA_SERVICO()
'On Error GoTo ERRO_TRATA

   lstServico.Visible = False
   lstServico.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from OSSERVICO WITH (NOLOCK) "
   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      VALOR_ITEM_N = 0 & TabCabeca.Fields("valor_servico").Value
      Set item = lstServico.ListItems.Add(, "seq." & CONT_N, TabCabeca.Fields("ostarefa_id").Value)

      item.SubItems(1) = "" & Trim(TabCabeca.Fields("descricao").Value)
      item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      If Not IsNull(TabCabeca.Fields("responsavel_id").Value) Then _
         item.SubItems(3) = "" & TRAZ_NOME_USUARIO(TabCabeca.Fields("responsavel_id").Value)

      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstServico.Refresh
   lstServico.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_SERVICO"
End Sub

Sub MOSTRA_OBS()
'On Error GoTo ERRO_TRATA

   lstOBs.Visible = False
   lstOBs.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * FROM OSOBS  WITH (NOLOCK) "
   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      Set item = lstOBs.ListItems.Add(, "seq." & CONT_N, Trim(TabCabeca.Fields("OBS").Value))
      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstOBs.Refresh
   lstOBs.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OBS"
End Sub

Sub MOSTRA_TERMO()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text

         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from OS "
         SQL = SQL & " where os_id = " & OS_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            CHAMADA_A = "TERMOGARANTIA"

            frmOBS.txtOBS.Enabled = False
            frmOBS.Show 1

            Else
               If TabCabeca.State = 1 Then _
                  TabCabeca.Close
               MsgBox "O.S. não informada !!!"
         End If
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
      End If
   End If
   CHAMADA_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TERMO"
End Sub
