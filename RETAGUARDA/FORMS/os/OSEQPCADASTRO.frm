VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSEqpCadastro 
   Caption         =   "Cadastro de Equipamento"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "OSEQPCADASTRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSEQPCADASTRO.frx":739E
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
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formul�rio"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "cons"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetiva��o da comiss�o"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impress�o"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "importa"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   7965
      _ExtentX        =   14049
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
      TabPicture(0)   =   "OSEQPCADASTRO.frx":76BE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCpf"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCNPJCPF"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDtIni"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCadCliente"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdConsultaCliente"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdConsultaEqp"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbMarca"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtReferencia"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtNome"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtANO"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMODELO"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmbTIPO"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtDescricao"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbCor"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtEqp"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbTipoAUX"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbCorAUX"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmbMarcaAUX"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&Hist�rico"
      TabPicture(1)   =   "OSEQPCADASTRO.frx":76DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstOS"
      Tab(1).Control(1)=   "lstProduto"
      Tab(1).Control(2)=   "lstServico"
      Tab(1).Control(3)=   "lstOBs"
      Tab(1).ControlCount=   4
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
         TabIndex        =   29
         Top             =   2520
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
         Left            =   5400
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
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
         Left            =   1560
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEqp 
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   0
         Top             =   480
         Width           =   1335
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
         Left            =   5400
         TabIndex        =   4
         Top             =   1920
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
         TabIndex        =   3
         Top             =   1920
         Width           =   2295
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
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2520
         Width           =   735
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
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   5
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNome 
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
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   10
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtReferencia 
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
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdConsultaEqp 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3000
         Picture         =   "OSEQPCADASTRO.frx":76F6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Pesquisa Equipamento"
         Top             =   480
         Width           =   405
      End
      Begin VB.CommandButton cmdConsultaCliente 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3050
         Picture         =   "OSEQPCADASTRO.frx":80F8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consultar Cliente"
         Top             =   3120
         Width           =   405
      End
      Begin VB.CommandButton cmdCadCliente 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3480
         Picture         =   "OSEQPCADASTRO.frx":8AFA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cadastrar Cliente"
         Top             =   3120
         Width           =   405
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   405
         Left            =   6360
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   3120
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
         TabIndex        =   30
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
         Left            =   -72285
         TabIndex        =   31
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
            Text            =   "C�digo"
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
         Left            =   -72285
         TabIndex        =   32
         Tag             =   "Servi�os O.S."
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
            Text            =   "Servi�o"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descri��o"
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
            Text            =   "Respons�vel"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView lstOBs 
         Height          =   1065
         Left            =   -72285
         TabIndex        =   33
         Tag             =   "Servi�os O.S."
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Left            =   4680
         TabIndex        =   25
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc./Modelo:"
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
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipamento:"
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
         TabIndex        =   23
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Eqp:"
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
         Left            =   600
         TabIndex        =   22
         Top             =   1920
         Width           =   900
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
         Left            =   2400
         TabIndex        =   20
         Top             =   2520
         Width           =   885
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
         Left            =   960
         TabIndex        =   19
         Top             =   2520
         Width           =   495
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
         TabIndex        =   18
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Refer�ncia:"
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
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   4650
         TabIndex        =   16
         Top             =   2520
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar barEQP 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   4785
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "OSEQPCADASTRO.frx":AC24
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   7890
      DesignHeight    =   5160
   End
End
Attribute VB_Name = "frmOSEqpCadastro"
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
   Dim EQP_ID_N     As Long
   Dim OSEQUIPAMENTO_ID_N As Long
   Dim COMBUSTIVEL_ID_N As Long
   Dim ANO_ID_N         As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   CARREGA_DESCRITORES

   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

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

   If SSTab1.Tab = 0 Then _
      txtEqp.SetFocus

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
      Case "cons"
         SQL3 = ""
         frmOSEqpConsulta.Show 1
         If SQL3 <> "" Then _
            txtEqp.Text = SQL3
         SQL3 = ""
         txtEqp.SetFocus
      Case "voltar"
         Unload Me
      Case "matar"
         If Trim(txtEqp.Text) <> "" Then
            If Not IsNumeric(txtEqp.Text) Then _
               Exit Sub

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from OSEQUIPAMENTO"
            SQL = SQL & " where equipamento_id = " & txtEqp.Text
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If TabAUX.State = 1 Then _
                  TabAUX.Close

               SQL = "select * from OS "
               SQL = SQL & " where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
               TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabAUX.EOF Then
                  If TabAUX.State = 1 Then _
                     TabAUX.Close

                  If TabTemp.State = 1 Then _
                     TabTemp.Close
                  
                  MsgBox "Imposs�vel excluir, Equipamento possue movimenta��o na oficina."
                  Exit Sub
               End If
               If TabAUX.State = 1 Then _
                  TabAUX.Close
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      Case "gravar"
         GRAVA_EQP
         txtEqp.SetFocus
      Case "limpar"
         LIMPA_EQP
      Case "imprimir"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F4 - Cadastra Cliente", "F7 - Consulta Clientes", "Informe propriet�rio do Equipamento", ""

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Mask = "##############"
      If CNPJCPF_A <> "" Then
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
      End If
   End If

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
         txtCNPJCPF.SetFocus
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.Mask = "##############"
            txtCNPJCPF.Text = CNPJCPF_A
         End If
         CNPJCPF_A = ""
         txtCNPJCPF.SetFocus
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
            CRITERIO_A = txtCNPJCPF.Text
      End If
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) <> "" Then
         CRITERIO_A = txtCNPJCPF.Text
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
         End If
         txtCNPJCPF.Text = CRITERIO_A
      End If
      txtCNPJCPF.PromptInclude = False
      txtNome.Enabled = True
      txtNome.SetFocus
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

      SQL = "select * from CLIENTE "
      SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF n�o Cadastrado.", vbOKOnly, "Aten��o !!!"
         'txtCNPJCPF.SetFocus
         Exit Sub
         Else
            If TabCliente!NOME <> "" Then
               txtNome.Text = TabCliente!NOME
               PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value

               'If Not IsNull(tabcliente!limite_credito) Then _
                  txtLIMITE.Text = Format(TabCliente!limite_credito, strFormatacao2Digitos)
               'SQL = "select sum(i.valor_item-i.valor_desconto) from ITEMLANCAMENTO i, LANCAMENTO l "
               'SQL = SQL & " where i.numr_doc = l.numr_doc "
               'SQL = SQL & " and l.prop = '" & tabcliente!CGCCPF & "'"
               'SQL = SQL & " and i.status = 'A' "
               'TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               'If Not TABAUX.EOF Then
               '   If Not IsNull(TABAUX.Fields(0).Value) Then
               '      txtPAGAR.Text = FORMAT(TABAUX.Fields(0).Value,strFormatacao2Digitos)
               '      txtPAGAR.Refresh
               '   End If
               'End If
               'TABAUX.Close
            End If
      End If

      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If
End Sub

Private Sub cmdConsultaCliente_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Mask = "##############"
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = "" & Trim(CNPJCPF_A)
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmdCadCliente_Click()
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaCadastro.Show 1
   txtCNPJCPF.SetFocus
End Sub

Private Sub txtNome_GotFocus()
   txtNome.SelStart = 0
   txtNome.SelLength = Len(txtNome)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEqp.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub

Private Sub txtReferencia_GotFocus()
   MOSTRA_RODAPE "Informe a Refer�ncia do Equipamento", "", "", "", ""
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTIPO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtReferencia_KeyPress"
End Sub

Private Sub txtEqp_GotFocus()
   MOSTRA_RODAPE "Informe a Identifica��o do Equipamento", "", "", "", ""
End Sub

Private Sub txtEqp_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSEqpConsulta.Show 1

         If Trim(SQL3) <> "" Then
            If Not IsNumeric(SQL3) Then _
               Exit Sub

            If TabAUX.State = 1 Then _
               TabAUX.Close

            SQL = "select * from OSEQUIPAMENTO "
            SQL = SQL & " where equipamento_id = " & Trim(SQL3)
            TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabAUX.EOF Then _
               MOSTRA_EQP
            If TabAUX.State = 1 Then _
               TabAUX.Close
         End If
         SQL3 = ""
         txtEqp.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEqp_KeyDown"
End Sub

Private Sub txtEqp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescricao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEqp_KeyPress"
End Sub

Private Sub txtEqp_LostFocus()
   If Trim(txtEqp.Text) = "" Then
      OSEQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "OSEQUIPAMENTO", "", "", "", "")
      txtEqp.Text = OSEQUIPAMENTO_ID_N
   End If

   If Not IsNumeric(txtEqp.Text) Then
      OSEQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "OSEQUIPAMENTO", "", "", "", "")
      txtEqp.Text = OSEQUIPAMENTO_ID_N
   End If

   MOSTRA_EQP
End Sub

Private Sub cmdConsultaEQP_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSEqpConsulta.Show 1
   If SQL3 <> "" Then _
      txtEqp.Text = SQL3
   SQL3 = ""
   If Trim(txtEqp.Text) <> "" Then _
      Call txtEqp_LostFocus
   txtEqp.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsultaEQP_Click"
End Sub

Private Sub txtDescricao_GotFocus()
   MOSTRA_RODAPE "Informe a descri��o do Equipamento", "", "", "", ""
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtREFERENCIA.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescricao_KeyPress"
End Sub

Private Sub txtReferencia_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtREFERENCIA.Text = "" Then
      txtREFERENCIA.Text = txtEqp.Text
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtReferencia_LostFocus"
End Sub

Private Sub txtAno_GotFocus()
   MOSTRA_RODAPE "Informe ano de fabrica��o do Equipamento", "", "", "", ""
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtMODELO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtANO_KeyPress"
End Sub

Private Sub txtmodelo_GotFocus()
   MOSTRA_RODAPE "Informe ano do modelo do Equipamento", "", "", "", ""
End Sub

Private Sub txtMODELO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbMarca.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMODELO_KeyPress"
End Sub

Private Sub cmbcor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtANO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_KeyPress"
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPO_KeyPress"
End Sub

Private Sub cmbTipo_GotFocus()
   MOSTRA_RODAPE "Informe tipo do Equipamento", "", "", "", ""
End Sub

Private Sub cmbTipo_Click()
'On Error GoTo ERRO_TRATA

   cmbTipoAUX.ListIndex = cmbTIPO.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipo_Click"
End Sub

Private Sub cmbMarca_GotFocus()
   MOSTRA_RODAPE "Informe marca do Equipamento", "", "", "", ""
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
   MOSTRA_RODAPE "Informe a cor do Equipamento", "", "", "", ""
End Sub

Private Sub cmbcor_Click()
'On Error GoTo ERRO_TRATA

   cmbCorAUX.ListIndex = cmbCor.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_Click"
End Sub

Private Sub MOSTRA_EQP()
'On Error GoTo ERRO_TRATA

   If Trim(txtEqp.Text) <> "" Then
      If Not IsNumeric(txtEqp.Text) Then _
         Exit Sub

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwEQUIPAMENTO "
      SQL = SQL & " where EQUIPAMENTO_ID = " & Trim(txtEqp.Text)
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then

         PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value
         OSEQUIPAMENTO_ID_N = TabTemp.Fields("EQUIPAMENTO_id").Value
         txtEqp.Text = OSEQUIPAMENTO_ID_N

         txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         txtNome.Text = "" & Trim(TabTemp.Fields("Nome_Cliente").Value)

         txtREFERENCIA.Text = "" & Trim(TabTemp.Fields("identificacao").Value)
         txtDescricao.Text = "" & TabTemp!DESCRICAO
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

         If Not IsNull(TabTemp!TIPO_EQP) Then
            If IsNumeric(TabTemp!TIPO_EQP) Then
               cmbTIPO.Text = "" & TRAZ_DESCRITOR("A", TabTemp!TIPO_EQP)
               cmbTipoAUX.Text = "" & TabTemp!TIPO_EQP
            End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_EQP"
End Sub

Private Sub LIMPA_EQP()
'On Error GoTo ERRO_TRATA

   COR_ID_N = 0
   MARCA_ID_N = 0
   TIPO_EQP_ID_N = 0
   ANO_N = 0
   MODELO_ID_N = 0
   EQP_ID_N = 0
   OSEQUIPAMENTO_ID_N = 0
   COMBUSTIVEL_ID_N = 0
   PESSOA_ID_N = 0
   txtEqp.Text = ""
   txtDescricao.Text = ""
   txtREFERENCIA.Text = ""
   cmbCor.Text = ""
   cmbCorAUX.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   cmbTIPO.Text = ""
   cmbTipoAUX.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtEqp.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_EQP"
End Sub

Private Sub GRAVA_EQP()
'On Error GoTo ERRO_TRATA

   If Trim(txtEqp.Text) = "" Then
      MsgBox "N�mero de Identifica��o deve ser informado."
      txtEqp.SetFocus
      Exit Sub
   End If
   'If Trim(txtREFERENCIA.Text) = "" Then
   '   MsgBox "N�mero de Chassi deve ser informado."
   '   txtREFERENCIA.SetFocus
   '   Exit Sub
   'End If
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Cliente deve ser informado."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Descri��o do Equipamento deve ser informada."
      txtDescricao.SetFocus
      Exit Sub
   End If

   COR_ID_N = 0 & cmbCorAUX.Text
   MARCA_ID_N = 0 & cmbMarcaAUX.Text
   TIPO_EQP_ID_N = 0 & cmbTipoAUX.Text
   ANO_ID_N = 0 & txtANO.Text
   MODELO_ID_N = 0 & txtMODELO.Text

   OSEQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "OSEQUIPAMENTO", "", "", "", "")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSEQUIPAMENTO "
   SQL = SQL & " where equipamento_id = " & Trim(txtEqp.Text)
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into OSEQUIPAMENTO "
      SQL = SQL & "(EQUIPAMENTO_ID,DT_CAD,DESCRICAO,IDENTIFICACAO,PESSOA_ID,"
      SQL = SQL & " COR_ID,MARCA_ID,TIPO_EQP,ANO,modelo,nome_cliente)"
      SQL = SQL & " values("
         SQL = SQL & OSEQUIPAMENTO_ID_N                       'EQUIPAMENTO_ID
         SQL = SQL & "," & DMA(Date)                        'DT_CAD
         SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'DESCRICAO
         SQL = SQL & ",'" & Trim(txtREFERENCIA.Text) & "'"  'IDENTIFICACAO
         SQL = SQL & "," & PESSOA_ID_N                      'CLIENTE_ID
         SQL = SQL & "," & COR_ID_N                         'COR_ID
         SQL = SQL & "," & MARCA_ID_N                       'MARCA_ID
         SQL = SQL & "," & TIPO_EQP_ID_N                    'TIPO_EQP
         SQL = SQL & "," & ANO_ID_N                         'ANO
         SQL = SQL & "," & MODELO_ID_N                      'modelo
         SQL = SQL & ",'" & Trim(txtNome.Text) & "'"        'NOME CLIENTE
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         OSEQUIPAMENTO_ID_N = TabTemp.Fields("equipamento_id").Value

         SQL = "update OSEQUIPAMENTO "
         SQL = SQL & "set "

            SQL = SQL & " descricao = '" & Trim(txtDescricao.Text) & "'"         'DESCRICAO
            SQL = SQL & ", IDENTIFICACAO = '" & Trim(txtREFERENCIA.Text) & "'"   'IDENTIFICACAO
            SQL = SQL & ", pessoa_ID = " & PESSOA_ID_N                           'pessoa_ID
            SQL = SQL & ", COR_ID = " & COR_ID_N                                 'COR_ID
            SQL = SQL & ", MARCA_ID = " & MARCA_ID_N                             'MARCA_ID
            SQL = SQL & ", TIPO_EQP = " & TIPO_EQP_ID_N                          'TIPO_EQP
            SQL = SQL & ", ANO = " & ANO_ID_N                                    'ANO
            SQL = SQL & ", nome_cliente = '" & Trim(txtNome.Text) & "'"          'NOME CLIENTE
            SQL = SQL & ", modelo = " & MODELO_ID_N                              'modelo

         SQL = SQL & " where equipamento_id = " & OSEQUIPAMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   LIMPA_EQP

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_EQP"
End Sub

Sub CARREGA_DESCRITORES()
'On Error GoTo ERRO_TRATA

'Tipo Fun��o
' A   Tipo Equipamento
' S   Cor
' U   Combustivel
' W   Marca

   cmbTipoAUX.Clear
   cmbTIPO.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
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

   SQL = "select * from DESCR "
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

   cmbMarcaAUX.Clear
   cmbMarca.Clear

   SQL = "select * from DESCR "
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

Public Sub MOSTRA_RODAPE(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   If Trim(Msg1) <> "" Then
      barEQP.Panels.Clear
      barEQP.Panels.Add (1)
      barEQP.Panels(1).Text = Trim(Msg1)
      barEQP.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         barEQP.Panels.Add (2)
         barEQP.Panels(2).Text = Trim(Msg2)
         barEQP.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            barEQP.Panels.Add (3)
            barEQP.Panels(3).Text = Trim(Msg3)
            barEQP.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               barEQP.Panels.Add (4)
               barEQP.Panels(4).Text = Trim(Msg4)
               barEQP.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  barEQP.Panels.Add (5)
                  barEQP.Panels(5).Text = Trim(Msg5)
                  barEQP.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If
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

Private Sub MOSTRA_HISTORICO()
'On Error GoTo ERRO_TRATA

   lstOS.ListItems.Clear
   NUMR_ID_N = 0

   If Trim(txtEqp.Text) <> "" Then
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from vwOSEQUIPAMENTO WITH (NOLOCK)"
      SQL = SQL & " where equipamento_id = " & Trim(txtEqp.Text)
      SQL = SQL & " order by dt_OS desc"

      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabAUX.EOF
         If NUMR_ID_N <> TabAUX.Fields("os_id").Value Then
            NUMR_ID_N = 0 & TabAUX.Fields("os_id").Value
            Set item = lstOS.ListItems.Add(, "seq." & NUMR_ID_N, TabAUX.Fields("os_id").Value)
            item.SubItems(1) = TabAUX.Fields("dt_os").Value
         End If
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
               MsgBox "O.S. n�o informada !!!"
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
