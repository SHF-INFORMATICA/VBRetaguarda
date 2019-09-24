VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOSEQPCADASTRO 
   Caption         =   "Cadastro de Equipamento"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "OSEQPCADASTRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   14
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
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
            Object.ToolTipText     =   "Limpar formulário"
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
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
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
      Height          =   4335
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7646
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
      TabCaption(1)   =   "C&onsulta"
      TabPicture(1)   =   "OSEQPCADASTRO.frx":76DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "txtCli2"
      Tab(1).Control(4)=   "lstEqp"
      Tab(1).Control(5)=   "txtNome2"
      Tab(1).Control(6)=   "txtDesc2"
      Tab(1).Control(7)=   "txtEqp2"
      Tab(1).ControlCount=   8
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
         Left            =   6120
         TabIndex        =   37
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
         Left            =   6120
         TabIndex        =   36
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
         Left            =   2280
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEqp2 
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
         Left            =   -73440
         MaxLength       =   8
         TabIndex        =   33
         Top             =   480
         Width           =   1095
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
         Left            =   2280
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
         Left            =   6120
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
         Left            =   2280
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
         Left            =   2280
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
         Left            =   4080
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
         Left            =   2280
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
         Top             =   3840
         Width           =   4455
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
         Left            =   2280
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
         Left            =   6120
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtDesc2 
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
         Left            =   -70440
         MaxLength       =   50
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtNome2 
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
         Left            =   -71880
         MaxLength       =   100
         TabIndex        =   13
         Top             =   960
         Width           =   5415
      End
      Begin VB.CommandButton cmdConsultaEqp 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3720
         Picture         =   "OSEQPCADASTRO.frx":76F6
         Style           =   1  'Graphical
         TabIndex        =   18
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
         TabIndex        =   17
         ToolTipText     =   "Consultar Cliente"
         Top             =   3840
         Width           =   405
      End
      Begin VB.CommandButton cmdCadCliente 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3480
         Picture         =   "OSEQPCADASTRO.frx":8AFA
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cadastrar Cliente"
         Top             =   3840
         Width           =   405
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   405
         Left            =   7080
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
         Top             =   3840
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
      Begin MSComctlLib.ListView lstEqp 
         Height          =   2745
         Left            =   -74955
         TabIndex        =   19
         Top             =   1440
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   4842
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
         BackColor       =   16777152
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EqpID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Identificação"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "ANO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "MODELO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "TIPO"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCli2 
         Height          =   405
         Left            =   -74040
         TabIndex        =   12
         Top             =   960
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   1320
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
         Left            =   5400
         TabIndex        =   31
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição/Modelo:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   840
         TabIndex        =   29
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
         Left            =   840
         TabIndex        =   28
         Top             =   1920
         Width           =   1380
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
         Left            =   5280
         TabIndex        =   27
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
         Left            =   3120
         TabIndex        =   26
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
         Left            =   1680
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência:"
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
         TabIndex        =   23
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
         Left            =   5370
         TabIndex        =   22
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição/Modelo:"
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
         Left            =   -72480
         TabIndex        =   21
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label14 
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar barEQP 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   4980
      Width           =   8670
      _ExtentX        =   15293
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
      DesignWidth     =   8670
      DesignHeight    =   5355
   End
End
Attribute VB_Name = "frmOSEQPCADASTRO"
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
   Dim EQUIPAMENTO_ID_N As Long
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

Private Sub lstEqp_DblClick()
   If Trim(lstEqp.SelectedItem.Text) <> "" Then
      txtEqp.Text = Trim(lstEqp.SelectedItem.Text)
      SSTab1.Tab = 0
      txtEqp.SetFocus
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then _
      txtEqp.SetFocus

   If SSTab1.Tab = 1 Then
      txtEqp2.SetFocus

      SETA_GRID_EQP
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
         frmOSEqpCONSULTA.Show 1
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

            SQL = "select * from EQUIPAMENTO"
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
                  
                  MsgBox "Impossível excluir, Equipamento possue movimentação na oficina."
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

Private Sub txtCli2_Change()
   SETA_GRID_EQP
End Sub

Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F4 - Cadastra Cliente", "F7 - Consulta Clientes", "Informe proprietário do Equipamento", ""

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

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCLIENTE.Show 1
         txtCNPJCPF.SetFocus
      Case vbKeyF7
         frmDISPLAYCLIENTE.Show 1
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

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
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
            CRITERIO = txtCNPJCPF.Text
      End If
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) <> "" Then
         CRITERIO = txtCNPJCPF.Text
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
         End If
         txtCNPJCPF.Text = CRITERIO
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

Private Sub TXTCNPJCPF_LostFocus()
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
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
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
   frmDISPLAYCLIENTE.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Mask = "##############"
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = "" & Trim(CNPJCPF_A)
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmdCadCliente_Click()
   frmCADASTROCLIENTE.Show 1
   txtCNPJCPF.SetFocus
End Sub

Private Sub txtDesc2_Change()
   SETA_GRID_EQP
End Sub

Private Sub txtEqp2_Change()
   SETA_GRID_EQP
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

Private Sub txtNome2_Change()
   SETA_GRID_EQP
End Sub

Private Sub txtReferencia_GotFocus()
   MOSTRA_RODAPE "Informe a Referência do Equipamento", "", "", "", ""
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtReferencia_KeyPress"
End Sub

Private Sub txtEqp_GotFocus()
   MOSTRA_RODAPE "Informe a Identificação do Equipamento", "", "", "", ""
End Sub

Private Sub txtEqp_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSEqpCONSULTA.Show 1

         If Trim(SQL3) <> "" Then
            If Not IsNumeric(SQL3) Then _
               Exit Sub

            If TabAUX.State = 1 Then _
               TabAUX.Close

            SQL = "select * from EQUIPAMENTO "
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
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEqp_KeyPress"
End Sub

Private Sub txtEqp_LostFocus()
   If Trim(txtEqp.Text) = "" Then
      EQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "EQUIPAMENTO", "", "", "", "")
      txtEqp.Text = EQUIPAMENTO_ID_N
   End If

   If Not IsNumeric(txtEqp.Text) Then
      EQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "EQUIPAMENTO", "", "", "", "")
      txtEqp.Text = EQUIPAMENTO_ID_N
   End If

   MOSTRA_EQP
End Sub

Private Sub cmdConsultaEQP_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmOSEqpCONSULTA.Show 1
   If SQL3 <> "" Then _
      txtEqp.Text = SQL3
   SQL3 = ""
   txtEqp.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsultaEQP_Click"
End Sub

Private Sub txtDescricao_GotFocus()
   MOSTRA_RODAPE "Informe a descrição do Equipamento", "", "", "", ""
End Sub

Private Sub TXTDESCRICAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescricao_KeyPress"
End Sub

Private Sub txtReferencia_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtReferencia.Text = "" Then
      txtReferencia.Text = txtEqp.Text
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtReferencia_LostFocus"
End Sub

Private Sub txtAno_GotFocus()
   MOSTRA_RODAPE "Informe ano de fabricação do Equipamento", "", "", "", ""
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMODELO_KeyPress"
End Sub

Private Sub cmbcor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_KeyPress"
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
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
      SendKeys ("{tab}")
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

      SQL = "select * from vwRel_EQUIPAMENTO "
      SQL = SQL & " where EQUIPAMENTO_ID = " & Trim(txtEqp.Text)
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then

         PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value
         EQUIPAMENTO_ID_N = TabTemp.Fields("EQUIPAMENTO_id").Value
         txtEqp.Text = EQUIPAMENTO_ID_N

         txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True

         txtNome.Text = "" & Trim(TabTemp.Fields("Nome_Cliente").Value)

         txtReferencia.Text = "" & Trim(TabTemp.Fields("identificacao").Value)
         txtDescricao.Text = "" & TabTemp!DESCRICAO
         txtANO.Text = "" & TabTemp!Ano
         txtMODELO.Text = "" & TabTemp!modelo

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
   EQUIPAMENTO_ID_N = 0
   COMBUSTIVEL_ID_N = 0
   PESSOA_ID_N = 0
   txtEqp.Text = ""
   txtDescricao.Text = ""
   txtReferencia.Text = ""
   cmbCor.Text = ""
   cmbCorAUX.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   cmbTIPO.Text = ""
   cmbTipoAUX.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   SETA_GRID_EQP
   txtEqp.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_EQP"
End Sub

Private Sub GRAVA_EQP()
'On Error GoTo ERRO_TRATA

   If Trim(txtEqp.Text) = "" Then
      MsgBox "Número de Identificação deve ser informado."
      txtEqp.SetFocus
      Exit Sub
   End If
   If Trim(txtReferencia.Text) = "" Then
      MsgBox "Número de Chassi deve ser informado."
      txtReferencia.SetFocus
      Exit Sub
   End If
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Cliente deve ser informado."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Descrição do Equipamento deve ser informada."
      txtDescricao.SetFocus
      Exit Sub
   End If

   COR_ID_N = 0 & cmbCorAUX.Text
   MARCA_ID_N = 0 & cmbMarcaAUX.Text
   TIPO_EQP_ID_N = 0 & cmbTipoAUX.Text
   ANO_ID_N = 0 & txtANO.Text
   MODELO_ID_N = 0 & txtMODELO.Text

   EQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "EQUIPAMENTO", "", "", "", "")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from EQUIPAMENTO "
   SQL = SQL & " where equipamento_id = " & Trim(txtEqp.Text)
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into EQUIPAMENTO "
      SQL = SQL & "(EQUIPAMENTO_ID,DT_CAD,DESCRICAO,IDENTIFICACAO,PESSOA_ID,"
      SQL = SQL & " COR_ID,MARCA_ID,TIPO_EQP,ANO,modelo,nome_cliente)"
      SQL = SQL & " values("
         SQL = SQL & EQUIPAMENTO_ID_N                       'EQUIPAMENTO_ID
         SQL = SQL & "," & DMA(Date)                        'DT_CAD
         SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'DESCRICAO
         SQL = SQL & ",'" & Trim(txtReferencia.Text) & "'"  'IDENTIFICACAO
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
         EQUIPAMENTO_ID_N = TabTemp.Fields("equipamento_id").Value

         SQL = "update EQUIPAMENTO "
         SQL = SQL & "set "
            SQL = SQL & " descricao = '" & Trim(txtDescricao.Text) & "'"         'DESCRICAO
            SQL = SQL & ", IDENTIFICACAO = '" & Trim(txtReferencia.Text) & "'"   'IDENTIFICACAO
            SQL = SQL & ", pessoa_ID = " & PESSOA_ID_N                           'pessoa_ID
            SQL = SQL & ", COR_ID = " & COR_ID_N                                 'COR_ID
            SQL = SQL & ", MARCA_ID = " & MARCA_ID_N                             'MARCA_ID
            SQL = SQL & ", TIPO_EQP = " & TIPO_EQP_ID_N                          'TIPO_EQP
            SQL = SQL & ", ANO = " & ANO_ID_N                                    'ANO
            SQL = SQL & ", nome_cliente = '" & Trim(txtNome.Text) & "'"          'NOME CLIENTE
            SQL = SQL & ", modelo = " & MODELO_ID_N                              'modelo

         SQL = SQL & " where equipamento_id = " & EQUIPAMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   LIMPA_EQP

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_EQP"
End Sub

Private Sub SETA_GRID_EQP()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 1
   lstEqp.ListItems.Clear

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "SELECT * from vwRel_EQUIPAMENTO "
   SQL = SQL & " where EQUIPAMENTO_ID > 0"

   If Trim(txtEqp2.Text) <> "" Then _
      SQL = SQL & " and equipamento_id = " & txtEqp2.Text

   If Trim(txtDesc2.Text) <> "" Then _
      SQL = SQL & " and descricao = like '" & Trim(txtDesc2.Text) & "%'"

   If Trim(txtNome2.Text) <> "" Then _
      SQL = SQL & " and nome_cliente like '" & Trim(txtNome2.Text) & "%'"

   txtCli2.PromptInclude = False
   If Trim(txtCli2.Text) <> "" Then _
      SQL = SQL & " and cnpjcpf = like '" & Trim(txtCli2.Text) & "%'"
   txtCli2.PromptInclude = True

   SQL = SQL & " order by ano asc "

   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabAUX.EOF
      Set Item = lstEqp.ListItems.Add(, "seq." & TabAUX.Fields("equipamento_id").Value, TabAUX.Fields("equipamento_id").Value)
      Item.SubItems(1) = "" & TabAUX.Fields("identificacao").Value
      Item.SubItems(2) = "" & TabAUX.Fields("nome_cliente").Value

      If Not IsNull(TabAUX!Ano) Then _
         Item.SubItems(3) = "" & TabAUX!Ano
      If Not IsNull(TabAUX!modelo) Then _
         Item.SubItems(4) = "" & TabAUX!modelo
      If Not IsNull(TabAUX!TIPO_EQP) Then _
         Item.SubItems(5) = "" & TRAZ_DESCRITOR("A", TabAUX!TIPO_EQP)

      TabAUX.MoveNext
   Wend
   If TabAUX.State = 1 Then _
      TabAUX.Close

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_EQP"
End Sub

Sub CARREGA_DESCRITORES()
'On Error GoTo ERRO_TRATA

'Tipo Função
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
      cmbTipoAUX.AddItem TabDESCR!codigo
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
      cmbCorAUX.AddItem TabDESCR!codigo
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
      cmbMarcaAUX.AddItem TabDESCR!codigo
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
