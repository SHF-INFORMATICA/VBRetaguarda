VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAcertoCliente 
   Caption         =   "Acerto Pedido Pendente Cliente"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AcertoCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   11820
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Acerto Pedido Pendente Cliente"
      TabPicture(0)   =   "AcertoCliente.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label7(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDtIni"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDtFim"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCNPJCPF"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lstPedido"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lstProduto"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtValorRevenda"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtValorCompra"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtQtdeProduto"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtValorProducao"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtReg"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtTotalVenda"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdConsCli"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtNome"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "optSintetico"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "optAnalitico"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmbForma"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmbAuxForma"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtTotVendas"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtTotDesconto"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "chkAbertos"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtValrAcerto"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "chkTodos"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "chkFunc"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      Begin VB.CheckBox chkFunc 
         Caption         =   "Somente Funcionários"
         Height          =   240
         Left            =   4080
         TabIndex        =   39
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtValrAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7695
         TabIndex        =   36
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkAbertos 
         Caption         =   "Pendentes"
         Height          =   240
         Left            =   4080
         TabIndex        =   35
         Top             =   600
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtTotDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   5685
         TabIndex        =   32
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox txtTotVendas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   8220
         TabIndex        =   31
         Top             =   6360
         Width           =   1095
      End
      Begin VB.ComboBox cmbAuxForma 
         BackColor       =   &H80000000&
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
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbForma 
         Height          =   360
         Left            =   1200
         TabIndex        =   0
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton optAnalitico 
         Caption         =   "&Analítico"
         Height          =   240
         Left            =   7440
         TabIndex        =   28
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optSintetico 
         Caption         =   "&Sintético"
         Height          =   240
         Left            =   6120
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3600
         MaxLength       =   100
         TabIndex        =   25
         Top             =   1080
         Width           =   5175
      End
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3120
         Picture         =   "AcertoCliente.frx":5C2E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   405
      End
      Begin VB.TextBox txtTotalVenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10530
         TabIndex        =   11
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox txtReg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3840
         TabIndex        =   10
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox txtValorProducao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtQtdeProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1560
         TabIndex        =   8
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox txtValorCompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtValorRevenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstProduto 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   4440
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   2990
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   14737632
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qtde"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Valor Item"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Peso"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total Item"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "cliente_id"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Dt.Pedido"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "NCM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TIPO_PROD"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstPedido 
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3836
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   128
         BackColor       =   16777152
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
            Text            =   "Pedido"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DtPedido"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Desconto"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Faturamento"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ValorAcerto ="
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6285
         TabIndex        =   37
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr.Vendas="
         Height          =   240
         Index           =   0
         Left            =   7005
         TabIndex        =   34
         Top             =   6360
         Width           =   1155
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Descontos="
         Height          =   240
         Left            =   4560
         TabIndex        =   33
         Top             =   6360
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fatura:"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Final:"
         Height          =   240
         Left            =   3120
         TabIndex        =   22
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Inicial:"
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr.Total = "
         Height          =   240
         Index           =   1
         Left            =   9465
         TabIndex        =   20
         Top             =   6360
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde.Pedidos="
         Height          =   240
         Left            =   2400
         TabIndex        =   19
         Top             =   6360
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   1
         X1              =   1080
         X2              =   12840
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   2
         X1              =   0
         X2              =   11760
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   11760
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Produção ="
         Height          =   240
         Left            =   8790
         TabIndex        =   18
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde.Produto="
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   6360
         Width           =   1380
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Compras ="
         Height          =   240
         Left            =   8865
         TabIndex        =   15
         Top             =   1560
         Width           =   1590
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Revenda ="
         Height          =   240
         Left            =   8835
         TabIndex        =   14
         Top             =   1080
         Width           =   1590
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1270
      ButtonWidth     =   3175
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
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "BaixarPedidos"
            Key             =   "baixar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "NFe"
            Key             =   "nfe"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   9840
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
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
               Picture         =   "AcertoCliente.frx":6630
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoCliente.frx":77CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoCliente.frx":8859
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoCliente.frx":980E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoCliente.frx":A919
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoCliente.frx":C8FB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   7560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11820
      DesignHeight    =   7785
   End
End
Attribute VB_Name = "frmAcertoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_PRODUCAO_N       As Double
   Dim VALOR_PERMITIDO_N      As Double
   Dim VALOR_ULTRAPASSADO_N   As Double
   Dim Conta_Produto_N        As Long
   Dim VALOR_COMPRA_N         As Double
   Dim VALOR_ACERTO_N         As Double
   Dim VALOR_REVENDA_N        As Double
   Dim INDR_ACHOU_REGISTRO    As Boolean
   Dim ST_PRODUTO_A           As String
   Dim strCFOP_ITEM           As String
   Dim PERCICMS_N             As Integer
   Dim PEDIDO_GRID_N          As Long

Private Sub Form_Load()

   If USA_NFe = True Then _
      Toolbar1.Buttons(6).Visible = True

   LIMPA_TUDO

   Call TXTDTINI_GotFocus
   Call TXTDTFIM_GotFocus
   cmbForma.Clear
   cmbAuxForma.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK) "
   SQL = SQL & " where receber = 'true' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbForma.AddItem TabTemp!DESCRICAO & " - " & TabTemp!TIPOVENDA_ID
      cmbAuxForma.AddItem TabTemp!TIPOVENDA_ID
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
   
End Sub

Private Sub lstPedido_Click()
   MOSTRA_VALOR_ACERTO
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "baixar"
         BAIXA_PEDIDO
      Case "limpar"
         LIMPA_TUDO
      Case "consultar"
         MONTA_CONSULTA_SQL
      Case "limpar"
         Call Form_Load
      Case "voltar"
         Unload Me
      Case "imprimir"
         MONTA_REL
      Case "nfe"
         MONTA_PEDIDO_NFE
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub chkAbertos_Click()
   If chkAbertos.Value = 1 Then
      chkAbertos.Caption = "Pendentes"
      Else: chkAbertos.Caption = "Liquidados"
   End If
   chkAbertos.Refresh
   DoEvents
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxForma.ListIndex = cmbForma.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbFORMA_Click"
End Sub

Private Sub lstpedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedido, ColumnHeader
End Sub

Private Sub lstProduto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstProduto, ColumnHeader
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      CLIENTE_ID_N = 0

      MOSTRA_CLIENTE
   End If
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

   MOSTRA_RODAPE "ESC - Sair", "F6 - Excluir Cliente", "F7 - Consultar Cliente", "Informe CPF do cliente", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
      Case vbKeyBack
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            txtCNPJCPF.Mask = "##############"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_CLIENTE
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtValorPermitido_Change()
   CALCULA_ULTRAPASSADO
End Sub

Private Sub chkTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstPedido.ListItems.Count > 0 Then
      For i = lstPedido.ListItems.Count To 1 Step -1
         If chkTodos.Value = 1 Then
            lstPedido.ListItems(i).Checked = True
            Else: lstPedido.ListItems(i).Checked = False
         End If
      Next i
   End If
   MOSTRA_VALOR_ACERTO

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "chkTodos_Click"
End Sub

Private Sub txtValrAcerto_GotFocus()
   txtCNPJCPF.SetFocus
End Sub
'=========================================
Sub CALCULA_ULTRAPASSADO()
'On Error GoTo ERRO_TRATA

   VALOR_ULTRAPASSADO_N = 0


   VALOR_ULTRAPASSADO_N = "" & VALOR_PRODUCAO_N - VALOR_PERMITIDO_N

   txtValorRevenda.Text = "" & Format(VALOR_REVENDA_N, strFormatacao2Digitos)

   DoEvents

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CALCULA_ULTRAPASSADO"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   chkFunc.Value = 0
   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0
   INDR_ACHOU_REGISTRO = False
   txtValrAcerto.Text = ""
   cmbAuxForma.Text = ""
   cmbForma.Text = ""
   CLIENTE_ID_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   CONT_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_PRODUCAO_N = 0
   VALOR_PERMITIDO_N = 0
   VALOR_ULTRAPASSADO_N = 0
   VALOR_REVENDA_N = 0

   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtDtIni.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtValorProducao.Text = ""
   txtQtdeProduto.Text = ""
   txtReg.Text = ""
   txtTotDesconto.Text = ""
   txtTotVendas.Text = ""
   txtTotalVenda.Text = ""
   lstPedido.ListItems.Clear
   lstProduto.ListItems.Clear
   txtValorRevenda.Text = ""
   txtValorCompra.Text = ""

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub CHECA_ULTIMO_DIA_MES()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
      txtDtFim.Text = ""

      txtDtIni.PromptInclude = True
      If IsDate(txtDtIni.Text) Then
         CRITERIO_A = FimDoMes(txtDtIni.Text, False)
         CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
         txtDtFim.Text = CRITERIO_A
         txtDtFim.PromptInclude = True
      End If
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CHECA_ULTIMO_DIA_MES"
End Sub

Sub MONTA_CONSULTA_SQL()
'On Error GoTo ERRO_TRATA

   HORA_INI = Time
   INDR_ACHOU_REGISTRO = False
   chkTodos.Value = 1
   txtValorProducao.Text = ""
   txtValorRevenda.Text = ""
   txtValorCompra.Text = ""
   txtQtdeProduto.Text = ""
   txtReg.Text = ""
   txtTotDesconto.Text = ""
   txtTotVendas.Text = ""
   txtTotalVenda.Text = ""
   txtValrAcerto.Text = ""

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", Format((HORA_INI), "hh:mm:ss")

   CHECA_ULTIMO_DIA_MES

   VALOR_TOTAL_N = 0

   txtTotalVenda.Text = ""
   txtReg.Text = ""

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   'SQL = "select * from vwACERTO_PEDIDO_CLIENTE WITH (NOLOCK)"

   SQL = "select * from vwACERTO_PEDIDO_CLIENTE WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   SQL = SQL & " ON vwACERTO_PEDIDO_CLIENTE.PEDIDO_ID = ITEMLANCAMENTO.NUMR_DOC"

   SQL = SQL & " where pedido_id Is Not Null"

   If chkAbertos.Value = 1 Then
      SQL = SQL & " and ITEMLANCAMENTO.status = 'A' "
      Else
         If chkAbertos.Value = 0 Then _
            SQL = SQL & " and ITEMLANCAMENTO.status = 'B' "
   End If

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and vwACERTO_PEDIDO_CLIENTE.status <> 9 "

   If CLIENTE_ID_N > 0 Then _
      SQL = SQL & " and cliente_id = " & CLIENTE_ID_N

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & Trim(txtDtIni.Text) & "'"
      SQL = SQL & " and dt_req <= '" & Trim(txtDtFim.Text) & "'"
   End If

   If Trim(cmbAuxForma.Text) <> "" Then _
      SQL = SQL & " and tipovenda_id = " & cmbAuxForma.Text

   SQL = SQL & " order by dt_req desc"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   SETA_GRID
   MOSTRA_VALOR_ACERTO

   lstProduto.Visible = True
   
   CALCULA_ULTRAPASSADO

   txtValorCompra.Text = "" & Format(VALOR_COMPRA_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim INDR_VAI   As Boolean
   Dim TabFunc    As New ADODB.Recordset

   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   CONT_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_PRODUCAO_N = 0
   VALOR_PERMITIDO_N = 0
   VALOR_ULTRAPASSADO_N = 0
   VALOR_COMPRA_N = 0
   VALOR_REVENDA_N = 0
   INDR_ACHOU_REGISTRO = False
   NUMR_SEQ_N = 0

   lstPedido.Visible = False
   lstPedido.ListItems.Clear
   lstProduto.ListItems.Clear
   lstProduto.Visible = False

   If Not TabTemp.EOF Then
      While Not TabTemp.EOF
         INDR_VAI = True
         If chkFunc.Value = 1 Then
            INDR_VAI = False
            If TabFunc.State = 1 Then _
               TabFunc.Close

            SQL = "select funcionario from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where pessoa_id = " & TabTemp.Fields("pessoa_id").Value
            TabFunc.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFunc.EOF Then _
               If Not IsNull(TabFunc.Fields(0).Value) Then _
                  If TabFunc.Fields(0).Value = True Then _
                     INDR_VAI = True
            If TabFunc.State = 1 Then _
               TabFunc.Close
         End If

         DoEvents
         PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value
         txtReg.Text = PEDIDO_ID_N

         If INDR_VAI = True Then
            INDR_VAI = False
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
   
            SQL = "select numr_doc from ITEMLANCAMENTO WITH (NOLOCK)"
            SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
            If chkAbertos.Value = 1 Then
               SQL = SQL & " and status = 'A' "
               Else
                  If chkAbertos.Value = 0 Then _
                     SQL = SQL & " and status = 'B' "
            End If
   
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               If Not IsNull(TabConsulta.Fields(0).Value) Then _
                  INDR_VAI = True
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
   
            If INDR_VAI = True Then
               If CONT_N < CONTA_REG_PROGRESSO Then
                  CONT_N = CONT_N + 1
                  If CONT_N <= 100 Then _
                     ProgressBar1.Value = CONT_N
               End If
   
               CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
               txtReg.Text = CONTA_REGISTRO_N
               txtReg.Refresh
   
               VALOR_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
               VALOR_ITEM_N = 0
   
               If TabPedidoItem.State = 1 Then _
                  TabPedidoItem.Close
   
               SQL = "select SUM(valor_item) from ITEMLANCAMENTO WITH (NOLOCK)"
               SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
               SQL = SQL & " and status = 'A' "
               TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabPedidoItem.EOF Then _
                  If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
                     VALOR_ITEM_N = 0 & TabPedidoItem.Fields(0).Value
               If TabPedidoItem.State = 1 Then _
                  TabPedidoItem.Close
NUMR_SEQ_N = NUMR_SEQ_N + 1

               Set item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, PEDIDO_ID_N)
   
               item.SubItems(1) = "" & TabTemp.Fields("dt_req").Value
               item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
               item.SubItems(3) = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
               item.SubItems(4) = "" & Format(VALOR_ITEM_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
               item.SubItems(5) = "" & Trim(TabTemp.Fields("descricao").Value)
   
               VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N
               VALOR_TOTAL_DESCONTO_N = VALOR_TOTAL_DESCONTO_N + VALOR_DESCONTO_N
   
               txtTotDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
               txtTotVendas.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
               txtTotalVenda.Text = Format(VALOR_TOTAL_N + VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
   
               item.Checked = True
               INDR_ACHOU_REGISTRO = True
   
               SETA_GRID_ITENS PEDIDO_ID_N
            End If
         End If
         TabTemp.MoveNext
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub SETA_GRID_ITENS(NUMR_PEDIDO_ID As Long)
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = " select pedido.CLIENTE_ID,pedido.dt_req,PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, "
   SQL = SQL & " PEDIDOITEM.PRODUTO_ID, PRODUTO.CODG_PRODUTO, PEDIDOITEM.QTD_PEDIDA, "
   SQL = SQL & " PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.STATUS, PEDIDOITEM.PRECO_CUSTO, "
   SQL = SQL & " PEDIDOITEM.PESO_ITEM, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " PRODUTO.SITUACAO, PRODUTO.TIPO_PROD, PRODUTO.PRECO_CUSTO AS precocusto, "
   SQL = SQL & " PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, PRODUTO.PRODUTO_BALANCA,codg_ncm,"
   SQL = SQL & " FAMILIAPRODUTO.FAMILIAPRODUTO_ID AS DescFamilia, "
   SQL = SQL & " FAMILIAPRODUTO.DESCRICAO AS DescFamilia, FAMILIAPRODUTO.PRODUCAO"
   SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"
   SQL = SQL & " INNER JOIN PEDIDO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = PEDIDO.PEDIDO_ID AND PEDIDOITEM.PEDIDO_ID = PEDIDO.PEDIDO_ID"

   SQL = SQL & " where PEDIDOITEM.pedido_id = " & NUMR_PEDIDO_ID
   SQL = SQL & " and pedidoitem.status <> 'C' "

   SQL = SQL & " order by PEDIDOITEM.pedido_id "

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Not TabProduto.EOF Then
      While Not TabProduto.EOF
         DoEvents

         CONT_N = CONT_N + 1

         Set ITEM2 = lstProduto.ListItems.Add(, "seq." & CONT_N, TabProduto.Fields("PEDIDO_ID").Value)

         ITEM2.SubItems(1) = "" & Trim(TabProduto.Fields("codg_produto").Value)
         ITEM2.SubItems(2) = "" & Trim(TabProduto.Fields("descricao").Value)
         ITEM2.SubItems(3) = "" & Format(TabProduto.Fields("QTD_PEDIDA").Value, strFormatacao3Digitos)
         ITEM2.SubItems(4) = "" & Format(TabProduto.Fields("valor_Item").Value, strFormatacao2Digitos)
         ITEM2.SubItems(5) = "" & Format(TabProduto.Fields("peso_Item").Value, strFormatacao3Digitos)
         ITEM2.SubItems(6) = "" & Format(TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value, strFormatacao2Digitos)
         ITEM2.SubItems(7) = "" & TabProduto.Fields("cliente_id").Value
         ITEM2.SubItems(8) = "" & TabProduto.Fields("dt_req").Value
         ITEM2.SubItems(9) = "" & TabProduto.Fields("codg_ncm").Value
         ITEM2.SubItems(10) = "" & TabProduto.Fields("PRODUCAO").Value

         VALOR_COMPRA_N = VALOR_COMPRA_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value

         'é item de produção
         If Not IsNull(TabProduto.Fields("PRODUCAO").Value) Then
            If TabProduto.Fields("PRODUCAO").Value = True Then
               ITEM2.SubItems(10) = "Produção"

               VALOR_PRODUCAO_N = VALOR_PRODUCAO_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value

               txtValorProducao.Text = Format(VALOR_PRODUCAO_N, strFormatacao2Digitos)
               txtValorProducao.Refresh

               ITEM2.ForeColor = vbRed
               ITEM2.ListSubItems(1).ForeColor = vbRed
               ITEM2.ListSubItems(2).ForeColor = vbRed
               ITEM2.ListSubItems(3).ForeColor = vbRed
               ITEM2.ListSubItems(4).ForeColor = vbRed
               ITEM2.ListSubItems(5).ForeColor = vbRed
               ITEM2.ListSubItems(6).ForeColor = vbRed

               Else
                  VALOR_REVENDA_N = VALOR_REVENDA_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value
                  ITEM2.SubItems(10) = "Revenda"
            End If
            Else: VALOR_REVENDA_N = VALOR_REVENDA_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value
         End If

         ITEM2.Checked = True

         Conta_Produto_N = Conta_Produto_N + 1
         txtQtdeProduto.Text = Conta_Produto_N

         TabProduto.MoveNext
      Wend
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ITENS"
End Sub

Sub BAIXA_PEDIDO()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma baixa das compras do cliente " & Trim(txtNome.Text) & " ?"
   PERGUNTA Msg, vbYesNo + 32, "Baixa de compras de Cliente", "DEMO.HLP", 1000
   If RESPOSTA = vbNo Then _
      Exit Sub

   Dim i                   As Integer

   INDR_PRI = False

   If lstPedido.ListItems.Count > 0 Then
      For i = lstPedido.ListItems.Count To 1 Step -1
         If lstPedido.ListItems(i).Checked = True Then

            If Trim(lstPedido.ListItems(i).Text) <> "" Then

               INDR_PRI = True
               PEDIDO_ID_N = 0 & Trim(lstPedido.ListItems(i).Text)

               ATUALIZA_ESTOQUE 0, PEDIDO_ID_N

               SQL = "update PEDIDO set "
               SQL = SQL & " status = 5 "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               CONECTA_RETAGUARDA.Execute SQL

               'baixando financeiro
               '=====
               SQL = "update itemLANCAMENTO set "
               SQL = SQL & " status = 'B' , dt_baixa = '" & Now & "'"
               SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
               PEDIDO_ID_N = 0
               '=====
            End If
         End If
      Next i
   End If

   If INDR_PRI = True Then
      MsgBox "Processo realizado com sucesso."
      'LIMPA_TUDO
      MONTA_CONSULTA_SQL
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_PEDIDO"
End Sub

Sub CALCULA_GRID_ITEM()
'On Error GoTo ERRO_TRATA

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "RETIRA_ITEM"
End Sub

Sub MOSTRA_CLIENTE()
'On Error GoTo ERRO_TRATA

   If Trim(txtCNPJCPF.Text) <> "" Then
      PESSOA_ID_N = 0
      CLIENTE_ID_N = 0
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select nome,cliente_id,pessoa_id from CLIENTE WITH (NOLOCK) "
      SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
         txtNome.Text = Trim(TabCliente.Fields("nome").Value)
         PESSOA_ID_N = 0 & TabCliente.Fields("pessoa_id").Value

         If Trim(txtCNPJCPF.Text) <> "" Then
            CRITERIO_A = Trim(txtCNPJCPF.Text)
            If Not IsNull(txtCNPJCPF.Text) Then
               If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
                  txtCNPJCPF.Mask = "###.###.###-##"
                  Else: txtCNPJCPF.Mask = "##.###.###/####-##"
               End If
            End If
            txtCNPJCPF.Text = CRITERIO_A
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CLIENTE"
End Sub

Sub MOSTRA_VALOR_ACERTO()
'On Error GoTo ERRO_TRATA

   Dim i             As Integer
   Dim Valr_Acerto_N As Double

   Valr_Acerto_N = 0
   txtValrAcerto.Text = tpMOEDA(Valr_Acerto_N)
   txtValrAcerto.Refresh

   INDR_PRI = False
   Valr_Acerto_N = 0

   If lstPedido.ListItems.Count > 0 Then
      For i = lstPedido.ListItems.Count To 1 Step -1
         If lstPedido.ListItems(i).Checked = True Then
            If Trim(lstPedido.ListItems(i).Text) <> "" Then

               INDR_PRI = True
               VALOR_ITEM_N = 0 & Trim(lstPedido.ListItems(i).SubItems(4))
               Valr_Acerto_N = VALOR_ITEM_N + Valr_Acerto_N
               txtValrAcerto.Text = tpMOEDA(Valr_Acerto_N)
               txtValrAcerto.Refresh
            End If
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VALOR_ACERTO"
End Sub

Sub MONTA_REL()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "REL_ACERTO_CLIENTE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "drop table REL_ACERTO_CLIENTE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "REL_ACERTO_CLIENTE", "") = False Then
      SQL = "CREATE TABLE [dbo].[REL_ACERTO_CLIENTE]("
      SQL = SQL & "[ESTABELECIMENTO_ID]   [int] NOT NULL,"
      SQL = SQL & "[CLIENTE_ID]           [bigint] NOT NULL,"
      SQL = SQL & "[PEDIDO_ID]            [bigint] NOT NULL,"
      SQL = SQL & "[CODG_PRODUTO]         [nvarchar](60) NOT NULL,"
      SQL = SQL & "[DESC_PRODUTO]         [nvarchar](100) NOT NULL,"
      SQL = SQL & "[QTDE_VENDIDA]         [float] NOT NULL,"
      SQL = SQL & "[VALOR_ITEM]           [float] NOT NULL,"
      SQL = SQL & "[VALOR_ITEM_PROD]      [float] NOT NULL,"
      SQL = SQL & "[VALOR_ITEM_REV]       [float] NOT NULL,"
      SQL = SQL & "[DT_PEDIDO]            [datetime] NOT NULL"
      SQL = SQL & ") ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   CONECTA_RETAGUARDA.Execute "delete REL_ACERTO_CLIENTE"

   Dim i                As Integer
   Dim Valr_Acerto_N    As Double

   Valr_Acerto_N = 0
   txtValrAcerto.Text = tpMOEDA(Valr_Acerto_N)
   txtValrAcerto.Refresh

   Valr_Acerto_N = 0

   If lstProduto.ListItems.Count > 0 Then
      For i = lstProduto.ListItems.Count To 1 Step -1
         If Trim(lstProduto.ListItems(i).Text) <> "" Then

            SQL = "insert into REL_ACERTO_CLIENTE "
               SQL = SQL & " (ESTABELECIMENTO_ID,CLIENTE_ID,PEDIDO_ID,CODG_PRODUTO,"
               SQL = SQL & " DESC_PRODUTO,QTDE_VENDIDA,VALOR_ITEM,dt_pedido,VALOR_ITEM_PROD,VALOR_ITEM_REV)"
            SQL = SQL & " values("
               SQL = SQL & ESTABELECIMENTO_ID_N                                     'ESTABELECIMENTO_ID
               SQL = SQL & "," & Trim(lstProduto.ListItems(i).SubItems(7))          'CLIENTE_ID
               SQL = SQL & "," & Trim(lstProduto.ListItems(i).Text)                 'PEDIDO_ID
               SQL = SQL & ",'" & Trim(lstProduto.ListItems(i).SubItems(1)) & "'"   'CODG_PRODUTO
               SQL = SQL & ",'" & Trim(lstProduto.ListItems(i).SubItems(2)) & "'"   'DESC_PRODUTO
               SQL = SQL & "," & tpMOEDA(lstProduto.ListItems(i).SubItems(3))       'QTDE_VENDIDA
               SQL = SQL & "," & tpMOEDA(lstProduto.ListItems(i).SubItems(4))       'VALOR_ITEM
               SQL = SQL & ",'" & DMA(lstProduto.ListItems(i).SubItems(8)) & "'"    'dt_pedido

               If Trim(lstProduto.ListItems(i).SubItems(10)) = "Produção" Then
                  SQL = SQL & "," & tpMOEDA(lstProduto.ListItems(i).SubItems(4))         'VALOR_ITEM_PROD
                  SQL = SQL & "," & tpMOEDA(0)                                            'VALOR_ITEM_REV
                  Else
                     If Trim(lstProduto.ListItems(i).SubItems(10)) = "Revenda" Then
                        SQL = SQL & "," & tpMOEDA(0)                                      'VALOR_ITEM_PROD
                        SQL = SQL & "," & tpMOEDA(lstProduto.ListItems(i).SubItems(4))   'VALOR_ITEM_REV
                     End If
               End If

            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL

         End If
      Next i
   End If

   FORMULA_REL = ""
   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If optSintetico.Value = True Then
      Nome_Relatorio = "acerto_cliente_sint.rpt"
      Else: Nome_Relatorio = "acerto_cliente_analit.rpt"
   End If
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL"
End Sub

Sub GERA_PEDIDO_NFE()
'On Error GoTo ERRO_TRATA

   Dim tabPEDIDONFe  As New ADODB.Recordset
   Dim i
   PEDIDO_ID_N = 1

   CONT_N = 999999999

   For i = 1 To CONT_N '> PEDIDO_ID_N
      PEDIDO_ID_N = i
      If tabPEDIDONFe.State = 1 Then _
         tabPEDIDONFe.Close

      SQL = "select pedido_id from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      tabPEDIDONFe.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If tabPEDIDONFe.EOF Then
         If tabPEDIDONFe.State = 1 Then _
            tabPEDIDONFe.Close

         SQL = "select numr_doc from LANCAMENTO WITH (NOLOCK)"
         SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
         tabPEDIDONFe.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If tabPEDIDONFe.EOF Then
            If tabPEDIDONFe.State = 1 Then _
               tabPEDIDONFe.Close

            SQL = "select pedido_id from PEDIDOTEMP WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            tabPEDIDONFe.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If tabPEDIDONFe.EOF Then
               If tabPEDIDONFe.State = 1 Then _
                  tabPEDIDONFe.Close

               SQL = "select os_id from OS WITH (NOLOCK)"
               SQL = SQL & " where os_id = " & PEDIDO_ID_N
               tabPEDIDONFe.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If tabPEDIDONFe.EOF Then _
                  Exit For
            End If
         End If
      End If
   Next
   If tabPEDIDONFe.State = 1 Then _
      tabPEDIDONFe.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_PEDIDO_NFE"
End Sub

Sub MONTA_PEDIDO_NFE()
'On Error GoTo ERRO_TRATA

Dim CFOP_ID_N As Integer
Dim strCFOP_ITEM As String

   If INDR_ACHOU_REGISTRO = True Then
      txtCNPJCPF.PromptInclude = False

      If Trim(UF_CLIENTE_A) = "" Then _
         TRATA_PESSOA Trim(txtCNPJCPF.Text)

      If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
         If Trim(UF_CLIENTE_A) = "" Then
            MsgBox "Cliente com cadastro incompleto !!!"
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      End If

      Msg = "Confirma emissão de Nota Fiscal Eletrônica para o grupo de pedidos selecionados ?"
      PERGUNTA Msg, vbYesNo + 32, "Emissão de Nota Fiscal Eletrônica", "DEMO.HLP", 1000
      If RESPOSTA = vbNo Then _
         Exit Sub

      Dim i                As Integer
      Dim AGRUPA_PEDIDO_A  As String

      If Trim(UF_EMPRESA_A) = "" Then _
         PEGA_DADOS_EMPRESA

      ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
      ALIQUOTA_ICMS_NORMAL_FORA_UF = 0
      PERCICMS_N = 0
      'Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, UF_CLIENTE_A, 0)

      strCFOP_ITEM = "5102"
      ST_PRODUTO_A = ""
      If CTR_EMPRESA_N = 1 Then
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
            strCFOP_ITEM = "5102"
            PERCICMS_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF
            Else
               PERCICMS_N = ALIQUOTA_ICMS_NORMAL_FORA_UF
               strCFOP_ITEM = "6102"
         End If
      End If


   If IsNumeric(strCFOP_ITEM) Then _
      CFOP_ID_N = 0 & strCFOP_ITEM
   Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, "", CFOP_ID_N)


      INDR_PRI = False
      AGRUPA_PEDIDO_A = ""
      PEDIDO_GRID_N = 0
      PEDIDO_ID_N = 0
      i = 0

      If lstPedido.ListItems.Count > 0 Then
         For i = lstPedido.ListItems.Count To 1 Step -1
            If lstPedido.ListItems(i).Checked = True Then
               If Trim(lstPedido.ListItems(i).Text) <> "" Then
                  If INDR_PRI = False Then
                     INDR_PRI = True

                     GERA_PEDIDO_NFE

                  End If
                  PEDIDO_GRID_N = 0 & Trim(lstPedido.ListItems(i).Text)
                  AGRUPA_PEDIDO_A = PEDIDO_GRID_N & ";" & AGRUPA_PEDIDO_A
               End If
            End If
         Next i
      End If
'ITENS
      i = 0
      If lstProduto.ListItems.Count > 0 Then
         For i = lstProduto.ListItems.Count To 1 Step -1
            If lstProduto.ListItems(i).Checked = True Then
               If Trim(lstProduto.ListItems(i).Text) <> "" Then

                  PEDIDO_GRID_N = 0 & Trim(lstProduto.ListItems(i).Text)
                  If PEDIDO_ID_N > 0 And INDR_PRI = True Then
                     INDR_PRI = False
                     GRAVA_CABECA_PEDIDO_NFE
                  End If

                  SEQ_ID_N = 0 & i
                  QTDE_N = 0 & lstProduto.ListItems(i).SubItems(3)
                  'VALOR_ITEM_N = 0 & lstProduto.ListItems(i).SubItems(4)
                  CODG_PRODUTO_A = "" & lstProduto.ListItems(i).SubItems(1)

                  GRAVA_ITENS_PEDIDO_NFE
               End If
            End If
         Next i
      End If
   End If

TIPO_NFe_GERAR = "R"
If USA_DOC_FISCAL = True Then _
   If USA_NFe = True Then _
      frmNOTAGERA.Show 1

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_PEDIDO_NFE"
End Sub

Sub GRAVA_CABECA_PEDIDO_NFE()
'On Error GoTo ERRO_TRATA

   VENDEDOR_ID_N = 0
   TIPOVENDA_ID_N = 9999
   txtCNPJCPF.PromptInclude = False

   SQL = "insert into PEDIDO "
   SQL = SQL & "("
      SQL = SQL & "PEDIDO_ID,CLIENTE_ID,EMPRESA_ID,ESTABELECIMENTO_ID,VENDEDOR_ID,"
      SQL = SQL & "CGCCPF,USUARIO_ID,DT_REQ,STATUS,TIPO_REGISTRO,NOME_CLIENTE,PREFIXO"
   SQL = SQL & ")"
   SQL = SQL & " values("
      SQL = SQL & PEDIDO_ID_N
      SQL = SQL & "," & CLIENTE_ID_N
      SQL = SQL & "," & EMPRESA_ID_N
      SQL = SQL & "," & ESTABELECIMENTO_ID_N
      SQL = SQL & "," & VENDEDOR_ID_N
      SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
      SQL = SQL & "," & USUARIO_ID_N
      SQL = SQL & ",'" & Now & "'"
      SQL = SQL & "," & 3
      SQL = SQL & ",'R'"
      SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
      SQL = SQL & ",'NFE'"
   SQL = SQL & ")"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA_PEDIDO_NFE"
End Sub

Sub GRAVA_ITENS_PEDIDO_NFE()
'On Error GoTo ERRO_TRATA

   Dim tabPedidoItemNFe  As New ADODB.Recordset
   Dim tabPedidoConsItem As New ADODB.Recordset
   Dim PRECO_CUSTO_N       As Double

   If tabPedidoItemNFe.State = 1 Then _
      tabPedidoItemNFe.Close

   SQL = "select produto_id,situacao_tributaria,preco_custo from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
   tabPedidoItemNFe.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabPedidoItemNFe.EOF Then
      ST_PRODUTO_A = "" & tabPedidoItemNFe.Fields("situacao_tributaria").Value
      PRODUTO_ID_N = 0 & tabPedidoItemNFe.Fields("produto_id").Value
      VALOR_ITEM_N = 0
      PRECO_CUSTO_N = 0 & tabPedidoItemNFe.Fields("preco_custo").Value

      'tras valor do item vendido
      If tabPedidoConsItem.State = 1 Then _
         tabPedidoConsItem.Close

SQL = "select valor_item,qtd_pedida from PEDIDOITEM "
SQL = SQL & " where pedido_id = " & PEDIDO_GRID_N
SQL = SQL & " and produto_id = " & PRODUTO_ID_N
tabPedidoConsItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
If Not tabPedidoConsItem.EOF Then
   QTDE_N = 0 & tabPedidoConsItem.Fields("qtd_pedida").Value
   VALOR_ITEM_N = 0 & tabPedidoConsItem.Fields("valor_item").Value
End If

      If tabPedidoConsItem.State = 1 Then _
         tabPedidoConsItem.Close

      SQL = "insert into PEDIDOITEM "
      SQL = SQL & "("
         SQL = SQL & "PEDIDO_ID,SEQ_ID,PRODUTO_ID,QTD_PEDIDA,VALOR_ITEM,CFOP_ID,STRIBUTARIA,STATUS,TIPO_REG,PERCICMS,PERC_desc,PRECO_CUSTO"
      SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & PEDIDO_ID_N
         SQL = SQL & "," & SEQ_ID_N
         SQL = SQL & "," & PRODUTO_ID_N
         SQL = SQL & ",'" & tpMOEDA(QTDE_N) & "'"
         SQL = SQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"
         SQL = SQL & ",'" & strCFOP_ITEM & "'"
         SQL = SQL & ",'" & ST_PRODUTO_A & "'"
         SQL = SQL & ",'P'"
         SQL = SQL & ",'PC'"
         SQL = SQL & ",'" & tpMOEDA(PERCICMS_N) & "'"
         SQL = SQL & ",'" & tpMOEDA(0) & "'"
         SQL = SQL & ",'" & tpMOEDA(PRECO_CUSTO_N) & "'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITENS_PEDIDO_NFE"
End Sub
