VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{EDF439C0-99E5-11CF-AFF3-004005100200}#8.0#0"; "PVMarq.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProdutoEntrega 
   BackColor       =   &H00000000&
   Caption         =   "Entrega de Mercadorias"
   ClientHeight    =   7305
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PRODUTOENTREGA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12000
      TabIndex        =   40
      Top             =   480
      Width           =   1455
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   13530
      DesignHeight    =   7305
   End
   Begin MSComctlLib.StatusBar BARRAECF 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6930
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   45
      TabIndex        =   16
      Top             =   720
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   10821
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Entregas &Agendadas"
      TabPicture(0)   =   "PRODUTOENTREGA.frx":5C12
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstENTREGA"
      Tab(0).Control(1)=   "TimerBar"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Entregas &Realizadas"
      TabPicture(1)   =   "PRODUTOENTREGA.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstRealizada"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Cadastra Entrega"
      TabPicture(2)   =   "PRODUTOENTREGA.frx":5C4A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Line1(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label6"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label7"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label8"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label12"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Line1(1)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Line1(2)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Line1(4)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label4"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label13"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "lblTxEntrega"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "lstPedidoItem"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "FlexTel"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtDtPedido"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtDtEntrega"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtDtAgenda"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "cmdPesquisar"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "cmdGravar"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "cmdSAIR"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "cmdLIMPAR"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "cmdExcluir"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtCliente"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtRua"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtBairro"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "txtCidade"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtComp"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "txtCep"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "txtUF"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "txtPedido"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "cmdResidencial"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "cmdCobrança"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "cmdComercial"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "txtCNPJCPF"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "cmbAtendente"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "cmbAtendenteAUX"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "cmbEntregador"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "cmbEntregadorAUX"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "txtOBS"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "cmdImprimir"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "txtTxEntrega"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "chkTxEntrega"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).ControlCount=   46
      Begin VB.CheckBox chkTxEntrega 
         Height          =   240
         Left            =   9840
         TabIndex        =   47
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtTxEntrega 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   11760
         MaxLength       =   12
         TabIndex        =   45
         Top             =   1320
         Width           =   1560
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Produção"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   4440
         Picture         =   "PRODUTOENTREGA.frx":5C66
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5160
         Width           =   900
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
         Height          =   1350
         Left            =   7200
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   3600
         Width           =   6105
      End
      Begin VB.ComboBox cmbEntregadorAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   10920
         TabIndex        =   39
         ToolTipText     =   "Selecione um usuario."
         Top             =   5640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbEntregador 
         Height          =   360
         Left            =   10800
         TabIndex        =   37
         ToolTipText     =   "Selecione um usuario."
         Top             =   5640
         Width           =   2535
      End
      Begin VB.ComboBox cmbAtendenteAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   10920
         TabIndex        =   36
         ToolTipText     =   "Selecione um usuario."
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbAtendente 
         Height          =   360
         Left            =   10800
         TabIndex        =   34
         ToolTipText     =   "Selecione um usuario."
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox txtCNPJCPF 
         Height          =   360
         Left            =   8400
         MaxLength       =   100
         TabIndex        =   33
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdComercial 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Comerial"
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
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
         Top             =   2400
         Width           =   1185
      End
      Begin VB.CommandButton cmdCobrança 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cobrança"
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
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
         Top             =   2160
         Width           =   1185
      End
      Begin VB.CommandButton cmdResidencial 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Residencial"
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
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
         Top             =   1920
         Width           =   1185
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtUF 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   11280
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtCep 
         Height          =   360
         Left            =   9480
         MaxLength       =   8
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtComp 
         Height          =   360
         Left            =   11280
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtCidade 
         Height          =   360
         Left            =   8040
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtBairro 
         Height          =   360
         Left            =   8880
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtRua 
         Height          =   360
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   4455
      End
      Begin VB.Timer TimerBar 
         Interval        =   100
         Left            =   -75000
         Top             =   0
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   3360
         Picture         =   "PRODUTOENTREGA.frx":2379F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5160
         Width           =   900
      End
      Begin VB.CommandButton cmdLIMPAR 
         Caption         =   "&Limpar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1200
         Picture         =   "PRODUTOENTREGA.frx":23E5A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5160
         Width           =   900
      End
      Begin VB.CommandButton cmdSAIR 
         Caption         =   "&Sair"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         Picture         =   "PRODUTOENTREGA.frx":2445C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5160
         Width           =   900
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2280
         Picture         =   "PRODUTOENTREGA.frx":24B2C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5160
         Width           =   900
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2400
         Picture         =   "PRODUTOENTREGA.frx":25F19
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   405
      End
      Begin MSMask.MaskEdBox txtDtAgenda 
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtEntrega 
         Height          =   360
         Left            =   5040
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSMask.MaskEdBox txtDtPedido 
         Height          =   360
         Left            =   11760
         TabIndex        =   19
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstENTREGA 
         Height          =   5600
         Left            =   -74955
         TabIndex        =   0
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   480
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   9869
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Previsto"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Complemento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Bairro"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cidade"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "UF"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "CEP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Entregador"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Montador"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstRealizada 
         Height          =   5600
         Left            =   -74955
         TabIndex        =   1
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   480
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   9869
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Previsto"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Entregue"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Complemento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Bairro"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cidade"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "UF"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "CEP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Entregador"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Montador"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid FlexTel 
         Height          =   705
         Left            =   7200
         TabIndex        =   41
         Top             =   2760
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   1244
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lstPedidoItem 
         Height          =   2985
         Left            =   60
         TabIndex        =   43
         Top             =   1920
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   5265
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12648384
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
            Object.Width           =   1960
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Produto"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qtde"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valr.Unitário"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Desconto"
            Object.Width           =   1877
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Valr.Total"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblTxEntrega 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Taxa Entrega = "
         Height          =   240
         Left            =   10215
         TabIndex        =   46
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Entregador:"
         Height          =   240
         Left            =   9600
         TabIndex        =   38
         Top             =   5640
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Atendente:"
         Height          =   240
         Left            =   9600
         TabIndex        =   35
         Top             =   5160
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   4
         X1              =   0
         X2              =   13440
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   2
         X1              =   0
         X2              =   13440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   13440
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   255
         Left            =   10920
         TabIndex        =   28
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   7200
         TabIndex        =   27
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cp:"
         Height          =   240
         Left            =   10800
         TabIndex        =   26
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rua:"
         Height          =   255
         Left            =   7200
         TabIndex        =   25
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Entrega:"
         Height          =   240
         Left            =   3960
         TabIndex        =   24
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Agendamento:"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   13440
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Pedido:"
         Height          =   240
         Left            =   10680
         TabIndex        =   22
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   240
         Left            =   3000
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
   End
   Begin PVMarqueeLib.PVMarquee PVMarquee1 
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   13575
      _Version        =   524288
      _ExtentX        =   23945
      _ExtentY        =   873
      _StockProps     =   29
      Text            =   "Controle Agenda Entrega"
      ForeColor       =   16711680
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Frame           =   5
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Text            =   "Controle Agenda Entrega"
   End
End
Attribute VB_Name = "frmProdutoEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Entrega_ID_N  As Long

Private Sub chkTxEntrega_Click()
   If chkTxEntrega.Value = 0 Then
      lblTxEntrega.Enabled = False
      txtTxEntrega.Enabled = False
      txtTxEntrega.Text = 0
      Else
         lblTxEntrega.Enabled = True
         txtTxEntrega.Enabled = True
         If Trim(txtTxEntrega.Text) <> "" Then
            If IsNumeric(txtTxEntrega.Text) Then
               VALOR_ITEM_N = 0 & txtTxEntrega.Text
               If VALOR_ITEM_N <= 0 Then _
                  txtTxEntrega.Text = "15,00"
            End If
         End If
   End If
End Sub

Private Sub Form_Load()
   ATUALIZA_TABELA_ENTREGA
   SETA_GRID_ENTREGA
   SETA_GRID_REALIZADA
   PESSOA_ID_N = 0
   ATUALIZA_GRID_FONE
   SETA_FONE
   Entrega_ID_N = 0
   txtOBS.Text = ""

   CARREGA_COMBO
   If Trim(cmbAtendente.Text) = "" Then
      cmbAtendenteAUX.Text = "" & USUARIO_ID_N
      cmbAtendente.Text = "" & TRAZ_NOME_USUARIO(USUARIO_ID_N)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
      Case vbKeyF9
         SETA_GRID_ENTREGA
         SETA_GRID_REALIZADA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lstENTREGA_GotFocus()
   MOSTRA_MSG "ESC-Sair", "F7-Imprimir Entrega", "F8-Imprimir Todos", "", ""
End Sub

Private Sub lstENTREGA_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstENTREGA.SelectedItem.Text) Then
      LIMPA_TELA
      txtPedido.Text = "" & lstENTREGA.SelectedItem.Text
      SSTab1.Tab = 2
      txtPedido.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstENTREGA_DblClick"
End Sub

Private Sub lstentrega_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstENTREGA, ColumnHeader
End Sub

Private Sub lstENTREGA_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         FORMULA_REL = ""
         If Not IsNull(lstENTREGA.SelectedItem.Text) Then _
            FORMULA_REL = "{vwRel_ENTREGA.pedido_id} = " & lstENTREGA.SelectedItem.Text

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_mapa_entrega.rpt"
         frmRELATORIO10.Show 1
      Case vbKeyF8
         FORMULA_REL = "isnull({vwRel_ENTREGA.dt_entrega}) "

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_mapa_entrega.rpt"
         frmRELATORIO10.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lstrealizada_GotFocus()
   MOSTRA_MSG "ESC-Sair", "F7-Imprimir Entrega", "F8-Imprimir Todos", "", ""
End Sub

Private Sub lstrealizada_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstRealizada, ColumnHeader
End Sub

Private Sub lstrealizada_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         FORMULA_REL = ""
         If Not IsNull(lstRealizada.SelectedItem.Text) Then _
            FORMULA_REL = "{vwRel_ENTREGA.pedido_id} = " & lstRealizada.SelectedItem.Text

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_mapa_entrega.rpt"
         frmRELATORIO10.Show 1
      Case vbKeyF8
         FORMULA_REL = "not isnull({vwRel_ENTREGA.dt_entrega}) "

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_mapa_entrega.rpt"
         frmRELATORIO10.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab <> 0 Then
      TimerBar.Enabled = False
      Else: TimerBar.Enabled = True
   End If
   If SSTab1.Tab = 0 Then
      SETA_GRID_ENTREGA
   End If
   If SSTab1.Tab = 1 Then
      SETA_GRID_REALIZADA
   End If
   If SSTab1.Tab = 2 Then
      MOSTRA_MSG "ESC-Sair", "", "", "", ""
      txtPedido.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSTab1_Click"
End Sub

Private Sub TimerBar_Timer()
'On Error GoTo ERRO_TRATA

   Dim aux As Integer
   aux = 1 + 10
   If aux <= 100 Then
      Else
         SETA_GRID_ENTREGA
         SETA_GRID_REALIZADA
   End If
   DoEvents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TimerBar_Timer"
   End
End Sub

Private Sub txtBairro_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe o Bairro e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtCep_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe o CEP e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtcidade_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe a Cidade e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtComp_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe Data de Complemento e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtDtAgenda_GotFocus()
   txtDtAgenda.PromptInclude = True
   MOSTRA_MSG "ESC-Sair", "Informe Data de Agendamento e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtDtEntrega_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe Data de Entrega e tecle <<Enter>>", "", "", ""
   txtDtEntrega.PromptInclude = True
End Sub

Private Sub txtPedido_GotFocus()
   MOSTRA_MSG "ESC-Sair", "F7-Consulta Pedido Venda", "Informe NºPedido e tecle <<Enter>>", "", ""
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_PEDIDO
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyPress"
End Sub

Private Sub txtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         BUSCA_VENDA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyDown"
End Sub

Private Sub txtDtAgenda_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      txtDtEntrega.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtAgenda_KeyPress"
End Sub

Private Sub txtDtAgenda_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtAgenda.PromptInclude = False
   If Trim(txtDtAgenda.Text) = "" Then _
      txtDtAgenda.Text = Now
   txtDtAgenda.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtAgenda_LostFocus"
End Sub

Private Sub txtDtEntrega_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      txtRua.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEntrega_KeyPress"
End Sub

Private Sub txtRua_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe a Rua e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtrua_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrua_KeyPress"
End Sub

Private Sub txtcomp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcomp_KeyPress"
End Sub

Private Sub txtbairro_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbairro_KeyPress"
End Sub

Private Sub txtcep_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcep_KeyPress"
End Sub

Private Sub txtcidade_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_KeyPress"
End Sub

Private Sub txtuf_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Informe Estado (UF) e tecle <<Enter>>", "", "", ""
End Sub

Private Sub txtuf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtuf_KeyPress"
End Sub

Private Sub cmdPesquisar_Click()
'On Error GoTo ERRO_TRATA

   BUSCA_VENDA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdPesquisar_Click"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   LIMPA_TELA
   txtPedido.SetFocus
End Sub

Private Sub cmdExcluir_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then
      MsgBox "Pedido inválido, não permitido."
      txtPedido.SetFocus
      Exit Sub
   End If

   Dim TabEntrega As New ADODB.Recordset

   If TabEntrega.State = 1 Then _
      TabEntrega.Close

   SQL = "select * from ENTREGA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   TabEntrega.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEntrega.EOF Then
      Msg = " Confirma exclusão desse agendamento ?"
      PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         CRITERIO_A = ""
         frmPedidoCancela.txtPedido.Text = 0 & txtPedido.Text
         frmPedidoCancela.Show 1
         SQL = ""
         CRITERIO_A = ""

         SQL = "spOBSEntrega " & 3 & "," & 0 & "," & TabEntrega.Fields("Entrega_ID").Value & "" & ",'" & Trim(txtOBS.Text) & "'" & ",'" & Now & "'"
         CONECTA_RETAGUARDA.Execute "EXEC " & SQL

         SQL = "delete from ENTREGA "
         SQL = SQL & " where pedido_id = " & txtPedido.Text
         CONECTA_RETAGUARDA.Execute SQL
         LIMPA_TELA
      End If
   End If
   If TabEntrega.State = 1 Then _
      TabEntrega.Close

   txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdExcluir_Click"
End Sub

Private Sub cmdImprimir_Click()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{vwRel_ENTREGA.pedido_id} = " & txtPedido.Text

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_mapa_producao.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdImprimir_Click"
End Sub

Private Sub CmdGravar_Click()
   GRAVA_ENTREGA
   txtPedido.SetFocus
End Sub

Private Sub cmdCobrança_Click()
   BUSCA_ENDEREÇO "B"
End Sub

Private Sub cmdResidencial_Click()
   BUSCA_ENDEREÇO "R"
End Sub

Private Sub cmdComercial_Click()
   BUSCA_ENDEREÇO "C"
End Sub

Private Sub cmbENTREGADOR_Click()
On Error Resume Next

   cmbEntregadorAUX.ListIndex = cmbEntregador.ListIndex
   
End Sub

Private Sub cmbAtendente_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Selecione o Entregador", "", "", ""
End Sub

Private Sub cmbENTREGADOR_GotFocus()
   MOSTRA_MSG "ESC-Sair", "Selecione o entregador", "", "", ""
End Sub

Private Sub cmbAtendente_Click()
On Error Resume Next

   cmbAtendenteAUX.ListIndex = cmbAtendente.ListIndex
   
End Sub

Sub GRAVA_ENTREGA()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then
      MsgBox "Pedido inválido, não permitido."
      txtPedido.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtPedido.Text) Then
      MsgBox "Pedido inválido, não permitido."
      txtPedido.SetFocus
      Exit Sub
   End If
   If Trim(txtCliente.Text) = "" Then
      MsgBox "Pedido inválido, não permitido."
      txtPedido.SetFocus
      Exit Sub
   End If
   txtDtPedido.PromptInclude = False
   If Trim(txtDtPedido.Text) = "" Then
      MsgBox "Pedido inválido, não permitido."
      txtPedido.SetFocus
      Exit Sub
   End If
   txtDtPedido.PromptInclude = True
   txtDtAgenda.PromptInclude = False
   If Trim(txtDtAgenda.Text) = "" Then
      MsgBox "Data de Agendamento inválida, não permitido."
      txtDtAgenda.SetFocus
      Exit Sub
   End If
   txtDtAgenda.PromptInclude = True

   txtDtEntrega.PromptInclude = False
   'If Trim(txtDtEntrega.Text) = "" Then
   '   MsgBox "Data de Agendamento inválida, não permitido."
   '   txtDtEntrega.SetFocus
   '   Exit Sub
   'End If
   txtDtEntrega.PromptInclude = True

   If Trim(txtRua.Text) = "" Then
      'MsgBox "Rua inválida, não permitido."
      'txtRua.SetFocus
      'Exit Sub
   End If
   If Trim(txtBairro.Text) = "" Then
      'MsgBox "Bairro inválido, não permitido."
      'txtBairro.SetFocus
      'Exit Sub
   End If
   If Trim(txtCidade.Text) = "" Then
      'MsgBox "Cidade inválida, não permitido."
      'txtCidade.SetFocus
      'Exit Sub
   End If

   If Trim(cmbAtendenteAUX.Text) = "" Then _
      cmbAtendenteAUX.Text = "null"
   If Trim(cmbEntregadorAUX.Text) = "" Then _
      cmbEntregadorAUX.Text = "null"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ENTREGA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      Entrega_ID_N = 0 & TabTemp.Fields("entrega_id").Value
      SQL = "update ENTREGA SET "
         SQL = SQL & " PESSOA_ID = " & PESSOA_ID_N                            'PESSOA_ID
         'SQL = SQL & ", PEDIDO_ID = " & PEDIDO_ID_N                           'PEDIDO_ID
         'SQL = SQL & ", EMPRESA_ID = " & EMPRESA_ID_N                         'EMPRESA_ID
         'SQL = SQL & ", DT_CAD = '" & Now & "'"                               'DT_CAD
         'SQL = SQL & ", DT_AGENDA = '" & txtDtAgenda.Text & "'"               'DT_AGENDA

         txtDtEntrega.PromptInclude = False
         If Trim(txtDtEntrega.Text) = "" Then
            SQL = SQL & ", DT_ENTREGA = null"                                 'DT_ENTREGA
            Else
               txtDtEntrega.PromptInclude = True
               SQL = SQL & ", DT_ENTREGA = '" & txtDtEntrega.Text & "'"    'DT_ENTREGA
         End If

         SQL = SQL & ", cep_id = '" & Trim(txtCep.Text) & "'"                 'CEP
         SQL = SQL & ", RUA = '" & Trim(txtRua.Text) & "'"                    'RUA
         SQL = SQL & ", COMPLEMENTO = '" & Trim(txtComp.Text) & "'"           'COMPLEMENTO
         SQL = SQL & ", BAIRRO = '" & Trim(txtBairro.Text) & "'"              'BAIRRO
         SQL = SQL & ", CIDADE = '" & Trim(txtCidade.Text) & "'"              'CIDADE
         SQL = SQL & ", UF = '" & Trim(txtUF.Text) & "'"                      'UF
         SQL = SQL & ", atendente_ID = " & cmbAtendenteAUX.Text               'atendente_ID
         SQL = SQL & ", entregador_ID = " & cmbEntregadorAUX.Text             'entregador_ID
         SQL = SQL & ", atendente = '" & Trim(cmbAtendente.Text) & "'"        'atendente
         SQL = SQL & ", entregador = '" & Trim(cmbEntregador.Text) & "'"      'entregador

      SQL = SQL & " where pedido_id = " & txtPedido.Text
      Else
         Entrega_ID_N = 0 & MAX_ID("entrega_id", "ENTREGA", "", "", "", "")

         SQL = "insert into ENTREGA "
            SQL = SQL & "(ENTREGA_ID,PESSOA_ID,PEDIDO_ID,EMPRESA_ID,DT_CAD,DT_AGENDA,DT_ENTREGA,"
            SQL = SQL & "CEP_ID,RUA,COMPLEMENTO,BAIRRO,CIDADE,UF,"
            SQL = SQL & "ATENDENTE_ID,ATENDENTE,ENTREGADOR_ID,ENTREGADOR)"
         SQL = SQL & " values("
            SQL = SQL & Entrega_ID_N                                    'ENTREGA_ID
            SQL = SQL & "," & PESSOA_ID_N                               'PESSOA_ID
            SQL = SQL & "," & PEDIDO_ID_N                               'PEDIDO_ID
            SQL = SQL & "," & EMPRESA_ID_N                              'EMPRESA_ID

            SQL = SQL & ",'" & Now & "'"                                'DT_CAD
            SQL = SQL & ",'" & txtDtAgenda.Text & "'"                   'DT_AGENDA

            txtDtEntrega.PromptInclude = False
            If Trim(txtDtEntrega.Text) = "" Then
               SQL = SQL & ", null"                                     'DT_ENTREGA
               Else
                  txtDtEntrega.PromptInclude = True
                  SQL = SQL & ",'" & txtDtEntrega.Text & "'"            'DT_ENTREGA
            End If

            SQL = SQL & ",'" & Trim(txtCep.Text) & "'"                  'CEP
            SQL = SQL & ",'" & Trim(txtRua.Text) & "'"                  'RUA
            SQL = SQL & ",'" & Trim(txtComp.Text) & "'"                 'COMPLEMENTO
            SQL = SQL & ",'" & Trim(txtBairro.Text) & "'"               'BAIRRO
            SQL = SQL & ",'" & Trim(txtCidade.Text) & "'"               'CIDADE
            SQL = SQL & ",'" & Trim(txtUF.Text) & "'"                   'UF

            SQL = SQL & "," & cmbAtendenteAUX.Text                      'ATENDENTE_ID
            SQL = SQL & ",'" & Trim(cmbAtendente.Text) & "'"            'ATENDENTE

            SQL = SQL & "," & cmbEntregadorAUX.Text                     'ENTREGADOR_ID
            SQL = SQL & ",'" & Trim(cmbEntregador.Text) & "'"           'ENTREGADOR

         SQL = SQL & " )"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   If Trim(txtOBS.Text) <> "" Then _
      GRAVA_OBS

   If Trim(txtTxEntrega.Text) <> "" Then
      If IsNumeric(txtTxEntrega.Text) Then
         VALOR_ITEM_N = 0 & txtTxEntrega.Text

         If TabTemp.State = 1 Then _
            TabTemp.Close
         SQL = "select pedido_id from PEDIDOENCOMENDA WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            spEncomenda 2, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_ITEM_N
            Else: spEncomenda 1, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_ITEM_N
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If

   MsgBox "Operação realizada com sucesso."
   LIMPA_TELA
   txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ENTREGA"
End Sub

Sub MOSTRA_PEDIDO()
'On Error GoTo ERRO_TRATA

   lstPedidoItem.ListItems.Clear

   If Trim(txtPedido.Text) = "" Then
      MsgBox "Pedido inválido, não permitido."
      txtPedido.SetFocus
      Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtCNPJCPF.Text = "" & Trim(TabTemp.Fields("cgccpf").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & Trim(TabTemp.Fields("cgccpf").Value) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         PESSOA_ID_N = TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value
      txtCliente.Text = Trim(TabTemp.Fields("nome_cliente").Value)
      txtDtPedido.PromptInclude = False
         txtDtPedido.Text = TabTemp.Fields("dt_req").Value
      txtDtPedido.PromptInclude = True

'===========================
      SETA_GRID_PRODUTOS
'==============================================

      ATUALIZA_GRID_FONE
      SETA_FONE
      MOSTRA_ENTREGA
      MOSTRA_OBS

      'SendKeys "{tab}"
      txtDtAgenda.SetFocus
      Else
         MsgBox "Pedido não encontrado."
         LIMPA_TELA
         txtPedido.SetFocus
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PEDIDO"
End Sub

Sub BUSCA_VENDA()
'On Error GoTo ERRO_TRATA

   Dim TabTipovenda  As New ADODB.Recordset

   If TabTipovenda.State = 1 Then _
      TabTipovenda.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where descricao = 'ENCOMENDA'"
   TabTipovenda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTipovenda.EOF Then
      frmPedidoConsulta.cmbForma.Text = "" & Trim(TabTipovenda.Fields("descricao").Value)
      frmPedidoConsulta.cmbFormaAUX.Text = "" & Trim(TabTipovenda.Fields("tipovenda_id").Value)
   End If
   If TabTipovenda.State = 1 Then _
      TabTipovenda.Close

   CRITERIO_A = ""
   CNPJCPF_A = ""
   frmPedidoConsulta.Show 1
   If PEDIDO_ID_N > 0 Then
      txtPedido.Text = PEDIDO_ID_N
      txtPedido.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_VENDA"
End Sub

Sub LIMPA_TELA()
'On Error GoTo ERRO_TRATA

   CARREGA_COMBO
   If Trim(cmbAtendente.Text) = "" Then
      cmbAtendenteAUX.Text = "" & USUARIO_ID_N
      cmbAtendente.Text = "" & TRAZ_NOME_USUARIO(USUARIO_ID_N)
   End If
   lstPedidoItem.ListItems.Clear
   txtOBS.Text = ""
   Entrega_ID_N = 0

   ATUALIZA_GRID_FONE
   txtCNPJCPF.Text = ""
   PEDIDO_ID_N = 0
   PESSOA_ID_N = 0
   txtPedido.Text = ""
   txtCliente.Text = ""
   txtDtPedido.PromptInclude = False
   txtDtPedido.Text = ""
   txtDtAgenda.PromptInclude = False
   txtDtAgenda.Text = ""
   txtDtEntrega.PromptInclude = False
   txtDtEntrega.Text = ""
   txtRua.Text = ""
   txtComp.Text = ""
   txtBairro.Text = ""
   txtCep.Text = ""
   txtCidade.Text = ""
   txtUF.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TELA"
End Sub

Private Sub SETA_GRID_ENTREGA()
'On Error GoTo ERRO_TRATA

   lstENTREGA.Visible = False
   lstENTREGA.ListItems.Clear

   SQL = "select * from vwRel_ENTREGA WITH (NOLOCK)"

   SQL = SQL & " where dt_entrega is null"
   SQL = SQL & " order by dt_agenda "

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      DoEvents
      DATA_INI = TabTemp.Fields("dt_agenda").Value

      Set item = lstENTREGA.ListItems.Add(, "seq." & TabTemp.Fields("pedido_id").Value, TabTemp.Fields("pedido_id").Value)

      item.SubItems(1) = "" & TabTemp.Fields("NOME_cliente").Value
      item.SubItems(2) = "" & TabTemp.Fields("dt_agenda").Value

      item.SubItems(3) = "" & TabTemp.Fields("RUA").Value
      item.SubItems(4) = "" & TabTemp.Fields("complemento").Value
      item.SubItems(5) = "" & TabTemp.Fields("bairro").Value
      item.SubItems(6) = "" & TabTemp.Fields("cidade").Value
      item.SubItems(7) = "" & TabTemp.Fields("uf").Value
      item.SubItems(8) = "" & TabTemp.Fields("cep_id").Value
'==================
      If Not IsNull(TabTemp.Fields("entregador_id").Value) Then
         If IsNumeric(TabTemp.Fields("entregador_id").Value) Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select nome from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & TabTemp.Fields("entregador_id").Value
            SQL = SQL & " and status = 1"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               item.SubItems(9) = "" & Trim(TabConsulta.Fields(0).Value)

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If

      If Not IsNull(TabTemp.Fields("entregador_id").Value) Then
         If IsNumeric(TabTemp.Fields("entregador_id").Value) Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select nome from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & TabTemp.Fields("entregador_id").Value
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               item.SubItems(10) = "" & Trim(TabConsulta.Fields(0).Value)

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
'==================
      If DATA_INI <= Date Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(3).ForeColor = vbRed
         item.ListSubItems(4).ForeColor = vbRed
         item.ListSubItems(5).ForeColor = vbRed
         item.ListSubItems(6).ForeColor = vbRed
         item.ListSubItems(7).ForeColor = vbRed
         item.ListSubItems(8).ForeColor = vbRed
         Else
            item.ForeColor = vbBlue
            item.ListSubItems(1).ForeColor = vbBlue
            item.ListSubItems(2).ForeColor = vbBlue
            item.ListSubItems(3).ForeColor = vbBlue
            item.ListSubItems(4).ForeColor = vbBlue
            item.ListSubItems(5).ForeColor = vbBlue
            item.ListSubItems(6).ForeColor = vbBlue
            item.ListSubItems(7).ForeColor = vbBlue
            item.ListSubItems(8).ForeColor = vbBlue
      End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstENTREGA.Visible = True
'vbBlack  vbRed  vbGreen  vbYellow  vbBlue  vbMagenta  vbCyan  vbWhite
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ENTREGA"
End Sub

Private Sub SETA_GRID_REALIZADA()
'On Error GoTo ERRO_TRATA

   lstRealizada.Visible = False
   lstRealizada.ListItems.Clear

   SQL = "select * from vwRel_ENTREGA WITH (NOLOCK)"

   SQL = SQL & " where dt_entrega is not null"
   SQL = SQL & " order by dt_agenda "

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      DoEvents
      'DATA_INI = TabTemp.Fields("dt_agenda").Value

      Set item = lstRealizada.ListItems.Add(, "seq." & TabTemp.Fields("pedido_id").Value, TabTemp.Fields("pedido_id").Value)

      item.SubItems(1) = "" & TabTemp.Fields("NOME_cliente").Value
      item.SubItems(2) = "" & TabTemp.Fields("dt_agenda").Value
      item.SubItems(3) = "" & TabTemp.Fields("dt_entrega").Value

      item.SubItems(4) = "" & TabTemp.Fields("RUA").Value
      item.SubItems(5) = "" & TabTemp.Fields("complemento").Value
      item.SubItems(6) = "" & TabTemp.Fields("bairro").Value
      item.SubItems(7) = "" & TabTemp.Fields("cidade").Value
      item.SubItems(8) = "" & TabTemp.Fields("uf").Value
      item.SubItems(9) = "" & TabTemp.Fields("cep_id").Value
'==================
      If Not IsNull(TabTemp.Fields("entregador_id").Value) Then
         If IsNumeric(TabTemp.Fields("entregador_id").Value) Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select nome from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & TabTemp.Fields("entregador_id").Value
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               item.SubItems(9) = "" & Trim(TabConsulta.Fields(0).Value)

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If

      If Not IsNull(TabTemp.Fields("entregador_id").Value) Then
         If IsNumeric(TabTemp.Fields("entregador_id").Value) Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select nome from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & TabTemp.Fields("entregador_id").Value
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               item.SubItems(10) = "" & Trim(TabConsulta.Fields(0).Value)

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
'==================
      'If DATA_INI <= Date Then
      '   Item.ForeColor = vbRed
      '   Item.ListSubItems(1).ForeColor = vbRed
      '   Item.ListSubItems(2).ForeColor = vbRed
      '   Item.ListSubItems(3).ForeColor = vbRed
      '   Item.ListSubItems(4).ForeColor = vbRed
      '   Item.ListSubItems(5).ForeColor = vbRed
      '   Item.ListSubItems(6).ForeColor = vbRed
      '   Item.ListSubItems(7).ForeColor = vbRed
      '   Item.ListSubItems(8).ForeColor = vbRed
      '   Else
            item.ForeColor = vbBlue
            item.ListSubItems(1).ForeColor = vbBlue
            item.ListSubItems(2).ForeColor = vbBlue
            item.ListSubItems(3).ForeColor = vbBlue
            item.ListSubItems(4).ForeColor = vbBlue
            item.ListSubItems(5).ForeColor = vbBlue
            item.ListSubItems(6).ForeColor = vbBlue
            item.ListSubItems(7).ForeColor = vbBlue
            item.ListSubItems(8).ForeColor = vbBlue
            'Item.ListSubItems(9).ForeColor = vbBlue
      'End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstRealizada.Visible = True
'vbBlack  vbRed  vbGreen  vbYellow  vbBlue  vbMagenta  vbCyan  vbWhite
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_REALIZADA"
End Sub

Sub MOSTRA_MSG(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   If Trim(Msg1) <> "" Then
      BARRAECF.Panels.Clear
      BARRAECF.Panels.Add (1)
      BARRAECF.Panels(1).Text = Trim(Msg1)
      BARRAECF.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         BARRAECF.Panels.Add (2)
         BARRAECF.Panels(2).Text = Trim(Msg2)
         BARRAECF.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            BARRAECF.Panels.Add (3)
            BARRAECF.Panels(3).Text = Trim(Msg3)
            BARRAECF.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               BARRAECF.Panels.Add (4)
               BARRAECF.Panels(4).Text = Trim(Msg4)
               BARRAECF.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  BARRAECF.Panels.Add (5)
                  BARRAECF.Panels(5).Text = Trim(Msg5)
                  BARRAECF.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If
End Sub

Sub MOSTRA_ENTREGA()
'On Error GoTo ERRO_TRATA

   CmdGravar.Enabled = True
   cmdExcluir.Enabled = True

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from ENTREGA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      Entrega_ID_N = 0 & TabConsulta.Fields("entrega_id").Value
      txtDtAgenda.PromptInclude = False
         txtDtAgenda.Text = "" & Trim(TabConsulta.Fields("DT_AGENDA").Value)
      txtDtAgenda.PromptInclude = True

      txtDtEntrega.PromptInclude = False
         txtDtEntrega.Text = "" & Trim(TabConsulta.Fields("DT_ENTREGA").Value)
      txtDtEntrega.PromptInclude = True

      If Not IsNull(TabConsulta.Fields("DT_ENTREGA").Value) Then
         If TabConsulta.Fields("DT_ENTREGA").Value >= TabConsulta.Fields("DT_agenda").Value Then
            MsgBox "Entrega já realizada, permitido somente consulta."
            CmdGravar.Enabled = False
            cmdExcluir.Enabled = False
         End If
      End If

      txtRua.Text = "" & Trim(TabConsulta.Fields("RUA").Value)
      txtComp.Text = "" & Trim(TabConsulta.Fields("complemento").Value)
      txtBairro.Text = "" & Trim(TabConsulta.Fields("bairro").Value)
      txtCep.Text = "" & Trim(TabConsulta.Fields("cep_id").Value)
      txtCidade.Text = "" & Trim(TabConsulta.Fields("cidade").Value)
      txtUF.Text = "" & Trim(TabConsulta.Fields("uf").Value)

      If Not IsNull(TabConsulta.Fields("atendente_id").Value) Then
         If IsNumeric(TabConsulta.Fields("atendente_id").Value) Then
            cmbAtendenteAUX.Text = "" & Trim(TabConsulta.Fields("atendente_id").Value)

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select nome from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & TabConsulta.Fields("atendente_id").Value
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               cmbAtendente.Text = "" & Trim(TabTemp.Fields(0).Value)
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If

      If Not IsNull(TabConsulta.Fields("entregador_id").Value) Then
         If IsNumeric(TabConsulta.Fields("entregador_id").Value) Then
            cmbEntregadorAUX.Text = "" & Trim(TabConsulta.Fields("entregador_id").Value)

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select nome from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & TabConsulta.Fields("entregador_id").Value
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then _
               cmbEntregador.Text = "" & Trim(TabTemp.Fields(0).Value)
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ENTREGA"
End Sub

Sub BUSCA_ENDEREÇO(Tipo_End As String)
'On Error GoTo ERRO_TRATA

   txtRua.Text = ""
   txtComp.Text = ""
   txtBairro.Text = ""
   txtCep.Text = ""
   txtCidade.Text = ""
   txtUF.Text = ""

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select ENDERECO.ENDERECO_ID, ENDERECO.PESSOA_ID, ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
   SQL = SQL & " Endereco.tipo , Endereco.Numero, PESSOA.CNPJCPF, CEP.Cidade, CEP.UF, CEP.IBGE_ID"
   SQL = SQL & " from PESSOA WITH (NOLOCK)"
   SQL = SQL & " RIGHT OUTER JOIN ENDERECO WITH (NOLOCK)"
   SQL = SQL & " ON PESSOA.PESSOA_ID = ENDERECO.PESSOA_ID "
   SQL = SQL & " LEFT OUTER JOIN CEP WITH (NOLOCK)"
   SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID "

   SQL = SQL & " where tipo = '" & Trim(Tipo_End) & "'"
   SQL = SQL & " and cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      txtRua.Text = "" & Trim(TabConsulta.Fields("rua").Value)
      txtComp.Text = "" & Trim(TabConsulta.Fields("complemento").Value)
      txtBairro.Text = "" & Trim(TabConsulta.Fields("bairro").Value)
      txtCep.Text = "" & Trim(TabConsulta.Fields("cep_id").Value)
      txtCidade.Text = "" & Trim(TabConsulta.Fields("cidade").Value)
      txtUF.Text = "" & Trim(TabConsulta.Fields("uf").Value)
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_ENDEREÇO"
End Sub

Sub CARREGA_COMBO()
'On Error GoTo ERRO_TRATA

   cmbEntregador.Clear
   cmbEntregadorAUX.Clear
   cmbAtendente.Clear
   cmbAtendenteAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, DESCR.DESCRICAO, USUARIO.NOME, USUARIO.STATUS"
   SQL = SQL & " from DESCR WITH (NOLOCK)"
   SQL = SQL & " RIGHT OUTER JOIN USUARIO WITH (NOLOCK)"
   SQL = SQL & " ON DESCR.codigo = USUARIO.TIPO "
   SQL = SQL & " WHERE DESCR.TIPO = 'T' "
   SQL = SQL & " AND UPPER(DESCR.DESCRICAO) = 'ATENDENTE' "
   SQL = SQL & " OR UPPER(DESCR.DESCRICAO) = 'ENTREGADOR'"
   SQL = SQL & " and status = 1 "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      If UCase(TabDESCR.Fields("DESCRICAO").Value) = "ATENDENTE" Then
         cmbAtendente.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & TabDESCR.Fields("usuario_id").Value
         cmbAtendenteAUX.AddItem TabDESCR.Fields("usuario_id").Value
         Else
            If UCase(TabDESCR.Fields("DESCRICAO").Value) = "ENTREGADOR" Then
               cmbEntregador.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & TabDESCR.Fields("usuario_id").Value
               cmbEntregadorAUX.AddItem TabDESCR.Fields("usuario_id").Value
            End If
      End If

      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub

Sub ATUALIZA_GRID_FONE()
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
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_GRID_FONE"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from FONE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   ATUALIZA_GRID_FONE
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
      TabAUX.MoveNext
   Loop
   If TabAUX.State = 1 Then _
      TabAUX.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_FONE"
End Sub

Sub SETA_GRID_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      lstPedidoItem.ListItems.Clear

      SQL = "select PRODUTO.CODG_PRODUTO, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "
      SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.PRECO_CUSTO, pedidoitem.seq_id,"
      SQL = SQL & " PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.CFOP_id, pedidoitem.status, "
      SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM, Produto.FORNECEDOR_ID"
      SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where pedidoitem.produto_id = produto.produto_id "
      SQL = SQL & " and PEDIDO_ID = " & txtPedido.Text
      SQL = SQL & " and pedidoitem.status <> 'C'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         VALOR_DESCONTO_N = 0 & TabConsulta.Fields("valor_desconto").Value
         VALOR_ITEM_N = TabConsulta.Fields("qtd_pedida").Value * (TabConsulta.Fields("valor_item").Value - VALOR_DESCONTO_N)

         Set item = lstPedidoItem.ListItems.Add(, "seq." & TabConsulta.Fields("seq_id").Value, Trim(TabConsulta.Fields("codg_produto").Value))
         item.SubItems(1) = "" & Trim(TabConsulta.Fields("descricao").Value)
         item.SubItems(2) = "" & Format(Trim(TabConsulta.Fields("qtd_pedida").Value), strFormatacao3Digitos)
         item.SubItems(3) = "" & Format(Trim(TabConsulta.Fields("valor_item").Value), strFormatacao2Digitos)
         item.SubItems(4) = "" & Format(Trim(TabConsulta.Fields("valor_desconto").Value), strFormatacao2Digitos)
         item.SubItems(5) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

         If Trim(TabConsulta.Fields("status").Value) = "A" Then
            item.ForeColor = vbBlue
            item.ListSubItems(1).ForeColor = vbBlue
            item.ListSubItems(2).ForeColor = vbBlue
            item.ListSubItems(3).ForeColor = vbBlue
            item.ListSubItems(4).ForeColor = vbBlue
            item.ListSubItems(5).ForeColor = vbBlue
         End If
         If Trim(TabConsulta.Fields("status").Value) = "P" Then
            item.ForeColor = vbBlack
            item.ListSubItems(1).ForeColor = vbBlack
            item.ListSubItems(2).ForeColor = vbBlack
            item.ListSubItems(3).ForeColor = vbBlack
            item.ListSubItems(4).ForeColor = vbBlack
            item.ListSubItems(5).ForeColor = vbBlack
         End If
         If Trim(TabConsulta.Fields("status").Value) = "C" Then
            item.ForeColor = vbRed
            item.ListSubItems(1).ForeColor = vbRed
            item.ListSubItems(2).ForeColor = vbRed
            item.ListSubItems(3).ForeColor = vbRed
            item.ListSubItems(4).ForeColor = vbRed
            item.ListSubItems(5).ForeColor = vbRed
         End If
         TabConsulta.MoveNext
         CRITERIO_A = ""
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      lstPedidoItem.Refresh
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PRODUTOS"
End Sub

Sub MOSTRA_OBS()
'On Error GoTo ERRO_TRATA

   Dim TabOBSENTREGA    As New ADODB.Recordset

   If TabOBSENTREGA.State = 1 Then _
      TabOBSENTREGA.Close
   SQL = "select obs from OBSENTREGA WITH (NOLOCK)"
   SQL = SQL & " where entrega_id = " & Entrega_ID_N
   TabOBSENTREGA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabOBSENTREGA.EOF Then _
      txtOBS.Text = "" & Trim(TabOBSENTREGA.Fields(0).Value)
   If TabOBSENTREGA.State = 1 Then _
      TabOBSENTREGA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OBS"
End Sub

Sub GRAVA_OBS()
'On Error GoTo ERRO_TRATA

   Dim TabOBSENTREGA    As New ADODB.Recordset
   Dim OBSENTREGA_ID_N  As Long

   txtOBS.Text = Trim(Replace(txtOBS.Text, "'", "´"))
   txtOBS.Text = Trim(Replace(txtOBS.Text, ",", ";"))

   OBSENTREGA_ID_N = 0
   If TabOBSENTREGA.State = 1 Then _
      TabOBSENTREGA.Close
   SQL = "select OBSENTREGA_id from OBSENTREGA WITH (NOLOCK)"
   SQL = SQL & " where entrega_id = " & Entrega_ID_N
   TabOBSENTREGA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabOBSENTREGA.EOF Then
      OBSENTREGA_ID_N = 0 & Trim(TabOBSENTREGA.Fields(0).Value)

      SQL = "spOBSEntrega " & 2 & "," & OBSENTREGA_ID_N & "," & Entrega_ID_N & "" & ",'" & Trim(txtOBS.Text) & "'" & ",'" & Now & "'"
      CONECTA_RETAGUARDA.Execute "EXEC " & SQL
      Else
         OBSENTREGA_ID_N = 0 & MAX_ID("OBSENTREGA_ID", "OBSENTREGA", "", "", "", "")
         SQL = "spOBSEntrega " & 1 & "," & OBSENTREGA_ID_N & "," & Entrega_ID_N & "" & ",'" & Trim(txtOBS.Text) & "'" & ",'" & Now & "'"

         CONECTA_RETAGUARDA.Execute "EXEC " & SQL
   End If
   If TabOBSENTREGA.State = 1 Then _
      TabOBSENTREGA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ENTREGA"
End Sub
