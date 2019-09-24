VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoqueMP 
   Caption         =   "Controle Mercadoria Fornecedor"
   ClientHeight    =   7815
   ClientLeft      =   1740
   ClientTop       =   450
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ESTOQUEMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Processo"
      TabPicture(0)   =   "ESTOQUEMP.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCNPJCPF"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lstDestino"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstOrigem"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtLOTE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdTransf"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtQtdeTransf"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtDesc2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdConsProd2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCodgProd"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CmdGravar"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CmdInsereTodos"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdRetira"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CmdRetirarTodos"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdInsere"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdLimpar"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbOrigem"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdConsulta"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtNome"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmbOrigemAUX"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Enviar Fornecedor"
      TabPicture(1)   =   "ESTOQUEMP.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkTrans"
      Tab(1).Control(1)=   "lstTransf"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Retorno"
      TabPicture(2)   =   "ESTOQUEMP.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid1"
      Tab(2).Control(1)=   "txtValorDig"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -64680
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbOrigemAUX 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   870
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
         Left            =   6720
         MaxLength       =   100
         TabIndex        =   27
         Top             =   1140
         Width           =   5055
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8640
         Picture         =   "ESTOQUEMP.frx":5C66
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cmbOrigem 
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   660
         Width           =   5055
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   1020
         Left            =   5520
         Picture         =   "ESTOQUEMP.frx":6668
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Confirma os acessos para este usuario."
         Top             =   780
         Width           =   1005
      End
      Begin VB.CommandButton cmdInsere 
         Height          =   615
         Left            =   5520
         Picture         =   "ESTOQUEMP.frx":6E93
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Insere a opção selecionada para este usuario."
         Top             =   1860
         Width           =   1005
      End
      Begin VB.CommandButton CmdRetirarTodos 
         Caption         =   "Retirar &Todos"
         Height          =   1215
         Left            =   5520
         Picture         =   "ESTOQUEMP.frx":72D5
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Retira todos os acessos do sistema para este usuario."
         Top             =   3120
         Width           =   1005
      End
      Begin VB.CommandButton cmdRetira 
         Height          =   615
         Left            =   5520
         Picture         =   "ESTOQUEMP.frx":7717
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Insere a opção selecionada para este usuario."
         Top             =   2460
         Width           =   1005
      End
      Begin VB.CommandButton CmdInsereTodos 
         Caption         =   "&Inserir Todos"
         Height          =   1335
         Left            =   5520
         Picture         =   "ESTOQUEMP.frx":7B59
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Insere todos os acessos para este usuario."
         Top             =   4380
         Width           =   1005
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "&Gravar"
         Height          =   1020
         Left            =   5520
         Picture         =   "ESTOQUEMP.frx":7F9B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Confirma os acessos para este usuario."
         Top             =   5820
         Width           =   1005
      End
      Begin VB.TextBox txtCodgProd 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "Informe o código do produto."
         Top             =   1260
         Width           =   2175
      End
      Begin VB.CommandButton cmdConsProd2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Picture         =   "ESTOQUEMP.frx":8606
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consulta Produto"
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox txtDesc2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   29
         TabIndex        =   8
         Top             =   1740
         Width           =   5055
      End
      Begin VB.TextBox txtQtdeTransf 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         ToolTipText     =   "Quantidade que tem em Estoque"
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton cmdTransf 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11280
         Picture         =   "ESTOQUEMP.frx":9008
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Transf. Realizadas"
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtLOTE 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   10560
         TabIndex        =   6
         ToolTipText     =   "Tecle Enter Para gerar Numero do Lote!"
         Top             =   660
         Width           =   735
      End
      Begin VB.CheckBox chkTrans 
         Caption         =   "Todos"
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstOrigem 
         Height          =   4215
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   2700
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7435
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   14737632
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qtde.Transf."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Produto_ID"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lstDestino 
         Height          =   5295
         Left            =   6720
         TabIndex        =   17
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   1620
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9340
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qtde.Atualiza"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "produto_id"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lstTransf 
         Height          =   6225
         Left            =   -74895
         TabIndex        =   18
         ToolTipText     =   "Duplo Click Impressão"
         Top             =   780
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   10980
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lote"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Seq."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Origem"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Qtde.Origem"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fornecedor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Qtde.À Enviar"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Situação"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Dt.Movimento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Dt.Retorno"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "produto_id"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "fornecedor_id"
            Object.Width           =   0
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   660
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6135
         Left            =   -74940
         TabIndex        =   29
         Top             =   840
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   10821
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Origem"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   5175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   5400
         X2              =   5400
         Y1              =   660
         Y2              =   6900
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Produto MP:"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Qtde.Transf."
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   255
         Left            =   6720
         TabIndex        =   20
         Top             =   420
         Width           =   5055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   255
         Left            =   9960
         TabIndex        =   19
         Top             =   660
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   6600
         X2              =   6600
         Y1              =   660
         Y2              =   6900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   5400
         X2              =   6600
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   5400
         X2              =   6600
         Y1              =   6900
         Y2              =   6900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1270
      ButtonWidth     =   2672
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
            Caption         =   "Voltar"
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpa"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Limpar a Tela"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atualizar"
            Key             =   "at"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confirmar"
            Key             =   "conf"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Impressão"
            Key             =   "imprimir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
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
         Left            =   10200
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   -120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":9A0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":ABA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":BDD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":D072
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":E17D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":F20C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":101C1
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":1143E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":126C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":13B65
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ESTOQUEMP.frx":14DE2
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
      DesignWidth     =   11940
      DesignHeight    =   7815
   End
End
Attribute VB_Name = "frmEstoqueMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conta_Transf  As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name
   txtDtLote = Format(Date, "dd/mm/yyyy")

   HABILITA_TRANSFERENCIA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
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

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "sair"
         Unload Me
      Case "limpa"
         LIMPA_TRANSF
      Case "consultar"
         MONTA_CONSULTA SSTab1.Caption
      Case "at"
         SETA_GRID SSTab1.Caption
      Case "conf"
         EXECUTA_PROCESSO SSTab1.Caption
      Case "imprimir"
         IMPRESSAO_REL
   End Select
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Toolbar1.Buttons(2).Visible = False 'Limpar
   Toolbar1.Buttons(3).Visible = False 'Consultar
   Toolbar1.Buttons(4).Visible = False 'Atualizar
   Toolbar1.Buttons(5).Visible = False 'Confirmar
   Toolbar1.Buttons(6).Visible = False 'Imprimir

   Select Case SSTab1.Tab
      Case 0
         Toolbar1.Buttons(2).Visible = True  'limpar
         txtCodgProd.SetFocus
      Case 1
         Toolbar1.Buttons(3).Visible = True  'Consultar
         Toolbar1.Buttons(4).Visible = True  'Atualizar
         Toolbar1.Buttons(5).Visible = True  'Confirmar
         Toolbar1.Buttons(5).Caption = "Enviar Fornec."
         Toolbar1.Buttons(6).Visible = True  'Imprimir
         SETA_GRID_TRANSITO
      Case 2
         Toolbar1.Buttons(3).Visible = True  'Consultar
         Toolbar1.Buttons(4).Visible = True  'Atualizar
         Toolbar1.Buttons(5).Visible = True  'Confirmar
         Toolbar1.Buttons(5).Caption = "Receber Fornec."
         Toolbar1.Buttons(6).Visible = True  'Imprimir
         SETA_GRID_FLEX
   End Select
End Sub

Private Sub chkTrans_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstTransf.ListItems.Count > 0 Then
      For i = lstTransf.ListItems.Count To 1 Step -1
         If chkTrans.Value = 1 Then
            lstTransf.ListItems(i).Checked = True
            Else: lstTransf.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "chkTrans_Click"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   PESSOA_ID_N = 0
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      txtCNPJCPF.PromptInclude = True
      Call txtCNPJCPF_LostFocus
      txtCNPJCPF.SetFocus
   End If
   CNPJCPF_A = ""
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub cmdLimpar_Click()
   LIMPA_TRANSF
End Sub

Private Sub cmdInsere_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOrigem.SelectedItem.Text) Then
      If Trim(lstOrigem.SelectedItem.Text) <> "" Then
         Conta_Transf = Conta_Transf + 1
         Set item = lstDestino.ListItems.Add(, "seq." & Conta_Transf, Trim(lstOrigem.SelectedItem.Text))
         item.SubItems(1) = "" & Trim(Trim(lstOrigem.SelectedItem.ListSubItems.item(1).Text))
         item.SubItems(2) = "" & Format(Trim(lstOrigem.SelectedItem.ListSubItems.item(2).Text), strFormatacao3Digitos)
         item.SubItems(3) = "" & Trim(Trim(lstOrigem.SelectedItem.ListSubItems.item(3).Text))

         lstOrigem.ListItems.Remove (lstOrigem.SelectedItem.Index)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdInsere_Click"
End Sub

Private Sub cmdRetira_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstDestino.SelectedItem.Text) Then
      If Trim(lstDestino.SelectedItem.Text) <> "" Then
         Conta_Transf = Conta_Transf + 1
         Set item = lstOrigem.ListItems.Add(, "seq." & Conta_Transf, Trim(lstDestino.SelectedItem.Text))
         item.SubItems(1) = "" & Trim(Trim(lstDestino.SelectedItem.ListSubItems.item(1).Text))
         item.SubItems(2) = "" & Format(Trim(lstDestino.SelectedItem.ListSubItems.item(2).Text), strFormatacao3Digitos)
         item.SubItems(3) = "" & Trim(Trim(lstDestino.SelectedItem.ListSubItems.item(3).Text))

         lstDestino.ListItems.Remove (lstDestino.SelectedItem.Index)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdRetira_Click"
End Sub

Sub CmdRetirarTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i As Integer

   If lstDestino.ListItems.Count > 0 Then
      For i = lstDestino.ListItems.Count To 1 Step -1
         If Trim(lstDestino.SelectedItem.Text) <> "" Then

            Conta_Transf = Conta_Transf + 1
            Set item = lstOrigem.ListItems.Add(, "seq." & Conta_Transf, Trim(lstDestino.SelectedItem.Text))
            item.SubItems(1) = "" & Trim(Trim(lstDestino.SelectedItem.ListSubItems.item(1).Text))
            item.SubItems(2) = "" & Format(Trim(lstDestino.SelectedItem.ListSubItems.item(2).Text), strFormatacao3Digitos)
            item.SubItems(3) = "" & Trim(Trim(lstDestino.SelectedItem.ListSubItems.item(3).Text))

            lstDestino.ListItems.Remove (lstDestino.SelectedItem.Index)
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CmdRetirarTodos_Click"
End Sub

Private Sub CmdInsereTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i As Integer

   If lstOrigem.ListItems.Count > 0 Then
      For i = lstOrigem.ListItems.Count To 1 Step -1
         If Trim(lstOrigem.SelectedItem.Text) <> "" Then

            Conta_Transf = Conta_Transf + 1
            Set item = lstDestino.ListItems.Add(, "seq." & Conta_Transf, Trim(lstOrigem.SelectedItem.Text))
            item.SubItems(1) = "" & Trim(Trim(lstOrigem.SelectedItem.ListSubItems.item(1).Text))
            item.SubItems(2) = "" & Format(Trim(lstOrigem.SelectedItem.ListSubItems.item(2).Text), strFormatacao3Digitos)
            item.SubItems(3) = "" & Trim(Trim(lstOrigem.SelectedItem.ListSubItems.item(3).Text))

            lstOrigem.ListItems.Remove (lstOrigem.SelectedItem.Index)
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CmdInsereTodos_Click"
End Sub

Private Sub cmborigem_Click()
On Error Resume Next

   cmbOrigemAUX.ListIndex = cmbOrigem.ListIndex
   txtCodgProd.SetFocus
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
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
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_DADOS_PESSOA

   txtCNPJCPF.PromptInclude = False
   If Len(Trim(txtCNPJCPF.Text)) > 0 Then
      If CInt(Len(Trim(txtCNPJCPF.Text))) = 11 Then
         If Not ValidaCPF(Trim(txtCNPJCPF.Text)) Then
            MsgBox "CPF com DV incorreto !!!"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF = ""
            'ssTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      ElseIf CInt(Len(Trim(txtCNPJCPF.Text))) = 14 Then
         If Not VALIDACNPJ(Trim(txtCNPJCPF.Text)) Then
            MsgBox "CNPJ com DV incorreto !!! "
            txtCNPJCPF.PromptInclude = False
            'ssTab.Tab = 0
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "CNPJ/CPF com DV incorreto !!! "
         txtCNPJCPF = ""
         'ssTab.Tab = 0
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   ElseIf Len(Trim(txtCNPJCPF.Text)) <> 0 Then
       MsgBox "CNPJ/CPF com DV incorreto !!! "
       txtCNPJCPF = ""
       'ssTab.Tab = 0
       TXTCNPJCPF_GotFocus
       txtCNPJCPF.SetFocus
       Exit Sub
   End If
   
   txtCNPJCPF.PromptInclude = False
   CRITERIO_A = Trim(txtCNPJCPF.Text)
   txtCNPJCPF.PromptInclude = False
   
   If Trim(txtCNPJCPF.Text) <> "" Then
      CRITERIO_A = Trim(txtCNPJCPF.Text)

      If Not IsNull(Trim(txtCNPJCPF.Text)) Then
          If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(Trim(txtCNPJCPF.Text)) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.Text = CRITERIO_A
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_LostFocus"
End Sub

Private Sub txtCODGPROD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtCodgProd.ForeColor = vbBlue
   txtDesc2.ForeColor = vbBlue

   If Trim(txtCodgProd.Text) = "" Then _
      Exit Sub

   If KeyAscii = 13 Then
      KeyAscii = 0
      LE_PRODUTO_TRANSF
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCODGPROD"
End Sub

Private Sub cmdTransf_Click()
   frmEstoqueFornecConsulta.Show 1
End Sub

Private Sub cmdConsProd2_Click()
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtCodgProd.Text = SQL3
      txtCodgProd.SetFocus
   End If
   SQL3 = ""
End Sub

Private Sub txtQtdeTransf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtQtdeTransf.Text) <> "" Then
         If IsNumeric(txtQtdeTransf.Text) Then
            SETA_GRID_ORIGEM

            txtCodgProd.Text = ""
            txtQtdeTransf.Text = ""
            txtCodgProd.SetFocus
         End If
      End If
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtdeTransf_KeyPress"
End Sub

Private Sub CmdGravar_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtLOTE.Text) = "" Then _
      txtLOTE.Text = MAX_ID("ESTOQUEFORNEC_ID", "ESTOQUEFORNEC", "", "", "", "")
   If Not IsNumeric(txtLOTE.Text) Then _
      txtLOTE.Text = MAX_ID("ESTOQUEFORNEC_ID", "ESTOQUEFORNEC", "", "", "", "")

   If Not IsNumeric(txtLOTE.Text) Then
      MsgBox "Problemas na geração do lote, verificar..."
      Exit Sub
   End If
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Selecionar Fornecedor para realização do processo."
      Exit Sub
   End If
   txtCNPJCPF.PromptInclude = True
   If Trim(cmbOrigemAUX.Text) = "" Then
      MsgBox "Selecionar Origem para realização do processo."
      Exit Sub
   End If
   If Not IsNumeric(cmbOrigemAUX.Text) Then
      MsgBox "Selecionar Origem para realização do processo."
      Exit Sub
   End If

   INDR_PRI = False

   If lstDestino.ListItems.Count > 0 Then
      For i = lstDestino.ListItems.Count To 1 Step -1
         If Trim(lstDestino.SelectedItem.Text) <> "" Then
            PRODUTO_ID_N = 0 & Trim(lstDestino.ListItems(i).SubItems(3))

            '==========================inserindo na tabela estoque caso não exista ORIGEM
            RODA_AT_ESTOQUE PRODUTO_ID_N, cmbOrigemAUX.Text

            '==========================inserindo na tabela estoque caso não exista DESTINO
            NUMR_SEQ_N = 0 & MAX_ID("SEQ_ID", "ESTOQUEFORNEC", "ESTOQUEFORNEC_ID", txtLOTE.Text, "", "")

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from ESTOQUEFORNEC WITH (NOLOCK)"
            SQL = SQL & " where ESTOQUEFORNEC_ID = " & txtLOTE.Text
            SQL = SQL & " and seq_ID = " & NUMR_SEQ_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabTemp.EOF Then
               SQL = "insert into ESTOQUEFORNEC "
                  SQL = SQL & "(ESTOQUEFORNEC_ID,SEQ_ID,ESTAB_ORIGEM_ID,fornecedor_id,"
                  SQL = SQL & "PRODUTO_ID,qtde_envio,dt_movimento,SITUACAO)"
               SQL = SQL & " values("
                  SQL = SQL & txtLOTE.Text                              'ESTOQUEFORNEC_ID
                  SQL = SQL & "," & NUMR_SEQ_N                                      'SEQ_ID
                  SQL = SQL & "," & cmbOrigemAUX.Text                               'ESTAB_ORIGEM_ID
                  SQL = SQL & "," & FORNEC_ID_N                                     'fornecedor_id
                  SQL = SQL & "," & Trim(lstDestino.ListItems(i).SubItems(3))       'PRODUTO_ID
                  SQL = SQL & "," & tpMOEDA(lstDestino.ListItems(i).SubItems(2))    'qtde
                  SQL = SQL & ",'" & (Now) & "'"                                    'dt_movimento
                  SQL = SQL & ",'T'"                                                'SITUACAO
               SQL = SQL & " )"
               CONECTA_RETAGUARDA.Execute SQL

'VAI MEXER NO ESTOQUE DA TABELA ESTOQUE
               ATUALIZA_ESTOQUE_origem lstDestino.ListItems(i).SubItems(2), Trim(lstDestino.ListItems(i).SubItems(3))

               INDR_PRI = True
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      Next i
      If INDR_PRI = True Then
         MsgBox "Operação realizada com sucesso !!!"
         LIMPA_TRANSF
         txtCodgProd.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CmdGravar_Click"
End Sub

Private Sub lstOrigem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstOrigem, ColumnHeader
End Sub

Private Sub lstOrigem_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(lstOrigem.SelectedItem.Text) Then _
            If Trim(lstOrigem.SelectedItem.Text) <> "" Then _
               lstOrigem.ListItems.Remove (lstOrigem.SelectedItem.Index)
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
End Sub

Sub MOSTRA_PRODUTO_TRANSF()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0 & TabProduto.Fields("produto_id").Value
   txtDesc2.Text = Trim(TabProduto!DESCRICAO)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO_TRANSF"
End Sub

Sub LE_PRODUTO_TRANSF()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodgProd.Text) = "" Then _
      Exit Sub

   PRODUTO_ID_N = 0
   txtQtdeTransf.Enabled = True

   txtCodgProd.Text = UCase(txtCodgProd.Text)

   CODG_PRODUTO_A = Trim(txtCodgProd.Text)

   'LE POR CODIGO DE PRODUTO
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   SQL = SQL & " and tipo_prod = 0 "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_PRODUTO_TRANSF

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQtdeTransf.SetFocus
      Exit Sub
   End If

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS_A = "" & Trim(CODG_PRODUTO_A)
   QTDE_N = 0

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   SQL = SQL & " and tipo_prod = 0 "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_PRODUTO_TRANSF

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQtdeTransf.SetFocus
      Exit Sub
   End If

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS_A = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS_A) = 13 Then
      '2 = produtos "in store" (sempre será 2)
      'C = código do produto (4,5 ou 6 dígitos)
      'T = total a pagar (sempre 6 dígitos)
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      'txtCodgProd.Text = "" & Int(Mid(CODIGO_BARRAS_A, 2, 6))
      'pegando codigo do produto no codigo de barras da etiqueta de balança
      'txtCodgProd.Text = "" & Int(Mid(CODIGO_BARRAS_A, 2, TamanhoCodgProdBarra_N))

txtCodgProd.Text = "" & Int(Mid(CODIGO_BARRAS_A, CasaInicioCodgProdBarra_N, TamanhoCodgProdBarra_N))

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtCodgProd.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      SQL = SQL & " and tipo_prod = 0 "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         MOSTRA_PRODUTO_TRANSF

         If UCase(PESO_VALOR_A) = UCase("valor") Then 'POR VALOR
            VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS_A, 8, TamanhoPesoValorBarra_N) / 100
            QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value) 'sta
            PESO_ITEM_N = QTDE_N
            txtQtdeTransf.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)

            Call txtQtdeTransf_KeyPress(13)

            txtCodgProd.SetFocus
            Else
               QTDE_N = 0 & Int(Mid(CODIGO_BARRAS_A, 8, 5))           'gramas
      
               If QTDE_N > 0 Then _
                  QTDE_N = Format(QTDE_N / 1000, strFormatacao3Digitos)
      
               PESO_ITEM_N = QTDE_N
               'txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
      
               txtQtdeTransf.Text = PESO_ITEM_N
      
               Call txtQtdeTransf_KeyPress(13)
      
               txtCodgProd.SetFocus
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close
txtQtdeTransf.SetFocus
         Exit Sub
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(CODIGO_BARRAS_A) = 12 Then
      'lendo codigo barras ultralav
      '100004360813
      '1-1 = masculino ou feminino
      '2-7 = código do produto
      '8-9 = numeração tamanho produto
      '10-11 = mes
      '12-13 = ano

      txtCodgProd.Text = "" & Mid(CODIGO_BARRAS_A, 1, 6)
      SqL2 = "" & Mid(CODIGO_BARRAS_A, 7, 2)

      SQL = "select * from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where referencia = '" & Trim(txtCodgProd.Text) & "'"
      SQL = SQL & " and RIGHT(descricao,2) = '" & Trim(SqL2) & "'"
      SQL = SQL & " and situacao <> 'C' "
      SQL = SQL & " and tipo_prod = 0 "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         MOSTRA_PRODUTO_TRANSF

         If TabProduto.State = 1 Then _
            TabProduto.Close

         txtQtdeTransf.Text = 1

         Call txtQtdeTransf_KeyPress(13)

         'txtCodgProd.SetFocus
         txtQtdeTransf.SetFocus
         Exit Sub
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MsgBox "Produto não cadastrado."
   txtCodgProd.SetFocus
   txtCodgProd.Text = ""
   txtQtdeTransf.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_PRODUTO_TRANSF"
End Sub

Sub HABILITA_TRANSFERENCIA()
'On Error GoTo ERRO_TRATA

   Toolbar1.Buttons(1).Visible = True
   Toolbar1.Buttons(2).Visible = False
   Toolbar1.Buttons(3).Visible = False
   Toolbar1.Buttons(4).Visible = False
   Toolbar1.Buttons(5).Visible = False

   cmbOrigem.Enabled = False

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select count(estabelecimento_id) from ESTABELECIMENTO WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields(0).Value) Then
         If TabTemp.Fields(0).Value > 1 Then
            If TabTemp.State = 1 Then _
               TabTemp.Close
            SQL = "select estabelecimento_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
            SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbOrigem.Enabled = True
               cmbOrigem.Text = Trim(TabTemp.Fields("descricao").Value)
               cmbOrigemAUX.Text = TabTemp.Fields("estabelecimento_id").Value
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "HABILITA_TRANSFERENCIA"
End Sub

Sub LIMPA_TRANSF()
'On Error GoTo ERRO_TRATA

   lstOrigem.ListItems.Clear
   lstDestino.ListItems.Clear
   txtCodgProd.Text = ""
   txtQtdeTransf.Text = ""
   txtDesc2.Text = ""
   Conta_Transf = 0
   PRODUTO_ID_N = -1
   txtLOTE.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CmdGravar_Click"
End Sub

Sub MOSTRA_DADOS_PESSOA()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   FORNEC_ID_N = 0
   txtNome.Text = ""

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If IsNumeric(txtCNPJCPF.Text) Then
         Dim TabPessoa     As New ADODB.Recordset

         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         SQL = "select  PESSOA.PESSOA_ID, PESSOA.CNPJCPF, FORNECEDOR.FORNECEDOR_ID, FORNECEDOR.STATUS, PESSOA.DESCRICAO"
         SQL = SQL & " from PESSOA  WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN FORNECEDOR  WITH (NOLOCK)"
         SQL = SQL & "ON PESSOA.PESSOA_ID = FORNECEDOR.PESSOA_ID"

         SQL = SQL & " where CNPJCPF = '" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & " and status = 'A' "
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then
            PESSOA_ID_N = 0 & TabPessoa.Fields("pessoa_id").Value
            FORNEC_ID_N = 0 & TabPessoa.Fields("fornecedor_id").Value
            txtNome.Text = "" & Trim(TabPessoa.Fields("descricao").Value)
         End If
         If TabPessoa.State = 1 Then _
            TabPessoa.Close
         Else: Exit Sub
      End If
      Else: Exit Sub
   End If
   If PESSOA_ID_N <= 0 Then _
      Exit Sub

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PESSOA"
End Sub

Sub SETA_GRID_ORIGEM()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodgProd.Text) <> "" And Trim(txtQtdeTransf.Text) <> "" Then
      If IsNumeric(txtQtdeTransf.Text) Then
         Conta_Transf = Conta_Transf + 1
         Set item = lstOrigem.ListItems.Add(, "seq." & Conta_Transf, Trim(txtCodgProd.Text))
         item.SubItems(1) = "" & Trim(txtDesc2.Text)
         item.SubItems(2) = "" & Format(txtQtdeTransf.Text, strFormatacao3Digitos)
         item.SubItems(3) = "" & Trim(PRODUTO_ID_N)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ORIGEM"
End Sub

Sub ATUALIZA_ESTOQUE_DESTINO(QTDE_N As Double, PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   SQL = "update ESTOQUE set "
   SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(QTDE_N)
   SQL = SQL & " where produto_id = " & Trim(PROD_ID_N)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_ESTOQUE_DESTINO"
End Sub

Sub ATUALIZA_ESTOQUE_origem(QTDE_N As Double, PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   SQL = "update ESTOQUE set "
   SQL = SQL & " qtde_estoque = qtde_estoque - " & tpMOEDA(QTDE_N)
   SQL = SQL & " where produto_id = " & PROD_ID_N
   SQL = SQL & " and estabelecimento_id = " & cmbOrigemAUX.Text
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_ESTOQUE_ORIGEM"
End Sub

Sub SETA_GRID(INDICE_A As String)
   If Trim(INDICE_A) = "Enviar Fornecedor" Then
      SETA_GRID_TRANSITO
      Else: SETA_GRID_FLEX
   End If
   CRITERIO_A = ""
End Sub

Sub SETA_GRID_TRANSITO()
'On Error GoTo ERRO_TRATA

   CONT_N = 0
   lstTransf.Visible = False
   lstTransf.ListItems.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, ESTOQUEFORNEC.*, pessoa_id"
   SQL = SQL & " from ESTOQUEFORNEC WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON ESTOQUEFORNEC.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON ESTOQUEFORNEC.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

   SQL = SQL & " where ESTAB_ORIGEM_ID = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and ESTOQUEFORNEC.situacao = 'T' "

   SQL = SQL & " order by ESTOQUEFORNEC_ID "

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      CONT_N = CONT_N + 1
      PRODUTO_ID_N = 0 & TabConsulta.Fields("produto_id").Value
      QTDE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      Set item = lstTransf.ListItems.Add(, "seq." & CONT_N, TabConsulta.Fields("ESTOQUEFORNEC_id").Value)
      item.SubItems(1) = "" & Trim(TabConsulta.Fields("seq_id").Value)
      item.SubItems(2) = "" & Trim(TabConsulta.Fields("codg_produto").Value)
      item.SubItems(3) = "" & Trim(TabConsulta.Fields("descricao").Value)
      item.SubItems(4) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estab_origem_id").Value)
      item.SubItems(5) = "" & Format(QTDE_N, strFormatacao3Digitos)

      item.SubItems(6) = "" & TRAZ_NOME_FORNECEDOR(TabConsulta.Fields("fornecedor_id").Value, TabConsulta.Fields("PESSOA_id").Value)
      item.SubItems(7) = "" & Format(TabConsulta.Fields("qtde_envio").Value, strFormatacao3Digitos)

      item.SubItems(8) = ""
      SqL2 = ""
      If Not IsNull(TabConsulta.Fields("SITUACAO").Value) Then
         If Trim(TabConsulta.Fields("SITUACAO").Value) = "A" Then _
            SqL2 = "Aberto"
         If Trim(TabConsulta.Fields("SITUACAO").Value) = "T" Then _
            SqL2 = "Transito"
         If Trim(TabConsulta.Fields("SITUACAO").Value) = "F" Then _
            SqL2 = "Fechado"
      End If
      item.SubItems(8) = "" & Trim(SqL2)

item.SubItems(9) = "" & TabConsulta.Fields("dt_movimento").Value

      item.SubItems(10) = "" & TabConsulta.Fields("dt_retorno").Value
      item.SubItems(11) = "" & TabConsulta.Fields("produto_id").Value
      item.SubItems(12) = "" & TabConsulta.Fields("fornecedor_id").Value

      item.ForeColor = vbBlue
      item.ListSubItems(1).ForeColor = vbBlue
      item.ListSubItems(2).ForeColor = vbRed
      item.ListSubItems(3).ForeColor = vbRed
      item.ListSubItems(4).ForeColor = vbRed
      item.ListSubItems(5).ForeColor = vbRed
      item.ListSubItems(6).ForeColor = vbRed
      item.ListSubItems(7).ForeColor = vbRed
      item.ListSubItems(8).ForeColor = vbRed
      item.ListSubItems(9).ForeColor = vbRed

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   lstTransf.Visible = True
   PRODUTO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_TRANSITO"
End Sub

Sub MONTA_CONSULTA(INDICE_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(INDICE_A) = "" Then _
      Exit Sub

   CRITERIO_A = Trim(INDICE_A)

   frmEstoqueFornecConsulta.Show 1

   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub

Sub EXECUTA_PROCESSO(INDICE_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(INDICE_A) = "Enviar Fornecedor" Then
      If lstTransf.ListItems.Count > 0 Then _
         RECEBE_SELECIONADOS_TRANSITO
      SETA_GRID_TRANSITO
   End If
   If Trim(INDICE_A) = "Retorno" Then
      'If lstRetorno.ListItems.Count > 0 Then _
         RECEBE_SELECIONADOS_RETORNO
      RECEBE_SELECIONADOS_RETORNO
      SETA_GRID_FLEX
   End If
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXECUTA_PROCESSO"
End Sub

Sub RECEBE_SELECIONADOS_TRANSITO()
'On Error GoTo ERRO_TRATA

   Dim i             As Integer
   Dim INDR_TRANSF   As Boolean

   INDR_PRI = True
   INDR_TRANSF = False

   If lstTransf.ListItems.Count > 0 Then
      For i = lstTransf.ListItems.Count To 1 Step -1
         If lstTransf.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               INDR_PRI = False
               Msg = "Confirma Transferência de mercadoria(s) Selecionada(s) para Fornecedore(s) ? "
               Style = vbYesNo + 32
               Title = "Atenção."
               Help = "DEMO.HLP"
               Ctxt = 1000
               RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
               If RESPOSTA = vbYes Then
                  INDR_TRANSF = True
                  Else
                     INDR_TRANSF = False
                     Exit Sub
               End If
            End If
            If INDR_TRANSF = True Then
               CONT_N = 1
               PRODUTO_ID_N = 0 & lstTransf.ListItems(i).SubItems(11)
               SEQ_ID_N = 0 & lstTransf.ListItems(i).SubItems(1)
               NUMR_ID_N = 0 & lstTransf.SelectedItem.Text

               '==========================inserindo na tabela estoque caso não exista
               RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N

               'ATUALIZA_ESTOQUE_DESTINO lstTransf.ListItems(i).SubItems(6), PRODUTO_ID_N

               SQL = "update ESTOQUEFORNEC set "
               SQL = SQL & " situacao = 'R'"  'transferida
               'SQL = SQL & ", dt_retorno = '" & Now & "'"

               SQL = SQL & " where ESTOQUEFORNEC_id = " & NUMR_ID_N
               SQL = SQL & " and seq_id = " & SEQ_ID_N
               SQL = SQL & " and situacao = 'T' "
               SQL = SQL & " and produto_id = " & PRODUTO_ID_N
               SQL = SQL & " and fornecedor_id = " & lstTransf.ListItems(i).SubItems(12)
               SQL = SQL & " and estab_origem_id = " & ESTABELECIMENTO_ID_N
'MsgBox SQL
               CONECTA_RETAGUARDA.Execute SQL

               PRODUTO_ID_N = 0
               SEQ_ID_N = 0
               NUMR_ID_N = 0
            End If
            DoEvents
         End If
      Next i
      If INDR_PRI = False Then _
         MsgBox "Processo realizado com sucesso !!!"
   End If
   PRODUTO_ID_N = 0
   SEQ_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RECEBE_SELECIONADOS_TRANSITO"
End Sub

Sub RECEBE_SELECIONADOS_RETORNO()
'On Error GoTo ERRO_TRATA

   Dim i             As Integer
   Dim INDR_TRANSF   As Boolean

   INDR_PRI = True
   INDR_TRANSF = False

   For i = 1 To MSFlexGrid1.Rows - 1
      If Trim(MSFlexGrid1.TextMatrix(i, 8)) <> "" Then
         If IsNumeric(MSFlexGrid1.TextMatrix(i, 8)) Then
            If INDR_PRI = True Then
               INDR_PRI = False
               Msg = "Confirma recebimento de mercadoria(s) Selecionada(s) de Fornecedor(s) ? "
               Style = vbYesNo + 32
               Title = "Atenção."
               Help = "DEMO.HLP"
               Ctxt = 1000
               RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
               If RESPOSTA = vbYes Then
                  INDR_TRANSF = True
                  Else
                     INDR_TRANSF = False
                     Exit Sub
               End If
            End If
            If INDR_TRANSF = True Then
               PRODUTO_ID_N = 0 & Trim(MSFlexGrid1.TextMatrix(i, 11))
               SEQ_ID_N = 0 & Trim(MSFlexGrid1.TextMatrix(i, 10))
               NUMR_ID_N = 0 & Trim(MSFlexGrid1.TextMatrix(i, 0))
               FORNEC_ID_N = 0 & Trim(MSFlexGrid1.TextMatrix(i, 13))
               QTDE_N = 0 & Trim(MSFlexGrid1.TextMatrix(i, 8))

               '==========================inserindo na tabela estoque caso não exista
               RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N

               ATUALIZA_ESTOQUE_DESTINO QTDE_N, PRODUTO_ID_N

               SQL = "update ESTOQUEFORNEC set "
               SQL = SQL & " situacao = 'F'"  'recebido do fornecedor
               SQL = SQL & ", dt_retorno = '" & Now & "'"
               SQL = SQL & ", qtde_retorno = " & tpMOEDA(QTDE_N)

               SQL = SQL & " where ESTOQUEFORNEC_id = " & NUMR_ID_N
               SQL = SQL & " and seq_id = " & SEQ_ID_N
               SQL = SQL & " and situacao = 'R' "
               SQL = SQL & " and produto_id = " & PRODUTO_ID_N
               SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
               SQL = SQL & " and estab_origem_id = " & ESTABELECIMENTO_ID_N

               CONECTA_RETAGUARDA.Execute SQL

               PRODUTO_ID_N = 0
               SEQ_ID_N = 0
               NUMR_ID_N = 0
               FORNEC_ID_N = 0
            End If
            DoEvents
         End If
      End If
   Next i

   If INDR_PRI = False And INDR_TRANSF = True Then _
      MsgBox "Processo realizado com sucesso !!!"

   PRODUTO_ID_N = 0
   SEQ_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RECEBE_SELECIONADOS_RETORNO"
End Sub

Sub IMPRESSAO_REL()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstTransf.SelectedItem.Text) Then
      FORMULA_REL = "{ESTOQUEFORNEC.ESTOQUEFORNEC_ID} = " & lstTransf.SelectedItem.Text

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "estoque_transf.rpt"
      frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRESSAO_REL"
End Sub
'==========================
Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   OcultarControles

   ExibirCelula

'   txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 8)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

   Select Case KeyAscii
      Case vbKeyReturn  ' Editar ao teclar ENTER
         KeyAscii = 0
         ExibirCelula
      Case vbKeyEscape  ' Cancelar ao pressionar ESC
         KeyAscii = 0
         AtribuiValorCelula
      Case 32 To 255    ' Editar ao pressinar qualquer tecla
         ExibirCelula
         With txtValorDig
            If .Visible Then
             .Text = Chr$(KeyAscii)
             .SelStart = Len(.Text) + 1
           End If
         End With
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyPress"
End Sub

Private Sub ExibirCelula()
'On Error GoTo ERRO_TRATA

   Static OK As Boolean

   If MSFlexGrid1.Col = 8 Then
      ' Se for celula fixa , sair
      If MSFlexGrid1.Col <= MSFlexGrid1.FixedCols - 1 Or MSFlexGrid1.Row <= MSFlexGrid1.FixedRows - 1 Then _
         Exit Sub
   
      If OK Then _
         Exit Sub

      OK = True

      OcultarControles

      LastRow = MSFlexGrid1.Row
      LastCol = MSFlexGrid1.Col

      Select Case LastCol
         Case Else
            txtValorDig.Move MSFlexGrid1.CellLeft - Screen.TwipsPerPixelX, MSFlexGrid1.CellTop + MSFlexGrid1.Top - Screen.TwipsPerPixelY, MSFlexGrid1.CellWidth + Screen.TwipsPerPixelX * 2, MSFlexGrid1.CellHeight + Screen.TwipsPerPixelY * 2
            txtValorDig.Text = MSFlexGrid1.Text

            If Len(MSFlexGrid1.Text) = 0 Then _
               If LastRow > 1 Then _
                  txtValorDig.Text = MSFlexGrid1.TextMatrix(LastRow - 1, LastCol)

            txtValorDig.Visible = True

            If txtValorDig.Visible Then
               txtValorDig.ZOrder
               txtValorDig.SetFocus
            End If
      End Select
   
      ControlVisible = True

      OK = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ExibirCelula"
End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorDig.SelStart = 0
   txtValorDig.SelLength = Len(txtValorDig)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_GotFocus"
End Sub

Private Sub txtValorDig_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         OcultarControles
         MSFlexGrid1.SetFocus
      Case vbKeyUp
         OcultarControles
         'move para a cima celula.
         With MSFlexGrid1
            If .Row > 1 Then
                .Row = .Row - 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
      Case vbKeyDown
         OcultarControles
         With MSFlexGrid1
             If .Row + 1 < .Rows Then
                .Row = .Row + 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyDown"
End Sub

Private Sub txtValorDig_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   ' ao pressionar ENTER aceitar a entrada de dados
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      If LastCol > 8 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      If Trim(txtValorDig.Text) <> "" Then _
         If IsNumeric(txtValorDig.Text) Then _
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = "" & Format(txtValorDig.Text, strFormatacao3Digitos)

      'AtribuiValorCelula
      'ProximaCelula
      OcultarControles

      'Qtde_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 8)

      'SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)

      'PRODUTO_ID_N = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 12))

      With MSFlexGrid1
         If .Row + 1 < .Rows Then
            .Row = .Row + 1
            '.Col = 0
            Else
               .Row = 1
               '.Col = 0
         End If
      End With
      txtValorDig.Text = ""
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            'ControlVisible = False
            Else
               If KeyAscii = 8 Or KeyAscii = 44 Then
                  Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyPress"
End Sub

Private Sub ProximaCelula()
'On Error GoTo ERRO_TRATA

   If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      Else
         MSFlexGrid1.Col = 1
         If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
             MSFlexGrid1.Row = MSFlexGrid1.Row + 1
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ProximaCelula"
End Sub

Private Sub AtribuiValorCelula()
'On Error GoTo ERRO_TRATA

   Dim texto As String

   ' atribuir o texto anterior a celula
   Select Case LastCol
      Case 8
         texto = txtValorDig.Text
         MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
      Case Else
         'texto = txtValorDig.Text
         'MSFlexGrid1.TextMatrix(LastRow, LastCol) = texto
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   'Ocultar o controle textbox
   txtValorDig.Visible = False
   'Toolbar1.Buttons(9).Visible = False
   'Toolbar1.Buttons(8).Visible = False

   'If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
   '   Toolbar1.Buttons(9).Visible = True
   '   Toolbar1.Buttons(8).Visible = True
   'End If
   'If MULT_EMPRESA_B = True Then _
      Toolbar1.Buttons(9).Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Private Sub SETA_GRID_FLEX()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   CONT_N = 0

   MSFlexGrid1.Clear
   MSFlexGrid1.Visible = False
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select ESTOQUEFORNEC.ESTOQUEFORNEC_ID as Lote, PRODUTO.CODG_PRODUTO as Código, PRODUTO.DESCRICAO as Descrição, "
   SQL = SQL & " ESTOQUEFORNEC.ESTAB_ORIGEM_ID as Origem, ESTOQUEFORNEC.DT_MOVIMENTO as DtMovimento, ESTOQUEFORNEC.QTDE_ENVIO as QtdeEnviado,"
   SQL = SQL & " ESTOQUEFORNEC.FORNECEDOR_ID as Fornecedor, ESTOQUEFORNEC.DT_RETORNO as DtRetorno, ESTOQUEFORNEC.QTDE_RETORNO as QtdeRetorno,"
   SQL = SQL & " ESTOQUEFORNEC.SITUACAO as Situação,"
   SQL = SQL & " ESTOQUEFORNEC.SEQ_ID , ESTOQUEFORNEC.PRODUTO_ID, FORNECEDOR.Pessoa_Id,ESTOQUEFORNEC.FORNECEDOR_ID"

   SQL = SQL & " from ESTOQUEFORNEC WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUEFORNEC.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUEFORNEC.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

   SQL = SQL & " where ESTAB_ORIGEM_ID = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and ESTOQUEFORNEC.situacao = 'R' "

   SQL = SQL & " order by ESTOQUEFORNEC_ID "

   TabTemp.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabTemp.EOF Then
      INDR_PEDIDO_VALIDO = True
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabTemp.Fields.Count

      ReDim largura_coluna(0 To TabTemp.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabTemp.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabTemp.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabTemp.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabTemp.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         FORNEC_ID_N = 0 & TabTemp.Fields("Fornecedor").Value
         PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value

         For Coluna = 0 To TabTemp.Fields.Count - 1
            If Coluna = 5 Or Coluna = 8 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabTemp.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna = 3 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TRAZ_ESTABELECIMENTO(TabTemp.Fields(Coluna).Value))
                     Else
                        If Coluna = 6 Then
                           MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TRAZ_NOME_FORNECEDOR(FORNEC_ID_N, PESSOA_ID_N))
                           Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabTemp.Fields(Coluna).Value)
                        End If
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabTemp.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabTemp.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabTemp.MoveNext
         Linha = Linha + 1
      Loop

      'define a largura das colunas do grid
      For Coluna = 0 To MSFlexGrid1.Cols - 1
         MSFlexGrid1.ColWidth(Coluna) = largura_coluna(Coluna) + 240
      Next Coluna

      MSFlexGrid1.ColWidth(0) = 0
      MSFlexGrid1.Refresh

      MSFlexGrid1.BackColor = vbWhite
      MSFlexGrid1.ForeColor = vbBlue

'Lote = 0
      MSFlexGrid1.ColWidth(0) = 1000
      MSFlexGrid1.ColAlignment(0) = 0

'Código Produto = 1
      MSFlexGrid1.ColWidth(1) = 1200
      MSFlexGrid1.ColAlignment(1) = 0

'Descrição Produto = 2
      MSFlexGrid1.ColWidth(2) = 5000
      MSFlexGrid1.ColAlignment(2) = 0

'Origem = 3
      MSFlexGrid1.ColWidth(3) = 2200
      MSFlexGrid1.ColAlignment(3) = 0

'DtMovimento = 4
      MSFlexGrid1.ColWidth(4) = 1700
      MSFlexGrid1.ColAlignment(4) = 0

'Qtde_Envio = 5
      MSFlexGrid1.ColWidth(5) = 2000
      MSFlexGrid1.ColAlignment(5) = 7

'Fornecedor = 6
      MSFlexGrid1.ColWidth(6) = 2200
      MSFlexGrid1.ColAlignment(6) = 0

'DtRetorno = 7
      MSFlexGrid1.ColWidth(7) = 1700
      MSFlexGrid1.ColAlignment(7) = 0

'Qtde_Retorno = 8
      MSFlexGrid1.ColWidth(8) = 2000
      MSFlexGrid1.ColAlignment(8) = 7

'Situação = 9
      MSFlexGrid1.ColWidth(9) = 100
      MSFlexGrid1.ColAlignment(9) = 0

'Seq_id
      MSFlexGrid1.ColWidth(10) = 1
      MSFlexGrid1.ColAlignment(10) = 0

'Produto_id
      MSFlexGrid1.ColWidth(11) = 1
      MSFlexGrid1.ColAlignment(11) = 0

'Pessoa_id
      MSFlexGrid1.ColWidth(12) = 1
      MSFlexGrid1.ColAlignment(12) = 0

'fornecedor_id
      MSFlexGrid1.ColWidth(13) = 1
      MSFlexGrid1.ColAlignment(13) = 0
   End If
   ' fecha o recordset e a conexao
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MSFlexGrid1.Visible = True
   FORNEC_ID_N = 0
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_FLEX"
End Sub
