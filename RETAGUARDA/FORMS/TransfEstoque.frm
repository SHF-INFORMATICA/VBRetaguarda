VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransfEstoque 
   Caption         =   "Transferência Estoque"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TransfEstoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Processo"
      TabPicture(0)   =   "TransfEstoque.frx":5C12
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Line1(1)"
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(6)=   "Line1(0)"
      Tab(0).Control(7)=   "Line2(0)"
      Tab(0).Control(8)=   "Line2(1)"
      Tab(0).Control(9)=   "lstDestino"
      Tab(0).Control(10)=   "lstOrigem"
      Tab(0).Control(11)=   "cmbOrigem"
      Tab(0).Control(12)=   "cmdLimpar"
      Tab(0).Control(13)=   "cmdInsere"
      Tab(0).Control(14)=   "CmdRetirarTodos"
      Tab(0).Control(15)=   "cmdRetira"
      Tab(0).Control(16)=   "CmdInsereTodos"
      Tab(0).Control(17)=   "CmdGravar"
      Tab(0).Control(18)=   "txtCodgProd"
      Tab(0).Control(19)=   "cmdConsProd2"
      Tab(0).Control(20)=   "txtDesc2"
      Tab(0).Control(21)=   "txtQtdeTransf"
      Tab(0).Control(22)=   "cmbDestino"
      Tab(0).Control(23)=   "cmbDestinoAUX"
      Tab(0).Control(24)=   "cmbOrigemAUX"
      Tab(0).Control(25)=   "cmdTransf"
      Tab(0).Control(26)=   "txtTransf_ID"
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Transito"
      TabPicture(1)   =   "TransfEstoque.frx":5C2E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line2(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line2(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstTransf"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdAt"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdTransfere"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkTodos"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdCons2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdCons2 
         Caption         =   "&Consultar"
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
         Left            =   10560
         Picture         =   "TransfEstoque.frx":5C4A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   480
         Width           =   1140
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdTransfere 
         Caption         =   "&Confirmar"
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
         Left            =   9360
         Picture         =   "TransfEstoque.frx":6F8C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         Width           =   1140
      End
      Begin VB.CommandButton cmdAt 
         Caption         =   "&Atualizar"
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
         Left            =   8145
         Picture         =   "TransfEstoque.frx":8107
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox txtTransf_ID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   -67560
         TabIndex        =   4
         ToolTipText     =   "Tecle Enter Para gerar Numero do Lote!"
         Top             =   1200
         Width           =   735
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
         Height          =   350
         Left            =   -66840
         Picture         =   "TransfEstoque.frx":945A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Transf. Realizadas"
         Top             =   1200
         Width           =   405
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
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox cmbDestinoAUX 
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
         Left            =   -68280
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox cmbDestino 
         Height          =   360
         Left            =   -68280
         TabIndex        =   3
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txtQtdeTransf 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -73680
         TabIndex        =   1
         ToolTipText     =   "Quantidade que tem em Estoque"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtDesc2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         MaxLength       =   29
         TabIndex        =   17
         Top             =   1800
         Width           =   5055
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
         Height          =   350
         Left            =   -71760
         Picture         =   "TransfEstoque.frx":9E5C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Consulta Produto"
         Top             =   1320
         Width           =   405
      End
      Begin VB.TextBox txtCodgProd 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   -74040
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "Informe o código do produto."
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "&Gravar"
         Height          =   1020
         Left            =   -69480
         Picture         =   "TransfEstoque.frx":A85E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Confirma os acessos para este usuario."
         Top             =   5640
         Width           =   1005
      End
      Begin VB.CommandButton CmdInsereTodos 
         Caption         =   "&Inserir Todos"
         Height          =   1335
         Left            =   -69480
         Picture         =   "TransfEstoque.frx":AEC9
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Insere todos os acessos para este usuario."
         Top             =   4200
         Width           =   1005
      End
      Begin VB.CommandButton cmdRetira 
         Height          =   615
         Left            =   -69480
         Picture         =   "TransfEstoque.frx":B30B
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Insere a opção selecionada para este usuario."
         Top             =   2280
         Width           =   1005
      End
      Begin VB.CommandButton CmdRetirarTodos 
         Caption         =   "Retirar &Todos"
         Height          =   1215
         Left            =   -69480
         Picture         =   "TransfEstoque.frx":B74D
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Retira todos os acessos do sistema para este usuario."
         Top             =   2940
         Width           =   1005
      End
      Begin VB.CommandButton cmdInsere 
         Height          =   615
         Left            =   -69480
         Picture         =   "TransfEstoque.frx":BB8F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Insere a opção selecionada para este usuario."
         Top             =   1680
         Width           =   1005
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   1020
         Left            =   -69480
         Picture         =   "TransfEstoque.frx":BFD1
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Confirma os acessos para este usuario."
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox cmbOrigem 
         Height          =   360
         Left            =   -74880
         TabIndex        =   2
         Top             =   720
         Width           =   5055
      End
      Begin MSComctlLib.ListView lstOrigem 
         Height          =   4185
         Left            =   -74880
         TabIndex        =   19
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   2760
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   7382
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
         Height          =   5265
         Left            =   -68280
         TabIndex        =   25
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   1680
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   9287
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   12648447
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
         Height          =   5385
         Left            =   105
         TabIndex        =   26
         ToolTipText     =   "Duplo Click Impressão"
         Top             =   1560
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   9499
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
         NumItems        =   12
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
            SubItemIndex    =   5
            Text            =   "Destino"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Qtde.Transf."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Qtde.Origem"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Situação"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Dt.Transferência"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Dt.Entrada"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "produto_id"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   3
         X1              =   11760
         X2              =   11760
         Y1              =   360
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   2
         X1              =   8040
         X2              =   8040
         Y1              =   360
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   3
         X1              =   8040
         X2              =   11760
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   2
         X1              =   8040
         X2              =   11760
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   1
         X1              =   -69600
         X2              =   -68400
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   0
         X1              =   -69600
         X2              =   -68400
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   0
         X1              =   -68400
         X2              =   -68400
         Y1              =   480
         Y2              =   6720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   240
         Left            =   -68160
         TabIndex        =   24
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   240
         Left            =   -68280
         TabIndex        =   20
         Top             =   480
         Width           =   5025
      End
      Begin VB.Label Label7 
         Caption         =   "Qtde.Transf."
         Height          =   240
         Left            =   -74880
         TabIndex        =   18
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Produto:"
         Height          =   240
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   1
         X1              =   -69600
         X2              =   -69600
         Y1              =   480
         Y2              =   6720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Origem"
         Height          =   240
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   5145
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "clear"
            Object.ToolTipText     =   "Limpar a Tela"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
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
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":C7FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":D996
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":EBC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":FE64
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":10F6F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":11FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":12FB3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":14230
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":154B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TransfEstoque.frx":16957
               Key             =   ""
            EndProperty
         EndProperty
      End
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
         TabIndex        =   6
         Top             =   240
         Width           =   1695
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
      DesignWidth     =   11895
      DesignHeight    =   7830
   End
End
Attribute VB_Name = "frmTransfEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conta_Transf As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.name
   txtDtLote = Format(Date, "dd/mm/yyyy")

   HABILITA_TRANSFERENCIA

   SSTab1.TabVisible(0) = False
   If TRAZ_TIPO_USUARIO = 4 Or TRAZ_TIPO_USUARIO = 5 Then _
      SSTab1.TabVisible(0) = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Unload"
End Sub

Private Sub lstTransf_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstTransf.SelectedItem.Text) Then
      FORMULA_REL = "{ESTOQUEtransf.transf_id} = " & lstTransf.SelectedItem.Text

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "estoque_transf.rpt"
      frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "lstTransf_DblClick"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      Toolbar1.Buttons(2).Visible = True
      Toolbar1.Buttons(3).Visible = True
      txtCodgProd.SetFocus
      Else
         If SSTab1.Tab = 1 Then
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False

            SETA_GRID_TRANSITO
         End If
   End If
End Sub

Private Sub cmdAt_Click()
   SETA_GRID_TRANSITO
End Sub

Private Sub cmdTransfere_Click()
   TRANSF_SELECIONADOS
   SETA_GRID_TRANSITO
End Sub

Private Sub chkTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstTransf.ListItems.Count > 0 Then
      For i = lstTransf.ListItems.Count To 1 Step -1
         If chkTodos.Value = 1 Then
            lstTransf.ListItems(i).Checked = True
            Else: lstTransf.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "chkTodos_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         frmTransfConsulta.Show 1
      Case "sair"
         Unload Me
      Case "clear"
         LIMPA_TRANSF
   End Select
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdCons2_Click()
   frmTransfConsulta.Show 1
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
   TRATA_ERROS Err.description, Me.name, "cmdInsere_Click"
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
   TRATA_ERROS Err.description, Me.name, "cmdRetira_Click"
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
   TRATA_ERROS Err.description, Me.name, "CmdRetirarTodos_Click"
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
   TRATA_ERROS Err.description, Me.name, "CmdInsereTodos_Click"
End Sub

Private Sub cmborigem_Click()
On Error Resume Next

   cmbOrigemAUX.ListIndex = cmbOrigem.ListIndex
   txtCodgProd.SetFocus
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
   TRATA_ERROS Err.description, Me.name, "txtCODGPROD"
End Sub

Private Sub cmbDESTINO_Click()
On Error Resume Next

   cmbDestinoAUX.ListIndex = cmbDestino.ListIndex

   txtTransf_ID.Text = MAX_ID("TRANSF_ID", "ESTOQUETRANSF", "", "", "", "")
End Sub

Private Sub cmdTransf_Click()
   frmTransfConsulta.Show 1
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
   TRATA_ERROS Err.description, Me.name, "txtQtdeTransf_KeyPress"
End Sub

Private Sub CmdGravar_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtTransf_ID.Text) = "" Then _
      txtTransf_ID.Text = MAX_ID("TRANSF_ID", "ESTOQUETRANSF", "", "", "", "")
   If Not IsNumeric(txtTransf_ID.Text) Then _
      txtTransf_ID.Text = MAX_ID("TRANSF_ID", "ESTOQUETRANSF", "", "", "", "")

   If Not IsNumeric(txtTransf_ID.Text) Then
      MsgBox "Problemas na geração do lote de transferência, verificar..."
      Exit Sub
   End If
   If Trim(cmbDestinoAUX.Text) = "" Then
      MsgBox "Selecionar destino para realização do processo."
      Exit Sub
   End If
   If Not IsNumeric(cmbDestinoAUX.Text) Then
      MsgBox "Selecionar destino para realização do processo."
      Exit Sub
   End If
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
            RODA_AT_ESTOQUE PRODUTO_ID_N, cmbDestinoAUX.Text

            NUMR_SEQ_N = 0 & MAX_ID("SEQ_ID", "ESTOQUETRANSF", "TRANSF_ID", txtTransf_ID.Text, "", "")

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from ESTOQUETRANSF WITH (NOLOCK)"
            SQL = SQL & " where TRANSF_ID = " & txtTransf_ID.Text
            SQL = SQL & " and seq_ID = " & NUMR_SEQ_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabTemp.EOF Then
               SQL = "insert into ESTOQUETRANSF "
                  SQL = SQL & "(TRANSF_ID,SEQ_ID,EMPRESA_ID,ESTAB_ORIGEM_ID,ESTAB_DESTINO_ID,"
                  SQL = SQL & "PRODUTO_ID,QTDE_TRANSF,DT_TRANSF,SITUACAO,"
                  SQL = SQL & " ESTAB_ORIGEM_DESC,ESTAB_DESTINO_DESC)"
               SQL = SQL & " values("
                  SQL = SQL & txtTransf_ID.Text                                     'TRANSF_ID
                  SQL = SQL & "," & NUMR_SEQ_N                                      'SEQ_ID
                  SQL = SQL & "," & EMPRESA_ID_N                                    'EMPRESA_ID
                  SQL = SQL & "," & cmbOrigemAUX.Text                               'ESTAB_ORIGEM_ID
                  SQL = SQL & "," & cmbDestinoAUX.Text                              'ESTAB_DESTINO_ID
                  SQL = SQL & "," & Trim(lstDestino.ListItems(i).SubItems(3))       'PRODUTO_ID
                  SQL = SQL & "," & tpMOEDA(lstDestino.ListItems(i).SubItems(2))    'QTDE_TRANSF
                  SQL = SQL & ",'" & (Now) & "'"                                    'DT_TRANSF
                  SQL = SQL & ",'T'"                                                'SITUACAO
                  SQL = SQL & ",'" & TRAZ_ESTABELECIMENTO(cmbOrigemAUX.Text) & "'"  'ESTAB_ORIGEM_DESC
                  SQL = SQL & ",'" & TRAZ_ESTABELECIMENTO(cmbDestinoAUX.Text) & "'" 'ESTAB_DESTINO_DESC
               SQL = SQL & " )"
               CONECTA_RETAGUARDA.Execute SQL

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
   TRATA_ERROS Err.description, Me.name, "CmdGravar_Click"
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
   TRATA_ERROS Err.description, Me.name, "MSFlexGrid1_KeyDown"
End Sub

Sub MOSTRA_PRODUTO_TRANSF()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0 & TabProduto.Fields("produto_id").Value
   txtDesc2.Text = Trim(TabProduto!DESCRICAO)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MOSTRA_PRODUTO_TRANSF"
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
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_PRODUTO_TRANSF

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQtdeTransf.SetFocus
      Exit Sub
   End If

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   Qtde_N = 0

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_PRODUTO_TRANSF

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQtdeTransf.SetFocus
      Exit Sub
   End If

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS) = 13 Then
      '2 = produtos "in store" (sempre será 2)
      'C = código do produto (4,5 ou 6 dígitos)
      'T = total a pagar (sempre 6 dígitos)
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      'txtCodgProd.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 6))
      'pegando codigo do produto no codigo de barras da etiqueta de balança
      'txtCodgProd.Text = "" & Int(Mid(CODIGO_BARRAS, 2, TamanhoCodgProdBarra_N))

txtCodgProd.Text = "" & Int(Mid(CODIGO_BARRAS, CasaInicioCodgProdBarra_N, TamanhoCodgProdBarra_N))

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtCodgProd.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         MOSTRA_PRODUTO_TRANSF

         If UCase(PESO_VALOR_A) = UCase("valor") Then 'POR VALOR
            VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS, 8, TamanhoPesoValorBarra_N) / 100
            Qtde_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value) 'sta
            PESO_ITEM_N = Qtde_N
            txtQtdeTransf.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)

            Call txtQtdeTransf_KeyPress(13)

            txtCodgProd.SetFocus
            Else
               Qtde_N = 0 & Int(Mid(CODIGO_BARRAS, 8, 5))           'gramas
      
               If Qtde_N > 0 Then _
                  Qtde_N = Format(Qtde_N / 1000, strFormatacao3Digitos)
      
               PESO_ITEM_N = Qtde_N
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

   If Len(CODIGO_BARRAS) = 12 Then
      'lendo codigo barras ultralav
      '100004360813
      '1-1 = masculino ou feminino
      '2-7 = código do produto
      '8-9 = numeração tamanho produto
      '10-11 = mes
      '12-13 = ano

      txtCodgProd.Text = "" & Mid(CODIGO_BARRAS, 1, 6)
      SqL2 = "" & Mid(CODIGO_BARRAS, 7, 2)

      SQL = "select * from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where referencia = '" & Trim(txtCodgProd.Text) & "'"
      SQL = SQL & " and RIGHT(descricao,2) = '" & Trim(SqL2) & "'"
      SQL = SQL & " and situacao <> 'C' "
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
   TRATA_ERROS Err.description, Me.name, "LE_PRODUTO_TRANSF"
End Sub

Sub HABILITA_TRANSFERENCIA()
'On Error GoTo ERRO_TRATA

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
            SQL = SQL & " where estabelecimento_id <> " & ESTABELECIMENTO_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbDestino.Enabled = True
               cmbDestino.Clear
               cmbDestinoAUX.Clear
            End If
            While Not TabTemp.EOF

               cmbDestino.AddItem Trim(TabTemp.Fields("descricao").Value)
               cmbDestinoAUX.AddItem TabTemp.Fields("estabelecimento_id").Value

               TabTemp.MoveNext
            Wend
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
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "HABILITA_TRANSFERENCIA"
End Sub

Sub LIMPA_TRANSF()
'On Error GoTo ERRO_TRATA

   lstOrigem.ListItems.Clear
   lstDestino.ListItems.Clear
   txtCodgProd.Text = ""
   txtQtdeTransf.Text = ""
   txtDesc2.Text = ""
   cmbDestinoAUX.Text = ""
   cmbDestino.Text = ""
   Conta_Transf = 0
   PRODUTO_ID_N = -1
   txtTransf_ID.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "CmdGravar_Click"
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
   TRATA_ERROS Err.description, Me.name, "SETA_GRID_ORIGEM"
End Sub

Sub ATUALIZA_ESTOQUE_DESTINO(Qtde_Transf_N As Double, PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   SQL = "update ESTOQUE set "
   SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(Qtde_Transf_N)
   SQL = SQL & " where produto_id = " & Trim(PROD_ID_N)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "ATUALIZA_ESTOQUE_DESTINO"
End Sub

Sub ATUALIZA_ESTOQUE_origem(Qtde_Transf_N As Double, PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   SQL = "update ESTOQUE set "
   SQL = SQL & " qtde_estoque = qtde_estoque - " & tpMOEDA(Qtde_Transf_N)
   SQL = SQL & " where produto_id = " & PROD_ID_N
   SQL = SQL & " and estabelecimento_id = " & cmbOrigemAUX.Text
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "ATUALIZA_ESTOQUE_ORIGEM"
End Sub

Sub SETA_GRID_TRANSITO()
'On Error GoTo ERRO_TRATA

   CONT_N = 0
   lstTransf.Visible = False
   lstTransf.ListItems.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, ESTOQUETRANSF.ESTAB_ORIGEM_ID, "
   SQL = SQL & " ESTOQUETRANSF.ESTAB_DESTINO_ID, ESTOQUETRANSF.QTDE_TRANSF, "
   SQL = SQL & " ESTOQUETRANSF.SITUACAO , ESTOQUETRANSF.DT_TRANSF,"
   SQL = SQL & " ESTOQUETRANSF.transf_id as Lote, ESTOQUETRANSF.dt_entrada, "
   SQL = SQL & " seq_id,ESTOQUETRANSF.produto_id"

   SQL = SQL & " from ESTOQUETRANSF WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON ESTOQUETRANSF.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where estab_destino_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and ESTOQUETRANSF.situacao = 'T' "

   SQL = SQL & " order by transf_id "

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      CONT_N = CONT_N + 1
      PRODUTO_ID_N = 0 & TabConsulta.Fields("produto_id").Value
      Qtde_N = 0 & TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      Set item = lstTransf.ListItems.Add(, "seq." & CONT_N, TabConsulta.Fields("lote").Value)
      item.SubItems(1) = "" & Trim(TabConsulta.Fields("seq_id").Value)
      item.SubItems(2) = "" & Trim(TabConsulta.Fields("codg_produto").Value)
      item.SubItems(3) = "" & Trim(TabConsulta.Fields("descricao").Value)
      item.SubItems(4) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estab_origem_id").Value)
      item.SubItems(5) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estab_destino_id").Value)
      item.SubItems(6) = "" & Format(TabConsulta.Fields("qtde_transf").Value, strFormatacao3Digitos)
      item.SubItems(7) = "" & Format(Qtde_N, strFormatacao3Digitos)

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

      item.SubItems(9) = "" & TabConsulta.Fields("dt_transf").Value
      item.SubItems(10) = "" & TabConsulta.Fields("dt_entrada").Value
      item.SubItems(11) = "" & TabConsulta.Fields("produto_id").Value

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
      item.ListSubItems(10).ForeColor = vbRed

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   lstTransf.Visible = True
   PRODUTO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "SETA_GRID_TRANSITO"
End Sub

Sub TRANSF_SELECIONADOS()
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
               Msg = "Confirma recebimento de mercadoria(s) Selecionada(s) ? "
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

               '==========================inserindo na tabela estoque caso não exista
               RODA_AT_ESTOQUE PRODUTO_ID_N, ESTABELECIMENTO_ID_N

               ATUALIZA_ESTOQUE_DESTINO lstTransf.ListItems(i).SubItems(6), PRODUTO_ID_N

               SQL = "update ESTOQUETRANSF set "
               SQL = SQL & " situacao = 'F'"  'transferida
               SQL = SQL & ", dt_entrada = '" & Now & "'"

               SQL = SQL & " where estab_destino_id = " & ESTABELECIMENTO_ID_N
               SQL = SQL & " and situacao = 'T' "
               SQL = SQL & " and produto_id = " & PRODUTO_ID_N
               SQL = SQL & " and seq_id = " & SEQ_ID_N
               CONECTA_RETAGUARDA.Execute SQL
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
   TRATA_ERROS Err.description, Me.name, "TRANSF_SELECIONADOS"
End Sub
