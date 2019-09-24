VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmINVENTARIO 
   Caption         =   "Inventário"
   ClientHeight    =   7830
   ClientLeft      =   390
   ClientTop       =   2130
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INVENTARIO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab stbInventario 
      Height          =   6375
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   11245
      _Version        =   393216
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
      TabCaption(0)   =   "Contagem"
      TabPicture(0)   =   "INVENTARIO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblConta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAtualiza"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstInventario"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "optSeg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "optPri"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdZeraEstoque"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Acerto Automático"
      TabPicture(1)   =   "INVENTARIO.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "cmddif"
      Tab(1).Control(3)=   "cmdimp"
      Tab(1).Control(4)=   "cmdacerto"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "INVENTARIO.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdZeraEstoque 
         Caption         =   "Zerar Estoque"
         Height          =   285
         Left            =   9840
         TabIndex        =   37
         Top             =   6000
         Width           =   1935
      End
      Begin VB.OptionButton optPri 
         Caption         =   "1ª Contagem"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   5430
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optSeg 
         Caption         =   "2ª contagem"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   6
         Top             =   5430
         Width           =   1935
      End
      Begin VB.Frame Frame1 
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
         Height          =   1215
         Left            =   100
         TabIndex        =   20
         Top             =   60
         Width           =   11655
         Begin VB.CommandButton cmdCadProd 
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
            Left            =   4680
            Picture         =   "INVENTARIO.frx":5C66
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Cadastro Produto"
            Top             =   240
            Width           =   405
         End
         Begin VB.CommandButton cmdMata 
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
            Left            =   50
            Picture         =   "INVENTARIO.frx":B268
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Excluir Item Inventário"
            Top             =   240
            Width           =   405
         End
         Begin VB.CommandButton cmdConsProd 
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
            Left            =   4125
            Picture         =   "INVENTARIO.frx":C0A9
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Consulta Produto"
            Top             =   240
            Width           =   405
         End
         Begin VB.TextBox txtRef 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10440
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtPesoItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   10440
            TabIndex        =   24
            ToolTipText     =   "Quantidade que tem em Estoque"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   500
            TabIndex        =   23
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSegCont 
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
            Height          =   360
            Left            =   8520
            TabIndex        =   3
            ToolTipText     =   "Quantidade que tem em Estoque"
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtProduto 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   1
            ToolTipText     =   "Informe o código do produto."
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtDescricao 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5280
            MaxLength       =   29
            TabIndex        =   22
            Top             =   240
            Width           =   5055
         End
         Begin VB.TextBox txtQtdeEstoque 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   1920
            TabIndex        =   21
            ToolTipText     =   "Quantidade que tem em Estoque"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtPriCont 
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
            Height          =   360
            Left            =   5280
            TabIndex        =   2
            ToolTipText     =   "Quantidade que tem em Estoque"
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
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
            Left            =   9690
            TabIndex        =   29
            Top             =   720
            Width           =   525
         End
         Begin VB.Label lblseg 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qtde. 2ª Contagem:"
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
            Left            =   6570
            TabIndex        =   28
            Top             =   720
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblcodgprod 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Left            =   1005
            TabIndex        =   27
            Top             =   240
            Width           =   810
         End
         Begin VB.Label lblqtdestoque 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qtde. Estoque:"
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
            Left            =   435
            TabIndex        =   26
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label lblcontagem 
            AutoSize        =   -1  'True
            Caption         =   "Qtde. 1ª Contagem:"
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
            Left            =   3360
            TabIndex        =   25
            Top             =   720
            Width           =   1845
         End
      End
      Begin VB.CommandButton cmdacerto 
         Caption         =   "Gera Diferença de Estoque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   17
         Top             =   1260
         Width           =   4095
      End
      Begin VB.CommandButton cmdimp 
         Caption         =   "Imprime Diferença"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -70560
         TabIndex        =   16
         Top             =   1260
         Width           =   3855
      End
      Begin VB.CommandButton cmddif 
         Caption         =   "Atualiza Diferenças"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -70560
         TabIndex        =   15
         Top             =   2340
         Width           =   3855
      End
      Begin MSComctlLib.ListView lstInventario 
         Height          =   3945
         Left            =   105
         TabIndex        =   30
         ToolTipText     =   "Clique para selecionar um produto ja gravado."
         Top             =   1380
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   6959
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sq."
            Object.Width           =   979
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Produto"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qtd. Estoque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qtd 1º Cont."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Diferença1º"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Qtd. 2º Cont."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Produto_ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Referencia"
            Object.Width           =   2540
         EndProperty
      End
      Begin Threed.SSCommand cmdAtualiza 
         Height          =   975
         Left            =   4800
         TabIndex        =   4
         Top             =   5370
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1720
         _Version        =   262144
         CaptionStyle    =   1
         BackColor       =   16777152
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "INVENTARIO.frx":CAAB
         Caption         =   "Atualizar Contagem"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin VB.Label lblConta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Qtde. Atualizados: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   8280
         TabIndex        =   31
         Top             =   5580
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Entradas Atualizadas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   -74640
         TabIndex        =   19
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Saídas Atualizadas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   -74640
         TabIndex        =   18
         Top             =   2700
         Width           =   2490
      End
   End
   Begin VB.OptionButton optEnt 
      Caption         =   "Entrada"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   850
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optSai 
      Caption         =   "Saída"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   850
      Width           =   1455
   End
   Begin VB.OptionButton optInvent 
      Caption         =   "Contagem"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   850
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1270
      ButtonWidth     =   3466
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contagem"
            Key             =   "relcontagem"
            Object.ToolTipText     =   "Imprimir Relatório de Contagem"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "clear"
            Object.ToolTipText     =   "Limpar a Tela"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "InventároVenda"
            Key             =   "relinventario"
            Object.ToolTipText     =   "Imprimir Relatório Final do Inventário"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "InventárioCusto"
            Key             =   "invcusto"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zerar Inv."
            Key             =   "zera"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir Lote"
            Key             =   "exclui_lote"
            ImageIndex      =   2
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
         Left            =   10920
         TabIndex        =   33
         Top             =   360
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
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":120BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":13257
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":14489
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":15725
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":16830
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":178BF
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":18874
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":19AF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":1AD75
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "INVENTARIO.frx":1C218
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDtLote 
      Height          =   360
      Left            =   4440
      TabIndex        =   10
      ToolTipText     =   "Data do lote"
      Top             =   900
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Tecle Enter Para gerar Numero do Lote!"
      Top             =   900
      Width           =   1335
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11910
      DesignHeight    =   7830
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      Height          =   615
      Left            =   50
      Top             =   750
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      Height          =   615
      Left            =   7080
      Top             =   750
      Width           =   4815
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Data Lote:"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lblLote 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Número Lote:"
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
      TabIndex        =   11
      Top             =   900
      Width           =   1335
   End
End
Attribute VB_Name = "frmINVENTARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim Tipo_Mov_A    As String
   Dim Conta_Transf  As Long
   Dim REGISTRO_A    As String

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name
   txtDtLote.Text = Now
   CHECA_TAB_TEMP

   cmdZeraEstoque.Visible = False
   If TIPO_USUARIO = 5 Then _
      cmdZeraEstoque.Visible = True

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

Private Sub cmdZeraEstoque_Click()
   Msg = "Deseja ZERA QTDE ESTOQUE DE TODOS PRODUTOS DESTE ESTABELECIMENTO ?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SQL = "update ESTOQUE set "
      SQL = SQL & " qtde_estoque = 0 "
      SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
      CONECTA_RETAGUARDA.Execute SQL
      MsgBox "Processo realizado com sucesso."
   End If
End Sub

Private Sub cmdCadProd_Click()
   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then _
      frmCADASTROPRODUTO.Show 1
End Sub

Private Sub cmdMata_Click()
   If Trim(txtSeq.Text) = "" Then _
      Exit Sub
   If Trim(txtLOTE.Text) = "" Then _
      Exit Sub

   MATA_ITEM txtSeq.Text
End Sub

Private Sub cmdacerto_Click()
'On Error GoTo ERRO_TRATA

   Msg = "Deseja Fazer Realmente esta operacao?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      LE_KARDEX
      
      Else
         RESPOSTA = ""
         Exit Sub
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdacerto_Click"
End Sub

Private Sub cmddif_Click()
'On Error GoTo ERRO_TRATA

   Msg = "Deseja Realmente Atualizar as Diferencas?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      ATUALIZA_DIF
      Else
         RESPOSTA = ""
         Exit Sub
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmddif_Click"
End Sub

Private Sub cmdimp_Click()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ESTACERTO WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by CODG_PRODUTO"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      FORMULA_REL = "{estacerto.CODG_PRODUTO} <> '" & "" & "'"
      If FORMULA_REL <> "" Then
         SqL2 = "+"

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_acerto.rpt"
         frmRELATORIO10.Show 1
       End If
   End If
   If TabInventario.State = 1 Then _
      TabInventario.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdimp_Click"
End Sub

Private Sub lstinventario_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstInventario, ColumnHeader
End Sub

Private Sub lstInventario_DblClick()
   If Not IsNull(lstInventario.SelectedItem.Text) Then _
      If Trim(lstInventario.SelectedItem.Text) <> "" Then _
         If IsNumeric(lstInventario.SelectedItem.Text) Then _
            txtSeq.Text = lstInventario.SelectedItem.Text
End Sub

Private Sub lstInventario_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete  'Excluir linhas selecionadas
         If IsNull(lstInventario.SelectedItem.Text) Then _
            Exit Sub
         If Trim(lstInventario.SelectedItem.Text) = "" Then _
            Exit Sub
         If Not IsNumeric(lstInventario.SelectedItem.Text) Then _
            Exit Sub
         If Trim(txtLOTE.Text) = "" Then _
            Exit Sub

         MATA_ITEM lstInventario.SelectedItem.Text
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstInventario_KeyDown"
End Sub

Private Sub optsai_Click()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 1"
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      TabUSU.Close
      MsgBox "Problemas com usuário, codigo=0"
      'txtent.Value = True
      Exit Sub
      Else
         Msg = "Digite Senha do Superior para Tirar Mercadoria do Estoque! "
         Msg = Msg & "Deseja liberar com senha superior ?"
         PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            frmSenha.Show 1

            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "select * from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where senha = '" & Trim(CRITERIO_A) & "'"
            SQL = SQL & " and status = 1"
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then
               If Not IsNull(TabUSU.Fields("tipo").Value) Then
                  If TabUSU.Fields("tipo").Value >= 4 And TabUSU.Fields("tipo").Value <= 5 Then
                     optEnt.Enabled = True
                     optSai.Enabled = True
                     cmdAtualiza.Enabled = True
                     Else
                        MsgBox "Usuario nao Permitido para esta alteração."
                        optEnt.Enabled = False
                        optSai.Enabled = False
                        cmdAtualiza.Enabled = False
                  End If
               End If
            End If

            If TabUSU.State = 1 Then _
               TabUSU.Close

            RESPOSTA = ""
            Else
               optEnt.Value = True
               optSai.Value = False
               cmdAtualiza.Enabled = True
         End If
         txtLOTE.SetFocus
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   DoEvents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optsai_Click"
End Sub

Private Sub stbInventario_Click(PreviousTab As Integer)

   Shape2.Visible = True
   Shape1.Visible = True
   lblLote.Visible = True
   txtLOTE.Visible = True
   lblData.Visible = True
   txtDtLote.Visible = True
   optEnt.Visible = True
   optSai.Visible = True
   optInvent.Visible = True

   If stbInventario.Tab = 1 Then _
      stbInventario.Tab = 0
   If stbInventario.Tab = 2 Then
      Shape2.Visible = False
      Shape1.Visible = False
      lblLote.Visible = False
      txtLOTE.Visible = False
      lblData.Visible = False
      txtDtLote.Visible = False
      optEnt.Visible = False
      optSai.Visible = False
      optInvent.Visible = False
   End If

End Sub

Private Sub txtLote_GotFocus()
   MOSTRA_RODAPE "ESC - SAIR", "Tecle <ENTER> para gerar novo Lote ou informe uma já existente", "", "", ""

   txtLOTE.SelStart = 0
   txtLOTE.SelLength = Len(txtLOTE)
   txtLOTE.BackColor = &HC0FFFF

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         frmInventarioConsulta.Show 1
      Case "sair"
         Unload Me
      Case "relcontagem"
         FORMULA_REL = "{ESTABELECIMENTO.ESTABELECIMENTO_ID} = " & ESTABELECIMENTO_ID_N
         FORMULA_REL = FORMULA_REL & " and {PRODUTO.situacao} <> 'C'"
         FORMULA_REL = FORMULA_REL & " and {PRODUTO.produto_ID} = {ESTOQUE.produto_ID}"

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Contagem.rpt"
         frmRELATORIO10.Show 1
      Case "clear"
         LIMPA_TUDO
         txtLOTE.SetFocus
      Case "invcusto"
         CRITERIO_A = "custo"
         MONTA_REL_INVENTARIO
      Case "relinventario"
         CRITERIO_A = "venda"
         MONTA_REL_INVENTARIO
      Case "exclui_lote"
         MATA_LOTE
      Case "zera"
         Beep
         Msg = "Deseja Zerar Arquivo de Inventario?"
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            SQL = "delete from INVENTARIO "
            SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL
            MsgBox "Arquivo Zerado com Sucesso"
            Exit Sub
         End If
         If RESPOSTA = vbNo Then _
            Exit Sub
   End Select

CRITERIO_A = ""
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtProduto.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtlote_KeyPress"
End Sub

Private Sub txtLote_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtLOTE.Text) = "" Then
      GERA_NUMR_LOTE
      txtLOTE.Text = NUMR_LOTE_N
      Else: NUMR_LOTE_N = txtLOTE.Text
   End If

   If TabInventario.State = 1 Then _
      TabInventario.Close

   SQL = "select * from INVENTARIO WITH (NOLOCK)"
   SQL = SQL & " where numr_lote = " & NUMR_LOTE_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabInventario.EOF Then
      txtDtLote.PromptInclude = False
      txtDtLote.Text = TabInventario!DT_LOTE
      txtDtLote.PromptInclude = True

      SETA_GRID

      If Trim(TabInventario!STATUS) = "F" Then
         If TabInventario.State = 1 Then _
            TabInventario.Close

         MsgBox "Lote já atualizado, impossível alterar !!!"

         LIMPA_TUDO

         'txtLote.SetFocus
         Else: MsgBox "Lote Esta aberto para alteração."
      End If
   End If
   If TabInventario.State = 1 Then _
      TabInventario.Close

   txtLOTE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtlote_LostFocus"
End Sub

Private Sub txtpricont_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtPriCont.Text) = "" Then _
      txtPriCont.Text = "" & Format(txtPriCont.Text, strFormatacao3Digitos)

   txtPriCont.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpricont_LostFocus"
End Sub

Private Sub TXTSEGCONT_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtSegCont.Text) = "" Then _
      txtSegCont.Text = "" & Format(txtSegCont.Text, strFormatacao3Digitos)
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEGCONT_LostFocus"
End Sub

Private Sub cmdConsProd_Click()
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
End Sub

Private Sub txtProduto_GotFocus()
   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.BackColor = &HC0FFFF
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtProduto.Text) <> "" Then _
            If Trim(txtLOTE.Text) <> "" Then _
               If Trim(txtSeq.Text) <> "" Then _
                  MATA_ITEM txtSeq.Text
      Case vbKeyF7
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtProduto.ForeColor = vbBlue
   txtDescricao.ForeColor = vbBlue

   If Trim(txtLOTE.Text) = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCESSA_DADOS_PRODUTOS
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub txtpricont_GotFocus()
'On Error GoTo ERRO_TRATA

   txtPriCont.SelStart = 0
   txtPriCont.SelLength = Len(txtPriCont)
   txtPriCont.BackColor = &HC0FFFF

   If optInvent.Value = True Then
      If TabInventario.State = 1 Then _
         TabInventario.Close

       SQL = "select * from INVENTARIO WITH (NOLOCK)"
       SQL = SQL & " where seq = " & txtSeq.Text
       SQL = SQL & " and numr_lote = " & txtLOTE.Text
       SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
       TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
       If Not TabInventario.EOF Then
          If TabInventario!QTD_PRIMEIRA <> "" Then
             txtPriCont.Enabled = False
             txtPriCont.Text = TabInventario!QTD_PRIMEIRA
             txtSegCont.SetFocus
             Exit Sub
          End If
       End If
      If TabInventario.State = 1 Then _
         TabInventario.Close

       If Trim(txtProduto.Text) = Empty Then
          MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
          txtProduto.Text = 99999999
          txtProduto.SetFocus
          Exit Sub
       End If
       If txtPriCont.Text <> "" Then
          txtPriCont.SelStart = 0
          txtPriCont.SelLength = Len(txtPriCont)
       End If
  End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpricont_GotFocus"
End Sub

Private Sub txtpricont_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtPriCont.Text) <> "" Then

         GRAVA_INVENTARIO

         LIMPA_BODY
         txtProduto.SetFocus
         Else: txtSegCont.SetFocus
      End If
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpricont_KeyPress"
End Sub

Private Sub TXTSEGCONT_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = Empty Then
      MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
      txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   If txtSegCont.Text <> "" Then
      txtSegCont.SelStart = 0
      txtSegCont.SelLength = Len(txtSegCont)
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEGCONT_GotFocus"
End Sub

Private Sub TXTSEGCONT_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtSegCont.Text <> "" Then
         GRAVA_INVENTARIO
         LIMPA_BODY
         txtProduto.SetFocus
         Exit Sub
         Else
            GRAVA_INVENTARIO
            LIMPA_BODY
            txtProduto.SetFocus
      End If
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEGCONT_KeyPress"
End Sub

'============================
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If stbInventario.Tab = 0 Then
      lstInventario.ListItems.Clear
      lstInventario.Visible = False
      NUMR_SEQ_N = 0
      CONT_N = 0
      QTDE_N = 0

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select INVENTARIO.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
      SQL = SQL & " PRODUTO.REFERENCIA, PRODUTO.TIPO_PROD, "
      SQL = SQL & " Produto.FORNECEDOR_ID , Produto.tamanho "
      SQL = SQL & " from INVENTARIO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON INVENTARIO.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where numr_lote = " & txtLOTE.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " order by seq desc"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         CONT_N = CONT_N + 1
         Set item = lstInventario.ListItems.Add(, "seq." & CONT_N, TabConsulta!SEQ)
         item.SubItems(1) = Trim(TabConsulta.Fields("codg_produto").Value)

         SQL3 = ""
         If Trim(TabConsulta.Fields("referencia").Value) <> "" Then _
            SQL3 = "  |  " & Trim(TabConsulta.Fields("referencia").Value)

         item.SubItems(2) = "" & Trim(TabConsulta!DESCRICAO) & SQL3
         'Item.SubItems(3) = "" & Format(TabConsulta!qtd_anterior, strFormatacao3Digitos)

         QTDE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabConsulta.Fields("produto_id").Value)
         item.SubItems(3) = "" & Format(QTDE_N, strFormatacao3Digitos)

         item.SubItems(4) = "" & Format(TabConsulta!QTD_PRIMEIRA, strFormatacao3Digitos)

         If QTDE_N < 0 Then
            item.SubItems(5) = "" & Format(QTDE_N + TabConsulta!QTD_PRIMEIRA, strFormatacao3Digitos)
            Else: item.SubItems(5) = "" & Format(QTDE_N - TabConsulta!QTD_PRIMEIRA, strFormatacao3Digitos)
         End If

         item.SubItems(6) = "" & Format(TabConsulta!qtd_segunda, strFormatacao3Digitos)
         item.SubItems(7) = "" & TabConsulta.Fields("produto_id").Value
         item.SubItems(8) = "" & Trim(TabConsulta.Fields("referencia").Value)

         If TabProduto.State = 1 Then _
            TabProduto.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      lstInventario.Visible = True
   End If
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GRAVA_INVENTARIO()
'On Error GoTo ERRO_TRATA

   txtDtLote.PromptInclude = True
   If Not IsDate(txtDtLote.Text) Then
      txtDtLote.PromptInclude = False
      txtDtLote.Text = Now
      txtDtLote.PromptInclude = True
   End If

   If optEnt.Value = True Then _
      Tipo_Mov_A = "E"
   If optSai.Value = True Then _
      Tipo_Mov_A = "S"
   If optInvent.Value = True Then _
      Tipo_Mov_A = "C"

   If stbInventario.Tab = 0 Then
      CRITERIO_A = txtLOTE.Text
      SqL2 = ESTABELECIMENTO_ID_N
      NUMR_SEQ_N = 0 & MAX_ID("seq", "INVENTARIO", "numr_lote", txtLOTE.Text, "estabelecimento_id", SqL2)
      REGISTRO_A = ""

      If TabInventario.State = 1 Then _
         TabInventario.Close

      SQL = "select produto_id from INVENTARIO WITH (NOLOCK)"
      SQL = SQL & " where produto_id = " & PRODUTO_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and numr_lote = " & Trim(txtLOTE.Text)
      TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabInventario.EOF Then _
         REGISTRO_A = "*"

      If TabInventario.State = 1 Then _
         TabInventario.Close

      SQL = "select seq from INVENTARIO WITH (NOLOCK)"
      SQL = SQL & " where seq = " & NUMR_SEQ_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and numr_lote = " & Trim(txtLOTE.Text)
      TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabInventario.EOF Then
         CRITERIO_A = txtDtLote.Text

         SQL = "insert into INVENTARIO "
         SQL = SQL & "("
            SQL = SQL & "NUMR_LOTE,SEQ,PRODUTO_ID,"
            SQL = SQL & " QTD_ANTERIOR,QTD_PRIMEIRA,"
            SQL = SQL & " QTD_SEGUNDA,QTD_ATUAL,DT_LOTE,STATUS,TIPO_MOV,"
            SQL = SQL & " ESTABELECIMENTO_ID, REGISTRO "
         SQL = SQL & ")"
         SQL = SQL & " values ("

            SQL = SQL & txtLOTE.Text                                    '[NUMR_LOTE]
            SQL = SQL & ",0" & NUMR_SEQ_N                               '[SEQ]
            SQL = SQL & ",0" & PRODUTO_ID_N                             '[PRODUTO_ID]
            SQL = SQL & "," & tpMOEDA(txtQtdeEstoque.Text)              '[QTD_ANTERIOR]

            If txtPriCont.Text <> "" Then
               SQL = SQL & "," & tpMOEDA(txtPriCont.Text)               '[QTD_PRIMEIRA]
               Else: SQL = SQL & ",0"                                   '[QTD_PRIMEIRA]
            End If

            If txtSegCont.Text <> "" Then
               SQL = SQL & ",0" & txtSegCont.Text                       '[QTD_SEGUNDA]
               Else: SQL = SQL & ",0"                                   '[QTD_SEGUNDA]
            End If
         
            SQL = SQL & "," & 0                                         '[QTD_ATUAL]
            SQL = SQL & ",'" & CRITERIO_A & "'"                           '[DT_LOTE]
            SQL = SQL & ",'" & "A" & "'"                                '[STATUS]
            SQL = SQL & ",'" & Tipo_Mov_A & "'"                         'tipo_mov
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                      'ESTABELECIMENTO_ID
            SQL = SQL & ",'" & REGISTRO_A & "'"                         'registro

         SQL = SQL & ")"

         CONECTA_RETAGUARDA.Execute SQL
         Else:
            If Trim(txtSegCont.Text) <> "" Then
               If IsNumeric(txtSegCont.Text) Then
                  SQL = "update INVENTARIO set "
                  SQL = SQL & " qtd_segunda = " & tpMOEDA(txtSegCont.Text)
                  SQL = SQL & ", numr_lote = " & txtLOTE.Text
                  SQL = SQL & ", seq = " & TabInventario.Fields("seq").Value
                  SQL = SQL & ", tipo_mov = '" & Tipo_Mov_A & "'"                   'tipo_mov

                  SQL = SQL & " where seq = " & TabInventario.Fields("seq").Value
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  SQL = SQL & " and numr_lote = " & Trim(txtLOTE.Text)

                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
            ' Senao for pela segunda contagem e acerto entao continua pela primeira contagem
            If Trim(txtPriCont.Text) <> "" Then
               If IsNumeric(txtPriCont.Text) Then
                  SQL = "update INVENTARIO set "
                  SQL = SQL & " qtd_primeira = " & tpMOEDA(txtPriCont.Text)
                  SQL = SQL & ", numr_lote = " & txtLOTE.Text
                  SQL = SQL & ", seq = " & TabInventario.Fields("seq").Value
                  SQL = SQL & ", tipo_mov = '" & Tipo_Mov_A & "'"                   'tipo_mov

                  SQL = SQL & " where seq = " & TabInventario.Fields("seq").Value
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  SQL = SQL & " and numr_lote = " & Trim(txtLOTE.Text)

                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
      End If
      If TabInventario.State = 1 Then _
         TabInventario.Close
   End If

   SETA_GRID
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_INVENTARIO"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtLOTE.Text = ""
   txtDtLote.Mask = "##/##/####"
   txtProduto.Text = ""
   txtRef.Text = ""
   txtPriCont.Text = ""
   txtSegCont.Text = ""
   txtQtdeEstoque.Text = ""
   lstInventario.ListItems.Clear
   NUMR_LOTE_N = 0
   NUMR_SEQ_N = 0
   QTDE_ESTOQUE_N = 0
txtDtLote.PromptInclude = False
   txtDtLote.Text = Now
txtDtLote.PromptInclude = True
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   txtRef.Text = ""
   txtSegCont.Text = ""
   txtDescricao.Text = ""
   txtQtdeEstoque.Text = ""
   txtPriCont.Text = ""
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub MATA_ITEM(NUMR_SEQUENCIA_N As Long)
'On Error GoTo ERRO_TRATA

   If NUMR_SEQUENCIA_N <= 0 Then _
      Exit Sub

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If stbInventario.Tab = 0 Then
      SQL = "select * from INVENTARIO WITH (NOLOCK)"
      SQL = SQL & " where seq = " & NUMR_SEQUENCIA_N
      SQL = SQL & " and numr_lote = " & txtLOTE.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Confirma exclusão desta sequencia ? " & TabTemp.Fields("seq").Value
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            NUMR_SEQ_N = 0 & TabTemp.Fields("seq").Value

            SQL = "delete from INVENTARIO "
            SQL = SQL & " where numr_lote = " & txtLOTE.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            SQL = SQL & " and seq = " & NUMR_SEQ_N
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_BODY
            SETA_GRID
         End If
         Else: MsgBox "Produto não encontrado neste lote."
      End If
      txtProduto.SetFocus
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_ITEM"
End Sub

Private Sub MATA_LOTE()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If txtLOTE.Text <> "" Then
      SQL = "select * from INVENTARIO WITH (NOLOCK)"
      SQL = SQL & " where numr_lote = " & txtLOTE.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         Msg = "Deseja Excluir Esse Lote?"
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            SQL = "delete from INVENTARIO "
            SQL = SQL & " where numr_lote = " & txtLOTE.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_TUDO
         End If
         Else: MsgBox "Lote nao Localizado."
      End If
      txtLOTE.SetFocus
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_LOTE"
End Sub

Private Sub cmdAtualiza_Click()
'On Error GoTo ERRO_TRATA

   Dim TOTAL_REG_PRI As Integer

   If Trim(txtLOTE.Text) = "" Then
      MsgBox "Lote inválido."
      txtLOTE.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtLOTE.Text) Then
      MsgBox "Lote inválido."
      txtLOTE.SetFocus
      Exit Sub
   End If

   If optEnt.Value = True Then _
      Tipo_Mov_A = "E"
   If optSai.Value = True Then _
      Tipo_Mov_A = "S"
   If optInvent.Value = True Then _
      Tipo_Mov_A = "C"

   NUMR_LOTE_N = Trim(txtLOTE.Text)
   TOTAL_REG_PRI = 0

   If TabInventario.State = 1 Then _
      TabInventario.Close

   SQL = "select * from INVENTARIO WITH (NOLOCK)"
   SQL = SQL & " where numr_lote = " & NUMR_LOTE_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabInventario.EOF Then
      If TabInventario!STATUS = "F" Then
         If TabInventario.State = 1 Then _
            TabInventario.Close

         MsgBox "Inventario ja Atualizado!, Favor Confirmar!"
         LIMPA_TUDO
         Exit Sub
      End If
   End If
   If TabInventario.State = 1 Then _
      TabInventario.Close

   ATUALIZA_INVENTARIO
   LIMPA_TUDO

   stbInventario.Tab = 0
   txtLOTE.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdatualiza_Click"
End Sub

Private Sub LE_KARDEX()
'On Error GoTo ERRO_TRATA

   Dim TOTAL_SAIDAS  As Double
   Dim TOTAL_ENTRADAS As Double

    TOTAL_SAIDAS = 0
    TOTAL_ENTRADAS = 0
    
   'Zerando Arquivo para nao duplicar quantidades
   SQL = "delete from ESTACERTO "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   CONECTA_RETAGUARDA.Execute SQL
   
   SQL = "select * from QryFinalKardex WITH (NOLOCK)"
   SQL = SQL & " where QryFinalKardex.status <> '" & 1 & "'" 'Orcamento
   SQL = SQL & " and QryFinalKardex.status <> '" & 2 & "'" 'Gerado
   SQL = SQL & " and QryFinalKardex.status <> '" & 9 & "'" 'Gerado
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      
      If TabCabeca!TIPO = "ENTRADA" Then
         SQL = "select * from ESTACERTO WITH (NOLOCK)"
         SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and CODG_PRODUTO = '" & TabCabeca!Codg_Produto & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            SQL = "update ESTACERTO set "
            SQL = SQL & " qtde_entrada = " & Replace(TabCabeca!QTDE_ENTRADA + TabTemp!QTDE_ENTRADA, ",", ".")
            SQL = SQL & ", USUARIO_ID = " & USUARIO_ID_N ' Codigo Usuario Gravado
            SQL = SQL & " where CODG_PRODUTO = '" & Trim(TabTemp!Codg_Produto) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            CONECTA_RETAGUARDA.Execute SQL
            Else 'Criando Registro
                SQL = "insert into ESTACERTO values ("
                SQL = SQL & EMPRESA_ID_N ' empresa ai
                SQL = SQL & ",'" & TabCabeca!Codg_Produto & "'" 'codigo do produto
                SQL = SQL & "," & 0 'qtd saida sempre 0 no caso de entrada
                SQL = SQL & "," & Replace(TabCabeca!QTDE_ENTRADA, ",", ".") 'atualizando entradas
                CRITERIO_A = Format(Date, "dd/mm/yyyy")

                SQL = SQL & ",'" & CRITERIO_A & "'" ' data de geracao do processo

                SqL2 = "select qtd from PRODUTO WITH (NOLOCK)"
                SqL2 = SqL2 & " where CODG_PRODUTO = '" & TabCabeca!Codg_Produto & "'"
                SqL2 = SqL2 & " and situacao <> 'C' "
                TabProduto.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
                If Not TabProduto.EOF Then
                   SQL = SQL & "," & Replace(TabProduto!QTD, ",", ".") 'qtd atual estoque
                   Else: SQL = SQL & "," & 0                      'qtd atual estoque
                End If
                TabProduto.Close

                SQL = SQL & "," & USUARIO_ID_N 'Codigo do usuario que realizou!
                SQL = SQL & ")"
                CONECTA_RETAGUARDA.Execute SQL

                TOTAL_ENTRADAS = TOTAL_ENTRADAS + 1
                'lblent.Caption = "= " & TOTAL_ENTRADAS
                'lblent.ForeColor = vbBlue
                'lblent.Refresh
                DoEvents
         End If
         TabTemp.Close
         Else 'Saida de mercadorias
             SQL = "select * from ESTACERTO WITH (NOLOCK)"
             SQL = SQL & " where estacerto.empresa_id = " & EMPRESA_ID_N
             SQL = SQL & " and estacerto.CODG_PRODUTO = '" & TabCabeca!Codg_Produto & "'"
             TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
             If Not TabTemp.EOF Then
                SQL = "update ESTACERTO set "
                SQL = SQL & " qtd_saida = " & Replace(TabCabeca!QTDE_ENTRADA + TabTemp!qtd_saida, ",", ".")
                SQL = SQL & ", USUARIO_ID = " & USUARIO_ID_N ' Codigo Usuario Gravado
                SQL = SQL & " where CODG_PRODUTO = '" & Trim(TabTemp!Codg_Produto) & "'"
                SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                CONECTA_RETAGUARDA.Execute SQL
                Else 'Criando Registro
                    SQL = "insert into ESTACERTO values ("
                    SQL = SQL & EMPRESA_ID_N ' empresa ai
                    SQL = SQL & ",'" & TabCabeca!Codg_Produto & "'" 'codigo do produto
                    SQL = SQL & "," & Replace(TabCabeca!QTDE_ENTRADA, ",", ".") 'atualizando saida
                    SQL = SQL & "," & 0 'qtd entrada sempre 0 no caso de saida
                    CRITERIO_A = Format(Date, "dd/mm/yyyy")
                    SQL = SQL & ",'" & CRITERIO_A & "'" ' data de geracao do processo
                    
                    SqL2 = "select qtd from PRODUTO WITH (NOLOCK)"
                    SqL2 = SqL2 & " where CODG_PRODUTO = '" & TabCabeca!Codg_Produto & "'"
                    SqL2 = SqL2 & " and situacao <> 'C' "
                    TabProduto.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
                    If Not TabProduto.EOF Then
                       SQL = SQL & "," & Replace(TabProduto!QTD, ",", ".") 'qtd atual estoque
                       Else: SQL = SQL & "," & 0                          'qtd atual estoque
                    End If
                    SQL = SQL & "," & USUARIO_ID_N 'Codigo do usuario que realizou!
                    SQL = SQL & ")"
                    CONECTA_RETAGUARDA.Execute SQL
                    TOTAL_SAIDAS = TOTAL_SAIDAS + 1
                    'lblsaida.Caption = "= " & TOTAL_SAIDAS
                    'lblsaida.ForeColor = vbRed
                    'lblsaida.Refresh
                    DoEvents
             End If
             TabTemp.Close
      End If
      TabCabeca.MoveNext
   Wend
   TabCabeca.Close
   MsgBox "Processo Atualizado com Sucesso!"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_KARDEX"
End Sub

Private Sub ATUALIZA_DIF()
'On Error GoTo ERRO_TRATA

    Dim TOTAL_ATU As Double
    TOTAL_ATU = 0
    
    SQL = "select * from ESTACERTO WITH (NOLOCK)"
    SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
    TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
    While Not TabTemp.EOF
       SqL2 = "select qtd from PRODUTO WITH (NOLOCK)"
       SqL2 = SqL2 & " where CODG_PRODUTO = '" & Trim(TabTemp!Codg_Produto) & "'"
       SqL2 = SqL2 & " and situacao <> 'C' "
       TabProduto.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
       If Not TabProduto.EOF Then
          If (TabTemp!QTDE_ENTRADA - TabTemp!qtd_saida) >= 0 Then
              SQL = "update PRODUTO set "
              SQL = SQL & " qtde = " & Replace(TabTemp!QTDE_ENTRADA - TabTemp!qtd_saida, ",", ".")
              SQL = SQL & " where CODG_PRODUTO = '" & Trim(TabTemp!Codg_Produto) & "'"
              CONECTA_RETAGUARDA.Execute SQL
              
              TOTAL_ATU = TOTAL_ATU + 1
              'lblatu.Caption = "QTD de Itens Atual = " & TOTAL_ATU
              'lblatu.ForeColor = vbGreen
              'lblatu.Refresh
              DoEvents
              Else 'Saldo Negativo vou zerar estoque
                  SQL = "update PRODUTO set "
                  SQL = SQL & " qtde = " & 0
                  SQL = SQL & " where CODG_PRODUTO = '" & Trim(TabTemp!Codg_Produto) & "'"
                  CONECTA_RETAGUARDA.Execute SQL
                    
                  TOTAL_ATU = TOTAL_ATU + 1
                  'lblatu.Caption = "QTD de Itens Atual = " & TOTAL_ATU
                  'lblatu.ForeColor = vbGreen
                  'lblatu.Refresh
                  DoEvents
          End If
      End If
      TabTemp.MoveNext
   Wend
   TabTemp.Close

   MsgBox "Atualizacao realizada Com sucesso!"
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_DIF"
End Sub

Sub ATUALIZA_INVENTARIO()
'On Error GoTo ERRO_TRATA

   If optInvent.Value = True Then
      Msg = "Deseja ZERAR a quantidade dos produtos do estoque antes de atualizar a contagem ?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         If TabInventario.State = 1 Then _
            TabInventario.Close

         SQL = "select distinct(produto_id) from INVENTARIO WITH (NOLOCK)"
         SQL = SQL & " where numr_lote = " & NUMR_LOTE_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabInventario.EOF

            SQL = "update ESTOQUE set "
            SQL = SQL & " qtde_estoque = 0 "
            SQL = SQL & " where produto_id = " & TabInventario.Fields("produto_id").Value
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            TabInventario.MoveNext
         Wend
      End If
   End If
   If TabInventario.State = 1 Then _
      TabInventario.Close

   SQL = "select * from INVENTARIO WITH (NOLOCK)"
   SQL = SQL & " where numr_lote = " & NUMR_LOTE_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabInventario.EOF
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where produto_id = " & TabInventario.Fields("produto_id").Value
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         '==========================inserindo na tabela estoque caso não exista
         If TabTemp.State = 1 Then _
            TabTemp.Close
      
         SQL = "select * from ESTOQUE WITH (NOLOCK)"
         SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then

            SQL = "insert into ESTOQUE values("
               SQL = SQL & MAX_ID("estoque_id", "estoque", "", "", "", "") 'ESTOQUE_ID
               SQL = SQL & "," & ESTABELECIMENTO_ID_N                      'ESTABELECIMENTO_ID
               SQL = SQL & "," & TabProduto.Fields("produto_id").Value     'PRODUTO_ID
               SQL = SQL & "," & 0                                         'QTDE_ESTOQUE
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
         '==========================

         TOTAL_REG_PRI = TOTAL_REG_PRI + 1
         lblConta.Caption = "1º Contagem = " & TOTAL_REG_PRI
         lblConta.ForeColor = vbRed
         lblConta.Refresh
         DoEvents

         QTDE_N = 0

         'Atualizando pela Primeira Contagem
         If optPri.Value = True Then _
            QTDE_N = 0 & TabInventario.Fields("qtd_primeira").Value

         If optSeg.Value = True Then _
            QTDE_N = 0 & TabInventario.Fields("qtd_segunda").Value


         'atualizando estoque
         SQL = "update ESTOQUE set "

         'foi modificado porque na contagem pode ter mais de uma vez o mesmo pedido
         'então no começo desta rotina ele zera a qtde_estoque tabela estoque e depois acumula
         'neste campo seguindo a filosofia
         If optInvent.Value = True Then _
            SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(QTDE_N)

         If optSai.Value = True Then _
            SQL = SQL & " qtde_estoque = qtde_estoque - " & tpMOEDA(QTDE_N)

         If optEnt.Value = True Then _
            SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(QTDE_N)

         SQL = SQL & " where produto_id = " & Trim(TabInventario.Fields("produto_id").Value)
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         CONECTA_RETAGUARDA.Execute SQL

         '==========================atualizando tabela inventário
         SQL = "update INVENTARIO set "
         SQL = SQL & " status = 'F'" 'status de Inventario Atualizado
         SQL = SQL & ", tipo_mov = '" & Tipo_Mov_A & "'"                   'tipo_mov

         SQL = SQL & " where produto_id = " & TabInventario.Fields("produto_id").Value
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and numr_lote = " & Trim(txtLOTE.Text)
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close

      TabInventario.MoveNext
   Wend
   If TabInventario.State = 1 Then _
      TabInventario.Close

   MsgBox "Processo realizado com sucesso !!!"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_INVENTARIO"
End Sub

Sub CHECA_TAB_TEMP()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "INVENTARIOREL", "U") = False Then
      SQL = "CREATE TABLE [dbo].[INVENTARIOREL]("
      SQL = SQL & " [LOTE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [SEQ_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CODG_PROD] [nvarchar](100) NOT NULL,"
      SQL = SQL & " [DESC_PROD] [nvarchar](100) NOT NULL,"
      SQL = SQL & " [SALDO_ATUAL] [float] NOT NULL,"
      SQL = SQL & " [CONTAGEM1] [float] NOT NULL,"
      SQL = SQL & " [CONTAGEM2] [float] NULL,"
      SQL = SQL & " [VALOR_ATACADO] [float] NOT NULL,"
      SQL = SQL & " [VALOR_VAREJO] [Float] Not null "
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_TAB_TEMP"
End Sub

Sub MONTA_REL_INVENTARIO()
'On Error GoTo ERRO_TRATA

   Dim DESC_PROD_A      As String
   Dim SALDO_ATUAL_N    As Double
   Dim VALOR_ATACADO_N  As Double
   Dim VALOR_VAREJO_N   As Double
   Dim TAB_PRECO_ID_N   As Integer

   TAB_PRECO_ID_N = 1
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select tabelapreco_id from TABELAPRECO "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TAB_PRECO_ID_N = 0 & TabConsulta.Fields(0).Value

   CONT_N = 0

   SQL = "delete from INVENTARIO "
   SQL = SQL & " where qtd_primeira = 0 and qtd_segunda = 0"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from INVENTARIOREL "
   'SQL = SQL & " where lote_id = " & txtLote.Text
   CONECTA_RETAGUARDA.Execute SQL

   If Trim(txtLOTE.Text) = "" Then _
      If Not IsNumeric(txtLOTE.Text) Then _
         Call txtLote_LostFocus

'==================
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from INVENTARIO WITH (NOLOCK)"
   SQL = SQL & " where numr_lote = " & txtLOTE.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      DESC_PROD_A = ""
      VALOR_ATACADO_N = 0
      VALOR_VAREJO_N = 0
      CODG_PRODUTO_A = ""
      SALDO_ATUAL_N = (TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabConsulta.Fields("produto_id").Value))

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select CODG_PRODUTO,DESCRICAO,preco_atacado,preco_venda,produto_id"
      SQL = SQL & " from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where produto_id = " & TabConsulta.Fields("produto_id").Value
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         DESC_PROD_A = "" & Trim(TabProduto.Fields("descricao").Value)
         VALOR_ATACADO_N = 0 & Trim(TabProduto.Fields("preco_atacado").Value)
         VALOR_VAREJO_N = 0 & Trim(TabProduto.Fields("preco_venda").Value)
         CODG_PRODUTO_A = "" & Trim(TabProduto.Fields("codg_produto").Value)

         If CRITERIO_A = "venda" Then
            If (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(TabProduto.Fields("produto_id").Value, TAB_PRECO_ID_N, 1)) > 0 Then _
               VALOR_VAREJO_N = 0 & (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(TabProduto.Fields("produto_id").Value, TAB_PRECO_ID_N, 1))
            Else
               If (TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(TabProduto.Fields("produto_id").Value, TAB_PRECO_ID_N, 1)) > 0 Then _
                  VALOR_VAREJO_N = 0 & (TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(TabProduto.Fields("produto_id").Value, TAB_PRECO_ID_N, 1))
         End If
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from INVENTARIOREL "
      SQL = SQL & " where lote_id = " & txtLOTE.Text
      SQL = SQL & " and produto_id = " & TabConsulta.Fields("produto_id").Value
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "insert into INVENTARIOREL "
            SQL = SQL & " (LOTE_ID,ESTABELECIMENTO_ID,PRODUTO_ID,SEQ_ID,"
            SQL = SQL & " CODG_PROD,DESC_PROD,SALDO_ATUAL,CONTAGEM1,CONTAGEM2,"
            SQL = SQL & " VALOR_ATACADO,VALOR_VAREJO)"
         SQL = SQL & " values("
            SQL = SQL & txtLOTE.Text                                             'LOTE_ID
            SQL = SQL & "," & TabConsulta.Fields("ESTABELECIMENTO_ID").Value     'ESTABELECIMENTO_ID
            SQL = SQL & "," & TabConsulta.Fields("produto_ID").Value             'PRODUTO_ID
            SQL = SQL & "," & TabConsulta.Fields("SEQ").Value                    'SEQ_ID
            SQL = SQL & ",'" & Trim(CODG_PRODUTO_A) & "'"                        'CODG_PROD
            SQL = SQL & ",'" & Trim(DESC_PROD_A) & "'"                           'DESC_PROD
            SQL = SQL & "," & tpMOEDA(SALDO_ATUAL_N)                             'SALDO_ATUAL
            SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("qtd_primeira").Value)  'CONTAGEM1
            SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("qtd_segunda").Value)   'CONTAGEM2
            SQL = SQL & "," & tpMOEDA(VALOR_ATACADO_N)                           'VALOR_ATACADO
            SQL = SQL & "," & tpMOEDA(VALOR_VAREJO_N)                            'VALOR_VAREJO
         SQL = SQL & " )"
         Else
            SQL = "update INVENTARIOREL set "
               SQL = SQL & " CONTAGEM1 = CONTAGEM1 + " & tpMOEDA(TabConsulta.Fields("qtd_primeira").Value)
               SQL = SQL & ",CONTAGEM2 = CONTAGEM2 + " & tpMOEDA(TabConsulta.Fields("qtd_segunda").Value)
            SQL = SQL & " where lote_id = " & txtLOTE.Text
            SQL = SQL & " and produto_id = " & TabConsulta.Fields("produto_id").Value
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute SQL

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
'====================

   Msg = "Deseja imprimir todos os produtos do estoque ?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select produto_id,CODG_PRODUTO,DESCRICAO,preco_atacado,preco_venda,preco_custo"
      SQL = SQL & " from PRODUTO WITH (NOLOCK)"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabProduto.EOF
         DESC_PROD_A = "" & Trim(TabProduto.Fields("descricao").Value)
         VALOR_ATACADO_N = 0 & Trim(TabProduto.Fields("preco_atacado").Value)
         PRODUTO_ID_N = 0 & Trim(TabProduto.Fields("produto_id").Value)
         VALOR_VAREJO_N = 0
         CODG_PRODUTO_A = "" & Trim(TabProduto.Fields("codg_produto").Value)
         SALDO_ATUAL_N = (TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value))
         SqL2 = ESTABELECIMENTO_ID_N
         DoEvents

         If CRITERIO_A = "venda" Then
            VALOR_VAREJO_N = 0 & Trim(TabProduto.Fields("preco_venda").Value)
            If (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(PRODUTO_ID_N, TAB_PRECO_ID_N, 1)) > 0 Then _
               VALOR_VAREJO_N = 0 & (TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(PRODUTO_ID_N, TAB_PRECO_ID_N, 1))
            Else
               VALOR_VAREJO_N = 0 & Trim(TabProduto.Fields("preco_custo").Value)
               If (TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PRODUTO_ID_N, TAB_PRECO_ID_N, 1)) > 0 Then _
                  VALOR_VAREJO_N = 0 & (TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PRODUTO_ID_N, TAB_PRECO_ID_N, 1))
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from INVENTARIOREL "
         SQL = SQL & " where lote_id = " & txtLOTE.Text
         SQL = SQL & " and produto_id = " & TabProduto.Fields("produto_id").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            SQL = "insert into INVENTARIOREL "
               SQL = SQL & " (LOTE_ID,ESTABELECIMENTO_ID,PRODUTO_ID,SEQ_ID,"
               SQL = SQL & " CODG_PROD,DESC_PROD,SALDO_ATUAL,CONTAGEM1,CONTAGEM2,"
               SQL = SQL & " VALOR_ATACADO,VALOR_VAREJO)"
            SQL = SQL & " values("
               SQL = SQL & txtLOTE.Text                                             'LOTE_ID
               SQL = SQL & "," & ESTABELECIMENTO_ID_N                               'ESTABELECIMENTO_ID
               SQL = SQL & "," & TabProduto.Fields("produto_ID").Value             'PRODUTO_ID
               SQL = SQL & "," & MAX_ID("seq_id", "INVENTARIOrel", "lote_id", txtLOTE.Text, "estabelecimento_id", SqL2) 'SEQ_ID
               SQL = SQL & ",'" & Trim(CODG_PRODUTO_A) & "'"                        'CODG_PROD
               SQL = SQL & ",'" & Trim(Left(DESC_PROD_A, 100)) & "'"                          'DESC_PROD
               SQL = SQL & "," & tpMOEDA(SALDO_ATUAL_N)                             'SALDO_ATUAL
               SQL = SQL & "," & tpMOEDA(0)  'CONTAGEM1
               SQL = SQL & "," & tpMOEDA(0)   'CONTAGEM2
               SQL = SQL & "," & tpMOEDA(VALOR_ATACADO_N)                           'VALOR_ATACADO
               SQL = SQL & "," & tpMOEDA(VALOR_VAREJO_N)                            'VALOR_VAREJO
            SQL = SQL & " )"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONT_N = CONT_N + 1
         lblConta.Caption = CONT_N
         lblConta.ForeColor = vbRed
         lblConta.Refresh

         TabProduto.MoveNext
      Wend
      If TabProduto.State = 1 Then _
         TabProduto.Close
   End If
'======================
   FORMULA_REL = "{INVENTARIO.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   'FORMULA_REL = FORMULA_REL & " and {ESTABELECIMENTO.ESTABELECIMENTO_ID} = " & ESTABELECIMENTO_ID_N
   'FORMULA_REL = FORMULA_REL & " and {PRODUTO.produto_ID} = {ESTOQUE.produto_ID}"

   'If Trim(txtLote.Text) <> "" Then _
      If IsNumeric(txtLote.Text) Then _
         FORMULA_REL = FORMULA_REL & " and {INVENTARIO.numr_lote} = " & Trim(txtLote.Text)

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   FORMULA_REL = ""
   Nome_Relatorio = "inventario.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL_INVENTARIO"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

   txtDescricao.Text = "" & DESC_PRODUTO_A
   txtProduto.Text = "" & CODG_PRODUTO_A
   txtRef.Text = "" & REFERENCIA_A
   txtQtdeEstoque.Text = Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N), strFormatacao3Digitos)

   SQL3 = ESTABELECIMENTO_ID_N
   NUMR_ID_N = MAX_ID("seq", "INVENTARIO", "numr_lote", txtLOTE.Text, "estabelecimento_id", SQL3)
   txtSeq.Text = NUMR_ID_N

   If TabInventario.State = 1 Then _
      TabInventario.Close

   SQL = "select * from INVENTARIO WITH (NOLOCK)"
   SQL = SQL & " where seq = " & txtSeq.Text
   SQL = SQL & " and numr_lote = " & txtLOTE.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabInventario.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabInventario.EOF Then
      txtQtdeEstoque.Text = TabInventario!qtd_anterior
      txtPriCont.Text = TabInventario!QTD_PRIMEIRA
      MsgBox "Produto já consta nesse Lote seqüência = " & TabInventario!SEQ
   End If
   If TabInventario.State = 1 Then _
      TabInventario.Close

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then _
      txtPriCont.Text = QTDE_N

   If Trim(txtPriCont.Text) <> "" Then
      Call txtpricont_KeyPress(13)
      Else: txtPriCont.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
