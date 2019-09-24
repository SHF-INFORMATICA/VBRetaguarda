VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoCompraCadastro 
   Caption         =   "Pedido de Compras"
   ClientHeight    =   8055
   ClientLeft      =   1485
   ClientTop       =   1920
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PEDIDOCOMPRACADASTRO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pedido de Compras"
      TabPicture(0)   =   "PEDIDOCOMPRACADASTRO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "flxGRID"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtValorDig"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtQtdeItensRelacionados"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtQtdeItensPedidos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTotalPedido"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtTotalPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   10560
         TabIndex        =   29
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtQtdeItensPedidos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   6240
         TabIndex        =   28
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtQtdeItensRelacionados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2400
         TabIndex        =   27
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   8888
         TabIndex        =   23
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00400000&
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   12015
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
            Height          =   375
            Left            =   5760
            Picture         =   "PEDIDOCOMPRACADASTRO.frx":5C2E
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Cadastro Produto"
            Top             =   1200
            Width           =   450
         End
         Begin VB.TextBox txtRef 
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
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdMata 
            Caption         =   "&Retirar itens zerados"
            Height          =   375
            Left            =   9600
            TabIndex        =   30
            Top             =   1680
            Width           =   2175
         End
         Begin VB.ComboBox cmbSituacao 
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
            Left            =   5400
            TabIndex        =   1
            ToolTipText     =   "Selecione a situação para este produto"
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtDtCad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   10320
            TabIndex        =   2
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFornec 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3720
            Picture         =   "PEDIDOCOMPRACADASTRO.frx":B230
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton cmdConsProd 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5280
            Picture         =   "PEDIDOCOMPRACADASTRO.frx":BC32
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPreco 
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
            Height          =   375
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   7
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtPedido 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   0
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtFornecDesc 
            DataField       =   "Nome"
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
            Height          =   375
            Left            =   4200
            MaxLength       =   80
            TabIndex        =   12
            Top             =   720
            Width           =   7575
         End
         Begin VB.TextBox txtProdutoDesc 
            DataField       =   "Nome"
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
            Height          =   375
            Left            =   6240
            MaxLength       =   80
            TabIndex        =   11
            Top             =   1200
            Width           =   5535
         End
         Begin VB.TextBox txtQtde 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   6
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1680
            Width           =   1935
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
            Height          =   375
            Left            =   3360
            MaxLength       =   15
            TabIndex        =   5
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   1200
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txtCNPJCPF 
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
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
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ref./Código:"
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
            Index           =   5
            Left            =   600
            TabIndex        =   31
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Situação:"
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
            Height          =   255
            Left            =   4440
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblpreco 
            Alignment       =   1  'Right Justify
            Caption         =   "Preço Mercadoria:"
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
            Left            =   4560
            TabIndex        =   17
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº Pedido:"
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
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Fornecedor:"
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
            Left            =   600
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Dt.Pedido:"
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
            Left            =   9240
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblquantidade 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantidade:"
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
            Left            =   600
            TabIndex        =   13
            Top             =   1680
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxGRID 
         Height          =   4095
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   7223
         _Version        =   393216
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "ValorTotalPedido = "
         Height          =   240
         Index           =   2
         Left            =   8580
         TabIndex        =   26
         Top             =   6840
         Width           =   1710
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "QtdeItensPedidos = "
         Height          =   240
         Index           =   1
         Left            =   4425
         TabIndex        =   25
         Top             =   6840
         Width           =   1785
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "QtdeItensRelacionados = "
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   6840
         Width           =   2250
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
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
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "sair"
            Description     =   "Sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "exclui"
            Object.ToolTipText     =   "Exclui itens"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Impressão"
            Key             =   "print"
            Object.ToolTipText     =   "Imprime Contagem"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consulta"
            Object.ToolTipText     =   "Consultar Produto"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar Pedido"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "clear"
            Object.ToolTipText     =   "Limpa a Tela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "At.Produto"
            Key             =   "produto"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
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
         Left            =   10440
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   11400
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
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":C634
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":DA5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":F159
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":10FD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":120DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":13343
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":143D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":159A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":16B3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":17D6D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10920
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   23
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":19009
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":19683
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":19CFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":1A377
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":1A511
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":1ADEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOCOMPRACADASTRO.frx":1B6C5
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
      DesignWidth     =   12270
      DesignHeight    =   8055
   End
End
Attribute VB_Name = "frmPedidoCompraCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private LastRow         As Long ' Ultima linha em que se editou
   Private LastCol         As Long ' ultima coluna em que se editou
   Private ControlVisible  As Boolean
   Dim strSQL              As String
   Dim LINHA_N             As Long
   Dim ULT_CLIK_LINHA_N    As Long
   Dim COLUNA_ATUAL_N      As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name
   txtDtCad.Text = Format(Date, "dd/mm/yyyy")

   cmbSituacao.Clear
   cmbSituacao.AddItem "Ativo"
   cmbSituacao.AddItem "Encerrado"
   cmbSituacao.AddItem "Cancelado"
   cmbSituacao.Text = "Ativo"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "sair"
         Unload Me
      Case "print"
         If Trim(txtPedido.Text) <> "" Then
            FORMULA_REL = "{PEDIDOCOMPRA.PEDIDOCOMPRA_id} = " & Trim(txtPedido.Text)

            If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "Pedido_Compra.rpt"
            frmRELATORIO10.Show 1
         End If
      Case "gravar"
         If FORNEC_ID_N > 0 And Trim(txtPedido.Text) <> "" Then
            GRAVA_TUDO
            txtPedido.SetFocus
            LIMPA_TUDO
         End If
      Case "clear"
         LIMPA_TUDO
      Case "exclui"
         If Trim(txtPedido.Text) <> "" Then
            Msg = " Confirma Exclusão do pedido de compra ?"
            PERGUNTA Msg, vbYesNo + 32, "Pedido de Compra", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "delete from PEDIDOCOMPRAITEM where PEDIDOCOMPRAITEM.PEDIDOCOMPRA_id = " & txtPedido.Text
               CONECTA_RETAGUARDA.Execute SQL
               SQL = "delete from PEDIDOCOMPRA where PEDIDOCOMPRA.PEDIDOCOMPRA_id = " & txtPedido.Text
               CONECTA_RETAGUARDA.Execute SQL
               LIMPA_TUDO
               txtPedido.SetFocus
            End If
         End If
      Case "consulta"
         SQL3 = ""
         frmPedidoCompraConsulta.Show 1
         txtPedido.Text = SQL3
         SQL3 = ""
      Case "exclui_lote"
      Case "produto"
         ATUALIZA_PRECO_PRODUTO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtPedido.Text) = "" Then _
         txtPedido.Text = GERA_NUMR_PEDIDO_COMPRA

      SETA_GRID

      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyPress"
End Sub

Private Sub cmdFornec_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   CRITERIO_A = ""
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
   End If
   txtCNPJCPF.SetFocus
   CNPJCPF_A = ""
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdFornec_Click"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         CRITERIO_A = ""
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
         txtCNPJCPF.SetFocus
         CNPJCPF_A = ""
         CRITERIO_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) = "" Then _
         Exit Sub

      If VALIDA_CNPJCPF(Trim(txtCNPJCPF.Text)) = False Then _
         Exit Sub

      KeyAscii = 0
      CRITERIO_A = txtCNPJCPF.Text

      If Trim(CRITERIO_A) <> "" Then
         If Len(txtCNPJCPF.Text) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
         txtCNPJCPF.Text = CRITERIO_A
      End If
      FORNEC_ID_N = 0

      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close
'AQUI SÓ ESTA CONSULTANDO CADASTRO DO FORNECEDOR
      SQL = "select * from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & Trim(CRITERIO_A) & "'"
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabFornecedor.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção."
         txtCNPJCPF.SetFocus
         Exit Sub
         Else
            txtFornecDesc.Text = Trim(TabFornecedor.Fields("descricao").Value)
            FORNEC_ID_N = Trim(TabFornecedor.Fields("fornecedor_id").Value)
      End If
      CRITERIO_A = ""
      If FORNEC_ID_N > 0 And Trim(txtPedido.Text) <> "" Then _
         CARREGA_ITENS_FORNECEDOR

      txtRef.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub cmdCadProd_Click()
   frmCADASTROPRODUTO.Show 1
End Sub

Private Sub cmdConsProd_Click()
   CONSULTA_PRODUTO
   txtProduto.SetFocus
End Sub

Private Sub txtRef_GotFocus()
'On Error GoTo ERRO_TRATA

   txtRef.SelStart = 0
   txtRef.SelLength = Len(txtRef)
   txtRef.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtref_GotFocus"
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtRef.Text) <> "" Then
         If FORNEC_ID_N <= 0 Then _
            TXTCNPJCPF_KeyPress (13)

         LE_REFERENCIA
         txtProduto.SetFocus
         Else
            txtProduto.SetFocus
            Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtref_KeyPress"
End Sub

Private Sub txtref_LostFocus()
   txtRef.BackColor = &HFFFFFF
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtProduto.Text = "" Then
      txtProduto.Text = 0
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_GotFocus"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

   If KeyAscii = 13 Then
      If Trim(txtPedido.Text) = "" Then
         MsgBox "Para Digitar os Produtos gere o numero do Pedido!"
         txtPedido.SetFocus
         Exit Sub
      End If

      PRODUTO_ID_N = 0

      If Trim(txtProduto.Text) <> "" Then
         KeyAscii = 0
         PROCESSA_DADOS_PRODUTOS
         Else
            txtCNPJCPF.SetFocus
            Exit Sub
      End If
      txtQTDE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CONSULTA_PRODUTO
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtqtde_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtQTDE.Text) <> "" Then _
         txtPreco.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtqtde_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
'On Error GoTo ERRO_TRATA

   QTDE_N = 0 & txtQTDE.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_LostFocus"
End Sub

Private Sub txtpreco_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtPreco.Text = "" Then
      txtPreco.Text = 0
      txtPreco.SelStart = 0
      txtPreco.SelLength = Len(txtPreco.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpreco_GotFocus"
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'If FORNEC_ID_N > 0 And PRODUTO_ID_N > 0 And QTDE_N > 0 And Trim(txtPedido.Text) <> "" Then
      If FORNEC_ID_N > 0 And PRODUTO_ID_N > 0 And Trim(txtPedido.Text) <> "" Then

         GRAVA_TUDO
         LIMPA_BODY
         txtProduto.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpreco_KeyPress"
End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtValorDig.Text) <> "" Then
      txtValorDig.SelStart = 0
      txtValorDig.SelLength = Len(txtValorDig.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_GotFocus"
End Sub

Private Sub txtValorDig_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         OcultarControles
         flxGRID.SetFocus
      Case vbKeyUp
         OcultarControles
         'move para a cima celula.
         With flxGRID
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
         With flxGRID
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
      If LastCol > 10 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           flxGRID.SetFocus
           Exit Sub
         End If
      End If

      Dim QTDE_RETIDO_ESTORNO As Double

      QTDE_RETIDO_ESTORNO = 0 & flxGRID.TextMatrix(flxGRID.Row, LastCol)

      AtribuiValorCelula
      'ProximaCelula
      OcultarControles

      QTDE_N = 0 & flxGRID.TextMatrix(LastRow, 11)
      VALOR_ITEM_N = 0 & flxGRID.TextMatrix(LastRow, 12)
      PRODUTO_ID_N = "" & Trim(flxGRID.TextMatrix(LastRow, 7))

      'VALOR_ITEM_N = 0 & txtValorDig.Text

      If TaBPedidoCompraItem.State = 1 Then _
         TaBPedidoCompraItem.Close

      SQL = "select PEDIDOCOMPRA_ID from PEDIDOCOMPRAITEM WITH (NOLOCK)"
      SQL = SQL & " where produto_id = " & PRODUTO_ID_N
      SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
      TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TaBPedidoCompraItem.EOF Then
         SQL = "update PEDIDOCOMPRAITEM set "
            SQL = SQL & " qtde = " & tpMOEDA(QTDE_N)
            SQL = SQL & ",preco = " & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",dt_altera = '" & Now & "'"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TaBPedidoCompraItem.State = 1 Then _
         TaBPedidoCompraItem.Close

      MOSTRA_TOTAIS

      With flxGRID
         If .Row + 1 < .Rows Then
            .Row = .Row + 1
            '.Col = 0
            Else
               .Row = 1
               '.Col = 0
         End If
      End With
      txtValorDig.Text = ""
      SETA_GRID
      'FLXGRID.TextMatrix(LastRow, LastCol)
      'FLXGRID.CellForeColor = vbRed
      'FLXGRID.CellBackColor = vbRed
      flxGRID.SetFocus
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            'ControlVisible = False
            flxGRID.SetFocus
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

Private Sub cmdMata_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) <> "" Then
      If IsNumeric(txtPedido.Text) Then
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select PEDIDOCOMPRA_ID from PEDIDOCOMPRA WITH (NOLOCK)"
         SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            Msg = "Confirma Exclusão do(s) item(ns) com sem qtde informada ?"
            PERGUNTA Msg, vbYesNo + 32, "Pedido de Compra", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "delete from PEDIDOCOMPRAITEM "
               SQL = SQL & " where PEDIDOCOMPRA_id = " & TabConsulta.Fields(0).Value
               SQL = SQL & " and qtde <= 0"
               CONECTA_RETAGUARDA.Execute SQL
               SETA_GRID
            End If
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMata_Click"
End Sub

Private Sub FLXGRID_Click()
'On Error GoTo ERRO_TRATA

   LINHA_N = 0 & flxGRID.RowSel
   COLUNA_ATUAL_N = 0 & flxGRID.Col

   flxGRID.Row = ULT_CLIK_LINHA_N
   flxGRID.Col = 9 'lCol
   flxGRID.CellBackColor = vbWhite

   flxGRID.Row = LINHA_N
   flxGRID.Col = 9 'lCol
   flxGRID.CellBackColor = vbGreen

   ULT_CLIK_LINHA_N = LINHA_N
   flxGRID.Col = COLUNA_ATUAL_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FLXGRID_Click"
End Sub

Private Sub FLXGRID_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = flxGRID.Row
   LastCol = flxGRID.Col

   OcultarControles

   ExibirCelula

   txtProduto.Text = "" & flxGRID.TextMatrix(LastRow, 0)

   txtProduto.Enabled = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FLXGRID_DblClick"
End Sub

Private Sub FLXGRID_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(flxGRID.TextMatrix(flxGRID.Row, 7)) Then       'produto_id
            If Trim(txtPedido.Text) <> "" Then
               Msg = " Confirma Exclusão do item " & Trim(flxGRID.TextMatrix(flxGRID.Row, 9)) & " ?"
               PERGUNTA Msg, vbYesNo + 32, "Pedido de Compra", "DEMO.HLP", 1000
               If RESPOSTA = vbNo Then _
                  Exit Sub

               PRODUTO_ID_N = flxGRID.TextMatrix(flxGRID.Row, 7)

               SQL = "delete PEDIDOCOMPRAitem "
               SQL = SQL & " where produto_id = " & PRODUTO_ID_N
               SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
               CONECTA_RETAGUARDA.Execute SQL

               SETA_GRID
               flxGRID.Refresh
               flxGRID.SetFocus
            End If
         End If
      Case vbKeyF12
         'frmobs.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FLXGRID_KeyDown"
End Sub

Private Sub FLXGRID_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "FLXGRID_KeyPress"
End Sub
'==================================
Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtPedido.Text = ""
   txtDtCad.Text = "##/##/####"
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtFornecDesc.Text = ""
   flxGRID.Clear
   LIMPA_BODY
   txtQtdeItensRelacionados.Text = ""
   txtQtdeItensPedidos.Text = ""
   txtTotalPedido.Text = ""
   LINHA_N = 0
   LastRow = 0
   LastCol = 0
   PRODUTO_ID_N = 0
   QTDE_N = 0
   txtPedido.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   txtProdutoDesc.Text = ""
   txtQTDE.Text = ""
   txtPreco.Text = ""
   PRODUTO_ID_N = 0
   QTDE_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub ExibirCelula()
'On Error GoTo ERRO_TRATA

   Static OK As Boolean

   If flxGRID.Col >= 11 And flxGRID.Col <= 12 Then
      ' Se for celula fixa , sair
      If flxGRID.Col <= flxGRID.FixedCols - 1 Or flxGRID.Row <= flxGRID.FixedRows - 1 Then _
         Exit Sub
   
      If OK Then _
         Exit Sub

      OK = True

      OcultarControles

      LastRow = flxGRID.Row
      LastCol = flxGRID.Col

      Select Case LastCol
         Case Else
            txtValorDig.Move flxGRID.CellLeft - Screen.TwipsPerPixelX, flxGRID.CellTop + flxGRID.Top - Screen.TwipsPerPixelY, flxGRID.CellWidth + Screen.TwipsPerPixelX * 2, flxGRID.CellHeight + Screen.TwipsPerPixelY * 2
            'txtValorDig.Text = FLXGRID.Text

            'If Len(FLXGRID.Text) = 0 Then _
               If LastRow > 1 Then _
                  txtValorDig.Text = FLXGRID.TextMatrix(LastRow - 1, LastCol)

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

Private Sub ProximaCelula()
'On Error GoTo ERRO_TRATA

   If flxGRID.Col < flxGRID.Cols - 1 Then
      flxGRID.Col = flxGRID.Col + 1
      Else
         flxGRID.Col = 1
         If flxGRID.Row < flxGRID.Rows - 1 Then
             flxGRID.Row = flxGRID.Row + 1
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
      Case 11 To 12
         texto = txtValorDig.Text

         If LastCol = 11 Then
            flxGRID.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
            Else: flxGRID.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao2Digitos)
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   'Ocultar o controle textbox
   txtValorDig.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then
      If Trim(txtPedido.Text) = "" Then _
         Exit Sub
      If Not IsNumeric(txtPedido.Text) Then _
         Exit Sub

      CONT_N = 0

      Dim Coluna, Linha, Largura_Campo

      flxGRID.Clear

      flxGRID.Gridlines = flexGridFlat
      flxGRID.FixedRows = 1
      flxGRID.FixedCols = 1
      flxGRID.ScrollBars = flexScrollBarBoth
      flxGRID.AllowUserResizing = flexResizeColumns

      'FLXGRID.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
      'FLXGRID.Rows = 2                   ' Número de linhas(com cabecalho)

      ' define linhas fixas igual a uma e não usa colunas fixas
      flxGRID.Rows = 2
      'FLXGRID.FixedRows = 3
      flxGRID.FixedCols = 0

      If TaBPedidoCompraItem.State = 1 Then _
         TaBPedidoCompraItem.Close

      strSQL = "select PEDIDOCOMPRA.PEDIDOCOMPRA_ID, PEDIDOCOMPRA.ESTABELECIMENTO_ID, "
      strSQL = strSQL & " PEDIDOCOMPRA.FORNECEDOR_ID, PEDIDOCOMPRA.USUARIO_ID, PEDIDOCOMPRA.DT_CADASTRO, "
      strSQL = strSQL & " PEDIDOCOMPRA.SITUACAO, PEDIDOCOMPRAITEM.PEDIDOCOMPRAITEM_ID, "
      strSQL = strSQL & " PEDIDOCOMPRAITEM.PRODUTO_ID, PRODUTO.CODG_PRODUTO AS Codigo, PRODUTO.DESCRICAO, "
      strSQL = strSQL & " Estoque.QTDE_ESTOQUE as QtdeEstoque,"
      strSQL = strSQL & " PEDIDOCOMPRAITEM.QTDE, PEDIDOCOMPRAITEM.PRECO, "
      strSQL = strSQL & " (PEDIDOCOMPRAITEM.QTDE * PEDIDOCOMPRAITEM.PRECO) AS TotalItem, FORNECEDOR.PESSOA_ID, "
      strSQL = strSQL & " PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFornec"

      strSQL = strSQL & " from PEDIDOCOMPRA WITH (NOLOCK) "
      strSQL = strSQL & " INNER JOIN PEDIDOCOMPRAITEM WITH (NOLOCK) "
      strSQL = strSQL & " ON PEDIDOCOMPRA.PEDIDOCOMPRA_ID = PEDIDOCOMPRAITEM.PEDIDOCOMPRA_ID "
      strSQL = strSQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
      strSQL = strSQL & " ON PEDIDOCOMPRAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      strSQL = strSQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
      strSQL = strSQL & " ON PEDIDOCOMPRA.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
      strSQL = strSQL & " INNER JOIN PESSOA WITH (NOLOCK) "
      strSQL = strSQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID "
      strSQL = strSQL & " INNER JOIN ESTOQUE WITH (NOLOCK) "
      strSQL = strSQL & " ON PRODUTO.PRODUTO_ID = ESTOQUE.PRODUTO_ID"

      strSQL = strSQL & " where PEDIDOCOMPRA.PEDIDOCOMPRA_id = " & txtPedido.Text

strSQL = strSQL & " order by dt_altera desc "

      TaBPedidoCompraItem.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TaBPedidoCompraItem.EOF Then

CONT_N = 0 & TaBPedidoCompraItem.Fields.Count

         txtDtCad.Text = "" & Trim(TaBPedidoCompraItem.Fields("DT_CADASTRO").Value)
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = "" & Trim(TaBPedidoCompraItem.Fields("cnpjcpf").Value)
         txtCNPJCPF.PromptInclude = True
         txtFornecDesc.Text = "" & Trim(TaBPedidoCompraItem.Fields("NomeFornec").Value)
         FORNEC_ID_N = Trim(TaBPedidoCompraItem.Fields("fornecedor_id").Value)

         cmbSituacao.Text = ""
         If Not IsNull(TaBPedidoCompraItem.Fields("situacao").Value) Then
            If Trim(TaBPedidoCompraItem.Fields("situacao").Value) <> "" Then
               If Trim(TaBPedidoCompraItem.Fields("situacao").Value) = "A" Then _
                  cmbSituacao.Text = "Ativo"
               If Trim(TaBPedidoCompraItem.Fields("situacao").Value) = "E" Then _
                  cmbSituacao.Text = "Encerrado"
               If Trim(TaBPedidoCompraItem.Fields("situacao").Value) = "C" Then _
                  cmbSituacao.Text = "Cancelado"
            End If
         End If

         ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
         flxGRID.Rows = 1
         flxGRID.Cols = TaBPedidoCompraItem.Fields.Count

         ReDim largura_coluna(0 To TaBPedidoCompraItem.Fields.Count - 1)

         ' exibe os cabeçalhos das colunas
         For Coluna = 0 To TaBPedidoCompraItem.Fields.Count - 1
            flxGRID.TextMatrix(0, Coluna) = Trim(TaBPedidoCompraItem.Fields(Coluna).Name)
            largura_coluna(Coluna) = TextWidth(Trim(TaBPedidoCompraItem.Fields(Coluna).Name))
         Next Coluna

         ' exibe o valor de cada linha
         Linha = 1

         Do While Not TaBPedidoCompraItem.EOF
            flxGRID.Rows = flxGRID.Rows + 1
            
            For Coluna = 0 To TaBPedidoCompraItem.Fields.Count - 1
               'If Coluna = 3 Or Coluna = 7 Then
               If Coluna = 10 Or Coluna = 11 Then
                  flxGRID.TextMatrix(Linha, Coluna) = Format(TaBPedidoCompraItem.Fields(Coluna).Value, strFormatacao3Digitos)
                  Else
                     If Coluna = 12 Or Coluna = 13 Then
                        flxGRID.TextMatrix(Linha, Coluna) = Format(TaBPedidoCompraItem.Fields(Coluna).Value, strFormatacao2Digitos)
                        Else: flxGRID.TextMatrix(Linha, Coluna) = "" & Trim(TaBPedidoCompraItem.Fields(Coluna).Value)
                     End If
               End If

               ' verifica o tamanho dos campos
               If Not IsNull(TaBPedidoCompraItem.Fields(Coluna).Value) Then _
                  Largura_Campo = TextWidth(TaBPedidoCompraItem.Fields(Coluna).Value)

               If largura_coluna(Coluna) < Largura_Campo Then _
                  largura_coluna(Coluna) = Largura_Campo

            Next Coluna

            TaBPedidoCompraItem.MoveNext
            Linha = Linha + 1
         Loop

         'define a largura das colunas do grid
         For Coluna = 0 To flxGRID.Cols - 1
            flxGRID.ColWidth(Coluna) = largura_coluna(Coluna) + 240
         Next Coluna

         flxGRID.ColWidth(0) = 0
         flxGRID.Refresh

         'PEDIDOCOMPRA.PEDIDOCOMPRA_ID
            flxGRID.ColWidth(0) = 0

         'PEDIDOCOMPRA.ESTABELECIMENTO_ID
            flxGRID.ColWidth(1) = 0

         'PEDIDOCOMPRA.FORNECEDOR_ID
            flxGRID.ColWidth(2) = 0

         'PEDIDOCOMPRA.USUARIO_ID
            flxGRID.ColWidth(3) = 0

         'PEDIDOCOMPRA.DT_CADASTRO
            flxGRID.ColWidth(4) = 0

         'PEDIDOCOMPRA.SITUACAO
            flxGRID.ColWidth(5) = 0

         'PEDIDOCOMPRAITEM.PEDIDOCOMPRAITEM_ID
            flxGRID.ColWidth(6) = 0

         'PEDIDOCOMPRAITEM.PRODUTO_ID
            flxGRID.ColWidth(7) = 0
'===================
         'PRODUTO.CODG_PRODUTO
            flxGRID.ColWidth(8) = 1500
            flxGRID.ColAlignment(8) = 0

         'PRODUTO.DESCRICAO
            flxGRID.ColWidth(9) = 6000
            flxGRID.ColAlignment(9) = 0

         'Estoque.QTDE_ESTOQUE"
            flxGRID.ColWidth(10) = 2200
            flxGRID.ColAlignment(10) = 7

         'PEDIDOCOMPRAITEM.QTDE
            flxGRID.ColWidth(11) = 2000
            flxGRID.ColAlignment(11) = 7

         'PEDIDOCOMPRAITEM.PRECO
            flxGRID.ColWidth(12) = 2000
            flxGRID.ColAlignment(12) = 7

         'PEDIDOCOMPRAITEM.PRECO
            flxGRID.ColWidth(13) = 2000
            flxGRID.ColAlignment(13) = 7
'===================
         'FORNECEDOR.PESSOA_ID
            flxGRID.ColWidth(14) = 0

         'FORNECEDOR.cnpjcpf
            flxGRID.ColWidth(15) = 0

         'FORNECEDOR.NOME
            flxGRID.ColWidth(16) = 0
      End If
      If TaBPedidoCompraItem.State = 1 Then _
         TaBPedidoCompraItem.Close
   End If

   MOSTRA_TOTAIS

   flxGRID.Row = 1
   flxGRID.Col = 9 'lCol
   flxGRID.CellBackColor = vbGreen

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   SQL3 = "NULL"
   If Trim(cmbSituacao.Text) = "" Then
      MsgBox "Informar Situação."
      cmbSituacao.SetFocus
      Exit Sub
      Else
         If Trim(Left(cmbSituacao.Text, 1)) = "E" Then _
            SQL3 = Date
   End If

   If TaBCompra.State = 1 Then _
      TaBCompra.Close

   SQL = "select dt_baixa from PEDIDOCOMPRA WITH (NOLOCK)"
   SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TaBCompra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TaBCompra.EOF Then
      SQL = "insert into PEDIDOCOMPRA "
         SQL = SQL & " (PEDIDOCOMPRA_ID,ESTABELECIMENTO_ID,FORNECEDOR_ID,USUARIO_ID,DT_CADASTRO,SITUACAO)"
      SQL = SQL & " values ("
         SQL = SQL & txtPedido.Text                            'PEDIDOCOMPRA_ID
         SQL = SQL & "," & ESTABELECIMENTO_ID_N                'ESTABELECIMENTO_ID
         SQL = SQL & "," & FORNEC_ID_N                         'FORNECEDOR_ID
         SQL = SQL & "," & USUARIO_ID_N                        'USUARIO_ID
         SQL = SQL & ",'" & DMA(txtDtCad.Text) & "'"           'DT_CADASTRO
         SQL = SQL & ",'" & Left(cmbSituacao.Text, 1) & "'"    'SITUACAO
      SQL = SQL & ")"
      Else
         SQL = "update PEDIDOCOMPRA SET"
            SQL = SQL & " situacao = '" & Left(cmbSituacao.Text, 1) & "'"  'SITUACAO
            SQL = SQL & ",USUARIO_ID = " & USUARIO_ID_N                    'USUARIO_ID
            SQL = SQL & ",FORNECEDOR_ID = " & FORNEC_ID_N                  'FORNECEDOR_ID
            SQL = SQL & ",DT_BAIXA = " & SQL3                              'DT_BAIXA
         SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   End If
   If TaBCompra.State = 1 Then _
      TaBCompra.Close

   CONECTA_RETAGUARDA.Execute SQL

   If PRODUTO_ID_N > 0 Then _
      GRAVA_ITEM

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub GRAVA_ITEM()
'On Error GoTo ERRO_TRATA

   If TaBPedidoCompraItem.State = 1 Then _
      TaBPedidoCompraItem.Close

   SQL = "select PEDIDOCOMPRA_ID from PEDIDOCOMPRAITEM WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
   TaBPedidoCompraItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TaBPedidoCompraItem.EOF Then
      SQL = "insert into PEDIDOCOMPRAITEM "
         SQL = SQL & " (PEDIDOCOMPRA_ID,PEDIDOCOMPRAITEM_ID,PRODUTO_ID,PRECO,QTDE) "
      SQL = SQL & " values ("
        SQL = SQL & txtPedido.Text                                                          'PEDIDOCOMPRA_ID
        SQL = SQL & "," & MAX_ID("PEDIDOCOMPRAITEM_ID", "PEDIDOCOMPRAITEM", "", "", "", "") 'PEDIDOCOMPRAITEM_ID
        SQL = SQL & "," & PRODUTO_ID_N                                                      'PRODUTO_ID
        SQL = SQL & "," & tpMOEDA(txtPreco.Text)                                            'PRECO
        SQL = SQL & "," & tpMOEDA(txtQTDE.Text)                                             'QTDE
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "update PEDIDOCOMPRAITEM set "
            'SQL = SQL & " qtde = " & tpMOEDA(QTDE_N)
            'SQL = SQL & ",preco = " & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & " dt_altera = '" & Now & "'"
         SQL = SQL & " where produto_id = " & PRODUTO_ID_N
         SQL = SQL & " and PEDIDOCOMPRA_id = " & txtPedido.Text
   End If
   If TaBPedidoCompraItem.State = 1 Then _
      TaBPedidoCompraItem.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITEM"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then _
      txtProduto.Text = SQL3
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Sub CARREGA_ITENS_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   'SQL = "select produto_id,codg_produto,descricao,preco_custo from PRODUTO WITH (NOLOCK)"

   SQL = "select PRODUTO.PRODUTO_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
   SQL = SQL & " PRODUTOFORNECEDOR.FORNECEDOR_ID, PRODUTOFORNECEDOR.CODG_PROD_FORNEC, "
   SQL = SQL & " PRODUTOFORNECEDOR.PRECO_CUSTO, PRODUTOFORNECEDOR.CODG_BARRA, "
   SQL = SQL & " FORNECEDOR.PESSOA_ID, PESSOA.CNPJCPF, PESSOA.DESCRICAO AS NomeFornec"
   SQL = SQL & " from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTOFORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOFORNECEDOR.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTOFORNECEDOR.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where produto.situacao = 'A' "
   SQL = SQL & " and PRODUTOFORNECEDOR.fornecedor_id = " & FORNEC_ID_N

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      Beep
      Msg = "Deseja Carregar os Itens Cadastrados Para Este Fornecedor: " & Trim(txtFornecDesc.Text) & " ?"
      Style = vbYesNo + 32
      Title = "Atenção."
      Help = "DEMO.HLP"
      Ctxt = 1000
      RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
      If RESPOSTA = vbNo Then
         If TabProduto.State = 1 Then _
            TabProduto.Close
         Exit Sub
      End If

      If TaBCompra.State = 1 Then _
         TaBCompra.Close

      SQL = "select dt_baixa from PEDIDOCOMPRA WITH (NOLOCK)"
      SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TaBCompra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TaBCompra.EOF Then
         SQL = "insert into PEDIDOCOMPRA "
            SQL = SQL & " (PEDIDOCOMPRA_ID,ESTABELECIMENTO_ID,FORNECEDOR_ID,USUARIO_ID,DT_CADASTRO,SITUACAO)"
         SQL = SQL & " values ("
            SQL = SQL & txtPedido.Text                            'PEDIDOCOMPRA_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                'ESTABELECIMENTO_ID
            SQL = SQL & "," & FORNEC_ID_N                         'FORNECEDOR_ID
            SQL = SQL & "," & USUARIO_ID_N                        'USUARIO_ID
            SQL = SQL & ",'" & DMA(txtDtCad.Text) & "'"           'DT_CADASTRO
            SQL = SQL & ",'A'"    'SITUACAO
         SQL = SQL & ")"
         Else
            SQL = "update PEDIDOCOMPRA SET"
               SQL = SQL & " situacao = 'A'"  'SITUACAO
               SQL = SQL & ",USUARIO_ID = " & USUARIO_ID_N                    'USUARIO_ID
               SQL = SQL & ",FORNECEDOR_ID = " & FORNEC_ID_N                  'FORNECEDOR_ID
            SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      End If
      If TaBCompra.State = 1 Then _
         TaBCompra.Close

      CONECTA_RETAGUARDA.Execute SQL

      While Not TabProduto.EOF
         DoEvents
         PRODUTO_ID_N = TabProduto.Fields("produto_id").Value
         txtQTDE.Text = 0
         txtPreco.Text = TabProduto.Fields("preco_custo").Value

         GRAVA_ITEM

         TabProduto.MoveNext
      Wend
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   PRODUTO_ID_N = 0
   txtQTDE.Text = ""
   txtPreco.Text = ""

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_ITENS_FORNECEDOR"
End Sub

Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   txtQtdeItensRelacionados.Text = ""
   txtQtdeItensPedidos.Text = ""
   txtTotalPedido.Text = ""

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select count(PEDIDOCOMPRA_ID) from PEDIDOCOMPRAITEM WITH (NOLOCK)"
   SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         txtQtdeItensRelacionados.Text = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum(qtde) from PEDIDOCOMPRAITEM WITH (NOLOCK)"
   SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         txtQtdeItensPedidos.Text = Format(TabConsulta.Fields(0).Value, strFormatacao3Digitos)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum(qtde*preco) from PEDIDOCOMPRAITEM WITH (NOLOCK)"
   SQL = SQL & " where PEDIDOCOMPRA_id = " & txtPedido.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         txtTotalPedido.Text = Format(TabConsulta.Fields(0).Value, strFormatacao2Digitos)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAIS"
End Sub

Sub LE_REFERENCIA()
'On Error GoTo ERRO_TRATA

   Dim TabProdFornec As New ADODB.Recordset
   PRODUTO_ID_N = 0

   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

   SQL = "select PRODUTOFORNECEDOR.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
   SQL = SQL & " from PRODUTOFORNECEDOR "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PRODUTOFORNECEDOR.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where codg_prod_fornec = '" & Trim(txtRef.Text) & "'"
   SQL = SQL & " and PRODUTOFORNECEDOR.fornecedor_id = " & FORNEC_ID_N
   TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProdFornec.EOF Then
      txtProduto.Text = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
      PRODUTO_ID_N = 0 & Trim(TabProdFornec.Fields("produto_id").Value)
      Else
         If TabProdFornec.State = 1 Then _
            TabProdFornec.Close

         SQL = "select codg_produto,descricao,produto_id,preco_custo,codg_ncm,unidade_medida,aliquota_icms,situacao_tributaria "
         SQL = SQL & " from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where referencia = '" & Trim(txtRef.Text) & "'"
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProdFornec.EOF Then
            txtProduto.Text = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
            PRODUTO_ID_N = 0 & Trim(TabProdFornec.Fields("produto_id").Value)
         End If
   End If
   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_REFERENCIA"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

   txtProdutoDesc.Text = "" & DESC_PRODUTO_A
   txtProduto.Text = "" & CODG_PRODUTO_A
   txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtPreco.Text = Format(PR_VAREJO_N, strFormatacao3Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub

Sub DESTACA_LINHA()
'On Error GoTo ERRO_TRATA

   Dim lCols   As Long
   Dim lRows   As Long
   Dim lCol    As Long
   Dim lRow    As Long
   Dim sColor  As String

   'numero de colunas
   lCols = CONT_N
   With flxGRID
      lRow = LINHA_N

      For lCol = 0 To lCols - 1
        .Col = lCol
        .Row = lRow

        'alterna a cor das linhas do grid
         If (lRow Mod 2) = 0 Then
            .CellBackColor = vbWhite
         Else
            '.CellBackColor = Comdlg.Color
            .CellBackColor = vbGreen
         End If
      Next
   End With

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "DESTACA_LINHA"
End Sub

Sub ATUALIZA_PRECO_PRODUTO()
'On Error GoTo ERRO_TRATA

'SELECT        PEDIDOCOMPRA.SITUACAO, PEDIDOCOMPRAITEM.*
'FROM            PEDIDOCOMPRA INNER JOIN
'                         PEDIDOCOMPRAITEM ON PEDIDOCOMPRA.PEDIDOCOMPRA_ID = PEDIDOCOMPRAITEM.PEDIDOCOMPRA_ID
 'order by DT_ALTERA desc

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_PRECO_PRODUTO"
End Sub
