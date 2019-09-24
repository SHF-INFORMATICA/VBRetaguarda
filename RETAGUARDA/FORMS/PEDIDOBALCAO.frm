VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPEDIDOBALCAO 
   Caption         =   "Pedido Venda"
   ClientHeight    =   8115
   ClientLeft      =   2085
   ClientTop       =   2475
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "PEDIDOBALCAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtValorDig 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   56
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtTotalPedido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   5640
      TabIndex        =   53
      Top             =   7380
      Width           =   1815
   End
   Begin VB.TextBox txtItens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   7680
      TabIndex        =   51
      Top             =   7380
      Width           =   1455
   End
   Begin VB.TextBox txtDescontoRodape 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   525
      Left            =   3840
      TabIndex        =   49
      Top             =   7380
      Width           =   1455
   End
   Begin VB.TextBox txtVlrUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   2040
      TabIndex        =   47
      Top             =   7380
      Width           =   1455
   End
   Begin VB.TextBox txtQtdeDisp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   120
      TabIndex        =   45
      Top             =   7380
      Width           =   1455
   End
   Begin VB.TextBox txtPesoTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   525
      Left            =   9360
      TabIndex        =   43
      Top             =   7380
      Width           =   1455
   End
   Begin VB.Frame FraReq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   10935
      Begin PVMaskEditLib.PVMaskEdit txtCNPJCPF 
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _Version        =   524288
         _ExtentX        =   3413
         _ExtentY        =   635
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BorderStyle     =   1
         Text            =   ""
      End
      Begin VB.ComboBox cmbDesconto 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   9360
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbFormaAUX 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8400
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbForma 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8400
         TabIndex        =   60
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3050
         Picture         =   "PEDIDOBALCAO.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3540
         MaxLength       =   100
         TabIndex        =   35
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtLIMITE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1920
         TabIndex        =   34
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtPAGAR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   4320
         TabIndex        =   33
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cmbTabPrecoAux 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8400
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTabPreco 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8400
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbVendAux 
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
         Left            =   8400
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbVend 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8400
         TabIndex        =   0
         ToolTipText     =   "Selecione um vendedor"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDtEmis 
         Height          =   360
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
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
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Forma:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7560
         TabIndex        =   61
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "*Cliente:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Crédito:"
         Height          =   240
         Left            =   1020
         TabIndex        =   37
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "À Pagar:"
         Height          =   240
         Left            =   3420
         TabIndex        =   36
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Pedido:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   2505
         TabIndex        =   30
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "TabPr:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "*Vendedor:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "*Pedido:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FraSeq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   10935
      Begin VB.CommandButton cmdDetalhe 
         BackColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   10320
         Picture         =   "PEDIDOBALCAO.frx":6614
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Registrar Detalhes"
         Top             =   840
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtPesoItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   9000
         TabIndex        =   41
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   10440
         TabIndex        =   40
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdMata 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   4040
         Picture         =   "PEDIDOBALCAO.frx":691E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtPreçoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   9840
         MaxLength       =   12
         TabIndex        =   32
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3590
         Picture         =   "PEDIDOBALCAO.frx":775F
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtVarejo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3720
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAtacado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1320
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optValor 
         Caption         =   "R$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9120
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optPerc 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtQTDE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8400
         TabIndex        =   7
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   360
         Left            =   4530
         MaxLength       =   29
         TabIndex        =   9
         Top             =   240
         Width           =   6255
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtValor_Unitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6360
         MaxLength       =   12
         TabIndex        =   6
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   360
         Left            =   8040
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Se houver algum desconto informe aqui. Pode ser em valor ou em percentual."
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCusto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Custo"
         Height          =   240
         Left            =   10260
         TabIndex        =   63
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Peso Item"
         Height          =   240
         Left            =   9885
         TabIndex        =   42
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Vlr.Varejo:"
         Height          =   240
         Left            =   2610
         TabIndex        =   23
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Atacado"
         Height          =   240
         Left            =   105
         TabIndex        =   22
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desc."
         Height          =   240
         Left            =   8100
         TabIndex        =   21
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Qtde:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7815
         TabIndex        =   18
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Vlr.Unitário:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5115
         TabIndex        =   17
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Produto:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1270
      ButtonWidth     =   3201
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cad. Cliente"
            Key             =   "CadCliente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Cad.Produto"
            Key             =   "produto"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Caixa"
            Key             =   "receber"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clonar"
            Key             =   "clonar"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   10200
         TabIndex        =   58
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
         Top             =   240
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
               Picture         =   "PEDIDOBALCAO.frx":8161
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":92FB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":A38A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":B33F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":C44A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":D5A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":D9F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":F869
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":10F1F
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAO.frx":12F01
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   915
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   1680
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10980
      DesignHeight    =   8115
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   55
      Top             =   3480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6165
      _Version        =   393216
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   9240
      X2              =   9240
      Y1              =   7080
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   7080
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   5520
      X2              =   5520
      Y1              =   7080
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   3720
      X2              =   3720
      Y1              =   7080
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   1800
      X2              =   1800
      Y1              =   7080
      Y2              =   7920
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Pedido"
      Height          =   240
      Left            =   6135
      TabIndex        =   54
      Top             =   7095
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Itens Pedido"
      Height          =   240
      Left            =   7965
      TabIndex        =   52
      Top             =   7095
      Width           =   1185
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Desconto"
      Height          =   240
      Left            =   4440
      TabIndex        =   50
      Top             =   7095
      Width           =   870
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor Unitário"
      Height          =   240
      Left            =   2190
      TabIndex        =   48
      Top             =   7095
      Width           =   1320
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "QtdeDisponível"
      Height          =   240
      Left            =   150
      TabIndex        =   46
      Top             =   7095
      Width           =   1440
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Peso Total (Kg)"
      Height          =   240
      Left            =   9390
      TabIndex        =   44
      Top             =   7095
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      Height          =   990
      Left            =   0
      Top             =   7080
      Width           =   10935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "*Vendedor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmPEDIDOBALCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim strInscEstadual        As String
   Dim dblTipoCliente         As Double
   Dim strCPFCNPJ             As String
   Dim bolRequisicaoJaExiste  As Boolean
   Dim rstEmpresa             As New ADODB.Recordset
   Dim Seq_N                  As Long
   Dim PRECO_PROD             As Double
   Dim CLIENTE_ID_N           As Long
   Dim TIPO_NOTA_A            As String
   Dim VALOR_UNITARIO_N       As Double
   Dim SITUAÇÃO_TRIBUTARIA_PRODUTO
   Dim INDR_PROD_BALANCA      As Boolean
   Dim Valr_Venda_Produto_n   As Double
   Dim PESO_ITEM_N            As Double
   Dim TabGridVaca            As New ADODB.Recordset
   Private LastRow            As Long ' Ultima linha em que se editou
   Private LastCol            As Long ' ultima coluna em que se editou
   Private ControlVisible     As Boolean
   Dim PRECO_CUSTO_N          As Double
   Dim INDR_TRAVA_TABELA      As Boolean
   Dim VALR_ATACADO_N         As Double
   Dim Valr_Digitado          As Double
   Dim VALR_VENDA_N             As Double
   Dim VALOR_VENDA_ORIGINAL_N As Double

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INICIALIZA_VENDA
   MOSTRA_VENDEDORES

   Call txtPedido_LostFocus

   OcultarControles
   lblCusto.Visible = False
   Label12.Visible = False
   txtPesoItem.Visible = False
   txtDesconto.Visible = False
   Label8.Visible = False
   optPerc.Visible = False
   optValor.Visible = False
   If USUARIO_ID_N = 144 Or USUARIO_ID_N = 1 Then
      'lblCusto.Visible = True
      txtPreçoCusto.Visible = True
   End If
   Toolbar1.Buttons(2).Visible = False
   If LIMPA_PEDIDO = True Then _
      Toolbar1.Buttons(2).Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF8
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1

         If NOME_A <> "" Then _
            txtNome.Text = NOME_A
         NOME_A = ""
      Case vbKeyF10
         INDR_GRAVA = False
         If Trim(txtPedido.Text) = "" Then _
            Exit Sub
         If Not IsNumeric(txtPedido.Text) Then _
            Exit Sub

         PEDIDO_ID_N = txtPedido.Text

         GERA_VENDA

         LIMPA_TUDO

         Call txtPedido_LostFocus
      Case vbKeyF11
      
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "clonar"
         CLONA_PEDIDO_VENDA
         
      Case "receber"
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
            frmDISPLAYEMISSOR.Show 1
            VALIDA_PEDIDO
            INICIALIZA_VENDA
         End If
      Case "gravar"
         INDR_GRAVA = False
         If txtPedido.Text <> "" Then
            PEDIDO_ID_N = txtPedido.Text
            Else
               MsgBox "Digite Numero da Requisicao para gravar!"
               Exit Sub

               Call txtPedido_LostFocus
         End If

         GERA_VENDA
         LIMPA_TUDO

         Call txtPedido_LostFocus
      Case "consultar"
         CRITERIO_A = ""
         CNPJCPF_A = ""
         frmPedidoConsulta.Show 1
         If PEDIDO_ID_N > 0 Then
            Dim NUMR_PEDIDO_N As Long

            NUMR_PEDIDO_N = PEDIDO_ID_N

            LIMPA_TUDO
            txtPedido.Text = NUMR_PEDIDO_N
            CRITERIO_A = ""
            NUMR_PEDIDO_N = 0
            Call txtPedido_LostFocus
            VALIDA_PEDIDO
         End If
         FraSeq.Enabled = True
         'txtProduto.SetFocus
      Case "print"
         GERA_IMPRESSAO
      Case "limpar"
         LIMPA_TUDO

         Call txtPedido_LostFocus
         FraSeq.Enabled = True
         txtProduto.SetFocus
      Case "voltar"
         Unload Me
      Case "produto"
         'If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
         If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
            frmCADASTROPRODUTO.Show 1
            Else: CHAMA_PRODUTO_SIMPLIFICADO
         End If
      Case "CadCliente"
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdDetalhe_Click()
'On Error GoTo ERRO_TRATA

   'pedido_id
   If IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then _
      Exit Sub
   'seq_id
   If IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then _
      Exit Sub
   'produto_id
   If IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then _
      Exit Sub

   'pedido_id
   If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then
      'seq_id
      If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then
         PEDIDO_ID_N = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
         SEQ_ID_N = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
         PRODUTO_ID_N = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select pedido_id,seq_id from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and seq_id = " & SEQ_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      'pedido_id
      If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then
         'seq_id
         If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then
            PEDIDO_ID_N = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
            SEQ_ID_N = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
            PRODUTO_ID_N = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)
   
            frmPEDIDODETALHE.Show 1
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdDetalhe_Click"
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Text = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = ""
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.Text = CNPJCPF_A
      Call txtCNPJCPF_LostFocus
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsCli_Click"
End Sub

Private Sub cmdConsProd_Click()
   CONSULTA_PRODUTO
   FraSeq.Enabled = True
   txtProduto.SetFocus
End Sub

Private Sub cmdMata_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then
      EXCLUIR_ITEM Trim(txtProduto.Text), Trim(txtPedido.Text), Trim(txtSeq.Text)
      Else: MsgBox "Informe código produto."
   End If
   FraSeq.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMata_Click"
End Sub

Private Sub optPerc_GotFocus()
'On Error GoTo ERRO_TRATA

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Else: SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optPerc_GotFocus"
End Sub

Private Sub cmbTabPreco_Click()
'On Error GoTo ERRO_TRATA

   cmbTabPrecoAux.ListIndex = cmbTabPreco.ListIndex
   cmbForma.Visible = False
   cmbForma.Clear
   cmbFormaAUX.Clear
   cmbDesconto.Clear
   CONT_N = 0

   If Trim(cmbVendAux.Text) <> "" And Trim(cmbTabPrecoAux.Text) <> "" Then
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select TABELAPRECO.CODG_TABELA, TABELAPRECO.DESCRICAO, tabelapreco.tabelapreco_id, "
      SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.PRODUTO_ID,"
      SQL = SQL & " TABELAPRECOITEM.FORMAPAGTO_ID, TABELAPRECOITEM.VALOR_VENDA, "
      SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO, TABELAPRECOITEM.PERC_COMISSAO,"
      SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescFormaPagto, FORMAPAGTO.STATUS"
      SQL = SQL & " from TABELAPRECO "
      SQL = SQL & " INNER JOIN TABELAPRECOITEM "
      SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
      SQL = SQL & " INNER JOIN VENDEDOR "
      SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = VENDEDOR.TABELAPRECO_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " INNER JOIN FORMAPAGTO "
      SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
      SQL = SQL & " where valor_venda > 0 "

      SQL = SQL & " and vendedor_id = " & VENDEDOR_ID_N

      If Trim(cmbTabPrecoAux.Text) <> "" Then _
         SQL = SQL & " and codg_tabela = '" & cmbTabPrecoAux.Text & "'"

      SQL = SQL & " order by formapagto_id"

      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

      cmbForma.Clear
      cmbFormaAUX.Clear
      cmbDesconto.Clear
      While Not TabVENDEDOR.EOF
         If CONT_N <> TabVENDEDOR.Fields("formapagto_id").Value Then
            cmbForma.AddItem Trim(TabVENDEDOR.Fields("descformapagto").Value) & "-" & Trim(TabVENDEDOR.Fields("formapagto_id").Value)
            cmbFormaAUX.AddItem Trim(TabVENDEDOR.Fields("formapagto_id").Value)

            CONT_N = TabVENDEDOR.Fields("formapagto_id").Value

            If TabVENDEDOR.Fields("formapagto_id").Value = 1 Then
               cmbForma.Text = "" & Trim(TabVENDEDOR.Fields("descformapagto").Value) & "-" & Trim(TabVENDEDOR.Fields("formapagto_id").Value)
               cmbFormaAUX.Text = "" & Trim(TabVENDEDOR.Fields("formapagto_id").Value)
            End If

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select permite_desconto from TIPOVENDA "
            SQL = SQL & " where formapagto_id = " & cmbFormaAUX.Text
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
                  cmbDesconto.AddItem TabTemp.Fields("permite_desconto").Value
                  Else: cmbDesconto.AddItem ""
               End If
               Else: cmbDesconto.AddItem ""
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
         TabVENDEDOR.MoveNext
      Wend
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      cmbForma.Visible = True

      If Trim(cmbFormaAUX.Text) <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select permite_desconto from TIPOVENDA "
         SQL = SQL & " where formapagto_id = " & Trim(cmbFormaAUX.Text)
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then _
               cmbDesconto.Text = "" & TabTemp.Fields("permite_desconto").Value
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If
   TABELAPRECO_ID_N = 0 & cmbTabPrecoAux.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTabPreco_Click"
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbFormaAUX.ListIndex = cmbForma.ListIndex
   cmbDesconto.ListIndex = cmbForma.ListIndex
   txtProduto.Enabled = True
   FraSeq.Enabled = True
   FORMAPAGTO_ID_N = 0 & cmbFormaAUX.Text
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbforma_Click"
End Sub

Private Sub cmbTabPreco_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Selecione Tabela Preço", "", "", ""
   cmbTabPreco.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTabPreco_GotFocus"
End Sub

Private Sub cmbTabPreco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTabPreco_KeyPress"
End Sub

Private Sub optValor_GotFocus()
'On Error GoTo ERRO_TRATA

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Else: SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optValor_GotFocus"
End Sub

Private Sub optPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDesconto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optPerc_KeyPress"
End Sub

Private Sub optvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDesconto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optvalor_KeyPress"
End Sub

Private Sub cmbForma_GotFocus()
   cmbForma.BackColor = &HC0FFFF
End Sub

Private Sub cmbForma_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbFormaAUX.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select permite_desconto from TIPOVENDA "
      SQL = SQL & " where formapagto_id = " & Trim(cmbFormaAUX.Text)
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
            cmbDesconto.Text = "" & TabTemp.Fields("permite_desconto").Value
            Else: cmbDesconto.Text = ""
         End If
         Else: cmbDesconto.Text = ""
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   cmbForma.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbForma_LostFocus"
End Sub

Private Sub cmbTabPreco_LostFocus()
   cmbTabPreco.BackColor = &HFFFFFF
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex
   cmbTabPreco.Clear
   cmbTabPrecoAux.Clear
   cmbForma.Clear
   cmbFormaAUX.Clear
   cmbDesconto.Clear

   CONT_N = 0

   If Trim(cmbVendAux.Text) <> "" Then
      VENDEDOR_ID_N = cmbVendAux.Text

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select TABELAPRECO.CODG_TABELA, TABELAPRECO.DESCRICAO, tabelapreco.tabelapreco_id, "
      SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.PRODUTO_ID,"
      SQL = SQL & " TABELAPRECOITEM.FORMAPAGTO_ID, TABELAPRECOITEM.VALOR_VENDA, "
      SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO, TABELAPRECOITEM.PERC_COMISSAO,"
      SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescFormaPagto, FORMAPAGTO.STATUS"
      SQL = SQL & " from TABELAPRECO "
      SQL = SQL & " INNER JOIN TABELAPRECOITEM "
      SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
      SQL = SQL & " INNER JOIN VENDEDOR "
      SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = VENDEDOR.TABELAPRECO_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " INNER JOIN FORMAPAGTO "
      SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
      SQL = SQL & " where valor_venda > 0 "

      SQL = SQL & " and vendedor_id = " & VENDEDOR_ID_N

      If Trim(cmbTabPrecoAux.Text) <> "" Then _
         SQL = SQL & " and codg_tabela = '" & cmbTabPrecoAux.Text & "'"
   
      SQL = SQL & " order by formapagto_id"
   
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

      While Not TabVENDEDOR.EOF
         If CONT_N <> TabVENDEDOR.Fields("codg_tabela").Value Then
            cmbTabPreco.AddItem TabVENDEDOR.Fields("descricao").Value
            cmbTabPrecoAux.AddItem TabVENDEDOR.Fields("codg_tabela").Value
            CONT_N = TabVENDEDOR.Fields("codg_tabela").Value
         End If
         TabVENDEDOR.MoveNext
      Wend
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If

   cmbTabPreco.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVend_Click"
End Sub

Private Sub cmbVend_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbVend.Text) = "" Then _
      MOSTRA_VENDEDORES

   MOSTRA_TOP "ESC - SAIR", "Selecione Vendedor e tecle <ENTER>", "", "", ""

   cmbVend.SelStart = 0
   cmbVend.SelLength = Len(cmbVend)
   cmbVend.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVend_GotFocus"
End Sub

Private Sub cmbvend_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTabPreco.SetFocus
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_KeyPress"
End Sub

Private Sub cmbVend_LostFocus()
   cmbVend.BackColor = &HFFFFFF
End Sub

Private Sub txtAtacado_Click()
'On Error GoTo ERRO_TRATA

   If txtAtacado.Text <> "" Then _
      txtValor_Unitario.Text = txtAtacado.Text
   txtValor_Unitario.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAtacado_Click"
End Sub

Private Sub txtAtacado_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtAtacado_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtAtacado.Text) <> "" Then
      If IsNumeric(txtAtacado.Text) Then
         txtAtacado.Text = Format(txtAtacado.Text, strFormatacao2Digitos)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAtacado_LostFocus"
End Sub

Private Sub txtDescontoRodape_GotFocus()
   FraSeq.Enabled = True
   txtProduto.SetFocus
End Sub

Private Sub txtDtEmis_GotFocus()
On Error Resume Next

   cmbVend.SetFocus

Err.Clear
End Sub

Private Sub txtITENS_GotFocus()
'On Error GoTo ERRO_TRATA

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtITENS_GotFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNome.Text) <> "" Then
      txtNome.SelStart = 0
      txtNome.SelLength = Len(txtNome)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtPesoTotal_GotFocus()
'On Error GoTo ERRO_TRATA

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   'txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPesoTotal_GotFocus"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtQtde_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         VALOR_RECEBIDO_N = 0
         VALOR_RECEBIDO_N = 0 & InputBox(VALOR_RECEBIDO_N, "Informe Valor da Venda.")

         If Not IsNull(VALOR_RECEBIDO_N) Then
            If IsNumeric(VALOR_RECEBIDO_N) Then
               If VALOR_RECEBIDO_N > 0 Then

                  If Not IsNull(txtValor_Unitario.Text) Then
                     If IsNumeric(txtValor_Unitario.Text) Then
                        VALOR_ITEM_N = txtValor_Unitario.Text
                        If VALOR_ITEM_N > 0 Then
                           txtQTDE.Text = VALOR_RECEBIDO_N / VALOR_ITEM_N
                           txtQTDE.Refresh
                        End If
                     End If
                  End If

               End If
            End If
         End If

   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_KeyDown"
End Sub

Private Sub txtQtdeDisp_GotFocus()
   FraSeq.Enabled = True
   txtProduto.SetFocus
End Sub

Private Sub txtTotalPedido_GotFocus()
   FraSeq.Enabled = True
   txtProduto.SetFocus
End Sub

Private Sub txtValorDig_LostFocus()
   txtValorDig.BackColor = &HFFFFFF
End Sub

Private Sub txtVarejo_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtVarejo_Click()
'On Error GoTo ERRO_TRATA

   If txtVarejo.Text <> "" Then _
      txtValor_Unitario.Text = txtVarejo.Text
   txtValor_Unitario.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtVarejo_Click"
End Sub

Private Sub txtVarejo_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtVarejo.Text) <> "" Then _
      If IsNumeric(txtVarejo.Text) Then _
         txtVarejo.Text = Format(txtVarejo.Text, strFormatacao2Digitos)


Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtVarejo_LostFocus"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtNome.Text = ""
   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0
   CNPJCPF_A = ""
   ''txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Text = "99999999999"
      txtNome.Enabled = True
      txtNome.Text = "Consumidor Final"
   End If

   If TRATA_PESSOA(txtCNPJCPF.Text) = False Then
      txtCNPJCPF.Text = "99999999999"
      txtNome.Enabled = True
      txtNome.Text = "Consumidor Final"
      Else: txtNome.Text = "" & NOME_CLIENTE_A
   End If

   txtPAGAR.Text = Format(VALOR_PENDENTE_N, strFormatacao2Digitos)
   txtPAGAR.Refresh

   If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
      txtCNPJCPF.Mask = "###.###.###-##"
      Else: txtCNPJCPF.Mask = "##.###.###/####-##"
   End If

   txtCNPJCPF.BackColor = &HFFFFFF
   'txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtDesconto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Informe desconto unitário", "F10 - Gravar", "", ""

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Else
         FraSeq.Enabled = True
         txtProduto.SetFocus
   End If

   txtDesconto.SelStart = 0
   txtDesconto.SelLength = Len(txtQTDE)
   txtDesconto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_GotFocus"
End Sub

Private Sub txtDesconto_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtDesconto.Text) <> "" Then _
      txtDesconto.Text = Format(txtDesconto.Text, strFormatacao2Digitos)

   VALOR_UNITARIO_N = txtValor_Unitario.Text
   'If TRATA_DESCONTO(VALOR_UNITARIO_N) = False Then _
      Exit Sub

   If Trim(UCase(cmbDesconto.Text)) = "TRUE" Or Trim(UCase(cmbDesconto.Text)) = "VERDADEIRO" Or Trim(cmbDesconto.Text) = "1" Then
      If TRATA_DESCONTO_GRID(VALOR_UNITARIO_N, VALOR_VENDA_ORIGINAL_N) = False Then _
         Exit Sub
      Else
         txtValor_Unitario.Text = "" & VALOR_VENDA_ORIGINAL_N
         MsgBox "Forma de faturamento não permitido desconto !!!"
         Exit Sub
   End If

   If Trim(UF_CLIENTE_A) = "" Then _
      TRATA_PESSOA txtCNPJCPF.Text

   PROCESSA_ITEM

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_LostFocus"
End Sub

Private Sub txtNome_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNome.Text) <> "" Then _
      txtNome.Text = UCase(txtNome.Text)
   txtNome.Enabled = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_LostFocus"
End Sub
'==================cnpjcpf
Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "F7-Consulta Clientes", "Inform CNPJ/CPF Cliente e Tecle <<Enter>>", "", ""
   txtNome.Enabled = True
   txtCNPJCPF.Mask = "###############"
   txtCNPJCPF.BackColor = &HC0FFFF
   'txtCNPJCPF.SelStart = 0
   'txtCNPJCPF.SelLength = Len(txtCNPJCPF.Text)
   'txtCNPJCPF.
   txtCNPJCPF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtCNPJCPF.Text = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.Text = ""
            txtCNPJCPF.Mask = "##############"

            txtCNPJCPF.Text = CNPJCPF_A
            Call txtCNPJCPF_LostFocus
            FraSeq.Enabled = True
            txtProduto.SetFocus
            Exit Sub
         End If
         CNPJCPF_A = ""
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         txtNome.Enabled = True
         txtNome.SetFocus
         Else
            txtProduto.Enabled = True
            FraSeq.Enabled = True
            txtProduto.SetFocus
            txtNome.Enabled = False
      End If
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      UCase (txtProduto.Text)
      FraSeq.Enabled = True
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDescricao.Enabled = False

   MOSTRA_TOP "ESC-SAIR", "F7-Consulta Produtos", "Delete-Excluir Produto", "F10-Gravar", ""

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto.Text)
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtPedido.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then _
            EXCLUIR_ITEM Trim(txtProduto.Text), Trim(txtPedido.Text), Trim(txtSeq.Text)
         FraSeq.Enabled = True
         txtProduto.SetFocus
      Case vbKeyF7
         CONSULTA_PRODUTO
         FraSeq.Enabled = True
         txtProduto.SetFocus
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

   If KeyAscii = 13 Then
      KeyAscii = 0
      LE_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "Informe a quantidade", "F10-Gravar", "", ""
   
   If Trim(txtProduto.Text) = Empty Then
   '   MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
   '   txtProduto.Text = 99999999
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If
   If Trim(txtQTDE.Text) <> "" Then
      txtQTDE.SelStart = 0
      txtQTDE.SelLength = Len(txtQTDE.Text)
   End If
   txtQTDE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Len(Trim(txtQTDE.Text)) > 10 Then
         FraSeq.Enabled = True
         txtProduto.SetFocus
         Exit Sub
      End If

      If Len(Trim(txtQTDE.Text)) > 10 Then
         FraSeq.Enabled = True
         txtProduto.SetFocus
         Exit Sub
      End If

      If Trim(txtQTDE.Text) = "" Then
         txtQTDE.Text = 1
         Else
            If IsNumeric(txtQTDE.Text) Then
               QTDE_N = txtQTDE.Text
               If QTDE_N <= 0 Then _
                  txtQTDE.Text = 1
            End If
      End If

      Call txtDesconto_LostFocus
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
'On Error GoTo ERRO_TRATA

   If Len(Trim(txtQTDE.Text)) > 10 Then
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If

   If Trim(txtQTDE.Text) = "" Then
      txtQTDE.Text = 1
      Else
         If IsNumeric(txtQTDE.Text) Then
            QTDE_N = txtQTDE.Text
            If QTDE_N <= 0 Then _
               txtQTDE.Text = 1
         End If
   End If
   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
   txtQTDE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_LostFocus"
End Sub

Private Sub txtPedido_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "Tecle <ENTER> para gerar nova Pedido ou informe uma já existente", "", "", ""
   cmbVend.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_GotFocus"
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyPress"
End Sub

Private Sub txtPedido_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then
      txtPedido.Enabled = False

      If Trim(cmbVendAux.Text) = "" Then
         cmbVend.Text = "Balcão"
         cmbVendAux.Text = 0
      End If

      If txtCNPJCPF.Text = "" Then
         txtCNPJCPF.Text = "99999999999"
         If Trim(txtNome.Text) = "" Then _
            txtNome.Text = "Consumidor Final"
      End If
   
      QUALIFICA_VENDEDOR
   End If

   'VALIDA_PEDIDO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_LostFocus"
End Sub

Private Sub TXTVALOR_UNITARIO_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Informe Valor Unitário", "", "", ""
   
   txtValor_Unitario.SelStart = 0
   txtValor_Unitario.SelLength = Len(txtValor_Unitario.Text)

   If INDR_TRAVA_TABELA = True Then _
      txtQTDE.SetFocus

   txtValor_Unitario.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_GotFocus"
End Sub

Private Sub TXTVALOR_UNITARIO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      VALOR_UNITARIO_N = 0 & txtValor_Unitario.Text
      If STATUS_PROD = "P" Then
         If VALOR_UNITARIO_N < PRECO_PROD Then
            MsgBox "Produto Tipo Promoção Impossível dar desconto."
            txtValor_Unitario.Text = 0
            txtValor_Unitario.SetFocus
            Else: txtQTDE.SetFocus
         End If
         Else
            If VALOR_UNITARIO_N <> VLR_ANTERIOR_N Then
                If VALOR_UNITARIO_N < PRECO_PROD Then
                   VALOR_DESCONTO_N = Format(PRECO_PROD - VALOR_UNITARIO_N, strFormatacao2Digitos)
                   PERC_DESCONTO_N = ((VALOR_DESCONTO_N * 100) / PRECO_PROD)
                   PERC_DESCONTO_N = Format(PERC_DESCONTO_N, strFormatacao2Digitos)
                   Else
                      VALOR_DESCONTO_N = 0
                      PERC_DESCONTO_N = 0
                End If
                Else
                    VALOR_DESCONTO_N = 0
                    PERC_DESCONTO_N = 0
            End If

checa_desconto_valor:

            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "select * from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabUSU.EOF Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close

               MsgBox "Problemas com usuário, codigo=0"
               txtDesconto.SetFocus
               Exit Sub
               Else: txtQTDE.SetFocus
            End If
            If TabUSU.State = 1 Then _
               TabUSU.Close
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_KeyPress"
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_KeyPress"
End Sub

Private Sub TXTVALOR_UNITARIO_LostFocus()
'On Error GoTo ERRO_TRATA

   VALOR_UNITARIO_N = 0 & txtValor_Unitario.Text
   txtValor_Unitario.Text = Format(VALOR_UNITARIO_N, strFormatacao2Digitos)

   If VALOR_UNITARIO_N <= 0 Then
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Exit Sub
      Else
         VALOR_ITEM_N = txtValor_Unitario.Text
         txtValor_Unitario.Text = Format(VALOR_UNITARIO_N, strFormatacao2Digitos)
         If VALOR_ITEM_N <= 0 Then
            MsgBox "Valor Unitário Inválido !!!"
            FraSeq.Enabled = True
            txtProduto.SetFocus
            Exit Sub
         End If
   End If

   If Trim(txtValor_Unitario.Text) <> "" Then _
      If IsNumeric(txtValor_Unitario.Text) Then _
         txtValor_Unitario.Text = Format(txtValor_Unitario.Text, strFormatacao2Digitos)

   txtValor_Unitario.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_LostFocus"
End Sub

Private Sub txtVlrUnit_GotFocus()
   FraSeq.Enabled = True
   txtProduto.SetFocus
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    'OcultarControles

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_Click"
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   OcultarControles

   ExibirCelula

   txtProduto.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 0)
   txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 11)
   'txtPesoItem.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 2)
   FraSeq.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then
            If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then
               If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then
                  If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then
                     If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) <> "" Then                'codg Produto
                        If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then             'pedido_id
                           If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then          'seq_id
                              If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then       'produto_id
                                 txtProduto.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
                                 txtSeq.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
                                 EXCLUIR_ITEM Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Case vbKeyF12
         'frmobs.Show 1
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
         'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

      'If INDR_TRAVA_TABELA = True Then
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

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorDig.SelStart = 0
   txtValorDig.SelLength = Len(txtValorDig)
   txtValorDig.BackColor = &HC0FFFF

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
      If LastCol = 4 Then
         'If Left(UCase(Trim(cmbForma.Text)), 6) = "CARTAO" Or _
            Left(UCase(Trim(cmbForma.Text)), 6) = "CARTÃO" Or _
            Left(UCase(Trim(cmbForma.Text)), 3) = "POS" Then

            VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
            CODG_PRODUTO_A = "" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)

            If Trim(UCase(cmbDesconto.Text)) = "TRUE" Or Trim(UCase(cmbDesconto.Text)) = "VERDADEIRO" Or Trim(cmbDesconto.Text) = "1" Then
               If TRATA_DESCONTO_GRID(txtValorDig.Text, VALOR_ITEM_N) = False Then _
                  Exit Sub
               Else
                  txtValor_Unitario.Text = "" & VALOR_VENDA_ORIGINAL_N
                  MsgBox "Forma de faturamento não permitido desconto !!!"
                  Exit Sub
            End If
         'End If
      End If

      KeyAscii = 0
      If LastCol > 3 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      Dim QTDE_RETIDO_ESTORNO As Double

      QTDE_RETIDO_ESTORNO = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)

      AtribuiValorCelula
      'ProximaCelula
      OcultarControles

'==========ATUALIZAR GRID colunas
'3 = qtde
'4 = valor venda
'5 = desconto

      QTDE_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 3)
      VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
      VALOR_DESCONTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 5)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)
      'PRECO_CUSTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 8)
      CODG_PRODUTO_A = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 0))
      PRODUTO_ID_N = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 12))

      If QTDE_N > 0 And _
         VALOR_ITEM_N > 0 And _
         VALOR_DESCONTO_N >= 0 And _
         SEQ_ID_N > 0 Then

         MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N) - VALOR_DESCONTO_N), strFormatacao2Digitos)  'total item
         'lucro MSFlexGrid1.TextMatrix(LastRow, 9) = Format(((VALOR_ITEM_N - PRECO_CUSTO_N) * QTDE_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

         If INDR_ESTQ_NEGATIVO = False Then
            QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)
            If QTDE_ESTOQUE_N < 0 Then
               Beep
               MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
               txtQTDE.SetFocus
               Exit Sub
            End If
         End If

         If TabGridVaca.State = 1 Then _
            TabGridVaca.Close

         SQL = "select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, "
         SQL = SQL & " produto.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
         SQL = SQL & " PEDIDOITEM.PERC_DESC , PEDIDOITEM.Valor_Desconto, PEDIDOITEM.Status, "
         SQL = SQL & " PEDIDOITEM.PRECO_CUSTO"
         SQL = SQL & " from PEDIDO WITH (NOLOCK) "
         SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
         SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
         SQL = SQL & " INNER JOIN PRODUTO "
         SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
         SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

         SQL = SQL & " where PEDIDO.pedido_id = " & txtPedido.Text
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and pedidoitem.status <> 'C' "

         TabGridVaca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

         If Not TabGridVaca.EOF Then
            SQL = "update PEDIDOITEM set "
            SQL = SQL & " QTD_PEDIDA = " & tpMOEDA(QTDE_N)
            SQL = SQL & ",Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",Valor_Desconto = " & tpMOEDA(VALOR_DESCONTO_N)
            SQL = SQL & ",peso_item = " & tpMOEDA(QTDE_N)

            SQL = SQL & " where pedido_id = " & TabGridVaca.Fields("pedido_id").Value
            SQL = SQL & " and seq_id = " & SEQ_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            QTDE_RETIDO_ESTORNO = 0
         End If
         If TabGridVaca.State = 1 Then _
            TabGridVaca.Close

         MOSTRA_TOTAIS
      End If

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

      MSFlexGrid1.SetFocus
      LIMPA_BODY
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
'======================================================
'======================================================
'======================================================
'============================subrotinas
Sub EXCLUIR_ITEM(CODG_PRODUTO_A As String, PEDIDO_ID_N As Long, SEQ_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Trim(PEDIDO_ID_N) > 0 And Trim(SEQ_ID_N) > 0 And Trim(CODG_PRODUTO_A) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select PEDIDOITEM.*, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
      SQL = SQL & " PRODUTO.PRECO_VENDA, PRODUTO.PRECO_CUSTO, "
      SQL = SQL & " Produto.Situacao_Tributaria"
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
      SQL = SQL & " and PEDIDOITEM.pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_ID_N
      SQL = SQL & " and pedido.estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and pedidoitem.status <> 'C' "

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Deseja Excluir Esse Item?"
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            If TabProduto.State = 1 Then _
               TabProduto.Close

            VALOR_TOTAL_N = Format(VALOR_TOTAL_N - (TabTemp!Valor_Item * TabTemp!QTD_PEDIDA), "##,##0.00")

            'mudando controle estoque, vai baixar quando fechar a venda
            'BAIXA_RETIDO TabTemp!QTD_PEDIDA

            SQL = "Delete from PEDIDOITEM "
            SQL = SQL & " Where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and seq_id = " & TabTemp.Fields("seq_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            CONECTA_RETAGUARDA.Execute SQL

            If TabTemp.State = 1 Then _
               TabTemp.Close

            LIMPA_BODY
            txtTotalPedido.Text = Format(VALOR_TOTAL_N, "##,##0.00")
   
            GRAVA_CABECA "R", 1
            SETA_GRID
            Else
               If TabTemp.State = 1 Then _
                  TabTemp.Close
         End If
         Else: MsgBox "Produto não encontrado."
      End If
      Else: MsgBox "Informe código produto."
   End If
FraSeq.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM"
End Sub

Sub MOSTRA_DADOS_REQ()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Text = TabCabeca!CGCCPF

   'MOSTRA VENDEDOR
   If Not IsNull(TabCabeca!VENDEDOR_ID) Then
      cmbVendAux.Text = TabCabeca!VENDEDOR_ID

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
      SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where vendedor_id = " & cmbVendAux.Text
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then _
         cmbVend.Text = TabVENDEDOR!DESCRICAO
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If

   If Not IsNull(TabCabeca.Fields("tabelapreco_id").Value) Then
      cmbTabPrecoAux.Text = TabCabeca.Fields("tabelapreco_id").Value

      If TabTemp.State = 1 Then _
         TabTemp.Close
      SQL = "select descricao from TABELAPRECO WITH (NOLOCK)"
      SQL = SQL & " where tabelapreco_id = " & TabCabeca.Fields("tabelapreco_id").Value
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         cmbTabPreco.Text = TabTemp!DESCRICAO
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   'MOSTRA CLIENTE
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome,status from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where cgccpf = '" & TabCabeca!CGCCPF & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      If TabCabeca!CGCCPF = "99999999999" Then
         If Not IsNull(TabCabeca!NOME_CLIENTE) Then
            If Trim(txtNome.Text) = "" Then _
               txtNome.Text = TabCabeca!NOME_CLIENTE
            Else
               If Trim(txtNome.Text) = "" Then _
                  txtNome.Text = TabCliente!NOME
         End If
         Else
            If Trim(txtNome.Text) = "" Then _
               txtNome.Text = TabCliente!NOME
      End If
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome_cliente from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         txtNome.Text = Trim(TabCliente.Fields(0).Value)

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_REQ"
End Sub

Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   Valr_Venda_Produto_n = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""
   txtQtdeDisp.Text = "" & Format(0, strFormatacao3Digitos)

   QTDE_PEDIDO = 0
   QTDE_ESTOQUE_N = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_DIFERENCA_N = 0
   PRODUTO_ID_N = 0

   txtAtacado.Text = Format(0, strFormatacao2Digitos)
   txtVarejo.Text = Format(0, strFormatacao2Digitos)
   txtValor_Unitario.Text = Format(0, strFormatacao2Digitos)
   txtPreçoCusto.Text = Format(0, strFormatacao2Digitos)
   txtQTDE.Text = Format(0, strFormatacao3Digitos)
   txtDesconto.Text = Format(0, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   MOSTRA_VENDEDORES

   MSFlexGrid1.Clear

   txtValorDig.Visible = False
   FraSeq.Enabled = False
   MSFlexGrid1.Enabled = True
   FraReq.Enabled = True
   txtProduto.Enabled = True

   Toolbar1.Buttons(3).Visible = True
   Toolbar1.Buttons(8).Visible = True
   Toolbar1.Buttons(9).Visible = False
   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
      Toolbar1.Buttons(9).Visible = True

   txtNome.Enabled = True
   txtPesoTotal.Text = ""
   txtItens.Text = ""
   txtTotalPedido.Text = ""
   txtDescontoRodape.Text = ""
   txtVlrUnit.Text = ""
   txtQtdeDisp.Text = "" & Format(0, strFormatacao3Digitos)

   PRODUTO_ID_N = 0
   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   txtPedido.Text = ""
   txtDtEmis = Format(Date, "dd/mm/yyyy")
   txtNome.Text = ""
   txtCNPJCPF.Text = ""
   cmbForma.Visible = False
   LIMPA_BODY
   
   VALOR_TOTAL_N = 0
   PEDIDO_ID_N = 0
   QTDE_PEDIDO = 0
   QTDE_ESTOQUE_N = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_TOTAL_N = 0
   USU_LIBERA_VENDA_N = 0
   txtLIMITE.Text = ""
   txtPAGAR.Text = ""
   INDR_RECEITA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub MOSTRA_VENDEDORES()
'On Error GoTo ERRO_TRATA

   cmbVend.Clear
   cmbVendAux.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close
   SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVENDEDOR.EOF
      cmbVend.AddItem Trim(TabVENDEDOR!DESCRICAO) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
      TabVENDEDOR.MoveNext
   Wend
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   cmbVend.Enabled = False

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select logon from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      CRITERIO_A = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
      SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where descricao like " & CRITERIO_A
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         cmbVend.Text = Trim(TabVENDEDOR!DESCRICAO)
         cmbVendAux.Text = Trim(TabVENDEDOR!VENDEDOR_ID)
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If

   INDR_TRAVA_TABELA = False
   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
      cmbVend.Enabled = True
      Else: INDR_TRAVA_TABELA = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDORES"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear

   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns
   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabGridVaca.State = 1 Then _
      TabGridVaca.Close

   SQL = "select produto.CODG_PRODuto as 'Código', "                                                                                '0
   SQL = SQL & " PRODUTO.REFERENCIA as Ref, "                                                                                       '1
   SQL = SQL & " PRODUTO.DESCRICAO as Produto, "                                                                                    '2
   SQL = SQL & " PEDIDOITEM.QTD_PEDIDA as Qtde,"                                                                                    '3
   SQL = SQL & " PEDIDOITEM.VALOR_ITEM as ValorItem, "                                                                              '4
   SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO as Desconto, "                                                                           '5
   SQL = SQL & " ((PEDIDOITEM.VALOR_ITEM - PEDIDOITEM.VALOR_DESCONTO) * PEDIDOITEM.QTD_PEDIDA) as TotItem, "                        '6
   SQL = SQL & " PRODUTO.SITUACAO_TRIBUTARIA as ST, "                                                                               '12
   SQL = SQL & " PRODUTO.ALIQUOTA_ICMS as ICMS, "                                                                                   '13
   SQL = SQL & " PRODUTO.CODG_NCM as NCM, "                                                                                         '14
   SQL = SQL & " PEDIDOITEM.PEDIDO_ID, "                                                                                            '15
   SQL = SQL & " PEDIDOITEM.SEQ_ID, "                                                                                               '16
   SQL = SQL & " PEDIDOITEM.PRODUTO_ID, "                                                                                           '17
   SQL = SQL & " PEDIDOITEM.STATUS AS StatusItem "                                                                                  '18

   SQL = SQL & " from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " order by seq_id desc"

   TabGridVaca.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGridVaca.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabGridVaca.Fields.Count

      ReDim largura_coluna(0 To TabGridVaca.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabGridVaca.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabGridVaca.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabGridVaca.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1
      Do While Not TabGridVaca.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabGridVaca.Fields.Count - 1
            'If Coluna = 3 Or Coluna = 7 Then
            If Coluna = 3 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGridVaca.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  'If Coluna = 4 Or Coluna = 5 Or Coluna = 6 Or Coluna = 7 Or Coluna = 8 Or Coluna = 9 Or Coluna = 10 Then
                  If Coluna = 4 Or Coluna = 5 Or Coluna = 6 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGridVaca.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabGridVaca.Fields(Coluna).Value)
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabGridVaca.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGridVaca.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabGridVaca.MoveNext
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

'CellFontName        - Define o nome da fonte para uma célula
'CellFontSize        - Define o tamanho da fonte para a célula
'CellFontBold        - Define se a fonte aparece em negrito.
'CellFontItalic      - Define se a fonte aparece em itálico.
'CellFontUnderline   - Define se a fonte aparece sublinhada.

'&H80C0FF = LARANJA
'&H8000000F = CINZA
'&HFF& = VERMELHO
'vbBlack 0x0
'vbRed 0xFF
'vbGreen 0xFF00
'vbYellow 0xFFFF
'vbBlue 0xFF0000
'vbMagenta 0xFF00FF
'vbCyan 0xFFFF00
'vbWhite 0xFFFFFF

'Codigo Produto
      MSFlexGrid1.ColWidth(0) = 2000
      MSFlexGrid1.ColAlignment(0) = 0

'Referencia
      MSFlexGrid1.ColWidth(1) = 0
      MSFlexGrid1.ColAlignment(1) = 0

'Descrição Produto
      MSFlexGrid1.ColWidth(2) = 7000
      MSFlexGrid1.ColAlignment(2) = 0

'QTDE
      MSFlexGrid1.ColWidth(3) = 2000
      MSFlexGrid1.ColAlignment(3) = 7

'Valor Item
      MSFlexGrid1.ColWidth(4) = 2000
      MSFlexGrid1.ColAlignment(4) = 7

'Desconto
      MSFlexGrid1.ColWidth(5) = 1500
      MSFlexGrid1.ColAlignment(5) = 7

'Total Item
      MSFlexGrid1.ColWidth(6) = 2000
      MSFlexGrid1.ColAlignment(6) = 7

'SITUAÇÃO TRIBUTARIA PRODUTO
      MSFlexGrid1.ColWidth(7) = 500
      MSFlexGrid1.ColAlignment(7) = 0

'ALIQUOTA ICMS
      MSFlexGrid1.ColWidth(8) = 500
      MSFlexGrid1.ColAlignment(8) = 0

'NCM
      MSFlexGrid1.ColWidth(9) = 500
      MSFlexGrid1.ColAlignment(9) = 0

'Pedido_id
      MSFlexGrid1.ColWidth(10) = 500
      MSFlexGrid1.ColAlignment(10) = 0

'seq_id
      MSFlexGrid1.ColWidth(11) = 500
      MSFlexGrid1.ColAlignment(11) = 0

'produto_id
      MSFlexGrid1.ColWidth(12) = 500
      MSFlexGrid1.ColAlignment(12) = 0

'SITUAÇÃO ITEM
      MSFlexGrid1.ColWidth(13) = 500
      MSFlexGrid1.ColAlignment(13) = 0
   End If

   ' fecha o recordset e a conexao
   If TabGridVaca.State = 1 Then _
      TabGridVaca.Close

   MOSTRA_TOTAIS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub GRAVA_CABECA(TIPO_REGISTRO_A As String, STATUS_N As Integer)
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""
   CLIENTE_ID_N = 0
   If Trim(cmbFormaAUX.Text) = "" Then _
      cmbFormaAUX.Text = "9999"

   txtCNPJCPF.Mask = "###############"

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value
      Else
         If TabCliente.State = 1 Then _
            TabCliente.Close
   
         MsgBox "Cliente não cadastrado, verificar."
         txtPedido.Text = ""
         Exit Sub
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

'PEDIDO_ID_N = 0 & MAX_ID("pedido_id", "PEDIDO", "", "", "", "")
PEDIDO_ID_N = 0 & txtPedido.Text

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCabeca.EOF Then
      SQL = "INSERT INTO PEDIDO "
         SQL = SQL & "("
            SQL = SQL & "PEDIDO_ID,Empresa_id, CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, Status, "
            SQL = SQL & " Tipo_Registro,usuario_id, CLIENTE_ID, Valor_ToTal,"
            SQL = SQL & " valor_desconto,perc_desc,NUMERO_CAIXA_CPU,estabelecimento_id,USUARIO_LIBERA_VENDA"
         SQL = SQL & ") "
         SQL = SQL & " VALUES ("
            SQL = SQL & PEDIDO_ID_N
            SQL = SQL & "," & EMPRESA_ID_N
            SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
            SQL = SQL & "," & cmbVendAux.Text & ","
            SQL = SQL & "'" & Now & "'"
            SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
            SQL = SQL & "," & STATUS_N
            SQL = SQL & ",'" & TIPO_REGISTRO_A & "'"
            SQL = SQL & "," & USUARIO_ID_N
            SQL = SQL & "," & CLIENTE_ID_N
            SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)
            SQL = SQL & "," & tpMOEDA(0)  'vai zerar e tratar somente na tela de desconto
            SQL = SQL & "," & tpMOEDA(0)
            SQL = SQL & "," & NUMERO_CAIXA_CPU           'NUMERO_CAIXA_CPU
            SQL = SQL & "," & ESTABELECIMENTO_ID_N       'estabelecimento_id
            SQL = SQL & "," & USU_LIBERA_VENDA_N
            'SQL = SQL & "," & 9999
         SQL = SQL & ")"
      Else
         PEDIDO_ID_N = 0 & TabCabeca.Fields("pedido_id").Value
         txtPedido.Text = PEDIDO_ID_N

         If Not IsNull(TabCabeca!STATUS) Then
            If TabCabeca!STATUS <> 3 Then
               If TabCabeca!STATUS <> 4 Then
                  If TabCabeca!STATUS <> 5 Then
                     If TabCabeca!STATUS <> 9 Then
                        SQL = "UPDATE PEDIDO SET "
                        SQL = SQL & " Valor_total = " & tpMOEDA(VALOR_TOTAL_N)
                        SQL = SQL & ",Valor_desconto = " & tpMOEDA(0)   'vai zerar e tratar somente na tela de desconto
                        SQL = SQL & ",Perc_desc = " & tpMOEDA(0)
                        SQL = SQL & ",CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
                        SQL = SQL & ",Vendedor_id = " & cmbVendAux.Text
                        SQL = SQL & ",dt_req = '" & Now & "'"
                        SQL = SQL & ",nome_cliente = '" & txtNome.Text & "'"
                        SQL = SQL & ",Status = " & STATUS_N
                        SQL = SQL & ",TIPO_REGISTRO = '" & TIPO_REGISTRO_A & "'"
                        SQL = SQL & ",usuario_id = " & USUARIO_ID_N
                        SQL = SQL & ",USUARIO_LIBERA_VENDA = " & USUARIO_ID_N
                        SQL = SQL & ",CLIENTE_ID = " & CLIENTE_ID_N

                        SQL = SQL & " where pedido_id = " & txtPedido.Text
                        SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     End If
                  End If
               End If
            End If
         End If
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   CONECTA_RETAGUARDA.Execute SQL

   SQL = "select pedido_id from PEDIDOFATURA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      Acao_N = 2
      Else: Acao_N = 1
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   If TABELAPRECO_ID_N <= 0 Then _
      TABELAPRECO_ID_N = 0 & cmbTabPrecoAux.Text

   If FORMAPAGTO_ID_N <= 0 Then _
      FORMAPAGTO_ID_N = 0 & cmbFormaAUX.Text

   TIPOVENDA_ID_N = 1
   SQL = "select tipovenda_id from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where formapagto_id = " & FORMAPAGTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then _
      TIPOVENDA_ID_N = 0 & TabCabeca.Fields(0).Value
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

spPEDIDOFATURA Acao_N, 0, PEDIDO_ID_N, TABELAPRECO_ID_N, FORMAPAGTO_ID_N, TIPOVENDA_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA"
End Sub

Sub GERA_VENDA()
'On Error GoTo ERRO_TRATA

   Dim strimpressoa As String

   PERC_DESCONTO_USUARIO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   USU_LIBERA_VENDA_N = 0

'   If INDR_LIBERA_DESCONTO = True Then
'      Msg = "Deseja informar desconto ?"
'      PERGUNTA Msg, vbYesNo + 32, "Desconto NFE", "DEMO.HLP", 1000
'      If RESPOSTA = vbYes Then
'         frmVENDADESCONTO.Show 1
'         If INDR_DESCONTO_AUTORIZADO = False Then _
            Exit Sub
'      End If
'   End If

   PERC_DESCONTO_USUARIO_N = 0
   PEDIDO_ID_N = txtPedido.Text
   CNPJCPF_A = txtCNPJCPF.Text

   If Trim(cmbTabPrecoAux.Text) = "" Then _
      cmbTabPrecoAux.Text = 0

   'atualizando desconto na cabeça
   SQL = "UPDATE PEDIDO SET "
   SQL = SQL & " Valor_desconto = " & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
   SQL = SQL & " , Perc_desc = " & tpMOEDA(PERC_DESCONTO_N)
   SQL = SQL & " , cgccpf = '" & CNPJCPF_A & "'"
   SQL = SQL & " , nome_cliente = '" & Trim(txtNome.Text) & "'"
   SQL = SQL & " , status = 2"
   If USU_LIBERA_VENDA_N > 0 Then
      SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
   End If
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   If RECEBE_PEDIDO_VENDA = True Then _
      FAZ_RECEBIMENTO

   Exit Sub

   Msg = "Deseja Imprimir Pedido?"
   Style = vbYesNo + 32
   Title = "Atenção."
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then
      Dim CEP_A As String
      FORMULA_REL = "{vwRelVenda.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
      FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & txtPedido.Text
      FORMULA_REL = FORMULA_REL & " and {vwRelVenda.statusitem} <> 'C' "

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_pedido_venda.rpt"
      If CNPJ_EMPRESA_N = "15333554000188" Then _
         Nome_Relatorio = "pedido_shf.rpt"

      frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_VENDA"
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Sub VALIDA_PEDIDO()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 1

   If Trim(txtPedido.Text) = "" Then
      GERA_PEDIDO_ID

      txtPedido.Text = PEDIDO_ID_N
      Else
         txtPedido.Enabled = True
            PEDIDO_ID_N = txtPedido.Text
         txtPedido.Enabled = False
   End If

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      bolRequisicaoJaExiste = False

      PEDIDO_ID_N = txtPedido.Text

      bolRequisicaoJaExiste = True

      MOSTRA_DADOS_REQ

      CRITERIO_A = ""

      txtDtEmis.Text = TabCabeca!DT_REQ

      If TabCabeca!STATUS = 9 Then
         MsgBox "Pedido cancelada, impossível alterar !!!"
         FraReq.Enabled = False
         FraSeq.Enabled = False
         txtProduto.Enabled = False
         Unload Me
         Exit Sub
         Else '1=ORÇAMENTO;2=GERADO;3=EMITIDA COM NOTA;4=EMITIDA COM CUPOM;5=ARECEBER;7=ECF/NF;9=CANCELADO
            If (TabCabeca!STATUS = 3 Or TabCabeca!STATUS = 5) Then
               If TabCabeca!STATUS = 3 Then
                  Toolbar1.Buttons(3).Visible = False
                  Toolbar1.Buttons(8).Visible = False
                  If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
                     Toolbar1.Buttons(9).Visible = True

                  PERGUNTA "Nota Processada para este pedido.", vbNo, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     If TabCabeca.State = 1 Then _
                        TabCabeca.Close

                     Else
                        FraSeq.Enabled = False
                        'LIMPA_BODY
                        'LIMPA_TUDO
                   End If
                   Exit Sub
               End If
               If TabCabeca!STATUS = 5 Then
                  Toolbar1.Buttons(3).Visible = False
                  Toolbar1.Buttons(8).Visible = False
                  FraSeq.Enabled = False
                  MSFlexGrid1.Enabled = False
                  FraReq.Enabled = False

                  'gerente / diretor
                  If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
                     Toolbar1.Buttons(9).Visible = True

                  PERGUNTA "Venda ja Faturada, Deseja imprimir ?", vbYesNo + 32, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     GERA_IMPRESSAO
                     Else
                        FraSeq.Enabled = False
                        MSFlexGrid1.Enabled = False
                  End If
               End If
               Exit Sub
            End If
            If TabCabeca!STATUS = 4 Then
               MsgBox "Permitido somente consulta, cupom fiscal emitido."
               Exit Sub
            End If
      End If
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VALIDA_PEDIDO"
End Sub

Sub GRAVA_TUDO_ITEM(strCFOP As String)
'On Error GoTo ERRO_TRATA

   'Tratamento da tributacao
   'fazer no final desta rotina
   'CODG_PRODUTO_A = Trim(txtProduto.Text)

   If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
      If Trim(UF_CLIENTE_A) = "" Then
         MsgBox "Cliente com cadastro incompleto !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   End If

   Dim tabEnd  As New ADODB.Recordset

   If Trim(txtPreçoCusto.Text) = "" And Trim(cmbTabPrecoAux.Text) <> "" And Trim(cmbFormaAUX.Text) <> "" Then _
      txtPreçoCusto.Text = "" & TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PRODUTO_ID_N, cmbTabPrecoAux.Text, cmbFormaAUX.Text)

   If Not IsNumeric(txtPreçoCusto.Text) Then _
      txtPreçoCusto.Text = 0

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
'=====================
   If Trim(strCFOP) = "" Then
      If tabEnd.State = 1 Then _
         tabEnd.Close

      SQL = "select CEP.UF from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK)"
      SQL = SQL & " ON CLIENTE.PESSOA_ID = ENDERECO.PESSOA_ID "
      SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
      SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

      SQL = SQL & " where CLIENTE.pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and tipo = 'C'"

      tabEnd.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabEnd.EOF Then _
         UF_CLIENTE_A = Trim(tabEnd.Fields("UF").Value)
      If tabEndereco.State = 1 Then _
         tabEndereco.Close
      
      If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
         strCFOP = "5102"
         Else: strCFOP = "6102"
      End If
   End If

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and seq_id = " & SEQ_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabPedidoItem.EOF Then
      SQL = "INSERT INTO PEDIDOITEM "
      SQL = SQL & " ("
         SQL = SQL & "PEDIDO_ID,SEQ_ID,PRODUTO_ID,Qtd_Pedida,Valor_item,cfop_id,"
         SQL = SQL & " PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG,PESO_ITEM "
      SQL = SQL & ") "
      SQL = SQL & " VALUES ("

         SQL = SQL & PEDIDO_ID_N                                                          'PEDIDO_id
         SQL = SQL & "," & SEQ_ID_N                                                       'SEQ_ID
         SQL = SQL & "," & PRODUTO_ID_N
         SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                                           'Qtd_Pedida
         SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                                          'Valor_item
         SQL = SQL & ",'" & Trim(strCFOP) & "'"
         SQL = SQL & "," & tpMOEDA(PERC_DESCONTO_N)                                       'PERC_DESC
         SQL = SQL & "," & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)  'valor_desconto
         SQL = SQL & ",'P'"                                                               'status
         SQL = SQL & "," & tpMOEDA(txtPreçoCusto.Text)                                    'PRECO_CUSTO
         SQL = SQL & ",'PC'"                                                              'TIPO_REG
         SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                                           'PESO_ITEM

      SQL = SQL & ")"
      Else
         'PEDIDO_ID_N = TabPedidoItem.Fields("pedido_id").Value

         SQL = "UPDATE PEDIDOITEM SET "
         SQL = SQL & " qtd_pedida = " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & ", Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
         SQL = SQL & ", PERC_desc = " & tpMOEDA(PERC_DESCONTO_N)
         SQL = SQL & ", valor_desconto = " & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)
         SQL = SQL & ", status = 'P'"
         SQL = SQL & ", preco_custo = " & tpMOEDA(txtPreçoCusto.Text)
         SQL = SQL & ", PESO_ITEM = " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & ", CFOP_ID = '" & Trim(strCFOP) & "'"

         SQL = SQL & " Where pedido_id = " & txtPedido.Text
         SQL = SQL & " and pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
   End If
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   CONECTA_RETAGUARDA.Execute SQL

   'Tratamento da tributacao
   CODG_PRODUTO_A = Trim(txtProduto.Text)

   PREPARA_TRIBUTACAO_PRODUTO Trim(txtCNPJCPF.Text), Trim(VALOR_ITEM_N), Trim(QTDE_PEDIDO)

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO_ITEM"
End Sub

Sub QUALIFICA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   cmbVend.Enabled = False

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select logon from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      CRITERIO_A = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select descricao, vendedor_id,tabelapreco_id from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where descricao like " & CRITERIO_A
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         cmbVend.Text = TabVENDEDOR!DESCRICAO
         cmbVendAux.Text = TabVENDEDOR!VENDEDOR_ID

         If Not IsNull(TabVENDEDOR.Fields("tabelapreco_id").Value) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from TABELAPRECO WITH (NOLOCK)"
            SQL = SQL & " where tabelapreco_id = " & TabVENDEDOR.Fields("tabelapreco_id").Value
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbTabPreco.Text = "" & Trim(TabTemp!DESCRICAO)
               cmbTabPrecoAux.Text = "" & Trim(TabTemp!TABELAPRECO_ID)
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then _
      cmbVend.Enabled = True

   If Trim(cmbVendAux.Text) = "" Then
      cmbVend.Text = "Balcão"
      cmbVendAux.Text = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "QUALIFICA_VENDEDOR"
End Sub

Sub GERA_IMPRESSAO()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" Then
      PEDIDO_ID_N = txtPedido.Text
      Else: PEDIDO_ID_N = InputBox(SQL3, "Informe número de Pedido a ser impressa ")
   End If

   FORMULA_REL = "{vwRelVenda.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & PEDIDO_ID_N
   FORMULA_REL = FORMULA_REL & " and {vwRelVenda.statusitem} <> 'C' "

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_pedido_venda.rpt"
   If CNPJ_EMPRESA_N = "15333554000188" Then _
      Nome_Relatorio = "pedido_shf.rpt"

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_IMPRESSAO"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      FraSeq.Enabled = True
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Sub INICIALIZA_VENDA()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name

   UF_CLIENTE_A = ""  'Variavel para tratamento Fiscal do item
   UF_EMPRESA_A = "" 'Variavel para tratamento Fiscal do item
   strInscEstadual = "" 'Variavel para tratamento Fiscal do item
   dblTipoCliente = -1 'Variavel para tratamento fiscal do item
   strCPFCNPJ = ""
   'bolRequisicaoJaExiste = False 'Indica se a requisicao atual é nova, ou se ja
                                 'esta no banco de dados ou nao.

   txtDtEmis = Format(Date, "dd/mm/yyyy")

   PEGA_DADOS_EMPRESA
   QUALIFICA_VENDEDOR

   If TIPO_USUARIO < 4 Then _
      Toolbar1.Buttons(8).Visible = False

   CONT_N = 0

   'If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   'BUSCA_TABELA_PRECO_VENDEDOR Trim(cmbVendAux.Text)
   'If Not TabVENDEDOR.EOF Then
      Call cmbTabPreco_Click

   '   While Not TabVENDEDOR.EOF
   '      If CONT_N <> TabVENDEDOR.Fields("TABELAPRECO_ID").Value Then
   '         cmbTabPreco.AddItem Trim(TabVENDEDOR!DESCRICAO)
   '         cmbTabPrecoAux.AddItem Trim(TabVENDEDOR!TABELAPRECO_ID)
   '         CONT_N = TabVENDEDOR.Fields("TABELAPRECO_ID").Value
   '      End If
   '      TabVENDEDOR.MoveNext
   '   Wend
   'End If
   'If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_VENDA"
End Sub

Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   Dim TOT_ITENS_PEDIDO_N As Long

   TOT_ITENS_PEDIDO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0

   txtVlrUnit.Text = Format(VALOR_ITEM_N, "##,##0.00")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'BUSCA VALOR TOTAL VENDA
   SQL = "select sum(valor_item*qtd_pedida) from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.empresa_id  = " & EMPRESA_ID_N
   SQL = SQL & " and PEDIDOITEM.pedido_id = " & txtPedido.Text

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_ITEM_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'CONTA QTDE ITENS NO PEDIDO
   SQL = "select count(pedidoitem.produto_id) from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      TOT_ITENS_PEDIDO_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'VALOR DESCONTO ITEM
   SQL = "select sum(PEDIDOITEM.valor_desconto) from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.empresa_id  = " & EMPRESA_ID_N
   SQL = SQL & " and PEDIDOITEM.pedido_id = " & txtPedido.Text

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'VALOR DESCONTO NA CABEÇA
   SQL = "select valor_desconto from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         If TabTemp.Fields(0).Value > 0 Then _
            VALOR_DESCONTO_N = TabTemp.Fields(0).Value + VALOR_DESCONTO_N

   If TabTemp.State = 1 Then _
      TabTemp.Close

   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N

   txtItens.Text = "" & TOT_ITENS_PEDIDO_N
   txtTotalPedido.Text = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtDescontoRodape.Text = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
   txtPesoTotal.Text = ""

   SQL = "select sum(peso_item) from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         If TabTemp.Fields(0).Value > 0 Then _
            txtPesoTotal.Text = "" & Format(TabTemp.Fields(0).Value, strFormatacao3Digitos)
            'txtPesoTotal.Text = "" & Format(TabTemp.Fields(0).Value / 1000, strFormatacao3Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtPesoTotal.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAIS"
End Sub

Sub LE_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   CODG_PRODUTO_A = Trim(txtProduto.Text)
   INDR_PROD_BALANCA = False

   'LE POR CODIGO DE PRODUTO
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close

      Exit Sub
   End If

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS_A = "" & Trim(CODG_PRODUTO_A)
   QTDE_N = 0
   CRITERIO_A = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close
'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   SQL = "select count(produto_id) from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO_A = Trim(CODIGO_BARRAS_A)

            frmPEDIDOBARRAS.Show 1

            If Trim(CRITERIO_A) <> "" Then
               txtProduto.Text = Trim(CRITERIO_A)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select * from PRODUTO WITH (NOLOCK)"
               SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and situacao <> 'C' "
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then _
                  MOSTRA_DADOS_PRODUTO
               If TabProduto.State = 1 Then _
                  TabProduto.Close

               CRITERIO_A = ""
               Exit Sub
            End If
         End If
      End If
   End If

CRITERIO_A = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close

      txtQTDE.Text = 1
      Call txtDesconto_LostFocus
      FraSeq.Enabled = True
      txtProduto.SetFocus

      Exit Sub
   End If

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS_A = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS_A) = 13 Then
      '2 = produtos "in store" (sempre será 2)     1
      'C = código do produto (4,5 ou 6 dígitos)    2 a 8
      'T = total a pagar (sempre 6 dígitos)        9 a 13
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, CasaInicioCodgProdBarra_N, TamanhoCodgProdBarra_N))

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         If UCase(PESO_VALOR_A) = UCase("valor") Then 'POR VALOR
            VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS_A, 8, TamanhoPesoValorBarra_N) / 100
            QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value) 'sta
            PESO_ITEM_N = QTDE_N
            txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
            txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)

            MOSTRA_DADOS_PRODUTO
            Else
               QTDE_N = 0 & Int(Mid(CODIGO_BARRAS_A, 8, 5))   'gramas
               QTDE_N = QTDE_N / 1000
               PESO_ITEM_N = QTDE_N
               txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
               txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)

               MOSTRA_DADOS_PRODUTO
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

         Exit Sub
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MsgBox "Produto não cadastrado."
   FraSeq.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_PRODUTO"
End Sub

Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   INDR_PROD_BALANCA = False
   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

   If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then _
      INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value

   If INDR_PROD_BALANCA = True Then
      Label13.Caption = "Preço/Kg"
      Else: Label13.Caption = "Preço/UN"
   End If

   txtProduto.Text = Trim(TabProduto.Fields("codg_produto").Value)
   STATUS_PROD = TabProduto!SITUACAO
   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      txtProduto.Text = TabProduto!Codg_Produto
      txtDescricao.Text = TabProduto!DESCRICAO
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            FraSeq.Enabled = True
            txtProduto.SetFocus
            Exit Sub
            Else: txtDescricao.Text = Trim(TabProduto!DESCRICAO)
         End If
   End If

   txtPesoItem.Text = Format(TabProduto.Fields("peso_liquido").Value, strFormatacao3Digitos)
   txtAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
   txtVarejo.Text = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
   STATUS_PROD = TabProduto!SITUACAO

   If Not IsNull(TabProduto!PRECO_Venda) Then
      txtVlrUnit.Text = "" & Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)

      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_Venda
      txtValor_Unitario.Text = Format(Valr_Venda_Produto_n, strFormatacao2Digitos)
      txtPreçoCusto.Text = "" & Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)

      If Trim(cmbTabPrecoAux.Text) <> "" And Trim(cmbFormaAUX.Text) <> "" Then _
         txtPreçoCusto.Text = "" & TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PRODUTO_ID_N, cmbTabPrecoAux.Text, cmbFormaAUX.Text)

      VLR_ANTERIOR_N = TabProduto!PRECO_Venda
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   PRECO_PROD = 0 & txtAtacado.Text

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   QTDE_ESTOQUE_N = Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value), strFormatacao3Digitos)
   txtQtdeDisp.Text = "" & Format(QTDE_ESTOQUE_N, strFormatacao3Digitos)
   
   CODG_PRODUTO_A = Trim(txtProduto.Text)

   If INDR_ESTQ_NEGATIVO = False Then
      QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      If QTDE_ESTOQUE_N <= 0 Then
         MsgBox "Produto sem estoque disponível."
         FraSeq.Enabled = True
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Not IsNull(TabProduto.Fields("codg_ncm").Value) Then
      If Len(TabProduto.Fields("codg_ncm").Value) > 2 Then
         If Len(TabProduto.Fields("codg_ncm").Value) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"
            LIMPA_BODY
            FraSeq.Enabled = True
            txtProduto.SetFocus
         End If
      End If
   End If

   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

   'If Trim(txtPedido.Text) = "" Then
   '   MsgBox "Falta numero pedido."
   '   Exit Sub
   'End If

If Trim(txtPedido.Text) <> "" Then

   PEDIDO_ID_N = Trim(txtPedido.Text)

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where PRODUTO_ID = " & PRODUTO_ID_N
   SQL = SQL & " and pedido_ID = " & PEDIDO_ID_N
   SQL = SQL & " and seq_ID = " & Trim(txtSeq.Text)
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      txtValor_Unitario.Text = Format(TabPedidoItem!Valor_Item, strFormatacao2Digitos)
      txtDesconto.Text = Format(TabPedidoItem!PERC_DESC, strFormatacao2Digitos)
      txtQTDE.Text = Format(TabPedidoItem!QTD_PEDIDA, strFormatacao3Digitos)
      QTDE_PEDIDO = TabPedidoItem!QTD_PEDIDA
      VALOR_ITEM_N = TabPedidoItem!Valor_Item
      VALOR_DIFERENCA_N = TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA
      txtSeq.Text = "" & TabPedidoItem.Fields("seq_id").Value

      QTDE_ESTOQUE_N = Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabPedidoItem.Fields("produto_id").Value), strFormatacao3Digitos)
      txtQtdeDisp.Text = "" & Format(QTDE_ESTOQUE_N, strFormatacao3Digitos)
   End If
End If

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

'vai pegar preço de tabela
   If Trim(cmbTabPrecoAux.Text) <> "" Then
      If Trim(cmbVendAux.Text) <> "" Then
         If Trim(cmbFormaAUX.Text) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select TABELAPRECOITEM.VALOR_VENDA from TABELAPRECO "
            SQL = SQL & " INNER JOIN TABELAPRECOITEM "
            SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
            SQL = SQL & " INNER JOIN VENDEDOR "
            SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = VENDEDOR.TABELAPRECO_ID "
            SQL = SQL & " INNER JOIN PRODUTO "
            SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " INNER JOIN FORMAPAGTO "
            SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
            SQL = SQL & " where valor_venda > 0 "

            SQL = SQL & " and codg_tabela = '" & cmbTabPrecoAux.Text & "'"
            SQL = SQL & " and vendedor.vendedor_id = " & cmbVendAux.Text
            SQL = SQL & " and TABELAPRECOITEM.formapagto_id = " & cmbFormaAUX.Text
            SQL = SQL & " and TABELAPRECOITEM.produto_id = " & PRODUTO_ID_N

            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               txtValor_Unitario.Text = Format(TabTemp.Fields("VALOR_VENDA").Value, strFormatacao2Digitos)
               txtAtacado.Text = Format(TabTemp.Fields("VALOR_VENDA").Value, strFormatacao2Digitos)
               Else  'NÃO ACHOU NA TABELA DE PREÇO VAI PEGAR DO CADASTRO DE PRODUTO
                  VALOR_VENDA_ORIGINAL_N = 0

                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select preco_venda,produto_id,preco_atacado from PRODUTO WITH (NOLOCK)"
                  SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
                  SQL = SQL & " and situacao <> 'C' "
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabProduto.EOF Then
                     VALR_ATACADO_N = 0 & TabProduto.Fields("preco_atacado").Value
                     txtAtacado.Text = "" & Format(VALR_ATACADO_N, strFormatacao2Digitos)
                     VALR_VENDA_N = 0 & TabProduto.Fields("preco_venda").Value
                     txtVarejo.Text = "" & Format(VALR_VENDA_N, strFormatacao2Digitos)
                     VALOR_VENDA_ORIGINAL_N = 0 & TabProduto.Fields("preco_venda").Value
                     PRODUTO_ID_N = 0 & TabProduto.Fields("produto_id").Value
                  End If
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  txtValor_Unitario.Text = "" & Format(VALOR_VENDA_ORIGINAL_N, strFormatacao2Digitos)

                  MsgBox "Produto sem cadastro na tabela de preço : " & Trim(cmbTabPreco.Text) & " ; pesquisando no cadastro de produto valores de venda !!!"
         End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If
   End If

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtValor_Unitario.Text) <> "" Then
            If IsNumeric(txtValor_Unitario.Text) Then
               Call txtDesconto_LostFocus

               CODIGO_BARRAS_A = ""
               FraSeq.Enabled = True
               txtProduto.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If
   CODIGO_BARRAS_A = ""

   txtValor_Unitario.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PRODUTO"
End Sub

Sub MOSTRA_DADOS_PRODUTOold()
'On Error GoTo ERRO_TRATA

   INDR_PROD_BALANCA = False
   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

   If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then _
      INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value

   If INDR_PROD_BALANCA = True Then
      Label13.Caption = "Preço/Kg"
      Else: Label13.Caption = "Preço/UN"
   End If

   txtProduto.Text = Trim(TabProduto.Fields("codg_produto").Value)
   STATUS_PROD = TabProduto!SITUACAO
   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      txtProduto.Text = TabProduto!Codg_Produto
      txtDescricao.Text = TabProduto!DESCRICAO
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            FraSeq.Enabled = True
            txtProduto.SetFocus
            Exit Sub
            Else: txtDescricao.Text = Trim(TabProduto!DESCRICAO)
         End If
   End If

   txtPesoItem.Text = Format(TabProduto.Fields("peso_liquido").Value, strFormatacao3Digitos)
   txtAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
   txtVarejo.Text = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
   STATUS_PROD = TabProduto!SITUACAO

   If Not IsNull(TabProduto!PRECO_Venda) Then
      txtVlrUnit.Text = "" & Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)

      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_Venda
      txtValor_Unitario.Text = Format(Valr_Venda_Produto_n, strFormatacao2Digitos)
      txtPreçoCusto.Text = "" & Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)

      If Trim(cmbTabPrecoAux.Text) <> "" And Trim(cmbFormaAUX.Text) <> "" Then _
         txtPreçoCusto.Text = "" & TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PRODUTO_ID_N, cmbTabPrecoAux.Text, cmbFormaAUX.Text)

      VLR_ANTERIOR_N = TabProduto!PRECO_Venda
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   PRECO_PROD = 0 & txtAtacado.Text

   'If Trim(txtPedido.Text) = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub
If Trim(txtProduto.Text) = "" Then
   Exit Sub
End If

   QTDE_ESTOQUE_N = Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabProduto.Fields("produto_id").Value), strFormatacao3Digitos)
   txtQtdeDisp.Text = "" & Format(QTDE_ESTOQUE_N, strFormatacao3Digitos)
   
   CODG_PRODUTO_A = Trim(txtProduto.Text)

   If INDR_ESTQ_NEGATIVO = False Then
      QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      If QTDE_ESTOQUE_N <= 0 Then
         MsgBox "Produto sem estoque disponível."
         FraSeq.Enabled = True
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Not IsNull(TabProduto.Fields("codg_ncm").Value) Then
      If Len(TabProduto.Fields("codg_ncm").Value) > 2 Then
         If Len(TabProduto.Fields("codg_ncm").Value) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"
            LIMPA_BODY
            FraSeq.Enabled = True
            txtProduto.SetFocus
         End If
      End If
   End If

   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

   'If Trim(txtPedido.Text) = "" Then
   '   MsgBox "Falta numero pedido."
   '   Exit Sub
   'End If

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================

   PEDIDO_ID_N = Trim(txtPedido.Text)

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where PRODUTO_ID = " & PRODUTO_ID_N
   SQL = SQL & " and pedido_ID = " & PEDIDO_ID_N
   SQL = SQL & " and seq_ID = " & Trim(txtSeq.Text)
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      txtValor_Unitario.Text = Format(TabPedidoItem!Valor_Item, strFormatacao2Digitos)
      txtDesconto.Text = Format(TabPedidoItem!PERC_DESC, strFormatacao2Digitos)
      txtQTDE.Text = Format(TabPedidoItem!QTD_PEDIDA, strFormatacao3Digitos)
      QTDE_PEDIDO = TabPedidoItem!QTD_PEDIDA
      VALOR_ITEM_N = TabPedidoItem!Valor_Item
      VALOR_DIFERENCA_N = TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA
      txtSeq.Text = "" & TabPedidoItem.Fields("seq_id").Value

      QTDE_ESTOQUE_N = Format(TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabPedidoItem.Fields("produto_id").Value), strFormatacao3Digitos)
      txtQtdeDisp.Text = "" & Format(QTDE_ESTOQUE_N, strFormatacao3Digitos)
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

'vai pegar preço de tabela
   If Trim(cmbTabPrecoAux.Text) <> "" Then
      If Trim(cmbVendAux.Text) <> "" Then
         If Trim(cmbFormaAUX.Text) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select TABELAPRECOITEM.VALOR_VENDA from TABELAPRECO "
            SQL = SQL & " INNER JOIN TABELAPRECOITEM "
            SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
            SQL = SQL & " INNER JOIN VENDEDOR "
            SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = VENDEDOR.TABELAPRECO_ID "
            SQL = SQL & " INNER JOIN PRODUTO "
            SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " INNER JOIN FORMAPAGTO "
            SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
            SQL = SQL & " where valor_venda > 0 "

            SQL = SQL & " and codg_tabela = '" & cmbTabPrecoAux.Text & "'"
            SQL = SQL & " and vendedor.vendedor_id = " & cmbVendAux.Text
            SQL = SQL & " and TABELAPRECOITEM.formapagto_id = " & cmbFormaAUX.Text
            SQL = SQL & " and TABELAPRECOITEM.produto_id = " & PRODUTO_ID_N

            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               txtValor_Unitario.Text = Format(TabTemp.Fields("VALOR_VENDA").Value, strFormatacao2Digitos)
               txtAtacado.Text = Format(TabTemp.Fields("VALOR_VENDA").Value, strFormatacao2Digitos)
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If
   End If

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtValor_Unitario.Text) <> "" Then
            If IsNumeric(txtValor_Unitario.Text) Then
               Call txtDesconto_LostFocus

               CODIGO_BARRAS_A = ""
               FraSeq.Enabled = True
               txtProduto.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If
   CODIGO_BARRAS_A = ""

   txtValor_Unitario.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PRODUTO"
End Sub

Private Sub ExibirCelula()
'On Error GoTo ERRO_TRATA

   Static OK As Boolean

   If MSFlexGrid1.Col >= 3 And MSFlexGrid1.Col <= 5 Then
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
      Case 3 To 5
         texto = txtValorDig.Text

         If LastCol = 3 Then
            MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
            Else: MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao2Digitos)
         End If

         VALOR_VAREJO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
         VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, LastCol)

'&H80C0FF = LARANJA
'&H8000000F = CINZA
'&HFF& = VERMELHO
'vbBlack 0x0
'vbRed 0xFF
'vbGreen 0xFF00
'vbYellow 0xFFFF
'vbBlue 0xFF0000
'vbMagenta 0xFF00FF
'vbCyan 0xFFFF00
'vbWhite 0xFFFFFF

         If VALOR_ITEM_N < VALOR_VAREJO_N Then
            MSFlexGrid1.CellForeColor = vbRed
            MSFlexGrid1.CellFontBold = True
            MSFlexGrid1.CellBackColor = &H8000000F
            Else
               If VALOR_ITEM_N = VALOR_VAREJO_N Then
                  MSFlexGrid1.CellForeColor = vbBlack
                  MSFlexGrid1.CellFontBold = True
                  MSFlexGrid1.CellBackColor = vbCyan
                  Else
                     MSFlexGrid1.CellForeColor = vbBlue
                     MSFlexGrid1.CellFontBold = True
                     MSFlexGrid1.CellBackColor = vbWhite
               End If
         End If
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
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(8).Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Toolbar1.Buttons(9).Visible = True
      Toolbar1.Buttons(8).Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Sub CLONA_PEDIDO_VENDA()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) <> "" Then
      If IsNumeric(txtPedido.Text) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from PEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and numero_caixa_cpu = " & NUMERO_CAIXA_CPU
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            Msg = "Deseja realmente clonar o pedido de venda : " & txtPedido.Text & " ?"
            PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               GERA_PEDIDO_ID

               SQL = "INSERT INTO PEDIDO "
                  SQL = SQL & "("
                     SQL = SQL & "PEDIDO_ID,Empresa_id, CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, Status, "
                     SQL = SQL & " Tipo_Registro,usuario_id, CLIENTE_ID, Valor_ToTal,"
                     SQL = SQL & " valor_desconto,perc_desc,NUMERO_CAIXA_CPU,estabelecimento_id"
                  SQL = SQL & ") "
                  SQL = SQL & " VALUES ("
                     SQL = SQL & PEDIDO_ID_N
                     SQL = SQL & "," & TabTemp.Fields("empresa_id").Value 'EMPRESA_ID_N
                     SQL = SQL & ",'" & Trim(TabTemp.Fields("cgccpf").Value) & "'"
                     SQL = SQL & "," & TabTemp.Fields("vendedor_id").Value
                     SQL = SQL & ",'" & Now & "'"
                     SQL = SQL & ",'" & Trim(TabTemp.Fields("nome_cliente").Value) & "'"
                     SQL = SQL & "," & 2
                     SQL = SQL & ",'R'"
                     SQL = SQL & "," & TabTemp.Fields("usuario_id").Value
                     SQL = SQL & "," & TabTemp.Fields("cliente_id").Value
                     SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_total").Value)
                     SQL = SQL & "," & tpMOEDA(0)  'vai zerar e tratar somente na tela de desconto
                     SQL = SQL & "," & tpMOEDA(0)
                     SQL = SQL & "," & NUMERO_CAIXA_CPU           'NUMERO_CAIXA_CPU
                     SQL = SQL & "," & ESTABELECIMENTO_ID_N           'estabelecimento_id
               SQL = SQL & ")"
               CONECTA_RETAGUARDA.Execute SQL

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select * from PEDIDOitem WITH (NOLOCK)"
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value

               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               While Not TabConsulta.EOF
                  SQL = "INSERT INTO PEDIDOITEM "
                  SQL = SQL & " ("
                     SQL = SQL & "PEDIDO_ID,SEQ_ID,PRODUTO_ID, Qtd_Pedida,Valor_item, "
                     SQL = SQL & " PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG,PESO_ITEM"
                  SQL = SQL & ") "
                  SQL = SQL & " VALUES ("

                     SQL = SQL & PEDIDO_ID_N                                               'PEDIDO_id
                     SQL = SQL & "," & TabConsulta.Fields("SEQ_ID").Value                 'SEQ_ID
                     SQL = SQL & "," & TabConsulta.Fields("PRODUTO_ID").Value
                     SQL = SQL & "'," & tpMOEDA(TabConsulta.Fields("QTD_PEDIDa").Value)   'Qtd_Pedida
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("VALOR_ITEM").Value)    'Valor_item
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("PERC_DESC").Value)     'PERC_DESC
                     SQL = SQL & "," & tpMOEDA((TabConsulta.Fields("VALOR_ITEM").Value * TabConsulta.Fields("QTD_PEDIDa").Value) * TabConsulta.Fields("PERC_DESC").Value / 100) 'valor_desconto
                     SQL = SQL & ", 'P'"                                                  'status
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("preco_custo").Value)   'PRECO_CUSTO
                     SQL = SQL & ",'PC'"                                                  'TIPO_REG
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTD_PEDIDa").Value)    'PESO_ITEM

                  SQL = SQL & ")"
                  CONECTA_RETAGUARDA.Execute SQL

                  TabConsulta.MoveNext
               Wend

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               MsgBox "Processo realizado com sucesso."
               SqL2 = PEDIDO_ID_N
               LIMPA_TUDO
               PEDIDO_ID_N = SqL2
               Call txtPedido_LostFocus
               FraSeq.Enabled = True
            End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CLONA_PEDIDO_VENDA"
End Sub

Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   Dim TabPedido As New ADODB.Recordset

   If PEDIDO_ID_N > 0 Then
      INDR_RECEITA = 1

      If INDR_FORM_ABERTO = True Then
         Unload frmFatura
         INDR_FORM_ABERTO = False
      End If
'===================================
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select contabiliza from TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " where tipovenda_id = 9999"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
            If TabTemp.Fields("contabiliza").Value = True Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

frmFatura.Show 1

               Else
                  SQL = "update PEDIDO set "
                  SQL = SQL & "status = 6 " 'não contabiliza
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
'===================================
      If INDR_CONTROLA_ESTOQUE = False Then _
         Exit Sub

      If TabPedido.State = 1 Then _
         TabPedido.Close

      SQL = "select * from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
         If TabPedido!STATUS = 5 Then
            CNPJCPF_A = Trim(TabPedido!CGCCPF)
'=============nota eletronica
            'txtCNPJCPF.PromptInclude = False
            If Trim(txtCNPJCPF.Text) = "99999999999" Or Trim(txtCNPJCPF.Text) = "99999999999999" Then
               Else
                  If USA_DOC_FISCAL = True Then
                     If USA_NFe = True Then
                        Msg = "Deseja Gerar Nota Fiscal Eletrônica ?"
                        PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                        If RESPOSTA = vbYes Then
                           If TabPedido.Fields("Status").Value = 5 Or TabPedido.Fields("Status").Value = 7 Then
                              CRITERIO_A = PEDIDO_ID_N
                              TIPO_NFe_GERAR = "R"
                              frmNOTAGERA.Show 1
                           End If
                        End If
                     End If
                  End If
            End If
'=============
'===================================
   If INDR_CONTROLA_ESTOQUE = True Then
'====================
ATUALIZA_ESTOQUE 0, PEDIDO_ID_N
'====================
   End If
         End If
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub

Sub BUSCA_TABELA_PRECO_VENDEDOR(VENDEDOR_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select TABELAPRECO.CODG_TABELA, TABELAPRECO.DESCRICAO, tabelapreco.tabelapreco_id, "
   SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.PRODUTO_ID,"
   SQL = SQL & " TABELAPRECOITEM.FORMAPAGTO_ID, TABELAPRECOITEM.VALOR_VENDA, "
   SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO, TABELAPRECOITEM.PERC_COMISSAO,"
   SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescFormaPagto, FORMAPAGTO.STATUS"
   SQL = SQL & " from TABELAPRECO "
   SQL = SQL & " INNER JOIN TABELAPRECOITEM "
   SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
   SQL = SQL & " INNER JOIN VENDEDOR "
   SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = VENDEDOR.TABELAPRECO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO "
   SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
   SQL = SQL & " where valor_venda > 0 "

   SQL = SQL & " and vendedor_id = " & VENDEDOR_ID_N

   If Trim(cmbTabPrecoAux.Text) <> "" Then _
      SQL = SQL & " and tabelapreco_id = '" & cmbTabPrecoAux.Text & "'"
      'SQL = SQL & " and codg_tabela = '" & cmbTabPrecoAux.Text & "'"

   SQL = SQL & " order by formapagto_id"

   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_TABELA_PRECO_VENDEDOR"
End Sub

Sub PROCESSA_ITEM()
'On Error GoTo ERRO_TRATA

   If Trim(cmbVendAux.Text) = "" Then
      MsgBox "Informe o Vendedor."
      cmbVend.SetFocus
      Exit Sub
   End If
   If Trim(cmbTabPrecoAux.Text) = "" Then
      MsgBox "Selecionar tabela de preço."
      cmbTabPreco.SetFocus
      Exit Sub
   End If
   If Trim(cmbFormaAUX.Text) = "" Then
      MsgBox "Selecionar Faturamento."
      cmbFormaAUX.SetFocus
      Exit Sub
   End If

   If Trim(txtDesconto.Text) <> "" Then
      VALOR_DESCONTO_N = 0 & txtDesconto.Text
      If VALOR_DESCONTO_N > 0 Then
         If STATUS_PROD = "P" Then
            MsgBox "Produto em Promoçao, Impossivel Conseder Desconto"
            txtDesconto.Text = 0
            Else
               'converte tudo para percentual
               If optValor.Value = True Then
                  VALOR_ITEM_N = txtValor_Unitario.Text
                  QTD_N = txtQTDE.Text

                  VALOR_DESCONTO_N = txtDesconto.Text
                  PERC_DESCONTO_N = VALOR_DESCONTO_N * 100 / (VALOR_ITEM_N * QTD_N)
                  Else: PERC_DESCONTO_N = txtDesconto.Text
               End If

CHECA_DESCONTO_USUARIO:

            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "select * from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabUSU.EOF Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close

               MsgBox "Problemas com usuário, codigo=0"
               txtDesconto.SetFocus
               Exit Sub
               Else
                  If Not IsNull(TabUSU!PERC_DESCONTO) Then
                     If PERC_DESCONTO_N > TabUSU!PERC_DESCONTO Then
                        Msg = "Percentual de desconto maior que o permitido para esse usuário. "
                        Msg = Msg & "Percentual cadastrado para " & TabUSU!NOME & " é de " & TabUSU!PERC_DESCONTO & " %. "
                        Msg = Msg & "Deseja liberar com senha superior ?"
                        TabUSU.Close
                        PERGUNTA Msg, vbYesNo + 32, "Desconto NFE", "DEMO.HLP", 1000
                        If RESPOSTA = vbYes Then
                           USUARIO_ATUAL = USUARIO_ID_N
                           frmSenha.Show 1
                           RESPOSTA = ""
                           GoTo CHECA_DESCONTO_USUARIO
                           Exit Sub
                           Else: If USUARIO_ATUAL > 0 _
                                 Then USUARIO_ID_N = USUARIO_ATUAL
                        End If
                        txtDesconto.SetFocus
                        Exit Sub
                        Else
                           If USUARIO_ATUAL > 0 Then _
                              USUARIO_ID_N = USUARIO_ATUAL
                           USU_LIBERA_VENDA_N = TabUSU!USUARIO_ID
                     End If
                     Else
                        If TabUSU.State = 1 Then _
                           TabUSU.Close
                        MsgBox "Percentual de desconto não cadastrado para " & TabUSU!NOME
                        Exit Sub
                  End If
            End If
            If TabUSU.State = 1 Then _
               TabUSU.Close
         End If
         Else '<= 0
            txtDesconto.Text = 0
            PERC_DESCONTO_N = 0
            VALOR_DESCONTO_N = 0
      End If
      Else
         txtDesconto.Text = 0
         PERC_DESCONTO_N = 0
         VALOR_DESCONTO_N = 0
   End If

   If txtPedido.Text = "" Then _
      VALIDA_PEDIDO

   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Text = "99999999999"

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Informe codigo de Produto.", vbOKOnly, "Atenção."
      FraSeq.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If

   If Not IsNull(txtValor_Unitario.Text) Then
      VALOR_ITEM_N = 0 & txtValor_Unitario.Text
      If VALOR_ITEM_N <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção."
         FraSeq.Enabled = True
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtQTDE.Text) = "" Then
      Beep
      MsgBox "Informe a quantidade.", vbOKOnly, "Atenção."
      txtQTDE.SetFocus
      Exit Sub
      Else
         'quantidade pedida
         QTDE_PEDIDO = txtQTDE.Text
         txtQTDE.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)
         If INDR_CONTROLA_ESTOQUE = True Then

            CODG_PRODUTO_A = Trim(txtProduto.Text)

            If INDR_ESTQ_NEGATIVO = False Then
               QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

               If QTDE_ESTOQUE_N < 0 Then
                  Beep
                  MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
                  txtQTDE.SetFocus
                  Exit Sub
               End If
            End If
         End If
         If QTDE_PEDIDO <= 0 Then
            Beep
            MsgBox "Quantidade pedida não permitido, deve ser maior que 0.", vbOKOnly, "Atenção."
            txtQTDE.SetFocus
            Exit Sub
         End If
   End If

   'valor venda item
   VALOR_ITEM_N = txtValor_Unitario.Text
   
   'valor desconto no produto
   If optPerc.Value = True Then
      VALOR_DESCONTO_N = Format(PERC_DESCONTO_N * (VALOR_ITEM_N * QTDE_PEDIDO), strFormatacao2Digitos)
      Else: VALOR_DESCONTO_N = 0 & Format(txtDesconto.Text, strFormatacao2Digitos)
   End If

   VALOR_TOTAL_DESCONTO_N = 0

   'valor total da Pedido, o desconto é armazenado no seu devido lugar, não entra no calculo do campo total da venda
   VALOR_TOTAL_N = VALOR_TOTAL_N + (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      If TabCabeca!STATUS <> 3 Then 'Emitido com Nota
         If TabCabeca!STATUS <> 4 Then ' Emitido com Cupom
            If TabCabeca!STATUS <> 5 Then ' Apenas Faturado
               GRAVA_CABECA "R", 1
            End If
         End If
      End If
      Else 'ainda nao gravou requisicao
         If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
            If Trim(UF_CLIENTE_A) = "" Then
               MsgBox "Cliente com cadastro incompleto !!!"
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If

         GRAVA_CABECA "R", 1
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   If Trim(txtPedido.Text) <> "" Then _
      If IsNumeric(txtPedido.Text) Then _
         GRAVA_TUDO_ITEM ""

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
End Sub

Function TRATA_DESCONTO(Valor_Informado_N As Double) As Boolean
'On Error GoTo ERRO_TRATA

   TRATA_DESCONTO = False
   PERC_DESCONTO_USUARIO_N = 0

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select perc_desconto from USUARIO "
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      If Not IsNull(TabUSU.Fields(0).Value) Then _
         PERC_DESCONTO_USUARIO_N = TabUSU.Fields(0).Value
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If PERC_DESCONTO_USUARIO_N <= 0 Then
      MsgBox "Permissão para desconto não concedida !!!"
      Exit Function
   End If

   VALR_VENDA_N = 0 & txtVarejo.Text
   VALR_ATACADO_N = 0 & txtAtacado.Text

   If VALR_ATACADO_N <= 0 Or VALR_VENDA_N <= 0 Then
      MsgBox "Produto sem valor de venda."
      txtValor_Unitario.Text = 0
      txtVarejo.Text = 0
      txtAtacado.Text = 0
      Exit Function
   End If

   If Valor_Informado_N < VALR_ATACADO_N Then
      Msg = "Valor informado menor que preço de atacado, não permitido !!!, deseja informar senha superior?"
      PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         CRITERIO_A = ""
         frmSenha.Show 1
         If Trim(CRITERIO_A) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from USUARIO WITH (NOLOCK)"
            SQL = SQL & " where senha = '" & Trim(CRITERIO_A) & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If IsNull(TabTemp.Fields("tipo").Value) Then
                  MsgBox "Não permitido."
                  txtValor_Unitario.Text = txtVarejo.Text
                  Exit Function
               End If
               If TabTemp.Fields("tipo").Value >= 4 And TabTemp.Fields("tipo").Value <= 5 Then
                  Else
                     MsgBox "Não permitido."
                     txtValor_Unitario.Text = txtVarejo.Text
                     Exit Function
               End If
               USU_LIBERA_VENDA_N = TabTemp.Fields("usuario_id").Value

               If USU_LIBERA_VENDA_N > 0 Then
                  SQL = "UPDATE PEDIDO SET "
                  SQL = SQL & " USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N

                  SQL = SQL & " where pedido_id = " & txtPedido.Text
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
               Else
                  MsgBox "Não permitido."
                  txtValor_Unitario.Text = txtVarejo.Text
                  txtVarejo.Text = 0
                  txtAtacado.Text = 0
                  Exit Function
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
            Else
               MsgBox "Não permitido."
               txtValor_Unitario.Text = txtVarejo.Text
               txtVarejo.Text = 0
               txtAtacado.Text = 0
               Exit Function
         End If
         Else
            MsgBox "Não permitido."
            txtValor_Unitario.Text = txtVarejo.Text
            txtVarejo.Text = 0
            txtAtacado.Text = 0
            Exit Function
      End If
      'txtValor_Unitario.Text = txtVarejo.Text
   End If

   TRATA_DESCONTO = True
   txtVarejo.Text = 0
   txtAtacado.Text = 0

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_DESCONTO"
End Function

Function TRATA_DESCONTO_GRID(Valor_Informado_N As Double, VLR_GRID_N As Double) As Boolean
'On Error GoTo ERRO_TRATA

   TRATA_DESCONTO_GRID = False
   PERC_DESCONTO_USUARIO_N = 0

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select perc_desconto from USUARIO "
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      If Not IsNull(TabUSU.Fields(0).Value) Then _
         PERC_DESCONTO_USUARIO_N = TabUSU.Fields(0).Value
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If PERC_DESCONTO_USUARIO_N <= 0 Then
      MsgBox "Permissão para desconto não concedida !!!"
      Exit Function
   End If

   VALOR_VENDA_ORIGINAL_N = 0

   VALR_VENDA_N = 0 & txtValor_Unitario.Text
   VALR_ATACADO_N = 0 & txtAtacado.Text

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select preco_venda,produto_id from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      VALOR_VENDA_ORIGINAL_N = 0 & TabProduto.Fields(0).Value
      PRODUTO_ID_N = 0 & TabProduto.Fields(1).Value
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

'   txtValor_Unitario.Text = Format(VALOR_VENDA_ORIGINAL_N, strFormatacao2Digitos)

'vai pegar preço de tabela
   If Trim(cmbTabPrecoAux.Text) <> "" Then
      If Trim(cmbVendAux.Text) <> "" Then
         If Trim(cmbFormaAUX.Text) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select TABELAPRECOITEM.VALOR_VENDA from TABELAPRECO "
            SQL = SQL & " INNER JOIN TABELAPRECOITEM "
            SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID "
            SQL = SQL & " INNER JOIN VENDEDOR "
            SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = VENDEDOR.TABELAPRECO_ID "
            SQL = SQL & " INNER JOIN PRODUTO "
            SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " INNER JOIN FORMAPAGTO "
            SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
            SQL = SQL & " where valor_venda > 0 "

            SQL = SQL & " and codg_tabela = '" & cmbTabPrecoAux.Text & "'"
            SQL = SQL & " and vendedor.vendedor_id = " & cmbVendAux.Text
            SQL = SQL & " and TABELAPRECOITEM.formapagto_id = " & cmbFormaAUX.Text
            SQL = SQL & " and TABELAPRECOITEM.produto_id = " & PRODUTO_ID_N

            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               'txtValor_Unitario.Text = Format(TabTemp.Fields("VALOR_VENDA").Value, strFormatacao2Digitos)
               VALOR_VENDA_ORIGINAL_N = 0 & TabTemp.Fields("VALOR_VENDA").Value
               txtAtacado.Text = Format(TabTemp.Fields("VALOR_VENDA").Value, strFormatacao2Digitos)
               VALR_VENDA_N = 0 & TabTemp.Fields("VALOR_VENDA").Value
               Else  'NÃO ACHOU NA TABELA DE PREÇO VAI PEGAR DO CADASTRO DE PRODUTO
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select preco_venda,produto_id,preco_atacado from PRODUTO WITH (NOLOCK)"
                  SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
                  SQL = SQL & " and situacao <> 'C' "
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabProduto.EOF Then
                     VALR_ATACADO_N = 0 & TabProduto.Fields("preco_atacado").Value
                     txtAtacado.Text = "" & Format(VALR_ATACADO_N, strFormatacao2Digitos)
                     VALR_VENDA_N = 0 & TabProduto.Fields("preco_venda").Value
                     txtVarejo.Text = "" & Format(VALR_VENDA_N, strFormatacao2Digitos)
                     VALOR_VENDA_ORIGINAL_N = 0 & TabProduto.Fields("preco_venda").Value
                     PRODUTO_ID_N = 0 & TabProduto.Fields("produto_id").Value
                  End If
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  txtValor_Unitario.Text = "" & Format(VALOR_VENDA_ORIGINAL_N, strFormatacao2Digitos)

                  MsgBox "Produto sem cadastro na tabela de preço : " & Trim(cmbTabPreco.Text) & " ; pesquisando no cadastro de produto valores de venda !!!"
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      End If
   End If

   If Valor_Informado_N < VALR_VENDA_N Then
      VALOR_DESCONTO_N = 0 & (VALR_VENDA_N - Valor_Informado_N)
      PERC_DESCONTO_N = 0 & ((VALOR_DESCONTO_N / VALR_VENDA_N) * 100)

      If PERC_DESCONTO_N > PERC_DESCONTO_USUARIO_N Then
         Msg = "Valor informado menor que preço de atacado, não permitido !!!, deseja informar senha superior?"
         PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            CRITERIO_A = ""
            frmSenha.Show 1
            If Trim(CRITERIO_A) <> "" Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close
   
               SQL = "select * from USUARIO WITH (NOLOCK)"
               SQL = SQL & " where senha = '" & Trim(CRITERIO_A) & "'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If IsNull(TabTemp.Fields("tipo").Value) Then
                     MsgBox "Não permitido."
                     txtValor_Unitario.Text = VLR_GRID_N
                     Exit Function
                  End If
                  If TabTemp.Fields("tipo").Value >= 4 And TabTemp.Fields("tipo").Value <= 5 Then
                     Else
                        MsgBox "Não permitido."
                        txtValor_Unitario.Text = txtVarejo.Text
                        Exit Function
                  End If
                  USU_LIBERA_VENDA_N = TabTemp.Fields("usuario_id").Value
   
                  If USU_LIBERA_VENDA_N > 0 Then
                     SQL = "UPDATE PEDIDO SET "
                     SQL = SQL & " USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
   
                     SQL = SQL & " where pedido_id = " & txtPedido.Text
                     SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
                  End If
                  Else
                     MsgBox "Não permitido."
                     txtValor_Unitario.Text = txtVarejo.Text
                     txtVarejo.Text = 0
                     txtAtacado.Text = 0
                     Exit Function
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close
               Else
                  MsgBox "Não permitido."
                  txtValor_Unitario.Text = txtVarejo.Text
                  txtVarejo.Text = 0
                  txtAtacado.Text = 0
                  Exit Function
            End If
            Else
               MsgBox "Não permitido."
               txtValor_Unitario.Text = txtVarejo.Text
               txtVarejo.Text = 0
               txtAtacado.Text = 0
               Exit Function
         End If
      End If
   End If

   TRATA_DESCONTO_GRID = True
   txtVarejo.Text = 0
   txtAtacado.Text = 0

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_DESCONTO_GRID"
End Function
