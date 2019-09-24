VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedidoBalcaoPANIFIC 
   Caption         =   "Pedido Venda"
   ClientHeight    =   7950
   ClientLeft      =   2085
   ClientTop       =   2475
   ClientWidth     =   10935
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
   Icon            =   "PEDIDOBALCAOpanif.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtValorDig 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   56
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtTotalPedido 
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
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   5760
      TabIndex        =   53
      Top             =   7500
      Width           =   1575
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
      Height          =   405
      Left            =   7680
      TabIndex        =   51
      Top             =   7500
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
      Height          =   405
      Left            =   3840
      TabIndex        =   49
      Top             =   7500
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
      Height          =   405
      Left            =   2040
      TabIndex        =   47
      Top             =   7500
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
      Height          =   405
      Left            =   120
      TabIndex        =   45
      Top             =   7500
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
      Height          =   405
      Left            =   9360
      TabIndex        =   43
      Top             =   7500
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
      Height          =   1215
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
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3050
         Picture         =   "PEDIDOBALCAOpanif.frx":058A
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
         Width           =   3615
      End
      Begin VB.TextBox txtLIMITE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   7920
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPAGAR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbFaturaAux 
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
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbFatura 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8400
         TabIndex        =   1
         Top             =   240
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
         Left            =   5760
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbVend 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5760
         TabIndex        =   0
         ToolTipText     =   "Selecione um vendedor"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDtEmis 
         Height          =   360
         Left            =   3120
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Cliente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   38
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Crédito"
         Height          =   240
         Left            =   7200
         TabIndex        =   37
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "À Pagar"
         Height          =   240
         Left            =   9000
         TabIndex        =   36
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt.Pedido:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   2025
         TabIndex        =   30
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Fat.:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7800
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Vendedor:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   4560
         TabIndex        =   19
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Pedido:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   15
         Top             =   240
         Width           =   810
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
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   10935
      Begin VB.CommandButton cmdDetalhe 
         BackColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   10320
         Picture         =   "PEDIDOBALCAOpanif.frx":0F8C
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Registrar Detalhes"
         Top             =   840
         Width           =   525
      End
      Begin VB.TextBox txtPesoItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   7800
         TabIndex        =   41
         Top             =   960
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
         Picture         =   "PEDIDOBALCAOpanif.frx":1296
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
         Left            =   840
         MaxLength       =   12
         TabIndex        =   32
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3590
         Picture         =   "PEDIDOBALCAOpanif.frx":20D7
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
         Left            =   1560
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAtacado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optValor 
         Caption         =   "R$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7080
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optPerc 
         Caption         =   "%"
         Height          =   195
         Left            =   6600
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtQTDE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4440
         TabIndex        =   7
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   960
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
         Left            =   3000
         MaxLength       =   12
         TabIndex        =   6
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   960
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
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Se houver algum desconto informe aqui. Pode ser em valor ou em percentual."
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Peso Item"
         Height          =   240
         Left            =   8205
         TabIndex        =   42
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Varejo"
         Height          =   240
         Left            =   1950
         TabIndex        =   23
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Atacado"
         Height          =   240
         Left            =   345
         TabIndex        =   22
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desc."
         Height          =   240
         Left            =   6060
         TabIndex        =   21
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4695
         TabIndex        =   18
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Unitário"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3255
         TabIndex        =   17
         Top             =   720
         Width           =   1080
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
      Height          =   1350
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   2381
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
               Picture         =   "PEDIDOBALCAOpanif.frx":2AD9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":3C73
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":4D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":5CB7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":6DC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":7F18
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":836A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":A1E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":B897
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PEDIDOBALCAOpanif.frx":D879
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
      DesignWidth     =   10935
      DesignHeight    =   7950
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   55
      Top             =   3240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6800
      _Version        =   393216
      AllowUserResizing=   3
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   9240
      X2              =   9240
      Y1              =   7200
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   7200
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   5520
      X2              =   5520
      Y1              =   7200
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   3720
      X2              =   3720
      Y1              =   7200
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   1800
      X2              =   1800
      Y1              =   7200
      Y2              =   7920
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Pedido"
      Height          =   240
      Left            =   6135
      TabIndex        =   54
      Top             =   7222
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Itens Pedido"
      Height          =   240
      Left            =   7965
      TabIndex        =   52
      Top             =   7222
      Width           =   1185
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Desconto"
      Height          =   240
      Left            =   4440
      TabIndex        =   50
      Top             =   7222
      Width           =   870
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor Unitário"
      Height          =   240
      Left            =   2190
      TabIndex        =   48
      Top             =   7222
      Width           =   1320
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "QtdeDisponível"
      Height          =   240
      Left            =   150
      TabIndex        =   46
      Top             =   7222
      Width           =   1440
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Peso Total (Kg)"
      Height          =   240
      Left            =   9390
      TabIndex        =   44
      Top             =   7222
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      Height          =   750
      Left            =   0
      Top             =   7200
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
Attribute VB_Name = "frmPedidoBalcaoPANIFIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim UF_CLIENTE             As String
   Dim UF_EMPRESA             As String
   Dim strInscEstadual        As String
   Dim dblTipoCliente         As Double
   Dim strCPFCNPJ             As String
   Dim bolRequisicaoJaExiste  As Boolean
   Dim rstEmpresa             As New ADODB.Recordset
   Dim seq_n                  As Long
   Dim PRECO_PROD             As Double
   Dim CLIENTE_ID_N           As Long
   Dim TIPO_NOTA_A            As String
   Dim VALOR_UNITARIO_N       As Double
   Dim TP2_DE_CONTRIB         As Double
   Dim TP2_DE_NCONTRIB        As Double
   Dim TP2_DE_CMAQ_IMP        As Double
   Dim TP2_DE_NMAQ_IMP        As Double
   Dim TP2_FE_CMAQ_IMP        As Double
   Dim TP2_FE_NMAQ_IMP        As Double
   Dim TP2_FE_CAP_INDU        As Double
   Dim TP2_FE_NAP_INDU        As Double
   Dim CFOP_SAIDA_DE          As String
   Dim CFOP_SAIDA_FE          As String
   Dim strCFOP                As String
   Dim SITUAÇÃO_TRIBUTARIA_PRODUTO
   Dim INDR_PROD_BALANCA      As Boolean

   Dim Valr_Venda_Produto_n   As Double
   Dim QTDE_N                 As Double
   Dim PESO_ITEM_N            As Double
   Dim CODIGO_BARRAS          As String

   Private CalculaIcmsG       As New MegasimCL.mCalculaIcms ' Yuri alterado em 01/05/2012

   Dim TabGridVaca         As New ADODB.Recordset
   Private LastRow         As Long ' Ultima linha em que se editou
   Private LastCol         As Long ' ultima coluna em que se editou
   Private ControlVisible  As Boolean
   Dim PRECO_CUSTO_N       As Double

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INICIALIZA_VENDA
   MOSTRA_VENDEDORES

   Call txtPedido_LostFocus

   OcultarControles

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF8
         frmCADASTROCLIENTE.Show 1
         If NOME_A <> "" Then _
            txtNome.Text = NOME_A
         NOME_A = ""
      Case vbKeyF10
         INDR_GRAVA = False
         If Trim(txtPedido.Text) = "" Then _
            Exit Sub
         If Not IsNumeric(txtPedido.Text) Then _
            Exit Sub

         NUMR_REQ_N = txtPedido.Text

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
            VALIDA_NUMR_REQ
            INICIALIZA_VENDA
         End If
      Case "gravar"
         INDR_GRAVA = False
         If txtPedido.Text <> "" Then
            NUMR_REQ_N = txtPedido.Text
            Else
               MsgBox "Digite Numero da Requisicao para gravar!"
               Exit Sub

               Call txtPedido_LostFocus
         End If

         GERA_VENDA
         LIMPA_TUDO

         Call txtPedido_LostFocus
      Case "consultar"
         frmPedidoConsulta.Show 1
         If NUMR_REQ_N > 0 Then
            Dim NUMR_PEDIDO_N As Long

            NUMR_PEDIDO_N = NUMR_REQ_N

            LIMPA_TUDO
            txtPedido.Text = NUMR_PEDIDO_N
            CRITERIO = ""
            NUMR_PEDIDO_N = 0
            Call txtPedido_LostFocus
         End If
         FraSeq.Enabled = True
         txtProduto.SetFocus
      Case "print"
         GERA_IMPRESSAO
      Case "gravar"
         INDR_GRAVA = False
         NUMR_REQ_N = txtPedido.Text

         GERA_VENDA
         LIMPA_TUDO

         Call txtPedido_LostFocus
      Case "limpar"
         LIMPA_TUDO

         Call txtPedido_LostFocus
         FraSeq.Enabled = True
         txtProduto.SetFocus
      Case "voltar"
         Unload Me
      Case "produto"
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
            frmCADASTROPRODUTO.Show 1
            Else: CHAMA_PRODUTO_SIMPLIFICADO
         End If
      Case "CadCliente"
          frmCADASTROCLIENTE.Show 1
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

   SQL = "select pedido_id,seq_id from PEDIDOITEM "
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
   frmDISPLAYCLIENTE.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = ""
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.Text = CNPJCPF_A
      Call TXTCNPJCPF_LostFocus
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
End Sub

Private Sub cmdMata_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then
      EXCLUIR_ITEM Trim(txtProduto.Text), Trim(txtPedido.Text), Trim(txtSeq.Text)
      Else: MsgBox "Informe código produto."
   End If

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

Private Sub cmbFATURA_Click()
'On Error GoTo ERRO_TRATA

   cmbFaturaAux.ListIndex = cmbFatura.ListIndex
   If cmbFaturaAux.Text <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & cmbFaturaAux.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp!parcela) Then _
            NUMR_PARCELA = TabTemp!parcela
         If Not IsNull(TabTemp!PRAZO) Then _
            DIAS_PRAZO = TabTemp!PRAZO
      End If
      Else
         MsgBox "Selecione tipo de venda."
         Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFATURA_Click"
End Sub

Private Sub cmbFATURA_GotFocus()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA "
   SQL = SQL & " order by TIPOVENDA_ID desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbFatura.AddItem Trim(TabTemp!Descricao)
      cmbFaturaAux.AddItem Trim(TabTemp!TipoVenda_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MOSTRA_TOP "ESC - SAIR", "Selecione Tipo Venda", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFATURA_GotFocus"
End Sub

Private Sub cmbFATURA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFATURA_KeyPress"
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

Private Sub cmbVend_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbVend.Text) = "" Then _
      MOSTRA_VENDEDORES

   MOSTRA_TOP "ESC - SAIR", "Selecione Vendedor e tecle <ENTER>", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVend_GotFocus"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex
   'cmbFatura.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVend_Click"
End Sub

Private Sub cmbvend_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbFatura.SetFocus
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_KeyPress"
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
   txtProduto.SetFocus
End Sub

Private Sub txtITENS_GotFocus()
   txtProduto.SetFocus
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtNome.Text <> "" Then
      txtNome.SelStart = 0
      txtNome.SelLength = Len(txtNome)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtPesoTotal_GotFocus()
   txtProduto.SetFocus
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
                           txtQtde.Text = VALOR_RECEBIDO_N / VALOR_ITEM_N
                           txtQtde.Refresh
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
   txtProduto.SetFocus
End Sub

Private Sub txtTotalPedido_GotFocus()
   txtProduto.SetFocus
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

Private Sub TXTCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Text = "99999999999"
      Else
         If Trim(txtCNPJCPF.Text) <> "99999999999" Then _
            TRATA_CLIENTE
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtDesconto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Informe desconto unitário", "F10 - Gravar", "", ""

   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Else: txtProduto.SetFocus
   End If

   txtDesconto.SelStart = 0
   txtDesconto.SelLength = Len(txtQtde)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_GotFocus"
End Sub

Private Sub txtDesconto_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtDesconto.Text) <> "" Then _
      txtDesconto.Text = Format(txtDesconto.Text, strFormatacao2Digitos)
   If UF_CLIENTE = "" Then _
      TRATA_CLIENTE

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
'==================cgccpf
Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "F7-Consulta Clientes", "Inform CNPJ/CPF Cliente e Tecle <<Enter>>", "", ""
   txtNome.Enabled = True
   txtCNPJCPF.Mask = "###############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtCNPJCPF.Text = ""
         frmDISPLAYCLIENTE.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.Text = ""
            txtCNPJCPF.Mask = "##############"

            txtCNPJCPF.Text = CNPJCPF_A
            Call TXTCNPJCPF_LostFocus
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

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         txtNome.Enabled = True
         txtNome.SetFocus
         Else
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

Private Sub txtnome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      UCase (txtProduto.Text)
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtPedido.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then _
            EXCLUIR_ITEM Trim(txtProduto.Text), Trim(txtPedido.Text), Trim(txtSeq.Text)
         txtProduto.SetFocus
      Case vbKeyF7
         CONSULTA_PRODUTO
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
      If INDR_PANIFICADORA = True And Trim(txtProduto.Text) <> "" And Trim(txtDescricao.Text) <> "" Then _
         txtQtde.SetFocus
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
      txtProduto.SetFocus
      Exit Sub
   End If
   If Trim(txtQtde.Text) <> "" Then
      txtQtde.SelStart = 0
      txtQtde.SelLength = Len(txtQtde.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      If Len(Trim(txtQtde.Text)) > 10 Then
         txtProduto.SetFocus
         Exit Sub
      End If
      txtDesconto.SetFocus
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

   If Len(Trim(txtQtde.Text)) > 10 Then
      txtProduto.SetFocus
      Exit Sub
   End If

   If Trim(txtQtde.Text) = "" Then
      txtQtde.Text = 1
      Else
         If IsNumeric(txtQtde.Text) Then
            QTDE_N = txtQtde.Text
            If QTDE_N <= 0 Then _
               txtQtde.Text = 1
         End If
   End If
   txtQtde.Text = Format(txtQtde.Text, strFormatacao3Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_LostFocus"
End Sub

Private Sub txtPEDIDO_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "Tecle <ENTER> para gerar nova Pedido ou informe uma já existente", "", "", ""

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

      If Trim(cmbFaturaAux.Text) = "" Then
         cmbFaturaAux.Text = 9999
         cmbFatura.Text = "A Vista"
      End If

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

   VALIDA_NUMR_REQ

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_LostFocus"
End Sub

Private Sub TXTVALOR_UNITARIO_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Informe Valor Unitário", "", "", ""
   
   txtValor_Unitario.SelStart = 0
   txtValor_Unitario.SelLength = Len(txtValor_Unitario.Text)

   'If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
   '   Else: SendKeys "{tab}"
   'End If

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
            Else: txtQtde.SetFocus
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

            SQL = "select * from USUARIO "
            SQL = SQL & " where usuario_id = " & CODG_USU_N
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabUSU.EOF Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close

               MsgBox "Problemas com usuário, codigo=0"
               txtDesconto.SetFocus
               Exit Sub
               Else: txtQtde.SetFocus
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

Private Sub txtdesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
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

   Dim Valr_Atacado  As Double
   Dim Valr_Digitado As Double
   Dim Valr_Venda    As Double

   VALOR_UNITARIO_N = 0 & txtValor_Unitario.Text
   If Trim(txtValor_Unitario.Text) = "" Then
      txtValor_Unitario.Text = Format(0, strFormatacao2Digitos)
      Else: txtValor_Unitario.Text = Format(VALOR_UNITARIO_N, strFormatacao2Digitos)
   End If
   If VALOR_UNITARIO_N <= 0 Then
      'MsgBox "Valor Unitário Inválido !!!"
      'txtValor_Unitario.SetFocus
      txtProduto.SetFocus
      Exit Sub
      Else
         VALOR_ITEM_N = txtValor_Unitario.Text
         txtValor_Unitario.Text = Format(VALOR_UNITARIO_N, strFormatacao2Digitos)
         If VALOR_ITEM_N <= 0 Then
            MsgBox "Valor Unitário Inválido !!!"
            txtProduto.SetFocus
            Exit Sub
         End If
   End If

   Valr_Venda = 0 & txtVarejo.Text
   Valr_Atacado = 0 & txtAtacado.Text

   If Valr_Atacado <= 0 Or Valr_Venda <= 0 Then
      MsgBox "Produto sem valor de venda."
      txtValor_Unitario.Text = 0
   End If

   Valr_Digitado = 0 & txtValor_Unitario.Text

   If Valr_Digitado < Valr_Atacado Then
      Msg = "Valor informado menor que preço de atacado, não permitido !!!, deseja informar senha superior?"
      PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         CRITERIO = ""
            frmSenha.Show 1
            If Trim(CRITERIO) <> "" Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from USUARIO "
               SQL = SQL & " where senha = '" & Trim(CRITERIO) & "'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If IsNull(TabTemp.Fields("tipo").Value) Then
                     MsgBox "Não permitido."
                     txtValor_Unitario.Text = txtVarejo.Text
                     Exit Sub
                  End If
                  If TabTemp.Fields("tipo").Value >= 4 Or TabTemp.Fields("tipo").Value <= 5 Then
                     Else
                        MsgBox "Não permitido."
                        txtValor_Unitario.Text = txtVarejo.Text
                        Exit Sub
                  End If
                  USU_LIBERA_VENDA_N = TabTemp.Fields("usuario_id").Value
                  Exit Sub
                  Else
                     MsgBox "Não permitido."
                     txtValor_Unitario.Text = txtVarejo.Text
                     Exit Sub
               End If

               If TabTemp.State = 1 Then _
                  TabTemp.Close
            End If
      End If
      txtValor_Unitario.Text = txtVarejo.Text
   End If

   If Trim(txtValor_Unitario.Text) <> "" Then _
      If IsNumeric(txtValor_Unitario.Text) Then _
         txtValor_Unitario.Text = Format(txtValor_Unitario.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_LostFocus"
End Sub

Private Sub txtVlrUnit_GotFocus()
   txtProduto.SetFocus
End Sub
'============================subrotinas
Private Sub EXCLUIR_ITEM(CODG_PRODUTO_A As String, PEDIDO_ID_N As Long, SEQ_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Trim(PEDIDO_ID_N) > 0 And Trim(SEQ_ID_N) > 0 And Trim(CODG_PRODUTO_A) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
   
      SQL = "SELECT PEDIDOITEM.*, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
      SQL = SQL & " PRODUTO.QTDE, PRODUTO.PRECO_VENDA, PRODUTO.PRECO_CUSTO, "
      SQL = SQL & " Produto.Situacao_Tributaria , Produto.QTDE_RETIDO"
      SQL = SQL & " FROM PEDIDO "
      SQL = SQL & " INNER JOIN PEDIDOITEM "
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   
      SQL = SQL & " where codg_prod = '" & Trim(CODG_PRODUTO_A) & "'"
      SQL = SQL & " and PEDIDOITEM.numr_req = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_ID_N
   
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

            BAIXA_RETIDO TabTemp!QTD_PEDIDA

            SQL = "Delete FROM PEDIDOITEM "
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

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM"
End Sub

Private Sub MOSTRA_DADOS_REQ()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Text = TabCABECA!CGCCPF

   'MOSTRA VENDEDOR
   If Not IsNull(TabCABECA!VENDEDOR_ID) Then
      SP_PROCURA_VENDEDOR 0, TabCABECA!VENDEDOR_ID, "", "", "", "", ""
      If Not TabVENDEDOR.EOF Then _
         cmbVend.Text = TabVENDEDOR!NOME_VEND

      cmbVendAux.Text = TabCABECA!VENDEDOR_ID

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If

   If Not IsNull(TabCABECA!TipoVenda_ID) Then
      If TabTipoVenda.State = 1 Then _
         If TabTipoVenda.State = 1 Then _
            TabTipoVenda.Close

      cmbFaturaAux.Text = TabCABECA!TipoVenda_ID

      SQL = "select * from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & cmbFaturaAux.Text
      TabTipoVenda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTipoVenda.EOF Then _
         cmbFatura.Text = TabTipoVenda!Descricao

      If TabTipoVenda.State = 1 Then _
         TabTipoVenda.Close
   End If

   'MOSTRA CLIENTE
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome,status from CLIENTE "
   SQL = SQL & " where cgccpf = '" & TabCABECA!CGCCPF & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      If TabCABECA!CGCCPF = "99999999999" Then
         If Not IsNull(TabCABECA!NOME_CLIENTE) Then
            If Trim(txtNome.Text) = "" Then _
               txtNome.Text = TabCABECA!NOME_CLIENTE
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

   SQL = "select nome_cliente from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
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

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   Aliquota_Icms = 0
   Valr_Venda_Produto_n = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""
   txtQtdeDisp.Text = ""

   QTDE_PEDIDO = 0
   QTDE_ESTOQUE = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_DIFERENCA_N = 0
   PRODUTO_ID_N = 0

   txtAtacado.Text = Format(0, strFormatacao2Digitos)
   txtVarejo.Text = Format(0, strFormatacao2Digitos)
   txtValor_Unitario.Text = Format(0, strFormatacao2Digitos)
   txtPreçoCusto.Text = Format(0, strFormatacao2Digitos)
   txtQtde.Text = Format(0, strFormatacao3Digitos)
   txtDesconto.Text = Format(0, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   MOSTRA_VENDEDORES

   MSFlexGrid1.Clear

   txtValorDig.Visible = False
   FraSeq.Enabled = False
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
   txtQtdeDisp.Text = ""

   PRODUTO_ID_N = 0
   Aliquota_Icms = 0
   txtPedido.Text = ""
   txtDtEmis = Format(Date, "dd/mm/yyyy")
   txtNome.Text = ""
   txtCNPJCPF.Text = ""
   cmbFatura.Text = ""
   cmbFaturaAux.Text = ""
   LIMPA_BODY
   
   VALOR_TOTAL_N = 0
   NUMR_REQ_N = 0
   QTDE_PEDIDO = 0
   QTDE_ESTOQUE = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_TOTAL_N = 0
   USU_LIBERA_VENDA_N = 0
   txtLIMITE.Text = ""
   txtPAGAR.Text = ""
   SINAL_INDICADOR_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub MOSTRA_VENDEDORES()
'On Error GoTo ERRO_TRATA

   cmbVend.Clear
   cmbVendAux.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SP_PROCURA_VENDEDOR 0, 0, "", "", "", "", "A"
   While Not TabVENDEDOR.EOF
      cmbVend.AddItem Trim(TabVENDEDOR!NOME_VEND) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
      TabVENDEDOR.MoveNext
   Wend
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbVend.Enabled = True
      Else
         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select logon from USUARIO "
         SQL = SQL & " where usuario_id = " & CODG_USU_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            cmbVend.Enabled = False

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            CRITERIO = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
            SQL = "select nome_vend, vendedor_id from VENDEDOR "
            SQL = SQL & " where nome_vend like " & CRITERIO
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVend.Text = Trim(TabVENDEDOR!NOME_VEND)
               cmbVendAux.Text = Trim(TabVENDEDOR!VENDEDOR_ID)
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDORES"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear

   MSFlexGrid1.GridLines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns
   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabGridVaca.State = 1 Then _
      TabGridVaca.Close

   SQL = "SELECT PEDIDOITEM.CODG_PROD as 'Código', "                                                                                '0
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

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.numr_req = " & txtPedido.Text
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N

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
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGridVaca.Fields(Coluna).Value, strFormatacao3Digitos)
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
      MSFlexGrid1.ColWidth(0) = 1500
      MSFlexGrid1.ColAlignment(0) = 0

'Referencia
      MSFlexGrid1.ColWidth(1) = 1000
      MSFlexGrid1.ColAlignment(1) = 0

'Descrição Produto
      MSFlexGrid1.ColWidth(2) = 6000
      MSFlexGrid1.ColAlignment(2) = 0

'QTDE
      MSFlexGrid1.ColWidth(3) = 1500
      MSFlexGrid1.ColAlignment(3) = 7

'Valor Item
      MSFlexGrid1.ColWidth(4) = 1500
      MSFlexGrid1.ColAlignment(4) = 7

'Desconto
      MSFlexGrid1.ColWidth(5) = 1500
      MSFlexGrid1.ColAlignment(5) = 7

'Total Item
      MSFlexGrid1.ColWidth(6) = 1500
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

Private Sub GRAVA_CABECA(TIPO_REGISTRO_A As String, STATUS_N As Integer)
'On Error GoTo ERRO_TRATA

   CRITERIO = ""
   CLIENTE_ID_N = 0

   txtCNPJCPF.Mask = "###############"

   If cmbFaturaAux.Text = "" Then _
      cmbFaturaAux.Text = 1

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select cliente_id from CLIENTE"
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

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PEDIDO "
   SQL = SQL & " where numr_req = " & txtPedido.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCABECA.EOF Then
      SQL = "INSERT INTO PEDIDO "
         SQL = SQL & "(PEDIDO_ID,Empresa_id, numr_req, CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, Status, "
         SQL = SQL & " Tipo_Registro,Codg_USU, TIPOVENDA_ID, CLIENTE_ID, Valor_ToTal,"
         SQL = SQL & " valor_desconto,perc_desc) "
         SQL = SQL & " VALUES ("
            SQL = SQL & PEDIDO_ID_N
            SQL = SQL & "," & EMPRESA_ID_N
            SQL = SQL & "," & txtPedido.Text
            SQL = SQL & ",'" & txtCNPJCPF.Text & "'"
            SQL = SQL & "," & cmbVendAux.Text & ","
            SQL = SQL & "'" & DMA(Date) & "'"
            SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
            SQL = SQL & "," & STATUS_N
            SQL = SQL & ",'" & TIPO_REGISTRO_A & "'"
            SQL = SQL & "," & CODG_USU_N
            SQL = SQL & "," & cmbFaturaAux.Text
            SQL = SQL & "," & CLIENTE_ID_N
            SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)
            SQL = SQL & "," & tpMOEDA(0)  'vai zerar e tratar somente na tela de desconto
            SQL = SQL & "," & tpMOEDA(0)
         SQL = SQL & ")"
      Else
         PEDIDO_ID_N = 0 & TabCABECA.Fields("pedido_id").Value
         txtPedido.Text = PEDIDO_ID_N

         If Not IsNull(TabCABECA!Status) Then
            If TabCABECA!Status <> 3 Then
               If TabCABECA!Status <> 4 Then
                  If TabCABECA!Status <> 5 Then
                     If TabCABECA!Status <> 9 Then
                        SQL = "UPDATE PEDIDO SET "
                        SQL = SQL & " Valor_total = " & tpMOEDA(VALOR_TOTAL_N)
                        SQL = SQL & ",numr_req = " & txtPedido.Text
                        SQL = SQL & ",Valor_desconto = " & tpMOEDA(0)   'vai zerar e tratar somente na tela de desconto
                        SQL = SQL & ",Perc_desc = " & tpMOEDA(0)
                        SQL = SQL & ",CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
                        SQL = SQL & ",Vendedor_id = " & cmbVendAux.Text
                        SQL = SQL & ",dt_req = '" & DMA(Date) & "'"
                        SQL = SQL & ",nome_cliente = '" & txtNome.Text & "'"
                        SQL = SQL & ",Status = " & STATUS_N
                        SQL = SQL & ",TIPO_REGISTRO = '" & TIPO_REGISTRO_A & "'"
                        SQL = SQL & ",CODG_USU = " & CODG_USU_N
                        SQL = SQL & ",EMPRESA_ID = " & EMPRESA_ID_N
                        SQL = SQL & ",TIPOvenda_id = " & cmbFaturaAux.Text
                        SQL = SQL & ",USUARIO_LIBERA_VENDA = " & CODG_USU_N
                        SQL = SQL & ",CLIENTE_ID = " & CLIENTE_ID_N

                        SQL = SQL & " where numr_req = " & txtPedido.Text
                        SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                     End If
                  End If
               End If
            End If
         End If
   End If

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA"
End Sub

Private Sub GRAVA_TUDO_ITEM()
'On Error GoTo ERRO_TRATA

   'Tratamento da tributacao
   'fazer no final desta rotina
   'CODG_PRODUTO_A = Trim(txtProduto.Text)

   If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
      If Trim(UF_CLIENTE) = "" Then
         MsgBox "Cliente com cadastro incompleto !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtPreçoCusto.Text) = "" Then _
      txtPreçoCusto.Text = 0

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

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and seq_id = " & SEQ_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabPedidoItem.EOF Then
      SQL = "INSERT INTO PEDIDOITEM "
      SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Numr_req, Codg_Prod, Qtd_Pedida,Valor_item, "
      SQL = SQL & " PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG,PESO_ITEM) "
      SQL = SQL & " VALUES ("

         SQL = SQL & PEDIDO_ID_N                                                          'PEDIDO_id
         SQL = SQL & "," & SEQ_ID_N                                                       'SEQ_ID
         SQL = SQL & "," & PRODUTO_ID_N
         SQL = SQL & "," & txtPedido.Text                                                 'Numr_req
         SQL = SQL & ",'" & Trim(txtProduto.Text)                                         'Codg_Prod
         SQL = SQL & "'," & tpMOEDA(QTDE_PEDIDO)                                          'Qtd_Pedida
         SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                                          'Valor_item
         SQL = SQL & "," & tpMOEDA(PERC_DESCONTO_N)                                       'PERC_DESC
         SQL = SQL & "," & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)  'valor_desconto
         SQL = SQL & ", 'P'"                                                              'status
         SQL = SQL & "," & tpMOEDA(txtPreçoCusto.Text)                                    'PRECO_CUSTO
         SQL = SQL & ",'PC'"                                                              'TIPO_REG
         SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                                           'PESO_ITEM

      SQL = SQL & ")"
      Else
         SQL = "UPDATE PEDIDOITEM SET "
         SQL = SQL & " qtd_pedida = " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & ", Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
         SQL = SQL & ", PERC_desc = " & tpMOEDA(PERC_DESCONTO_N)
         SQL = SQL & ", valor_desconto = " & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)
         SQL = SQL & ", status = 'P'"
         SQL = SQL & ", preco_custo = " & tpMOEDA(txtPreçoCusto.Text)
         SQL = SQL & ", PESO_ITEM = " & tpMOEDA(QTDE_PEDIDO)

         SQL = SQL & " Where numr_req = " & txtPedido.Text
         SQL = SQL & " and pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
   End If
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   CONECTA_RETAGUARDA.Execute SQL

   'Atualiza Qt Balcao
   SQL = "UPDATE Produto SET "
   SQL = SQL & " qtde_retido = qtde_retido + " & tpMOEDA(QTDE_PEDIDO)
   SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"
   CONECTA_RETAGUARDA.Execute SQL

   'Tratamento da tributacao
   CODG_PRODUTO_A = Trim(txtProduto.Text)

   PREPARA_TRIBUTAÇÃO_PRODUTO Trim(txtCNPJCPF.Text)

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO_ITEM"
End Sub

Private Sub PEGA_DADOS_EMPRESA()
'On Error GoTo ERRO_TRATA

   Dim RstTemp As New ADODB.Recordset
   Dim strTemp As String
   Dim CEP_EMP_A As String

   If rstEmpresa.State = 1 Then _
      rstEmpresa.Close

   SQL = "Select TP2_DE_CONTRIB,TP2_DE_NCONTRIB,TP2_DE_CMAQ_IMP,"
   SQL = SQL & " TP2_DE_NMAQ_IMP,TP2_FE_CMAQ_IMP,TP2_FE_NMAQ_IMP,"
   SQL = SQL & " TP2_FE_CAP_INDU,TP2_FE_NAP_INDU,"
   SQL = SQL & " CFOP_SAIDA_DE,CFOP_SAIDA_FE"

   SQL = SQL & " From EMPRESA "
   SQL = SQL & " where EMPRESA_ID = " & EMPRESA_ID_N
   SQL = SQL & " and cgc = '" & Trim(CNPJ_GERAL) & "'"
   rstEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstEmpresa.EOF Then
      If rstEmpresa.State = 1 Then _
         rstEmpresa.Close

      MsgBox "O sistema não obteve sucesso ao tentar localizar a empresa corrente."
      Unload Me
      Exit Sub
   End If

   ' yuri 01/05/2012 para pegar tambem outras informações referentes a importos
   'g_trabalhacomtare_empresa = rstEmpresa!optante_tare não retirar sergio vamos precisar
   'so to colocando aqui com comentário para nao te atrapalhar
   
   TP2_DE_CONTRIB = rstEmpresa!TP2_DE_CONTRIB
   TP2_DE_NCONTRIB = rstEmpresa!TP2_DE_NCONTRIB
   TP2_DE_CMAQ_IMP = rstEmpresa!TP2_DE_CMAQ_IMP
   TP2_DE_NMAQ_IMP = rstEmpresa!TP2_DE_NMAQ_IMP
   TP2_FE_CMAQ_IMP = rstEmpresa!TP2_FE_CMAQ_IMP
   TP2_FE_NMAQ_IMP = rstEmpresa!TP2_FE_NMAQ_IMP
   TP2_FE_CAP_INDU = rstEmpresa!TP2_FE_CAP_INDU
   TP2_FE_NAP_INDU = rstEmpresa!TP2_FE_NAP_INDU
   CFOP_SAIDA_DE = rstEmpresa!CFOP_SAIDA_DE
   CFOP_SAIDA_FE = rstEmpresa!CFOP_SAIDA_FE

   If RstTemp.State = 1 Then _
      RstTemp.Close

   SQL = "Select * From ENDERECO "

   'SQL = SQL & " Where PROP = '" & Trim(rstEmpresa!CGC) & "'"
   SQL = SQL & " Where PROP = '" & Trim(CNPJ_GERAL) & "'"

   RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not RstTemp.EOF Then
      If Not IsNull(RstTemp!CEP) Then
         CEP_EMP_A = "" & RstTemp!CEP
         Else
            If rstEmpresa.State = 1 Then _
               rstEmpresa.Close
            
            If RstTemp.State = 1 Then _
               RstTemp.Close

            MsgBox "Verificar cadastro de empresa !!! " & CEP_EMP_A
            Unload Me
            Exit Sub
      End If
      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "Select * From CEP "
      SQL = SQL & " Where CEP = " & CEP_EMP_A
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not RstTemp.EOF Then
         UF_EMPRESA = RstTemp!UF
         Else
            If RstTemp.State = 1 Then _
               RstTemp.Close

            If rstEmpresa.State = 1 Then _
               rstEmpresa.Close

            MsgBox "Verificar cadastro de empresa, endereço não cadastrado"
            End
            Exit Sub
      End If
      Else
         If rstEmpresa.State = 1 Then _
            rstEmpresa.Close

         If RstTemp.State = 1 Then _
            RstTemp.Close

         MsgBox "Verificar cadastro de empresa, endereço não cadastrado"
         End
         Exit Sub
   End If
   If RstTemp.State = 1 Then _
      RstTemp.Close

   If rstEmpresa.State = 1 Then _
      rstEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PEGA_DADOS_EMPRESA"
End Sub

Private Sub GERA_VENDA()
'On Error GoTo ERRO_TRATA

   Dim strimpressoa As String

   PERC_DESCONTO_USUARIO = 0
   VALOR_TOTAL_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   USU_LIBERA_VENDA_N = 0

   If INDR_LIBERA_DESCONTO = True Then
      Msg = "Deseja informar desconto ?"
      PERGUNTA Msg, vbYesNo + 32, "Desconto NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then _
         LIBERA_DESCONTO
   End If

   PERC_DESCONTO_USUARIO = 0
   NUMR_REQ_N = txtPedido.Text
   CNPJCPF_A = txtCNPJCPF.Text

   'atualizando desconto na cabeça
   SQL = "UPDATE PEDIDO SET "
   SQL = SQL & " Valor_desconto = " & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
   SQL = SQL & " , Perc_desc = " & tpMOEDA(PERC_DESCONTO_N)
   SQL = SQL & " , cgccpf = '" & CNPJCPF_A & "'"
   SQL = SQL & " , nome_cliente = '" & Trim(txtNome.Text) & "'"
   SQL = SQL & " , status = 2"
   SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N

    If Trim(cmbFaturaAux.Text) <> "" Then _
        If IsNumeric(cmbFaturaAux.Text) Then _
            SQL = SQL & " , tipovenda_id = " & Trim(cmbFaturaAux.Text)

   SQL = SQL & " where numr_req = " & txtPedido.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
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
      FORMULA_REL = "{vwRelVenda.empresa_id} = " & EMPRESA_ID_N
      FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & txtPedido.Text

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_pedido_venda.rpt"
      frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_VENDA"
End Sub

Private Sub BAIXA_RETIDO(QTDE_BAIXAR As Double)
'On Error GoTo ERRO_TRATA

   If QTDE_BAIXAR > 0 Then

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select qtde_retido from PRODUTO "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         If TabProduto!QTDE_RETIDO >= QTDE_BAIXAR Then
            SQL = "UPDATE Produto SET "
            SQL = SQL & " qtde_retido = qtde_retido - " & tpMOEDA(QTDE_BAIXAR)
            SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
            CONECTA_RETAGUARDA.Execute SQL
         End If
      End If

      If TabProduto.State = 1 Then _
         TabProduto.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_RETIDO"
End Sub

Private Sub BAIXA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   SQL = "select PEDIDOITEM.* "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.numr_req = " & NUMR_REQ_N
   SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      SQL = "UPDATE Produto SET "
      SQL = SQL & " qtde = qtde - " & tpMOEDA(QTDE_PEDIDO)
      SQL = SQL & ", qtde_retido = qtde_retido - " & tpMOEDA(QTDE_PEDIDO)
      SQL = SQL & ", DT_ULT_VENDA =  '" & DMA(Date) & "'"
      SQL = SQL & "  Where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"
      CONECTA_RETAGUARDA.Execute SQL

      TabPedidoItem.MoveNext
   Wend

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_ESTOQUE"
End Sub

Public Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub

Public Sub VALIDA_NUMR_REQ()
'On Error GoTo ERRO_TRATA

   NUMR_REQ_N = 1

   If Trim(txtPedido.Text) = "" Then
      GERA_NUMR_REQ

      txtPedido.Text = NUMR_REQ_N
      Else
         txtPedido.Enabled = True
            NUMR_REQ_N = txtPedido.Text
         txtPedido.Enabled = False
   End If

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      bolRequisicaoJaExiste = False

      NUMR_REQ_N = txtPedido.Text

      bolRequisicaoJaExiste = True

      MOSTRA_DADOS_REQ

      CRITERIO = ""

      txtDtEmis.Text = TabCABECA!DT_REQ

      If TabCABECA!Status = 9 Then
         MsgBox "Pedido cancelada, impossível alterar !!!"
         Exit Sub
         Else '1=ORÇAMENTO;2=GERADO;3=EMITIDA COM NOTA;4=EMITIDA COM CUPOM;5=ARECEBER;7=ECF/NF;9=CANCELADO
            If (TabCABECA!Status = 3 Or TabCABECA!Status = 5) Then
               If TabCABECA!Status = 3 Then
                  Toolbar1.Buttons(3).Visible = False
                  Toolbar1.Buttons(8).Visible = False
                  If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
                     Toolbar1.Buttons(9).Visible = True

                  PERGUNTA "Nota Processada para este pedido.", vbNo, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     If TabCABECA.State = 1 Then _
                        TabCABECA.Close

                     Else
                        FraSeq.Enabled = False
                        'LIMPA_BODY
                        'LIMPA_TUDO
                   End If
                   Exit Sub
               End If
               If TabCABECA!Status = 5 Then
                  Toolbar1.Buttons(3).Visible = False
                  Toolbar1.Buttons(8).Visible = False
                  If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
                     Toolbar1.Buttons(9).Visible = True

                  PERGUNTA "Venda ja Faturada, Deseja imprimir ?", vbYesNo + 32, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     GERA_IMPRESSAO
                     Else
                        FraSeq.Enabled = False
                        'LIMPA_BODY
                        'LIMPA_TUDO
                   End If
               End If
               Exit Sub
            End If
            If TabCABECA!Status = 4 Then
               MsgBox "Permitido somente consulta, cupom fiscal emitido."
               Exit Sub
            End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VALIDA_NUMR_REQ"
End Sub

Sub PROCESSA_ITEM()
'On Error GoTo ERRO_TRATA

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
                  QTD_N = txtQtde.Text

                  VALOR_DESCONTO_N = txtDesconto.Text
                  PERC_DESCONTO_N = VALOR_DESCONTO_N * 100 / (VALOR_ITEM_N * QTD_N)
                  Else: PERC_DESCONTO_N = txtDesconto.Text
               End If

CHECA_DESCONTO_USUARIO:

            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "select * from USUARIO "
            SQL = SQL & " where usuario_id = " & CODG_USU_N
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
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
                           USUARIO_ATUAL = CODG_USU_N
                           frmSenha.Show 1
                           RESPOSTA = ""
                           GoTo CHECA_DESCONTO_USUARIO
                           Exit Sub
                           Else: If USUARIO_ATUAL > 0 _
                                 Then CODG_USU_N = USUARIO_ATUAL
                        End If
                        txtDesconto.SetFocus
                        Exit Sub
                        Else
                           If USUARIO_ATUAL > 0 Then _
                              CODG_USU_N = USUARIO_ATUAL
                           USU_LIBERA_VENDA_N = TabUSU!usuario_id
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

   If Trim(cmbFaturaAux.Text) = "" Then _
      cmbFaturaAux.Text = 9999

   If cmbVendAux.Text = "" Then
      cmbVend.Text = "Balcão"
      cmbVendAux.Text = 0
   End If

   If txtPedido.Text = "" Then _
      VALIDA_NUMR_REQ

   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Text = "99999999999"

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Informe codigo de Produto.", vbOKOnly, "Atenção."
      txtProduto.SetFocus
      Exit Sub
   End If

   If Not IsNull(txtValor_Unitario.Text) Then
      VALOR_ITEM_N = 0 & txtValor_Unitario.Text
      If VALOR_ITEM_N <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção."
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtQtde.Text) = "" Then
      Beep
      MsgBox "Informe a quantidade.", vbOKOnly, "Atenção."
      txtQtde.SetFocus
      Exit Sub
      Else
         'quantidade pedida
         QTDE_PEDIDO = txtQtde.Text
         txtQtde.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)
         If INDR_CONTROLA_ESTOQUE = True Then

            CODG_PRODUTO_A = Trim(txtProduto.Text)

            If INDR_ESTQ_NEGATIVO = False Then
               CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO

               If QTDE_ESTOQUE < 0 Then
                  Beep
                  MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
                  txtQtde.SetFocus
                  Exit Sub
               End If
            End If
         End If
         If QTDE_PEDIDO <= 0 Then
            Beep
            MsgBox "Quantidade pedida não permitido, deve ser maior que 0.", vbOKOnly, "Atenção."
            txtQtde.SetFocus
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

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      If TabCABECA!Status <> 3 Then 'Emitido com Nota
         If TabCABECA!Status <> 4 Then ' Emitido com Cupom
            If TabCABECA!Status <> 5 Then ' Apenas Faturado
               GRAVA_CABECA "R", 1
               'grava itens
               If Trim(txtPedido.Text) <> "" Then
                  If IsNumeric(txtPedido.Text) Then
                     GRAVA_TUDO_ITEM

                     If INDR_BAIXA_ESTQ_PEDIDO = True Then _
                        BAIXA_ESTOQUE
                  End If
               End If
            End If
         End If
      End If
      Else 'ainda nao gravou requisicao
         If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
            If UF_CLIENTE = "" Then
               MsgBox "Cliente com cadastro incompleto !!!"
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If

         GRAVA_CABECA "R", 1

         If Trim(txtPedido.Text) <> "" Then
            If IsNumeric(txtPedido.Text) Then
               GRAVA_TUDO_ITEM

               If INDR_BAIXA_ESTQ_PEDIDO = True Then _
                  BAIXA_ESTOQUE
            End If
         End If
   End If

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
End Sub

Sub QUALIFICA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbVend.Enabled = True
      Else
         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select logon from USUARIO "
         SQL = SQL & " where usuario_id = " & CODG_USU_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            cmbVend.Enabled = False

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            CRITERIO = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
            SQL = "select nome_vend, vendedor_id from VENDEDOR "
            SQL = SQL & " where nome_vend like " & CRITERIO
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVend.Text = TabVENDEDOR!NOME_VEND
               cmbVendAux.Text = TabVENDEDOR!VENDEDOR_ID
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
         If TabUSU.State = 1 Then _
            TabUSU.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "QUALIFICA_VENDEDOR"
End Sub

Sub GERA_IMPRESSAO()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" Then
      NUMR_REQ_N = txtPedido.Text
      Else: NUMR_REQ_N = InputBox(SQL3, "Informe número de Pedido a ser impressa ")
   End If

   FORMULA_REL = "{vwRelVenda.empresa_id} = " & EMPRESA_ID_N
   FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & NUMR_REQ_N

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_pedido_venda.rpt"
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
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Sub TRATA_CLIENTE()
'On Error GoTo ERRO_TRATA

   Dim rstCliente       As New ADODB.Recordset
   Dim rstAux           As New ADODB.Recordset
   Dim rstEndereco      As New ADODB.Recordset
   Dim rstCep           As New ADODB.Recordset

   Dim VALOR_LIMITE_N   As Double
   Dim VALOR_PENDENTE_N As Double

   ENDERECO_A = ""
   PESSOA_ID_N = 0
   If txtCNPJCPF.Text = "" Then
      txtCNPJCPF.Text = "99999999999"
      Else
         If CHECA_CNPJCPF(Trim(txtCNPJCPF.Text)) = True Then
            CRITERIO = txtCNPJCPF.Text
            Else
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCNPJCPF = ""
               txtCNPJCPF.SetFocus
               Exit Sub
         End If
   End If

   If Trim(txtCNPJCPF.Text) <> "" Then
      CRITERIO = Trim(txtCNPJCPF.Text)
      If Not IsNull(txtCNPJCPF.Text) Then
         If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
      End If
      txtCNPJCPF.Text = CRITERIO
   End If

   txtNome.Enabled = True

   If rstCliente.State = 1 Then _
      rstCliente.Close

   SQL = "select pessoa_id,nome,cliente_id,limite_credito,tipo_cliente,cgccpf,ie,Status from CLIENTE "
   SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
   SQL = SQL & " and status = 'A'"
   rstCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstCliente.EOF Then
      If rstCliente.State = 1 Then _
         rstCliente.Close

      Beep
      MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção."
      txtCNPJCPF.SetFocus
      Exit Sub
      Else
         txtNome.Enabled = False
         PESSOA_ID_N = rstCliente.Fields("pessoa_id").Value

         If Trim(rstCliente!NOME) <> "" And Trim(txtCNPJCPF.Text) <> "99999999999" Then _
            txtNome.Text = rstCliente!NOME

         SQL = "update PEDIDO set nome_cliente = '" & Trim(txtNome.Text) & "'"
         SQL = SQL & " where numr_req = " & NUMR_REQ_N
         CONECTA_RETAGUARDA.Execute SQL

         CLIENTE_ID_N = rstCliente.Fields("cliente_id").Value
         If Not IsNull(rstCliente!limite_credito) Then _
            txtLIMITE.Text = Format(rstCliente!limite_credito, strFormatacao2Digitos)

         'Pegou o tipo do cliente
         If Not IsNull(rstCliente!TIPO_CLIENTE) Then _
            dblTipoCliente = rstCliente!TIPO_CLIENTE

         If Not IsNull(rstCliente!CGCCPF) Then _
            strCPFCNPJ = rstCliente!CGCCPF

         If Not IsNull(rstCliente!IE) Then 'O Cara ja tem no Cadastro de Cliente
            strInscEstadual = rstCliente!IE
            Else ' Se ele nao tiver no Cadastro de Cliente pega aqui!
               If rstCliente.State = 1 Then _
                  rstCliente.Close
               MsgBox "Inscrição estatual invalida para este cliente, atualizar."
               Exit Sub
         End If

         If rstAux.State = 1 Then _
            rstAux.Close

         SQL = "select sum(i.valor_item) from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = l.numr_doc "
         SQL = SQL & " and l.pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and i.status = 'A' "
         SQL = SQL & " and l.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = 1"
         SQL = SQL & " and i.formapagto_id <> 1"
         rstAux.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not rstAux.EOF Then
            If Not IsNull(rstAux.Fields(0).Value) Then
               VALOR_PENDENTE_N = 0 & rstAux.Fields(0).Value
               txtPAGAR.Text = Format(rstAux.Fields(0).Value, strFormatacao2Digitos)
               txtPAGAR.Refresh
            End If
         End If
         If rstAux.State = 1 Then _
            rstAux.Close

         VALOR_LIMITE_N = 0 & rstCliente.Fields("LIMITE_CREDITO").Value

         If VALOR_LIMITE_N > 0 Then
            If VALOR_PENDENTE_N >= VALOR_LIMITE_N Then
               MsgBox "Valor limite de credito para esse cliente ultrapassado, não permitido venda, verificar com departamento financeiro."
               txtCNPJCPF.Text = ""
               txtNome.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If

         If rstEndereco.State = 1 Then _
            rstEndereco.Close

         SQL = "select * from ENDERECO "
         SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & " and tipo = 'C'"
         rstEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not rstEndereco.EOF Then
            If Not IsNull(rstEndereco!Rua) Then _
               ENDERECO_A = rstEndereco!Rua
            If Not IsNull(rstEndereco!Complemento) Then _
               ENDERECO_A = ENDERECO_A & "," & rstEndereco!Complemento
            If Not IsNull(rstEndereco!Bairro) Then _
               ENDERECO_A = ENDERECO_A & "," & rstEndereco!Bairro

            'Pegou o CEP do cliente
            If IsNull(rstEndereco!CEP) Then
               If rstEndereco.State = 1 Then _
                  rstEndereco.Close
   
               MsgBox "O Cadastro do cliente não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
               txtCNPJCPF.Text = ""
               txtNome.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
            If rstCep.State = 1 Then _
               rstCep.Close
      
            'Pegar a uf do cliente
            rstCep.Open "Select * From CEP Where CEP = " & rstEndereco!CEP, CONECTA_RETAGUARDA, , , adCmdText
            If Not rstCep.EOF Then
               If Not IsNull(rstCep!UF) Then
                  UF_CLIENTE = rstCep!UF
                  Else 'UF nao localizada
                     If rstCep.State = 1 Then _
                        rstCep.Close
                     MsgBox "O Cadastro do cliente não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                     txtCNPJCPF.Text = ""
                     txtNome.Text = ""
                     txtCNPJCPF.SetFocus
                     Exit Sub
               End If
               Else
                  If rstCep.State = 1 Then _
                     rstCep.Close

                  MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais completos. Verique-os, principalmente o Estado(UF) da empresa"
                  txtCNPJCPF.Text = ""
                  txtNome.Text = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End If
            If rstCep.State = 1 Then _
               rstCep.Close
         End If
         If rstEndereco.State = 1 Then _
            rstEndereco.Close

         If rstCliente!Status = "C" Then
            If rstCliente.State = 1 Then _
               rstCliente.Close

            Beep
            MsgBox "Cliente Esta Bloqueado!, Verifique Cadastro!.", vbOKOnly, "Atenção."
            txtCNPJCPF.Text = ""
            txtNome.Text = ""
            txtCNPJCPF.SetFocus
            Exit Sub
         End If
   End If
   If rstCliente.State = 1 Then _
      rstCliente.Close

   SQL = "select nome_cliente from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   rstCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not rstCliente.EOF Then _
      If Not IsNull(rstCliente.Fields(0).Value) Then _
         txtNome.Text = Trim(rstCliente.Fields(0).Value)
   If rstCliente.State = 1 Then _
      rstCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_CLIENTE"
End Sub

Private Sub PREPARA_TRIBUTAÇÃO_PRODUTO(ClienteId As String)
'On Error GoTo ERRO_TRATA

'Duvidas
'- 13/06/2006 Quando o item for subsituicao ou do tipo tributario = 60, ele terá dois valores de icms
' ou somente um valor. Exemplificando, se for 100,00 ele tera uma aliquota de 17% e outra de 10% por exemplo
'ou somente sera cobrado uma aliquota? Pergunto isto pois se houver dois valores para o mesmo item devera
'ser criado um outro registro no banco de dados.

   Dim rstProduto                As New ADODB.Recordset
   Dim RstTemp                   As New ADODB.Recordset

   Dim strSql                    As String
   Dim dblValorBaseICMS          As Double
   Dim dblPercentualICMS         As Double
   Dim VALOR_ICMS_PRODUTO        As Double
   Dim dblValorBaseICMSSubst     As Double
   Dim VALOR_ICMS_PRODUTOSubst   As Double
   Dim dblPercentualICMSSubst    As Double
   Dim dblPercReducICMS          As Double
   Dim dblPercIVA                As Double
   Dim dblTotalItem              As Double

   If CODG_PRODUTO_A = "" Or ClienteId = "" Then
      MsgBox "O sistema esta esperando alguns parametros que nao forma  localizados. Verifique"
      Exit Sub
   End If

   dblValorBaseICMS = 0
   dblPercentualICMS = 0
   VALOR_ICMS_PRODUTO = 0
   dblValorBaseICMSSubst = 0
   VALOR_ICMS_PRODUTOSubst = 0
   dblPercentualICMSSubst = 0
   dblPercReducICMS = 0
   dblPercIVA = 0
   dblTotalItem = 0
   strCFOP = ""
   SITUAÇÃO_TRIBUTARIA_PRODUTO = ""

   If UF_CLIENTE = "" Then _
      TRATA_CLIENTE

   If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
      If UF_CLIENTE = "" Then
         MsgBox "Cliente com cadastro incompleto !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
   End If

   If UF_EMPRESA = "" Then _
      PEGA_DADOS_EMPRESA

   dblTotalItem = (txtQtde.Text * txtValor_Unitario.Text)

   If rstProduto.State = 1 Then _
      rstProduto.Close

   strSql = "Select situacao_tributaria,perciva,comp_tributaria From PRODUTO "
   strSql = strSql & " Where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   strSql = strSql & " And EMPRESA_ID = " & EMPRESA_ID_N
   strSql = strSql & " and situacao <> 'C' "
   rstProduto.Open strSql, CONECTA_RETAGUARDA, , , adCmdText
   If rstProduto.EOF Then
      If rstProduto.State = 1 Then _
         rstProduto.Close

      MsgBox "O sistema nao localizou nenhum produto com o seguinte codigo: " & CODG_PRODUTO_A & vbCrLf & "Verique"
      Exit Sub
   End If

   'Inicio yuri 01/05/2012
     ' Aqui será colocado a rotina para calcular os tributos em substituição a toda essa regra que esta
     ' nesta instrução
     ' busca aliquota do Unidade federativa do Cliente
     ' aqui nao retirar aqui vamos dar o inicio a toda carga tributaria
     ' comentei aqui para nao atraplhar se codigo
   'Call BuscaAliquota(strUFCliente, CLng(ClienteId))

   ' fim yuri 01/05/2012

   'Tentando fazer igual o dataflex faz
   '//Impostos  Tributos
   '// ---- Calculo das Reducoes de ICMS e Substituicao Tributaria -------- //
    '  //0 = Tributado integralmente
    '  //1 = Tributado e com cobranca do ICMS por Substituicao Tributaria
    '  //2 = Com Reducao de Base de Calculo
    '  //3 = Isenta ou nao tributada e com cobranca do ICMS por Sub. Tributaria
    '  //4 = Isenta ou nao Tributado
    '  //5 = Com Suspensao ou diferimento
    '  //6 = ICMS cobrado anteriormente por subst. Tributaria
    '  //7 = Com reducao de base de Calculo e Cobranca do icms por Subst. Tributaria
    '  //9 = Outras
    '  //Compensacao Tribuaria
    '  //0 = Mercadorias Normais
    '  //1 = Maquinas e Implementos Agricolas
    '  //2 = Maquinas Aparelhos Equipamentos Industriais

'==========banco de dados
'CODIGO  DESCRICAO
'00      Tributada integralmente
'10      Tributada  e com cobrança do ICMS por substituição tributária
'20      Com redução de base de cálculo
'30      Isenta ou não tributada e com cobrança do ICMS por substituição tributária
'40      Isenta
'41      Não tributada
'50      Suspensão
'51      Diferimento
'60      ICMS cobrado anteriormente por substituição tributária
'70      Com redução de base de cálculo e cobrança de ICMS por substituição tributária
'90      Outras
'==========banco de dados

   'Tributada integralmente
   If rstProduto!SITUACAO_TRIBUTARIA = "00" Then
      'Desconto nao entra no valor do ICMS de acordo com informacoes
      'da CONTABILIDADE
      dblValorBaseICMS = dblTotalItem

      'Criar campo de TIPO DE CLIENTE NO CADASTRO DE CLIENTE
      If dblTipoCliente = 2 Then
         If UF_CLIENTE = UF_EMPRESA Then
            dblValorBaseICMS = ((dblTotalItem * TP2_DE_CONTRIB) / 100)  'Valor da Reducao da base
            dblPercentualICMS = TP2_DE_CONTRIB                ' Percentual da reducao
         End If
      End If
   End If

   'Tributada e com cobrança do ICMS por substituição tributária
   If rstProduto!SITUACAO_TRIBUTARIA = 10 Then 'Substituicao Tributaria
      dblValorBaseICMS = dblTotalItem

      If UF_CLIENTE = UF_EMPRESA Then
         'Campo IVA nao existe nao tabela verificar se precisa
         If Not IsNull(rstProduto!PERCIVA) Then _
           dblValorBaseICMSSubst = ((dblValorBaseICMS * rstProduto!PERCIVA) / 100)  'Valor da Reducao da base

         'dblValorBaseICMSSubst = ((dblValorBaseICMS * 1) / 100)  'Valor da Reducao da base
         VALOR_ICMS_PRODUTOSubst = ((dblValorBaseICMSSubst * 17) / 100) 'é fixo o percentual, procurar saber se tem como parametrizar
         dblPercentualICMSSubst = 17
      End If
   End If

   'Com redução de base de cálculo
   If rstProduto!SITUACAO_TRIBUTARIA = 20 Then 'Reducao da base de calculo
      If rstProduto!COMP_TRIBUTARIA = 0 Then 'tipos de maquinas, normais, agricolas, industriais
         If strInscEstadual <> "" Then   'Tem que ter inscricao estadual
            dblValorBaseICMS = ((dblTotalItem * TP2_DE_CONTRIB) / 100)
            dblPercReducICMS = TP2_DE_CONTRIB
            Else  'Sem inscricao estadual
               dblValorBaseICMS = ((dblTotalItem * TP2_DE_NCONTRIB) / 100)
               dblPercReducICMS = TP2_DE_NCONTRIB
         End If
      End If

      'Maquinas agricolas
      If rstProduto!COMP_TRIBUTARIA = 1 Then
         If UF_CLIENTE = UF_EMPRESA Then 'Dentro do estado
            If strInscEstadual <> "" Then
               dblValorBaseICMS = ((dblTotalItem * TP2_DE_CMAQ_IMP) / 100)
               dblPercReducICMS = TP2_DE_CMAQ_IMP
               Else
                  dblValorBaseICMS = ((dblTotalItem * TP2_DE_NMAQ_IMP) / 100)
                  dblPercReducICMS = TP2_DE_NMAQ_IMP
            End If
            Else 'Fora do Estado
               If strInscEstadual <> "" Then
                  dblValorBaseICMS = ((dblTotalItem * TP2_FE_CMAQ_IMP) / 100)
                  dblPercReducICMS = TP2_FE_CMAQ_IMP
                  Else
                     dblValorBaseICMS = ((dblTotalItem * TP2_FE_NMAQ_IMP) / 100)
                     dblPercReducICMS = TP2_FE_NMAQ_IMP
               End If
         End If
      End If

      If rstProduto!COMP_TRIBUTARIA = 2 Then 'Maquinas industriais
         If UF_CLIENTE = UF_EMPRESA Then 'Dentro do estado
            If strInscEstadual <> "" Then
               dblValorBaseICMS = ((dblTotalItem * TP2_DE_CONTRIB) / 100)
               dblPercReducICMS = TP2_DE_CONTRIB
               Else
                  dblValorBaseICMS = ((dblTotalItem * TP2_DE_NCONTRIB) / 100)
                  dblPercReducICMS = TP2_DE_NCONTRIB
            End If
            Else 'Fora do Estado
               If strInscEstadual <> "" Then
                  dblValorBaseICMS = ((dblTotalItem * TP2_FE_CAP_INDU) / 100)
                  dblPercReducICMS = TP2_FE_CAP_INDU
                  Else
                     dblValorBaseICMS = ((dblTotalItem * TP2_FE_NAP_INDU) / 100)
                     dblPercReducICMS = TP2_FE_NAP_INDU
               End If
         End If
      End If
   End If

   'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
   If rstProduto!SITUACAO_TRIBUTARIA = 30 Then '//Isenta ou nao Tributada Com ICMS por Subs. Trib
      dblValorBaseICMS = 0
      dblPercentualICMS = 0

      If UCase(UF_CLIENTE) <> UCase(UF_EMPRESA) Then
          '//Desconto nao entra no valor de ICMS de Acordo com as
          '//Informacoes Contabeis
          '//move (ITENS.TOTAL_ITEM - ITENS.VLR_DESC_RATEIO)  ;
          '//                                     To   ITENS.VLR_BASE_ICMS
          dblValorBaseICMS = dblTotalItem
          '??? nao grava o percentual do aliquota?
      End If
   End If

   'Isenta ou Não tributada
   If rstProduto!SITUACAO_TRIBUTARIA = 40 Or rstProduto!SITUACAO_TRIBUTARIA = 41 Then '//Isento ou nao Tributado
      dblValorBaseICMS = 0
      dblPercentualICMS = 0
   End If

'50      Suspensão
'51      Diferimento

   'ICMS cobrado anteriormente por substituição tributária
   If rstProduto!SITUACAO_TRIBUTARIA = 60 Then '//Situacao Tributaria com Substitui‡ao Tributaria
      '//Desconto nao entra no valor de ICMS de Acordo com as
      '//Informacoes Contabeis

      dblValorBaseICMS = dblTotalItem
      If UCase(UF_CLIENTE) = UCase(UF_EMPRESA) Then
         If dblTipoCliente = 2 Then 'Atacado
            '//Dentro do Estado e Cliente Contribuinte ele e Isento
            '/Emanoel Informacoes Contabilidade dia 30/05/2006
            dblValorBaseICMS = 0
            dblPercentualICMS = 0
         End If

         'Só é tratado o tipo de cliente 2, atacado, e os outros tipos de clientes (varejo),
         'nao precisa tratar?
         Else 'Fora do estado
            If dblTipoCliente = 2 Then 'Atacado
               dblValorBaseICMS = dblTotalItem
               'nao grava o percentual? porque?
            End If
      End If
   End If

'70      Com redução de base de cálculo e cobrança de ICMS por substituição tributária
'90      Outras

'========================================================================
'========================================================================
'========================================================================

   'If Not IsNull(rstProduto.Fields("cfop").Value) Then
      
   'End If

   'DENTRO DO ESTADO
   If UCase(UF_CLIENTE) = UCase(UF_EMPRESA) Then
      If rstProduto!SITUACAO_TRIBUTARIA = 60 Then
         'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
         'CFOP 5405 - Venda de mercadoria adquirida/recebida de terceiros em operação _
                      com mercadoria sujeita ao regime de substituição tributária, na condição de _
                      contrib substituído
 
'portanto o que vai diferenciar se será um codigo ou outro será a mercadoria em
'si...se ela é substituiçao tributaria ou nao...se for varias mercadorias vc tem que
'verificar uma por uma pra saber.

         strCFOP = "5405"
         Else: strCFOP = CFOP_SAIDA_DE                     'cfop de venda dentro do estado
      End If

      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "Select * From CFOP "
      SQL = SQL & " Where codigo = '" & Trim(strCFOP) & "'"
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If RstTemp.EOF Then
         If RstTemp.State = 1 Then _
            RstTemp.Close

         If rstProduto.State = 1 Then _
            rstProduto.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      'if rstTEMP!Tipo = 0 then 'Dentro do Estado
      VALOR_ICMS_PRODUTO = ((dblTotalItem * RstTemp!PERC_ICMS) / 100)
      dblPercentualICMS = RstTemp!PERC_ICMS

      If RstTemp.State = 1 Then _
         RstTemp.Close
   End If

   'FORA DO ESTADO
   If UCase(UF_CLIENTE) <> UCase(UF_EMPRESA) Then
      If rstProduto!SITUACAO_TRIBUTARIA = 60 Then
         strCFOP = "6403"  'Fixo por enquanto
         '6403 Venda de mercadoria adquirida ou recebida de terceiros em operação _
               com mercadoria sujeita ao regime de substituição tributária, _
               na condição de contribuinte substituto _
               Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros, _
               na condição de contribuinte substituto, em operação com mercadorias sujeitas _
               ao regime de substituição tributária.

         strCFOP = "6404"
         '6404 Venda de mercadoria sujeita ao regime de substituição tributária, _
               cujo imposto já tenha sido retido anteriormente _
               Classificam-se neste código as vendas de mercadorias sujeitas ao regime de substituição tributária, _
               na condição de substituto tributário, exclusivamente nas hipóteses em que o _
               imposto já tenha sido retido anteriormente

         Else: strCFOP = CFOP_SAIDA_FE                  'cfop de venda fora do estado do estado
      End If

      SQL = "Select * From CFOP "
      SQL = SQL & " Where CODIGO = '" & Trim(strCFOP) & "'"
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If RstTemp.EOF Then
         If RstTemp.State = 1 Then _
            RstTemp.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      If Trim(Len(strCPFCNPJ)) > 11 Then ' Se for pessoa juridica
         VALOR_ICMS_PRODUTO = ((dblTotalItem * RstTemp!PERC_ICMS) / 100)  'CFOP.P_ICMS_VND_F_UF - verificar se existe
         dblPercentualICMS = RstTemp!PERC_ICMS ' CFOP.P_ICMS_VND_F_UF'duas aliquotas para  o mesmo cfop
         Else ' Pessoa fisica
            VALOR_ICMS_PRODUTO = ((dblTotalItem * RstTemp!ICMS_PJ_F_UF) / 100)
            dblPercentualICMS = RstTemp!ICMS_PJ_F_UF
      End If

      If RstTemp.State = 1 Then _
         RstTemp.Close
   End If

   'HOJE 12/06/2006 22:00
   'FALTA VERIFICAR SE EXISTE DUAS ALIQUOTAS PARA O MESMO CFOP
   'FALTA GRAVAR OS DADOS CORRETAMENTE NA TABELA
   'FALTA VER O LANCE ABAIXO
   
   'Ver depois com o emanoel para que estes campos
   'se for necessarario mesmo, acho que criarei um campo asc de tamanho x
   ' vou appendando os CFOPS que existir separando-os com com um ';"
   'farei uma funcao para tratar os cfops appendando depois
   '   //Testa Cfop para Cabeca!
   '   if PRODUTOS.COD_TRIBUTACAO eq 60 begin
   '      if CIDADE.UF eq DOCUMENT.UF begin
   '         move 5405                               To   CFOP1_D
   '      End
   '      if CIDADE.UF ne DOCUMENT.UF move 6403      To   CFOP1_F
   '   End
   '   if PRODUTOS.COD_TRIBUTACAO ne 60 begin
   '      if CIDADE.UF eq DOCUMENT.UF begin
   '          move CFOP.VND_MERC_D_UF                To   CFOP_D
   '      End
   '      if CIDADE.UF ne DOCUMENT.UF move CFOP.VND_MERC_F_UF;
   '                                                 To   CFOP_F
   '   End

   SITUAÇÃO_TRIBUTARIA_PRODUTO = "" & rstProduto!SITUACAO_TRIBUTARIA

   'If Not isnull(rstProduto!PERCIVA) Then dblPercIVA = rstProduto!PERCIVA

   If dblValorBaseICMS = 0 Then _
      dblPercentualICMS = 0
   
   If rstProduto.State = 1 Then _
      rstProduto.Close

   If RstTemp.State = 1 Then _
      RstTemp.Close

   SQL = "Select PEDIDO.pedido_id FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " Where PEDIDO.EMPRESA_ID = " & EMPRESA_ID_N
   SQL = SQL & " And PEDIDO.NUMR_REQ = " & txtPedido.Text
   SQL = SQL & " And CODG_PROD = '" & Trim(txtProduto.Text) & "'"

   RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not RstTemp.EOF Then
      PEDIDO_ID_N = RstTemp.Fields(0).Value

      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "UPDATE PEDIDOITEM SET "
      SQL = SQL & " VlrBaseIcms = " & tpMOEDA(dblValorBaseICMS)
      SQL = SQL & ", PERCICMS = " & tpMOEDA(dblPercentualICMS)
      SQL = SQL & ", VlrIcms = " & tpMOEDA(VALOR_ICMS_PRODUTO)
      SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(dblValorBaseICMSSubst)
      SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(dblPercentualICMSSubst)
      SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(VALOR_ICMS_PRODUTOSubst)
      SQL = SQL & ", cfop = '" & strCFOP & "'"
      SQL = SQL & ", STRIBUTARIA = '" & SITUAÇÃO_TRIBUTARIA_PRODUTO & "'"
      SQL = SQL & ", status = 'P'"

      SQL = SQL & " Where numr_req = " & txtPedido.Text
      SQL = SQL & " and pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"

      CONECTA_RETAGUARDA.Execute SQL
   End If
   If RstTemp.State = 1 Then _
      RstTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREPARA_TRIBUTAÇÃO_PRODUTO"
End Sub

Sub INICIALIZA_VENDA()
'On Error GoTo ERRO_TRATA
   
   Me.Caption = Me.Caption & " - " & Me.Name
   
   UF_CLIENTE = ""  'Variavel para tratamento Fiscal do item
   UF_EMPRESA = "" 'Variavel para tratamento Fiscal do item
   strInscEstadual = "" 'Variavel para tratamento Fiscal do item
   dblTipoCliente = -1 'Variavel para tratamento fiscal do item
   strCPFCNPJ = ""
   'bolRequisicaoJaExiste = False 'Indica se a requisicao atual é nova, ou se ja
                                 'esta no banco de dados ou nao.
   
   txtDtEmis = Format(Date, "dd/mm/yyyy")
   
   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close
   
   SQL = "select * from EMPRESA "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmpresa.EOF Then
      If Not IsNull(TabEmpresa!baixa_estoque_req) Then
         INDR_BAIXA_ESTQ_PEDIDO = TabEmpresa!baixa_estoque_req
         Else: INDR_BAIXA_ESTQ_PEDIDO = False
      End If
   End If
   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   PEGA_DADOS_EMPRESA
   QUALIFICA_VENDEDOR

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_VENDA"
End Sub

Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   INDR_PROD_BALANCA = False

   If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then _
      INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value

   txtProduto.Text = Trim(TabProduto.Fields("codg_produto").Value)
   STATUS_PROD = TabProduto!SITUACAO
   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      txtProduto.Text = TabProduto!CODG_PRODUTO
      txtDescricao.Text = TabProduto!Descricao
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            txtProduto.SetFocus
            Exit Sub
            Else: txtDescricao.Text = Trim(TabProduto!Descricao)
         End If
   End If

   txtPesoItem.Text = Format(TabProduto.Fields("peso_liquido").Value, strFormatacao3Digitos)
   txtAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
   txtVarejo.Text = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
   STATUS_PROD = TabProduto!SITUACAO

   If Not IsNull(TabProduto.Fields("qtde_retido").Value) Then _
      QTDE_ESTOQUE = TabProduto!QTDE - TabProduto.Fields("qtde_retido").Value

   If Not IsNull(TabProduto!PRECO_VENDA) Then
      txtVlrUnit.Text = "" & Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)

      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_VENDA
      txtValor_Unitario.Text = Format(Valr_Venda_Produto_n, strFormatacao2Digitos)
      txtPreçoCusto.Text = 0 & Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)

      VLR_ANTERIOR_N = TabProduto!PRECO_VENDA
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   PRECO_PROD = 0 & txtAtacado.Text

   If txtPedido.Text = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

   txtQtdeDisp.Text = Format(QTDE_ESTOQUE, strFormatacao3Digitos)
   CODG_PRODUTO_A = Trim(txtProduto.Text)

   If INDR_ESTQ_NEGATIVO = False Then
      CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO

      If QTDE_ESTOQUE <= 0 Then
         MsgBox "Produto sem estoque disponível."
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Not IsNull(TabProduto.Fields("codg_ncm").Value) Then
      If Len(TabProduto.Fields("codg_ncm").Value) > 2 Then
         If Len(TabProduto.Fields("codg_ncm").Value) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"
            LIMPA_BODY
            txtProduto.SetFocus
         End If
      End If
   End If

   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

   If Trim(txtPedido.Text) = "" Then
      MsgBox "Falta numero pedido."
      Exit Sub
   End If

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

'panificadora
   'If INDR_PANIFICADORA = True And Trim(CODIGO_BARRAS) <> "" And INDR_PROD_BALANCA = True Then
   '   VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS, 8, 5) / 100
   '   txtValor_Unitario.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

   '   QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value)

   '   PESO_ITEM_N = QTDE_N
   '   txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
   '   txtQtde.Text = txtPesoItem.Text
   'End If

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * FROM PEDIDOITEM "
   SQL = SQL & " where codg_prod = '" & Trim(txtProduto.Text) & "'"
   SQL = SQL & " and PRODUTO_ID = " & PRODUTO_ID_N
   SQL = SQL & " and pedido_ID = " & PEDIDO_ID_N
   SQL = SQL & " and seq_ID = " & Trim(txtSeq.Text)
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      txtValor_Unitario.Text = Format(TabPedidoItem!Valor_Item, strFormatacao2Digitos)
      txtDesconto.Text = Format(TabPedidoItem!PERC_desc, strFormatacao2Digitos)
      txtQtde.Text = Format(TabPedidoItem!QTD_PEDIDA, strFormatacao3Digitos)
      QTDE_PEDIDO = TabPedidoItem!QTD_PEDIDA
      VALOR_ITEM_N = TabPedidoItem!Valor_Item
      VALOR_DIFERENCA_N = TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA

      SQL = "UPDATE Produto SET qtde_retido = qtde_retido - " & tpMOEDA(QTDE_PEDIDO)
      SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N & " and codg_produto = '" & txtProduto.Text & "' and qtde_retido >= " & tpMOEDA(QTDE_PEDIDO)
      CONECTA_RETAGUARDA.Execute SQL

      txtQtdeDisp.Text = "" & Format(TabProduto!QTDE, strFormatacao3Digitos)
      If Not IsNull(TabProduto!QTDE) Then _
         QTDE_ESTOQUE = TabProduto!QTDE
      txtSeq.Text = "" & TabPedidoItem.Fields("seq_id").Value
   End If

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   If Len(Trim(CODIGO_BARRAS)) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtValor_Unitario.Text) <> "" Then
            If IsNumeric(txtValor_Unitario.Text) Then
'=======
'foi para le_produto
               'txtQtde.Text = QTDE_N
               'If INDR_PANIFICADORA = True And INDR_PROD_BALANCA = True Then
               '   txtQtde.Text = Format(QTDE_N, strFormatacao3Digitos)
               '   Else: txtQtde.Text = Format(QTDE_N / 1000, strFormatacao3Digitos)
               'End If

'================
               Call txtDesconto_LostFocus

               CODIGO_BARRAS = ""
               txtProduto.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If
   CODIGO_BARRAS = ""

   txtValor_Unitario.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PRODUTO"
End Sub

Sub CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "SELECT PEDIDO.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.Codg_Prod, "
      SQL = SQL & " PEDIDOITEM.QTD_PEDIDA, Produto.Descricao, Produto.Qtde , Produto.QTDE_RETIDO "
      SQL = SQL & " FROM PEDIDO "
      SQL = SQL & " INNER JOIN PEDIDOITEM "
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
      SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and tipo_registro in ('S','R','D') "
      SQL = SQL & " and PEDIDO.status < 3 "

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         QTDE_RETIDO = TabConsulta.Fields("qtde_retido").Value

         QTDE_ESTOQUE = TabConsulta.Fields("qtde").Value - _
                        QTDE_RETIDO - _
                        QTDE_PEDIDO
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      Else
         MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)
         txtProduto.SetFocus
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO"
End Sub

Private Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   Dim TabPedido As New ADODB.Recordset

   If NUMR_REQ_N > 0 Then
      SINAL_INDICADOR_N = 1

      If INDR_FORM_ABERTO = True Then
         Unload frmCADRECEBVENDA
         INDR_FORM_ABERTO = False
      End If

'===================================
'===================================
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select contabiliza from TIPOVENDA "
         SQL = SQL & " where tipovenda_id = " & cmbFaturaAux.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then

            If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
               If TabTemp.Fields("contabiliza").Value = True Then
                  If TabTemp.State = 1 Then _
                     TabTemp.Close

'frmRECEBECAIXA.Show 1
frmCADRECEBVENDA.Show 1

                  'Exit Sub
                  Else
                     SQL = "update PEDIDO set "
                     SQL = SQL & "status = 6 " 'não contabiliza
                     SQL = SQL & " where numr_req = " & NUMR_REQ_N
                     SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
'===================================
'===================================
'===================================
   
      If INDR_CONTROLA_ESTOQUE = False Then _
         Exit Sub

      If TabPedido.State = 1 Then _
         TabPedido.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
         If TabPedido!Status = 5 Then
            CNPJCPF_A = Trim(TabPedido!CGCCPF)
'=============================================================================
            If USA_ECF = True Then
               Msg = "Confirma Faturamento ?"
               PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
'====================
frmDISPLAYEMISSOR.IMPRIME_CUPOM_FISCAL
'====================
                  If NUMR_REQ_N > 0 Then
                     SQL = "update PEDIDO set "
                     SQL = SQL & "status = 7 " 'CUPOM FISCAL
                     'SQL = SQL & ", numr_cupom =  " & NUMEROCUPOM
                     SQL = SQL & " where numr_req = " & NUMR_REQ_N
                     SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                     CONECTA_RETAGUARDA.Execute SQL
                  End If
               End If
            End If
         End If
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close

'====================
frmDISPLAYEMISSOR.CONTROLE_ESTOQUE_2  'CONTROLE
'====================

   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

   If USA_NFe = True Then
      SQL = "select status from PEDIDO "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         If Not IsNull(TabPedido.Fields(0).Value) Then
            If TabPedido.Fields(0).Value > 2 And TabPedido.Fields(0).Value < 9 Then
               Msg = "Deseja Gerar Nota Fiscal Eletrônica ?"
               PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then _
                  GERA_NOTA
            End If
         End If
      End If
   End If

   If TabPedido.State = 1 Then _
      TabPedido.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      If Not IsNull(TabCABECA!Status) Then
         If TabCABECA!Status <> "" Then
            If TabCABECA!Status = 5 Or TabCABECA!Status = 7 Then
               CRITERIO = NUMR_REQ_N
               If TabCABECA.State = 1 Then _
                  TabCABECA.Close
               TIPO_NFe_GERAR = "S"
               frmNOTAGERA.Show 1
            End If
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
End Sub

Private Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   Dim TOT_ITENS_PEDIDO_N As Long

   TOT_ITENS_PEDIDO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0

   txtVlrUnit.Text = Format(VALOR_ITEM_N, "##,##0.00")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'BUSCA VALOR TOTAL VENDA
   SQL = "select sum(valor_item*qtd_pedida) "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.empresa_id  = " & EMPRESA_ID_N
   SQL = SQL & " and PEDIDOITEM.numr_req = " & txtPedido.Text

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_ITEM_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'CONTA QTDE ITENS NO PEDIDO
   SQL = "select count(pedidoitem.produto_id) "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.numr_req = " & txtPedido.Text
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      TOT_ITENS_PEDIDO_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'VALOR DESCONTO ITEM
   SQL = "select sum(PEDIDOITEM.valor_desconto) "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.empresa_id  = " & EMPRESA_ID_N
   SQL = SQL & " and PEDIDOITEM.numr_req = " & txtPedido.Text
   
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'VALOR DESCONTO NA CABEÇA
   SQL = "select valor_desconto from PEDIDO "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and numr_req = " & txtPedido.Text
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

   SQL = "select sum(peso_item) from PEDIDOITEM "
   SQL = SQL & " where numr_req = " & txtPedido.Text
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

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close

      Exit Sub
   End If

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   QTDE_N = 0
   CRITERIO = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close
'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   SQL = "select count(produto_id) from PRODUTO "
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO = Trim(CODIGO_BARRAS)

            frmPEDIDOBARRAS.Show 1

            If Trim(CRITERIO) <> "" Then
               txtProduto.Text = Trim(CRITERIO)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select * from PRODUTO "
               SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and situacao <> 'C' "
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then _
                  MOSTRA_DADOS_PRODUTO

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               CRITERIO = ""
               Exit Sub
            End If
         End If
      End If
   End If

CRITERIO = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQtde.Text = 1
Call txtDesconto_LostFocus
txtProduto.SetFocus
      Exit Sub
   End If

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS) = 13 Then
      '2 = produtos "in store" (sempre será 2)     1
      'C = código do produto (4,5 ou 6 dígitos)    2 a 8
      'T = total a pagar (sempre 6 dígitos)        9 a 13
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      If INDR_PANIFICADORA = True Then
         txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 4))

         If TabProduto.State = 1 Then _
            TabProduto.Close
      
         SQL = "select * from PRODUTO "
         SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then
               INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value
               QTDE_N = 1

'removido do mostra_dados
'panificadora
   If INDR_PANIFICADORA = True And Trim(CODIGO_BARRAS) <> "" And INDR_PROD_BALANCA = True Then
      VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS, 8, 5) / 100
      txtValor_Unitario.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

      QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value)

      PESO_ITEM_N = QTDE_N
      txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
      txtQtde.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
   End If

               MOSTRA_DADOS_PRODUTO

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               Exit Sub
               Else: MsgBox "Verificar cadastro produto."
            End If
            Else: MsgBox "Verificar cadastro produto."
         End If
         Else
            txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 6))  'PORTO SEGURO

            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select * from PRODUTO "
            SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then
               QTDE_N = 0 & Int(Mid(CODIGO_BARRAS, 8, 5))   'gramas
               PESO_ITEM_N = QTDE_N
               txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)

               MOSTRA_DADOS_PRODUTO

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               Exit Sub
            End If
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(CODIGO_BARRAS) = 12 Then
      'lendo codigo barras ultralav
      '100004360813
      '100002361113
      '1-1 = masculino ou feminino
      '2-7 = código do produto
      '8-9 = numeração tamanho produto
      '10-11 = mes
      '12-13 = ano

      txtProduto.Text = "" & Mid(CODIGO_BARRAS, 1, 6)
      SqL2 = "" & Mid(CODIGO_BARRAS, 7, 2)

      SQL = "select * from PRODUTO "
      SQL = SQL & " where referencia = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and RIGHT(descricao,2) = '" & Trim(SqL2) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         MOSTRA_DADOS_PRODUTO

         If TabProduto.State = 1 Then _
            TabProduto.Close

         txtQtde.Text = 1
         'txtQTDE.SetFocus
         'Call txtQtde_LostFocus

         'txtDesconto.SetFocus
         Call txtDesconto_LostFocus

         txtProduto.SetFocus

         Exit Sub
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MsgBox "Produto não cadastrado."
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_PRODUTO"
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
'======================================================
'======================================================
'======================================================
Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   'txtValorDig.SelStart = 0
   'txtValorDig.SelLength = Len(txtValorDig)

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

      If QTDE_N > 0 And VALOR_ITEM_N > 0 And VALOR_DESCONTO_N >= 0 And SEQ_ID_N > 0 Then

MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N) - VALOR_DESCONTO_N), strFormatacao2Digitos)  'total item
'lucro MSFlexGrid1.TextMatrix(LastRow, 9) = Format(((VALOR_ITEM_N - PRECO_CUSTO_N) * QTDE_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

         If INDR_ESTQ_NEGATIVO = False Then
            CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO

            If QTDE_ESTOQUE < 0 Then
               Beep
               MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
               txtQtde.SetFocus
               Exit Sub
            End If
         End If

         If TabGridVaca.State = 1 Then _
            TabGridVaca.Close

         SQL = "SELECT PEDIDO.NUMR_REQ, PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, "
         SQL = SQL & " PEDIDOITEM.CODG_PROD, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.Valor_Item, "
         SQL = SQL & " PEDIDOITEM.PERC_DESC, PEDIDOITEM.Valor_Desconto, PEDIDOITEM.Status, PEDIDOITEM.PRECO_CUSTO "
         SQL = SQL & " FROM PEDIDO "
         SQL = SQL & " INNER JOIN PEDIDOITEM "
         SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "

         SQL = SQL & " where PEDIDO.numr_req = " & txtPedido.Text
         SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N

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

            BAIXA_RETIDO QTDE_RETIDO_ESTORNO

            SQL = "update PRODUTO set "
            SQL = SQL & " qtde_retido = qtde_retido + " & tpMOEDA(QTDE_N)
            SQL = SQL & " where codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
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

         SQL = "select * from PEDIDO "
         SQL = SQL & " where numr_req = " & txtPedido.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            Msg = "Deseja realmente clonar o pedido de venda : " & txtPedido.Text & " ?"
            PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               GERA_NUMR_REQ

               SQL = "INSERT INTO PEDIDO "
                  SQL = SQL & "(PEDIDO_ID,Empresa_id, numr_req, CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, Status, "
                  SQL = SQL & " Tipo_Registro,Codg_USU, TIPOVENDA_ID, CLIENTE_ID, Valor_ToTal,"
                  SQL = SQL & " valor_desconto,perc_desc) "
                  SQL = SQL & " VALUES ("
                     SQL = SQL & NUMR_REQ_N
                     SQL = SQL & "," & TabTemp.Fields("empresa_id").Value 'EMPRESA_ID_N
                     SQL = SQL & "," & NUMR_REQ_N
                     SQL = SQL & ",'" & Trim(TabTemp.Fields("cgccpf").Value) & "'"
                     SQL = SQL & "," & TabTemp.Fields("vendedor_id").Value
                     SQL = SQL & ",'" & DMA(Date) & "'"
                     SQL = SQL & ",'" & Trim(TabTemp.Fields("nome_cliente").Value) & "'"
                     SQL = SQL & "," & 2
                     SQL = SQL & ",'R'"
                     SQL = SQL & "," & TabTemp.Fields("codg_usu").Value
                     SQL = SQL & "," & TabTemp.Fields("tipovenda_id").Value
                     SQL = SQL & "," & TabTemp.Fields("cliente_id").Value
                     SQL = SQL & "," & tpMOEDA(TabTemp.Fields("valor_total").Value)
                     SQL = SQL & "," & tpMOEDA(0)  'vai zerar e tratar somente na tela de desconto
                     SQL = SQL & "," & tpMOEDA(0)
               SQL = SQL & ")"
               CONECTA_RETAGUARDA.Execute SQL

'set itens

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select * from PEDIDOitem "
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               While Not TabConsulta.EOF
                  SQL = "INSERT INTO PEDIDOITEM "
                  SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Numr_req, Codg_Prod, Qtd_Pedida,Valor_item, "
                  SQL = SQL & " PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG,PESO_ITEM) "
                  SQL = SQL & " VALUES ("

                     SQL = SQL & NUMR_REQ_N                                                          'PEDIDO_id
                     SQL = SQL & "," & TabConsulta.Fields("SEQ_ID").Value                                                        'SEQ_ID
                     SQL = SQL & "," & TabConsulta.Fields("PRODUTO_ID").Value
                     SQL = SQL & "," & NUMR_REQ_N                                                 'Numr_req
                     SQL = SQL & ",'" & TabConsulta.Fields("codg_prod").Value                                          'Codg_Prod
                     SQL = SQL & "'," & tpMOEDA(TabConsulta.Fields("QTD_PEDIDa").Value)                                           'Qtd_Pedida
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("VALOR_ITEM").Value)                                           'Valor_item
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("PERC_DESC").Value)                                        'PERC_DESC
                     SQL = SQL & "," & tpMOEDA((TabConsulta.Fields("VALOR_ITEM").Value * TabConsulta.Fields("QTD_PEDIDa").Value) * TabConsulta.Fields("PERC_DESC").Value / 100) 'valor_desconto
                     SQL = SQL & ", 'P'"                                                              'status
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("preco_custo").Value)                                    'PRECO_CUSTO
                     SQL = SQL & ",'PC'"                                                              'TIPO_REG
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTD_PEDIDa").Value)                                           'PESO_ITEM

                  SQL = SQL & ")"
                  CONECTA_RETAGUARDA.Execute SQL

                  TabConsulta.MoveNext
               Wend

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               MsgBox "Processo realizado com sucesso."
               SqL2 = NUMR_REQ_N
               LIMPA_TUDO
               NUMR_REQ_N = SqL2
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
