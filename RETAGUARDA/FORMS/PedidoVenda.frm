VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoVenda 
   Caption         =   "Pedido Venda"
   ClientHeight    =   7350
   ClientLeft      =   2085
   ClientTop       =   2475
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00400040&
   Icon            =   "PedidoVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleMode       =   0  'User
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtValorDig 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTotalPedido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   9000
      TabIndex        =   28
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox txtItens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   6960
      TabIndex        =   26
      Top             =   6480
      Width           =   615
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
      Height          =   700
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   10935
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   6100
         Picture         =   "PedidoVenda.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6615
         MaxLength       =   100
         TabIndex        =   22
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtLIMITE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   7920
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPAGAR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
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
         Left            =   10560
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbFatura 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   10560
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
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
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbVendedor 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Selecione um vendedor"
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   4080
         TabIndex        =   39
         Top             =   240
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
      Begin MSMask.MaskEdBox txtDtEmis 
         Height          =   360
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3315
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Crédito"
         Height          =   240
         Left            =   7200
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "À Pagar"
         Height          =   240
         Left            =   9000
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Fat.:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9960
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   495
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
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   10935
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPreçoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   19
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2865
         Picture         =   "PedidoVenda.frx":6614
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtVarejo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtAtacado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4320
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         MaxLength       =   29
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtValor_Unitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7800
         MaxLength       =   12
         TabIndex        =   4
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Vlr.Varejo ="
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Qtde ="
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Preço ="
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Produto:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   -240
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10965
      DesignHeight    =   7350
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4215
      Left            =   45
      TabIndex        =   30
      Top             =   2160
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7435
      _Version        =   393216
      GridLinesFixed  =   1
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
   Begin MSComctlLib.StatusBar barPedido 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   6975
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "PedidoVenda.frx":7016
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   14250
      _ExtentX        =   25135
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
               Picture         =   "PedidoVenda.frx":7468
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":8602
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":9691
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":A646
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":B751
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":C8A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":CCF9
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":EB70
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":10226
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PedidoVenda.frx":12208
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   9480
         TabIndex        =   35
         Top             =   360
         Width           =   1455
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
         TabIndex        =   36
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Label lblIDComanda 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1800
      TabIndex        =   41
      Top             =   6480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblComanda 
      AutoSize        =   -1  'True
      Caption         =   "Comanda: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7755
      TabIndex        =   29
      Top             =   6480
      Width           =   1020
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Itens = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   27
      Top             =   6495
      Width           =   990
   End
End
Attribute VB_Name = "frmPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
   Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
   Private Const MF_BYPOSITION = &H400&
   Private LastRow            As Long ' Ultima linha em que se editou
   Private LastCol            As Long ' ultima coluna em que se editou
   Private ControlVisible     As Boolean

   'Private CalculaIcmsG       As New MegasimCL.mCalculaIcms ' Yuri alterado em 01/05/2012

   Dim TabCOMANDA As New ADODB.Recordset

   Dim TIPO_CLIENTE_N         As Integer
   Dim VALOR_UNITARIO_N       As Double
   Dim INDR_PROD_BALANCA      As Boolean
   Dim COMANDA_ID_N           As Long
   Dim SITUACAO_COMANDA_A     As String
   Dim SEQ_PEDIDO_ID_N        As Long
   Dim SEQ_COMANDA_ID_N       As Long

   'TRIBUTAÇÃO
   Private cTritutacao As New cTributacao

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INICIALIZA_VENDA

   ABRE_PEDIDO

   OcultarControles

   Toolbar1.Buttons(2).Visible = True
   If LIMPA_PEDIDO = True Then _
      Toolbar1.Buttons(2).Visible = False

   REMOVE_MENU

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Dim QUAL_TECLA As String

   Select Case KeyCode
      Case vbKeyEscape
         QUAL_TECLA = "vbKeyEscape"
         If TRAZ_TIPO_USUARIO = 1 Then
            LIMPA_TUDO
            Unload Me
         End If
      Case vbKeyF1
         QUAL_TECLA = "vbKeyf1"
            COMANDA_CHAMA "GRAVAR"
         txtProduto.SetFocus
      Case vbKeyF3
         Msg = "Tem certeza que deseja excluir todos itens de comanda ?"
         PERGUNTA Msg, vbYesNo + 32, "Cancelar", "DEMO.HLP", 1000
         If RESPOSTA = vbNo Then _
            Exit Sub
         QUAL_TECLA = "vbKeyF3"
         If TRAZ_TIPO_USUARIO <> 1 Then
            SEQ_PEDIDO_ID_N = 0
            COMANDA_ID_N = 0
            CARTAOBARRA_ID_N = 0

            frmCOMANDA.Show 1
            If CARTAOBARRA_ID_N > 0 Then
               If TabCOMANDA.State = 1 Then _
                  TabCOMANDA.Close
               SQL = "SELECT comanda_id FROM COMANDA WITH (NOLOCK) "
               SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N
               TabCOMANDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCOMANDA.EOF Then
                  Msg = "Todos os itens da comanda nº: " & CARTAOBARRA_ID_N & " serão excluidos, CONFIRMA ?"
                  PERGUNTA Msg, vbYesNo + 32, "Cancelar", "DEMO.HLP", 1000
                  If RESPOSTA = vbNo Then _
                     Exit Sub

                  COMANDA_ID_N = 0 & TabCOMANDA.Fields("COMANDA_ID").Value

                  SQL = "spCOMANDAITEM " & 3 & "," & COMANDA_ID_N & "," & 0 & "," & PRODUTO_ID_N & ",'" & tpMOEDA(QTDE_N) & "','" & tpMOEDA(VALOR_ITEM_N) & "','" & "" & "'," & ATENDENTE_ID_N
                  CONECTA_RETAGUARDA.Execute "EXEC " & SQL

                  SQL = "spCOMANDA " & 3 & "," & COMANDA_ID_N & "," & CARTAOBARRA_ID_N & "," & USUARIO_ID_N & ",'" & Now & "','" & "" & "'"
                  CONECTA_RETAGUARDA.Execute "EXEC " & SQL
               End If
               If TabCOMANDA.State = 1 Then _
                  TabCOMANDA.Close
            End If

            'ATUALIZAR TABELA PEDIDOCOMANDA PELO CAMPO cartaobarra_id, TABELA TEMPORÁRIA
            'EXCLUINDO
            spPedidoComanda 3, 0, 0, ""

            txtProduto.Enabled = True
            FraSeq.Enabled = True
            txtProduto.SetFocus
         End If
      Case vbKeyF8
         QUAL_TECLA = "vbKeyF8"
         Call cmdConsCli_Click
      Case vbKeyF10
         VAI_VENDA
      Case vbKeyF12
         QUAL_TECLA = "vbKeyF12"
         If TRATA_SAIDA_TELA = True Then
            LIMPA_TUDO

            ABRE_PEDIDO

            FraSeq.Enabled = True
            txtProduto.Enabled = True
            txtProduto.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & QUAL_TECLA, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If PEDIDO_ID_N > 0 Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select status from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            If TabTemp.Fields(0).Value <= 2 Then
               Msg = "Pedido pendente, deseja realmente cancelar essa venda ?"
               PERGUNTA Msg, vbYesNo + 32, "Cancelar", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
   
                  SQL = "delete from ITEMLANCAMENTO where numr_doc = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
                  SQL = "delete from LANCAMENTO where numr_doc = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
                  SQL = "delete from pedidotemp where pedido_id = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

                  SQL = "update pedido set "
                  SQL = SQL & " status = 9 "
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

'ATUALIZAR TABELA PEDIDOCOMANDA PELO CAMPO cartaobarra_id, TABELA TEMPORÁRIA
'EXCLUINDO
spPedidoComanda 3, 0, 0, ""

                  Else: Cancel = 1
               End If
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "clonar"
         CLONA_PEDIDO_VENDA
      Case "receber"
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
            frmDISPLAYEMISSOR.Show 1

            If Trim(txtPedido.Text) <> "" Then
               If IsNumeric(txtPedido.Text) Then
                  txtPedido.Enabled = True
                     PEDIDO_ID_N = txtPedido.Text
                  txtPedido.Enabled = False

                  VALIDA_PEDIDO txtPedido.Text
                  SETA_GRID
               End If
            End If

            INICIALIZA_VENDA
         End If
      Case "gravar"
         VAI_VENDA
      Case "consultar"
         If TRATA_SAIDA_TELA = True Then
            CRITERIO_A = ""
            CNPJCPF_A = ""
            frmPedidoConsulta.Show 1
            If PEDIDO_ID_N > 0 Then
               Dim NUMR_PEDIDO_N As Long

               NUMR_PEDIDO_N = 0 & PEDIDO_ID_N

               LIMPA_TUDO

               txtPedido.Text = "" & NUMR_PEDIDO_N
               CRITERIO_A = ""
               NUMR_PEDIDO_N = 0

               ABRE_PEDIDO
            End If
            FraSeq.Enabled = True
            txtProduto.Enabled = True
            txtProduto.SetFocus
         End If
      Case "print"
         GERA_IMPRESSAO
      Case "limpar"
         If TRATA_SAIDA_TELA = True Then
            LIMPA_TUDO

            ABRE_PEDIDO

            FraSeq.Enabled = True
            txtProduto.Enabled = True
            txtProduto.SetFocus
         End If
      Case "voltar"
         Unload Me
      Case "produto"
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
            frmCADASTROPRODUTO.Show 1
            Else: CHAMA_PRODUTO_SIMPLIFICADO
         End If
      Case "CadCliente"
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1
   End Select

   If MULT_EMPRESA_B = True Then _
      If USUARIO_ID_N <> 144 Then _
         Toolbar1.Buttons(9).Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.Text = CNPJCPF_A
      Call txtCNPJCPF_LostFocus
      FraSeq.Enabled = True
      'txtCNPJCPF.PromptInclude = True

      txtProduto.Enabled = True

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
'On Error GoTo ERRO_TRATA

   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Private Sub cmbVENDEDOR_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE_PEDIDO "F1-Incluir Comanda", "Selecione Vendedor e tecle <ENTER>", "", "", ""

   If MULT_EMPRESA_B = True Then
      If TIPO_USUARIO < 4 Then
         If Trim(cmbVendedor.Text) = "" Then _
            MOSTRA_VENDEDORES

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus
      End If
   End If
   cmbVendedor.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVendedor_GotFocus"
End Sub

Private Sub cmbVENDEDOR_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVendedor.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVendedor_Click"
End Sub

Private Sub cmbvendedor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      FraSeq.Enabled = True
      txtProduto.Enabled = True

      txtProduto.SetFocus
      Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVendedor_KeyPress"
End Sub

Private Sub cmbVendedor_LostFocus()
   cmbVendedor.BackColor = &HFFFFFF
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
'On Error GoTo ERRO_TRATA

   KeyAscii = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAtacado_KeyPress"
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

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   If MULT_EMPRESA_B = True Then
      
         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus
      
      Else: cmbVendedor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_GotFocus"
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
'==================cnpjcpf
Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = "" & Trim(txtCNPJCPF.Text)
      txtCNPJCPF.Mask = "##############"
      txtCNPJCPF.Mask = ""
   txtCNPJCPF.Text = CNPJCPF_A

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.SelLength = Len(txtCNPJCPF)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         FraSeq.Enabled = True
         txtProduto.Enabled = True

         txtProduto.SetFocus
         SQL3 = txtNome.Text
         txtNome.Text = Trim(InputBox("Informe Nome do cliente", "Emissão de Cupom Fiscal", SQL3))
      Case vbKeyF7
         CNPJCPF_A = ""
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
      CNPJCPF_A = Trim(txtCNPJCPF.Text)
      txtCNPJCPF.Mask = ""
      If Trim(CNPJCPF_A) = "99999999999" Then
         txtNome.Enabled = True
         FraSeq.Enabled = True
         Else
            FraSeq.Enabled = True
            txtProduto.Enabled = True
      End If
      txtNome.Enabled = True
      txtNome.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0
   NOME_CLIENTE_A = ""

   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = "" & Trim(txtCNPJCPF.Text)
   If Trim(CNPJCPF_A) = "" Then
      txtCNPJCPF.Text = "99999999999"
      CNPJCPF_A = "99999999999"
      txtNome.Enabled = True
      If Trim(txtNome.Text) = "" Then _
         txtNome.Text = "" & TRAZ_NOME_PESSOA(0, Trim(CNPJCPF_A))
   End If
   If TRATA_PESSOA(CNPJCPF_A) = False Then
      txtCNPJCPF.Text = "99999999999"
      CNPJCPF_A = "99999999999"
      txtNome.Enabled = True
      txtNome.Text = "" & TRAZ_NOME_PESSOA(0, Trim(CNPJCPF_A))
      txtCNPJCPF.SetFocus
      Exit Sub
      Else
         txtNome.Text = "" & TRAZ_NOME_PESSOA(0, Trim(CNPJCPF_A))
         CHECA_CLIENTE CNPJCPF_A, Trim(txtNome.Text)
   End If

   txtPAGAR.Text = Format(VALOR_PENDENTE_N, strFormatacao2Digitos)
   txtPAGAR.Refresh

   If Trim(CNPJCPF_A) <> "" Then
      If Not IsNull(CNPJCPF_A) Then
          If Len(CNPJCPF_A) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(CNPJCPF_A) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.Text = CNPJCPF_A
   End If

   txtCNPJCPF.BackColor = &HFFFFFF
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNome.Text) <> "" Then
      txtNome.SelStart = 0
      txtNome.SelLength = Len(txtNome)
   End If
   txtNome.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
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

Private Sub txtTotalPedido_GotFocus()
'On Error GoTo ERRO_TRATA

      FraSeq.Enabled = True
      txtProduto.Enabled = True

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTotalPedido_GotFocus"
End Sub

Private Sub txtVarejo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtVarejo_KeyPress"
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

Private Sub txtNome_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNome.Text) <> "" Then _
      txtNome.Text = UCase(txtNome.Text)
   'txtNome.Enabled = False
   txtNome.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_LostFocus"
End Sub
Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      UCase (txtProduto.Text)
            FraSeq.Enabled = True
      txtProduto.Enabled = True

      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDescricao.Enabled = False

   MOSTRA_RODAPE_PEDIDO "F1-Incluir Comanda", "F3-Limpar Comanda", "F7-Consulta Produtos", "F8-Consulta Cliente", "F10-Gravar"

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
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
         txtProduto.Enabled = True
         txtProduto.SetFocus
      Case vbKeyF7
         CONSULTA_PRODUTO
         FraSeq.Enabled = True
         txtProduto.Enabled = True
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
      'If txtPedido.Text = "" Or Trim(txtProduto.Text) = "" Then _
      If Trim(txtProduto.Text) = "" Then _
         Exit Sub

      KeyAscii = 0

      PROCESSA_DADOS_PRODUTOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE_PEDIDO "F1-Incluir Comanda", "Informe a quantidade", "F10-Gravar", "", ""
   
   If Trim(txtProduto.Text) = Empty Then
   '   MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
   '   txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   QTDE_N = 0 & txtQTDE.Text
   If QTDE_N <= 0 Then _
      txtQTDE.Text = 1

   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE.Text)
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
         txtProduto.SetFocus
         Exit Sub
      End If
      QTDE_N = 0 & txtQTDE.Text
      If QTDE_N < 0 Then _
         txtQTDE.Text = 1

      Call PROCESSA_ITEM

      FraSeq.Enabled = True
      txtProduto.Enabled = True
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

   If Len(Trim(txtQTDE.Text)) >= 10 Then

      FraSeq.Enabled = True
      txtProduto.Enabled = True

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

   MOSTRA_RODAPE_PEDIDO "F1-Incluir Comanda", "Tecle <ENTER> para gerar nova Pedido ou informe uma já existente", "", "", ""

   If MULT_EMPRESA_B = True Then

      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus

      Else: cmbVendedor.SetFocus
   End If

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

Private Sub TXTVALOR_UNITARIO_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE_PEDIDO "F1-Incluir Comanda", "Informe Valor Unitário", "", "", ""
   
   txtValor_Unitario.SelStart = 0
   txtValor_Unitario.SelLength = Len(txtValor_Unitario.Text)
   txtValor_Unitario.BackColor = &HC0FFFF
   txtQTDE.SetFocus

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
         If VALOR_UNITARIO_N < PR_ATACADO_N Then
            MsgBox "Produto Tipo Promoção Impossível dar desconto."
            txtValor_Unitario.Text = 0
            txtValor_Unitario.SetFocus
            Else: txtQTDE.SetFocus
         End If
         Else
            If VALOR_UNITARIO_N <> VLR_ANTERIOR_N Then
                If VALOR_UNITARIO_N < PR_ATACADO_N Then
                   VALOR_DESCONTO_N = Format(PR_ATACADO_N - VALOR_UNITARIO_N, strFormatacao2Digitos)
                   PERC_DESCONTO_N = ((VALOR_DESCONTO_N * 100) / PR_ATACADO_N)
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
      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus
      Exit Sub
      Else
         VALOR_ITEM_N = txtValor_Unitario.Text
         txtValor_Unitario.Text = Format(VALOR_UNITARIO_N, strFormatacao2Digitos)
         If VALOR_ITEM_N <= 0 Then
            MsgBox "Valor Unitário Inválido !!!"

            FraSeq.Enabled = True
            txtProduto.Enabled = True

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
      Msg = "Valor informado menor que preço de atacado, não permitido !!!, informar senha superior?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
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
                     Exit Sub
                  End If
                  If TabTemp.Fields("tipo").Value >= 4 And TabTemp.Fields("tipo").Value <= 5 Then
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

   txtValor_Unitario.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_LostFocus"
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

   FraSeq.Enabled = True

      txtProduto.Enabled = True

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
      If LastCol > 3 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      Dim TabDig              As New ADODB.Recordset
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
      SEQ_PEDIDO_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)
      CODG_PRODUTO_A = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 0))
      PRODUTO_ID_N = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 12))

      If QTDE_N > 0 And VALOR_ITEM_N > 0 And VALOR_DESCONTO_N >= 0 And SEQ_PEDIDO_ID_N > 0 Then

         'MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * Qtde_N) - VALOR_DESCONTO_N), strFormatacao2Digitos)  'total item
         MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N)), strFormatacao2Digitos)  'total item
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

'===================
   'checa se o funcionário pode comprar produtos de produção conforme a cota diária estabelecida
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "99999999999" And Valor_Compra_Dia_Permitida > 0 Then
      If CHECA_FUNCIONARIO(txtCNPJCPF.Text) = True Then
         If CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO(ESTABELECIMENTO_ID_N, Date, Trim(txtCNPJCPF.Text), (VALOR_ITEM_N * QTDE_N)) = True Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Cota diária de compra de produtos de produção ultrapassada, não permitido."
            Exit Sub
         End If
      End If
   End If
'===================

         If TabDig.State = 1 Then _
            TabDig.Close

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

         SQL = SQL & " where PEDIDO.PEDIDO_ID = " & txtPedido.Text
         SQL = SQL & " and seq_id = " & SEQ_PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and pedidoitem.status <> 'C' "

         TabDig.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

         If Not TabDig.EOF Then
            SQL = "update PEDIDOITEM set "
            SQL = SQL & " QTD_PEDIDA = " & tpMOEDA(QTDE_N)
            SQL = SQL & ",Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",Valor_Desconto = " & tpMOEDA(VALOR_DESCONTO_N)
            SQL = SQL & ",peso_item = " & tpMOEDA(QTDE_N)

            SQL = SQL & " where pedido_id = " & TabDig.Fields("pedido_id").Value
            SQL = SQL & " and seq_id = " & SEQ_PEDIDO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            QTDE_RETIDO_ESTORNO = 0

            SETA_GRID
         End If
         If TabDig.State = 1 Then _
            TabDig.Close
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
      LIMPA_BODY
      txtProduto.SetFocus
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
'============================subrotinas
Sub INICIALIZA_VENDA()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & " - " & Me.Name

   UF_CLIENTE_A = ""  'Variavel para tratamento Fiscal do item
   UF_EMPRESA_A = "" 'Variavel para tratamento Fiscal do item
   CCE_CLIENTE_A = ""  'Variavel para tratamento Fiscal do item
   TIPO_CLIENTE_N = -1 'Variavel para tratamento fiscal do item
   CNPJCPF_A = ""

   txtDtEmis = Format(Date, "dd/mm/yyyy")

   PEGA_DADOS_EMPRESA
   MOSTRA_VENDEDORES

   Toolbar1.Buttons(8).Visible = True
   If TIPO_USUARIO < 4 Then _
      Toolbar1.Buttons(8).Visible = False

   If MULT_EMPRESA_B = True Then
      If USUARIO_ID_N <> 144 Then _
         Toolbar1.Buttons(9).Visible = False
      txtAtacado.Visible = False
      Label13.Caption = "Preço"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_VENDA"
End Sub

Sub ABRE_PEDIDO()
'On Error GoTo ERRO_TRATA

   txtPedido.Enabled = False
   TIPOVENDA_ID_N = 9999

   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = Trim(txtCNPJCPF.Text)
   If CNPJCPF_A = "" Then
      CNPJCPF_A = "99999999999"
      txtCNPJCPF.Text = "" & CNPJCPF_A
      txtNome.Text = "" & TRAZ_NOME_PESSOA(0, Trim(CNPJCPF_A))
   End If
   If Trim(txtPedido.Text) <> "" Then
      If IsNumeric(txtPedido.Text) Then
         txtPedido.Enabled = True
            PEDIDO_ID_N = txtPedido.Text
         txtPedido.Enabled = False

         VALIDA_PEDIDO PEDIDO_ID_N
      End If
   End If

   SETA_GRID

   Dim TabPedido As New ADODB.Recordset

   lblIDComanda.Caption = ""
   lblIDComanda.Visible = False
   lblComanda.Visible = False

   If TabPedido.State = 1 Then _
      TabPedido.Close
'AQUI LENDO QUANTAS COMANDAS TEM ABERTAS PARA ESSE PEDIDO
   SQL = "select cartaobarra_id from PEDIDOCOMANDA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedido.EOF
      If Trim(lblIDComanda.Caption) = "" Then
         lblIDComanda.Caption = TabPedido.Fields(0).Value
         Else: lblIDComanda.Caption = lblIDComanda.Caption & "," & TabPedido.Fields(0).Value
      End If
      TabPedido.MoveNext
   Wend
   If TabPedido.State = 1 Then _
      TabPedido.Close

   If Trim(lblIDComanda.Caption) <> "" Then
      lblIDComanda.Visible = True
      lblComanda.Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ABRE_PEDIDO"
End Sub

Private Sub EXCLUIR_ITEM(CODG_PRODUTO_A As String, PEDIDO_ID_N As Long, SEQ_PEDIDO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Trim(PEDIDO_ID_N) > 0 And Trim(SEQ_PEDIDO_ID_N) > 0 And Trim(CODG_PRODUTO_A) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select PEDIDOITEM.*, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
      SQL = SQL & " PRODUTO.PRECO_VENDA, PRODUTO.PRECO_CUSTO, Produto.Situacao_Tributaria"
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where PEDIDOITEM.PEDIDO_ID = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_PEDIDO_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
      'SQL = SQL & " and pedidoitem.status <> 'C' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Deseja cancelar esse item ?  " & Trim(TabTemp.Fields("descricao").Value)
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            If TabProduto.State = 1 Then _
               TabProduto.Close

            VALOR_TOTAL_N = Format(VALOR_TOTAL_N - (TabTemp!Valor_Item * TabTemp!QTD_PEDIDA), "##,##0.00")

            'BAIXA_RETIDO  TabTemp!QTD_PEDIDA, TabTemp.Fields("produto_id").Value

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "Delete from PEDIDOITEM "
            SQL = SQL & " Where pedido_id = " & PEDIDO_ID_N
            SQL = SQL & " and seq_id = " & SEQ_PEDIDO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            'ATUALIZAR TABELA PEDIDOCOMANDA PELO CAMPO cartaobarra_id, TABELA TEMPORÁRIA
            'EXCLUINDO
            spPedidoComanda 3, 0, SEQ_PEDIDO_ID_N, ""

            LIMPA_BODY
            txtTotalPedido.Text = Format(VALOR_TOTAL_N, "##,##0.00")
            txtTotalPedido.Text = Format(VALOR_TOTAL_N, "currency")
   
            GRAVA_PEDIDO_CABECA "P", 1
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
      txtProduto.Enabled = True

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""
   txtQTDE.Text = ""

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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 0
   CARTAOBARRA_ID_N = 0
   lblIDComanda.Caption = ""
   lblIDComanda.Visible = False
   txtCNPJCPF.Enabled = True
   cmdConsCli.Enabled = True
   INDR_VENDA = False
   cmbVendedor.Enabled = False

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
   txtItens.Text = ""
   txtTotalPedido.Text = ""

   PRODUTO_ID_N = 0
   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   txtPedido.Text = ""
   txtDtEmis = Format(Date, "dd/mm/yyyy")
   txtNome.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCNPJCPF.Mask = "##############"
   
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

Private Sub MOSTRA_VENDEDORES()
'On Error GoTo ERRO_TRATA

   cmbVendedor.Clear
   cmbVendAux.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVENDEDOR.EOF
      cmbVendedor.AddItem Trim(TabVENDEDOR!DESCRICAO) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
      TabVENDEDOR.MoveNext
   Wend
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   QUALIFICA_VENDEDOR

   cmbVendedor.Enabled = False
   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
      cmbVendedor.Enabled = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDORES"
End Sub

Private Sub GERA_VENDA()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = txtPedido.Text
   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = txtCNPJCPF.Text

   txtNome.Text = "" & TRAZ_NOME_PESSOA(0, Trim(CNPJCPF_A))

   If Trim(cmbVendAux.Text) = "" Then
      cmbVendAux.Text = 0

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select VENDEDOR.VENDEDOR_ID FROM VENDEDOR WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
      SQL = SQL & " ON VENDEDOR.PESSOA_ID = PESSOA.PESSOA_ID"
      SQL = SQL & " where descricao = 'BALCAO'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         cmbVendAux.Text = "" & TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

   'atualizando desconto na cabeça
   SQL = "UPDATE PEDIDO SET "
   SQL = SQL & " Valor_desconto = 0"
   SQL = SQL & " , Perc_desc = 0"
   SQL = SQL & " , cgccpf = '" & Trim(CNPJCPF_A) & "'"
   SQL = SQL & " , nome_cliente = '" & Trim(txtNome.Text) & "'"
   SQL = SQL & " , status = 2"
   SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
   SQL = SQL & " , vendedor_id = " & cmbVendAux.Text

   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N 'txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   If RECEBE_PEDIDO_VENDA = True Then
      FAZ_RECEBIMENTO
      Else

         LIMPA_TUDO
         PEDIDO_ID_N = 0
         txtPedido.Text = ""

   End If

   'rotina que trabalha com a comanda eletronica
   If Trim(lblIDComanda.Caption) <> "" Then _
      lblIDComanda.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_VENDA"
End Sub

Sub QUALIFICA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   Dim TabVend As New ADODB.Recordset
   Dim TabUser As New ADODB.Recordset

   cmbVendedor.Enabled = False
   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
      cmbVendedor.Enabled = True

   'buscar vendedor por pessoa_id
   If TabUser.State = 1 Then _
      TabUser.Close

   SQL = "select pessoa_id from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
   TabUser.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUser.EOF Then
      PESSOA_ID_N = 0 & TabUser.Fields("pessoa_id").Value

      If TabVend.State = 1 Then _
         TabVend.Close

      SQL = "select descricao, vendedor_id "

      SQL = SQL & " FROM PESSOA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN VENDEDOR WITH (NOLOCK)"
      SQL = SQL & " ON PESSOA.PESSOA_ID = VENDEDOR.PESSOA_ID"

      SQL = SQL & " where VENDEDOR.PESSOA_ID = " & PESSOA_ID_N

      TabVend.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVend.EOF Then
         VENDEDOR_ID_N = 0 & TabVend.Fields("vendedor_id").Value
         cmbVendedor.Text = "" & Trim(TabVend.Fields("descricao").Value)
         cmbVendAux.Text = "" & VENDEDOR_ID_N
      End If
      If TabVend.State = 1 Then _
         TabVend.Close
   End If
   If TabUser.State = 1 Then _
      TabUser.Close

   PESSOA_ID_N = 0

   If Trim(cmbVendedor.Text) = "" Then
      If TabUser.State = 1 Then _
         TabUser.Close

      SQL = "select logon,usuario_id,pessoa_id from USUARIO WITH (NOLOCK)"
      SQL = SQL & " where logon = 'BALCAO' "
      TabUser.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabUser.EOF Then
         PESSOA_ID_N = 0 & TabUser.Fields("pessoa_id").Value
   
         If TabVend.State = 1 Then _
            TabVend.Close
   
         SQL = "select descricao, vendedor_id "
   
         SQL = SQL & " FROM PESSOA WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN VENDEDOR WITH (NOLOCK)"
         SQL = SQL & " ON PESSOA.PESSOA_ID = VENDEDOR.PESSOA_ID"
   
         SQL = SQL & " where VENDEDOR.PESSOA_ID = " & PESSOA_ID_N
   
         TabVend.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVend.EOF Then
            VENDEDOR_ID_N = 0 & TabVend.Fields("vendedor_id").Value
            cmbVendedor.Text = "" & Trim(TabVend.Fields("descricao").Value)
            cmbVendAux.Text = "" & VENDEDOR_ID_N
         End If
         If TabVend.State = 1 Then _
            TabVend.Close
      End If
      If TabUser.State = 1 Then _
         TabUser.Close
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
      FraSeq.Enabled = True
      txtProduto.Enabled = True

      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Private Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N <= 0 Then
      MsgBox "Pedido inválido, verifique !!!"
      Exit Sub
   End If

   Dim TabPedido           As New ADODB.Recordset
   Dim INDR_PERGUNTA       As Boolean
   Dim SITUACAO_PEDIDO_N   As Integer

SITUACAO_PEDIDO_N = 0

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
      SQL = SQL & " where tipovenda_id = " & TIPOVENDA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("contabiliza").Value) Then  'VERIFICANDO SE O TIPO DE VENDA CONTABILIZA
            If TabTemp.Fields("contabiliza").Value = True Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

frmFatura.Show 1

               Else  'SE NÃO CONTABILIZA ATUALIZA TABELA PEDIDO E NÃO GERA FINANCEIRO
                  SQL = "update PEDIDO set "
                  SQL = SQL & "status = 6 " 'não contabiliza
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

'essa rotina tem que testar quando a condição da empresa for esta, não contabiliza
                  If Trim(lblIDComanda.Caption) <> "" Then _
                     lblIDComanda.Visible = True
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
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
         SITUACAO_PEDIDO_N = 0 & TabPedido.Fields("STATUS").Value

         If SITUACAO_PEDIDO_N = 5 Then
            CNPJCPF_A = Trim(TabPedido!CGCCPF)

            If USA_DOC_FISCAL = True Then 'SE UTILIZA DOCUMENTO FISCAL VAI ENTRAR AQUI PARA EMITIR NFe OU NFCe
               If USA_NFC_E = True Then  'se usa NFC-e ENTRA AQUI
                  RESPOSTA = ""
                  Msg = ""
                  If INDR_VENDA_CARTAO = True Then
                     RESPOSTA = vbYes
                     Else
                        Msg = "Deseja Gerar Cupom Eletrônico ?"
                        PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                  End If

                  If RESPOSTA = vbYes Then
                     frmDISPLAYEMISSOR.ROTINA_NFC
                     Else
                        txtCNPJCPF.PromptInclude = False
                        If Trim(txtCNPJCPF.Text) <> "99999999999" Then
                           If USA_NFe = True And INDR_CAIXA = False Then
                              Msg = "Deseja Gerar Nota Fiscal Eletrônica ?"
                              PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                              If RESPOSTA = vbYes Then
                                 If SITUACAO_PEDIDO_N = 5 Or SITUACAO_PEDIDO_N = 7 Then
                                    CRITERIO_A = PEDIDO_ID_N
                                    TIPO_NFe_GERAR = "R"
                                    frmNOTAGERA.Show 1
                                 End If
                              End If
                           End If
                        End If
                  End If
                  Else
                     If USA_NFe = True And INDR_CAIXA = False Then
                        Msg = "Deseja Gerar Nota Fiscal Eletrônica ?"
                        PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                        If RESPOSTA = vbYes Then
                           If SITUACAO_PEDIDO_N = 5 Or SITUACAO_PEDIDO_N = 7 Then
                              CRITERIO_A = PEDIDO_ID_N
                              TIPO_NFe_GERAR = "R"
                              frmNOTAGERA.Show 1
                           End If
                        End If
                     End If
               End If
            End If   'FIM  If USA_DOC_FISCAL = True Then
         End If      'FIM  If SITUACAO_PEDIDO_N = 5 Then
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close

'==== VERIFICAR SE ESTA COM STATUS RECEBIDO

      If TabPedido.State = 1 Then _
         TabPedido.Close
      SQL = "select STATUS from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then _
         SITUACAO_PEDIDO_N = 0 & TabPedido.Fields("STATUS").Value
      If TabPedido.State = 1 Then _
         TabPedido.Close

'===================================
      If SITUACAO_PEDIDO_N = 3 Or SITUACAO_PEDIDO_N = 5 Then
      'If SITUACAO_PEDIDO_N = 5 Then
         If INDR_CONTROLA_ESTOQUE = True Then
            '====================
               ATUALIZA_ESTOQUE 0, PEDIDO_ID_N
            '====================
         End If
         '=============================================================================
         Dim i                         As Integer
         Dim Contador_CartaoBarra(10)  As String
         Dim Contador_Comanda(10)      As String

         lblIDComanda.Caption = ""

         CONT_N = 0
         If TabPedido.State = 1 Then _
            TabPedido.Close
'AQUI LENDO QUANTAS COMANDAS TEM ABERTAS PARA ESSE PEDIDO
         SQL = "select cartaobarra_id from PEDIDOCOMANDA WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabPedido.EOF
            If Trim(lblIDComanda.Caption) = "" Then
               lblIDComanda.Caption = TabPedido.Fields(0).Value
               Else: lblIDComanda.Caption = lblIDComanda.Caption & "," & TabPedido.Fields(0).Value
            End If

            CONT_N = CONT_N + 1
            Contador_CartaoBarra(CONT_N) = "" & TabPedido.Fields("cartaobarra_id").Value
            Contador_Comanda(CONT_N) = "" & TRAZ_COMANDA_ID(TabPedido.Fields("cartaobarra_id").Value)

'MsgBox Contador_CartaoBarra(CONT_N) & "   " & Contador_Comanda(CONT_N)

            TabPedido.MoveNext
         Wend
         If TabPedido.State = 1 Then _
            TabPedido.Close

'DEPOIS DE AGRUPADO O NUMERO DAS COMANDAS RODA AS ROTINAS DE
'ATUALIZAÇÃO DA COMANDA E PEDIDOCOMANDA
         For i = 0 To CONT_N

            If Contador_CartaoBarra(i) <> "" Then
               CARTAOBARRA_ID_N = 0 & Contador_CartaoBarra(i)
               COMANDA_ID_N = 0 & Contador_Comanda(i)

               If IsNumeric(CARTAOBARRA_ID_N) Then
                  If CARTAOBARRA_ID_N > 0 Then
                     frmPedidoComanda.GRAVA_COMANDA_ITEM "EXCLUIR", 0, 0, COMANDA_ID_N
                     frmPedidoComanda.GRAVA_CABECA_COMANDA "EXCLUIR", COMANDA_ID_N
                  End If
               End If

            End If
         Next i

         'ATUALIZAR TABELA PEDIDOCOMANDA PELO CAMPO cartaobarra_id, TABELA TEMPORÁRIA
         'EXCLUINDO
         spPedidoComanda 3, 0, 0, ""
      End If

      SQL = "delete pedidoTEMP where pedido_id = " & PEDIDO_ID_N
      CONECTA_RETAGUARDA.Execute SQL
'===================================
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
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
   If MULT_EMPRESA_B = True Then _
      If USUARIO_ID_N <> 144 Then _
         Toolbar1.Buttons(9).Visible = False

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
         SQL = SQL & " where PEDIDO_ID = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            Msg = "Deseja realmente clonar o pedido de venda : " & txtPedido.Text & " ?"
            PERGUNTA Msg, vbYesNo + 32, "Clonar Pedido", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               GERA_PEDIDO_ID

               SQL = "INSERT INTO PEDIDO "
                  SQL = SQL & "("
                     SQL = SQL & "PEDIDO_ID,Empresa_id,CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, Status, "
                     SQL = SQL & " Tipo_Registro,usuario_id, CLIENTE_ID, Valor_ToTal,"
                     SQL = SQL & " valor_desconto,perc_desc,NUMERO_CAIXA_CPU,estabelecimento_id,cartaobarra_id"
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
                     SQL = SQL & "," & EMPRESA_ID_N
                     SQL = SQL & "," & CARTAOBARRA_ID_N
               SQL = SQL & ")"
               CONECTA_RETAGUARDA.Execute SQL

'set itens

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select * from PEDIDOitem WITH (NOLOCK)"
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
               SQL = SQL & " and pedidoitem.status <> 'C' "
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               While Not TabConsulta.EOF
                  SQL = "INSERT INTO PEDIDOITEM "
                  SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Qtd_Pedida,Valor_item, "
                  SQL = SQL & " PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG,PESO_ITEM,usu_atende) "
                  SQL = SQL & " VALUES ("

                     SQL = SQL & PEDIDO_ID_N                                              'PEDIDO_id
                     SQL = SQL & "," & TabConsulta.Fields("SEQ_ID").Value                 'SEQ_ID
                     SQL = SQL & "," & TabConsulta.Fields("PRODUTO_ID").Value
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTD_PEDIDa").Value)   'Qtd_Pedida
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("VALOR_ITEM").Value)    'Valor_item
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("PERC_DESC").Value)     'PERC_DESC
                     SQL = SQL & "," & tpMOEDA((TabConsulta.Fields("VALOR_ITEM").Value * _
                                       TabConsulta.Fields("QTD_PEDIDa").Value) * _
                                       TabConsulta.Fields("PERC_DESC").Value / 100)       'valor_desconto
                     SQL = SQL & ", 'P'"                                                  'status
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("preco_custo").Value)   'PRECO_CUSTO
                     SQL = SQL & ",'PC'"                                                  'TIPO_REG
                     SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTD_PEDIDa").Value)    'PESO_ITEM
                     SQL = SQL & "," & ATENDENTE_ID_N                                        'USU_ATENDE

                  SQL = SQL & ")"
                  
'MsgBox SQL
                  
                  CONECTA_RETAGUARDA.Execute SQL
ATENDENTE_ID_N = 0
                  TabConsulta.MoveNext
               Wend
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               MsgBox "Processo realizado com sucesso."
               SqL2 = PEDIDO_ID_N
               LIMPA_TUDO
               PEDIDO_ID_N = SqL2

ABRE_PEDIDO

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

Sub MOSTRA_RODAPE_PEDIDO(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
'On Error GoTo ERRO_TRATA

   If Trim(Msg1) <> "" Then
      barPedido.Panels.Clear
      barPedido.Panels.Add (1)
      barPedido.Panels(1).Text = Trim(Msg1)
      barPedido.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         barPedido.Panels.Add (2)
         barPedido.Panels(2).Text = Trim(Msg2)
         barPedido.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            barPedido.Panels.Add (3)
            barPedido.Panels(3).Text = Trim(Msg3)
            barPedido.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               barPedido.Panels.Add (4)
               barPedido.Panels(4).Text = Trim(Msg4)
               barPedido.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  barPedido.Panels.Add (5)
                  barPedido.Panels(5).Text = Trim(Msg5)
                  barPedido.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_RODAPE_PEDIDO"
End Sub

Function CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO(ESTAB_ID As Long, DIA_COMPRA As String, CPF_ID As String, VALOR_N As Double) As Boolean
'On Error GoTo ERRO_TRATA

   Dim TabPed                       As New ADODB.Recordset
   Dim Valor_Compra_Dia             As Double

   CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO = False
   Valor_Compra_Dia = 0

   If ESTAB_ID > 0 And IsDate(DIA_COMPRA) And CPF_ID <> "" Then
      If TabPed.State = 1 Then _
         TabPed.Close

      SQL = "select (PEDIDOITEM.QTD_PEDIDA * PEDIDOITEM.VALOR_ITEM) AS TotalCompra"
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
      SQL = SQL & " INNER JOIN USUARIO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.CGCCPF = USUARIO.CPF"

      SQL = SQL & " WHERE pedido.estabelecimento_id = " & ESTAB_ID
      SQL = SQL & " AND pedido.STATUS <> 9 "
      SQL = SQL & " and dt_req >= '" & DMA(DIA_COMPRA, "i") & "'"
      SQL = SQL & " and dt_req <= '" & DMA(DIA_COMPRA, "f") & "'"
      SQL = SQL & " and PRODUCAO = 1"         'aqui somente produtos de produção
      SQL = SQL & " and cpf = '" & (CPF_ID) & "'"
      SQL = SQL & " and funcionario = 1"
      SQL = SQL & " and pedidoitem.status <> 'C' "

      TabPed.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabPed.EOF
                     
         Valor_Compra_Dia = Valor_Compra_Dia + TabPed.Fields("TotalCompra").Value

         TabPed.MoveNext
      Wend
      If TabPed.State = 1 Then _
         TabPed.Close

      If ((Valor_Compra_Dia + VALOR_N) > Valor_Compra_Dia_Permitida) And Valor_Compra_Dia_Permitida > 0 Then _
         CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO = True
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO"
End Function

Function CHECA_FUNCIONARIO(CPF_ID As String) As Boolean
'On Error GoTo ERRO_TRATA

   Dim TabPed                       As New ADODB.Recordset

   CHECA_FUNCIONARIO = False
   txtCNPJCPF.PromptInclude = False

   If Trim(CPF_ID) <> "" Then
      If TabPed.State = 1 Then _
         TabPed.Close

      SQL = "select funcionario from USUARIO WITH (NOLOCK)"
      SQL = SQL & " where cpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabPed.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPed.EOF Then _
         If Not IsNull(TabPed.Fields("funcionario").Value) Then _
            If TabPed.Fields("funcionario").Value = True Then _
               CHECA_FUNCIONARIO = True
      If TabPed.State = 1 Then _
         TabPed.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_FUNCIONARIO"
End Function

Private Sub REMOVE_MENU()
   Dim hMenu As Long
   hMenu = GetSystemMenu(hwnd, False)
   DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Function TRATA_SAIDA_TELA() As Boolean
'On Error GoTo ERRO_TRATA

   TRATA_SAIDA_TELA = False

   If PEDIDO_ID_N > 0 And INDR_CAIXA = True Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select status from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            If TabTemp.Fields(0).Value <= 2 Then
               Msg = "Pedido pendente, deseja realmente cancelar essa venda ?"
               PERGUNTA Msg, vbYesNo + 32, "Cancela Pedido", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
   
                  SQL = "delete from ITEMLANCAMENTO where numr_doc = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
                  SQL = "delete from LANCAMENTO where numr_doc = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
                  SQL = "delete from pedidotemp where pedido_id = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

                  SQL = "update pedido set "
                  SQL = SQL & " status = 9 "
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

                  TRATA_SAIDA_TELA = True
                  Else
                     TRATA_SAIDA_TELA = False
                     Exit Function
               End If
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   TRATA_SAIDA_TELA = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_SAIDA_TELA"
End Function

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtPedido.Text) Then _
      Exit Sub
   If PEDIDO_ID_N <= 0 Then _
      Exit Sub

   Dim TabGrid As New ADODB.Recordset
   Dim Coluna, Linha, Largura_Campo
   Dim VALOR_ITENS_PRODUCAO   As Double
   Dim VALOR_ITENS_REVENDA    As Double

   CONT_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
   VALOR_TOTAL_N = 0
   VALOR_ITENS_PRODUCAO = 0
   VALOR_ITENS_REVENDA = 0

   txtItens.Text = "" & CONT_N
   txtTotalPedido.Text = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTotalPedido.Text = Format(VALOR_TOTAL_N, "currency")

   MSFlexGrid1.Clear
   MSFlexGrid1.Visible = False
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabGrid.State = 1 Then _
      TabGrid.Close

   SQL = "select * from vwPedidoVendaItens WITH (NOLOCK)"

   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and StatusItem <> 'C' "
   SQL = SQL & " order by seq_id desc"

   TabGrid.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGrid.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabGrid.Fields.Count

      ReDim largura_coluna(0 To TabGrid.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabGrid.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabGrid.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabGrid.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabGrid.EOF
         INDR_PRI = False
         If Not IsNull(TabGrid.Fields("producao").Value) Then _
            If TabGrid.Fields("producao").Value = True Then _
               INDR_PRI = True

'=======totais
         CONT_N = CONT_N + 1
         VALOR_ITEM_N = VALOR_ITEM_N + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtde").Value)
         If Not IsNull(TabGrid.Fields("desconto").Value) Then _
            VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabGrid.Fields("desconto").Value

         If INDR_PRI = True Then
            'VALOR_ITENS_PRODUCAO = 0
            VALOR_ITENS_PRODUCAO = VALOR_ITENS_PRODUCAO + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtde").Value)
            Else
               'VALOR_ITENS_REVENDA = 0
               VALOR_ITENS_REVENDA = VALOR_ITENS_REVENDA + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtde").Value)
         End If
'========= verificando se o produto é de produção
'=========

         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabGrid.Fields.Count - 1
            'If Coluna = 3 Or Coluna = 7 Then
            If Coluna = 3 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  'If Coluna = 4 Or Coluna = 5 Or Coluna = 6 Or Coluna = 7 Or Coluna = 8 Or Coluna = 9 Or Coluna = 10 Then
                  If Coluna = 4 Or Coluna = 5 Or Coluna = 6 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabGrid.Fields(Coluna).Value)
                  End If
            End If
'=========se o produto for de produção pintar linha
            If INDR_PRI = True Then
               MSFlexGrid1.Row = Linha
               MSFlexGrid1.Col = Coluna
               'flex_tst.Text = "Bold Font"
               'flex_tst.CellFontBold = True
               'flex_tst.CellForeColor = vbRed
               MSFlexGrid1.CellForeColor = &H4000&   '&H40&
            End If
'=========

            ' verifica o tamanho dos campos
            If Not IsNull(TabGrid.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGrid.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

'=========totais

'--------------
         TabGrid.MoveNext
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
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(5) = 0
         Else: MSFlexGrid1.ColWidth(5) = 1500
      End If
      MSFlexGrid1.ColAlignment(5) = 7

'Total Item
      MSFlexGrid1.ColWidth(6) = 2000
      MSFlexGrid1.ColAlignment(6) = 7

'SITUAÇÃO TRIBUTARIA PRODUTO
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(7) = 1
         Else: MSFlexGrid1.ColWidth(7) = 500
      End If
      MSFlexGrid1.ColAlignment(7) = 0

'ALIQUOTA ICMS
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(8) = 1
         Else: MSFlexGrid1.ColWidth(8) = 500
      End If
      MSFlexGrid1.ColAlignment(8) = 0

'NCM
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(9) = 1
         Else: MSFlexGrid1.ColWidth(9) = 500
      End If
      MSFlexGrid1.ColAlignment(9) = 0

'Pedido_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(10) = 1
         Else: MSFlexGrid1.ColWidth(10) = 500
      End If
      MSFlexGrid1.ColAlignment(10) = 0

'seq_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(11) = 1
         Else: MSFlexGrid1.ColWidth(11) = 500
      End If
      MSFlexGrid1.ColAlignment(11) = 0

'produto_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(12) = 1
         Else: MSFlexGrid1.ColWidth(12) = 500
      End If
      MSFlexGrid1.ColAlignment(12) = 0

'SITUAÇÃO ITEM
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(13) = 1
         Else: MSFlexGrid1.ColWidth(13) = 500
      End If
      MSFlexGrid1.ColAlignment(13) = 0

'familiaproduto_id
      MSFlexGrid1.ColWidth(14) = 50
      MSFlexGrid1.ColAlignment(14) = 0

'producao
      MSFlexGrid1.ColWidth(15) = 0
      MSFlexGrid1.ColAlignment(15) = 0
   End If

   ' fecha o recordset e a conexao
   If TabGrid.State = 1 Then _
      TabGrid.Close

   txtItens.Text = "" & CONT_N

   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N

   txtTotalPedido.Text = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   DoEvents

MSFlexGrid1.Visible = True
   VALOR_ITENS_PRODUCAO = 0
   VALOR_ITENS_REVENDA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then
      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto)
      Exit Sub
      Else: txtProduto.Text = "" & CODG_PRODUTO_A
   End If

   If INDR_PROD_BALANCA = True Then
      Label13.Caption = "Preço/Kg"
      Else: Label13.Caption = "Preço/UN"
   End If

   txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtValor_Unitario.Text = "" & Format(PR_VAREJO_N, strFormatacao2Digitos)
   txtVarejo.Text = "" & Format(PR_VAREJO_N, strFormatacao2Digitos)
   txtAtacado.Text = "" & Format(PR_ATACADO_N, strFormatacao2Digitos)
   txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtPreçoCusto.Text = "" & Format(PR_CUSTO_PRODUTO_N, strFormatacao2Digitos)
   txtDescricao.Text = "" & Trim(DESC_PRODUTO_A)

   If INDR_ESTQ_NEGATIVO = False Then
      QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      If QTDE_ESTOQUE_N <= 0 Then
         MsgBox "Produto sem estoque disponível."

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)

         txtProduto.SetFocus
         Exit Sub
      End If
   End If
   If Len(CODG_NCM_A) > 2 Then
      If Len(CODG_NCM_A) < 8 Then
         MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"

         LIMPA_BODY

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)

         txtProduto.SetFocus
         Exit Sub
      End If
   End If
   If PR_VAREJO_N < 0 Then
      MsgBox "Valor do produto invalido !!!"
      Exit Sub
   End If

   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            FraSeq.Enabled = True
            txtProduto.Enabled = True
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            txtProduto.SetFocus
            Exit Sub
         End If
   End If
'=====================
If Trim(txtPedido.Text) <> "" Then
   If Trim(txtSeq.Text) = "" Then
      SEQ_PEDIDO_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Str(PEDIDO_ID_N), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_PEDIDO_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Str(PEDIDO_ID_N), "", "")
            Else: SEQ_PEDIDO_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_PEDIDO_ID_N
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and pedido_ID = " & PEDIDO_ID_N
   SQL = SQL & " and seq_ID = " & Trim(txtSeq.Text)
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      txtValor_Unitario.Text = "" & Format(TabPedidoItem!Valor_Item, strFormatacao2Digitos)
      txtQTDE.Text = "" & Format(TabPedidoItem!QTD_PEDIDA, strFormatacao3Digitos)
      QTDE_PEDIDO = 0 & TabPedidoItem!QTD_PEDIDA
      VALOR_ITEM_N = 0 & TabPedidoItem!Valor_Item
      VALOR_DIFERENCA_N = 0 & TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA
      txtSeq.Text = "" & TabPedidoItem.Fields("seq_id").Value
   End If
End If
   If TabProduto.State = 1 Then _
      TabProduto.Close
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close
'=====================

   FraSeq.Enabled = True

   If INDR_LEU_POR_CODG_BARRAS = True Then
      txtQTDE.Text = 1

      Call PROCESSA_ITEM

      CODIGO_BARRAS_A = ""
      txtProduto.Enabled = True
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto)
      txtProduto.SetFocus
      CODIGO_BARRAS_A = ""
      Exit Sub
   End If

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtValor_Unitario.Text) <> "" Then
            If IsNumeric(txtValor_Unitario.Text) Then

               Call PROCESSA_ITEM

               CODIGO_BARRAS_A = ""
               txtProduto.Enabled = True
               txtProduto.SelStart = 0
               txtProduto.SelLength = Len(txtProduto)
               txtProduto.SetFocus
               CODIGO_BARRAS_A = ""
               Exit Sub
            End If
         End If
      End If
      Else: txtQTDE.SetFocus
   End If
   CODIGO_BARRAS_A = ""

   'If MULT_EMPRESA_B = True And Trim(txtProduto.Text) <> "" And Trim(txtDescricao.Text) <> "" Then _
      txtQtde.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub

Sub VAI_VENDA()
'On Error GoTo ERRO_TRATA

   INDR_GRAVA = False
   If Trim(txtPedido.Text) <> "" Then
      PEDIDO_ID_N = txtPedido.Text
      Else
         MsgBox "Digite Numero da Requisicao para gravar!"
         Exit Sub

         ABRE_PEDIDO
   End If

'===================
   'checa se o funcionário pode comprar produtos de produção conforme a cota diária estabelecida
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "99999999999" And Valor_Compra_Dia_Permitida > 0 Then
      If CHECA_FUNCIONARIO(txtCNPJCPF.Text) = True Then
         If CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO(ESTABELECIMENTO_ID_N, Date, Trim(txtCNPJCPF.Text), 0) = True Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Cota diária de compra de produtos de produção ultrapassada, não permitido."
            Exit Sub
         End If
      End If
   End If
'===================
   txtCNPJCPF.PromptInclude = False
   CHECA_CLIENTE txtCNPJCPF.Text, txtNome.Text

   GERA_VENDA

'========================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select status from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where PEDIDO_ID = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If IsNull(TabTemp.Fields(0).Value) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus

         Exit Sub
      End If
      If TabTemp.Fields(0).Value <= 2 Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus

         Exit Sub
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   '===================================

   LIMPA_TUDO
   ABRE_PEDIDO

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VAI_VENDA"
End Sub

Sub VALIDA_PEDIDO(NUMR_PEDIDO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Dim TabPed              As New ADODB.Recordset
   Dim SITUACAO_PEDIDO_A   As String

   CRITERIO_A = ""
   SITUACAO_PEDIDO_A = ""

   If TabPed.State = 1 Then _
      TabPed.Close

   SQL = "select CGCCPF,VENDEDOR_ID,nome_cliente,DT_REQ,status"
   SQL = SQL & " from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & NUMR_PEDIDO_ID_N
   TabPed.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPed.EOF Then
      SITUACAO_PEDIDO_A = "" & TabPed.Fields("STATUS").Value
      INDR_VENDA = True

      CNPJCPF_A = "" & Trim(TabPed.Fields("CGCCPF").Value)
      VENDEDOR_ID_N = 0 & TabPed.Fields("VENDEDOR_ID").Value
      'TIPOVENDA_ID_N = 0 & TabPed.Fields("tipovenda_id").Value
      NOME_CLIENTE_A = "" & Trim(TabPed.Fields("nome_cliente").Value)

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = "" & Trim(CNPJCPF_A)

      cmbVendAux.Text = "" & VENDEDOR_ID_N
      cmbVendedor.Text = "" & TRAZ_NOME_VENDEDOR(VENDEDOR_ID_N)

      txtNome.Text = "" & NOME_CLIENTE_A

      txtDtEmis.Text = TabPed.Fields("DT_REQ").Value

      If SITUACAO_PEDIDO_A = 9 Then
         MsgBox "Pedido cancelado, impossível alterar !!!"
         Exit Sub
         Else '1=ORÇAMENTO;2=GERADO;3=EMITIDA COM NOTA;4=EMITIDA COM CUPOM;5=ARECEBER;7=ECF/NF;9=CANCELADO
            If (SITUACAO_PEDIDO_A = 3 Or SITUACAO_PEDIDO_A = 5) Then
               Toolbar1.Buttons(3).Visible = False
               Toolbar1.Buttons(8).Visible = False

               If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
                  Toolbar1.Buttons(9).Visible = True

               MsgBox "Venda ja Faturada !!!", vbOKOnly + 32, "PedidoVenda", "DEMO.HLP", 1000

               Exit Sub
            End If
            If SITUACAO_PEDIDO_A = 4 Then
               MsgBox "Permitido somente consulta !!!"
               Exit Sub
            End If
      End If
   End If
   If TabPed.State = 1 Then _
      TabPed.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VALIDA_PEDIDO"
End Sub

Sub PROCESSA_ITEM()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = "" & Trim(txtCNPJCPF.Text)
   If Trim(CNPJCPF_A) = "" Then
      txtCNPJCPF.Text = "99999999999"
      CNPJCPF_A = "99999999999"
      Else: CHECA_CLIENTE CNPJCPF_A, Trim(txtNome.Text)
   End If

   If Trim(UF_CLIENTE_A) = "" Then _
      TRATA_PESSOA CNPJCPF_A

   If Trim(UF_CLIENTE_A) = "" Then _
      UF_CLIENTE_A = "GO"

   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
   If TabCliente.State = 1 Then _
      TabCliente.Close

   QTDE_N = 0 & txtQTDE.Text
   VALOR_ITEM_N = 0 & txtValor_Unitario.Text
   TIPOVENDA_ID_N = 9999

   If QTDE_N <= 0 Then
      Msg = "Atenção quantidade informada inválida !!!"
      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus
      Exit Sub
      Else
         If QTDE_N > 99 Then
            Msg = "Atenção quantidade informada muito alta, deseja continuar ???? !!!"
            PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            If RESPOSTA = vbNo Then
               FraSeq.Enabled = True
               txtProduto.Enabled = True
               txtProduto.SetFocus
               Exit Sub
            End If
         End If
   End If
   'quantidade pedida
   If INDR_CONTROLA_ESTOQUE = True Then
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

   If VALOR_ITEM_N <= 0 Then
      MsgBox "Atenção Valor informado inválido !!!", vbOKOnly, "Atenção."
      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Informe codigo de Produto.", vbOKOnly, "Atenção."
      FraSeq.Enabled = True
      txtProduto.Enabled = True
      txtProduto.SetFocus
      Exit Sub
   End If

   VALOR_TOTAL_DESCONTO_N = 0

   'valor total da Pedido, o desconto é armazenado no seu devido lugar, não entra no calculo do campo total da venda
   'VALOR_TOTAL_N = VALOR_TOTAL_N + (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N

'===================
   'checa se o funcionário pode comprar produtos de produção conforme a cota diária estabelecida
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "99999999999" And Valor_Compra_Dia_Permitida > 0 And INDR_PRODUTO_PRODUCAO_B = True Then
      If CHECA_FUNCIONARIO(txtCNPJCPF.Text) = True Then
         If CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO(ESTABELECIMENTO_ID_N, Date, Trim(txtCNPJCPF.Text), (QTDE_PEDIDO * VALOR_ITEM_N)) = True Then
            MsgBox "Cota diária de compra de produtos de produção ultrapassada, não permitido."
            Exit Sub
         End If
      End If
   End If

   If PEDIDO_ID_N >= 0 Then
      GRAVA_PEDIDO_CABECA "P", 1
      GRAVA_PEDIDO_ITEM "VENDA"
   End If

   SETA_GRID
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
End Sub

Sub COMANDA_CHAMA(Tipo_Chamada As String)
'On Error GoTo ERRO_TRATA

   SEQ_PEDIDO_ID_N = 0
   COMANDA_ID_N = 0
   CARTAOBARRA_ID_N = 0

   frmCOMANDA.Show 1

''''''''''''''aqui

'MsgBox "voltando da tela de comanda " & CARTAOBARRA_ID_N

   COMANDA_VERIFICA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "COMANDA_CHAMA"
End Sub

Sub COMANDA_VERIFICA()
'On Error GoTo ERRO_TRATA

   SITUACAO_COMANDA_A = ""
'=========================LENDO PELO NUMERO DO PEDIDO
   If TabCOMANDA.State = 1 Then _
      TabCOMANDA.Close

'CARTAOBARRA_ID_N FOI CARREGADA NA ROTINA COMANDA_CHAMA AO CHAMAR frmCOMANDA.Show 1
'VERIFICANDO SE A COMANDA ESTÁ (ATIVA ; PENDENTE)

   If PEDIDO_ID_N > 0 Then
      SQL = "select * from PEDIDOCOMANDA WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N

'LENDO PELO NUMERO DE PEDIDO PARA CERIFICAR SE A COMANDA INFORMADA PERTENCE AO MESMO NUMERO DE PEDIDO ATUAL DA TELA

      TabCOMANDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCOMANDA.EOF Then
         If PEDIDO_ID_N <> TabCOMANDA.Fields("pedido_id").Value Then
            SQL = TabCOMANDA.Fields("pedido_id").Value
            If TabCOMANDA.State = 1 Then _
               TabCOMANDA.Close
            MsgBox "Comanda nº: " & CARTAOBARRA_ID_N & " , " & "já registrada no pedido nº: " & SQL
            Exit Sub
         End If
         SEQ_PEDIDO_ID_N = 0 & TabCOMANDA.Fields("seq_PEDIDO_id").Value
         SEQ_COMANDA_ID_N = 0 & TabCOMANDA.Fields("seq_COMANDA_id").Value
         'CARTAOBARRA_ID_N = 0 & TabComanda.Fields("cartaobarra_id").Value
         'Else: MsgBox "Não existe itens nesta comanda nº: " & CARTAOBARRA_ID_N
      End If
      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close
   End If   'If PEDIDO_ID_N > 0 Then
'=========================

   If CARTAOBARRA_ID_N > 0 Then
      Dim TabComandaItem   As New ADODB.Recordset
'=========================LENDO PELO NUMERO DA COMANDA
      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close

      SQL = "select * from PEDIDOCOMANDA WITH (NOLOCK)"
      SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N

'LE PELO NUMERO DA COMANDA PARA CERTIFICAR QUE PERTENCE AO MESMO NUMERO DE PEDIDO DA TELA ATUAL

      TabCOMANDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCOMANDA.EOF Then
         If PEDIDO_ID_N <> TabCOMANDA.Fields("pedido_id").Value Then
            SQL = TabCOMANDA.Fields("pedido_id").Value
            If TabCOMANDA.State = 1 Then _
               TabCOMANDA.Close
            MsgBox "Comanda nº: " & CARTAOBARRA_ID_N & " , " & "já registrada no pedido nº: " & SQL
            Exit Sub
         End If
         SEQ_PEDIDO_ID_N = 0 & TabCOMANDA.Fields("SEQ_PEDIDO_ID").Value
         SEQ_COMANDA_ID_N = 0 & TabCOMANDA.Fields("SEQ_COMANDA_ID").Value
         'Else: MsgBox "Não existe itens nesta comanda nº: " & CARTAOBARRA_ID_N
      End If
      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close
'=========================

      COMANDA_ID_N = 0

      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close
      'procurando registro comanda POR NUMERO DA COMANDA INFORMADA
      SQL = "SELECT COMANDA.*, "
      SQL = SQL & " COMANDAITEM.SEQ_ID, COMANDAITEM.PRODUTO_ID, COMANDAITEM.QTDE, COMANDAITEM.VALOR_ITEM, "
      SQL = SQL & " COMANDAITEM.SITUACAO AS SitItem, COMANDAITEM.USUARIO_ID AS ATENDENTE_ID"
      SQL = SQL & " FROM COMANDA WITH (NOLOCK) "
      SQL = SQL & " INNER JOIN COMANDAITEM WITH (NOLOCK) "
      SQL = SQL & " ON COMANDA.COMANDA_ID = COMANDAITEM.COMANDA_ID"

      SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N
      'SQL = SQL & " and upper(COMANDAITEM.SITUACAO) = 'ABERTA'"
      TabCOMANDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCOMANDA.EOF Then
         COMANDA_ID_N = 0 & TabCOMANDA.Fields("comanda_id").Value
         SITUACAO_COMANDA_A = "" & TabCOMANDA.Fields("situacao").Value
         SEQ_COMANDA_ID_N = 0 & TabCOMANDA.Fields("seq_id").Value

         If SITUACAO_COMANDA_A = "ABERTA" Then
            If Trim(txtPedido.Text) <> "" Then
               If IsNumeric(txtPedido.Text) Then
                  PEDIDO_ID_N = 0 & txtPedido.Text
               End If
            End If

            'VAI ENTRAR AQUI POIS TEM ITENS NA COMANDA
            SEQ_PEDIDO_ID_N = 0

            COMANDA_GRAVA_ITENS_PEDIDO
            'Else: MsgBox "Não permitido, comanda com situação: " & SITUACAO_COMANDA_A
         End If
      End If
      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close

      SETA_GRID
   End If   'If CARTAOBARRA_ID_N > 0 Then

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "COMANDA_VERIFICA"
End Sub

Private Sub GRAVA_PEDIDO_CABECA(TIPO_REGISTRO_A As String, SITUACAO_N As Integer)
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""
   TIPOVENDA_ID_N = 9999
   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = "" & txtCNPJCPF.Text

'=======================
   If Trim(txtPedido.Text) = "" Then
      GERA_PEDIDO_ID
      txtPedido.Text = PEDIDO_ID_N
   End If
   If PEDIDO_ID_N <= 0 Then
      GERA_PEDIDO_ID
      txtPedido.Text = PEDIDO_ID_N
   End If
'====================

   If CLIENTE_ID_N <= 0 Then
      Call txtCNPJCPF_LostFocus
   End If

   txtPedido.Text = PEDIDO_ID_N

   If TabCabeca.State = 1 Then _
      TabCabeca.Close
   SQL = "select status,nome_cliente,pedido_id,cliente_id from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      Acao_N = 2
      CLIENTE_ID_N = 0 & TabCabeca.Fields("cliente_id").Value
      Else: Acao_N = 1
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   spPedido Acao_N, Trim(CNPJCPF_A), Now, SITUACAO_N, TIPO_REGISTRO_A, 0, 0, 0, Trim(txtNome.Text), 0, VALOR_TOTAL_N, "", 0

'====================
   SQL = "select pedido_id from PEDIDOFATURA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      Acao_N = 2
      Else: Acao_N = 1
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   'If TABELAPRECO_ID_N <= 0 Then _

   TABELAPRECO_ID_N = 0 '& cmbTabPrecoAux.Text

   'If FORMAPAGTO_ID_N <= 0 Then _

   FORMAPAGTO_ID_N = 1  '0 & cmbFormaAUX.Text

   TIPOVENDA_ID_N = 9999

spPEDIDOFATURA Acao_N, 0, PEDIDO_ID_N, TABELAPRECO_ID_N, FORMAPAGTO_ID_N, TIPOVENDA_ID_N
'====================

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PEDIDO_CABECA"
End Sub

Private Sub GRAVA_PEDIDO_ITEM(ORIGEM_ITEM_A As String)
'On Error GoTo ERRO_TRATA

'set
'aqui quando for o caso de já existirem itens no pedido e chamar a comanda tem que criar
'jeito de relacionar e lincar sequencia de itens da comanda com sequencia de itens do pedido em tela
'pois podem usar o mesmo numero de sequencia.
'talvez colocar decimal na sequencia
'FOI CRIADO VARIAVEL ORIGEM_ITEM_A, É DE ONDE VEM A CHAMADA DA ROTINA, SE FOR NA HORA DA COMANDA ENTRA NO PRIMEIRO IF

   If Trim(ORIGEM_ITEM_A) = "COMANDA" And SEQ_PEDIDO_ID_N <= 0 Then
      SEQ_PEDIDO_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Str(PEDIDO_ID_N), "", "")
      Else
      '=====================
         If Trim(txtSeq.Text) = "" Then
            SEQ_PEDIDO_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Str(PEDIDO_ID_N), "", "")
            Else
               If Not IsNumeric(txtSeq.Text) Then
                  SEQ_PEDIDO_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Str(PEDIDO_ID_N), "", "")
                  Else: SEQ_PEDIDO_ID_N = txtSeq.Text
               End If
         End If
      '=====================
   End If

   Dim TabPedidoItem    As New ADODB.Recordset
'====================

If Trim(ORIGEM_ITEM_A) = "COMANDA" Then
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select seq_pedido_id from PEDIDOCOMANDA WITH (NOLOCK)"
   SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N

   SQL = SQL & " and seq_COMANDA_id = " & SEQ_COMANDA_ID_N

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      Acao_N = 2
      SEQ_PEDIDO_ID_N = 0 & TabConsulta.Fields("seq_pedido_id").Value
      Else
         Acao_N = 1
         SQL = "spPedidoComanda " & Acao_N & "," & PEDIDO_ID_N & "," & CARTAOBARRA_ID_N & "," & SEQ_COMANDA_ID_N & "," & SEQ_PEDIDO_ID_N
         CONECTA_RETAGUARDA.Execute "EXEC " & SQL
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
End If

'====================

   SQL = "select pedido_id from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N

SQL = SQL & " and seq_id = " & SEQ_PEDIDO_ID_N

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabPedidoItem.EOF Then
      Acao_N = 1
      Else: Acao_N = 2
   End If
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   spPedidoItem Acao_N, _
                QTDE_N, _
                VALOR_ITEM_N, _
                0, _
                "5102", _
                "00", _
                VALOR_ITEM_N, _
                17, _
                (VALOR_ITEM_N * 17 / 100), 0, 0, 0, 0, 0, 0, 0, 0, _
                "P", PR_CUSTO_PRODUTO_N, "PC", 0, 0, 0, 0, _
                SEQ_PEDIDO_ID_N, SEQ_COMANDA_ID_N, ORIGEM_ITEM_A

txtCNPJCPF.PromptInclude = False

   PREPARA_TRIBUTACAO_PRODUTO Trim(txtCNPJCPF.Text), tpMOEDA(VALOR_ITEM_N), tpMOEDA(QTDE_N)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PEDIDO_ITEM"
End Sub

Sub COMANDA_GRAVA_ITENS_PEDIDO()
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N > 0 Then
      Msg = "Incluir itens da comanda nº: " & CARTAOBARRA_ID_N & " neste pedido, deseja continuar ???? !!!"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbNo Then
         If TabCOMANDA.State = 1 Then _
            TabCOMANDA.Close
         MsgBox "Finalize esta venda antes de prosseguir !!!"
         Exit Sub
      End If
   End If

   lblIDComanda.Visible = False
   lblComanda.Visible = False

   If CARTAOBARRA_ID_N > 0 Then
      lblIDComanda.Visible = True
      lblComanda.Visible = True
      If Trim(lblIDComanda.Caption) <> "" Then
         lblIDComanda.Caption = lblIDComanda.Caption & " ; " & CARTAOBARRA_ID_N
         Else: lblIDComanda.Caption = CARTAOBARRA_ID_N
      End If
   End If

   GRAVA_PEDIDO_CABECA "P", 1

   While Not TabCOMANDA.EOF

      VALOR_ITEM_N = 0 & TabCOMANDA.Fields("valor_item").Value
      QTDE_N = 0 & TabCOMANDA.Fields("qtde").Value
      PRODUTO_ID_N = 0 & TabCOMANDA.Fields("produto_id").Value
      PR_CUSTO_PRODUTO_N = 0
      SEQ_COMANDA_ID_N = 0 & TabCOMANDA.Fields("SEQ_ID").Value
      SEQ_PEDIDO_ID_N = 0
      txtSeq.Text = ""

'txtSeq.Text = "" & SEQ_PEDIDO_ID_N

      GRAVA_PEDIDO_ITEM "COMANDA"

      TabCOMANDA.MoveNext
   Wend

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "COMANDA_GRAVA_ITENS_PEDIDO"
End Sub

Public Function TRAZ_COMANDA_ID(CARTAOBARRA_ID_N As Long) As Long
'On Error GoTo ERRO_TRATA

   TRAZ_COMANDA_ID = 0
   If CARTAOBARRA_ID_N > 0 Then
      Dim TaBCartaoBarra   As New ADODB.Recordset
      Dim strSQL           As String

      If TaBCartaoBarra.State = 1 Then _
         TaBCartaoBarra.Close

      strSQL = "select comanda_id from COMANDA WITH (NOLOCK)"
      strSQL = strSQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N
      TaBCartaoBarra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TaBCartaoBarra.EOF Then _
         TRAZ_COMANDA_ID = "" & Trim(TaBCartaoBarra.Fields(0).Value)
      If TaBCartaoBarra.State = 1 Then _
         TaBCartaoBarra.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRAZ_COMANDA_ID"
End Function

Sub TRIBUTA()

   'cTritutacao.
End Sub
