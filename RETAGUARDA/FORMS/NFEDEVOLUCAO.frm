VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNFEDEVOLUCAO 
   Caption         =   "NFe Diversas"
   ClientHeight    =   7950
   ClientLeft      =   2085
   ClientTop       =   2475
   ClientWidth     =   11010
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
   Icon            =   "NFEDEVOLUCAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
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
      TabIndex        =   15
      Top             =   1680
      Width           =   10935
      Begin VB.TextBox txtValor_Unitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   3
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   360
         Left            =   4530
         MaxLength       =   29
         TabIndex        =   24
         Top             =   240
         Width           =   6255
      End
      Begin VB.TextBox txtQTDE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   9480
         TabIndex        =   4
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAtacado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1560
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtVarejo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3000
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3540
         Picture         =   "NFEDEVOLUCAO.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtDolar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtValorDolar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   19
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtPreçoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5040
         MaxLength       =   12
         TabIndex        =   18
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdMata 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   4040
         Picture         =   "NFEDEVOLUCAO.frx":0F8C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   10440
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Produto:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   31
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Unitário"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7455
         TabIndex        =   30
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9735
         TabIndex        =   29
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Atacado"
         Height          =   240
         Left            =   1785
         TabIndex        =   28
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Varejo"
         Height          =   240
         Left            =   3390
         TabIndex        =   27
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Dolar Dia"
         Height          =   240
         Left            =   225
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.em Dolar"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4695
         TabIndex        =   25
         Top             =   720
         Width           =   1200
      End
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
      TabIndex        =   5
      Top             =   600
      Width           =   10935
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3080
         Picture         =   "NFEDEVOLUCAO.frx":1DCD
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPAGAR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtLIMITE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   360
         Left            =   7920
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3540
         MaxLength       =   100
         TabIndex        =   6
         Top             =   720
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtDtEmis 
         Height          =   360
         Left            =   9480
         TabIndex        =   9
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
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Informe o CNPJ/CPF/Código do cliente, F7-Consultar"
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Pedido:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   14
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt.Pedido:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   8385
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "À Pagar"
         Height          =   240
         Left            =   9000
         TabIndex        =   12
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Crédito"
         Height          =   240
         Left            =   7200
         TabIndex        =   11
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Cliente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   810
      End
   End
   Begin ComctlLib.StatusBar stBarReq 
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   7560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   10
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Diponível="
            TextSave        =   "Diponível="
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   "disponivel"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Vlr.Unitário="
            TextSave        =   "Vlr.Unitário="
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   ""
            Key             =   "unitario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Desconto="
            TextSave        =   "Desconto="
            Key             =   "descvalr_unit"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   ""
            Key             =   "desconto"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Itens="
            TextSave        =   "Itens="
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "Valor Total ="
            TextSave        =   "Valor Total ="
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "total"
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
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
         NumButtons      =   8
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Caixa"
            Key             =   "receber"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   10200
         TabIndex        =   37
         Top             =   240
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
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":27CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":3969
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":49F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":59AD
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":6AB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":7C0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":8060
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":9ED7
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDEVOLUCAO.frx":B58D
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
         TabIndex        =   34
         Top             =   0
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView LISTAPRODUTO 
      Height          =   4305
      Left            =   0
      TabIndex        =   35
      ToolTipText     =   "Clique para selecionar um produto ja gravado."
      Top             =   3240
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   7594
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Produto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vlr.Unitário"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total Produto"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "ST"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "NCM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "seq_id"
         Object.Width           =   1411
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11010
      DesignHeight    =   7950
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
      TabIndex        =   36
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmNFEDEVOLUCAO"
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
   Dim CLIENTE_ID_N       As Long
   Dim TIPO_NOTA_A            As String

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

   Dim Valr_Venda_Produto_n   As Double
   Dim QTDE_N                 As Double

   Dim CODIGO_BARRAS          As String

   Private CalculaIcmsG As New MegasimCL.mCalculaIcms ' Yuri alterado em 01/05/2012

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INICIALIZA_VENDA

   If VALR_DOLAR_DIA(Date) <= 0 Then
      Label12.Visible = False
      txtDolar.Visible = False
      Label14.Visible = False
      txtValorDolar.Visible = False
   End If

   Call txtPedido_LostFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
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
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub LISTAPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(LISTAPRODUTO.SelectedItem.Text) Then
             txtProduto.Text = LISTAPRODUTO.SelectedItem.Text
             txtSeq.Text = Trim(LISTAPRODUTO.SelectedItem.ListSubItems.Item(8).Text)
         End If

         EXCLUIR_ITEM
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTAPRODUTO_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
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

         'Call txtPedido_LostFocus
         txtPedido.Enabled = True
         txtPedido.SetFocus
      Case "voltar"
         Unload Me
      Case "produto"
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then _
            frmCADASTROPRODUTO.Show 1
      Case "CadCliente"
          frmCADASTROCLIENTE.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdPesquisar_Click()
   CONSULTA_PRODUTO
End Sub

Private Sub cmdConsCli_Click()
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   frmDISPLAYCLIENTE.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = "99999999999"
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.Text = CNPJCPF_A
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus
End Sub

Private Sub cmdMata_Click()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" And Trim(txtProduto.Text) <> "" Then
      EXCLUIR_ITEM
      Else: MsgBox "Informe código produto."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMata_Click"
End Sub

Private Sub listaproduto_DblClick()
'On Error GoTo ERRO_TRATA

    If Not IsNull(LISTAPRODUTO.SelectedItem.Text) Then
        txtProduto.Text = LISTAPRODUTO.SelectedItem.Text
        txtSeq.Text = Trim(LISTAPRODUTO.SelectedItem.ListSubItems.Item(8).Text)
        txtProduto.SetFocus
    End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "listaproduto_DblClick"
End Sub

Private Sub listaproduto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTAPRODUTO, ColumnHeader
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

Private Sub txtDolar_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtDolar.Text) <> "" Then
      If IsNumeric(txtDolar.Text) Then
         txtDolar.Text = Format(txtDolar.Text, strFormatacao2Digitos)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDolar_LostFocus"
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

Private Sub txtValor_Unitario_Change()
'On Error GoTo ERRO_TRATA

   If Trim(txtValor_Unitario.Text) <> "" Then
      If IsNumeric(txtValor_Unitario.Text) Then
         txtDolar.Text = Format(VALR_DOLAR_DIA(Date), strFormatacao2Digitos)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_Unitario_Change"
End Sub

Private Sub txtValorDolar_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If Trim(txtValorDolar.Text) <> "" Then
      If IsNumeric(txtValorDolar.Text) Then
         Dim DOLAR_DIA              As Double
         Dim valor_dolar_digitado   As Double

         valor_dolar_digitado = 0 & txtValorDolar.Text
         DOLAR_DIA = VALR_DOLAR_DIA(Date)

         VALOR_ITEM_N = DOLAR_DIA * valor_dolar_digitado

         txtValor_Unitario.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
         txtValor_Unitario.Refresh
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDolar_KeyPress"
End Sub

Private Sub txtValorDolar_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtValorDolar.Text) <> "" Then
      If IsNumeric(txtValorDolar.Text) Then
         txtValorDolar.Text = Format(txtValorDolar.Text, strFormatacao2Digitos)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDolar_LostFocus"
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

   If Trim(txtVarejo.Text) <> "" Then
      If IsNumeric(txtVarejo.Text) Then
         txtVarejo.Text = Format(txtVarejo.Text, strFormatacao2Digitos)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtVarejo_LostFocus"
End Sub

Private Sub TXTCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtNome_LostFocus()
'On Error GoTo ERRO_TRATA

   txtNome.Text = UCase(txtNome.Text)
   txtNome.Enabled = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_LostFocus"
End Sub
'==================cgccpf
Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "F7-Consulta Clientes", "Inform Cliente", "", ""
   
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Mask = "###############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
         frmDISPLAYCLIENTE.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = "99999999999"
            txtCNPJCPF.Mask = "##############"

            txtCNPJCPF.Text = CNPJCPF_A
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

      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) = "" Then
         MsgBox "Informe cliente para devolução."
         Exit Sub
      End If

      TRATA_CLIENTE

      txtCNPJCPF.PromptInclude = True

      SendKeys "{tab}"
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If txtPedido.Text <> "" And Trim(txtProduto.Text) <> "" Then
            EXCLUIR_ITEM
            Else: MsgBox "Informe código produto."
         End If
      Case vbKeyF7
         CONSULTA_PRODUTO
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

      UCase (txtProduto.Text)
      CODIGO_BARRAS = ""

      If Trim(txtProduto.Text) <> "" Then
         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select * from PRODUTO "
         SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabProduto.EOF Then
            CODIGO_BARRAS = "" & Trim(txtProduto.Text)
            QTDE_N = 0

            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select * from PRODUTO "
            SQL = SQL & " where CODG_barra = '" & Trim(txtProduto.Text) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabProduto.EOF Then
               If Len(CODIGO_BARRAS) = 13 Then
                  '2 = produtos "in store" (sempre será 2)
                  'C = código do produto (4,5 ou 6 dígitos)
                  'T = total a pagar (sempre 6 dígitos)
                  'P = peso (sempre 5 dígitos)
                  'Q = quantidade (sempre 5 dígitos)
                  '0 = zero fixo
                  'DV = dígito verificador do EAN-13

                  txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 6))
                  QTDE_N = 0 & Int(Mid(CODIGO_BARRAS, 8, 5))   'gramas

                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select * from PRODUTO "
                  SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  SQL = SQL & " and situacao <> 'C' "
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabProduto.EOF Then
                     MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
                     txtProduto.SelStart = 0
                     txtProduto.SelLength = Len(txtProduto)
                     txtProduto.SetFocus
                     Exit Sub
                     Else: MOSTRA_DADOS_PRODUTO
                  End If
               End If
               Else: MOSTRA_DADOS_PRODUTO
            End If
            Else: MOSTRA_DADOS_PRODUTO
         End If
         Else: MsgBox "Informe código do produto !!!"
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "Informe a quantidade", "F10-Gravar", "", ""
   
   If Trim(txtProduto.Text) = Empty Then
      MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
      txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   If txtQTDE.Text <> "" Then
      txtQTDE.SelStart = 0
      txtQTDE.SelLength = Len(txtQTDE)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
'On Error GoTo ERRO_TRATA

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

   PROCESSA_ITEM

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

   If Trim(txtPedido.Text) = "" Then _
      txtPedido.Enabled = False

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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_GotFocus"
End Sub

Private Sub TXTVALOR_UNITARIO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If STATUS_PROD = "P" Then
         If txtValor_Unitario.Text < PRECO_PROD Then
            MsgBox "Produto Tipo Promoção Impossível dar desconto."
            txtValor_Unitario.Text = 0
            txtValor_Unitario.SetFocus
            Else: txtQTDE.SetFocus
         End If
         Else
            If txtValor_Unitario.Text <> VLR_ANTERIOR_N Then
                If txtValor_Unitario.Text < PRECO_PROD Then
                   VALOR_DESCONTO_N = Format(PRECO_PROD - txtValor_Unitario.Text, strFormatacao2Digitos)
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
               SendKeys "{tab}"
               Exit Sub
               Else: SendKeys "{tab}"
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

   If txtValor_Unitario.Text = "" Then
      txtValor_Unitario.Text = Format(0, strFormatacao2Digitos)
      Else: txtValor_Unitario.Text = Format(txtValor_Unitario.Text, strFormatacao2Digitos)
   End If
   If txtValor_Unitario.Text = "" Then
      MsgBox "Valor Unitário Inválido !!!"
      txtValor_Unitario.SetFocus
      Exit Sub
      Else
         VALOR_ITEM_N = txtValor_Unitario.Text
         txtValor_Unitario.Text = Format(txtValor_Unitario.Text, strFormatacao2Digitos)
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

   If Trim(txtValor_Unitario.Text) <> "" Then _
      If IsNumeric(txtValor_Unitario.Text) Then _
         txtValor_Unitario.Text = Format(txtValor_Unitario.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_LostFocus"
End Sub
'============================subrotinas
Private Sub EXCLUIR_ITEM()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) <> "" And Trim(txtSeq.Text) <> "" Then
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
   
      SQL = SQL & " where codg_prod = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and PEDIDOITEM.numr_req = " & txtPedido.Text
      SQL = SQL & " and PEDIDOITEM.seq_id = " & txtSeq.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

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

            SQL = "Delete FROM PEDIDOITEM "
            SQL = SQL & " Where pedido_id = " & TabTemp.Fields("pedido_id").Value
            SQL = SQL & " and seq_id = " & TabTemp.Fields("seq_id").Value
            SQL = SQL & " and tipo_reg = 'PC' "
            CONECTA_RETAGUARDA.Execute SQL

            If TabTemp.State = 1 Then _
               TabTemp.Close

            LIMPA_BODY
            stBarReq.Panels(10).Text = Format(VALOR_TOTAL_N, "##,##0.00")

            GRAVA_CABECA "D", 1
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

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = TabCABECA!CGCCPF
   txtCNPJCPF.PromptInclude = True

   'MOSTRA CLIENTE
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome,status from CLIENTE "
   SQL = SQL & " where cgccpf = '" & TabCABECA!CGCCPF & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      txtNome.Text = TabCliente!NOME
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_REQ"
End Sub

Private Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
   stBarReq.Panels(6).Text = Format(VALOR_ITEM_N, "##,##0.00")

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

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_ITEM_N = TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select count(CODG_PROD) "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.numr_req = " & txtPedido.Text
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      stBarReq.Panels(8).Text = TabTemp.Fields(0).Value
      stBarReq.Refresh
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   stBarReq.Panels(6).Text = Format(VALOR_DESCONTO_N, "##,##0.00")
   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
   stBarReq.Panels(10).Text = Format(VALOR_TOTAL_N, "##,##0.00")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAIS"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   Aliquota_Icms = 0
   Valr_Venda_Produto_n = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""

   stBarReq.Panels(2).Text = ""
   stBarReq.Panels(6).Text = ""
   stBarReq.Panels(4).Text = ""

   QTDE_PEDIDO = 0
   QTDE_ESTOQUE = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_DIFERENCA_N = 0
   PRODUTO_ID_N = 0

   txtValorDolar.Text = Format(0, strFormatacao2Digitos)
   txtDolar.Text = Format(0, strFormatacao2Digitos)
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

   If TabUSU.State = 1 Then _
      TabUSU.Close

   PRODUTO_ID_N = 0
   Aliquota_Icms = 0
   txtPedido.Text = ""
   txtDtEmis = Format(Date, "dd/mm/yyyy")
   txtNome.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""

   LIMPA_BODY
   LISTAPRODUTO.ListItems.Clear
   stBarReq.Panels(6).Text = ""
   stBarReq.Panels(2).Text = ""
   stBarReq.Panels(10).Text = ""
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

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   LISTAPRODUTO.ListItems.Clear
   CONT_N = 0
   NOME_A = ""

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select PEDIDOITEM.seq_id, PEDIDOITEM.NUMR_REQ, PEDIDOITEM.CODG_PROD, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "
   SQL = SQL & " PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.PRECO_CUSTO, PEDIDOITEM.PERC_DESC, "
   SQL = SQL & " PEDIDO.Status, PEDIDO.TIPO_REGISTRO, PRODUTO.*"

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDOITEM.numr_req = " & txtPedido.Text
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF
      NOME_A = "" & Trim(TabPedidoItem!Descricao)

      CONT_N = CONT_N + 1
      Set Item = LISTAPRODUTO.ListItems.Add(, "seq." & CONT_N, Trim(TabPedidoItem!Codg_Prod))

      Item.SubItems(1) = "" & NOME_A
      Item.SubItems(2) = Format(TabPedidoItem!QTD_PEDIDA, strFormatacao3Digitos)
      Item.SubItems(3) = Format(TabPedidoItem!Valor_Item, strFormatacao3Digitos)

      'If TabPedidoItem!PERC_desc = 0 Then
         'If TabPedidoItem!PRECO_VENDA > TabPedidoItem!Valor_Item Then
            'Item.SubItems(4) = Format((TabPedidoItem!PRECO_VENDA - TabPedidoItem!Valor_Item) * TabPedidoItem!QTD_PEDIDA, strFormatacao2Digitos)
         'Else: Item.SubItems(4) = Format((TabPedidoItem!PERC_desc * (TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA) / 100), strFormatacao2Digitos)
      'End If
Item.SubItems(4) = Format(TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA, strFormatacao2Digitos)
      'Item.SubItems(5) = Format(TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA - (TabPedidoItem!PERC_desc * (TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA) / 100), strFormatacao2Digitos)

      Item.SubItems(5) = TabPedidoItem.Fields("situacao_tributaria").Value

      Item.SubItems(6) = "" & TabPedidoItem.Fields("codg_ncm").Value

      Item.SubItems(7) = "" & TabPedidoItem.Fields("seq_id").Value

      QTDE_ESTOQUE = TabPedidoItem!Qtde

      If TabPedidoItem.Fields("situacao").Value = "A" Then
         Item.ForeColor = vbBlue
         Item.ListSubItems(1).ForeColor = vbBlue
         Item.ListSubItems(2).ForeColor = vbBlue
         Item.ListSubItems(3).ForeColor = vbBlue
         Item.ListSubItems(4).ForeColor = vbBlue
         Item.ListSubItems(5).ForeColor = vbBlue
         Item.ListSubItems(6).ForeColor = vbBlue
      End If
      If TabPedidoItem.Fields("situacao").Value = "P" Then
         Item.ForeColor = vbRed
         Item.ListSubItems(1).ForeColor = vbRed
         Item.ListSubItems(2).ForeColor = vbRed
         Item.ListSubItems(3).ForeColor = vbRed
         Item.ListSubItems(4).ForeColor = vbRed
         Item.ListSubItems(5).ForeColor = vbRed
         Item.ListSubItems(6).ForeColor = vbRed
      End If

      TabPedidoItem.MoveNext
   Wend

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   MOSTRA_TOTAIS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GRAVA_CABECA(TIPO_REGISTRO_A As String, STATUS_N As Integer)
'On Error GoTo ERRO_TRATA

   CRITERIO = ""
   CLIENTE_ID_N = 0

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Mask = "###############"

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "Select * From Cliente Where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value
   If TabCliente.State = 1 Then _
      TabCliente.Close

   PEDIDO_ID_N = 0 & txtPedido.Text

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * FROM PEDIDO "
   SQL = SQL & " where numr_req = " & txtPedido.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and numero_caixa_cpu = " & NUMERO_CAIXA_CPU
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCABECA.EOF Then
      SQL = "INSERT INTO PEDIDO "
         SQL = SQL & "("
            SQL = SQL & "PEDIDO_ID,Empresa_id, numr_req, CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, Status, "
            SQL = SQL & " Tipo_Registro,Codg_USU, TIPOVENDA_ID, CLIENTE_ID, Valor_ToTal,"
            SQL = SQL & " valor_desconto,perc_desc,NUMERO_CAIXA_CPU,estabelecimento_id "
            SQL = SQL & ") "
         SQL = SQL & " VALUES ("
            SQL = SQL & PEDIDO_ID_N
            SQL = SQL & "," & EMPRESA_ID_N
            SQL = SQL & "," & txtPedido.Text
            SQL = SQL & ",'" & txtCNPJCPF.Text & "'"
            SQL = SQL & ",0"
            SQL = SQL & ",'" & DMA(Date) & "'"
            SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
            SQL = SQL & "," & STATUS_N
            SQL = SQL & ",'" & TIPO_REGISTRO_A & "'"
            SQL = SQL & "," & CODG_USU_N
            SQL = SQL & ",9999"
            SQL = SQL & "," & CLIENTE_ID_N
            SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)
            SQL = SQL & "," & tpMOEDA(0)  'vai zerar e tratar somente na tela de desconto
            SQL = SQL & "," & tpMOEDA(0)
            SQL = SQL & "," & NUMERO_CAIXA_CPU           'NUMERO_CAIXA_CPU
            SQL = SQL & "," & ESTABELECIMENTO_ID_N           'estabelecimento_id
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
                        SQL = SQL & ",Vendedor_id = 0"
                        SQL = SQL & ",dt_req = '" & DMA(Date) & "'"
                        SQL = SQL & ",nome_cliente = '" & txtNome.Text & "'"
                        SQL = SQL & ",Status = " & STATUS_N
                        SQL = SQL & ",TIPO_REGISTRO = '" & TIPO_REGISTRO_A & "'"
                        SQL = SQL & ",CODG_USU = " & CODG_USU_N
                        SQL = SQL & ",EMPRESA_ID = " & EMPRESA_ID_N
                        SQL = SQL & ",TIPOvenda_id = 9999"
                        SQL = SQL & ",USUARIO_LIBERA_VENDA = " & CODG_USU_N
                        SQL = SQL & ",CLIENTE_ID = " & CLIENTE_ID_N

                        SQL = SQL & " where numr_req = " & txtPedido.Text
                        SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                        SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
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

   txtCNPJCPF.PromptInclude = False

   If UF_CLIENTE = "" Then
      MsgBox "Cliente com cadastro incompleto !!!"
      txtCNPJCPF.SetFocus
      Exit Sub
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
      SQL = SQL & " (PEDIDO_ID,SEQ_ID,PRODUTO_ID, Numr_req, Codg_Prod, Qtd_Pedida, "
      SQL = SQL & " Valor_item, PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG) "
      SQL = SQL & " VALUES ("

         SQL = SQL & PEDIDO_ID_N                                                          'PEDIDO_id
         SQL = SQL & "," & SEQ_ID_N                                                       'SEQ_ID
         SQL = SQL & "," & PRODUTO_ID_N
         SQL = SQL & "," & txtPedido.Text                                                 'Numr_req
         SQL = SQL & ",'" & Trim(txtProduto.Text)                                         'Codg_Prod
         SQL = SQL & "'," & tpMOEDA(QTDE_PEDIDO)                                          'Qtd_Pedida
         SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                                          'Valor_item
         SQL = SQL & ",0" '& tpMOEDA(PERC_DESCONTO_N)                                       'PERC_DESC
         'SQL = SQL & "," & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)  'valor_desconto
         SQL = SQL & ",0" '& tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)  'valor_desconto
         SQL = SQL & ", 'P'"                                                              'status
         SQL = SQL & "," & tpMOEDA(txtPreçoCusto.Text)                                    'PRECO_CUSTO
         SQL = SQL & ",'PC'"                                                              'TIPO_REG

      SQL = SQL & ")"
      Else
         SQL = "UPDATE PEDIDOITEM SET "
         SQL = SQL & " qtd_pedida = " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & ", Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
         'SQL = SQL & ", PERC_desc = " & tpMOEDA(PERC_DESCONTO_N)
         'SQL = SQL & ", valor_desconto = " & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)
         SQL = SQL & ", status = 'P'"
         SQL = SQL & ", preco_custo = " & tpMOEDA(txtPreçoCusto.Text)
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
   Dim dblTemp As String

   If rstEmpresa.State = 1 Then _
      rstEmpresa.Close

   SQL = "Select empresa.* From EMPRESA "

   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   rstEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstEmpresa.EOF Then
      If rstEmpresa.State = 1 Then _
         rstEmpresa.Close

      MsgBox "O sistema não obteve sucesso ao tentar localizar a empresa corrente."
      Unload Me
      Exit Sub
      Else
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

         SQL = "Select * From ENDERECO Where PROP = '" & rstEmpresa!CGC & "'"
         RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not RstTemp.EOF Then
            If Not IsNull(RstTemp!CEP) Then
               dblTemp = "" & RstTemp!CEP
               Else
                  If rstEmpresa.State = 1 Then _
                     rstEmpresa.Close
                  
                  If RstTemp.State = 1 Then _
                     RstTemp.Close

                  dblTemp = "74000000"
                  MsgBox "Verificar cadastro de empresa!!!"
                  Unload Me
                  Exit Sub
            End If

            If RstTemp.State = 1 Then _
               RstTemp.Close

            SQL = "Select * From CEP Where CEP = " & dblTemp

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
   End If
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

   PERC_DESCONTO_USUARIO = 0
   NUMR_REQ_N = txtPedido.Text
   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = txtCNPJCPF.Text

   'atualizando desconto na cabeça
   SQL = "UPDATE PEDIDO SET "
   'SQL = SQL & " Valor_desconto = " & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
   'SQL = SQL & " , Perc_desc = " & tpMOEDA(PERC_DESCONTO_N)
   SQL = SQL & "  cgccpf = '" & CNPJCPF_A & "'"
   SQL = SQL & " , nome_cliente = '" & Trim(txtNome.Text) & "'"
   SQL = SQL & " , status = 2"
   SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
   SQL = SQL & " , tipo_registro = 'D' "
   SQL = SQL & " where numr_req = " & txtPedido.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

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

   SQL = "select * FROM PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N

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
                  'Msg = "Nota ja Processada, Deseja Reativar para imprimir?"
                  PERGUNTA "Nota Processada para este pedido.", vbNo, "Venda NFE", "DEMO.HLP", 1000
                  'RESPOSTA = vbNo
                  If RESPOSTA = vbYes Then
                     If TabCABECA.State = 1 Then _
                        TabCABECA.Close

                     Else
                        LIMPA_BODY
                        LIMPA_TUDO
                   End If
                   Exit Sub
               End If
               If TabCABECA!Status = 5 Then
                  'Msg = "Venda ja Faturada, Deseja Reativar para imprimir ?"
                  PERGUNTA "Venda ja Faturada, Deseja imprimir ?", vbYesNo + 32, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     GERA_IMPRESSAO
                     Else
                        LIMPA_BODY
                        LIMPA_TUDO
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

   If STATUS_PROD = "P" Then
      MsgBox "Produto em Promoçao, Impossivel fazer devolução."
      txtProduto.SetFocus
   End If

   If txtPedido.Text = "" Then _
      VALIDA_NUMR_REQ

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Informe cliente para devolução."
      Exit Sub
   End If

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Informe codigo de Produto.", vbOKOnly, "Atenção."
      txtProduto.SetFocus
      Exit Sub
   End If

   If Not IsNull(txtValor_Unitario.Text) Then
      If txtValor_Unitario.Text <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção."
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If txtQTDE.Text = "" Then
      Beep
      MsgBox "Informe a quantidade.", vbOKOnly, "Atenção."
      txtQTDE.SetFocus
      Exit Sub
      Else
         'quantidade pedida
         QTDE_PEDIDO = txtQTDE.Text
         txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
         If INDR_CONTROLA_ESTOQUE = True Then

            CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO

         End If
   End If

   'valor venda item
   VALOR_ITEM_N = txtValor_Unitario.Text

   VALOR_DESCONTO_N = 0

   VALOR_TOTAL_DESCONTO_N = 0

   VALOR_TOTAL_N = VALOR_TOTAL_N + (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N

   SQL = "select * FROM PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If TabCABECA.State = 1 Then TabCABECA.Close
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      If TabCABECA!Status <> 3 Then 'Emitido com Nota
         If TabCABECA!Status <> 4 Then ' Emitido com Cupom
            If TabCABECA!Status <> 5 Then ' Apenas Faturado
               GRAVA_CABECA "D", 1

               GRAVA_TUDO_ITEM
            End If
         End If
      End If
      Else 'ainda nao gravou requisicao
         txtCNPJCPF.PromptInclude = False

         If UF_CLIENTE = "" Then
            MsgBox "Cliente com cadastro incompleto !!!"
            txtCNPJCPF.SetFocus
            Exit Sub
         End If

         GRAVA_CABECA "D", 1
         GRAVA_TUDO_ITEM
   End If

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
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

   Dim VALOR_LIMITE_N As Double
   Dim VALOR_PENDENTE_N As Double

   ENDERECO_A = ""
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Informe cliente para devolução."
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

   If txtCNPJCPF.Text <> "" Then
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

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from CLIENTE "
   SQL = SQL & " where CGCCPF = '" & txtCNPJCPF.Text & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCliente.EOF Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      Beep
      MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção."
      txtCNPJCPF.SetFocus
      Exit Sub
      Else
         If TabCliente!NOME <> "" Then _
            txtNome.Text = TabCliente!NOME

         CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value
         If Not IsNull(TabCliente!limite_credito) Then _
            txtLIMITE.Text = Format(TabCliente!limite_credito, strFormatacao2Digitos)

         'Pegou o tipo do cliente
         If Not IsNull(TabCliente!TIPO_CLIENTE) Then _
            dblTipoCliente = TabCliente!TIPO_CLIENTE

         If Not IsNull(TabCliente!CGCCPF) Then _
            strCPFCNPJ = TabCliente!CGCCPF

         If Not IsNull(TabCliente!IE) Then 'O Cara ja tem no Cadastro de Cliente
            strInscEstadual = TabCliente!IE
            Else ' Se ele nao tiver no Cadastro de Cliente pega aqui!
               TabCliente.Close
               MsgBox "Inscrição estatual invalida para este cliente, atualizar."
               Exit Sub
         End If

         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select sum(i.valor_item) from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = l.numr_doc "
         SQL = SQL & " and l.prop = '" & Trim(TabCliente!CGCCPF) & "'"
         SQL = SQL & " and i.status = 'A' "
         SQL = SQL & " and l.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = 1"
         SQL = SQL & " and i.formapagto_id <> 1"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If Not IsNull(TabAUX.Fields(0).Value) Then
               VALOR_PENDENTE_N = 0 & TabAUX.Fields(0).Value
               txtPAGAR.Text = Format(TabAUX.Fields(0).Value, strFormatacao2Digitos)
               txtPAGAR.Refresh
            End If
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         VALOR_LIMITE_N = 0 & TabCliente.Fields("LIMITE_CREDITO").Value

         If VALOR_LIMITE_N > 0 Then
            If VALOR_PENDENTE_N >= VALOR_LIMITE_N Then
               MsgBox "Valor limite de credito para esse cliente ultrapassado, não permitido venda, verificar com departamento financeiro."
               txtCNPJCPF.Text = ""
               txtNome.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If

         If tabEndereco.State = 1 Then _
            tabEndereco.Close

         SQL = "select * from ENDERECO "
         SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & " and tipo = 'C'"
         tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not tabEndereco.EOF Then
            If Not IsNull(tabEndereco!Rua) Then _
               ENDERECO_A = tabEndereco!Rua
            If Not IsNull(tabEndereco!Complemento) Then _
               ENDERECO_A = ENDERECO_A & "," & tabEndereco!Complemento
            If Not IsNull(tabEndereco!Bairro) Then _
               ENDERECO_A = ENDERECO_A & "," & tabEndereco!Bairro

            'Pegou o CEP do cliente
            If IsNull(tabEndereco!CEP) Then
               If tabEndereco.State = 1 Then _
                  tabEndereco.Close
   
               MsgBox "O Cadastro do cliente não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
               txtCNPJCPF.Text = ""
               txtNome.Text = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If

            If TabCEP.State = 1 Then _
               TabCEP.Close
      
            'Pegar a uf do cliente
            TabCEP.Open "Select * From CEP Where CEP = " & tabEndereco!CEP, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCEP.EOF Then
               If Not IsNull(TabCEP!UF) Then
                  UF_CLIENTE = TabCEP!UF
                  Else 'UF nao localizada
                     TabCEP.Close
                     MsgBox "O Cadastro do cliente não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                     txtCNPJCPF.Text = ""
                     txtNome.Text = ""
                     txtCNPJCPF.SetFocus
               End If
               Else
                  If TabCEP.State = 1 Then _
                     TabCEP.Close

                  MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais completos. Verique-os, principalmente o Estado(UF) da empresa"
                  txtCNPJCPF.Text = ""
                  txtNome.Text = ""
                  txtCNPJCPF.SetFocus
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
         End If
         If tabEndereco.State = 1 Then _
            tabEndereco.Close

         If TabCliente!Status = "C" Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            Beep
            MsgBox "Cliente Esta Bloqueado!, Verifique Cadastro!.", vbOKOnly, "Atenção."
            txtCNPJCPF.Text = ""
            txtNome.Text = ""
            txtCNPJCPF.SetFocus
         End If
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

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

   Dim rstProduto             As New ADODB.Recordset
   Dim TabTemp                As New ADODB.Recordset
   Dim strSql                 As String
   Dim dblValorBaseICMS       As Double
   Dim dblPercentualICMS      As Double
   Dim dblValorICMS           As Double
   Dim dblValorBaseICMSSubst  As Double
   Dim dblValorICMSSubst      As Double
   Dim dblPercentualICMSSubst As Double
   Dim dblPercReducICMS       As Double
   Dim dblPercIVA             As Double
   Dim dblTotalItem           As Double

   If CODG_PRODUTO_A = "" Or ClienteId = "" Then
      MsgBox "O sistema esta esperando alguns parametros que nao forma  localizados. Verifique"
      Exit Sub
   End If

   dblValorBaseICMS = 0
   dblPercentualICMS = 0
   dblValorICMS = 0
   dblValorBaseICMSSubst = 0
   dblValorICMSSubst = 0
   dblPercentualICMSSubst = 0
   dblPercReducICMS = 0
   dblPercIVA = 0
   dblTotalItem = 0
   strCFOP = ""
   SITUAÇÃO_TRIBUTARIA_PRODUTO = ""

   If UF_CLIENTE = "" Then _
      TRATA_CLIENTE

   txtCNPJCPF.PromptInclude = False

   If UF_CLIENTE = "" Then
      MsgBox "Cliente com cadastro incompleto !!!"
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   If UF_EMPRESA = "" Then _
      PEGA_DADOS_EMPRESA

   dblTotalItem = (txtQTDE.Text * txtValor_Unitario.Text)

   If rstProduto.State = 1 Then _
      rstProduto.Close

   strSql = "Select * From PRODUTO "
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
   If rstProduto!Situacao_Tributaria = "00" Then
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
   If rstProduto!Situacao_Tributaria = 10 Then 'Substituicao Tributaria
      dblValorBaseICMS = dblTotalItem

      If UF_CLIENTE = UF_EMPRESA Then
         'Campo IVA nao existe nao tabela verificar se precisa
         If Not IsNull(rstProduto!PercIva) Then _
           dblValorBaseICMSSubst = ((dblValorBaseICMS * rstProduto!PercIva) / 100)  'Valor da Reducao da base

         'dblValorBaseICMSSubst = ((dblValorBaseICMS * 1) / 100)  'Valor da Reducao da base
         dblValorICMSSubst = ((dblValorBaseICMSSubst * 17) / 100) 'é fixo o percentual, procurar saber se tem como parametrizar
         dblPercentualICMSSubst = 17
      End If
   End If

   'Com redução de base de cálculo
   If rstProduto!Situacao_Tributaria = 20 Then 'Reducao da base de calculo
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
   If rstProduto!Situacao_Tributaria = 30 Then '//Isenta ou nao Tributada Com ICMS por Subs. Trib
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
   If rstProduto!Situacao_Tributaria = 40 Or rstProduto!Situacao_Tributaria = 41 Then '//Isento ou nao Tributado
      dblValorBaseICMS = 0
      dblPercentualICMS = 0
   End If

'50      Suspensão
'51      Diferimento

   'ICMS cobrado anteriormente por substituição tributária
   If rstProduto!Situacao_Tributaria = 60 Then '//Situacao Tributaria com Substitui‡ao Tributaria
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
      If rstProduto!Situacao_Tributaria = 60 Then
         strCFOP = "5405"  'Fixo por enquanto
         strCFOP = "5405"
         Else: strCFOP = CFOP_SAIDA_DE                     'cfop de venda dentro do estado
      End If

      SQL = "Select * From CFOP "
      SQL = SQL & " Where codigo = '" & strCFOP & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         If rstProduto.State = 1 Then _
            rstProduto.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      'if tabtemp!Tipo = 0 then 'Dentro do Estado
      dblValorICMS = ((dblTotalItem * TabTemp!PERC_ICMS) / 100)
      dblPercentualICMS = TabTemp!PERC_ICMS

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   'FORA DO ESTADO
   If UCase(UF_CLIENTE) <> UCase(UF_EMPRESA) Then
      If rstProduto!Situacao_Tributaria = 60 Then
         strCFOP = "6403"  'Fixo por enquanto
         strCFOP = "6404"
         Else: strCFOP = CFOP_SAIDA_FE                  'cfop de venda fora do estado do estado
      End If

      TabTemp.Open "Select * From CFOP Where CODIGO = '" & strCFOP & "'", CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      If Trim(Len(strCPFCNPJ)) > 11 Then ' Se for pessoa juridica
         dblValorICMS = ((dblTotalItem * TabTemp!PERC_ICMS) / 100)  'CFOP.P_ICMS_VND_F_UF - verificar se existe
         dblPercentualICMS = TabTemp!PERC_ICMS ' CFOP.P_ICMS_VND_F_UF'duas aliquotas para  o mesmo cfop
         Else ' Pessoa fisica
            dblValorICMS = ((dblTotalItem * TabTemp!ICMS_PJ_F_UF) / 100)
            dblPercentualICMS = TabTemp!ICMS_PJ_F_UF
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close
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

SITUAÇÃO_TRIBUTARIA_PRODUTO = "" & rstProduto!Situacao_Tributaria

   'If Not isnull(rstProduto!PERCIVA) Then dblPercIVA = rstProduto!PERCIVA

   If dblValorBaseICMS = 0 Then _
      dblPercentualICMS = 0
   
   If rstProduto.State = 1 Then _
      rstProduto.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "Select * "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " Where PEDIDO.EMPRESA_ID = " & EMPRESA_ID_N
   SQL = SQL & " And PEDIDO.NUMR_REQ = " & txtPedido.Text
   SQL = SQL & " And CODG_PROD = '" & Trim(txtProduto.Text) & "'"

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "UPDATE PEDIDOITEM SET "
      SQL = SQL & " VlrBaseIcms = " & tpMOEDA(dblValorBaseICMS)
      SQL = SQL & ", PERCICMS = " & tpMOEDA(dblPercentualICMS)
      SQL = SQL & ", VlrIcms = " & tpMOEDA(dblValorICMS)
      SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(dblValorBaseICMSSubst)
      SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(dblPercentualICMSSubst)
      SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(dblValorICMSSubst)
      SQL = SQL & ", cfop = '" & strCFOP & "'"
      SQL = SQL & ", STRIBUTARIA = '" & SITUAÇÃO_TRIBUTARIA_PRODUTO & "'"
      SQL = SQL & ", status = 'P'"
      SQL = SQL & " Where numr_req = " & txtPedido.Text
      SQL = SQL & " and pedido_id = " & TabTemp.Fields("pedido_id").Value
      SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREPARA_TRIBUTAÇÃO_PRODUTO"
   Exit Sub
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

   PEGA_DADOS_EMPRESA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_VENDA"
End Sub

Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

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

   txtAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
   txtVarejo.Text = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
   STATUS_PROD = TabProduto!SITUACAO

   stBarReq.Panels(2).Text = Format(TabProduto!Qtde, strFormatacao3Digitos)

   QTDE_ESTOQUE = TabProduto!Qtde

   If Not IsNull(TabProduto!PRECO_VENDA) Then
      stBarReq.Panels(4).Text = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)

      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_VENDA
      txtValor_Unitario.Text = Format(Valr_Venda_Produto_n, strFormatacao2Digitos)
      txtPreçoCusto.Text = 0 & Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)

      txtValorDolar.Text = Format(CONVERTE_DOLAR(Valr_Venda_Produto_n), strFormatacao2Digitos)

      VLR_ANTERIOR_N = TabProduto!PRECO_VENDA
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   PRECO_PROD = 0 & txtAtacado.Text

   If txtPedido.Text = "" Or Trim(txtProduto.Text) = "" Then _
      Exit Sub

CHECA_QTDE_ATUAL_ESTOQUE_PRODUTO

stBarReq.Panels(2).Text = Format(QTDE_ESTOQUE, strFormatacao3Digitos)

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

      txtQTDE.Text = Format(TabPedidoItem!QTD_PEDIDA, strFormatacao3Digitos)

      QTDE_PEDIDO = TabPedidoItem!QTD_PEDIDA

      VALOR_ITEM_N = TabPedidoItem!Valor_Item
      VALOR_DIFERENCA_N = TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA

      SQL = "UPDATE Produto SET qtde_retido = qtde_retido - " & tpMOEDA(QTDE_PEDIDO)
      SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N & " and codg_produto = '" & txtProduto.Text & "' and qtde_retido >= " & tpMOEDA(QTDE_PEDIDO)
      CONECTA_RETAGUARDA.Execute SQL

      stBarReq.Panels(2).Text = "" & Format(TabProduto!Qtde, strFormatacao3Digitos)
      QTDE_ESTOQUE = TabProduto!Qtde
      txtSeq.Text = "" & TabPedidoItem.Fields("seq_id").Value
   End If

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   If Len(CODIGO_BARRAS) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtValor_Unitario.Text) <> "" Then
            If IsNumeric(txtValor_Unitario.Text) Then
               txtQTDE.Text = Format(QTDE_N / 1000, strFormatacao3Digitos)

               Call txtQtde_LostFocus

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
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
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

   SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_registro in ('D') "
   SQL = SQL & " and PEDIDO.status < 3 "

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         QTDE_ESTOQUE = TabConsulta.Fields("qtde").Value - _
                        TabConsulta.Fields("qtde_retido").Value - _
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
