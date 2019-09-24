VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNFEDIVERSAS 
   Caption         =   "NFe Diversas"
   ClientHeight    =   7950
   ClientLeft      =   2085
   ClientTop       =   2475
   ClientWidth     =   11040
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
   Icon            =   "NFEDIVERSAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11040
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
      Top             =   2280
      Width           =   10935
      Begin VB.TextBox txtValor_Unitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   8
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   375
         Left            =   4560
         MaxLength       =   29
         TabIndex        =   22
         Top             =   240
         Width           =   6255
      End
      Begin VB.TextBox txtQTDE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9480
         TabIndex        =   9
         ToolTipText     =   "Informe a quantidade de venda deste produto."
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAtacado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtVarejo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3480
         Picture         =   "NFEDIVERSAS.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtValorDolar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   7
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtPreçoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   18
         ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdMata 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         Picture         =   "NFEDIVERSAS.frx":6614
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Unitário"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7440
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9600
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Atacado"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Varejo"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   720
         Width           =   975
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
      Height          =   1695
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   10935
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   39
         Text            =   "55"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtCFOP 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optTransf 
         Caption         =   "Transferência"
         Height          =   240
         Left            =   6000
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optDC 
         Caption         =   "Devolução Compra"
         Height          =   240
         Left            =   -120
         TabIndex        =   37
         Top             =   -3120
         Width           =   2175
      End
      Begin VB.OptionButton optDV 
         Caption         =   "Devolução Venda"
         Height          =   240
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbCFOPAux 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4200
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbCFOP 
         Height          =   360
         Left            =   4140
         TabIndex        =   2
         Text            =   "-- Selecione --"
         ToolTipText     =   "CFOP"
         Top             =   720
         Width           =   4335
      End
      Begin VB.CommandButton cmdConsCli 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3080
         Picture         =   "NFEDIVERSAS.frx":7455
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1200
         Width           =   405
      End
      Begin VB.TextBox txtNF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3540
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1200
         Width           =   7215
      End
      Begin MSMask.MaskEdBox txtDtEmis 
         Height          =   360
         Left            =   9360
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483637
         Enabled         =   0   'False
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
         TabIndex        =   4
         ToolTipText     =   "Informe o CNPJ/CPF/Código do cliente, F7-Consultar"
         Top             =   1200
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
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   0
         X2              =   10920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   240
         Left            =   2520
         TabIndex        =   34
         Top             =   780
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NºNFe:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   330
         TabIndex        =   14
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt.NFe:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   8565
         TabIndex        =   13
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblCliFor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Cliente:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   1200
         Width           =   810
      End
   End
   Begin ComctlLib.StatusBar stBarReq 
      Height          =   375
      Left            =   0
      TabIndex        =   28
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
      TabIndex        =   29
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
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
         NumButtons      =   4
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
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   10200
         TabIndex        =   32
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
               Picture         =   "NFEDIVERSAS.frx":7E57
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":8FF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":A080
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":B035
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":C140
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":D296
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":D6E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":F55F
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEDIVERSAS.frx":10C15
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
         TabIndex        =   30
         Top             =   0
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lstNfItem 
      Height          =   3705
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Clique para selecionar um produto ja gravado."
      Top             =   3840
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   6535
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
      DesignWidth     =   11040
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
      TabIndex        =   31
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmNFEDIVERSAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim strInscEstadual_A      As String
   Dim dblTipoCliente         As Double
   Dim strCPFCNPJ             As String
   Dim rstEmpresa             As New ADODB.Recordset
   Dim Seq_N                  As Long
   Dim PRECO_PROD             As Double
   Dim CLIENTE_ID_N           As Long
   Dim TIPO_NOTA_A            As String
   Dim strCFOP                As String
   Dim SITUAÇÃO_TRIBUTARIA_PRODUTO
   Dim Valr_Venda_Produto_n   As Double

   'Private CalculaIcmsG As New MegasimCL.mCalculaIcms ' Yuri alterado em 01/05/2012

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   If optDC.Value = True Then
      lblCliFor.Caption = "Fornec.:"
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   End If
   If optDV.Value = True Then
      lblCliFor.Caption = "Cliente:"
      TIPO_PESSOA_CADASTRO = "CLIENTE"
   End If

   INICIA_TELA
   CARREGA_CFOP

   Call TXTNF_LostFocus

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
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaCadastro.Show 1

         If NOME_A <> "" Then _
            txtNome.Text = NOME_A
         NOME_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub cmbCFOP_GotFocus()
   cmbCFOP.SelStart = 0
   cmbCFOP.SelLength = Len(cmbCFOP)
   cmbCFOP.BackColor = &HC0FFFF
End Sub

Private Sub cmbcfop_LostFocus()
   cmbCFOP.BackColor = &HFFFFFF
End Sub

Private Sub cmbCFOP_Click()
On Error Resume Next

   cmbCFOPAux.ListIndex = cmbCFOP.ListIndex
   txtCNPJCPF.SetFocus

End Sub

Private Sub cmbCFOP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfop_KeyPress"
End Sub

Private Sub lstNfItem_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(lstNfItem.SelectedItem.Text) Then
             txtProduto.Text = lstNfItem.SelectedItem.Text
             txtSeq.Text = Trim(lstNfItem.SelectedItem.ListSubItems.item(7).Text)
         End If

         EXCLUIR_ITEM
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstNfItem_KeyDown"
End Sub

Private Sub optDV_Click()
   If optDC.Value = True Then
      lblCliFor.Caption = "Fornec.:"
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   End If
   If optDV.Value = True Then
      lblCliFor.Caption = "Cliente:"
      TIPO_PESSOA_CADASTRO = "CLIENTE"
   End If
   txtCFOP.SetFocus
End Sub

Private Sub optDC_Click()
   If optDC.Value = True Then
      lblCliFor.Caption = "Fornec.:"
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   End If
   If optDV.Value = True Then
      lblCliFor.Caption = "Cliente:"
      TIPO_PESSOA_CADASTRO = "CLIENTE"
   End If
   txtCFOP.SetFocus
End Sub

Private Sub optTransf_Click()
   If optDC.Value = True Then
      lblCliFor.Caption = "Fornec.:"
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   End If
   If optDV.Value = True Then
      lblCliFor.Caption = "Cliente:"
      TIPO_PESSOA_CADASTRO = "CLIENTE"
   End If
   txtCFOP.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         MsgBox "Processo realizado com sucesso !!! Vá em Recebimento para enviar nota para o SEFAZ !!!"
         LIMPA_TUDO
         txtCFOP.SetFocus
      Case "consultar"
         CRITERIO_A = ""
         CNPJCPF_A = ""
         frmNFECONSULTA.Show 1
         If NF_ID_N > 0 Then
            'LIMPA_TUDO
            CRITERIO_A = ""
            'mostra_nfe
         End If
      Case "limpar"
         LIMPA_TUDO
         txtNF.Enabled = True
         txtNF.SetFocus
      Case "voltar"
         Unload Me
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

   If optDC.Value = True Then
      lblCliFor.Caption = "Fornec.:"
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   End If
   If optDV.Value = True Then
      lblCliFor.Caption = "Cliente:"
      TIPO_PESSOA_CADASTRO = "CLIENTE"
   End If

   frmPessoaConsulta.Show 1
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

   EXCLUIR_ITEM

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMata_Click"
End Sub

Private Sub lstNfItem_DblClick()
'On Error GoTo ERRO_TRATA

    If Not IsNull(lstNfItem.SelectedItem.Text) Then
        txtProduto.Text = lstNfItem.SelectedItem.Text
        txtSeq.Text = Trim(lstNfItem.SelectedItem.ListSubItems.item(7).Text)
        txtProduto.SetFocus
    End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstNfItem_DblClick"
End Sub

Private Sub lstNfItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstNfItem, ColumnHeader
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

Private Sub TXTPRODUTO_LostFocus()
   If STATUS_PROD = "P" Then
      MsgBox "Produto em Promoçao, Impossivel fazer devolução."
      txtProduto.Text = ""
      txtProduto.SetFocus
   End If
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

Private Sub txtValorDolar_GotFocus()
   txtValorDolar.SelStart = 0
   txtValorDolar.SelLength = Len(txtValorDolar)
   txtValorDolar.BackColor = &HC0FFFF
End Sub

Private Sub txtValorDolar_LostFocus()
   txtValorDolar.BackColor = &HFFFFFF
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

Private Sub txtNome_LostFocus()
'On Error GoTo ERRO_TRATA

   txtNome.Text = UCase(txtNome.Text)
   txtNome.Enabled = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_LostFocus"
End Sub
'==================CNPJCPF
Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "F7-Consulta Clientes", "Inform Cliente", "", ""
   
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Mask = "###############"

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF)
   txtCNPJCPF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
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

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) = "" Then
         MsgBox "Informe cliente para devolução."
         Exit Sub
      End If
      txtCNPJCPF.PromptInclude = True

      txtProduto.SetFocus
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
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then _
      If TRATA_PESSOA(txtCNPJCPF.Text) = False Then _
         txtCNPJCPF.SetFocus
   txtCNPJCPF.PromptInclude = True
   txtNome.Text = "" & NOME_CLIENTE_A
   txtCNPJCPF.BackColor = &HFFFFFF
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
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
   MOSTRA_TOP "ESC-SAIR", "F7-Consulta Produtos", "Delete-Excluir Produto", "", ""

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
         EXCLUIR_ITEM
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
      CODIGO_BARRAS_A = ""

      If Trim(txtProduto.Text) <> "" Then
         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select * from PRODUTO With (NOLOCK)"
         SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabProduto.EOF Then
            CODIGO_BARRAS_A = "" & Trim(txtProduto.Text)
            QTDE_N = 0

            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select * from PRODUTO With (NOLOCK)"
            SQL = SQL & " where CODG_barra = '" & Trim(txtProduto.Text) & "'"
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabProduto.EOF Then
               If Len(CODIGO_BARRAS_A) = 13 Then
                  '2 = produtos "in store" (sempre será 2)
                  'C = código do produto (4,5 ou 6 dígitos)
                  'T = total a pagar (sempre 6 dígitos)
                  'P = peso (sempre 5 dígitos)
                  'Q = quantidade (sempre 5 dígitos)
                  '0 = zero fixo
                  'DV = dígito verificador do EAN-13

                  txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, 2, 6))
                  CONT_N = 0 & Int(Mid(CODIGO_BARRAS_A, 2, 2))

                  If CONT_N > 9 Then
                     txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, 4, TamanhoCodgProdBarra_N))
                     Else:
                        If CONT_N > 0 Then
                           txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, 3, TamanhoCodgProdBarra_N + 1))
                           Else
                              txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, 3, TamanhoCodgProdBarra_N))
                        End If
                  End If

                  QTDE_N = 0 & Int(Mid(CODIGO_BARRAS_A, 8, 5))   'gramas

                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select * from PRODUTO With (NOLOCK)"
                  SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
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

   MOSTRA_TOP "ESC-SAIR", "Informe a quantidade", "", "", ""
   
   If Trim(txtProduto.Text) = Empty Then
      MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro."
      txtProduto.Text = 99999999
      txtProduto.SetFocus
      Exit Sub
   End If
   If txtQTDE.Text <> "" Then
      txtQTDE.SelStart = 0
      txtQTDE.SelLength = Len(txtQTDE)
      txtQTDE.BackColor = &HC0FFFF
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
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
   txtQTDE.BackColor = &HFFFFFF
End Sub

Private Sub TXTNF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC-SAIR", "Tecle <ENTER> para gerar nova Pedido ou informe uma já existente", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTNF_GotFocus"
End Sub

Private Sub TXTNF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCFOP.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTNF_KeyPress"
End Sub

Private Sub TXTNF_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNF.Text) = "" Then _
      txtNF.Enabled = False

   txtNF.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTNF_LostFocus"
End Sub

Private Sub txtCFOP_GotFocus()
   txtCFOP.SelStart = 0
   txtCFOP.SelLength = Len(txtCFOP.Text)
   txtCFOP.BackColor = &HC0FFFF
End Sub

Private Sub txtCFOP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCFOP_KeyPress"
End Sub

Private Sub txtCFOP_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCFOP.BackColor = &HFFFFFF


   If Trim(txtCFOP.Text) <> "" Then
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP With (NOLOCK)"
   SQL = SQL & " where cfop_id = " & Trim(txtCFOP.Text)
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      cmbCFOPAux.Text = "" & Trim(TabDESCR!CFOP_ID)
      cmbCFOP.Text = "" & Trim(TabDESCR!CFOP_ID) & "-" & Trim(TabDESCR!DESCRICAO)
      Else
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         MsgBox "Cadastro de CFOP com problemas. Não foi localizado nenhum codigo de CFOP cadastrado", vbCritical
         Exit Sub
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close
End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCFOP_LostFocus"
End Sub

Private Sub TXTVALOR_UNITARIO_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "ESC - SAIR", "Informe Valor Unitário", "", "", ""
   
   txtValor_Unitario.SelStart = 0
   txtValor_Unitario.SelLength = Len(txtValor_Unitario.Text)
   txtValor_Unitario.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_GotFocus"
End Sub

Private Sub TXTVALOR_UNITARIO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If TabUSU.State = 1 Then _
         TabUSU.Close

      SQL = "select * from USUARIO With (NOLOCK)"
      SQL = SQL & " where usuario_id = " & USUARIO_ID_N
      SQL = SQL & " and status = 1"
      TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabUSU.EOF Then
         If TabUSU.State = 1 Then _
            TabUSU.Close

         MsgBox "Problemas com usuário, codigo=0"
         Exit Sub
      End If

      If TabUSU.State = 1 Then _
         TabUSU.Close

      txtQTDE.SetFocus
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

   txtValor_Unitario.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_UNITARIO_LostFocus"
End Sub
'============================SUBROTINAS
Private Sub EXCLUIR_ITEM()
'On Error GoTo ERRO_TRATA

   If Trim(txtNF.Text) <> "" And Trim(txtSeq.Text) <> "" Then
      If IsNumeric(txtNF.Text) And IsNumeric(txtSeq.Text) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select nf_id,seq_id from NFITEM With (NOLOCK)"
         SQL = SQL & " where nf_id = " & NF_ID_N
         SQL = SQL & " and seq_id = " & txtSeq.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            Msg = "Deseja Excluir Esse Item ?"
            Style = vbYesNo + 32
            Title = "Atenção."
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "Delete from NFITEM "
               SQL = SQL & " where nf_id = " & NF_ID_N
               SQL = SQL & " and seq_id = " & txtSeq.Text
               CONECTA_RETAGUARDA.Execute SQL

               SETA_GRID
            End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   Valr_Venda_Produto_n = 0
   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""

   stBarReq.Panels(2).Text = ""
   stBarReq.Panels(6).Text = ""
   stBarReq.Panels(4).Text = ""

   QTDE_PEDIDO = 0
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

   If TabUSU.State = 1 Then _
      TabUSU.Close

   PRODUTO_ID_N = 0
   txtCFOP.Text = ""
   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   txtNF.Text = ""
   txtDtEmis = Format(Date, "dd/mm/yyyy")
   txtNome.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""

   LIMPA_BODY
   lstNfItem.ListItems.Clear
   stBarReq.Panels(6).Text = ""
   stBarReq.Panels(2).Text = ""
   stBarReq.Panels(10).Text = ""
   VALOR_TOTAL_N = 0
   NF_ID_N = 0
   QTDE_PEDIDO = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_TOTAL_N = 0
   USU_LIBERA_VENDA_N = 0
   INDR_RECEITA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstNfItem.ListItems.Clear
   CONT_N = 0

   If TabNFITEM.State = 1 Then _
      TabNFITEM.Close

   SQL = "select * from NFITEM WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " where nf_id = " & NF_ID_N
   TabNFITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabNFITEM.EOF
      CONT_N = CONT_N + 1
      Set item = lstNfItem.ListItems.Add(, "seq." & CONT_N, Trim(TabNFITEM!Codg_Produto))
      item.SubItems(1) = "" & Trim(TabNFITEM.Fields("DESCRICAO").Value)
      item.SubItems(2) = "" & Format(TabNFITEM.Fields("qtde").Value, strFormatacao3Digitos)
      item.SubItems(3) = "" & Format(TabNFITEM.Fields("valor").Value, strFormatacao2Digitos)
      item.SubItems(4) = "" & Format(TabNFITEM.Fields("valor").Value * TabNFITEM.Fields("qtde").Value, strFormatacao2Digitos)
      item.SubItems(5) = "" & TabNFITEM.Fields("stributaria").Value
      item.SubItems(6) = "" & TabNFITEM.Fields("codg_ncm").Value
      item.SubItems(7) = "" & TabNFITEM.Fields("seq_id").Value

      If TabNFITEM.Fields("situacao").Value = "A" Then
         item.ForeColor = vbBlue
         item.ListSubItems(1).ForeColor = vbBlue
         item.ListSubItems(2).ForeColor = vbBlue
         item.ListSubItems(3).ForeColor = vbBlue
         item.ListSubItems(4).ForeColor = vbBlue
         item.ListSubItems(5).ForeColor = vbBlue
         item.ListSubItems(6).ForeColor = vbBlue
      End If
      If TabNFITEM.Fields("situacao").Value = "P" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
         item.ListSubItems(3).ForeColor = vbRed
         item.ListSubItems(4).ForeColor = vbRed
         item.ListSubItems(5).ForeColor = vbRed
         item.ListSubItems(6).ForeColor = vbRed
      End If

      TabNFITEM.MoveNext
   Wend

   If TabNFITEM.State = 1 Then _
      TabNFITEM.Close

   MOSTRA_TOTAIS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GERA_NFe()
'On Error GoTo ERRO_TRATA

   If Trim(txtNF.Text) = "" Then _
      txtNF.Text = "" & GERA_NUMERO_NFe_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NFe"
End Sub

Public Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
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

Sub INICIA_TELA()
'On Error GoTo ERRO_TRATA
   
   Me.Caption = Me.Caption & " - " & Me.Name
   
   UF_CLIENTE_A = ""  'Variavel para tratamento Fiscal do item
   UF_EMPRESA_A = "" 'Variavel para tratamento Fiscal do item
   strInscEstadual_A = "" 'Variavel para tratamento Fiscal do item
   dblTipoCliente = -1 'Variavel para tratamento fiscal do item
   strCPFCNPJ = ""

   txtDtEmis = Format(Date, "dd/mm/yyyy")

   PEGA_DADOS_EMPRESA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIA_TELA"
End Sub

Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

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
            txtProduto.SetFocus
            Exit Sub
            Else: txtDescricao.Text = Trim(TabProduto!DESCRICAO)
         End If
   End If

   txtAtacado.Text = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
   txtVarejo.Text = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
   STATUS_PROD = TabProduto!SITUACAO

   If Not IsNull(TabProduto!PRECO_Venda) Then
      stBarReq.Panels(4).Text = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)

      Valr_Venda_Produto_n = 0 & TabProduto!PRECO_Venda
      txtValor_Unitario.Text = Format(Valr_Venda_Produto_n, strFormatacao2Digitos)
      txtPreçoCusto.Text = "" & Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)

      VLR_ANTERIOR_N = TabProduto!PRECO_Venda
      If VLR_ANTERIOR_N < 0 Then
         MsgBox "Valor do produto invalido !!!"
         Exit Sub
      End If
   End If

   PRECO_PROD = 0 & txtAtacado.Text

   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

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

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If TabNFITEM.State = 1 Then _
      TabNFITEM.Close

   If Len(CODIGO_BARRAS_A) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtValor_Unitario.Text) <> "" Then
            If IsNumeric(txtValor_Unitario.Text) Then
               txtQTDE.Text = Format(QTDE_N / 1000, strFormatacao3Digitos)

               Call txtQtde_LostFocus

               CODIGO_BARRAS_A = ""
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

Sub CARREGA_CFOP()
'On Error GoTo ERRO_TRATA

   'CFOP
   cmbCFOPAux.Clear
   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP With (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      Do Until TabDESCR.EOF
         DoEvents
         cmbCFOPAux.AddItem Trim(TabDESCR!CFOP_ID)
         cmbCFOP.AddItem Trim(TabDESCR!CFOP_ID) & "-" & Trim(TabDESCR!DESCRICAO)
         TabDESCR.MoveNext
      Loop
      Else
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         MsgBox "Cadastro de CFOP com problemas. Não foi localizado nenhum codigo de CFOP cadastrado", vbCritical
         Exit Sub
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_CFOP"
End Sub

'Redução da base de cálculo conforme Lei 12.431/2003.

'- Transferência de mercadoria fora do estado: CFOP 6152; CST ICMS: 300; base de cálculo ICMS: 100%; alíquota ICMS: 4%; CST PIS: 08; base de cálculo PIS: 0,0; alíquota PIS: 0,0; CST COFINS: 08; base de cálculo COFINS: 0,0; alíquota COFINS: 0,0.

'- Transferência de mercadoria p/ Fortaleza: CFOP 6152; CST ICMS: 300; base de cálculo ICMS: 100%; alíquota ICMS: 7%; CST PIS: 08; base de cálculo PIS: 0,0; alíquota PIS: 0,0; CST COFINS: 08; base de cálculo COFINS: 0,0; alíquota COFINS: 0,0.


Private Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
   CONT_N = 0

   stBarReq.Panels(6).Text = ""
   stBarReq.Panels(8).Text = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'BUSCA VALOR TOTAL VENDA
   SQL = "select sum(valor*qtde) from NFITEM With (NOLOCK)"
   SQL = SQL & " where nf_id = " & NF_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_ITEM_N = TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select count(produto_id) from NFITEM With (NOLOCK)"
   SQL = SQL & " where nf_id = " & NF_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         CONT_N = 0 & TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   stBarReq.Panels(8).Text = CONT_N
   stBarReq.Refresh

   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
   stBarReq.Panels(10).Text = Format(VALOR_TOTAL_N, "##,##0.00")



'SET
'         SQL = SQL & "," & Replace(Qtd_Volume_N, ",", ".")     'Qtd_Volume
'         SQL = SQL & "," & Replace(PESO_BRUTO_A, ",", ".")     'Peso_Bruto
'         SQL = SQL & "," & Replace(PESO_LIQUI_A, ",", ".")     'Peso_Liquido


Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAIS"
End Sub

Sub PROCESSA_ITEM()
'On Error GoTo ERRO_TRATA

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

   If Trim(txtQTDE.Text) = "" Then
      Beep
      MsgBox "Informe a quantidade.", vbOKOnly, "Atenção."
      txtQTDE.SetFocus
      Exit Sub
      Else
         'quantidade pedida
         QTDE_PEDIDO = txtQTDE.Text
         txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
   End If

   'valor venda item
   VALOR_ITEM_N = txtValor_Unitario.Text
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_TOTAL_N = VALOR_TOTAL_N + (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N

   GERA_NFe
   GRAVA_CABECA
   GRAVA_TUDO_ITEM

   txtCNPJCPF.PromptInclude = False

   If Trim(UF_CLIENTE_A) = "" Then
      MsgBox "Cliente com cadastro incompleto !!!"
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
End Sub

Private Sub GRAVA_CABECA()
'On Error GoTo ERRO_TRATA

   Dim TIPO_REGISTRO_A  As String
   Dim PESO_BRUTO_N     As Double
   Dim PESO_LIQUI_N     As Double
   Dim Qtd_Volume_N     As Double
   Dim INDDEST_N        As Integer
   Dim TRANSP_ID_N      As Long

   If optDC.Value = True Then _
      TIPO_REGISTRO_A = "DC"
   If optDV.Value = True Then _
      TIPO_REGISTRO_A = "DV"
   If optTransf.Value = True Then _
      TIPO_REGISTRO_A = "TF"

   CRITERIO_A = ""
   PESO_BRUTO_N = 0
   PESO_LIQUI_N = 0
   Qtd_Volume_N = 0
   INDDEST_N = 1

   If PESSOA_ID_N <= 0 Then
      txtCNPJCPF.PromptInclude = False
         TRATA_PESSOA txtCNPJCPF.Text
      txtCNPJCPF.PromptInclude = True
   End If

   If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
      INDDEST_N = 1
      Else: INDDEST_N = 2
   End If

   TRANSP_ID_N = 0 & TRAZ_ID_TABELA("vwTRANSPORTADORA", "transp_id", "cnpjcpf", CNPJ_EMPRESA_N)

   If TabNF.State = 1 Then _
      TabNF.Close

   SQL = "select nf_id from NF With (NOLOCK)"
   SQL = SQL & " where numr_nota = " & txtNF.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and serie_nota = " & txtSerie.Text
   TabNF.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabNF.EOF Then
      NF_ID_N = MAX_ID("nf_id", "nf", "", "", "", "")
      Acao_N = 1
      Else
         NF_ID_N = 0 & TabNF.Fields("nf_id").Value
         Acao_N = 2
   End If
   If TabNF.State = 1 Then _
      TabNF.Close

   SQL = "spNF " & Acao_N & "," & _
                  NF_ID_N & "," & _
                  PESSOA_ID_N & "," & _
                  TRANSP_ID_N & ",'" & _
                  Trim(TIPO_REGISTRO_A) & "'," & _
                  Trim(txtNF.Text) & ",'" & _
                  Trim(txtSerie.Text) & "','" & _
                  Now & "','" & _
                  Now & "','" & _
                  "A" & "'," & _
                  "NULL" & ",'" & _
                  Replace(Qtd_Volume_N, ",", ".") & "'," & _
                  0 & "," & _
                  0 & "," & _
                  0 & "," & _
                  ESTABELECIMENTO_ID_N & "," & _
                  1 & "," & _
                  1 & ",'" & _
                  "NFE" & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA"
End Sub

Private Sub GRAVA_TUDO_ITEM()
'On Error GoTo ERRO_TRATA

   If Trim(txtPreçoCusto.Text) = "" Then _
      txtPreçoCusto.Text = 0
   If Not IsNumeric(txtPreçoCusto.Text) Then _
      txtPreçoCusto.Text = 0

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "NFITEM", "nf_id", Str(NF_ID_N), "", "")
      Else: SEQ_ID_N = 0 & txtSeq.Text
   End If
'=====================

   If Trim(txtQTDE.Text) = "" Then
      MsgBox "Informar a quantidade !!!"
      txtQTDE.SetFocus
      Exit Sub
   End If
   If Trim(txtValor_Unitario.Text) = "" Then
      MsgBox "Informar valor !!!"
      txtValor_Unitario.SetFocus
      Exit Sub
   End If

   QTDE_PEDIDO = 0 & txtQTDE.Text
   VALOR_ITEM_N = 0 & txtValor_Unitario.Text

   If TabNFITEM.State = 1 Then _
      TabNFITEM.Close

   SQL = "select * from NFITEM With (NOLOCK)"
   SQL = SQL & " where nf_id = " & NF_ID_N
   SQL = SQL & " and seq_id = " & SEQ_ID_N
   TabNFITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabNFITEM.EOF Then
      Acao_N = 1
      Else
         Acao_N = 2
         NF_ID_N = 0 & TabNFITEM.Fields("nf_id").Value
   End If
   If TabNFITEM.State = 1 Then _
      TabNFITEM.Close

   spNFITEM Acao_N, NF_ID_N, SEQ_ID_N, PRODUTO_ID_N, VALOR_ITEM_N, 0, QTDE_PEDIDO, Trim(txtCFOP.Text), "00", 0, 0, 0, 0, 0, 0, 0, 0, 0

   'Tratamento da tributacao
   txtCNPJCPF.PromptInclude = False
   PREPARA_TRIBUTACAO_PRODUTO Trim(txtCNPJCPF.Text), tpMOEDA(VALOR_ITEM_N), tpMOEDA(QTDE_PEDIDO)

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO_ITEM"
End Sub
