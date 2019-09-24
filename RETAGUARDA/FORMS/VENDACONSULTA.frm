VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedidoConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Venda"
   ClientHeight    =   7305
   ClientLeft      =   1845
   ClientTop       =   2565
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VENDACONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtDescProd 
      Enabled         =   0   'False
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
      Left            =   3540
      MaxLength       =   100
      TabIndex        =   31
      Top             =   2400
      Width           =   4905
   End
   Begin VB.ComboBox cmbAuxSituacao 
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
      Left            =   9960
      TabIndex        =   30
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtCupom 
      Alignment       =   2  'Center
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
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtNOTA 
      Alignment       =   2  'Center
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
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbAuxVend 
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
      Left            =   1680
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbSITUACAO 
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
      Left            =   9960
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Emissão"
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   8550
      TabIndex        =   23
      Top             =   1270
      Width           =   3255
      Begin Threed.SSOption optSintetico 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Sintético"
         Value           =   -1
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optAnalitico 
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Analítico"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final"
         Height          =   240
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtReq 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtCli 
      Enabled         =   0   'False
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
      Left            =   6000
      MaxLength       =   100
      TabIndex        =   20
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   18
      Top             =   6900
      Width           =   1335
   End
   Begin VB.TextBox txtTotalVenda 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3690
      TabIndex        =   16
      Top             =   6900
      Width           =   2055
   End
   Begin VB.ComboBox cmbAuxForma 
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
      Left            =   1680
      TabIndex        =   14
      Top             =   1380
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbForma 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.ComboBox cmbVend 
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   3585
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
      ButtonWidth     =   3175
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir Tela"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Pedido"
            Key             =   "pedido"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Emissão NFe"
            Key             =   "caixa"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8760
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACONSULTA.frx":47C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACONSULTA.frx":48DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACONSULTA.frx":49E73
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACONSULTA.frx":4AE28
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACONSULTA.frx":4BF33
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11910
      DesignHeight    =   7305
   End
   Begin MSComctlLib.ListView lstPedidoItem 
      Height          =   2625
      Left            =   30
      TabIndex        =   33
      Top             =   4200
      Visible         =   0   'False
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   4630
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2252
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   7508
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "QTD."
         Object.Width           =   1668
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valr.Unitário"
         Object.Width           =   3003
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
         Object.Width           =   3003
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Req."
         Object.Width           =   1502
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   3735
      Left            =   0
      TabIndex        =   34
      Top             =   3000
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   6588
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      Appearance      =   1
      MousePointer    =   99
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cupom"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NFe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Vlr.Venda"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Vlr.Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dt.Emisão"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Faturamento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Vendedor"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ID"
         Object.Width           =   176
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   240
      Left            =   720
      TabIndex        =   32
      Top             =   2445
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Cupom:"
      Height          =   255
      Left            =   5520
      TabIndex        =   29
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N.F.:"
      Height          =   255
      Left            =   6120
      TabIndex        =   28
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   255
      Left            =   9000
      TabIndex        =   26
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   240
      Left            =   3060
      TabIndex        =   22
      Top             =   885
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Registros:"
      Height          =   240
      Left            =   6840
      TabIndex        =   19
      Top             =   6960
      Width           =   1470
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vendido no Período:"
      Height          =   240
      Left            =   1080
      TabIndex        =   17
      Top             =   6960
      Width           =   2505
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento:"
      Height          =   240
      Left            =   300
      TabIndex        =   15
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor(a):"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmPedidoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmPedidoConsulta)
   frmPedidoConsulta.Top = 570

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Activate()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_N = 0
   cmbForma.Clear
   cmbAuxForma.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbForma.AddItem TabTemp!Descricao & " - " & TabTemp!TIPOVENDA_ID
      cmbAuxForma.AddItem TabTemp!TIPOVENDA_ID
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbVend.Clear
   SQL = "select vendedor_id,nome_vend from VENDEDOR "
   SQL = SQL & " order by nome_vend "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbVend.AddItem Trim(TabTemp!NOME_VEND) & " - " & Trim(TabTemp!VENDEDOR_ID)
      cmbAuxVend.AddItem Trim(TabTemp!VENDEDOR_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbSITUACAO.AddItem "Emitido"
   cmbAuxSituacao.AddItem "3457"

'7 = cupom

   cmbSITUACAO.AddItem "Cancelado"
   cmbAuxSituacao.AddItem "9"

   cmbVend.Enabled = False
   Toolbar1.Buttons(11).Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then
      cmbVend.Enabled = True
      Toolbar1.Buttons(11).Visible = True
      Else
         SQL = "select logon from USUARIO "
         SQL = SQL & " where usuario_id = " & CODG_USU_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         If TabUSU.State = 1 Then TabUSU.Close
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            CRITERIO = Chr$(39) & Trim(TabUSU.Fields("logon").Value) & "%" & Chr(39)

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            SQL = "select nome_vend, vendedor_id from VENDEDOR "
            SQL = SQL & " where nome_vend like " & CRITERIO
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVend.Text = TabVENDEDOR!NOME_VEND
               cmbAuxVend.Text = TabVENDEDOR!VENDEDOR_ID
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Activate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lstPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            If Trim(lstPedido.SelectedItem.Text) <> "" Then
               If IsNumeric(lstPedido.SelectedItem.Text) Then
                  CRITERIO = ""
                  frmPedidoCancela.txtPedido.Text = 0 & lstPedido.SelectedItem.Text
                  frmPedidoCancela.Show 1
                  CRITERIO = ""
               End If
            End If
         End If
      Case vbKeyF7
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            'If TabTemp.State = 1 Then _
               TabTemp.Close

            'SQL = "select cgccpf from PEDIDO "
            'SQL = SQL & " where numr_req = " & lstPedido.SelectedItem.Text
            'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            'TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            'If Not TabTemp.EOF Then
            '   If Not IsNull(TabTemp!CGCCPF) Then
            '      txtCNPJCPF.PromptInclude = False
            '         txtCNPJCPF.Text = TabTemp!CGCCPF
            '      txtCNPJCPF.PromptInclude = True
            '   End If
            'End If

            'If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.ListItems.Clear

            SQL = "select i.codg_prod,i.qtd_pedida,i.valor_item,"
            SQL = SQL & " p.descricao, p.situacao, i.seq_id"
            SQL = SQL & " FROM PEDIDOITEM i, PRODUTO p "
            SQL = SQL & " where i.codg_prod = p.CODG_PRODUTO "
            SQL = SQL & " and i.PEDIDO_ID = " & lstPedido.SelectedItem.ListSubItems.Item(11).Text
            SQL = SQL & " and p.empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and I.tipo_reg = 'PC' "
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               MOSTRA_TOP "Duplo Click no grid ocultar", " ", " ", " ", ""
               lstPedidoItem.Visible = True
            End If
            While Not TabTemp.EOF
               Set Item = lstPedidoItem.ListItems.Add(, "seq." & TabTemp.Fields("seq_id"), Trim(TabTemp.Fields("codg_prod").Value))
               Item.SubItems(1) = Trim(TabTemp.Fields("descricao").Value)
               Item.SubItems(2) = Trim(TabTemp.Fields("qtd_pedida").Value)
               Item.SubItems(3) = Format(Trim(TabTemp.Fields("valor_item").Value), strFormatacao2Digitos)
               Item.SubItems(6) = lstPedido.SelectedItem.Text

               If TabTemp.Fields("situacao").Value = "A" Then
                  Item.ForeColor = vbBlue
                  Item.ListSubItems(1).ForeColor = vbBlue
                  Item.ListSubItems(2).ForeColor = vbBlue
                  Item.ListSubItems(3).ForeColor = vbBlue
                  Item.ListSubItems(4).ForeColor = vbBlue
                  Item.ListSubItems(5).ForeColor = vbBlue
                  Item.ListSubItems(6).ForeColor = vbBlue
               End If
               If TabTemp.Fields("situacao").Value = "P" Then
                  Item.ForeColor = vbRed
                  Item.ListSubItems(1).ForeColor = vbRed
                  Item.ListSubItems(2).ForeColor = vbRed
                  Item.ListSubItems(3).ForeColor = vbRed
                  Item.ListSubItems(4).ForeColor = vbRed
                  Item.ListSubItems(5).ForeColor = vbRed
                  Item.ListSubItems(6).ForeColor = vbRed
               End If

               TabTemp.MoveNext
               CRITERIO = ""
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close

            lstPedidoItem.Refresh
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoVENDA_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmPRODUTOCONSULTA.Show 1
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

Private Sub txtReq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtReq_KeyPress"
End Sub

Private Sub lstPedido_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstPedido.SelectedItem.Text) Then
      CRITERIO = lstPedido.SelectedItem.Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_DblClick"
End Sub

Private Sub lstPedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedido, ColumnHeader
End Sub

Private Sub lstPedido_Click()
'On Error GoTo ERRO_TRATA

   NUMR_REQ_N = 0
   If Not IsNull(lstPedido.SelectedItem.Text) Then _
      If IsNumeric(lstPedido.SelectedItem.Text) Then _
         NUMR_REQ_N = lstPedido.SelectedItem.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedido_Click"
End Sub

Private Sub lstPedidoitem_DblClick()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstPedidoitem_DblClick"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "caixa"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False

         CRITERIO = ""
         Indr_Consulta = True
         If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then _
            If Not IsNull(lstPedido.SelectedItem.Text) Then _
               If IsNumeric(lstPedido.SelectedItem.Text) Then _
                  GERA_NOTA
         CRITERIO = ""
         Indr_Consulta = False
      Case "consultar"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False
         CONSULTA_TUDO
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
      Case "pedido"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False
         FORMULA_REL = ""
         If Not IsNull(lstPedido.SelectedItem.Text) Then
            FORMULA_REL = lstPedido.SelectedItem.Text

            If Not IsNumeric(FORMULA_REL) Then _
               Exit Sub

            NUMR_REQ_N = FORMULA_REL

            FORMULA_REL = "{vwRelVenda.empresa_id} = " & EMPRESA_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & NUMR_REQ_N

ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_pedido_venda.rpt"
            frmRELATORIO10.Show 1
         End If
      Case "print"
         lstPedidoItem.ListItems.Clear
         lstPedidoItem.Visible = False
         FORMULA_REL = "{vwRelVenda.pedido_id} > 0 "

         txtDtIni.PromptInclude = False
         txtDtFim.PromptInclude = False

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then

            FORMULA_REL = "{vwRelVenda.dt_req} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.dt_req} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

         End If

         txtDtIni.PromptInclude = True
         txtDtFim.PromptInclude = True

         If Trim(cmbSITUACAO.Text) = "Emitido" Then
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.status} >= 3 "
            FORMULA_REL = FORMULA_REL & " and {vwRelVenda.status} <= 7 "
            Else
               If Trim(cmbSITUACAO.Text) = "Cancelado" Then _
                  FORMULA_REL = FORMULA_REL & " and {vwRelVenda.status} = 9 "
         End If

ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
'set aqui
         If optSintetico.Value = True Then
            Nome_Relatorio = "rel_venda_sintetico.rpt"
            Else: Nome_Relatorio = "rel_venda_analitico.rpt"
         End If

         frmRELATORIO10.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxForma.ListIndex = cmbForma.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFORMA_Click"
End Sub

Private Sub cmbSituacao_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxSituacao.ListIndex = cmbSITUACAO.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbsituacao_Click"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxVend.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_Click"
End Sub

Private Sub TXTCGCCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""

   If CPF_N <> "" Then
      txtCGCCPF.Text = CPF_N
      CPF_N = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_GotFocus"
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYCLIENTE.Show 1
         If CPF_N <> "" Then _
            txtCGCCPF.Text = CPF_N
         CPF_N = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_KeyDown"
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCGCCPF.Text = "" Then _
         txtCGCCPF.Text = "99999999999"

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE "
      SQL = SQL & " where CGCCPF = '" & txtCGCCPF.Text & "'"
      'SQL = SQL & " and status = 'A'"

      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCGCCPF.SetFocus
         Exit Sub
         Else: If TabCliente!NOME <> "" _
               Then txtCli.Text = TabCliente!NOME
      End If
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_KeyPress"
End Sub

Private Sub txtDTfim_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDTfim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub CONSULTA_TUDO()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_N = 0
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SQL = "SELECT PEDIDO.tipovenda_id,PEDIDO.dt_req,PEDIDO.nome_cliente,PEDIDO.cgccpf as CNPJCPF,"
   SQL = SQL & " PEDIDO.vendedor_id,PEDIDO.status as SIT_PEDIDO, PEDIDO.valor_desconto as Desc_Cabeca, "
   SQL = SQL & " PEDIDO.NUMR_REQ, PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.CODG_PROD, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
   SQL = SQL & " PEDIDOITEM.CFOP, PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.VALOR_DESCONTO, PEDIDOITEM.STATUS, PEDIDOITEM.PRECO_CUSTO, PRODUTO.DESCRICAO,"
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.SITUACAO, PRODUTO.QTDE, PRODUTO.SITUACAO_TRIBUTARIA,"
   SQL = SQL & " PRODUTO.CODG_NCM, PRODUTO.QTDE_RETIDO, PRODUTO.PRECO_CUSTO AS Preço_Custo_Produto, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, NF.NUMR_NOTA,"
   SQL = SQL & " NF.SERIE_NOTA, NF.DT_EMISSAO, NF.QTD_VOLUME, NF.DT_CANCELA, NF.PESO_BRUTO, NF.TIPO_ESPECIE, NF.PESO_LIQUIDO, "
   SQL = SQL & " CUPOM.NUMR_CUPOM, cupom.VALOR_CUPOM "

   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " LEFT OUTER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " LEFT OUTER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " LEFT OUTER JOIN NF "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = NF.PEDIDO_ID "
   SQL = SQL & " LEFT OUTER JOIN CUPOM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = CUPOM.PEDIDO_ID"

   SQL = SQL & " where PEDIDOITEM.pedido_id Is Not Null"
   SQL = SQL & " and PEDIDO.empresa_id = " & EMPRESA_ID_N

'CRITERIO = "{PEDIDO.PEDIDO_ID} > 0"

   If Trim(txtProduto.Text) <> "" Then
      SQL = SQL & " and produto.produto_id = " & PRODUTO_ID_N
'CRITERIO = CRITERIO & " and {PEDIDO.numr_req} = " & txtproduto.Text
   End If


   If Trim(txtCupom.Text) <> "" Then
      SQL = SQL & " and CUPOM.numr_cupom = " & txtCupom.Text
'CRITERIO = CRITERIO & " and {PEDIDO.numr_req} = " & txtCupom.Text
   End If

   If Trim(txtNOTA.Text) <> "" Then
      SQL = SQL & " and NF.numr_nota = " & txtNOTA.Text
'CRITERIO = CRITERIO & " and {PEDIDO.numr_doc} = " & txtNOTA.Text
   End If

   If Trim(txtReq.Text) <> "" Then
      SQL = SQL & " and PEDIDO.numr_req = " & txtReq.Text
'CRITERIO = CRITERIO & " and {PEDIDO.numr_req} = " & txtReq.Text
   End If

   txtCGCCPF.PromptInclude = False
   If Trim(txtCGCCPF.Text) <> "" Then
      SQL = SQL & " and PEDIDO.cgccpf = '" & txtCGCCPF.Text & "'"
'CRITERIO = CRITERIO & " and {PEDIDO.cgccpf} = '" & txtCGCCPF.Text & "'"
   End If
   txtCGCCPF.PromptInclude = True

   If Trim(cmbAuxVend.Text) <> "" Then
      SQL = SQL & " and PEDIDO.vendedor_id = " & cmbAuxVend.Text
'CRITERIO = CRITERIO & " and {PEDIDO.vendedor} = " & cmbAuxVend.Text
   End If

   If Trim(cmbSITUACAO.Text) = "Emitido" Then
      SQL = SQL & " and PEDIDO.status >= 3 "
      SQL = SQL & " and PEDIDO.status <= 7 "

'CRITERIO = CRITERIO & " and {PEDIDO.status} >= 3 "
'CRITERIO = CRITERIO & " and {PEDIDO.status} <= 7 "
      Else
         If Trim(cmbSITUACAO.Text) = "Cancelado" Then
            SQL = SQL & " and PEDIDO.status = 9 "
'CRITERIO = CRITERIO & " and {PEDIDO.status} = 9 "
         End If
   End If

   If Trim(cmbAuxForma.Text) <> "" Then
      SQL = SQL & " and PEDIDO.tipovenda_id = " & cmbAuxForma.Text
'CRITERIO = CRITERIO & " and {PEDIDO.tipovenda_id} = " & cmbAuxForma.Text
   End If

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and PEDIDO.dt_req >= '" & Format(txtDtIni.Text, "dd/mm/yyyy") & "'"
      SQL = SQL & " and PEDIDO.dt_req <= '" & Format(txtDtFim.Text, "dd/mm/yyyy") & "'"
'CRITERIO = CRITERIO & " and {PEDIDO.dt_req} >= DATE (" & Year(txtDtIni.Text) & "," & Month(txtDtIni.Text) & "," & Day(txtDtIni.Text) & ") "
'CRITERIO = CRITERIO & " and {PEDIDO.dt_req} <= DATE (" & Year(txtDtFim.Text) & "," & Month(txtDtFim.Text) & "," & Day(txtDtFim.Text) & ") "
   End If
   SQL = SQL & " order by PEDIDO.dt_req,PEDIDO_ID,SEQ_ID"

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_TUDO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 0
   VALOR_TOTAL_N = 0
   NUMR_SEQ_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0

   lstPedido.Visible = False
   lstPedido.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then _
      MsgBox "Não existe dados para essa consulta."
   While Not TabTemp.EOF
      DoEvents
      If PEDIDO_ID_N <> TabTemp.Fields("pedido_id").Value Then
         PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value

         NUMR_SEQ_N = NUMR_SEQ_N + 1
         Set Item = lstPedido.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp.Fields("NUMR_REQ").Value)

         Item.SubItems(11) = "" & TabTemp.Fields("PEDIDO_ID").Value

         Item.SubItems(1) = "" & TabTemp.Fields("numr_cupom").Value
         Item.SubItems(2) = "" & TabTemp.Fields("numr_nota").Value

         If TabCliente.State = 1 Then _
            TabCliente.Close

         Item.SubItems(3) = "" & Trim(TabTemp!NOME_CLIENTE)

         If IsNull(TabTemp!NOME_CLIENTE) Or Trim(TabTemp!NOME_CLIENTE) = "" Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select nome from CLIENTE "
            SQL = SQL & " where cgccpf = '" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               Item.SubItems(3) = "" & TabCliente!NOME

            If TabCliente.State = 1 Then _
               TabCliente.Close
         End If

         Item.SubItems(7) = TabTemp!DT_REQ

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         Item.SubItems(8) = ""

         SQL = "select * from TIPOVENDA "
         SQL = SQL & " where tipovenda_id = " & TabTemp!TIPOVENDA_ID
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            Item.SubItems(8) = TabDESCR!Descricao

         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         If TabUSU.State = 1 Then _
            TabUSU.Close

         Item.SubItems(9) = ""

         SQL = "select * from VENDEDOR "
         SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id").Value
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then _
            Item.SubItems(9) = TabUSU!NOME_VEND

         If TabUSU.State = 1 Then _
            TabUSU.Close

         Item.SubItems(10) = ""

         If Not IsNull(TabTemp.Fields("Status")) Then
            If TabTemp.Fields("Status") = 2 Then
               If TabTemp.Fields("tipo_registro") = "O" Then
                  Item.SubItems(10) = "Orcamento"
                  Else: Item.SubItems(10) = "Pedido"
               End If
            End If
            If TabTemp.Fields("Status") = 3 Then _
               Item.SubItems(10) = "3-Faturado"
            If TabTemp.Fields("Status") = 4 Then _
               Item.SubItems(10) = "4-Cupom"
            If TabTemp.Fields("Status") = 5 Then _
               Item.SubItems(10) = "5-Faturado"
            If TabTemp.Fields("Status") = 7 Then _
               Item.SubItems(10) = "7-Cupom Fiscal"
            If TabTemp.Fields("Status") = 9 Then _
               Item.SubItems(10) = "9-Cancelado"
         End If

         VALOR_DESCONTO_N = 0
         VALOR_TOTAL_DESCONTO_N = 0

         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         SQL = SQL & " and tipo_reg = 'PC' "
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedidoItem.EOF Then _
            If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
               VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         VALOR_DESCONTO_N = VALOR_DESCONTO_N + Desc_Cabeca
         VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_N + VALOR_TOTAL_DESCONTO_N

         'BUSCA VALOR TOTAL VENDA
         VALOR_ITEM_N = 0
         SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         SQL = SQL & " and tipo_reg = 'PC' "
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
            VALOR_ITEM_N = TabPedidoItem.Fields(0).Value
         TabPedidoItem.Close

         VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N

         Item.SubItems(4) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
         Item.SubItems(5) = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
         Item.SubItems(6) = Format(VALOR_ITEM_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

         txtTotalVenda.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
         txtTotalVenda.Refresh

         CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
         txtReg.Text = CONTA_REGISTRO_N
         txtReg.Refresh

         If TabTemp.Fields("SIT_PEDIDO").Value = 1 Then
            Item.ForeColor = vbRed
            Item.ListSubItems(1).ForeColor = vbRed
            Item.ListSubItems(2).ForeColor = vbRed
            Item.ListSubItems(3).ForeColor = vbRed
            Item.ListSubItems(4).ForeColor = vbRed
            Item.ListSubItems(5).ForeColor = vbRed
            Item.ListSubItems(6).ForeColor = vbRed
            Item.ListSubItems(7).ForeColor = vbRed
            Item.ListSubItems(8).ForeColor = vbRed
            Item.SubItems(10) = "" & "Em Aberto"
         End If
         If TabTemp.Fields("SIT_PEDIDO").Value = 2 Then
            Item.ForeColor = vbBlue
            Item.ListSubItems(1).ForeColor = vbBlue
            Item.ListSubItems(2).ForeColor = vbBlue
            Item.ListSubItems(3).ForeColor = vbBlue
            Item.ListSubItems(4).ForeColor = vbBlue
            Item.ListSubItems(5).ForeColor = vbBlue
            Item.ListSubItems(6).ForeColor = vbBlue
            Item.ListSubItems(7).ForeColor = vbBlue
            Item.ListSubItems(8).ForeColor = vbBlue
            Item.ListSubItems(9).ForeColor = vbBlue
            Item.SubItems(10) = "" & "A Faturar"
         End If
         If TabTemp.Fields("SIT_PEDIDO").Value = 3 Then
            Item.ForeColor = vbBlack
            Item.ListSubItems(1).ForeColor = vbBlack
            Item.ListSubItems(2).ForeColor = vbBlack
            Item.ListSubItems(3).ForeColor = vbBlack
            Item.ListSubItems(4).ForeColor = vbBlack
            Item.ListSubItems(5).ForeColor = vbBlack
            Item.ListSubItems(6).ForeColor = vbBlack
            Item.ListSubItems(7).ForeColor = vbBlack
            Item.ListSubItems(8).ForeColor = vbBlack
            Item.ListSubItems(9).ForeColor = vbBlack
            Item.SubItems(10) = "" & "Faturado"
         End If
         If TabTemp.Fields("SIT_PEDIDO").Value = 5 Then
            Item.ForeColor = vbBlack
            Item.ListSubItems(1).ForeColor = vbBlack
            Item.ListSubItems(2).ForeColor = vbBlack
            Item.ListSubItems(3).ForeColor = vbBlack
            Item.ListSubItems(4).ForeColor = vbBlack
            Item.ListSubItems(5).ForeColor = vbBlack
            Item.ListSubItems(6).ForeColor = vbBlack
            Item.ListSubItems(7).ForeColor = vbBlack
            Item.ListSubItems(8).ForeColor = vbBlack
            Item.ListSubItems(9).ForeColor = vbBlack
            Item.SubItems(10) = "" & "Faturado"
         End If
         If TabTemp.Fields("SIT_PEDIDO").Value = 7 Then
            Item.ForeColor = vbMagenta
            Item.ListSubItems(1).ForeColor = vbMagenta
            Item.ListSubItems(2).ForeColor = vbMagenta
            Item.ListSubItems(3).ForeColor = vbMagenta
            Item.ListSubItems(4).ForeColor = vbMagenta
            Item.ListSubItems(5).ForeColor = vbMagenta
            Item.ListSubItems(6).ForeColor = vbMagenta
            Item.ListSubItems(7).ForeColor = vbMagenta
            Item.ListSubItems(8).ForeColor = vbMagenta
            Item.ListSubItems(9).ForeColor = vbMagenta
            Item.SubItems(10) = "" & "Cupom Fiscal"
         End If
         If TabTemp.Fields("SIT_PEDIDO").Value = 9 Then
            Item.ForeColor = vbWhite
            Item.ListSubItems(1).ForeColor = vbWhite
            Item.ListSubItems(2).ForeColor = vbWhite
            Item.ListSubItems(3).ForeColor = vbWhite
            Item.ListSubItems(4).ForeColor = vbWhite
            Item.ListSubItems(5).ForeColor = vbWhite
            Item.ListSubItems(6).ForeColor = vbWhite
            Item.ListSubItems(7).ForeColor = vbWhite
            Item.ListSubItems(8).ForeColor = vbWhite
            Item.ListSubItems(9).ForeColor = vbWhite
            Item.SubItems(10) = "" & "Cancelado"
         End If
      End If
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
'vbBlack  vbRed  vbGreen  vbYellow  vbBlue  vbMagenta  vbCyan  vbWhite
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   MOSTRA_TOP "Consuta Pedido Venda", "", "", "", ""
   lstPedidoItem.ListItems.Clear
   lstPedidoItem.Visible = False
   PRODUTO_ID_N = 0
   txtDescProd.Text = ""
   txtProduto.Text = ""
   txtNOTA.Text = ""
   txtCupom.Text = ""
   txtTotalVenda.Text = ""
   txtReg.Text = ""
   lstPedido.ListItems.Clear
   txtReq.Text = ""
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   txtCli.Text = ""
   If cmbVend.Enabled = True Then
      cmbVend.Text = ""
   End If
   cmbForma.Text = ""
   cmbAuxForma.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   cmbSITUACAO.Text = ""
   cmbAuxSituacao.Text = ""
   lstPedido.Visible = True
   optSintetico.Value = True
   txtReq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub GERA_NOTA()
'On Error GoTo ERRO_TRATA

   NUMR_REQ_N = lstPedido.SelectedItem.Text
   CPF_N = ""

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status, cgccpf from PEDIDO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      CPF_N = "" & Trim(TabCABECA.Fields("cgccpf").Value)
      If Not IsNull(TabCABECA!Status) Then
         If TabCABECA!Status <> "9" Then
            If Trim(CPF_N) = "99999999999" Then
               Msg = "Para geração de nota fiscal eletrônica, os dados do cliente devem ser cadastrados, deseja continuar essa operação ?"
               PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CPF_N = ""
                  frmDISPLAYCLIENTE.Show 1
                  If Trim(CPF_N) <> "" Then
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

SQL = "select nome,cgccpf from CLIENTE "
SQL = SQL & " where cgccpf = '" & Trim(CPF_N) & "'"
TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
If TabConsulta.EOF Then
   Msg = "CNPF/CPF informado não cadastrado, deseja realizar cadastro de cliente agora ?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      frmCADASTROCLIENTE.Show 1
      MsgBox "Repetir operação."
      Else
         If TabCABECA.State = 1 Then _
            TabCABECA.Close
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Exit Sub
   End If
   Else
      If TabCABECA.State = 1 Then _
         TabCABECA.Close
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
End If
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

                     SQL = "update PEDIDO set cgccpf = '" & Trim(CPF_N) & "'"
                     SQL = SQL & " where numr_req = " & NUMR_REQ_N
                     CONECTA_RETAGUARDA.Execute SQL
                     Else
                        If TabCABECA.State = 1 Then _
                           TabCABECA.Close
                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close
                        Exit Sub
                  End If
                  Else
                     If TabCABECA.State = 1 Then _
                        TabCABECA.Close
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
                     Exit Sub
               End If
            End If

            CRITERIO = NUMR_REQ_N
            If TabCABECA.State = 1 Then _
               TabCABECA.Close
            frmNOTAGERA.Show 1
         End If
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_NOTA"
End Sub

Sub PROCURA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select produto_id,descricao from PRODUTO "
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDescProd.Text = TabConsulta.Fields("descricao").Value
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
End Sub

Sub MOSTRA_TOP(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   Me.Caption = Msg1 & " | " & Msg2 & " | " & Msg3 & " | " & Msg4 & " | " & Msg5
End Sub
