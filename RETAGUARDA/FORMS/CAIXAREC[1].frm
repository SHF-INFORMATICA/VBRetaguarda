VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCAIXAREC 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recebimento Caixa "
   ClientHeight    =   7305
   ClientLeft      =   2175
   ClientTop       =   2235
   ClientWidth     =   10980
   Icon            =   "CAIXAREC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Faturamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1935
      Left            =   30
      TabIndex        =   24
      Top             =   2115
      Width           =   10935
      Begin VB.TextBox txtRecebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Height          =   375
         Left            =   8880
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtValorVenda 
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
         Height          =   375
         Left            =   1935
         TabIndex        =   30
         Text            =   "0,00"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtDesconto 
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
         Height          =   375
         Left            =   1935
         TabIndex        =   29
         Text            =   "0,00"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8895
         TabIndex        =   28
         Text            =   "0,00"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Height          =   375
         Left            =   8880
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbAuxTIPOVENDA 
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTIPOVENDA 
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
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   5355
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Informe Valor Recebido = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5955
         TabIndex        =   35
         Top             =   360
         Width           =   2940
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Venda = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   34
         Top             =   840
         Width           =   1650
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desconto = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   465
         TabIndex        =   33
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Troco = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7965
         TabIndex        =   32
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Total = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7305
         TabIndex        =   31
         Top             =   1320
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1275
      Left            =   30
      TabIndex        =   14
      Top             =   780
      Width           =   10935
      Begin VB.CommandButton cmdPesquisaCliente 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3150
         Picture         =   "CAIXAREC.frx":47C4A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Pesquisar Clientes"
         Top             =   750
         Width           =   315
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0C0C0&
         DataField       =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3600
         MaxLength       =   100
         TabIndex        =   22
         Top             =   720
         Width           =   7215
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8400
         TabIndex        =   17
         Top             =   210
         Width           =   2415
      End
      Begin VB.TextBox txtPEDIDO 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   405
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtVendedor 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   405
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   714
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   23
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblPRAZO 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6360
         TabIndex        =   21
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6690
         TabIndex        =   19
         Top             =   255
         Width           =   1665
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2265
         TabIndex        =   18
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seq.   Modalidade de Pagamento              Vlr.Pagamento    Dias       Dt.Emissão      Dt.Vencimento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   30
      TabIndex        =   10
      Top             =   4111
      Width           =   10935
      Begin VB.TextBox txtDias 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.ComboBox cmbAuxLanc 
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
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtValorItem 
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
         Height          =   405
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1755
      End
      Begin VB.ComboBox cmbMODALIDADE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   3675
      End
      Begin MSMask.MaskEdBox txtDTVENC 
         Height          =   360
         Left            =   9210
         TabIndex        =   8
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDTEMIS 
         Height          =   360
         Left            =   7560
         TabIndex        =   7
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":4864C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":48AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":48DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":49210
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":49664
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":49984
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":49DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":4A0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":4AB0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXAREC.frx":4B51C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1270
      ButtonWidth     =   3625
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Confirmar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir Seqüência"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Sequência"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9360
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXAREC.frx":4BF2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXAREC.frx":4D0B9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXAREC.frx":4E7B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXAREC.frx":4F845
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListaLanc 
      Height          =   2355
      Left            =   0
      TabIndex        =   13
      Top             =   4900
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   4154
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Doc."
         Object.Width           =   2351
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2743
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modalidade"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Dt.Lanç."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Dt.Venc."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Juros"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   6780
      Y1              =   3540
      Y2              =   3600
   End
End
Attribute VB_Name = "frmCAIXAREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_RECEBIDO_N As Double, NUMR_PARCELA As Integer
   Dim VALOR_TROCO_N As Double, VALOR_TOTAL_LANÇADO As Double
   Dim VALOR_ENTRADA As Double, PERC_JUROS_N As Double, DIAS_PRAZO As Integer
   Dim TabTipoVenda As New ADODB.Recordset, INDR_FINALIZA_RECEBIMENTO As Boolean
   Dim VALOR_DESCONTO_CABECA_N As Double
   Dim VALOR_TOTAL_DESCONTO_N As Double
   Dim VALOR_DESCONTO_ITEM_N As Double
   Dim VALOR_DIFERENCA_N As Double
   Dim VALOR_VENDA_N As Double
   Dim strCPFFunc As String
   Dim strDESCFORMAPGTO As String

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INICIALIZA_RECEBIMENTO

   'Call CentralizaJanela2(frmCAIXAREC)

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF2
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KEYDOWN"
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
      Case "limpar"
         EXCLUIR_TUDO
         SETA_GRID
         LIMPA_BODY
         VALOR_ITEM_N = 0
         VALOR_ENTRADA = 0
         cmbTIPOVENDA.Text = ""
         cmbAuxTIPOVENDA.Text = ""
         Frame1.Enabled = False
      Case "matar"
         MATA_ITEM_LANCAMENTO
      Case "voltar"
         CONFIRMAR_RECEBIMENTO_PARCELADO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdPesquisaCliente_Click()
'On Error GoTo ERRO_TRATA

    CPF_N = ""
    frmDISPLAYCLIENTE.Show 1
    If Len(CPF_N) > 0 Then
         txtCNPJCPF.Text = CPF_N
         MostraCliente
    End If
    txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdPesquisaCliente_Click"
End Sub

Private Sub cmbTIPOVENDA_LostFocus()
'On Error GoTo ERRO_TRATA

   If cmbAuxTIPOVENDA.Text <> "" Then
      cmbMODALIDADE.Clear
      cmbAuxLanc.Clear
      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " Where forma_id >= 1 "
      If TabDESCR.State = 1 Then TabDESCR.Close
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabDESCR.EOF
         cmbMODALIDADE.AddItem TabDESCR!Descricao
         cmbAuxLanc.AddItem TabDESCR!FORMA_ID
         TabDESCR.MoveNext
      Wend
      TabDESCR.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_LostFocus"
End Sub

Private Sub cmbTipoVenda_Click()
'On Error GoTo ERRO_TRATA

   lblPRAZO.Caption = ""
   cmbAuxTIPOVENDA.ListIndex = cmbTIPOVENDA.ListIndex
   VALOR_ITEM_N = 0
   VALOR_ENTRADA = 0

   EXCLUIR_TUDO

   NUMR_PARCELA = 0
   DIAS_PRAZO = 0

   If Trim(cmbAuxTIPOVENDA.Text) <> "" Then
      SQL = "select * from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
      If TabTipoVenda.State = 1 Then TabTipoVenda.Close
      TabTipoVenda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTipoVenda.EOF Then
         If Not IsNull(TabTipoVenda!parcela) Then _
            NUMR_PARCELA = TabTipoVenda!parcela

         If Not IsNull(TabTipoVenda!Prazo) Then _
            DIAS_PRAZO = TabTipoVenda!Prazo
         
         If Not IsNull(TabTipoVenda!Descricao) Then _
            strDESCFORMAPGTO = TabTipoVenda!Descricao
      End If
      Else
         MsgBox "Selecione tipo de venda."
         Exit Sub
   End If

   Frame1.Enabled = True
   If DIAS_PRAZO > 0 Or Trim(cmbAuxTIPOVENDA.Text) = "9999" Then
      'GERA TITULOS
      SQL = "select * from CABECAREQ "
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCABECA.EOF Then
         GERA_FATURAMENTO
         SETA_GRID
      End If
      TabCABECA.Close

      'Calcula

      'Quando Form A vista entende -se que ele podera escolher varias forma de pagamento a vista como
      'Cartao Credito, Cartao Debito, Cheque, Dinheiro
      If Trim(cmbAuxTIPOVENDA.Text) = "9999" Then
         CONFIRMAR_RECEBIMENTO_PARCELADO
         Else
            txtSeq.Text = 1
            txtSeq.SetFocus
      End If
   End If
   TabTipoVenda.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_Click"
End Sub

Private Sub cmbTIPOVENDA_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtpedido.Text = "" Then _
      BuscaLancamento

   Frame1.Enabled = False

   MOSTRA_RODAPE "ESC - SAIR", "Selecione Tipo Venda", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_GotFocus"
End Sub

Private Sub cmbTIPOVENDA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Frame1.Enabled = True Then _
         txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_KeyPress"
End Sub

Private Sub cmbMODALIDADE_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxLanc.ListIndex = cmbMODALIDADE.ListIndex
   
   If cmbMODALIDADE.Text <> "" Then
      If Left(UCase(cmbMODALIDADE.Text), 6) = "CHEQUE" Then
         frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = False
            frmCHEQUECADASTRO.txtPORTADOR.Text = CPF_N
         frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = True
         frmCHEQUECADASTRO.Show 1
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub cmbmodalidade_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione Forma de Pagto."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmodalidade_GotFocus"
End Sub

Private Sub txtcnpjcpf_GotFocus()
'On Error GoTo ERRO_TRATA
   'TXTCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe CPF do cliente"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         MostraCliente
         txtCNPJCPF.PromptInclude = True
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtdesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      'Calcula
      txtValorVenda.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_KeyPress"
End Sub

Private Sub txtDias_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtDias.Text) <> "" Then _
      If IsNumeric(txtDias.Text) Then _
         DIAS_PRAZO = txtDias.Text

   If Trim(txtDias.Text) = "" Then _
      txtDias.Text = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDias_LostFocus"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDTEMIS.PromptInclude = True
   If Not IsDate(txtDTEMIS.Text) Then
      txtDTEMIS.PromptInclude = False
         txtDTEMIS.Text = Date
      txtDTEMIS.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_GotFocus"
End Sub

Private Sub txtDTEMIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDTVENC.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMIS_KeyPress"
End Sub

Private Sub txtDTEMIS_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDTEMIS.PromptInclude = True
   If Not IsDate(txtDTEMIS.Text) Then
      txtDTEMIS.PromptInclude = False
         txtDTEMIS.Text = Date
      txtDTEMIS.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMIS_LostFocus"
End Sub

Private Sub txtDTVENC_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDTVENC.PromptInclude = True

   MOSTRA_RODAPE "Informe Data Vencimento da parcela", "ESC - Confirma", "", "", ""

   If DIAS_PRAZO > 0 Then
      NUMR_SEQ_N = 0 & txtSeq.Text
      DATA_INI = txtDTEMIS.Text
      txtDTVENC.Text = DATA_INI + DIAS_PRAZO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_GotFocus"
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtDTVENC.PromptInclude = True
      If Not IsDate(txtDTVENC.Text) Then
         txtDTVENC.SetFocus
         txtDTVENC.PromptInclude = False
            txtDTVENC.Text = Date
         txtDTVENC.PromptInclude = True
         Exit Sub
      End If
      If txtSeq.Text = "" Then
         MsgBox "Seqüência deve ser gerada ou informada."
         txtSeq.SetFocus
         Exit Sub
      End If
      If cmbAuxLanc.Text = "" Then
         MsgBox "Selecione Forma de Pagamento !!!"
         cmbMODALIDADE.SetFocus
         Exit Sub
      End If
      If txtValorItem.Text = "" Then
         MsgBox "Valor Incorreto !!!"
         txtValorItem.SetFocus
         Exit Sub
      End If
      txtDTEMIS.PromptInclude = True
      If Not IsDate(txtDTEMIS.Text) Then
         MsgBox "Data de emissão inválida !!!"
         txtDTVENC.SetFocus
         Exit Sub
      End If
      txtDTVENC.PromptInclude = True
      If CDate(txtDTVENC.Text) < CDate(txtDTEMIS.Text) Then
         MsgBox "Data de vencimento não pode ser menor que data de emissão !!!"
         txtDTVENC.SetFocus
         Exit Sub
      End If

      KeyAscii = 0
      VALOR_ITEM_N = txtValorItem.Text

      GRAVAR_TUDO
      
      'Criando Relacao com a Tabela de Recebimentos a Vista para efeito de Cupom Fiscal
      'If TABTEMP.State = 1 Then _
         TABTEMP.Close

      'SQL = "Select isnull(max(Sequencia),0) as Sequencia "
      'SQL = SQL & " From ITEMRECEBIMENTO "
      'SQL = SQL & " Where Numr_Req = " & NUMR_REQ_N

      'TABTEMP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      'If TABTEMP!sequencia > 0 Then

      '   If TabConsulta.State = 1 Then _
            TabConsulta.Close

      '   SQL = "Select * From ITEMRECEBIMENTO "
      '   SQL = SQL & " Where Numr_Req = " & NUMR_REQ_N
      '   SQL = SQL & " and Sequencia = " & TABTEMP!sequencia + 1
      '   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      '   If Not TabConsulta.EOF Then
      '      'so altera Valor e forma de pagamento
      '      SQL = "UPDATE ITEMRECEBIMENTO set "
      '      SQL = SQL & " Forma_id = " & cmbAuxLanc.Text
      '      SQL = SQL & " , Valor_item = " & tpMoeda(VALOR_ITEM_N)
      '      SQL = SQL & " Where Numr_Req = " & NUMR_REQ_N
      '      SQL = SQL & " and Sequencia = " & TABTEMP!sequencia
      '      CONECTA_RETAGUARDA.Execute SQL
      '      Else
      '         SQL = "INSERT INTO ITEMRECEBIMENTO "
      '         SQL = SQL & " (Numr_req, Forma_id, Sequencia, Valor_item, Dt_Recebimento) "
      '         SQL = SQL & " VALUES ("
      '            SQL = SQL & NUMR_REQ_N
      '            SQL = SQL & "," & cmbAuxLanc.Text
      '            SQL = SQL & "," & (TABTEMP!sequencia + 1)
      '            SQL = SQL & "," & tpMoeda(VALOR_ITEM_N)
      '            SQL = SQL & ",'" & DMA(Date) & "'"
      '         SQL = SQL & ")"
      '         CONECTA_RETAGUARDA.Execute SqL2
      '   End If
      '   Else
      '      SqL2 = "INSERT INTO ITEMRECEBIMENTO (Numr_req, Forma_id, Sequencia, Valor_item, Dt_Recebimento) "
      '      SqL2 = SqL2 & " VALUES (" & NUMR_REQ_N & "," & cmbAuxLanc.Text & "," & 1 & "," & tpMoeda(VALOR_ITEM_N) & ",'" & DMA(Date) & "')"
      '      CONECTA_RETAGUARDA.Execute SqL2
      'End If
      'TABTEMP.Close

      LIMPA_BODY

      SETA_GRID

      CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_KeyPress"
End Sub

Private Sub txtRecebe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtRecebe.Text <> "" Then
         If IsNumeric(txtRecebe.Text) Then
            VALOR_VENDA_N = 0 & txtValorVenda
            VALOR_RECEBIDO_N = 0 & txtRecebe.Text
            VALOR_TROCO_N = VALOR_RECEBIDO_N - VALOR_VENDA_N

            txtTroco.Text = Format(VALOR_TROCO_N, strFormatacao2Digitos)
            txtTroco.Refresh

            txtRecebe.Text = Format(txtRecebe.Text, strFormatacao2Digitos)
            txtRecebe.Refresh
'MsgBox Format(VALOR_RECEBIDO_N - VALOR_TROCO_N - VALOR_DESCONTO_N, strFormatacao2Digitos) & "       " & Format(VALOR_VENDA_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
            If Format(VALOR_RECEBIDO_N - VALOR_TROCO_N - VALOR_DESCONTO_N, strFormatacao2Digitos) = Format(VALOR_VENDA_N - VALOR_DESCONTO_N, strFormatacao2Digitos) Then _
               Call cmbTipoVenda_Click
         End If
      End If
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF6
         MATA_ITEM_LANCAMENTO
   End Select
End Sub

Private Sub txtTroco_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub tXTVALORVENDA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      CalculaRecebimento
      'Calcula
      'Label10.Visible = True
      'cmbTIPOVENDA.Enabled = True
      If txtTroco.Text >= 0 Then _
         cmbTIPOVENDA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "tXTVALORVENDA_KeyPress"
End Sub

Private Sub txtValorItem_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe o valor da parcela"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   VALR_DESCONTO_N = txtTotal.Text
   VALOR_ITEM_N = 0
   VALR_DESCONTO_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_GotFocus"
End Sub

Private Sub txtValorItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtValorItem.Text <> "" Then
         VALOR_ITEM_N = txtValorItem.Text
         VALOR_ITEM_N = Format(VALOR_ITEM_N, strFormatacao2Digitos)
         VALOR_TOTAL_N = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
         If VALOR_ITEM_N >= VALOR_TOTAL_N Then
            VALOR_TROCO_N = VALOR_RECEBIDO_N - VALOR_TOTAL_N
            txtTroco.Text = Format(VALOR_TROCO_N, strFormatacao2Digitos)
         End If
      End If
      KeyAscii = 0
      txtDias.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_KeyPress"
End Sub

Private Sub txtdias_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe quantidade de dias sua vaca "
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdias_GotFocus"
End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDTEMIS.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdias_KeyPress"
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   SETA_GRID

   VALOR_DIFERENCA_N = 0

   MOSTRA_RODAPE "Tecle <<ENTER>> para nova seqüência, ou selecione", "ESC - Confirma", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtSeq.Text = "" Then
         NUMR_SEQ_N = 1

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l "
         SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
         SQL = SQL & " and i.numr_doc = l.numr_doc "
         SQL = SQL & " and i.lancamento_id = l.lancamento_id "
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and l.tipo_lancamento = " & SINAL
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then _
            If Not IsNull(TabLancamento!ultimo_reg) Then _
               NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         txtSeq.Text = NUMR_SEQ_N
         Else
            SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
            SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
            SQL = SQL & " and i.numr_doc = l.numr_doc "
            SQL = SQL & " and i.lancamento_id = l.lancamento_id "
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and seq = " & txtSeq.Text
            SQL = SQL & " and l.tipo_lancamento = " & SINAL
            TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLancamento.EOF Then
               'valor lançamento
               txtValorItem.Text = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
               VALOR_DIFERENCA_N = TabLancamento!Valor_Item
               'descrição da modalidade
               SQL = "select * from FORMAPAGTO "
               SQL = SQL & " where forma_id = " & TabLancamento!FORMA_ID
               TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabDESCR.EOF Then
                  cmbMODALIDADE.Text = TabDESCR!Descricao
                  cmbAuxLanc.Text = TabDESCR!FORMA_ID
               End If
               TabDESCR.Close
               txtDTVENC.PromptInclude = False
               txtDTEMIS.PromptInclude = False
               txtDTVENC.Text = TabLancamento!DT_VENCIMENTO
               'txtDTEMIS.Text = data_lancamento
               'else
            End If
            TabLancamento.Close
      End If
      cmbMODALIDADE.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyPress"
End Sub

Private Sub cmbMODALIDADE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorItem.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_KeyPress"
End Sub

Private Sub txtValorItem_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorItem.Text <> "" Then _
      txtValorItem.Text = Format(txtValorItem.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_LostFocus"
End Sub

Private Sub txtValorEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe Valor da Entrada"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_GotFocus"
End Sub

Private Sub txtpercEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Confirma"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe Percentual(%) da Entrada"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercEntrada_GotFocus"
End Sub

Private Sub txtPercEntrada_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'If txtPercEntrada.Text <> "" Then
      '   VALOR_ITEM_N = txtPercEntrada.Text
      '   txtValorEntrada.Text = Format(((VALOR_ITEM_N * VALOR_TOTAL_N) / 100), strFormatacao2Digitos)
      '   txtValorEntrada.Refresh
      'End If
      cmbTIPOVENDA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercEntrada_KeyPress"
End Sub

Private Sub txtValorEntrada_keypress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'If txtValorEntrada.Text <> "" Then
      '   VALOR_ITEM_N = txtValorEntrada.Text
      '   txtPercEntrada.Text = Format(((VALOR_ITEM_N / VALOR_TOTAL_N) * 100), strFormatacao2Digitos)
      '   cmbTIPOVENDA.SetFocus
      '   Else: txtPercEntrada.SetFocus
      'End If
      'Else
      '   If KeyAscii = 8 Or KeyAscii = 44 Then
      '      Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
      '   End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_keypress"
End Sub

Private Sub txtValorEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   'If txtValorEntrada.Text <> "" Then
   '   txtValorEntrada.Text = Format(txtValorEntrada.Text, strFormatacao2Digitos)
   'End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_LostFocus"
End Sub

Private Sub txtpercEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   'If txtPercEntrada.Text <> "" Then
   '   txtPercEntrada.Text = Format(txtPercEntrada.Text, strFormatacao2Digitos)
   'End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercEntrada_LostFocus"
End Sub
'============================================================='subrotinas
Private Sub GRAVAR_TUDO()
'On Error GoTo ERRO_TRATA

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      NUMR_ID_N = TabLancamento!LANCAMENTO_ID

      SQL = "UPDATE LANCAMENTO SET "
      SQL = SQL & " Lancamento_id = " & NUMR_ID_N
      SQL = SQL & ", Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Prop = '" & CPF_N & "'"
      SQL = SQL & ", dt_lanc = '" & DMA(Date) & "'"
      SQL = SQL & ", Valor_Lanc = " & Str(txtPago.Text)
      SQL = SQL & ", Total_Desconto = " & Str(VALOR_TOTAL_DESCONTO_N)
      SQL = SQL & ", Tipo_pagto = " & cmbAuxTIPOVENDA.Text
      SQL = SQL & " WHERE Empresa_Id = " & EMPRESA_ID_N
      SQL = SQL & " and Numr_Doc = " & NUMR_REQ_N
      SQL = SQL & " and Tipo_Lancamento = " & SINAL
      Else
          NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

          SQL = "INSERT INTO LANCAMENTO "
            SQL = SQL & " (Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, "
            SQL = SQL & " Total_Desconto, Tipo_Lancamento, Empresa_id, Tipo_pagto) "
          SQL = SQL & " VALUES ("
          SQL = SQL & NUMR_ID_N
          SQL = SQL & "," & NUMR_REQ_N
          SQL = SQL & ",'" & CPF_N & "'"
          SQL = SQL & ",'" & Date & "'"
          SQL = SQL & "," & Str(txtPago.Text)
          SQL = SQL & "," & VALOR_TOTAL_DESCONTO_N
          SQL = SQL & "," & SINAL
          SQL = SQL & "," & EMPRESA_ID_N
          SQL = SQL & "," & cmbAuxTIPOVENDA.Text
          SQL = SQL & ")"

          NUMR_ID_N = TabLancamento!LANCAMENTO_ID
          SQL3 = NUMR_REQ_N
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   CONECTA_RETAGUARDA.Execute SQL

   'ITENS
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and seq = " & txtSeq.Text
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " usu_alt = " & CODG_USU_N
      SQL = SQL & ", Dt_Alt = '" & DMA(Date) & "'"
      SQL = SQL & ", Dt_Cad = '" & DMA(Date) & "'"
      SQL = SQL & ", lancamento_id = " & NUMR_ID_N
      SQL = SQL & ", Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Numr_Dp = " & NUMR_REQ_N
      SQL = SQL & ", Seq = " & txtSeq.Text
      SQL = SQL & ", Valor_Item = " & Str(VALOR_ITEM_N)
      SQL = SQL & ", Status = '" & "A" & "'"
      SQL = SQL & ", FORMA_ID = " & cmbAuxLanc.Text
      SQL = SQL & ", DT_VENCIMENTO = '" & txtDTVENC.Text & "'"
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & txtSeq.Text
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
          SQL = SQL & " (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, "
          SQL = SQL & " seq, Valor_Item, Status, FORMA_ID, DT_VENCIMENTO, acerto) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CODG_USU_N
            SQL = SQL & ",'" & DMA(Date) & "'"
            SQL = SQL & ",'" & DMA(Date) & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & txtSeq.Text
            SQL = SQL & "," & Str(VALOR_ITEM_N)
            SQL = SQL & ",'A'"
            SQL = SQL & "," & cmbAuxLanc.Text
            SQL = SQL & ",'" & DMA(txtDTVENC.Text) & "'"
            SQL = SQL & ", 1 "
         SQL = SQL & ")"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SQL

   If Left(UCase(cmbMODALIDADE.Text), 8) = "DINHEIRO" _
   Or Left(UCase(cmbMODALIDADE.Text), 6) = "CHEQUE" _
   Or Left(UCase(cmbMODALIDADE.Text), 6) = "CARTAO" _
   Or Left(UCase(cmbMODALIDADE.Text), 6) = "CARTÃO" Then

      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " Status = 'B'"
      SQL = SQL & ", DT_BAIXA = '" & DMA(Date) & "'"
      SQL = SQL & ", CODG_USU_BAIXA = " & CODG_USU_N
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & txtSeq.Text

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVAR_TUDO"
End Sub

Private Sub MATA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select lancamento_id from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      If Not IsNull(TabLancamento.Fields(0).Value) Then
         SQL = "delete from ITEMLANCAMENTO "
         SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
         SQL = SQL & " and seq = " & txtSeq.Text
         CONECTA_RETAGUARDA.Execute SQL

         'SQL = "delete from LANCAMENTO "
         'SQL = SQL & " where lancamento_id = " & TABLANCAMENTO.Fields(0).Value
         'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         'SQL = SQL & " and tipo_lancamento = " & SINAL
         'CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   BUSCA_LANCAMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_LANCAMENTO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 0
   ListaLanc.ListItems.Clear

   If TabLANCAMENTOITEM.State = 1 Then
      TabLANCAMENTOITEM.Close
   End If

   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLANCAMENTOITEM.EOF
      'sequencia

      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set Item = ListaLanc.ListItems.Add(, "seq." & NUMR_SEQ_N, TabLANCAMENTOITEM!SEQ)
      'numero documento
      Item.SubItems(1) = TabLANCAMENTOITEM!NUMR_DOC
      'valor lançamento
      Item.SubItems(2) = Format(TabLANCAMENTOITEM!Valor_Item, strFormatacao2Digitos)

      'descrição da modalidade
      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where FORMA_ID = " & TabLANCAMENTOITEM!FORMA_ID
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      If TabDESCR.State = 1 Then TabDESCR.Close
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         Item.SubItems(3) = TabDESCR!Descricao
      TabDESCR.Close

      Item.SubItems(4) = Date
      Item.SubItems(5) = TabLANCAMENTOITEM!DT_VENCIMENTO

      If cmbAuxTIPOVENDA.Text <> "" Then
         SQL = "select * from TIPOVENDA "
         SQL = SQL & " where TIPOVENDA_id = " & cmbAuxTIPOVENDA.Text
         'SQL = SQL & " and empresa_id = " & EMPRESA_ID_n
         If TabAUX.State = 1 Then TabAUX.Close
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If Not IsNull(TabAUX!PERC_JUROS) Then
               Item.SubItems(6) = TabAUX!PERC_JUROS & "%"
               Else: Item.SubItems(6) = "00,00 %"
            End If
         End If
         TabAUX.Close
      End If
      TabLANCAMENTOITEM.MoveNext
   Wend
   TabLANCAMENTOITEM.Close

   BUSCA_LANCAMENTO

   txtTotal.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   lblPRAZO.Caption = ""
   cmbTIPOVENDA.Text = ""
   cmbAuxTIPOVENDA.Text = ""
   txtpedido.Text = ""
   txtVendedor.Text = ""
   txtData.Text = ""
   txtTotal.Text = ""
   txtTroco.Text = ""
   txtValorVenda.Text = ""
   txtNome.Text = ""
   cmbAuxLanc.Clear
   cmbMODALIDADE.Clear
   txtValorItem.Text = ""
   txtDTEMIS.PromptInclude = False
   txtDTVENC.PromptInclude = False
   txtDTEMIS.Text = ""
   txtDTVENC.Text = ""
   ListaLanc.ListItems.Clear
   
   txtSeq.Text = ""
   VALOR_TOTAL_LANÇADO = 0
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_LANCAMENTO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   VALOR_DIFERENCA_N = 0
   VALOR_ITEM_N = 0
   txtSeq.Text = ""

   cmbAuxLanc.Text = ""
   If Left(UCase(cmbMODALIDADE.Text), 8) <> "DINHEIRO" Then _
      cmbMODALIDADE.Text = ""

   txtValorItem.Text = ""
   txtDias.Text = ""
   txtDTEMIS.PromptInclude = False
   txtDTVENC.PromptInclude = False
   txtDTEMIS.Text = ""
   txtDTVENC.Text = ""
   VALOR_TOTAL_LANÇADO = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub GERA_FATURAMENTO()
'On Error GoTo ERRO_TRATA

   Dim Valor_Tot_n As Double

   NUMR_PARCELA = 0
   SQL = "select * from TIPOVENDA "
   SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      NUMR_PARCELA = TabTemp!parcela
      
      VALOR_DESCONTO_N = 0
      SQL = "select perc_desc from CABECAREQ "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & TabCABECA!numr_req
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         PERC_DESCONTO_N = TabPedidoItem.Fields(0).Value
      TabPedidoItem.Close
      
      VALOR_DESCONTO_CABECA_N = 0
      SQL = "select valor_desconto from CABECAREQ "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & TabCABECA!numr_req
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_CABECA_N = TabPedidoItem.Fields(0).Value
      TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from ITEMREQ "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & TabCABECA!numr_req
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      TabPedidoItem.Close

      'BUSCA VALOR TOTAL VENDA
      Valor_Tot_n = 0
      SQL = "select sum(valor_item*qtd_pedida) from ITEMREQ "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and numr_req = " & TabCABECA!numr_req
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         Valor_Tot_n = TabPedidoItem.Fields(0).Value
      TabPedidoItem.Close
      
      'VALOR_DESCONTO_N = VALOR_DESCONTO_N + (Valor_Tot_n * IIf(Perc_Desconto_n > 0, Perc_Desconto_n / 100, 1))
      VALOR_DESCONTO_N = VALOR_DESCONTO_N + VALOR_DESCONTO_CABECA_N
      
      VALOR_ITEM_N = 0
      DATA_INI = Date
      If NUMR_PARCELA > 0 Then _
         VALOR_ITEM_N = (Valor_Tot_n - VALOR_DESCONTO_N) / NUMR_PARCELA

      'CABEÇA
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      SQL = "select * from LANCAMENTO "
      SQL = SQL & " where numr_doc = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and tipo_lancamento = " & SINAL
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         NUMR_ID_N = TabLancamento!LANCAMENTO_ID

         SQL = "UPDATE LANCAMENTO SET "
         SQL = SQL & " Numr_doc = " & NUMR_REQ_N
         SQL = SQL & ", Prop = '" & CPF_N & "'"
         SQL = SQL & ", dt_lanc = '" & TabCABECA!DT_REQ & "'"
         SQL = SQL & ", Valor_Lanc = " & Str(Format(Valor_Tot_n, strFormatacao2Digitos))
         SQL = SQL & ", Total_Desconto = " & Str(Format(VALOR_DESCONTO_N, strFormatacao2Digitos))
         SQL = SQL & ", Tipo_Lancamento = " & SINAL
         SQL = SQL & ", Empresa_Id = " & EMPRESA_ID_N
         SQL = SQL & ", Tipo_pagto = " & cmbAuxTIPOVENDA.Text
         SQL = SQL & " WHERE Empresa_Id = " & EMPRESA_ID_N
         SQL = SQL & " and Numr_Doc = " & NUMR_REQ_N
         SQL = SQL & " and Tipo_Lancamento = " & SINAL
         Else
            NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

            SQL = "INSERT INTO LANCAMENTO "
               SQL = SQL & " (Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, "
               SQL = SQL & " Total_Desconto, Tipo_Lancamento, Empresa_id, Tipo_pagto) "
            SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & NUMR_REQ_N
               SQL = SQL & ",'" & CPF_N & "'"
               SQL = SQL & ",'" & TabCABECA!DT_REQ & "'"
               SQL = SQL & "," & Str(Format(Valor_Tot_n, strFormatacao2Digitos))
               SQL = SQL & "," & Str(Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos))
               SQL = SQL & "," & SINAL
               SQL = SQL & "," & EMPRESA_ID_N
               SQL = SQL & "," & cmbAuxTIPOVENDA.Text
            SQL = SQL & ")"
      End If
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      CONECTA_RETAGUARDA.Execute SQL

      SQL3 = NUMR_REQ_N
      SqL2 = EMPRESA_ID_N
      CONT_N = 0
      'ITENS
      While CONT_N < NUMR_PARCELA
         GRAVA_LANÇAMENTO
         CONT_N = CONT_N + 1
      Wend
   End If
   TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FATURAMENTO"
End Sub

Private Sub GRAVA_LANÇAMENTO()
'On Error GoTo ERRO_TRATA

   Dim Situacao_a As String
   Dim DATA_TEXTO As String

   Situacao_a = "A"
   If Trim(cmbAuxTIPOVENDA.Text) = "9999" Then _
      Situacao_a = "B"

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   NUMR_SEQ_N = 1
   SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento!ultimo_reg) Then _
         NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   'tem que dividir os dias de prazo
   If NUMR_PARCELA <= 0 Then
      DATA_INI = DATA_INI + DIAS_PRAZO
      Else: DATA_INI = DATA_INI + TabTemp!Prazo
   End If
   DATA_TEXTO = DATA_INI

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where seq = " & NUMR_SEQ_N
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " usu_alt = " & CODG_USU_N
      SQL = SQL & ", Dt_Alt = '" & DMA(Date) & "'"
      SQL = SQL & ", Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Seq = " & NUMR_SEQ_N
      SQL = SQL & ", Valor_Item = " & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA / NUMR_PARCELA))
      SQL = SQL & ", Status = '" & Situacao_a & "'"
      SQL = SQL & ", FORMA_ID = " & TabTemp!FORMA_ID
      SQL = SQL & ", DT_VENCIMENTO = '" & DMA(DATA_TEXTO) & "'"
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & NUMR_SEQ_N
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
            SQL = SQL & " (Usu_Alt, Dt_Alt, Lancamento_id, Numr_doc, NUMR_DP, seq, "
            SQL = SQL & " Valor_Item, Status, FORMA_ID, DT_VENCIMENTO, ACERTO) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CODG_USU_N
            SQL = SQL & ",'" & DMA(Date) & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & "," & NUMR_SEQ_N
            SQL = SQL & "," & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA / NUMR_PARCELA))
            SQL = SQL & ",'" & Situacao_a & "'"
            SQL = SQL & "," & TabTemp!FORMA_ID
            SQL = SQL & ",'" & DMA(DATA_TEXTO) & "'"
            SQL = SQL & "," & 0
         SQL = SQL & ")"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_LANÇAMENTO"
End Sub

Private Sub BUSCA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_LANÇADO = 0
   VALOR_RECEBIDO_N = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select sum(valor_item) from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento.Fields(0).Value) Then _
         VALOR_TOTAL_LANÇADO = TabLancamento.Fields(0).Value

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_LANCAMENTO"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO

   VALOR_TOTAL_LANÇADO = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   VALOR_TOTAL_N = Format((VALOR_TOTAL_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

   If VALOR_TOTAL_LANÇADO = VALOR_TOTAL_N Then
      If Left(UCase(cmbMODALIDADE.Text), 8) = "DINHEIRO" Then
         Msg = "Confirma recebimento ?"
         PERGUNTA Msg, vbYesNo + 32, "Recebimento NFE", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            SQL = "select * from CABECAREQ "
            SQL = SQL & " where numr_req = " & txtpedido.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCABECA.EOF Then
               If TabCABECA!Status = 2 Then
                  SQL = "update CABECAREQ set "
                  SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
                  SQL = SQL & ",TipoVenda_id = " & cmbAuxTIPOVENDA.Text 'Tipo venda
                  SQL = SQL & ",CGCCPF = '" & txtCNPJCPF.Text & "'" 'CPF do Cliente Atualizado
                  SQL = SQL & ",Valor_desconto = " & tpMoeda(txtDesconto.Text)  'Valor Desconto
                  SQL = SQL & ",Perc_desc = " & tpMoeda(PERC_DESCONTO_N)  'Perc_desconto
                  SQL = SQL & ",Valor_Recebido = " & tpMoeda(txtPago.Text)  'Valor Recebido
                  SQL = SQL & " where numr_req = " & txtpedido.Text
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
            If TabCABECA.State = 1 Then _
               TabCABECA.Close

            ROTINACUPOM

            Unload Me
            Exit Sub
         End If
         Else
             If VALOR_TOTAL_LANÇADO >= VALOR_TOTAL_N Then
                If txtTroco.Text <> "" Then _
                   If VALOR_TROCO_N > 0 Then _
                      REGISTRA_TROCO

                Msg = "Confirma recebimento ?"
                PERGUNTA Msg, vbYesNo + 32, "Recebimento NFE", "DEMO.HLP", 1000
                If RESPOSTA = vbYes Then
                   SQL = "select * from CABECAREQ "
                   SQL = SQL & " where numr_req = " & txtpedido.Text
                   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                   If Not TabCABECA.EOF Then
                      If TabCABECA!Status = 2 Then
                         SQL = "update CABECAREQ set "
                         SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
                         SQL = SQL & ",TipoVenda_id = " & cmbAuxTIPOVENDA.Text 'Tipo venda
                         SQL = SQL & ",CGCCPF = '" & txtCNPJCPF.Text & "'" 'CPF do Cliente Atualizado
                         SQL = SQL & ",Valor_desconto = " & tpMoeda(txtDesconto.Text) 'Valor Desconto
                         SQL = SQL & ",Perc_desc = " & tpMoeda(PERC_DESCONTO_N) 'Perc_desconto
                         SQL = SQL & ",Valor_Recebido = " & tpMoeda(txtPago.Text)  'Valor Recebido
                         SQL = SQL & " where numr_req = " & txtpedido.Text
                         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                         CONECTA_RETAGUARDA.Execute SQL
                      End If
                   End If
                   TabCABECA.Close
                   ROTINACUPOM
                   Unload Me
                   Exit Sub
                End If
             End If
      End If
   End If
   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA"
End Sub

Private Sub REGISTRA_TROCO()
'On Error GoTo ERRO_TRATA

   SQL = "select lancamento_id from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      If Not IsNull(TabLancamento.Fields(0).Value) Then
         SQL = "update itemLANCAMENTO set "
         SQL = SQL & " valor_desconto = valor_desconto + " & Replace(VALOR_TROCO_N, ",", ".")
         SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
         SQL = SQL & " and seq = " & NUMR_SEQ_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REGISTRA_TROCO"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_PARCELADO()
'On Error GoTo ERRO_TRATA

   If cmbAuxTIPOVENDA.Text = "" Then
      MsgBox "Selecione Forma de Pgto, Para Confirmar Recebimento!", vbExclamation, "MEGASIM"
      cmbTIPOVENDA.SetFocus
      Exit Sub
   End If

   BUSCA_LANCAMENTO

   If VALOR_TOTAL_LANÇADO <= 0 Then
      MsgBox "Atenção, realizar faturamento."
      Exit Sub
   End If

   VALOR_ITEM_N = 0 & txtTotal.Text
   If Format(VALOR_ITEM_N, strFormatacao2Digitos) = Format(VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos) Then
      INDR_FINALIZA_RECEBIMENTO = False

      Msg = "Confirma recebimento ?"
      PERGUNTA Msg, vbYesNo + 32, "Recebimento NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         txtpedido.Text = NUMR_REQ_N

         If txtpedido.Text = "" Then _
            txtpedido.Text = 1

         SQL = "select * from CABECAREQ "
         SQL = SQL & " where numr_req = " & txtpedido.Text
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCABECA.EOF Then
            If TabCABECA!Status = 2 Then
               SQL = "update CABECAREQ set "
               SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
               SQL = SQL & ",TipoVenda_id = " & cmbAuxTIPOVENDA.Text 'Tipo venda

               If Not IsNull(TabCABECA.Fields("CLIENTE_ID").Value) Then
                  SQL = SQL & ",CLIENTE_ID = " & TabCABECA.Fields("CLIENTE_ID").Value
                  Else: SQL = SQL & ",CLIENTE_ID = 1 "
               End If

               SQL = SQL & ",Valor_desconto = " & Str(txtDesconto.Text)  'Valor Desconto
               SQL = SQL & ",Perc_desc = " & Str(PERC_DESCONTO_N) 'Perc_desconto
               SQL = SQL & ",Valor_Recebido = " & tpMoeda(txtTotal.Text)  'Valor Recebido
               SQL = SQL & " where numr_req = " & txtpedido.Text
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
         TabCABECA.Close

         'ROTINACUPOM

         LIMPA_LANCAMENTO

         ControlaReceb = True
         Me.Hide
         Exit Sub
      End If
      Else: MsgBox "Valores de venda com recebimento não conferem."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub

Private Sub EXCLUIR_TUDO()
'On Error GoTo ERRO_TRATA

   SQL = "delete ITEMLANCAMENTO from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and l.tipo_lancamento = " & SINAL
   CONECTA_RETAGUARDA.Execute SQL
   
   SQL = "delete from ITEMRECEBIMENTO "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   CONECTA_RETAGUARDA.Execute SQL
         
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_TUDO"
End Sub

Private Sub MostraCliente()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   intCodigo = 0

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from CLIENTE "
   SQL = SQL & " where CGCCPF = '" & txtCNPJCPF.Text & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      txtNome.Text = TabCliente!NOME
      intCodigo = TabTemp.Fields("cliente_id").Value
      Else
         If TabCliente.State = 1 Then _
            TabCliente.Close

         MsgBox "Cliente Nao Cadastrado Favor Verificar", vbCritical, "MEGASIM"
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
         txtNome.Text = ""
         txtCNPJCPF.SetFocus
         Exit Sub
   End If
   txtDesconto.SetFocus

   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MostraCliente"
End Sub

Private Sub BuscaLancamento()
'On Error GoTo ERRO_TRATA

   SINAL = 1
   VALOR_TOTAL_N = 0
   Frame1.Enabled = False
   NUMR_PARCELA = 0
   LIMPA_LANCAMENTO
   ControlaReceb = False
   txtData.Text = Now
   txtpedido.Text = NUMR_REQ_N
   If NUMR_REQ_N > 0 Then
      SETA_GRID
      Else
         MsgBox "Número de lançamento não foi informado. verifique."
         Unload Me
   End If

   SQL = "select * from CABECAREQ "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   If TabCABECA.State = 1 Then TabCABECA.Close
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      SQL = "select nome_vend from VENDEDOR v, EQUIPE e "
      SQL = SQL & " where v.vendedor_id = " & TabCABECA!VENDEDOR_ID
      SQL = SQL & " and e.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and v.codg_eq = e.codg_eq "
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         txtVendedor.Text = TabVENDEDOR!NOME_VEND
         txtVendedor.Refresh
      End If
      TabVENDEDOR.Close
        
      SQL = "select nome from CLIENTE "
      SQL = SQL & " where cgccpf = '" & TabCABECA!CGCCPF & "'"
      SQL = SQL & " and status = 'A'"
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         txtCNPJCPF.Text = TabCABECA!CGCCPF
         If TabCABECA!CGCCPF = "00000000000000" Then
            txtNome.Text = TabCABECA!NOME_CLIENTE
            Else
                If TabCABECA!CGCCPF = "00000000000" Then
                   txtNome.Text = TabCABECA!NOME_CLIENTE
                   Else
                       txtNome.Text = TabVENDEDOR!NOME
                End If
                
         End If
         txtNome.Refresh
      End If
      CPF_N = TabCABECA!CGCCPF
      TabVENDEDOR.Close

      PERC_DESCONTO_N = 0
      VALOR_DESCONTO_N = 0
      VALOR_DESCONTO_CABECA_N = 0
      VALOR_TOTAL_DESCONTO_N = 0
      VALOR_ITEM_N = 0
      If Not IsNull(TabCABECA!Valor_Total) Then
         VALOR_ITEM_N = TabCABECA!Valor_Total
      Else
         SQL = "Select sum(Valor_Item*Qtd_Pedida) as ValorTotal from ITEMREQ Where Empresa_id = " & EMPRESA_ID_N & " and Numr_req = " & NUMR_REQ_N
         If TabPedidoItem.State = 1 Then TabPedidoItem.Close
         TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedidoItem.EOF Then
            VALOR_ITEM_N = TabPedidoItem!valortotal
         End If
         TabPedidoItem.Close
      End If
      
      VALOR_TOTAL_DESCONTO_N = VALOR_ITEM_N * PERC_DESCONTO_N / 100
      VALOR_DESCONTO_N = VALOR_DESCONTO_N + VALOR_DESCONTO_CABECA_N + VALOR_TOTAL_DESCONTO_N
      VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
   End If
   TabCABECA.Close
   
   txtValorVenda.Text = Format(txtValorVenda.Text, "###,##0.00")
   txtValorVenda.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTotal.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTroco.Text = "0,00"
   txtPago.Text = "0,00"
   txtDesconto.Text = "0,00"
   
   txtTroco.Text = Format(txtTroco.Text, "###,##0.00")
   txtDesconto.Text = Format(txtDesconto.Text, "###,##0.00")
   txtPago.Text = Format(txtPago.Text, "###,##0.00")

   txtCNPJCPF.PromptInclude = False
   BUSCA_LANCAMENTO
   
   txtTotal.Refresh

   Call CentralizaJanela2(frmCAIXAREC)

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   'txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BuscaLancamento"
End Sub

Private Sub Calcula()
'On Error GoTo ERRO_TRATA

   If IsNumeric(txtDesconto.Text) Then
      If txtDesconto.Text > 0 Then
         VALOR_TOTAL_DESCONTO_N = 0
         If txtDesconto.Text > 0 And VALOR_ITEM_N > 0 Then
            txtValorVenda.Text = VALOR_ITEM_N
            VALOR_TOTAL_DESCONTO_N = Format(((txtValorVenda.Text * txtDesconto.Text) / 100), strFormatacao3Digitos)
            txtValorVenda.Text = Format((txtValorVenda.Text - (txtValorVenda.Text * txtDesconto.Text) / 100), strFormatacao2Digitos)
            PERC_DESCONTO_N = txtDesconto.Text
         End If
         Else
            If VALOR_ITEM_N > 0 Then
               txtValorVenda.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
               Else: txtValorVenda.Text = Format(txtPago.Text, strFormatacao2Digitos)
            End If
            VALOR_TOTAL_DESCONTO_N = 0
            PERC_DESCONTO_N = 0
      End If

      txtPago.Text = Format((VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N), strFormatacao2Digitos)
      txtDesconto.Text = Format(txtDesconto.Text, "###,##0.00")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Calcula"
End Sub

Private Sub CalculaRecebimento()
'On Error GoTo ERRO_TRATA

   If IsNumeric(txtValorVenda.Text) Then
      If txtValorVenda.Text > 0 Then
         txtTroco.Text = Format(txtValorVenda.Text - (VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N), strFormatacao2Digitos)
         txtPago.Text = Format((VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N), strFormatacao2Digitos)
         If txtTroco.Text < 0 Then
            MsgBox "Valor Recebido Menor que Valor da Compra!", vbCritical, "MEGASIM"
            txtValorVenda.SetFocus
            Exit Sub
         End If
      End If
      txtValorVenda.Text = Format(txtValorVenda.Text, "###,##0.00")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CalculaRecebimento"
End Sub

Private Sub ROTINACUPOM()
   Dim logCARTÃO     As Boolean
   Dim DESCONTO_N    As Currency
   Dim FORMA_PAGTO   As Long

   'Verifica se o cliente e Dentro do Estado ou Fora do Estado Caso seja fora vai gerar NFE Direto
   If booUsaImpFiscal = True Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      txtCNPJCPF.PromptInclude = False

      SQL = "Select CGCCPF from Cliente "
      SQL = SQL & " Inner Join ENDERECO "
      SQL = SQL & " ON CLIENTE.CGCCPF = ENDERECO.PROP "
      SQL = SQL & " INNER JOIN CEP "
      SQL = SQL & " ON ENDERECO.CEP = CEP.CEP "
      SQL = SQL & " Where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
      SQL = SQL & " and CEP.UF = 'GO'"
      SQL = SQL & " and status = 'A'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         'verifica se cartão de crédito ou débito
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "SELECT Forma_id, Descricao FROM FormaPagto "
         SQL = SQL & " WHERE Forma_id = " & cmbAuxTIPOVENDA.Text
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            FORMA_PAGTO = TabConsulta!FORMA_ID
            If TabConsulta!Descricao = "CARTAO CREDITO" Or TabConsulta!Descricao = "CARTAO DEBITO" Then
               logCARTÃO = True
               DESCONTO_N = txtDesconto.Text
               Else: logCARTÃO = False
            End If
         End If
         TabConsulta.Close

         Me.MousePointer = 11
         If logCARTÃO = True Then
            'se esta variavel estiver como false o tef e Discado entao precisa verificar o gerenciador esta ativo
            If booTipoTef = False Then
               If VerificaGerenciadorPadrao(1) = True Then
                  ImprimeCupom NUMR_REQ_N, logCARTÃO, DESCONTO_N, FORMA_PAGTO
                  Else
                     MsgBox "Gerenciador Padrão não está ativo.", 48, Me.Caption
                     'cmdImprimirNotaFiscal.Enabled = True
                     Exit Sub
               End If
               Else: ImprimeCupom NUMR_REQ_N, logCARTÃO, DESCONTO_N, FORMA_PAGTO
            End If
            Else: ImprimeCupom NUMR_REQ_N, logCARTÃO, DESCONTO_N, FORMA_PAGTO
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      txtCNPJCPF.PromptInclude = True
   End If
End Sub

Sub INICIALIZA_RECEBIMENTO()

   LIMPA_LANCAMENTO

   SINAL = 1
   Me.Caption = Me.Caption & " - " & Me.Name & " Sinal = " & SINAL & " nr " & NUMR_REQ_N
   ControlaReceb = False
   VALOR_TOTAL_N = 0
   Frame1.Enabled = False
   NUMR_PARCELA = 0

   txtData.Text = Now
   txtpedido.Text = NUMR_REQ_N
   If NUMR_REQ_N > 0 Then
      SETA_GRID
      Else
         MsgBox "Número de lançamento não foi informado. verifique."
         Unload Me
         Exit Sub
   End If

   'Preechendo forma de pagamento
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbTIPOVENDA.Clear
   cmbAuxTIPOVENDA.Clear

   SQL = "select * from TIPOVENDA "
   SQL = SQL & " order by tipovenda_id desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbTIPOVENDA.AddItem Trim(TabTemp!Descricao)
      cmbAuxTIPOVENDA.AddItem Trim(TabTemp!tipovenda_id)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   VALOR_DESCONTO_CABECA_N = 0
   VALOR_DESCONTO_ITEM_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_VENDA_N = 0

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select * from CABECAREQ "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   If TabCABECA.State = 1 Then TabCABECA.Close
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      SQL = "select nome_vend from VENDEDOR v, EQUIPE e "
      SQL = SQL & " where v.vendedor_id = " & TabCABECA!VENDEDOR_ID
      SQL = SQL & " and e.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and v.codg_eq = e.codg_eq "
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         txtVendedor.Text = TabVENDEDOR!NOME_VEND
         txtVendedor.Refresh
      End If
      TabVENDEDOR.Close

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "select nome from CLIENTE "
      If Not IsNull(TabCABECA.Fields("CLIENTE_ID").Value) Then
         SQL = SQL & " where CLIENTE_ID = " & TabCABECA.Fields("CLIENTE_ID").Value
         Else: SQL = SQL & " where CLIENTE_ID = " & 1
      End If
      SQL = SQL & " and status = 'A'"
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         txtCNPJCPF.Text = TabCABECA!CGCCPF
         If TabCABECA!CGCCPF = "00000000000000" Then
            txtNome.Text = TabCABECA!NOME_CLIENTE
            Else
               If TabCABECA!CGCCPF = "00000000000" Then
                  txtNome.Text = TabCABECA!NOME_CLIENTE
                  Else: txtNome.Text = TabVENDEDOR!NOME
               End If
         End If
         txtNome.Refresh
      End If
      CPF_N = TabCABECA!CGCCPF
      TabVENDEDOR.Close

      VALOR_DESCONTO_CABECA_N = 0 & TabCABECA.Fields("valor_desconto").Value

      If Not IsNull(TabCABECA!Valor_Total) Then _
         VALOR_VENDA_N = TabCABECA!Valor_Total

      If Not IsNull(TabCABECA.Fields("tipovenda_id").Value) Then
         'forma de pagamento
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from TIPOVENDA "
         SQL = SQL & " where tipovenda_id = " & TabCABECA.Fields("tipovenda_id").Value
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            cmbTIPOVENDA.Text = Trim(TabTemp!Descricao)
            cmbAuxTIPOVENDA.Text = Trim(TabTemp!tipovenda_id)
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

'VALOR ITENS
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "Select sum(Valor_Item*Qtd_Pedida) as ValorTotal from ITEMREQ "
   SQL = SQL & " Where Empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and Numr_req = " & NUMR_REQ_N
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then _
      VALOR_VENDA_N = TabPedidoItem!valortotal
   TabPedidoItem.Close

'desconto itens
   SQL = "select sum(valor_desconto) from ITEMREQ "
   SQL = SQL & " where numr_req = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then _
      If Not IsNull(TabCABECA.Fields(0).Value) Then _
         VALOR_DESCONTO_ITEM_N = TabCABECA.Fields(0).Value

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   VALOR_TOTAL_DESCONTO_N = VALOR_DESCONTO_ITEM_N + VALOR_DESCONTO_CABECA_N

   VALOR_TOTAL_N = VALOR_VENDA_N - VALOR_DESCONTO_ITEM_N

   txtValorVenda.Text = Format(VALOR_VENDA_N, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
   txtTotal.Text = Format(VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)

   txtTroco.Text = "0,00"
   txtPago.Text = "0,00"
   txtTroco.Text = Format(txtTroco.Text, "###,##0.00")
   txtDesconto.Text = Format(txtDesconto.Text, "###,##0.00")
   txtPago.Text = Format(txtPago.Text, "###,##0.00")

   BUSCA_LANCAMENTO

   txtTotal.Refresh
   txtCNPJCPF.PromptInclude = False
End Sub

Sub MATA_ITEM_LANCAMENTO()
   If txtSeq.Text <> "" Then
      Msg = "Confirma Exclusão do Item =  ?" & txtSeq.Text
      Style = vbYesNo + 32
      Title = "Atenção !!!"
      Help = "DEMO.HLP"
      Ctxt = 1000
      RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
      If RESPOSTA = vbYes Then
         MATA_LANCAMENTO
         SETA_GRID
      End If
      Else: MsgBox "Informe número da seqüência."
   End If
End Sub
