VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNOTAENTRADAFINANC 
   Caption         =   "Geração de Títulos à Pagar"
   ClientHeight    =   7560
   ClientLeft      =   2295
   ClientTop       =   2475
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADFINENTRADA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   11055
      Begin VB.TextBox txtNota 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5085
         MaxLength       =   9
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtDesconto 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   27
         Top             =   1800
         Width           =   1815
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
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTipoVenda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   6015
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
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
         Left            =   9000
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtVendaComDesconto 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtRecebido 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   15
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtVendaSemDesconto 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtFornec 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   6015
      End
      Begin VB.TextBox txtVendedor 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   450
         Left            =   1560
         TabIndex        =   30
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nº.NF:"
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
         Height          =   240
         Left            =   4320
         TabIndex        =   38
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblEntrada_ID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4830
         TabIndex        =   32
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "S. NF:"
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
         Height          =   240
         Left            =   6180
         TabIndex        =   31
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Desconto:"
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
         Left            =   7395
         TabIndex        =   28
         Top             =   1800
         Width           =   1500
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
         TabIndex        =   26
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label10 
         Caption         =   "Forma Faturamento:"
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
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   1515
         Width           =   1950
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Compra com Desconto:"
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
         Left            =   6690
         TabIndex        =   22
         Top             =   2280
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Pagamento:"
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
         Left            =   7170
         TabIndex        =   20
         Top             =   2760
         Width           =   1725
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Compra sem Desconto: "
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
         Left            =   6645
         TabIndex        =   19
         Top             =   1320
         Width           =   2250
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
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
         Left            =   7605
         TabIndex        =   17
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsável:"
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
         Height          =   240
         Left            =   7635
         TabIndex        =   16
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   11055
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   615
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
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtValorItem 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5400
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbMODALIDADE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin MSMask.MaskEdBox txtDTVENC 
         Height          =   450
         Left            =   9240
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   794
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
         Height          =   450
         Left            =   7320
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   794
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Vencimento"
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
         Index           =   3
         Left            =   9240
         TabIndex        =   37
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Emissão"
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
         Index           =   2
         Left            =   7320
         TabIndex        =   36
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor"
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
         Index           =   1
         Left            =   5445
         TabIndex        =   35
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Modalidade"
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
         Left            =   870
         TabIndex        =   34
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Seq."
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
         Index           =   0
         Left            =   210
         TabIndex        =   33
         Top             =   240
         Width           =   435
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADFINENTRADA.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1270
      ButtonWidth     =   2937
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
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
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
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5880
         Top             =   0
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
               Picture         =   "CADFINENTRADA.frx":76BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADFINENTRADA.frx":8849
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADFINENTRADA.frx":9F46
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADFINENTRADA.frx":AFD5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListaLanc 
      Height          =   2505
      Left            =   50
      TabIndex        =   10
      Top             =   5040
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   4419
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
         Text            =   "Seq."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Doc."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Modalidade"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Dt.Lanç."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dt.Venc."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Juros"
         Object.Width           =   1764
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11070
      DesignHeight    =   7560
   End
End
Attribute VB_Name = "frmNOTAENTRADAFINANC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
   Dim VALOR_RECEBIDO_N As Double, NUMR_PARCELA As Integer
   Dim VALOR_TROCO_N As Double, VALOR_TOTAL_LANÇADO As Double
   Dim VALOR_ENTRADA As Double, PERC_JUROS_N As Double, DIAS_PRAZO As Integer
   Dim VALOR_ICMS_SUB_N As Currency, LANCAMENTO_ID_N As Long

Private Sub Form_Activate()

   txtNOTA.SetFocus

End Sub

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

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
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         Msg = "Deseja Cancelar todo lançamento ?"
         Style = vbYesNo + 32
         Title = "Atenção !!!"
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then _
            MATA_LANCAMENTO
         SETA_GRID
         LIMPA_BODY
         VALOR_ITEM_N = 0
         VALOR_ENTRADA = 0
         cmbTipoVenda.Text = ""
         cmbAuxTIPOVENDA.Text = ""
         cmbTipoVenda.SetFocus
      Case "matar"
         If txtSeq.Text <> "" Then
            Msg = "Confirma Exclusão do Item =  ?" & txtSeq.Text
            Style = vbYesNo + 32
            Title = "Atenção !!!"
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then _
               MATA_LANCAMENTO
            SETA_GRID
            Else: MsgBox "Informe número da seqüência."
         End If
      Case "voltar"
         CONFIRMAR_RECEBIMENTO_PARCELADO
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub listalanc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView ListaLanc, ColumnHeader
End Sub

Private Sub cmbMODALIDADE_LostFocus()
   cmbModalidade.BackColor = &HFFFFFF
End Sub

Private Sub txtNota_GotFocus()
'On Error GoTo ERRO_TRATA

   If INDR_RECEITA <> 2 Then
      MsgBox "Registro não é do contas a pagar, não permitido, chamar suporte."
      Unload Me
      Exit Sub
   End If

   If NOTAENTRADA_ID_N <= 0 Then
      MsgBox "Parametro de registro não informado."
      Unload Me
      Exit Sub
   End If
   NUMR_PARCELA = 0

   LIMPA_TUDO

   txtData.Text = Now

   'GERA_PEDIDO_ID

   'If PEDIDO_ID_N > 0 Then
      SETA_GRID
   '   Else
   '      MsgBox "Número de lançamento não foi informado. verifique."
   '      Unload Me
   'End If

   lblEntrada_ID.Caption = NOTAENTRADA_ID_N
   Dim TabNOTA As New ADODB.Recordset

   If TIPO_ENTRADA_N = 1 Then _
      MOSTRA_NOTA
   If TIPO_ENTRADA_N = 2 Then _
      MOSTRA_NOTA_AVULSA

   txtRecebido.Refresh
   txtVendaSemDesconto.Refresh

   cmbTipoVenda.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNota_GotFocus"
End Sub

Private Sub cmbTIPOVENDA_LostFocus()
'On Error GoTo ERRO_TRATA

   cmbTipoVenda.BackColor = &HFFFFFF

   If cmbAuxTIPOVENDA.Text <> "" Then
      cmbModalidade.Clear
      cmbAuxLanc.Clear

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK) "
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and status = 'true' "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabDESCR.EOF
         cmbModalidade.AddItem TabDESCR!DESCRICAO
         cmbAuxLanc.AddItem TabDESCR!FORMAPAGTO_ID
         TabDESCR.MoveNext
      Wend
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_LostFocus"
End Sub

Private Sub cmbTipoVenda_Click()
'On Error GoTo ERRO_TRATA

   lblPRAZO.Caption = ""
   cmbAuxTIPOVENDA.ListIndex = cmbTipoVenda.ListIndex
   VALOR_ITEM_N = 0
   VALOR_ENTRADA = 0

   SETA_GRID
   If Trim(cmbAuxTIPOVENDA.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
   
      SQL = "select * from TIPOVENDA WITH (NOLOCK) "
      SQL = SQL & " where TIPOVENDA_id = " & cmbAuxTIPOVENDA.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         NUMR_PARCELA = 0
         DIAS_PRAZO = 0
         If Not IsNull(TabTemp!parcela) Then _
            NUMR_PARCELA = TabTemp!parcela
         If Not IsNull(TabTemp!PRAZO) Then _
            DIAS_PRAZO = TabTemp!PRAZO
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      Else
         MsgBox "Selecione tipo de venda."
         Exit Sub
   End If

   Frame1.Enabled = True
   txtSeq.SetFocus
   Exit Sub
   
   If cmbAuxTIPOVENDA.Text = 2 Then
      If TabTipovenda!parcela = 0 Then
         Frame1.Enabled = True
         txtSeq.SetFocus
      Else
         lblPRAZO.Caption = TabTipovenda!PRAZO & " dias"
         lblPRAZO.Refresh

         If NUMR_PARCELA = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue parcelas. " & TabTipovenda!TIPOVENDA_ID & " - " & TabTipovenda!DESCRICAO

            If TabTipovenda.State = 1 Then _
               TabTipovenda.Close

            cmbTipoVenda.SetFocus
            Exit Sub
         End If
         If DIAS_PRAZO = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue dias de vencimento. " & TabTipovenda!TIPOVENDA_ID & " - " & TabTipovenda!DESCRICAO

            If TabTipovenda.State = 1 Then _
               TabTipovenda.Close

            cmbTipoVenda.SetFocus
            Exit Sub
         End If
         NUMR_SEQ_N = 1
         CONT_N = 0

         'GERA TITULOS
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from NOTAENTRADA WITH (NOLOCK) "
         SQL = SQL & " where entrada_id = " & lblEntrada_ID.Caption
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            GERA_FATURAMENTO
            SETA_GRID
         End If
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         CONFIRMAR_RECEBIMENTO_PARCELADO
      End If
      Else   'se parcelas parametrizadas não for 0 entra aqui
         lblPRAZO.Caption = TabTemp!PRAZO & " dias"
         lblPRAZO.Refresh
         If NUMR_PARCELA = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue parcelas. " & TabTemp!TIPOVENDA_ID & " - " & TabTemp!DESCRICAO

            If TabTemp.State = 1 Then _
               TabTemp.Close

            cmbTipoVenda.SetFocus
            Exit Sub
         End If
         If DIAS_PRAZO = 0 Then
            MsgBox "Impossível faturar, tipo de venda não possue dias de vencimento. " & TabTemp!TIPOVENDA_ID & " - " & TabTemp!DESCRICAO

            If TabTemp.State = 1 Then _
               TabTemp.Close

            cmbTipoVenda.SetFocus
            Exit Sub
         End If

         CONT_N = 0

         'GERA TITULOS
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from NOTAENTRADA WITH (NOLOCK) "
   SQL = SQL & " where entrada_id = " & lblEntrada_ID.Caption
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            GERA_FATURAMENTO
            SETA_GRID
         End If
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         CONFIRMAR_RECEBIMENTO_PARCELADO
   End If

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_Click"
End Sub

Private Sub cmbTIPOVENDA_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbTipoVenda.SelStart = 0
   cmbTipoVenda.SelLength = Len(cmbTipoVenda)
   cmbTipoVenda.BackColor = &HC0FFFF

   cmbTipoVenda.SelStart = 0
   cmbTipoVenda.SelLength = Len(cmbTipoVenda)
   cmbTipoVenda.BackColor = &HC0FFFF
   cmbTipoVenda.Clear
   cmbAuxTIPOVENDA.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK) "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and pagar = 'true' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbTipoVenda.AddItem TabTemp!DESCRICAO
      cmbAuxTIPOVENDA.AddItem TabTemp!TIPOVENDA_ID
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MOSTRA_RODAPE "ESC - SAIR", "Selecione Tipo Venda", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_GotFocus"
End Sub

Private Sub cmbTIPOVENDA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_KeyPress"
End Sub

Private Sub cmbMODALIDADE_Click()
'On Error GoTo ERRO_TRATA
   cmbAuxLanc.ListIndex = cmbModalidade.ListIndex
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub cmbmodalidade_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbModalidade.SelStart = 0
   cmbModalidade.SelLength = Len(cmbModalidade)
   cmbModalidade.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmodalidade_GotFocus"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEmis.SelStart = 0
   txtDtEmis.SelLength = Len(txtDtEmis)
   txtDtEmis.BackColor = &HC0FFFF
   txtDtEmis.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_GotFocus"
End Sub

Private Sub txtDTEMIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtVenc.SetFocus
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

   txtDtEmis.BackColor = &HFFFFFF
   txtDtEmis.PromptInclude = True
   If Not IsDate(txtDtEmis.Text) Then
      txtDtEmis.PromptInclude = False
         txtDtEmis.Text = Date
      txtDtEmis.PromptInclude = True
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMIS_LostFocus"
End Sub

Private Sub txtDTVENC_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtVenc.PromptInclude = True
   txtDtVenc.SelStart = 0
   txtDtVenc.SelLength = Len(txtDtVenc)
   txtDtVenc.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_GotFocus"
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtDtVenc.PromptInclude = True
      If Not IsDate(txtDtVenc.Text) Then
         'MsgBox "Data Informada Inválida !!!"
         txtDtVenc.SetFocus
         txtDtVenc.PromptInclude = False
            txtDtVenc.Text = Date
         txtDtVenc.PromptInclude = True
         Exit Sub
      End If
      If txtSeq.Text = "" Then
         MsgBox "Seqüência deve ser gerada ou informada."
         txtSeq.SetFocus
         Exit Sub
      End If
      If cmbAuxLanc.Text = "" Then
         MsgBox "Selecione Forma de Pagamento !!!"
         cmbModalidade.SetFocus
         Exit Sub
      End If
      If txtValorItem.Text = "" Then
         MsgBox "Valor Incorreto !!!"
         txtValorItem.SetFocus
         Exit Sub
      End If
      txtDtEmis.PromptInclude = True
      If Not IsDate(txtDtEmis.Text) Then
         MsgBox "Data de emissão inválida !!!"
         txtDtVenc.SetFocus
         Exit Sub
      End If
      txtDtVenc.PromptInclude = True
      If CDate(txtDtVenc.Text) < CDate(txtDtEmis.Text) Then
         MsgBox "Data de vencimento não pode ser menor que data de emissão !!!"
         txtDtVenc.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      
      GRAVAR_TUDO
      
      LIMPA_BODY
      SETA_GRID
      txtSeq.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_KeyPress"
End Sub

Private Sub txtDTVENC_LostFocus()
'On Error GoTo ERRO_TRATA

   CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA
   txtDtVenc.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_LostFocus"
End Sub

Private Sub txtseq_LostFocus()
   txtSeq.BackColor = &HFFFFFF
End Sub

Private Sub txtValorItem_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorItem.SelStart = 0
   txtValorItem.SelLength = Len(txtValorItem)
   txtValorItem.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_GotFocus"
End Sub

Private Sub txtValorItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(txtValorItem.Text) <> "" Then
         KeyAscii = 0
         txtDtEmis.SetFocus
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_KeyPress"
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   SETA_GRID
   VALOR_DIFERENCA_N = 0

   txtSeq.SelStart = 0
   txtSeq.SelLength = Len(txtSeq)
   txtSeq.BackColor = &HC0FFFF

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
         
         SQL = "select max(seq) as ultimo_reg "
         SQL = SQL & " from ITEMLANCAMENTO "
         SQL = SQL & " where lancamento_id =  " & LANCAMENTO_ID_N
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then _
            If Not IsNull(TabLancamento!ultimo_reg) Then _
               NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         txtSeq.Text = NUMR_SEQ_N
         Else
            If TabLancamento.State = 1 Then _
               TabLancamento.Close

            SQL = "select * from ITEMLANCAMENTO "
            SQL = SQL & " where seq = " & txtSeq.Text
            SQL = SQL & " and lancamento_id =  " & LANCAMENTO_ID_N
            TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLancamento.EOF Then
               'valor lançamento
               txtValorItem.Text = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
               VALOR_DIFERENCA_N = TabLancamento!Valor_Item

               'descrição da modalidade
               If TabDESCR.State = 1 Then _
                  TabDESCR.Close
               SQL = "select descricao,formapagto_id from FORMAPAGTO WITH (NOLOCK) "
               SQL = SQL & " where formapagto_id = " & TabLancamento!FORMAPAGTO_ID
               SQL = SQL & " and status = 'true' "
               TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabDESCR.EOF Then
                  cmbModalidade.Text = TabDESCR!DESCRICAO
                  cmbAuxLanc.Text = TabDESCR!FORMAPAGTO_ID
               End If
               If TabDESCR.State = 1 Then _
                  TabDESCR.Close

               txtDtVenc.PromptInclude = False
               txtDtEmis.PromptInclude = False
               txtDtVenc.Text = TabLancamento!DT_VENCIMENTO
            End If
            If TabLancamento.State = 1 Then _
               TabLancamento.Close
      End If
      cmbModalidade.SetFocus
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

Private Sub txtRecebido_GotFocus()
'On Error GoTo ERRO_TRATA
   txtSeq.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrecebido_gotfocus"
End Sub

Private Sub txtValorItem_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorItem.Text <> "" Then _
      txtValorItem.Text = Format(txtValorItem.Text, strFormatacao2Digitos)
   txtValorItem.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_LostFocus"
End Sub

Private Sub txtVendaSemDesconto_GotFocus()
'On Error GoTo ERRO_TRATA
   txtSeq.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, " txtvendasemdesconto_GotFocus"
End Sub

'subrotinas
Private Sub GRAVAR_TUDO()
'On Error GoTo ERRO_TRATA

   Dim strSQL As String
   Dim strSQL2 As String

   VALOR_ITEM_N = 0 & txtValorItem.Text
   VALOR_DESCONTO_N = 0 & txtDesconto.Text

   INDR_RECEITA = 2

   If LANCAMENTO_ID_N <= 0 Then _
      LANCAMENTO_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   strSQL = "select * from LANCAMENTO WITH (NOLOCK) "
   strSQL = strSQL & " where lancamento_id = " & LANCAMENTO_ID_N
   TabLancamento.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabLancamento.EOF Then
      strSQL = "INSERT INTO LANCAMENTO "
      strSQL = strSQL & " ("
         strSQL = strSQL & " Lancamento_id, Numr_doc, dt_cad, Tipo_Lancamento, tipovenda_id,pessoa_id,estabelecimento_id"
      strSQL = strSQL & " ) "
      strSQL = strSQL & " VALUES ("
         strSQL = strSQL & LANCAMENTO_ID_N
         strSQL = strSQL & "," & NOTAENTRADA_ID_N
         strSQL = strSQL & ",'" & Now & "'"
         strSQL = strSQL & "," & INDR_RECEITA
         strSQL = strSQL & "," & cmbAuxTIPOVENDA.Text
         strSQL = strSQL & "," & PESSOA_ID_N
         strSQL = strSQL & "," & ESTABELECIMENTO_ID_N
      strSQL = strSQL & ")"
      CONECTA_RETAGUARDA.Execute strSQL
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close
  
   'ITENS
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   strSQL = "select * from ITEMLANCAMENTO WITH (NOLOCK) "
   strSQL = strSQL & " where seq = " & txtSeq.Text
   strSQL = strSQL & " and lancamento_id = " & LANCAMENTO_ID_N
   TabLANCAMENTOITEM.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      strSQL2 = "UPDATE ITEMLANCAMENTO SET "
         strSQL2 = strSQL2 & " usu_alt = " & USUARIO_ID_N
         strSQL2 = strSQL2 & ",Dt_Alt = '" & Date & "'"
         strSQL2 = strSQL2 & ",lancamento_id = " & LANCAMENTO_ID_N
         strSQL2 = strSQL2 & ",Valor_Item = " & Str(VALOR_ITEM_N)
         strSQL2 = strSQL2 & ",Status = 'A'"
         strSQL2 = strSQL2 & ",formapagto_id = " & cmbAuxLanc.Text
         strSQL2 = strSQL2 & ",DT_VENCIMENTO = '" & txtDtVenc.Text & "'"
      strSQL2 = strSQL2 & " Where Lancamento_id = " & LANCAMENTO_ID_N
      strSQL2 = strSQL2 & " and Seq = " & txtSeq.Text
      Else
         strSQL2 = "INSERT INTO ITEMLANCAMENTO "
            strSQL2 = strSQL2 & " (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, "
            strSQL2 = strSQL2 & " Numr_doc, NUMR_DP, seq, Valor_Item, Status, "
            strSQL2 = strSQL2 & " formapagto_id, DT_VENCIMENTO, CODG_USU_BAIXA, Acerto) "
         strSQL2 = strSQL2 & " VALUES ("
            strSQL2 = strSQL2 & USUARIO_ID_N
            strSQL2 = strSQL2 & ",'" & Now & "'"
            strSQL2 = strSQL2 & ",'" & Now & "'"
            strSQL2 = strSQL2 & "," & LANCAMENTO_ID_N
            strSQL2 = strSQL2 & "," & NOTAENTRADA_ID_N
            strSQL2 = strSQL2 & "," & txtNOTA.Text
            strSQL2 = strSQL2 & "," & txtSeq.Text
            strSQL2 = strSQL2 & "," & Str(VALOR_ITEM_N)
            strSQL2 = strSQL2 & ",'A'"
            strSQL2 = strSQL2 & "," & cmbAuxLanc.Text
            strSQL2 = strSQL2 & ",'" & DMA(txtDtVenc.Text) & "'"
            strSQL2 = strSQL2 & "," & USUARIO_ID_N
            strSQL2 = strSQL2 & "," & 1
         strSQL2 = strSQL2 & ")"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute strSQL2

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVAR_TUDO"
End Sub

Private Sub MATA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   SQL = "delete from ITEMLANCAMENTO "
   SQL = SQL & " where lancamento_id = " & LANCAMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from LANCAMENTO "
   SQL = SQL & " where lancamento_id = " & LANCAMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   BUSCA_LANCAMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_LANCAMENTO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Valor_Tot_n   As Double
   Dim VALOR_ITEM_N  As Double

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   Valor_Tot_n = 0
   ListaLanc.ListItems.Clear

   SQL = "select * from vwFATURAMENTO WITH (NOLOCK) "

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   SQL = SQL & " and numr_doc = " & NOTAENTRADA_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLANCAMENTOITEM.EOF
      LANCAMENTO_ID_N = 0 & TabLANCAMENTOITEM.Fields("lancamento_id").Value
      'sequencia
      Set item = ListaLanc.ListItems.Add(, "seq." & TabLANCAMENTOITEM!SEQ, TabLANCAMENTOITEM!SEQ)
      'numero documento
      item.SubItems(1) = "" & TabLANCAMENTOITEM!Numr_doc

      'valor lançamento
      PERC_DESCONTO_N = 0 & TabLANCAMENTOITEM!PERC_DESCONTO
      VALOR_ITEM_N = TabLANCAMENTOITEM!Valor_Item - (PERC_DESCONTO_N * TabLANCAMENTOITEM!Valor_Item / 100)
      Valor_Tot_n = VALOR_ITEM_N + Valor_Tot_n

      item.SubItems(2) = "" & Format(TabLANCAMENTOITEM!Valor_Item, strFormatacao2Digitos)
      item.SubItems(3) = "" & Format(0 & TabLANCAMENTOITEM!PERC_DESCONTO * TabLANCAMENTOITEM!Valor_Item / 100, strFormatacao2Digitos)
      item.SubItems(4) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      item.SubItems(5) = "" & Trim(TabLANCAMENTOITEM!FormaPagto)
      item.SubItems(6) = "" & Date
      item.SubItems(7) = "" & TabLANCAMENTOITEM!DT_VENCIMENTO
      item.SubItems(8) = "" & "00,00 %"

      If cmbAuxTIPOVENDA.Text <> "" Then
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from TIPOVENDA WITH (NOLOCK) "
         SQL = SQL & " where TIPOVENDA_id = " & cmbAuxTIPOVENDA.Text
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            If Not IsNull(TabAUX!PERC_JUROS) Then _
               item.SubItems(8) = "" & TabAUX!PERC_JUROS & "%"
         If TabAUX.State = 1 Then _
            TabAUX.Close
      End If
      TabLANCAMENTOITEM.MoveNext
   Wend
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   txtRecebido.Refresh
   txtVendaSemDesconto.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   lblPRAZO.Caption = ""
   cmbTipoVenda.Text = ""
   cmbAuxTIPOVENDA.Text = ""
   txtVendedor.Text = ""
   txtData.Text = ""
   txtVendaSemDesconto.Text = ""
   txtVendaComDesconto.Text = ""
   txtRecebido.Text = ""
   txtFornec.Text = ""
   cmbAuxLanc.Clear
   cmbModalidade.Clear
   txtValorItem.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtVenc.PromptInclude = False
   txtDtEmis.Text = ""
   txtDtVenc.Text = ""
   ListaLanc.ListItems.Clear
   txtSeq.Text = ""
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   VALOR_DIFERENCA_N = 0
   VALOR_ITEM_N = 0
   txtSeq.Text = ""
   cmbAuxLanc.Text = ""
   cmbModalidade.Text = ""
   txtValorItem.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtVenc.PromptInclude = False
   txtDtEmis.Text = ""
   txtDtVenc.Text = ""
   VALOR_TOTAL_LANÇADO = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub GERA_FATURAMENTO()
'On Error GoTo ERRO_TRATA

   INDR_RECEITA = 2
   NUMR_PARCELA = 0
   VALOR_DESCONTO_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK) "
   SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
         If TabTemp.Fields("contabiliza").Value = False Then
            If TabTemp.State = 1 Then _
               TabTemp.Close
            Exit Sub
         End If
      End If

      NUMR_PARCELA = TabTemp!parcela
      VALOR_DESCONTO_N = 0 & txtDesconto.Text
      VALOR_ITEM_N = 0
      DATA_INI = Date
      VALOR_ITEM_N = VALOR_TOTAL_N / NUMR_PARCELA

      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      'CABEÇA
      SQL = "select * from vwFornecedor WITH (NOLOCK) "
      SQL = SQL & " where fornecedor_id = " & TabCabeca!FORNECEDOR_ID
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      SQL = "select * from LANCAMENTO WITH (NOLOCK) "
      SQL = SQL & " where numr_doc = " & NOTAENTRADA_ID_N
      SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         LANCAMENTO_ID_N = TabLancamento!LANCAMENTO_ID
         Else
            LANCAMENTO_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

            SQL = "INSERT INTO LANCAMENTO "
            SQL = SQL & " ("
               SQL = SQL & " Lancamento_id, Numr_doc, dt_cad, Tipo_Lancamento, tipovenda_id,pessoa_id, estabelecimento_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & LANCAMENTO_ID_N
               SQL = SQL & "," & NOTAENTRADA_ID_N
               SQL = SQL & ",'" & Date & "'"
               SQL = SQL & "," & INDR_RECEITA
               SQL = SQL & "," & cmbAuxTIPOVENDA.Text
               SQL = SQL & "," & PESSOA_ID_N
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      VALOR_DESCONTO_N = VALOR_DESCONTO_N / NUMR_PARCELA
      While CONT_N <> NUMR_PARCELA
         GRAVA_LANÇAMENTO
         CONT_N = CONT_N + 1
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FATURAMENTO"
End Sub

Private Sub GRAVA_LANÇAMENTO()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 1

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO "
   SQL = SQL & " where lancamento_id = " & LANCAMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento!ultimo_reg) Then _
         NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   'ITENS
   DATA_INI = DATA_INI + TabTemp!PRAZO

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " where seq = " & NUMR_SEQ_N
   SQL = SQL & " and lancamento_id = " & LANCAMENTO_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SqL2 = "UPDATE ITEMLANCAMENTO SET usu_alt = " & USUARIO_ID_N & ", Dt_Alt = '" & Date & "', Dt_Cad = '" & Date & "',"
      SqL2 = SqL2 & " lancamento_id = " & LANCAMENTO_ID_N & ", Numr_doc = " & NOTAENTRADA_ID_N & ",  Numr_Dp = " & txtNOTA.Text & ", Seq = " & txtSeq.Text & ", Valor_Item = " & Str(VALOR_ITEM_N - (VALOR_ENTRADA / NUMR_PARCELA)) & ","
      SqL2 = SqL2 & " Status = '" & "A" & "', formapagto_id = " & TabTemp!FORMAPAGTO_ID & ", DT_VENCIMENTO = '" & DATA_INI & "' Where Lancamento_id = " & LANCAMENTO_ID_N & " and Seq = " & NUMR_SEQ_N
      CONECTA_RETAGUARDA.Execute SqL2
      Else
         SqL2 = "INSERT INTO ITEMLANCAMENTO (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, seq, Valor_Item, Status, formapagto_id, DT_VENCIMENTO, CODG_USU_BAIXA, Acerto) "
         SqL2 = SqL2 & " VALUES (" & USUARIO_ID_N & ",'" & Date & "','" & Date & "'," & LANCAMENTO_ID_N & "," & NOTAENTRADA_ID_N & "," & txtNOTA.Text & "," & NUMR_SEQ_N & "," & Str(VALOR_ITEM_N - (VALOR_ENTRADA / NUMR_PARCELA)) & ",'" & "A" & "'," & TabTemp!FORMAPAGTO_ID & ",'" & DATA_INI & "'," & USUARIO_ID_N & "," & 1 & ")"
         CONECTA_RETAGUARDA.Execute SqL2
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_LANÇAMENTO"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO
   VALOR_DESCONTO_N = 0 & txtDesconto.Text
   If Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos) >= (Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)) Then
      Msg = "Confirma recebimento ?"
      PERGUNTA Msg, vbYesNo + 32, "Recebimento Entrada de Mercadoria", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         Unload Me
         Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA"
End Sub

Private Sub CONFIRMAR_RECEBIMENTO_PARCELADO()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO
   If VALOR_TOTAL_LANÇADO >= (VALOR_TOTAL_N - VALOR_DESCONTO_N) Then
      Msg = "Confirma lançamento ?"
      PERGUNTA Msg, vbYesNo + 32, "Recebimento Entrada", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         Me.Hide
         Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub

Sub MOSTRA_NOTA()
'On Error GoTo ERRO_TRATA

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NOTAENTRADA WITH (NOLOCK) "
   SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      txtVendedor.Text = "" & TRAZ_NOME_USUARIO(TabNOTA.Fields("usuario_id").Value)
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Mask = "##############"
      txtNOTA.Text = "" & TabNOTA.Fields("numr_nota").Value
      txtSerie.Text = "" & TabNOTA.Fields("serie_nota").Value

      If TabTemp.State = 1 Then _
         TabTemp.Close
      SQL = "select descricao,cnpjcpf,pessoa_id from vwFornecedor WITH (NOLOCK) "
      SQL = SQL & " where fornecedor_id = " & TabNOTA!FORNECEDOR_ID
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         PESSOA_ID_N = TabTemp.Fields("pessoa_id").Value
         CNPJCPF_A = "" & Trim(TabTemp.Fields("cnpjcpf").Value)
         txtCNPJCPF.Text = CNPJCPF_A
         txtFornec.Text = "" & Trim(TabTemp.Fields("descricao").Value)
         txtFornec.Refresh
         Else
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Fornecedor não encontrado !!!"
            Unload Me
            Exit Sub
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If Not IsNull(txtCNPJCPF.Text) Then
          If Len(txtCNPJCPF.Text) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(txtCNPJCPF.Text) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.PromptInclude = False

      VALOR_TOTAL_N = 0
      SQL = "select sum(preco_custo*qtde_entrada) from NOTAENTRADAITEM WITH (NOLOCK) "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_TOTAL_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close

      VALOR_DESCONTO_N = 0
      SQL = "select valor_desconto from NOTAENTRADA WITH (NOLOCK) "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_DESCONTO_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      VALOR_ICMS_SUB_N = 0
      SQL = "select valor_icms_subst from NOTAENTRADA WITH (NOLOCK) "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_ICMS_SUB_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      VALOR_IPI_N = 0
      SQL = "select valor_ipi from NOTAENTRADA WITH (NOLOCK) "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_IPI_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select sum(valor_desconto*qtde_entrada) from NOTAENTRADAITEM WITH (NOLOCK) "
      SQL = SQL & " where entrada_id = " & NOTAENTRADA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_DESCONTO_N = 0 & TabTemp.Fields(0).Value + VALOR_DESCONTO_N
      If TabTemp.State = 1 Then _
         TabTemp.Close

      VALOR_TOTAL_N = (VALOR_TOTAL_N + VALOR_ICMS_SUB_N + VALOR_IPI_N - VALOR_DESCONTO_N)
      txtVendaSemDesconto.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
      If VALOR_DESCONTO_N > 0 Then _
         txtVendaComDesconto.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
      Else
         MsgBox "Nota fiscal não encontrada."
         Unload Me
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   BUSCA_LANCAMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub

Private Sub BUSCA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_LANÇADO = 0
   VALOR_RECEBIDO_N = 0
   txtRecebido.Text = ""

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select sum(valor_item) from ITEMLANCAMENTO "
   SQL = SQL & " where lancamento_id = " & LANCAMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento.Fields(0).Value) Then _
         VALOR_TOTAL_LANÇADO = 0 & TabLancamento.Fields(0).Value
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   txtRecebido.Text = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   txtRecebido.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_LANCAMENTO"
End Sub

Sub MOSTRA_NOTA_AVULSA()
'On Error GoTo ERRO_TRATA

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select ENTRADAESTOQUE.ENTRADAESTOQUE_ID, ENTRADAESTOQUE.ESTABELECIMENTO_ID, "
   SQL = SQL & " ENTRADAESTOQUE.FORNECEDOR_ID, ENTRADAESTOQUE.USUARIO_ID, ENTRADAESTOQUE.DT_CADASTRO, "
   SQL = SQL & " ENTRADAESTOQUE.SITUACAO , ENTRADAESTOQUE.DT_BAIXA, FORNECEDOR.PESSOA_ID, "
   SQL = SQL & " PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO"

   SQL = SQL & " from ENTRADAESTOQUE WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN FORNECEDOR WITH (NOLOCK) "
   SQL = SQL & " ON ENTRADAESTOQUE.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON FORNECEDOR.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where ENTRADAESTOQUE_ID = " & NOTAENTRADA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      txtVendedor.Text = "" & TRAZ_NOME_USUARIO(TabNOTA.Fields("usuario_id").Value)
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Mask = "##############"
      txtNOTA.Text = "" & TabNOTA.Fields("ENTRADAESTOQUE_ID").Value
      txtSerie.Text = "" '& TabNOTA.Fields("serie_nota").Value
      PESSOA_ID_N = TabNOTA.Fields("pessoa_id").Value
      CNPJCPF_A = "" & Trim(TabNOTA.Fields("cnpjcpf").Value)
      txtCNPJCPF.Text = CNPJCPF_A
      txtFornec.Text = "" & Trim(TabNOTA.Fields("descricao").Value)
      txtFornec.Refresh

      If Not IsNull(txtCNPJCPF.Text) Then
          If Len(txtCNPJCPF.Text) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(txtCNPJCPF.Text) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.PromptInclude = True

      VALOR_TOTAL_N = 0
      SQL = "select sum(preco*qtde) from ENTRADAESTOQUEitem WITH (NOLOCK) "
      SQL = SQL & " where ENTRADAESTOQUE_ID = " & NOTAENTRADA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         VALOR_TOTAL_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close

      VALOR_DESCONTO_N = 0
      VALOR_ICMS_SUB_N = 0
      VALOR_IPI_N = 0

      VALOR_TOTAL_N = (VALOR_TOTAL_N + VALOR_ICMS_SUB_N + VALOR_IPI_N - VALOR_DESCONTO_N)
      txtVendaSemDesconto.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
      If VALOR_DESCONTO_N > 0 Then _
         txtVendaComDesconto.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
      Else
         MsgBox "Nota fiscal não encontrada."
         Unload Me
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   BUSCA_LANCAMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub
