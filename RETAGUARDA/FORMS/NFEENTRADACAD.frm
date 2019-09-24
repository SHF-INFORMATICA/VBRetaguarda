VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNFEENTRADA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Produtos Estoque"
   ClientHeight    =   8685
   ClientLeft      =   1950
   ClientTop       =   2235
   ClientWidth     =   12180
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NFEENTRADACAD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12180
   WindowState     =   2  'Maximized
   Begin VB.Frame FraSeq 
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   50
      TabIndex        =   64
      Top             =   4440
      Width           =   12060
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   960
         MaxLength       =   30
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   3120
         MaxLength       =   100
         TabIndex        =   21
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10560
         MaxLength       =   9
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPrecoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   7800
         MaxLength       =   12
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtPrecoVenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   10560
         MaxLength       =   12
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtICMS_SUBST_Item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   25
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtIPI_Item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   960
         MaxLength       =   6
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtpercfrt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtICMS_Item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FTR%:"
         Height          =   240
         Left            =   5160
         TabIndex        =   72
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SUB%:"
         Height          =   240
         Left            =   3675
         TabIndex        =   71
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Preço Custo:"
         Height          =   240
         Left            =   6600
         TabIndex        =   70
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Preço Venda:"
         Height          =   240
         Left            =   9360
         TabIndex        =   69
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   240
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Qtde:"
         Height          =   240
         Left            =   10080
         TabIndex        =   67
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IPI%:"
         Height          =   240
         Left            =   360
         TabIndex        =   66
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ICMS%:"
         Height          =   240
         Left            =   1905
         TabIndex        =   65
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Importação de XML - NotaFiscal - Modelo 55"
      Height          =   735
      Left            =   5760
      TabIndex        =   59
      Top             =   720
      Width           =   6375
      Begin VB.TextBox txt_caminho_xml 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   62
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmd_explorer 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmd_ler_xml 
         Caption         =   "Preencher dados"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   200
         Width           =   1155
      End
      Begin VB.Label lbl_caminho_xml_demo 
         Alignment       =   2  'Center
         Caption         =   "Caminho XML:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Documento"
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   50
      TabIndex        =   56
      Top             =   720
      Width           =   5700
      Begin Threed.SSOption optEntrada 
         Height          =   270
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Entrada Nota Fiscal"
         Value           =   -1
      End
      Begin Threed.SSOption optDevolucao 
         Height          =   255
         Left            =   2880
         TabIndex        =   58
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Devolução de Entrada"
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   50
      TabIndex        =   45
      Top             =   3000
      Width           =   12060
      Begin VB.TextBox txtBaseCalculoICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPercICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtBaseICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7560
         MaxLength       =   12
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValorICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7560
         MaxLength       =   12
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPercICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7560
         MaxLength       =   12
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtValorOutras 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtUF 
         Height          =   375
         Left            =   3600
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor  Frete:"
         Height          =   240
         Left            =   9255
         TabIndex        =   55
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Base Cálculo do ICMS:"
         Height          =   240
         Left            =   150
         TabIndex        =   54
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS:"
         Height          =   240
         Left            =   855
         TabIndex        =   53
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percentual do ICMS:"
         Height          =   240
         Left            =   375
         TabIndex        =   52
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Base Cálculo do ICMS Substituto:"
         Height          =   240
         Left            =   4485
         TabIndex        =   51
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS Substituto:"
         Height          =   240
         Left            =   5190
         TabIndex        =   50
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percentual do ICMS Substituto:"
         Height          =   240
         Left            =   4710
         TabIndex        =   49
         Top             =   960
         Width           =   2745
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Outras:"
         Height          =   240
         Left            =   9180
         TabIndex        =   48
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do IPI:"
         Height          =   240
         Left            =   9270
         TabIndex        =   47
         Top             =   960
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   50
      TabIndex        =   32
      Top             =   1440
      Width           =   12060
      Begin VB.TextBox txtPedidoCompra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbCFOP 
         Height          =   360
         Left            =   1080
         TabIndex        =   9
         Text            =   "-- Selecione --"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8400
         MaxLength       =   12
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtValorTotalNota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   10560
         MaxLength       =   12
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbAuxTrans 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7515
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   360
         Left            =   7515
         TabIndex        =   10
         Top             =   1200
         Width           =   4395
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2925
         MaxLength       =   9
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   7800
         TabIndex        =   33
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtPEDIDO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5400
         MaxLength       =   9
         TabIndex        =   0
         Top             =   720
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtEntrada 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   6
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtDTEMISSAO 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "P.Compra:"
         Height          =   240
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CFOP:"
         Height          =   240
         Left            =   390
         TabIndex        =   43
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do Desconto:"
         Height          =   240
         Left            =   6720
         TabIndex        =   42
         Top             =   765
         Width           =   1665
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Nota:"
         Height          =   240
         Left            =   9600
         TabIndex        =   41
         Top             =   765
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Transportadora:"
         Height          =   240
         Left            =   6120
         TabIndex        =   39
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "S. NF:"
         Height          =   240
         Left            =   4020
         TabIndex        =   38
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nº.NF:"
         Height          =   240
         Left            =   2280
         TabIndex        =   37
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt.Entrada:"
         Height          =   240
         Left            =   2955
         TabIndex        =   36
         Top             =   765
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt.Emissão:"
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fornec:"
         Height          =   240
         Left            =   5205
         TabIndex        =   34
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":47C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":4809E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":483BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":4880E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":48C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":48F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":493D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFEENTRADACAD.frx":496F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   1270
      ButtonWidth     =   2884
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Confirma Nota"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir lançamento nota"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Relatório"
            Key             =   "print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fornecedor"
            Key             =   "CadFornec"
            Object.ToolTipText     =   "Devolucao de Mercadoria"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":49B4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":4AF76
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":4C005
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":4D26D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":4E96A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":4FC06
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":50D11
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":5210E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFEENTRADACAD.frx":532A8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTAITENS 
      Height          =   2670
      Left            =   45
      TabIndex        =   29
      Top             =   5595
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   4710
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   ""
      SmallIcons      =   ""
      ColHdrIcons     =   ""
      ForeColor       =   4194304
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Codg."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qtd."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Pr.Custo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pr.Venda"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Desconto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valr.Total "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "IPI"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "ICMS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ICMS SUB"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "FRETE"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.StatusBar stBarReq 
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   8280
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   8
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2222
            MinWidth        =   2222
            Text            =   "Diponível:"
            TextSave        =   "Diponível:"
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
            Object.Width           =   3152
            MinWidth        =   3152
            Text            =   "Preço Venda:"
            TextSave        =   "Preço Venda:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Key             =   "unitario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3510
            MinWidth        =   3510
            Text            =   "Preço Custo:"
            TextSave        =   "Preço Custo:"
            Key             =   "descvalr_unit"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Key             =   "desconto"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2294
            Text            =   "Valor Total:"
            TextSave        =   "Valor Total:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3134
            MinWidth        =   3134
            TextSave        =   ""
            Key             =   "total"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12180
      DesignHeight    =   8685
   End
   Begin MSComDlg.CommonDialog buscaxml 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmNFEENTRADA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim rstEmpresa             As New ADODB.Recordset
   Dim Valor_Tot_n            As Double
   Dim TIPO_ENTRADA_N         As Integer
   Dim VLR_FRETE_N            As Currency
   Dim VLR_OUTROS_N           As Currency
   Dim VALOR_TAXA_VAREJO_N    As Currency
   Dim VALOR_TAXA_ATACADO_N   As Currency
   Dim STATUS_NOTA_ENTRADA    As String
   Dim ALIQUOTA_FORNEC        As Integer
   Dim CFOP_DV_ENT_FE         As Integer
   Dim CFOP_DV_ENT_DE         As Integer
   Dim CFOP_ENTRADA_FE        As Integer
   Dim CFOP_ENTRADA_DE        As Integer

   Private gb_Recordset As New ADODB.Recordset ' yuri 11/05/2012
   'Importar XML
   'Private cNotaEntrada As New cNotaEntrada ' yuri 11/05/2012

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   frmNFEENTRADA.Top = 100
   TIPO_ENTRADA_N = 1
   Me.Caption = Me.Caption & Me.Name

   txtDTEMISSAO.PromptInclude = False
      txtDTEMISSAO.Text = Date
      txtDtEntrada.Text = Date
   txtDTEMISSAO.PromptInclude = True
   stBarReq.Refresh

   cmbTrans.Clear
   cmbAuxTrans.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TRANSPORTADORA "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by nome"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbTrans.AddItem Trim(TabTemp!Nome) & "-" & Trim(TabTemp!CGCCPF)
      cmbAuxTrans.AddItem Trim(TabTemp!TRANSP_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
   
   preencheComboCfop
   
   PegaDadosEmpresa 'ficar em memoria os cfops padroes e outros dados
   
   If Trim(txtPedido.Text) = "" Then
      GERA_NUMR_REQ
      txtPedido.Text = NUMR_REQ_N
   End If
   
   If Indr_Consulta = True Then
      'LIMPA_NOTA_ENTRADA
      txtPedido.Text = NUMR_REQ_N
      PROCURA_NOTA_ENTRADA
      NUMR_REQ_N = 0
   End If

   If Trim(cmbAuxTrans.Text) = "" Then _
      cmbAuxTrans.Text = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF9
         NUMR_REQ_N = 0
         LIMPA_NOTA_ENTRADA
         txtPedidoCompra.SetFocus
      Case vbKeyF10
         If Indr_Consulta = False Then
            If txtNota.Text <> "" And txtNota.Text <> "" And txtCNPJCPF.Text <> "" And txtValorTotalNota.Text <> "" Then
               If optDevolucao.Value = False Then
                  GRAVA_CABECA_NOTA

                  If optDevolucao.Value = False Then
                     If INDR_CONTROLA_ESTOQUE = True Then _
                        GRAVA_ESTOQUE

                     FINANCEIRO_FORM
                  End If

                  NUMR_REQ_N = 0
                  LIMPA_NOTA_ENTRADA
               Else
                  GRAVA_CABECA_NOTA
                  LIMPA_NOTA_ENTRADA
               End If
            End If
            MsgBox "Processo realizado com sucesso."
            txtPedido.Text = ""
            txtPedidoCompra.SetFocus
         End If
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub optDevolucao_Click(Value As Integer)
'On Error GoTo ERRO_TRATA

   TIPO_ENTRADA_N = 1
   STATUS_NOTA_ENTRADA = "D"
   cmbCFOP.Text = CFOP_DV_ENT_DE
   PegaDescricaoCFOP
   txtPedidoCompra.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optDevolucao_Click"
End Sub

Private Sub optEntrada_Click(Value As Integer)
'On Error GoTo ERRO_TRATA

    TIPO_ENTRADA_N = 1
    STATUS_NOTA_ENTRADA = "A"
    txtPedidoCompra.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optEntrada_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "fin"
         txtPedido.Text = 213
         FINANCEIRO_FORM
      Case "consultar"
         CRITERIO = ""
         frmNOTAENTRADACONSULTA.Show 1
         If CRITERIO <> "" Then
            txtPedido.Text = CRITERIO
            txtPedidoCompra.SetFocus
         End If
         CRITERIO = ""
      Case "nota"
         FORMULA_REL = ""
         If txtNota.Text <> "" Then
            FORMULA_REL = "{NF.NUMR_NOTA} = " & txtNota.Text
            FORMULA_REL = FORMULA_REL & " and {NFITEM.seq} < 10 "
         End If
ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
         Nome_Relatorio = "nf_notaentra.rpt"
         frmRELATORIO10.Show 1
      Case "print"
         FORMULA_REL = ""
         If txtNota.Text <> "" Then
         If txtNota.Text <> "" Then _
            FORMULA_REL = "{NOTAENTRADA.numr_nota} = " & txtNota.Text
         End If
ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
         Nome_Relatorio = "nf_entra.rpt"
         frmRELATORIO10.Show 1
      Case "matar"
         If txtPedido.Text <> "" Then
            If TabNOTA.State = 1 Then _
               TabNOTA.Close

            SQL = "select * from NOTAENTRADA "
            SQL = SQL & " where NUMR_PEDIDO_COMPRA = " & txtPedido.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabNOTA.EOF Then
               Msg = "Confirma exclusão de nota ?"
               PERGUNTA Msg & txtPedido.Text, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  CONECTA_RETAGUARDA.Execute "Delete from NotaEntrada Where empresa_id = " & EMPRESA_ID_N & " and NUMR_PEDIDO_COMPRA  = " & txtPedido.Text
                  NUMR_REQ_N = 0
                  LIMPA_NOTA_ENTRADA
               End If
            End If
            If TabNOTA.State = 1 Then _
               TabNOTA.Close
            Else: MsgBox "Informe um número de nota."
         End If
         txtPedidoCompra.SetFocus
      Case "voltar"
         Unload Me
      Case "gravar"
         If Indr_Consulta = False Then
            GRAVA_CABECA_NOTA

            If optDevolucao.Value = False Then
               If INDR_CONTROLA_ESTOQUE = True Then _
                  GRAVA_ESTOQUE

               FINANCEIRO_FORM
            End If

            NUMR_REQ_N = 0
            LIMPA_NOTA_ENTRADA
            MsgBox "Processo realizado com sucesso."
            txtPedidoCompra.SetFocus
         End If
      Case "limpar"
         NUMR_REQ_N = 0
         LIMPA_NOTA_ENTRADA
         txtPedidoCompra.SetFocus
      Case "CadFornec"
         frmCADASTROFORNECEDOR.Show 1
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtBaseCalculoICMS_GotFocus()
'On Error GoTo ERRO_TRATA

    txtBaseCalculoICMS.SelStart = 0
    txtBaseCalculoICMS.SelLength = Len(txtBaseCalculoICMS.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseCalculoICMS_GotFocus"
End Sub

'==================CNPJcpf
Private Sub txtcnpjcpf_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Fornecedores", "", "", ""
   
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      txtCNPJCPF.Mask = "##############"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         txtCNPJCPF.PromptInclude = False
         If CPF_N <> "" Then _
            txtCNPJCPF.Text = CPF_N
         CPF_N = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   Dim strTemp As String
   Dim dblTemp As Double

   txtCNPJCPF.PromptInclude = False

   If Trim(txtCNPJCPF.Text) = "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select CGC from EMPRESA "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select CGCcpf,nome,fornecedor_id from FORNECEDOR "
         SQL = SQL & " where CGCcpf = '" & Trim(TabTemp.Fields(0).Value) & "'"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            txtCNPJCPF.Text = "" & TabConsulta.Fields(0).Value
            txtNome.Text = "" & TabConsulta.Fields(1).Value
            FORNEC_ID_N = 0 & TabConsulta.Fields(2).Value
         End If

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close
      Else
         If optDevolucao.Value = True Then
            If TabEND.State = 1 Then _
               TabEND.Close

            SQL = "select * from FONE "
            SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
            TabEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabEND.EOF Then
               If TabEND.State = 1 Then _
                  TabEND.Close

               MsgBox "Imposivel Processar Nota Fiscal de Devolução sem o Numero do Telefone do Fornecedor!", vbExclamation, ""
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
            If TabEND.State = 1 Then _
               TabEND.Close
         End If

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

         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then _
               txtCNPJCPF.Mask = "###.###.###-##"
            If Len(txtCNPJCPF.Text) > 11 Then _
               txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
         txtCNPJCPF.Text = CRITERIO

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select * from FORNECEDOR "
         SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabCliente.EOF Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            Beep
            MsgBox "Fornecedor não Cadastrado.", vbOKOnly, "Atenção !!!"
            txtCNPJCPF.SetFocus
            Exit Sub
            Else
               If TabCliente!Nome <> "" Then _
                  txtNome.Text = TabCliente!Nome

               FORNEC_ID_N = TabCliente!fornecedor_id

               If Not IsNull(TabCliente!Status) Then
                  If TabCliente!Status <> "A" Then
                     MsgBox "Fornecedor Desativado, Favor Atualizar Cadastro!"
                     txtCNPJCPF.SetFocus
                     Exit Sub
                  End If
               End If

               If Not IsNull(TabCliente!aliquotafornec) Then
                  ALIQUOTA_FORNEC = TabCliente!aliquotafornec
                  Else: MsgBox "Fornecedor Com Aliquota de Icms Zerada, Favor Verificar Caso a Compra tenha ICMS Substituição!", vbExclamation
               End If

               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "Select * From ENDERECO "
               SQL = SQL & " Where PROP = '" & Trim(txtCNPJCPF.Text) & "'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  'Pegou o CEP do cliente
                  If Not IsNull(TabTemp!CEP) Then
                     dblTemp = TabTemp!CEP
                     Else 'Não tem cadastrado cep, impossivel fazer tributacao sem a uf
                        If TabTemp.State = 1 Then _
                           TabTemp.Close
                        MsgBox "O Cadastro do Fornecedor não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                        txtCNPJCPF.SetFocus
                        Exit Sub
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close

                  'Pegar a uf do cliente
                  If TabTemp.State = 1 Then _
                     TabTemp.Close

                  TabTemp.Open "Select * From CEP Where CEP=" & dblTemp, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     If Not IsNull(TabTemp!UF) Then
                        txtUF.Text = TabTemp!UF

                        If TabTemp.State = 1 Then _
                           TabTemp.Close

                        Else 'UF nao localizada
                           If TabTemp.State = 1 Then _
                              TabTemp.Close

                           MsgBox "O Cadastro do fornecedor não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                           txtCNPJCPF.SetFocus
                           Exit Sub
                     End If
                     Else
                        If TabTemp.State = 1 Then _
                           TabTemp.Close

                        MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais incompletos. Verique-os, principalmente o Estado(UF) da empresa"
                        txtCNPJCPF.SetFocus
                        Exit Sub
                  End If

                  If optDevolucao.Value = True Then
                     If txtUF.Text = "GO" Then
                        cmbCFOP.Text = CFOP_DV_ENT_DE
                        Else: cmbCFOP.Text = CFOP_DV_ENT_FE
                     End If
                     Else
                        If txtUF.Text = "GO" Then
                           cmbCFOP.Text = CFOP_ENTRADA_DE
                           Else: cmbCFOP.Text = CFOP_ENTRADA_FE
                        End If
                  End If

                     PegaDescricaoCFOP

                  Else
                     If TabTemp.State = 1 Then _
                        TabTemp.Close

                     MsgBox "O Sistema verificou que este Fornecedor esta com cadastrais incompletos."
                     'txtCNPJCPF.SetFocus
                     Exit Sub
               End If
         End If
   End If

   txtCNPJCPF.PromptInclude = False

   If Trim(txtNota.Text) <> "" And Trim(txtSerie.Text) <> "" And Trim(txtCNPJCPF.Text) <> "" Then _
      MOSTRA_NOTA

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_LostFocus"
End Sub

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub txtDescItem_GotFocus()
'On Error GoTo ERRO_TRATA

   'VALOR_DIFERENCA_N = "0" & txtDescItem.Text

    'txtDescItem.SelStart = 0
    'txtDescItem.SelLength = Len(txtDescItem.Text)
    
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescItem_GotFocus"
End Sub

Private Sub txtdescitem_LostFocus()
'On Error GoTo ERRO_TRATA
   
   'VALOR_ITEM_N = "0" & txtDescItem.Text
   'txtDescItem.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   'txtDescItem.Text = VALOR_ITEM_N
   If VALOR_ITEM_N > (txtQtd.Text * txtPrecoCusto.Text) Then
      MsgBox "O valor do desconto é maior que o valor total deste produto."
      'txtDescItem.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdescitem_LostFocus"
End Sub

Private Sub txtDesconto_GotFocus()
'On Error GoTo ERRO_TRATA

    txtDesconto.SelStart = 0
    txtDesconto.SelLength = Len(txtDesconto.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_GotFocus"
End Sub

Private Sub txtDTEMISSAO_Change()
'On Error GoTo ERRO_TRATA

   txtDtEntrada.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMISSAO_Change"
End Sub

Private Sub txtDTEMISSAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMISSAO_KeyPress"
End Sub

Private Sub txtdtemissao_LostFocus()
'On Error GoTo ERRO_TRATA

   If Not IsDate(txtDTEMISSAO.Text) Then
      txtDTEMISSAO.PromptInclude = False
         txtDTEMISSAO.Text = Date
      txtDTEMISSAO.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtemissao_LostFocus"
End Sub

Private Sub txtDtEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEntrada.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEntrada_GotFocus"
End Sub

Private Sub txtdtentrada_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtentrada_KeyPress"
End Sub

Private Sub cmbTrans_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxTrans.ListIndex = cmbTrans.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTrans_Click"
End Sub

Private Sub cmbtrans_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBaseCalculoICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtrans_KeyPress"
End Sub

Private Sub txtBaseCalculoICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseCalculoICMS_KeyPress"
End Sub

Private Sub txtBaseCalculoICMS_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtBaseCalculoICMS.Text = "" Then _
      txtBaseCalculoICMS.Text = 0
   If IsNumeric(txtBaseCalculoICMS.Text) Then
      VALOR_ITEM_N = txtBaseCalculoICMS.Text
      'txtBaseCalculoICMS.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      txtBaseCalculoICMS.Text = VALOR_ITEM_N
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseCalculoICMS_LostFocus"
End Sub

Private Sub txtICMS_Item_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtICMS_Item.Text = "" Then _
      txtICMS_Item.Text = 0

   VALOR_DIFERENCA_N = txtICMS_Item.Text

   txtICMS_Item.SelStart = 0
   txtICMS_Item.SelLength = Len(txtICMS_Item.Text)
            
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_Item_GotFocus"
End Sub

Private Sub txticms_item_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtICMS_Item.Text = "" Then _
      txtICMS_Item.Text = 0

   VALOR_ITEM_N = txtICMS_Item.Text
   'txtICMS_Item.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txticms_item_LostFocus"
End Sub

Private Sub txtIPI_Item_GotFocus()
'On Error GoTo ERRO_TRATA

    If txtIPI_Item.Text = "" Then
         txtIPI_Item.Text = 0
    End If
   VALOR_DIFERENCA_N = txtIPI_Item.Text
    
    txtIPI_Item.SelStart = 0
    txtIPI_Item.SelLength = Len(txtIPI_Item.Text)
    
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIPI_Item_GotFocus"
End Sub

Private Sub txtipi_item_LostFocus()
'On Error GoTo ERRO_TRATA

    If txtIPI_Item.Text = "" Then
         txtIPI_Item.Text = 0
    End If
    
   VALOR_ITEM_N = txtIPI_Item.Text
   'txtIPI_Item.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIPI_Item_LostFocus"
End Sub
Private Sub txtICMS_SUBST_Item_GotFocus()
'On Error GoTo ERRO_TRATA

    If txtICMS_SUBST_Item.Text = "" Then
         txtICMS_SUBST_Item.Text = 0
    End If
   VALOR_DIFERENCA_N = txtICMS_SUBST_Item.Text

    txtICMS_SUBST_Item.SelStart = 0
    txtICMS_SUBST_Item.SelLength = Len(txtICMS_SUBST_Item.Text)
            
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_SUBST_Item_GotFocus"
End Sub

Private Sub txtICMS_SUBST_Item_LostFocus()
'On Error GoTo ERRO_TRATA

    If txtICMS_SUBST_Item.Text = "" Then
         txtICMS_SUBST_Item.Text = 0
    End If
   
   VALOR_ITEM_N = txtICMS_SUBST_Item.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_SUBST_Item_LostFocus"
End Sub

Private Sub txtPedido_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then
      GERA_NUMR_REQ
      txtPedido.Text = NUMR_REQ_N
   End If
   PROCURA_NOTA_ENTRADA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_LostFocus"
End Sub

Private Sub txtpercfrt_GotFocus()
'On Error GoTo ERRO_TRATA

    If txtpercfrt.Text = "" Then
         txtpercfrt.Text = 0
    End If
   VALOR_DIFERENCA_N = txtpercfrt.Text

    txtpercfrt.SelStart = 0
    txtpercfrt.SelLength = Len(txtpercfrt.Text)
            
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercfrt_GotFocus"
End Sub

Private Sub txtpercfrt_LostFocus()
'On Error GoTo ERRO_TRATA

    If txtpercfrt.Text = "" Then
         txtpercfrt.Text = 0
    End If
   
   VALOR_ITEM_N = txtpercfrt.Text
   
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercfrt_LostFocus"
End Sub

Private Sub txtPercICMS_GotFocus()
'On Error GoTo ERRO_TRATA

    txtPercICMS.SelStart = 0
    txtPercICMS.SelLength = Len(txtPercICMS.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercICMS_GotFocus"
End Sub

Private Sub txtPrecoCusto_GotFocus()
'On Error GoTo ERRO_TRATA

    txtPrecoCusto.SelStart = 0
    txtPrecoCusto.SelLength = Len(txtPrecoCusto.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPrecoCusto_GotFocus"
End Sub

Private Sub txtQtd_GotFocus()
'On Error GoTo ERRO_TRATA

    If txtQtd.Text = "" Then
         txtQtd.Text = 0
    End If
    txtQtd.SelStart = 0
    txtQtd.SelLength = Len(txtQtd.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtd_GotFocus"
End Sub

Private Sub txtSERIE_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtSerie.Text) = "" Then _
      txtSerie.Text = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtSERIE_LostFocus"
End Sub

Private Sub txtPEDIDOCOMPRA_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedidoCompra.Text) <> "" Then
      MONTA_NOTA_ENTRADA
      If PedidoCompra = True Then ' So Grava aqui se o pedido de compra existir
         GRAVA_CABECA_NOTA
         GRAVA_CORPO_NOTA
      End If
      Else
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPEDIDOCOMPRA_LostFocus"
End Sub

Private Sub txtValorICMS_GotFocus()
'On Error GoTo ERRO_TRATA

    txtValorICMS.SelStart = 0
    txtValorICMS.SelLength = Len(txtValorICMS.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMS_GotFocus"
End Sub

Private Sub txtValorICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMS_KeyPress"
End Sub

Private Sub txtValorICMS_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorICMS.Text = "" Then _
      txtValorICMS.Text = 0
   If IsNumeric(txtValorICMS.Text) Then
      VALOR_ITEM_N = txtValorICMS.Text
      'txtValorICMS.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      txtValorICMS.Text = VALOR_ITEM_N
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMS_LostFocus"
End Sub

Private Sub txtpercICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBaseICMSSubst.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercICMS_KeyPress"
End Sub

Private Sub txtpercICMS_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtPercICMS.Text = "" Then _
      txtPercICMS.Text = 0
   If IsNumeric(txtPercICMS.Text) Then
      VALOR_ITEM_N = txtPercICMS.Text
      'txtPercICMS.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercICMS_LostFocus"
End Sub

Private Sub txtBaseICMSSubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorICMSSubst.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseICMSSubst_KeyPress"
End Sub

Private Sub txtBaseICMSSubst_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtBaseICMSSubst.Text = "" Then _
      txtBaseICMSSubst.Text = 0
   If IsNumeric(txtBaseICMSSubst.Text) Then
      VALOR_ITEM_N = txtBaseICMSSubst.Text
      'txtBaseICMSSubst.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseICMSSubst_LostFocus"
End Sub

Private Sub txtValorICMSSubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercICMSSubst.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMSSubst_KeyPress"
End Sub

Private Sub txtValorICMSSubst_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorICMSSubst.Text = "" Then _
      txtValorICMSSubst.Text = 0
   If IsNumeric(txtValorICMSSubst.Text) Then
      VALOR_ITEM_N = txtValorICMSSubst.Text
      'txtValorICMSSubst.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMSSubst_LostFocus"
End Sub

Private Sub txtPercICMSSubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorOutras.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercICMSSubst_KeyPress"
End Sub

Private Sub txtPercICMSSubst_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtPercICMSSubst.Text = "" Then _
      txtPercICMSSubst.Text = 0
   If IsNumeric(txtPercICMSSubst.Text) Then
      VALOR_ITEM_N = txtPercICMSSubst.Text
      'txtPercICMSSubst.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercICMSSubst_LostFocus"
End Sub

Private Sub txtValorOutras_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtFrete.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorOutras_KeyPress"
End Sub

Private Sub txtValorOutras_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorOutras.Text = "" Then _
      txtValorOutras.Text = 0
   If IsNumeric(txtValorOutras.Text) Then
      VALOR_ITEM_N = txtValorOutras.Text
      'txtValorOutras.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorOutras_LostFocus"
End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorIPI.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFrete_KeyPress"
End Sub

Private Sub txtFrete_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtFrete.Text = "" Then _
      txtFrete.Text = 0
   If IsNumeric(txtFrete.Text) Then
      VALOR_ITEM_N = txtFrete.Text
      'txtFrete.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFrete_LostFocus"
End Sub

Private Sub txtvaloripi_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorIPI_KeyPress"
End Sub

Private Sub txtValorIPI_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorIPI.Text = "" Then _
      txtValorIPI.Text = 0
   If IsNumeric(txtValorIPI.Text) Then
      VALOR_ITEM_N = txtValorIPI.Text
      
      
      'txtValorIPI.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorIPI_LostFocus"
End Sub

Private Sub txtdesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCFOP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_KeyPress"
End Sub

Private Sub txtDesconto_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtDesconto.Text = "" Then _
      txtDesconto.Text = 0
   If IsNumeric(txtDesconto.Text) Then
      VALOR_ITEM_N = txtDesconto.Text
      'txtDesconto.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_LostFocus"
End Sub

Private Sub txtDtEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   If Not IsDate(txtDtEntrada.Text) Then
      txtDtEntrada.PromptInclude = False
         txtDtEntrada.Text = Date
      txtDtEntrada.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEntrada_LostFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNota.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPedido_KeyPress"
End Sub

Private Sub txtPEDIDOCOMPRA_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtPEDIDOCOMPRA_KeyPress """
End Sub

Private Sub txtnota_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtnota_KeyPress"
   'Resume
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtserie_KeyPress"
   'Resume
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "F3 - Consulta Produtos", "F6 - Excluir Item", "F10 - Gravar", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtProduto.Text) <> "" And Trim(txtPedido.Text) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select i.entrada_id from NOTAENTRADAITEM i, NOTAENTRADA n "
            SQL = SQL & " where n.numr_pedido_compra = " & txtPedido.Text
            SQL = SQL & " and i.codg_prod = '" & Trim(txtProduto.Text) & "'"
            SQL = SQL & " and n.entrada_id = i.entrada_id "
            SQL = SQL & " and n.empresa_id = " & EMPRESA_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               Msg = "Confirma exclusão desse produto ?"
               PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  SQL = "delete  from NOTAENTRADAITEM "
                  SQL = SQL & " where entrada_id = " & TabTemp.Fields(0).Value
                  SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"
                  CONECTA_RETAGUARDA.Execute SQL
                  LIMPA_BODY
                  MOSTRA_TOTAL_NOTA
               End If
            End If
            txtProduto.SetFocus
         End If
      Case vbKeyF7
         frmPRODUTOCONSULTA.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtProduto.Text <> "" Then
         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select * from PRODUTO "
         SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabProduto.EOF Then
            MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção !!!"
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            txtProduto.SetFocus
            Exit Sub
            Else
               txtDesc = Trim(TabProduto!Descricao)
               stBarReq.Panels(2).Text = TabProduto!Qtde - TabProduto!QTDE_RETIDO
               QTDE_ESTOQUE = stBarReq.Panels(2).Text

               If Not IsNull(TabProduto!PRECO_CUSTO) Then
                  stBarReq.Panels(4).Text = TabProduto!PRECO_CUSTO
                  txtPrecoCusto.Text = TabProduto!PRECO_CUSTO
               End If

               If Not IsNull(TabProduto!PRECO_VENDA) Then
                  txtPrecoVenda.Text = TabProduto!PRECO_VENDA
                  stBarReq.Panels(6).Text = TabProduto!PRECO_VENDA
               End If

               If TabPedidoItem.State = 1 Then _
                  TabPedidoItem.Close

               SQL = "select i.PRECO_VENDA,i.QTD_ENTRADA, i.PRECO_CUSTO,  "
               SQL = SQL & " i.PERC_IPI, i.PERC_ICMS, i.VALOR_DESCONTO, "
               SQL = SQL & " i.PERC_ICMS_SUB "
               SQL = SQL & " from NOTAENTRADAITEM i, NOTAENTRADA n "

               SQL = SQL & " where n.empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and i.codg_prod = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and n.entrada_id = i.entrada_id "

               SQL = SQL & " and n.numr_nota = " & txtNota.Text
               SQL = SQL & " and n.serie_nota = " & txtSerie.Text
               SQL = SQL & " and n.fornecedor_id = " & FORNEC_ID_N

               If txtPedidoCompra.Text <> "" Then _
                  SQL = SQL & " and n.numr_pedido_compra = " & txtPedido.Text

               TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabPedidoItem.EOF Then
                  QTDE_PEDIDO = TabPedidoItem!QTD_ENTRADA
                  If IsNull(TabPedidoItem!PRECO_CUSTO) Then VALOR_ITEM_N = 0
                  VALOR_ITEM_N = TabPedidoItem!PRECO_CUSTO
                  VALOR_DIFERENCA_N = (TabPedidoItem!PRECO_CUSTO * TabPedidoItem!QTD_ENTRADA)
                  txtPrecoCusto.Text = TabPedidoItem!PRECO_CUSTO
                  txtPrecoVenda.Text = TabPedidoItem!PRECO_VENDA
                  txtQtd.Text = TabPedidoItem!QTD_ENTRADA
                  txtIPI_Item.Text = TabPedidoItem!PERC_IPI
                  txtICMS_Item.Text = TabPedidoItem!PERC_ICMS
                  txtICMS_SUBST_Item.Text = TabPedidoItem!PERC_ICMS_SUB
                  'txtDescItem.Text = TabPedidoItem!Valor_Desconto
               End If
               If TabPedidoItem.State = 1 Then _
                  TabPedidoItem.Close
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

         txtQtd.SetFocus
         Else: MsgBox "Informe código do produto."
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtqtd_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtQtd.Text = "" Then
         MsgBox "Quantidade não pode ser 0."
         txtQtd.SetFocus
         Exit Sub
      End If
      QTDE_PEDIDO = txtQtd.Text
      txtIPI_Item.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtqtd_KeyPress"
End Sub

Private Sub txtQtd_LostFocus()
'On Error GoTo ERRO_TRATA

    If txtQtd.Text = "" Then
         txtQtd.Text = 0
    End If
   VALOR_ITEM_N = txtQtd.Text
   txtQtd.Text = VALOR_ITEM_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtd_LostFocus"
End Sub

Private Sub txtipi_item_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0


      If txtIPI_Item.Text = "" Then
         txtIPI_Item.Text = 0
      End If
      
      If txtValorIPI.Text = "" Then
         txtValorIPI.Text = 0
      End If
      
      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      txtICMS_Item.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtipi_item_KeyPress"
End Sub

Private Sub txticms_item_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      
      If txtICMS_Item.Text = "" Then
         txtICMS_Item.Text = 0
      End If
      
      If txtValorICMS.Text = "" Then
         txtValorICMS.Text = 0
      End If

      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      txtICMS_SUBST_Item.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txticms_item_KeyPress"
End Sub

Private Sub txtICMS_SUBST_item_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      
      If txtICMS_SUBST_Item.Text = "" Then
         txtICMS_SUBST_Item.Text = 0
      End If
      
      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      txtpercfrt.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_SUBST_item_KeyPress"
End Sub

Private Sub txtpercfrt_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtpercfrt.Text = "" Then
         txtpercfrt.Text = 0
      End If
            
      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0
      
      'perc_frete_n = txtpercfrt.Text
      
      txtPrecoCusto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercfrt_KeyPress"
End Sub

Private Sub txtprecocusto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
'      If SINAL = "DC" Then
'         GRAVA_NOTA_Devolução
'         GRAVA_CABECA_Devolução
'         LIMPA_BODY
'         txtProduto.SetFocus
'      End If
      Indr_Consulta = False
      If Indr_Consulta = False Then
         GRAVA_NOTA_ENTRADA
         LIMPA_BODY
         txtProduto.SetFocus
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   
   TRATA_ERROS Err.Description, Me.Name, "txtprecocusto_KeyPress"
End Sub
Private Sub txtprecocusto_LostFocus()
'On Error GoTo ERRO_TRATA

      VALOR_ITEM_N = "0" & txtPrecoCusto.Text
      QTDE_PEDIDO = "0" & txtQtd.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprecocusto_LostFocus"
End Sub

Private Sub cmbCFOP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

    cmbCFOP.Refresh
   If KeyAscii = 13 Then
      'If cmbCFOP.ListIndex = -1 Then
      If cmbCFOP.Text = "" Then
         MsgBox "Selecione CFOP"
         
         cmbCFOP.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      cmbTrans.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfop_KeyPress"
End Sub

Private Sub cmbCFOP_GotFocus()
'On Error GoTo ERRO_TRATA

   If cmbCFOP.Text = "" Then 'nao escolheu cfop, coloca o default
      If txtUF.Text = "GO" Then
         cmbCFOP.Text = CFOP_ENTRADA_DE
         Else: cmbCFOP.Text = CFOP_ENTRADA_FE
      End If
      PegaDescricaoCFOP
   End If
    
    'cmbCFOP.Refresh
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCFOP_GotFocus"
End Sub

Private Sub cmbTrans_LostFocus()
   If Trim(cmbAuxTrans.Text) = "" Then _
      cmbAuxTrans.Text = 1
End Sub
'================================
Private Sub LIMPA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   SINAL = 0
   cmbCFOP.Text = ""
   txtUF.Text = ""
   CRITERIO = ""
   txtSerie.Text = ""
   txtProduto.Text = ""
   txtPedido.Text = ""
   txtDtEntrada.PromptInclude = False
   txtDtEntrada.Text = ""
   cmbTrans.Text = ""
   cmbAuxTrans.Text = ""
   txtPedidoCompra.Text = ""
   txtBaseCalculoICMS.Text = 0
   txtValorICMS.Text = 0
   txtPercICMS.Text = 0
   txtBaseICMSSubst.Text = 0
   txtValorICMSSubst.Text = 0
   txtPercICMSSubst.Text = 0
   txtValorOutras.Text = 0
   txtFrete.Text = 0
   txtValorIPI.Text = 0
   txtDesconto.Text = 0
   
   txtProduto.Text = ""
   txtNota.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtDtEntrada.PromptInclude = False
   txtDtEntrada.Text = ""
   stBarReq.Panels(8).Text = ""
   txtValorTotalNota.Text = ""
   LIMPA_BODY
   VALOR_TOTAL_N = 0
   VALOR_DESCONTO_N = 0
   Valor_Tot_n = 0
   
   'MOSTRA_TOTAL_NOTA
   
   stBarReq.Refresh
   SETA_GRID

Exit Sub
ERRO_TRATA:

   TRATA_ERROS Err.Description, Me.Name, "LIMPA_NOTA_ENTRADA"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   txtDesc.Text = ""
   txtQtd.Text = ""
   
   'txtIPI_Item.Text = Format(0, strFormatacao2Digitos)
   'txtICMS_Item.Text = Format(0, strFormatacao2Digitos)
   'txtDescItem.Text = Format(0, strFormatacao2Digitos)
   'txtPrecoCusto.Text = Format(0, strFormatacao2Digitos)
   'txtPrecoVenda.Text = Format(0, strFormatacao2Digitos)
   
   txtIPI_Item.Text = 0
   txtICMS_Item.Text = 0
   txtICMS_SUBST_Item.Text = 0
   txtpercfrt.Text = 0
   txtPrecoCusto.Text = 0
   txtPrecoVenda.Text = 0
   
   'cmbAuxCFOP.Text = ""
   'cmbCFOP.Text = ""
   stBarReq.Panels(2).Text = ""
   stBarReq.Panels(4).Text = ""
   stBarReq.Panels(6).Text = ""

   VALOR_ITEM_N = 0
   VALOR_DIFERENCA_N = 0
   
   'If SINAL = "DC" Then 'Devolução de Compra
   '   SETA_GRID_DEV
   '   Else: SETA_GRID
   'End If
   SETA_GRID
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub MOSTRA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   txtNota.Text = "" & TabNOTA!NUMR_NOTA
   txtSerie.Text = "" & TabNOTA!SERIE_NOTA
   NUMR_REQ_N = 0 & TabNOTA!NUMR_PEDIDO_COMPRA
   txtPedido.Text = 0 & TabNOTA!NUMR_PEDIDO_COMPRA

   txtCNPJCPF.PromptInclude = False
      FORNEC_ID_N = TabNOTA!fornecedor_id
      'TXTCNPJCPF.Text = TABNOTA!CGCCPF

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select nome, cgccpf from FORNECEDOR "
   SQL = SQL & " where fornecedor_id = " & FORNEC_ID_N
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      txtNome.Text = TabCliente!Nome

   txtCNPJCPF.Text = TabCliente!CGCCPF

   If TabCliente.State = 1 Then _
      TabCliente.Close

   txtDtEntrada.PromptInclude = False
      txtDtEntrada.Text = TabNOTA!DT_ENTRADA
   txtDtEntrada.PromptInclude = True
   txtDTEMISSAO.PromptInclude = False
      txtDTEMISSAO.Text = TabNOTA!DT_EMISSAO
   txtDTEMISSAO.PromptInclude = True
   txtBaseCalculoICMS.Text = Format(TabNOTA!BASE_CALC_ICMS, strFormatacao2Digitos)
   txtValorICMS.Text = Format(TabNOTA!VALOR_ICMS, strFormatacao2Digitos)
   txtPercICMS.Text = Format(TabNOTA!PERC_ICMS, strFormatacao2Digitos)
   txtBaseICMSSubst.Text = TabNOTA!BASE_CALC_ICMS_SUBST
   txtValorICMSSubst.Text = TabNOTA!VALOR_ICMS_SUBST
   txtPercICMSSubst.Text = TabNOTA!PERC_ICMS_SUBST
   txtValorOutras.Text = TabNOTA!VALOR_OUTRAS
   txtFrete.Text = TabNOTA!valor_frete
   txtValorIPI.Text = TabNOTA!VALOR_IPI
   txtDesconto.Text = TabNOTA!Valor_Desconto

   If Not IsNull(TabNOTA!TRANSP_ID) Then
      cmbAuxTrans.Text = TabNOTA!TRANSP_ID

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TRANSPORTADORA "
      SQL = SQL & " where transp_id = " & cmbAuxTrans.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         cmbTrans.Text = Trim(TabTemp!Nome) & "-" & Trim(TabTemp!CGCCPF)
         'cmbAuxTrans.Text= Trim(TABTEMP!transp_id)
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   
   cmbCFOP.Text = Trim(TabNOTA!CFOP)
   PegaDescricaoCFOP

   SETA_GRID
   MOSTRA_TOTAL_NOTA
   stBarReq.Refresh

Exit Sub
ERRO_TRATA:
   
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_NOTA_ENTRADA"
End Sub

Private Sub GRAVA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   If optEntrada.Value = True Then
    If txtNota.Text = "" Then
       MsgBox "Informe número de nota."
       txtNota.SetFocus
       Exit Sub
    End If
   End If
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe Fornecedor."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   txtDtEntrada.PromptInclude = True
   If Not IsDate(txtDtEntrada.Text) Then
      MsgBox "Informe data de entrada."
      txtDtEntrada.SetFocus
      Exit Sub
   End If
   txtDTEMISSAO.PromptInclude = True
   If Not IsDate(txtDTEMISSAO.Text) Then
      MsgBox "Informe data de emissão."
      txtDTEMISSAO.SetFocus
      Exit Sub
   End If
   If cmbCFOP.Text = "" Then
      MsgBox "Selecione CFOP"
      cmbCFOP.SetFocus
      Exit Sub
   End If
   If txtProduto.Text = "" Then
      MsgBox "Seqüência sem codigo de Produto.", vbOKOnly, "Atenção !!!"
      txtProduto.SetFocus
      Exit Sub
   End If
   If Not IsNull(txtPrecoCusto.Text) Then
      If txtPrecoCusto.Text <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção !!!"
         txtProduto.SetFocus
         Exit Sub
      End If
   End If
   If txtQtd.Text = "" Then
      MsgBox "Quantidade inválida.", vbOKOnly, "Atenção !!!"
      txtQtd.SetFocus
      Exit Sub
   End If
   If txtPrecoCusto.Text = "" Then
      MsgBox "Valor total inválido."
      txtPrecoCusto.SetFocus
      Exit Sub
   Else
      VALOR_ITEM_N = "0" & txtPrecoCusto.Text
      VALOR_TOTAL_N = (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N + VALOR_TOTAL_N
      MOSTRA_TOTAL_NOTA
      stBarReq.Refresh
   End If
   
   GRAVA_CABECA_NOTA
   GRAVA_CORPO_NOTA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_NOTA_ENTRADA"
End Sub

Private Sub GRAVA_NOTA_Devolução()
'On Error GoTo ERRO_TRATA

   If txtNota.Text = "" Then
      MsgBox "Informe número de nota."
      txtNota.SetFocus
      Exit Sub
   End If
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe Fornecedor."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   txtDtEntrada.PromptInclude = True
   If Not IsDate(txtDtEntrada.Text) Then
      MsgBox "Informe data de entrada."
      txtDtEntrada.SetFocus
      Exit Sub
   End If
   txtDTEMISSAO.PromptInclude = True
   If Not IsDate(txtDTEMISSAO.Text) Then
      MsgBox "Informe data de emissão."
      txtDTEMISSAO.SetFocus
      Exit Sub
   End If
   If cmbCFOP.Text = "" Then
      MsgBox "Selecione CFOP"
      cmbCFOP.SetFocus
      Exit Sub
   End If
   If txtProduto.Text = "" Then
      MsgBox "Seqüência sem codigo de Produto.", vbOKOnly, "Atenção !!!"
      txtProduto.SetFocus
      Exit Sub
   End If
   If Not IsNull(txtPrecoCusto.Text) Then
      If txtPrecoCusto.Text <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção !!!"
         txtProduto.SetFocus
         Exit Sub
      End If
   End If
   If txtQtd.Text = "" Then
      MsgBox "Quantidade inválida.", vbOKOnly, "Atenção !!!"
      txtQtd.SetFocus
      Exit Sub
   End If
   If txtPrecoCusto.Text = "" Then
      MsgBox "Valor total inválido."
      txtPrecoCusto.SetFocus
      Exit Sub
   Else
      VALOR_ITEM_N = "0" & txtPrecoCusto.Text
      'VALOR_TOTAL_N = (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N + VALOR_TOTAL_N
      MOSTRA_TOTAL_NOTA_DEV
      stBarReq.Refresh
   End If
   GRAVA_CORPO_DEV_QTD

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_NOTA_Devolução"
End Sub

Private Sub GRAVA_CABECA_NOTA()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   txtDtEntrada.PromptInclude = True
   txtDTEMISSAO.PromptInclude = True

   If Not IsDate(txtDtEntrada.Text) Then _
      txtDtEntrada.Text = Date

   If Trim(txtNota.Text) = "" Then _
      txtNota.Text = 0

   If Trim(cmbAuxTrans.Text) = "" Then _
      cmbAuxTrans.Text = 1

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where numr_pedido_compra = " & txtPedido.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      NUMR_ID_N = TabNOTA!entrada_id

      SQL = "UPDATE NOTAENTRADA SET "
      SQL = SQL & " TRANSP_ID = " & Trim(cmbAuxTrans.Text)
      SQL = SQL & ", TIPOENTRADA_id = " & TIPO_ENTRADA_N
      SQL = SQL & ", CODG_USU = " & CODG_USU_N
      SQL = SQL & ", NUMR_NOTA = " & txtNota.Text
      SQL = SQL & ", serie_nota = " & txtSerie.Text
      SQL = SQL & ", numr_pedido_compra = " & txtPedido.Text
      SQL = SQL & ", fornecedor_id = " & FORNEC_ID_N
      SQL = SQL & ", dt_entrada = '" & DMA(txtDtEntrada.Text) & "'"
      SQL = SQL & ", dt_Emissao = '" & DMA(txtDTEMISSAO.Text) & "'"
      SQL = SQL & ", Status = '" & STATUS_NOTA_ENTRADA & "'"
      SQL = SQL & ", BASE_CALC_ICMS = " & tpMoeda(txtBaseCalculoICMS.Text)
      SQL = SQL & ", VALOR_ICMS = " & tpMoeda(txtValorICMS.Text)
      SQL = SQL & ", PERC_ICMS = " & tpMoeda(txtPercICMS.Text)
      SQL = SQL & ", BASE_CALC_ICMS_SUBST = " & tpMoeda(txtBaseICMSSubst.Text)
      SQL = SQL & ", VALOR_ICMS_SUBST = " & tpMoeda(txtValorICMSSubst.Text)
      SQL = SQL & ", PERC_ICMS_SUBST = " & tpMoeda(txtPercICMSSubst.Text)
      SQL = SQL & ", VALOR_OUTRAS = " & tpMoeda(txtValorOutras.Text)
      SQL = SQL & ", valor_frete = " & tpMoeda(txtFrete.Text)
      SQL = SQL & ", VALOR_IPI = " & tpMoeda(txtValorIPI.Text)
      SQL = SQL & ", Valor_Desconto = " & tpMoeda(txtDesconto.Text)
      SQL = SQL & ", CFOP = " & Left(cmbCFOP.Text, 4)
      SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and numr_pedido_compra = " & txtPedido.Text
      Else
         NUMR_ID_N = MAX_ID("entrada_id", "notaentrada", "", "", "", "")

         SQL = "INSERT INTO NOTAENTRADA "
            SQL = SQL & " (Empresa_id, entrada_id, numr_pedido_compra, NUMR_NOTA, "
            SQL = SQL & " CODG_USU, SERIE_NOTA, fornecedor_id, dt_entrada, dt_Emissao, "
            SQL = SQL & " Status, BASE_CALC_ICMS, VALOR_ICMS , PERC_ICMS, BASE_CALC_ICMS_SUBST, "
            SQL = SQL & " VALOR_ICMS_SUBST, PERC_ICMS_SUBST,  VALOR_OUTRAS, valor_frete, "
            SQL = SQL & " VALOR_IPI, Valor_Desconto, CFOP, TIPOENTRADA_ID, TRANSP_ID) "
         SQL = SQL & " VALUES ("
            SQL = SQL & EMPRESA_ID_N
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & txtPedido.Text
            SQL = SQL & "," & txtNota.Text
            SQL = SQL & "," & CODG_USU_N
            SQL = SQL & "," & txtSerie.Text
            SQL = SQL & "," & FORNEC_ID_N
            SQL = SQL & ",'" & DMA(txtDtEntrada.Text) & "'"
            SQL = SQL & ",'" & DMA(Date) & "'"
            SQL = SQL & ",'" & STATUS_NOTA_ENTRADA & "'"
            SQL = SQL & "," & tpMoeda(txtBaseCalculoICMS.Text)
            SQL = SQL & "," & tpMoeda(txtValorICMS.Text)
            SQL = SQL & "," & tpMoeda(txtPercICMS.Text)
            SQL = SQL & "," & tpMoeda(txtBaseICMSSubst.Text)
            SQL = SQL & "," & tpMoeda(txtValorICMSSubst.Text)
            SQL = SQL & "," & tpMoeda(txtPercICMSSubst.Text)
            SQL = SQL & "," & tpMoeda(txtValorOutras.Text)
            SQL = SQL & "," & tpMoeda(txtFrete.Text)
            SQL = SQL & "," & tpMoeda(txtValorIPI.Text)
            SQL = SQL & "," & tpMoeda(txtDesconto.Text)
            SQL = SQL & "," & Left(cmbCFOP.Text, 4)
            SQL = SQL & "," & TIPO_ENTRADA_N
            SQL = SQL & "," & Trim(cmbAuxTrans.Text)
         SQL = SQL & ")"
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA_NOTA"
End Sub

Private Sub GRAVA_CORPO_NOTA()
'On Error GoTo ERRO_TRATA

   ' se fora false e porque nao existe pedido de compra
   If PedidoCompra = False Then _
      GoTo Grava_Itens_Nota

   If Trim(txtPedidoCompra.Text) <> "" Then
      NUMR_SEQ_N = NUMR_SEQ_N + 1

      If tabPEDIDOCOMPRAITEM.State = 1 Then _
         tabPEDIDOCOMPRAITEM.Close

      SQL = "select * from PEDIDOCOMPRAITEM i "
      SQL = SQL & " where i.pedido = " & txtPedidoCompra.Text
      SQL = SQL & " and i.qtd > " & 0
      SQL = SQL & " order by i.sequencia"
      tabPEDIDOCOMPRAITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not tabPEDIDOCOMPRAITEM.EOF
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SqL2 = "select max(seq) from NOTAENTRADAITEM "
         SqL2 = SqL2 & " where entrada_id = " & NUMR_ID_N
         TabTemp.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp.Fields(0).Value) Then
               NUMR_SEQ_N = TabTemp.Fields(0).Value + 1
               Else: NUMR_SEQ_N = 1
            End If
            Else: NUMR_SEQ_N = 1
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         PRODUTO_ID_N = 0

         SQL = "select produto_id from PRODUTO "
         SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and situacao <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            PRODUTO_ID_N = TabTemp.Fields(0).Value
         End If

         If TabTemp.State = 1 Then
            TabTemp.Close
         End If

         If TabItemNota.State = 1 Then _
            TabItemNota.Close

         SqL2 = "select * from NOTAENTRADAITEM "
         SqL2 = SqL2 & " where entrada_id = " & NUMR_ID_N
         SqL2 = SqL2 & " and codg_prod = '" & tabPEDIDOCOMPRAITEM!Produto & "'"
         TabItemNota.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
         If TabItemNota.EOF Then
            SQL = "INSERT INTO NOTAENTRADAITEM "
               SQL = SQL & " (entrada_id,seq,Codg_Prod,preco_custo,PRECO_venda,qtd_Entrada,Status,CFOP,produto_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & NUMR_SEQ_N
               SQL = SQL & ",'" & Trim(tabPEDIDOCOMPRAITEM!Produto) & "'"
               SQL = SQL & "," & tpMoeda(tabPEDIDOCOMPRAITEM!Preco)
               SQL = SQL & "," & tpMoeda(tabPEDIDOCOMPRAITEM!Preco)
               SQL = SQL & "," & tpMoeda(tabPEDIDOCOMPRAITEM!qtd)
               SQL = SQL & ",'" & STATUS_NOTA_ENTRADA & "'"
               SQL = SQL & "," & Left(cmbCFOP.Text, (InStr(1, cmbCFOP.Text, "-") - 1))
               SQL = SQL & "," & PRODUTO_ID_N
            SQL = SQL & ")"
            Else
               SQL = "UPDATE NOTAENTRADAITEM SET "
               SQL = SQL & " Codg_Prod = '" & tabPEDIDOCOMPRAITEM!Produto & "'"
               SQL = SQL & ", preco_custo = " & tpMoeda(tabPEDIDOCOMPRAITEM!Preco)
               SQL = SQL & " PRECO_venda = " & tpMoeda(tabPEDIDOCOMPRAITEM!Preco)
               SQL = SQL & ", qtd_Entrada = " & tpMoeda(tabPEDIDOCOMPRAITEM!qtd)
               SQL = SQL & ", Status = '" & STATUS_NOTA_ENTRADA & "'"
               SQL = SQL & ", CFOP =  " & Left(cmbCFOP.Text, (InStr(1, cmbCFOP.Text, "-") - 1))
               SQL = SQL & " Where entrada_id = " & NUMR_ID_N
               SQL = SQL & " and seq = " & TabItemNota.Fields("seq").Value
         End If
         If TabItemNota.State = 1 Then _
            TabItemNota.Close

         CONECTA_RETAGUARDA.Execute SQL

         tabPEDIDOCOMPRAITEM.MoveNext
      Wend
      If tabPEDIDOCOMPRAITEM.State = 1 Then _
         tabPEDIDOCOMPRAITEM.Close
      Else

Grava_Itens_Nota:

         QTDE_PEDIDO = txtQtd.Text
         NUMR_SEQ_N = 1

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select max(seq) from NOTAENTRADAITEM "
         SQL = SQL & " where entrada_id = " & NUMR_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               NUMR_SEQ_N = TabTemp.Fields(0).Value + 1
         If TabTemp.State = 1 Then _
            TabTemp.Close


         PRODUTO_ID_N = 0

         SQL = "select produto_id from PRODUTO "
         SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and situacao <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            PRODUTO_ID_N = TabTemp.Fields(0).Value
         End If

         If TabTemp.State = 1 Then
            TabTemp.Close
         End If

         If TabItemNota.State = 1 Then _
            TabItemNota.Close

         SQL = "select * from NOTAENTRADAITEM "
         SQL = SQL & " where entrada_id = " & NUMR_ID_N
         SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"
         TabItemNota.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItemNota.EOF Then
            SQL = "INSERT INTO NOTAENTRADAITEM "
               SQL = SQL & " (entrada_id, seq, Codg_Prod, "
               SQL = SQL & " preco_custo, PRECO_venda, qtd_Entrada, Status, CFOP, PERC_IPI, "
               SQL = SQL & " PERC_ICMS_SUB, PERC_ICMS, perc_frete,produto_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & NUMR_SEQ_N
               SQL = SQL & ",'" & Trim(txtProduto.Text) & "'"
               SQL = SQL & "," & tpMoeda(txtPrecoCusto.Text)
               SQL = SQL & "," & tpMoeda(txtPrecoVenda.Text)
               SQL = SQL & "," & tpMoeda(QTDE_PEDIDO)
               SQL = SQL & ",'A'"
               SQL = SQL & "," & Left(cmbCFOP.Text, 4)
               SQL = SQL & "," & tpMoeda(txtIPI_Item.Text)
               SQL = SQL & "," & tpMoeda(txtICMS_SUBST_Item.Text)
               SQL = SQL & "," & tpMoeda(txtICMS_Item.Text)
               SQL = SQL & "," & tpMoeda(txtpercfrt.Text)
               SQL = SQL & "," & PRODUTO_ID_N
            SQL = SQL & ")"
            Else
               SQL = "UPDATE NOTAENTRADAITEM SET "
               
               SQL = SQL & " seq = " & NUMR_SEQ_N
               SQL = SQL & ", Codg_Prod = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & ", preco_custo = " & tpMoeda(txtPrecoCusto.Text)
               SQL = SQL & ", PRECO_venda = " & tpMoeda(txtPrecoVenda.Text)
               SQL = SQL & ", qtd_Entrada = " & tpMoeda(QTDE_PEDIDO)
               SQL = SQL & ", Status = 'A'"
               SQL = SQL & ", CFOP = " & Left(cmbCFOP.Text, (InStr(1, cmbCFOP.Text, "-") - 1))
               SQL = SQL & ", PERC_IPI = " & tpMoeda(txtIPI_Item.Text)
               SQL = SQL & ", PERC_ICMS_SUB = " & tpMoeda(txtICMS_SUBST_Item.Text)
               SQL = SQL & ", PERC_ICMS = " & tpMoeda(txtICMS_Item.Text)
               SQL = SQL & ", perc_frete = " & tpMoeda(txtpercfrt.Text)

               SQL = SQL & "  Where entrada_id = " & NUMR_ID_N
               SQL = SQL & " and codg_prod = '" & Trim(txtProduto.Text) & "'"
         End If
         If TabItemNota.State = 1 Then _
            TabItemNota.Close

         CONECTA_RETAGUARDA.Execute SQL

         SETA_GRID
         MOSTRA_TOTAL_NOTA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CORPO_NOTA"
End Sub

Private Sub MOSTRA_TOTAL_NOTA()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" Then _
      NUMR_REQ_N = txtPedido.Text

   VALOR_TOTAL_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum((i.PRECO_CUSTO * i.QTD_ENTRADA)-isnull(i.VALOR_DESCONTO,0)) as ValorTotal "
   SQL = SQL & " from NOTAENTRADAITEM i, NOTAENTRADA n "
   SQL = SQL & " where n.numr_pedido_compra = " & NUMR_REQ_N
   SQL = SQL & " and n.entrada_id = i.entrada_id "
   SQL = SQL & " and n.empresa_id = " & EMPRESA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_TOTAL_N = TabTemp!valortotal
   If TabTemp.State = 1 Then _
      TabTemp.Close

   VALOR_DESCONTO_N = 0

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select isnull(valor_desconto,0) from NOTAENTRADA "
   SQL = SQL & " where NUMR_PEDIDO_COMPRA = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVENDEDOR.EOF Then _
      VALOR_DESCONTO_N = 0 & TabVENDEDOR.Fields(0).Value + VALOR_DESCONTO_N
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   'stBarReq.Panels(8).Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, "currency")
   txtValorTotalNota.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
   stBarReq.Panels(8).Text = txtValorTotalNota.Text
   stBarReq.Refresh

   'txtValorTotalNota.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, "currency")
   txtValorTotalNota.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAL_NOTA"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   LISTAITENS.ListItems.Clear
   If txtNota.Text <> "" Then
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select * from NOTAENTRADAITEM i, NOTAENTRADA n "
      SQL = SQL & " where n.numr_pedido_compra = " & txtPedido.Text
      SQL = SQL & " and n.entrada_id = i.entrada_id "
      SQL = SQL & " and n.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " order by i.seq desc"
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabPedidoItem.EOF
         Set Item = LISTAITENS.ListItems.Add(, "seq." & TabPedidoItem!SEQ, TabPedidoItem!SEQ)
         Item.SubItems(1) = TabPedidoItem!Codg_Prod

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select descricao from PRODUTO "
         SQL = SQL & " where codg_produto = '" & TabPedidoItem!Codg_Prod & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            Item.SubItems(2) = TabTemp!Descricao
         If TabTemp.State = 1 Then _
            TabTemp.Close

         Item.SubItems(3) = TabPedidoItem!QTD_ENTRADA
         'ITEM.SubItems(4) = Format(TabPedidoItem!PRECO_CUSTO, strFormatacao2Digitos)
         'ITEM.SubItems(5) = Format(TabPedidoItem!PRECO_VENDA, strFormatacao2Digitos)
         
         Item.SubItems(4) = TabPedidoItem!PRECO_CUSTO
         Item.SubItems(5) = TabPedidoItem!PRECO_VENDA
         'ITEM.SubItems(6) = Format(TabPedidoItem.Fields("i.Valor_Desconto").Value, strFormatacao2Digitos)
         Item.SubItems(6) = TabPedidoItem!Valor_Desconto
         'ITEM.SubItems(7) = Format((TabPedidoItem!PRECO_CUSTO - TabPedidoItem.Fields("i.Valor_Desconto").Value) * TabPedidoItem!QTD_ENTRADA, strFormatacao2Digitos)
         'Debug.Print "(" & TabPedidoItem!PRECO_CUSTO & "-" & TabPedidoItem.Fields("i.Valor_Desconto").Value & ")" & "*" & TabPedidoItem!QTD_ENTRADA
         Item.SubItems(7) = Format((TabPedidoItem!PRECO_CUSTO * TabPedidoItem!QTD_ENTRADA) - TabPedidoItem!Valor_Desconto, strFormatacao2Digitos)
         
         Item.SubItems(8) = TabPedidoItem!PERC_IPI
         Item.SubItems(9) = TabPedidoItem!PERC_ICMS
         Item.SubItems(10) = TabPedidoItem!PERC_ICMS_SUB
         If Not IsNull(TabPedidoItem!perc_frete) Then
            Item.SubItems(11) = TabPedidoItem!perc_frete
         End If
         TabPedidoItem.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GRAVA_ESTOQUE()
'On Error GoTo ERRO_TRATA
   
   Dim intRetorno  As Integer
   Dim intRetorno_2 As Integer
   
   Dim VALOR_IPI_N As Double
   Dim VALOR_ICMS_N As Double
   Dim VALOR_DIF_ICMS_SUB As Currency
   Dim VALOR_ACUMULADO_SUB As Currency
   Dim PERC_ICMS_SUB As Currency
   Dim VALOR_ICMS_SUB_ITEM As Currency
   Dim VALOR_SUB_CALCULADO As Currency
   Dim VALOR_ICMS_NORMAL As Currency
   Dim VALOR_NOTA_SEM_SUB As Currency
   Dim ValorICMSSubst As Double
   Dim ValorIPI As Double

   VALOR_ICMS_N = 0
   VALOR_IPI_N = 0
   VALOR_DESCONTO_N = 0 & txtDesconto.Text
   VALOR_TOTAL_N = 0 & txtValorTotalNota.Text
   ValorICMSSubst = 0 & txtValorICMSSubst.Text
   ValorIPI = 0 & txtValorIPI.Text

   VALOR_NOTA_SEM_SUB = VALOR_TOTAL_N - VALOR_DESCONTO_N
   txtValorTotalNota.Text = (VALOR_NOTA_SEM_SUB + ValorICMSSubst + ValorIPI)

  intRetorno = MsgBox("Deseja atualizar preço de custo dos itens no estoque?", vbQuestion + vbYesNo + vbDefaultButton2)
  
  'Pegando Calculo Valor Icms Substituicao se tiver
   If IsNumeric(txtBaseICMSSubst.Text) Then
      If txtBaseICMSSubst.Text > 0 Then
         If VALOR_NOTA_SEM_SUB <= txtValorTotalNota.Text Then
            VALOR_DIF_ICMS_SUB = (txtValorTotalNota.Text - txtValorIPI.Text - VALOR_NOTA_SEM_SUB)
            Else: VALOR_DIF_ICMS_SUB = (VALOR_NOTA_SEM_SUB - txtValorIPI.Text - txtBaseICMSSubst.Text)
         End If
      End If
   End If
   If VALOR_DIF_ICMS_SUB > 0 Then
      VALOR_ACUMULADO_SUB = (VALOR_DIF_ICMS_SUB * 100)
     'Percentual correspondente a valor agregado da substituicao do valor total da compra conforme formula indicada
      PERC_ICMS_SUB = (VALOR_ACUMULADO_SUB / VALOR_NOTA_SEM_SUB)
   End If

   If TabItemNota.State = 1 Then _
      TabItemNota.Close

   SQL = "select * from NOTAENTRADAITEM i , NOTAENTRADA n "
   SQL = SQL & " where n.numr_pedido_compra = " & txtPedido.Text
   SQL = SQL & " and n.entrada_id = i.entrada_id "
   SQL = SQL & " and n.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by i.seq desc"
   TabItemNota.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabItemNota.EOF
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_produto = '" & Trim(TabItemNota!Codg_Prod) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         SQL = "UPDATE Produto SET "
         SQL = SQL & " Dt_Ult_Compra = '" & DMA(Date) & "'"
         SQL = SQL & ", qtde = " & Str(TabProduto!Qtde + TabItemNota!QTD_ENTRADA)
         SQL = SQL & " Where empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and codg_produto = '" & Trim(TabItemNota!Codg_Prod) & "'"
         CONECTA_RETAGUARDA.Execute SQL

         If intRetorno = vbYes Then
               'Calculando Preco de Custo da Mercadoria , Falta Frete
               'VALOR_ICMS_N = TABITEMNOTA!preco_custo * TABITEMNOTA!PERC_ICMS / 100
               VALOR_ICMS_N = TabItemNota!PRECO_CUSTO * txtPercICMS.Text / 100
               VALOR_IPI_N = TabItemNota!PRECO_CUSTO * TabItemNota!PERC_IPI / 100
               
               'Calculo Preco Custo Mercadoria
               
               'Calculo Valor Substituicao por item automatico
               If PERC_ICMS_SUB > 0 Then
                  VALOR_ICMS_SUB_ITEM = (((TabItemNota!PRECO_CUSTO + VALOR_IPI_N) * PERC_ICMS_SUB) / 100)
                  'VALOR_SUB_CALCULADO = (VALOR_ICMS_SUB_ITEM + (TABITEMNOTA!Preco_Custo + VALOR_IPI_N))
                  'VALOR_SUB_CALCULADO = ((VALOR_SUB_CALCULADO * 17) / 100) 'pegar sempre aliquota interna
                  'VALOR_ICMS_NORMAL = ((TABITEMNOTA!Preco_Custo * ALIQUOTA_FORNEC) / 100) ' pegar a aliquota externa da nota
                  'VALOR_ICMS_SUB_ITEM = (VALOR_SUB_CALCULADO - VALOR_ICMS_NORMAL)
               End If
               
               'Calculo para achar valor de frete e outros cobre os itens da nota de entrada
               If txtFrete.Text > 0 Then
                  VLR_FRETE_N = VALOR_NOTA_SEM_SUB - txtFrete.Text
                  VLR_FRETE_N = (VLR_FRETE_N * 100)
                  VLR_FRETE_N = (VLR_FRETE_N / VALOR_NOTA_SEM_SUB)
                  VLR_FRETE_N = 100 - VLR_FRETE_N
               End If
               If txtValorOutras.Text > 0 Then
                  VLR_OUTROS_N = VALOR_NOTA_SEM_SUB - txtValorOutras.Text
                  VLR_OUTROS_N = (VLR_OUTROS_N * 100)
                  VLR_OUTROS_N = (VLR_OUTROS_N / VALOR_NOTA_SEM_SUB)
                  VLR_OUTROS_N = 100 - VLR_OUTROS_N
               End If
               If VLR_FRETE_N > 0 Then
                  VLR_FRETE_N = TabItemNota!PRECO_CUSTO * VLR_FRETE_N / 100
               End If
               If VLR_OUTROS_N > 0 Then
                  VLR_OUTROS_N = TabItemNota!PRECO_CUSTO * VLR_OUTROS_N / 100
               End If
               
               
               'If TabProduto!preco_custo <> TabProduto!PRECO_CUSTO_ANTERIOR Then
               '   If TabProduto!PRECO_CUSTO_ANTERIOR > TabProduto!preco_custo Then
               '      intRetorno_2 = MsgBox("Custo Anterior Maior que Custo Atual,Deseja atualizar preço de custo Pela media de Custo?", vbQuestion + vbYesNo + vbDefaultButton2)
               '      Else
               '         intRetorno_2 = MsgBox("Custo Anterior Menor que Custo Atual,Deseja atualizar preço de custo Pela media de Custo?", vbQuestion + vbYesNo + vbDefaultButton2)
               '   End If
               '   If intRetorno_2 = vbYes Then
               '      TabProduto!preco_custo = Format(TabProduto!PRECO_CUSTO_ANTERIOR + TabProduto!preco_custo / TabProduto!qtd, strFormatacao2Digitos)
               '   End If
               'End If
               
               Dim VALOR_TAXA_MARC_VAREJO_N As Double
               Dim VALOR_TAXA_MARC_ATACADO_N As Double
               
               VALOR_TOTAL_N = 0
               VALOR_TAXA_VAREJO_N = 0
               VALOR_TAXA_ATACADO_N = 0
               
               'Busca Taxa Marcacao Varejo
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select sum(perc_taxa) from ITENTAXA"
               SQL = SQL & " where taxamarc_id = " & 1
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp.Fields(0).Value) Then
                  'Calculo de Valores da Taxa de Marcacao
                     VALOR_TOTAL_N = 100 - TabTemp.Fields(0).Value
                     VALOR_TAXA_MARC_VAREJO_N = Format(100 / VALOR_TOTAL_N, strFormatacao2Digitos)
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               VALOR_TOTAL_N = 0 ' Zerando para fazer marcacao de atacado
               
               'Busca Taxa Marcacao atacado
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select sum(perc_taxa) from ITENTAXA"
               SQL = SQL & " where taxamarc_id = " & 2
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp.Fields(0).Value) Then
                  'Calculo de Valores da Taxa de Marcacao
                     VALOR_TOTAL_N = 100 - TabTemp.Fields(0).Value
                     VALOR_TAXA_MARC_ATACADO_N = Format(100 / VALOR_TOTAL_N, strFormatacao2Digitos)
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               'Busca Taxa Marcacao por Fornecedor
               SQL = "select * from FORNECEDOR"
               SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp!markup_varejo) Then
                     VALOR_TAXA_MARC_VAREJO_N = TabTemp!markup_varejo
                  End If
                  If Not IsNull(TabTemp!markup_atacado) Then
                     VALOR_TAXA_MARC_ATACADO_N = TabTemp!markup_atacado
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               VALOR_TOTAL_N = TabItemNota!PRECO_CUSTO + VALOR_IPI_N + VALOR_ICMS_SUB_ITEM + VLR_FRETE_N + VLR_OUTROS_N

               SqL2 = "UPDATE Produto SET "
               SqL2 = SqL2 & " preco_custo = " & Str(VALOR_TOTAL_N)
               SqL2 = SqL2 & ",PRECO_venda = " & tpMoeda(VALOR_TOTAL_N * VALOR_TAXA_MARC_VAREJO_N)
               SqL2 = SqL2 & ",PRECO_atacado = " & tpMoeda(VALOR_TOTAL_N * VALOR_TAXA_MARC_ATACADO_N)
               SqL2 = SqL2 & ",PRECO_CUSTO_ANTERIOR = " & tpMoeda(TabProduto!PRECO_CUSTO)
               SqL2 = SqL2 & ",PRECO_ATACADO_ANTERIOR = " & tpMoeda(TabProduto!Preco_Atacado)
               SqL2 = SqL2 & ",PRECO_VAREJO_ANTERIOR = " & tpMoeda(TabProduto!PRECO_VENDA)
               SqL2 = SqL2 & " Where empresa_id = " & EMPRESA_ID_N
               SqL2 = SqL2 & " and codg_produto = '" & TabProduto!CODG_PRODUTO & "'"
               CONECTA_RETAGUARDA.Execute SqL2
         End If
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close

      TabItemNota.MoveNext
   Wend
   If TabItemNota.State = 1 Then _
      TabItemNota.Close

   If optEntrada.Value = True Then
      SQL = "update NOTAENTRADA set "
      SQL = SQL & " status = 'E' "
      SQL = SQL & " where numr_pedido_compra = " & txtPedido.Text
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ESTOQUE"
End Sub

Private Sub FINANCEIRO_FORM()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   CPF_N = txtCNPJCPF.Text
   NUMR_REQ_N = txtPedido.Text
   SINAL = 2
   frmCADFINENTRADA.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FINANCEIRO_FORM"
End Sub

Private Sub PROCURA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   NUMR_REQ_N = 0
   If txtPedido.Text <> "" Then _
      If IsNumeric(txtPedido.Text) Then _
         NUMR_REQ_N = txtPedido.Text

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NOTAENTRADA "
   SQL = SQL & " where numr_pedido_compra = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      NUMR_REQ_N = 0

      If SINAL = 0 Then _
         SINAL = 2

      If SINAL <> 2 Then _
         LIMPA_NOTA_ENTRADA

      MOSTRA_NOTA_ENTRADA

      SETA_GRID

      If TabNOTA!Status = "C" Then
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         MsgBox "Nota fiscal cancelada, impossível alterar."
         Indr_Consulta = True
         'LIMPA_NOTA_ENTRADA
         'txtPEDIDOCOMPRA.SetFocus
         Exit Sub
      End If
      If TabNOTA!Status = "E" Then
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         MsgBox "Nota fiscal já atualizada, impossível alterar."
         Indr_Consulta = True
         'LIMPA_NOTA_ENTRADA
         'txtPEDIDOCOMPRA.SetFocus
         Exit Sub
      End If
      Indr_Consulta = False
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
ERRO_TRATA:

   TRATA_ERROS Err.Description, Me.Name, "PROCURA_NOTA_ENTRADA"
End Sub

Private Sub GRAVA_Devolução()
End Sub

Private Sub GRAVA_CABECA_Devolução()
End Sub

Private Sub GRAVA_CORPO_Devolução()
End Sub

Private Sub GRAVA_CORPO_DEV_QTD()
End Sub

Private Sub SETA_GRID_DEV()
End Sub

Private Sub MOSTRA_TOTAL_NOTA_DEV()
End Sub

Private Sub MONTA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   NUMR_COMPRA_N = 0
   If txtPedidoCompra.Text <> "" Then _
      If IsNumeric(txtPedidoCompra.Text) Then _
         NUMR_COMPRA_N = txtPedidoCompra.Text
   PedidoCompra = True

   If TABCOMPRA.State = 1 Then _
      TABCOMPRA.Close

   SQL = "select * from PEDIDOCOMPRA "
   SQL = SQL & " where pedido = " & txtPedidoCompra.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TABCOMPRA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TABCOMPRA.EOF Then
      txtCNPJCPF.Text = TABCOMPRA!CGCCPF

      If TabFORNEC.State = 1 Then _
         TabFORNEC.Close

      SqL2 = "select * from FORNECEDOR "
      SqL2 = SqL2 & " where CGCCPF = '" & TABCOMPRA!CGCCPF & "'"
      TabFORNEC.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      txtNome.Text = TabFORNEC!Nome

      If cmbCFOP.Text = "" Then 'nao escolheu cfop, coloca o default
         If txtUF.Text = "GO" Then
            cmbCFOP.Text = CFOP_ENTRADA_DE
            Else: cmbCFOP.Text = CFOP_ENTRADA_FE
         End If
         PegaDescricaoCFOP
      End If
      
      SETA_GRID_COMPRA
   Else
       PedidoCompra = False
   End If
   If TABCOMPRA.State = 1 Then _
      TABCOMPRA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_NOTA_ENTRADA"
End Sub
   
Private Sub SETA_GRID_COMPRA()
'On Error GoTo ERRO_TRATA

   LISTAITENS.ListItems.Clear
   If txtPedidoCompra.Text <> "" Then
      If tabPEDIDOCOMPRAITEM.State = 1 Then _
         tabPEDIDOCOMPRAITEM.Close

      SQL = "select * from PEDIDOCOMPRAITEM i "
      SQL = SQL & " where i.pedido = " & txtPedidoCompra.Text
      SQL = SQL & " and i.qtd > " & 0
      SQL = SQL & " order by i.sequencia"
      tabPEDIDOCOMPRAITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not tabPEDIDOCOMPRAITEM.EOF
         Set Item = LISTAITENS.ListItems.Add(, "seq." & tabPEDIDOCOMPRAITEM!sequencia, tabPEDIDOCOMPRAITEM!sequencia)
         Item.SubItems(1) = tabPEDIDOCOMPRAITEM!Produto

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select descricao from PRODUTO "
         SQL = SQL & " where codg_prod='" & tabPEDIDOCOMPRAITEM!Produto & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            Item.SubItems(2) = TabTemp!Descricao
         If TabTemp.State = 1 Then _
            TabTemp.Close

         Item.SubItems(3) = tabPEDIDOCOMPRAITEM!qtd
         Item.SubItems(4) = tabPEDIDOCOMPRAITEM!Preco
         Item.SubItems(5) = tabPEDIDOCOMPRAITEM!Preco
         Item.SubItems(7) = Format((tabPEDIDOCOMPRAITEM!Preco * tabPEDIDOCOMPRAITEM!qtd), strFormatacao2Digitos)
         tabPEDIDOCOMPRAITEM.MoveNext
      Wend
      If tabPEDIDOCOMPRAITEM.State = 1 Then _
         tabPEDIDOCOMPRAITEM.Close
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SqL2 = "select sum(i.preco * i.qtd) from PEDIDOCOMPRAITEM i"
   SqL2 = SqL2 & " where i.pedido = " & NUMR_COMPRA_N
   TabTemp.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_TOTAL_N = TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtValorTotalNota.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   stBarReq.Panels(8).Text = txtValorTotalNota.Text
   stBarReq.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_COMPRA"
End Sub

' Inicio FUNÇÕES PARA A IMPORTAÇÃO DO XML, CONFORME O CAMINHO DADO PARA O USUARIO
' yuri 11/05/2012
'*****************************************************************************
'Criação: Yuri Grandinetti                                    Data: 11/05/2012
'Propósito: Preenche todos os campos do cabeçalho da nota.
'           A busca do fornecedor é feita pelo CNPJ da mesma. O fornecedor tem
'           que estar cadastrado no sistema com o CNPJ da NF-e!
' Substituir os textbox,etc pelo nossos no megasim, se nao tiver cria como invisivel
'*****************************************************************************
Private Sub PreencheCabecalho()
'On Error GoTo Err_PreencheCabecalho

'    Call BuscaFornecedor(0, cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "emit//CNPJ"))                                  'Fornecedor
'    txt_numero_nf.Text = Val(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "ide//nNF"))                                   'Número
'    txt_serie_nf.Text = Val(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "ide//serie"))                                  'Série
'    txt_modelo_nf.Text = Val(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "ide//mod"))                                   'Modelo
'    lbl_forma.Caption = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "ide//natOp")                                       'Natureza
'    msk_emissao.Text = Format(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "ide//dEmi"), "dd/mm/yyyy")                   'Data da emissão
'    txt_chave_acesso.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "infProt//chNFe")                               'Chave de acesso
'    txt_total.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vNF"), ".", ",")              'Total da nota
'    txt_bc_icms.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vBC"), ".", ",")            'Base Calculo ICMS
'    txt_icms.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vICMS"), ".", ",")             'Valor ICMS
'    txt_bc_substituicao.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vBCST"), ".", ",")  'Base Subs. Trib.
'    txt_substituicao.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vST"), ".", ",")       'Valor ST
'    txt_ipi.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vIPI"), ".", ",")               'Valor IPI
'    txt_frete.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vFrete"), ".", ",")           'Valor Frete
'    txt_seguro.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vSeg"), ".", ",")            'Valor Seguro
'    txt_desconto.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vDesc"), ".", ",")         'Valor Desconto
'    txt_outras.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "total//ICMSTot//vOutro"), ".", ",")          'Valor Outro

Exit Sub
'Err_PreencheCabecalho: ValidaErros Err, Me.Caption & " - PreencheCabecalho"
End Sub
'Yuri 11/05/2012
' CASO NAO EXISTE O FORNECEDOR PELA PESQUISA
' VOCE DEIXA A O CAMPO TXTCNPJCPF EM BRANCO PARA O USUARIO SELECIONAR
' VOCE PODE AUTOMATIZAR ESSE PROCESSO FAZENDO A INCLUSÃO DOS DADOS DO FORNECEDOR QUE ESTA NO XML
' SE NÃO VOCE TERA QUE CONSTRUIR UMA PESQUISA E DA PESQUISAR FAZER O USUARIO CADASTRAR O FORNECEDOR
' POIS AQUI NÃO FOI IMPLEMENTADO
' NÃO LEMBRO SE TEM ESSE VALIDADOR vALIDAeRROS SE NAO TIVER USA O SEU
Private Sub BuscaFornecedor(xcodigo As String, Optional strCNPJ As String)
'On Error GoTo BuscaFornecedor

Dim strCondicao As String


'    If strCNPJ <> "" Then
'        strCondicao = " CGC = '" & strCNPJ & "'"
'    Else
'        strCondicao = " codigo = '" & xcodigo & "'"
'    End If
'    gb_Recordset.Source = "SELECT * FROM fornecedor WHERE " & strCondicao
'    gb_Recordset.Open
'
'
'    If gb_Recordset.RecordCount > 0 Then
'        LimpaTelaFornecedor
'        AtualTelaFornecedor
'
'        'Efetuado teste para evitar nos casos de consulta ou conferência de produtos
'        'gerar algum erro por causa do campo desabilitado ***
'        If frmDados.Enabled = True Then txt_fornecedor.SetFocus
'
'    Else
'        LimpaTelaFornecedor
'    End If
'    gb_Recordset.Close

Exit Sub
'BuscaFornecedor: ValidaErros Err, Me.Caption & " - buscafornecedor"
End Sub

' Yuri 11/05/2012
' coloque aqui os componentes  relacionado com a tela do megasim
Private Sub LimpaTelaFornecedor()
    'txt_codigo_Fornecedor = ""
    'txt_fornecedor = ""
    'lbl_cgc = ""
    'lbl_inscricao = ""
    'lbl_endereco = ""
    'lbl_cidade = ""
    'lbl_uf = ""
    'lbl_pais = ""
    'x_perc_icms_subst = 0
    'txt_porc_red_icms = 0
End Sub
' yuri 11/05/2012
' aqui preenche conforme nossa tela
' por exemplo se não tiver a label lporcentagemreducao cria ele como invisivel para o usuário
' a variavel global g_reducao_invertido voce criar no modDeclara
Private Sub AtualTelaFornecedor()
On Error GoTo file

'    With gb_Recordset
'       If !PESSOA = 1 Then
'           lbl_cgc = !CGC
'           lbl_inscricao = !Identidade
'           lCGCFornecedor = ""
'       ElseIf !PESSOA = 3 Then
'           lbl_cgc = !CGC
'           lbl_inscricao = !inscricao_estadual
'           lCGCFornecedor = ""
'       Else
'           lbl_cgc = !CGC
'           lCGCFornecedor = !CGC
'           lbl_inscricao = !inscricao_estadual
'       End If
'       lPessoa = !PESSOA
'       txt_codigo_Fornecedor = !Codigo
'       txt_fornecedor.Text = !razao_social
'       lbl_endereco = !Endereco & " - " & !Bairro
'       lbl_cidade = !Cidade
'       If IsNull(!Pais) Or Trim(!Pais) = "" Then
'           lbl_pais = "BRASIL"
'       Else
'           lbl_pais = !Pais
'       End If
'
'       lbl_uf = !UF
'       lbl_telefone = !Telefone
'       lbl_bairro = !Bairro
'       lbl_cep = !Cep
'
'       lporcentagemreducao = !red_icms
'       x_perc_icms_subst = !icms_subst
'       xplanoconta = !conta_contabil
'       g_reducao_invertido = !reducao_invertido
'       txt_porc_red_icms = !red_icms
'    End With
             
Exit Sub
file: If Err.Number = 94 Then Resume Next
End Sub

' 11/05/2012 yuri
'*****************************************************************************
'Criação: Yuri Grandinetti                                         Data: 11/05/2012
'Propósito: Preenche os campos da transportadora com os dados no XML.
'           O sistema busca os dados pelo CNPJ da transp., onde o mesmo tem
'           que estar cadastrado.
'
'*****************************************************************************
Private Sub PreencheTransporte()
'On Error GoTo Err_PreencheTransporte

'    Call BuscaTransportadora(0, cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//transporta//CNPJ"))
'    txt_volume.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//qVol")
'    txt_especie.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//esp")
'    txt_peso_liquido.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//pesoL")
'    txt_peso_bruto.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//vol//pesoB")
'    txt_transportadora_placa.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//veicTransp//placa")
'    txt_transportadora_placa_uf.Text = cNotaEntrada.RetornaTagXML((Trim(txt_caminho_xml)), "transp//veicTransp//UF")

Exit Sub
'Err_PreencheTransporte: ValidaErros Err, Me.Caption & " - PreencheTransporte"
End Sub
' 11/05/2012
' aqui voce substitui pelo nosso cadastro de transportadora
Private Sub BuscaTransportadora(strCodigo As String, Optional strCNPJ As String)
'On Error GoTo BuscaTransportadora

Dim strCondicao As String

'    If strCNPJ <> "" Then
'        strCondicao = " CGC = '" & strCNPJ & "'"
'    Else
'        strCondicao = " codigo = '" & strcodigo & "'"
'    End If
'    gb_Recordset.Source = "SELECT * FROM transportadora WHERE " & strCondicao
'    gb_Recordset.Open
'    'Set gb_Recordset = Conexao.GeraRecordset("SELECT * FROM transportadora WHERE " & strCondicao, 1)
'    If gb_Recordset.RecordCount > 0 Then
'        Call PreencheTransportadora
'        txt_transportadora_nome.SetFocus
'    Else
'        Alerta "Transportadora não cadastrada!"
'    End If
'    gb_Recordset.Close

Exit Sub
'BuscaTransportadora: ValidaErros Err, Me.Caption & " - BuscaTransportadora"
End Sub

'*****************************************************************************
'Criação: Yuri Grandinetti                                     Data: 11/05/2012
'Propósito: Preenche dados da transportadora
' Sugestão caso nao tenha alguns textbox cria como invisivel para o usuario
'*****************************************************************************
Private Sub PreencheTransportadora()
'On Error GoTo Err_PreencheTransportadora

'    With gb_Recordset
'        txt_codigo_transportadora = !Codigo
'        txt_transportadora_nome = !Nome
'        txt_transportadora_endereco = !Endereco
'        txt_transportadora_bairro = !Bairro
'        txt_transportadora_cidade = !Cidade
'        txt_transportadora_uf = !UF
'        txt_transportadora_cnpj = !CGC
'        txt_transportadora_inscricao_estadual = !inscricao_estadual
'    End With

Exit Sub
'Err_PreencheTransportadora: ValidaErros Err, Me.Caption & " - PreencheTransportadora"
End Sub
' YURI 11/05/2012
' aqui voce vai precisar adaptar para o listview do megasim, pois aqui a grade faz referencia ao
' MSHFLEXGRID
'xmlElem.SelectSingleNode SÃO OS atributos do xml da notafiscal eletronica
' pega uma nota e voce entendera
' se nao tiver algum elemento no listview, neste caso voce precisara mudar o listview
' OBS .: Acho que voce vai precisar infelizmente trocar o listview pelo componente
' mshflexgrid, pois não tem como sabermos o codigo do produto, pois esse código e do nosso fornecedor
' como ja existe muitos produtos cadastrados nos nossos clientes não temos como amarrar isso, pois não vem no
' xml do cliente
' SUGESTÃO : SE VOCE NÃO CONSEGUIR UMA MANEIRA DE ALTERAR O PRODUTO QUE VAI VIR NO LISTVIEW, INFELIZMENTE VOCE TERA QUE
' TRABALHAR COM ESSE COMPONENTE
' UMA COISA IMPORTANTE CASO QUEIRA TRABALHAR E NÃO RECOMENDO ISSO PELO TRABALHO QUE VAI LHE DAR EU TE PASSO
' A ROTNA DA NOTA FISCAL DE ENTRADA PARA VOCE TER UMA IDÉIA DE COMO FAZER, EU ACHO UM POUCO COMPLEXO, PARA FAZER ISSO
' A TOQUE DE CAIXA
Public Sub PreencheProduto(strCaminhoXML As String)
On Error Resume Next

   Dim lngItem As Long
   Dim lngLinha As Long
   Dim XML As DOMDocument
   Dim xmlElem As IXMLDOMNode
   Dim bolFimProdutos As Boolean
   Set XML = New DOMDocument
    
'    lngItem = 0
'    lngLinha = 1
'    XML.async = False
'    bolFimProdutos = False
'    If XML.Load(strCaminhoXML) Then
'
'        Do Until bolFimProdutos
'            Set xmlElem = XML.SelectNodes("/nfeProc/NFe/infNFe/det").Item(lngItem).FirstChild
'
'                '*** Verfica se tem algum valor no código do produto, senão tiver finaliza o loop ***
'                If xmlElem.SelectSingleNode("cProd").Text = "" Then bolFimProdutos = True: Exit Do
'
'                grade1.TextMatrix(lngLinha, 2) = xmlElem.SelectSingleNode("cProd").Text                           'Codigo produto
'                grade1.TextMatrix(lngLinha, 3) = xmlElem.SelectSingleNode("xProd").Text                           'Descrição produto
'                grade1.TextMatrix(lngLinha, 4) = xmlElem.SelectSingleNode("uCom").Text                            'Unidade produto
'                grade1.TextMatrix(lngLinha, 5) = Trim$(Format(xmlElem.SelectSingleNode("qCom").Text, "currency")) 'Quantidade Produto
'                grade1.TextMatrix(lngLinha, 6) = Replace(xmlElem.SelectSingleNode("vUnCom").Text, ".", ",")       'Valor unitário Produto
'                grade1.TextMatrix(lngLinha, 7) = Replace(xmlElem.SelectSingleNode("vDesc").Text, ".", ",")        'Valor Desconto Produto
'                grade1.TextMatrix(lngLinha, 8) = Replace(xmlElem.SelectSingleNode("vUnCom").Text, ".", ",")       'Valor unitário Produto
'                grade1.TextMatrix(lngLinha, 11) = Replace(xmlElem.SelectSingleNode("vProd").Text, ".", ",")       'Valor Total Produto
'                grade1.TextMatrix(lngLinha, 17) = "NF" ' sergio aqui é so para referenciar que é uma nota fiscal e nao um cupom fiscal
'                grade1.TextMatrix(lngLinha, 28) = xmlElem.SelectSingleNode("CFOP").Text                           'CFOP Produto
'                grade1.TextMatrix(lngLinha, 38) = xmlElem.SelectSingleNode("cEAN").Text                           'Cód. barras produto
'                grade1.TextMatrix(lngLinha, 40) = xmlElem.SelectSingleNode("NCM").Text                            'Código NCM Produto
'
'                xmlElem.SelectSingleNode("cProd").Text = "" 'Limpa o objeto para setar um novo.
'
'                lngItem = lngItem + 1
'                lngLinha = lngLinha + 1
'                Call ChamaCelula
'        Loop
'    Else
'        MsgBox "Não foi possível abrir o arquivo XML da NFe especificada para Leitura.", vbCritical, "Erro."
'    End If
    
End Sub
' 11/05/2012
' Aqui não sei como voce vai fazer, mas aqui no caso do componente mshflexgrid
' é para adicionar um novo produto dos itens do xml na grid
Private Sub ChamaCelula()
    
'    LastRow = grade1.Row
'    LastCol = grade1.Col
'
'    'Nova Celula
'    With grade1
'        If .TextMatrix(LastRow, 0) = NovaLinha Then
'            .Rows = .Rows + 1
'            .TextMatrix(LastRow, 0) = LastRow
'            .TextMatrix(.Rows - 1, 0) = NovaLinha
'            LastRow = LastRow + 1
'            ZeraGrade
'       End If
'    End With
'
'grade1.Col = 1
'grade1.Row = grade1.Row + 1
End Sub

' FIM DAS FUNÇÕES PARA A IMPORTAÇÃO DO XML, CONFORME O CAMINHO DADO PARA O USUARIO

Private Sub cmd_explorer_Click()
    buscaxml.InitDir = "C:\"
    buscaxml.ShowOpen
    txt_caminho_xml.Text = buscaxml.FileName
    buscaxml.FileName = ""
End Sub

Private Sub cmd_ler_xml_Click()
    ' yuri 11/05/2012
    'Call PreencheCabecalho
    'Call PreencheTransporte
    'Call PreencheProduto(txt_caminho_xml)
End Sub

Private Sub PegaDadosEmpresa()
'On Error GoTo ERRO_TRATA

   If rstEmpresa.State = 1 Then _
      rstEmpresa.Close

   SQL = "Select * From EMPRESA where EMPRESA_ID = " & EMPRESA_ID_N
   rstEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstEmpresa.EOF Then
      If rstEmpresa.State = 1 Then _
         rstEmpresa.Close
      MsgBox "O sistema nao obteve sucesso ao tentar localizar a empresa corrente."
      Unload Me
   End If

   CFOP_DV_ENT_FE = rstEmpresa.Fields("CFOP_DV_ENT_FE").Value
   CFOP_DV_ENT_DE = rstEmpresa.Fields("CFOP_DV_ENT_DE").Value
   CFOP_ENTRADA_FE = rstEmpresa.Fields("CFOP_ENTRADA_FE").Value
   CFOP_ENTRADA_DE = rstEmpresa.Fields("CFOP_ENTRADA_dE").Value

   If rstEmpresa.State = 1 Then _
      rstEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PegaDadosEmpresa"
End Sub

Private Sub PegaDescricaoCFOP()
'On Error GoTo ERRO_TRATA

   Dim strCFOP As String

   If Len(cmbCFOP.Text) > 4 Then
      strCFOP = Mid(cmbCFOP.Text, 1, (InStr(1, cmbCFOP.Text, "-")) - 1)
      'strCFOP = Mid(cmbCFOP.Text, 1, (InStr(1, cmbCFOP.Text, "-")))
      'strCFOP = Left(cmbCFOP.Text, (InStr(1, cmbCFOP.Text, "-")))
      Else: strCFOP = cmbCFOP.Text
   End If
   'strCFOP = Left(cmbCFOP.Text, (InStr(1, cmbCFOP.Text, "-")))

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP "
   SQL = SQL & " where codigo = '" & strCFOP & "'"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      cmbCFOP.Text = TabDESCR!Codigo & "-" & TabDESCR!Descricao
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PegaDescricaoCFOP"
End Sub

Private Sub preencheComboCfop()
'On Error GoTo ERRO_TRATA

   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      Do Until TabDESCR.EOF
         DoEvents
         'cmbAuxCFOP.AddItem TABDESCR!Codigo
         cmbCFOP.AddItem TabDESCR!Codigo & "-" & TabDESCR!Descricao
         TabDESCR.MoveNext
      Loop
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "preencheComboCfop"
End Sub

Sub MOSTRA_NOTA()
'On Error GoTo ERRO_TRATA

   If optEntrada.Value = True Then
      If txtNota.Text = "" Then
         MsgBox "Informe número de nota."
         txtNota.SetFocus
         Exit Sub
      End If
   End If

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe fornecedor."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   If optEntrada.Value = True Then
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "select * from NOTAENTRADA "
      SQL = SQL & " where numr_nota = " & txtNota.Text
      SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         NUMR_REQ_N = 0
         LIMPA_NOTA_ENTRADA
         MOSTRA_NOTA_ENTRADA
         'SETA_GRID

         If TabNOTA!Status = "C" Then
             If TabNOTA.State = 1 Then _
                TabNOTA.Close

            MsgBox "Nota fiscal cancelada, impossível alterar."
            LIMPA_NOTA_ENTRADA
            txtNota.SetFocus
            Exit Sub
         End If

         If TabNOTA!Status = "E" Then
             If TabNOTA.State = 1 Then _
                TabNOTA.Close

            MsgBox "Nota fiscal já atualizada, impossível alterar."
            LIMPA_NOTA_ENTRADA
            txtPedidoCompra.SetFocus
            Exit Sub
         End If
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

   End If
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_NOTA"
End Sub
