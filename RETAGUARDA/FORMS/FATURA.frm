VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFatura 
   BackColor       =   &H000000C0&
   Caption         =   "Recebimento Caixa "
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   1845
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FATURA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10950
   StartUpPosition =   1  'CenterOwner
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
      Height          =   3855
      Left            =   0
      TabIndex        =   13
      Top             =   550
      Width           =   10935
      Begin VB.TextBox txtTxEntrega 
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
         Height          =   480
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   53
         Top             =   1200
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtDiaVencto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6660
         MaxLength       =   8
         TabIndex        =   50
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtCNPJCPF 
         Height          =   420
         Left            =   5160
         TabIndex        =   49
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cmbBandeiraAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   48
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbBandeira 
         Height          =   405
         Left            =   1920
         TabIndex        =   46
         Top             =   3240
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.ComboBox cmbTipoVendaINDR 
         BackColor       =   &H80000000&
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   2640
         TabIndex        =   41
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtPercEntrada 
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
         Height          =   495
         Left            =   4080
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtValorEntrada 
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
         Height          =   480
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   0
         Top             =   1800
         Width           =   2040
      End
      Begin VB.ComboBox cmbTipoVendaAUX 
         BackColor       =   &H80000000&
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   1920
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTipoVenda 
         Height          =   405
         Left            =   1920
         TabIndex        =   2
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtData 
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
         Height          =   480
         Left            =   8640
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   9000
         TabIndex        =   26
         Top             =   3120
         Width           =   1800
      End
      Begin VB.TextBox txtVendaComDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   9000
         TabIndex        =   24
         Top             =   1680
         Width           =   1800
      End
      Begin VB.TextBox txtRecebido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   555
         Left            =   9000
         TabIndex        =   18
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtVendaSemDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   9000
         TabIndex        =   16
         Top             =   960
         Width           =   1800
      End
      Begin VB.TextBox txtCli 
         Height          =   420
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtVendedor 
         Height          =   420
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblFat 
         AutoSize        =   -1  'True
         Caption         =   "Negociação Pedido Venda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   5160
         TabIndex        =   54
         Top             =   1320
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblTxEntrega 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Taxa Entrega = "
         Height          =   285
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label lblDiaVencto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DiasVencto:"
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
         Left            =   5520
         TabIndex        =   51
         Top             =   2400
         Width           =   1125
      End
      Begin VB.Label lblBandeira 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Band.Cartão:"
         Height          =   285
         Left            =   285
         TabIndex        =   47
         Top             =   3240
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblPemite_Desconto 
         Caption         =   "Não Permite Desconto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1920
         TabIndex        =   42
         Top             =   2880
         Visible         =   0   'False
         Width           =   1905
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
         TabIndex        =   33
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Entrada = "
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Venda:"
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Troco = "
         Height          =   285
         Left            =   7965
         TabIndex        =   27
         Top             =   3120
         Width           =   930
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desconto = "
         Height          =   285
         Left            =   7530
         TabIndex        =   25
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Recebido = "
         Height          =   285
         Left            =   7545
         TabIndex        =   23
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total = "
         Height          =   285
         Left            =   8070
         TabIndex        =   21
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente:"
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendedor:"
         Height          =   330
         Left            =   3630
         TabIndex        =   19
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame fraSeq 
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
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   4440
      Width           =   10935
      Begin VB.ComboBox cmbCCAux 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbCC 
         BackColor       =   &H00FFFFFF&
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
         Left            =   9000
         TabIndex        =   43
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtDias 
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
         Height          =   360
         Left            =   5760
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtSeq 
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
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbModalidadeAUX 
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
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtValorItem 
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
         Height          =   360
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbModalidade 
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
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtDTVENC 
         Height          =   360
         Left            =   7680
         TabIndex        =   8
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Left            =   6360
         TabIndex        =   7
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Centro Custo"
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
         Left            =   9045
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Vencto"
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
         Index           =   6
         Left            =   7680
         TabIndex        =   39
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Index           =   5
         Left            =   6360
         TabIndex        =   38
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
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
         Index           =   4
         Left            =   5760
         TabIndex        =   37
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor "
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
         Left            =   4320
         TabIndex        =   36
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagto"
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
         Left            =   840
         TabIndex        =   35
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Seq.:"
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
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   120
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
            Picture         =   "FATURA.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FATURA.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FATURA.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FATURA.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FATURA.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FATURA.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FATURA.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1164
      ButtonWidth     =   2805
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confirmar"
            Key             =   "conf"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Informar Desconto"
            Key             =   "desconto"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lstFatura 
      Height          =   1905
      Left            =   0
      TabIndex        =   12
      Top             =   5430
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   3360
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   14737632
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Doc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   3528
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
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Dt.Venc."
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Juros"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CentroCusto"
         Object.Width           =   2540
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   200
      DesignWidth     =   10950
      DesignHeight    =   7725
   End
   Begin MSComctlLib.StatusBar barRodape 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   7350
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "FATURA.frx":76BE
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
End
Attribute VB_Name = "frmFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
   Dim VALOR_RECEBIDO_N          As Double
   Dim NUMR_PARCELA_N            As Integer
   Dim VALOR_TROCO_N             As Double
   Dim VALOR_TOTAL_LANÇADO       As Double
   Dim VALOR_ENTRADA_N           As Double
   Dim PERC_JUROS_N              As Double
   Dim DIAS_PRAZO                As Integer
   Dim DIA_VENCTO                As Integer
   Dim TabTipovenda              As New ADODB.Recordset
   Dim TabVai                    As New ADODB.Recordset
   Dim INDR_FINALIZA_RECEBIMENTO As Boolean
   Dim VALOR_DESCONTO_CABECA_N   As Double
   Dim TOTAL_DESCONTO_N          As Double
   Dim VALOR_DIFERENCA_N         As Double
   Dim VALOR_DESCONTO_ITEM_N     As Double
   Dim VALOR_VENDA_BRUTA_N       As Double
   Dim CC_ID_N                   As Integer
   Dim CNPJCPF_CLI_A             As String
   Dim VALOR_TX_ENTREGA_N        As Double

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INDR_FUNCIONARIO = False
   PESSOA_ID_N = 0
   CC_ID_N = 0
   lblDiaVencto.Visible = False
   txtDiaVencto.Visible = False

   INDR_FORM_ABERTO = True
   Me.Caption = Me.Caption & " - " & Me.Name
      
   VALOR_TOTAL_N = 0
   FraSeq.Enabled = False
   NUMR_PARCELA_N = 0

   LIMPA_LANCAMENTO

   txtData.Text = Now
   txtPedido.Text = PEDIDO_ID_N
   If PEDIDO_ID_N > 0 Then
      SETA_GRID
      Else
         MsgBox "Número de lançamento não foi informado. verifique."
         Unload frmFatura
   End If

   MOSTRA_DADOS_PEDIDO
   BUSCA_LANCAMENTO

   txtVendaSemDesconto.Refresh

   MOSTRA_RODAPE_AQUI "ESC - SAIR", "", "", "", ""

   If MULT_EMPRESA_B = True Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select descricao from vwVendedor "
      SQL = SQL & " where vendedor_id = " & VENDEDOR_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtVendedor.Text = "" & TabConsulta.Fields("descricao").Value
         txtVendedor.Refresh

'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

         SQL = "UPDATE PEDIDO set "
         SQL = SQL & " vendedor_id = " & VENDEDOR_ID_N
         SQL = SQL & ",valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
SQL = SQL & ",cliente_id = " & CLIENTE_ID_N
         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

   CARREGA_COMBO_TIPO_VENDA

   Toolbar1.Buttons(7).Visible = False

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select funcionario from USUARIO "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   'SQL = SQL & " where cpf = '" & Trim(CNPJCPF_CLI_A) & "'"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      If Not IsNull(TabConsulta.Fields("funcionario").Value) Then _
         INDR_FUNCIONARIO = TabConsulta.Fields("funcionario").Value
      Else
         If INDR_DESCONTO_CLIENTE = True Then _
            Toolbar1.Buttons(7).Visible = True
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   If INDR_DESCONTO_FUNCIONARIO = True And INDR_FUNCIONARIO = True Then _
      Toolbar1.Buttons(7).Visible = True

   cmbBandeiraAUX.Clear
   cmbBandeira.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'G'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbBandeira.AddItem Trim(TabTemp!DESCRICAO)
      cmbBandeiraAUX.AddItem "0" & TabTemp!Codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbTipoVenda.Enabled = True
   txtValorEntrada.Enabled = True
   FraSeq.Enabled = True
   lblFat.Visible = False
   lblFat.Caption = ""
   'vai verificar se é permitido alterar o tipo de recebimento,
   'caso não vai prevalecer registro que foi informado no pedido de venda
   If ALTERA_FATURA_B = False Then
      'ler dados do pedido pra saber qual tipo de faturamento foi informado durante o pedido de venda

      SQL = "SELECT PEDIDOFATURA.PEDIDO_ID, PEDIDOFATURA.FORMAPAGTO_ID, PEDIDOFATURA.TIPOVENDA_ID, "
      SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescForma, TIPOVENDA.DESCRICAO AS DescTipoVenda"
      SQL = SQL & " FROM PEDIDOFATURA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOFATURA.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID "
      SQL = SQL & " INNER JOIN TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " ON FORMAPAGTO.FORMAPAGTO_ID = TIPOVENDA.FORMAPAGTO_ID"

      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         cmbTipoVendaAUX.Text = "" & TabTemp.Fields("tipovenda_id").Value
         cmbTipoVenda.Text = "" & TabTemp.Fields("desctipovenda").Value
         lblFat.Visible = True
         lblFat.Caption = TabTemp.Fields("tipovenda_id").Value & "-" & TabTemp.Fields("desctipovenda").Value
         'cmbTipoVenda.Enabled = False
         'txtValorEntrada.Enabled = False
         'FraSeq.Enabled = False

         'Call cmbTipoVenda_Click
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload frmFatura
      Case vbKeyF1
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from TIPOVENDA WITH (NOLOCK)"
         SQL = SQL & " where tipovenda_id = 9999 "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            cmbTipoVenda.Text = Trim(TabTemp!DESCRICAO)
            cmbTipoVendaAUX.Text = Trim(TabTemp!TIPOVENDA_ID)
            CC_ID_N = 0 & TabTemp.Fields("cc_id").Value

            If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
               Else: cmbTipoVendaINDR.Text = "False"
            End If

            Call cmbTipoVenda_Click
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      Case vbKeyF3
         If Trim(CNPJCPF_CLI_A) <> "99999999999" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from TIPOVENDA WITH (NOLOCK)"

            SQL = SQL & " where contabiliza = 1 "
            SQL = SQL & " and (upper(left(descricao,8)) = upper('Convenio') or upper(left(descricao,12)) = upper('Venda Futura'))"
            SQL = SQL & " and receber = 'true' "
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               cmbTipoVenda.Text = Trim(TabTemp!DESCRICAO)
               cmbTipoVendaAUX.Text = Trim(TabTemp!TIPOVENDA_ID)
               CC_ID_N = 0 & TabTemp.Fields("cc_id").Value

               If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
                  Else: cmbTipoVendaINDR.Text = "False"
               End If

               Call cmbTipoVenda_Click
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      Case vbKeyF4
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from TIPOVENDA WITH (NOLOCK)"
         SQL = SQL & " where descricao = 'Cartao de Debito' "
         SQL = SQL & " and contabiliza = 1 "
         SQL = SQL & " and receber = 'true' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            cmbTipoVenda.Text = Trim(TabTemp!DESCRICAO)
            cmbTipoVendaAUX.Text = Trim(TabTemp!TIPOVENDA_ID)
            CC_ID_N = 0 & TabTemp.Fields("cc_id").Value

            If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
               Else: cmbTipoVendaINDR.Text = "False"
            End If

            Call cmbTipoVenda_Click
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      Case vbKeyF5
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from TIPOVENDA WITH (NOLOCK)"
         SQL = SQL & " where descricao = 'Cartao de Credito' "
         SQL = SQL & " and contabiliza = 1 "
         SQL = SQL & " and receber = 'true' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            cmbTipoVenda.Text = Trim(TabTemp!DESCRICAO)
            cmbTipoVendaAUX.Text = Trim(TabTemp!TIPOVENDA_ID)
            CC_ID_N = 0 & TabTemp.Fields("cc_id").Value

            If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
               Else: cmbTipoVendaINDR.Text = "False"
            End If

            Call cmbTipoVenda_Click
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      Case vbKeyF12
         Msg = "Confirma cadastro desse CPF como funcionário ?"
         PERGUNTA Msg, vbYesNo + 32, "Inclusão de Funcionário", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            If TabUSU.State = 1 Then _
               TabUSU.Close

            SQL = "select nome,pessoa_id from USUARIO"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            'SQL = SQL & " where cpf = '" & Trim(CNPJCPF_CLI_A) & "'"
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabUSU.EOF Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close
                  PESSOA_ID_N = 0
   
                  If TabCliente.State = 1 Then _
                     TabCliente.Close
   
                  SQL = "select pessoa_id from PESSOA "
                  SQL = SQL & " where CNPJcpf = '" & Trim(CNPJCPF_CLI_A) & "'"
                  TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabCliente.EOF Then
                     PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
                     Else
                        PESSOA_ID_N = MAX_ID("pessoa_id", "pessoa", "", "", "", "")
   
                        SQL = "insert into PESSOA "
                           SQL = SQL & "(PESSOA_ID,CNPJCPF,DESCRICAO,RAZAO,DATA_CAD,SITUACAO)"
                        SQL = SQL & " values("
                           SQL = SQL & PESSOA_ID_N 'PESSOA_ID
                           SQL = SQL & "," & Trim(CNPJCPF_CLI_A) & "'" 'CNPJCPF
                           SQL = SQL & "," & Trim(txtCli.Text) & "'" 'DESCRICAO
                           SQL = SQL & "," & Trim(txtCli.Text) & "'" 'RAZAO
                           SQL = SQL & "," & Now & "'" 'DATA_CAD
                           SQL = SQL & ",A'" 'SITUACAO
                        SQL = SQL & ")"
                        CONECTA_RETAGUARDA.Execute SQL
                  End If
                  If TabCliente.State = 1 Then _
                     TabCliente.Close
   
                  SQL = "INSERT INTO USUARIO "
                     SQL = SQL & " (empresa_id,usuario_id,Nome,Cpf,Status,Pessoa_id,FUNCIONARIO) "
                  SQL = SQL & " VALUES ("
                     SQL = SQL & EMPRESA_ID_N
                     SQL = SQL & "," & MAX_ID("usuario_id", "usuario", "empresa_id", "1", "", "")
                     SQL = SQL & ",'" & Trim(txtCli.Text) & "'"
                     SQL = SQL & ",'" & Trim(CNPJCPF_CLI_A) & "'"
                     SQL = SQL & ",'TRUE'"
                     SQL = SQL & "," & PESSOA_ID_N
                     SQL = SQL & ",'true'"
                     SQL = SQL & ")"
                  CONECTA_RETAGUARDA.Execute SQL
   
                  MsgBox "Processo realizado com sucesso."
               Else
                  SQL = "update USUARIO set "
                  SQL = SQL & " funcionario = 1"
                  SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
                  'SQL = SQL & " where cpf = '" & Trim(CNPJCPF_CLI_A) & "'"
                  CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
         If TabUSU.State = 1 Then _
            TabUSU.Close
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KEYDOWN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   INDR_FORM_ABERTO = False
End Sub

Private Sub lstFATURA_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstFatura, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "desconto"
         EXCLUIR_TUDO
         TRATA_DESCONTO
      Case "voltar"
         EXCLUIR_TUDO
         Unload frmFatura
      Case "limpar"
'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

         SQL = "UPDATE PEDIDO set "
         SQL = SQL & " status = 2 " 'foi recebido mas ainda não emitiu documento fiscal
         SQL = SQL & ", valor_recebido = 0"
         SQL = SQL & ", Valor_desconto = 0"
         SQL = SQL & ", Perc_desc = 0"
         SQL = SQL & ", valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
         SQL = SQL & ", cliente_id = " & CLIENTE_ID_N

         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         EXCLUIR_TUDO
         SETA_GRID

         LIMPA_BODY

         cmbBandeiraAUX.Text = ""
         cmbBandeiraAUX.Visible = False
         cmbBandeira.Text = ""
         cmbBandeira.Visible = False
         lblBandeira.Visible = False
         VALOR_ITEM_N = 0
         VALOR_ENTRADA_N = 0
         txtValorEntrada.Text = ""
         txtPercEntrada.Text = ""
         cmbTipoVenda.Text = ""
         cmbTipoVendaAUX.Text = ""
         cmbTipoVendaINDR.Text = ""
         FraSeq.Enabled = False
         txtTroco.Text = ""
         txtVendaComDesconto.Text = txtVendaSemDesconto.Text
         txtTxEntrega.Text = ""
         txtTxEntrega.Visible = False
         lblTxEntrega.Visible = False

         MOSTRA_DADOS_PEDIDO

         cmbTipoVenda.SetFocus
      Case "conf"
         If Trim(txtValorEntrada.Text) <> "" Then _
            txtRecebido.Text = txtValorEntrada.Text

         CONFIRMAR_RECEBIMENTO_PARCELADO
         Unload frmFatura
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbTIPOVENDA_LostFocus()
'On Error GoTo ERRO_TRATA

   cmbTipoVenda.BackColor = &HFFFFFF
   If Trim(cmbTipoVendaAUX.Text) <> "" Then
      cmbModalidade.Clear
      cmbModalidadeAux.Clear

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where formapagto_id < 9999 "
      SQL = SQL & " and status = 'true' "
      If Trim(cmbTipoVendaAUX.Text) <> "" Then _
         If IsNumeric(cmbTipoVendaAUX.Text) Then _
            SQL = SQL & " and formapagto_id >= 1 "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         cmbModalidade.AddItem TabConsulta!DESCRICAO
         cmbModalidadeAux.AddItem TabConsulta!FORMAPAGTO_ID
         TabConsulta.MoveNext
      Wend

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_LostFocus"
End Sub

Private Sub cmbTipoVenda_Click()
'On Error GoTo ERRO_TRATA
'set ta passando com liberação de venda, não pode

   lblPRAZO.Caption = ""
   cmbTipoVendaAUX.ListIndex = cmbTipoVenda.ListIndex
   cmbTipoVendaINDR.ListIndex = cmbTipoVenda.ListIndex
   INDR_VENDA_CARTAO = False
   INDR_ERRO_TEF = False

   VALOR_ITEM_N = 0
   VALOR_ENTRADA_N = 0
   INDR_PreFatura = False

   NUMR_PARCELA_N = 0
   DIAS_PRAZO = 0
   DIA_VENCTO = 0
   EXCLUIR_TUDO

   If Trim(cmbTipoVendaAUX.Text) <> "" Then

   If TabTemp.State = 1 Then _
      TabTemp.Close
   
   'vai verificar se é permitido alterar o tipo de recebimento,
   'caso não vai prevalecer registro que foi informado no pedido de venda
   lblFat.Visible = False
   lblFat.Caption = ""
   If ALTERA_FATURA_B = False Then
      'ler dados do pedido pra saber qual tipo de faturamento foi informado durante o pedido de venda

      SQL = "SELECT PEDIDOFATURA.PEDIDO_ID, PEDIDOFATURA.FORMAPAGTO_ID, PEDIDOFATURA.TIPOVENDA_ID, "
      SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescForma, TIPOVENDA.DESCRICAO AS DescTipoVenda"
      SQL = SQL & " FROM PEDIDOFATURA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOFATURA.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID "
      SQL = SQL & " INNER JOIN TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " ON FORMAPAGTO.FORMAPAGTO_ID = TIPOVENDA.FORMAPAGTO_ID"

      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         lblFat.Visible = True
         lblFat.Caption = TabTemp.Fields("tipovenda_id").Value & "-" & TabTemp.Fields("desctipovenda").Value
         If cmbTipoVendaAUX.Text <> TabTemp.Fields("tipovenda_id").Value Then
            MsgBox "Não é permitido alterar faturamento, verificar pedido de venda !!!"
            cmbTipoVendaAUX.Text = "" & TabTemp.Fields("tipovenda_id").Value
            cmbTipoVenda.Text = "" & TabTemp.Fields("desctipovenda").Value
            If TabTemp.State = 1 Then _
               TabTemp.Close
            Exit Sub
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

      If Trim(UCase(cmbTipoVenda.Text)) = "ENCOMENDA" Then
         txtTxEntrega.Visible = True
         lblTxEntrega.Visible = True
      End If
      If Left(UCase(Trim(cmbTipoVenda.Text)), 6) = "CARTAO" Or Left(UCase(Trim(cmbTipoVenda.Text)), 6) = "CARTÃO" Then
         If USA_TEF = True Then
            '========== primeiro checar se a venda com cartao efetivou
            EXCLUIR_TUDO

            If frmDISPLAYEMISSOR.TRATA_RECEBIMENTO_CARTAO(Left(UCase(Trim(cmbTipoVenda.Text)), 6), txtVendaComDesconto.Text) = False Then
               INDR_VENDA_CARTAO = False
               MsgBox "Venda com Cartao nao realizada."
               Exit Sub
            End If

            If INDR_ERRO_TEF = True Then
               INDR_VENDA_CARTAO = False
               Exit Sub
            End If

         End If   'If USA_TEF = True Then
      End If

      CHAMA_FATURAMENTO

      Else
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         MsgBox "Selecione tipo de venda."
         Exit Sub
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Exit Sub
ERRO_TRATA:
   INDR_VENDA_CARTAO = False
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_Click"
End Sub

Private Sub cmbTIPOVENDA_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbTipoVenda.SelStart = 0
   cmbTipoVenda.SelLength = Len(cmbTipoVenda)
   cmbTipoVenda.BackColor = &HC0FFFF

   FraSeq.Enabled = False
   CARREGA_COMBO_TIPO_VENDA

   If Trim(CNPJCPF_CLI_A) <> "99999999999" Then
      MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F3-Convênio", "F4-Cartão de Debito", "F5-Cartão de Credito"
      Else: MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F4-Cartão de Debito", "F5-Cartão de Credito", ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_GotFocus"
End Sub

Private Sub cmbTIPOVENDA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If FraSeq.Enabled = True Then _
         txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPOVENDA_KeyPress"
End Sub

Private Sub cmbmodalidade_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbModalidade.SelStart = 0
   cmbModalidade.SelLength = Len(cmbModalidade)
   cmbModalidade.BackColor = &HC0FFFF

   If Trim(CNPJCPF_CLI_A) <> "99999999999" Then
      MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F3-Convênio", "F4-Cartão de Debito", "F5-Cartão de Credito"
      Else: MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F4-Cartão de Debito", "F5-Cartão de Credito", ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmodalidade_GotFocus"
End Sub

Private Sub cmbModalidade_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF4
         cmbModalidadeAux.ListIndex = cmbModalidade.ListIndex
         
         If Trim(cmbModalidade.Text) <> "" Then
            If Left(UCase(cmbModalidade.Text), 6) = "CHEQUE" Then
               INDR_PRI = True
               frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = False
                  frmCHEQUECADASTRO.txtPORTADOR.Text = Trim(CNPJCPF_CLI_A)
               frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = True
               frmCHEQUECADASTRO.Show 1
               INDR_PRI = False
            End If
         End If
         txtValorItem.SetFocus
   End Select
End Sub

Private Sub cmbMODALIDADE_LostFocus()

   cmbModalidade.BackColor = &HFFFFFF
   If Trim(cmbModalidade.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where descricao = '" & Trim(cmbModalidade.Text) & "'"
      SQL = SQL & " and status = 'true' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         cmbModalidadeAux.Text = "" & TabConsulta!FORMAPAGTO_ID

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

End Sub

Private Sub txtDias_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDias.BackColor = &HFFFFFF
   If txtDias.Text <> "" Then _
      If IsNumeric(txtDias.Text) Then _
         DIAS_PRAZO = txtDias.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDias_LostFocus"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEmis.SelStart = 0
   txtDtEmis.SelLength = Len(txtDtEmis)
   txtDtEmis.BackColor = &HC0FFFF

   txtDtEmis.PromptInclude = True
   If Not IsDate(txtDtEmis.Text) Then
      txtDtEmis.PromptInclude = False
         txtDtEmis.Text = Date
      txtDtEmis.PromptInclude = True
   End If

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

Private Sub cmbCC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then _
        txtDTVENC_KeyPress (13)
End Sub

Private Sub txtDTVENC_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtVenc.SelStart = 0
   txtDtVenc.SelLength = Len(txtDtVenc)
   txtDtVenc.BackColor = &HC0FFFF

   txtDtVenc.PromptInclude = True

   MOSTRA_RODAPE_AQUI "Informe Data Vencimento da parcela", "ESC-Sair", "", "", ""

   If DIAS_PRAZO > 0 Then
      NUMR_SEQ_N = 0 & txtSeq.Text
      DATA_INI = txtDtEmis.Text
      txtDtVenc.Text = DATA_INI + DIAS_PRAZO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTVENC_GotFocus"
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtDtVenc.PromptInclude = True
      If Not IsDate(txtDtVenc.Text) Then
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
      If cmbModalidadeAux.Text = "" Then
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
      VALOR_ITEM_N = txtValorItem.Text

      GRAVAR_TUDO_SEQ
      LIMPA_BODY
      SETA_GRID

      Dim VALOR_VENDA_N As Double
      Dim VALOR_REC_N   As Double

      VALOR_VENDA_N = 0 & txtVendaComDesconto.Text
      VALOR_REC_N = 0 & txtRecebido.Text

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

Private Sub cmbcc_GotFocus()
   cmbCC.SelStart = 0
   cmbCC.SelLength = Len(cmbCC)
   cmbCC.BackColor = &HC0FFFF
End Sub

Private Sub cmbCC_Click()
On Error Resume Next

   cmbCCAux.ListIndex = cmbCC.ListIndex
   Call cmbcc_LostFocus
End Sub

Private Sub cmbcc_LostFocus()

   cmbCC.BackColor = &HFFFFFF
   If Trim(cmbCC.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from DESCR d WITH (NOLOCK)"
      SQL = SQL & " where TIPO = 'O' "
      SQL = SQL & " and codigo = '" & Trim(cmbCCAux.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         cmbCC.Text = "" & TabTemp.Fields("descricao").Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

End Sub

Private Sub txtDTVENC_LostFocus()
   txtDtVenc.BackColor = &HFFFFFF
End Sub

Private Sub txtPedido_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtSeq.Text) <> "" Then
            If IsNumeric(txtSeq.Text) Then
               If TabLancamento.State = 1 Then _
                  TabLancamento.Close
      
               SQL = "select lancamento_id from LANCAMENTO WITH (NOLOCK)"
               SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
               SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabLancamento.EOF Then
                  If Not IsNull(TabLancamento.Fields(0).Value) Then
      
                     Msg = "Confirma Exclusão do Item =  ?" & txtSeq.Text
                     Style = vbYesNo + 32
                     Title = "Atenção !!!"
                     Help = "DEMO.HLP"
                     Ctxt = 1000
                     RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                     If RESPOSTA = vbYes Then
      
                        SQL = "delete from ITEMLANCAMENTO "
                        SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
                        SQL = SQL & " and seq = " & txtSeq.Text
                        CONECTA_RETAGUARDA.Execute SQL
      
                        SETA_GRID
                     End If
                  End If
               End If
               If TabLancamento.State = 1 Then _
                  TabLancamento.Close
      
               BUSCA_LANCAMENTO
            End If
         End If
         txtSeq.SetFocus

   End Select
End Sub

Private Sub txtseq_LostFocus()
   txtSeq.BackColor = &HFFFFFF
End Sub

Private Sub txtVENDEDOR_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtDATA_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtCli_gotfocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtVendaSemDesconto_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtVendaComDesconto_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtRecebido_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txttROCO_GotFocus()
   cmbTipoVenda.SetFocus
End Sub

Private Sub txtValorItem_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorItem.SelStart = 0
   txtValorItem.SelLength = Len(txtValorItem)
   txtValorItem.BackColor = &HC0FFFF

   Dim VALOR_TOTAL_VENDA As Double

   MOSTRA_RODAPE_AQUI "Informe o valor da parcela", "ESC-Sair", "", "", ""

   If Trim(txtVendaSemDesconto.Text) <> "" Then
      VALOR_ITEM_N = txtVendaSemDesconto.Text
      txtValorItem.Text = Format(VALOR_ITEM_N - VALR_DESCONTO_N, strFormatacao2Digitos)

      If Trim(txtVendaComDesconto.Text) <> "" Then
         VALOR_ITEM_N = txtVendaComDesconto.Text
         txtValorItem.Text = Format(VALOR_ITEM_N - VALR_DESCONTO_N, strFormatacao2Digitos)
      End If
   End If

   VALOR_ITEM_N = 0
   VALR_DESCONTO_N = 0

   BUSCA_LANCAMENTO

   VALOR_TOTAL_VENDA = 0 & txtVendaComDesconto.Text
   txtValorItem.Text = Format(VALOR_TOTAL_VENDA - VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   txtValorItem.Refresh
   txtValorItem.SelStart = 0
   txtValorItem.SelLength = Len(txtValorItem.Text)

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
            VALOR_TROCO_N = VALOR_ITEM_N - VALOR_TOTAL_N
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

   txtDias.SelStart = 0
   txtDias.SelLength = Len(txtDias)
   txtDias.BackColor = &HC0FFFF

   MOSTRA_RODAPE_AQUI "Informe quantidade de dias sua vaca ", "ESC-Sair", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdias_GotFocus"
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtEmis.SetFocus
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

   txtSeq.SelStart = 0
   txtSeq.SelLength = Len(txtSeq)
   txtSeq.BackColor = &HC0FFFF

   SETA_GRID

   VALOR_DIFERENCA_N = 0

   MOSTRA_RODAPE_AQUI "Tecle <<ENTER>> para nova seqüência, ou selecione", "ESC-Sair", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtSeq.Text = "" Then
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         NUMR_SEQ_N = 1
         SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l WITH (NOLOCK)"
         SQL = SQL & " where i.numr_doc = " & PEDIDO_ID_N
         SQL = SQL & " and i.numr_doc = l.numr_doc "
         SQL = SQL & " and i.lancamento_id = l.lancamento_id "
         SQL = SQL & " and l.tipo_lancamento = " & INDR_RECEITA
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

            SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l WITH (NOLOCK)"
            SQL = SQL & " where i.numr_doc = " & PEDIDO_ID_N
            SQL = SQL & " and i.numr_doc = l.numr_doc "
            SQL = SQL & " and i.lancamento_id = l.lancamento_id "
            SQL = SQL & " and seq = " & txtSeq.Text
            SQL = SQL & " and l.tipo_lancamento = " & INDR_RECEITA
            TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLancamento.EOF Then
               'valor lançamento
               txtValorItem.Text = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
               VALOR_DIFERENCA_N = TabLancamento!Valor_Item

               If TabDESCR.State = 1 Then _
                  TabDESCR.Close

               'descrição da modalidade
               SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
               SQL = SQL & " where formapagto_id = " & TabLancamento!FORMAPAGTO_ID
               SQL = SQL & " and status = 'true' "
               TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabDESCR.EOF Then
                  cmbModalidade.Text = TabDESCR!DESCRICAO
                  cmbModalidadeAux.Text = TabDESCR!FORMAPAGTO_ID
               End If
               If TabDESCR.State = 1 Then _
                  TabDESCR.Close

               txtDtVenc.PromptInclude = False
               txtDtEmis.PromptInclude = False
               txtDtVenc.Text = TabLancamento!DT_VENCIMENTO
               'txtDTEMIS.Text = data_lancamento
               'else
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

Private Sub txtValorItem_LostFocus()
'On Error GoTo ERRO_TRATA

   txtValorItem.BackColor = &HFFFFFF
   If txtValorItem.Text <> "" Then _
      txtValorItem.Text = Format(txtValorItem.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_LostFocus"
End Sub

Private Sub txtTxEntrega_GotFocus()
'On Error GoTo ERRO_TRATA

   txtTxEntrega.SelStart = 0
   txtTxEntrega.SelLength = Len(txtTxEntrega)
   txtTxEntrega.BackColor = &HC0FFFF

   If Trim(CNPJCPF_CLI_A) <> "99999999999" Then
      MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F3-Convênio", "F4-Cartão de Debito", "F5-Cartão de Credito"
      Else: MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F4-Cartão de Debito", "F5-Cartão de Credito", ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTxEntrega_GotFocus"
End Sub

Private Sub txtTxEntrega_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      'If Trim(UCase(cmbTipoVenda.Text)) = "ENCOMENDA" Then
         VALOR_TX_ENTREGA_N = 0 & txtTxEntrega.Text
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select pedido_id from PEDIDOENCOMENDA WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then
            spEncomenda 2, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_TX_ENTREGA_N
            Else: spEncomenda 1, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_TX_ENTREGA_N
         End If
         If TabVai.State = 1 Then _
            TabVai.Close

      'End If

      cmbTipoVenda.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTxEntrega_KeyPress"
End Sub

Private Sub txtTxEntrega_lostFocus()
'On Error GoTo ERRO_TRATA

   txtTxEntrega.BackColor = &HFFFFFF
   If Trim(txtTxEntrega.Text) <> "" Then
      If IsNumeric(txtTxEntrega.Text) Then
         txtTxEntrega.Text = "" & Format(txtTxEntrega.Text, strFormatacao2Digitos)
         txtTxEntrega.Visible = True
         lblTxEntrega.Visible = True
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTxEntrega_lostFocus"
End Sub

Private Sub txtValorEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorEntrada.SelStart = 0
   txtValorEntrada.SelLength = Len(txtValorEntrada)
   txtValorEntrada.BackColor = &HC0FFFF

   If Trim(CNPJCPF_CLI_A) <> "99999999999" Then
      MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F3-Convênio", "F4-Cartão de Debito", "F5-Cartão de Credito"
      Else: MOSTRA_RODAPE_AQUI "ESC-Sair", "F1-A Vista", "F4-Cartão de Debito", "F5-Cartão de Credito", ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_GotFocus"
End Sub

Private Sub txtValorEntrada_keypress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtValorEntrada.Text) <> "" Then
         VALOR_ENTRADA_N = 0 & txtValorEntrada.Text
         VALOR_TX_ENTREGA_N = 0 & txtTxEntrega
         txtValorEntrada.Text = Format(VALOR_ENTRADA_N, strFormatacao2Digitos)

         VALOR_VENDA_BRUTA_N = 0 & txtVendaSemDesconto.Text
         VALOR_ITEM_N = 0 & txtVendaComDesconto.Text

         VALOR_DESCONTO_N = VALOR_VENDA_BRUTA_N - VALOR_ITEM_N
         VALOR_TOTAL_N = VALOR_VENDA_BRUTA_N - VALOR_DESCONTO_N
         VALOR_RECEBIDO_N = 0 & VALOR_ENTRADA_N

         txtRecebido.Text = "" & Format(VALOR_RECEBIDO_N, strFormatacao2Digitos)
         txtTroco.Text = "" & Format(VALOR_RECEBIDO_N - VALOR_TOTAL_N - VALOR_TX_ENTREGA_N, strFormatacao2Digitos)

         If VALOR_ENTRADA_N >= VALOR_TOTAL_N Then
            cmbModalidade.Clear
            cmbModalidadeAux.Clear

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
            SQL = SQL & " where formapagto_id < 9999 "
            SQL = SQL & " and status = 'true' "
            If Trim(cmbTipoVendaAUX.Text) <> "" Then _
               If IsNumeric(cmbTipoVendaAUX.Text) Then _
                  SQL = SQL & " and formapagto_id >= 1 "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            While Not TabConsulta.EOF
               cmbModalidade.AddItem TabConsulta!DESCRICAO
               cmbModalidadeAux.AddItem TabConsulta!FORMAPAGTO_ID
               TabConsulta.MoveNext
            Wend
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            'txtPercEntrada.Text = Format(VALOR_RECEBIDO_N / VALOR_VENDA_BRUTA_N * 100, strFormatacao2Digitos)
            INDR_ERRO_TEF = False
            cmbTipoVenda.Text = "A VISTA"
            cmbTipoVenda.Refresh
            cmbTipoVendaAUX.Text = 9999
            cmbTipoVendaAUX.Refresh
            txtSeq.Text = 1
            cmbModalidade.Text = "Dinheiro"
            cmbModalidadeAux.Text = 1
            'txtValorItem.Text = VALOR_RECEBIDO_N
            txtValorItem.Text = VALOR_ITEM_N
            txtDias.Text = 0
            txtDtEmis.PromptInclude = False
            txtDtEmis.Text = Date
            txtDtVenc.PromptInclude = False
            txtDtVenc.Text = Date
            cmbTipoVendaINDR.Text = "False"

            If TabAUX.State = 1 Then _
               TabAUX.Close

            SQL = "select permite_desconto from TIPOVENDA WITH (NOLOCK)"
            SQL = SQL & " where TIPOVENDA_id = " & cmbTipoVendaAUX.Text
            SQL = SQL & " and contabiliza = 1 "
            TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabAUX.EOF Then _
               If Not IsNull(TabAUX.Fields(0).Value) Then _
                  cmbTipoVendaINDR.Text = TabAUX.Fields(0).Value
            If TabAUX.State = 1 Then _
               TabAUX.Close

'============================
            If TRATA_RECEBIMENTO = False Then _
               Exit Sub

            FraSeq.Enabled = True
            txtDTVENC_KeyPress (13)
            Else
               txtPercEntrada.Text = Format(((VALOR_ITEM_N / VALOR_TOTAL_N) * 100), strFormatacao2Digitos)
               cmbTipoVenda.SetFocus
         End If
         'Else: txtPercEntrada.SetFocus
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorEntrada_keypress"
End Sub

Private Sub txtValorEntrada_LostFocus()
   txtValorEntrada.BackColor = &HFFFFFF
End Sub

Private Sub txtpercEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   txtPercEntrada.SelStart = 0
   txtPercEntrada.SelLength = Len(txtPercEntrada)
   txtPercEntrada.BackColor = &HC0FFFF
   
   MOSTRA_RODAPE_AQUI "ESC-Sair", "Informe Percentual(%) da Entrada", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercEntrada_GotFocus"
End Sub

Private Sub txtPercEntrada_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPercEntrada.Text <> "" Then
         VALOR_ITEM_N = txtPercEntrada.Text
         txtValorEntrada.Text = Format(((VALOR_ITEM_N * VALOR_TOTAL_N) / 100), strFormatacao2Digitos)
         txtValorEntrada.Refresh
      End If
      cmbTipoVenda.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercEntrada_KeyPress"
End Sub

Private Sub txtpercEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   txtPercEntrada.BackColor = &HFFFFFF
   If txtPercEntrada.Text <> "" Then _
      txtPercEntrada.Text = Format(txtPercEntrada.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercEntrada_LostFocus"
End Sub
'============================================================='subrotinas
Sub GRAVAR_TUDO()
'On Error GoTo ERRO_TRATA

   cmbBandeiraAUX.Visible = False
   cmbBandeira.Visible = False
   lblBandeira.Visible = False
   cmbBandeiraAUX.Text = ""
   cmbBandeira.Text = ""

   If Trim(cmbCCAux.Text) = "" Then _
      cmbCCAux.Text = "NULL"

   'somente para pegar id da pessoa pelo cpf ou cnpj
   If TabPessoa.State = 1 Then _
      TabPessoa.Close
   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_CLI_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then _
      PESSOA_ID_N = TabPessoa.Fields(0).Value
   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   VALOR_TOTAL_LANÇADO = 0 & txtValorEntrada.Text
   VALOR_TOTAL_N = 0 & txtVendaSemDesconto.Text
   VALOR_ITEM_N = 0 & txtVendaComDesconto.Text
   VALOR_DESCONTO_N = VALOR_TOTAL_N - VALOR_ITEM_N

   Dim SITUACAO_A As String
   Dim DT_BAIXA   As String

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from LANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      NUMR_ID_N = TabLancamento!LANCAMENTO_ID
      Else
         NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

         SQL3 = "INSERT INTO LANCAMENTO "
         SQL3 = SQL3 & " ("
            SQL3 = SQL3 & " Lancamento_id, Numr_doc, dt_cad, Tipo_Lancamento, tipovenda_id,pessoa_id,estabelecimento_id) "
         SQL3 = SQL3 & " VALUES ("
            SQL3 = SQL3 & NUMR_ID_N
            SQL3 = SQL3 & "," & PEDIDO_ID_N
            SQL3 = SQL3 & ",'" & Now & "'"
            SQL3 = SQL3 & "," & INDR_RECEITA
            SQL3 = SQL3 & "," & cmbTipoVendaAUX.Text
            SQL3 = SQL3 & "," & PESSOA_ID_N
            SQL3 = SQL3 & "," & ESTABELECIMENTO_ID_N
         SQL3 = SQL3 & ")"
         CONECTA_RETAGUARDA.Execute SQL3
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SITUACAO_A = "A"
   DT_BAIXA = 0
   CRITERIO_A = ""

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and seq = " & txtSeq.Text
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " usu_alt = " & USUARIO_ID_N
      SQL = SQL & ", Dt_Alt = '" & Now & "'"
      SQL = SQL & ", lancamento_id = " & NUMR_ID_N
      SQL = SQL & ", Numr_doc = " & PEDIDO_ID_N
      SQL = SQL & ", Numr_Dp = " & PEDIDO_ID_N
      SQL = SQL & ", Seq = " & txtSeq.Text
      SQL = SQL & ", Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
      SQL = SQL & ", Status = '" & SITUACAO_A & "'"
      SQL = SQL & ", formapagto_id = " & cmbModalidadeAux.Text
      SQL = SQL & ", DT_VENCIMENTO = '" & DMA(txtDtVenc.Text) & "'"
      SQL = SQL & ", cc_id = " & CC_ID_N

      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & txtSeq.Text
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
         SQL = SQL & " (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, seq, "
         SQL = SQL & " Valor_Item, Status, formapagto_id, DT_VENCIMENTO, acerto,cc_id) "
         SQL = SQL & " VALUES ("
            SQL = SQL & USUARIO_ID_N
            SQL = SQL & ",'" & Now & "'"
            SQL = SQL & ",'" & Now & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & txtSeq.Text
            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",'" & SITUACAO_A & "'"
            SQL = SQL & "," & cmbModalidadeAux.Text
            SQL = SQL & ",'" & DMA(txtDtVenc.Text) & "'"
            SQL = SQL & "," & 1
            SQL = SQL & "," & CC_ID_N
         SQL = SQL & " )"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SQL

   If TabLancamento.State = 1 Then _
      TabLancamento.Close
'================================================
   SQL = "select ITEMLANCAMENTO.FORMAPAGTO_ID, FORMAPAGTO.DESCRICAO, FORMAPAGTO.contab_tesora, FORMAPAGTO.BAIXAAUTO"
   SQL = SQL & " from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and Seq = " & txtSeq.Text

   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      If Not IsNull(TabLANCAMENTOITEM.Fields("BAIXAAUTO").Value) Then
         If TabLANCAMENTOITEM.Fields("BAIXAAUTO").Value = True Then

            SQL = "UPDATE ITEMLANCAMENTO SET "
            SQL = SQL & " Status = 'B'"
            SQL = SQL & ", DT_BAIXA = '" & Now & "'"
            SQL = SQL & ", CODG_USU_BAIXA = " & USUARIO_ID_N
            SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
            SQL = SQL & " and Seq = " & txtSeq.Text
            CONECTA_RETAGUARDA.Execute SQL

         End If
      End If
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close
'=========================================
   GRAVA_TX_ENTREGA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & "-" & SQL3, Me.Name, "GRAVAR_TUDO"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstFatura.ListItems.Clear

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   'SQL = "select ITEMLANCAMENTO.SEQ,LANCAMENTO.NUMR_DOC,ITEMLANCAMENTO.VALOR_ITEM, "
   'SQL = SQL & " FORMAPAGTO.DESCRICAO,ITEMLANCAMENTO.DT_VENCIMENTO,TIPOVENDA.PERC_JUROS,"
   'SQL = SQL & " ITEMLANCAMENTO.CC_ID, TIPOVENDA.DESCRICAO AS TipoVenda"
   'SQL = SQL & " from LANCAMENTO "
   'SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   'SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   'SQL = SQL & " INNER JOIN FORMAPAGTO "
   'SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID "
   'SQL = SQL & " INNER JOIN TIPOVENDA "
   'SQL = SQL & " ON FORMAPAGTO.FORMAPAGTO_ID = TIPOVENDA.FORMAPAGTO_ID"


SQL = "select ITEMLANCAMENTO.SEQ, ITEMLANCAMENTO.NUMR_DOC, ITEMLANCAMENTO.VALOR_ITEM, "
SQL = SQL & " FORMAPAGTO.DESCRICAO, ITEMLANCAMENTO.DT_VENCIMENTO, ITEMLANCAMENTO.CC_ID"
SQL = SQL & " from ITEMLANCAMENTO "
SQL = SQL & " INNER JOIN FORMAPAGTO "
SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " where itemlancamento.numr_doc = " & PEDIDO_ID_N

   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLANCAMENTOITEM.EOF
      'sequencia
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set item = lstFatura.ListItems.Add(, "seq." & NUMR_SEQ_N, TabLANCAMENTOITEM!SEQ)
      'numero documento
      item.SubItems(1) = TabLANCAMENTOITEM.Fields("numr_doc").Value
      'valor lançamento
      item.SubItems(2) = Format(TabLANCAMENTOITEM!Valor_Item, strFormatacao2Digitos)
      item.SubItems(3) = Trim(TabLANCAMENTOITEM!DESCRICAO)
      item.SubItems(4) = Now
      item.SubItems(5) = TabLANCAMENTOITEM!DT_VENCIMENTO

      If Not IsNull(TabLANCAMENTOITEM.Fields("CC_ID").Value) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from DESCR d WITH (NOLOCK)"
         SQL = SQL & " where TIPO = 'O' "
         SQL = SQL & " and codigo = '" & Trim(TabLANCAMENTOITEM.Fields("CC_ID").Value) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            item.SubItems(7) = "" & TabTemp.Fields("descricao").Value
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      TabLANCAMENTOITEM.MoveNext
   Wend
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   'BUSCA_LANCAMENTO

   txtVendaSemDesconto.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub LIMPA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   INDR_FUNCIONARIO = False
   lblPemite_Desconto.Visible = False
   lblPRAZO.Caption = ""
   cmbTipoVenda.Text = ""
   cmbTipoVendaAUX.Text = ""
   txtPedido.Text = ""
   txtData.Text = ""
   txtVendaSemDesconto.Text = ""
   txtVendaComDesconto.Text = ""
   txtRecebido.Text = ""
   txtTroco.Text = ""
   txtCli.Text = ""
   cmbModalidadeAux.Clear
   cmbModalidade.Clear
   txtValorItem.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtVenc.PromptInclude = False
   txtDtEmis.Text = ""
   txtDtVenc.Text = ""
   lstFatura.ListItems.Clear
   cmbBandeiraAUX.Text = ""
   cmbBandeiraAUX.Visible = False
   cmbBandeira.Text = ""
   cmbBandeira.Visible = False
   lblBandeira.Visible = False
   txtSeq.Text = ""
   VALOR_TOTAL_LANÇADO = 0
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_LANCAMENTO"
End Sub

Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   VALOR_DIFERENCA_N = 0
   VALOR_ITEM_N = 0
   txtSeq.Text = ""
   cmbModalidadeAux.Text = ""
   cmbModalidade.Text = ""
   txtValorItem.Text = ""
   txtDias.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtVenc.PromptInclude = False
   txtDtEmis.Text = ""
   txtDtVenc.Text = ""
   VALOR_TOTAL_LANÇADO = 0
   FORMAPAGTO_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub BUSCA_LANCAMENTO()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_LANÇADO = 0
   VALOR_RECEBIDO_N = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select sum(valor_item) from ITEMLANCAMENTO i, LANCAMENTO l WITH (NOLOCK)"
   SQL = SQL & " where i.numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento.Fields(0).Value) Then _
         VALOR_TOTAL_LANÇADO = TabLancamento.Fields(0).Value
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   txtRecebido.Text = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   txtRecebido.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_LANCAMENTO"
End Sub

Sub REGISTRA_TROCO()
'On Error GoTo ERRO_TRATA

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select lancamento_id from LANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
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
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REGISTRA_TROCO"
End Sub

Sub EXCLUIR_TUDO()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   SQL = "delete from ITEMLANCAMENTO "
   SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID

   MOSTRA_TOTAIS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_TUDO"
End Sub

Sub MOSTRA_TOTAIS()
'On Error GoTo ERRO_TRATA

   VALOR_ITEM_N = 0

   'BUSCA VALOR TOTAL VENDA
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabTemp.Fields(0).Value) Then _
      VALOR_ITEM_N = TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select valor_recebido,valor_desconto from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      txtTroco.Text = "" & Format(0, strFormatacao2Digitos)
      txtVendaSemDesconto.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      txtVendaComDesconto.Text = "" & Format(VALOR_ITEM_N - TabConsulta.Fields("valor_desconto").Value, strFormatacao2Digitos)
      txtRecebido.Text = "" & Format(TabConsulta.Fields("valor_recebido").Value, strFormatacao2Digitos)
      If Not IsNull(TabConsulta.Fields("valor_recebido").Value) Then _
         If TabConsulta.Fields("valor_recebido").Value > VALOR_ITEM_N Then _
            txtTroco.Text = "" & Format(TabConsulta.Fields("valor_recebido").Value - VALOR_ITEM_N - VALOR_TX_ENTREGA_N, strFormatacao2Digitos)
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAIS"
End Sub

Sub MOSTRA_RODAPE_AQUI(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
'On Error GoTo ERRO_TRATA

   If Trim(Msg1) <> "" Then
      barRodape.Panels.Clear
      barRodape.Panels.Add (1)
      barRodape.Panels(1).Text = Trim(Msg1)
      barRodape.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         barRodape.Panels.Add (2)
         barRodape.Panels(2).Text = Trim(Msg2)
         barRodape.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            barRodape.Panels.Add (3)
            barRodape.Panels(3).Text = Trim(Msg3)
            barRodape.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               barRodape.Panels.Add (4)
               barRodape.Panels(4).Text = Trim(Msg4)
               barRodape.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  barRodape.Panels.Add (5)
                  barRodape.Panels(5).Text = Trim(Msg5)
                  barRodape.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "MOSTRA_RODAPE_AQUI"
End Sub

Sub GERA_FATURAMENTO()
'On Error GoTo ERRO_TRATA

   Dim Valor_Tot_n As Double

   NUMR_PARCELA_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where tipovenda_id = " & cmbTipoVendaAUX.Text
   SQL = SQL & " and contabiliza = 1 "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      CC_ID_N = 0 & TabTemp.Fields("cc_id").Value
      If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
         If TabTemp.Fields("contabiliza").Value = False Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            If Trim(txtRecebido.Text) = "" Then _
               txtRecebido.Text = 0

            MsgBox "Esse Tipo de Venda está configurado para não contabilizar !!!"
            VALOR_RECEBIDO_N = 0 & txtRecebido.Text

'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

            SQL = "update PEDIDO set "
            SQL = SQL & "status = 6 " 'não contabiliza
            SQL = SQL & " , valor_recebido = " & tpMOEDA(VALOR_RECEBIDO_N)
            SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
            SQL = SQL & " , cliente_id = " & CLIENTE_ID_N
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            Exit Sub
            Unload frmFatura
         End If
      End If

      NUMR_PARCELA_N = TabTemp!parcela

      VALOR_DESCONTO_N = 0

      If Not IsNull(TabCabeca.Fields("perc_desc").Value) Then _
         PERC_DESCONTO_N = TabCabeca.Fields("perc_desc").Value

      VALOR_DESCONTO_CABECA_N = 0

      If Not IsNull(TabCabeca.Fields("valor_desconto").Value) Then _
         VALOR_DESCONTO_CABECA_N = TabCabeca.Fields("valor_desconto").Value

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabCabeca!PEDIDO_ID
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      'BUSCA VALOR TOTAL VENDA
      Valor_Tot_n = 0
      SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabCabeca!PEDIDO_ID
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
         Valor_Tot_n = TabPedidoItem.Fields(0).Value
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close
      
      'VALOR_DESCONTO_N = VALOR_DESCONTO_N + (Valor_Tot_n * IIf(PERC_DESCONTO_N > 0, PERC_DESCONTO_N / 100, 1))
      VALOR_DESCONTO_N = VALOR_DESCONTO_N + VALOR_DESCONTO_CABECA_N
      
      VALOR_ITEM_N = 0
      DATA_INI = Date
      If DIA_VENCTO > 0 Then _
         DATA_INI = DIA_VENCTO & "/" & Month(Date) & "/" & Year(Date)
      If NUMR_PARCELA_N > 0 Then _
         VALOR_ITEM_N = (Valor_Tot_n - VALOR_DESCONTO_N) / NUMR_PARCELA_N

      'CABEÇA lançamento
      If TabLancamento.State = 1 Then _
         TabLancamento.Close

      SQL = "select * from LANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
      SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLancamento.EOF Then
         NUMR_ID_N = TabLancamento!LANCAMENTO_ID
         Else
            NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

            SQL = "INSERT INTO LANCAMENTO "
            SQL = SQL & " ("
               SQL = SQL & " Lancamento_id, Numr_doc, dt_cad, Tipo_Lancamento, tipovenda_id,pessoa_id,estabelecimento_id) "
            SQL = SQL & " VALUES ("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & "," & PEDIDO_ID_N
               SQL = SQL & ",'" & Now & "'"
               SQL = SQL & "," & INDR_RECEITA
               SQL = SQL & "," & cmbTipoVendaAUX.Text
               SQL = SQL & "," & PESSOA_ID_N
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      SQL3 = PEDIDO_ID_N
      SqL2 = EMPRESA_ID_N
      CONT_N = 0

      'ITENS
      While CONT_N < NUMR_PARCELA_N
         GRAVA_LANÇAMENTO
         CONT_N = CONT_N + 1
      Wend
      GRAVA_TX_ENTREGA
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FATURAMENTO"
End Sub

Sub GRAVA_LANÇAMENTO()
'On Error GoTo ERRO_TRATA

   Dim SITUACAO_A As String

   SITUACAO_A = "A"
   NUMR_SEQ_N = 1
   If Trim(cmbCCAux.Text) = "" Then _
      cmbCCAux.Text = "NULL"

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l WITH (NOLOCK)"
   SQL = SQL & " where i.numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento!ultimo_reg) Then _
         NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   'tem que dividir os dias de prazo
   If DIA_VENCTO > 0 Then
      'DATA_INI = DIA_VENCTO & "/" & Month(Date) & "/" & Year(Date)
      DATA_INI = DATA_INI + 30
      DATA_INI = DIA_VENCTO & "/" & Month(DATA_INI) & "/" & Year(DATA_INI)
      Else
         If NUMR_PARCELA_N <= 0 Then _
            DATA_INI = DATA_INI + DIAS_PRAZO
   End If

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where seq = " & NUMR_SEQ_N
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
         SQL = SQL & "  usu_alt = " & USUARIO_ID_N
         SQL = SQL & ", Dt_Alt = '" & Now & "'"
         SQL = SQL & ", Numr_doc = " & PEDIDO_ID_N
         SQL = SQL & ", Seq = " & NUMR_SEQ_N
         SQL = SQL & ", Valor_Item = " & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA_N / NUMR_PARCELA_N))
         SQL = SQL & ", Status = '" & SITUACAO_A & "'"
         SQL = SQL & ", formapagto_id = " & TabTemp!FORMAPAGTO_ID
         SQL = SQL & ", DT_VENCIMENTO = '" & DATA_INI & "'"
         SQL = SQL & ", cc_id = " & CC_ID_N
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & NUMR_SEQ_N
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
            SQL = SQL & " (Usu_cad, Dt_cad, Lancamento_id, Numr_doc, "
            SQL = SQL & " NUMR_DP, seq, Valor_Item, Status, formapagto_id, "
            SQL = SQL & " DT_VENCIMENTO, ACERTO,cc_id) "
         SQL = SQL & " VALUES ("
            SQL = SQL & USUARIO_ID_N
            SQL = SQL & ",'" & Now & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & NUMR_SEQ_N
            SQL = SQL & "," & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA_N / NUMR_PARCELA_N))
            SQL = SQL & ",'" & SITUACAO_A & "'"
            SQL = SQL & "," & TabTemp!FORMAPAGTO_ID
            SQL = SQL & ",'" & DATA_INI & "'"
            SQL = SQL & "," & 0
            SQL = SQL & "," & CC_ID_N
         SQL = SQL & ")"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SQL
'================================================
   SQL = "select ITEMLANCAMENTO.FORMAPAGTO_ID, FORMAPAGTO.DESCRICAO, FORMAPAGTO.contab_tesora, FORMAPAGTO.BAIXAAUTO"
   SQL = SQL & " from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and Seq = " & NUMR_SEQ_N

   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      If Not IsNull(TabLANCAMENTOITEM.Fields("BAIXAAUTO").Value) Then
         If TabLANCAMENTOITEM.Fields("BAIXAAUTO").Value = True Then

            SQL = "UPDATE ITEMLANCAMENTO SET "
            SQL = SQL & " Status = 'B'"
            SQL = SQL & ", DT_BAIXA = '" & Now & "'"
            SQL = SQL & ", CODG_USU_BAIXA = " & USUARIO_ID_N
            SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
            SQL = SQL & " and Seq = " & NUMR_SEQ_N
            CONECTA_RETAGUARDA.Execute SQL

         End If
      End If
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close
'=========================================

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_LANÇAMENTO"
End Sub

Sub TRATA_DESCONTO()
'On Error GoTo ERRO_TRATA

   PERC_DESCONTO_USUARIO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   USU_LIBERA_VENDA_N = 0

   If IsNumeric(txtPedido.Text) Then _
      PEDIDO_ID_N = txtPedido.Text

   If INDR_LIBERA_DESCONTO = True Then
      frmVENDADESCONTO.Show 1
      If INDR_DESCONTO_AUTORIZADO = False Then
         MsgBox "Não autorizado !!!"
         Exit Sub
      End If
      Else
         MsgBox "Não permitido"
         Exit Sub
   End If

'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

   'atualizando desconto na cabeça
   SQL = "UPDATE PEDIDO SET "
   SQL = SQL & " Valor_desconto = " & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
   SQL = SQL & " , Perc_desc = " & tpMOEDA(PERC_DESCONTO_N)
   SQL = SQL & " , cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
   SQL = SQL & " , nome_cliente = '" & Trim(txtCli.Text) & "'"
   SQL = SQL & " , status = 2"
   SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
   SQL = SQL & " , vendedor_id = " & VENDEDOR_ID_N
   SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
   SQL = SQL & " , cliente_id = " & CLIENTE_ID_N
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   If Trim(cmbTipoVendaAUX.Text) <> "" Then
      SQL = "update PEDIDOFATURA set "
         SQL = SQL & " tipovenda_id = " & Trim(cmbTipoVendaAUX.Text)
      SQL = SQL & " where pedido_id = " & txtPedido.Text
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If TabVai.State = 1 Then _
      TabVai.Close
   SQL = "select pedido_id from PEDIDOFATURA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVai.EOF Then
      Acao_N = 2
      Else: Acao_N = 1
   End If
   If TabVai.State = 1 Then _
      TabVai.Close

spPEDIDOFATURA Acao_N, 0, PEDIDO_ID_N, TABELAPRECO_ID_N, FORMAPAGTO_ID_N, TIPOVENDA_ID_N

   MOSTRA_TOTAIS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_DESCONTO"
End Sub

Sub CARREGA_COMBO_TIPO_VENDA()
'On Error GoTo ERRO_TRATA

   cmbTipoVenda.Clear
   cmbTipoVendaAUX.Clear
   cmbTipoVendaINDR.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where contabiliza = 1 "
   SQL = SQL & " and receber = 'true' "
   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CC_ID_N = 0 & TabTemp.Fields("cc_id").Value

      cmbTipoVenda.AddItem Trim(TabTemp!DESCRICAO) & "-" & Trim(TabTemp.Fields("TIPOVENDA_ID").Value)
      cmbTipoVendaAUX.AddItem Trim(TabTemp!TIPOVENDA_ID)

      'cmbTipoVendaAUX.AddItem Trim(TabTemp!FORMAPAGTO_ID)
      If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
         cmbTipoVendaINDR.AddItem TabTemp.Fields("permite_desconto").Value
         Else: cmbTipoVendaINDR.AddItem "False"
      End If

      TabTemp.MoveNext
   Wend

   cmbCCAux.Clear
   cmbCC.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'O'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCC.AddItem Trim(TabTemp!DESCRICAO)
      cmbCCAux.AddItem TabTemp!Codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO_TIPO_VENDA"
End Sub

Function TRATA_RECEBIMENTO() As Boolean
'On Error GoTo ERRO_TRATA

   Dim Situacao_Pedido  As Integer
   TRATA_RECEBIMENTO = False

   If INDR_LIBERA_DESCONTO = True Then
      lblPemite_Desconto.Visible = False

      Dim Valor_Sem_Desconto As Double
      Dim Valor_Com_Desconto As Double

Valor_Sem_Desconto = 0 & txtVendaSemDesconto.Text
Valor_Com_Desconto = 0 & txtVendaComDesconto.Text

      If Valor_Sem_Desconto <> Valor_Com_Desconto Then _
         INDR_DESCONTO_AUTORIZADO = True
      If INDR_DESCONTO_AUTORIZADO = True Then 'foi la no trata_desconto formulario frmvendadesconto.frm
         lblPemite_Desconto.Visible = True
         If Trim(UCase(cmbTipoVendaINDR.Text)) = UCase("true") Then
            lblPemite_Desconto.Visible = False
            Else  'tipo de venda não liberada para desconto.
               If Trim(txtRecebido.Text) = "" Then _
                  txtRecebido.Text = 0
'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

               SQL = "UPDATE PEDIDO set "
               SQL = SQL & "status = 2 " 'foi recebido mas ainda não emitiu documento fiscal
               SQL = SQL & " , valor_recebido = 0"
               SQL = SQL & " , Valor_desconto = 0"
               SQL = SQL & " , Perc_desc = 0"
               SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
               SQL = SQL & " , cliente_id = " & CLIENTE_ID_N

               SQL = SQL & " where pedido_id = " & txtPedido.Text
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               CONECTA_RETAGUARDA.Execute SQL

               txtVendaComDesconto.Text = txtVendaSemDesconto.Text
               txtRecebido.Text = ""
               txtTroco.Text = ""
               INDR_DESCONTO_AUTORIZADO = False

               MsgBox "Atenção, tipo venda não liberada para desconto, verifique !!!"
               Exit Function
         End If
         Else
            txtVendaComDesconto.Text = txtVendaSemDesconto.Text
            txtRecebido.Text = ""
            INDR_DESCONTO_AUTORIZADO = False
            If Trim(txtRecebido.Text) = "" Then _
               txtRecebido.Text = 0
'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close


            SQL = "UPDATE PEDIDO set "
            SQL = SQL & "status = 2 " 'foi recebido mas ainda não emitiu documento fiscal
            SQL = SQL & " , valor_recebido = 0"
            SQL = SQL & " , Valor_desconto = 0"
            SQL = SQL & " , Perc_desc = 0"
            SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
            SQL = SQL & " , cliente_id = " & CLIENTE_ID_N

            SQL = SQL & " where pedido_id = " & txtPedido.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            'Exit Function
      End If
      Else
         txtVendaComDesconto.Text = txtVendaSemDesconto.Text
         If Trim(txtRecebido.Text) = "" Then _
            txtRecebido.Text = 0

         INDR_DESCONTO_AUTORIZADO = False

         Situacao_Pedido = 2
         If INDR_PreFatura = True Then _
            Situacao_Pedido = 8
'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

         SQL = "UPDATE PEDIDO set "
         SQL = SQL & "status = " & Situacao_Pedido
         SQL = SQL & " , valor_recebido = " & tpMOEDA(txtRecebido.Text)
         SQL = SQL & " , Valor_desconto = 0"
         SQL = SQL & " , Perc_desc = 0"
         SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
         SQL = SQL & " , cliente_id = " & CLIENTE_ID_N

         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         'Exit Function
   End If

   TRATA_RECEBIMENTO = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRATA_RECEBIMENTO"
End Function

Sub CONFIRMAR_RECEBIMENTO_NA_SEQUENCIA()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO

   VALOR_TOTAL_LANÇADO = 0 & txtValorEntrada.Text
   VALOR_TOTAL_N = 0 & txtVendaSemDesconto.Text
   VALOR_ITEM_N = 0 & txtVendaComDesconto.Text
   VALOR_DESCONTO_N = VALOR_TOTAL_N - VALOR_ITEM_N
   VALOR_TOTAL_N = VALOR_TOTAL_N - VALOR_DESCONTO_N
   VALOR_RECEBIDO_N = 0 & txtRecebido.Text

   If Round(VALOR_RECEBIDO_N) >= Round(VALOR_TOTAL_N) Then
      If Left(UCase(cmbModalidade.Text), 8) = UCase("Dinheiro") Then
         Msg = "Confirma recebimento ?"
         PERGUNTA Msg, vbYesNo + 32, "Recebimento ", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            If TabCabeca.State = 1 Then _
               TabCabeca.Close
   
            SQL = "select status from PEDIDO WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & txtPedido.Text
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCabeca.EOF Then
               If TabCabeca!STATUS = 2 Then
                  VALOR_RECEBIDO_N = 0 & txtRecebido.Text
'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

                  SQL = "UPDATE PEDIDO set "
                  SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
                  SQL = SQL & " , valor_recebido = " & tpMOEDA(VALOR_RECEBIDO_N)
                  SQL = SQL & " , valor_desconto = " & tpMOEDA(VALOR_DESCONTO_N)
                  SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
                  SQL = SQL & " , cliente_id = " & CLIENTE_ID_N

                  SQL = SQL & " where pedido_id = " & txtPedido.Text
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            Unload frmFatura

            Exit Sub
         End If
         Else
            If VALOR_RECEBIDO_N >= VALOR_TOTAL_N Then
               If txtTroco.Text <> "" Then _
                  If VALOR_TROCO_N > 0 Then _
                     REGISTRA_TROCO

               Msg = "Confirma recebimento ?"
               'PERGUNTA Msg, vbYesNo + 32, "Recebimento ", "DEMO.HLP", 1000

               If CONFIRMA_PERGUNTA(Msg, vbYesNo + 32, "Recebimento ", "DEMO.HLP", 1000) = True Then
                 If TabCabeca.State = 1 Then _
                    TabCabeca.Close
   
                  SQL = "select * from PEDIDO WITH (NOLOCK)"
                  SQL = SQL & " where pedido_id = " & txtPedido.Text
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabCabeca.EOF Then
                     If TabCabeca!STATUS = 2 Then
                        VALOR_RECEBIDO_N = 0 & txtRecebido.Text
'SO PRA GARANTIR
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

                        SQL = "UPDATE PEDIDO set "
                        SQL = SQL & "status = 5 " 'foi recebido mas ainda não emitiu documento fiscal
                        SQL = SQL & " , valor_recebido = " & tpMOEDA(VALOR_RECEBIDO_N)
                        SQL = SQL & " , valor_desconto = " & tpMOEDA(VALOR_DESCONTO_N)
                        SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
                        SQL = SQL & " , cliente_id = " & CLIENTE_ID_N

                        SQL = SQL & " where pedido_id = " & txtPedido.Text
                        SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                        CONECTA_RETAGUARDA.Execute SQL
                     End If
                  End If
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close

                  Unload frmFatura

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

Sub CONFIRMAR_RECEBIMENTO_PARCELADO()
'On Error GoTo ERRO_TRATA

   BUSCA_LANCAMENTO

   Dim Situacao_Pedido  As Integer

   'VALOR_TOTAL_LANÇADO = 0 & txtValorEntrada.Text
   VALOR_TOTAL_N = 0 & txtVendaSemDesconto.Text
   VALOR_ITEM_N = 0 & txtVendaComDesconto.Text
   VALOR_DESCONTO_N = VALOR_TOTAL_N - VALOR_ITEM_N
   VALOR_TOTAL_N = VALOR_TOTAL_N - VALOR_DESCONTO_N
   VALOR_RECEBIDO_N = 0 & txtRecebido.Text

   If VALOR_TOTAL_LANÇADO <= 0 Then _
      VALOR_TOTAL_LANÇADO = txtRecebido.Text

   VALOR_TOTAL_LANÇADO = Format(VALOR_TOTAL_LANÇADO, strFormatacao2Digitos)
   VALOR_TOTAL_N = Format(VALOR_TOTAL_N, strFormatacao2Digitos)

   If Round(VALOR_TOTAL_LANÇADO) >= Round(VALOR_TOTAL_N) Then
   'If Format(Round(VALOR_TOTAL_LANÇADO), strFormatacao2Digitos) >= Format(Round(VALOR_TOTAL_N), strFormatacao2Digitos) Then
      INDR_FINALIZA_RECEBIMENTO = False
      Msg = "Confirma recebimento ?"
      Situacao_Pedido = 5

      If Trim(UCase(cmbTipoVenda.Text)) = "ENCOMENDA" Then
         VALOR_TX_ENTREGA_N = 0 & txtTxEntrega.Text
         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select pedido_id from PEDIDOENCOMENDA WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then
            spEncomenda 2, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_TX_ENTREGA_N
            Else: spEncomenda 1, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_TX_ENTREGA_N
         End If
         If TabVai.State = 1 Then _
            TabVai.Close

         Msg = "Confirma Agendamento de Encomenda ?"
         Situacao_Pedido = 8
      End If
      PERGUNTA Msg, vbYesNo + 32, "Recebimento NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         txtPedido.Text = PEDIDO_ID_N

         If TabVai.State = 1 Then _
            TabVai.Close
         SQL = "select cliente_id from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLI_A) & "'"
         SQL = SQL & " and status = 'A'"
         TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVai.EOF Then _
            CLIENTE_ID_N = TabVai.Fields("cliente_id").Value
         If TabVai.State = 1 Then _
            TabVai.Close

         If txtPedido.Text = "" Then _
            txtPedido.Text = 1
'set aqui nao precia lar, update direto
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from PEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & txtPedido.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            If TabCabeca!STATUS = 2 Then
               SQL = "UPDATE PEDIDO set "
               SQL = SQL & " status = " & Situacao_Pedido
               '5 = foi recebido mas ainda não emitiu documento fiscal
               '8 = agendamento de encomenda
               SQL = SQL & " , valor_recebido = " & tpMOEDA(VALOR_RECEBIDO_N)
               SQL = SQL & " , valor_desconto = " & tpMOEDA(VALOR_DESCONTO_N)
               SQL = SQL & " , valor_total = " & tpMOEDA(txtVendaSemDesconto.Text)
               SQL = SQL & " , cliente_id = " & CLIENTE_ID_N

               SQL = SQL & " where pedido_id = " & txtPedido.Text
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

GRAVA_CARTAO_PEDIDO cmbTipoVenda.Text

'MsgBox Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(3).Text)

         'If Trim(frmDISPLAYEMISSOR.lstPedidos.SelectedItem.ListSubItems.item(3).Text) = "ENCOMENDA" Then
         If Trim(TIPO_PEDIDO_A) = "ENCOMENDA" Then
            If TabVai.State = 1 Then _
               TabVai.Close

            SQL = "select * from ENTREGA WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVai.EOF Then
               SQL = "update ENTREGA SET "
                  SQL = SQL & " DT_ENTREGA = '" & Now & "'"                                     'DT_ENTREGA
                  SQL = SQL & ", entregador_ID = " & USUARIO_ID_N                               'entregador_ID
                  SQL = SQL & ", entregador = '" & Trim(TRAZ_NOME_USUARIO(USUARIO_ID_N)) & "'"  'entregador
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabVai.State = 1 Then _
               TabVai.Close

            VALOR_TX_ENTREGA_N = 0 & txtTxEntrega.Text
            If TabVai.State = 1 Then _
               TabVai.Close
            SQL = "select pedido_id from PEDIDOENCOMENDA WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVai.EOF Then
               spEncomenda 2, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_TX_ENTREGA_N
               Else: spEncomenda 1, 0, PEDIDO_ID_N, USUARIO_ID_N, VALOR_TX_ENTREGA_N
            End If
            If TabVai.State = 1 Then _
               TabVai.Close
         End If

         LIMPA_LANCAMENTO
TIPO_PEDIDO_A = ""

         'Unload FRMFATURA
         frmFatura.Hide
         Else
            txtRecebido.Text = 0
            txtTroco.Text = 0
            txtSeq.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONFIRMAR_RECEBIMENTO_PARCELADO"
End Sub

Sub MOSTRA_DADOS_PEDIDO()
'On Error GoTo ERRO_TRATA

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, "
   SQL = SQL & " PEDIDO.VENDEDOR_ID, PEDIDO.CGCCPF, PEDIDO.DT_REQ, "
   SQL = SQL & " PEDIDO.STATUS, PEDIDO.TIPO_REGISTRO, PEDIDO.VALOR_DESCONTO, "
   SQL = SQL & " PEDIDO.PERC_DESC, PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_RECEBIDO,"
   SQL = SQL & " PEDIDO.VALOR_TOTAL, PEDIDO.NUMERO_CAIXA_CPU, PEDIDO.ESTABELECIMENTO_ID, "
   SQL = SQL & " CLIENTE.STATUS AS sit_cli, CLIENTE.LIMITE_CREDITO, Cliente.CPFFUNCCONVENIO, "
   SQL = SQL & " Cliente.PERC_DESC_CONVENIO, PESSOA.PESSOA_ID, PESSOA.Descricao"
   SQL = SQL & " from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN CLIENTE WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID "
   SQL = SQL & " INNER JOIN PESSOA  WITH (NOLOCK)"
   SQL = SQL & " ON CLIENTE.PESSOA_ID = PESSOA.PESSOA_ID "

   SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      PESSOA_ID_N = TabCabeca.Fields("pessoa_id").Value
      VENDEDOR_ID_N = TabCabeca.Fields("vendedor_id").Value
      
      CNPJCPF_CLI_A = Trim(TabCabeca.Fields("CGCCPF").Value)

      If Trim(CNPJCPF_CLI_A) <> "" Then _
         txtCNPJCPF.Text = FORMATA_CNPJCPF(CNPJCPF_CLI_A)

      txtCli.Text = Trim(TabCabeca.Fields("nome_cliente").Value)
      PEDIDO_ID_N = Trim(TabCabeca.Fields("pedido_id").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select descricao from vwVendedor WITH (NOLOCK) "
      SQL = SQL & " where vendedor_id = " & VENDEDOR_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtVendedor.Text = TabConsulta!DESCRICAO
         txtVendedor.Refresh
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      PERC_DESCONTO_N = 0
      VALOR_DESCONTO_N = 0
      VALOR_DESCONTO_CABECA_N = 0
      TOTAL_DESCONTO_N = 0
      VALOR_DESCONTO_ITEM_N = 0
      VALOR_DESCONTO_CABECA_N = 0 & TabCabeca.Fields("valor_desconto").Value

      'PEGANDO DESCONTO INDIVIDUAL ITEM
      SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabCabeca.Fields("pedido_id").Value
      'SQL = SQL & " and tipo_reg = 'PC' "
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_ITEM_N = TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'BUSCA VALOR TOTAL VENDA
      VALOR_ITEM_N = 0
      SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & TabCabeca.Fields("pedido_id").Value
      'SQL = SQL & " and tipo_reg = 'PC' "
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_ITEM_N = TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TOTAL_DESCONTO_N = VALOR_DESCONTO_CABECA_N + VALOR_DESCONTO_ITEM_N
      VALOR_TOTAL_N = VALOR_ITEM_N - TOTAL_DESCONTO_N

      If TOTAL_DESCONTO_N > 0 And VALOR_ITEM_N > 0 Then _
         PERC_DESCONTO_N = 0 & (TOTAL_DESCONTO_N / VALOR_ITEM_N) * 100

      txtVendaSemDesconto.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      txtVendaComDesconto.Text = Format((VALOR_ITEM_N - TOTAL_DESCONTO_N), strFormatacao2Digitos)
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   If TabCabeca.State = 1 Then _
      TabCabeca.Close
   SQL = "select VLR_TX_ENTREGA from PEDIDOENCOMENDA WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      If Not IsNull(TabCabeca.Fields(0).Value) Then _
         txtTxEntrega.Text = "" & Format(TabCabeca.Fields(0).Value, strFormatacao2Digitos)
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close
   
   If Trim(txtTxEntrega.Text) <> "" Then
      If IsNumeric(txtTxEntrega.Text) Then
         txtTxEntrega.Text = "" & Format(txtTxEntrega.Text, strFormatacao2Digitos)
         txtTxEntrega.Visible = True
         lblTxEntrega.Visible = True
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PEDIDO"
End Sub

Sub GRAVAR_TUDO_SEQ()
'On Error GoTo ERRO_TRATA

   If Trim(cmbCCAux.Text) = "" Then _
      cmbCCAux.Text = "NULL"

   'somente para pegar id da pessoa pelo cpf ou cnpj
   If TabPessoa.State = 1 Then _
      TabPessoa.Close
   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_CLI_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then _
      PESSOA_ID_N = TabPessoa.Fields(0).Value
   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   'PEGANDO TOTAL DESCONTO
   VALOR_VENDA_BRUTA_N = 0 & txtVendaSemDesconto.Text
    VALOR_ITEM_N = 0 & txtVendaComDesconto.Text
   VALOR_DESCONTO_N = VALOR_VENDA_BRUTA_N - VALOR_ITEM_N

   VALOR_TOTAL_N = VALOR_VENDA_BRUTA_N - VALOR_DESCONTO_N

   VALOR_ITEM_N = 0 & txtValorItem.Text

   Dim SITUACAO_A As String
   Dim DT_BAIXA   As String

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from LANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      NUMR_ID_N = TabLancamento!LANCAMENTO_ID
      Else
         NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

         SQL3 = "INSERT INTO LANCAMENTO "
         SQL3 = SQL3 & " ("
            SQL3 = SQL3 & " Lancamento_id, Numr_doc, dt_cad, Tipo_Lancamento, tipovenda_id,pessoa_id,estabelecimento_id ) "
         SQL3 = SQL3 & " VALUES ("
            SQL3 = SQL3 & NUMR_ID_N
            SQL3 = SQL3 & "," & PEDIDO_ID_N
            SQL3 = SQL3 & ",'" & Now & "'"
            SQL3 = SQL3 & "," & INDR_RECEITA
            SQL3 = SQL3 & "," & cmbTipoVendaAUX.Text
            SQL3 = SQL3 & "," & PESSOA_ID_N
            SQL3 = SQL3 & "," & ESTABELECIMENTO_ID_N
         SQL3 = SQL3 & ")"
         CONECTA_RETAGUARDA.Execute SQL3
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SITUACAO_A = "A"
   DT_BAIXA = 0
   CRITERIO_A = ""

   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and seq = " & txtSeq.Text
   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " usu_alt = " & USUARIO_ID_N
      SQL = SQL & ", Dt_Alt = '" & Now & "'"
      SQL = SQL & ", lancamento_id = " & NUMR_ID_N
      SQL = SQL & ", Numr_doc = " & PEDIDO_ID_N
      SQL = SQL & ", Numr_Dp = " & PEDIDO_ID_N
      SQL = SQL & ", Seq = " & txtSeq.Text
      SQL = SQL & ", Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
      SQL = SQL & ", Status = '" & SITUACAO_A & "'"
      SQL = SQL & ", formapagto_id = " & cmbModalidadeAux.Text
      SQL = SQL & ", DT_VENCIMENTO = '" & DMA(txtDtVenc.Text) & "'"
      SQL = SQL & ", cc_id = " & CC_ID_N

      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & txtSeq.Text
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
         SQL = SQL & " (Usu_Alt, Dt_Alt, Dt_Cad, Lancamento_id, Numr_doc, NUMR_DP, seq, "
         SQL = SQL & " Valor_Item, Status, formapagto_id, DT_VENCIMENTO, acerto,cc_id) "
         SQL = SQL & " VALUES ("
            SQL = SQL & USUARIO_ID_N
            SQL = SQL & ",'" & Now & "'"
            SQL = SQL & ",'" & Now & "'"
            SQL = SQL & "," & NUMR_ID_N
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & txtSeq.Text
            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",'" & SITUACAO_A & "'"
            SQL = SQL & "," & cmbModalidadeAux.Text
            SQL = SQL & ",'" & DMA(txtDtVenc.Text) & "'"
            SQL = SQL & "," & 1
            SQL = SQL & "," & CC_ID_N
         SQL = SQL & " )"
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

   CONECTA_RETAGUARDA.Execute SQL

   If TabLancamento.State = 1 Then _
      TabLancamento.Close
'====baixa titulo automatico se estiver parametrizado
   SQL = "select ITEMLANCAMENTO.FORMAPAGTO_ID, FORMAPAGTO.DESCRICAO, "
   SQL = SQL & " FORMAPAGTO.contab_tesora, FORMAPAGTO.BAIXAAUTO"
   SQL = SQL & " from ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
   SQL = SQL & " and Seq = " & txtSeq.Text

   TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLANCAMENTOITEM.EOF Then
      If Not IsNull(TabLANCAMENTOITEM.Fields("BAIXAAUTO").Value) Then
         If TabLANCAMENTOITEM.Fields("BAIXAAUTO").Value = True Then

            SQL = "UPDATE ITEMLANCAMENTO SET "
            SQL = SQL & " Status = 'B'"
            SQL = SQL & ", DT_BAIXA = '" & Now & "'"
            SQL = SQL & ", CODG_USU_BAIXA = " & USUARIO_ID_N
            SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
            SQL = SQL & " and Seq = " & txtSeq.Text
            CONECTA_RETAGUARDA.Execute SQL

         End If
      End If
   End If
   If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close
'=========================================

   'SQL = "select sum(valor_item) from ITEMLANCAMENTO WITH (NOLOCK)"
   'SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
   'TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'If Not TabLANCAMENTOITEM.EOF Then _
      If Not IsNull(TabLANCAMENTOITEM.Fields(0).Value) Then _
         txtRecebido.Text = Format(TabLANCAMENTOITEM.Fields(0).Value, strFormatacao2Digitos)
   'If TabLANCAMENTOITEM.State = 1 Then _
      TabLANCAMENTOITEM.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & "-" & SQL3, Me.Name, "GRAVAR_TUDO_SEQ"
End Sub

Sub CHAMA_FATURAMENTO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where tipovenda_id = " & cmbTipoVendaAUX.Text
   SQL = SQL & " and contabiliza = 1 "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      FORMAPAGTO_ID_N = TabConsulta.Fields("formapagto_id").Value
      CC_ID_N = 0 & TabConsulta.Fields("cc_id").Value
      If Not IsNull(TabConsulta.Fields("prefatura").Value) Then _
         INDR_PreFatura = TabConsulta.Fields("prefatura").Value
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      'descrição da modalidade
      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where formapagto_id = " & FORMAPAGTO_ID_N
      SQL = SQL & " and status = 'true' "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then
         cmbModalidade.Text = TabDESCR!DESCRICAO
         cmbModalidadeAux.Text = TabDESCR!FORMAPAGTO_ID

         If TabDESCR.Fields("FUNC").Value = True Then
            If Trim(UCase(TabDESCR!DESCRICAO)) = UCase("VALE") Or _
               Trim(UCase(TabDESCR!DESCRICAO)) = UCase("refeicao") Or _
               Trim(UCase(TabDESCR!DESCRICAO)) = UCase("refeição") Or _
               Trim(UCase(TabDESCR!DESCRICAO)) = UCase("refeiçao") Or _
               Trim(UCase(TabDESCR!DESCRICAO)) = UCase("refeicão") Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select funcionario from USUARIO "
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               'SQL = SQL & " where cpf = '" & Trim(CNPJCPF_CLI_A) & "'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then _
                  If Not IsNull(TabTemp.Fields("funcionario").Value) Then _
                     INDR_FUNCIONARIO = TabTemp.Fields("funcionario").Value
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               If INDR_FUNCIONARIO = False Then
                  If TabDESCR.State = 1 Then _
                     TabDESCR.Close
                  If TabConsulta.State = 1 Then _
                     TabConsulta.Close

                  MsgBox "Tipo de Venda permitida somente para funcionários."

                  cmbModalidade.Text = ""
                  cmbModalidadeAux.Text = ""

                  Exit Sub
               End If
            End If
         End If
      End If
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      If Not IsNull(TabConsulta!parcela) Then _
         NUMR_PARCELA_N = 0 & TabConsulta!parcela
      If Not IsNull(TabConsulta!PRAZO) Then _
         DIAS_PRAZO = 0 & TabConsulta!PRAZO
      If Not IsNull(TabConsulta!DIAVENCTO) Then _
         DIA_VENCTO = 0 & TabConsulta!DIAVENCTO
   End If

   If DIA_VENCTO > 0 Then
      'lblDiaVencto.Visible = True
      'txtDiaVencto.Visible = True
   End If
'============================
   If TRATA_RECEBIMENTO = False Then _
      Exit Sub

   FraSeq.Enabled = True
   If DIAS_PRAZO > 0 Then
      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      'GERA TITULOS
      SQL = "select * from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabeca.EOF Then
         GERA_FATURAMENTO
         SETA_GRID
         Else
            If TabCabeca.State = 1 Then _
               TabCabeca.Close
            MsgBox "Pedido não encontrado."
            Exit Sub
      End If
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
      Else 'If DIAS_PRAZO > 0 Then
         NUMR_SEQ_N = 1

         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select lancamento_id from LANCAMENTO  WITH (NOLOCK)"
         SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
         SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then _
            If Not IsNull(TabLancamento.Fields(0).Value) Then _
               NUMR_SEQ_N = MAX_ID("seq", "itemlancamento", "lancamento_id", TabLancamento.Fields(0).Value, "", "")
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         txtSeq.Text = NUMR_SEQ_N
         VALOR_ITEM_N = 0 & txtVendaComDesconto.Text
         VALOR_ENTRADA_N = 0 & txtValorEntrada.Text
         
         If Trim(cmbModalidadeAux.Text) = "" Then _
            cmbModalidadeAux.Text = 1

         txtDtVenc.PromptInclude = False
            txtDtVenc.Text = Date
         txtDtVenc.PromptInclude = True

         If Trim(txtRecebido.Text) = "" Then
            txtRecebido.Text = Format(txtVendaComDesconto.Text, strFormatacao2Digitos)
            txtTroco.Text = Format(0, strFormatacao2Digitos)
         End If

         GRAVAR_TUDO
   End If

   SETA_GRID
   LIMPA_BODY

   CONFIRMAR_RECEBIMENTO_PARCELADO

   lblDiaVencto.Visible = False
   txtDiaVencto.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & "-" & SQL3, Me.Name, "CHAMA_FATURAMENTO"
End Sub

Sub GRAVA_CARTAO_PEDIDO(TIPO_VENDA_A As String)
'On Error GoTo ERRO_TRATA

   'GRAVANDO RELAÇÃO DA VENDA PEDIDO COM CARTÃO DEBITO/CREDITO/BANDEIRA CARTÃO
   If Trim(TIPO_VENDA_A) <> "" Then
      If Left(UCase(Trim(TIPO_VENDA_A)), 6) = "CARTAO" Or _
         Left(UCase(Trim(TIPO_VENDA_A)), 6) = "CARTÃO" Or _
         Left(UCase(Trim(TIPO_VENDA_A)), 3) = "POS" Then

         Dim TabCartao As New ADODB.Recordset
         SQL = ""

         If TabCartao.State = 1 Then _
            TabCartao.Close

         SQL = "select cartaopedido_id from CARTAOPEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & txtPedido.Text
         TabCartao.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCartao.EOF Then
            SQL = "update CARTAOPEDIDO set "
               SQL = " BANDEIRA_ID = '" & Trim(cmbBandeiraAUX.Text) & "'"
               SQL = SQL & ", CNPJ_CARTAO = '" & Trim(CNPJ_CRED_CARTAO_ESTAB) & "'"
            SQL = SQL & " where cartaopedido_id = " & TabCartao.Fields("cartaopedido_id").Value
            Else
               SQL = "insert into CARTAOPEDIDO "
               SQL = SQL & "("
                  SQL = SQL & "CARTAOPEDIDO_ID,PEDIDO_ID,BANDEIRA_ID,CNPJ_CARTAO"
               SQL = SQL & ")"
               SQL = SQL & " values("
                  SQL = SQL & MAX_ID("CARTAOPEDIDO_ID", "CARTAOPEDIDO", "", "", "", "")
                  SQL = SQL & "," & txtPedido.Text
                  SQL = SQL & ",'" & Trim(cmbBandeiraAUX.Text) & "'"
                  SQL = SQL & ",'" & Trim(CNPJ_CRED_CARTAO_ESTAB) & "'"
               SQL = SQL & ")"
         End If
         If TabCartao.State = 1 Then _
            TabCartao.Close

         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & "-" & SQL3 & " ; " & "Perdeu conexão com o banco MEGASIM", Me.Name, "GRAVA_CARTAO_PEDIDO"
End Sub

Sub GRAVA_TX_ENTREGA()
'On Error GoTo ERRO_TRATA

   If Trim(txtTxEntrega.Text) <> "" Then
      If IsNumeric(txtTxEntrega.Text) Then
         VALOR_TX_ENTREGA_N = 0 & txtTxEntrega.Text

         If VALOR_TX_ENTREGA_N > 0 Then
            If TabLancamento.State = 1 Then _
               TabLancamento.Close
   
            SQL = "select lancamento_id from LANCAMENTO WITH (NOLOCK)"
            SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
            SQL = SQL & " and tipo_lancamento = " & INDR_RECEITA
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabLancamento.EOF Then
               NUMR_ID_N = 0 & TabLancamento.Fields(0).Value

               If TabLANCAMENTOITEM.State = 1 Then _
                  TabLANCAMENTOITEM.Close

               SQL = "select formapagto_id from FORMAPAGTO WITH (NOLOCK)"
               SQL = SQL & " where descricao = 'TAXA ENTREGA'"
               TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabLANCAMENTOITEM.EOF Then _
                  FORMAPAGTO_ID_N = 0 & TabLANCAMENTOITEM.Fields(0).Value
               If TabLANCAMENTOITEM.State = 1 Then _
                  TabLANCAMENTOITEM.Close

               SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
               SQL = SQL & " where lancamento_id = " & NUMR_ID_N
               SQL = SQL & " and formapagto_id = " & FORMAPAGTO_ID_N
               TabLANCAMENTOITEM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabLANCAMENTOITEM.EOF Then
                  NUMR_SEQ_N = 0 & TabLANCAMENTOITEM.Fields("seq").Value

                  SQL = "UPDATE ITEMLANCAMENTO SET "
                     SQL = SQL & "  usu_alt = " & USUARIO_ID_N
                     SQL = SQL & ", Dt_Alt = '" & Now & "'"
                     SQL = SQL & ", Numr_doc = " & PEDIDO_ID_N
                     SQL = SQL & ", Seq = " & NUMR_SEQ_N
                     SQL = SQL & ", Valor_Item = " & Str(Format(VALOR_TX_ENTREGA_N, strFormatacao2Digitos))
                     SQL = SQL & ", Status = 'B'"
                     SQL = SQL & ", formapagto_id = " & TabTemp!FORMAPAGTO_ID
                     SQL = SQL & ", DT_VENCIMENTO = '" & Now & "'"
                     SQL = SQL & ", DT_baixa = '" & Now & "'"
                     SQL = SQL & ", cc_id = " & CC_ID_N
                  SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
                  SQL = SQL & " and Seq = " & NUMR_SEQ_N
                  Else
                     NUMR_SEQ_N = 0 & MAX_ID("SEQ", "ITEMLANCAMENTO", "LANCAMENTO_ID", Str(NUMR_ID_N), "", "")

                     SQL = "INSERT INTO ITEMLANCAMENTO "
                        SQL = SQL & " (Usu_cad, Dt_cad, Lancamento_id, Numr_doc, "
                        SQL = SQL & " NUMR_DP, seq, Valor_Item, Status, formapagto_id, "
                        SQL = SQL & " DT_VENCIMENTO, ACERTO,cc_id,dt_baixa) "
                     SQL = SQL & " VALUES ("
                        SQL = SQL & USUARIO_ID_N
                        SQL = SQL & ",'" & Now & "'"
                        SQL = SQL & "," & NUMR_ID_N
                        SQL = SQL & "," & PEDIDO_ID_N
                        SQL = SQL & "," & PEDIDO_ID_N
                        SQL = SQL & "," & NUMR_SEQ_N
                        SQL = SQL & "," & Str(Format(VALOR_TX_ENTREGA_N, strFormatacao2Digitos))
                        SQL = SQL & ",'B'"
                        SQL = SQL & "," & FORMAPAGTO_ID_N
                        SQL = SQL & ",'" & Now & "'"
                        SQL = SQL & "," & 0
                        SQL = SQL & "," & CC_ID_N
                        SQL = SQL & ",'" & Now & "'"
                     SQL = SQL & ")"
               End If
               If TabLANCAMENTOITEM.State = 1 Then _
                  TabLANCAMENTOITEM.Close

               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If

      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & "-" & SQL3, Me.Name, "GRAVA_TX_ENTREGA"
End Sub
