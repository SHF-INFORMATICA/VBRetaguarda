VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmNOTADISPLAY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Nota Fiscal"
   ClientHeight    =   8115
   ClientLeft      =   2280
   ClientTop       =   2565
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NOTADISPLAY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Vendedor"
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
      Height          =   855
      Left            =   7050
      TabIndex        =   29
      Top             =   2550
      Visible         =   0   'False
      Width           =   3915
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
         Left            =   75
         TabIndex        =   30
         Top             =   285
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbVend 
         Height          =   360
         Left            =   75
         TabIndex        =   6
         Top             =   285
         Width           =   3765
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
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   60
      TabIndex        =   21
      Top             =   690
      Width           =   5055
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   6
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtReq 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbSTATUS 
         Height          =   360
         ItemData        =   "NOTADISPLAY.frx":47C4A
         Left            =   3000
         List            =   "NOTADISPLAY.frx":47C54
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NFe:"
         Height          =   240
         Left            =   420
         TabIndex        =   35
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         Height          =   240
         Left            =   165
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   240
         Left            =   2265
         TabIndex        =   22
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   5070
      TabIndex        =   19
      Top             =   690
      Width           =   5865
      Begin VB.OptionButton optFor 
         Caption         =   "&Fornecedor"
         Height          =   240
         Left            =   2400
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optCli 
         Caption         =   "&Cliente"
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCli 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   20
         Top             =   660
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtCGCCPF 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   660
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
   End
   Begin VB.TextBox txtTotalVenda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   2520
      TabIndex        =   18
      Top             =   7650
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   1800
      Width           =   10875
      Begin VB.ComboBox cmbCFOP 
         Height          =   360
         Left            =   1080
         TabIndex        =   5
         Text            =   "-------------------- Selecione ----------------------"
         Top             =   240
         Width           =   9705
      End
      Begin VB.ComboBox cmbAuxCFOP 
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
         Height          =   345
         Left            =   2670
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.TextBox txtDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   2520
      TabIndex        =   14
      Top             =   7150
      Width           =   1935
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   9000
      TabIndex        =   13
      Top             =   7150
      Width           =   1935
   End
   Begin VB.TextBox txtCotas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   9000
      TabIndex        =   12
      Top             =   7650
      Width           =   1935
   End
   Begin VB.Frame Frame3 
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
      Height          =   1095
      Left            =   60
      TabIndex        =   7
      Top             =   2550
      Width           =   5145
      Begin VB.OptionButton optBaixa 
         Caption         =   "Data &Cancelamento"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   2175
      End
      Begin VB.OptionButton optEmis 
         Caption         =   "Data &Emissão"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   360
         Left            =   3480
         TabIndex        =   0
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   360
         Left            =   3480
         TabIndex        =   1
         Top             =   680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   11
         Top             =   700
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   10
         Top             =   300
         Width           =   1020
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":47C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":480C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":483DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":48832
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":48C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":48FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":493FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":4971A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":4BECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":4C322
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":4CD34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":4D746
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NOTADISPLAY.frx":4E158
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1270
      ButtonWidth     =   2540
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Relatório"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reimprimir"
            Key             =   "reimp"
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
               Picture         =   "NOTADISPLAY.frx":4EB6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTADISPLAY.frx":4FD04
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTADISPLAY.frx":50D93
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTADISPLAY.frx":5202F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTADISPLAY.frx":5313A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTAASS 
      Height          =   3105
      Left            =   60
      TabIndex        =   24
      Top             =   3840
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   5477
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Nota"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Serie"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Valor Venda"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor Desc."
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Dt.Emis."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Forma Pagto"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Vendedor"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "CFOP"
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
      MaxFontSize     =   100
      DesignWidth     =   10995
      DesignHeight    =   8115
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   0
      X2              =   11040
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   11040
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Faturado = "
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
      Left            =   6975
      TabIndex        =   28
      Top             =   7200
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desconto = "
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
      Left            =   420
      TabIndex        =   27
      Top             =   7200
      Width           =   1995
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Geral = "
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
      Left            =   930
      TabIndex        =   26
      Top             =   7680
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Notas Emitidas = "
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
      Left            =   6285
      TabIndex        =   25
      Top             =   7680
      Width           =   2610
   End
End
Attribute VB_Name = "frmNOTADISPLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmNOTADISPLAY)
   Me.Caption = Me.Caption & " - " & Me.Name

   cmbAuxCFOP.Clear
   cmbCFOP.Clear
   SQL = "select * from DESCR "
   SQL = SQL & " where tipo_a = 'H' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbAuxCFOP.AddItem TabDESCR!Codigo
      cmbCFOP.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!desc_a)
      TabDESCR.MoveNext
   Wend
   TabDESCR.Close

   cmbVend.Clear
   cmbAuxVend.Clear
   SQL = "select * from USUARIO "
   SQL = SQL & " where tipo=2 " 'vendedores
   SQL = SQL & " order by nome "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbVend.AddItem TabDESCR!NOME
      cmbAuxVend.AddItem TabDESCR.Fields("USUARIO_ID").Value
      TabDESCR.MoveNext
   Wend
   TabDESCR.Close

   cmbStatus.Clear
   cmbStatus.AddItem "Emitidas"
   cmbStatus.AddItem "Canceladas"
   
   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub LISTAASS_Click()
On Error Resume Next

   NUMR_REQ_N = LISTAASS.SelectedItem.Text
   CRITERIO = LISTAASS.SelectedItem.ListSubItems.Item(LISTAASS.ColumnHeaders(1).Position)
   SQL3 = LISTAASS.SelectedItem.ListSubItems.Item(LISTAASS.ColumnHeaders(2).Position)
   Err.Clear
End Sub

Private Sub optEmis_Click()
'On Error GoTo ERRO_TRATA

   optEmis.Value = True
   optBaixa.Value = False
   txtDtIni.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optEmis_Click"
End Sub

Private Sub optbaixa_Click()
'On Error GoTo ERRO_TRATA

   optEmis.Value = False
   optBaixa.Value = True
   txtDtIni.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optbaixa_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "reimp"
         If Not IsNull(LISTAASS.SelectedItem.Text) Then
            CRITERIO = LISTAASS.SelectedItem.Text
            If IsNumeric(CRITERIO) Then
               NUMR_REQ_N = CRITERIO

               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from NF "
               SQL = SQL & " where numr_req = " & NUMR_REQ_N
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp!Status) Then
                     If TabTemp!Status = "C" Then
                        TabTemp.Close
                        MsgBox "Operação não permitida, nota fiscal cancelada."
                        Exit Sub
                     End If
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from CABECAREQ "
               SQL = SQL & " where numr_req = " & NUMR_REQ_N
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp!Status) Then
                     If TabTemp!Status = 9 Then
                        TabTemp.Close
                        MsgBox "Operação não permitida, Pedido cancelada."
                        Exit Sub
                     End If
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               frmNOTAGERA.Show 1
            End If
         End If
         'CONSULTA_TUDO
      Case "consultar"
         CONSULTA_TUDO
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
      Case "print"
         Dim rstNF As New ADODB.Recordset
         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
            FORMULA_REL = ""

            SQL = "select * from NF "
            SQL = SQL & " where nf_tipo = '" & "S" & "'"
            SQL = SQL & " and dt_emissao >= '" & DMA(txtDtIni.Text) & "'"
            SQL = SQL & " and dt_emissao <= '" & DMA(txtDtFim.Text) & "'"
            SQL = SQL & " order by numr_nota"
            rstNF.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not rstNF.EOF Then
               FORMULA_REL = "{NF.NF_TIPO} = '" & "S" & "'"
               If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
                  If FORMULA_REL <> "" Then
                     FORMULA_REL = FORMULA_REL & " and {NF.dt_emissao} >= DATE (" & Year(txtDtIni.Text) & "," & Month(txtDtIni.Text) & "," & Day(txtDtIni.Text) & ")"
                     FORMULA_REL = FORMULA_REL & " and {NF.dt_emissao} <= DATE (" & Year(txtDtFim.Text) & "," & Month(txtDtFim.Text) & "," & Day(txtDtFim.Text) & ")"
                  End If
               End If
               If FORMULA_REL <> "" Then
                  SqL2 = "+" & "{NF.numr_nota}"
ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
                  Nome_Relatorio = "rel_nf_emitidas.rpt"
                  frmRELATORIO10.Show 1
                End If
            End If
            rstNF.Close
            Else
               MsgBox "Digite Periodo de Emissao Para Emitir este Relatorio Sr: Usuario"
               txtDtIni.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbCFOP_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxCFOP.ListIndex = cmbCFOP.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCFOP_Click"
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

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Clientes"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   If CPF_N <> "" Then
      txtCGCCPF.Text = CPF_N
      CPF_N = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_GotFocus"
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
      SQL = "select * from CLIENTE "
      SQL = SQL & " where CGCCPF = '" & txtCGCCPF.Text & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCGCCPF.SetFocus
         Exit Sub
         Else: If TabCliente!NOME <> "" _
               Then txtCli.Text = TabCliente!NOME
      End If
      txtCli.SetFocus
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
      txtDtIni.PromptInclude = True
      If Not IsDate(txtDtIni.Text) Then
         txtDtIni.PromptInclude = False
            txtDtIni.Text = Date
         txtDtIni.PromptInclude = True
      End If
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
      txtDtFim.PromptInclude = True
      If Not IsDate(txtDtFim.Text) Then
         txtDtFim.PromptInclude = False
            txtDtFim.Text = Date
         txtDtFim.PromptInclude = True
      End If
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtnota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnota_KeyPress"
End Sub

Private Sub txtReq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'CONSULTA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtReq_KeyPress"
End Sub

Private Sub CONSULTA_TUDO()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SQL = "select * from NF "
   SQL = SQL & " where numr_nota > 0 "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N

   If txtReq.Text <> "" Then
      SQL = SQL & " and numr_req = " & txtReq.Text
      
   End If

   txtCGCCPF.PromptInclude = False
   If txtCGCCPF.Text <> "" Then
      SQL = SQL & " and prop = '" & Trim(txtCGCCPF.Text) & "'"

   End If
   txtCGCCPF.PromptInclude = True

   If cmbAuxVend.Text <> "" Then
      SQL = SQL & " and c.vendedor_id = " & cmbAuxVend.Text
      CRITERIO = CRITERIO & " and {CABECAREQ.vendedor_id} = " & cmbVend.Text
   End If

   'cfop
   If cmbAuxCFOP.Text <> "" Then
      SQL = SQL & " and cfop = '" & Trim(cmbAuxCFOP.Text) & "'"

   End If

   'statux
   If Trim(cmbStatus.Text) <> "" Then _
      SQL = SQL & " and status = '" & Left(cmbStatus.Text, 1) & "'"

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   If ((optEmis.Value = False) And (optBaixa.Value = False)) Then _
      optEmis.Value = True

   If ((IsDate(txtDtIni.Text)) And (IsDate(txtDtFim.Text))) Then
      If optEmis.Value = True Then
         SQL = SQL & " and dt_emissao >= '" & DMA(txtDtIni.Text) & "'"
         SQL = SQL & " and dt_emissao <= '" & DMA(txtDtFim.Text) & "'"
      End If
   End If
   If txtNota.Text <> "" Then _
      SQL = SQL & " and numr_nota = " & txtNota.Text

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_TUDO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   txtDesconto.Text = ""
   txtTot.Text = ""
   txtCotas.Text = ""
   txtTotalVenda.Text = ""

   VALOR_TOTAL_N = 0
   CONT_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_DESCONTO_CABECA_N = 0
   
   LISTAASS.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If TabTemp.EOF Then _
      MsgBox "Não existe notas emitidas para esses parametros."

   While Not TabTemp.EOF
      Set Item = LISTAASS.ListItems.Add(, "seq." & TabTemp.Fields("numr_req").Value & CONT_N, TabTemp.Fields("numr_req").Value)
      Item.SubItems(1) = TabTemp!NUMR_NOTA
      Item.SubItems(2) = TabTemp!SERIE_NOTA

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select nome from CLIENTE "
      SQL = SQL & " where cgccpf = '" & TabTemp.Fields("prop") & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         Item.SubItems(3) = TabCliente!NOME
         Else
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select nome from FORNECEDOR "
            SQL = SQL & " where cgccpf = '" & TabTemp.Fields("prop") & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               Item.SubItems(3) = TabCliente!NOME
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      'aqui pega o total da venda
      VALOR_ITEM_N = 0
      SQL = "select sum(valor_item*qtd_pedida) "
      SQL = SQL & " FROM CABECAREQ "
      SQL = SQL & " INNER JOIN ITEMREQ "
      SQL = SQL & " ON CABECAREQ.PEDIDO_ID = ITEMREQ.PEDIDO_ID"
      SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and CABECAREQ.numr_req = " & TabTemp.Fields("numr_req").Value
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabCliente.Fields(0).Value) Then
         VALOR_TOTAL_N = VALOR_TOTAL_N + TabCliente.Fields(0).Value
         VALOR_ITEM_N = TabCliente.Fields(0).Value
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      PERC_DESCONTO_N = 0
      SQL = "select perc_desc from CABECAREQ "
      SQL = SQL & " where numr_req = " & TabTemp.Fields("numr_req").Value
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then _
         If Not IsNull(TabCliente.Fields(0).Value) Then _
            PERC_DESCONTO_N = TabCliente.Fields(0).Value
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select valor_desconto from CABECAREQ "
      SQL = SQL & " where numr_req = " & TabTemp.Fields("numr_req").Value
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then _
         If Not IsNull(TabCliente.Fields(0).Value) Then _
            VALOR_DESCONTO_CABECA_N = TabCliente.Fields(0).Value
      If TabCliente.State = 1 Then _
         TabCliente.Close

      'aqui o desconto vem valor, se houve desconto individual, Nao Existe desconto no item
      'VALOR_DESCONTO_N = 0
      'SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from ITEMREQ "
      'SQL = SQL & " where empresa_id  = " & EMPRESA_ID_n_n
      'SQL = SQL & " and numr_req = " & TABTEMP.Fields("c.numr_req").Value
      'TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      'If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
      '   VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
      'TabPedidoItem.Close

      'desconto da tabela cabecareq                ***
      VALOR_DESCONTO_N = VALOR_DESCONTO_N + (VALOR_ITEM_N * PERC_DESCONTO_N / 100)
      
      Item.SubItems(4) = Format(VALOR_ITEM_N, strFormatacao2Digitos)
      Item.SubItems(5) = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

      CONT_N = CONT_N + 1
            
      Item.SubItems(6) = TabTemp!DT_EMISSAO

      'SQL = "select * from TIPOVENDA "
      'SQL = SQL & " where tipovenda_id = " & TabTemp!tipovenda_id
      'TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      'If Not TabCliente.EOF Then _
      '   Item.SubItems(7) = TabCliente!Descricao
      If TabCliente.State = 1 Then _
         TabCliente.Close

      If Not IsNull(TabTemp.Fields("cfop").Value) Then
         If Trim(TabTemp.Fields("cfop").Value) <> "" Then
            If IsNumeric(TabTemp.Fields("cfop").Value) Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close

               SQL = "select * from CFOP "
               SQL = SQL & " where codigo = " & Trim(TabTemp.Fields("cfop").Value)
               TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabUSU.EOF Then _
                  Item.SubItems(10) = Trim(TabUSU.Fields("codigo").Value) & "-" & Trim(TabUSU.Fields("descricao").Value)

               If TabUSU.State = 1 Then _
                  TabUSU.Close
            End If
         End If
      End If

      If TabUSU.State = 1 Then _
         TabUSU.Close

      'SQL = "select * from VENDEDOR "
      'SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id")
      'TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      'If Not TabUSU.EOF Then _
         Item.SubItems(8) = TabUSU!NOME_VEND
      'If TabUSU.State = 1 Then _
         TabUSU.Close

      Item.SubItems(9) = ""

      If Not IsNull(TabTemp.Fields("Status")) Then
         If TabTemp.Fields("Status") = 1 Then _
            Item.SubItems(9) = "Vendido"
         If TabTemp.Fields("Status") = 2 Then _
            Item.SubItems(9) = "Lançamento"
         If TabTemp.Fields("Status") = "E" Then _
            Item.SubItems(9) = "Emitida"
         If TabTemp.Fields("Status") = "N" Then _
            Item.SubItems(9) = "Venda Nao Registrada"
      End If

      If Not IsNull(TabTemp.Fields("Status")) Then _
         If TabTemp.Fields("Status") <> "" Then _
            If TabTemp.Fields("Status") = "C" Then _
               Item.SubItems(9) = "Cancelado"

      txtTotalVenda.Text = Format(VALOR_TOTAL_N, "currency")
      txtDesconto.Text = Format(VALOR_DESCONTO_N, "currency")
      txtTot.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, "currency")
      txtCotas.Text = CONT_N
      txtCotas.Refresh
      txtTot.Refresh
      txtTotalVenda.Refresh
      txtDesconto.Refresh
      TabTemp.MoveNext
   Wend
   TabTemp.Close
   txtTot.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   cmbAuxCFOP.Text = ""
   cmbCFOP.Text = ""
   cmbStatus.Text = ""
   optEmis.Value = False
   optBaixa.Value = False
   LISTAASS.ListItems.Clear
   txtReq.Text = ""
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   txtCli.Text = ""
   cmbVend.Text = ""
   cmbAuxVend.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtNota.Text = ""
   txtTotalVenda.Text = ""
   txtDesconto.Text = ""
   txtTot.Text = ""
   txtCotas.Text = ""
   txtReq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub
