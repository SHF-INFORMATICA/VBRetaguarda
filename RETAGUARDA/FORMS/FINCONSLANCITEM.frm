VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFINCONSULTAFATURA 
   Caption         =   "Consulta Faturas"
   ClientHeight    =   8475
   ClientLeft      =   1515
   ClientTop       =   1905
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FINCONSLANCITEM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRecibo 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   4080
      Picture         =   "FINCONSLANCITEM.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Recibo"
      Top             =   8040
      Width           =   405
   End
   Begin VB.TextBox txtTotaljuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   10620
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   8070
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   8070
      Width           =   1575
   End
   Begin VB.TextBox txtQtde 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   8070
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1350
      Left            =   60
      TabIndex        =   15
      Top             =   705
      Width           =   12255
      Begin VB.CommandButton cmdForCli 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3120
         Picture         =   "FINCONSLANCITEM.frx":69A5
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbCCaux 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   9360
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbAuxModalidade 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4680
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9360
         TabIndex        =   26
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox cmbModalidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4680
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtLanc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox cmbSituacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9360
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "CentroCusto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Título:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Pgto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Situação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   60
      TabIndex        =   9
      Top             =   1920
      Width           =   12255
      Begin VB.ComboBox cmbEstabAUX 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox cmbEstab 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5160
         TabIndex        =   41
         Top             =   840
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         TabIndex        =   36
         Top             =   720
         Width           =   2295
         Begin Threed.SSOption optSintetico 
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   120
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
         End
         Begin Threed.SSOption optAnalitico 
            Height          =   255
            Left            =   1200
            TabIndex        =   38
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
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
            Value           =   -1
         End
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbVendAux 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   9960
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbVend 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   9840
         TabIndex        =   32
         ToolTipText     =   "Selecione um vendedor"
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optVencimento 
         Caption         =   "Dt.&Venc."
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optBaixa 
         Caption         =   "Dt.&Baixa"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optEmissao 
         Caption         =   "Dt.&Emis."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   315
         Left            =   5130
         TabIndex        =   8
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   315
         Left            =   7620
         TabIndex        =   10
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Estab.:"
         Height          =   240
         Left            =   3960
         TabIndex        =   40
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   9780
         TabIndex        =   33
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         Height          =   225
         Left            =   6570
         TabIndex        =   12
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         Height          =   225
         Left            =   3990
         TabIndex        =   11
         Top             =   360
         Width           =   1020
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   1270
      ButtonWidth     =   2963
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
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixar Titulo"
            Key             =   "baixa"
            Object.ToolTipText     =   "Baixar Título"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
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
         Left            =   11160
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   11520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":73A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":8541
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":95D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":A585
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":B690
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":C90E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":DDBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FINCONSLANCITEM.frx":F178
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstTitulos 
      Height          =   4575
      Left            =   45
      TabIndex        =   14
      ToolTipText     =   "Clique para selecionar"
      Top             =   3240
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8070
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Título"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Parcela"
         Object.Width           =   1960
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Forma Pagto."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Juros"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dias Atrazo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor Atualizado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Vencimento"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Dt.Baixa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "pedido"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CNPJCPF"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Estabelecimento"
         Object.Width           =   2540
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12345
      DesignHeight    =   8475
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   12360
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor Com Juros = "
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
      Index           =   1
      Left            =   8430
      TabIndex        =   25
      Top             =   8100
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   4620
      TabIndex        =   23
      Top             =   8100
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total Registro(s) = "
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
      TabIndex        =   21
      Top             =   8100
      Width           =   2145
   End
End
Attribute VB_Name = "frmFINCONSULTAFATURA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim NossoNumero      As String
   Dim SequenciaRemessa As Long
   Dim intQtdTitulos    As Long
   Dim VlrTotalTitulos  As Double
   Dim CANCELA_LOOP     As Boolean

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   CANCELA_LOOP = False
   PESSOA_ID_N = 0
   MOSTRA_VENDEDORES

   Label15.Visible = False
   cmbEstab.Visible = False
   cmbEstabAUX.Clear
   cmbEstab.Clear
   cmbEstab.AddItem "Todos"
   cmbEstabAUX.AddItem ""

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select count(ESTABELECIMENTO_id) from ESTABELECIMENTO WITH (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      If TabDESCR.Fields(0).Value > 1 Then
         cmbEstab.Visible = True

         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      
         SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
         SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
         SQL = SQL & " order by DESCRICAO"
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabDESCR.EOF
            cmbEstab.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
            cmbEstabAUX.AddItem Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
            TabDESCR.MoveNext
         Wend
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      End If
   End If
   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstab.Text = "" & TRAZ_ESTABELECIMENTO(cmbEstabAUX.Text)

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
'======================
   'txtDtIni.PromptInclude = False
   'txtDtFim.PromptInclude = False
   'txtDtIni.Text = Date
   'txtDtFim.Text = Date
   'txtDtIni.PromptInclude = True
   'txtDtFim.PromptInclude = True

   CRITERIO_A = Month(Date)
   If Len(CRITERIO_A) = 1 Then _
      CRITERIO_A = "0" & CRITERIO_A

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False

   txtDtIni.Text = "01/" & CRITERIO_A & "/" & Year(Date)

   txtDtIni.PromptInclude = True
   CRITERIO_A = FimDoMes(txtDtIni.Text, False)
   CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
   txtDtFim.Text = CRITERIO_A

If UCase(NOME_BANCO_DADOS) <> UCase("SHFINFO") Then
   If MULT_EMPRESA_B = True Then
      txtDtIni.Text = Date
      txtDtFim.Text = Date
   End If
End If

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Activate()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", ""

   If INDR_RECEITA = "2" Then
      frmFINCONSULTAFATURA.Caption = "Consulta Contas à Pagar"
      Label4.Caption = "Fornecedor :"
      'optBaixa.Caption = "Data Pagamento"
   End If

   If INDR_RECEITA = "1" Then
      frmFINCONSULTAFATURA.Caption = "Consulta Contas à Receber"
      Label4.Caption = "Cliente :"
      'optBaixa.Caption = "Data Recebimento"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Activate"
End Sub

Private Sub Form_Resize()
'On Error GoTo ERRO_TRATA

   'MODALIDADE

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbAuxModalidade.Clear
   cmbModalidade.Clear

   SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " where formapagto_id < 9999 "
   SQL = SQL & " and status = 'true' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbModalidade.AddItem TabDESCR!DESCRICAO
      cmbAuxModalidade.AddItem TabDESCR!FORMAPAGTO_ID
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbSituacao.Clear
   cmbSituacao.AddItem "Aberto"
   cmbSituacao.AddItem "Baixado"
   cmbSituacao.AddItem "Cancelado"
   cmbSituacao.ListIndex = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Resize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         CANCELA_LOOP = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub cmdRecibo_Click()
'On Error GoTo ERRO_TRATA

   If lstTitulos.ListItems.Count > 0 Then
      Dim i
      For i = lstTitulos.ListItems.Count To 1 Step -1
         If lstTitulos.ListItems(i).Checked = True Then
            If Trim(lstTitulos.ListItems(i).Text) <> "" And Trim(lstTitulos.ListItems(i).SubItems(1)) <> "" Then
               FORMULA_REL = "{ITEMLANCAMENTO.numr_doc} = " & Trim(lstTitulos.ListItems(i).Text)
               FORMULA_REL = FORMULA_REL & " and {ITEMLANCAMENTO.seq} = " & Trim(lstTitulos.ListItems(i).SubItems(1))

               If chkImp.Value = 1 Then _
                  ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

               Nome_Relatorio = "reciboshf.rpt"
               frmRELATORIO10.Show 1
            End If
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdRecibo_Click"
End Sub

Private Sub chkTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstTitulos.ListItems.Count > 0 Then
      For i = lstTitulos.ListItems.Count To 1 Step -1
         If chkTodos.Value = 1 Then
            lstTitulos.ListItems(i).Checked = True
            Else: lstTitulos.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "chkTodos_Click"
End Sub

Private Sub cmdForCli_Click()
   If INDR_RECEITA = 1 Then
      TIPO_PESSOA_CADASTRO = "CLIENTE"
      frmPessoaConsulta.Show 1
   End If
   If INDR_RECEITA = 2 Then
      TIPO_PESSOA_CADASTRO = "FORNECEDOR"
      frmPessoaConsulta.Show 1
   End If
   If Trim(CNPJCPF_A) <> "" Then _
      txtCNPJCPF.Mask = "##############" 'verificar com Sergio

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = CNPJCPF_A
   CNPJCPF_A = ""

   MostraCliente
   txtCNPJCPF.SetFocus
End Sub

Private Sub LSTTITULOS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstTitulos, ColumnHeader
End Sub

Private Sub LSTTITULOS_DblClick()
'On Error GoTo ERRO_TRATA

    If Not IsNull(lstTitulos.SelectedItem.Text) Then
        If INDR_RECEITA = 1 Then
           CRITERIO_A = lstTitulos.SelectedItem.Text
           Indr_Consulta = True
           Unload Me
        Else
           If INDR_RECEITA = 2 Then
              CRITERIO_A = lstTitulos.SelectedItem.Text
              Indr_Consulta = True
              'frmNOTAENTRADA.Show 1
              Indr_Consulta = False
              'CRITERIO_A = ""
              Unload Me
           End If
        End If
    End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LSTTITULOS_DblClick"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "print"
         MONTA_REL
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
      Case "consultar"
         CONSULTA_TUDO
         Me.Enabled = True
      Case "baixa"
         BAIXA_TITULOS_SELECIONADOS
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbMODALIDADE_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxModalidade.ListIndex = cmbModalidade.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub cmbCC_Click()
'On Error GoTo ERRO_TRATA

   cmbCCAux.ListIndex = cmbCC.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMODALIDADE_Click"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then _
      txtCNPJCPF.Mask = "##############"

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         If INDR_RECEITA = 1 Then
            TIPO_PESSOA_CADASTRO = "CLIENTE"
            frmPessoaConsulta.Show 1
         End If
         If INDR_RECEITA = 2 Then
            TIPO_PESSOA_CADASTRO = "FORNECEDOR"
            frmPessoaConsulta.Show 1
         End If

         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Mask = "##############" 'verificar com Sergio

         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""

         MostraCliente
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         txtCli.Text = TabCliente!NOME
         CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value
         PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
         Else
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

            SQL = "select * from vwFornecedor WITH (NOLOCK)"
            SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
            TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFornecedor.EOF Then
               txtCli.Text = Trim(TabFornecedor.Fields("descricao").Value)
               PESSOA_ID_N = TabFornecedor.Fields("pessoa_id").Value
               Else
                  MsgBox "Cliente não cadastrado."
                  txtCNPJCPF.SetFocus
            End If
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSituacao.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub optEmissao_Click()
   txtDtIni.SetFocus
End Sub

Private Sub optVencimento_Click()
   txtDtIni.SetFocus
End Sub

Private Sub optBaixa_Click()
   txtDtIni.SetFocus
End Sub

Private Sub TXTDTFIM_GotFocus()
   txtDtFim.PromptInclude = False
   If txtDtFim.Text = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDtFim_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
         txtDtFim.Text = Date
      txtDtFim.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtfim_LostFocus"
End Sub

Private Sub TXTDTINI_GotFocus()
   txtDtIni.PromptInclude = False
   If txtDtIni.Text = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
End Sub

Private Sub txtDtIni_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   If Not IsDate(txtDtIni.Text) Then
      txtDtIni.PromptInclude = False
         txtDtIni.Text = Date
      txtDtIni.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtIni_LostFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
   Exit Sub

ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVend_Click"
End Sub
'====================================
Private Sub CONSULTA_TUDO()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""

   SQL = "select LANCAMENTO.LANCAMENTO_ID, LANCAMENTO.PESSOA_ID, LANCAMENTO.ESTABELECIMENTO_ID, LANCAMENTO.NUMR_DOC, LANCAMENTO.TIPO_LANCAMENTO, lancamento.dt_cad, "
   SQL = SQL & " LANCAMENTO.tipovenda_id, ITEMLANCAMENTO.SEQ, ITEMLANCAMENTO.FORMAPAGTO_ID, ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.STATUS,LANCAMENTO.NOME_PESSOA,"
   SQL = SQL & " ITEMLANCAMENTO.DT_VENCIMENTO, ITEMLANCAMENTO.DT_BAIXA, ITEMLANCAMENTO.DT_CANCELA, ITEMLANCAMENTO.VALOR_DESCONTO, ITEMLANCAMENTO.PERC_DESCONTO,"
   SQL = SQL & " ITEMLANCAMENTO.NUMR_DP , ITEMLANCAMENTO.CC_ID, ITEMLANCAMENTO.HISTORICO, PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO"
   SQL = SQL & " from LANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON LANCAMENTO.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " where tipo_lancamento = " & INDR_RECEITA
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   If Trim(cmbEstab.Text) <> "" Then _
      If Trim(UCase(cmbEstab.Text)) <> UCase("todos") Then _
         SQL = SQL & " and estabelecimento_id = " & cmbEstabAUX.Text

   'Numero Pedido
   If IsNumeric(txtLanc.Text) Then _
      SQL = SQL & " and LANCAMENTO.Numr_doc = '" & txtLanc.Text & "'"

   'Tipo pagamento
   If cmbAuxModalidade.Text <> "" Then _
      SQL = SQL & " and formapagto_id = " & cmbAuxModalidade.Text

   'Situacao
   If cmbSituacao.Text <> "" Then _
      SQL = SQL & " and status = '" & Left(cmbSituacao.Text, 1) & "'"

   'CLIENTE
   If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and LANCAMENTO.pessoa_id = " & PESSOA_ID_N

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optEmissao.Value = True Then
         SQL = SQL & " and lancamento.dt_cad >= '" & DMA(txtDtIni.Text, "i") & "'"
         SQL = SQL & " and lancamento.dt_cad <= '" & DMA(txtDtFim.Text, "f") & "'"
      End If
      If optVencimento.Value = True Then
         SQL = SQL & " and dt_vencimento >= '" & DMA(txtDtIni.Text, "i") & "'"
         SQL = SQL & " and dt_vencimento <= '" & DMA(txtDtFim.Text, "f") & "'"
      End If
      If optBaixa.Value = True Then
         SQL = SQL & " and dt_baixa >= '" & DMA(txtDtIni.Text, "i") & "'"
         SQL = SQL & " and dt_baixa <= '" & DMA(txtDtFim.Text, "f") & "'"
      End If
   End If

   If Trim(cmbCC.Text) <> "" Then _
      SQL = SQL & " and cc_id = " & Trim(cmbCCAux.Text)

   CONT_N = 0
'==========================
   Dim INDR_VACA As Boolean

   NUMR_SEQ_N = 1
   VALOR_TOTAL_N = 0
   VALOR_TOTAL_JUROS_N = 0
   VLR_JUROS_ATUALIZADO = 0
   VLR_TITULO_ATUALIZADO = 0
   MORA_JUROS = 0

   lstTitulos.ListItems.Clear
   lstTitulos.Visible = False

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      Me.Enabled = False

   While Not TabLancamento.EOF
      If CANCELA_LOOP = True Then
         Me.Enabled = True
         Exit Sub
      End If
      INDR_VACA = False
      If Trim(cmbVendAux.Text) <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

            SQL = "select vendedor_id from PEDIDO WITH (NOLOCK)"
            SQL = SQL & " where pedido_id = " & TabLancamento.Fields("numr_doc").Value
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            SQL = SQL & " and vendedor_id = " & cmbVendAux.Text
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If IsNull(TabTemp.Fields(0).Value) Then _
                  INDR_VACA = True
               Else: INDR_VACA = True
            End If

         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      If INDR_VACA = False Then
         DoEvents
         Set item = lstTitulos.ListItems.Add(, "seq." & NUMR_SEQ_N, Trim(TabLancamento.Fields("Numr_doc").Value))
         txtQTDE.Text = CONT_N

         item.SubItems(1) = "" & TabLancamento!SEQ

         'Pagamento
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
         SQL = SQL & " where formapagto_id = " & TabLancamento!FORMAPAGTO_ID
         SQL = SQL & " and status = 'true' "
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            item.SubItems(2) = "" & TabDESCR!DESCRICAO
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         VALOR_TOTAL_N = VALOR_TOTAL_N + TabLancamento!Valor_Item
         txtTotal.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      
         item.SubItems(3) = "" & Format(TabLancamento!Valor_Item, strFormatacao2Digitos)

         'Calculo Juros de atrazo -> Feito Por Emanoel
         If TabLancamento!DT_VENCIMENTO < Date Then
            DIAS_ATRAZO = Date - TabLancamento!DT_VENCIMENTO
            If TabLancamento!STATUS = "A" Then
               If tabEmpresa.State = 1 Then _
                  tabEmpresa.Close
   
               SQL = "select QTD_DIAS_ATRAZO,PERC_JUROS_ATRAZO from ESTABELECIMENTO WITH (NOLOCK)"
               SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not tabEmpresa.EOF Then
                  If tabEmpresa!QTD_DIAS_ATRAZO > 0 Then
                      MORA_JUROS = (((TabLancamento!Valor_Item * tabEmpresa!PERC_JUROS_ATRAZO) / 100) _
                                      / tabEmpresa!QTD_DIAS_ATRAZO)
                      MORA_JUROS = Format(MORA_JUROS, strFormatacao2Digitos)
                  End If
               End If
               If tabEmpresa.State = 1 Then _
                  tabEmpresa.Close
   
               VLR_JUROS_ATUALIZADO = MORA_JUROS * DIAS_ATRAZO
               VLR_TITULO_ATUALIZADO = VLR_JUROS_ATUALIZADO + TabLancamento!Valor_Item
            End If
         End If

         'No Caso de Restaurante nao tem isso
         DIAS_ATRAZO = 0
         MORA_JUROS = 0
         VLR_JUROS_ATUALIZADO = 0
         VLR_TITULO_ATUALIZADO = 0
         If TabLancamento!DT_VENCIMENTO >= Date Then
            DIAS_ATRAZO = 0
            MORA_JUROS = 0
            VLR_JUROS_ATUALIZADO = 0
            VLR_TITULO_ATUALIZADO = 0
         End If
          
         VALOR_TOTAL_JUROS_N = (VALOR_TOTAL_JUROS_N + (VLR_JUROS_ATUALIZADO + TabLancamento!Valor_Item))
         txtTotaljuros.Text = Format(VALOR_TOTAL_JUROS_N, strFormatacao2Digitos)
     
         item.SubItems(4) = "" & Format(VLR_JUROS_ATUALIZADO, strFormatacao2Digitos)
         item.SubItems(5) = "" & DIAS_ATRAZO 'Dias Atrazo

         If VLR_TITULO_ATUALIZADO = 0 Then _
            VLR_TITULO_ATUALIZADO = TabLancamento!Valor_Item

         item.SubItems(6) = "" & Format(VLR_TITULO_ATUALIZADO, strFormatacao2Digitos)  'Valor Atualizado
         item.SubItems(7) = "" & TabLancamento!DT_VENCIMENTO   'Vencimento

         If Not IsNull(TabLancamento!DT_BAIXA) Then
            If IsDate(TabLancamento!DT_BAIXA) Then
               If Year(TabLancamento!DT_BAIXA) > 2000 Then
                  item.SubItems(8) = "" & TabLancamento!DT_BAIXA
               End If
            End If
         End If

         SQL3 = "" & Trim(TabLancamento.Fields("status").Value)
         If (SQL3) <> "" Then
            If (SQL3) = "C" Then
               item.SubItems(9) = "Cancelado"
               item.SubItems(8) = ""
               Else
                  If (SQL3) = "A" Then
                     item.SubItems(9) = "Aberto"
                     Else
                        If (SQL3) = "B" Then
                           If INDR_RECEITA = 1 Then _
                              item.SubItems(9) = "Recebido"
                           If INDR_RECEITA = 2 Then _
                              item.SubItems(9) = "Pago"
                        End If
                        If TabLancamento.Fields("dt_baixa").Value >= TabLancamento.Fields("dt_cad").Value Then
                           If INDR_RECEITA = 1 Then _
                              item.SubItems(9) = "Recebido"
                           If INDR_RECEITA = 2 Then _
                              item.SubItems(9) = "Pago"
                        End If
                  End If
            End If
         End If
'=============
         If Trim(TabLancamento.Fields("NOME_PESSOA").Value) <> "" Then
            item.SubItems(10) = Trim(TabLancamento.Fields("NOME_PESSOA").Value)
            Else
               item.SubItems(10) = "" & Trim(TabLancamento.Fields("descricao").Value)

               If TabTemp.State = 1 Then _
                  TabTemp.Close
                  SQL = "select NOME_CLIENTE from PEDIDO WITH (NOLOCK)"
                  SQL = SQL & " where pedido_id = " & TabLancamento.Fields("numr_doc").Value
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then _
                     If Not IsNull(TabTemp.Fields(0).Value) Then _
                        item.SubItems(10) = "" & Trim(TabTemp.Fields(0).Value)
               If TabTemp.State = 1 Then _
                  TabTemp.Close
         End If
         item.SubItems(12) = "" & Trim(TabLancamento.Fields("cnpjcpf").Value)

         NUMR_SEQ_N = NUMR_SEQ_N + 1
         CONT_N = CONT_N + 1

         item.SubItems(11) = "" & TabLancamento!Numr_doc
         item.SubItems(13) = "" & TRAZ_ESTABELECIMENTO(TabLancamento.Fields("estabelecimento_id").Value)
      End If
      TabLancamento.MoveNext
   Wend
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   MOSTRA_RODAPE "ESC - SAIR", "Duplo Click ocultar itens", "", "", ""

   Me.Enabled = True
'1===================================================
   txtQTDE.Text = CONT_N
   txtTotal.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTotaljuros.Text = Format(VALOR_TOTAL_JUROS_N, strFormatacao2Digitos)
   lstTitulos.Visible = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_TUDO"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstab.Text = "" & TRAZ_ESTABELECIMENTO(cmbEstabAUX.Text)
   lstTitulos.ListItems.Clear
   txtLanc.Text = ""
   cmbModalidade.Text = ""
   cmbAuxModalidade.Text = ""
   cmbSituacao.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""
   optEmissao.Value = False
   optVencimento.Value = False
   optBaixa.Value = False
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.Text = ""
   txtTotal.Text = ""
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub preencheCombovend(NomeCombo As ComboBox)
'On Error GoTo ERRO_TRATA

    Dim rstVEND As ADODB.Recordset

    If rstVEND.State = 1 Then _
      rstVEND.Close

    SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK) "
    SQL = SQL & " where status = 'A' "
    SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
    SQL = SQL & " order by vendedor_id"
    rstVEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
        
    NomeCombo.Clear
    If Not rstVEND.EOF Then
       'Mundando o ponteiro do mouse, para mostrar para o usuario que esta processando...
        Screen.MousePointer = vbHourglass
        
        rstVEND.MoveFirst
        Do Until rstVEND.EOF
            'Importantissimo
            DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"
            
            NomeCombo.AddItem rstVEND!VENDEDOR_ID & "-" & rstVEND!DESCRICAO
            rstVEND.MoveNext
        Loop
    End If
    
    'Voltando o ponteiro do mouse para o tipo default, ponteiro.
    Screen.MousePointer = vbDefault
    If rstVEND.State = 1 Then _
      rstVEND.Close
    
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "preencheCombovend"
End Sub

Private Sub MostraCliente()
'On Error GoTo ERRO_TRATA

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      txtCli.Text = TabCliente!NOME
      CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value
   End If

   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MostraCliente"
End Sub

Sub MONTA_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = ""
   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   FORMULA_REL = "{vwFATURAMENTO.Tipo_Lancamento} = " & INDR_RECEITA

   If Trim(cmbEstab.Text) <> "" Then _
      If Trim(UCase(cmbEstab.Text)) <> UCase("todos") Then _
         FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.estabelecimento_id} = " & cmbEstabAUX.Text

   'Numero Pedido
   If IsNumeric(txtLanc.Text) Then _
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.Numr_doc} = " & txtLanc.Text

   If cmbAuxModalidade.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.formapagto_id} = " & cmbAuxModalidade

   If cmbSituacao.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.status} = '" & Left(cmbSituacao.Text, 1) & "'"

   txtCNPJCPF.PromptInclude = False
   If IsNumeric(txtCNPJCPF.Text) Then _
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.pessoa_id} = " & PESSOA_ID_N
   txtCNPJCPF.PromptInclude = True

   'Datas
   If optEmissao.Value = True Then _
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.dt_cad} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

   If optVencimento.Value = True Then
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.dt_vencimento} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.dt_vencimento} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
   End If

   If optBaixa.Value = True Then _
      FORMULA_REL = FORMULA_REL & " and {vwFATURAMENTO.dt_baixa} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

   If INDR_RECEITA = 1 Then
      If optSintetico.Value = True Then
         Nome_Relatorio = "FINCLIENTESINT.rpt"
         Else: Nome_Relatorio = "FINCLIENTE.rpt"
      End If
      Else
         If optSintetico.Value = True Then
            Nome_Relatorio = "FINforSINT.rpt"
            Else: Nome_Relatorio = "FINfor.rpt"
         End If
   End If

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL"
End Sub

Sub BAIXA_TITULOS_SELECIONADOS()
'On Error GoTo ERRO_TRATA

   Dim i                   As Integer
   Dim INDR_BAIXA_TITULO   As Boolean

   INDR_PRI = True
   INDR_BAIXA_TITULO = False

   If lstTitulos.ListItems.Count > 0 Then
      For i = lstTitulos.ListItems.Count To 1 Step -1
         If lstTitulos.ListItems(i).Checked = True Then
            If Trim(UCase(lstTitulos.ListItems(i).SubItems(9))) = "ABERTO" Then

               If INDR_PRI = True Then
                  INDR_PRI = False
                  Msg = "Deseja Baixar Título(s) Selecionado(s) ? "
                  Style = vbYesNo + 32
                  Title = "Atenção."
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                     INDR_BAIXA_TITULO = True
                     Else
                        INDR_BAIXA_TITULO = False
                        Exit Sub
                  End If
               End If
               If INDR_BAIXA_TITULO = True Then
                  CONT_N = 1

                  SQL = "UPDATE ITEMLANCAMENTO SET "
                     SQL = SQL & " Status = 'B'"
                     SQL = SQL & ", DT_baixa = '" & Now & "'"
                     SQL = SQL & ", usu_alt = " & USUARIO_ID_N
                     SQL = SQL & ", dt_alt = '" & Now & "'"
                  SQL = SQL & " where numr_doc = " & Trim(lstTitulos.ListItems(i).Text)
                  SQL = SQL & " and seq = " & lstTitulos.ListItems(i).SubItems(1)
                  SQL = SQL & " and status <> 'C' "
                  CONECTA_RETAGUARDA.Execute SQL
               End If
               txtQTDE.Text = Trim(lstTitulos.ListItems(i).Text)
               DoEvents
            End If
         End If
      Next i
      If INDR_PRI = False Then _
         MsgBox "Baixa realizada com sucesso !!!"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_TITULOS_SELECIONADOS"
End Sub

Private Sub MOSTRA_VENDEDORES()
'On Error GoTo ERRO_TRATA

   cmbVend.Clear
   cmbVendAux.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close
   SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVENDEDOR.EOF
      cmbVend.AddItem Trim(TabVENDEDOR!DESCRICAO) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
      TabVENDEDOR.MoveNext
   Wend
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDORES"
End Sub
