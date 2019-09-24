VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNFECONSULTA 
   Caption         =   "Consulta Nota Fiscal"
   ClientHeight    =   8115
   ClientLeft      =   2295
   ClientTop       =   2580
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NFECONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   -90
      TabIndex        =   21
      Top             =   690
      Width           =   11145
      Begin VB.ComboBox cmbAuxVend 
         BackColor       =   &H80000008&
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
         Left            =   6120
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbVend 
         Height          =   360
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   5205
      End
      Begin VB.ComboBox cmbSTATUS 
         Height          =   360
         ItemData        =   "NFECONSULTA.frx":5C12
         Left            =   1200
         List            =   "NFECONSULTA.frx":5C1C
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtReq 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3315
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCli 
         Height          =   360
         Left            =   6840
         MaxLength       =   100
         TabIndex        =   27
         Top             =   720
         Width           =   4095
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   4680
         TabIndex        =   4
         Top             =   720
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   240
         Left            =   4665
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   240
         Left            =   3840
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   240
         Left            =   195
         TabIndex        =   30
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         Height          =   240
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NFe:"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   495
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
      TabIndex        =   20
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
      Left            =   -180
      TabIndex        =   17
      Top             =   1800
      Width           =   11355
      Begin VB.ComboBox cmbCFOP 
         Height          =   360
         Left            =   1320
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
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   240
         Left            =   600
         TabIndex        =   19
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      Height          =   735
      Left            =   -180
      TabIndex        =   11
      Top             =   2430
      Width           =   11625
      Begin VB.OptionButton optBaixa 
         Caption         =   "Data &Cancelamento"
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   300
         Width           =   2175
      End
      Begin VB.OptionButton optEmis 
         Caption         =   "Data &Emissão"
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   300
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   360
         Left            =   6960
         TabIndex        =   8
         Top             =   300
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
         Left            =   9600
         TabIndex        =   9
         Top             =   300
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
         Height          =   240
         Left            =   8520
         TabIndex        =   13
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial:"
         Height          =   240
         Left            =   5760
         TabIndex        =   12
         Top             =   300
         Width           =   1140
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
            Picture         =   "NFECONSULTA.frx":5C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":608A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":63A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":67FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":6C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":6F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":73C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":76E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":9E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":A2EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":ACFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":B70E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NFECONSULTA.frx":C120
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
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
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   8280
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
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
               Picture         =   "NFECONSULTA.frx":CB32
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECONSULTA.frx":DCCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECONSULTA.frx":ED5B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECONSULTA.frx":FFF7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFECONSULTA.frx":11102
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTAASS 
      Height          =   3585
      Left            =   60
      TabIndex        =   10
      Top             =   3360
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   6324
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
         Text            =   "Tipo NFe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Vendedor"
         Object.Width           =   0
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
         Object.Width           =   5292
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
      DesignWidth     =   10950
      DesignHeight    =   8115
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   11040
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   11040
      Y1              =   3240
      Y2              =   3240
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   7680
      Width           =   2610
   End
End
Attribute VB_Name = "frmNFECONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmNFECONSULTA)
   Me.Caption = Me.Caption & " - " & Me.Name

   cmbAuxCFOP.Clear
   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP WITH (NOLOCK)"
   SQL = SQL & " order by CFOP_ID"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbAuxCFOP.AddItem TabDESCR.Fields("cfop_id").Value
      cmbCFOP.AddItem TabDESCR.Fields("cfop_id").Value & "-" & Trim(TabDESCR!DESCRICAO)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
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
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbSTATUS.Clear
   cmbSTATUS.AddItem "Emitidas"
   cmbSTATUS.AddItem "Canceladas"
   
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

   PEDIDO_ID_N = LISTAASS.SelectedItem.Text
   CRITERIO_A = LISTAASS.SelectedItem.ListSubItems.item(LISTAASS.ColumnHeaders(1).Position)
   SQL3 = LISTAASS.SelectedItem.ListSubItems.item(LISTAASS.ColumnHeaders(2).Position)
   Err.Clear
End Sub

Private Sub listaass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTAASS, ColumnHeader
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

Private Sub optBaixa_Click()
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
            CRITERIO_A = LISTAASS.SelectedItem.Text
            If IsNumeric(CRITERIO_A) Then
               PEDIDO_ID_N = CRITERIO_A

               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from NF "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp!STATUS) Then
                     If TabTemp!STATUS = "C" Then
                        TabTemp.Close
                        MsgBox "Operação não permitida, nota fiscal cancelada."
                        Exit Sub
                     End If
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from PEDIDO "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  If Not IsNull(TabTemp!STATUS) Then
                     If TabTemp!STATUS = 9 Then
                        TabTemp.Close
                        MsgBox "Operação não permitida, Pedido cancelada."
                        Exit Sub
                     End If
                  End If
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

TIPO_NFe_GERAR = "R"
If USA_DOC_FISCAL = True Then _
   If USA_NFe = True Then _
      frmNOTAGERA.Show 1
            End If
         End If
         'CONSULTA_TUDO
      Case "consultar"
         FORMULA_REL = ""
         CONSULTA_TUDO
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         FORMULA_REL = ""
         Unload Me
      Case "print"
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_nf_saida.rpt"
         frmRELATORIO10.Show 1
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

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Clientes", "", "", ""

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.Text = CNPJCPF_A
      CNPJCPF_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCNPJCPF.Text = "" Then _
         txtCNPJCPF.Text = "99999999999"
      SQL = "select * from CLIENTE "
      SQL = SQL & " where CGCCPF = '" & txtCNPJCPF.Text & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
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
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
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

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
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

Private Sub txtNota_KeyPress(KeyAscii As Integer)
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

   FORMULA_REL = "{vwRel_Nf_Saida.estabelecimento_id} = " & ESTABELECIMENTO_ID_N

   If txtReq.Text <> "" Then
      SQL = SQL & " and pedido_id = " & txtReq.Text

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.pedido_id} = " & Trim(txtReq.Text)
   End If

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" Then
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.CNPJCPF_aF} = '" & Trim(txtCNPJCPF.Text) & "'"
   End If
   txtCNPJCPF.PromptInclude = True

   If cmbAuxVend.Text <> "" Then
      SQL = SQL & " and vendedor_id = " & cmbAuxVend.Text

      FORMULA_REL = FORMULA_REL & " and {PEDIDO.vendedor_id} = " & cmbVend.Text
   End If

   'cfop
   If cmbAuxCFOP.Text <> "" Then
      SQL = SQL & " and cfop = '" & Trim(cmbAuxCFOP.Text) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.cfop} = '" & Trim(cmbAuxCFOP.Text) & "'"
   End If

   'statux
   If Trim(cmbSTATUS.Text) <> "" Then
      SQL = SQL & " and STATUS = '" & Left(cmbSTATUS.Text, 1) & "'"

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.STATUS_NF} = '" & Left(cmbSTATUS.Text, 1) & "'"
   End If

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   If ((optEmis.Value = False) And (optBaixa.Value = False)) Then _
      optEmis.Value = True

   If ((IsDate(txtDtIni.Text)) And (IsDate(txtDtFim.Text))) Then
      If optEmis.Value = True Then
         SQL = SQL & " and dt_emissao >= '" & DMA(txtDtIni.Text) & "'"
         SQL = SQL & " and dt_emissao <= '" & DMA(txtDtFim.Text) & "'"

FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.dt_emissao} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.dt_emissao} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
      End If
   End If

   If txtNOTA.Text <> "" Then
      SQL = SQL & " and numr_nota = " & txtNOTA.Text

      FORMULA_REL = FORMULA_REL & " and {vwRel_Nf_Saida.numr_nota} = " & txtNOTA.Text
   End If

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
      Set item = LISTAASS.ListItems.Add(, "seq." & TabTemp.Fields("pedido_id").Value & CONT_N, TabTemp.Fields("pedido_id").Value)
      item.SubItems(1) = "" & Trim(TabTemp!NUMR_NOTA)
      item.SubItems(2) = "" & Trim(TabTemp!SERIE_NOTA)

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select nome from CLIENTE "
      SQL = SQL & " where cgccpf = '" & TabTemp.Fields("prop") & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         item.SubItems(3) = Trim(TabCliente!NOME)
         Else
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select descricao from vwFornecedor "
            SQL = SQL & " where cnpjcpf = '" & TabTemp.Fields("prop") & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then _
               item.SubItems(3) = Trim(TabCliente!NOME)
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      'aqui pega o total da venda
      VALOR_ITEM_N = 0
      SQL = "select sum(valor_item*qtd_pedida) "
      SQL = SQL & " from PEDIDO "
      SQL = SQL & " INNER JOIN PEDIDOITEM "
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"
      SQL = SQL & " where PEDIDO.pedido_id = " & TabTemp.Fields("pedido_id").Value
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and empresa_id  = " & EMPRESA_ID_N
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not IsNull(TabCliente.Fields(0).Value) Then
         VALOR_TOTAL_N = VALOR_TOTAL_N + TabCliente.Fields(0).Value
         VALOR_ITEM_N = TabCliente.Fields(0).Value
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      PERC_DESCONTO_N = 0
      SQL = "select perc_desc from PEDIDO "
      SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then _
         If Not IsNull(TabCliente.Fields(0).Value) Then _
            PERC_DESCONTO_N = TabCliente.Fields(0).Value
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select valor_desconto from PEDIDO "
      SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then _
         If Not IsNull(TabCliente.Fields(0).Value) Then _
            VALOR_DESCONTO_CABECA_N = TabCliente.Fields(0).Value
      If TabCliente.State = 1 Then _
         TabCliente.Close

      'desconto da tabela PEDIDO                ***
      VALOR_DESCONTO_N = VALOR_DESCONTO_N + (VALOR_ITEM_N * PERC_DESCONTO_N / 100)
      
      item.SubItems(4) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      item.SubItems(5) = "" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos)

      CONT_N = CONT_N + 1
            
      item.SubItems(6) = "" & TabTemp!DT_EMISSAO
      item.SubItems(7) = "" & Trim(TabTemp.Fields("nf_tipo").Value)

      If Not IsNull(TabTemp.Fields("CFOP_id").Value) Then
         If Trim(TabTemp.Fields("CFOP_id").Value) <> "" Then
            If IsNumeric(TabTemp.Fields("CFOP_id").Value) Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close

               SQL = "select * from CFOP WITH (NOLOCK)"
               SQL = SQL & " where CFOP_ID = " & Trim(TabTemp.Fields("CFOP_id").Value)
               TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabUSU.EOF Then _
                  item.SubItems(10) = "" & Trim(TabUSU.Fields("codigo").Value) & "-" & Trim(TabUSU.Fields("descricao").Value)

               If TabUSU.State = 1 Then _
                  TabUSU.Close
            End If
         End If
      End If
      If TabUSU.State = 1 Then _
         TabUSU.Close

      item.SubItems(9) = ""

      If Not IsNull(TabTemp.Fields("STATUS")) Then
         If TabTemp.Fields("STATUS") = "E" Then _
            item.SubItems(9) = "Emitida"

         If TabTemp.Fields("STATUS") = "C" Then _
            item.SubItems(9) = "Cancelado"
      End If

      txtTotalVenda.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
      txtTot.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
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

   FORMULA_REL = ""
   cmbAuxCFOP.Text = ""
   cmbCFOP.Text = ""
   cmbSTATUS.Text = ""
   optEmis.Value = False
   optBaixa.Value = False
   LISTAASS.ListItems.Clear
   txtReq.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""
   cmbVend.Text = ""
   cmbAuxVend.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtNOTA.Text = ""
   txtTotalVenda.Text = ""
   txtDesconto.Text = ""
   txtTot.Text = ""
   txtCotas.Text = ""
   txtReq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub
