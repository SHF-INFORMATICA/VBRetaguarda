VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVENDAECF2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venda"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "VENDAECF2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   4200
      Picture         =   "VENDAECF2.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox txtDesconto 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   2280
      MaxLength       =   12
      TabIndex        =   28
      Top             =   5925
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.TextBox txtItens 
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
      ForeColor       =   &H00C000C0&
      Height          =   450
      Left            =   11040
      MaxLength       =   12
      TabIndex        =   20
      Top             =   5925
      Width           =   735
   End
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   8880
      MaxLength       =   12
      TabIndex        =   8
      Top             =   5925
      Width           =   1815
   End
   Begin VB.TextBox txtValorTroco 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   6720
      MaxLength       =   12
      TabIndex        =   7
      Top             =   5925
      Width           =   1815
   End
   Begin VB.TextBox txtValorRecebido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   450
      Left            =   4560
      MaxLength       =   12
      TabIndex        =   3
      Top             =   5925
      Width           =   1815
   End
   Begin VB.TextBox txtValorCompra 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      MaxLength       =   12
      TabIndex        =   6
      Top             =   5925
      Width           =   1815
   End
   Begin VB.TextBox txtValorItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtUN 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      MaxLength       =   4
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtQtde 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox txtPRODUTO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin MSComctlLib.ListView lstECF 
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   6165
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   17639
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtde."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pr.Venda"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total Item"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Código Barras"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "pedido_id"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ICMS"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ST"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "seq_id"
         Object.Width           =   176
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
      DesignWidth     =   11910
      DesignHeight    =   6900
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   9360
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483648
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar BARRAECF 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6525
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor Desc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   2400
      TabIndex        =   29
      Top             =   5595
      Width           =   1590
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   12
      X1              =   2040
      X2              =   2040
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Label lblST2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   7200
      TabIndex        =   26
      Top             =   1320
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      TabIndex        =   25
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblST 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6840
      TabIndex        =   24
      Top             =   1320
      Width           =   315
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   7
      X1              =   10920
      X2              =   10920
      Y1              =   5520
      Y2              =   6720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   6
      X1              =   8760
      X2              =   8760
      Y1              =   5520
      Y2              =   6720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   5
      X1              =   6600
      X2              =   6600
      Y1              =   5520
      Y2              =   6720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   4
      X1              =   4320
      X2              =   4320
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   4
      X1              =   0
      X2              =   12000
      Y1              =   5535
      Y2              =   5535
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Itens"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   11070
      TabIndex        =   21
      Top             =   5595
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   1125
      Left            =   120
      Top             =   720
      Width           =   11655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   11
      X1              =   10200
      X2              =   10200
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   10
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   9
      X1              =   5880
      X2              =   5880
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   8
      X1              =   4080
      X2              =   4080
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   8880
      TabIndex        =   19
      Top             =   5595
      Width           =   1515
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Valor Troco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   6720
      TabIndex        =   18
      Top             =   5595
      Width           =   1650
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor Pagto."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   4680
      TabIndex        =   17
      Top             =   5595
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor Venda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   5595
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vlr.Item ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Qtde ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4260
      TabIndex        =   14
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label lblUsu 
      Caption         =   "Usuário"
      Height          =   240
      Left            =   10275
      TabIndex        =   13
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label lblTabela 
      Caption         =   "Tabela Preço"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   8040
      TabIndex        =   12
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label lblData 
      Caption         =   "Data Emis."
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6000
      TabIndex        =   11
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label lblPedido 
      Caption         =   "Pedido"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   4200
      TabIndex        =   10
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label lblEstabelecimento 
      Caption         =   "Estabelecimento"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   3615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   3
      X1              =   11760
      X2              =   11760
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   2
      X1              =   135
      X2              =   135
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   3
      X1              =   120
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   2
      X1              =   120
      X2              =   11760
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   1
      X1              =   0
      X2              =   12000
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   1
      X1              =   11880
      X2              =   11880
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   0
      X1              =   -120
      X2              =   11880
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Index           =   0
      X1              =   30
      X2              =   30
      Y1              =   0
      Y2              =   6840
   End
End
Attribute VB_Name = "frmVENDAECF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim QTDE_DISPONIVEL        As Double
   Dim Qtde_lstECF_Atual      As Double
   Dim Indr_Achou_Produto     As Boolean
   Dim NUMR_CUPOM_ABERTO      As Long
   Dim NUMEROCUPOMCancelado   As String
   Dim NUMEROCUPOM            As String
   Dim NOME_VENDEDOR          As String
   Dim NOME_CLIENTE           As String
   Dim NUMR_PEDIDO_ID_N       As Long
   Dim QTDE_N                 As Double
   Dim Mensagem_Final         As String
   Dim Tipo_Venda             As String
   Dim Aliquota_Icms_N        As Long
   Dim CONTA_TENTATIVA        As Long
   Dim f                      As Variant
   Dim LocalRetorno           As String
   Dim CNPJCPF_CLIENTE        As String
   Dim Parametros             As Variant
   Dim OperacaoECFOK          As Boolean
   Dim ITEM_DESCONTO_N        As Double
   Dim Descr_Forma_Pagto      As String
   Const FORMATO_MONEY        As String = "#0.00"
   Const CUPOM_FISCAL         As String = "Cupom Fiscal"
   Const FORMA_PGTO_CARTAO    As String = "Cartao"
   Const FORMA_PGTO_CHEQUE    As String = "Cheque"
   Dim INDR_PROD_BALANCA      As Boolean
   Dim CODIGO_BARRAS          As String
   Dim PESO_ITEM_N            As Double

Private Sub Form_Load()
   INICIALIZA_TELA
End Sub

Private Sub Form_Activate()
'On Error GoTo ERRO_TRATA

   'verifica se o programa já está sendo executado
   If (App.PrevInstance) Then
       Dim nome_tela As String
       nome_tela = App.Title
       App.Title = "Já estou em execução, frmVENDAECF2 !!!"
       AppActivate nome_tela
       SendKeys "%R", True
       MsgBox "Já em execução !!!"
       Unload frmVENDAECF2
       'End
   End If

   If CHECA_ABERTURA_DIA = False Then
      MsgBox "Solicitar Contra Senha ao suporte da SHF INFORMÁTICA."
      End
   End If

Exit Sub
ERRO_TRATA:
MsgBox Err.Description
   TRATA_ERROS Err.Description, Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF10
         FAZ_RECEBIMENTO
         INICIALIZA_TELA
         txtProduto.SetFocus
      Case vbKeyF2
         CNPJCPF_CLIENTE = "99999999999"
         CNPJCPF_CLIENTE = Trim(InputBox("Informe CPF/CNPJ do cliente", "Emissão de Cupom Fiscal", CNPJCPF_CLIENTE))

         NOME_CLIENTE = "Consumidor Final"
         NOME_CLIENTE = Trim(InputBox("Informe Nome do cliente", "Emissão de Cupom Fiscal", NOME_CLIENTE))
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   If INDR_VENDA = True Then
      Msg = "Cupom fiscal aberto, deseja cancelar essa venda? "
      Msg = "Deseja cancelar essa venda? "
      PERGUNTA Msg, vbYesNo + 32, "Pedido Venda", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         CANCELA_CUPOM_ABERTO
         Else: Cancel = 1
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub lstECF_GotFocus()
   MOSTRA_MSG "F3 - Cancela Venda", "F6 - Cancela Item", "", "F11 - Abre Gaveta", ""
End Sub

Private Sub lstecf_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstECF, ColumnHeader
End Sub

Private Sub Timer1_Timer()
   Me.Caption = "Venda de Produtos  |  " & NOME_ESTABELEC & "  |  " & Time & "  |  Pedido: " & NUMR_REQ_N & "  |  Cupom: " & NUMR_CUPOM_ABERTO
End Sub

Private Sub lstECF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6 Or vbKeyDelete
         If Not IsNull(lstECF.SelectedItem.Text) Then
            If Trim(lstECF.SelectedItem.Text) <> "" Then
               If IsNumeric(lstECF.SelectedItem.Text) Then
                  If NUMR_REQ_N > 0 Then

                     SQL3 = "" & lstECF.SelectedItem.ListSubItems.Item(lstECF.ColumnHeaders(9).Position)

                     RETORNO_ECF = Bematech_FI_CancelaItemGenerico(SQL3)

                     If TABCUPOM.State = 1 Then _
                        TABCUPOM.Close

                     SQL = "select * FROM PEDIDOITEM "
                     SQL = SQL & " where pedido_id = " & NUMR_REQ_N
                     SQL = SQL & " and seq_id = " & SQL3
                     SQL = SQL & " and status <> 'C' "
                     SQL = SQL & " and tipo_reg = 'PC' "
   'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

                     TABCUPOM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TABCUPOM.EOF Then
                        QTDE_PEDIDO = lstECF.SelectedItem.ListSubItems.Item(lstECF.ColumnHeaders(3).Position)

                        SQL = "UPDATE PEDIDOITEM set "
                        SQL = SQL & " status = 'C' "
                        SQL = SQL & " where pedido_id = " & NUMR_REQ_N
                        SQL = SQL & " and seq_id = " & SQL3
                        CONECTA_RETAGUARDA.Execute SQL

                        SQL = "UPDATE PRODUTO SET "
                        SQL = SQL & " Qtde = Qtde + " & Replace(QTDE_PEDIDO, ",", ".")
                        SQL = SQL & " where codg_produto = '" & Trim(lstECF.SelectedItem.Text) & "'"
                        SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                        CONECTA_RETAGUARDA.Execute SQL
                     End If

                     If TABCUPOM.State = 1 Then _
                        TABCUPOM.Close
                  End If
               End If
            End If
         End If
         SETA_GRID
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstECF_KeyDown"
End Sub

Private Sub cmdConsProd_Click()
   CRITERIO = ""
   SQL3 = ""
   frmProdutoConsulta.Show 1
   txtProduto.Text = SQL3
   txtProduto.SetFocus
End Sub

Private Sub txtDescricao_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescricao_GotFocus"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_MSG "Produto/Barras", "F2-Incluir Cliente", "F3-Cancela Venda", "F7-Consulta Produtos", "F10-Fechar Venda  F11-Abre Gaveta"
   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_GotFocus"
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF3
         If INDR_CUPOM_ABERTO = True Then
            Call frmECFOperacao.CANCELA_CUPOM_FISCAL
            INICIALIZA_TELA
         End If
      Case vbKeyF7
         CRITERIO = ""
         SQL3 = ""
         frmProdutoConsulta.Show 1
         txtProduto.Text = SQL3
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = Asc(",") Then _
      KeyAscii = Asc(".")

   If KeyAscii = 13 Then
      If Trim(txtProduto.Text) = "" And Trim(txtValorTotal.Text) <> "" Then
         txtValorRecebido.SetFocus
         Exit Sub
      End If

      LE_PRODUTO

      txtProduto.Enabled = True
      txtQTDE.Enabled = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
   txtQTDE.Enabled = False
   If Trim(txtQTDE.Text) = "" Then
      txtProduto.SetFocus
      Exit Sub
   End If
   If Not IsNumeric(txtQTDE.Text) Then
      txtProduto.SetFocus
      Exit Sub
   End If

   QTDE_PEDIDO = txtQTDE.Text

   If QTDE_PEDIDO <= 0 Then
      txtProduto.SetFocus
      Exit Sub
   End If
End Sub

Private Sub txtITENS_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtITENS_GotFocus"
End Sub

'=======================================================================================================================
Private Sub txtValorCompra_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorCompra_GotFocus"
End Sub

Private Sub txtDesconto_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDESCONTO_GotFocus"
End Sub

Private Sub txtValorRecebido_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_MSG "Informe Valor Recebido", "", "", "", ""

   txtValorRecebido.SelStart = 0
   txtValorRecebido.SelLength = Len(txtValorRecebido)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorRecebido_GotFocus"
End Sub

Private Sub txtValorRecebido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = Asc(".") Then _
      KeyAscii = Asc(",")

   If KeyAscii = 13 Then
      KeyAscii = 0
      If INDR_CUPOM_ABERTO = True Then
         If Trim(txtValorRecebido.Text) <> "" Then

            PERC_DESCONTO_USUARIO = 0
            VALOR_TOTAL_DESCONTO_N = 0
            PERC_DESCONTO_N = 0
            USU_LIBERA_VENDA_N = 0
            VALOR_TOTAL_N = 0 & txtValorCompra.Text
            VALOR_TOTAL_DESCONTO_N = 0

            If INDR_LIBERA_DESCONTO = True Then
               Msg = "Deseja informar desconto ?"
               PERGUNTA Msg, vbYesNo + 32, "Desconto NFE", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  frmVENDADESCONTO.Show 1
                  If DESCONTO_AUTORIZADO = False Then _
                     Exit Sub
               End If
            End If

            txtValorRecebido.Text = Format(txtValorRecebido.Text, strFormatacao2Digitos)
            VALOR_TOTAL_N = 0
            VALOR_RECEBIDO_N = 0
            VALOR_RECEBIDO_N = txtValorRecebido.Text

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select sum(VALOR_ITEM*QTD_PEDIDA) FROM PEDIDOITEM "
            SQL = SQL & " where pedido_id = " & NUMR_REQ_N
            SQL = SQL & " and status <> 'C' "
            SQL = SQL & " and tipo_reg = 'PC' "
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               If Not IsNull(TabTemp.Fields(0).Value) Then _
                  VALOR_TOTAL_N = TabTemp.Fields(0).Value
            If TabTemp.State = 1 Then _
               TabTemp.Close

            VALOR_TOTAL_N = VALOR_TOTAL_N - VALOR_TOTAL_DESCONTO_N

            txtDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
            'txtValorRecebido.Text = VALOR_RECEBIDO_N-VALOR_TOTAL_N
            'VALOR_RECEBIDO_N = VALOR_TOTAL_N
            txtValorTroco.Text = ""

            VALOR_TOTAL_N = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
            VALOR_RECEBIDO_N = Format(VALOR_RECEBIDO_N, strFormatacao2Digitos)

            If VALOR_TOTAL_N > VALOR_RECEBIDO_N Then
               MsgBox "Valor recebido menor que valor da venda, não permitido."
               Exit Sub
            End If

            txtValorTroco.Text = Format(VALOR_RECEBIDO_N - VALOR_TOTAL_N, strFormatacao2Digitos)
            KeyAscii = 0

            PERC_DESCONTO_USUARIO = 0
            txtCNPJCPF.PromptInclude = False
            CNPJCPF_A = txtCNPJCPF.Text

            'atualizando desconto na cabeça
            SQL = "UPDATE PEDIDO SET "
            SQL = SQL & " Valor_desconto = " & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
            SQL = SQL & " , Perc_desc = " & tpMOEDA(PERC_DESCONTO_N)
            SQL = SQL & " , cgccpf = '" & Trim(CNPJCPF_CLIENTE) & "'"
            SQL = SQL & " , nome_cliente = '" & Trim(NOME_CLIENTE) & "'"
            SQL = SQL & " , status = 2"
            SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
            SQL = SQL & " , valor_recebido = " & tpMOEDA(VALOR_RECEBIDO_N)
            SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

            CONECTA_RETAGUARDA.Execute SQL

            FAZ_RECEBIMENTO
            INICIALIZA_TELA
            txtProduto.SetFocus
         End If
      End If
      Else
         If KeyAscii = 8 Then
            Else
               If KeyAscii = 44 Then
                  Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorRecebido_KeyPress"
End Sub

Private Sub txtValorTotal_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorTotal_GotFocus"
End Sub
'=======================================================================================================================
Private Sub txtValorTroco_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorTroco_GotFocus"
End Sub
'=======================================================================================================================
Private Sub txtUN_GotFocus()
'On Error GoTo ERRO_TRATA

   SendKeys ("{tab}")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtun_GotFocus"
End Sub
'=======================================================================================================================
Private Sub txtValorItem_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_MSG "Informe Valor Item", "F3 - Cancela Venda", "", "Disponível estoque = " & QTDE_DISPONIVEL, ""

   If Trim(txtValorItem.Text) = "" Then _
      txtValorItem.Text = 0

   txtValorItem.SelStart = 0
   txtValorItem.SelLength = Len(txtValorItem)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_GotFocus"
End Sub

Private Sub txtValorItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'If KeyAscii = Asc(",") Then _
      KeyAscii = Asc(".")

   If KeyAscii = 13 Then
      VALOR_ITEM_N = 0 & txtValorItem.Text

      If VALOR_ITEM_N <= 0 Then
         MsgBox "Valor informado não permitido."
         Exit Sub
      End If

      KeyAscii = 0
      txtQTDE.Enabled = True
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorItem_KeyPress"
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   txtQTDE.Enabled = True

   MOSTRA_MSG "Informe Quantidade", "F3 - Cancela Venda", "", "Disponível estoque = " & QTDE_DISPONIVEL, ""

   If Trim(txtQTDE.Text) = "" Then _
      txtQTDE.Text = 0

   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'If KeyAscii = Asc(",") Then _
      KeyAscii = Asc(".")

   If KeyAscii = 13 Then

      KeyAscii = 0

      txtQTDE.Enabled = False

      If Trim(txtQTDE.Text) = "" Then
         txtProduto.SetFocus
         Exit Sub
      End If
      If Not IsNumeric(txtQTDE.Text) Then
         txtProduto.SetFocus
         Exit Sub
      End If

      QTDE_PEDIDO = txtQTDE.Text

      If QTDE_PEDIDO < 0 Then
         txtProduto.SetFocus
         Exit Sub
      End If
      If QTDE_PEDIDO = 0 Then
         QTDE_PEDIDO = 1
         txtQTDE.Text = 1
      End If
      txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)

      If QTDE_PEDIDO <= 0 Then
         txtValorRecebido.SetFocus
         QTDE_PEDIDO = 0
         txtProduto.SetFocus
         Exit Sub
      End If

      ABRE_VENDA
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtqtde_KeyPress"
End Sub
'===========================SUBROTINAS
Sub INICIALIZA_TELA()
'On Error GoTo ERRO_TRATA

   LIMPA_VENDA
   QUALIFICA_VENDEDOR

   txtItens.Text = ""
   NUMR_REQ_N = 0
   INDR_CUPOM_ABERTO = False

   lblEstabelecimento.Caption = ""
   lblPedido.Caption = ""
   lblData.Caption = Format(Date, "dd/mm/yyyy")
   lblTabela.Caption = ""
   lblUsu.Caption = ""
   txtValorTroco.Text = ""

   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   SQL = SQL & " SELECT PESSOA.CNPJCPF, PESSOA.DESCRICAO"
   SQL = SQL & " FROM EMPRESA "
   SQL = SQL & " INNER JOIN PESSOA "
   SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmpresa.EOF Then
      TabEmpresa.MoveFirst
      CNPJCPF_A = Trim(TabEmpresa.Fields("cnpjcpf").Value)

      lblEstabelecimento.Caption = Trim(TabEmpresa.Fields("descricao").Value)
      lblData.Caption = "Dt.Emis: " & Format(Date, "dd/mm/yyyy")

      If TabUSU.State = 1 Then _
         TabUSU.Close
      SQL = "select logon from USUARIO "
      SQL = SQL & " where usuario_id = " & CODG_USU_N
      TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabUSU.EOF Then _
         lblUsu.Caption = Trim(TabUSU.Fields(0).Value)
      If TabUSU.State = 1 Then _
         TabUSU.Close
   End If
   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   'Verifica se a Impressa esta ligada ou nao
   RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
   If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
      RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
      MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!", vbCritical, "MEGASIM"
      Exit Sub
   End If

   'Verifica Cupom em Aberto
   INDR_CUPOM_ABERTO = False
   Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
   If INDR_CUPOM_ABERTO = True Then _
      CANCELA_CUPOM_ABERTO

Me.WindowState = 2

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_TELA"
End Sub

Sub PROCURA_PRODUTO(Codigo_Produto_A As String)
'On Error GoTo ERRO_TRATA

   Indr_Achou_Produto = False

   If Trim(Codigo_Produto_A) <> "" Then
      If TabProduto.State = 1 Then _
          TabProduto.Close

      SQL = "SELECT PRODUTO_ID,EMPRESA_ID,FAMILIAPRODUTO_ID,FORNECEDOR_ID,CODG_PRODUTO,"
      SQL = SQL & " DESCRICAO,UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,QTDE,"
      SQL = SQL & " SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,CODG_NCM,"
      SQL = SQL & " PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda "

      SQL = SQL & " from Produto"

      SQL = SQL & " where codg_barra = '" & Trim(Codigo_Produto_A) & "'"
      SQL = SQL & " and situacao = 'A' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         If TabTemp.State = 1 Then _
             TabTemp.Close

         SQL = "select count(codg_barra) from PRODUTO "
         SQL = SQL & " where codg_barra = '" & Trim(Codigo_Produto_A) & "'"
         SQL = SQL & " and situacao = 'A' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp.Fields(0).Value) Then
               If TabTemp.Fields(0).Value > 1 Then
                  CRITERIO = Codigo_Produto_A
                  frmPEDIDOBARRAS.Show 1
                  Codigo_Produto_A = CRITERIO

                  If TabProduto.State = 1 Then _
                      TabProduto.Close

                  SQL = "SELECT PRODUTO_ID,EMPRESA_ID,FAMILIAPRODUTO_ID,FORNECEDOR_ID,CODG_PRODUTO,"
                  SQL = SQL & " DESCRICAO,UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,QTDE,"
                  SQL = SQL & " SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,CODG_NCM,"
                  SQL = SQL & " PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda "

                  SQL = SQL & " from Produto"

                  SQL = SQL & " where codg_produto = '" & Trim(Codigo_Produto_A) & "'"
                  SQL = SQL & " and situacao = 'A' "
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabProduto.EOF Then
                     If TabProduto.State = 1 Then _
                         TabProduto.Close

                     SQL = "SELECT PRODUTO_ID,EMPRESA_ID,FAMILIAPRODUTO_ID,FORNECEDOR_ID,CODG_PRODUTO,"
                     SQL = SQL & " DESCRICAO,UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,"
                     SQL = SQL & " SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,CODG_NCM,"
                     SQL = SQL & " PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda "

                     SQL = SQL & " from Produto"

                     SQL = SQL & " where codg_barra = '" & Trim(Codigo_Produto_A) & "'"
                     SQL = SQL & " and situacao = 'A' "
                     TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabProduto.EOF Then _
                        MOSTRA_PRODUTO
                     Else: MOSTRA_PRODUTO
                  End If

                  VALOR_ITEM_N = 0
                  CODG_PRODUTO_A = ""
            
                  If txtValorItem.Text <> "" Then _
                     VALOR_ITEM_N = txtValorItem.Text
                  If Codigo_Produto_A <> "" Then _
                     CODG_PRODUTO_A = Codigo_Produto_A
            
                  If CODG_PRODUTO_A = "" Then
                     MsgBox "Produto  informado inválido."
                     txtProduto.SetFocus
                     Exit Sub
                  End If
                  If VALOR_ITEM_N <= 0 Then
                     MsgBox "Produto sem preço."
                     txtProduto.SetFocus
                     Exit Sub
                  End If

                  Exit Sub

                  Else
                     Codigo_Produto_A = "" & TabProduto!CODG_PRODUTO
                     txtProduto.Text = "" & TabProduto!CODG_PRODUTO
                     PRODUTO_ID_N = 0 & TabProduto!PRODUTO_ID
               End If
               Else
                  txtProduto.Text = "" & TabProduto!CODG_PRODUTO
                  Codigo_Produto_A = "" & TabProduto!CODG_PRODUTO
                  PRODUTO_ID_N = 0 & TabProduto!PRODUTO_ID
            End If
            Else
               txtProduto.Text = "" & TabProduto!CODG_PRODUTO
               Codigo_Produto_A = "" & TabProduto!CODG_PRODUTO
               PRODUTO_ID_N = 0 & TabProduto!PRODUTO_ID
         End If
         If TabTemp.State = 1 Then _
             TabTemp.Close

         If TabProduto.State = 1 Then _
             TabProduto.Close

         SQL = "SELECT PRODUTO_ID,EMPRESA_ID,FAMILIAPRODUTO_ID,FORNECEDOR_ID,CODG_PRODUTO,"
         SQL = SQL & " DESCRICAO,UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,QTDE,"
         SQL = SQL & " SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,CODG_NCM,"
         SQL = SQL & " PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda "

         SQL = SQL & " from Produto"

         SQL = SQL & " where codg_produto = '" & Trim(Codigo_Produto_A) & "'"
         SQL = SQL & " and situacao = 'A' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then _
            MOSTRA_PRODUTO
         Else
            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "SELECT PRODUTO_ID,EMPRESA_ID,FAMILIAPRODUTO_ID,FORNECEDOR_ID,CODG_PRODUTO,"
            SQL = SQL & " DESCRICAO,UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,QTDE,"
            SQL = SQL & " SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,CODG_NCM,"
            SQL = SQL & " PRECO_CUSTO,PRECO_ATACADO,PRECO_Venda "

            SQL = SQL & " from Produto"

            SQL = SQL & " where codg_produto = '" & Trim(Codigo_Produto_A) & "'"
            SQL = SQL & " and situacao = 'A' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then
               txtProduto.Text = TabProduto!CODG_PRODUTO
               Codigo_Produto_A = TabProduto!CODG_PRODUTO
               PRODUTO_ID_N = 0 & TabProduto!PRODUTO_ID
               MOSTRA_PRODUTO
               Else
                  MsgBox "Produto não cadastrado."
                  txtProduto.Text = ""
                  PRODUTO_ID_N = 0
                  Codigo_Produto_A = ""
                  Indr_Achou_Produto = False
                  txtProduto.SetFocus
                  Exit Sub
            End If
      End If
      If TabProduto.State = 1 Then _
          TabProduto.Close

      VALOR_ITEM_N = 0
      CODG_PRODUTO_A = ""

      If txtValorItem.Text <> "" Then _
         VALOR_ITEM_N = txtValorItem.Text

      If Trim(txtProduto.Text) <> "" Then _
         CODG_PRODUTO_A = txtProduto.Text

      If Trim(CODG_PRODUTO_A) = "" Then
         MsgBox "Produto  informado inválido."
         txtProduto.SetFocus
         Exit Sub
      End If

      If VALOR_ITEM_N <= 0 Then
         MsgBox "Produto sem preço."
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
End Sub

Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

   txtDescricao.Text = ""
   txtUN.Text = ""
   txtValorItem.Text = ""

   Indr_Achou_Produto = True

   txtProduto.Text = TabProduto!CODG_PRODUTO
   PRODUTO_ID_N = 0 & TabProduto!PRODUTO_ID
   txtDescricao.Text = TabProduto!Descricao
   txtUN.Text = "" & TabProduto!Unidade_Medida
   lblST.Caption = ""
   
   lblST2.Caption = ""

   If Not IsNull(TabProduto.Fields("aliquota_icms").Value) Then _
      Aliquota_Icms_N = 0 & TabProduto.Fields("aliquota_icms").Value

   If Not IsNull(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value) Then
      If Trim(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value) <> "" Then
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         SQL = "select * from CST "
         SQL = SQL & " where codigo = " & Trim(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value)
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            lblST.Caption = "" & TabDESCR!codigo
            lblST2.Caption = "" & Trim(TabDESCR!Descricao)
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      End If
   End If

   If Trim(TabProduto!SITUACAO) <> "A" Then
      MsgBox "Produto não liberado para venda."
      txtProduto.Text = ""
      Indr_Achou_Produto = False
      txtProduto.SetFocus
      Exit Sub
   End If

   If IsNull(TabProduto!PRECO_VENDA) Then
      MsgBox "Produto Sem Preço de Venda."
      txtProduto.Text = ""
      Indr_Achou_Produto = False
      txtProduto.SetFocus
      Exit Sub
   End If

   VALOR_ITEM_N = 0 & TabProduto.Fields("preco_venda").Value
   If VALOR_ITEM_N <= 0 Then
      MsgBox "Produto Sem Preço de Venda."
      txtProduto.Text = ""
      Indr_Achou_Produto = False
      txtProduto.SetFocus
      Exit Sub
   End If

   txtValorItem.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)

   Qtde_lstECF_Atual = 0

   Set ITEM2 = Nothing
   Set ITEM2 = lstECF.ListItems
   Set ITEM2 = lstECF.FindItem(Trim(txtProduto.Text), , , 1)
   If Not ITEM2 Is Nothing Then _
      Qtde_lstECF_Atual = ITEM2.SubItems(2)
   Set ITEM2 = Nothing

   QTDE_DISPONIVEL = 0
   QTDE_PEDIDO = 0

   If Not IsNull(TabProduto!Qtde) Then
      QTDE_DISPONIVEL = TabProduto!Qtde - Qtde_lstECF_Atual
      QTDE_PEDIDO = TabProduto!Qtde - Qtde_lstECF_Atual
   End If

   Indr_Achou_Produto = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub LIMPA_VENDA()
'On Error GoTo ERRO_TRATA

   txtValorTroco.Text = ""
   txtDesconto.Text = ""
   txtValorCompra.Text = ""
   txtValorRecebido.Text = ""
   txtValorTotal.Text = ""
   txtItens.Text = ""
   INDR_VENDA = False
   lstECF.ListItems.Clear
   CNPJCPF_CLIENTE = ""
   NOME_CLIENTE = ""

   LIMPA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_VENDA"
End Sub

Sub LIMPA_GRID()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0
   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtQTDE.Text = ""
   txtUN.Text = ""
   txtValorItem.Text = ""
   lblST.Caption = ""
   lblST2.Caption = ""
   SEQ_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_GRID"
End Sub

Sub QUALIFICA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select logon from USUARIO "
   SQL = SQL & " where usuario_id = " & CODG_USU_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      CRITERIO = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
      SQL = "select nome_vend, vendedor_id from VENDEDOR "
      SQL = SQL & " where nome_vend like " & CRITERIO
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         NOME_VENDEDOR = Trim(TabVENDEDOR!NOME_VEND)
         VENDEDOR_ID_N = TabVENDEDOR!VENDEDOR_ID
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "QUALIFICA_VENDEDOR"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstECF.ListItems.Clear
   NUMR_SEQ_N = 0
   VALOR_ITEM_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.VENDEDOR_ID, "
   SQL = SQL & " PEDIDO.TIPOVENDA_ID, PEDIDO.NUMR_CUPOM, PEDIDO.NUMR_REQ, PEDIDO.CGCCPF,"
   SQL = SQL & " PEDIDO.CODG_USU, PEDIDO.DT_REQ, PEDIDO.STATUS as SituaçãoPedido, PEDIDO.TIPO_REGISTRO,"
   SQL = SQL & " PEDIDO.VALOR_DESCONTO DescontoCabeça, PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_RECEBIDO,"
   SQL = SQL & " PEDIDO.VALOR_TOTAL, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.CODG_PROD, PEDIDOITEM.QTD_PEDIDA,"
   SQL = SQL & " PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.CFOP, PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.VALOR_DESCONTO AS DescontoItem,"
   SQL = SQL & " PEDIDOITEM.STATUS AS SituaçãoItem, PEDIDOITEM.PRECO_CUSTO, PRODUTO.FAMILIAPRODUTO_ID,"
   SQL = SQL & " PRODUTO.FORNECEDOR_ID, PRODUTO.DESCRICAO, PRODUTO.UNIDADE_MEDIDA, PRODUTO.CODG_BARRA,"
   SQL = SQL & " PRODUTO.SITUACAO as SituaçãoProduto, PRODUTO.QTDE, PRODUTO.produto_balanca,"
   SQL = SQL & " PRODUTO.SITUACAO_TRIBUTARIA, PRODUTO.ALIQUOTA_ICMS, PRODUTO.TIPO_PROD, PRODUTO.CODG_NCM,QTDE_BALANCA,"
   SQL = SQL & " PRODUTO.PRECO_CUSTO AS CustoProduto, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, VENDEDOR.NOME_VEND"
   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN VENDEDOR "
   SQL = SQL & " ON PEDIDO.VENDEDOR_ID = VENDEDOR.VENDEDOR_ID"

   SQL = SQL & " where PEDIDO.PEDIDO_ID = " & NUMR_REQ_N
   SQL = SQL & " and PEDIDOITEM.PEDIDO_ID = " & NUMR_REQ_N
   SQL = SQL & " and PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "

   SQL = SQL & " where pedido.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " order by PEDIDOITEM.SEQ_ID desc"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      INDR_VENDA = True
      'NUMR_SEQ_N = NUMR_SEQ_N + 1

      Set Item = lstECF.ListItems.Add(, "seq." & TabTemp.Fields("SEQ_ID").Value, TabTemp.Fields("codg_prod").Value)
      Item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
      Item.SubItems(2) = "" & Format(TabTemp.Fields("QTD_PEDIDA").Value, strFormatacao3Digitos)
      Item.SubItems(3) = "" & Format(TabTemp.Fields("VALOR_ITEM").Value, strFormatacao2Digitos)
      Item.SubItems(4) = "" & Format(TabTemp.Fields("VALOR_ITEM").Value * TabTemp.Fields("QTD_PEDIDA").Value, strFormatacao2Digitos)

If Not IsNull(TabTemp.Fields("produto_balanca").Value) Then
   If TabTemp.Fields("produto_balanca").Value = True Then
      Item.SubItems(2) = "" & Format(TabTemp.Fields("QTDE_BALANCA").Value, strFormatacao3Digitos)
      Item.SubItems(3) = "" & Format(TabTemp.Fields("VALOR_ITEM").Value, strFormatacao2Digitos)
      'Item.SubItems(4) = "" & Format(TabTemp.Fields("VALOR_ITEM").Value * TabTemp.Fields("QTDE_BALANCA").Value, strFormatacao2Digitos)
   End If
End If

'txtValorRecebido.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
'txtValorCompra.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
'txtValorTotal.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)

      Item.SubItems(5) = "" & TabTemp.Fields("codg_barra").Value
      Item.SubItems(6) = "" & TabTemp.Fields("pedido_id").Value

      '1-Tributado;2-Isento;3-Outros;4-Base Calculo Reduzido;5-Diferido
      If Not IsNull(TabTemp.Fields("Situacao_Tributaria").Value) Then
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         SQL = "select descricao from CST "
         SQL = SQL & " where codigo = '" & Trim(TabTemp.Fields("SITUACAO_TRIBUTARIA").Value) & "'"
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            Item.SubItems(7) = "" & Trim(TabTemp.Fields("SITUACAO_TRIBUTARIA").Value) & "-" & Trim(TabDESCR.Fields(0).Value)
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      End If

      Item.SubItems(9) = "" & TabTemp.Fields("SEQ_ID").Value

      If Trim(TabTemp.Fields("SituaçãoItem").Value) = "C" Then
         Item.ForeColor = vbRed
         Item.ListSubItems(1).ForeColor = vbRed
         Item.ListSubItems(2).ForeColor = vbRed
         Item.ListSubItems(3).ForeColor = vbRed
         Item.ListSubItems(4).ForeColor = vbRed
         Item.ListSubItems(5).ForeColor = vbRed
         Item.ListSubItems(6).ForeColor = vbRed
         Item.ListSubItems(7).ForeColor = vbRed
         Item.ListSubItems(8).ForeColor = vbRed
         Else
            Item.ForeColor = vbBlue
            Item.ListSubItems(1).ForeColor = vbBlue
            Item.ListSubItems(2).ForeColor = vbBlue
            Item.ListSubItems(3).ForeColor = vbBlue
            Item.ListSubItems(4).ForeColor = vbBlue
            Item.ListSubItems(5).ForeColor = vbBlue
            Item.ListSubItems(6).ForeColor = vbBlue
            Item.ListSubItems(7).ForeColor = vbBlue
            Item.ListSubItems(8).ForeColor = vbBlue
      End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   TOTALIZA_JANELAS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub ABRE_VENDA()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Produto não informado."
      Exit Sub
   End If

   VALOR_ITEM_N = 0 & txtValorItem.Text
   If VALOR_ITEM_N <= 0 Then
      MsgBox "Valor de Venda Incorreto."
      Exit Sub
   End If

   If QTDE_N <= 0 Then
      MsgBox "Quantidade informada inválida."
      Exit Sub
   End If

   CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
   Select Case CRITERIO
      Case "BEMATECH"
         'Verifica se a Impressa esta ligada ou nao
         RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
         If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
            RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
            MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!", vbCritical, "MEGASIM"
            'Exit Sub
         End If

         If INDR_CUPOM_ABERTO = False Then _
            GRAVA_CABEÇA_PEDIDO

         If Trim(txtProduto.Text) <> "" Then _
            GRAVA_ITEM_PEDIDO
      Case "DARUMA"
      Case "Sweda"
   End Select

   SETA_GRID

   lblPedido.Caption = "Pedido : " & NUMR_REQ_N
   lblTabela.Caption = "Estabelecimento: " & EMPRESA_ID_N

   MOSTRA_MSG "Venda de Mercadoria ", NOME_ESTABELEC, "", "", ""

   LIMPA_GRID

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ABRE_VENDA"
End Sub

Sub GRAVA_CABEÇA_PEDIDO()
'On Error GoTo ERRO_TRATA

   MOSTRA_MSG "Inicializando Cupom Fiscal", "", "", "", ""

   Indr_Erro = False
   INDR_VENDA = True

   NUMR_CUPOM_ABERTO = 0

   CONTA_TENTATIVA = 0
   If Trim(CNPJCPF_CLIENTE) = "" Then _
      CNPJCPF_CLIENTE = "99999999999"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PESSOA_ID,nome,cliente_id,cgccpf from CLIENTE "
   SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLIENTE) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      PESSOA_ID_N = TabTemp.Fields(0).Value
      NOME_CLIENTE = Trim(TabTemp.Fields(1).Value)
      CLIENTE_ID_N = TabTemp.Fields(2).Value
      CNPJCPF_CLIENTE = Trim(TabTemp.Fields(3).Value)
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   Msg = ""

   MOSTRA_MSG "Abrindo Gaveta", NOME_ESTABELEC, "", "", ""

      GERA_NUMR_REQ

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from PEDIDO "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      INDR_PRI = True
      Indr_Erro = False

ABRINDO_CUPOM_FISCAL:

      DoEvents

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select PESSOA_ID,nome,cliente_id,cgccpf from CLIENTE "
      SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_CLIENTE) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         PESSOA_ID_N = TabTemp.Fields(0).Value
         NOME_CLIENTE = Trim(TabTemp.Fields(1).Value)
         CLIENTE_ID_N = TabTemp.Fields(2).Value
         CNPJCPF_CLIENTE = "" & Trim(TabTemp.Fields(3).Value)
         Else
            If Trim(CNPJCPF_CLIENTE) = "" Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close
      
               SQL = "select PESSOA_ID,nome,cliente_id,cgccpf from CLIENTE "
               SQL = SQL & " where cgccpf = '99999999999'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  PESSOA_ID_N = TabTemp.Fields(0).Value
                  NOME_CLIENTE = Trim(TabTemp.Fields(1).Value)
                  CLIENTE_ID_N = TabTemp.Fields(2).Value
                  CNPJCPF_CLIENTE = "" & Trim(TabTemp.Fields(3).Value)
               End If
            End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If CLIENTE_ID_N <= 0 Then _
         CLIENTE_ID_N = 1

      If TIPOVENDA_ID_N <= 0 Then _
         TIPOVENDA_ID_N = 9999

      SQL = "insert into PEDIDO "
         SQL = SQL & " ("
         SQL = SQL & " PEDIDO_ID,CLIENTE_ID,EMPRESA_ID,VENDEDOR_ID,"
         SQL = SQL & " TIPOVENDA_ID,NUMR_CUPOM, NUMR_REQ,CGCCPF,"
         SQL = SQL & " CODG_USU,DT_REQ,STATUS,TIPO_REGISTRO,NOME_CLIENTE,NUMERO_CAIXA_CPU,establecimento_id"
         SQL = SQL & " )"
      SQL = SQL & " values("
         SQL = SQL & NUMR_REQ_N                       'PEDIDO_ID
         SQL = SQL & "," & CLIENTE_ID_N               'CLIENTE_ID
         SQL = SQL & "," & EMPRESA_ID_N               'EMPRESA_ID
         SQL = SQL & "," & VENDEDOR_ID_N              'VENDEDOR_ID
         SQL = SQL & "," & TIPOVENDA_ID_N             'TIPOVENDA_ID
         SQL = SQL & ",0" & NUMEROCUPOM                'NUMR_CUPOM
         SQL = SQL & "," & NUMR_REQ_N                 'NUMR_REQ
         SQL = SQL & "," & Trim(CNPJCPF_CLIENTE)      'CGCCPF
         SQL = SQL & "," & CODG_USU_N                 'CODG_USU
         SQL = SQL & ",'" & DMA(Date) & "'"           'DT_REQ
         SQL = SQL & "," & 2                          'STATUS
         SQL = SQL & "," & "'R'"                      'TIPO_REGISTRO
         SQL = SQL & ",'" & Trim(NOME_CLIENTE) & "'"  'NOME_CLIENTE "
         SQL = SQL & "," & NUMERO_CAIXA_CPU           'NUMERO_CAIXA_CPU
         SQL = SQL & "," & ESTABELECIMENTO_ID_N           'estabelecimento_id
      SQL = SQL & ")"

      CONECTA_RETAGUARDA.Execute SQL

      INDR_VENDA = True
      INDR_CUPOM_ABERTO = True
   End If
   Indr_Erro = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABEÇA_PEDIDO"
End Sub

Sub GRAVA_ITEM_PEDIDO()
'On Error GoTo ERRO_TRATA

If Trim(txtProduto.Text) <> "" Then
   SQL = "select * FROM PEDIDO"
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      SQL = "insert into PEDIDO "
         SQL = SQL & " ("
         SQL = SQL & " PEDIDO_ID,CLIENTE_ID,EMPRESA_ID,VENDEDOR_ID,"
         SQL = SQL & " TIPOVENDA_ID,NUMR_CUPOM, NUMR_REQ,CGCCPF,"
         SQL = SQL & " CODG_USU,DT_REQ,STATUS,TIPO_REGISTRO,NOME_CLIENTE,NUMERO_CAIXA_CPU,estabelecimento_id"
         SQL = SQL & " )"
      SQL = SQL & " values("
         SQL = SQL & NUMR_REQ_N                       'PEDIDO_ID
         SQL = SQL & "," & CLIENTE_ID_N               'CLIENTE_ID
         SQL = SQL & "," & EMPRESA_ID_N               'EMPRESA_ID
         SQL = SQL & "," & VENDEDOR_ID_N              'VENDEDOR_ID
         SQL = SQL & "," & TIPOVENDA_ID_N             'TIPOVENDA_ID
         SQL = SQL & ",0" & NUMEROCUPOM                'NUMR_CUPOM
         SQL = SQL & "," & NUMR_REQ_N                 'NUMR_REQ
         SQL = SQL & "," & Trim(CNPJCPF_CLIENTE)      'CGCCPF
         SQL = SQL & "," & CODG_USU_N                 'CODG_USU
         SQL = SQL & ",'" & DMA(Date) & "'"           'DT_REQ
         SQL = SQL & "," & 2                          'STATUS
         SQL = SQL & "," & "'R'"                      'TIPO_REGISTRO
         SQL = SQL & ",'" & Trim(NOME_CLIENTE) & "'"  'NOME_CLIENTE "
         SQL = SQL & "," & NUMERO_CAIXA_CPU           'NUMERO_CAIXA_CPU
         SQL = SQL & "," & ESTABELECIMENTO_ID_N           'estabelecimento_id
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   SEQ_ID_N = 1

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select max(seq_id) FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and tipo_reg = 'PC' "
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      If Not IsNull(TabUSU.Fields(0).Value) Then _
         SEQ_ID_N = TabUSU.Fields(0).Value + 1
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If TabUSU.State = 1 Then _
      TabUSU.Close
VALOR_ITEM_N = 0
QTDE_BALANCA = 0
   SQL = "select preco_venda,produto_balanca FROM PRODUTO "
   SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      If Not IsNull(TabUSU.Fields("preco_venda").Value) Then _
         VALOR_ITEM_N = TabUSU.Fields("preco_venda").Value

      INDR_PROD_BALANCA = False
      If TabUSU.Fields("produto_balanca").Value Then _
         INDR_PROD_BALANCA = TabUSU.Fields("produto_balanca").Value

      If INDR_PROD_BALANCA = True Then
         VALOR_ITEM_N = Format(TabUSU.Fields("preco_venda").Value * QTDE_PEDIDO, strFormatacao2Digitos)
         QTDE_BALANCA = QTDE_PEDIDO
         QTDE_PEDIDO = 1
      End If
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "insert into PEDIDOITEM "
      SQL = SQL & " ("
         SQL = SQL & " PEDIDO_ID,SEQ_ID,PRODUTO_ID,NUMR_REQ,CODG_PROD,QTD_PEDIDA,"
         SQL = SQL & " VALOR_ITEM,STRIBUTARIA,STATUS,TIPO_REG,QTDE_BALANCA  "
      SQL = SQL & " )"
   SQL = SQL & " values("
      SQL = SQL & NUMR_REQ_N                          'PEDIDO_ID
      SQL = SQL & "," & SEQ_ID_N                      'seq_ID
      SQL = SQL & "," & PRODUTO_ID_N                  'PRODUTO_ID
      SQL = SQL & "," & NUMR_REQ_N                    'NUMR_REQ
      SQL = SQL & ",'" & Trim(txtProduto.Text) & "'"  'Codg_Prod
      SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)          'QTD_PEDIDA
      SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)         'Valor_Item
      SQL = SQL & ",'" & Trim(lblST.Caption) & "'"    'STRIBUTARIA
      SQL = SQL & ",'P'"                              'STATUS
      SQL = SQL & ",'PC'"                             'TIPO_REG
      SQL = SQL & "," & tpMOEDA(QTDE_BALANCA)         'QTDE_BALANCA
   SQL = SQL & ")"
   CONECTA_RETAGUARDA.Execute SQL

   If INDR_PROD_BALANCA = True Then _
      QTDE_PEDIDO = QTDE_BALANCA

'=============baixa estoque INICIO
   SQL = "UPDATE PRODUTO SET "
   SQL = SQL & " Qtde = Qtde - " & Replace(QTDE_PEDIDO, ",", ".")
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "UPDATE PEDIDOITEM set "
   SQL = SQL & " status = 'B' "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and status <> 'B' "
   SQL = SQL & " and produto_id = " & PRODUTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL
'=============baixa estoque FIM

   INDR_CUPOM_ABERTO = True
End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITEM_PEDIDO"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Sub GERA_FINANCEIRO()
'On Error GoTo ERRO_TRATA

   Dim Valor_Tot_n As Double
   Dim VALOR_DESCONTO_CABECA_N As Double
   Dim TOTAL_DESCONTO_N As Double
   Dim VALOR_ENTRADA As Double

   NUMR_PARCELA = 1

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   VALOR_DESCONTO_N = 0
   SQL = "select perc_desc from PEDIDO "
   SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
   SQL = SQL & " and pedido_id = " & NUMR_REQ_N

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
      PERC_DESCONTO_N = TabPedidoItem.Fields(0).Value
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   VALOR_DESCONTO_CABECA_N = 0
   SQL = "select valor_desconto from PEDIDO "
   SQL = SQL & " where empresa_id  = " & EMPRESA_ID_N
   SQL = SQL & " and pedido_id = " & NUMR_REQ_N

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
      VALOR_DESCONTO_CABECA_N = TabPedidoItem.Fields(0).Value
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and status <> 'C' "
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
      VALOR_DESCONTO_N = TabPedidoItem.Fields(0).Value
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   'BUSCA VALOR TOTAL VENDA
   Valor_Tot_n = 0
   SQL = "select sum(valor_item*qtd_pedida) FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and status <> 'C' "
   SQL = SQL & " and tipo_reg = 'PC' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not IsNull(TabPedidoItem.Fields(0).Value) Then _
      Valor_Tot_n = TabPedidoItem.Fields(0).Value
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   'VALOR_DESCONTO_N = VALOR_DESCONTO_N + (Valor_Tot_n * IIf(PERC_DESCONTO_N > 0, PERC_DESCONTO_N / 100, 1))
   VALOR_DESCONTO_N = VALOR_DESCONTO_N + VALOR_DESCONTO_CABECA_N

   VALOR_ITEM_N = 0
   DATA_INI = Date
   If NUMR_PARCELA > 0 Then _
      VALOR_ITEM_N = (Valor_Tot_n - VALOR_DESCONTO_N) / NUMR_PARCELA

   'CABEÇA
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SqL2 = Date
   NUMR_ID_N = 0

   SQL = "select * from LANCAMENTO "

   SQL = SQL & " where numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = 1"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      NUMR_ID_N = TabLancamento!Lancamento_id

      SQL = "UPDATE LANCAMENTO SET "
      SQL = SQL & " Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Prop = '" & CNPJCPF_A & "'"
      SQL = SQL & ", dt_lanc = '" & DMA(SqL2) & "'"
      SQL = SQL & ", Valor_Lanc = " & Str(Format(Valor_Tot_n, strFormatacao2Digitos))
      SQL = SQL & ", Total_Desconto = " & Str(Format(VALOR_DESCONTO_N, strFormatacao2Digitos))
      SQL = SQL & ", Tipo_Lancamento = 1 "
      SQL = SQL & ", Empresa_Id = " & EMPRESA_ID_N
      SQL = SQL & ", Tipo_pagto = 1"

      SQL = SQL & " WHERE Empresa_Id = " & EMPRESA_ID_N
      SQL = SQL & " and Numr_Doc = " & NUMR_REQ_N
      SQL = SQL & " and Tipo_Lancamento = 1"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      Else
         NUMR_ID_N = MAX_ID("lancamento_id", "lancamento", "", "", "", "")

         SQL = "INSERT INTO LANCAMENTO "
         SQL = SQL & " ("
            SQL = SQL & " Lancamento_id, Numr_doc, Prop, dt_lanc, Valor_Lanc, Total_Desconto, "
            SQL = SQL & " Tipo_Lancamento, Empresa_id, Tipo_pagto,pessoa_id,estabelecimento_id"
         SQL = SQL & " ) "
         SQL = SQL & " VALUES ("
            SQL = SQL & NUMR_ID_N
            SQL = SQL & "," & NUMR_REQ_N
            SQL = SQL & ",'" & Trim(CNPJCPF_A) & "'"
            SQL = SQL & ",'" & DMA(SqL2) & "'"
            SQL = SQL & "," & Str(Format(Valor_Tot_n, strFormatacao2Digitos))
            SQL = SQL & "," & Str(Format(TOTAL_DESCONTO_N, strFormatacao2Digitos))
            SQL = SQL & ", 1"
            SQL = SQL & "," & EMPRESA_ID_N
            SQL = SQL & "," & 1
            SQL = SQL & "," & PESSOA_ID_N
            SQL = SQL & "," & ESTABELECIMENTO_ID_N
         SQL = SQL & ")"
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   CONECTA_RETAGUARDA.Execute SQL

   Dim Situacao_A As String

   Situacao_A = "B"

   NUMR_SEQ_N = 1

   SQL = "select max(seq) as ultimo_reg from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where i.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and i.numr_doc = l.numr_doc "
   SQL = SQL & " and i.lancamento_id = l.lancamento_id "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento!ultimo_reg) Then _
         NUMR_SEQ_N = NUMR_SEQ_N + TabLancamento!ultimo_reg
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   DATA_INI = Date

   SQL = "select * from ITEMLANCAMENTO "
   SQL = SQL & " where seq = " & NUMR_SEQ_N
   SQL = SQL & " and lancamento_id = " & NUMR_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      SQL = "UPDATE ITEMLANCAMENTO SET "
      SQL = SQL & " usu_alt = " & CODG_USU_N
      SQL = SQL & ", Dt_Alt = '" & DMA(SqL2) & "'"
      SQL = SQL & ", Numr_doc = " & NUMR_REQ_N
      SQL = SQL & ", Seq = " & NUMR_SEQ_N
      SQL = SQL & ", Valor_Item = " & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA / NUMR_PARCELA))
      SQL = SQL & ", Status = 'B'"
      SQL = SQL & ", formapagto_id = 1 "
      SQL = SQL & ", DT_VENCIMENTO = '" & DMA(SqL2) & "'"
      SQL = SQL & ", DT_baixa = '" & DMA(SqL2) & "'"
      SQL = SQL & " Where Lancamento_id = " & NUMR_ID_N
      SQL = SQL & " and Seq = " & NUMR_SEQ_N
      Else
         SQL = "INSERT INTO ITEMLANCAMENTO "
            SQL = SQL & " (Usu_Alt, Dt_Alt, Lancamento_id, Numr_doc, NUMR_DP, seq, Valor_Item, Status, formapagto_id, DT_VENCIMENTO, ACERTO) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CODG_USU_N                                                                                'Usu_Alt
            SQL = SQL & ",'" & DMA(SqL2) & "'"                                                                    'Dt_Alt
            SQL = SQL & "," & NUMR_ID_N                                                                           'Lancamento_id
            SQL = SQL & "," & NUMR_REQ_N                                                                          'Numr_doc
            SQL = SQL & "," & NUMR_REQ_N                                                                          'NUMR_DP
            SQL = SQL & "," & NUMR_SEQ_N                                                                          'seq
            SQL = SQL & "," & Str(Format(VALOR_ITEM_N, strFormatacao2Digitos) - (VALOR_ENTRADA / NUMR_PARCELA))   'Valor_Item
            SQL = SQL & ",'B'"                                                                                    'Status
            SQL = SQL & "," & 1                                                                                'formapagto_id
            SQL = SQL & ",'" & DMA(SqL2) & "'"                                                                    'DT_VENCIMENTO
            SQL = SQL & "," & 0                                                                                   'ACERTO
         SQL = SQL & ")"
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

      CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FINANCEIRO"
End Sub

Sub GRAVANDO_ERRO_EMISSAO_CUPOM()
'On Error GoTo ERRO_TRATA

   If Indr_Erro = True Then
      Dim RETORNOSTATUS As String

      NUMR_SEQ_N = 0

LE_ULTIMO_ECF2:

      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOMCancelado = Space(1)
         Else: NUMEROCUPOMCancelado = Space(6)
      End If

      RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOMCancelado)
      'Função que analisa o retorno da impressora
      'Call VerificaRetornoImpressora("Número do Último Cupom: ", _
           NumeroCupomCancelado, "Informações da Impressora")
      MOSTRA_MSG "ERRO", Msg & " Ultimo Cupom Impresso", "", "", ""

      If NUMEROCUPOMCancelado = "" Then
         If Not IsNumeric(NUMEROCUPOMCancelado) Then
            MsgBox "Erro na leitura do ultimo cupom impresso.  \" & NUMEROCUPOM
            NUMR_SEQ_N = NUMR_SEQ_N + 1
            'If NUMR_SEQ_N < 3 Then _
               GoTo LE_ULTIMO_ECF2
         End If
      End If

      If IsNumeric(NUMEROCUPOMCancelado) Then
         If NUMEROCUPOMCancelado = NUMR_CUPOM_ABERTO Then
            RETORNO_ECF = Bematech_FI_CancelaCupom()
            'Função que analisa o retorno da impressora
            Call VerificaRetornoImpressora("Bematech_FI_CancelaCupom", "", "Emissão de Cupom Fiscal")
            MOSTRA_MSG "ERRO", Msg & " Cancelando Cupom Fiscal", "", "", ""

            'GRAVA_CUPOM NUMEROCUPOMCancelado

            NUMR_ID_N = 0

            If TABCUPOM.State = 1 Then _
               TABCUPOM.Close

            SQL = "select pedido_id from CUPOM "
            SQL = SQL & " where numr_cupom = " & NUMEROCUPOMCancelado
            TABCUPOM.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TABCUPOM.EOF Then
               If Not IsNull(TABCUPOM.Fields(0).Value) Then
                  SQL = "update PEDIDO set "
                  SQL = SQL & " dt_cancela = '" & DMA(Date) & "'"
                  SQL = SQL & " , status = 'C'"
                  SQL = SQL & " where pedido_id = " & TABCUPOM.Fields(0).Value
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
               Else
                  MsgBox "Cupom não encontrado. " & NUMEROCUPOMCancelado
                  Exit Sub
            End If
            If TABCUPOM.State = 1 Then _
               TABCUPOM.Close

            'TRATA_ERROS Msg & "  " & NumeroCupomCancelado, Me.Name, "IMPRIME_CUPOM_FISCAL"

            Else: MsgBox "Erro, cupom fiscal diferente do impresso, não cancelado."
         End If
      End If
   End If

   MOSTRA_MSG "OK", Msg & " / " & "Fim Impressão, ECF cancelado  " & NUMEROCUPOMCancelado, "", "", ""
   INDR_VENDA = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVANDO_ERRO_EMISSAO_CUPOM"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Sub TOTALIZA_JANELAS()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(valor_ITEM*qtd_pedida) FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and status <> 'C' "
   SQL = SQL & " and tipo_reg = 'PC' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      VALOR_TOTAL_N = Format(TabTemp.Fields(0).Value, strFormatacao4Digitos)
      txtValorRecebido.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtValorCompra.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
      txtValorTotal.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select count(produto_id) FROM PEDIDOITEM "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and status <> 'C' "
   SQL = SQL & " and tipo_reg = 'PC' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtItens.Text = TabTemp.Fields(0).Value
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOTALIZA_JANELAS"
End Sub

Sub GRAVA_CUPOM(Numero_Pedido As String, Numero_Cupom As String)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   If IsNumeric(Numero_Pedido) And IsNumeric(Numero_Cupom) Then
      SQL = "select pedido_id from PEDIDO "
      SQL = SQL & " where pedido_id = " & Numero_Pedido
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then _
         Exit Sub

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'GRAVA TABELA CUPOM
      SQL = "select * from CUPOM"
      SQL = SQL & " where numr_cupom = " & Numero_Pedido
      SQL = SQL & " and Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         SqL2 = "update CUPOM set "
            SqL2 = SqL2 & " VALOR_CUPOM = " & tpMOEDA(VALOR_TOTAL_N)               'VALOR_CUPOM
            SqL2 = SqL2 & ", IMPRESSORA_ID = " & IMPRESSORA_ID_N                   'IMPRESSORA_ID
            SqL2 = SqL2 & ", Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO   'Numr_Contador_Reinicio
         SqL2 = SqL2 & " where numr_cupom = " & Numero_Cupom
         SqL2 = SqL2 & " and Numr_Contador_Reinicio = " & NUMR_CONTADOR_REINICIO   'Numr_Contador_Reinicio
         Else
            SqL2 = "insert into CUPOM "
            SqL2 = SqL2 & " (CUPOM_ID,NUMR_CUPOM,PEDIDO_ID,VALOR_CUPOM,IMPRESSORA_ID,Numr_Contador_Reinicio)"
            SqL2 = SqL2 & " VALUES("
               SqL2 = SqL2 & MAX_ID("cupom_id", "cupom", "", "", "", "")  'CUPOM_ID
               SqL2 = SqL2 & "," & Numero_Cupom                           'NUMR_CUPOM
               SqL2 = SqL2 & "," & Numero_Pedido                          'PEDIDO_ID
               SqL2 = SqL2 & "," & tpMOEDA(VALOR_TOTAL_N)                 'VALOR_CUPOM
               SqL2 = SqL2 & "," & IMPRESSORA_ID_N                        'IMPRESSORA_ID
               SqL2 = SqL2 & "," & NUMR_CONTADOR_REINICIO                 'Numr_Contador_Reinicio
            SqL2 = SqL2 & ")"
      End If

      CONECTA_RETAGUARDA.Execute SqL2
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
End Sub

Sub MOSTRA_MSG(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   If Trim(Msg1) <> "" Then
      BARRAECF.Panels.Clear
      BARRAECF.Panels.Add (1)
      BARRAECF.Panels(1).Text = Trim(Msg1)
      BARRAECF.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         BARRAECF.Panels.Add (2)
         BARRAECF.Panels(2).Text = Trim(Msg2)
         BARRAECF.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            BARRAECF.Panels.Add (3)
            BARRAECF.Panels(3).Text = Trim(Msg3)
            BARRAECF.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               BARRAECF.Panels.Add (4)
               BARRAECF.Panels(4).Text = Trim(Msg4)
               BARRAECF.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  BARRAECF.Panels.Add (5)
                  BARRAECF.Panels(5).Text = Trim(Msg5)
                  BARRAECF.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If
End Sub

Sub CANCELA_CUPOM_ABERTO()
'On Error GoTo ERRO_TRATA

   If NUMR_REQ_N > 0 Then
      SQL = "update PEDIDO set "
      SQL = SQL & " status = 9"
      SQL = SQL & " where numr_req = " & NUMR_REQ_N
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      CONECTA_RETAGUARDA.Execute SQL
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CANCELA_CUPOM_ABERTO"
End Sub

Private Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   Dim TabPedido As New ADODB.Recordset

   If Not IsNull(NUMR_REQ_N) Then
      If NUMR_REQ_N > 0 Then
         
         SINAL_INDICADOR_N = 1

         If INDR_FORM_ABERTO = True Then
            Unload frmCADRECEBVENDA
            INDR_FORM_ABERTO = False
         End If

'===================================
'===================================
         'atualizando cabeça
         SQL = "UPDATE PEDIDO SET "
         SQL = SQL & " cgccpf = '" & Trim(CNPJCPF_CLIENTE) & "'"
         SQL = SQL & " , nome_cliente = '" & Trim(NOME_CLIENTE) & "'"
         SQL = SQL & " , status = 2"
         SQL = SQL & " , valor_recebido = " & tpMOEDA(txtValorRecebido.Text)
         SQL = SQL & " , valor_total = " & tpMOEDA(txtValorTotal.Text)

SQL = SQL & " , valor_desconto = " & tpMOEDA(txtDesconto.Text)

         SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         CONECTA_RETAGUARDA.Execute SQL

         'If TabTemp.State = 1 Then _
            TabTemp.Close

         'SQL = "select contabiliza from TIPOVENDA "
         'SQL = SQL & " where tipovenda_id = " & cmbAuxTIPOVENDA.Text
         'TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         'If Not TabTemp.EOF Then

         '   If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
         '      If TabTemp.Fields("contabiliza").Value = True Then
         '         If TabTemp.State = 1 Then _
                     TabTemp.Close

         frmCADRECEBVENDA.Show 1

                  'Exit Sub
         '         Else
         '            SQL = "update PEDIDO set "
         '            SQL = SQL & "status = 6 " 'não contabiliza
         '            SQL = SQL & " where numr_req = " & NUMR_REQ_N
         '            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         '            CONECTA_RETAGUARDA.Execute SQL
         '      End If
         '   End If
         'End If
         'If TabTemp.State = 1 Then _
            TabTemp.Close
'===================================
'===================================
'===================================
         If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And CODG_USU_N = 144) Then
            CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
            Select Case CRITERIO
               Case "BEMATECH"
                  'Verifica se a Impressa esta ligada ou nao
                  RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
                  If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
                     RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
                     MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!", vbCritical, "MEGASIM"
                     Exit Sub
                  End If
               Case "DARUMA"
                  
               Case "Sweda"
                  
            End Select
         End If
'===================================
         If TabPedido.State = 1 Then _
            TabPedido.Close

         SQL = "select * from PEDIDO "
         SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPedido.EOF Then
            PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
            If TabPedido!Status = 5 Then
               CNPJCPF_A = Trim(TabPedido!CGCCPF)
   
               '====================================
               If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And CODG_USU_N = 144) Then
                  Msg = "Confirma Fechamento Venda ?"
                  PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     '==============================
                     CRITERIO = Trim(UCase(TRAZ_DESCRITOR("C", IMPRESSORA_FISCAL_N)))
                     Select Case CRITERIO
                        Case "BEMATECH"
                           'Verifica se a Impressa esta ligada ou nao
                           RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
                           If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
                              RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
                              MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!!!", vbCritical, "MEGASIM"
                              Exit Sub
                              Else

                                 'BlockInput True   'Bloqueia o teclado
'====================
frmDISPLAYEMISSOR.IMPRIME_CUPOM_FISCAL
'====================

'                                    FECHA_CUPOM_FISCAL
                                 BlockInput False  'Desbloqueia o teclado
                           End If
                        Case "DARUMA"

                        Case "Sweda"
                     End Select

                     '=======================
                     Me.WindowState = 0

                     If Trim(NUMEROCUPOM) <> "" And NUMR_REQ_N > 0 Then
                        SQL = "update PEDIDO set "
                        SQL = SQL & "status = 7 " 'CUPOM FISCAL
                        SQL = SQL & ", numr_cupom =  " & NUMEROCUPOM
                        SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

                        CONECTA_RETAGUARDA.Execute SQL
                     End If
                  End If
               End If
            End If
         End If
         If TabPedido.State = 1 Then _
            TabPedido.Close
      End If
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub

Sub FECHA_CUPOM_FISCALold()
'On Error GoTo ERRO_TRATA

   Dim INDR_CARTAO As Boolean
   INDR_CARTAO = False

   'Faz as verificações de TEF
   If frmINICIO.UsarTEF Then
      Dim Formas              As Variant
      Dim Valores             As Variant
      Dim i                   As Integer
      Dim OperacaoECFOK       As Boolean
      Dim Parametros          As Variant
      Dim ValorTotal          As Double

      Screen.MousePointer = vbHourglass

      ' inicia variáveis
      OperacaoECFOK = False

      i = 0
      Formas = Array("")
      Valores = Array("")

      SQL = "SELECT ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
      SQL = SQL & " FROM LANCAMENTO "
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
      SQL = SQL & " INNER JOIN FORMAPAGTO "
      SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

      SQL = SQL & " where LANCAMENTO.numr_doc = " & NUMR_REQ_N
      SQL = SQL & " and LANCAMENTO.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         ' se for uma forma de pagamento de cartão
         If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CARTAO") > 0 Then
            INDR_CARTAO = True
            If i > 0 Then
               ReDim Preserve Formas(UBound(Formas) + 1)
               ReDim Preserve Valores(UBound(Valores) + 1)
            End If
            Formas(i) = Left(UCase(Trim(TabTemp.Fields("descricao").Value)), 15)
            Valores(i) = Format(TabTemp.Fields("valor_item").Value, strFormatacao2Digitos)
            i = i + 1
         ElseIf InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CHEQUE") > 0 Then

            If MsgBox("Gostaria de consultar este cheque junto à SERASA (Somente Redecard)?", _
               vbYesNo + vbQuestion, "Consulta de Cheque") = vbYes Then

               If Not frmINICIO.ConsultarCheque(TabTemp.Fields("valor_item").Value, _
                  Left(UCase(Trim(TabTemp.Fields("descricao").Value)), 15)) Then
                  Exit Sub
               End If
            End If
         End If

         TabTemp.MoveNext
      Wend

      ' se encontrou alguma forma de pagamento de cartão
      If i > 0 Then
         If Not frmINICIO.tratarPagamentoComCartao(Valores, Formas) Then
            Exit Sub
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

'===================================================================
'===================================================================
'===================================================================
'===================================================================

   VALOR_ITEM_N = 0

   Mensagem_Final = "Obrigado, Volte Sempre."
   While Len(Mensagem_Final) < 48
      Mensagem_Final = Mensagem_Final & " "
   Wend

'=======================
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If Trim(txtCNPJCPF.Text) <> "99999999999" Then
         If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
         txtCNPJCPF.PromptInclude = True

         SQL = "Cliente: " & Trim(txtCNPJCPF.Text)

         While Len(SQL) < 48
            SQL = SQL & " "
         Wend

         Mensagem_Final = Mensagem_Final & SQL
      End If
   End If

   If Trim(NOME_CLIENTE) <> "Consumidor Final" Then
      SQL = Trim(Left(NOME_CLIENTE, 48))

      While Len(SQL) < 48
         SQL = SQL & " "
      Wend

      Mensagem_Final = Mensagem_Final & SQL
   End If
'=======================

   SQL = "NºPedido =  " & NUMR_REQ_N
   While Len(SQL) < 48
      SQL = SQL & " "
   Wend

   Mensagem_Final = Mensagem_Final & SQL

   NOME_VENDEDOR = "Balcão"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select nome_vend from VENDEDOR, PEDIDO "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and VENDEDOR.vendedor_id = PEDIDO.vendedor_id "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         NOME_VENDEDOR = Trim(TabTemp.Fields(0).Value)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "Vendedor: " & Trim(NOME_VENDEDOR)
   While Len(SQL) < 48
      SQL = SQL & " "
   Wend

   Mensagem_Final = Mensagem_Final & SQL

   CONTA_TENTATIVA = 0

INICIANDO_FECHAMENTO_CUPOM_FISCAL:

   ' desconto cielo premia
   Parametros = Array("D", "$", Format(frmINICIO.EasyTEF.ValorCampo709_000 + VALOR_TOTAL_DESCONTO_N, "#0.00"))
   Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeIniciarFechamentoCupomFiscal, Parametros, OperacaoECFOK)

   If OperacaoECFOK = False Then
      MsgBox "Não foi possível iniciar o fechamento do cupom fiscal.", vbCritical
      Exit Sub
   End If

   Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Iniciando Fechamento Cupom Fiscal"

   CONTA_TENTATIVA = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
   SQL = SQL & " FROM LANCAMENTO "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO "
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " where LANCAMENTO.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and LANCAMENTO.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

EFETUANDO_FORMA_DE_PAGAMENTO:

      ITEM_DESCONTO_N = 0 & TabTemp.Fields("valor_desconto").Value
      Descr_Forma_Pagto = "" & Trim(TabTemp.Fields("descricao").Value)
      If UCase(TabTemp.Fields("descricao").Value) = UCase("Dinheiro") Then _
         Descr_Forma_Pagto = "Dinheiro"

      ' Formas de pagamento que NÃO são de cartão
      If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CARTAO") = 0 Then
         Parametros = Array(Trim(TabTemp.Fields("descricao").Value), _
             Replace(Format(TabTemp.Fields("valor_item").Value - ITEM_DESCONTO_N, strFormatacao2Digitos), ",", "."))

         Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeEfetuarFormaPagamento, Parametros, OperacaoECFOK)

         ' A variável operacaoECFOK retorna se o comando da ECF foi executado
         ' com sucesso ou não
         If Not OperacaoECFOK Then
            MsgBox "Não foi possível efetuar a forma de pagamento.", vbCritical
            Exit Sub
         End If
      End If

      Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Efetuando Forma de Pagamento"
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   ' se houve pagamento com cartão
   ' usa o método automático para efetuar as formas de pagamento de maneira
   ' simples, ou seja, somente descrição da forma de pagamento de cartão
   ' e o valor de cada forma de pagamento
   If Not (frmINICIO.EasyTEF.OperacaoTEFAtual = ttCheque) Then
      If Not frmINICIO.EasyTEF.EfetuarFormasPagamentoCartao Then
         MsgBox "Não foi possível efetuar a(s) forma(s) de pagamento de cartão.", vbCritical
         Exit Sub
      End If
   End If

   CONTA_TENTATIVA = 0

Finalizando_Fechamento_Cupom_Fiscal:

   Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeTerminarFechamentoCupomFiscal, Array(Mensagem_Final), OperacaoECFOK)

   If Not OperacaoECFOK Then
      MsgBox "Não foi possível terminar o fechamento do cupom fiscal.", vbCritical
      Exit Sub
   End If
BlockInput False  'Desbloqueia o teclado
   Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Finalizando Fechamento Cupom Fiscal"

If INDR_CARTAO = True Then
   '=================================================
   '=================================================
   'imprime todos os cupons tef de transações aprovadas
   Call frmINICIO.EasyTEF.ImprimirCuponsECF
   '=================================================
   '=================================================
End If

   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NUMEROCUPOM = Space(1)
      Else: NUMEROCUPOM = Space(6)
   End If

   NUMR_SEQ_N = 0
   CONTA_TENTATIVA = 0

LE_ULTIMO_ECF:
   Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Ultimo Cupom Impresso"
      RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOM)
   Me.Caption = "Aguarde, Imprimindo Cupom Fiscal, " & Msg & " Ultimo Cupom Impresso"

   If Trim(NUMEROCUPOM) = "" Then
      'MsgBox "Atenção, erro de comunicação com impressora. Cupom Fiscal não gravado."
      Else
         If IsNumeric(NUMEROCUPOM) Then
            GRAVA_CUPOM Str(NUMR_REQ_N), NUMEROCUPOM

            If IsNumeric(NUMEROCUPOM) Then
               SQL = "update PEDIDO set "
               SQL = SQL & " numr_CUPOM = " & NUMEROCUPOM
               SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

               CONECTA_RETAGUARDA.Execute SQL
            End If

            Me.Caption = "OK, " & Msg & " " & "Fim Impressão"
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_CUPOM_FISCAL"
   INDR_VENDA = False
   CRITERIO = ""
End Sub

Private Function ObterValoresTransacaoAnteriorCartao() As Double
   Dim i As Integer
   Dim Acumulador As Double

   Acumulador = 0

   If frmINICIO.EasyTEF.OperacaoTEFAtual <> ttCheque Then
      For i = LBound(frmINICIO.EasyTEF.ValoresCartoes) To UBound(frmINICIO.EasyTEF.ValoresCartoes)
         Acumulador = Acumulador + Val(frmINICIO.EasyTEF.ValoresCartoes(i))
      Next i
   End If

   ObterValoresTransacaoAnteriorCartao = Acumulador
End Function

Private Function tratarPagamentoComCartao(ByRef ValorCartao As Double) As Boolean
   Dim i          As Integer
   Dim Resultado  As Boolean

   ' inicia as variáveis
   'ValorCartao = 0
   Resultado = True

   'frmINICIO.EasyTEF.NumeroDeCartoes = 0
   frmINICIO.EasyTEF.NumeroDeCartoes = 1

   ' Se houver um pagamento com 1 cartão
   'If IIf(ValorCartao = "", "0,00", ValorCartao) > 0 Then

   '   frmINICIO.EasyTEF.NumeroDeCartoes = 1

      ' Se houver um pagamento com 2 cartões
      'If IIf(edtValorCartao2.Text = "", "0,00", edtValorCartao2.Text) > 0 Then

      '    EasyTEF.NumeroDeCartoes = 2

      '    ' Se houver um pagamento com 3 cartões
      '    If IIf(edtValorCartao3.Text = "", "0,00", edtValorCartao3.Text) > 0 Then

      '        EasyTEF.NumeroDeCartoes = 3

      '    End If
      'End If
   'End If

   frmINICIO.EasyTEF.ImprimirComprovante = False
   For i = 1 To frmINICIO.EasyTEF.NumeroDeCartoes
      'If i = 1 Then
      '    ValorCartao = edtValorCartao1.Text
      'ElseIf i = 2 Then
      '    ValorCartao = edtValorCartao2.Text
      'ElseIf i = 3 Then
      '    ValorCartao = edtValorCartao3.Text
      'End If

      Call frmINICIO.EasyTEF.PagarNoCartao(ValorCartao, tmReal, NUMEROCUPOM, _
          i = 1, i = frmINICIO.EasyTEF.NumeroDeCartoes, FORMA_PGTO_CARTAO)

      Resultado = frmINICIO.EasyTEF.TransacaoAprovada
      If Not frmINICIO.EasyTEF.TransacaoAprovada Then
         MsgBox "Não foi possível finalizar com sucesso o pagamento com cartão", vbCritical
         Exit For
      End If

      ' Caso fosse necessário mudar a descrição da forma de pagamento
      ' após a transação ser aprovada, o método a ser usado é o seguinte
      '
      'EasyTEF.AlterarNomeUltimaFormaPagamento(NomeDaFormaDePagamento)
   Next i

   'If UsuarioNaoQuerOutraFormaPgto Then
      'Call LimparTela
   'End If

   tratarPagamentoComCartao = Resultado
End Function

Sub FECHA_CUPOM_FISCAL()
'On Error GoTo ERRO_TRATA

   Dim ValorTotal    As Double
   Dim ValorDinheiro As Double
   Dim ValorCheque   As Double
   Dim ValorCartao   As Double
   Dim Parametros    As Variant
   Dim Desconto      As String
   Dim Tipodesc      As String
   Dim Valor         As String
   Dim i             As Integer
   Dim RETORNO_ECF       As String
   Dim OperacaoECFOK As Boolean

   Screen.MousePointer = vbHourglass

   ' inicia variáveis
   OperacaoECFOK = False
   ValorTotal = 0
   ValorDinheiro = 0
   ValorCheque = 0
   ValorCartao = 0


   ' obtem o total do cupom fiscal
   ' esta chamada é obrigatória mesmo que não se deseje obter o total do cupom fiscal
   'Parametros = Array("0")
   'ValorTotal = frmINICIO.EasyTEF.TratarCupomFiscal(tmeSubTotalizarCupom, Parametros, OperacaoECFOK) / 100

'======================================
   i = 0
   'Formas = Array("")
   'Valores = Array("")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT ITEMLANCAMENTO.VALOR_ITEM, ITEMLANCAMENTO.VALOR_DESCONTO, FORMAPAGTO.DESCRICAO"
   SQL = SQL & " FROM LANCAMENTO "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO "
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " where LANCAMENTO.numr_doc = " & NUMR_REQ_N
   SQL = SQL & " and LANCAMENTO.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      ValorTotal = ValorTotal + TabTemp.Fields("valor_item").Value

      ' se for uma forma de pagamento de cartão
      If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CARTAO") > 0 Then
         ValorCartao = ValorCartao + TabTemp.Fields("valor_item").Value
         Else
            If InStr(1, UCase(Trim(TabTemp.Fields("descricao").Value)), "CHEQUE") > 0 Then
               ValorCheque = ValorCheque + TabTemp.Fields("valor_item").Value

               'If MsgBox("Gostaria de consultar este cheque junto à SERASA (Somente Redecard)?", _
               '   vbYesNo + vbQuestion, "Consulta de Cheque") = vbYes Then

               '   If Not frmINICIO.ConsultarCheque(TabTemp.Fields("valor_item").Value, _
               '      Left(UCase(Trim(TabTemp.Fields("descricao").Value)), 15)) Then
               '      Exit Sub
               '   End If
               'End If
               Else: ValorDinheiro = ValorDinheiro + TabTemp.Fields("valor_item").Value
            End If
      End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
'======================================
   'ValorDinheiro = ValorDinheiro
   'ValorCheque = ValorCheque
   'ValorCartao = CDbl(edtValorCartao1.Text) + CDbl(edtValorCartao2.Text) + CDbl(edtValorCartao3.Text)

   If (ValorTotal < (ValorDinheiro + ValorCheque + ValorCartao + ObterValoresTransacaoAnteriorCartao)) Then
      Screen.MousePointer = vbDefault
      MsgBox "Total das formas de pagamento diferente do total do cupom.", vbCritical
      Exit Sub
   End If

   'If CDbl(edtValorCartao1.Text) > 0 Then
   If CDbl(ValorCartao) > 0 Then
      If Not tratarPagamentoComCartao(ValorTotal) Then
         MsgBox "Não foi possível terminar o pagamento com cartão.", vbCritical
         'Call voltarCursorAoNormal
         Exit Sub
      End If
   End If

   Desconto = "D"
   Tipodesc = "$"
   Valor = Format(frmINICIO.EasyTEF.ValorCampo709_000, "#0.00") ' desconto cielo premia

   Parametros = Array(Desconto, Tipodesc, Valor)
   RETORNO_ECF = frmINICIO.EasyTEF.TratarCupomFiscal(tmeIniciarFechamentoCupomFiscal, Parametros, OperacaoECFOK)

   If OperacaoECFOK = False Then
      'Call voltarCursorAoNormal
      MsgBox "Não foi possível iniciar o fechamento do cupom fiscal.", vbCritical
      Exit Sub
   End If

   If Val(ValorDinheiro) > 0 Then
      Parametros = Array("Dinheiro", Format(ValorDinheiro, FORMATO_MONEY))
      Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeEfetuarFormaPagamento, Parametros, OperacaoECFOK)

      ' A variável operacaoECFOK retorna se o comando da ECF foi executado
      ' com sucesso ou não
      If Not OperacaoECFOK Then
         MsgBox "Não foi possível efetuar a forma de pagamento 'Dinheiro'.", vbCritical
         Exit Sub
      End If
   End If

   If Val(ValorCheque) > 0 Then
      Parametros = Array(FORMA_PGTO_CHEQUE, Format(ValorCheque, FORMATO_MONEY))
      Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeEfetuarFormaPagamento, Parametros, OperacaoECFOK)

      If Not OperacaoECFOK Then
         MsgBox "Não foi possível efetuar a forma de pagamento 'Cheque'.", vbCritical
         Exit Sub
      End If
   End If

   ' se houve pagamento com cartão
   ' usa o método automático para efetuar as formas de pagamento de maneira
   ' simples, ou seja, somente descrição da forma de pagamento de cartão
   ' e o valor de cada forma de pagamento
   If Not (frmINICIO.EasyTEF.OperacaoTEFAtual = ttCheque) Then
      If Not frmINICIO.EasyTEF.EfetuarFormasPagamentoCartao Then
         MsgBox "Não foi possível efetuar a(s) forma(s) de pagamento de cartão.", vbCritical
         Exit Sub
      End If
   End If

'===========
   Mensagem_Final = "Obrigado, Volte Sempre."
   While Len(Mensagem_Final) < 48
      Mensagem_Final = Mensagem_Final & " "
   Wend

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      If Trim(txtCNPJCPF.Text) <> "99999999999" Then
         If Len(Trim(txtCNPJCPF.Text)) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
         txtCNPJCPF.PromptInclude = True

         SQL = "Cliente: " & Trim(txtCNPJCPF.Text)

         While Len(SQL) < 48
            SQL = SQL & " "
         Wend

         Mensagem_Final = Mensagem_Final & SQL
      End If
   End If

   If Trim(NOME_CLIENTE) <> "Consumidor Final" Then
      SQL = Trim(Left(NOME_CLIENTE, 48))

      While Len(SQL) < 48
         SQL = SQL & " "
      Wend

      Mensagem_Final = Mensagem_Final & SQL
   End If

   SQL = "NºPedido =  " & NUMR_REQ_N
   While Len(SQL) < 48
      SQL = SQL & " "
   Wend

   Mensagem_Final = Mensagem_Final & SQL

   NOME_VENDEDOR = "Balcão"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select nome_vend from VENDEDOR, PEDIDO "
   SQL = SQL & " where pedido_id = " & NUMR_REQ_N
   SQL = SQL & " and VENDEDOR.vendedor_id = PEDIDO.vendedor_id "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         NOME_VENDEDOR = Trim(TabTemp.Fields(0).Value)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "Vendedor: " & Trim(NOME_VENDEDOR)
   While Len(SQL) < 48
      SQL = SQL & " "
   Wend

   Mensagem_Final = Mensagem_Final & SQL
'===========

   Parametros = Array(Mensagem_Final)
   Call frmINICIO.EasyTEF.TratarCupomFiscal(tmeTerminarFechamentoCupomFiscal, Parametros, OperacaoECFOK)

   If Not OperacaoECFOK Then
      MsgBox "Não foi possível terminar o fechamento do cupom fiscal.", vbCritical
      Exit Sub
   End If

Sleep 6000

   'Verifica se a Impressa esta ligada ou nao
   RETORNO_ECF = Bematech_FI_VerificaImpressoraLigada()
   If RETORNO_ECF <> 1 Then 'Se For + a 1 esta perfeito , diferente de 1 ela esta desligada
      RETORNO_ECF = 0 'Aqui eu zero a variavel para que caia no loop de impressora desligada
      MsgBox "ECF Desligado, Ligue a Impressora Para Continuar!", vbCritical, "MEGASIM"
      Exit Sub
   End If

   'INDR_CUPOM_ABERTO = False
   Call VerificaRetornoImpressora("Bematech_FI_AbreCupom", "", "Emissão de Cupom Fiscal")
   'If INDR_CUPOM_ABERTO = True Then _

   ' imprime todos os cupons tef de transações aprovadas
   Call frmINICIO.EasyTEF.ImprimirCuponsECF

   Screen.MousePointer = vbDefault

Exit Sub
ERRO_TRATA:
   'TRATA_ERROS Err.Description, Me.Name, "CANCELA_CUPOM_ABERTO"
   End
End Sub

Sub LE_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   CODG_PRODUTO_A = Trim(txtProduto.Text)
   INDR_PROD_BALANCA = False

   'LE POR CODIGO DE PRODUTO
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQTDE.Enabled = True
txtQTDE.SetFocus
      Exit Sub
   End If

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   QTDE_N = 0
   CRITERIO = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close
'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   SQL = "select count(produto_id) from PRODUTO "
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO = Trim(CODIGO_BARRAS)

            frmPEDIDOBARRAS.Show 1

            If Trim(CRITERIO) <> "" Then
               txtProduto.Text = Trim(CRITERIO)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select * from PRODUTO "
               SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               SQL = SQL & " and situacao <> 'C' "
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then
                  MOSTRA_DADOS_PRODUTO
                  Call txtQTDE_KeyPress(13)
               End If

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               CRITERIO = ""
               Exit Sub
            End If
         End If
      End If
   End If

CRITERIO = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO "
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS) & "'"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      MOSTRA_DADOS_PRODUTO

      If TabProduto.State = 1 Then _
         TabProduto.Close
txtQTDE.Text = 1
Call txtQTDE_KeyPress(13)
txtProduto.SetFocus
      Exit Sub
   End If

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS) = 13 Then
      '2 = produtos "in store" (sempre será 2)     1
      'C = código do produto (4,5 ou 6 dígitos)    2 a 8
      'T = total a pagar (sempre 6 dígitos)        9 a 13
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      If INDR_PANIFICADORA = True Then
         txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 4))

         If TabProduto.State = 1 Then _
            TabProduto.Close
      
         SQL = "select * from PRODUTO "
         SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            INDR_PROD_BALANCA = False
            If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then
               INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value
               QTDE_N = 1

'removido do mostra_dados
'panificadora
   If INDR_PANIFICADORA = True And Trim(CODIGO_BARRAS) <> "" And INDR_PROD_BALANCA = True Then
      VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS, 8, 6) / 1000
      'txtValorItem.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

      QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value)

      QTDE_PEDIDO = QTDE_N
         'txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
      txtQTDE.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)
      QTDE_N = 1
   End If

               MOSTRA_DADOS_PRODUTO
               Call txtQTDE_KeyPress(13)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               Exit Sub
               Else: MsgBox "Verificar cadastro produto."
            End If
            Else: MsgBox "Verificar cadastro produto."
         End If
         Else
            txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS, 2, 6))  'PORTO SEGURO

            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select * from PRODUTO "
            SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then
               QTDE_N = 0 & Int(Mid(CODIGO_BARRAS, 8, 5))   'gramas
               PESO_ITEM_N = QTDE_N
               'txtPesoItem.Text = Format(PESO_ITEM_N, strFormatacao3Digitos)

               MOSTRA_DADOS_PRODUTO
               Call txtQTDE_KeyPress(13)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               Exit Sub
            End If
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(CODIGO_BARRAS) = 12 Then
      'lendo codigo barras ultralav
      '100004360813
      '100002361113
      '1-1 = masculino ou feminino
      '2-7 = código do produto
      '8-9 = numeração tamanho produto
      '10-11 = mes
      '12-13 = ano

      txtProduto.Text = "" & Mid(CODIGO_BARRAS, 1, 6)
      SqL2 = "" & Mid(CODIGO_BARRAS, 7, 2)

      SQL = "select * from PRODUTO "
      SQL = SQL & " where referencia = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and RIGHT(descricao,2) = '" & Trim(SqL2) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         MOSTRA_DADOS_PRODUTO

         If TabProduto.State = 1 Then _
            TabProduto.Close

         txtQTDE.Text = 1
         'txtQTDE.SetFocus
         'Call txtQtde_LostFocus

         'txtDesconto.SetFocus
         Call txtQTDE_KeyPress(13)

         txtProduto.SetFocus

         Exit Sub
      End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MsgBox "Produto não cadastrado."
   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_PRODUTO"
End Sub

Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   txtDescricao.Text = ""
   txtUN.Text = ""
   txtValorItem.Text = ""

   Indr_Achou_Produto = True

   txtProduto.Text = TabProduto!CODG_PRODUTO
   PRODUTO_ID_N = 0 & TabProduto!PRODUTO_ID
   txtDescricao.Text = TabProduto!Descricao
   txtUN.Text = "" & TabProduto!Unidade_Medida
   lblST.Caption = ""
   
   lblST2.Caption = ""

   If Not IsNull(TabProduto.Fields("aliquota_icms").Value) Then _
      Aliquota_Icms_N = 0 & TabProduto.Fields("aliquota_icms").Value

   If Not IsNull(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value) Then
      If Trim(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value) <> "" Then
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
         SQL = "select * from CST "
         SQL = SQL & " where codigo = " & Trim(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value)
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            lblST.Caption = "" & TabDESCR!codigo
            lblST2.Caption = "" & Trim(TabDESCR!Descricao)
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      End If
   End If

   If Trim(TabProduto!SITUACAO) <> "A" Then
      MsgBox "Produto não liberado para venda."
      txtProduto.Text = ""
      Indr_Achou_Produto = False
      txtProduto.SetFocus
      Exit Sub
   End If

   If IsNull(TabProduto!PRECO_VENDA) Then
      MsgBox "Produto Sem Preço de Venda."
      txtProduto.Text = ""
      Indr_Achou_Produto = False
      txtProduto.SetFocus
      Exit Sub
   End If

   VALOR_ITEM_N = 0 & TabProduto.Fields("preco_venda").Value
   If VALOR_ITEM_N <= 0 Then
      MsgBox "Produto Sem Preço de Venda."
      txtProduto.Text = ""
      Indr_Achou_Produto = False
      txtProduto.SetFocus
      Exit Sub
   End If

   txtValorItem.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)

   INDR_PROD_BALANCA = False
   If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then _
      INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value

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

   txtValorItem.Text = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
   STATUS_PROD = TabProduto!SITUACAO
   CODG_PRODUTO_A = Trim(txtProduto.Text)

   If INDR_ESTQ_NEGATIVO = False Then
      QTDE_ESTOQUE = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      If QTDE_ESTOQUE <= 0 Then
         MsgBox "Produto sem estoque disponível."
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Not IsNull(TabProduto.Fields("codg_ncm").Value) Then
      If Len(TabProduto.Fields("codg_ncm").Value) > 2 Then
         If Len(TabProduto.Fields("codg_ncm").Value) < 8 Then
            MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"
            LIMPA_GRID
            txtProduto.SetFocus
         End If
      End If
   End If

   PRODUTO_ID_N = TabProduto.Fields("produto_id").Value

'=====================
   'If Trim(txtSeq.Text) = "" Then
   '   SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
   '   Else
   '      If Not IsNumeric(txtSeq.Text) Then
   '         SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
   '         Else: SEQ_ID_N = txtSeq.Text
   '      End If
   'End If
   'txtSeq.Text = SEQ_ID_N
'=====================

   'PEDIDO_ID_N = Trim(txtPedido.Text)

'panificadora
   If QTDE_N <= 0 Then _
      QTDE_N = 1

   If INDR_PANIFICADORA = True And Trim(CODIGO_BARRAS) <> "" And INDR_PROD_BALANCA = True Then
      VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS, 8, 6) / 1000
      txtValorItem.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

      QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value)

      QTDE_PEDIDO = QTDE_N
      QTDE_N = 1
   End If

   If TabProduto.State = 1 Then _
      TabProduto.Close

   If Len(Trim(CODIGO_BARRAS)) = 13 Then
      If Trim(txtValorItem.Text) <> "" Then
         If IsNumeric(txtValorItem.Text) Then
'=======
'foi para le_produto
            'txtQtde.Text = QTDE_N
            If INDR_PANIFICADORA = True And INDR_PROD_BALANCA = True Then
               txtQTDE.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)
               'Else: txtQtde.Text = Format(QTDE_N / 1000, strFormatacao3Digitos)
            End If
'================

            CODIGO_BARRAS = ""
            txtProduto.SetFocus
         End If
      End If
   End If
   CODIGO_BARRAS = ""

   'txtValorItem.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PRODUTO"
End Sub
