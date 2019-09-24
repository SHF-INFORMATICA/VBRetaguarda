VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAcertoFunc 
   Caption         =   "Acerto Funcionário"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AcertoFunc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Acerto Pedido Funcionário"
      TabPicture(0)   =   "AcertoFunc.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label14"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lstPedido"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lstProduto"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDtFim"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtDtIni"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbFunc"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbFuncAUX"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTotalVenda"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtReg"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkConceder"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtValorProducao"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtQtdeProduto"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtQtdeItemProducao"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtValorAcerto"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtValorCompra"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtValorRevenda"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtValorPermitido"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtUltrapassado"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtValorDiario"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmdConsulta"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "chkAbertos"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      Begin VB.CheckBox chkAbertos 
         Caption         =   "Pendentes"
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
         Left            =   10200
         TabIndex        =   36
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   8760
         Picture         =   "AcertoFunc.frx":5C2E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5760
         Width           =   405
      End
      Begin VB.TextBox txtValorDiario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   10545
         TabIndex        =   33
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txtUltrapassado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   31
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtValorPermitido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   29
         Text            =   "100,00"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtValorRevenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   27
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtValorCompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   25
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtValorAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   10560
         TabIndex        =   23
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtQtdeItemProducao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4200
         TabIndex        =   21
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox txtQtdeProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1920
         TabIndex        =   19
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox txtValorProducao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10560
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkConceder 
         Caption         =   "Conceder produção?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   3
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtReg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   8040
         TabIndex        =   13
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox txtTotalVenda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10530
         TabIndex        =   11
         Top             =   6360
         Width           =   1095
      End
      Begin VB.ComboBox cmbFuncAUX 
         BackColor       =   &H80000000&
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
         Left            =   8280
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbFunc 
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
         Left            =   8280
         TabIndex        =   0
         Top             =   600
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   8280
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Height          =   375
         Left            =   10440
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
      Begin MSComctlLib.ListView lstProduto 
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7223
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   14737632
         Appearance      =   1
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qtde"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Valor Item"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Peso"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total Item"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lstPedido 
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2143
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   128
         BackColor       =   16777152
         Appearance      =   1
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DtPedido"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vendedor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Caixa"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Diário Permitido ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9330
         TabIndex        =   34
         Top             =   5640
         Width           =   1470
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ultrapassado ="
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
         TabIndex        =   32
         Top             =   3960
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Permitido ="
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
         Left            =   8775
         TabIndex        =   30
         Top             =   3480
         Width           =   1680
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Revenda ="
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
         Left            =   8835
         TabIndex        =   28
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Compras ="
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
         Left            =   8865
         TabIndex        =   26
         Top             =   3000
         Width           =   1590
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         X1              =   8666
         X2              =   8666
         Y1              =   1920
         Y2              =   6240
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Acerto ="
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
         Left            =   9000
         TabIndex        =   24
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ItensProdução = "
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
         Left            =   2520
         TabIndex        =   22
         Top             =   6360
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde.Produto(s) = "
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
         Left            =   120
         TabIndex        =   20
         Top             =   6360
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Func.(a):"
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
         Left            =   7200
         TabIndex        =   18
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Produção ="
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
         Left            =   8790
         TabIndex        =   17
         Top             =   2040
         Width           =   1665
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   11760
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   2
         X1              =   0
         X2              =   11760
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde.Pedidos = "
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
         Left            =   6480
         TabIndex        =   14
         Top             =   6360
         Width           =   1515
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr.Total.Pedidos = "
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
         Left            =   8640
         TabIndex        =   12
         Top             =   6360
         Width           =   1875
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Inicial:"
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
         Left            =   7200
         TabIndex        =   7
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Final:"
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
         Left            =   9480
         TabIndex        =   6
         Top             =   1080
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1270
      ButtonWidth     =   3043
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
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "BaixarPedido"
            Key             =   "baixar"
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
               Picture         =   "AcertoFunc.frx":B230
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoFunc.frx":C3CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoFunc.frx":D459
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoFunc.frx":E40E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AcertoFunc.frx":F519
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   7560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11880
      DesignHeight    =   7785
   End
End
Attribute VB_Name = "frmAcertoFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_PRODUCAO_N       As Double
   Dim VALOR_PERMITIDO_N      As Double
   Dim VALOR_ULTRAPASSADO_N   As Double
   Dim Conta_Produto_N        As Long
   Dim ITEM_PRODUCAO_N        As Long
   Dim VALOR_COMPRA_N         As Double
   Dim VALOR_ACERTO_N         As Double
   Dim VALOR_REVENDA_N        As Double

Private Sub Form_Load()

   CARREGA_USUARIO

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select VLR_DIA_COMPRA_PROD from estabelecimento WITH (NOLOCK) "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         txtValorDiario.Text = Format(TabTemp.Fields(0).Value, strFormatacao2Digitos)
   If TabTemp.State = 1 Then _
      TabTemp.Close

   Call TXTDTINI_GotFocus
   Call TXTDTFIM_GotFocus
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "baixar"
         BAIXA_PEDIDO
      Case "limpar"
         LIMPA_TUDO
      Case "consultar"
         MONTA_CONSULTA_SQL
      Case "limpar"
         Call Form_Load
      Case "voltar"
         Unload Me
      Case "imprimir"
         'lstPedidoItem.ListItems.Clear
         'lstPedidoItem.Visible = False

         'MONTA_CONSULTA_SQL False
         'GERA_REL
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsulta_Click()
   If Trim(txtValorDiario.Text) <> "" Then
      If IsNumeric(txtValorDiario.Text) Then

         SQL = "update ESTABELECIMENTO set "
         SQL = SQL & " VLR_DIA_COMPRA_PROD = " & tpMOEDA(txtValorDiario.Text)
         SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         MsgBox "Processo realizado com sucesso."
      End If
   End If
End Sub

Private Sub lstpedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstPedido, ColumnHeader
End Sub

Private Sub lstProduto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstProduto, ColumnHeader
End Sub

Private Sub cmbFunc_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select pessoa_id,nome,dt_nasc from CLIENTE WITH (NOLOCK) "
            SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_A) & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then
               If TabUSU.State = 1 Then _
                  TabUSU.Close
   
               SQL = "select nome,pessoa_id from USUARIO WITH (NOLOCK) "
               SQL = SQL & " where cpf = '" & Trim(CNPJCPF_A) & "'"
               TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabUSU.EOF Then
                  Msg = "Confirma cadastro desse CPF como funcionário ?"
                  PERGUNTA Msg, vbYesNo + 32, "Inclusão de Funcionário", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     SQL = "INSERT INTO USUARIO "
                        SQL = SQL & " (empresa_id,usuario_id,Nome,Cpf,Status,Pessoa_id,FUNCIONARIO) "
                     SQL = SQL & " VALUES ("
                        SQL = SQL & EMPRESA_ID_N
                        SQL = SQL & "," & MAX_ID("usuario_id", "usuario", "empresa_id", "1", "", "")
                        SQL = SQL & ",'" & Trim(TabCliente.Fields("nome").Value) & "'"
                        SQL = SQL & ",'" & Trim(CNPJCPF_A) & "'"
                        SQL = SQL & ",'TRUE'"
                        SQL = SQL & "," & TabCliente.Fields("pessoa_id").Value
                        SQL = SQL & ",'true'"
                        SQL = SQL & ")"
                     CONECTA_RETAGUARDA.Execute SQL

                     CARREGA_USUARIO
                  End If
               End If
               If TabUSU.State = 1 Then _
                  TabUSU.Close
            End If
            If TabCliente.State = 1 Then _
               TabCliente.Close
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFunc_DropDown"
End Sub

Private Sub cmbFunc_GotFocus()
   cmbFunc.SelStart = 0
   cmbFunc.SelLength = Len(cmbFunc)
   cmbFunc.BackColor = &HC0FFFF
End Sub

Private Sub cmbfunc_Click()
'On Error GoTo ERRO_TRATA

   cmbFuncAUX.ListIndex = cmbFunc.ListIndex

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbfunc_Click"
End Sub

Private Sub cmbFunc_LostFocus()
   cmbFunc.BackColor = &HFFFFFF
End Sub

Private Sub txtDtFim_LostFocus()
   txtDtFim.BackColor = &HFFFFFF
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni)
   txtDtIni.BackColor = &HC0FFFF

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim)
   txtDtFim.BackColor = &HC0FFFF

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error   GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub txtValorDiario_GotFocus()
   txtValorDiario.SelStart = 0
   txtValorDiario.SelLength = Len(txtValorDiario)
   txtValorDiario.BackColor = &HC0FFFF
End Sub

Private Sub txtValorDiario_LostFocus()
   txtValorDiario.BackColor = &HFFFFFF
End Sub

Private Sub txtValorPermitido_Change()
   CALCULA_ULTRAPASSADO
End Sub

Private Sub txtValorPermitido_GotFocus()
   txtValorPermitido.SelStart = 0
   txtValorPermitido.SelLength = Len(txtValorPermitido)
   txtValorPermitido.BackColor = &HC0FFFF
End Sub

Private Sub txtValorPermitido_LostFocus()
   txtValorPermitido.BackColor = &HFFFFFF
End Sub

Private Sub chkConceder_Click()
   CALCULA_ULTRAPASSADO
End Sub

Private Sub chkAbertos_Click()
   If chkAbertos.Value = 1 Then
      chkAbertos.Caption = "Pendentes"
      Else: chkAbertos.Caption = "Liquidados"
   End If
   chkAbertos.Refresh
   DoEvents
End Sub

'=========================================
Sub CALCULA_ULTRAPASSADO()
'On Error GoTo ERRO_TRATA

   VALOR_ULTRAPASSADO_N = 0

   If chkConceder.Value = 1 Then
      VALOR_PERMITIDO_N = txtValorPermitido.Text
      Else: VALOR_PERMITIDO_N = 0
   End If

   VALOR_ULTRAPASSADO_N = "" & VALOR_PRODUCAO_N - VALOR_PERMITIDO_N
   txtUltrapassado.Text = "" & Format(VALOR_ULTRAPASSADO_N, strFormatacao2Digitos)

   If chkConceder.Value = 1 Then
      VALOR_ACERTO_N = VALOR_COMPRA_N + VALOR_ULTRAPASSADO_N - VALOR_PERMITIDO_N
      Else: VALOR_ACERTO_N = VALOR_COMPRA_N
   End If

   txtValorAcerto.Text = "" & Format(VALOR_ACERTO_N, strFormatacao2Digitos)

   txtValorRevenda.Text = "" & Format(VALOR_REVENDA_N, strFormatacao2Digitos)

   DoEvents

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CALCULA_ULTRAPASSADO"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   CONT_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_PRODUCAO_N = 0
   VALOR_PERMITIDO_N = 0
   VALOR_ULTRAPASSADO_N = 0
   ITEM_PRODUCAO_N = 0
   VALOR_REVENDA_N = 0
   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   CONT_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_PERMITIDO_N = 0
   VALOR_ULTRAPASSADO_N = 0
   VALOR_COMPRA_N = 0
   VALOR_REVENDA_N = 0

   lstPedido.ListItems.Clear
   lstProduto.ListItems.Clear
   txtValorProducao.Text = ""
   txtValorRevenda.Text = ""
   txtValorCompra.Text = ""
   txtUltrapassado.Text = ""
   txtValorAcerto.Text = ""

   cmbFunc.Text = ""
   cmbFuncAUX.Text = ""
   txtDtIni.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   chkConceder.Value = 1
   txtValorPermitido.Text = "100,00"
   txtUltrapassado.Text = ""
   txtQtdeProduto.Text = ""
   txtReg.Text = ""
   txtTotalVenda.Text = ""

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub CARREGA_USUARIO()
'On Error GoTo ERRO_TRATA

   cmbFunc.Clear
   cmbFuncAUX.Clear

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select nome,pessoa_id from USUARIO WITH (NOLOCK) "
   SQL = SQL & " where funcionario = 1"
   SQL = SQL & " and status = 1"
   SQL = SQL & " order by nome"
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabUSU.EOF

      cmbFunc.AddItem Trim(TabUSU.Fields("nome").Value)
      cmbFuncAUX.AddItem Trim(TabUSU.Fields("pessoa_id").Value)

      TabUSU.MoveNext
   Wend
   If TabUSU.State = 1 Then _
      TabUSU.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Sub CHECA_ULTIMO_DIA_MES()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
      txtDtFim.Text = ""

      txtDtIni.PromptInclude = True
      If IsDate(txtDtIni.Text) Then
         CRITERIO_A = FimDoMes(txtDtIni.Text, False)
         CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
         txtDtFim.Text = CRITERIO_A
         txtDtFim.PromptInclude = True
      End If
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CHECA_ULTIMO_DIA_MES"
End Sub

Sub MONTA_CONSULTA_SQL()
'On Error GoTo ERRO_TRATA

   If Trim(cmbFunc.Text) = "" Then
      MsgBox "Selecionar Funcionário."
      cmbFunc.SetFocus
      Exit Sub
   End If

   HORA_INI = Time

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", Format((HORA_INI), "hh:mm:ss")

   CHECA_ULTIMO_DIA_MES

   VALOR_TOTAL_N = 0

   txtTotalVenda.Text = ""
   txtReg.Text = ""
   
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   SqL2 = "select PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.VENDEDOR_ID, "
   SqL2 = SqL2 & " PEDIDO.CGCCPF, PEDIDO.DT_REQ, PEDIDO.STATUS, PEDIDO.VALOR_TOTAL , "
   SqL2 = SqL2 & " PEDIDO.NUMERO_CAIXA_CPU, PEDIDO.ESTABELECIMENTO_ID, USUARIO.USUARIO_ID, "
   SqL2 = SqL2 & " USUARIO.NOME, USUARIO.PESSOA_ID"

   SQL3 = "select count(PEDIDO.PEDIDO_ID) "

   SQL = " from PEDIDO WITH (NOLOCK) "

   SQL = SQL & " INNER JOIN USUARIO WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDO.CGCCPF = USUARIO.CPF"

   SQL = SQL & " where PEDIDO.pedido_id Is Not Null"

   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and PEDIDO.status in ('7','5')"

   'If Trim(cmbFunc.Text) = "" Then
      SQL = SQL & " and usuario.pessoa_id = " & cmbFuncAUX.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and PEDIDO.dt_req >= '" & DMA(txtDtIni.Text, "i") & "'"
      SQL = SQL & " and PEDIDO.dt_req <= '" & DMA(txtDtFim.Text, "f") & "'"
   End If

   SQL3 = SQL3 & " " & SQL

   SQL = SQL & " order by PEDIDO_ID desc"

   SqL2 = SqL2 & " " & SQL


   Dim TabTemp    As New ADODB.Recordset
   Dim INDR_VAI   As Boolean

   PEDIDO_ID_N = 0
   Conta_Produto_N = 0
   VALOR_TOTAL_N = 0
   CONT_N = 0
   CONTA_REGISTRO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_PRODUCAO_N = 0
   VALOR_PERMITIDO_N = 0
   VALOR_ULTRAPASSADO_N = 0
   ITEM_PRODUCAO_N = 0
   VALOR_COMPRA_N = 0
   VALOR_REVENDA_N = 0

   lstPedido.ListItems.Clear
   lstProduto.ListItems.Clear
   txtValorProducao.Text = ""
   txtValorRevenda.Text = ""
   txtValorCompra.Text = ""
   txtUltrapassado.Text = ""
   txtValorAcerto.Text = ""

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL3, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then _
      CONTA_REG_PROGRESSO = TabTemp.Fields(0).Value
'============================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If CONTA_REG_PROGRESSO > 0 Then
      ProgressBar1.Min = 0                   'Indica o valor inicial
      ProgressBar1.Max = CONTA_REG_PROGRESSO 'Indica o valor final
      'frmProgresso.Show 1
   End If
   CONT_N = 0

   TabTemp.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Nenhuma venda registrada para essa pesquisa."
      Me.Enabled = True
      Me.KeyPreview = True
      Exit Sub
   End If

   If Not TabTemp.EOF Then
      TabTemp.MoveFirst
      While Not TabTemp.EOF
         PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value
         '=====================
         'atualizando cabeça pedido valor total
         SQL = "update pedido set "
         SQL = SQL & " pedido.VALOR_TOTAL = (pedidoitem.QTD_PEDIDA * PEDIDOITEM.VALOR_ITEM) "
         SQL = SQL & " from PEDIDO WITH (NOLOCK) "
         SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
         SQL = SQL & " ON pedido.pedido_ID = pedidoitem.pedido_ID"
         SQL = SQL & " Where pedido.PEDIDO_ID = " & PEDIDO_ID_N
         SQL = SQL & " and pedidoitem.status <> 'C' "
         CONECTA_RETAGUARDA.Execute SQL
         '=====================
         TabTemp.MoveNext
      Wend
      TabTemp.MoveFirst
      While Not TabTemp.EOF
         DoEvents

         If CONT_N < CONTA_REG_PROGRESSO Then
            CONT_N = CONT_N + 1
            ProgressBar1.Value = CONT_N
         End If

         CONTA_REGISTRO_N = CONTA_REGISTRO_N + 1
         txtReg.Text = CONTA_REGISTRO_N
         txtReg.Refresh

         PEDIDO_ID_N = TabTemp.Fields("pedido_id").Value

         INDR_VAI = False
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select numr_doc from ITEMLANCAMENTO WITH (NOLOCK)"
         SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
         If chkAbertos.Value = 1 Then
            SQL = SQL & " and status = 'A' "
            Else
               If chkAbertos.Value = 0 Then _
                  SQL = SQL & " and status = 'B' "
         End If
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               INDR_VAI = True
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         If INDR_VAI = True Then
            Set item = lstPedido.ListItems.Add(, "seq." & TabTemp.Fields("PEDIDO_ID").Value, TabTemp.Fields("PEDIDO_ID").Value)

            item.SubItems(1) = "" & TabTemp.Fields("dt_req").Value
            item.SubItems(2) = "" & Format(TabTemp.Fields("valor_total").Value, strFormatacao2Digitos)

            item.SubItems(3) = ""
            If TabUSU.State = 1 Then _
               TabUSU.Close
            SQL = "select * from vwVendedor WITH (NOLOCK) "
            SQL = SQL & " where vendedor_id = " & TabTemp.Fields("vendedor_id").Value
            TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabUSU.EOF Then _
               item.SubItems(3) = Trim(TabUSU!DESCRICAO)
            If TabUSU.State = 1 Then _
               TabUSU.Close

            If Not IsNull(TabTemp.Fields("numero_caixa_cpu").Value) Then _
               item.SubItems(4) = TabTemp.Fields("numero_caixa_cpu").Value
         
            VALOR_TOTAL_N = VALOR_TOTAL_N + TabTemp.Fields("valor_total").Value
            txtTotalVenda.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
            txtTotalVenda.Refresh

            item.Checked = True

            SETA_GRID_ITENS TabTemp.Fields("PEDIDO_ID").Value
         End If
         TabTemp.MoveNext
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstPedido.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True

   HORA_FIM = Time

   MOSTRA_RODAPE "ESC - SAIR", "", "", "", Format((HORA_FIM - HORA_INI), "hh:mm:ss")

   lstProduto.Visible = True
   
   CALCULA_ULTRAPASSADO

   txtValorCompra.Text = "" & Format(VALOR_COMPRA_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA_SQL"
End Sub

Private Sub SETA_GRID_ITENS(NUMR_PEDIDO_ID As Long)
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = " select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, produto.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
   SQL = SQL & " PEDIDOITEM.STATUS, PEDIDOITEM.PRECO_CUSTO, PEDIDOITEM.PESO_ITEM, PRODUTO.DESCRICAO,"
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.SITUACAO, PRODUTO.TIPO_PROD, PRODUTO.PRECO_CUSTO AS precocusto, PRODUTO.PRECO_ATACADO,"
   SQL = SQL & " PRODUTO.PRECO_Venda, PRODUTO.PRODUTO_BALANCA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia, produto.conceder_producao,"
   SQL = SQL & " FAMILIAPRODUTO.PRODUCAO"
   SQL = SQL & " from PEDIDOITEM WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"

   SQL = SQL & " where pedido_id = " & NUMR_PEDIDO_ID
   SQL = SQL & " and pedidoitem.status <> 'C' "

   SQL = SQL & " order by pedido_id "

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Not TabProduto.EOF Then
      While Not TabProduto.EOF
         DoEvents

         CONT_N = CONT_N + 1

         Set ITEM2 = lstProduto.ListItems.Add(, "seq." & CONT_N, TabProduto.Fields("PEDIDO_ID").Value)

         ITEM2.SubItems(1) = "" & Trim(TabProduto.Fields("codg_produto").Value)
         ITEM2.SubItems(2) = "" & Trim(TabProduto.Fields("descricao").Value)
         ITEM2.SubItems(3) = "" & Format(TabProduto.Fields("QTD_PEDIDA").Value, strFormatacao3Digitos)
         ITEM2.SubItems(4) = "" & Format(TabProduto.Fields("valor_Item").Value, strFormatacao2Digitos)
         ITEM2.SubItems(5) = "" & Format(TabProduto.Fields("peso_Item").Value, strFormatacao3Digitos)
         ITEM2.SubItems(6) = "" & Format(TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value, strFormatacao2Digitos)

         VALOR_COMPRA_N = VALOR_COMPRA_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value

         'é item de produção
         'If Not IsNull(TabProduto.Fields("PRODUCAO").Value) Then
         '   If TabProduto.Fields("PRODUCAO").Value = True Then

         If Not IsNull(TabProduto.Fields("conceder_producao").Value) Then
            If TabProduto.Fields("conceder_producao").Value = True Then

               VALOR_PRODUCAO_N = VALOR_PRODUCAO_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value

               txtValorProducao.Text = Format(VALOR_PRODUCAO_N, strFormatacao2Digitos)
               txtValorProducao.Refresh

               ITEM_PRODUCAO_N = ITEM_PRODUCAO_N + 1
               txtQtdeItemProducao.Text = ITEM_PRODUCAO_N

               ITEM2.ForeColor = vbRed
               ITEM2.ListSubItems(1).ForeColor = vbRed
               ITEM2.ListSubItems(2).ForeColor = vbRed
               ITEM2.ListSubItems(3).ForeColor = vbRed
               ITEM2.ListSubItems(4).ForeColor = vbRed
               ITEM2.ListSubItems(5).ForeColor = vbRed
               ITEM2.ListSubItems(6).ForeColor = vbRed

               Else: VALOR_REVENDA_N = VALOR_REVENDA_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value
            End If
            Else: VALOR_REVENDA_N = VALOR_REVENDA_N + TabProduto.Fields("QTD_PEDIDA").Value * TabProduto.Fields("valor_Item").Value
         End If

         ITEM2.Checked = True

         Conta_Produto_N = Conta_Produto_N + 1
         txtQtdeProduto.Text = Conta_Produto_N

         TabProduto.MoveNext
      Wend
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ITENS"
End Sub

Sub BAIXA_PEDIDO()
'On Error GoTo ERRO_TRATA

   Msg = "Confirma baixa das compras do funcionário " & Trim(cmbFunc.Text) & " ?"
   PERGUNTA Msg, vbYesNo + 32, "Baixa de compras de Funcionário", "DEMO.HLP", 1000
   If RESPOSTA = vbNo Then _
      Exit Sub

   Dim i                   As Integer

   INDR_PRI = False

   If lstPedido.ListItems.Count > 0 Then
      For i = lstPedido.ListItems.Count To 1 Step -1
         If lstPedido.ListItems(i).Checked = True Then

            If Trim(lstPedido.ListItems(i).Text) <> "" Then

               INDR_PRI = True
               PEDIDO_ID_N = 0 & Trim(lstPedido.ListItems(i).Text)

               SQL = "update PEDIDO set "
               SQL = SQL & " status = 5 "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
'pegar pelo pai do financeiro
               SQL = "UPDATE ITEMLANCAMENTO SET "
               SQL = SQL & " Status = 'B'"
               SQL = SQL & ", DT_BAIXA = '" & Now & "'"
               SQL = SQL & ", CODG_USU_BAIXA = " & USUARIO_ID_N
               SQL = SQL & " Where numr_doc = " & PEDIDO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
      Next i
   End If

   If INDR_PRI = True Then
      MsgBox "Processo realizado com sucesso."
      LIMPA_TUDO
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "BAIXA_PEDIDO"
End Sub

Sub CALCULA_GRID_ITEM()
'On Error GoTo ERRO_TRATA

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "RETIRA_ITEM"
End Sub
