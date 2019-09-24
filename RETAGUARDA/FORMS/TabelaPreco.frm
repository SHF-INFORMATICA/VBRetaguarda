VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTabelaPreco 
   Caption         =   "Tabela de Preço Produtos"
   ClientHeight    =   6390
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TabelaPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   23
      Top             =   720
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "TabelaPreco.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line1(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblConta"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lstFormaPagto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lstFamilia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtValidade"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MSFlexGrid1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPreco"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdConsulta"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDescProd"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtProduto"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtDtCad"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtValorDig"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCodgTabela"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDescricao"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtCusto"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPerc"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkFamilia"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkVenda"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "chkCusto"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "chkPagto"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Command1"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Clonar"
      TabPicture(1)   =   "TabelaPreco.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Line1(0)"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(4)=   "lblMSG"
      Tab(1).Control(5)=   "Label13"
      Tab(1).Control(6)=   "Label14(0)"
      Tab(1).Control(7)=   "Label15"
      Tab(1).Control(8)=   "Label16"
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(10)=   "Label6"
      Tab(1).Control(11)=   "Label14(1)"
      Tab(1).Control(12)=   "lblProc"
      Tab(1).Control(13)=   "lstDestino"
      Tab(1).Control(14)=   "lstOrigem"
      Tab(1).Control(15)=   "txtDtValDestino"
      Tab(1).Control(16)=   "txtDtValOrigem"
      Tab(1).Control(17)=   "txtDescOrigem"
      Tab(1).Control(18)=   "txtOrigem"
      Tab(1).Control(19)=   "txtDtCadOrigem"
      Tab(1).Control(20)=   "txtDescDestino"
      Tab(1).Control(21)=   "txtDestino"
      Tab(1).Control(22)=   "txtDtCadDestino"
      Tab(1).Control(23)=   "CmdGravar"
      Tab(1).Control(24)=   "cmbDestino"
      Tab(1).Control(25)=   "cmbDestinoAUX"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Cadastradas"
      TabPicture(2)   =   "TabelaPreco.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstTabelaItem"
      Tab(2).Control(1)=   "lstTabela"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   9600
         TabIndex        =   57
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkPagto 
         Caption         =   "Todas Forma Pagto"
         Height          =   285
         Left            =   5520
         TabIndex        =   55
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkCusto 
         Caption         =   "AlteraPreçoCusto?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   51
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox chkVenda 
         Caption         =   "AlteraPreçoVenda?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   50
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox cmbDestinoAUX 
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
         Left            =   -73680
         TabIndex        =   46
         Top             =   3240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox cmbDestino 
         Height          =   405
         Left            =   -73680
         TabIndex        =   13
         Top             =   3240
         Width           =   5175
      End
      Begin VB.CheckBox chkFamilia 
         Caption         =   "Familia Produto"
         Height          =   285
         Left            =   200
         TabIndex        =   44
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.TextBox txtPerc 
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
         Left            =   9840
         MaxLength       =   12
         TabIndex        =   8
         ToolTipText     =   "Valor unitario de venda(varejo) do produto."
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "&Gravar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65280
         Picture         =   "TabelaPreco.frx":5C66
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Confirma os acessos para este usuario."
         Top             =   3240
         Width           =   1005
      End
      Begin VB.TextBox txtDtCadDestino 
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
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   -67440
         TabIndex        =   19
         ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDestino 
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
         Left            =   -73920
         TabIndex        =   14
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtDescDestino 
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
         Left            =   -71400
         TabIndex        =   18
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox txtDtCadOrigem 
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
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   -67440
         TabIndex        =   16
         ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtOrigem 
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
         Left            =   -73920
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDescOrigem 
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
         Left            =   -71400
         TabIndex        =   15
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtCusto 
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
         Left            =   5880
         MaxLength       =   12
         TabIndex        =   7
         ToolTipText     =   "Valor unitario de venda(varejo) do produto."
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   3600
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtCodgTabela 
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
         Left            =   1080
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   9240
         TabIndex        =   25
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDtCad 
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
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   7560
         TabIndex        =   2
         ToolTipText     =   "Informe Locação do Produto Com 6 Digitos (Alfanumerico)"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtProduto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtDescProd 
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
         Height          =   360
         Left            =   3555
         TabIndex        =   11
         Top             =   3240
         Width           =   7095
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3075
         Picture         =   "TabelaPreco.frx":6EBE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3240
         Width           =   405
      End
      Begin VB.TextBox txtPreco 
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
         Height          =   375
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   6
         ToolTipText     =   "Valor unitario de venda(varejo) do produto."
         Top             =   2640
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1695
         Left            =   45
         TabIndex        =   10
         Top             =   3840
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2990
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
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
      Begin MSMask.MaskEdBox txtValidade 
         Height          =   360
         Left            =   9600
         TabIndex        =   3
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtValOrigem 
         Height          =   360
         Left            =   -65400
         TabIndex        =   17
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtValDestino 
         Height          =   360
         Left            =   -67440
         TabIndex        =   20
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstFamilia 
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1720
         View            =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Forma Pagto."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   176
         EndProperty
      End
      Begin MSComctlLib.ListView lstFormaPagto 
         Height          =   975
         Left            =   5520
         TabIndex        =   5
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1720
         View            =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Forma Pagto."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   176
         EndProperty
      End
      Begin MSComctlLib.ListView lstTabela 
         Height          =   1215
         Left            =   -74925
         TabIndex        =   48
         Top             =   480
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   2143
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estabelecimento"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DtCad."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dt.Valid."
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lstTabelaItem 
         Height          =   3705
         Left            =   -74925
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   6535
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Produto"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FormaPagto"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "PreçoVenda"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "PreçoCusto"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lstOrigem 
         Height          =   1215
         Left            =   -74925
         TabIndex        =   52
         Top             =   1320
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   2143
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   255
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estabelecimento"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DtCad."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dt.Valid."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstDestino 
         Height          =   1335
         Left            =   -74925
         TabIndex        =   53
         Top             =   4200
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   2355
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483647
         BackColor       =   16777215
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
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estabelecimento"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DtCad."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dt.Valid."
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblConta 
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
         Left            =   4170
         TabIndex        =   56
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label lblProc 
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
         Left            =   -66120
         TabIndex        =   54
         Top             =   3840
         Width           =   675
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
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
         Index           =   1
         Left            =   -72480
         TabIndex        =   47
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label6 
         Caption         =   "Destino:"
         Height          =   285
         Left            =   -74760
         TabIndex        =   45
         Top             =   3240
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Index           =   4
         X1              =   0
         X2              =   12360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   12360
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "%Comis."
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
         Left            =   8940
         TabIndex        =   43
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Produto:"
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
         Left            =   600
         TabIndex        =   42
         Top             =   3240
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Index           =   2
         X1              =   0
         X2              =   12360
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Valida:"
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
         Left            =   -68220
         TabIndex        =   40
         Top             =   3720
         Width           =   675
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Cad.:"
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
         Left            =   -68295
         TabIndex        =   39
         Top             =   3240
         Width           =   750
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabela:"
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
         Left            =   -74745
         TabIndex        =   38
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
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
         Index           =   0
         Left            =   -72480
         TabIndex        =   37
         Top             =   3720
         Width           =   990
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Destino"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   36
         Top             =   2760
         Width           =   10530
      End
      Begin VB.Label lblMSG 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Origem"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   10530
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Valida:"
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
         Left            =   -66180
         TabIndex        =   34
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Cad.:"
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
         Left            =   -68295
         TabIndex        =   33
         Top             =   840
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Index           =   0
         X1              =   -75000
         X2              =   -62640
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabela:"
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
         Left            =   -74745
         TabIndex        =   32
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Pr.Custo="
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
         Left            =   4515
         TabIndex        =   31
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
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
         Left            =   2505
         TabIndex        =   30
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabela:"
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
         Left            =   255
         TabIndex        =   29
         Top             =   480
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   12360
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Cad.:"
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
         Left            =   6600
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Valida:"
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
         Left            =   8820
         TabIndex        =   27
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Pr.Venda="
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
         Left            =   345
         TabIndex        =   26
         Top             =   2640
         Width           =   1125
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   3069
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "mata"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vendedores"
            Key             =   "vend"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Custo/Venda"
            Key             =   "custovenda"
            ImageIndex      =   4
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
         Height          =   285
         Left            =   9000
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6960
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":78C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":8A5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":9AE9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":AA9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":BCBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":CDC9
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TabelaPreco.frx":E47F
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   10
      ResizeFonts     =   0   'False
      DesignWidth     =   10830
      DesignHeight    =   6390
   End
End
Attribute VB_Name = "frmTabelaPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABELAPRECO_ORIGEM_ID_N As Long

Private Sub Form_Load()

   Command1.Visible = False
   'If USUARIO_ID_N = 144 Then _
      Command1.Visible = True

   CARREGA_COMBO
   txtDtCad.Text = Date
   txtProduto.Enabled = False
   cmdConsulta.Enabled = False
   txtDescProd.Enabled = False

End Sub

Private Sub lstFormaPagto_Click()
   MOSTRA_TABELA_PRECO
   txtPreco.SetFocus
End Sub

Private Sub lstFamilia_Click()
   MOSTRA_TABELA_PRECO
   txtPreco.SetFocus
End Sub

Private Sub lstOrigem_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(lstOrigem.SelectedItem.ListSubItems.item(5).Text) Then
            Msg = "Confirma exclução da tabela ?"
            PERGUNTA Msg, vbYesNo + 32, "Exclusão Tabela Preço", "DEMO.HLP", 1000
            If RESPOSTA = vbNo Then _
               Exit Sub

            SQL = "delete from TABELAPRECOITEM where tabelapreco_id = " & lstOrigem.SelectedItem.ListSubItems.item(5).Text
            CONECTA_RETAGUARDA.Execute SQL
            SQL = "delete from TABELAPRECO where tabelapreco_id = " & lstOrigem.SelectedItem.ListSubItems.item(5).Text
            CONECTA_RETAGUARDA.Execute SQL
            MOSTRA_GRID_ORIGEM
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstOrigem_KeyDown"
End Sub

Private Sub lstTabela_DblClick()
   If Not IsNull(lstTabela.SelectedItem.Text) Then _
      MOSTRA_GRID_TABPRECO_ITEM lstTabela.SelectedItem.Text
End Sub

Private Sub lstOrigem_DblClick()
   If Not IsNull(lstOrigem.SelectedItem.Text) Then _
      txtDestino.Text = "" & lstOrigem.SelectedItem.Text
End Sub

Private Sub chkVenda_Click()
   txtPreco.SetFocus
End Sub

Private Sub chkCusto_Click()
   txtCusto.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

   Toolbar1.Buttons(4).Visible = True
   Toolbar1.Buttons(5).Visible = True
   Toolbar1.Buttons(6).Visible = True

   If SSTab1.Tab <> 0 Then
      Toolbar1.Buttons(4).Visible = False
      Toolbar1.Buttons(5).Visible = False
      Toolbar1.Buttons(6).Visible = False
   End If

   If SSTab1.Tab = 0 Then _
      txtCodgTabela.SetFocus
   If SSTab1.Tab = 1 Then
      txtOrigem.SetFocus
      MOSTRA_GRID_ORIGEM
   End If
   If SSTab1.Tab = 2 Then _
      MOSTRA_GRID_TABPRECO

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "custovenda"
         FORMULA_REL = ""
         SELECAO_FORMAPAGTO_A = ""
         INDR_PRI = True

         FORMULA_REL = "{TABELAPRECOitem.TABELAPRECO_ID} = " & TABELAPRECO_ORIGEM_ID_N

         If lstFormaPagto.ListItems.Count > 0 Then
            For i = lstFormaPagto.ListItems.Count To 1 Step -1
               If lstFormaPagto.ListItems(i).Checked = True Then
                  If INDR_PRI = True Then
                     SELECAO_FORMAPAGTO_A = lstFormaPagto.ListItems(i).SubItems(1)
                     Else: SELECAO_FORMAPAGTO_A = SELECAO_FORMAPAGTO_A & "," & lstFormaPagto.ListItems(i).SubItems(1)
                  End If
                  INDR_PRI = False
               End If
            Next i
            If Trim(SELECAO_FORMAPAGTO_A) <> "" Then _
               FORMULA_REL = FORMULA_REL & " and {TABELAPRECOitem.FORMAPAGTO_ID} in [" & Trim(SELECAO_FORMAPAGTO_A) & "]"
         End If

         SELECAO_FAMILIA_A = ""
         INDR_PRI = True

         If lstFamilia.ListItems.Count > 0 Then
            For i = lstFamilia.ListItems.Count To 1 Step -1
               If lstFamilia.ListItems(i).Checked = True Then
                  If INDR_PRI = True Then
                     SELECAO_FAMILIA_A = lstFamilia.ListItems(i).SubItems(1)
                     Else: SELECAO_FAMILIA_A = SELECAO_FAMILIA_A & "," & lstFamilia.ListItems(i).SubItems(1)
                  End If
                  INDR_PRI = False
               End If
            Next i
            If Trim(SELECAO_FAMILIA_A) <> "" Then _
               FORMULA_REL = FORMULA_REL & " and {PRODUTO.familiaproduto_ID} in [" & Trim(SELECAO_FAMILIA_A) & "]"
         End If

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_tab_CUSTO_VENDA.rpt"
         frmRELATORIO10.Show 1
      Case "vend"
         FORMULA_REL = ""
         SELECAO_FORMAPAGTO_A = ""
         INDR_PRI = True

         FORMULA_REL = "{TABELAPRECOitem.TABELAPRECO_ID} = " & TABELAPRECO_ORIGEM_ID_N

         If lstFormaPagto.ListItems.Count > 0 Then
            For i = lstFormaPagto.ListItems.Count To 1 Step -1
               If lstFormaPagto.ListItems(i).Checked = True Then
                  If INDR_PRI = True Then
                     SELECAO_FORMAPAGTO_A = lstFormaPagto.ListItems(i).SubItems(1)
                     Else: SELECAO_FORMAPAGTO_A = SELECAO_FORMAPAGTO_A & "," & lstFormaPagto.ListItems(i).SubItems(1)
                  End If
                  INDR_PRI = False
               End If
            Next i
            If Trim(SELECAO_FORMAPAGTO_A) <> "" Then _
               FORMULA_REL = FORMULA_REL & " and {TABELAPRECOitem.FORMAPAGTO_ID} in [" & Trim(SELECAO_FORMAPAGTO_A) & "]"
         End If

         SELECAO_FAMILIA_A = ""
         INDR_PRI = True

         If lstFamilia.ListItems.Count > 0 Then
            For i = lstFamilia.ListItems.Count To 1 Step -1
               If lstFamilia.ListItems(i).Checked = True Then
                  If INDR_PRI = True Then
                     SELECAO_FAMILIA_A = lstFamilia.ListItems(i).SubItems(1)
                     Else: SELECAO_FAMILIA_A = SELECAO_FAMILIA_A & "," & lstFamilia.ListItems(i).SubItems(1)
                  End If
                  INDR_PRI = False
               End If
            Next i
            If Trim(SELECAO_FAMILIA_A) <> "" Then _
               FORMULA_REL = FORMULA_REL & " and {PRODUTO.familiaproduto_ID} in [" & Trim(SELECAO_FAMILIA_A) & "]"
         End If

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_tab_preco.rpt"
         frmRELATORIO10.Show 1
      Case "consultar"
         CRITERIO_A = ""
         frmTabelaPrecoConsulta.Show 1
         If SSTab1.Tab = 0 Then
            txtCodgTabela.Text = CRITERIO_A
            txtCodgTabela.SetFocus
            Else
               txtOrigem.Text = CRITERIO_A
               txtOrigem.SetFocus
         End If
         CRITERIO_A = ""
      Case "gravar"
         GRAVA_CABECA
         GRAVA_ITEM

         lblConta.Caption = ""
         LIMPA_BODY
         MOSTRA_TABELA_PRECO

         txtCodgTabela.SetFocus
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
         txtCodgTabela.SetFocus
      Case "mata"
         TABELAPRECO_ORIGEM_ID_N = 0
         If Trim(txtCodgTabela.Text) <> "" Then
            Msg = "Confirma exclução da tabela ?"
            PERGUNTA Msg, vbYesNo + 32, "Exclusão Tabela Preço", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close
               SQL = "select tabelapreco_id from TABELAPRECO WITH (NOLOCK)"
               SQL = SQL & " where codg_tabela = '" & Trim(txtCodgTabela.Text) & "'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then _
                  If Not IsNull(TabTemp.Fields(0).Value) Then _
                     TABELAPRECO_ORIGEM_ID_N = TabTemp.Fields(0).Value
               If TabTemp.State = 1 Then _
                  TabTemp.Close
            
               SQL = "delete TABELAPRECOITEM "
               SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
               CONECTA_RETAGUARDA.Execute SQL
            
               SQL = "delete TABELAPRECO "
               SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
               CONECTA_RETAGUARDA.Execute SQL

               LIMPA_TUDO
            End If
         End If
         TABELAPRECO_ORIGEM_ID_N = 0
         txtCodgTabela.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbDESTINO_Click()
On Error Resume Next

   cmbDestinoAUX.ListIndex = cmbDestino.ListIndex

   txtTransf_ID.Text = MAX_ID("TRANSF_ID", "ESTOQUETRANSF", "", "", "", "")

   MOSTRA_GRID_DESTINO
End Sub

Private Sub chkFamilia_Click()
   If chkFamilia.Value = 0 Then
      txtProduto.Enabled = True
      cmdConsulta.Enabled = True
      lstFamilia.Enabled = False
      txtProduto.SetFocus
      Else
         txtProduto.Enabled = False
         cmdConsulta.Enabled = False
         lstFamilia.Enabled = True
         lstFamilia.SetFocus
   End If
End Sub

Private Sub chkPagto_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If lstFormaPagto.ListItems.Count > 0 Then
      For i = lstFormaPagto.ListItems.Count To 1 Step -1
         If chkPagto.Value = 1 Then
            lstFormaPagto.ListItems(i).Checked = True
            Else: lstFormaPagto.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "chkPagto_Click"
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    'OcultarControles

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_Click"
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   OcultarControles

   ExibirCelula

   txtCodgTabela.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 0)
   txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 11)
   'txtPesoItem.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 2)
   FraSeq.Enabled = True
   txtCodgTabela.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      'Case vbKeyF2      'Editar ao pressionar F2
      '   ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) <> "" Then _
            EXCLUIR_TABELA Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)), _
                           Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)), _
                          Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

   Select Case KeyAscii
      Case vbKeyReturn  ' Editar ao teclar ENTER
         KeyAscii = 0
         ExibirCelula
      Case vbKeyEscape  ' Cancelar ao pressionar ESC
         KeyAscii = 0
         AtribuiValorCelula
      Case 32 To 255    ' Editar ao pressinar qualquer tecla
         ExibirCelula
         With txtValorDig
            If .Visible Then
             .Text = Chr$(KeyAscii)
             .SelStart = Len(.Text) + 1
           End If
         End With
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyPress"
End Sub

Private Sub txtCodgTabela_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC-SAIR", "F7-Consulta Tabela", "Delete-Excluir Tabela", "F10-Gravar", ""

   txtCodgTabela.SelStart = 0
   txtCodgTabela.SelLength = Len(txtCodgTabela.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgTabela_GotFocus"
End Sub

Private Sub txtCodgTabela_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtCodgTabela.ForeColor = vbBlue
   txtDescricao.ForeColor = vbBlue

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCodgTabela.Text) = "" Then _
         txtCodgTabela.Text = MAX_ID("tabelapreco_id", "tabelapreco", "", "", "", "")

      txtDescricao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgTabela_KeyPress"
End Sub

Private Sub txtCodgTabela_LostFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_TABELA_PRECO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgTabela_LostFocus"
End Sub

Private Sub txtDESTINO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtDestino.ForeColor = vbBlue
   txtDescDestino.ForeColor = vbBlue

   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_TABELA_DESTINO
      txtDescDestino.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDESTINO_KeyPress"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then _
            txtProduto.Text = SQL3

         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

   If KeyAscii = 13 Then
      KeyAscii = 0
      LE_PRODUTO_LOCAL
      If Trim(txtDescProd.Text) <> "" Then _
         txtPreco.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then _
      txtProduto.Text = SQL3
   SQL3 = ""
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_CABECA
      GRAVA_ITEM

      lblConta.Caption = ""
      LIMPA_BODY
      MOSTRA_TABELA_PRECO

      txtPreco.SetFocus
      If chkFamilia.Value = 0 Then _
         txtProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPerc_KeyPress"
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCusto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPreco_KeyPress"
End Sub

Private Sub txtcusto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPerc.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcusto_KeyPress"
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtDescricao.Text) <> "" Then _
         txtValidade.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtvALIDADE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

   If KeyAscii = 13 Then
      KeyAscii = 0
      lstFamilia.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtValidade_LostFocus()
   txtValidade.PromptInclude = True
   If Not IsDate(txtValidade.Text) Then _
      txtValidade.Text = Date
End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   'txtValorDig.SelStart = 0
   'txtValorDig.SelLength = Len(txtValorDig)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_GotFocus"
End Sub

Private Sub txtValorDig_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         OcultarControles
         MSFlexGrid1.SetFocus
      Case vbKeyUp
         OcultarControles
         'move para a cima celula.
         With MSFlexGrid1
            If .Row > 1 Then
                .Row = .Row - 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
      Case vbKeyDown
         OcultarControles
         With MSFlexGrid1
             If .Row + 1 < .Rows Then
                .Row = .Row + 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyDown"
End Sub

Private Sub txtValorDig_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   ' ao pressionar ENTER aceitar a entrada de dados
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      If LastCol > 3 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      Dim QTDE_RETIDO_ESTORNO As Double

      QTDE_RETIDO_ESTORNO = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)

      AtribuiValorCelula
      'ProximaCelula
      OcultarControles

'==========ATUALIZAR GRID colunas
'3 = qtde
'4 = valor venda
'5 = desconto

      QTDE_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 3)
      VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
      VALOR_DESCONTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 5)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)
      'PRECO_CUSTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 8)
      CODG_PRODUTO_A = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 0))
      PRODUTO_ID_N = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 12))

      If QTDE_N > 0 And _
         VALOR_ITEM_N > 0 And _
         VALOR_DESCONTO_N >= 0 And _
         SEQ_ID_N > 0 Then

         MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N) - VALOR_DESCONTO_N), strFormatacao2Digitos)  'total item
         'lucro MSFlexGrid1.TextMatrix(LastRow, 9) = Format(((VALOR_ITEM_N - PRECO_CUSTO_N) * QTDE_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

         If INDR_ESTQ_NEGATIVO = False Then
            QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)
            If QTDE_ESTOQUE_N < 0 Then
               Beep
               MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
               txtQTDE.SetFocus
               Exit Sub
            End If
         End If

         If TabGridVaca.State = 1 Then _
            TabGridVaca.Close

         SQL = SQL & " select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, "
         SQL = SQL & " produto.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
         SQL = SQL & " PEDIDOITEM.PERC_DESC , PEDIDOITEM.Valor_Desconto, PEDIDOITEM.Status, PEDIDOITEM.PRECO_CUSTO"
         SQL = SQL & " from PEDIDO WITH (NOLOCK) "
         SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
         SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
         SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
         SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
         SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

         SQL = SQL & " where PEDIDO.pedido_id = " & txtPedido.Text
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and pedidoitem.status <> 'C' "

         TabGridVaca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

         If Not TabGridVaca.EOF Then
            SQL = "update PEDIDOITEM set "
            SQL = SQL & " QTD_PEDIDA = " & tpMOEDA(QTDE_N)
            SQL = SQL & ",Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",Valor_Desconto = " & tpMOEDA(VALOR_DESCONTO_N)
            SQL = SQL & ",peso_item = " & tpMOEDA(QTDE_N)

            SQL = SQL & " where pedido_id = " & TabGridVaca.Fields("pedido_id").Value
            SQL = SQL & " and seq_id = " & SEQ_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            QTDE_RETIDO_ESTORNO = 0
         End If

         If TabGridVaca.State = 1 Then _
            TabGridVaca.Close

         'MOSTRA_TOTAIS
      End If

      With MSFlexGrid1
         If .Row + 1 < .Rows Then
            .Row = .Row + 1
            '.Col = 0
            Else
               .Row = 1
               '.Col = 0
         End If
      End With
      txtValorDig.Text = ""

      MSFlexGrid1.SetFocus
      'LIMPA_BODY

      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            'ControlVisible = False
            Else
               If KeyAscii = 8 Or KeyAscii = 44 Then
                  Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyPress"
End Sub



Private Sub txtOrigem_GotFocus()
'On Error GoTo ERRO_TRATA

   txtOrigem.SelStart = 0
   txtOrigem.SelLength = Len(txtOrigem.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOrigem_GotFocus"
End Sub

Private Sub txtOrigem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtOrigem.ForeColor = vbBlue
   txtDescOrigem.ForeColor = vbBlue

   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_TABELA_ORIGEM
      txtDescOrigem.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOrigem_KeyPress"
End Sub

Private Sub CmdGravar_Click()
'On Error GoTo ERRO_TRATA

   Dim TabOrigemItem As New ADODB.Recordset

   lblProc.Caption = ""

   If Trim(cmbDestinoAUX.Text) = "" Then
      MsgBox "Selecione o estabelecimento de destino."
      cmbDestino.SetFocus
      Exit Sub
   End If

   If Trim(txtDestino.Text) = "" Then _
      Exit Sub

   If Trim(txtOrigem.Text) <> "" Then
      If Trim(txtDescDestino.Text) = "" Then _
         txtDescDestino.Text = "Não Informado"

      If Trim(txtDtCadDestino.Text) = "" Then _
         txtDtCadDestino.Text = Date

      If Trim(txtDtCadDestino.Text) = "" Then _
         txtDtCadDestino.Text = Date

      txtDtValDestino.PromptInclude = False
      If Trim(txtDtValDestino.Text) = "" Then _
         txtDtValDestino.Text = Date
      txtDtValDestino.PromptInclude = True

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from TABELAPRECO WITH (NOLOCK)"
      SQL = SQL & " where codg_tabela = '" & Trim(txtDestino.Text) & "'"
      SQL = SQL & " and estabelecimento_id = " & cmbDestinoAUX.Text
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         NUMR_ID_N = TabConsulta.Fields("tabelapreco_id").Value
         Msg = "Tabela de preço para o código = " & Trim(txtDestino.Text) & " já cadastrada com descrição : " & Trim(TabConsulta.Fields("descricao").Value) & ". Deseja sobrepor ?"
         PERGUNTA Msg, vbYesNo + 32, "Tabela Preço", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            SQL = "delete from TABELAPRECOITEM where tabelapreco_id = " & NUMR_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "update TABELAPRECO set "
            SQL = SQL & " DESCRICAO = '" & Trim(txtDescDestino.Text) & "'"
            SQL = SQL & ",DT_CAD = '" & DMA(txtDtCadDestino.Text) & "'"
            SQL = SQL & ",DT_validade = '" & DMA(txtDtValDestino.Text) & "'"

            SQL = SQL & " where tabelapreco_id = " & NUMR_ID_N
            SQL = SQL & " and estabelecimento_id = " & cmbDestinoAUX.Text
            CONECTA_RETAGUARDA.Execute SQL
         End If
         Else
            NUMR_ID_N = MAX_ID("tabelapreco_id", "tabelapreco", "", "", "", "")
            SQL = "insert into TABELAPRECO "
               SQL = SQL & "(TABELAPRECO_ID,ESTABELECIMENTO_ID,CODG_TABELA,DESCRICAO,DT_CAD,DT_VALIDADE)"
            SQL = SQL & " values("
               SQL = SQL & NUMR_ID_N                              'TABELAPRECO_ID
               SQL = SQL & "," & cmbDestinoAUX.Text               'ESTABELECIMENTO_ID
               SQL = SQL & ",'" & Trim(txtDestino.Text) & "'"     'CODG_TABELA
               SQL = SQL & ",'" & Trim(txtDescDestino.Text) & "'" 'Descricao
               SQL = SQL & ",'" & Now & "'"                 'DT_CAD
               SQL = SQL & ",'" & DMA(txtDtValDestino.Text) & "'" 'DT_VALIDADE
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
      End If

      If TabOrigemItem.State = 1 Then _
         TabOrigemItem.Close
      
      SQL = "select * from TABELAPRECOitem WITH (NOLOCK)"
      SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
      TabOrigemItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabOrigemItem.EOF
         SQL = "insert into TABELAPRECOITEM "
            SQL = SQL & "(TABELAPRECO_ID,TABELAPRECOITEM_ID,PRODUTO_ID,FORMAPAGTO_ID,VALOR_VENDA,VALOR_CUSTO,PERC_COMISSAO)"
         SQL = SQL & "values("
            SQL = SQL & NUMR_ID_N                                                'TABELAPRECO_ID
            SQL = SQL & "," & TabOrigemItem.Fields("TABELAPRECOITEM_ID").Value     'TABELAPRECOITEM_ID
            SQL = SQL & "," & TabOrigemItem.Fields("PRODUTO_ID").Value             'PRODUTO_ID
            SQL = SQL & "," & TabOrigemItem.Fields("formapagto_ID").Value          'FORMAPAGTO_ID
            SQL = SQL & "," & tpMOEDA(TabOrigemItem.Fields("VALOR_VENDA").Value)   'VALOR_VENDA
            SQL = SQL & "," & tpMOEDA(TabOrigemItem.Fields("VALOR_CUSTO").Value)   'VALOR_CUSTO
            SQL = SQL & "," & tpMOEDA(TabOrigemItem.Fields("PERC_COMISSAO").Value) 'PERC_COMISSAO
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL

         lblProc.Caption = TabOrigemItem.Fields("TABELAPRECOITEM_ID").Value
         lblProc.Refresh
         DoEvents

         TabOrigemItem.MoveNext
      Wend
      If TabOrigemItem.State = 1 Then _
         TabOrigemItem.Close
   End If

   MOSTRA_GRID_DESTINO
   MsgBox "Ok"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CmdGravar_Click"
End Sub

Private Sub ExibirCelula()
'On Error GoTo ERRO_TRATA

   Static OK As Boolean

   If MSFlexGrid1.Col >= 3 And MSFlexGrid1.Col <= 5 Then

      ' Se for celula fixa , sair
      If MSFlexGrid1.Col <= MSFlexGrid1.FixedCols - 1 Or MSFlexGrid1.Row <= MSFlexGrid1.FixedRows - 1 Then _
         Exit Sub
   
      If OK Then _
         Exit Sub

      OK = True

      OcultarControles

      LastRow = MSFlexGrid1.Row
      LastCol = MSFlexGrid1.Col

      Select Case LastCol
         Case Else
            txtValorDig.Move MSFlexGrid1.CellLeft - Screen.TwipsPerPixelX, MSFlexGrid1.CellTop + MSFlexGrid1.Top - Screen.TwipsPerPixelY, MSFlexGrid1.CellWidth + Screen.TwipsPerPixelX * 2, MSFlexGrid1.CellHeight + Screen.TwipsPerPixelY * 2
            txtValorDig.Text = MSFlexGrid1.Text

            If Len(MSFlexGrid1.Text) = 0 Then _
               If LastRow > 1 Then _
                  txtValorDig.Text = MSFlexGrid1.TextMatrix(LastRow - 1, LastCol)

            txtValorDig.Visible = True

            If txtValorDig.Visible Then
               txtValorDig.ZOrder
               txtValorDig.SetFocus
            End If
      End Select
   
      ControlVisible = True

      OK = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ExibirCelula"
End Sub

Private Sub ProximaCelula()
'On Error GoTo ERRO_TRATA

   If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      Else
         MSFlexGrid1.Col = 1
         If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
             MSFlexGrid1.Row = MSFlexGrid1.Row + 1
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ProximaCelula"
End Sub

Private Sub AtribuiValorCelula()
'On Error GoTo ERRO_TRATA

   Dim texto As String

   ' atribuir o texto anterior a celula
   Select Case LastCol
      Case 3 To 5
         texto = txtValorDig.Text

         If LastCol = 3 Then
            MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
            Else: MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao2Digitos)
         End If

         VALOR_VAREJO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
         VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, LastCol)

'&H80C0FF = LARANJA
'&H8000000F = CINZA
'&HFF& = VERMELHO
'vbBlack 0x0
'vbRed 0xFF
'vbGreen 0xFF00
'vbYellow 0xFFFF
'vbBlue 0xFF0000
'vbMagenta 0xFF00FF
'vbCyan 0xFFFF00
'vbWhite 0xFFFFFF

         If VALOR_ITEM_N < VALOR_VAREJO_N Then
            MSFlexGrid1.CellForeColor = vbRed
            MSFlexGrid1.CellFontBold = True
            MSFlexGrid1.CellBackColor = &H8000000F
            Else
               If VALOR_ITEM_N = VALOR_VAREJO_N Then
                  MSFlexGrid1.CellForeColor = vbBlack
                  MSFlexGrid1.CellFontBold = True
                  MSFlexGrid1.CellBackColor = vbCyan
                  Else
                     MSFlexGrid1.CellForeColor = vbBlue
                     MSFlexGrid1.CellFontBold = True
                     MSFlexGrid1.CellBackColor = vbWhite
               End If
         End If
      Case Else
         'texto = txtValorDig.Text
         'MSFlexGrid1.TextMatrix(LastRow, LastCol) = texto
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   'Ocultar o controle textbox
   txtValorDig.Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(8).Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Toolbar1.Buttons(9).Visible = True
      Toolbar1.Buttons(8).Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Sub MOSTRA_TABELA_PRECO()
'On Error GoTo ERRO_TRATA

   SELECAO_FORMAPAGTO_A = ""
   INDR_PRI = True

   If lstFormaPagto.ListItems.Count > 0 Then
      For i = lstFormaPagto.ListItems.Count To 1 Step -1
         If lstFormaPagto.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               SELECAO_FORMAPAGTO_A = lstFormaPagto.ListItems(i).SubItems(1)
               Else: SELECAO_FORMAPAGTO_A = SELECAO_FORMAPAGTO_A & "," & lstFormaPagto.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   SELECAO_FAMILIA_A = ""
   INDR_PRI = True

   If lstFamilia.ListItems.Count > 0 Then
      For i = lstFamilia.ListItems.Count To 1 Step -1
         If lstFamilia.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               SELECAO_FAMILIA_A = lstFamilia.ListItems(i).SubItems(1)
               Else: SELECAO_FAMILIA_A = SELECAO_FAMILIA_A & "," & lstFamilia.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   MSFlexGrid1.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PRODUTO.CODG_PRODUTO as Codg, PRODUTO.DESCRICAO AS Produto, "
   SQL = SQL & " TABELAPRECOITEM.FORMAPAGTO_ID as PagtoID, FORMAPAGTO.DESCRICAO AS DescPagto, "
   SQL = SQL & " TABELAPRECO.TABELAPRECO_ID, TABELAPRECO.CODG_TABELA, "
   SQL = SQL & " TABELAPRECO.DESCRICAO, TABELAPRECO.DT_CAD,"
   SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.TABELAPRECOITEM_ID,"
   SQL = SQL & " TABELAPRECOITEM.PRODUTO_ID, TABELAPRECOITEM.VALOR_VENDA as ValorVenda,"
   SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO as ValorCusto, TABELAPRECOITEM.PERC_COMISSAO as Comissao,"
   SQL = SQL & " Produto.Referencia"
   SQL = SQL & " from TABELAPRECO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN TABELAPRECOITEM WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " where TABELAPRECO.codg_tabela = '" & Trim(txtCodgTabela.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(SELECAO_FAMILIA_A) <> "" Then _
      SQL = SQL & " and familiaproduto_id in ( " & Trim(SELECAO_FAMILIA_A) & ")"

   If Trim(SELECAO_FORMAPAGTO_A) <> "" Then _
      SQL = SQL & " and TABELAPRECOITEM.formapagto_id in ( " & Trim(SELECAO_FORMAPAGTO_A) & ")"

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      TABELAPRECO_ORIGEM_ID_N = 0 & Trim(TabConsulta.Fields("TABELAPRECO_ID").Value)
      txtDescricao.Text = "" & Trim(TabConsulta.Fields("DESCRICAO").Value)
      txtDtCad.Text = "" & Trim(TabConsulta.Fields("dt_cad").Value)
      If Not IsNull(TabConsulta.Fields("dt_validade").Value) Then
         If IsDate(TabConsulta.Fields("dt_validade").Value) Then
            txtValidade.PromptInclude = False
               txtValidade.Text = TabConsulta.Fields("dt_validade").Value
            txtValidade.PromptInclude = True
         End If
      End If

      SETA_GRID

   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TABELA_PRECO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear

   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   ' define linhas fixas igual a uma e não usa colunas fixas
   MSFlexGrid1.Rows = 2
   'MSFlexGrid1.FixedRows = 3
   MSFlexGrid1.FixedCols = 0

   ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Cols = TabConsulta.Fields.Count

   ReDim largura_coluna(0 To TabConsulta.Fields.Count - 1)

   ' exibe os cabeçalhos das colunas
   For Coluna = 0 To TabConsulta.Fields.Count - 1
      MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabConsulta.Fields(Coluna).Name)
      largura_coluna(Coluna) = TextWidth(Trim(TabConsulta.Fields(Coluna).Name))
   Next Coluna

   ' exibe o valor de cada linha
   Linha = 1

   Do While Not TabConsulta.EOF
      MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

      For Coluna = 0 To TabConsulta.Fields.Count - 1
         'If Coluna = 3 Or Coluna = 7 Then
         If Coluna = 11 Or Coluna = 12 Then
            MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabConsulta.Fields(Coluna).Value, strFormatacao3Digitos)
            Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabConsulta.Fields(Coluna).Value)
         End If

'=========se o produto for de produção pintar linha
         'If INDR_PRI = True Then
         '   MSFlexGrid1.Row = Linha
         '   MSFlexGrid1.Col = Coluna
            'flex_tst.Text = "Bold Font"
            'flex_tst.CellFontBold = True
            'flex_tst.CellForeColor = vbRed
         '   MSFlexGrid1.CellForeColor = &H4000&   '&H40&
         'End If
'=========

         ' verifica o tamanho dos campos
         If Not IsNull(TabConsulta.Fields(Coluna).Value) Then _
            Largura_Campo = TextWidth(TabConsulta.Fields(Coluna).Value)

         If largura_coluna(Coluna) < Largura_Campo Then _
            largura_coluna(Coluna) = Largura_Campo

      Next Coluna

      TabConsulta.MoveNext
      Linha = Linha + 1
   Loop

   'define a largura das colunas do grid
   For Coluna = 0 To MSFlexGrid1.Cols - 1
      MSFlexGrid1.ColWidth(Coluna) = largura_coluna(Coluna) + 240
   Next Coluna

   MSFlexGrid1.ColWidth(0) = 0
   MSFlexGrid1.Refresh

   MSFlexGrid1.BackColor = vbWhite
   MSFlexGrid1.ForeColor = vbBlue

'CellFontName        - Define o nome da fonte para uma célula
'CellFontSize        - Define o tamanho da fonte para a célula
'CellFontBold        - Define se a fonte aparece em negrito.
'CellFontItalic      - Define se a fonte aparece em itálico.
'CellFontUnderline   - Define se a fonte aparece sublinhada.

'CODG_PRODUTO
      MSFlexGrid1.ColWidth(0) = 1500
      MSFlexGrid1.ColAlignment(0) = 0

'DescProd
      MSFlexGrid1.ColWidth(1) = 5000
      MSFlexGrid1.ColAlignment(1) = 0

'FORMAPAGTO_ID
      MSFlexGrid1.ColWidth(2) = 1000
      MSFlexGrid1.ColAlignment(2) = 7

'DescPagto
      MSFlexGrid1.ColWidth(3) = 4000
      MSFlexGrid1.ColAlignment(3) = 0

'TABELAPRECO_ID
      MSFlexGrid1.ColWidth(4) = 0
      MSFlexGrid1.ColAlignment(4) = 7

'DESCRICAO
      MSFlexGrid1.ColWidth(5) = 0
      MSFlexGrid1.ColAlignment(5) = 7

'DT_CAD
      MSFlexGrid1.ColWidth(6) = 0
      MSFlexGrid1.ColAlignment(6) = 7

'DT_VALIDADE
      MSFlexGrid1.ColWidth(7) = 1
      MSFlexGrid1.ColAlignment(7) = 0

'TABELAPRECOITEM_ID
      MSFlexGrid1.ColWidth(8) = 1
      MSFlexGrid1.ColAlignment(8) = 0

'PRODUTO_ID
      MSFlexGrid1.ColWidth(9) = 0
      MSFlexGrid1.ColAlignment(9) = 7

'Pedido_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(10) = 1
         Else: MSFlexGrid1.ColWidth(10) = 0
      End If
      MSFlexGrid1.ColAlignment(10) = 0

'VALOR_VENDA
      MSFlexGrid1.ColWidth(11) = 1500
      MSFlexGrid1.ColAlignment(11) = 7

'VALOR_CUSTO
      MSFlexGrid1.ColWidth(12) = 1500
      MSFlexGrid1.ColAlignment(12) = 7

'PERC_COMISSAO
      MSFlexGrid1.ColWidth(13) = 0
      MSFlexGrid1.ColAlignment(13) = 7

'Referencia
      MSFlexGrid1.ColWidth(14) = 0
      MSFlexGrid1.ColAlignment(14) = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub LE_PRODUTO_LOCAL()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   CODG_PRODUTO_A = Trim(txtProduto.Text)
   INDR_PROD_BALANCA = False

   'le por codigo de barras gravado no cadastro de produto
   CODIGO_BARRAS_A = "" & Trim(CODG_PRODUTO_A)
   QTDE_N = 0
   CRITERIO_A = ""

   If TabProduto.State = 1 Then _
      TabProduto.Close
   'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   SQL = "select count(produto_id) from PRODUTO  WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO_A = Trim(CODIGO_BARRAS_A)

            frmPEDIDOBARRAS.Show 1

            If Trim(CRITERIO_A) <> "" Then
               txtProduto.Text = Trim(CRITERIO_A)

               If TabProduto.State = 1 Then _
                  TabProduto.Close

               SQL = "select familiaproduto_id,produto_id,produto_balanca,codg_produto,SITUACAO,descricao,"
               SQL = SQL & " peso_liquido,PRECO_ATACADO,PRECO_VENDA,preco_custo,codg_ncm"
               SQL = SQL & " from PRODUTO WITH (NOLOCK)"
               SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
               SQL = SQL & " and situacao <> 'C' "
               TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabProduto.EOF Then
                  txtDescProd.Text = Trim(TabProduto.Fields("descricao").Value)
                  
               End If
               If TabProduto.State = 1 Then _
                  TabProduto.Close

               CRITERIO_A = ""
               Exit Sub
            End If
         End If
      End If
   End If

CRITERIO_A = ""

'produto de revenda
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select familiaproduto_id,produto_id,produto_balanca,codg_produto,SITUACAO,descricao,"
   SQL = SQL & " peso_liquido,PRECO_ATACADO,PRECO_VENDA,preco_custo,codg_ncm"
   SQL = SQL & " from PRODUTO  WITH (NOLOCK)"
   SQL = SQL & " where CODG_barra = '" & Trim(CODIGO_BARRAS_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      txtDescProd.Text = Trim(TabProduto.Fields("descricao").Value)
      txtCusto.SetFocus

      If TabProduto.State = 1 Then _
         TabProduto.Close
      Exit Sub
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   'le por codigo de barras ean 13 etiqueta balança
   CODIGO_BARRAS_A = "" & Trim(CODG_PRODUTO_A)
   If Len(CODIGO_BARRAS_A) = 13 Then
      '2 = produtos "in store" (sempre será 2)     1
      'C = código do produto (4,5 ou 6 dígitos)    2 a 8
      'T = total a pagar (sempre 6 dígitos)        9 a 13
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, CasaInicioCodgProdBarra_N, TamanhoCodgProdBarra_N))

      'If MULT_EMPRESA_B = True Then
         txtProduto.Text = "" & Int(Mid(CODIGO_BARRAS_A, 2, 4))

         If TabProduto.State = 1 Then _
            TabProduto.Close
      
         SQL = "select familiaproduto_id,produto_id,produto_balanca,codg_produto,SITUACAO,descricao,"
         SQL = SQL & " peso_liquido,PRECO_ATACADO,PRECO_VENDA,preco_custo,codg_ncm"
         SQL = SQL & " from PRODUTO  WITH (NOLOCK)"
         SQL = SQL & " where CODG_produto = '" & Trim(txtProduto.Text) & "'"
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            txtDescProd.Text = Trim(TabProduto.Fields("descricao").Value)
            txtCusto.SetFocus

            If TabProduto.State = 1 Then _
               TabProduto.Close

            Exit Sub
            Else: MsgBox "Verificar cadastro produto."
         End If
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   'LE POR CODIGO DE PRODUTO
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select familiaproduto_id,produto_id,produto_balanca,codg_produto,SITUACAO,descricao,"
   SQL = SQL & " peso_liquido,PRECO_ATACADO,PRECO_VENDA,preco_custo,codg_ncm"
   SQL = SQL & " from PRODUTO  WITH (NOLOCK) "
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(CODG_PRODUTO_A) & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      txtDescProd.Text = Trim(TabProduto.Fields("descricao").Value)
      txtCusto.SetFocus

      If TabProduto.State = 1 Then _
         TabProduto.Close

      Exit Sub
   End If

   MsgBox "Produto não cadastrado."
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_PRODUTO_LOCAL"
End Sub

Sub CARREGA_COMBO()
'On Error GoTo ERRO_TRATA

   Dim i

   lstFamilia.ListItems.Clear
   i = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select familiaproduto_id,descricao from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " order by DESCRICAO"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstFamilia.ListItems.Add(, "seq." & Trim(TabTemp.Fields("familiaproduto_id").Value), Trim(TabTemp.Fields("descricao").Value))
      item.SubItems(1) = "" & Trim(TabTemp.Fields("familiaproduto_id").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'If lstFamilia.ListItems.Count > 0 Then
   '   For i = lstFamilia.ListItems.Count To 1 Step -1
   '      lstFamilia.ListItems(i).Checked = True
   '   Next i
   'End If

   lstFormaPagto.ListItems.Clear
   i = 0

   SQL = "select formapagto_id,descricao from FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " where contab_balcao = 1"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstFormaPagto.ListItems.Add(, "seq." & TabTemp.Fields("formapagto_id").Value, Trim(TabTemp.Fields("descricao").Value))
      item.SubItems(1) = "" & TabTemp.Fields("formapagto_id").Value
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'If lstFormaPagto.ListItems.Count > 0 Then
   '   For i = lstFormaPagto.ListItems.Count To 1 Step -1
   '      lstFormaPagto.ListItems(i).Checked = True
   '   Next i
   'End If

   cmbDestino.Clear
   cmbDestinoAUX.Clear
   SQL = "select * from ESTABELECIMENTO "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      cmbDestino.AddItem Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("ESTABELECIMENTO_ID").Value)
      cmbDestinoAUX.AddItem TabTemp.Fields("ESTABELECIMENTO_ID").Value

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   TABELAPRECO_ORIGEM_ID_N = 0
   chkVenda.Value = 0
   chkCusto.Value = 0
   txtCodgTabela.Text = ""
   txtDescricao.Text = ""
   txtDtCad.Text = ""
   txtValidade.PromptInclude = False
   txtValidade.Text = ""
   txtValidade.PromptInclude = True
   MSFlexGrid1.Clear
   LIMPA_BODY

   txtOrigem.Text = ""
   txtDescOrigem.Text = ""
   txtDtCadOrigem.Text = ""
   txtDtValOrigem.PromptInclude = False
   txtDtValOrigem.Text = ""
   txtDtValOrigem.PromptInclude = True

   txtDestino.Text = ""
   txtDescDestino.Text = ""
   txtDtCadDestino.Text = ""
   txtDtValDestino.PromptInclude = False
   txtDtValDestino.Text = ""
   txtDtValDestino.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtCusto.Text = ""
   txtPerc.Text = ""
   txtProduto.Text = ""
   txtDescProd.Text = ""
   txtPreco.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub GRAVA_CABECA()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodgTabela.Text) = "" Then
      MsgBox "Informar codigo tabela !!!"
      txtCodgTabela.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Informar descrição tabela !!!"
      txtDescricao.SetFocus
      Exit Sub
   End If
   txtValidade.PromptInclude = True
   If Not IsDate(txtValidade.Text) Then _
      txtValidade.Text = Date

   TABELAPRECO_ORIGEM_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TABELAPRECO WITH (NOLOCK)"
   SQL = SQL & " where codg_tabela = '" & Trim(txtCodgTabela.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      TABELAPRECO_ORIGEM_ID_N = TabTemp.Fields("tabelapreco_id").Value
      SQL = "update TABELAPRECO set"
         SQL = SQL & " descricao = '" & Trim(txtDescricao.Text) & "'"
         SQL = SQL & " ,dt_validade = '" & DMA(txtValidade.Text) & "'"
      SQL = SQL & " where codg_tabela = '" & Trim(txtCodgTabela.Text) & "'"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      Else
         TABELAPRECO_ORIGEM_ID_N = MAX_ID("tabelapreco_id", "tabelapreco", "", "", "", "")

         SQL = "insert into TABELAPRECO "
            SQL = SQL & "(TABELAPRECO_ID,ESTABELECIMENTO_ID,CODG_TABELA,DESCRICAO,DT_CAD,DT_VALIDADE)"
         SQL = SQL & " values("
            SQL = SQL & TABELAPRECO_ORIGEM_ID_N                'TABELAPRECO_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N             'ESTABELECIMENTO_ID
            SQL = SQL & ",'" & Trim(txtCodgTabela.Text) & "'"  'CODG_TABELA
            SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'Descricao
            SQL = SQL & ",'" & Now & "'"                       'DT_CAD
            SQL = SQL & ",'" & DMA(txtValidade.Text) & "'"     'DT_VALIDADE
         SQL = SQL & ")"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA"
End Sub

Sub GRAVA_ITEM()
'On Error GoTo ERRO_TRATA

   TABELAPRECO_ORIGEM_ID_N = 0
   If Trim(txtCodgTabela.Text) = "" Then
      MsgBox "Informar codigo tabela !!!"
      txtCodgTabela.SetFocus
      Exit Sub
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select tabelapreco_id from TABELAPRECO WITH (NOLOCK)"
         SQL = SQL & " where codg_tabela = '" & Trim(txtCodgTabela.Text) & "'"
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Tabela de preço não encontrada."
            txtCodgTabela.SetFocus
            Exit Sub
            Else: TABELAPRECO_ORIGEM_ID_N = TabTemp.Fields(0).Value
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
   End If
   'If Trim(txtPreco.Text) = "" Then
   '   MsgBox "Informar valor !!!"
   '   txtPreco.SetFocus
   '   Exit Sub
   'End If
   If Trim(txtPerc.Text) = "" Then _
      txtPerc.Text = 0
   If Trim(txtCusto.Text) = "" Then _
      txtCusto.Text = 0

   SELECAO_FORMAPAGTO_A = ""
   INDR_PRI = True

   If lstFormaPagto.ListItems.Count > 0 Then
      For i = lstFormaPagto.ListItems.Count To 1 Step -1
lblConta.Caption = "" & i
DoEvents
         If lstFormaPagto.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               SELECAO_FORMAPAGTO_A = lstFormaPagto.ListItems(i).SubItems(1)
               Else: SELECAO_FORMAPAGTO_A = SELECAO_FORMAPAGTO_A & "," & lstFormaPagto.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If
   If Trim(SELECAO_FORMAPAGTO_A) = "" Then
      MsgBox "Informar selecione forma pagto !!!"
      lstFormaPagto.SetFocus
      Exit Sub
   End If
'============================ vai começar
   If chkFamilia.Value = 0 Then        'vai atualizar por produto
      If Trim(txtProduto.Text) <> "" Then
         PRODUTO_ID_N = 0
   
         If TabTemp.State = 1 Then _
            TabTemp.Close
   
         SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where codg_produto = '" & Trim(txtProduto.Text) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            PRODUTO_ID_N = TabTemp.Fields(0).Value
            Else
               If TabTemp.State = 1 Then _
                  TabTemp.Close
   
               MsgBox "Produto não cadastrado !!!"
               txtProduto.SetFocus
               Exit Sub
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

'passou dos teste vai atualizar por produto
'vai ler as forma pagto e gravar todas para um produto somente

         If lstFormaPagto.ListItems.Count > 0 Then
            For i = lstFormaPagto.ListItems.Count To 1 Step -1
lblConta.Caption = "" & PRODUTO_ID_N
DoEvents
               If lstFormaPagto.ListItems(i).Checked = True Then
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
         
                  SQL = "select tabelapreco_id,tabelaprecoitem_id from TABELAPRECOITEM WITH (NOLOCK)"
                  SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
                  SQL = SQL & " and formapagto_id = " & Trim(lstFormaPagto.ListItems(i).SubItems(1))
                  SQL = SQL & " and produto_id = " & PRODUTO_ID_N
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     SQL = "update TABELAPRECOITEM set"
                        SQL = SQL & " perc_comissao = " & tpMOEDA(txtPerc.Text)           'PERC_COMISSAO

                        If chkVenda.Value = 1 Then _
                           SQL = SQL & ", valor_venda = " & tpMOEDA(txtPreco.Text)        'VALOR_VENDA
                        If chkCusto.Value = 1 Then _
                           SQL = SQL & ", VALOR_CUSTO = " & tpMOEDA(txtCusto.Text)        'VALOR_CUSTO

                     SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
                     SQL = SQL & " and formapagto_id = " & Trim(lstFormaPagto.ListItems(i).SubItems(1))
                     SQL = SQL & " and produto_id = " & PRODUTO_ID_N
                     Else
                        SQL = "insert into TABELAPRECOITEM "
                           SQL = SQL & "(TABELAPRECO_ID,TABELAPRECOITEM_ID,PRODUTO_ID,FORMAPAGTO_ID,VALOR_VENDA,VALOR_CUSTO,PERC_COMISSAO)"
                        SQL = SQL & "values("
                           SQL = SQL & TABELAPRECO_ORIGEM_ID_N                                                               'TABELAPRECO_ID
                           SQL = SQL & "," & MAX_ID("TABELAPRECOITEM_ID", "tabelaprecoitem", "", "", "", "")   'TABELAPRECOITEM_ID
                           SQL = SQL & "," & PRODUTO_ID_N                                                      'PRODUTO_ID
                           SQL = SQL & "," & Trim(lstFormaPagto.ListItems(i).SubItems(1))                      'FORMAPAGTO_ID
                           SQL = SQL & "," & tpMOEDA(txtPreco.Text)                                            'VALOR_VENDA
                           SQL = SQL & "," & tpMOEDA(txtCusto.Text)                                            'VALOR_CUSTO
                           SQL = SQL & "," & tpMOEDA(txtPerc.Text)                                             'PERC_COMISSAO
                        SQL = SQL & ")"
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
               
                  CONECTA_RETAGUARDA.Execute SQL
               End If
            Next i
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      Else                             'vai atualizar por familia
'======================================
         SELECAO_FAMILIA_A = ""
         INDR_PRI = True

         If lstFamilia.ListItems.Count > 0 Then
            For i = lstFamilia.ListItems.Count To 1 Step -1
lblConta.Caption = "" & i
DoEvents
               If lstFamilia.ListItems(i).Checked = True Then
                  If INDR_PRI = True Then
                     SELECAO_FAMILIA_A = lstFamilia.ListItems(i).SubItems(1)
                     Else: SELECAO_FAMILIA_A = SELECAO_FAMILIA_A & "," & lstFamilia.ListItems(i).SubItems(1)
                  End If
                  INDR_PRI = False
               End If
            Next i
         End If
         If Trim(SELECAO_FAMILIA_A) = "" Then
            MsgBox "Informar selecione Familia Produto !!!"
            lstFormaPagto.SetFocus
            Exit Sub
         End If

         'LENDO FAMILIAS SELECIONADAS
         If lstFamilia.ListItems.Count > 0 Then
            For i_Familia = lstFamilia.ListItems.Count To 1 Step -1
               If lstFamilia.ListItems(i_Familia).Checked = True Then
                  If Trim(lstFamilia.ListItems(i_Familia).SubItems(1)) <> "" Then

                     'LENDO PAGTO SELECIONADAS
                     If lstFormaPagto.ListItems.Count > 0 Then
                        For i_Pagto = lstFormaPagto.ListItems.Count To 1 Step -1
                           If lstFormaPagto.ListItems(i_Pagto).Checked = True Then
                              If Trim(lstFormaPagto.ListItems(i_Pagto).SubItems(1)) <> "" Then
'======================================================================================================
                                 If TabProduto.State = 1 Then _
                                    TabProduto.Close

                                 SQL = "select produto_id,familiaproduto_id,codg_produto from PRODUTO WITH (NOLOCK)"
                                 SQL = SQL & " where familiaproduto_id = " & Trim(lstFamilia.ListItems(i_Familia).SubItems(1))
                                 TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                                 While Not TabProduto.EOF
lblConta.Caption = "" & TabProduto.Fields("codg_produto").Value
DoEvents
                                    If TabTemp.State = 1 Then _
                                       TabTemp.Close

                                    SQL = "select tabelapreco_id,tabelaprecoitem_id from TABELAPRECOITEM WITH (NOLOCK)"
                                    SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
                                    SQL = SQL & " and formapagto_id = " & Trim(lstFormaPagto.ListItems(i_Pagto).SubItems(1))
                                    SQL = SQL & " and produto_id = " & TabProduto.Fields("PRODUTO_ID").Value
                                    TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                                    If Not TabTemp.EOF Then
                                       SQL = "update TABELAPRECOITEM set"
                                          SQL = SQL & " formapagto_id = " & Trim(lstFormaPagto.ListItems(i_Pagto).SubItems(1)) 'FORMAPAGTO_ID

                                       If chkVenda.Value = 1 Then _
                                          SQL = SQL & ", valor_venda = " & tpMOEDA(txtPreco.Text)        'VALOR_VENDA
                                       If chkCusto.Value = 1 Then _
                                          SQL = SQL & ", VALOR_CUSTO = " & tpMOEDA(txtCusto.Text)        'VALOR_CUSTO

                                          SQL = SQL & ", perc_comissao = " & tpMOEDA(txtPerc.Text)                       'PERC_COMISSAO
                                       SQL = SQL & " where tabelapreco_id = " & TABELAPRECO_ORIGEM_ID_N
                                       SQL = SQL & " and formapagto_id = " & Trim(lstFormaPagto.ListItems(i_Pagto).SubItems(1))
                                       SQL = SQL & " and produto_id = " & TabProduto.Fields("PRODUTO_ID").Value
                                       Else
                                          SQL = "insert into TABELAPRECOITEM "
                                             SQL = SQL & "(TABELAPRECO_ID,TABELAPRECOITEM_ID,PRODUTO_ID,FORMAPAGTO_ID,VALOR_VENDA,VALOR_CUSTO,PERC_COMISSAO)"
                                          SQL = SQL & "values("
                                             SQL = SQL & TABELAPRECO_ORIGEM_ID_N                                                               'TABELAPRECO_ID
                                             SQL = SQL & "," & MAX_ID("TABELAPRECOITEM_ID", "tabelaprecoitem", "", "", "", "")   'TABELAPRECOITEM_ID
                                             SQL = SQL & "," & TabProduto.Fields("PRODUTO_ID").Value                             'PRODUTO_ID
                                             SQL = SQL & "," & Trim(lstFormaPagto.ListItems(i_Pagto).SubItems(1))                      'FORMAPAGTO_ID
                                             SQL = SQL & "," & tpMOEDA(txtPreco.Text)                                            'VALOR_VENDA
                                             SQL = SQL & "," & tpMOEDA(txtCusto.Text)                                            'VALOR_CUSTO
                                             SQL = SQL & "," & tpMOEDA(txtPerc.Text)                                             'PERC_COMISSAO
                                          SQL = SQL & ")"
                                    End If
                                    If TabTemp.State = 1 Then _
                                       TabTemp.Close

                                    CONECTA_RETAGUARDA.Execute SQL

                                    TabProduto.MoveNext
                                 Wend
                                 If TabProduto.State = 1 Then _
                                    TabProduto.Close
'======================================================================================================
                              End If
                           End If
                        Next i_Pagto
                     End If

                  End If
               End If
            Next i_Familia
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITEM"
End Sub

Sub EXCLUIR_TABELA(Prod_Id As Long, Tab_Id As Long, Forma_Id As Long)
'On Error GoTo ERRO_TRATA

   If Prod_Id > 0 And Tab_Id > 0 And Forma_Id > 0 Then
      Msg = "Confirma Exclusão ?"
      PERGUNTA Msg, vbYesNo + 32, "Tabela Preço", "DEMO.HLP", 1000
      If RESPOSTA = vbNo Then _
         Exit Sub

      SQL = "delete TABELAPRECOITEM "
      SQL = SQL & " where tabelapreco_id = " & Tab_Id
      SQL = SQL & " and formapagto_id = " & Forma_Id
      SQL = SQL & " and produto_id = " & Prod_Id
      CONECTA_RETAGUARDA.Execute SQL
      MOSTRA_TABELA_PRECO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_TABELA"
End Sub

Sub MOSTRA_TABELA_ORIGEM()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PRODUTO.CODG_PRODUTO as Codg, PRODUTO.DESCRICAO AS Produto, "
   SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescPagto, TABELAPRECOITEM.FORMAPAGTO_ID, "
   SQL = SQL & " TABELAPRECO.TABELAPRECO_ID, TABELAPRECO.CODG_TABELA, "
   SQL = SQL & " TABELAPRECO.DESCRICAO, TABELAPRECO.DT_CAD,"
   SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.TABELAPRECOITEM_ID,"
   SQL = SQL & " TABELAPRECOITEM.PRODUTO_ID, TABELAPRECOITEM.VALOR_VENDA as ValorVenda,"
   SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO, TABELAPRECOITEM.PERC_COMISSAO as Comissao,"
   SQL = SQL & " Produto.Referencia"
   SQL = SQL & " from TABELAPRECO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN TABELAPRECOITEM WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
   SQL = SQL & " where TABELAPRECO.codg_tabela = '" & Trim(txtOrigem.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      TABELAPRECO_ORIGEM_ID_N = TabConsulta.Fields("TABELAPRECO_ID").Value
      If Not IsNull(TabConsulta.Fields("dt_validade").Value) Then
         If IsDate(TabConsulta.Fields("dt_validade").Value) Then
            txtDtValOrigem.PromptInclude = False
               txtDtValOrigem.Text = TabConsulta.Fields("dt_validade").Value
            txtDtValOrigem.PromptInclude = True
         End If
      End If

      txtDestino.Text = MAX_ID("tabelapreco_id", "tabelapreco", "", "", "", "")
      txtDescOrigem.Text = "" & Trim(TabConsulta.Fields("DESCRICAO").Value)
      txtDtCadOrigem.Text = "" & Trim(TabConsulta.Fields("dt_cad").Value)
      If Not IsNull(TabConsulta.Fields("dt_validade").Value) Then
         If IsDate(TabConsulta.Fields("dt_validade").Value) Then
            txtDtValDestino.PromptInclude = False
               txtDtValDestino.Text = TabConsulta.Fields("dt_validade").Value
            txtDtValDestino.PromptInclude = True
         End If
      End If
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TABELA_ORIGEM"
End Sub

Sub MOSTRA_TABELA_DESTINO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PRODUTO.CODG_PRODUTO as Codg, PRODUTO.DESCRICAO AS Produto, "
   SQL = SQL & " FORMAPAGTO.DESCRICAO AS DescPagto, TABELAPRECOITEM.FORMAPAGTO_ID, "
   SQL = SQL & " TABELAPRECO.TABELAPRECO_ID, TABELAPRECO.CODG_TABELA, "
   SQL = SQL & " TABELAPRECO.DESCRICAO, TABELAPRECO.DT_CAD,"
   SQL = SQL & " TABELAPRECO.DT_VALIDADE, TABELAPRECOITEM.TABELAPRECOITEM_ID,"
   SQL = SQL & " TABELAPRECOITEM.PRODUTO_ID, TABELAPRECOITEM.VALOR_VENDA as ValorVenda,"
   SQL = SQL & " TABELAPRECOITEM.VALOR_CUSTO, TABELAPRECOITEM.PERC_COMISSAO as Comissao,"
   SQL = SQL & " Produto.Referencia"
   SQL = SQL & " from TABELAPRECO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN TABELAPRECOITEM WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON TABELAPRECOITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
   SQL = SQL & " where TABELAPRECO.codg_tabela = '" & Trim(txtDestino.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & cmbDestinoAUX.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      If Not IsNull(TabConsulta.Fields("dt_validade").Value) Then
         If IsDate(TabConsulta.Fields("dt_validade").Value) Then
            txtDtValDestino.PromptInclude = False
               txtDtValDestino.Text = TabConsulta.Fields("dt_validade").Value
            txtDtValDestino.PromptInclude = True
         End If
      End If

      txtDescDestino.Text = "" & Trim(TabConsulta.Fields("DESCRICAO").Value)
      txtDtCadDestino.Text = "" & Trim(TabConsulta.Fields("dt_cad").Value)
      If Not IsNull(TabConsulta.Fields("dt_validade").Value) Then
         If IsDate(TabConsulta.Fields("dt_validade").Value) Then
            txtDtValDestino.PromptInclude = False
               txtDtValDestino.Text = TabConsulta.Fields("dt_validade").Value
            txtDtValDestino.PromptInclude = True
         End If
      End If
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TABELA_DESTINO"
End Sub

Sub MOSTRA_GRID_TABPRECO()
'On Error GoTo ERRO_TRATA

   lstTabela.ListItems.Clear
   lstTabelaItem.ListItems.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from TABELAPRECO"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      Set item = lstTabela.ListItems.Add(, "seq." & TabConsulta.Fields("tabelapreco_ID").Value, TabConsulta.Fields("codg_tabela").Value)
      item.SubItems(1) = "" & Trim(TabConsulta.Fields("descricao").Value)
      item.SubItems(2) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estabelecimento_id").Value)
      item.SubItems(3) = "" & DMA(TabConsulta.Fields("dt_cad").Value)
      item.SubItems(4) = "" & DMA(TabConsulta.Fields("dt_validade").Value)

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_GRID_TABPRECO"
End Sub

Sub MOSTRA_GRID_TABPRECO_ITEM(TABELA_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Msg = "Confirma mostrar itens tabela de preço ? Essa rotina pode demorar alguns minutos."
   PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SqL2 = Me.Caption
      Me.Caption = "Aguarde ..."
      lstTabelaItem.ListItems.Clear
      lstTabelaItem.Visible = False
      CONT_N = 0

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from TABELAPRECOITEM"
      SQL = SQL & " where tabelapreco_id = " & TABELA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF

         Set item = lstTabelaItem.ListItems.Add(, "seq." & CONT_N, TRAZ_DESCRICAO_PRODUTO(TabConsulta.Fields("produto_id").Value, ""))
         item.SubItems(1) = "" & TRAZ_DESCRICAO_FORMAPAGTO(TabConsulta.Fields("formapagto_id").Value)
         item.SubItems(2) = "" & TabConsulta.Fields("valor_venda").Value
         item.SubItems(3) = "" & TabConsulta.Fields("valor_custo").Value

         TabConsulta.MoveNext
         CONT_N = CONT_N + 1
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      Me.Caption = SqL2
   End If
   lstTabelaItem.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_GRID_TABPRECO_ITEM"
End Sub

Sub MOSTRA_GRID_ORIGEM()
'On Error GoTo ERRO_TRATA

   lstOrigem.ListItems.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from TABELAPRECO"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      Set item = lstOrigem.ListItems.Add(, "seq." & TabConsulta.Fields("tabelapreco_ID").Value, TabConsulta.Fields("codg_tabela").Value)
      item.SubItems(1) = "" & Trim(TabConsulta.Fields("descricao").Value)
      item.SubItems(2) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estabelecimento_id").Value)
      item.SubItems(3) = "" & DMA(TabConsulta.Fields("dt_cad").Value)
      item.SubItems(4) = "" & DMA(TabConsulta.Fields("dt_validade").Value)
      item.SubItems(5) = "" & TabConsulta.Fields("tabelapreco_ID").Value

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_GRID_ORIGEM"
End Sub

Sub MOSTRA_GRID_DESTINO()
'On Error GoTo ERRO_TRATA

   lstDestino.ListItems.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from TABELAPRECO"
   SQL = SQL & " where estabelecimento_id = " & cmbDestinoAUX.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      Set item = lstDestino.ListItems.Add(, "seq." & TabConsulta.Fields("tabelapreco_ID").Value, TabConsulta.Fields("codg_tabela").Value)
      item.SubItems(1) = "" & Trim(TabConsulta.Fields("descricao").Value)
      item.SubItems(2) = "" & TRAZ_ESTABELECIMENTO(TabConsulta.Fields("estabelecimento_id").Value)
      item.SubItems(3) = "" & DMA(TabConsulta.Fields("dt_cad").Value)
      item.SubItems(4) = "" & DMA(TabConsulta.Fields("dt_validade").Value)

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_GRID_destino"
End Sub

Private Sub Command1_Click()

   Dim TabFDP  As New ADODB.Recordset

   GRAVA_CABECA

   If TabFDP.State = 1 Then _
      TabFDP.Close

   SQL = "select produto_id,codg_produto,preco_venda,preco_custo,descricao from PRODUTO "
   SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
   TabFDP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabFDP.EOF

      txtProduto.Text = "" & TabFDP.Fields("codg_produto").Value
      txtPreco.Text = "" & TabFDP.Fields("preco_venda").Value
      txtCusto.Text = "" & TabFDP.Fields("preco_custo").Value
      txtDescProd.Text = "" & TabFDP.Fields("descricao").Value
      txtPerc.Text = "" & 0
      DoEvents

      GRAVA_ITEM

      TabFDP.MoveNext
   Wend
   If TabFDP.State = 1 Then _
      TabFDP.Close
lblConta.Caption = ""
   LIMPA_BODY
   MOSTRA_TABELA_PRECO

MsgBox "fim"
End Sub
