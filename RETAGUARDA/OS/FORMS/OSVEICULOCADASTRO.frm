VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmOSVEICULOCADASTRO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Veículo"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "OSVEICULOCADASTRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "OSVEICULOCADASTRO.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCpf"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCNPJCPF"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDtIni"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbComb"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbCor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMotor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDescricao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbTIPO"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMODELO"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtANO"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtNome"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCHASSI"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbCombAUX"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbCorAUX"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbTipoAUX"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmbMarca"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmbMarcaAUX"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtPlaca"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "C&onsulta"
      TabPicture(1)   =   "OSVEICULOCADASTRO.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label14"
      Tab(1).Control(3)=   "MaskEdBox1"
      Tab(1).Control(4)=   "LISTACHASSI"
      Tab(1).Control(5)=   "Text1"
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(7)=   "txtPlaca2"
      Tab(1).ControlCount=   8
      Begin PVMaskEditLib.PVMaskEdit txtPlaca2 
         Height          =   375
         Left            =   -74040
         TabIndex        =   11
         Top             =   480
         Width           =   1215
         _Version        =   524288
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "@@@-####"
      End
      Begin PVMaskEditLib.PVMaskEdit txtPlaca 
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _Version        =   524288
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "@@@-####"
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Left            =   -71880
         MaxLength       =   100
         TabIndex        =   37
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox Text1 
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
         Left            =   -70800
         MaxLength       =   50
         TabIndex        =   35
         Top             =   480
         Width           =   4335
      End
      Begin VB.ComboBox cmbMarcaAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbMarca 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cmbTipoAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbCorAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbCombAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCHASSI 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
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
         Left            =   3120
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3840
         Width           =   5295
      End
      Begin VB.TextBox txtANO 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   4
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtMODELO 
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
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   5
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox cmbTIPO 
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
         Left            =   6120
         TabIndex        =   8
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   6135
      End
      Begin VB.TextBox txtMotor 
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
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1920
         Width           =   6135
      End
      Begin VB.ComboBox cmbCor 
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
         Left            =   6120
         TabIndex        =   6
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox cmbComb 
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
         Left            =   2280
         TabIndex        =   7
         Top             =   2880
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   405
         Left            =   7080
         TabIndex        =   15
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   16777215
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   405
         Left            =   960
         TabIndex        =   10
         Top             =   3840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   18
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
      Begin MSComctlLib.ListView LISTACHASSI 
         Height          =   2745
         Left            =   -74955
         TabIndex        =   27
         Top             =   1440
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   4842
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777152
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Chassi"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Placa"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "ANO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "MODELO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "TIPO"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   405
         Left            =   -74040
         TabIndex        =   38
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   18
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
      Begin VB.Label Label14 
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
         Height          =   240
         Left            =   -74880
         TabIndex        =   39
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição/Modelo:"
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
         Left            =   -72840
         TabIndex        =   36
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
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
         Left            =   1095
         TabIndex        =   32
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chassi:"
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
         Left            =   1440
         TabIndex        =   26
         Top             =   1440
         Width           =   735
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
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
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
         Left            =   1680
         TabIndex        =   24
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Modelo:"
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
         Left            =   3480
         TabIndex        =   23
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Cadastro:"
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
         Left            =   5280
         TabIndex        =   22
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Veículo:"
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
         Left            =   4680
         TabIndex        =   21
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
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
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição/Modelo:"
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
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motor:"
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
         Left            =   1560
         TabIndex        =   18
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cor:"
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
         Left            =   5400
         TabIndex        =   17
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Combustível:"
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
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
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
            Picture         =   "OSVEICULOCADASTRO.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":0498
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":07B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":0C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":105C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":137C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSVEICULOCADASTRO.frx":17D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "importa"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   8700
      DesignHeight    =   5385
   End
   Begin MSComctlLib.StatusBar barVeiculo 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   5010
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "OSVEICULOCADASTRO.frx":1AF0
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
Attribute VB_Name = "frmOSVEICULOCADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim COR_ID_N         As Long
   Dim MARCA_ID_N       As Long
   Dim TIPO_EQP_ID_N    As Long
   Dim ANO_N            As Long
   Dim MODELO_ID_N      As Long
   Dim VEICULO_ID_N     As Long
   Dim EQUIPAMENTO_ID_N As Long
   Dim COMBUSTIVEL_ID_N As Long
   Dim ANO_ID_N         As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_MEGASIM NOME_BANCO_DADOS

   CARREGA_DESCRITORES

   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

CRIA_CONSULTAS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If conecta_retaguarda.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   If SSTab1.Tab = 0 Then _
      txtPLACA.SetFocus

   If SSTab1.Tab = 2 Then
      txtPlaca2.SetFocus

      SETA_GRID_VEICULO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSTab1_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "matar"
         If Trim(txtPLACA.Text) <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from VEICULO "
            SQL = SQL & "where placa = '" & Trim(txtPLACA.Text) & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If TabAUX.State = 1 Then _
                  TabAUX.Close

               SQL = "select * from OS "
               SQL = SQL & "where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
               TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabAUX.EOF Then
                  If TabAUX.State = 1 Then _
                     TabAUX.Close

                  If TabTemp.State = 1 Then _
                     TabTemp.Close
                  
                  MsgBox "Impossível excluir, veículo possue movimentação na oficina."
                  Exit Sub
               End If
               If TabAUX.State = 1 Then _
                  TabAUX.Close

               Msg = "Confirma Exclusão do VEICULO ?"
               PERGUNTA Msg, vbYesNo + 32, "", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  SQL = "delete from VEICULO "
                  SQL = SQL & " where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
                  CONECTA_RETAGUARDA.Execute SQL

                  SQL = "delete from EQUIPAMENTO "
                  SQL = SQL & " where equipamento_id = " & TabTemp.Fields("equipamento_id").Value
                  CONECTA_RETAGUARDA.Execute SQL

                  LIMPA_VEICULO
                  txtPLACA.SetFocus
               End If
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
      Case "gravar"
         GRAVA_VEICULO
         txtPLACA.SetFocus
      Case "limpar"
         LIMPA_VEICULO
      Case "imprimir"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F4 - Cadastra Cliente", "F7 - Consulta Clientes", "Informe proprietário do Veículo", ""

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Mask = "##############"
      If CPF_N <> "" Then
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = CPF_N
         CPF_N = ""
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCLIENTE.Show 1
      Case vbKeyF7
         frmDISPLAYCLIENTE.Show 1
         If Trim(CPF_N) <> "" Then
            txtCNPJCPF.Mask = "##############"
            txtCNPJCPF.Text = CPF_N
         End If
         CPF_N = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      ENDERECO_A = ""
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         'MsgBox "Informe CNPJ/CPF corretamente"
         txtCNPJCPF.Text = "99999999999"
         Else
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
      End If
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         CRITERIO = txtCNPJCPF.Text
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then
               txtCNPJCPF.Mask = "###.###.###-##"
               Else: txtCNPJCPF.Mask = "##.###.###/####-##"
            End If
         End If
         txtCNPJCPF.Text = CRITERIO
      End If
      txtCNPJCPF.PromptInclude = False

      txtPLACA.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub TXTCNPJCPF_LostFocus()
   txtCNPJCPF.PromptInclude = False
   PESSOA_ID_N = 0

   If Trim(txtCNPJCPF.Text) <> "" Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE "
      SQL = SQL & "where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
         Else
            If TabCliente!NOME <> "" Then
               txtNome.Text = TabCliente!NOME
               PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value

               'If Not IsNull(tabcliente!limite_credito) Then _
                  txtLIMITE.Text = Format(tabcliente!limite_credito, "fixed")
               'SQL = "select sum(i.valor_item-i.valor_desconto) from ITEMLANCAMENTO i, LANCAMENTO l "
               'SQL = SQL & "where i.numr_doc = l.numr_doc "
               'SQL = SQL & " and l.prop = '" & tabcliente!CGCCPF & "'"
               'SQL = SQL & " and i.status = 'A' "
               'TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               'If Not TABAUX.EOF Then
               '   If Not IsNull(TABAUX.Fields(0).Value) Then
               '      txtPAGAR.Text = Format(TABAUX.Fields(0).Value, "fixed")
               '      txtPAGAR.Refresh
               '   End If
               'End If
               'TABAUX.Close
            End If
      End If

      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If
End Sub

Private Sub txtCHASSI_GotFocus()
   MOSTRA_RODAPE "Informe a chassi do Veículo", "", "", "", ""
End Sub

Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCHASSI_KeyPress"
End Sub

Private Sub txtPlaca_GotFocus()
   MOSTRA_RODAPE "Informe a Placa do Veículo", "", "", "", ""
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSCONSULTAVEICULO.Show 1
         If Trim(SQL3) <> "" Then
            If TabAUX.State = 1 Then _
               TabAUX.Close

            SQL = "select placa from VEICULO "
            SQL = SQL & "where placa = '" & Trim(SQL3) & "'"
            TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabAUX.EOF Then _
               MOSTRA_VEICULO
            If TabAUX.State = 1 Then _
               TabAUX.Close
         End If
         SQL3 = ""
         txtPLACA.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPLACA_KeyDown"
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtplaca_KeyPress"
End Sub

Private Sub txtPLACA_LostFocus()
   If Trim(txtPLACA.Text) <> "" Then _
      MOSTRA_VEICULO
End Sub

Private Sub txtDescricao_GotFocus()
   MOSTRA_RODAPE "Informe a descrição do Veículo", "", "", "", ""
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescricao_KeyPress"
End Sub

Private Sub txtMotor_GotFocus()
   MOSTRA_RODAPE "Informe nº motor do Veículo", "", "", "", ""
End Sub

Private Sub txtmotor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtmotor_KeyPress"
End Sub

Private Sub txtCHASSI_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtCHASSI.Text = "" Then
      txtCHASSI.Text = txtPLACA.Text
   '   MsgBox "Chassi inválido."
   '   txtCHASSI.SetFocus
   '   Exit Sub
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCHASSI_LostFocus"
End Sub

Private Sub txtAno_GotFocus()
   MOSTRA_RODAPE "Informe ano de fabricação do Veículo", "", "", "", ""
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtANO_KeyPress"
End Sub

Private Sub txtmodelo_GotFocus()
   MOSTRA_RODAPE "Informe ano do modelo do Veículo", "", "", "", ""
End Sub

Private Sub txtMODELO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMODELO_KeyPress"
End Sub

Private Sub cmbcor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_KeyPress"
End Sub

Private Sub cmbComb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbComb_KeyPress"
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTIPO_KeyPress"
End Sub

Private Sub cmbTipo_GotFocus()
   MOSTRA_RODAPE "Informe tipo do Veículo", "", "", "", ""
End Sub

Private Sub cmbTipo_Click()
'On Error GoTo ERRO_TRATA

   cmbTipoAUX.ListIndex = cmbTipo.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipo_Click"
End Sub

Private Sub cmbMarca_GotFocus()
   MOSTRA_RODAPE "Informe marca do Veículo", "", "", "", ""
End Sub

Private Sub cmbmarca_Click()
'On Error GoTo ERRO_TRATA

   cmbMarcaAUX.ListIndex = cmbMarca.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmarca_Click"
End Sub

Private Sub cmbmarca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbmarca_KeyPress"
End Sub

Private Sub cmbCor_GotFocus()
   MOSTRA_RODAPE "Informe a cor do Veículo", "", "", "", ""
End Sub

Private Sub cmbcor_Click()
'On Error GoTo ERRO_TRATA

   cmbCorAUX.ListIndex = cmbCor.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcor_Click"
End Sub

Private Sub cmbComb_GotFocus()
   MOSTRA_RODAPE "Informe combustível do Veículo", "", "", "", ""
End Sub

Private Sub cmbComb_Click()
'On Error GoTo ERRO_TRATA

   cmbCombAUX.ListIndex = cmbComb.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbComb_Click"
End Sub

Private Sub MOSTRA_VEICULO()
'On Error GoTo ERRO_TRATA

   If Trim(txtPLACA.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwRel_VEICULO "
      SQL = SQL & "where placa = '" & Trim(txtPLACA.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then

         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         SQL = "select * from PESSOA "
         SQL = SQL & "where pessoa_id = " & TabTemp.Fields("pessoa_id").Value
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then

            txtCNPJCPF.PromptInclude = False
               txtCNPJCPF.Text = "" & Trim(TabPessoa.Fields("cnpjcpf").Value)
            txtCNPJCPF.PromptInclude = True

            txtNome.Text = "" & Trim(TabPessoa.Fields("descricao").Value)
         End If
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         txtCHASSI.Text = "" & Trim(TabTemp!chassi)
         txtDescricao.Text = "" & TabTemp!Descricao
         txtMotor.Text = "" & TabTemp!motor
         txtANO.Text = "" & TabTemp!Ano
         txtMODELO.Text = "" & TabTemp!Modelo_id

         If Not IsNull(TabTemp.Fields("cor_id").Value) Then
            If IsNumeric(TabTemp.Fields("cor_id").Value) Then
               cmbCor.Text = "" & TRAZ_DESCRITOR("S", TabTemp.Fields("cor_id").Value)
               cmbCorAUX.Text = "" & TabTemp.Fields("cor_id").Value
            End If
         End If

         If Not IsNull(TabTemp.Fields("marca_id").Value) Then
            If IsNumeric(TabTemp.Fields("marca_id").Value) Then
               cmbMarca.Text = "" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value)
               cmbMarcaAUX.Text = "" & TabTemp.Fields("marca_id").Value
            End If
         End If

         If Not IsNull(TabTemp!combustivel_id) Then
            If IsNumeric(TabTemp.Fields("combustivel_id").Value) Then
               cmbComb.Text = "" & TRAZ_DESCRITOR("U", TabTemp.Fields("combustivel_id").Value)
               cmbCombAUX.Text = "" & TabTemp.Fields("combustivel_id").Value
            End If
         End If

         If Not IsNull(TabTemp!tipo_eqp) Then
            If IsNumeric(TabTemp!tipo_eqp) Then
               cmbTipo.Text = "" & TRAZ_DESCRITOR("A", TabTemp!tipo_eqp)
               cmbTipoAUX.Text = "" & TabTemp!tipo_eqp
            End If
         End If

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select nome,pessoa_id,cliente_id from CLIENTE "
         SQL = SQL & "where CGCCPF = '" & txtCNPJCPF.Text & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then
            txtNome.Text = TabCliente!NOME
            PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
         End If

         If TabCliente.State = 1 Then _
            TabCliente.Close
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VEICULO"
End Sub

Private Sub LIMPA_VEICULO()
'On Error GoTo ERRO_TRATA

   COR_ID_N = 0
   MARCA_ID_N = 0
   TIPO_EQP_ID_N = 0
   ANO_N = 0
   MODELO_ID_N = 0
   VEICULO_ID_N = 0
   EQUIPAMENTO_ID_N = 0
   COMBUSTIVEL_ID_N = 0
   PESSOA_ID_N = 0
   txtPLACA.Text = ""
   txtDescricao.Text = ""
   txtMotor.Text = ""
   txtCHASSI.Text = ""
   cmbCor.Text = ""
   cmbCorAUX.Text = ""
   cmbCombAUX.Text = ""
   cmbComb.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   cmbTipo.Text = ""
   cmbTipoAUX.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   SETA_GRID_VEICULO
   txtPLACA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_VEICULO"
End Sub

Private Sub GRAVA_VEICULO()
'On Error GoTo ERRO_TRATA

   If Trim(txtPLACA.Text) = "" Then
      MsgBox "Número de Placa deve ser informado."
      txtPLACA.SetFocus
      Exit Sub
   End If
   If Trim(txtCHASSI.Text) = "" Then
      MsgBox "Número de Chassi deve ser informado."
      txtCHASSI.SetFocus
      Exit Sub
   End If
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Cliente deve ser informado."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Descrição do Veículo deve ser informada."
      txtDescricao.SetFocus
      Exit Sub
   End If

   COR_ID_N = 0 & cmbCorAUX.Text
   MARCA_ID_N = 0 & cmbMarcaAUX.Text
   TIPO_EQP_ID_N = 0 & cmbTipoAUX.Text
   ANO_ID_N = 0 & txtANO.Text
   MODELO_ID_N = 0 & txtMODELO.Text
   COMBUSTIVEL_ID_N = 0 & cmbCombAUX.Text

   EQUIPAMENTO_ID_N = MAX_ID("equipamento_id", "EQUIPAMENTO", "", "", "", "")
   VEICULO_ID_N = MAX_ID("veiculo_id", "VEICULO", "", "", "", "")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from VEICULO "
   SQL = SQL & "where placa = '" & Trim(txtPLACA.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into EQUIPAMENTO "
      SQL = SQL & "(EQUIPAMENTO_ID,DT_CAD,DESCRICAO,IDENTIFICACAO,PESSOA_ID,COR_ID,MARCA_ID,TIPO_EQP,ANO,MODELO_ID)"
      SQL = SQL & " values("
         SQL = SQL & EQUIPAMENTO_ID_N                       'EQUIPAMENTO_ID
         SQL = SQL & "," & DMA(Date)                        'DT_CAD
         SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'DESCRICAO
         SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'IDENTIFICACAO
         SQL = SQL & "," & PESSOA_ID_N                      'CLIENTE_ID
         SQL = SQL & "," & COR_ID_N                         'COR_ID
         SQL = SQL & "," & MARCA_ID_N                       'MARCA_ID
         SQL = SQL & "," & TIPO_EQP_ID_N                    'TIPO_EQP
         SQL = SQL & "," & ANO_ID_N                         'ANO
         SQL = SQL & "," & MODELO_ID_N                      'MODELO_ID
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into VEICULO "
      SQL = SQL & "(VEICULO_ID,EQUIPAMENTO_ID,COMBUSTIVEL_ID,PLACA,DESCRICAO,MOTOR,CHASSI)"
      SQL = SQL & " values("
         SQL = SQL & VEICULO_ID_N                           'VEICULO_ID
         SQL = SQL & "," & EQUIPAMENTO_ID_N                 'EQUIPAMENTO_ID
         SQL = SQL & "," & COMBUSTIVEL_ID_N                 'COMBUSTIVEL_ID
         SQL = SQL & ",'" & Trim(txtPLACA.Text) & "'"       'PLACA
         SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'DESCRICAO
         SQL = SQL & ",'" & Trim(txtMotor.Text) & "'"       'MOTOR
         SQL = SQL & ",'" & Trim(txtCHASSI.Text) & "'"      'CHASSI
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         EQUIPAMENTO_ID_N = TabTemp.Fields("equipamento_id").Value
         VEICULO_ID_N = TabTemp.Fields("veiculo_id").Value

         SQL = "update EQUIPAMENTO "
         SQL = SQL & "set "
            SQL = SQL & " descricao = '" & Trim(txtDescricao.Text) & "'"      'DESCRICAO
            SQL = SQL & ", IDENTIFICACAO = '" & Trim(txtDescricao.Text) & "'" 'IDENTIFICACAO
            SQL = SQL & ", pessoa_ID = " & PESSOA_ID_N                        'pessoa_ID
            SQL = SQL & ", COR_ID = " & COR_ID_N                              'COR_ID
            SQL = SQL & ", MARCA_ID = " & MARCA_ID_N                          'MARCA_ID
            SQL = SQL & ", TIPO_EQP = " & TIPO_EQP_ID_N                       'TIPO_EQP
            SQL = SQL & ", ANO = " & ANO_ID_N                                 'ANO
            SQL = SQL & ", MODELO_ID = " & MODELO_ID_N                        'MODELO_ID
         SQL = SQL & " where equipamento_id = " & EQUIPAMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "update VEICULO "
         SQL = SQL & "set "
            SQL = SQL & " COMBUSTIVEL_ID = " & COMBUSTIVEL_ID_N               'COMBUSTIVEL_ID
            SQL = SQL & ", PLACA = '" & Trim(txtPLACA.Text) & "'"             'PLACA
            SQL = SQL & ", DESCRICAO = '" & Trim(txtDescricao.Text) & "'"     'DESCRICAO
            SQL = SQL & ", MOTOR = '" & Trim(txtMotor.Text) & "'"             'MOTOR
            SQL = SQL & ", CHASSI = '" & Trim(txtCHASSI.Text) & "'"           'CHASSI
         SQL = SQL & " where veiculo_id = " & VEICULO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   LIMPA_VEICULO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_VEICULO"
End Sub

Private Sub SETA_GRID_VEICULO()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 1
   LISTACHASSI.ListItems.Clear

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from vwRel_VEICULO "
   SQL = SQL & " where chassi <> '' "
   If txtPLACA.Text <> "" Then _
      SQL = SQL & " and placa like " & Chr$(39) & Replace(txtPLACA.Text, "-", "") & "*" & Chr(39)
   SQL = SQL & " order by ano asc "

   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabAUX.EOF
      Set Item = LISTACHASSI.ListItems.Add(, "seq." & TabAUX!placa, TabAUX!chassi)
      Item.SubItems(1) = TabAUX!placa

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from PESSOA "
      SQL = SQL & "where pessoa_id = " & TabAUX.Fields("pessoa_id").Value
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         Item.SubItems(2) = TabTemp.Fields(0).Value

      If TabTemp.State = 1 Then _
         TabTemp.Close

      If Not IsNull(TabAUX!Ano) Then _
         Item.SubItems(3) = TabAUX!Ano
      If Not IsNull(TabAUX!Modelo_id) Then _
         Item.SubItems(4) = TabAUX!Modelo_id
      If Not IsNull(TabAUX!tipo_eqp) Then _
         Item.SubItems(5) = TRAZ_DESCRITOR("A", TabAUX!tipo_eqp)

      TabAUX.MoveNext
   Wend
   If TabAUX.State = 1 Then _
      TabAUX.Close

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_VEICULO"
End Sub

Sub CARREGA_DESCRITORES()
'On Error GoTo ERRO_TRATA

'Tipo Função
' A   Tipo Veículo
' S   Cor
' U   Combustivel
' W   Marca

   cmbTipoAUX.Clear
   cmbTipo.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'A' "
   SQL = SQL & "order by desc_a"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipo.AddItem Trim(TabDESCR!desc_a)
      cmbTipoAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbCorAUX.Clear
   cmbCor.Clear

   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'S' "
   SQL = SQL & "order by desc_a"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbCor.AddItem Trim(TabDESCR!desc_a)
      cmbCorAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close
   
   cmbCombAUX.Clear
   cmbComb.Clear

   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'U' "
   SQL = SQL & "order by desc_a"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbComb.AddItem Trim(TabDESCR!desc_a)
      cmbCombAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbMarcaAUX.Clear
   cmbMarca.Clear

   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'W' "
   SQL = SQL & "order by desc_a"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbMarca.AddItem Trim(TabDESCR!desc_a)
      cmbMarcaAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_DESCRITORES"
End Sub

Public Sub MOSTRA_RODAPE(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   If Trim(Msg1) <> "" Then
      barVeiculo.Panels.Clear
      barVeiculo.Panels.Add (1)
      barVeiculo.Panels(1).Text = Trim(Msg1)
      barVeiculo.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         barVeiculo.Panels.Add (2)
         barVeiculo.Panels(2).Text = Trim(Msg2)
         barVeiculo.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            barVeiculo.Panels.Add (3)
            barVeiculo.Panels(3).Text = Trim(Msg3)
            barVeiculo.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               barVeiculo.Panels.Add (4)
               barVeiculo.Panels(4).Text = Trim(Msg4)
               barVeiculo.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  barVeiculo.Panels.Add (5)
                  barVeiculo.Panels(5).Text = Trim(Msg5)
                  barVeiculo.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If
End Sub
