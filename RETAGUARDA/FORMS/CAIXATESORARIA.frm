VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCAIXATESORARIA 
   Caption         =   "Movimento Caixa Tesouraria"
   ClientHeight    =   7830
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CAIXATESORARIA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lançamentos"
      TabPicture(0)   =   "CAIXATESORARIA.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTotReg"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lstListaLançamentos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lstMod"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lstSaldo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDEBITO"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSALDOBALCAO"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtSaldoAtual"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCREDITO"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FraSeq"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtSaldoAnterior"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDATA"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Consultas"
      TabPicture(1)   =   "CAIXATESORARIA.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "Label16(0)"
      Tab(1).Control(3)=   "Label16(1)"
      Tab(1).Control(4)=   "lblProc"
      Tab(1).Control(5)=   "lstCC"
      Tab(1).Control(6)=   "lstForma"
      Tab(1).Control(7)=   "txtDtFim"
      Tab(1).Control(8)=   "txtDtIni"
      Tab(1).Control(9)=   "chkImp"
      Tab(1).ControlCount=   10
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   -64560
         TabIndex        =   44
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtDATA 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   30
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtSaldoAnterior 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2760
         TabIndex        =   29
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Frame FraSeq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   50
         TabIndex        =   19
         Top             =   1920
         Width           =   11775
         Begin VB.ComboBox cmbCCAux 
            BackColor       =   &H00FFC0C0&
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
            Left            =   5160
            TabIndex        =   41
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbCC 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5160
            TabIndex        =   6
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox txtValrSaida 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   10200
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtHistorico 
            Height          =   360
            Left            =   6240
            MaxLength       =   100
            TabIndex        =   3
            Top             =   240
            Width           =   5415
         End
         Begin VB.TextBox txtMOD 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   2
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtTIPO 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4560
            MaxLength       =   1
            TabIndex        =   5
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSeq 
            Height          =   360
            Left            =   720
            MaxLength       =   5
            TabIndex        =   0
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtValrEntra 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   10200
            MaxLength       =   12
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtDOC 
            Height          =   360
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   4
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtORIGEM 
            Enabled         =   0   'False
            Height          =   360
            Left            =   2640
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdMata 
            BackColor       =   &H00FFFFFF&
            Height          =   350
            Left            =   1320
            Picture         =   "CAIXATESORARIA.frx":5C4A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
            Height          =   240
            Left            =   5160
            TabIndex        =   40
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Valor Saída ="
            Height          =   240
            Left            =   8760
            TabIndex        =   28
            Top             =   1080
            Width           =   1305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   240
            Left            =   3960
            TabIndex        =   27
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
            Height          =   240
            Left            =   5205
            TabIndex        =   26
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Seq.:"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor Entrada ="
            Height          =   240
            Left            =   8640
            TabIndex        =   24
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Doc./Req.:"
            Height          =   240
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Origem:"
            Height          =   240
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mod.:"
            Height          =   240
            Index           =   1
            Left            =   3960
            TabIndex        =   21
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.TextBox txtCREDITO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   10080
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtSaldoAtual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   10080
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtSALDOBALCAO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2760
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtDEBITO 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   390
         Left            =   10080
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstSaldo 
         Height          =   1065
         Left            =   6240
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   1879
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12640511
         Appearance      =   0
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mod."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modalidade"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComctlLib.ListView lstMod 
         Height          =   2745
         Left            =   4680
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   4842
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777152
         Appearance      =   0
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modalidade"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstListaLançamentos 
         Height          =   3345
         Left            =   50
         TabIndex        =   31
         Top             =   3600
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   5900
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Seq."
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Origem"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Histórico"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Doc./Req."
            Object.Width           =   11465
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tipo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Vlr.Entrada"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Vlr.Saida"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CC"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   300
         Left            =   -73605
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   300
         Left            =   -70965
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSComctlLib.ListView lstForma 
         Height          =   855
         Left            =   -68880
         TabIndex        =   11
         Top             =   1200
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1508
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSComctlLib.ListView lstCC 
         Height          =   855
         Left            =   -74880
         TabIndex        =   45
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1508
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin VB.Label lblProc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   -74400
         TabIndex        =   48
         Top             =   6720
         Width           =   75
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Formas Pagto"
         Height          =   240
         Index           =   1
         Left            =   -68880
         TabIndex        =   47
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Centro Custo"
         Height          =   240
         Index           =   0
         Left            =   -74880
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial:"
         Height          =   240
         Left            =   -74805
         TabIndex        =   43
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final:"
         Height          =   240
         Left            =   -72165
         TabIndex        =   42
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sd. Diário Venda ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         TabIndex        =   39
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2040
         TabIndex        =   37
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   36
         Top             =   1440
         Width           =   1710
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Atual ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8595
         TabIndex        =   35
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label lblTotReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   34
         Top             =   7800
         Width           =   60
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Saída ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8565
         TabIndex        =   33
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Entrada ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8265
         TabIndex        =   32
         Top             =   480
         Width           =   1620
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1270
      ButtonWidth     =   2487
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
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7200
         Top             =   120
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
               Picture         =   "CAIXATESORARIA.frx":6A8B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXATESORARIA.frx":7EB3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXATESORARIA.frx":8F42
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXATESORARIA.frx":A0DC
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
      MaxFontSize     =   100
      DesignWidth     =   12105
      DesignHeight    =   7830
   End
End
Attribute VB_Name = "frmCAIXATESORARIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_SALDO_ATUAL         As Double
   Dim SALDO_DIA_VENDA           As Double
   Dim STATUS_CX_FECHADO         As String

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   frmCAIXATESORARIA.Caption = "Movimento Caixa Tesoraria"

   LIMPA_LANC

   frmCAIXATESORARIA.Caption = "Movimento Caixa Tesoraria" & " ; Usuário = " & TRAZ_NOME_USUARIO(USUARIO_ID_N)

   BUSCA_DIA_ATUAL               'verifica se caixa foi aberto
   Verifica_Caixa_Dia_Anterior   'verifica se o caixa dia anterior está fechado
   Toolbar1.Buttons(5).Visible = False

   SETA_SALDO_ANTERIOR
   SETA_SALDO_ATUAL

   BUSCA_TITULOS_BAIXADOS
   SETA_GRID_LANÇAMENTOS

   cmbCCAux.Clear
   cmbCC.Clear
   lstCC.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR WITH (NOLOCK) "
   SQL = SQL & " where TIPO = 'O'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCC.AddItem Trim(TabTemp!DESCRICAO)
      cmbCCAux.AddItem TabTemp!Codigo

      Set item = lstCC.ListItems.Add(, "seq." & TabTemp!Codigo, Trim(TabTemp!DESCRICAO))
      item.SubItems(1) = "" & TabTemp!Codigo

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub cmbCC_Click()
On Error Resume Next

   cmbCCAux.ListIndex = cmbCC.ListIndex

   txtValrSaida.Enabled = False
   txtValrEntra.Enabled = False

   If txtTIPO.Text = "D" Then
      txtValrSaida.Enabled = True
      txtValrEntra.Enabled = False
      txtValrSaida.SetFocus
      Exit Sub
      Else
         If txtTIPO.Text = "C" Then
            txtValrSaida.Enabled = False
            txtValrEntra.Enabled = True
            txtValrEntra.SetFocus
         End If
   End If
End Sub

Private Sub lstLISTALANÇAMENTOS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstListaLançamentos, ColumnHeader
End Sub

Private Sub LSTMOD_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstMod.SelectedItem) Then
      If lstMod.SelectedItem <> "" Then
         If IsNumeric(lstMod.SelectedItem) Then
            If TabDESCR.State = 1 Then _
               TabDESCR.Close
            SQL = "select * from FORMAPAGTO WITH (NOLOCK) "
            SQL = SQL & " where formapagto_id = " & lstMod.SelectedItem
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and status = 'true' "
            TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabDESCR.EOF Then
               txtMod.Text = TabDESCR!FORMAPAGTO_ID
               txtHistorico.Text = TabDESCR!DESCRICAO
               lstMod.Enabled = False
            End If
            If TabDESCR.State = 1 Then _
               TabDESCR.Close
         End If
      End If
   End If
   lstMod.Visible = False
   lstListaLançamentos.Refresh
   txtHistorico.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LSTMOD_DblClick"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Toolbar1.Buttons(5).Visible = True
   If SSTab1.Tab = 0 Then
      Toolbar1.Buttons(5).Visible = False
      txtSeq.SetFocus
      Else
         CARREGA_FORMAPAGTO
         txtDtIni.SetFocus
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "print"
         CRIA_TAB_TEMP
         MONTA_RELATORIO
      Case "limpar"
         If SSTab1.Tab = 0 Then
            LIMPA_BODY
            txtSeq.SetFocus
            If SSTab1.Tab = 1 Then
               txtDtIni.PromptInclude = False
               txtDtFim.PromptInclude = False
               txtDtIni.Text = ""
               txtDtFim.Text = ""
               lstCC.ListItems.Clear
               lstForma.ListItems.Clear
               txtDtIni.SetFocus
            End If
         End If
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub cmdMata_Click()
   MATA_SEQ
End Sub

Private Sub txtDOC_LostFocus()
   txtDoc.BackColor = &HFFFFFF
End Sub

Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub txtDtFim_LostFocus()
   txtDtFim.BackColor = &HFFFFFF
End Sub

Private Sub txtOrigem_GotFocus()
   txtOrigem.SelStart = 0
   txtOrigem.SelLength = Len(txtOrigem.Text)
   txtOrigem.BackColor = &HC0FFFF
End Sub

Private Sub txtOrigem_LostFocus()
   txtOrigem.BackColor = &HFFFFFF
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   txtSeq.SelStart = 0
   txtSeq.SelLength = Len(txtSeq.Text)
   txtSeq.BackColor = &HC0FFFF

   MOSTRA_RODAPE "ESC - SAIR", "Tecle <ENTER> para gerar nova seqüência ou informe uma já existente", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         MATA_SEQ
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyDown"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
    
      NUMR_ID_N = 1
   
      SQL = "select max(CAIXATESORARIAITEM_ID) from CAIXATESORARIAITEM WITH (NOLOCK) "
      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
      SQL = SQL & " and CAIXATESORARIAITEM_ID < 9999 "

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            NUMR_ID_N = TabTemp.Fields(0).Value + 1
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If Trim(txtSeq.Text) <> "" Then _
         If IsNumeric(txtSeq.Text) Then _
            NUMR_ID_N = 0 & txtSeq.Text

      txtSeq.Text = NUMR_ID_N
      txtOrigem.Text = "Tesoraria"

      PROCURA_SEQ

      KeyAscii = 0
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyPress"
End Sub

Private Sub txtseq_LostFocus()
   txtSeq.BackColor = &HFFFFFF
End Sub


Private Sub txtMOD_GotFocus()
'On Error GoTo ERRO_TRATA

   txtMod.SelStart = 0
   txtMod.SelLength = Len(txtMod.Text)
   txtMod.BackColor = &HC0FFFF

   MOSTRA_RODAPE "ESC - SAIR", "Informe modalidade ou tecle <<ENTER>> ", "", "", ""

   lstMod.Visible = True
   SETA_GRID_MOD

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMOD_GotFocus"
End Sub

Private Sub txtmod_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(txtMod.Text) = "" Then _
         txtMod.Text = 1

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
      SQL = "select * from FORMAPAGTO WITH (NOLOCK) "
      SQL = SQL & " where formapagto_id = " & txtMod.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and status = 'true' "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then
         txtMod.Text = TabDESCR!FORMAPAGTO_ID
         txtHistorico.Text = TabDESCR!DESCRICAO
         lstMod.Visible = False
      End If
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
      
      KeyAscii = 0

      lstMod.Visible = False
      txtHistorico.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtmod_KeyPress"
End Sub

Private Sub txtMOD_LostFocus()
   txtMod.BackColor = &HFFFFFF
End Sub

Private Sub txthistorico_GotFocus()
'On Error GoTo ERRO_TRATA

   txtHistorico.SelStart = 0
   txtHistorico.SelLength = Len(txtHistorico.Text)
   txtHistorico.BackColor = &HC0FFFF

   MOSTRA_RODAPE "ESC - SAIR", "Informe Histórico", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txthistorico_gotFocus"
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDoc.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtHISTORICO_KeyPress"
End Sub

Private Sub txthistorico_LostFocus()
   txtHistorico.BackColor = &HFFFFFF
End Sub

Private Sub txtDOC_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDoc.SelStart = 0
   txtDoc.SelLength = Len(txtDoc.Text)
   txtDoc.BackColor = &HC0FFFF

   lstMod.Visible = False
   lstMod.Refresh
   lstListaLançamentos.Refresh

   MOSTRA_RODAPE "ESC - SAIR", "Informe númedo de documento", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDOC_GotFocus"
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   lstMod.Visible = False
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtTIPO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDoc_KeyPress"
End Sub

Private Sub txtTIPO_GotFocus()
'On Error GoTo ERRO_TRATA

   txtTIPO.SelStart = 0
   txtTIPO.SelLength = Len(txtTIPO.Text)
   txtTIPO.BackColor = &HC0FFFF

   MOSTRA_RODAPE "ESC - SAIR", "Informe (D) = Débito ; (C) = Crédito", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTIPO_GotFocus"
End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      txtValrSaida.Enabled = False
      txtValrEntra.Enabled = False

      If txtTIPO.Text = "D" Then
         txtValrSaida.Enabled = True
         txtValrEntra.Enabled = False
         txtValrSaida.SetFocus
         Exit Sub
         Else
            If txtTIPO.Text = "C" Then
               txtValrSaida.Enabled = False
               txtValrEntra.Enabled = True
               txtValrEntra.SetFocus
            End If
      End If

   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTipo_KeyPress"
End Sub

Private Sub txtTIPO_LostFocus()
   txtValrSaida.Enabled = False
   txtValrEntra.Enabled = False

   If txtTIPO.Text = "D" Then
      txtValrSaida.Enabled = True
      txtValrEntra.Enabled = False
      txtValrSaida.SetFocus
      Exit Sub
      Else
         If txtTIPO.Text = "C" Then
            txtValrSaida.Enabled = False
            txtValrEntra.Enabled = True
            txtValrEntra.SetFocus
         End If
   End If
   txtTIPO.BackColor = &HFFFFFF
End Sub

Private Sub txtValrEntra_GotFocus()
   txtValrEntra.SelStart = 0
   txtValrEntra.SelLength = Len(txtValrEntra.Text)
   txtValrEntra.BackColor = &HC0FFFF
End Sub

Private Sub txtValrEntra_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(cmbCCAux.Text) = "" Then _
         cmbCCAux.Text = "0"

      GRAVA_TUDO
      LIMPA_BODY
      SETA_GRID_LANÇAMENTOS

      txtSeq.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValrEntra_KeyPress"
End Sub

Private Sub txtValrEntra_LostFocus()
   txtValrEntra.BackColor = &HFFFFFF
End Sub

Private Sub txtValrSaida_GotFocus()
   txtValrSaida.SelStart = 0
   txtValrSaida.SelLength = Len(txtValrSaida.Text)
   txtValrSaida.BackColor = &HC0FFFF
End Sub

Private Sub txtValrSaida_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(cmbCCAux.Text) = "" Then _
         cmbCCAux.Text = "0"

      GRAVA_TUDO
      LIMPA_BODY
      SETA_GRID_LANÇAMENTOS

      txtSeq.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValrsaida_KeyPress"
End Sub

Private Sub txtValrSaida_LostFocus()
   txtValrSaida.BackColor = &HFFFFFF
End Sub
'==========================
Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni.Text)
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
   txtDtFim.SelLength = Len(txtDtFim.Text)
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
'On Error GoTo ERRO_TRATA

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

Private Sub LIMPA_LANC()
'On Error GoTo ERRO_TRATA

   txtData.Text = Date
   txtSaldoAnterior.Text = ""
   txtSaldoAtual.Text = ""
   txtSALDOBALCAO.Text = ""
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_LANC"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtCredito.Text = ""
   txtDebito.Text = ""
   txtSeq.Text = ""
   txtMod.Text = ""
   txtHistorico.Text = ""
   txtTIPO.Text = "D"
   txtDoc.Text = ""
   txtOrigem.Text = ""
   txtValrEntra.Text = ""
   txtValrSaida.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub BUSCA_DIA_ATUAL()
'On Error GoTo ERRO_TRATA

   CAIXA_DIA_ID_N = 0
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select * from CAIXATESORARIA WITH (NOLOCK) "
   SQL = SQL & " where dt_abertura >= '" & DMA(Date, "I") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(Date, "F") & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      CAIXA_DIA_ID_N = TabCAIXA!CAIXATESORARIA_ID
      If Not IsNull(TabCAIXA!dt_fechamento) Then
         If TabCAIXA!dt_fechamento > 0 Then
            If TabCAIXA.State = 1 Then _
               TabCAIXA.Close

            FraSeq.Enabled = False
            MsgBox "Caixa já fechado para essa data, permitido somente consulta."
            Else: FraSeq.Enabled = True
         End If
         Else: FraSeq.Enabled = True
      End If
      Else
         If TabCAIXA.State = 1 Then _
            TabCAIXA.Close
         MsgBox "Caixa não foi aberto para essa data. " & Now
         Unload Me
         End
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Busca_Dia_Atual"
End Sub

Private Sub Verifica_Caixa_Dia_Anterior()
'On Error GoTo ERRO_TRATA

   STATUS_CX_FECHADO = ""
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select * from CAIXATESORARIA WITH (NOLOCK) "
   SQL = SQL & " where dt_abertura < '" & DMA(Date) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & "  order by DT_ABERTURA desc"
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      If Not IsNull(TabCAIXA!dt_fechamento) Then
         If TabCAIXA!dt_fechamento > 0 Then
            Else
               If TabCAIXA.State = 1 Then _
                  TabCAIXA.Close

               MsgBox "Caixa dia anterior não foi fechado, verifique."
               STATUS_CX_FECHADO = "F"
               Unload Me
               Exit Sub
         End If
         Else
            If TabCAIXA.State = 1 Then _
               TabCAIXA.Close

            'MsgBox "Caixa dia anterior não foi fechado, verifique."
            'STATUS_CX_FECHADO = "F"
            'Unload Me
            'Exit Sub
      End If
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Verifica_Caixa_Dia_Anterior"
End Sub

Private Sub SETA_GRID_LANÇAMENTOS()
'On Error GoTo ERRO_TRATA

   lstListaLançamentos.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select FORMAPAGTO.FORMAPAGTO_ID, FORMAPAGTO.DESCRICAO, CAIXATESORARIAITEM.CAIXATESORARIA_ID, "
   SQL = SQL & " CAIXATESORARIAITEM.CAIXATESORARIAITEM_ID, CAIXATESORARIAITEM.NUMR_DOC, CAIXATESORARIAITEM.VALOR, "
   SQL = SQL & " CAIXATESORARIAITEM.STATUS, CAIXATESORARIAITEM.Origem , CAIXATESORARIAITEM.tipo, "
   SQL = SQL & " CAIXATESORARIAITEM.HISTORICO, CAIXATESORARIAITEM.CC_ID "
   SQL = SQL & " from FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " ON FORMAPAGTO.FORMAPAGTO_ID = CAIXATESORARIAITEM.FORMAPAGTO_ID"

   SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N

   'SQL = SQL & " order by FORMAPAGTO.DESCRICAO "
   SQL = SQL & " order by CAIXATESORARIAITEM_ID"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      Set item = lstListaLançamentos.ListItems.Add(, "seq." & TabTemp.Fields("CAIXATESORARIAITEM_ID").Value, TabTemp.Fields("CAIXATESORARIAITEM_ID").Value)

      If TabTemp!ORIGEM = "B" Then _
         item.SubItems(1) = "Balcão"
      If TabTemp!ORIGEM = "T" Then _
         item.SubItems(1) = "Tesoraria"

      item.SubItems(2) = "" & TabTemp!FORMAPAGTO_ID & " - " & Trim(TabTemp!HISTORICO)

      If Not IsNull(TabTemp!Numr_doc) Then _
         If Trim(TabTemp!Numr_doc) <> "" Then _
            item.SubItems(3) = TabTemp!Numr_doc

      If Not IsNull(TabTemp.Fields("tipo").Value) Then
         If Trim(TabTemp.Fields("tipo").Value) = "D" Then
            item.SubItems(4) = "Débito"
            'Item.SubItems(6) = "" & Format((TabTemp!Valor * (-1)), strFormatacao2Digitos)
            item.SubItems(6) = "" & Format(TabTemp!VALOR, strFormatacao2Digitos)
         End If
         If Trim(TabTemp.Fields("tipo").Value) = "C" Then
            item.SubItems(4) = "Crédito"
            item.SubItems(5) = "" & Format(TabTemp!VALOR, strFormatacao2Digitos)
         End If
      End If

      If Not IsNull(TabTemp.Fields("cc_id").Value) Then
         item.SubItems(7) = "" & TabTemp.Fields("tipo").Value
         item.SubItems(7) = "" & TRAZ_DESCRITOR("O", TabTemp.Fields("cc_id").Value)
      End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SETA_SALDO_ATUAL
   SETA_GRID_SALDO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_LANÇAMENTOS"
End Sub

Private Sub SETA_GRID_MOD()
'On Error GoTo ERRO_TRATA

   lstMod.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " where formapagto_id < 9999 "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 'true' "
   SQL = SQL & " order by formapagto_id"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstMod.ListItems.Add(, "seq." & TabTemp!FORMAPAGTO_ID, TabTemp!FORMAPAGTO_ID)
      item.SubItems(1) = TabTemp!DESCRICAO
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_MOD"
End Sub

Private Sub PROCURA_SEQ()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
   SQL = SQL & " and CAIXATESORARIAITEM_ID = " & txtSeq.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtSeq.Text = "" & TabTemp!CAIXATESORARIAITEM_ID
      txtMod.Text = "" & TabTemp!FORMAPAGTO_ID
      txtHistorico.Text = "" & TabTemp!HISTORICO
      txtTIPO.Text = "" & TabTemp!TIPO
      If Not IsNull(TabTemp!Numr_doc) Then
         If TabTemp!ORIGEM = "B" Then
            txtDoc.Text = TabTemp!Numr_doc & "/" & TabTemp!CAIXATESORARIAITEM_ID
            Else: txtDoc.Text = TabTemp!Numr_doc
         End If
      End If

      If TabTemp!ORIGEM = "B" Then
         txtOrigem.Text = "Balcão"
         Else: txtOrigem.Text = "Tesoraria"
      End If

      If txtTIPO.Text = "D" Then
         txtValrSaida.Text = TabTemp!VALOR
         Else: txtValrEntra.Text = TabTemp!VALOR
      End If

      Msg = "Seqüência já existente, deseja alterar ?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         txtMod.SetFocus
         Else
            LIMPA_BODY
            txtSeq.SetFocus
      End If
      Else: txtMod.SetFocus
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_SEQ"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   If txtSeq.Text = "" Then
      MsgBox "Seqüência inválida."
      txtSeq.SetFocus
      Exit Sub
   End If
   If txtMod.Text = "" Then
      MsgBox "Modalidade inválida."
      txtMod.SetFocus
      Exit Sub
   End If
   If txtHistorico.Text = "" Then
      MsgBox "Histórico inválido."
      txtHistorico.SetFocus
      Exit Sub
   End If
   If txtTIPO.Text = "" Then
      MsgBox "Informe de é débito ou crédito."
      txtTIPO.SetFocus
      Exit Sub
   End If
   If txtTIPO.Text = "D" Then
      If txtValrSaida.Text = "" Then
         MsgBox "Valor Saida inválido."
         txtValrSaida.SetFocus
         Exit Sub
      End If
   End If
   If txtTIPO.Text = "C" Then
      If txtValrEntra.Text = "" Then
         MsgBox "Valor Entrada inválido."
         txtValrEntra.SetFocus
         Exit Sub
      End If
   End If
   VALOR_ITEM_N = 0

   If txtTIPO.Text = "D" Then
      VALOR_ITEM_N = txtValrSaida.Text
      VALOR_ITEM_N = (txtValrSaida.Text * (-1))
   End If
   If txtTIPO.Text = "C" Then _
      VALOR_ITEM_N = txtValrEntra.Text

   If Trim(cmbCCAux.Text) = "" Then _
      cmbCCAux.Text = "0"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
   SQL = SQL & " and CAIXATESORARIAITEM_ID = " & txtSeq.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtSeq.Text = TabTemp!CAIXATESORARIAITEM_ID

      SQL = "UPDATE CAIXATESORARIAITEM SET "
      SQL = SQL & " Valor = " & tpMOEDA(VALOR_ITEM_N + TabTemp!VALOR)
      SQL = SQL & ",cc_id = " & Trim(cmbCCAux.Text)

      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
      SQL = SQL & " and CAIXATESORARIAITEM_ID = " & txtSeq.Text
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "INSERT INTO CAIXATESORARIAITEM "
         SQL = SQL & " ("
            SQL = SQL & " CAIXATESORARIA_ID, Valor, CAIXATESORARIAITEM_ID, formapagto_id, "
            SQL = SQL & " historico,Tipo, numr_doc, origem, Status,cc_id"
         SQL = SQL & " )"
         SQL = SQL & " VALUES ("
            SQL = SQL & CAIXA_DIA_ID_N
            SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & "," & txtSeq.Text
            SQL = SQL & "," & txtMod.Text
            SQL = SQL & ",'" & txtHistorico.Text & "'"
            SQL = SQL & ",'" & txtTIPO.Text & "'"
            SQL = SQL & ",'" & txtDoc.Text & "'"
            SQL = SQL & ",'" & Left(txtOrigem.Text, 1) & "'"
            SQL = SQL & ",'A'"
            SQL = SQL & "," & Trim(cmbCCAux.Text)
         SQL = SQL & " )"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub SETA_SALDO_ATUAL()
'On Error GoTo ERRO_TRATA

   VALOR_DEBITO = 0
   VALOR_CREDITO = 0
   CONT_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'Para Buscar Saldo Atual tem que sempre pegar o saldo gravado anteriormente
   'pois nao tem nada gravado nos itens do caixa tesouraria
   SQL = "select * from CAIXATESORARIA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN CAIXATESORARIAITEM "
   SQL = SQL & " ON CAIXATESORARIA.CAIXATESORARIA_ID = CAIXATESORARIAITEM.CAIXATESORARIA_ID"
   SQL = SQL & " where CAIXATESORARIA.dt_abertura >= '" & DMA(Date, "i") & "'"
   SQL = SQL & " and CAIXATESORARIA.dt_abertura <= '" & DMA(Date, "f") & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If TabTemp!TIPO = "D" Then _
         VALOR_DEBITO = (VALOR_DEBITO + TabTemp!VALOR)
      If TabTemp!TIPO = "C" Then _
         VALOR_CREDITO = VALOR_CREDITO + TabTemp!VALOR
      CONT_N = CONT_N + 1
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtCredito.Text = Format(VALOR_CREDITO, strFormatacao2Digitos)
   'txtDEBITO.Text = Format((VALOR_DEBITO * (-1)), strFormatacao2Digitos)
   txtDebito.Text = Format(VALOR_DEBITO, strFormatacao2Digitos)
   VALOR_SALDO_ATUAL = (VALOR_CREDITO + VALOR_DEBITO)

   txtSaldoAtual.Text = Format(VALOR_SALDO_ATUAL, strFormatacao2Digitos)
   If VALOR_SALDO_ATUAL < 0 Then
      txtSaldoAtual.ForeColor = vbRed
      Else: txtSaldoAtual.ForeColor = vbBlack
   End If
   txtDebito.ForeColor = vbRed
   txtCredito.ForeColor = vbBlue
   lblTotReg.Caption = "Total Registros =  " & CONT_N
   lblTotReg.Refresh
   txtCredito.Refresh
   txtDebito.Refresh
   txtSaldoAtual.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_SALDO_ATUAL"
End Sub

Private Sub SETA_SALDO_ANTERIOR()
'On Error GoTo ERRO_TRATA

   Dim VALOR_SALDO_dia_ANTERIOR  As Double
   Dim CAIXATESORARIA_ID_N       As Long

'--update CAIXATESORARIAITEM set VALOR = (VALOR * -1) where TIPO = 'D' and VALOR > 0

   VALOR_DEBITO = 0
   VALOR_CREDITO = 0
   VALOR_SALDO_dia_ANTERIOR = 0
   CAIXATESORARIA_ID_N = 0

'============================
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select max(CAIXATESORARIA_id) from CAIXATESORARIA WITH (NOLOCK) "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N

      SQL = SQL & " and dt_abertura < '" & DMA(Date, "I") & "'"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         If TabTemp.Fields(0).Value > 0 Then _
            CAIXATESORARIA_ID_N = 0 & TabTemp.Fields(0).Value
'===================================================================
      'CREDITOS
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select SUM(CAIXATESORARIAITEM.VALOR) from CAIXATESORARIAITEM WITH (NOLOCK) "
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK) "
      SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID "

      'para pegar o dia anterior tem que validar se funciona
      SQL = SQL & " where CAIXATESORARIAITEM.CAIXATESORARIA_ID = " & CAIXATESORARIA_ID_N
      SQL = SQL & " and tipo = 'C' "   'creditos

      'SQL = SQL & " and formapagto_id = 1 " 'SOMENTE DINHEIRO
      SQL = SQL & " and contab_tesora = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            If TabTemp.Fields(0).Value > 0 Then _
               VALOR_CREDITO = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close

      'DEBITOS
      SQL = "select SUM(CAIXATESORARIAITEM.VALOR) from CAIXATESORARIAITEM WITH (NOLOCK) "
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK) "
      SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID "

      'para pegar o dia anterior tem que validar se funciona
      SQL = SQL & " where CAIXATESORARIAITEM.CAIXATESORARIA_ID = " & CAIXATESORARIA_ID_N
      SQL = SQL & " and tipo = 'D' "   'debitos
      'SQL = SQL & " and formapagto_id = 1 " 'SOMENTE DINHEIRO
      SQL = SQL & " and contab_tesora = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            VALOR_DEBITO = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
'===================================================================

   If VALOR_DEBITO > 0 Then
      VALOR_SALDO_dia_ANTERIOR = (VALOR_CREDITO - VALOR_DEBITO)
      Else: VALOR_SALDO_dia_ANTERIOR = (VALOR_CREDITO + VALOR_DEBITO)
   End If

   txtSaldoAnterior.Text = Format(VALOR_SALDO_dia_ANTERIOR, strFormatacao2Digitos)

   If VALOR_SALDO_dia_ANTERIOR < 0 Then
      txtSaldoAnterior.ForeColor = vbRed
      txtTIPO.Text = "D"
      'VALOR_SALDO_dia_ANTERIOR = (VALOR_SALDO_dia_ANTERIOR * (-1))
      Else
         txtSaldoAnterior.ForeColor = vbBlack
         txtTIPO.Text = "C"
   End If
   txtSaldoAnterior.Refresh

   SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
   SQL = SQL & " and CAIXATESORARIAITEM_ID = 9999 "
   SQL = SQL & " and formapagto_id = 9999 "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "INSERT INTO CAIXATESORARIAITEM "
      SQL = SQL & " (CAIXATESORARIA_ID, Valor, CAIXATESORARIAITEM_ID, formapagto_id, historico, "
      SQL = SQL & " Tipo, numr_doc, origem, Status )"
      SQL = SQL & " VALUES ("
         SQL = SQL & CAIXA_DIA_ID_N                                     'CAIXATESORARIA_ID
         SQL = SQL & "," & tpMOEDA(VALOR_SALDO_dia_ANTERIOR)      'Valor
         SQL = SQL & "," & 9999                                   'CAIXATESORARIAITEM_ID
         SQL = SQL & "," & 9999                                   'formapagto_id
         SQL = SQL & ",'" & "Saldo Anterior Transferido" & "'"    'historico
         SQL = SQL & ",'" & Trim(txtTIPO.Text) & "'"              'Tipo
         SQL = SQL & ",'9999'"                                    'numr_doc
         SQL = SQL & ",'" & "T" & "'"                             'origem
         SQL = SQL & ",'" & "A" & "'"                             'Status
      SQL = SQL & " )"

      'CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_SALDO_ANTERIOR"
End Sub

Private Sub BUSCA_TITULOS_BAIXADOS()
'On Error GoTo ERRO_TRATA

   SALDO_DIA_VENDA = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " inner join LANCAMENTO WITH (NOLOCK) "
   SQL = SQL & " on ITEMLANCAMENTO.lancamento_id = LANCAMENTO.lancamento_id "
   SQL = SQL & " where itemlancamento.dt_baixa >= '" & DMA(Date, "i") & "'"
   SQL = SQL & " and itemlancamento.dt_baixa <= '" & DMA(Date, "f") & "'"
   SQL = SQL & " and itemlancamento.status = 'B' "
   SQL = SQL & " and Lancamento.Tipo_lancamento = 1"
   SQL = SQL & " and itemLancamento.formapagto_id = 1"   'dinheiro
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If Not IsNull(TabTemp!Valor_Desconto) Then
         SALDO_DIA_VENDA = SALDO_DIA_VENDA + TabTemp!Valor_Item - TabTemp!Valor_Desconto
         Else: SALDO_DIA_VENDA = SALDO_DIA_VENDA + TabTemp!Valor_Item
      End If
      txtSALDOBALCAO.Text = Format(SALDO_DIA_VENDA, strFormatacao2Digitos)
      txtSALDOBALCAO.Refresh

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
   
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_TITULOS_BAIXADOS"
End Sub
'========================
Private Sub SETA_GRID_SALDO()
'On Error GoTo ERRO_TRATA

   lstSaldo.ListItems.Clear
   NUMR_SEQ_N = 0
   VALOR_DEBITO = 0
   VALOR_CREDITO = 0
   INDR_PRI = False
   SQL3 = ""

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select FORMAPAGTO.FORMAPAGTO_ID, FORMAPAGTO.DESCRICAO, CAIXATESORARIAITEM.CAIXATESORARIA_ID, "
   SQL = SQL & " CAIXATESORARIAITEM.CAIXATESORARIAITEM_ID, CAIXATESORARIAITEM.NUMR_DOC, CAIXATESORARIAITEM.VALOR, "
   SQL = SQL & " CAIXATESORARIAITEM.STATUS, CAIXATESORARIAITEM.Origem , CAIXATESORARIAITEM.tipo, "
   SQL = SQL & " CAIXATESORARIAITEM.HISTORICO "
   SQL = SQL & " from FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " ON FORMAPAGTO.FORMAPAGTO_ID = CAIXATESORARIAITEM.FORMAPAGTO_ID"

   SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N

   'SQL = SQL & " order by FORMAPAGTO.DESCRICAO "

SQL = SQL & " order by CAIXATESORARIAITEM.formapagto_id "

   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabAUX.EOF Then
      NUMR_SEQ_N = TabAUX.Fields("formapagto_id").Value
      SQL3 = Trim(TabAUX.Fields("descricao").Value)
      If TabAUX.Fields("formapagto_id").Value = 9999 Then _
         SQL3 = Trim(TabAUX.Fields("historico").Value)
   End If
   While Not TabAUX.EOF
      If Not IsNull(TabAUX.Fields("tipo").Value) Then
         If Trim(TabAUX.Fields("tipo").Value) <> "" Then

If Not IsNull(TabAUX.Fields("valor").Value) Then

   If TabAUX!TIPO = "D" Then _
      VALOR_DEBITO = VALOR_DEBITO + (TabAUX.Fields("valor").Value)
   If TabAUX!TIPO = "C" Then _
      VALOR_CREDITO = VALOR_CREDITO + (TabAUX.Fields("valor").Value)

End If

'==========indicador
               If TabAUX.Fields("formapagto_id").Value <> NUMR_SEQ_N Then
                  Set item = lstSaldo.ListItems.Add(, "seq." & NUMR_SEQ_N, NUMR_SEQ_N)
                  item.SubItems(1) = "" & SQL3
                  item.SubItems(2) = "" & Format(VALOR_CREDITO + VALOR_DEBITO, strFormatacao2Digitos)

                  VALOR_DEBITO = 0
                  VALOR_CREDITO = 0

                  INDR_PRI = True
                  NUMR_SEQ_N = TabAUX.Fields("formapagto_id").Value
                  SQL3 = Trim(TabAUX.Fields("descricao").Value)
                  If TabAUX.Fields("formapagto_id").Value = 9999 Then _
                     SQL3 = Trim(TabAUX.Fields("historico").Value)
               End If
         End If
      End If

      TabAUX.MoveNext
   Wend
   If TabAUX.State = 1 Then _
      TabAUX.Close

   If INDR_PRI = False Then
      Set item = lstSaldo.ListItems.Add(, "seq." & NUMR_SEQ_N, NUMR_SEQ_N)
      item.SubItems(1) = "" & Trim(SQL3)
      item.SubItems(2) = "" & Format(VALOR_CREDITO + VALOR_DEBITO, strFormatacao2Digitos)

      VALOR_DEBITO = 0
      VALOR_CREDITO = 0
      INDR_PRI = True
   End If

   lstSaldo.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_SALDO"
End Sub

Sub MATA_SEQ()
'On Error GoTo ERRO_TRATA

   If txtSeq.Text <> "" And txtSeq.Text <> "9999" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK) "
      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
      SQL = SQL & " and CAIXATESORARIAITEM_ID = " & txtSeq.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Confirma Exclusão ?"
         PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
          If RESPOSTA = vbYes Then
            SQL = "Delete from CAIXATESORARIAITEM "
            SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
            SQL = SQL & " and CAIXATESORARIAITEM_ID = " & txtSeq.Text
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_BODY
            SETA_GRID_LANÇAMENTOS

            Exit Sub
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_SEQ"
End Sub
'============================================
Sub CARREGA_FORMAPAGTO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select formapagto_id,descricao from FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " where status = 1 "
   'SQL = SQL & " and contab_tesora = 'true' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      Set item = lstForma.ListItems.Add(, "seq." & TabConsulta.Fields(0).Value, Trim(TabConsulta.Fields("descricao").Value))
      item.SubItems(1) = "" & TabConsulta.Fields(0).Value

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Dim i

   If lstForma.ListItems.Count > 0 Then
      For i = lstForma.ListItems.Count To 1 Step -1
         lstForma.ListItems(i).Checked = True
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_FORMAPAGTO"
End Sub

Sub CRIA_TAB_TEMP()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "REL_CX_TES", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP table [dbo].[REL_CX_TES]"

   SQL = "CREATE TABLE [dbo].[REL_CX_TES]("
   SQL = SQL & " [NOME_ESTAB] [nvarchar](60) NOT NULL,"
   SQL = SQL & " [DT_MOVIMENTO] [datetime] NOT NULL,"
   SQL = SQL & " [SEQ_ID] [bigint] NOT NULL,"
   SQL = SQL & " [ORIGEM] [char](1) NOT NULL,"
   SQL = SQL & " [HISTORICO] [varchar](100) NULL,"
   SQL = SQL & " [NUMR_DOC] [varchar](30) NULL,"
   SQL = SQL & " [TIPO] [char](1) NOT NULL,"
   SQL = SQL & " [VALOR] [float] NOT NULL,"
   SQL = SQL & " [CC_ID] [bigint] NULL,"
   SQL = SQL & " [CC_DESC] [varchar](30) NULL"
   SQL = SQL & " ) ON [PRIMARY]"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from REL_CX_TES"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CRIA_TAB_TEMP"
End Sub

Sub MONTA_RELATORIO()
'On Error GoTo ERRO_TRATA

   Dim SELECAO_FORMAPAGTO, i, SELECAO_CC
   Dim CC_ID_N As Long
   SELECAO_FORMAPAGTO = ""
   SELECAO_CC = ""
   lblProc.Caption = ""
   INDR_PRI = True

   If lstForma.ListItems.Count > 0 Then
      For i = lstForma.ListItems.Count To 1 Step -1
         If lstForma.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               SELECAO_FORMAPAGTO = lstForma.ListItems(i).SubItems(1)
               Else: SELECAO_FORMAPAGTO = SELECAO_FORMAPAGTO & "," & lstForma.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   INDR_PRI = True

   If lstCC.ListItems.Count > 0 Then
      For i = lstCC.ListItems.Count To 1 Step -1
         If lstCC.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               SELECAO_CC = lstCC.ListItems(i).SubItems(1)
               Else: SELECAO_CC = SELECAO_CC & "," & lstCC.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select ESTABELECIMENTO.DESCRICAO, CAIXATESORARIA.CAIXATESORARIA_ID, "
   SQL = SQL & " CAIXATESORARIA.ESTABELECIMENTO_ID, CAIXATESORARIA.DT_ABERTURA, "
   SQL = SQL & " CAIXATESORARIA.DT_FECHAMENTO, CAIXATESORARIA.usuario_id, "
   SQL = SQL & " CAIXATESORARIA.STATUS,"
   SQL = SQL & " CAIXATESORARIAITEM.CAIXATESORARIAITEM_ID, CAIXATESORARIAITEM.FORMAPAGTO_ID, "
   SQL = SQL & " CAIXATESORARIAITEM.NUMR_DOC, CAIXATESORARIAITEM.VALOR,"
   SQL = SQL & " CAIXATESORARIAITEM.STATUS AS Status_Item, CAIXATESORARIAITEM.ORIGEM, "
   SQL = SQL & " CAIXATESORARIAITEM.TIPO, CAIXATESORARIAITEM.HISTORICO,"
   SQL = SQL & " FORMAPAGTO.DESCRICAO AS Desc_PAGTO, FORMAPAGTO.STATUS AS Status_PAGTO, "
   SQL = SQL & " FORMAPAGTO.CONTAB_TESORA, CAIXATESORARIAITEM.CC_ID"

   SQL = SQL & " from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN CAIXATESORARIA WITH (NOLOCK)"
   SQL = SQL & " ON ESTABELECIMENTO.ESTABELECIMENTO_ID = CAIXATESORARIA.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN CAIXATESORARIAITEM WITH (NOLOCK)"
   SQL = SQL & " ON CAIXATESORARIA.CAIXATESORARIA_ID = CAIXATESORARIAITEM.CAIXATESORARIA_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " where ESTABELECIMENTO.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and DT_abertura >= '" & Format(txtDtIni.Text, "dd/mm/yyyy") & " 00:00:00'"
      SQL = SQL & " and DT_abertura <= '" & Format(txtDtFim.Text, "dd/mm/yyyy") & " 23:59:59'"
   End If

   If Trim(SELECAO_FORMAPAGTO) <> "" Then _
      SQL = SQL & " and CAIXATESORARIAITEM.formapagto_id in ( " & Trim(SELECAO_FORMAPAGTO) & ")"

   If Trim(SELECAO_CC) <> "" Then _
      SQL = SQL & " and CAIXATESORARIAITEM.cc_id in ( " & Trim(SELECAO_CC) & ")"

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      lblProc.Caption = "Processados = " & TabConsulta.Fields("CAIXATESORARIAITEM_ID").Value
      DoEvents

      CC_ID_N = 0 & TabConsulta.Fields("CC_ID").Value
      SqL2 = CC_ID_N
      SQL3 = Replace(TabConsulta.Fields("historico").Value, ",", ".")
      SQL3 = SQL3 & Replace(TabConsulta.Fields("historico").Value, "'", ".")

      SQL = "insert into REL_CX_TES "
      SQL = SQL & "("
         SQL = SQL & "NOME_ESTAB,DT_MOVIMENTO,SEQ_ID,ORIGEM,HISTORICO,NUMR_DOC,TIPO,VALOR,CC_ID,CC_DESC"
      SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & "'" & Trim(TabConsulta.Fields("DESCRICAO").Value) & "'"     'NOME_ESTAB
         SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_abertura").Value) & "'"   'DT_MOVIMENTO
         SQL = SQL & "," & TabConsulta.Fields("CAIXATESORARIAITEM_ID").Value     'SEQ_ID
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("origem").Value) & "'"       'ORIGEM
         SQL = SQL & ",'" & Trim(SQL3) & "'"                                     'HISTORICO
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("NUMR_DOC").Value) & "'"     'NUMR_DOC
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("TIPO").Value) & "'"         'TIPO
         SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("VALOR").Value)            'VALOR
         SQL = SQL & "," & CC_ID_N                                               'CC_ID
         SQL = SQL & ",'" & TRAZ_DESCRITOR("O", SqL2) & "'"                      'CC_DESC
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   lblProc.Caption = ""

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   FORMULA_REL = ""

   Nome_Relatorio = "CX_TES_CC.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_RELATORIO"
End Sub
