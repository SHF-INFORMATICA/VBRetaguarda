VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROUSUARIO 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   7005
   ClientLeft      =   2970
   ClientTop       =   2700
   ClientWidth     =   11430
   Icon            =   "CADASTROUSUARIO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11430
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   2055
      Left            =   -1200
      TabIndex        =   47
      Top             =   4920
      Width           =   13575
      Begin VB.TextBox txtL 
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
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   22
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtN 
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
         Left            =   3360
         MaxLength       =   30
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtDDD 
         Alignment       =   2  'Center
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin MSComctlLib.Toolbar TOOBARFONE 
         Height          =   390
         Left            =   9240
         TabIndex        =   48
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "gravar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "matar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Grid 
         Bindings        =   "CADASTROUSUARIO.frx":5C12
         Height          =   1215
         Left            =   1800
         TabIndex        =   53
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         WrapCellPointer =   -1  'True
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ddd"
            Caption         =   "DDD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "numero"
            Caption         =   "Telefone"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "local"
            Caption         =   "Local"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4110,236
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc ADOCabeca 
         Height          =   330
         Left            =   840
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Grid Cabeça"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstEstab 
         Height          =   1335
         Left            =   9720
         TabIndex        =   61
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2355
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
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso Estabelecimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9795
         TabIndex        =   60
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Número:"
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
         TabIndex        =   51
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblLabels 
         Caption         =   "Local:"
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
         Index           =   13
         Left            =   5400
         TabIndex        =   50
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "DDD:"
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
         Left            =   1260
         TabIndex        =   49
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   -75
      TabIndex        =   40
      Top             =   3720
      Width           =   12255
      Begin VB.TextBox txtEndR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9120
         MaxLength       =   80
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtUFR 
         Alignment       =   2  'Center
         DataField       =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   19
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtCidadeR 
         DataField       =   "Cidade"
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
         Left            =   4680
         MaxLength       =   80
         TabIndex        =   18
         Top             =   690
         Width           =   2775
      End
      Begin VB.TextBox txtBairroR 
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
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
         Left            =   960
         MaxLength       =   80
         TabIndex        =   17
         Top             =   690
         Width           =   2760
      End
      Begin VB.TextBox txtRuaR 
         DataField       =   "Endereco_Res"
         DataSource      =   "Data1"
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
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txtCepR 
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
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
         Left            =   8745
         TabIndex        =   46
         Top             =   735
         Width           =   315
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7680
         TabIndex        =   45
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblCep 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   420
         TabIndex        =   44
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   43
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblBairro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   210
         TabIndex        =   42
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         Caption         =   "Rua:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   -75
      TabIndex        =   34
      Top             =   2640
      Width           =   12255
      Begin VB.ComboBox cmbDpAUX 
         BackColor       =   &H80000001&
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
         ItemData        =   "CADASTROUSUARIO.frx":5C2A
         Left            =   9720
         List            =   "CADASTROUSUARIO.frx":5C2C
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkFunc 
         Caption         =   "&Funcionário ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   8160
         TabIndex        =   57
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtComis 
         Alignment       =   1  'Right Justify
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
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbDp 
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
         ItemData        =   "CADASTROUSUARIO.frx":5C2E
         Left            =   9720
         List            =   "CADASTROUSUARIO.frx":5C30
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmbTipoUsu 
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
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5400
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox txtLogon 
         DataField       =   "Endereco_Res"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   11
         Top             =   690
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Tipo Usuário:"
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
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8205
         TabIndex        =   38
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Senha:"
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
         Left            =   4545
         TabIndex        =   37
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Comissão:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4200
         TabIndex        =   36
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Login:"
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
         Left            =   780
         TabIndex        =   35
         Top             =   720
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   -75
      TabIndex        =   28
      Top             =   1605
      Width           =   12255
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
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
         Left            =   8160
         MaxLength       =   5
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "&Email"
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
         Left            =   9720
         TabIndex        =   54
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtRG 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   5
         Top             =   495
         Width           =   1935
      End
      Begin VB.TextBox txtOrigem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   6
         Top             =   495
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtDTEXP 
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Top             =   495
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox txtDtNasc 
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   495
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2160
         TabIndex        =   33
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origem:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4200
         TabIndex        =   32
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Expedição:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6360
         TabIndex        =   31
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8160
         TabIndex        =   30
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Nascimento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   -75
      TabIndex        =   24
      Top             =   750
      Width           =   12615
      Begin VB.ComboBox cmbStatus 
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
         Left            =   9960
         TabIndex        =   3
         Top             =   330
         Width           =   1095
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   1440
         Picture         =   "CADASTROUSUARIO.frx":5C32
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtPath_foto 
         DataField       =   "Note"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9720
         TabIndex        =   52
         Top             =   1320
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox txtNome 
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
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
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
         Left            =   360
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtCPF 
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "*Status:"
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
         Left            =   9960
         TabIndex        =   58
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Código:"
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
         Left            =   360
         TabIndex        =   27
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*CPF:"
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
         Left            =   1920
         TabIndex        =   26
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Nome:"
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
         Left            =   3960
         TabIndex        =   25
         Top             =   120
         Width           =   690
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   23
      ToolTipText     =   "* Campos são obrigatórios"
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   1270
      ButtonWidth     =   2858
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Cadastro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Permisões"
            Key             =   "perm"
            ImageIndex      =   9
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
         Left            =   9720
         TabIndex        =   56
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   4560
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":6634
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3000
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
               Picture         =   "CADASTROUSUARIO.frx":7485
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":88AD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":993C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":ABA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":C2A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":D256
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":E3F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":F622
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROUSUARIO.frx":1062B
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
      DesignWidth     =   11430
      DesignHeight    =   7005
   End
End
Attribute VB_Name = "frmCADASTROUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim TIPO_USUARIO_N As Integer

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela2(frmCADASTROUSUARIO)
   
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   ATUALIZA_TABELA_USUARIO

   cmbSTATUS.Clear
   cmbSTATUS.AddItem "Ativo"
   cmbSTATUS.AddItem "Cancelado"

   cmbTipoUsu.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close
'1  OPERADOR
'2  VENDEDOR (a)
'3  ADMINISTRATIVO
'4  GERENTE
'5  DIRETOR
'6  FINANCEIRO
'7  CAIXA

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'T'"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipoUsu.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbDp.Clear
   cmbDpAUX.Clear

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'V'"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbDp.AddItem Trim(TabDESCR!DESCRICAO)
      cmbDpAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   MOSTRA_ESTABELECIMENTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Resize()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = 0
   txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Resize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
      Case vbKeyF9
         LIMPA_TUDO
         txtCodigo.SetFocus
      Case vbKeyF10
         GRAVA_TUDO
         txtCodigo.SetFocus
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

Private Sub chkFunc_Click()
   txtLogon.SetFocus
End Sub

Private Sub cmbStatus_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe situação cliente", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidade_GotFocus"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   frmCONSULTAUSUARIO.Show 1
   If CRITERIO_A <> "" Then _
      txtCodigo.Text = CRITERIO_A
   CRITERIO_A = ""
   txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub cmdEmail_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   txtCPF.PromptInclude = False
   If Trim(txtCPF.Text) <> "" Then
      CNPJCPF_A = Trim(txtCPF.Text)
      frmEmail.Show 1
   End If
   txtCPF.PromptInclude = True
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
End Sub

Private Sub picimagem_Click()
'On Error GoTo ERRO_TRATA

   frmINICIO.Dialogo.DialogTitle = "Selecione imagem com código do usuário !!!"
   frmINICIO.Dialogo.Filter = "*.jpg;*.gif;*.bmp;*.ico;*.cur"
   frmINICIO.Dialogo.ShowOpen
   If frmINICIO.Dialogo.FileName <> "" Then
      NUMR_SEQ_N = Len(frmINICIO.Dialogo.FileName)
      'txtPath_foto.Text = Right(FRMinicio.Dialogo.FileName, NUMR_SEQ_N - 1)
      txtPath_foto.Text = frmINICIO.Dialogo.FileName
      On Error GoTo PULA_FOTO
      'picimagem.Picture = LoadPicture( Right(FRMinicio.Dialogo.FileName, NUMR_SEQ_N - 1))
'      picimagem.Picture = LoadPicture(Left(FRMinicio.Dialogo.FileName, 1) & Right(FRMinicio.Dialogo.FileName, NUMR_SEQ_N - 1))
   End If
PULA_FOTO:
   Err.Clear

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "picimagem_Click"
End Sub

Private Sub TOOBARFONE_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         txtCPF.PromptInclude = False
         If txtCPF.Text <> "" And txtN.Text <> "" Then
            SQL = "delete  from FONE WITH (NOLOCK)"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and numero = '" & txtN.Text & "'"
            CONECTA_RETAGUARDA.Execute SQL
            CRITERIO_A = txtCPF.Text
         End If
         LIMPA_FONE
         CRITERIO_A = txtCPF.Text
         txtCPF.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOOBARFONE_ButtonClick"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         frmCONSULTAUSUARIO.Show 1
         If CRITERIO_A <> "" Then _
            txtCodigo.Text = CRITERIO_A
         CRITERIO_A = ""
         txtCodigo.SetFocus
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
         txtCodigo.SetFocus
      Case "print"
         FORMULA_REL = "         "
         If txtCodigo.Text <> "" Then _
            FORMULA_REL = "{USUARIO.codigo} = " & txtCodigo.Text

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_usu.rpt"
         frmRELATORIO10.Show 1
      Case "matar"
         MATA_USU
         txtCodigo.SetFocus
      Case "gravar"
         GRAVA_TUDO
      Case "foto"
         frmINICIO.Dialogo.Filter = "*.BMP"
         frmINICIO.Dialogo.InitDir = App.Path & "\FOTO"
         '.Picture = LoadPicture(FRMinicio.Dialogo.FileName)
         'PATH_FOTO = FRMinicio.Dialogo.FileName
      Case "perm"
         FrmControleAcesso.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtBairroR_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o bairro"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBairroR_GotFocus"
End Sub

Private Sub txtCidadeR_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe a cidade"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCidadeR_GotFocus"
End Sub

Private Sub txtCodigo_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe o código ou tecle <<ENTER>>", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCODIGO_GotFocus"
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         MATA_USU
      Case vbKeyF7
         frmCONSULTAUSUARIO.Show 1
         If CRITERIO_A <> "" Then _
            txtCodigo.Text = CRITERIO_A
         CRITERIO_A = ""
         txtCodigo.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodigo_KeyDown"
End Sub

Private Sub txtCodigo_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtCodigo.Text = "" Then _
      txtCodigo.Text = MAX_ID("usuario_id", "usuario", "empresa_id", "1", "", "")

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON USUARIO.PESSOA_ID = PESSOA.PESSOA_ID"
   SQL = SQL & " where usuario_id = " & txtCodigo.Text
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      LIMPA_BODY
      Else: MOSTRA_DADOS
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   txtNome.Enabled = True
   txtCPF.Enabled = True

   'SendKeys "{tab}"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodigo_LostFocus"
End Sub

Private Sub txtComis_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe comissão de venda", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtComis_GotFocus"
End Sub

Private Sub txtcpf_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCPF.SelStart = 0
   txtCPF.SelLength = Len(txtCPF.Mask)

   txtCPF.PromptInclude = False
   If txtCPF.Text = "" Then _
      txtCPF.Mask = "##############"

   txtCPF.PromptInclude = True

   MOSTRA_RODAPE "ESC - Sair", "F7 - Consultar Cliente", "Informe CPF do cliente", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcpf_GotFocus"
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      txtCPF.PromptInclude = False
      If Trim(txtCPF.Text) = "" Then
         MsgBox "Campo obrigatório."
         txtCPF.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      If Len(Trim(txtCPF.Text)) > 0 Then
         Select Case Len(Trim(txtCPF.Text))
            Case Is = 11
               If Not CALCULACPF(Trim(txtCPF.Text)) Then
                  MsgBox "CPF com DV incorreto !!!"
                  txtCPF.PromptInclude = False
                  txtCPF = ""
                  txtCPF.SetFocus
                  Exit Sub
               End If
            Case Is = 14
               If Not VALIDACGC(Trim(txtCPF.Text)) Then
                  MsgBox "CNPJ com DV incorreto !!! "
                  txtCPF.PromptInclude = False
                  txtCPF = ""
                  txtCPF.SetFocus
                  Exit Sub
               End If
            Case Is > 14
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCPF = ""
               txtCPF.SetFocus
               Exit Sub
            Case Is < 11
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCPF = ""
               txtCPF.SetFocus
               Exit Sub
         End Select
         Else
            MsgBox "CNPJ/CPF com DV incorreto !!! "
            txtCPF = ""
            txtCPF.SetFocus
            Exit Sub
      End If

      txtCPF.PromptInclude = False
      CRITERIO_A = Trim(txtCPF.Text)

      If Trim(txtCPF.Text) <> "" Then
         If Not IsNull(txtCPF.Text) Then
            If Len(Trim(txtCPF.Text)) <= 11 Then
               txtCPF.Mask = "###.###.###-##"
               Else: txtCPF.Mask = "##.###.###/####-##"
            End If
         End If
         txtCPF.Text = CRITERIO_A
      End If

      PROCURA_USU

      txtNome.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCPF_KeyPress"
End Sub

Private Sub txtCepR_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCepR.PromptInclude = True

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F4 - Cadastra Cep"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F7 - Consulta Cep"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   txtCPF.PromptInclude = False
   If txtCPF.Text <> "" And Trim(txtNome.Text) <> "" Then
      frmINICIO.BARI.Panels.Add (4)
      frmINICIO.BARI.Panels(4).Text = "F10 - Gravar"
      frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   End If
   txtCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepR_GotFocus"
End Sub

Private Sub txtCepR_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCEP.Show 1
         txtCepR.PromptInclude = False
         txtCepR.Text = CRITERIO_A
         txtCepR.PromptInclude = True
      Case vbKeyF7
         frmCONSULTACEP.Show 1
         txtCepR.PromptInclude = False
         txtCepR.Text = CRITERIO_A
         txtCepR.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepR_KeyDown"
End Sub

Private Sub txtcepr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepR.PromptInclude = False
      If Trim(txtCepR.Text) <> "" Then
         SP_PROCURA_CEP txtCepR.Text
         If TabCEP.EOF Then
            If TabCEP.State = 1 Then _
               TabCEP.Close

            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepR.SetFocus
            Exit Sub
            Else
               txtCidadeR.Text = TabCEP!CIDADE
               txtUFR.Text = TabCEP!UF
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
      txtRuaR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcepr_KeyPress"
End Sub

Private Sub txtDDD_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o DDD"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDDD_GotFocus"
End Sub

Private Sub txtDesconto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe desconto máximo do usuário na venda", "", "", ""

   'txtDesconto.SelStart = 0
   'txtDesconto.Text = FORMAT(txtDesconto.Text, "##.00")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_GotFocus"
End Sub

Private Sub txtEndR_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o endereço"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtEndR_GotFocus"
End Sub

Private Sub txtL_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe descrição/local", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtL_GotFocus"
End Sub

Private Sub txtLogon_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe logon usuário", "", "", ""
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLogon_GotFocus"
End Sub

Private Sub txtLogon_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSenha.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLogon_KeyPress"
End Sub

Private Sub txtN_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe o número telefone", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtN_GotFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe o nome do usuário", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtOrigem_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe origem do número da Identidade", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOrigem_GotFocus"
End Sub

Private Sub txtRG_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o número da Identidade"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRG_GotFocus"
End Sub

Private Sub txtRuaR_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe a rua"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRuaR_GotFocus"
End Sub

Private Sub txtruar_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtruar_KeyPress"
End Sub

Private Sub txtendr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairroR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtendr_KeyPress"
End Sub

Private Sub txtbairror_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidadeR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbairror_KeyPress"
End Sub

Private Sub txtcidader_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUFR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidader_KeyPress"
End Sub

Private Sub txtSenha_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe senha usuário"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtSenha_GotFocus"
End Sub

Private Sub txtUFR_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o estado"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUFR_GotFocus"
End Sub

Private Sub txtufr_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtufr_KeyPress"
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDDD_KeyPress"
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtL.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtN_KeyPress"
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCPF.PromptInclude = False
      If txtN.Text <> "" And txtCPF.Text <> "" Then
         GRAVA_TUDO
         SETA_FONE
         LIMPA_FONE
      End If
      txtCPF.PromptInclude = True
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtl_KeyPress"
End Sub

Private Sub txtOrigem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDTEXP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOrigem_KeyPress"
End Sub

Private Sub txtDTEXP_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDTEXP.PromptInclude = True

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe data expedição do número da Identidade"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEXP_GotFocus"
End Sub

Private Sub txtDTEXP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDesconto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEXP_KeyPress"
End Sub

Private Sub txtOrigem_LostFocus()
'On Error GoTo ERRO_TRATA

   txtOrigem.Text = UCase(txtOrigem.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtOrigem_LostFocus"
End Sub

Private Sub txtRG_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtOrigem.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRg_KeyPress"
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbTipoUsu.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_KeyPress"
End Sub

Private Sub cmbTipoUsu_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe tipo de usuário"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoUsu_GotFocus"
End Sub

Private Sub cmbTipoUsu_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtComis.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoUsu_KeyPress"
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCPF.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodigo_KeyPress"
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSTATUS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtNasc.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbStatus_KeyPress"
End Sub

Private Sub txtDtNasc_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtNasc.PromptInclude = True

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe data de nascimento"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTNASC_GotFocus"
End Sub

Private Sub txtDTNasc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtRg.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtNasc_KeyPress"
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbDp.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtSenha_KeyPress"
End Sub

Private Sub cmbDp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbDp_KeyPress"
End Sub

Private Sub txtcomis_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtLogon.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcomis_KeyPress"
End Sub

Private Sub GRAVA_FONE()
'On Error GoTo ERRO_TRATA

   If Trim(txtN.Text) <> "" Then
      Dim FONE_ID As Integer

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FONE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero = '" & txtN.Text & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "INSERT INTO FONE "
         SQL = SQL & " (PESSOA_ID,NUMERO,DDD,LOCAL) "
         SQL = SQL & " VALUES ("
            SQL = SQL & PESSOA_ID_N
            SQL = SQL & ",'" & Trim(txtN.Text) & "'"
            SQL = SQL & "," & 0 & Trim(txtDDD.Text)
            SQL = SQL & ",'" & Trim(txtL.Text) & "'"
            SQL = SQL & ",0"
         SQL = SQL & ")"
         Else
            SQL = "UPDATE FONE SET "
            SQL = SQL & "Numero = '" & txtN.Text & "'"
            SQL = SQL & ", ddd = 0" & txtDDD.Text & ""
            SQL = SQL & ", local = '" & txtL.Text & "'"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and Numero = '" & Trim(txtN.Text) & "'"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_FONE"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   If txtCodigo.Text = "" Then
      MsgBox "Informe código do funcionário."
      txtCodigo.SetFocus
      Exit Sub
   End If
   txtCPF.PromptInclude = False
   If txtCPF.Text = "" Then
      MsgBox "Informe cpf do funcionário."
      txtCPF.SetFocus
      Exit Sub
   End If
   If Trim(txtNome.Text) = "" Then
      MsgBox "Informe nome do funcionário."
      txtNome.SetFocus
      Exit Sub
   End If
   If cmbTipoUsu.Text = "" Then
      MsgBox "Informe o Tipo do Usuario."
      cmbTipoUsu.SetFocus
      Exit Sub
   End If
   If txtLogon.Text = "" Then
      MsgBox "Informe nome do Login do Usuario para entrar no Sistema."
      txtLogon.SetFocus
      Exit Sub
   End If
   If txtSenha.Text = "" Then
      MsgBox "Informe a Senha do Login do Usuario para entrar no Sistema."
      txtSenha.SetFocus
      Exit Sub
   End If
'=========================PESSOA
   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close

   'executa stored procedure spPessoa
   CONT_N = 1
   If PESSOA_ID_N > 0 Then _
      CONT_N = 2

   spPessoa CONT_N, PESSOA_ID_N, Trim(txtCPF.Text), Trim(txtNome.Text), "", Trim(Left(cmbSTATUS, 1))

   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close
'=========================

   TIPO_USUARIO_N = numeros(cmbTipoUsu.Text)

   Dim INDR_FUNC As Integer

   If chkFunc.Value = 0 Then
      INDR_FUNC = 0
      Else: INDR_FUNC = 1
   End If

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & txtCodigo.Text
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      SQL = "INSERT INTO USUARIO "
         SQL = SQL & " (empresa_id, usuario_id, Nome, Senha, Cpf, "
         SQL = SQL & " DtNasc, Tipo, Perc_desconto, Perc_Comissao, "
         SQL = SQL & " Status, Logon, Classe, Pessoa_id,FUNCIONARIO) "
      SQL = SQL & " VALUES ("
         SQL = SQL & EMPRESA_ID_N                        'empresa_id
         SQL = SQL & "," & txtCodigo.Text                'usuario_id
         SQL = SQL & ",'" & Trim(txtNome.Text) & "'"     'Nome
         SQL = SQL & ",'" & Trim(txtSenha.Text) & "'"    'Senha
         SQL = SQL & ",'" & Trim(txtCPF.Text) & "'"      'Cpf
         SQL = SQL & ",'" & DMA(txtDtNasc.Text) & "'"    'DtNasc
         SQL = SQL & "," & TIPO_USUARIO_N                'Tipo
         SQL = SQL & "," & tpMOEDA(txtDesconto.Text)     'Perc_desconto
         SQL = SQL & "," & tpMOEDA(txtComis.Text)        'Perc_Comissao
         SQL = SQL & "," & 1                             'Status
         SQL = SQL & ",'" & Trim(txtLogon.Text) & "'"    'Logon
         SQL = SQL & ",'A'"                              'Classe
         SQL = SQL & "," & PESSOA_ID_N                   'Pessoa_id
         SQL = SQL & "," & INDR_FUNC               'FUNCIONARIO
         SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "UPDATE USUARIO SET "
         SQL = SQL & " Tipo = " & TIPO_USUARIO_N
         SQL = SQL & ", Logon = '" & Trim(txtLogon.Text) & "'"
         SQL = SQL & ", DtNasc = '" & DMA(txtDtNasc.Text) & "'"
         SQL = SQL & ", Nome = '" & Trim(txtNome.Text) & "'"
         SQL = SQL & ", Perc_desconto = " & tpMOEDA(txtDesconto.Text)
         SQL = SQL & ", Perc_Comissao = " & tpMOEDA(txtComis.Text)
         SQL = SQL & ", senha = '" & Trim(txtSenha.Text) & "'"
         SQL = SQL & ", pessoa_id = " & PESSOA_ID_N
         SQL = SQL & ", funcionario = " & INDR_FUNC
         
         If Trim(Left(cmbSTATUS.Text, 1)) = "A" Then
            SQL = SQL & ", status = 1"
            Else: SQL = SQL & ", status = 0"
         End If

         SQL = SQL & "  Where Usuario_id = " & TabUSU!USUARIO_ID
'SQL = SQL & "  and Empresa_id = " & EMPRESA_ID_N
         CONECTA_RETAGUARDA.Execute SQL
   End If

   If Trim(txtRg.Text) <> "" Then _
      GRAVA_RG Trim(txtRg.Text), Trim(txtOrigem.Text), Trim(txtDTEXP.Text)
   
   GRAVA_FONE
'   GRAVA_ENDERECO
'ENDEREÇO RESIDENCIAL
   txtCepR.PromptInclude = False
   'If Not IsNumeric(txtIBGE.Text) Then _
      txtIBGE.Text = "5201211"

   If txtCepR.Text <> "" Or txtRuaR.Text <> "" Or txtBairroR.Text <> "" Or txtEndR.Text <> "" Then
      sp_Grava_Endereco txtCepR.Text, txtRuaR.Text, txtBairroR.Text, txtEndR.Text, "R", 0
      Else: SP_MATA_ENDEREÇO "R"
   End If

   GRAVA_ACESSO_ESTABELECIMENTO

   If TabUSU.State = 1 Then _
      TabUSU.Close

   LIMPA_TUDO
   txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub MATA_USU()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & txtCodigo.Text
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      MsgBox "Registro não encontrado.", vbOKOnly, "Atenção !!!"
      Else
         Msg = "Confirma Exclusão?"
         Style = vbYesNo + 32
         Title = "Atenção !!!"
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            SQL = "UPDATE USUARIO SET Status = 0"
            SQL = SQL & " Where usuario_id = " & txtCodigo.Text
            CONECTA_RETAGUARDA.Execute SQL
            LIMPA_TUDO
         End If
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_USU"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtCodigo.Text = ""
   'picimagem.Picture = LoadPicture("")
   txtPath_foto.Text = ""
   txtLogon.Text = ""
   chkFunc.Value = False
   cmbSTATUS.Text = ""
   LIMPA_BODY
   txtCPF.PromptInclude = False
   txtCPF.Text = ""

   MOSTRA_ESTABELECIMENTO
   SETA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtCPF.PromptInclude = False
   txtCPF.Text = ""
   txtNome.Text = ""
   txtDtNasc.PromptInclude = False
   txtDtNasc.Text = ""
   txtRg.Text = ""
   txtOrigem.Text = ""
   txtDTEXP.PromptInclude = False
   txtDTEXP.Text = ""
   txtDesconto.Text = ""
   txtComis.Text = ""
   cmbTipoUsu.Text = ""
   cmbDp.Text = ""
   cmbDpAUX.Text = ""
   txtSenha.Text = ""
   txtCepR.PromptInclude = False
   txtCepR.Text = ""
   txtRuaR.Text = ""
   txtEndR.Text = ""
   txtBairroR.Text = ""
   txtCidadeR.Text = ""
   txtUFR.Text = ""
   LIMPA_FONE
   CRITERIO_A = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub LIMPA_FONE()
'On Error GoTo ERRO_TRATA

   txtN.Text = ""
   txtDDD.Text = ""
   txtL.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_FONE"
End Sub

Private Sub PROCURA_USU()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where cpf = '" & txtCPF.Text & "'"
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      MOSTRA_DADOS
   If TabUSU.State = 1 Then _
      TabUSU.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_USU"
End Sub

Private Sub MOSTRA_DADOS()
'On Error GoTo ERRO_TRATA

   Dim i

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select ESTABELECIMENTO_ID from ESTABELECIMENTOACESSO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & txtCodigo.Text
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      If lstEstab.ListItems.Count > 0 Then
         For i = lstEstab.ListItems.Count To 1 Step -1
            If lstEstab.ListItems(i).SubItems(1) = TabConsulta.Fields(0).Value Then _
               lstEstab.ListItems(i).Checked = True
         Next i
      End If

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   chkFunc.Value = False
   PESSOA_ID_N = 0 & TabUSU.Fields("PESSOA_ID").Value

   If Not IsNull(TabUSU.Fields("funcionario").Value) Then
      If TabUSU.Fields("funcionario").Value = False Then
         chkFunc.Value = 0
         Else: chkFunc.Value = 1
      End If
   End If

   If Not IsNull(TabUSU.Fields("status").Value) Then
      If TabUSU.Fields("status").Value = False Then
         cmbSTATUS.Text = "Cancelado"
         Else: cmbSTATUS.Text = "Ativo"
      End If
   End If

   txtCodigo.Text = TabUSU!USUARIO_ID
   txtNome.Text = TabUSU!NOME

   If Not IsNull(TabUSU!Logon) Then _
      txtLogon.Text = "" & TabUSU!Logon

   'If Not IsNull(TABUSU!PATH_FOTO) Then
      'txtPath_foto.Text = TABUSU!PATH_FOTO
      'On Error Resume Next
      'picimagem.Picture = LoadPicture(txtPath_foto.Text)
  ' End If

   If TabUSU!Senha <> "" Then _
      txtSenha.Text = TabUSU!Senha

   If Trim(TabUSU!CPF) <> "" Then
      txtCPF.PromptInclude = False
      txtCPF.Text = TabUSU!CPF
      CRITERIO_A = TabUSU!CPF

      MOSTRA_ENDERECO

      Else: CRITERIO_A = 0
   End If

   txtDtNasc.PromptInclude = False
   If IsDate(TabUSU!DtNasc) Then _
      txtDtNasc.Text = TabUSU!DtNasc

   If Not IsNull(TabUSU!TIPO) Then
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from DESCR WITH (NOLOCK)"
      SQL = SQL & " where TIPO = 'T' "
      SQL = SQL & "and codigo = " & TabUSU!TIPO
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         cmbTipoUsu.Text = TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

   txtDesconto.Text = "" & Format(TabUSU!PERC_DESCONTO, strFormatacao2Digitos)
   txtComis.Text = "" & Format(TabUSU!PERC_COMISSAO, strFormatacao2Digitos)

   'RG
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from RG WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtRg.Text = TabTemp!numero_rg
      If Not IsNull(TabTemp!orgao) Then _
         txtOrigem.Text = TabTemp!orgao
      If IsDate(TabTemp!Dt_Exp) Then
         txtDTEXP.PromptInclude = False
         txtDTEXP.Text = TabTemp!Dt_Exp
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SETA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS"
End Sub
'==============================================
Private Sub MOSTRA_ENDERECO()
'On Error GoTo ERRO_TRATA

   txtCPF.PromptInclude = False
'ok
   BUSCA_ENDERECO_PESSOA "R", ""
   If Not tabEndereco.EOF Then
      txtRuaR.Text = "" & tabEndereco.Fields("Rua")
      txtBairroR.Text = "" & tabEndereco!Bairro
      txtEndR.Text = "" & tabEndereco!Complemento
      If Not IsNull(tabEndereco!CEP_ID) Then
         If tabEndereco!CEP_ID <> "" Then
            txtCepR.Text = tabEndereco!CEP_ID

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from CEP WITH (NOLOCK)"
            SQL = SQL & " where cep_ID = '" & tabEndereco!CEP_ID & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCepR.Text = tabEndereco!CEP_ID
               txtCidadeR.Text = TabConsulta!CIDADE
               txtUFR.Text = TabConsulta!UF
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ENDERECO"
End Sub

Private Sub GRAVA_ENDERECO()
'On Error GoTo ERRO_TRATA

   sp_Grava_Endereco txtCepR.Text, txtRuaR.Text, txtBairroR.Text, txtEndR.Text, "R", 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ENDERECO"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   ADOCabeca.Enabled = True
   ADOCabeca.ConnectionString = AUTENTICA_GRID

   SQL = "select * from FONE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " order by NUMERO"

   ADOCabeca.RecordSource = SQL
   ADOCabeca.Enabled = True
   ADOCabeca.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_FONE"
End Sub

Sub GRAVA_ACESSO_ESTABELECIMENTO()
'On Error GoTo ERRO_TRATA

   Dim i

   If Trim(txtCodigo.Text) <> "" Then
      If IsNumeric(txtCodigo.Text) Then
         SQL = "delete from ESTABELECIMENTOACESSO "
         SQL = SQL & " where usuario_id = " & txtCodigo.Text
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If lstEstab.ListItems.Count > 0 Then
      For i = lstEstab.ListItems.Count To 1 Step -1
         If lstEstab.ListItems(i).Checked = True Then _
            GRAVA_ESTABELECIMENTOACESSO txtCodigo.Text, lstEstab.ListItems(i).SubItems(1)
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ACESSO_ESTABELECIMENTO"
End Sub

Sub MOSTRA_ESTABELECIMENTO()
'On Error GoTo ERRO_TRATA

   lstEstab.ListItems.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close
   SQL = "select ESTABELECIMENTO_id,LOCALIZACAO from ESTABELECIMENTO WITH (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      Set item = lstEstab.ListItems.Add(, "seq." & TabDESCR.Fields(0).Value, Trim(TabDESCR.Fields(1).Value))
      item.SubItems(1) = "" & TabDESCR.Fields(0).Value

      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ESTABELECIMENTO"
End Sub

