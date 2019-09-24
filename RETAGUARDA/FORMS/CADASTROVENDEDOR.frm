VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROVENDEDOR 
   Caption         =   "Cadastro de Vendedores"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   900
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROVENDEDOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   10440
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   -120
      TabIndex        =   37
      Top             =   2760
      Width           =   10815
      Begin VB.TextBox txtNumeroR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9600
         MaxLength       =   50
         TabIndex        =   13
         Top             =   480
         Width           =   825
      End
      Begin VB.TextBox txtIBGE 
         Appearance      =   0  'Flat
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         LinkTimeout     =   7
         MaxLength       =   50
         TabIndex        =   44
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtUF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8160
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtRua 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   240
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtComp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   360
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   11
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
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9600
         TabIndex        =   46
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "IBGE:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9000
         TabIndex        =   45
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8160
         TabIndex        =   43
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   41
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rua/Avenida/Praça"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2160
         TabIndex        =   40
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6240
         TabIndex        =   38
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   -120
      TabIndex        =   31
      Top             =   4320
      Width           =   10695
      Begin VB.TextBox txtDDD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   840
         MaxLength       =   2
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtL 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4800
         MaxLength       =   30
         TabIndex        =   19
         Top             =   240
         Width           =   3855
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   3480
         TabIndex        =   32
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "gravar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "matar"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Grid 
         Bindings        =   "CADASTROVENDEDOR.frx":5C12
         Height          =   1215
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
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
         Left            =   600
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
         Height          =   1215
         Left            =   7440
         TabIndex        =   50
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2143
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DDD:"
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "LOCAL:"
         Height          =   240
         Index           =   13
         Left            =   3960
         TabIndex        =   33
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   -120
      TabIndex        =   21
      Top             =   600
      Width           =   10935
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
         Left            =   5040
         TabIndex        =   52
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdRg 
         Caption         =   "RG"
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
         Left            =   4440
         Picture         =   "CADASTROVENDEDOR.frx":5C2A
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox cmbTbPrecoAUX 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         TabIndex        =   49
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTbPreco 
         Height          =   360
         Left            =   6840
         TabIndex        =   4
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cmbEquipeAUX 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtCodg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbEquipe 
         Height          =   360
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   360
         ItemData        =   "CADASTROVENDEDOR.frx":6D70
         Left            =   5880
         List            =   "CADASTROVENDEDOR.frx":6D7A
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optValor 
         Caption         =   "V&alor"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optQtd 
         Caption         =   "&Quantidade"
         Height          =   255
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1680
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPerc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtCPF 
         Height          =   360
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox txtDTNasc 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox txtDtCad 
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox txtDtBaixa 
         Height          =   375
         Left            =   9000
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tb.Preço:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5820
         TabIndex        =   48
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipe:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6015
         TabIndex        =   47
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Vendedor:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1560
         TabIndex        =   29
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Nascimento:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inclusão:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Baixa:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7800
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Comissão:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   1680
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar barVendedor 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "sair"
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
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   240
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
            Picture         =   "CADASTROVENDEDOR.frx":6D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVENDEDOR.frx":71E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVENDEDOR.frx":74FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVENDEDOR.frx":7952
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVENDEDOR.frx":7DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVENDEDOR.frx":81FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVENDEDOR.frx":8516
            Key             =   ""
         EndProperty
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
      DesignWidth     =   10440
      DesignHeight    =   6345
   End
End
Attribute VB_Name = "frmCADASTROVENDEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
   CARREGA_COMBO

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

Private Sub barVendedor_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         CRITERIO_A = ""
         frmCONSULTAVENDEDOR.Show 1
         If CRITERIO_A <> "" Then _
            txtCodg.Text = CRITERIO_A
      Case "matar"
         MATAR_VENDEDOR
      Case "gravar"
         If Trim(cmbSTATUS.Text) = "" Then
            MsgBox "Informe Situação."
            cmbSTATUS.SetFocus
            Exit Sub
         End If
      
         If Trim(txtCodg.Text) = "" Then
            MsgBox "Informe codigo."
            txtCodg.SetFocus
            Exit Sub
         End If
      
         If Not IsNumeric(txtCodg.Text) Then
            MsgBox "Informe codigo."
            txtCodg.SetFocus
            Exit Sub
         End If
      
         If Trim(txtNome.Text) = "" Then
            MsgBox "Informe nome."
            txtNome.SetFocus
            Exit Sub
         End If
      
         If Trim(cmbEquipeAUX.Text) = "" Then
            MsgBox "Informe representante."
            cmbEquipe.SetFocus
            Exit Sub
         End If
      
         txtCPF.PromptInclude = False
         If Trim(txtCPF.Text) = "" Then
            MsgBox "Informe CPF."
            txtCPF.SetFocus
            Exit Sub
         End If

         GRAVA_VENDEDOR

         LIMPA_VENDEDOR
         txtCodg.SetFocus
      Case "limpar"
         LIMPA_VENDEDOR
      Case "sair"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "barVendedor_ButtonClick"
End Sub

Private Sub cmdRG_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      CNPJCPF_A = Trim(txtCNPJCPF.Text)
      frmCADASTRORG.Show 1
   End If
   txtCNPJCPF.PromptInclude = True
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCep.SetFocus
   End If
End Sub

Private Sub cmbEquipe_Click()
On Error Resume Next

   cmbEquipeAUX.ListIndex = cmbEquipe.ListIndex

End Sub

Private Sub cmbtbpreco_Click()
On Error Resume Next

   cmbTbPrecoAUX.ListIndex = cmbTbPreco.ListIndex

End Sub

Private Sub cmbEquipe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCPF.SetFocus
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   txtCPF.PromptInclude = False
      If txtCPF.Text <> "" And txtN.Text <> "" Then

         EXCLUIR_REGISTRO_FONE Trim(txtN.Text)

         LIMPA_FONE
         SETA_FONE
         txtN.Text = ""
      End If
   txtCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJ_KeyDown"
End Sub

Private Sub txtbairro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidade.SetFocus
   End If
End Sub

Private Sub txtcep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtRua.SetFocus
   End If
End Sub

Private Sub txtCEP_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCEP.Show 1
         txtCep.PromptInclude = False
         txtCep.Text = CRITERIO_A
         txtCep.PromptInclude = True
      Case vbKeyF7
         frmCONSULTACEP.Show 1
         txtCep.PromptInclude = False
         txtCep.Text = CRITERIO_A
         txtCep.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCEP_KeyDown"
End Sub

Private Sub txtCep_LostFocus()
   txtCep.PromptInclude = False
   If Trim(txtCep.Text) <> "" Then
      SP_PROCURA_CEP Trim(txtCep.Text)
      If TabCEP.EOF Then
         If TabCEP.State = 1 Then _
            TabCEP.Close

         MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
         Exit Sub
         Else
            txtCidade.Text = TabCEP!CIDADE
            txtUF.Text = TabCEP!UF
            If Not IsNull(TabCEP!IBGE_ID) Then _
               txtIBGE.Text = TabCEP!IBGE_ID
      End If
      If TabCEP.State = 1 Then _
         TabCEP.Close
   End If
End Sub

Private Sub txtcidade_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUF.SetFocus
   End If
End Sub

Private Sub txtcodg_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNome.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodg_KeyPress"
End Sub

Private Sub txtcomp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNumeroR.SetFocus
   End If
End Sub

Private Sub txtcpf_GotFocus()
'On Error GoTo ERRO_TRATA

    txtCPF.SelStart = 0
    txtCPF.SelLength = Len(txtCPF.Mask)
    
   txtCPF.PromptInclude = False
   If txtCPF.Text = "" Then _
      txtCPF.Mask = "##############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcpf_GotFocus"
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCPF.PromptInclude = False
      cmbTbPreco.SetFocus
   ElseIf KeyAscii = vbKeyDelete Then
      If Not IsNumeric(txtCPF.Text) Then
         txtCPF.Mask = "##############"
      End If
   ElseIf KeyAscii = vbKeyBack Then
      If Not IsNumeric(txtCPF.Text) Then
         txtCPF.Mask = "##############"
      End If
   End If
End Sub

Private Sub txtCPF_LostFocus()
   txtCPF.PromptInclude = False
      PROCURA_VENDEDOR
   txtCPF.PromptInclude = True
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
   End If
End Sub

Private Sub txtDtBaixa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPerc.SetFocus
   End If
End Sub

Private Sub txtDtCad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtBaixa.SetFocus
   End If
End Sub
Private Sub cmbTbPreco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtNasc.SetFocus
   End If
End Sub

Private Sub txtDTNasc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtCad.SetFocus
   End If
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtN.Text) <> "" Then
         GRAVA_VENDEDOR
'GRAVA_FONE_PESSOA Trim(txtN.Text), Trim(txtDDD.Text), Trim(txtL.Text), "0"
         'GRAVA_FONE Trim(txtCPF.Text), Trim(txtN.Text), 0 & Trim(txtDDD.Text), Trim(txtL.Text), Trim(0)
      End If

      SETA_FONE
      LIMPA_FONE
   End If
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtL.SetFocus
   End If
End Sub

Private Sub txtNumeroR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairro.SetFocus
   End If
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSTATUS.SetFocus
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbEquipe.SetFocus
   End If
End Sub

Private Sub txtCodg_LostFocus()
   MOSTRA_DADOS
End Sub

Private Sub txtrua_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtComp.SetFocus
   End If
End Sub

Private Sub txtuf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtIBGE.SetFocus
   End If
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
'=======================================================
Sub CARREGA_COMBO()
'On Error GoTo ERRO_TRATA

   cmbSTATUS.Clear
   cmbSTATUS.AddItem "Ativo"
   cmbSTATUS.AddItem "Cancelado"
   cmbSTATUS.Text = "Ativo"

   cmbEquipeAUX.Clear
   cmbEquipe.Clear

   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close

   SQL = "select EQUIPE_ID,descricao from EQUIPE "
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by descricao"
   TabEQUIPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEQUIPE.EOF Then
      cmbEquipeAUX.Text = "" & TabEQUIPE!EQUIPE_ID
      cmbEquipe.Text = "" & TabEQUIPE!DESCRICAO
   End If
   While Not TabEQUIPE.EOF
      cmbEquipeAUX.AddItem TabEQUIPE!EQUIPE_ID
      cmbEquipe.AddItem TabEQUIPE!DESCRICAO
      TabEQUIPE.MoveNext
   Wend
   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close

   cmbTbPrecoAUX.Clear
   cmbTbPreco.Clear

   SQL = "select * from TABELAPRECO "
   'SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by descricao"
   TabEQUIPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabEQUIPE.EOF
      cmbTbPrecoAUX.AddItem TabEQUIPE.Fields("tabelapreco_id").Value
      cmbTbPreco.AddItem TabEQUIPE.Fields("descricao").Value
      TabEQUIPE.MoveNext
   Wend
   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close

   lstEstab.ListItems.Clear

   SQL = "select ESTABELECIMENTO_id,localizacao from ESTABELECIMENTO WITH (NOLOCK)"
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
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub

Sub MOSTRA_DADOS()
'On Error GoTo ERRO_TRATA

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   If Trim(txtCodg.Text) <> "" Then
      If IsNumeric(txtCodg.Text) Then
         SQL = "select * from vwVendedor"
         SQL = SQL & " where vendedor_id = " & txtCodg.Text
         TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVENDEDOR.EOF Then
            VENDEDOR_ID_N = TabVENDEDOR.Fields("vendedor_id").Value
            If Not IsNull(TabVENDEDOR.Fields("tabelapreco_id").Value) Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               SQL = "select * from TABELAPRECO WITH (NOLOCK)"
               SQL = SQL & " where tabelapreco_id = " & TabVENDEDOR.Fields("tabelapreco_id").Value
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  cmbTbPreco.Text = "" & Trim(TabTemp!DESCRICAO)
                  cmbTbPrecoAUX.Text = "" & Trim(TabTemp!TABELAPRECO_ID)
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close
            End If

            MOSTRA_VEND
'===============
            Dim i
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select ESTABELECIMENTO.ESTABELECIMENTO_ID, ESTABELECIMENTO.DESCRICAO"
            SQL = SQL & " from ESTABELECIMENTO "
            SQL = SQL & " INNER JOIN ESTABVENDEDOR "
            SQL = SQL & " ON ESTABELECIMENTO.ESTABELECIMENTO_ID = ESTABVENDEDOR.ESTABELECIMENTO_ID"
            SQL = SQL & " where vendedor_id = " & VENDEDOR_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            While Not TabTemp.EOF
               If lstEstab.ListItems.Count > 0 Then
                  For i = lstEstab.ListItems.Count To 1 Step -1
                     If lstEstab.ListItems(i).SubItems(1) = TabTemp.Fields(0).Value Then _
                        lstEstab.ListItems(i).Checked = True
                  Next i
               End If
               TabTemp.MoveNext
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close
'================

            'ENDEREÇO
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from ENDERECO"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and tipo = 'R' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               If Not IsNull(TabConsulta!Rua) Then _
                  txtRua.Text = TabConsulta!Rua
               If Not IsNull(TabConsulta!Bairro) Then _
                  txtBairro.Text = TabConsulta!Bairro
               If Not IsNull(TabConsulta!Complemento) Then _
                  txtComp.Text = TabConsulta!Complemento
               If Not IsNull(TabConsulta!CEP_ID) Then _
                  txtCep.Text = TabConsulta!CEP_ID

               If TabCEP.State = 1 Then _
                  TabCEP.Close

               SQL = "select * from CEP "
               SQL = SQL & " where cep_ID = '" & Trim(TabConsulta.Fields("cep_id").Value) & "'"
               TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCEP.EOF Then
                  If Not IsNull(TabCEP.Fields("Cep_id").Value) Then _
                     txtCep.Text = TabCEP.Fields("Cep_id").Value

                  If Not IsNull(TabCEP!CIDADE) Then _
                     txtCidade.Text = TabCEP!CIDADE

                  If Not IsNull(TabCEP!UF) Then _
                     txtUF.Text = TabCEP!UF
               End If
               If TabCEP.State = 1 Then _
                  TabCEP.Close
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SETA_FONE
         End If
      End If
      Else: GERA_VENDEDOR
   End If
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS"
End Sub

Private Sub MOSTRA_VEND()
'On Error GoTo ERRO_TRATA

   txtNome.Text = TabVENDEDOR.Fields("descricao").Value
   cmbEquipeAUX.Text = TabVENDEDOR.Fields("EQUIPE_ID").Value
   cmbEquipe.Text = TabVENDEDOR.Fields("descequipe").Value

   optQtd.Value = False
   optValor.Value = False
   optQtd.Value = False
   optValor.Value = False

   If Not IsNull(TabVENDEDOR!TIPO_COMIS) Then
      If TabVENDEDOR!TIPO_COMIS = 1 Then
         optValor.Value = True
         Else: optQtd.Value = True
      End If
   End If

   If IsDate(TabVENDEDOR!DT_NASCIMENTO) Then
      txtDtNasc.PromptInclude = False
         txtDtNasc.Text = TabVENDEDOR!DT_NASCIMENTO
      txtDtNasc.PromptInclude = True
   End If

   If IsDate(TabVENDEDOR!DATA_CAD) Then
      txtDtCad.PromptInclude = False
         txtDtCad.Text = TabVENDEDOR!DATA_CAD
      txtDtCad.PromptInclude = True
   End If

   If IsDate(TabVENDEDOR!DT_BAIXA) Then
      txtDtBaixa.PromptInclude = False
         txtDtBaixa.Text = TabVENDEDOR!DT_BAIXA
      txtDtBaixa.PromptInclude = True
   End If

   If Not IsNull(TabVENDEDOR!STATUS) Then
      If TabVENDEDOR!STATUS = "A" Then
         cmbSTATUS.Text = "ATIVO"
         Else: cmbSTATUS.Text = "Cancelado"
      End If
   End If

   txtCPF.PromptInclude = False
   txtCPF.Text = TabVENDEDOR.Fields("cnpjcpf").Value
   CRITERIO_A = txtCPF.Text

   If Not IsNull(TabVENDEDOR!PERC_COMISSAO) Then _
      txtPerc.Text = TabVENDEDOR!PERC_COMISSAO

   txtCPF.PromptInclude = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VEND"
End Sub

Private Sub LIMPA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   lstEstab.ListItems.Clear

   SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      Set item = lstEstab.ListItems.Add(, "seq." & TabDESCR.Fields(0).Value, Trim(TabDESCR.Fields(1).Value))
      item.SubItems(1) = "" & TabDESCR.Fields(0).Value

      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   txtIBGE.Text = ""
   txtNumeroR.Text = ""
   txtPerc.Text = ""
   optValor.Value = False
   txtDtCad.PromptInclude = False
   txtDtCad.Text = ""
   txtDtBaixa.PromptInclude = False
   txtDtBaixa.Text = ""
   txtDtNasc.PromptInclude = False
   txtDtNasc.Text = ""
   cmbSTATUS.Text = "Ativo"

   DT_EXP_D = 0
   VENDEDOR_ID_N = 0
   CODG_EQUIPE_N = 0
   NOME_A = ""
   DATA_INI = 0
   STATUS_A = ""
   EMAIL_A = "   "
   CRITERIO_A = ""
   txtCodg.Text = ""
   txtNome.Text = ""
   cmbTbPreco.Text = ""
   cmbTbPrecoAUX.Text = ""
   cmbEquipeAUX.Text = ""
   cmbEquipe.Text = ""
   txtCPF.PromptInclude = False
   txtCPF.Text = ""
   txtCep.Text = ""
   txtRua.Text = ""
   txtComp.Text = ""
   txtBairro.Text = ""
   txtCidade.Text = ""
   txtUF.Text = ""
   LIMPA_FONE
   SETA_FONE

   txtCodg.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_VENDEDOR"
End Sub

Sub GRAVA_VENDEDOR()
'On Error GoTo ERRO_TRATA

'=========================
   PESSOA_ID_N = 0
'=========================PESSOA
   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA "
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

   spPessoa CONT_N, PESSOA_ID_N, Trim(txtCPF.Text), Trim(txtNome.Text), "", Left(cmbSTATUS.Text, 1)

   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA "
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close
'=========================

   If Trim(cmbTbPreco.Text) = "" Then _
      cmbTbPrecoAUX.Text = "NULL"

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select * from vwVendedor "
   SQL = SQL & " where vendedor_id = " & txtCodg.Text
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVENDEDOR.EOF Then
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      SQL = "update VENDEDOR set "
      SQL = SQL & " EQUIPE_ID = " & cmbEquipeAUX.Text
      SQL = SQL & " ,Status = '" & Left(cmbSTATUS.Text, 1) & "'"
      SQL = SQL & " ,DT_NASCIMENTO = '" & DMA(txtDtNasc.Text) & "'"
      SQL = SQL & " ,DT_BAIXA = '" & DMA(txtDtBaixa.Text) & "'"
      SQL = SQL & " ,TIPO_COMIS = 0"
      SQL = SQL & " ,Categoria = 0"
      SQL = SQL & " ,PERC_COMISSAO = 0" & Replace(txtPerc.Text, ",", ".")
      SQL = SQL & " ,TABELAPRECO_ID = " & cmbTbPrecoAUX.Text
      SQL = SQL & " where vendedor_id = " & txtCodg.Text
      Else
         If TabVENDEDOR.State = 1 Then _
            TabVENDEDOR.Close

         SQL = "insert into VENDEDOR "
            SQL = SQL & " (VENDEDOR_ID,PESSOA_ID,EQUIPE_ID,"
            SQL = SQL & " STATUS,DT_NASCIMENTO,DT_BAIXA,TIPO_COMIS,"
            SQL = SQL & " CATEGORIA,PERC_COMISSAO,TABELAPRECO_ID) "
         SQL = SQL & " VALUES( "
            SQL = SQL & txtCodg.Text                              'VENDEDOR_ID
            SQL = SQL & "," & PESSOA_ID_N                         'PESSOA_ID
            SQL = SQL & "," & cmbEquipeAUX.Text                   'EQUIPE_ID
            SQL = SQL & ",'" & Left(cmbSTATUS.Text, 1) & "'"      'STATUS
            SQL = SQL & ",'" & DMA(txtDtNasc.Text) & "'"          'DT_NASCIMENTO
            SQL = SQL & ",'" & DMA(txtDtBaixa.Text) & "'"         'DT_BAIXA
            SQL = SQL & "," & 0                                   'TIPO_COMIS
            SQL = SQL & "," & 0                                   'CATEGORIA
            SQL = SQL & ",0" & Replace(txtPerc.Text, ",", ".")    'PERC_COMISSAO
            SQL = SQL & "," & cmbTbPrecoAUX.Text                  'TABELAPRECO_ID
         SQL = SQL & ")"
   End If
   CONECTA_RETAGUARDA.Execute SQL

   Dim i

   If Trim(txtCodg.Text) <> "" Then
      If IsNumeric(txtCodg.Text) Then
         SQL = "delete from ESTABVENDEDOR "
         SQL = SQL & " where vendedor_id = " & txtCodg.Text
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If lstEstab.ListItems.Count > 0 Then
      For i = lstEstab.ListItems.Count To 1 Step -1
         If lstEstab.ListItems(i).Checked = True Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from ESTABVENDEDOR WITH (NOLOCK)"
            SQL = SQL & " where estabelecimento_id = " & lstEstab.ListItems(i).SubItems(1)
            SQL = SQL & " and vendedor_id = " & txtCodg.Text
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabConsulta.EOF Then
               SQL = "insert into ESTABVENDEDOR "
               SQL = SQL & "values("
                  SQL = SQL & lstEstab.ListItems(i).SubItems(1)
                  SQL = SQL & "," & txtCodg.Text
               SQL = SQL & " )"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      Next i
   End If

'ENDEREÇO RESIDENCIAL
   If Not IsNumeric(txtIBGE.Text) Then _
      txtIBGE.Text = "5201211"

   txtCep.PromptInclude = False
   If Trim(txtCep.Text) <> "" Or Trim(txtRua.Text) <> "" Or Trim(txtBairro.Text) <> "" Or Trim(txtComp.Text) <> "" Then
      If txtCep.Text <> "" Then _
         SP_GRAVA_CEP txtCep.Text, txtCidade.Text, txtUF.Text, txtIBGE.Text

      sp_Grava_Endereco txtCep.Text, txtRua.Text, txtBairro.Text, txtComp.Text, "R", txtNumeroR.Text
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_VENDEDOR"
End Sub

Sub MATAR_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodg.Text) = "" Then
      MsgBox "Informe codigo."
      txtCodg.SetFocus
      Exit Sub
   End If

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select * from vwVendedor "
   SQL = SQL & " where vendedor_id = " & txtCodg.Text
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVENDEDOR.EOF Then
      TabVENDEDOR.Close
      Msg = "Confirma desativação?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         SQL = "update VENDEDOR set "
         SQL = SQL & " status = 'C' "
         SQL = SQL & " where vendedor_id = " & txtCodg.Text
         CONECTA_RETAGUARDA.Execute SQL

         LIMPA_VENDEDOR
      End If
      Else
         TabVENDEDOR.Close
         MsgBox "Registro não encontrado."
   End If
   txtCodg.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATAR_VENDEDOR"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   txtCPF.PromptInclude = False

   ADOCabeca.Enabled = True
   ADOCabeca.ConnectionString = AUTENTICA_GRID

   SQL = "select * from FONE"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " order by NUMERO"

   ADOCabeca.RecordSource = SQL
   ADOCabeca.Enabled = True
   ADOCabeca.Refresh

   txtCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcpf_KeyDown"
End Sub

Private Sub LIMPA_FONE()
'On Error GoTo ERRO_TRATA

   txtN.Text = ""
   txtDDD.Text = ""
   txtL.Text = ""
   txtL.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_FONE"
End Sub

Sub GERA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   txtCodg.Text = MAX_ID("vendedor_id", "vendedor", "", "", "", "")
   txtCodg.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_VENDEDOR"
End Sub

Sub PROCURA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   If Trim(txtCPF.Text) <> "" Then
      SQL = "select * from vwVENDEDOR "
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCPF.Text) & "'"
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         txtCodg.Text = TabVENDEDOR.Fields("vendedor_id").Value
         VENDEDOR_ID_N = txtCodg.Text

         MOSTRA_VEND

         'ENDEREÇO
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from ENDERECO"
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         SQL = SQL & " and tipo = 'R' "
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If Not IsNull(TabConsulta!Rua) Then _
               txtRua.Text = TabConsulta!Rua
            If Not IsNull(TabConsulta!Bairro) Then _
               txtBairro.Text = TabConsulta!Bairro
            If Not IsNull(TabConsulta!Complemento) Then _
               txtComp.Text = TabConsulta!Complemento
            If Not IsNull(TabConsulta!CEP_ID) Then _
               txtCep.Text = TabConsulta!CEP_ID

            If TabCEP.State = 1 Then _
               TabCEP.Close

            SQL = "select * from CEP "
            SQL = SQL & " where cep_ID = '" & Trim(TabConsulta.Fields("cep_id").Value) & "'"
            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCEP.EOF Then
               If Not IsNull(TabCEP.Fields("Cep_id").Value) Then _
                  txtCep.Text = TabCEP.Fields("Cep_id").Value

               If Not IsNull(TabCEP!CIDADE) Then _
                  txtCidade.Text = TabCEP!CIDADE

               If Not IsNull(TabCEP!UF) Then _
                  txtUF.Text = TabCEP!UF
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SETA_FONE
      End If
   End If
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_VENDEDOR"
End Sub
