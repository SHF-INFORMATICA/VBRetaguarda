VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCADASTROTRANS 
   Caption         =   "Cadastro de Transportadora"
   ClientHeight    =   7755
   ClientLeft      =   2085
   ClientTop       =   2700
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROTRANS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcluirFone 
      Height          =   375
      Left            =   10200
      Picture         =   "CADASTROTRANS.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5985
      Width           =   375
   End
   Begin VB.TextBox txtL 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   20
      Top             =   6000
      Width           =   6615
   End
   Begin VB.TextBox txtN 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtDDD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      MaxLength       =   2
      TabIndex        =   18
      Top             =   6000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   " Endereço Cobrança "
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   0
      TabIndex        =   35
      Top             =   4320
      Width           =   10935
      Begin VB.CommandButton cmdCopEnd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copiar Endereço Comercial"
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Clique aqui para copiar o endereço pessoal para o endereço comercial."
         Top             =   1150
         Width           =   3105
      End
      Begin VB.TextBox txtIbge 
         Alignment       =   1  'Right Justify
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   720
         MaxLength       =   80
         TabIndex        =   45
         Top             =   1150
         Width           =   1455
      End
      Begin VB.TextBox txtRuaB 
         DataField       =   "Endereco_Res"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   13
         Top             =   350
         Width           =   4575
      End
      Begin VB.TextBox txtBaIrroB 
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   720
         MaxLength       =   80
         TabIndex        =   15
         Top             =   750
         Width           =   2535
      End
      Begin VB.TextBox txtCidadeB 
         DataField       =   "Cidade"
         Height          =   345
         Left            =   4680
         MaxLength       =   80
         TabIndex        =   16
         Top             =   750
         Width           =   4575
      End
      Begin VB.TextBox txtUFB 
         Alignment       =   2  'Center
         DataField       =   "Estado"
         Height          =   345
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   17
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox txtEndB 
         Height          =   345
         Left            =   8400
         MaxLength       =   80
         TabIndex        =   14
         Top             =   350
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtCepB 
         Height          =   345
         Left            =   720
         TabIndex        =   12
         Top             =   350
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
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
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IBGE:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Rua:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1920
         TabIndex        =   41
         Top             =   380
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3960
         TabIndex        =   39
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   255
         TabIndex        =   38
         Top             =   380
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7200
         TabIndex        =   37
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   225
         Left            =   9840
         TabIndex        =   36
         Top             =   780
         Width           =   255
      End
   End
   Begin VB.Frame FraPessoa 
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   0
      TabIndex        =   28
      Top             =   720
      Width           =   10935
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
         Left            =   9360
         TabIndex        =   55
         Top             =   240
         Width           =   1455
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
         Left            =   9480
         TabIndex        =   54
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2640
         Picture         =   "CADASTROTRANS.frx":6A53
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   405
      End
      Begin VB.ComboBox cmbAuxProf 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6480
         TabIndex        =   51
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   9960
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
               Picture         =   "CADASTROTRANS.frx":7455
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1080
         TabIndex        =   47
         Top             =   240
         Width           =   1425
      End
      Begin VB.ComboBox cmbProf 
         Height          =   360
         Left            =   6480
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtRazao 
         DataField       =   "Nome"
         Height          =   360
         Left            =   4800
         MaxLength       =   80
         TabIndex        =   1
         Top             =   700
         Width           =   6015
      End
      Begin VB.TextBox txtFant 
         DataField       =   "Nome"
         Height          =   360
         Left            =   4800
         MaxLength       =   80
         TabIndex        =   3
         Top             =   1080
         Width           =   6015
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   345
         Left            =   1080
         TabIndex        =   0
         Top             =   700
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
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
      Begin MSMask.MaskEdBox txtIE 
         Height          =   345
         Left            =   1920
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   25
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
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblInsc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estatual:"
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
         TabIndex        =   34
         Top             =   1560
         Width           =   1725
      End
      Begin VB.Label lblProf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ramo de Atividade:"
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
         Left            =   4560
         TabIndex        =   33
         Top             =   1560
         Width           =   1875
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*CNPJ:"
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
         Left            =   330
         TabIndex        =   32
         Top             =   700
         Width           =   645
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Razão Social:"
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
         Left            =   3300
         TabIndex        =   31
         Top             =   700
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia:"
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
         Left            =   3210
         TabIndex        =   30
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   330
         TabIndex        =   29
         Top             =   1080
         Width           =   645
      End
   End
   Begin VB.Frame fraCom 
      Caption         =   " Endereço Comercial "
      ForeColor       =   &H00400000&
      Height          =   1335
      Left            =   0
      TabIndex        =   21
      Top             =   2880
      Width           =   10935
      Begin VB.TextBox txtEndC 
         Height          =   345
         Left            =   8400
         MaxLength       =   80
         TabIndex        =   8
         Top             =   350
         Width           =   2415
      End
      Begin VB.TextBox txtUFC 
         Alignment       =   2  'Center
         DataField       =   "Estado_Com"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   11
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtCidadeC 
         DataField       =   "Cidade_Com"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   4680
         MaxLength       =   80
         TabIndex        =   10
         Top             =   780
         Width           =   4575
      End
      Begin VB.TextBox txtBairroC 
         DataField       =   "Bairro_Com"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   720
         MaxLength       =   80
         TabIndex        =   9
         Top             =   780
         Width           =   2535
      End
      Begin VB.TextBox txtRuaC 
         DataField       =   "Endereco_Com"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   7
         Top             =   350
         Width           =   4575
      End
      Begin MSMask.MaskEdBox txtCepC 
         Height          =   345
         Left            =   720
         TabIndex        =   6
         Top             =   350
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7200
         TabIndex        =   27
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label lblEstadoCom 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9840
         TabIndex        =   26
         Top             =   780
         Width           =   255
      End
      Begin VB.Label lblCepCom 
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   380
         Width           =   345
      End
      Begin VB.Label lblCidadeCom 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3960
         TabIndex        =   24
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblBairroCom 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   780
         Width           =   570
      End
      Begin VB.Label lblRuaCom 
         AutoSize        =   -1  'True
         Caption         =   "Rua:"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1920
         TabIndex        =   22
         Top             =   380
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":82A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":86FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":8A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":8E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":92BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":95DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":9A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTRANS.frx":9D52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10320
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
               Picture         =   "CADASTROTRANS.frx":A1A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTRANS.frx":B340
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTRANS.frx":C3CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTRANS.frx":D637
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTRANS.frx":ED34
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTRANS.frx":FE3F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTRANS.frx":10DF4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexTel 
      Height          =   1425
      Left            =   0
      TabIndex        =   49
      Top             =   6360
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   2514
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   16711680
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      DesignHeight    =   7755
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Local:"
      Height          =   225
      Index           =   13
      Left            =   2880
      TabIndex        =   44
      Top             =   6030
      Width           =   480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "DDD:"
      Height          =   225
      Left            =   240
      TabIndex        =   43
      Top             =   6030
      Width           =   405
   End
End
Attribute VB_Name = "frmCADASTROTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   cmbStatus.Clear
   cmbStatus.AddItem "Ativo"
   cmbStatus.AddItem "Inativo"
   cmbProf.Clear
   cmbAuxProf.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & "where TIPO = 'E' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbAuxProf.AddItem TabDESCR!codigo
      cmbProf.AddItem Trim(TabDESCR!DESCRICAO)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   GERA_TRANSP

   Call CentralizaJanela(frmCADASTROTRANS)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF10
         GRAVA_TUDO
         LIMPA_TUDO
         GERA_TRANSP
         txtCNPJCPF.SetFocus
      Case vbKeyEscape
         Unload Me
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "gravar"
         GRAVA_TUDO
         LIMPA_TUDO
         GERA_TRANSP
      Case "matar"
         txtCNPJCPF.PromptInclude = False
         If txtCNPJCPF.Text <> "" Then
            Msg = "Confirma exclusão do TRANSPORTADORA ?"
            PERGUNTA Msg, vbYesNo + 32, "Cadastro Transportadora NFE", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "update TRANSPORTADORA set "
               SQL = SQL & "status='C'"
               SQL = SQL & "where CGCCPF='" & Trim(txtCNPJCPF.Text) & "'"
               CONECTA_RETAGUARDA.Execute SQL
               LIMPA_TUDO
               txtCNPJCPF.SetFocus
            End If
         End If
         txtCNPJCPF.PromptInclude = True
      Case "print"
         REL_TRANSP
      Case "limpar"
         LIMPA_TUDO
         GERA_TRANSP
         txtCNPJCPF.SetFocus
      Case "consultar"
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmConsultaPessoa.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
            txtCNPJCPF.PromptInclude = True
            PROCURA_DADOS
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdEmail_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then
      CNPJCPF_A = Trim(txtCNPJCPF.Text)
      frmEmail.Show 1
   End If
   txtCNPJCPF.PromptInclude = True
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdEmail_Click"
End Sub

Private Sub cmdCopEnd_Click()
   txtCepC.PromptInclude = False
   txtCepB.PromptInclude = False
   txtCepB.Text = "" & txtCepC.Text
   txtRuaB.Text = "" & txtRuaC.Text
   txtEndB.Text = "" & txtEndC.Text
   txtBaIrroB.Text = "" & txtBairroC.Text
   txtCidadeB.Text = "" & txtCidadeC.Text
   txtUFB.Text = "" & txtUFC.Text
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   frmConsultaPessoa.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      txtCNPJCPF.PromptInclude = True
      PROCURA_DADOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub

Private Sub cmdExcluirFone_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" And txtN.Text <> "" Then
      Dim Achou As Boolean
      Achou = False
      Select Case Button.key
         Case "matar"
            FlexTel.Col = 1
            For i = 1 To FlexTel.Rows - 1
               If Replace(txtN.Text, "-", "") = FlexTel.TextMatrix(i, 1) Then
                  Achou = True
                  Exit For
               End If
            Next
            If Achou = True Then
               If FlexTel.Rows > 2 Then
                  FlexTel.RemoveItem (i)
                  Else
                     FlexTel.AddItem ""
                     FlexTel.RemoveItem (i)
                     refresh_GRID
               End If
            End If
            txtCNPJCPF.PromptInclude = True
      End Select
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdExcluirFone_ButtonClick"
End Sub

Private Sub cmbProf_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxProf.ListIndex = cmbProf.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbProf_Click"
End Sub

Private Sub txtCepb_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCepB.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepb_GotFocus"
End Sub

Private Sub txtCepb_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCEP.Show 1
         txtCepB.PromptInclude = False
         txtCepB.Text = CRITERIO
         txtCepB.PromptInclude = True
      Case vbKeyF7
         frmCONSULTACEP.Show 1
         txtCepB.PromptInclude = False
         txtCepB.Text = CRITERIO
         txtCepB.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepb_KeyDown"
End Sub

Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.Mask = "##############"
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_GotFocus"
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
        txtCNPJCPF.PromptInclude = False
        If txtCNPJCPF.Text = "" Then
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
                         MsgBox "CGC com DV incorreto !!! "
                         txtCNPJCPF.PromptInclude = False
                         txtCNPJCPF = ""
                         txtCNPJCPF.SetFocus
                         Exit Sub
                      End If
                    Case Is > 14
                       MsgBox "CGC/CPF com DV incorreto !!! "
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    Case Is < 11
                       MsgBox "CGC/CPF com DV incorreto !!! "
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                 End Select
                 Else
                    MsgBox "CGC/CPF com DV incorreto !!! "
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

      PROCURA_DADOS

      txtRazao.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub TXTRAZAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbStatus.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTRAZAO_KeyPress"
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtFant.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbStatus_KeyPress"
End Sub

Private Sub txtFant_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtIE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFant_KeyPress"
End Sub

Private Sub cmbprof_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbprof_KeyPress"
End Sub

Private Sub txtIE_GotFocus()
'On Error GoTo ERRO_TRATA

   txtIE.PromptInclude = False
   If txtIE.Text = "" Then
      'txtIE.Mask = "##.###.###-#"
      'Else: txtIE.Text = Format(txtIE.Text, "##.###.###-#")
   End If
   txtIE.PromptInclude = True

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe Inscrição Estadual"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" And Trim(txtRazao.Text) <> "" Then
      frmINICIO.BARI.Panels.Add (3)
      frmINICIO.BARI.Panels(3).Text = "F10 - Gravar"
      frmINICIO.BARI.Panels(3).AutoSize = sbrContents
   End If
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIE_GotFocus"
End Sub

Private Sub txtie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbProf.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtie_KeyPress"
End Sub

Private Sub txtcepc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepC.PromptInclude = False
      If txtCepC.Text <> "" Then
         SP_PROCURA_CEP txtCepC.Text
         If TabCEP.EOF Then
            TabCEP.Close
            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepC.SetFocus
            Exit Sub
            Else
               txtCidadeC.Text = TabCEP!Cidade
               txtUFC.Text = TabCEP!UF
         End If
         TabCEP.Close
      End If
      txtRuaC.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcepc_KeyPress"
End Sub

Private Sub txtCepc_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         frmCADASTROCEP.Show 1
         txtCepC.PromptInclude = False
         txtCepC.Text = CRITERIO
         txtCepC.PromptInclude = True
      Case vbKeyF7
         frmCONSULTACEP.Show 1
         txtCepC.PromptInclude = False
         txtCepC.Text = CRITERIO
         txtCepC.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCepc_KeyDown"
End Sub

Private Sub txtruac_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtruac_KeyPress"
End Sub

Private Sub txtendc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairroC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtendc_KeyPress"
End Sub

Private Sub txtbairroc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidadeC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbairroc_KeyPress"
End Sub

Private Sub txtcidadec_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUFC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidadec_KeyPress"
End Sub

Private Sub txtufc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtufc_KeyPress"
End Sub

Private Sub txtcepb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCepB.PromptInclude = False
      If txtCepB.Text <> "" Then
         SP_PROCURA_CEP txtCepB.Text
         If TabCEP.EOF Then
            TabCEP.Close
            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepB.SetFocus
            Exit Sub
            Else
               txtCidadeB.Text = TabCEP!Cidade
               txtUFB.Text = TabCEP!UF
         End If
         TabCEP.Close
      End If
      txtRuaB.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcepb_KeyPress"
End Sub

Private Sub txtruab_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtEndB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtruab_KeyPress"
End Sub

Private Sub txtendb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBaIrroB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtendb_KeyPress"
End Sub

Private Sub txtbairrob_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCidadeB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbairrob_KeyPress"
End Sub

Private Sub txtcidadeb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUFB.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcidadeb_KeyPress"
End Sub

Private Sub txtufb_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtufb_KeyPress"
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
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
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtN_KeyPress"
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      If txtN.Text <> "" And txtCNPJCPF.Text <> "" Then _
         GRAVA_FONE_TEMP
      txtCNPJCPF.PromptInclude = True
      txtN.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtl_KeyPress"
End Sub
'=======================================================
Private Sub PROCURA_DADOS()
'On Error GoTo ERRO_TRATA

   LIMPA_QUASE_TUDO
   PESSOA_ID_N = 0
   txtCNPJCPF.PromptInclude = False

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select razao,descricao,pessoa_id from PESSOA "
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtRazao.Text = "" & Trim(TabTemp.Fields("razao").Value)
      txtFant.Text = "" & Trim(TabTemp.Fields("descricao").Value)
      PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close
         MsgBox "Cadastro tabela pessoa não encontrado."
         Exit Sub
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TRANSPORTADORA "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtCodigo.Text = TabTemp!TRANSP_ID
      If Not IsNull(TabTemp!Status) Then
         If TabTemp!Status = "A" Then
            cmbStatus.Text = "Ativo"
            Else: cmbStatus.Text = "Inativo"
         End If
      End If
      txtIE.Text = TabTemp!IE
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CRITERIO = txtCNPJCPF.Text
   SETA_FONE
   
   'ENDEREÇO COMERCIAL
'ok
   BUSCA_ENDERECO_PESSOA "C", ""
   If Not tabEndereco.EOF Then
      If Not IsNull(tabEndereco!Rua) Then _
         txtRuaC.Text = tabEndereco!Rua
      If Not IsNull(tabEndereco!Bairro) Then _
         txtBairroC.Text = tabEndereco!Bairro
      If Not IsNull(tabEndereco!Complemento) Then _
         txtEndC.Text = tabEndereco!Complemento
      If Not IsNull(tabEndereco!CEP) Then
         If tabEndereco!CEP <> "" Then
            txtCepC.Text = tabEndereco!CEP

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from CEP "
            SQL = SQL & "where cep = '" & tabEndereco!CEP & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCidadeC.Text = TabConsulta!Cidade
               txtUFC.Text = TabConsulta!UF
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   'ENDEREÇO COBRANÇA
'ok
   BUSCA_ENDERECO_PESSOA "B", ""
   If Not tabEndereco.EOF Then
      If Not IsNull(tabEndereco!Rua) Then txtRuaB.Text = tabEndereco!Rua
      If Not IsNull(tabEndereco!Bairro) Then txtBaIrroB.Text = tabEndereco!Bairro
      If Not IsNull(tabEndereco!Complemento) Then txtEndB.Text = tabEndereco!Complemento
      If Not IsNull(tabEndereco!CEP) Then
         If tabEndereco!CEP <> "" Then
            txtCepB.Text = tabEndereco!CEP
            SQL = "select * from CEP "
            SQL = SQL & "where cep = '" & tabEndereco!CEP & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCidadeB.Text = TabConsulta!Cidade
               txtUFB.Text = TabConsulta!UF
            End If
            TabConsulta.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_DADOS"
End Sub

Private Sub LIMPA_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   txtRazao.Text = ""
   cmbStatus.Text = ""
   cmbProf.Text = ""
   txtIE.PromptInclude = False
   txtIE.Text = ""
   txtCepC.PromptInclude = False
   txtCepC.Text = ""
   txtRuaC.Text = ""
   txtEndC.Text = ""
   txtBairroC.Text = ""
   txtCidadeC.Text = ""
   txtUFC.Text = ""
   txtCepB.PromptInclude = False
   txtCepB.Text = ""
   txtRuaB.Text = ""
   txtEndB.Text = ""
   txtBaIrroB.Text = ""
   txtCidadeB.Text = ""
   txtUFB.Text = ""
   txtFant.Text = ""
   LIMPA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_QUASE_TUDO"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtRazao.Text = ""
   cmbStatus.Text = ""
   cmbProf.Text = ""
   cmbAuxProf.Text = ""
   txtIE.PromptInclude = False
   txtIE.Text = ""
   txtCepC.PromptInclude = False
   txtCepC.Text = ""
   txtRuaC.Text = ""
   txtEndC.Text = ""
   txtBairroC.Text = ""
   txtCidadeC.Text = ""
   txtUFC.Text = ""
   txtCepB.PromptInclude = False
   txtCepB.Text = ""
   txtRuaB.Text = ""
   txtEndB.Text = ""
   txtBaIrroB.Text = ""
   txtCidadeB.Text = ""
   txtUFB.Text = ""
   txtFant.Text = ""
   LIMPA_FONE
   CRITERIO = 0
   SETA_FONE
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe CPF do TRANSPORTADORA."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If Trim(txtRazao.Text) = "" Then
      MsgBox "Informe Nome do TRANSPORTADORA."
      txtRazao.SetFocus
      Exit Sub
   End If
'=========================PESSOA
   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close
   SQL = "select pessoa_id from PESSOA "
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close

   'executa stored procedure sp_pessoa
   CONT_N = 1
   If PESSOA_ID_N > 0 Then _
      CONT_N = 2

   SP_PESSOA CONT_N, PESSOA_ID_N, Trim(txtCNPJCPF.Text), Trim(txtFant.Text), Trim(txtRazao.Text), Left(cmbStatus.Text, 1)

   PESSOA_ID_N = 0
   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select pessoa_id from PESSOA "
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then _
      If Not IsNull(TabCliente.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabCliente.Fields(0).Value)
   If TabCliente.State = 1 Then _
      TabCliente.Close
'=========================

   SQL = "select * from TRANSPORTADORA "
   SQL = SQL & "where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      TRANSP_ID = TabCliente!TRANSP_ID
      SQL = "UPDATE TRANSPORTADORA SET NOME = '" & txtRazao.Text & "', IE = '" & txtIE.Text & "', RAZAO_SOCIAL = '" & txtFant.Text & "', STATUS = '" & Left(cmbStatus.Text, 1) & "'"
      SQL = SQL & "where TRANSP_ID = " & TRANSP_ID
      Else
         SQL = "INSERT INTO TRANSPORTADORA (TRANSP_ID, DT_CAD, CGCCPF, NOME, IE, RAZAO_SOCIAL, STATUS, EMPRESA_ID,pessoa_id) "
         SQL = SQL & " VALUES (" & TRANSP_ID & ",'" & Now & "','" & txtCNPJCPF.Text & "','" & txtRazao.Text & "','" & txtIE.Text & "','" & txtFant.Text & "','" & Left(cmbStatus.Text, 1) & "'," & EMPRESA_ID_N & "," & PESSOA_ID_N & ")"
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

   CONECTA_RETAGUARDA.Execute SQL

   If txtRuaC.Text <> "" And txtBairroC.Text <> "" Then
      'ENDEREÇO COMERCIAL
      txtIBGE.Text = 0
      txtCepC.PromptInclude = False
      If txtCepC.Text <> "" Or txtRuaC.Text <> "" Or txtBairroC.Text <> "" Or txtEndC.Text <> "" Then
         If txtCepC.Text <> "" Then _
            SP_GRAVA_CEP txtCepC.Text, txtCidadeC.Text, txtUFC.Text, txtIBGE.Text
   
         sp_Grava_Endereco txtCepC.Text, txtRuaC.Text, txtBairroC.Text, txtEndC.Text, "C", 0
         Else: SP_MATA_ENDEREÇO "C"
      End If
   End If
   If txtRuaB.Text <> "" And txtBaIrroB.Text <> "" Then
      'ENDEREÇO COBRANÇA
      txtCepB.PromptInclude = False
      If txtCepB.Text <> "" Or txtRuaB.Text <> "" Or txtBaIrroB.Text <> "" Or txtEndB.Text <> "" Then
         If txtCepB.Text <> "" Then _
            SP_GRAVA_CEP txtCepB.Text, txtCidadeB.Text, txtUFB.Text, txtIBGE.Text
   
         sp_Grava_Endereco txtCepB.Text, txtRuaB.Text, txtBaIrroB.Text, txtEndB.Text, "B", 0
         Else: SP_MATA_ENDEREÇO "B"
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub
'============fone
Private Sub TOOBARFONE_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         txtCNPJCPF.PromptInclude = False
         If txtCNPJCPF.Text <> "" And txtN.Text <> "" Then
            SQL = "delete  from FONE "
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and numero = '" & txtN.Text & "'"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         LIMPA_FONE
         CRITERIO = txtCNPJCPF.Text
         SETA_FONE
         txtCNPJCPF.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOOBARFONE_ButtonClick"
End Sub

Private Sub GRAVA_FONE_TEMP()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False

   GRAVA_TUDO

   If TabFone.State = 1 Then _
      TabFone.Close

   SQL = "select * from FONE "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and numero = '" & txtN.Text & "'"
   TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabFone.EOF Then
      SQL = "INSERT INTO FONE "
      SQL = SQL & " (prop, Numero, ddd, local,pessoa_id ) "
      SQL = SQL & " VALUES ("
      SQL = SQL & "'" & Trim(txtCNPJCPF.Text) & "'"
      SQL = SQL & "," & txtN.Text & ""
      SQL = SQL & "," & txtDDD.Text
      SQL = SQL & ",'" & txtL.Text & "'"
      SQL = SQL & "," & PESSOA_ID_N
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "UPDATE FONE SET Numero = '" & txtN.Text & "', ddd = " & txtDDD.Text & ", local = '" & txtL.Text & "'"
         SQL = SQL & " where prop='" & txtCNPJCPF.Text & "' and numero = '" & txtN.Text & "'"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabFone.State = 1 Then _
      TabFone.Close

   CRITERIO = txtCNPJCPF.Text
   SETA_FONE
   LIMPA_FONE
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_FONE_TEMP"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   SQL = "select * from FONE "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   If TabAUX.State = 1 Then TabAUX.Close
   TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   refresh_GRID
   Do While Not TabAUX.EOF
      FlexTel.AddItem ""
      FlexTel.Row = FlexTel.Rows - 1
      FlexTel.Col = 0
      FlexTel.Text = TabAUX!DDD & ""
      FlexTel.Col = 1
      FlexTel.Text = TabAUX!Numero
      FlexTel.Col = 2
      FlexTel.Text = TabAUX!local & ""
      FlexTel.Col = 3
      txtCNPJCPF.PromptInclude = False
      FlexTel.Text = txtCNPJCPF.Text
      txtCNPJCPF.PromptInclude = True
      TabAUX.MoveNext
   Loop

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_FONE"
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

Public Sub GERA_TRANSP()
'On Error GoTo ERRO_TRATA

   SQL = "select max(transp_id) from TRANSPORTADORA "
   If TabEmpresa.State = 1 Then TabEmpresa.Close
   TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmpresa.EOF Then _
      If Not IsNull(TabEmpresa.Fields(0).Value) Then _
         TRANSP_ID = TabEmpresa.Fields(0).Value + 1
         txtCodigo.Text = TRANSP_ID
   TabEmpresa.Close
   If TRANSP_ID = 0 Then
      TRANSP_ID = 1
      txtCodigo.Text = TRANSP_ID
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_TRANSP"
End Sub

Private Sub refresh_GRID()
'On Error GoTo ERRO_TRATA

   FlexTel.Clear
   FlexTel.Row = 0
   FlexTel.Col = 0: FlexTel.ColWidth(0) = (FlexTel.Width / 8) - 100: FlexTel.Text = "DDD": FlexTel.ColAlignment(0) = 3
   FlexTel.Col = 1: FlexTel.ColWidth(1) = FlexTel.Width / 4: FlexTel.Text = "NÚMERO": FlexTel.ColAlignment(1) = 1
   FlexTel.Col = 2: FlexTel.ColWidth(2) = FlexTel.Width / 1.65: FlexTel.Text = "LOCAL"
   FlexTel.Col = 3: FlexTel.ColWidth(3) = 0: FlexTel.Text = "CNPJCPF"
   FlexTel.Rows = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "refresh_GRID"
End Sub

Sub REL_TRANSP()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{transportadora.empresa_id} = " & EMPRESA_ID_N
   If IsNumeric(txtCNPJCPF.Text) Then
      txtCNPJCPF.PromptInclude = False
      FORMULA_REL = FORMULA_REL & " and {transportadora.cgccpf} = '" & Trim(txtCNPJCPF.Text) & "'"
   End If

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_Transp.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REL_FORNECEDOR"
End Sub
