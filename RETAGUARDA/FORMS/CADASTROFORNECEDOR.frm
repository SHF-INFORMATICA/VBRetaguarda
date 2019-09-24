VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCADASTROFORNECEDOR 
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   7680
   ClientLeft      =   2085
   ClientTop       =   3135
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROFORNECEDOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   0
      TabIndex        =   50
      Top             =   5640
      Width           =   10935
      Begin VB.CommandButton cmdExcluirFone 
         Height          =   375
         Left            =   10440
         Picture         =   "CADASTROFORNECEDOR.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox txtL 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   23
         Top             =   200
         Width           =   6975
      End
      Begin VB.TextBox txtN 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   22
         Top             =   200
         Width           =   1335
      End
      Begin VB.TextBox txtDDD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   600
         MaxLength       =   2
         TabIndex        =   21
         Top             =   200
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid FlexTel 
         Height          =   1305
         Left            =   0
         TabIndex        =   24
         Top             =   600
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   2302
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
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Local:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   13
         Left            =   2880
         TabIndex        =   52
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "DDD:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame fraCom 
      Caption         =   " Endereço Comercial "
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   0
      TabIndex        =   36
      Top             =   2760
      Width           =   10935
      Begin VB.TextBox txtIbge1 
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   9345
         MaxLength       =   80
         TabIndex        =   59
         Top             =   800
         Width           =   1455
      End
      Begin VB.TextBox txtRuaC 
         DataField       =   "Endereco_Com"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   8
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtBairroC 
         DataField       =   "Bairro_Com"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   840
         MaxLength       =   80
         TabIndex        =   10
         Top             =   800
         Width           =   1935
      End
      Begin VB.TextBox txtCidadeC 
         DataField       =   "Cidade_Com"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   3600
         MaxLength       =   80
         TabIndex        =   11
         Top             =   800
         Width           =   3615
      End
      Begin VB.TextBox txtUFC 
         Alignment       =   2  'Center
         DataField       =   "Estado_Com"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   12
         Top             =   800
         Width           =   615
      End
      Begin VB.TextBox txtEndC 
         Height          =   360
         Left            =   8760
         MaxLength       =   80
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtCepC 
         Height          =   360
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
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
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IBGE:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   8760
         TabIndex        =   60
         Top             =   840
         Width           =   525
      End
      Begin VB.Label lblRuaCom 
         AutoSize        =   -1  'True
         Caption         =   "Rua:"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2160
         TabIndex        =   42
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lblBairroCom 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   830
         Width           =   570
      End
      Begin VB.Label lblCidadeCom 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2880
         TabIndex        =   40
         Top             =   825
         Width           =   615
      End
      Begin VB.Label lblCepCom 
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   360
         TabIndex        =   39
         Top             =   390
         Width           =   345
      End
      Begin VB.Label lblEstadoCom 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   7320
         TabIndex        =   38
         Top             =   825
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7440
         TabIndex        =   37
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame FraPessoa 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   0
      TabIndex        =   32
      Top             =   720
      Width           =   10935
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2790
         Picture         =   "CADASTROFORNECEDOR.frx":6A53
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   360
         Width           =   405
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   10320
         Top             =   -240
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
               Picture         =   "CADASTROFORNECEDOR.frx":7455
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmbAuxProf 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   8160
         TabIndex        =   47
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtFant 
         DataField       =   "Nome"
         Height          =   360
         Left            =   4680
         MaxLength       =   80
         TabIndex        =   3
         Top             =   720
         Width           =   6135
      End
      Begin VB.TextBox txtRazao 
         DataField       =   "Nome"
         Height          =   360
         Left            =   4680
         MaxLength       =   80
         TabIndex        =   1
         Top             =   360
         Width           =   6135
      End
      Begin VB.ComboBox cmbEmail 
         Height          =   360
         ItemData        =   "CADASTROFORNECEDOR.frx":82A6
         Left            =   120
         List            =   "CADASTROFORNECEDOR.frx":82A8
         TabIndex        =   4
         Top             =   1440
         Width           =   5055
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   360
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cmbProf 
         Height          =   360
         Left            =   8160
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   345
         Left            =   840
         TabIndex        =   0
         Top             =   360
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   5205
         TabIndex        =   49
         Top             =   1425
         Width           =   405
         _ExtentX        =   714
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
      Begin MSMask.MaskEdBox txtIE 
         Height          =   345
         Left            =   5760
         TabIndex        =   5
         Top             =   1440
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "*Razão Social:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*CNPJ:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   90
         TabIndex        =   44
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblProf 
         AutoSize        =   -1  'True
         Caption         =   "Ramo de Atividade:"
         Height          =   225
         Left            =   8160
         TabIndex        =   35
         Top             =   1185
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   1185
         Width           =   510
      End
      Begin VB.Label lblInsc 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estatual:"
         Height          =   225
         Left            =   5760
         TabIndex        =   33
         Top             =   1185
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Endereço Cobrança "
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   0
      TabIndex        =   25
      Top             =   4080
      Width           =   11055
      Begin VB.CommandButton cmdCopEnd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copiar endereço comercial"
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Clique aqui para copiar o endereço comercial para o endereço cobrança"
         Top             =   1111
         Width           =   2895
      End
      Begin VB.TextBox txtaliq 
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   4800
         MaxLength       =   80
         TabIndex        =   20
         ToolTipText     =   "Informe Aliquota do Estado para Efeito de Entrada de Nota Fiscal:"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtIbge2 
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   840
         MaxLength       =   80
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtEndB 
         Height          =   360
         Left            =   8640
         MaxLength       =   80
         TabIndex        =   15
         Top             =   290
         Width           =   2175
      End
      Begin VB.TextBox txtUFB 
         Alignment       =   2  'Center
         DataField       =   "Estado"
         Height          =   360
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   18
         Top             =   670
         Width           =   615
      End
      Begin VB.TextBox txtCidadeB 
         DataField       =   "Cidade"
         Height          =   360
         Left            =   4800
         MaxLength       =   80
         TabIndex        =   17
         Top             =   670
         Width           =   4935
      End
      Begin VB.TextBox txtBaIrroB 
         DataField       =   "Bairro_Res"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   840
         MaxLength       =   80
         TabIndex        =   16
         Top             =   670
         Width           =   2535
      End
      Begin VB.TextBox txtRuaB 
         DataField       =   "Endereco_Res"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   14
         Top             =   290
         Width           =   4575
      End
      Begin MSMask.MaskEdBox txtCepB 
         Height          =   360
         Left            =   840
         TabIndex        =   13
         Top             =   290
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
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
      Begin VB.Label Label16 
         Caption         =   "%"
         Height          =   255
         Left            =   5400
         TabIndex        =   56
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "*Aliq. Estado:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3600
         TabIndex        =   55
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IBGE:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   255
         TabIndex        =   53
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9840
         TabIndex        =   31
         Top             =   740
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7440
         TabIndex        =   30
         Top             =   340
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   330
         TabIndex        =   29
         Top             =   345
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4080
         TabIndex        =   28
         Top             =   740
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Rua:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2160
         TabIndex        =   26
         Top             =   340
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   0
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
            Picture         =   "CADASTROFORNECEDOR.frx":82AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":86FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":8A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":8E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":92C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":95E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":9A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROFORNECEDOR.frx":9D56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imp"
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
         TabIndex        =   62
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   360
         Left            =   9720
         TabIndex        =   57
         Top             =   120
         Width           =   1185
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10200
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
               Picture         =   "CADASTROFORNECEDOR.frx":A1AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROFORNECEDOR.frx":B344
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROFORNECEDOR.frx":C3D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROFORNECEDOR.frx":D63B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROFORNECEDOR.frx":ED38
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROFORNECEDOR.frx":FE43
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROFORNECEDOR.frx":10DF8
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
      DesignWidth     =   10995
      DesignHeight    =   7680
   End
End
Attribute VB_Name = "frmCADASTROFORNECEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   CRITERIO = 0
   SETA_FONE

   GERA_FORNEC

   cmbStatus.Clear
   cmbStatus.AddItem "Ativo"
   cmbStatus.AddItem "Inativo"
   cmbProf.Clear
   cmbAuxProf.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where TIPO = 'E' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbAuxProf.AddItem TabDESCR!codigo
      cmbProf.AddItem Trim(TabDESCR!DESCRICAO)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

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
         GERA_FORNEC
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
      Case "imp"
         'IMPRIME_FORNECEDOR
      Case "voltar"
         Unload Me
      Case "gravar"
         GRAVA_TUDO
         LIMPA_TUDO
         GERA_FORNEC
      Case "matar"
         MATA_FORNECEDOR
      Case "print"
          REL_FORNECEDOR
      Case "limpar"
         LIMPA_TUDO
         GERA_FORNEC
         txtCNPJCPF.SetFocus
      Case "consultar"
         CNPJCPF_A = ""
         frmDISPLAYFORNECEDOR.Show 1
         SETA_FONE
         If Trim(CNPJCPF_A) <> "" Then
               txtCNPJCPF.PromptInclude = False
               txtCNPJCPF.Text = CNPJCPF_A
               txtCNPJCPF.PromptInclude = True
            PROCURA_DADOS
         End If
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   frmDISPLAYFORNECEDOR.Show 1
   SETA_FONE
   If Trim(CNPJCPF_A) <> "" Then
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = CNPJCPF_A
         txtCNPJCPF.PromptInclude = True
      PROCURA_DADOS
   End If
   CNPJCPF_A = ""

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

Private Sub cmdCopEnd_Click()
'On Error GoTo ERRO_TRATA

   txtCepB.PromptInclude = False
   txtCepC.PromptInclude = False
   txtCepB.Text = txtCepC.Text
   txtRuaB.Text = txtRuaC.Text
   txtEndB.Text = txtEndC.Text
   txtBaIrroB.Text = txtBairroC.Text
   txtCidadeB.Text = txtCidadeC.Text
   txtUFB.Text = txtUFC.Text
   txtIbge2.Text = txtIbge1.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCopEnd_Click"
End Sub

Private Sub cmdExcluirFone_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If Trim(txtN.Text) <> "" Then
      SQL = "delete from FONE "
      SQL = SQL & " where numero = '" & Trim(txtN.Text) & "'"
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
      txtN.Text = ""
      txtDDD.Text = ""
      txtL.Text = ""
      SETA_FONE
   End If
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdExcluirFone_Click"
End Sub

Private Sub cmbProf_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxProf.ListIndex = cmbProf.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbProf_Click"
End Sub

Private Sub txtaliq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbairrob_KeyPress"
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

   txtCNPJCPF.PromptInclude = False

   If Trim(txtCNPJCPF.Text = "") Then _
      If txtCNPJCPF.Mask = "" Then _
         txtCNPJCPF.Mask = "##############"

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         MATA_FORNECEDOR
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         SETA_FONE
         If Trim(CNPJCPF_A) <> "" Then
               txtCNPJCPF.PromptInclude = False
               txtCNPJCPF.Text = CNPJCPF_A
               txtCNPJCPF.PromptInclude = True
            PROCURA_DADOS
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
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
        If Trim(txtCNPJCPF.Text) <> "" Then
           CRITERIO = txtCNPJCPF.Text
           If Not IsNull(txtCNPJCPF.Text) Then
              If Len(txtCNPJCPF.Text) <= 11 Then
                 txtCNPJCPF.Mask = "###.###.###-##"
                 Else: txtCNPJCPF.Mask = "##.###.###/####-##"
              End If
           End If
           txtCNPJCPF.Text = CRITERIO
           Else: txtCNPJCPF.Mask = "##############"
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
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
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
      cmbEmail.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFant_KeyPress"
End Sub

Private Sub cmbEmail_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If cmbEmail.Text <> "" Then
         cmbEmail.AddItem cmbEmail.Text
         cmbEmail.Text = ""
         cmbEmail.SetFocus
         Else: txtIE.SetFocus
      End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbEmail_KeyPress"
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

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         txtCNPJCPF.PromptInclude = False
         If cmbEmail.Text <> "" And txtCNPJCPF.Text <> "" Then
            SQL = "delete from EMAIL "
            SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
            SQL = SQL & " and email = '" & cmbEmail.Text & "'"
            CONECTA_RETAGUARDA.Execute SQL

            cmbEmail.Clear

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from EMAIL "
            SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            While Not TabTemp.EOF
               cmbEmail.AddItem TabTemp!Email
               TabTemp.MoveNext
            Wend
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
         txtCNPJCPF.PromptInclude = True
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar2_ButtonClick"
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
   If txtCNPJCPF.Text <> "" And txtRazao.Text <> "" Then
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
      'SendKeys "{tab}"
      txtCepC.SetFocus
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
            If TabCEP.State = 1 Then _
               TabCEP.Close

            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepC.SetFocus
            Exit Sub
            Else
               txtCidadeC.Text = TabCEP!Cidade
               txtUFC.Text = TabCEP!UF
               If Not IsNull(TabCEP!CODIGO_IBGE) Then _
                  txtIbge1.Text = TabCEP!CODIGO_IBGE
         End If
         If TabCEP.State = 1 Then _
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
            If TabCEP.State = 1 Then _
               TabCEP.Close

            MsgBox "Cep não cadastrado. F4 - Cadastra Cep !!!"
            txtCepB.SetFocus
            Exit Sub
            Else
               txtCidadeB.Text = TabCEP!Cidade
               txtUFB.Text = TabCEP!UF
               If Not IsNull(TabCEP!CODIGO_IBGE) Then _
                  txtIbge2.Text = TabCEP!CODIGO_IBGE
         End If
         If TabCEP.State = 1 Then _
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
      txtaliq.SetFocus
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

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select razao,descricao,pessoa_id from PESSOA "
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      txtRazao.Text = "" & Trim(TabFornecedor.Fields("razao").Value)
      txtFant.Text = "" & Trim(TabFornecedor.Fields("descricao").Value)
      PESSOA_ID_N = 0 & TabFornecedor.Fields("pessoa_id").Value
      Else
         If TabFornecedor.State = 1 Then _
            TabFornecedor.Close
         MsgBox "Cadastro tabela pessoa não encontrado."
         Exit Sub
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select * from vwFornecedor "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      txtCodigo.Text = TabFornecedor!FORNECEDOR_ID
      txtRazao.Text = TabFornecedor!NOME

      If Not IsNull(TabFornecedor!razao_social) Then _
         txtFant.Text = TabFornecedor!razao_social

      If Not IsNull(TabFornecedor!Status) Then
         If TabFornecedor!Status = "A" Then
            cmbStatus.Text = "Ativo"
            Else: cmbStatus.Text = "Inativo"
         End If
      End If
      If Not IsNull(TabFornecedor!IE) Then _
         txtIE.Text = TabFornecedor!IE
   
      If Not IsNull(TabFornecedor.Fields("status").Value) Then _
         If Trim(TabFornecedor.Fields("status").Value) = "C" Then _
            MsgBox "Fornecedor Cancelado."
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close

   'EMAIL
   SQL = "select * from EMAIL "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbEmail.AddItem TabTemp!Email
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
   
   'FONE
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
            SQL = SQL & " where cep = '" & tabEndereco!CEP & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCidadeC.Text = TabConsulta!Cidade
               txtUFC.Text = TabConsulta!UF
               txtIbge1.Text = "" & TabConsulta.Fields("codigo_ibge").Value
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
      txtRuaB.Text = "" & tabEndereco!Rua
      txtBaIrroB.Text = "" & tabEndereco!Bairro
      If Not IsNull(tabEndereco!Complemento) Then
         txtEndB.Text = tabEndereco!Complemento
      End If
      If Not IsNull(tabEndereco!CEP) Then
         If tabEndereco!CEP <> "" Then
            txtCepB.Text = tabEndereco!CEP

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from CEP "
            SQL = SQL & " where cep = '" & tabEndereco!CEP & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCidadeB.Text = TabConsulta!Cidade
               txtUFB.Text = TabConsulta!UF
               txtIbge2.Text = "" & TabConsulta.Fields("codigo_ibge").Value
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
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_DADOS"
End Sub

Private Sub LIMPA_QUASE_TUDO()
'On Error GoTo ERRO_TRATA

   txtRazao.Text = ""
   cmbStatus.Text = ""
   cmbProf.Text = ""
   cmbEmail.Text = ""
   cmbEmail.Clear
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

   txtIbge1.Text = ""
   txtIbge2.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtRazao.Text = ""
   cmbStatus.Text = ""
   cmbProf.Text = ""
   cmbAuxProf.Text = ""
   cmbEmail.Text = ""
   cmbEmail.Clear
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
   txtaliq.Text = "0"
   txtCNPJCPF.Mask = "##############"
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

   Dim strRazaoSocial As String
   Dim intStatus As Integer
   Dim strIe As String
   
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe CPF do FORNECEDOR."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   If txtRazao.Text = "" Then
      MsgBox "Informe Nome do FORNECEDOR."
      txtRazao.SetFocus
      Exit Sub
   End If

   'se for vazio colocar aliquota local
   If txtaliq.Text = "" Then _
      txtaliq.Text = 17

   If txtIE.Text <> "" Then
      strIe = txtIE.Text
      Else: strIe = "ISENTO"
   End If

   If Not IsNull(txtFant.Text) Then _
      strRazaoSocial = txtFant.Text
   If cmbStatus.Text = "" Then _
      cmbStatus.Text = "Ativo"
   If txtaliq.Text = "" Then _
      txtaliq.Text = 0

'=========================
'PESSOA
   PESSOA_ID_N = 0
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select pessoa_id from PESSOA "
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then _
      If Not IsNull(TabFornecedor.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabFornecedor.Fields(0).Value)
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   CONT_N = 2
   If PESSOA_ID_N <= 0 Then _
      CONT_N = 1

   'executa stored procedure sp_pessoa
   SP_PESSOA CONT_N, PESSOA_ID_N, Trim(txtCNPJCPF.Text), Trim(txtFant.Text), Trim(txtRazao.Text), Left(cmbStatus.Text, 1)

   PESSOA_ID_N = 0
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select pessoa_id from PESSOA "
   SQL = SQL & " where CNPJcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then _
      If Not IsNull(TabFornecedor.Fields(0).Value) Then _
         PESSOA_ID_N = Trim(TabFornecedor.Fields(0).Value)
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close
'=========================

   SQL = "select * from vwFornecedor "
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      FORNEC_ID_N = TabFornecedor!FORNECEDOR_ID

      SQL = "UPDATE FORNECEDOR SET "
         SQL = SQL & " Status = '" & Left(cmbStatus.Text, 1) & "'"
      SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
      Else
         GERA_FORNEC

         SQL = "INSERT INTO FORNECEDOR "
            SQL = SQL & " (fornecedor_id,pessoa_id,EMPRESA_ID, dt_cad, CGCCPF, Nome, RAZAO_SOCIAL, IE, Status) "
         SQL = SQL & " VALUES ("
         SQL = SQL & FORNEC_ID_N
         SQL = SQL & "," & PESSOA_ID_N
         SQL = SQL & "," & EMPRESA_ID_N
         SQL = SQL & ",'" & Now & "'"
         SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
         SQL = SQL & ",'" & txtRazao.Text & "'"
         SQL = SQL & ",'" & strRazaoSocial & "'"
         SQL = SQL & ",'" & strIe & "'"
         SQL = SQL & ",'" & Left(cmbStatus.Text, 1) & "'"
         SQL = SQL & "," & txtaliq.Text
         SQL = SQL & ")"
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   CONECTA_RETAGUARDA.Execute SQL
   
'FONE
   SQL = "Delete FONE "
   SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
   CONECTA_RETAGUARDA.Execute SQL

   Dim i As Integer
   For i = 1 To FlexTel.Rows - 1
      FlexTel.Row = i
      SQL = "insert into FONE (pessoa_id,PROP,DDD,NUMERO,LOCAL) values ("
      SQL = SQL & PESSOA_ID_N
      SQL = SQL & ",'" & FlexTel.TextMatrix(i, 3)
      SQL = SQL & "','" & FlexTel.TextMatrix(i, 0)
      SQL = SQL & "','" & FlexTel.TextMatrix(i, 1)
      SQL = SQL & "','" & Replace(FlexTel.TextMatrix(i, 2), "|", "/") & "')"
      CONECTA_RETAGUARDA.Execute SQL
   Next

   'EMAIL
   If cmbEmail.ListCount > 0 Then
      NUMR_SEQ_N = 0
      While NUMR_SEQ_N <> cmbEmail.ListCount
         cmbEmail.ListIndex = NUMR_SEQ_N

         Dim EMAIL_ID As Integer

         If cmbEmail.Text <> "" Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from EMAIL "
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and email = '" & Trim(cmbEmail.Text) & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabTemp.EOF Then

               SQL = "INSERT INTO EMAIL "
               SQL = SQL & " VALUES ("
                  SQL = SQL & PESSOA_ID_N
                  SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
                  SQL = SQL & ",'" & Trim(cmbEmail.Text) & "'"
               SQL = SQL & ")"

               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
         NUMR_SEQ_N = NUMR_SEQ_N + 1
      Wend
   End If

   If txtRuaC.Text <> "" And txtBairroC.Text <> "" Then
      'ENDEREÇO COMERCIAL
      txtCepC.PromptInclude = False
      If txtCepC.Text <> "" Or txtRuaC.Text <> "" Or txtBairroC.Text <> "" Or txtEndC.Text <> "" Then
         If txtCepC.Text <> "" Then _
            SP_GRAVA_CEP txtCepC.Text, txtCidadeC.Text, txtUFC.Text, txtIbge1.Text
            
            
         sp_Grava_Endereco txtCepC.Text, txtRuaC.Text, txtBairroC.Text, txtEndC.Text, "C", 0
         Else: SP_MATA_ENDEREÇO "C"
      End If
   End If
   If txtRuaB.Text <> "" And txtBaIrroB.Text <> "" Then
      'ENDEREÇO COBRANÇA
      txtCepB.PromptInclude = False
      If txtCepB.Text <> "" Or txtRuaB.Text <> "" Or txtBaIrroB.Text <> "" Or txtEndB.Text <> "" Then
         If txtCepB.Text <> "" Then _
            SP_GRAVA_CEP txtCepB.Text, txtCidadeB.Text, txtUFB.Text, txtIbge2.Text
   
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

Sub GRAVA_FONE_TEMP()
'On Error GoTo ERRO_TRATA

   Dim strAux As String * 2
   Dim Achou As Boolean
   Dim i As Integer

   Achou = False

   strAux = "DP"
   txtCNPJCPF.PromptInclude = False
   FlexTel.Col = 1
   
   For i = 1 To FlexTel.Rows - 1
      FlexTel.Row = i
      If Replace(txtN.Text, "-", "") = FlexTel.Text Then
         Achou = True
         Exit For
      End If
   Next

   If Not Achou Then
      FlexTel.AddItem ""
      FlexTel.Row = FlexTel.Rows - 1
      FlexTel.Col = 0
      FlexTel.Text = txtDDD.Text
      FlexTel.Col = 1
      FlexTel.Text = Replace(txtN.Text, "-", "")
      FlexTel.Col = 2
      FlexTel.Text = strAux & " / " & txtL.Text
      FlexTel.Col = 3
      FlexTel.Text = txtCNPJCPF.Text
      Else
         FlexTel.Col = 0
         FlexTel.Text = txtDDD.Text
         FlexTel.Col = 1
         FlexTel.Text = Replace(txtN.Text, "-", "")
         FlexTel.Col = 2
         FlexTel.Text = strAux & " / " & txtL.Text
         FlexTel.Col = 3
         FlexTel.Text = txtCNPJCPF.Text
   End If
   FlexTel.Refresh
   LIMPA_FONE
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_FONE_TEMP"
End Sub

Private Sub SETA_FONE()
'On Error GoTo ERRO_TRATA

   If TabAUX.State = 1 Then _
      TabAUX.Close

   SQL = "select * from FONE "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
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

Public Sub GERA_FORNEC()
'On Error GoTo ERRO_TRATA

   FORNEC_ID_N = 0 & MAX_ID("fornecedor_id", "FORNECEDOR", "", "", "", "")
   txtCodigo.Text = FORNEC_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_FORNEC"
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
   TRATA_ERROS Err.Description, Me.Name, "REFRESH_GRID"
End Sub

Sub MATA_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" Then
      Msg = "Confirma exclusão do FORNECEDOR ?"
      PERGUNTA Msg, vbYesNo + 32, "Cadastro Fornecedor NFE", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         'Exclui Cadastro do Fornecedor e dados Relacionados ao Cadastro

         'Exclui email
         SQL = "delete  from EMAIL "
         SQL = SQL & " where prop = '" & Trim(txtCNPJCPF.Text) & "'"
         CONECTA_RETAGUARDA.Execute SQL

         'Exclui Fone
         SQL = "delete  from FONE "
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "update FORNECEDOR set "
         SQL = SQL & "status='C'"
         SQL = SQL & "where CGCCPF='" & Trim(txtCNPJCPF.Text) & "'"
         CONECTA_RETAGUARDA.Execute SQL
         LIMPA_TUDO
         txtCNPJCPF.SetFocus
      End If
   End If
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_FORNECEDOR"
End Sub

Sub IMPRIME_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   SQL = "select * from vwFornecedor "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select * from ENDERECO "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      'end comercial
      If Not IsNull(tabEndereco!CEP) Then
         If tabEndereco!CEP <> "" Then
            If TabCEP.State = 1 Then _
               TabCEP.Close

            SQL = "select * from CEP"
            SQL = SQL & " where cep = '" & tabEndereco!CEP & "'"
            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabCEP.EOF Then
               SqL2 = "INSERT INTO CEP (cep, Cidade, UF ) "
               SqL2 = SqL2 & " VALUES ('" & tabEndereco!CEP & "','" & Cidade_Com & "','" & Estado_Com & "')"
               CONECTA_RETAGUARDA.Execute SqL2
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
         End If
      End If

      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select * from ENDERECO "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If tabEndereco.EOF Then
         SqL2 = "INSERT INTO ENDERECO (prop, Complemento, Bairro, Rua, Cep ) "
         SqL2 = SqL2 & " VALUES ('" & TabTemp!CGCCPF & "','" & TabTemp!Endereco_Com & "','" & TabTemp!Bairro_Com & "','" & TabTemp!Rua_com & "','" & TabTemp!CEP & "')"
         CONECTA_RETAGUARDA.Execute SqL2
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from EMAIL "
      SQL = SQL & " where prop = '" & TabTemp!CGCCPF & "'"
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabAUX.EOF Then
         SqL2 = "INSERT INTO EMAIL (empresa_id, prop, EMAIL ) "
         SqL2 = SqL2 & " VALUES (" & EMPRESA_ID_N & ",'" & TabTemp!CGCCPF & "','" & "email@email.com.br" & "')"
         'CONECTA_RETAGUARDA.Execute SqL2
      End If
      If TabAUX.State = 1 Then _
         TabAUX.Close

      'FONE
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select * from FONE "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If tabEndereco.EOF Then
         SQL = "INSERT INTO FONE (prop, ddd, numero, Local ) "
         SQL = SQL & " VALUES ("
         SQL = SQL & "'" & TabTemp!CGCCPF & "'"
         SQL = SQL & "," & TabTemp!DDD1 & ","
         SQL = SQL & "'" & TabTemp!Fone_Com & "'"
         SQL = SQL & ",'" & Trim(txtL.Text) & "')"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIME_FORNECEDOR"
End Sub

Sub REL_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{Fornecedor.empresa_id} = " & EMPRESA_ID_N
   If IsNumeric(txtCNPJCPF.Text) Then
      txtCNPJCPF.PromptInclude = False
      FORMULA_REL = FORMULA_REL & " and {Fornecedor.cgccpf} = '" & Trim(txtCNPJCPF.Text) & "'"
   End If
   If txtCepC.Text <> "" Then
      FORMULA_REL = FORMULA_REL & " and {CEP.CEP} = " & txtCepC.Text
   End If
   If txtCidadeC.Text <> "" Then
      FORMULA_REL = FORMULA_REL & " and {CEP.Cidade} = '" & txtCidadeC.Text & "'"
   End If
   If txtUFC.Text <> "" Then
      FORMULA_REL = FORMULA_REL & " and {CEP.uf} = '" & txtUFC.Text & "'"
   End If
   If Left(cmbStatus.Text, 1) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {Fornecedor.Status} = '" & Left(cmbStatus.Text, 1) & "'"
   End If

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_Fornecedor.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REL_FORNECEDOR"
End Sub
