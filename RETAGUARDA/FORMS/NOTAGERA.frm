VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNOTAGERA 
   BackColor       =   &H80000004&
   Caption         =   "Emissor de Nota Fiscal"
   ClientHeight    =   7620
   ClientLeft      =   1740
   ClientTop       =   2355
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NOTAGERA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   1020
      Left            =   9600
      MaskColor       =   &H00FF8080&
      Picture         =   "NOTAGERA.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   115
      ToolTipText     =   "Confirma os acessos para este usuario."
      Top             =   6600
      Width           =   1125
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Voltar"
      CausesValidation=   0   'False
      Height          =   1020
      Left            =   10800
      Picture         =   "NOTAGERA.frx":6D8D
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   6600
      Width           =   1140
   End
   Begin VB.ComboBox cmbTpEmisAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F17
      Left            =   9840
      List            =   "NOTAGERA.frx":7F19
      TabIndex        =   113
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbTpEmis 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F1B
      Left            =   9840
      List            =   "NOTAGERA.frx":7F1D
      TabIndex        =   112
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtChaveNFe 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      MaxLength       =   50
      TabIndex        =   99
      ToolTipText     =   "Informe a quantidade"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtNFeDev 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9840
      MaxLength       =   50
      TabIndex        =   98
      ToolTipText     =   "Informe a quantidade"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbFinalidadeAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F1F
      Left            =   9840
      List            =   "NOTAGERA.frx":7F21
      TabIndex        =   96
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFinalidade 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F23
      Left            =   9840
      List            =   "NOTAGERA.frx":7F25
      TabIndex        =   94
      Top             =   4440
      Width           =   1935
   End
   Begin VB.ComboBox cmbMarcaAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F27
      Left            =   9840
      List            =   "NOTAGERA.frx":7F29
      TabIndex        =   93
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbMarca 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F2B
      Left            =   9840
      List            =   "NOTAGERA.frx":7F2D
      TabIndex        =   91
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox cmbCFinalAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F2F
      Left            =   9840
      List            =   "NOTAGERA.frx":7F31
      TabIndex        =   90
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbCFinal 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F33
      Left            =   9840
      List            =   "NOTAGERA.frx":7F35
      TabIndex        =   88
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox cmbAmbienteAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F37
      Left            =   9840
      List            =   "NOTAGERA.frx":7F39
      TabIndex        =   87
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbAmbiente 
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F3B
      Left            =   9840
      List            =   "NOTAGERA.frx":7F3D
      TabIndex        =   85
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cmbTipoOperaAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F3F
      Left            =   9840
      List            =   "NOTAGERA.frx":7F41
      TabIndex        =   84
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbTipoOpera 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F43
      Left            =   9840
      List            =   "NOTAGERA.frx":7F45
      TabIndex        =   82
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox cmbLocalAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F47
      Left            =   9840
      List            =   "NOTAGERA.frx":7F49
      TabIndex        =   81
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbPresencaAUX 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F4B
      Left            =   9840
      List            =   "NOTAGERA.frx":7F4D
      TabIndex        =   80
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbLocal 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F4F
      Left            =   9840
      List            =   "NOTAGERA.frx":7F51
      TabIndex        =   78
      Top             =   3720
      Width           =   1935
   End
   Begin VB.ComboBox cmbPresenca 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "NOTAGERA.frx":7F53
      Left            =   9840
      List            =   "NOTAGERA.frx":7F55
      TabIndex        =   76
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame Frame8 
      Caption         =   "Outras Informações"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   0
      TabIndex        =   64
      Top             =   2460
      Width           =   8415
      Begin VB.ComboBox cmbFreteAUX 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "NOTAGERA.frx":7F57
         Left            =   5520
         List            =   "NOTAGERA.frx":7F59
         TabIndex        =   109
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbFrete 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "NOTAGERA.frx":7F5B
         Left            =   5400
         List            =   "NOTAGERA.frx":7F5D
         TabIndex        =   108
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox TxtEspecie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   18
         Text            =   "UN"
         ToolTipText     =   "Informe a especie"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtPesoLiquido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "1"
         ToolTipText     =   "Peso liquido da nota"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtPesoBruto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "1"
         ToolTipText     =   "Peso bruto da nota"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtQuantidadeRodapeNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   17
         Text            =   "1"
         ToolTipText     =   "Informe a quantidade"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Frete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5415
         TabIndex        =   110
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso líquido:"
         Height          =   240
         Left            =   4080
         TabIndex        =   69
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso bruto:"
         Height          =   240
         Left            =   2760
         TabIndex        =   68
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie:"
         Height          =   240
         Left            =   1440
         TabIndex        =   67
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   240
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   59
      Top             =   4560
      Width           =   8445
      Begin VB.TextBox txtDadosAdicionais 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1470
         MultiLine       =   -1  'True
         TabIndex        =   32
         Tag             =   " "
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   690
         Width           =   6885
      End
      Begin VB.TextBox txtMSG 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1470
         MultiLine       =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "Insira um texto e sera impresso no campo DADOS ADICIONAIS  da nota fiscal."
         Top             =   210
         Width           =   6885
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionais:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   61
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Msg.Rodapé:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   60
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transportadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   0
      TabIndex        =   58
      Top             =   1800
      Width           =   8415
      Begin VB.ComboBox cmbAuxCNPJCPF_TRANSP 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "NOTAGERA.frx":7F5F
         Left            =   3960
         List            =   "NOTAGERA.frx":7F61
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbCNPJCPF_TRANSP 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         ItemData        =   "NOTAGERA.frx":7F63
         Left            =   3960
         List            =   "NOTAGERA.frx":7F65
         TabIndex        =   16
         Top             =   240
         Width           =   4125
      End
      Begin MSMask.MaskEdBox txtCNPJCPF_TRANSP 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         ToolTipText     =   "Se houver uma transportadora informe aqui. F7 Consulta transportadora"
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   62
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame fraEmitente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1335
      Left            =   0
      TabIndex        =   49
      Top             =   480
      Width           =   11895
      Begin VB.ComboBox cmbIE 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         ItemData        =   "NOTAGERA.frx":7F67
         Left            =   9720
         List            =   "NOTAGERA.frx":7F69
         TabIndex        =   104
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cmbEmail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         ItemData        =   "NOTAGERA.frx":7F6B
         Left            =   8760
         List            =   "NOTAGERA.frx":7F6D
         TabIndex        =   14
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtIBGE 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5820
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   12
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txtEmitente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   5
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtCep 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   600
         Width           =   2715
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtUF_CLIENTE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtFone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7260
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   13
         Top             =   960
         Width           =   1395
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   330
         Left            =   8400
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483634
         ForeColor       =   8388608
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IBGE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5340
         TabIndex        =   72
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destinatário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   71
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7440
         TabIndex        =   57
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   56
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7470
         TabIndex        =   55
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4005
         TabIndex        =   54
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Município:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   53
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6765
         TabIndex        =   52
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4320
         TabIndex        =   51
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblInsc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CCE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9225
         TabIndex        =   50
         Top             =   600
         Width           =   390
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fatura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      TabIndex        =   47
      Top             =   3360
      Width           =   8415
      Begin MSComctlLib.ListView ListaDP 
         Height          =   915
         Left            =   90
         TabIndex        =   48
         Top             =   210
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   1614
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Seq."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor da Duplicata"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "NºDocumento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Venc.DP"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Histórico"
            Object.Width           =   10583
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   0
      TabIndex        =   46
      Top             =   5760
      Width           =   8415
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         TabIndex        =   103
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbCFOPAux 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   102
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbCFOP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   101
         Text            =   "-- Selecione --"
         ToolTipText     =   "CFOP"
         Top             =   1320
         Visible         =   0   'False
         Width           =   4815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2778
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraImposto 
      Caption         =   "Cálculo Imposto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      TabIndex        =   39
      Top             =   7680
      Width           =   11925
      Begin VB.CheckBox chkImposto 
         Caption         =   "Calcular Imposto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1560
         TabIndex        =   105
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox txtValorOutros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8760
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtVlrIcmsSub 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6720
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtBaseIcmsSub 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6720
         TabIndex        =   23
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   10800
         TabIndex        =   25
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtValorTotalNota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   10800
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtValorProdutos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4200
         TabIndex        =   22
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtBaseCalculo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1920
         TabIndex        =   21
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   8760
         TabIndex        =   24
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr ICMS Sub.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5430
         TabIndex        =   75
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outros:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   74
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. Icms Sub.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5520
         TabIndex        =   73
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VLR IPI.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   70
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Nota:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   45
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total dos Produtos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr.ICMS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   43
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base cálculo ICMS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   40
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame fraNota 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   33
      Top             =   -120
      Width           =   11895
      Begin VB.TextBox txtModelo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   10935
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   106
         ToolTipText     =   "Série Nota Fiscal"
         Top             =   200
         Width           =   855
      End
      Begin VB.TextBox txtDtEmis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Data Emissão Nota Fiscal"
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtDtSaida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3330
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Data Saida Nota Fiscal"
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtHoraSaida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Hora de Saída Nota Fiscal"
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         ToolTipText     =   "Número Nota Fiscal"
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   9120
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "Série Nota Fiscal"
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10080
         TabIndex        =   107
         Top             =   200
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Emissão:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   38
         Top             =   200
         Width           =   990
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Saída:"
         Height          =   240
         Left            =   2445
         TabIndex        =   37
         Top             =   200
         Width           =   870
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Saída:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4515
         TabIndex        =   36
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº  Nota:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   35
         Top             =   200
         Width           =   705
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Série:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8505
         TabIndex        =   34
         Top             =   195
         Width           =   510
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
      DesignWidth     =   11940
      DesignHeight    =   7620
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tp.Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8730
      TabIndex        =   111
      Top             =   3360
      Width           =   1005
   End
   Begin VB.Label lblNFeDev 
      BackStyle       =   0  'Transparent
      Caption         =   "NFeDevolução:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8520
      TabIndex        =   97
      Top             =   4800
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fin.Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8700
      TabIndex        =   95
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   9120
      TabIndex        =   92
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cons.Final:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8835
      TabIndex        =   89
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ambiente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8880
      TabIndex        =   86
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tp.Operação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8655
      TabIndex        =   83
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id.Dest.Oper.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8625
      TabIndex        =   79
      Top             =   3720
      Width           =   1110
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id.Comprador:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8550
      TabIndex        =   77
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   66
      Top             =   3000
      Width           =   1020
   End
End
Attribute VB_Name = "frmNOTAGERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Tipo_Endereço             As String * 1
   Dim NOME_VEND_A               As String
   Dim NossoNumero               As String
   Dim CODG_SUFRAMA_A            As String
   Dim FONE_EMPRESA_N            As String * 10
   Dim strCFOP                   As String
   Dim NaturezaOperacao_A        As String * 60
   Dim CFOP_ID_N                 As Long
   Dim TRANSP_ID_N               As Long
   Dim Vlr_BaseCalculo_N         As Double
   Dim Aliquota_ICMS_Normal_N    As Double

   Dim Vlr_TotICMS_N             As Double
   Dim Vlr_BaseICMSub_N          As Double
   Dim Vlr_ICMSub_N              As Double
   Dim Vlr_TotProdutos_N         As Double
   Dim VLR_FRETE_N               As Double
   Dim Vlr_Desconto_N            As Double
   Dim Vlr_TotIPI_N              As Double
   Dim VLR_OUTROS_N              As Double
   Dim Vlr_TotNFe_N              As Double

   Private ControlVisible        As Boolean
   Private LastRow               As Long ' Ultima linha em que se editou
   Private LastCol               As Long ' ultima coluna em que se editou

Private Sub Form_Load()
'set so gerar quando clicr em gravar

   INICIALIZA_NF
   TRIBUTOS_LEI12741

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         FORMULA_REL = "{LANCAMENTO.Tipo_Lancamento} = " & 1
         FORMULA_REL = FORMULA_REL & " and {ITEMLANCAMENTO.Numr_Doc} = " & PEDIDO_ID_N

         'If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "BoletoItau.rpt"
         frmRELATORIO10.Show 1
      Case vbKeyEscape
         Unload Me
      Case vbKeyF10
         'EXCLUIR_400
            GERAR_NFE
         'EXCLUIR_400
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   TIPO_NFe_GERAR = ""
End Sub

Private Sub CmdGravar_Click()
   GERAR_NFE
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmbCFOP_Click()
On Error Resume Next

   cmbCFOPAux.ListIndex = cmbCFOP.ListIndex
   txtValorDig.Text = cmbCFOPAux.Text
   txtValorDig.SetFocus

End Sub

Private Sub cmblocal_Click()
On Error Resume Next

   cmbLocalAUX.ListIndex = cmbLocal.ListIndex

End Sub

Private Sub cmbPresenca_Click()
On Error Resume Next

   cmbPresencaAUX.ListIndex = cmbPresenca.ListIndex

End Sub

Private Sub cmbtipoopera_Click()
On Error Resume Next

   cmbTipoOperaAUX.ListIndex = cmbTipoOpera.ListIndex

End Sub

Private Sub cmbAmbiente_Click()
On Error Resume Next

   cmbAmbienteAUX.ListIndex = cmbAmbiente.ListIndex

End Sub

Private Sub cmbCfinal_Click()
On Error Resume Next

   cmbCFinalAUX.ListIndex = cmbCFinal.ListIndex

End Sub

Private Sub cmbmarca_Click()
On Error Resume Next

   cmbMarcaAUX.ListIndex = cmbMarca.ListIndex

End Sub

Private Sub cmbfinalidade_Click()
On Error Resume Next

   txtNFeDev.Visible = False
   lblNFeDev.Visible = False
   'lblChave.Visible = False
   txtChaveNFe.Visible = False

   cmbFinalidadeAUX.ListIndex = cmbFinalidade.ListIndex

   If Trim(cmbFinalidadeAUX.Text) = "4" Then

      txtNFeDev.Visible = True
      lblNFeDev.Visible = True
      txtChaveNFe.Visible = True
      txtNFeDev.SetFocus
      Else
         txtNFeDev.Visible = False
         lblNFeDev.Visible = False
         txtChaveNFe.Visible = False
   End If
End Sub

Private Sub txtChaveNFe_LostFocus()
   If Trim(txtChaveNFe.Text) <> "" Then
      If Len(Trim(txtChaveNFe.Text)) <> 44 Then
         MsgBox "Chave informada inválida, verifique."
         txtChaveNFe.Text = ""
         txtChaveNFe.SetFocus
      End If
   End If
End Sub

Private Sub TxtEspecie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
       KeyAscii = 0
       TxtPesoBruto.SetFocus
   End If
End Sub

Private Sub txtNFeDev_LostFocus()
'On Error GoTo ERRO_TRATA

   txtChaveNFe.Text = ""

   If Trim(txtNFeDev.Text) <> "" Then
      If IsNumeric(txtNFeDev.Text) Then
         txtChaveNFe.Text = "" & TRAZ_DADOS_NFe("MFA010", "MFACHAVENFE", "MFADOC", Trim(txtNFeDev.Text))
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFrete_KeyPress"
End Sub

Private Sub TxtPesoBruto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      TxtPesoLiquido.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub TxtPesoLiquido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       KeyAscii = 0
       txtDadosAdicionais.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub TxtQuantidadeRodapeNota_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      TxtEspecie.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
    End If
End Sub

Private Sub txtNFeDev_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
    End If
End Sub

Private Sub txtDadosAdicionais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtCNPJCPF_TRANSP_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_TRANSP.PromptInclude = False
   If txtCNPJCPF_TRANSP.Text = "" Then
      txtCNPJCPF_TRANSP.Mask = "##############"
      Else
         If Len(txtCNPJCPF_TRANSP.Text) <= 11 Then
            txtCNPJCPF_TRANSP.Mask = "###.###.###-##"
            Else
               If Len(txtCNPJCPF_TRANSP.Text) > 11 Then
                  txtCNPJCPF_TRANSP.Mask = "##.###.###/####-##"
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_GotFocus"
End Sub

Private Sub txtCNPJCPF_TRANSP_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

    Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF_TRANSP.PromptInclude = False
            txtCNPJCPF_TRANSP.Text = Trim(CNPJCPF_A)
         End If

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select cnpj,descricao from vwTRANSPORTADORA WITH (NOLOCK)"
         SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               cmbCNPJCPF_TRANSP.Text = Trim(TabTemp.Fields("descricao").Value)
         If TabTemp.State = 1 Then _
            TabTemp.Close

         txtCNPJCPF_TRANSP.PromptInclude = True
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_KeyDown"
End Sub

Private Sub txtCNPJCPF_TRANSP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      TxtQuantidadeRodapeNota.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_KeyPress"
End Sub

Private Sub txtCNPJCPF_TRANSP_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_TRANSP.PromptInclude = False

   If Trim(txtCNPJCPF_TRANSP.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select descricao from vwTRANSPORTADORA WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         MsgBox "Informar trasportadora."
         txtCNPJCPF_TRANSP.Text = ""
         cmbCNPJCPF_TRANSP.Text = ""
         Exit Sub
         Else: If Not IsNull(TabTemp.Fields(0).Value) _
               Then cmbCNPJCPF_TRANSP.Text = Trim(TabTemp.Fields(0).Value)
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   txtCNPJCPF_TRANSP.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_TRANSP_LostFocus"
End Sub

Private Sub cmbCNPJCPF_TRANSP_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbCNPJCPF_TRANSP.Clear
   cmbAuxCNPJCPF_TRANSP.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwTRANSPORTADORA WITH (NOLOCK)"
   'SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCNPJCPF_TRANSP.AddItem Trim(TabTemp!CNPJCPF) & " - " & Trim(TabTemp!DESCRICAO)
      cmbAuxCNPJCPF_TRANSP.AddItem Trim(TabTemp!CNPJCPF)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCNPJCPF_TRANSP_GotFocus"
End Sub

Private Sub cmbCNPJCPF_TRANSP_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxCNPJCPF_TRANSP.ListIndex = cmbCNPJCPF_TRANSP.ListIndex
   txtCNPJCPF_TRANSP.PromptInclude = False
      Select Case Len(cmbAuxCNPJCPF_TRANSP.Text)
         Case Is <= 11
            txtCNPJCPF_TRANSP.Mask = "###.###.###-##"
         Case Is = 14
            txtCNPJCPF_TRANSP.Mask = "##.###.###/####-##"
      End Select
      txtCNPJCPF_TRANSP.Text = "" & Trim(cmbAuxCNPJCPF_TRANSP.Text)
   txtCNPJCPF_TRANSP.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCNPJCPF_TRANSP_Click"
End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorDig.SelStart = 0
   txtValorDig.SelLength = Len(txtValorDig)

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
      If LastCol = 5 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      CFOP_ID_N = 0 & txtValorDig.Text

      If Trim(TRAZ_CFOP(CFOP_ID_N)) = "" Then
         txtValorDig.SelStart = 0
         txtValorDig.SelLength = Len(txtValorDig)
         MsgBox "CFOP inválido."
         Exit Sub
      End If

      Dim CFOP_ID_ANT As Double

      CFOP_ID_ANT = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
      'PEDIDO_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
      PRODUTO_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)

      'AtribuiValorCelula
      'ProximaCelula
      OcultarControles

      If CFOP_ID_N > 0 Then
         SQL = "update PEDIDOITEM set "
         SQL = SQL & " cfop_id = " & CFOP_ID_N

         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         CFOP_ID_ANT = 0
         PRODUTO_ID_N = 0
         SEQ_ID_N = 0

         SETA_GRID
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
      NaturezaOperacao_A = ""
      MSFlexGrid1.SetFocus
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            cmbCFOP.Visible = False
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

Private Sub txtFrete_LostFocus()
   CALCULA_IMPOSTO
End Sub

Private Sub txtValorOutros_LostFocus()
   CALCULA_IMPOSTO
End Sub

Private Sub chkImposto_Click()

   If chkImposto.Value = 1 Then
      If Trim(cmbIE.Text) <> "" Then _
         If Trim(cmbIE.Text) <> "ISENTO" Then _
            CALCULA_IMPOSTO
      Else: TOTAIS_NOTA
   End If

End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    OcultarControles

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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
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
         'AtribuiValorCelula
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
'============= S U B R O T I N A S
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear
   MSFlexGrid1.Visible = False
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select codg_produto as Código,descricao as Descrição, "
   SQL = SQL & " QTD_PEDIDA as Qtde,Valor_Item as PreçoVenda, (QTD_PEDIDA*Valor_Item) as Total,"
   SQL = SQL & " cfop_id as CFOP, stributaria as ST, PercIcms as ICMS,"
   SQL = SQL & " codg_ncm as NCM,Unidade_Medida as UN, pedidoitem.produto_id,pedido_id,seq_id"

   SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "

'Debug.Print SQL

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabProduto.Fields.Count

      ReDim largura_coluna(0 To TabProduto.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabProduto.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabProduto.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabProduto.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabProduto.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabProduto.Fields.Count - 1
            If Coluna = 2 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabProduto.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna = 3 Or Coluna = 4 Or Coluna = 7 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabProduto.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabProduto.Fields(Coluna).Value)
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabProduto.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabProduto.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabProduto.MoveNext
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

'Codigo Produto
      MSFlexGrid1.ColWidth(0) = 1000
      MSFlexGrid1.ColAlignment(0) = 0

'Descrição Produto
      MSFlexGrid1.ColWidth(1) = 4000
      MSFlexGrid1.ColAlignment(1) = 0

'QTDE
      MSFlexGrid1.ColWidth(2) = 1500
      MSFlexGrid1.ColAlignment(2) = 7

'Valor Item
      MSFlexGrid1.ColWidth(3) = 1500
      MSFlexGrid1.ColAlignment(3) = 7

'Total Item
      MSFlexGrid1.ColWidth(4) = 1500
      MSFlexGrid1.ColAlignment(4) = 7

'cfop
      MSFlexGrid1.ColWidth(5) = 1000
      MSFlexGrid1.ColAlignment(5) = 7

'SITUAÇÃO TRIBUTARIA PRODUTO
      MSFlexGrid1.ColWidth(6) = 500
      MSFlexGrid1.ColAlignment(6) = 0

'ALIQUOTA ICMS
      MSFlexGrid1.ColWidth(7) = 1000
      MSFlexGrid1.ColAlignment(7) = 7

'NCM
      MSFlexGrid1.ColWidth(8) = 1000
      MSFlexGrid1.ColAlignment(8) = 0

'UN
      MSFlexGrid1.ColWidth(9) = 500
      MSFlexGrid1.ColAlignment(9) = 0

'produto_id
      MSFlexGrid1.ColWidth(10) = 0
      MSFlexGrid1.ColAlignment(10) = 0

'pedido_id
      MSFlexGrid1.ColWidth(11) = 0
      MSFlexGrid1.ColAlignment(11) = 0

'seq_id
      MSFlexGrid1.ColWidth(12) = 0
      MSFlexGrid1.ColAlignment(12) = 0
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MSFlexGrid1.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub TOTAIS_NOTA()
'On Error GoTo ERRO_TRATA

   Dim VALOR_TOTAL_PRODUTO_N  As Double

   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_PRODUTO_N = 0

txtBaseCalculo.Text = "" & Format(0, strFormatacao2Digitos)
txtValorICMS.Text = "" & Format(0, strFormatacao2Digitos)
txtBaseIcmsSub.Text = "" & Format(0, strFormatacao2Digitos)
txtFrete.Text = "" & Format(0, strFormatacao2Digitos)
txtValorIPI.Text = "" & Format(0, strFormatacao2Digitos)
txtValorProdutos.Text = "" & Format(0, strFormatacao2Digitos)
txtDesconto.Text = "" & Format(0, strFormatacao2Digitos)
txtVlrIcmsSub.Text = "" & Format(0, strFormatacao2Digitos)
txtValorOutros.Text = "" & Format(0, strFormatacao2Digitos)
txtValorTotalNota.Text = "" & Format(0, strFormatacao2Digitos)

  'valor de desconto na cabeça
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'desconto individual por item
   SQL = "select sum((valor_item*qtd_pedida)*perc_desc/100) from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'desconto na cabeça do pedido
   SQL = "select valor_desconto from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_TOTAL_PRODUTO_N = 0

   SQL = "select sum(valor_item*qtd_pedida) from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_TOTAL_PRODUTO_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_TOTAL_N = VALOR_TOTAL_PRODUTO_N - VALOR_DESCONTO_N

   txtValorTotalNota.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
   txtValorProdutos.Text = Format(VALOR_TOTAL_PRODUTO_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOTAIS_NOTA"
End Sub

Sub CALCULA_IMPOSTO()
'On Error GoTo ERRO_TRATA

'============================================================
'(+) vProd     (id:W07)                'vem da rotina total nota (pedido/devolução)
'(-) vDesc     (id:W10)                'vem da rotina total nota (pedido/devolução)
'(+) vICMSST   (id:W06)
'(+) vFrete    (id:W09)                'informado no campo
'(+) vSeg      (id:W10)
'(+) vOutro    (id:W15)                'informado no campo
'(+) vII       (id:W11)
'(+) vIPI      (id:W12)                'informado no campo
'(+) vServ     (id:W19) (NT 2011/004)
'============================================================
Dim strCFOP_ITEM As String
Dim CFOP_ID_N As Integer

   VLR_FRETE_N = 0 & txtFrete.Text
   Vlr_Desconto_N = 0 & txtDesconto.Text
   VLR_OUTROS_N = 0 & txtValorOutros.Text

   Vlr_BaseICMSub_N = 0 & txtBaseIcmsSub.Text
   Vlr_ICMSub_N = 0 & txtVlrIcmsSub.Text

   Vlr_TotIPI_N = 0 & txtValorIPI.Text

   'valor total dos produtos
   Vlr_TotProdutos_N = 0 & txtValorProdutos.Text

   Vlr_TotNFe_N = Vlr_TotProdutos_N + VLR_FRETE_N + VLR_OUTROS_N + Vlr_TotIPI_N + Vlr_ICMSub_N - Vlr_Desconto_N
   txtValorTotalNota.Text = Format(Vlr_TotNFe_N, strFormatacao2Digitos)

   'base de calculo do icms normal
   'A base de cálculo do ICMS é o montante da operação, incluindo o frete e despesas acessórias cobradas do adquirente/consumidor.
   'Sobre a respectiva base de cálculo se aplicará a alíquota do ICMS respectiva
   Vlr_BaseCalculo_N = Vlr_TotNFe_N

   'aliquota ICMS normal
   'Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, Trim(txtUF_CLIENTE.Text), 0)


   If IsNumeric(strCFOP_ITEM) Then _
      CFOP_ID_N = 0 & strCFOP_ITEM
   Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, "", CFOP_ID_N)


   If Trim(txtUF_CLIENTE.Text) = Trim(UF_EMPRESA_A) Then      'dentro do estado
      Aliquota_ICMS_Normal_N = 0 & ALIQUOTA_ICMS_NORMAL_DENTRO_UF
      Else: Aliquota_ICMS_Normal_N = 0 & ALIQUOTA_ICMS_NORMAL_FORA_UF
   End If
   Vlr_TotICMS_N = 0 & Vlr_BaseCalculo_N * Aliquota_ICMS_Normal_N / 100

'ICMS NORMAL
   txtBaseCalculo.Text = Format(Vlr_BaseCalculo_N, strFormatacao2Digitos)
   txtValorICMS.Text = Format(Vlr_TotICMS_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CALCULA_IMPOSTO"
End Sub

Private Sub MONTA_NOTA_SAIDA()
'On Error GoTo ERRO_TRATA

   PESSOA_ID_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then  'CHECA TIPO VENDA
      If TabCabeca!STATUS = 1 Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         MsgBox "Não é permitido emitir nota para Orçamento."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'fraImposto.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If
      If TabCabeca!STATUS = 2 And Trim(TIPO_NFe_GERAR) = "R" Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
         MsgBox "É necessário fazer faturamento antes de emitir nota fiscal."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'fraImposto.Enabled = False
         Frame8.Enabled = False
         Unload Me
         Exit Sub
      End If
      If TabCabeca!STATUS = 4 Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         MsgBox "Cupom fiscal já emitido para essa Pedido."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'fraImposto.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      'passou do PEDIDO, checar na tabela nf agora
      SQL = "SELECT NF.NF_ID, NF.PESSOA_ID, NF.NF_TIPO, NF.NUMR_NOTA, NF.SERIE_NOTA, NF.DT_EMISSAO, NF.DT_ENTRASAI, "
      SQL = SQL & " NF.TRANSP_ID, NF.STATUS AS SitNota, NF.DT_CANCELA, NF.QTD_VOLUME, NF.PESO_BRUTO, "
      SQL = SQL & " NF.PESO_LIQUIDO, NF.NUMR_REQ_DEV, NF.indPres, NF.idDest, NF.ESTABELECIMENTO_ID, "
      SQL = SQL & " NF.MODELO_DOC, PEDIDO.PEDIDO_ID, PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.VENDEDOR_ID,"
      SQL = SQL & " Pedido.CGCCPF , Pedido.USUARIO_ID, Pedido.TIPO_REGISTRO, Pedido.NOME_CLIENTE"
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDONF WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDONF.PEDIDO_ID"
      SQL = SQL & " INNER JOIN NF WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDONF.NF_ID = NF.NF_ID "

      SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         txtNOTA.Text = "" & TabNOTA!NUMR_NOTA
         txtSerie.Text = "" & TabNOTA!SERIE_NOTA
         txtMODELO.Text = "" & TabNOTA.Fields("modelo_doc").Value
         txtDtEmis.Text = "" & Format(TabNOTA!DT_EMISSAO, "dd/mm/yyyy")
         txtDtSaida.Text = "" & Format(TabNOTA!DT_ENTRASAI, "dd/mm/yyyy")

         If Not IsNull(TabNOTA!TRANSP_ID) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cnpjcpf,descricao from vwTRANSPORTADORA WITH (NOLOCK)"
            SQL = SQL & " where cnpjcpf = '" & TabNOTA!TRANSP_ID & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp.Fields(0).Value) Then
                  cmbCNPJCPF_TRANSP.Text = "" & Trim(TabTemp!CNPJCPF) & " - " & Trim(TabTemp!DESCRICAO)
                  txtCNPJCPF_TRANSP.Text = "" & Trim(TabTemp.Fields(0).Value)
                  'Volumes
                  TxtQuantidadeRodapeNota.Text = "" & TabNOTA!Qtd_Volume
                  TxtEspecie.Text = "UN"
                  TxtPesoBruto.Text = "" & TabNOTA!Peso_Bruto
                  TxtPesoLiquido.Text = "" & TabNOTA!PESO_LIQUIDO
               End If
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         MOSTRA_NOTA_TELA

         If TabCabeca.State = 1 Then _
            TabCabeca.Close
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         'aqui
         MsgBox "Já existe nota fiscal emitida para Pedido = " & PEDIDO_ID_N & " ; Nota Fiscal = " & txtNOTA & " ; Empresa = " & EMPRESA_ID_N

         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'fraImposto.Enabled = False
         Frame8.Enabled = False
         Exit Sub
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      MOSTRA_NOTA_TELA
      Else
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         MsgBox "Registro de venda não encontrado."
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         'fraImposto.Enabled = False
         Exit Sub
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_NOTA_SAIDA"
End Sub

Private Sub MOSTRA_NOTA_TELA()
'On Error GoTo ERRO_TRATA

   Dim NOME_A        As String
   Dim VENDEDOR_ID_N As Long
   Dim TIPO_NOTA_A   As String

   NOME_A = ""

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from vwVendedor WITH (NOLOCK) "
   SQL = SQL & " where vendedor_id = " & VENDEDOR_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      NOME_A = Trim(TabUSU!DESCRICAO)
      VENDEDOR_ID_N = TabUSU!VENDEDOR_ID

      If Trim(TIPO_NFe_GERAR) = "R" Then _
         Me.Caption = Me.Caption & "Emissão Nota Fiscal de Saída" & " ; Vendedor = " & Trim(TabUSU!DESCRICAO)
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If txtDtEmis.Text = "" Then _
      txtDtEmis.Text = Format(Date, "dd/mm/yyyy")
   If txtDtSaida.Text = "" Then _
      txtDtSaida.Text = Format(Date, "dd/mm/yyyy")

   If txtSerie.Text = "" Then
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      'SQL = "select seq_nota_saida, serie_nota_saida from EMPRESA WITH (NOLOCK)"
      'SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
      'SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
      'SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
      'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      'SQL = "select seq_nota_saida from ESTABELECIMENTO WITH (NOLOCK)"
      'SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      'TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      'If Not TabNOTA.EOF Then
         'txtNota.Text = TabNOTA.Fields(0).Value + 1
         'txtNota.Refresh
         txtSerie.Text = 1
         txtSerie.Refresh
      'End If
      'If TabNOTA.State = 1 Then _
         TabNOTA.Close
   End If

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select PESSOA.PESSOA_ID from PEDIDO "
   SQL = SQL & " INNER JOIN CLIENTE "
   SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID "
   SQL = SQL & " INNER JOIN PESSOA "
   SQL = SQL & " ON CLIENTE.PESSOA_ID = PESSOA.PESSOA_ID"
SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      PESSOA_ID_N = 0 & TabUSU.Fields(0).Value
   If TabUSU.State = 1 Then _
      TabUSU.Close
'=======================
      cmbEmail.Text = ""
      INDR_PRI = False

      If TabEmail.State = 1 Then _
         TabEmail.Close

         SQL = "select EMAIL.EMAIL from PESSOA WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN CLIENTE WITH (NOLOCK)"
         SQL = SQL & " ON PESSOA.PESSOA_ID = CLIENTE.PESSOA_ID "
         SQL = SQL & " INNER JOIN EMAIL WITH (NOLOCK)"
         SQL = SQL & " ON PESSOA.PESSOA_ID = EMAIL.PESSOA_ID"
         SQL = SQL & " where email.pessoa_id = " & PESSOA_ID_N
         TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabEmail.EOF
            cmbEmail.AddItem "" & Trim(TabEmail.Fields(0).Value)
            cmbEmail.Text = "" & Trim(TabEmail.Fields(0).Value)
            TabEmail.MoveNext
         Wend
      If TabEmail.State = 1 Then _
         TabEmail.Close

      Dim i
      For i = 1 To Len(cmbEmail.Text)
         If Mid(cmbEmail.Text, i, 1) <> " " Then
            If Mid(cmbEmail.Text, i, 1) = "@" Then
               INDR_PRI = False
               Exit For
            End If
            Else
               Exit For
         End If
      Next

      cmbEmail.ForeColor = vbRed
      cmbEmail.Refresh
      '================================
      cmbIE.Clear
      If TabEmail.State = 1 Then _
         TabEmail.Close

      SQL = "select numr_ie from IE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEmail.EOF Then
         If Trim(TabEmail.Fields("numr_ie").Value) <> "" Then
            cmbIE.Text = Trim(TabEmail.Fields("numr_ie").Value)
            Else: MsgBox "Inscrição Estadual inválida !!!"
         End If
         'Else: MsgBox "Inscrição Estadual inválida !!!"
      End If
      If TabEmail.State = 1 Then _
         TabEmail.Close

      SQL = "select * from FONE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero <> ''"
      TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEmail.EOF Then
         txtFone.Text = "" & Trim(Left(TabEmail.Fields("numero").Value, 10))
         Else: MsgBox "Fone de cliente não encontrado !!!"
      End If
      If TabEmail.State = 1 Then _
         TabEmail.Close
'=============================
      MOSTRA_CLIENTE

   If Trim(cmbCFOPAux.Text) <> "" Then
      NaturezaOperacao_A = "" & Trim(TRAZ_CFOP(Trim(cmbCFOPAux.Text)))
      Else: MsgBox "Problemas no CFOP."
   End If

   TOTAIS_NOTA

   GRID_DP

   SETA_GRID

   If Trim(TIPO_NFe_GERAR) = "R" Then _
      txtMSG.Text = " Numero Pedido: " & PEDIDO_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_NOTA_TELA"
End Sub

Private Sub MOSTRA_CLIENTE()
'On Error GoTo ERRO_TRATA

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select CLIENTE.CLIENTE_ID, CLIENTE.PESSOA_ID, CLIENTE.estabelecimento_ID, "
   SQL = SQL & " CLIENTE.STATUS, PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO,"
   SQL = SQL & " codg_suframa "
   SQL = SQL & " from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON CLIENTE.PESSOA_ID = PESSOA.PESSOA_ID"
   SQL = SQL & " where CLIENTE.pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
      txtEmitente.Text = Trim(TabCliente.Fields("descricao").Value)
      Select Case Len(Trim(TabCliente.Fields("cnpjcpf").Value))
         Case Is <= 11
            txtCNPJCPF.Mask = "###.###.###-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = Trim(TabCliente.Fields("cnpjcpf").Value)
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "R"
         Case Is = 14
            txtCNPJCPF.Mask = "##.###.###/####-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = Trim(TabCliente.Fields("cnpjcpf").Value)
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "C"
      End Select
'========================== endereço cliente
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select ENDERECO_ID, PESSOA_ID, ENDERECO.CEP_ID, RUA, BAIRRO, COMPLEMENTO, "
      SQL = SQL & " TIPO , Numero, Cidade, UF, IBGE_ID"
      SQL = SQL & " from ENDERECO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
      SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

      SQL = SQL & " where ENDERECO.pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and tipo = 'C'"

      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

      If tabEndereco.EOF Then
         If tabEndereco.State = 1 Then _
            tabEndereco.Close
         MsgBox "Não achou endereço do cliente, verifique."
         Exit Sub
      End If
      txtEnd.Text = "" & tabEndereco!Rua
      txtBairro.Text = "" & tabEndereco!Bairro

      If Not IsNull(tabEndereco!Complemento) Then
         If txtEnd.Text = "" Then
            txtEnd.Text = tabEndereco!Complemento
            Else: txtEnd.Text = txtEnd.Text & " , " & tabEndereco!Complemento
         End If
      End If

      If Not IsNull(tabEndereco!CEP_ID) Then  'CEP
         txtCep.Text = "" & tabEndereco!CEP_ID

         If TabCEP.State = 1 Then _
            TabCEP.Close

         SQL = "select * from CEP WITH (NOLOCK)"
         SQL = SQL & " where cep_ID = '" & Trim(txtCep.Text) & "'"
         TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCEP.EOF Then
            If Not IsNull(TabCEP!CIDADE) Then _
               txtCidade.Text = TabCEP!CIDADE
            If Not IsNull(TabCEP!UF) Then _
               txtUF_CLIENTE.Text = TabCEP!UF
            If IsNull(TabCEP!IBGE_ID) Then
               MsgBox "Código IBGE inválido !!!"
               Else
                  If Trim(TabCEP!IBGE_ID) = "" Then
                     MsgBox "Código IBGE inválido !!!"
                     Else: txtIBGE.Text = TabCEP!IBGE_ID
                  End If
            End If
         End If
         If TabCEP.State = 1 Then _
            TabCEP.Close
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close
'==========================
      txtSerie.Text = SERIE_NFe_A

      If Trim(UF_EMPRESA_A) = "" Then _
         PEGA_DADOS_EMPRESA

      If Trim(TIPO_NFe_GERAR) = "R" Then
         If Trim(txtUF_CLIENTE.Text) = Trim(UF_EMPRESA_A) Then      'dentro do estado
            cmbCFOPAux.Text = "" & CFOP_SAIDA_DENTRO_UF_N

            cmbLocalAUX.Text = 1
            cmbLocal.Text = "Opereção Interna"

'===================================== 'para cupom fiscal vinculado a nota fiscal
            If Not IsNull(TabCabeca.Fields("status").Value) Then
               If TabCabeca.Fields("status").Value = 7 Then
                  cmbCFOPAux.Text = 5929  'Lançamento efetuado em decorrência de emissão de documento fiscal relativo a operação ou prestação também registrada em equipamento Emissor de Cupom Fiscal - ECF
                  txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & "NFe referente cupom fiscal nº: " & TRAZ_NUMERO_CUPOM

                  If TabConsulta.State = 1 Then _
                     TabConsulta.Close

                  SQL = "select nf_id from NF WITH (NOLOCK)"
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabConsulta.EOF Then
                     SQL = "update NFITEM set CFOP_id = 5929"
                     SQL = SQL & " where nf_id = " & TabConsulta.Fields("nf_id").Value
                     CONECTA_RETAGUARDA.Execute SQL
                  End If
                  If TabConsulta.State = 1 Then _
                     TabConsulta.Close
               End If
            End If
'=====================================
            Else                                                        'fora do estado
               cmbCFOPAux.Text = CFOP_SAIDA_FORA_UF_N

               cmbLocalAUX.Text = 2
               cmbLocal.Text = "Operação interestadual"

'===================================== 'para cupom fiscal vinculado a nota fiscal
               If Not IsNull(TabCabeca.Fields("status").Value) Then
                  If TabCabeca.Fields("status").Value = 7 Then
                     cmbCFOPAux.Text = 6929  'Lançamento efetuado em decorrência de emissão de documento fiscal relativo a operação ou prestação também registrada em equipamento Emissor de Cupom Fiscal - ECF
                     txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & "NFe referente cupom fiscal nº: " & TRAZ_NUMERO_CUPOM
                  
                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close

                     SQL = "select nf_id from NF WITH (NOLOCK)"
                     SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                     TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabConsulta.EOF Then
                        SQL = "update NFITEM set CFOP_id = 6929"
                        SQL = SQL & " where nf_id = " & TabConsulta.Fields("nf_id").Value
                        CONECTA_RETAGUARDA.Execute SQL
                     End If

                     If TabConsulta.State = 1 Then _
                        TabConsulta.Close
                  End If
               End If
'=====================================
         End If
         Else  'não é nota de saida
            If Trim(txtUF_CLIENTE.Text) = Trim(UF_EMPRESA_A) Then
               cmbCFOPAux.Text = CFOP_DEVOLUCAO_SAI_DENTRO_UF_N
               Else: cmbCFOPAux.Text = CFOP_DEVOLUCAO_SAI_FORA_UF_N
            End If
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
'===========suframa
      If Not IsNull(TabCliente.Fields("codg_suframa").Value) Then
         If Trim(TabCliente.Fields("codg_suframa").Value) <> "" Then
            If IsNumeric(TabCliente.Fields("codg_suframa").Value) Then
               If Len(TabCliente.Fields("codg_suframa").Value) > 4 Then
                  CODG_SUFRAMA_A = Trim(TabCliente.Fields("codg_suframa").Value)
                  cmbCFOPAux.Text = 6110  'Venda de mercadoria adquirida ou recebida de terceiros, destinada à Zona Franca de Manaus ou Áreas de Livre Comércio
                  NaturezaOperacao_A = Trim(TRAZ_CFOP(Trim(cmbCFOPAux.Text)))
                  cmbCFOP.Text = 6110 & "-" & NaturezaOperacao_A
                  txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & " " & "Codigo Suframa: " & Trim(CODG_SUFRAMA_A)
               End If
            End If
         End If
      End If
      Else
         If TabCliente.State = 1 Then _
            TabCliente.Close

         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         MsgBox "Cliente não cadastrado !!!"
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         Exit Sub
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CLIENTE"
End Sub

Private Sub MOSTRA_FORNECEDOR()
'On Error GoTo ERRO_TRATA

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select * from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      cmbIE.Clear
      cmbIE.Text = TabFornecedor!IE
      cmbIE.AddItem TabFornecedor!IE
      txtEmitente.Text = TabFornecedor!NOME
      Select Case Len(TabFornecedor!CNPJCPF)
         Case Is <= 11
            txtCNPJCPF.Mask = "###.###.###-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = TabFornecedor!CNPJCPF
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "R"
         Case Is = 14
            txtCNPJCPF.Mask = "##.###.###/####-##"
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = TabFornecedor!CNPJCPF
            txtCNPJCPF.PromptInclude = True
            Tipo_Endereço = "C"
      End Select

      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      'endereço
      SQL = "select * from ENDERECO WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and tipo = '" & Tipo_Endereço & "'"
      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabEndereco.EOF Then
         If Not IsNull(tabEndereco!Rua) Then _
            txtEnd.Text = tabEndereco!Rua
         If Not IsNull(tabEndereco!Complemento) Then
            If txtEnd.Text = "" Then
               txtEnd.Text = tabEndereco!Complemento
               Else: txtEnd.Text = txtEnd.Text & " , " & tabEndereco!Complemento
            End If
         End If

         If Not IsNull(tabEndereco!Bairro) Then _
            txtBairro.Text = tabEndereco!Bairro

         If Not IsNull(tabEndereco!CEP_ID) Then  'CEP
            txtCep.Text = "" & tabEndereco!CEP_ID

            If TabCEP.State = 1 Then _
               TabCEP.Close

            SQL = "select * from CEP WITH (NOLOCK)"
            SQL = SQL & " where cep_ID = '" & Trim(txtCep.Text) & "'"
            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCEP.EOF Then
               If Not IsNull(TabCEP!CIDADE) Then _
                  txtCidade.Text = TabCEP!CIDADE
               If Not IsNull(TabCEP!UF) Then _
                  txtUF_CLIENTE.Text = TabCEP!UF
               If IsNull(TabCEP!IBGE_ID) Then
                  MsgBox "Código IBGE inválido !!!"
                  Else
                     If Trim(TabCEP!IBGE_ID) = "" Then
                        MsgBox "Código IBGE inválido !!!"
                        Else: txtIBGE.Text = TabCEP!IBGE_ID
                     End If
               End If
            End If
            If TabCEP.State = 1 Then _
               TabCEP.Close
'=================
         End If
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      'Transportadora

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from EMPRESA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
      SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
      SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabEmpresa.EOF Then
         'SERIE NOTA SAIDA
         If Not IsNull(tabEmpresa!SERIE_NOTA_SAIDA) Then _
            txtSerie.Text = tabEmpresa!SERIE_NOTA_SAIDA

         'If Not IsNull(TabEmpresa!Instrucao_Fisco) Then _
            txtMSG.Text = TabEmpresa!Instrucao_Fisco

         txtNOTA.Refresh
         Else
            If TabTemp.State = 1 Then _
               TabTemp.Close

            MsgBox "Erro no arquivo de empresa."
            Exit Sub
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      If Trim(UF_EMPRESA_A) = "" Then _
         PEGA_DADOS_EMPRESA

      If Trim(UF_EMPRESA_A) = "" Then
         MsgBox "Impossivel continuar, inconsitencia na UNIDADE FEDERAÇÃO cadastrada para empresa."
         Exit Sub
         Else
            'devolução de compra
            If Trim(TIPO_NFe_GERAR) = "DC" Then
               If Trim(txtUF_CLIENTE.Text) = Trim(UF_EMPRESA_A) Then
               'tem que ver a nota de origem para determinar o cfop, por enquanto vai o que está no combo
                  cmbCFOPAux.Text = CFOP_DEVOLUCAO_ENTRADA_DENTRO_UF_N     '5411
                  Else: cmbCFOPAux.Text = CFOP_DEVOLUCAO_ENTRADA_FORA_UF_N '6411
                End If
            End If
      End If
      NaturezaOperacao_A = Trim(TRAZ_CFOP(Trim(cmbCFOPAux.Text)))
      cmbCFOP.Text = cmbCFOPAux.Text & "-" & NaturezaOperacao_A
      Else
         If TabFornecedor.State = 1 Then _
            TabFornecedor.Close
         If TabCABENTRA.State = 1 Then _
            TabCABENTRA.Close

         MsgBox "Fornecedor Não Cadastrado !!!"
         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         Frame6.Enabled = False
         'fraImposto.Enabled = False
         Exit Sub
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   txtCNPJCPF.PromptInclude = False

   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   SQL = "select * from FONE WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and numero <> ''"
   tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabEndereco.EOF Then
      txtFone.Text = "" & Trim(Left(tabEndereco.Fields("numero").Value, 10))
      Else: MsgBox "Fone de cliente não encontrado !!!"
   End If

   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_FORNECEDOR"
End Sub

Private Sub TOTAIS_NOTA_DVolução() 'Devolução de Entrada
'On Error GoTo ERRO_TRATA

   PERC_DESCONTO_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0

   If TabItem.State = 1 Then _
      TabItem.Close

   SQL = "select sum(preco_custo*qtde_entrada) WITH (NOLOCK)"
   SQL = SQL & " from NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM WITH (NOLOCK)"
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"
   SQL = SQL & " where NOTAENTRADA.entrada_id = " & TabCABENTRA!ENTRADA_ID
   TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabItem.EOF Then _
      If Not IsNull(TabItem.Fields(0).Value) Then _
         VALOR_ITEM_N = TabItem.Fields(0).Value
   If TabItem.State = 1 Then _
      TabItem.Close

   txtValorTotalNota.Text = Format(VALOR_ITEM_N + TabCABENTRA!VALOR_IPI, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
   txtValorProdutos.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   txtBaseCalculo.Text = Format(TabCABENTRA!BASE_CALC_ICMS, strFormatacao2Digitos)
   txtValorICMS.Text = Format(TabCABENTRA!VALOR_ICMS, strFormatacao2Digitos)
   txtValorIPI.Text = Format(TabCABENTRA!VALOR_IPI, strFormatacao2Digitos)

   txtBaseIcmsSub.Text = Format(0, strFormatacao2Digitos)
   txtVlrIcmsSub.Text = Format(0, strFormatacao2Digitos)
   txtFrete.Text = Format(0, strFormatacao2Digitos)
   txtValorOutros.Text = Format(0, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOTAIS_NOTA_DVolução"
End Sub

Private Sub GRID_DP()
'On Error GoTo ERRO_TRATA

   Dim Desconto_Item_Fat   As Double
   Dim Valr_Item_Fat       As Double

   ListaDP.ListItems.Clear

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l WITH (NOLOCK)"
   SQL = SQL & " where l.numr_doc = " & PEDIDO_ID_N   'TabCABECA!PEDIDO_ID
   SQL = SQL & " and l.lancamento_id = i.lancamento_id "
   SQL = SQL & " and l.tipo_lancamento = 1 "
   SQL = SQL & " order by i.seq"
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLancamento.EOF
      If TabLancamento!FORMAPAGTO_ID = 9999 Then
         ListaDP.ForeColor = vbRed
         Set item = ListaDP.ListItems.Add(, "seq." & TabLancamento!SEQ, TabLancamento!SEQ)
         Else
            ListaDP.ForeColor = vbBlue
            Set item = ListaDP.ListItems.Add(, "seq." & TabLancamento!SEQ, TabLancamento!SEQ)
      End If

      Desconto_Item_Fat = 0 & TabLancamento!Valor_Desconto
      Valr_Item_Fat = 0 & TabLancamento!Valor_Item

      item.SubItems(1) = Format(Valr_Item_Fat - Desconto_Item_Fat, strFormatacao2Digitos)
      item.SubItems(2) = TabLancamento!Numr_doc & "-" & TabLancamento!SEQ
      item.SubItems(3) = TabLancamento!DT_VENCIMENTO

      If Not IsNull(TabLancamento!FORMAPAGTO_ID) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
         SQL = SQL & " where formapagto_id = " & TabLancamento!FORMAPAGTO_ID
         SQL = SQL & " and status = 'true' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            item.SubItems(4) = Trim(TabTemp!DESCRICAO) '& " ; "
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      TabLancamento.MoveNext
   Wend
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRID_DP"
End Sub
'=================== ROTINA DE IMPRESSÃO DE NOTA FISCAL
Sub GERAR_NFE()
'On Error GoTo ERRO_TRATA

   Dim VALOR_01      As Double
   Dim VALOR_02      As Double
   Dim strTributacao As String

   If Trim(cmbPresencaAUX.Text) = "" Then
      MsgBox "Selecione (Indicador de presença do comprador)."
      Exit Sub
   End If
   If Trim(cmbLocalAUX.Text) = "" Then
      MsgBox "Selecione (Identificador de local de destino da operação)."
      Exit Sub
   End If

   'validar transportadora
   If Trim(cmbCNPJCPF_TRANSP.Text) = "" Then
      MsgBox "Informar trasportadora."
      txtCNPJCPF_TRANSP.SetFocus
      Exit Sub
   End If

   txtCNPJCPF_TRANSP.PromptInclude = False
   TRANSP_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PESSOA.CNPJCPF, PESSOA.DESCRICAO, TRANSPORTADORA.PESSOA_ID, transp_id"
   SQL = SQL & " from TRANSPORTADORA WITH (NOLOCK)"
   SQL = SQL & " Inner Join PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON TRANSPORTADORA.PESSOA_ID = PESSOA.PESSOA_ID"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      'MsgBox "Informar trasportadora."
      'txtCNPJCPF_TRANSP.SetFocus
      'Exit Sub
      Else: TRANSP_ID_N = 0 & TabTemp.Fields("transp_id").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'validar cliente
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "CNPJ/CPF inválido !!!"
      Exit Sub
   End If
   If Trim(txtEnd.Text) = "" Then
      MsgBox "Endereço inválido !!!"
      txtEnd.SetFocus
      Exit Sub
   End If
   If Trim(txtBairro.Text) = "" Then
      MsgBox "Bairro inválido !!!"
      txtBairro.SetFocus
      Exit Sub
   End If
   If Trim(txtUF_CLIENTE.Text) = "" Then
      MsgBox "UF inválido !!!"
      txtUF_CLIENTE.SetFocus
      Exit Sub
   End If
   If Trim(txtCep.Text) = "" Then
      MsgBox "CEP inválido !!!"
      txtCep.SetFocus
      Exit Sub
   End If
   If Trim(txtChaveNFe.Text) <> "" Then
      If Len(Trim(txtChaveNFe.Text)) <> 44 Then
         MsgBox "Chave informada inválida, verifique."
         txtChaveNFe.Text = ""
         txtChaveNFe.SetFocus
         Exit Sub
      End If
    End If

   CRITERIO_A = txtCep.Text
   CRITERIO_A = Replace(CRITERIO_A, "-", "")
   If Len(Trim(CRITERIO_A)) < 8 Then
      MsgBox "CEP inválido, deve conter 8 digitos !!!"
      txtCep.SetFocus
      Exit Sub
   End If

   If Trim(cmbIE.Text) = "" Then _
      cmbIE.Text = "ISENTO"

   If Trim(txtCidade.Text) = "" Then
      MsgBox "Cidade inválido !!!"
      txtCidade.SetFocus
      Exit Sub
   End If
   If Trim(txtIBGE.Text) = "" Then
      MsgBox "Código IBGE inválido !!!"
      txtIBGE.SetFocus
      Exit Sub
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from IBGE WITH (NOLOCK)"
         SQL = SQL & " where IBGE_ID = " & txtIBGE.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
            Exit Sub
            Else
               If Not IsNull(TabTemp.Fields("estado").Value) Then
                  If Trim(UCase(TabTemp.Fields("estado").Value)) <> Trim(UCase(txtUF_CLIENTE.Text)) Then
                     MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
                     Exit Sub
                  End If
                  Else
                     MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
                     Exit Sub
               End If
               If Not IsNull(TabTemp.Fields("municipio").Value) Then
                  Else
                     MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
                     Exit Sub
               End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
   End If

frmINTEGRA.INTEGRA_IBGE txtIBGE.Text

   If Trim(txtFone.Text) = "" Then
      MsgBox "Fone inválido !!!"
      txtFone.Enabled = True
      txtFone.SetFocus
      Exit Sub
   End If

   If TxtQuantidadeRodapeNota.Text = "" Then
      MsgBox "Digite Quantidade de Produtos a Transportar!"
      TxtQuantidadeRodapeNota.SetFocus
      Exit Sub
      Else
         If Not IsNumeric(TxtQuantidadeRodapeNota.Text) Then
            MsgBox "Favor Digite Quantidade de Produtos a Transportar!"
            TxtQuantidadeRodapeNota.SetFocus
            Exit Sub
         End If
   End If

   If TxtEspecie.Text = "" Then
      MsgBox "Digite a Especie a Transportar!"
      TxtEspecie.SetFocus
      Exit Sub
   End If

   If TxtPesoBruto.Text = "" Then
      MsgBox "Digite o Peso Bruto a Transportar!"
      TxtPesoBruto.SetFocus
      Exit Sub
   End If

   If TxtPesoLiquido.Text = "" Then
      MsgBox "Digite o Peso Liquido a Transportar!"
      TxtPesoLiquido.SetFocus
      Exit Sub
   End If

   txtCNPJCPF.PromptInclude = False

'===================== NOTA DE PEDIDO VENDA
   If Trim(TIPO_NFe_GERAR) = "R" Then
      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      'chegando itens
      SQL = "select PEDIDO.*, CLIENTE.PESSOA_ID "
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN CLIENTE WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID"

      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCabeca.EOF Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
         MsgBox "Pedido não encontrado."
         Exit Sub
      End If
      If Not TabCabeca.EOF Then
         PESSOA_ID_N = TabCabeca.Fields("pessoa_id").Value

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabCabeca.Fields("pedido_id").Value
         SQL = SQL & " and tipo_reg = 'PC' "
         SQL = SQL & " and pedidoitem.status <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close
            MsgBox "Pedido não possue itens, não permitido."
            Exit Sub
         End If
         While Not TabTemp.EOF
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from PRODUTO WITH (NOLOCK)"
            SQL = SQL & " where produto_id = " & TabTemp.Fields("produto_id").Value
            SQL = SQL & " and situacao <> 'C' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabConsulta.EOF Then
               'If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               MsgBox "Pedido Produto não cadastrado, não permitido."
               Exit Sub
               Else
                  If IsNull(TabConsulta.Fields("situacao").Value) Then
                     MsgBox "Situação do produto inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        If Trim(TabConsulta.Fields("situacao").Value) <> "A" Then
                           MsgBox "Situação do produto inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                        End If
                  End If

                  QTDE_PEDIDO = 0
                  QTDE_PEDIDO = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabConsulta.Fields("produto_id").Value)

                  If INDR_ESTQ_NEGATIVO = False Then
                     If Indr_Consulta = False Then
                        If Not IsNull(TabTemp.Fields("STATUS").Value) Then
                           If Trim(UCase((TabTemp.Fields("STATUS").Value))) <> "B" Then
                              If IsNull(QTDE_PEDIDO) Then
                                 MsgBox "Qtde disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                                 Exit Sub
                                 Else
                                    If QTDE_PEDIDO <= 0 Then
                                       MsgBox "Qtde disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                                       Exit Sub
                                    End If
                              End If
                           End If
                           Else: MsgBox "Verificar situação do item no pedido."
                        End If
                     End If
                  End If

                  If IsNull(TabConsulta.Fields("PRECO_VENDA").Value) Then
                     MsgBox "Valor_venda disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        VALOR_ITEM_N = TabConsulta.Fields("PRECO_VENDA").Value
                        If VALOR_ITEM_N <= 0 Then
                           MsgBox "Valor_venda disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                           Else
                              VALOR_01 = TabConsulta.Fields("PRECO_VENDA").Value
                              VALOR_01 = TabTemp.Fields("valor_item").Value
                              VALOR_02 = TabConsulta.Fields("PRECO_CUSTO").Value
                              If VALOR_01 < VALOR_02 Then
                                 MsgBox "Produto : " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value) & " , venda abaixo preço de custo."
                              End If
                        End If
                  End If
                  If IsNull(TabConsulta.Fields("CODG_NCM").Value) Then
                     MsgBox "CODG_NCM disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        VALOR_01 = TabConsulta.Fields("CODG_NCM").Value
                        If VALOR_01 <= 0 Then
                           MsgBox "Código NCM do produto inválido. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                        End If
                  End If
            End If

            TabTemp.MoveNext

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF_TRANSP.PromptInclude = False

         If Trim(txtNOTA.Text) = "" Then
            txtNOTA.Text = "" & GERA_NUMERO_NFe_N
            txtNOTA.Refresh
         End If

         GRAVA_NOTA txtNOTA.Text, _
                    txtSerie.Text, _
                    txtMODELO.Text, _
                    TIPO_NFe_GERAR, _
                    TxtQuantidadeRodapeNota.Text, _
                    TxtPesoBruto.Text, _
                    TxtPesoLiquido.Text, _
                    cmbPresencaAUX.Text, _
                    cmbLocalAUX.Text, _
                    cmbCFOPAux.Text, _
                    Trim(txtCNPJCPF_TRANSP.Text)

         IMPRESSAO_NF
      End If
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
   End If  'If Trim(TIPO_NFe_GERAR) = "R" Then
   If Trim(TIPO_NFe_GERAR) = "DV" Then
      IMPRESSAO_NF
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_NFE"
End Sub

Private Sub IMPRESSAO_NF()
'On Error GoTo ERRO_TRATA

   CHAMA_SP_GLOBAL
   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRESSAO_NF"
End Sub

Sub CHAMA_SP_GLOBAL()
'On Error GoTo ERRO_TRATA

   Dim TRANSP_ID_N               As Long
   Dim ID_NF_N                   As Long
   Dim DESC_NATUREZA_OPERACAO_A  As String
   Dim CFOP_N                    As String
   Dim MFATIPO_tpnf              As String

   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0
   ID_NF_N = 0
   NF_ID_N = 0
   txtCNPJCPF.PromptInclude = False

   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "Cliente não encontrado, verifique."
      Exit Sub
      Else
         If TabCliente.State = 1 Then _
            TabCliente.Close
         SQL = "select pessoa_id,cliente_id from CLIENTE "
         SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then
            PESSOA_ID_N = 0 & TabCliente.Fields("pessoa_id").Value
            CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
         End If
         If TabCliente.State = 1 Then _
            TabCliente.Close
   End If
   If Trim(txtIBGE.Text) = "" Then
      MsgBox "IBGE não encontrado, verifique."
      Exit Sub
   End If

   Call frmINTEGRA.TRANSPORTADORA_INTEGRA(txtCNPJCPF_TRANSP.Text)

   Call frmINTEGRA.CLIENTE_INTEGRA(txtCNPJCPF.Text)

   Call frmINTEGRA.INTEGRA_IBGE(txtIBGE.Text)

   If TabProduto.State = 1 Then _
      TabProduto.Close
   SQL = "select NFITEM.NF_ID, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID"
   SQL = SQL & " from NF "
   SQL = SQL & " INNER JOIN NFITEM "
   SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID"

   SQL = SQL & " where numr_nota = " & Trim(txtNOTA.Text)
   SQL = SQL & " and modelo_doc = '" & Trim(txtMODELO.Text) & "'"
   SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   
'Debug.Print SQL
   
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      ID_NF_N = 0 & TabProduto.Fields("nf_id").Value
      NF_ID_N = 0 & TabProduto.Fields("nf_id").Value
      frmINTEGRA.INTEGRA_PRODUTO (TabProduto.Fields("produto_id").Value)
      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   DESC_NATUREZA_OPERACAO_A = ""

   SQL = "select distinct(cfop_id) from NFITEM"
   SQL = SQL & " where nf_id = " & ID_NF_N
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      CFOP_N = "" & TabCliente.Fields(0).Value

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select descricao from CFOP "
      SQL = SQL & " where cfop_id = " & CFOP_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then _
         DESC_NATUREZA_OPERACAO_A = "" & Trim(TabCliente.Fields(0).Value)
      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

   If Trim(TxtPesoBruto.Text) = "" Then _
      TxtPesoBruto.Text = 0
   If Trim(TxtPesoLiquido.Text) = "" Then _
      TxtPesoLiquido.Text = 0
   If Trim(TxtQuantidadeRodapeNota.Text) = "" Then _
      TxtQuantidadeRodapeNota.Text = 0

   txtCNPJCPF_TRANSP.PromptInclude = False
   TRANSP_ID_N = 0 & TRAZ_ID_TABELA("vwTRANSPORTADORA", "transp_id", "cnpjcpf", txtCNPJCPF_TRANSP.Text)

   MFATIPO_tpnf = "N"
   'aqui é se for devolução
   If Trim(cmbFinalidadeAUX.Text) <> "" Then
      If Trim(cmbFinalidadeAUX.Text) = "4" Then
         MFATIPO_tpnf = "D"
      End If
   End If

   If Trim(txtFrete.Text) = "" Then _
      txtFrete.Text = "0"

   Call frmINTEGRA.PEDIDO_INTEGRA_MFA010(ID_NF_N, _
                                         TRANSP_ID_N, _
                                         "NFE", _
                                         Trim(txtMSG.Text) & ", " & Trim(txtDadosAdicionais.Text), _
                                         cmbCFinalAUX.Text, _
                                         cmbLocalAUX.Text, _
                                         cmbPresencaAUX.Text, _
                                         txtChaveNFe.Text, _
                                         cmbFinalidadeAUX.Text, _
                                         TxtPesoLiquido.Text, _
                                         TxtPesoBruto.Text, _
                                         cmbFreteAUX.Text, _
                                         DESC_NATUREZA_OPERACAO_A, _
                                         TxtPesoBruto.Text, _
                                         TxtPesoLiquido.Text, _
                                         TxtQuantidadeRodapeNota.Text, _
                                         cmbTpEmisAUX.Text, _
                                         MFATIPO_tpnf, _
                                         txtFrete.Text)

'CAMPO : PESO BRUTO =  MFAVALIMP5  E  peso liquido =  MFAVALIMP6 e Volume = MFAVOLUME1
'MFAVALIMP5="" & TxtPesoBruto.       'peso BRUTO
'peso liquido = MFAVALIMP6
'Volume = MFAVOLUME1

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHAMA_SP_GLOBAL"
End Sub

Sub INICIALIZA_NF()
'On Error GoTo ERRO_TRATA

   Dim TOTAL_N As Long

   fraNota.Enabled = True
   fraEmitente.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   Frame5.Enabled = True
   Frame6.Enabled = True
   txtNFeDev.Visible = False
   lblNFeDev.Visible = False
   ''lblChave.Visible = False
   txtChaveNFe.Visible = False
   txtChaveNFe.Text = ""
   Msg = ""

   txtBaseCalculo.Text = 0 & Format(0, strFormatacao2Digitos)
   txtValorICMS.Text = 0 & Format(0, strFormatacao2Digitos)
   txtBaseIcmsSub.Text = 0 & Format(0, strFormatacao2Digitos)
   txtFrete.Text = 0 & Format(0, strFormatacao2Digitos)
   txtValorIPI.Text = 0 & Format(0, strFormatacao2Digitos)
   txtValorProdutos.Text = 0 & Format(0, strFormatacao2Digitos)
   txtDesconto.Text = 0 & Format(0, strFormatacao2Digitos)
   txtVlrIcmsSub.Text = 0 & Format(0, strFormatacao2Digitos)
   txtValorOutros.Text = 0 & Format(0, strFormatacao2Digitos)
   txtValorTotalNota.Text = 0 & Format(0, strFormatacao2Digitos)
   txtMODELO.Text = "" & "NFE"

   TOTAL_N = 0
   PRODUTO_ID_N = 0

   Me.Caption = Me.Caption & " - " & Me.Name

   cmbFreteAUX.Clear
   cmbFrete.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'L' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFrete.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      cmbFreteAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   cmbFreteAUX.Text = 1
   cmbFrete.Text = "1-Contratação do Frete por conta do Destinatário (FOB)"

   cmbTpEmisAUX.Clear
   cmbTpEmis.Clear
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'A1' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTpEmis.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      cmbTpEmisAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   cmbTpEmisAUX.Text = 1
   cmbTpEmis.Text = "1-Emissão normal (não em contingência)"

'SE É ENTRADA OU SAIDA
   cmbTipoOperaAUX.Clear
   cmbTipoOpera.Clear
   cmbTipoOperaAUX.AddItem 0
   cmbTipoOpera.AddItem "0-Entrada"
   cmbTipoOperaAUX.AddItem 1
   cmbTipoOpera.AddItem "1-Saída"

   cmbTipoOperaAUX.Text = 1
   cmbTipoOpera.Text = "1-Saída"

'====================================================
'Finalidade de emissão da NF-e
'1=NF-e normal;
'2=NF-e complementar;
'3=NF-e de ajuste;
'4=Devolução/Retorno

   cmbFinalidadeAUX.Clear
   cmbFinalidade.Clear
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'D' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbFinalidade.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      cmbFinalidadeAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbFinalidadeAUX.Text = 1
   cmbFinalidade.Text = "1-NF-e normal"

   If Trim(TIPO_NFe_GERAR) = "DV" Or Trim(TIPO_NFe_GERAR) = "DC" Then
      cmbFinalidadeAUX.Text = 4
      cmbFinalidade.Text = "4-Devolução/Retorno"
      txtNFeDev.Visible = True
      lblNFeDev.Visible = True
      ''lblChave.Visible = True
      txtChaveNFe.Visible = True
      'txtNFeDev.SetFocus

      cmbTipoOperaAUX.Text = 0
      cmbTipoOpera.Text = "0-Entrada"
      txtMSG.Text = "Dev.Ref.NFe = "
   End If

   'Marca produto
   cmbMarcaAUX.Clear
   cmbMarca.Clear
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'W' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbMarca.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      cmbMarcaAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close
   cmbMarcaAUX.Text = 0
   cmbMarca.Text = "0-PROPRIO"

   'Indicador de presença do comprador
   cmbPresencaAUX.Clear
   cmbPresenca.Clear
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'K' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbPresenca.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      cmbPresencaAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbPresencaAUX.Text = 1
   cmbPresenca.Text = "1-Operação presencial"

   'Identificador de local de destino da operação
   cmbLocalAUX.Clear
   cmbLocal.Clear
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'Y' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbLocal.AddItem TabDESCR!Codigo & "-" & Trim(TabDESCR!DESCRICAO)
      cmbLocalAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbLocalAUX.Text = 1
   cmbLocal.Text = "1-Opereção Interna"

   cmbAmbienteAUX.Clear
   cmbAmbiente.Clear
   cmbAmbienteAUX.AddItem 1
   cmbAmbiente.AddItem "1-Produção"
   cmbAmbienteAUX.AddItem 2
   cmbAmbiente.AddItem "2-Homologação"

   cmbAmbienteAUX.Text = 1
   cmbAmbiente.Text = "1-Produção"

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then _
      Exit Sub

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select NFETPAMBI from EMPRES WITH (NOLOCK)"
   SQL = SQL & " where empresa = '0" & EMPRESA_ID_N & "'"
   TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabConsulta.EOF Then
      If Not IsNull(TabConsulta.Fields("NFETPAMBI").Value) Then
         If Trim(UCase(TabConsulta.Fields("NFETPAMBI").Value)) = "H" Then
            cmbAmbienteAUX.Text = 2
            cmbAmbiente.Text = "2-Homologação"
            Else
               If Trim(UCase(TabConsulta.Fields("NFETPAMBI").Value)) = "P" Then
                  cmbAmbienteAUX.Text = 1
                  cmbAmbiente.Text = "1-Produção"
               End If
         End If
      End If
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close
'============================================

   If USUARIO_ID_N = 144 Then _
      cmbAmbiente.Enabled = True

   cmbCFinalAUX.Clear
   cmbCFinal.Clear
   cmbCFinalAUX.AddItem 0
   cmbCFinal.AddItem "0-Não"
   cmbCFinalAUX.AddItem 1
   cmbCFinal.AddItem "1-Consumidor Final"

   cmbCFinalAUX.Text = 1
   cmbCFinal.Text = "1-Consumidor Final"

   CARREGA_TRANSPORTADORA
   CARREGA_CFOP

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from OBS WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_EMPRESA_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      If TabConsulta.Fields("seq").Value = 2 Then _
         txtMSG.Text = "" & Trim(TabConsulta.Fields("obs").Value)
      If TabConsulta.Fields("seq").Value = 3 Then _
         txtDadosAdicionais.Text = "" & Trim(TabConsulta.Fields("obs").Value)
      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   If Trim(UF_EMPRESA_A) = "" Then _
      PEGA_DADOS_EMPRESA

   If IsNull(TIPO_NFe_GERAR) Then
      MsgBox "Parametro não informado, erro, TIPO_NFe_GERAR"
      Exit Sub
      Unload Me
   End If
   If Trim(TIPO_NFe_GERAR) = "" Then
      MsgBox "Parametro não informado, erro, TIPO_NFe_GERAR"
      Exit Sub
      Unload Me
   End If

'TIPO_NFe_GERAR : foi carregada lá na tela displayemissor
'================================== PEDIDO DE VENDA
   If PEDIDO_ID_N > 0 Then _
      MONTA_NOTA_SAIDA
'================================== TRANSFERENCIA
   'If UCase(Trim(TIPO_NFe_GERAR)) = UCase("T") Then

      'NOTA DE TRANSFERENCIA
      'If Trim(TIPO_NFe_GERAR) = "T" Then _
         If ((PEDIDO_ID_N > 0) And (EMPRESA_ID_N > 0)) Then _
            MONTA_TRANSFERENCIA

   '   Exit Sub
   'End If

   'If Trim(TIPO_NFe_GERAR ) = "E" Then   'NOTA DE ENTRADA
   'End If
   'If Trim(TIPO_NFe_GERAR ) = "R" Then   'NOTA DE SIMPLES REMESSA
   'End If

   'pegando mensagem fisco que está na tabela cfop
   If Trim(cmbCFOPAux.Text) <> "" Then
      If Trim(txtDadosAdicionais.Text) = "" Then
         txtDadosAdicionais.Text = "" & TRAZ_CFOP_MSG(cmbCFOPAux.Text)
         Else: txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & " " & TRAZ_CFOP_MSG(cmbCFOPAux.Text)
      End If
   End If
   If Trim(TIPO_NFe_GERAR) = "DV" Then
      Me.Caption = Me.Caption & "Emissão Nota Fiscal de Devolução Venda"
      MONTA_NOTA_DV
   End If
   If Trim(TIPO_NFe_GERAR) = "DC" Then _
      Me.Caption = Me.Caption & "Emissão Nota Fiscal de Devolução Compra"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_NF"
End Sub

Private Sub CARREGA_TRANSPORTADORA()
'On Error GoTo ERRO_TRATA

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select * from vwTRANSPORTADORA "
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJ_EMPRESA_N) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      txtCNPJCPF_TRANSP.PromptInclude = False
         txtCNPJCPF_TRANSP.Text = "" & Trim(TabFornecedor!CNPJCPF)
         cmbCNPJCPF_TRANSP.Text = "" & Trim(TabFornecedor!DESCRICAO)
      txtCNPJCPF_TRANSP.PromptInclude = True
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_TRANSPORTADORA"
End Sub

Private Sub CARREGA_CFOP()
'On Error GoTo ERRO_TRATA

   'CFOP
   cmbCFOPAux.Clear
   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from CFOP With (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      Do Until TabDESCR.EOF
         DoEvents
         cmbCFOPAux.AddItem Trim(TabDESCR!CFOP_ID)
         cmbCFOP.AddItem Trim(TabDESCR!CFOP_ID) & "-" & Trim(TabDESCR!DESCRICAO)
         TabDESCR.MoveNext
      Loop
      Else
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         MsgBox "Cadastro de CFOP com problemas. Não foi localizado nenhum codigo de CFOP cadastrado", vbCritical
         Exit Sub
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

 Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_CFOP"
End Sub

Private Sub GERAR_CABEÇALHO_NFe()
'On Error GoTo ERRO_TRATA

   Dim Numero_A         As String
   Dim DDD_N            As Integer
   Dim CODG_IBGE_A      As String

   Print #1, Tab(1); "NOTAFISCAL|1";   'saida
   Print #1, Tab(1); "A|3.10|NFe52110310628919000188550010000000011203740010|";

'ok, checar
   'BUSCA_ENDERECO_PESSOA "C", ""
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   SQL = "select ENDERECO_ID, PESSOA_ID, ENDERECO.CEP_ID, RUA, BAIRRO, COMPLEMENTO, "
   SQL = SQL & " TIPO , Numero, Cidade, UF, IBGE_ID"
   SQL = SQL & " from ENDERECO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
   SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

   SQL = SQL & " where ENDERECO.pessoa_id = " & PESSOA_ID_EMPRESA_N
   SQL = SQL & " and tipo = 'C'"

   tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If tabEndereco.EOF Then
      MsgBox "Não achou endereço da empresa, verifique."
      Exit Sub
   End If

   SP_PROCURA_CEP tabEndereco!CEP_ID

   CODG_IBGE_A = "" & TabCEP!IBGE_ID

   If Len(CODG_IBGE_A) < 7 Then
      MsgBox "IBGE errado para o cep_id = " & tabEndereco!CEP_ID
      Exit Sub
   End If
'================================================
   Dim cUF, cNF, NatOp, indPag, mode, serie, nNF, dhEmi, dhSaiEnt, tpNF, idDest, cMunFG, TpImp, TpEmis, cDV, TpAmb, FinNFe, indFinal, indPres, ProcEmi, VerProc, dhCont, xJust
   Dim LINHA_B

   cUF = ""
   cNF = ""
   NatOp = ""
   indPag = ""
   mode = ""
   serie = ""
   nNF = ""
   dhEmi = ""
   dhSaiEnt = ""
   tpNF = ""
   idDest = ""
   cMunFG = ""
   TpImp = ""
   TpEmis = ""
   cDV = ""
   TpAmb = ""
   FinNFe = ""
   indFinal = ""
   indPres = ""
   ProcEmi = ""
   VerProc = ""
   dhCont = ""
   xJust = ""
   LINHA_B = ""

'Segmento B ficou asim o layout
'"§B|cUF|cNF|NatOp|indPag|mod|serie|nNF|dhEmi|dhSaiEnt|tpNF|idDest|cMunFG|TpImp|TpEmis|cDV|TpAmb|FinNFe|indFinal|indPres|ProcEmi|VerProc|dhCont|xJust"

'obs 1 - Sendo que a Data de Emissao(dhEmi) e Data da saida Entrada(dhSaiEnt) mandar com data e hora no fotmato do arquivo que esta em anexo
'obs 2 - a tag Hora da Saida(hSaiEnt) foi descontinuada retirar do seu layout
'obs 3 - foram incluidos as seguintes tags :  idDest,indFinal,indPres
'================================================

'tem que configurar de acordo com a tabela do estado
cUF = "52"  'Código da UF do emitente do Documento Fiscal.
            'Código da UF do emitente do Documento Fiscal. Utilizar a
            'Tabela do IBGE de código de unidades da federação
            '(Anexo IX - Tabela de UF, Município e País).

cNF = PEDIDO_ID_N  'Código Numérico que compõe a Chave de Acesso.
                  'Número aleatório gerado pelo emitente para cada NF-e
                  'para evitar acessos indevidos da NF-e. (v2.0)
NaturezaOperacao_A = Trim(TRAZ_CFOP(Trim(cmbCFOPAux.Text)))
NaturezaOperacao_A = Trim(TRAZ_CFOP(Trim(CFOP_ID_N)))
NatOp = Trim(Left(NaturezaOperacao_A, 60))         'Descrição da Natureza da Operação
                                                   'Informar a natureza da operação de que decorrer a saída
                                                   'ou a entrada, tais como: venda, compra, transferência,
                                                   'devolução, importação, consignação, remessa (para fins
                                                   'de demonstração, de industrialização ou outra), conforme
                                                   'previsto na alínea 'i', inciso I, art. 19 do CONVÊNIO S/Nº,
                                                   'de 15 de dezembro de 1970.

indPag = 0  'Indicador da forma de pagamento       '0=Pagamento à vista;
                                                   '1=Pagamento a prazo;
                                                   '2=Outros.

mode = 55                                          'Código do Modelo do Documento Fiscal
                                                   '55=NF-e emitida em substituição ao modelo 1 ou 1A;
                                                   '65=NFC-e, utilizada nas operações de venda no varejo
                                                   '(a critério da UF aceitar este modelo de documento).

serie = Trim(txtSerie.Text)                        'Série do Documento Fiscal
                                                   'Série do Documento Fiscal, preencher com zeros na
                                                   'hipótese de a NF-e não possuir série. (v2.0)
                                                   'Série 890-899: uso exclusivo para emissão de NF-e
                                                   'avulsa, pelo contribuinte com seu certificado digital,
                                                   'através do site do Fisco (procEmi=2). (v2.0)
                                                   'Serie 900-999: uso exclusivo de NF-e emitidas no SCAN.(v2.0)

nNF = Trim(txtNOTA.Text)                           'Número do Documento Fiscal

dhEmi = Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) 'Data e hora de emissão do Documento
                                                                        'Data e hora no formato UTC (Universal Coordinated
                                                                        'Time): AAAA-MM-DDThh:mm:ssTZD

dhSaiEnt = Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2) & " " & Time   'Data e hora de Saída ou da Entrada da Mercadoria / Produto
                                                                                          'Data e hora no formato UTC (Universal Coordinated
                                                                                          'Time): AAAA-MM-DDThh:mm:ssTZD.
                                                                                          'Nota: Não informar este campo para a NFC-e.
If Trim(cmbTipoOperaAUX.Text) = "" Then _
    cmbTipoOperaAUX.Text = 1
tpNF = cmbTipoOperaAUX.Text   'Tipo de Operação
                              '0=Entrada;
                              '1=Saída

If Trim(txtUF_CLIENTE.Text) <> "" Then
   If Trim(UF_EMPRESA_A) = Trim(txtUF_CLIENTE.Text) Then
      cmbLocalAUX.Text = 1
      cmbLocal.Text = "Opereção Interna"
      Else
         cmbLocalAUX.Text = 2
         cmbLocal.Text = "Operação interestadual"
   End If
   Else: cmbLocalAUX.Text = 1
End If

idDest = cmbLocalAUX.Text     'Identificador de local de destino da operação
                              '1=Operação interna;
                              '2=Operação interestadual;
                              '3=Operação com exterior

cMunFG = Trim(CODG_IBGE_A)    'Código do Município de Ocorrência do Fato Gerador
                              'Informar o município de ocorrência do fato gerador do
                              'ICMS. Utilizar a Tabela do IBGE (Anexo IX - Tabela de
                              'UF, Município e País)

TpImp = 0                     'Formato de Impressão do DANFE
                              '0=Sem geração de DANFE;
                              '1=DANFE normal, Retrato;
                              '2=DANFE normal, Paisagem;
                              '3=DANFE Simplificado;
                              '4=DANFE NFC-e;
                              '5=DANFE NFC-e em mensagem eletrônica (o envio de
                              'mensagem eletrônica pode ser feita de forma simultânea
                              'com a impressão do DANFE; usar o tpImp=5 quando
                              'esta for a única forma de disponibilização do DANFE).

TpEmis = 1                    'Tipo de Emissão da NF-e
                              '1=Emissão normal (não em contingência);
                              '2=Contingência FS-IA, com impressão do DANFE em formulário de segurança;
                              '3=Contingência SCAN (Sistema de Contingência do Ambiente Nacional);
                              '4=Contingência DPEC (Declaração Prévia da Emissão em Contingência);
                              '5=Contingência FS-DA, com impressão do DANFE em formulário de segurança;
                              '6=Contingência SVC-AN (SEFAZ Virtual de Contingência do AN);
                              '7=Contingência SVC-RS (SEFAZ Virtual de Contingência do RS);
                              '9=Contingência off-line da NFC-e (as demais opções de contingência são válidas também para a NFC-e);
                              'Nota: Para a NFC-e somente estão disponíveis e são válidas as opções de contingência 5 e 9.

cDV = 1                       'Dígito Verificador da Chave de Acesso da NF - E
                              'Informar o DV da Chave de Acesso da NF-e, o DV será
                              'calculado com a aplicação do algoritmo módulo 11
                              '(base 2,9) da Chave de Acesso. (vide item 5 do Manual de Orientação)

TpAmb = cmbAmbienteAUX.Text   'Identificação do Ambiente
                              '1=Produção; 2=Homologação
'set aqui VI
FinNFe = cmbFinalidadeAUX.Text 'Finalidade de emissão da NF-e
                               '1=NF-e normal;
                               '2=NF-e complementar;
                               '3=NF-e de ajuste;
                               '4=Devolução/Retorno

indFinal = Trim(cmbCFinalAUX.Text)  'Indica operação com Consumidor final
                                    '0=Não;
                                    '1=Consumidor final;

indPres = Trim(cmbPresencaAUX.Text) 'Indicador de presença do comprador 'no estabelecimento comercial no momento da operação
                                    '0=Não se aplica (por exemplo, Nota Fiscal complementar ou de ajuste);
                                    '1=Operação presencial;
                                    '2=Operação não presencial, pela Internet;
                                    '3=Operação não presencial, Teleatendimento;
                                    '4=NFC-e em operação com entrega a domicílio;
                                    '9=Operação não presencial, outros

ProcEmi = 0                         'Processo de emissão da NF-e
                                    '0=Emissão de NF-e com aplicativo do contribuinte;
                                    '1=Emissão de NF-e avulsa pelo Fisco;
                                    '2=Emissão de NF-e avulsa, pelo contribuinte com seu certificado digital, através do site do Fisco;
                                    '3=Emissão NF-e pelo contribuinte com aplicativo fornecido pelo Fisco.

VerProc = "2.0.7"                   'Versão do Processo de emissão da NF -E
                                    'Informar a versão do aplicativo emissor de NF-e

dhCont = ""                         'Data e Hora da entrada em contingência
                                    'Justificativa da entrada em

xJust = ""                          'Justificativa da entrada em contingência

   LINHA_B = Trim(cUF) & "|"
   LINHA_B = LINHA_B & Trim(cNF) & "|"
   LINHA_B = LINHA_B & Trim(NatOp) & "|"
   LINHA_B = LINHA_B & Trim(indPag) & "|"
   LINHA_B = LINHA_B & Trim(mode) & "|"
   LINHA_B = LINHA_B & Trim(serie) & "|"
   LINHA_B = LINHA_B & Trim(nNF) & "|"
   LINHA_B = LINHA_B & Trim(dhEmi) & "|"
   LINHA_B = LINHA_B & Trim(dhSaiEnt) & "|"
   LINHA_B = LINHA_B & Trim(tpNF) & "|"            '0=Entrada;1=Saída
   LINHA_B = LINHA_B & Trim(idDest) & "|"          'Identificador de local de destino da operação
   LINHA_B = LINHA_B & Trim(cMunFG) & "|"
   LINHA_B = LINHA_B & Trim(TpImp) & "|"
   LINHA_B = LINHA_B & Trim(TpEmis) & "|"
   LINHA_B = LINHA_B & Trim(cDV) & "|"
   LINHA_B = LINHA_B & Trim(TpAmb) & "|"           'Identificação do Ambiente
   LINHA_B = LINHA_B & Trim(FinNFe) & "|"          'Finalidade de emissão da NF-e '1=NF-e normal; '4=Devolução/Retorno
   LINHA_B = LINHA_B & Trim(indFinal) & "|"
   LINHA_B = LINHA_B & Trim(indPres) & "|"
   LINHA_B = LINHA_B & Trim(ProcEmi) & "|"
   LINHA_B = LINHA_B & Trim(VerProc) & "|"
   LINHA_B = LINHA_B & Trim(dhCont) & "|"
   LINHA_B = LINHA_B & Trim(xJust) & "|"

   'If TIPO_NFe_GERAR = "R" Then

   Print #1, Tab(1); "B" & "|" & LINHA_B

'==================================== DEVOLUÇÃO
   If Trim(cmbFinalidadeAUX.Text) = 4 And Trim(txtNFeDev.Text) <> "" And Trim(txtChaveNFe.Text) <> "" Then
'Segmento BA02 layout = "§" +BA02 + |refNFe; //ok

'obs.: no caso de devolução voce tem que incluir agora esse segmento, a tag refNFe é a chave de referencia da nota de origem,
'manda somente uma chave mesmo que por exemplo o outro item se refere a outra chave caso de nota de devolução com duas notas de origem.

      Dim NFref, refNFe, refNF, AAMM, CNPJ, modelodocfisc, refNFP, cUF_refNFP, LINHA_DEVOLUCAO
      Dim AAMM_refNFP, CNPJ_refNFP, CPF_refNFP, IE_refNFP, modelodocfisc_refNFP
      Dim serie_refNFP, nNF_refNFP, refCTe_refNFP, refECF, mod_refECF, nECF_refECF, nCOO_refECF

      NFref = ""
      refNFe = ""
      refNF = ""
      cUF = ""
      AAMM = ""
      CNPJ = ""
      modelodocfisc = ""
      serie = ""
      nNF = ""

      'Informação de Documentos Fiscais referenciados
      NFref = Trim(txtChaveNFe.Text)            'Grupo com informações de Documentos Fiscais
                                                'referenciados. Informação utilizada nas hipóteses
                                                'previstas na legislação. (Ex.: Devolução de Mercadorias,
                                                'Substituição de NF cancelada, Complementação de NF,etc.).

      refNFe = Trim(txtNFeDev.Text)             'Chave de acesso da NF-e referenciada
                                                'Referencia uma NF-e (modelo 55) emitida anteriormente,
                                                'vinculada a NF-e atual, ou uma NFC-e (modelo 65),

      refNF = 55                                'Informação da NF modelo 1/1A referenciada

      cUF = txtIBGE.Text                        'Código da UF do emitente
                                                'Utilizar a Tabela do IBGE (Anexo IX - Tabela de UF,Município e País)

      AAMM = Mid(Date, 7, 4) & Mid(Date, 4, 2)  'Ano e Mês de emissão da NF-e

      CNPJ = Trim(CNPJ_EMPRESA_N)               'CNPJ do emitente

      modelodocfisc = "01"                      'Modelo do Documento Fiscal

      serie = 1                                 'Série do Documento Fiscal

      nNF = Trim(txtNOTA.Text)                  'Número do Documento Fiscal

      LINHA_DEVOLUCAO = Trim(NFref) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(refNFe) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(refNF) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(cUF) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(AAMM) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(CNPJ) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(modelodocfisc) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(serie) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(nNF) & "|"

'Informações da NF de produtor rural referenciada
      refNFP = ""
      cUF_refNFP = ""
      AAMM_refNFP = ""
      CNPJ_refNFP = ""
      CPF_refNFP = ""
      IE_refNFP = ""
      modelodocfisc_refNFP = ""
      serie_refNFP = ""
      nNF_refNFP = ""
      refCTe_refNFP = ""

'refNFP
'cUF_refNFP
'AAMM_refNFP
'CNPJ_refNFP
'CPF_refNFP
'IE_refNFP
'modelodocfisc_refNFP
'serie_refNFP
'nNF_refNFP
'refCTe_refNFP

      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(cUF_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(AAMM_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(CNPJ_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(CPF_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(IE_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(modelodocfisc_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(serie_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(nNF_refNFP) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(refCTe_refNFP) & "|"

'Informações do Cupom Fiscal referenciado
      refECF = ""
      mod_refECF = ""
      nECF_refECF = ""
      nCOO_refECF = ""

'refECF=
'mod_refECF=
'nECF_refECF=
'nCOO_refECF=

      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(refECF) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(mod_refECF) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(nECF_refECF) & "|"
      'LINHA_DEVOLUCAO = LINHA_DEVOLUCAO & Trim(nCOO_refECF) & "|"

      Print #1, Tab(1); "BA02" & "|" & LINHA_DEVOLUCAO
   End If
'====================================

   'encerra Registro B52
   'Começa registro B13 a Chave que depois eu vou gerar aqui
   'Print #1, Tab(1); "B|13|0000000000000000000000000000000000000000000"
   'Registro Codigo do estado ano e mes conforme yuri nao passar
   'Print #1, Tab(1); "B|14|35|" & Mid(Date, 7, 4) & Mid(Date, 4, 2);
   'Dados do Emitente
   'Colocar Campo Inscricao Municipaç empresa
   Print #1, Tab(1); "C|" & _
                     Trim(NOME_EMPRESA_A) & "|" & _
                     Trim("") & "|" & _
                     Trim(CCE_EMPRESA_N) & _
                     "||      ||" & _
                     Trim(CTR_EMPRESA_N) & _
                     "|"

   Print #1, Tab(1); "C02|" & Trim(CNPJ_EMPRESA_N)

   'Buscar Endereco
   'arrumae isso TABEND!Numero
   Print #1, Tab(1); "C05|" & _
                     Trim(tabEndereco!Rua) & "|" & _
                     "0" & "||" & _
                     Trim(tabEndereco!Bairro) & "|" & _
                     Trim(TabCEP!IBGE_ID) & "|" & _
                     Trim(TabCEP!CIDADE) & "|" & _
                     Trim(TabCEP!UF) & "|" & _
                     Trim(txtCep.Text) & _
                     "|1058|BRASIL|" & _
                     Trim(FONE_EMPRESA_N);

   'Dados do Destinatario
   If Left(TIPO_NFe_GERAR, 1) = "D" Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from vwFornecedor "
      SQL = SQL & " where pessoa_id = " & Trim(TabNOTA.Fields("pessoa_id").Value)
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then

         Print #1, Tab(1); "E|" & Left(Trim(TabCliente!DESCRICAO), 60) & "|" & UCase(Trim(cmbIE.Text)) & "|||";

         If Len(Trim(TabCliente.Fields("cnpjcpf").Value)) > 11 Then
            Print #1, Tab(1); "E02|" & Trim(TabCliente.Fields("cnpjcpf").Value);
            Else 'Nesta Caso e CNPJ
               Print #1, Tab(1); "E03|" & Trim(TabCliente.Fields("cnpjcpf").Value);
         End If

         BUSCA_ENDERECO_PESSOA "C", ""
         If Not tabEndereco.EOF Then
            Numero_A = "000"
            If Not IsNull(tabEndereco.Fields("numero").Value) Then _
               If Trim(tabEndereco.Fields("numero").Value) <> "" Then _
                  Numero_A = Trim(tabEndereco.Fields("numero").Value)

            SP_PROCURA_CEP tabEndereco!CEP_ID
            SP_PROCURA_FONE ""

            Print #1, Tab(1); "E05|" & tabEndereco!Rua & "|" & Numero_A & "| |" & tabEndereco!Bairro & "|" & TabCEP!IBGE_ID & "|" & TabCEP!CIDADE & "|" & TabCEP!UF & "|" & Trim(txtCep.Text) & "|1058|BRASIL|" & Right(TabFone!DDD, 2) & TabFone!Numero;
         End If
         If tabEndereco.State = 1 Then _
            tabEndereco.Close
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close
      Else  'buscando cliente
         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select CLIENTE.CLIENTE_ID, CLIENTE.PESSOA_ID, CLIENTE.estabelecimento_ID, "
         SQL = SQL & " CLIENTE.STATUS, PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO,"
         SQL = SQL & " codg_suframa "
         SQL = SQL & " from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
         SQL = SQL & " ON CLIENTE.PESSOA_ID = PESSOA.PESSOA_ID"

         SQL = SQL & " where PESSOA.pessoa_id = " & Trim(TabNOTA.Fields("pessoa_id").Value)
         SQL = SQL & " and status = 'A'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then
            CODG_SUFRAMA_A = "" & Trim(TabCliente.Fields("codg_suframa").Value)
            If Len(CODG_SUFRAMA_A) <= 3 Then _
               CODG_SUFRAMA_A = ""

            Print #1, Tab(1); "E|" & Left(Trim(TabCliente!DESCRICAO), 60) & "|" & "|" & UCase(Trim(cmbIE.Text)) & "|" & "|" & Trim(CODG_SUFRAMA_A) & "|" & Trim(cmbEmail.Text);

            If Len(Trim(TabCliente.Fields("cnpjcpf").Value)) > 11 Then
               Print #1, Tab(1); "E02|" & Trim(TabCliente.Fields("cnpjcpf").Value);
               Else 'Nesta Caso e CNPJ
                  Print #1, Tab(1); "E03|" & Trim(TabCliente.Fields("cnpjcpf").Value);
            End If

            BUSCA_ENDERECO_PESSOA "C", ""
            If Not tabEndereco.EOF Then
               Numero_A = "000"
               If Not IsNull(tabEndereco.Fields("numero").Value) Then _
                  If Trim(tabEndereco.Fields("numero").Value) <> "" Then _
                     Numero_A = Trim(tabEndereco.Fields("numero").Value)

               SP_PROCURA_CEP tabEndereco!CEP_ID

               If TabCEP.EOF Then _
                  MsgBox "NÃO ENCONTROU CEP"

               SP_PROCURA_FONE ""
               If TabFone.EOF Then
                  MsgBox "NÃO ENCONTROU TABFONE"
                  Else: DDD_N = 0 & TabFone!DDD
               End If

               Print #1, Tab(1); _
               "E05|" & _
               Trim(txtEnd.Text) & "|" & _
               Numero_A & "| |" & _
               Trim(txtBairro.Text) & "|" & _
               Trim(txtIBGE.Text) & "|" & _
               Trim(txtCidade.Text) & "|" & _
               Trim(txtUF_CLIENTE.Text) & "|" & _
               Trim(txtCep.Text) & _
               "|1058|BRASIL|" & _
               Right(DDD_N, 2) & _
               Trim(Left(txtFone.Text, 10));
            End If
            If tabEndereco.State = 1 Then _
               tabEndereco.Close
         End If
         If TabCliente.State = 1 Then _
            TabCliente.Close
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   'Dados Local da retirada
   SQL = "select * from ENDERECO "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and tipo = 'C'"
   tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabEndereco.EOF Then
      SP_PROCURA_CEP tabEndereco!CEP_ID

      Numero_A = "000"
      If Not IsNull(tabEndereco.Fields("numero").Value) Then _
         If Trim(tabEndereco.Fields("numero").Value) <> "" Then _
            Numero_A = Trim(tabEndereco.Fields("numero").Value)

      'Print #1, Tab(1); "F|" & TABEND!Rua & "|" & "000|  " & "|" & TABEND!Bairro & "|" & TABCEP!IBGE_ID & "|" & TABCEP!Cidade & "|" & TABCEP!UF & "|";
      Print #1, Tab(1); "F|" & tabEndereco!Rua & "|" & Numero_A & "|  " & "|" & tabEndereco!Bairro & "|" & TabCEP!IBGE_ID & "|" & TabCEP!CIDADE & "|" & TabCEP!UF & "|";
      Print #1, Tab(1); "F02|" & Trim(CNPJ_EMPRESA_N);
      'Print #1, Tab(1); "G|" & TABEND!Rua & "|" & "000|  " & "|" & TABEND!Bairro & "|" & TABCEP!IBGE_ID & "|" & TABCEP!Cidade & "|" & TABCEP!UF & "|";
      Print #1, Tab(1); "G|" & tabEndereco!Rua & "|" & Numero_A & "|  " & "|" & tabEndereco!Bairro & "|" & TabCEP!IBGE_ID & "|" & TabCEP!CIDADE & "|" & TabCEP!UF & "|";
      Print #1, Tab(1); "G02|" & Trim(CNPJ_EMPRESA_N);
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_CABEÇALHO_NFe"
End Sub

Private Sub GERAR_PRODUTOS_NFe()
'On Error GoTo ERRO_TRATA

   Dim Linhas_Impressas_N  As Long
   Dim strCFOP_ITEM        As String
   Dim TabItemNota         As New ADODB.Recordset
   Dim intTributacao       As Integer
   Dim cProd, cEAN, xProd, NCM, NVE, EXTIPI, CFOP, uCom, qCom, vUnCom, vProd, cEANTrib, uTrib, qTrib, vUnTrib, vFrete, vSeg, vDesc, vOutro, indTot

   NUMR_SEQ_N = 0

   If TabItemNota.State = 1 Then _
      TabItemNota.Close

   SQL = "select NF.PEDIDO_ID, NF.NF_TIPO, NFITEM.*, "
   SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.UNIDADE_MEDIDA, PRODUTO.CODG_NCM, PRODUTO.TIPO_PROD, "
   SQL = SQL & " Produto.Situacao_Tributaria, produto.codg_produto, produto.origem_mercado "
   SQL = SQL & " from NF "
   SQL = SQL & " INNER JOIN NFITEM "
   SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where NF.nf_id = " & TabNOTA!NF_ID
   TabItemNota.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabItemNota.EOF Then
      TabItemNota.MoveFirst
      While Not TabItemNota.EOF
'=====zerando variaveis
         cProd = ""
         cEAN = "   "
         xProd = ""
         NCM = ""
         NVE = ""
         EXTIPI = ""
         CFOP = ""
         uCom = ""
         qCom = ""
         vUnCom = ""
         vProd = ""
         cEANTrib = " "
         uTrib = ""
         qTrib = ""
         vUnTrib = ""
         vFrete = " "
         vSeg = " "
         vDesc = " "
         vOutro = ""
         indTot = ""
         '===================================
         cProd = Trim(TabItemNota!Codg_Produto)
         cEAN = "   "
         xProd = Trim(TabItemNota!DESCRICAO)
         NCM = Left(TabItemNota!CODG_NCM, 8)
         NVE = ""
         EXTIPI = ""
         CFOP = Trim(TabItemNota!CFOP_ID)
         uCom = Trim(TabItemNota!UNIDADE_MEDIDA)
         qCom = Format$(TabItemNota!QTDE, strFormatacao4Digitos)
         vUnCom = Format(TabItemNota!VALOR, strFormatacao2Digitos)
         vProd = Format(TabItemNota!QTDE * TabItemNota!VALOR, strFormatacao2Digitos)
         cEANTrib = " "
         uTrib = Trim(TabItemNota!UNIDADE_MEDIDA)
         qTrib = Format$(TabItemNota!QTDE, strFormatacao4Digitos)
         vUnTrib = Format(TabItemNota!VALOR, strFormatacao2Digitos)
         vFrete = " " 'Format(0, strFormatacao2Digitos)
         vSeg = " "   'Format(0, strFormatacao2Digitos)
         vDesc = " "  'Format(0, strFormatacao2Digitos)
         vOutro = "" 'Format(0, strFormatacao2Digitos)
         indTot = 1
'==============================
         strCFOP_ITEM = "" & Trim(TabItemNota!CFOP_ID)
         cmbCFOPAux.Text = "" & Trim(TabItemNota!CFOP_ID)

         'quando é cupom com nfe
         '5929-Lançamento efetuado em decorrência de emissão de documento fiscal relativo a operação ou prestação também registrada em equipamento Emissor de Cupom Fiscal - ECF
         '6929-Lançamento efetuado em decorrência de emissão de documento fiscal relativo a operação ou prestação também registrada em equipamento Emissor de Cupom Fiscal - ECF
         If Trim(cmbCFOPAux.Text) = 5929 Or Trim(cmbCFOPAux.Text) = 6929 Then

            SQL = "update NFITEM set CFOP_id = " & Trim(cmbCFOPAux.Text)
            SQL = SQL & " where nf_id = " & TabItemNota.Fields("nf_id").Value
            CONECTA_RETAGUARDA.Execute SQL

            strCFOP_ITEM = "" & Trim(cmbCFOPAux.Text)
            Else
               'quando é substituição tributária
'10 Tributada  e com cobrança do ICMS por substituição tributária
'60 ICMS cobrado anteriormente por substituição tributária
'70 Com redução de base de cálculo e cobrança de ICMS por substituição tributária
'set
               If Trim(TabItemNota!SITUACAO_TRIBUTARIA) = 10 Or _
                  Trim(TabItemNota!SITUACAO_TRIBUTARIA) = 60 Or _
                  Trim(TabItemNota!SITUACAO_TRIBUTARIA) = 70 Then
'ajustado aqui para quando for devolução de venda pegar do combo o cfop
                  If Trim(TIPO_NFe_GERAR) <> "DC" Then
                     If Trim(TIPO_NFe_GERAR) <> "DV" Then
                        If Trim(UF_EMPRESA_A) = Trim(txtUF_CLIENTE.Text) Then

                           If INDR_INDUSTRIA_B = False Then
                              '5405  Venda de mercadoria, adquirida ou recebida de terceiros,
                              'sujeita ao regime de substituição tributária,
                              'na condição de contribuinte-substituído
                              strCFOP_ITEM = "5405"   'não é industria
                              Else
                                 '
                                 strCFOP_ITEM = "5403"
                           End If
                           Else
                              If INDR_INDUSTRIA_B = False Then
                                 strCFOP_ITEM = "6404"         'não é industria
                                 Else: strCFOP_ITEM = "6404"   'é industria
                              End If
                        End If
                     End If
                  End If
                  Else: strCFOP_ITEM = "" & Trim(cmbCFOPAux.Text) 'segue o cfop informado do combo da tela
               End If
         End If
'=====================================

         frmINTEGRA.INTEGRA_CFOP Int(strCFOP_ITEM)

         NUMR_SEQ_N = NUMR_SEQ_N + 1

         Print #1, Tab(1); "H|" & NUMR_SEQ_N & "|" & TabItemNota!DESCRICAO;

cProd = Trim(TabItemNota!Codg_Produto)
cEAN = "   "
xProd = Trim(TabItemNota!DESCRICAO)
NCM = Left(TabItemNota!CODG_NCM, 8)
NVE = ""
EXTIPI = ""
CFOP = Trim(TabItemNota!CFOP_ID)
uCom = Trim(TabItemNota!UNIDADE_MEDIDA)
qCom = Format$(TabItemNota!QTDE, strFormatacao4Digitos)
vUnCom = Format(TabItemNota!VALOR, strFormatacao2Digitos)
vProd = Format(TabItemNota!QTDE * TabItemNota!VALOR, strFormatacao2Digitos)
cEANTrib = " "
uTrib = Trim(TabItemNota!UNIDADE_MEDIDA)
qTrib = Format$(TabItemNota!QTDE, strFormatacao4Digitos)
vUnTrib = Format(TabItemNota!VALOR, strFormatacao2Digitos)
vFrete = " " 'Format(0, strFormatacao2Digitos)
vSeg = " "   'Format(0, strFormatacao2Digitos)
vDesc = " "  'Format(0, strFormatacao2Digitos)
vOutro = "" 'Format(0, strFormatacao2Digitos)
indTot = 1

SQL = "|" & cProd
SQL = SQL & "|" & cEAN
SQL = SQL & "|" & xProd
SQL = SQL & "|" & NCM
SQL = SQL & "|" & NVE
SQL = SQL & "|" & EXTIPI
SQL = SQL & "|" & CFOP
SQL = SQL & "|" & uCom
SQL = SQL & "|" & qCom
SQL = SQL & "|" & vUnCom
SQL = SQL & "|" & vProd
SQL = SQL & "|" & cEANTrib
SQL = SQL & "|" & uTrib
SQL = SQL & "|" & qTrib
SQL = SQL & "|" & vUnTrib
SQL = SQL & "|" & vFrete
SQL = SQL & "|" & vSeg
SQL = SQL & "|" & vDesc
SQL = SQL & "|" & vOutro
SQL = SQL & "|" & indTot
SQL = SQL & "||"

Print #1, Tab(1); "I" & SQL

'verificar aqui
'Print #1, Tab(1); "I|" & Trim(TabItemNota!CODG_PRODUTO) & "|   |" & Trim(TabItemNota!Descricao) & "|" & Trim(Left(TabItemNota!codg_ncm, 8)) & "||" & "||" & strCFOP_ITEM & "|" & TabItemNota!Unidade_Medida & "|" & Format$(TabItemNota!Qtde, strFormatacao4Digitos) & "|" & Format(TabItemNota!Valor, strFormatacao2Digitos) & "|" & Format(TabItemNota!Qtde * TabItemNota!Valor, strFormatacao2Digitos) & "| |" & TabItemNota!Unidade_Medida & "|" & Format$(TabItemNota!Qtde, strFormatacao4Digitos) & "|" & Format(TabItemNota!Valor, strFormatacao2Digitos) & "| | | ||||";  '" & Format$(TABITEMNOTA!DESCONTO, strFormatacao2Digitos);
'===================================================
         'Parte de impostos dos produtos
         Print #1, Tab(1); "M";
         Print #1, Tab(1); "N";

'If CTR_EMPRESA_N = 1 Then

intTributacao = "400"
'Else


'intTributacao = BUSCA_TRIBUTACAO_PRODUTO(TabItemNota.Fields("origem_mercado").Value, TabItemNota.Fields("SITUACAO_TRIBUTARIA").Value)
'End If

'If TabItemNota.Fields("SITUACAO_TRIBUTARIA").Value = "00" Then
'   intTributacao=
'   Else
'End If
'verificar quando for tributado


         Print #1, Tab(1); "N10d|0|" & intTributacao & "|";

         'Finaliza Itens PIS
         Print #1, Tab(1); "Q";
         Print #1, Tab(1); "Q05|99|0,00";
         Print #1, Tab(1); "Q07|0,00|0,00";

         'Finaliza Itens Cofins
         Print #1, Tab(1); "S";
         Print #1, Tab(1); "S05|99|0,00";
         Print #1, Tab(1); "S07|0,00|0,00";

'=============baixa estoque INICIO
         If TabItemNota.Fields("nf_tipo").Value = "S" Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select status from PEDIDOITEM "
            SQL = SQL & " where pedido_id = " & TabItemNota.Fields("pedido_id").Value
            SQL = SQL & " and produto_id = " & TabItemNota.Fields("produto_id").Value
            SQL = SQL & " and status not in ('B','C') "
            SQL = SQL & " and tipo_reg = 'PC' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               PRODUTO_ID_N = TabItemNota.Fields("produto_id").Value
               QTDE_PEDIDO = TabItemNota.Fields("qtde").Value

               SQL = "UPDATE PEDIDOITEM set "
               SQL = SQL & " status = 'B' "
               SQL = SQL & " where pedido_id = " & TabItemNota.Fields("pedido_id").Value
               SQL = SQL & " and produto_id = " & TabItemNota.Fields("produto_id").Value
               SQL = SQL & " and status <> 'B' "
               CONECTA_RETAGUARDA.Execute SQL
            End If
         End If
'=============baixa estoque FIM

         TabItemNota.MoveNext
      Wend
   End If
   If TabItemNota.State = 1 Then _
      TabItemNota.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_PRODUTOS_NFe"
End Sub

Private Sub GERAR_TOTAIS_NFe()
'On Error GoTo ERRO_TRATA

   Vlr_BaseCalculo_N = 0 & txtBaseCalculo.Text
   Vlr_TotICMS_N = 0 & txtValorICMS.Text
   Vlr_BaseICMSub_N = 0 & txtBaseIcmsSub.Text
   Vlr_ICMSub_N = 0 & txtVlrIcmsSub.Text
   Vlr_TotProdutos_N = 0 & txtValorProdutos.Text
   VLR_FRETE_N = 0 & txtFrete.Text
   Vlr_Desconto_N = 0 & txtDesconto.Text
   Vlr_TotIPI_N = 0 & txtValorIPI.Text
   VLR_OUTROS_N = 0 & txtValorOutros.Text
   Vlr_TotNFe_N = 0 & txtValorTotalNota.Text

'============================================================
'(+) vProd (id:W07)                 'vem da rotina total nota (pedido/devolução)
'(-) vDesc (id:W10)                 'vem da rotina total nota (pedido/devolução)
'(+) vICMSST (id:W06)
'(+) vFrete (id:W09)                'informado no campo
'(+) vSeg (id:W10)
'(+) vOutro (id:W15)                'informado no campo
'(+) vII (id:W11)
'(+) vIPI (id:W12)                  'informado no campo
'(+) vServ (id:W19) (NT 2011/004)

   Vlr_TotNFe_N = Vlr_TotProdutos_N + VLR_FRETE_N + VLR_OUTROS_N + Vlr_TotIPI_N + Vlr_ICMSub_N - Vlr_Desconto_N

   Print #1, Tab(1); "W";

'============================================
'segmento W02
'novo layout : "§W02|vBC|vICMS|vICMSDeson|vBCST|vST|vProd|vFrete|vSeg|vDesc|vII|vIPI|vPIS|vCOFINS|vOutro|vNF|vTotTrib"
'obs.: obs.: eles acrrescentaram essa tag  vICMSDeson manda 0 assim acrescentando isso |0|, conforme layout
'============================================

   SQL = "|" & Format(Vlr_BaseCalculo_N, strFormatacao2Digitos)         'vBC - Base de Cálculo do ICMS
   SQL = SQL & "|" & Format(Vlr_TotICMS_N, strFormatacao2Digitos)       'vICMS - Valor Total do ICMS
   
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vICMSDeson - Valor Total do ICMS desonerado
   
   SQL = SQL & "|" & Format(Vlr_BaseICMSub_N, strFormatacao2Digitos)    'vBCST - Base de Cálculo do ICMS ST
   SQL = SQL & "|" & Format(Vlr_ICMSub_N, strFormatacao2Digitos)        'vST - Valor Total do ICMS ST
   SQL = SQL & "|" & Format(Vlr_TotProdutos_N, strFormatacao2Digitos)   'vProd - Valor Total dos produtos e serviços
   SQL = SQL & "|" & Format(VLR_FRETE_N, strFormatacao2Digitos)         'vFrete - Valor Total do Frete
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vSeg - Valor Total do Seguro
   SQL = SQL & "|" & Format(Vlr_Desconto_N, strFormatacao2Digitos)      'vDesc - Valor Total do Desconto
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vII - Valor Total do II
   SQL = SQL & "|" & Format(Vlr_TotIPI_N, strFormatacao2Digitos)        'vIPI - Valor Total do IPI
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vPIS - Valor do PIS
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vCOFINS - Valor do COFINS
   SQL = SQL & "|" & Format(VLR_OUTROS_N, strFormatacao2Digitos)        'vOutro - Outras Despesas acessórias
   SQL = SQL & "|" & Format(Vlr_TotNFe_N, strFormatacao2Digitos)        'vNF - Valor Total da NF-e
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'ISSQNtot - Grupo de Valores Totais referentes ao ISSQN

   Print #1, Tab(1); "W02" & SQL;
'================================
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_TOTAIS_NFe"
End Sub

Private Sub GERAR_TRANSPORTADORA_NFe()
'On Error GoTo ERRO_TRATA

   Dim Quantidade_n              As Double
   Dim PESO_BRUTO_N              As Double
   Dim PESO_LIQUIDO_N            As Double

   Quantidade_n = 0 & TxtQuantidadeRodapeNota.Text
   PESO_BRUTO_N = 0 & TxtPesoBruto.Text
   PESO_LIQUIDO_N = 0 & TxtPesoLiquido.Text

   If Trim(txtCNPJCPF_TRANSP.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwTRANSPORTADORA "
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("transp_id").Value) Then
            BUSCA_ENDERECO_PESSOA "C", ""

            Print #1, Tab(1); "X|1";
            If Not tabEndereco.EOF Then
               Print #1, Tab(1); "X03|" & Trim(TabTemp!DESCRICAO) & "|" & TabTemp!numr_IE & "|" & tabEndereco!Rua & "|" & tabEndereco!UF & "|" & tabEndereco!CIDADE;
               Else
               Print #1, Tab(1); "X03|" & Trim(TabTemp!DESCRICAO) & "|" & "ISENTO" & "|" & "GOIANIA" & "|" & "GO" & "|" & "GOIANIA";
            End If
            Print #1, Tab(1); "X04|" & Trim(TabTemp!CNPJCPF);

SQL = "|" & Quantidade_n                                          'Quantidade de volumes transportados
SQL = SQL & "|" & Trim(TxtEspecie.Text)                           'Espécie dos volumes transportados
SQL = SQL & "|" & Trim(cmbMarca.Text)                             'Marca dos volumes transportados
SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                'Numeração dos volumes transportados
SQL = SQL & "|" & Format(PESO_LIQUIDO_N, strFormatacao3Digitos)   'Peso Líquido (em kg)
SQL = SQL & "|" & Format(PESO_LIQUIDO_N, strFormatacao3Digitos)   'Peso Bruto (em kg)

Print #1, Tab(1); "X26" & SQL
SQL = ""
         End If
         Else
            MsgBox "Transportadora inexistente, impossível continuar !!!"
            Unload Me
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      If tabEndereco.State = 1 Then _
         tabEndereco.Close
      If TabCEP.State = 1 Then _
         TabCEP.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_TRANSPORTADORA_NFe"
End Sub

Private Sub GERAR_FATURAS_NFe()
'On Error GoTo ERRO_TRATA

   Print #1, Tab(1); "Y";
   'Print #1, Tab(1); "Y02|" & PEDIDO_ID_N & "|" & Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos) & "|" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos) & "|" & Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
   Print #1, Tab(1); "Y02|" & PEDIDO_ID_N & "|" & Format(txtValorTotalNota.Text, strFormatacao2Digitos) & "|" & Format(VALOR_DESCONTO_N, strFormatacao2Digitos) & "|" & Format(txtValorTotalNota.Text, strFormatacao2Digitos)

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from ITEMLANCAMENTO i, LANCAMENTO l "
   SQL = SQL & " where l.numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and l.lancamento_id = i.lancamento_id "
   SQL = SQL & " and l.tipo_lancamento = 1 "
   SQL = SQL & " and formapagto_id > 1 "
   SQL = SQL & " order by i.seq"
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabLancamento.EOF
      Print #1, Tab(1); "Y07|" & txtNOTA.Text & "|" & Mid(TabLancamento!DT_VENCIMENTO, 7, 4) & "-" & Mid(TabLancamento!DT_VENCIMENTO, 4, 2) & "-" & Mid(TabLancamento!DT_VENCIMENTO, 1, 2) & "|" & Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
      TabLancamento.MoveNext
   Wend
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_FATURAS_NFe"
End Sub

Private Sub GERAR_RODAPE_NFe()
'On Error GoTo ERRO_TRATA

   Dim TabTipovenda        As New ADODB.Recordset
   Dim Descricao_Pgto      As String

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select descricao from PEDIDO "
   SQL = SQL & " INNER JOIN vwVendedor "
   SQL = SQL & " ON PEDIDO.VENDEDOR_ID = vwVendedor.VENDEDOR_ID"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      NOME_VEND_A = "" & Trim(TabConsulta.Fields(0).Value)
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'Condicoes de Pgto
   SQL = "select tipovenda_id from LANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
   SQL = SQL & " and tipo_lancamento = " & 1 'RECEBER
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      If Not IsNull(TabConsulta!TIPOVENDA_ID) Then
         If IsNumeric(TabConsulta!TIPOVENDA_ID) Then
            If TabTipovenda.State = 1 Then _
               TabTipovenda.Close

            SqL2 = "select descricao from TIPOVENDA WITH (NOLOCK)"
            SqL2 = SqL2 & " where tipovenda_id = " & TabConsulta!TIPOVENDA_ID
            TabTipovenda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTipovenda.EOF Then _
               If Not IsNull(TabTipovenda.Fields(0).Value) Then _
                  Descricao_Pgto = TabTipovenda.Fields(0).Value
            If TabTipovenda.State = 1 Then _
               TabTipovenda.Close
         End If
      End If
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Msg = ""
   'Print #1, Tab(1); "Z|" & Trim(txtMSG.Text) & ", Numero Pedido: " & PEDIDO_ID_N & ", Vendedor: " & Trim(NOME_VEND_A) & ", " & Trim(txtDadosAdicionais.Text);
   Print #1, Tab(1); "Z|" & Trim(txtMSG.Text) & Msg & ", " & Trim(txtDadosAdicionais.Text);

   If TIPO_NFe_GERAR = "DC" Then
      Print #1, Tab(1); "Z04|" & "DV ENT REF. " & Trim(txtNFeDev.Text) & "|0";
      Else
         If TIPO_NFe_GERAR = "DV" Then
            Print #1, Tab(1); "Z04|" & "DV SAI REF. " & 0 & "|0";
            Else: Print #1, Tab(1); "Z04|" & "VOLTE SEMPRE" & "|0";
         End If
   End If

   Print #1, Tab(1); "Z10|" & "PROCESSO INTERNO" & "|1";
   Msg = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_RODAPE_NFe"
End Sub

Sub TRIBUTOS_LEI12741()
'On Error GoTo ERRO_TRATA

   Dim TabIBPTax           As New ADODB.Recordset
   Dim ALIQ_IBPT_N         As Double
   Dim VALOR_TOTAL_IMPOSTO As Double
   Dim VALOR_TOTAL_N       As Double

'===========================================lei 12.741
ALIQ_IBPT_N = 0
VALOR_TOTAL_IMPOSTO = 0
VALOR_TOTAL_N = 0 & txtValorTotalNota.Text
'==========================
'CALCULO IMPOSTO LEI 12.741 (BUSCA CODIGO NCM DO CADASTRO DO PRODUTO,
'LÊ TABELA 'IBPTax' QUE CONTEM A ALIQUOTA RELACIONADA AO NCM DO PRODUTO

   If Trim(TIPO_NFe_GERAR) = "R" Then
      If INDR_LEI_12741 = True Then
         If TabIBPTax.State = 1 Then _
            TabIBPTax.Close

         'chegando itens
         SQL = "select PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.STRIBUTARIA, PRODUTO.CODG_NCM, PRODUTO.ORIGEM_MERCADO,"
         SQL = SQL & " PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.Valor_Item"
         SQL = SQL & " from PEDIDO "
         SQL = SQL & " INNER JOIN PEDIDOITEM "
         SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
         SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
         SQL = SQL & " INNER JOIN PRODUTO "
         SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
         SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

         SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and pedidoitem.status <> 'C' "

         TabIBPTax.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabIBPTax.EOF
            If Not IsNull(TabIBPTax.Fields("codg_ncm").Value) Then
               If Trim(TabIBPTax.Fields("codg_ncm").Value) <> "" Then
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
                  If Not IsNull(TabIBPTax.Fields("origem_mercado").Value) Then
                     If Trim(TabIBPTax.Fields("origem_mercado").Value) <> "" Then
                        SQL = "select ALIQNAC,ALIQIMP from IBPTax "
                        SQL = SQL & " where codg_ncm = '" & Trim(TabIBPTax.Fields("codg_ncm").Value) & "'"
                        SQL = SQL & " and tabela = " & TabIBPTax.Fields("origem_mercado").Value
                        TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If Not TabTemp.EOF Then
                           ALIQ_IBPT_N = 0

                           If TabIBPTax.Fields("origem_mercado").Value = 0 Then _
                              ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqnac").Value
                           If TabIBPTax.Fields("origem_mercado").Value = 1 Then _
                              ALIQ_IBPT_N = 0 & TabTemp.Fields("aliqimp").Value

                           VALOR_ITEM_N = 0 & TabIBPTax.Fields("valor_item").Value
                           QTDE_N = 0 & TabIBPTax.Fields("QTD_PEDIDA").Value

                           VALOR_TOTAL_IMPOSTO = VALOR_TOTAL_IMPOSTO + ((VALOR_ITEM_N * QTDE_N) * ALIQ_IBPT_N / 100)
                        End If
                     End If
                  End If
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
               End If
            End If

            TabIBPTax.MoveNext
         Wend
         If TabIBPTax.State = 1 Then _
            TabIBPTax.Close

         If (VALOR_TOTAL_IMPOSTO / VALOR_TOTAL_N) > 0 Then
            SQL = "Val Aprox Tributos R$ " & Format(VALOR_TOTAL_IMPOSTO, strFormatacao2Digitos) & _
                  "(" & Format((VALOR_TOTAL_IMPOSTO / VALOR_TOTAL_N), strFormatacao2Digitos) & "%)"
            txtDadosAdicionais.Text = Trim(txtDadosAdicionais.Text) & " ; " & SQL
         End If
      End If
   End If
'=====================================

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRIBUTOS_LEI12741"
End Sub

Sub ATUALIZA_GLOBAL_MFA010()
'On Error GoTo ERRO_TRATA

   Dim VLR_FRETE_N   As Double
   Dim PESO_BRUTO_N  As Double
   Dim PESO_LIQUI_N  As Double
   Dim Quantidade_n  As Double

   VLR_FRETE_N = 0 & txtFrete.Text
   PESO_BRUTO_N = 0 & TxtPesoBruto.Text
   PESO_LIQUI_N = 0 & TxtPesoLiquido.Text
   Quantidade_n = 0 & TxtQuantidadeRodapeNota.Text

   ABRE_BANCO_GLOBAL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MFADOC from MFA010"
   SQL = SQL & " where mfadoc = '" & Trim(txtNOTA.Text) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update MFA010 set "
      SQL = SQL & " MFAFRETE = " & tpMOEDA(VLR_FRETE_N)
      SQL = SQL & ", MFAPBRUTO = " & tpMOEDA(PESO_BRUTO_N)
      SQL = SQL & ", MFAPLIQUI  = " & tpMOEDA(PESO_LIQUI_N)
      SQL = SQL & ", MFAESPECIE = '" & Trim(TxtEspecie.Text) & "'"
      SQL = SQL & ", MFAVOLUME4 = " & tpMOEDA(Quantidade_n)
      SQL = SQL & ", MFATIFRETE = '" & Trim(cmbFreteAUX.Text) & "'"

      SQL = SQL & " where mfadoc = '" & Trim(txtNOTA.Text) & "'"
      CONECTA_GLOBAL.Execute SQL
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_GLOBAL_MFA010"
End Sub

Private Sub GERAR_TOTAIS_NFe_VEIA()
'On Error GoTo ERRO_TRATA

   Dim Vlr_BaseCalculo_N   As Double
   Dim Vlr_TotICMS_N       As Double
   Dim Vlr_BaseICMSub_N    As Double
   Dim Vlr_ICMSub_N        As Double
   Dim Vlr_TotProdutos_N   As Double
   Dim VLR_FRETE_N         As Double
   Dim Vlr_Desconto_N      As Double
   Dim Vlr_TotIPI_N        As Double
   Dim VLR_OUTROS_N        As Double
   Dim Vlr_TotNFe_N        As Double

   Vlr_BaseCalculo_N = 0 & txtBaseCalculo.Text
   Vlr_TotICMS_N = 0 & txtValorICMS.Text
   Vlr_BaseICMSub_N = 0 & txtBaseIcmsSub.Text
   Vlr_ICMSub_N = 0 & txtVlrIcmsSub.Text
   Vlr_TotProdutos_N = 0 & txtValorProdutos.Text
   VLR_FRETE_N = 0 & txtFrete.Text
   Vlr_Desconto_N = 0 & txtDesconto.Text
   Vlr_TotIPI_N = 0 & txtValorIPI.Text
   VLR_OUTROS_N = 0 & txtValorOutros.Text
   Vlr_TotNFe_N = 0 & txtValorTotalNota.Text

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum((valor*qtde)) from NFITEM "
   SQL = SQL & " where nf_id = " & TabNOTA!NF_ID
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_TOTAL_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_TOTAL_N = VALOR_TOTAL_N - Vlr_Desconto_N

   Print #1, Tab(1); "W";

   SQL = "|" & Format(Vlr_BaseCalculo_N, strFormatacao2Digitos)         'vBC - Base de Cálculo do ICMS
   SQL = SQL & "|" & Format(Vlr_TotICMS_N, strFormatacao2Digitos)       'vICMS - Valor Total do ICMS
   SQL = SQL & "|" & Format(Vlr_BaseICMSub_N, strFormatacao2Digitos)    'vBCST - Base de Cálculo do ICMS ST
   SQL = SQL & "|" & Format(Vlr_ICMSub_N, strFormatacao2Digitos)        'vST - Valor Total do ICMS ST
   SQL = SQL & "|" & Format(Vlr_TotProdutos_N, strFormatacao2Digitos)   'vProd - Valor Total dos produtos e serviços
   SQL = SQL & "|" & Format(VLR_FRETE_N, strFormatacao2Digitos)         'vFrete - Valor Total do Frete
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vSeg - Valor Total do Seguro
   SQL = SQL & "|" & Format(Vlr_Desconto_N, strFormatacao2Digitos)      'vDesc - Valor Total do Desconto
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vII - Valor Total do II
   SQL = SQL & "|" & Format(Vlr_TotIPI_N, strFormatacao2Digitos)        'vIPI - Valor Total do IPI
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vPIS - Valor do PIS
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'vCOFINS - Valor do COFINS
   SQL = SQL & "|" & Format(VLR_OUTROS_N, strFormatacao2Digitos)        'vOutro - Outras Despesas acessórias
   SQL = SQL & "|" & Format(Vlr_TotNFe_N, strFormatacao2Digitos)        'vNF - Valor Total da NF-e
   SQL = SQL & "|" & Format(0, strFormatacao2Digitos)                   'ISSQNtot - Grupo de Valores Totais referentes ao ISSQN

   Print #1, Tab(1); "W02" & SQL;
'================================
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_TOTAIS_NFe"
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
            cmbCFOP.Visible = True

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
      Case 5
         texto = txtValorDig.Text
         MSFlexGrid1.TextMatrix(LastRow, LastCol) = Trim(texto)
         MSFlexGrid1.CellForeColor = vbRed
         MSFlexGrid1.CellFontBold = True
         MSFlexGrid1.CellBackColor = &H8000000F
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   'Ocultar o controle textbox
   txtValorDig.Visible = False
   cmbCFOP.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Sub GERAR_NFE_old()
'On Error GoTo ERRO_TRATA

   Dim VALOR_01      As Double
   Dim VALOR_02      As Double
   Dim strTributacao As String

   If Trim(cmbPresencaAUX.Text) = "" Then
      MsgBox "Selecione (Indicador de presença do comprador)."
      Exit Sub
   End If
   If Trim(cmbLocalAUX.Text) = "" Then
      MsgBox "Selecione (Identificador de local de destino da operação)."
      Exit Sub
   End If

   'validar transportadora
   If Trim(cmbCNPJCPF_TRANSP.Text) = "" Then
      MsgBox "Informar trasportadora."
      txtCNPJCPF_TRANSP.SetFocus
      Exit Sub
   End If

   txtCNPJCPF_TRANSP.PromptInclude = False
   TRANSP_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PESSOA.CNPJCPF, PESSOA.DESCRICAO, TRANSPORTADORA.PESSOA_ID, transp_id"
   SQL = SQL & " from TRANSPORTADORA WITH (NOLOCK)"
   SQL = SQL & " Inner Join PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON TRANSPORTADORA.PESSOA_ID = PESSOA.PESSOA_ID"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF_TRANSP.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      'MsgBox "Informar trasportadora."
      'txtCNPJCPF_TRANSP.SetFocus
      'Exit Sub
      Else: TRANSP_ID_N = 0 & TabTemp.Fields("transp_id").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   'validar cliente
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      MsgBox "CNPJ/CPF inválido !!!"
      Exit Sub
   End If
   If Trim(txtEnd.Text) = "" Then
      MsgBox "Endereço inválido !!!"
      txtEnd.SetFocus
      Exit Sub
   End If
   If Trim(txtBairro.Text) = "" Then
      MsgBox "Bairro inválido !!!"
      txtBairro.SetFocus
      Exit Sub
   End If
   If Trim(txtUF_CLIENTE.Text) = "" Then
      MsgBox "UF inválido !!!"
      txtUF_CLIENTE.SetFocus
      Exit Sub
   End If
   If Trim(txtCep.Text) = "" Then
      MsgBox "CEP inválido !!!"
      txtCep.SetFocus
      Exit Sub
   End If
   If Trim(txtChaveNFe.Text) <> "" Then
      If Len(Trim(txtChaveNFe.Text)) <> 44 Then
         MsgBox "Chave informada inválida, verifique."
         txtChaveNFe.Text = ""
         txtChaveNFe.SetFocus
         Exit Sub
      End If
    End If

   CRITERIO_A = txtCep.Text
   CRITERIO_A = Replace(CRITERIO_A, "-", "")
   If Len(Trim(CRITERIO_A)) < 8 Then
      MsgBox "CEP inválido, deve conter 8 digitos !!!"
      txtCep.SetFocus
      Exit Sub
   End If

   If Trim(cmbIE.Text) = "" Then _
      cmbIE.Text = "ISENTO"

   If Trim(txtCidade.Text) = "" Then
      MsgBox "Cidade inválido !!!"
      txtCidade.SetFocus
      Exit Sub
   End If
   If Trim(txtIBGE.Text) = "" Then
      MsgBox "Código IBGE inválido !!!"
      txtIBGE.SetFocus
      Exit Sub
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from IBGE WITH (NOLOCK)"
         SQL = SQL & " where IBGE_ID = " & txtIBGE.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
            Exit Sub
            Else
               If Not IsNull(TabTemp.Fields("estado").Value) Then
                  If Trim(UCase(TabTemp.Fields("estado").Value)) <> Trim(UCase(txtUF_CLIENTE.Text)) Then
                     MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
                     Exit Sub
                  End If
                  Else
                     MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
                     Exit Sub
               End If
               If Not IsNull(TabTemp.Fields("municipio").Value) Then
                  Else
                     MsgBox "Erro no IBGE, verificar." & txtIBGE.Text
                     Exit Sub
               End If
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
   End If

frmINTEGRA.INTEGRA_IBGE txtIBGE.Text

   If Trim(txtFone.Text) = "" Then
      MsgBox "Fone inválido !!!"
      txtFone.Enabled = True
      txtFone.SetFocus
      Exit Sub
   End If

   If TxtQuantidadeRodapeNota.Text = "" Then
      MsgBox "Digite Quantidade de Produtos a Transportar!"
      TxtQuantidadeRodapeNota.SetFocus
      Exit Sub
      Else
         If Not IsNumeric(TxtQuantidadeRodapeNota.Text) Then
            MsgBox "Favor Digite Quantidade de Produtos a Transportar!"
            TxtQuantidadeRodapeNota.SetFocus
            Exit Sub
         End If
   End If

   If TxtEspecie.Text = "" Then
      MsgBox "Digite a Especie a Transportar!"
      TxtEspecie.SetFocus
      Exit Sub
   End If

   If TxtPesoBruto.Text = "" Then
      MsgBox "Digite o Peso Bruto a Transportar!"
      TxtPesoBruto.SetFocus
      Exit Sub
   End If

   If TxtPesoLiquido.Text = "" Then
      MsgBox "Digite o Peso Liquido a Transportar!"
      TxtPesoLiquido.SetFocus
      Exit Sub
   End If

   txtCNPJCPF.PromptInclude = False

'=============================== NOTA DE DEVOLUÇÃO
   If Trim(Left(TIPO_NFe_GERAR, 1)) = "D" Then
      If Indr_Consulta = False Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select CLIENTE.CLIENTE_ID, CLIENTE.PESSOA_ID, CLIENTE.CGCCPF, "
         SQL = SQL & " PEDIDO.PEDIDO_ID, PEDIDO.EMPRESA_ID, "
         SQL = SQL & " PEDIDO.VENDEDOR_ID, "
         SQL = SQL & " PEDIDO.DT_REQ, PEDIDO.STATUS, PEDIDO.TIPO_REGISTRO, "
         SQL = SQL & " PEDIDO.VALOR_DESCONTO, PEDIDO.VALOR_TOTAL, cliente.pessoa_id, "
         SQL = SQL & " PEDIDO.ESTABELECIMENTO_ID, Cliente.NOME"

         SQL = SQL & " from PEDIDO WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN CLIENTE WITH (NOLOCK)"
         SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID"

         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and left(tipo_registro,1) = 'D'"
         SQL = SQL & " and cliente.status = 'A'"
         SQL = SQL & " and CLIENTE.cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            PESSOA_ID_N = TabCabeca.Fields("pessoa_id").Value

            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF_TRANSP.PromptInclude = False

If Trim(txtNOTA.Text) = "" Then
   txtNOTA.Text = "" & GERA_NUMERO_NFe_N
   txtNOTA.Refresh
End If

GRAVA_NOTA txtNOTA.Text, _
           txtSerie.Text, _
           txtMODELO.Text, _
           TIPO_NFe_GERAR, _
           TxtQuantidadeRodapeNota.Text, _
           TxtPesoBruto.Text, _
           TxtPesoLiquido.Text, _
           cmbPresencaAUX.Text, _
           cmbLocal.Text, _
           cmbCFOPAux.Text, _
           Trim(txtCNPJCPF_TRANSP.Text)

            SQL = "select * from NFITEM WITH (NOLOCK)"
            SQL = SQL & " where nf_id = " & NUMR_ID_N
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               TabTemp.MoveFirst
               While Not TabTemp.EOF
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  'Ler Tab. Produtos Para pegar tributacao e nacionalidade
                  SQL = "select SITUACAO_TRIBUTARIA from PRODUTO WITH (NOLOCK)"
                  SQL = SQL & " where produto_id = " & TabTemp.Fields("produto_id").Value
                  SQL = SQL & " and situacao <> 'C' "
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabProduto.EOF Then _
                     strTributacao = TabProduto!SITUACAO_TRIBUTARIA
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  If TabAUX.State = 1 Then _
                     TabAUX.Close

                  TabTemp.MoveNext
               Wend
               GRAVA_STATUS_EMITIDO
            End If
'======================================
               IMPRESSAO_NF
               Unload Me
'======================================
         End If
         Else
            IMPRESSAO_NF
            Unload Me
      End If
   End If
'===================== NOTA DE PEDIDO VENDA
   If Trim(TIPO_NFe_GERAR) = "R" Then
      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      'chegando itens
      SQL = "select PEDIDO.*, CLIENTE.PESSOA_ID "
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN CLIENTE WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID"

      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCabeca.EOF Then
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
         MsgBox "Pedido não encontrado."
         Exit Sub
      End If
      If Not TabCabeca.EOF Then
         PESSOA_ID_N = TabCabeca.Fields("pessoa_id").Value

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & TabCabeca.Fields("pedido_id").Value
         SQL = SQL & " and tipo_reg = 'PC' "
         SQL = SQL & " and pedidoitem.status <> 'C' "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close
            MsgBox "Pedido não possue itens, não permitido."
            Exit Sub
         End If
         While Not TabTemp.EOF
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from PRODUTO WITH (NOLOCK)"
            SQL = SQL & " where produto_id = " & TabTemp.Fields("produto_id").Value
            SQL = SQL & " and situacao <> 'C' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabConsulta.EOF Then
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               MsgBox "Pedido Produto não cadastrado, não permitido."
               Exit Sub
               Else
                  If IsNull(TabConsulta.Fields("situacao").Value) Then
                     MsgBox "Situação do produto inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        If Trim(TabConsulta.Fields("situacao").Value) <> "A" Then
                           MsgBox "Situação do produto inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                        End If
                  End If

                  QTDE_PEDIDO = 0
                  QTDE_PEDIDO = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabConsulta.Fields("produto_id").Value)

                  If INDR_ESTQ_NEGATIVO = False Then
                     If Indr_Consulta = False Then
                        If Not IsNull(TabTemp.Fields("STATUS").Value) Then
                           If Trim(UCase((TabTemp.Fields("STATUS").Value))) <> "B" Then
                              If IsNull(QTDE_PEDIDO) Then
                                 MsgBox "Qtde disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                                 Exit Sub
                                 Else
                                    If QTDE_PEDIDO <= 0 Then
                                       MsgBox "Qtde disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                                       Exit Sub
                                    End If
                              End If
                           End If
                           Else: MsgBox "Verificar situação do item no pedido."
                        End If
                     End If
                  End If

                  If IsNull(TabConsulta.Fields("PRECO_VENDA").Value) Then
                     MsgBox "Valor_venda disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        VALOR_ITEM_N = TabConsulta.Fields("PRECO_VENDA").Value
                        If VALOR_ITEM_N <= 0 Then
                           MsgBox "Valor_venda disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                           Else
                              VALOR_01 = TabConsulta.Fields("PRECO_VENDA").Value
                              VALOR_01 = TabTemp.Fields("valor_item").Value
                              VALOR_02 = TabConsulta.Fields("PRECO_CUSTO").Value
                              If VALOR_01 < VALOR_02 Then
                                 MsgBox "Produto : " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value) & " , venda abaixo preço de custo."
                              End If
                        End If
                  End If
                  If IsNull(TabConsulta.Fields("CODG_NCM").Value) Then
                     MsgBox "CODG_NCM disponível inválida. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                     Exit Sub
                     Else
                        VALOR_01 = TabConsulta.Fields("CODG_NCM").Value
                        If VALOR_01 <= 0 Then
                           MsgBox "Código NCM do produto inválido. " & Trim(TabConsulta.Fields("codg_produto").Value) & "-" & Trim(TabConsulta.Fields("descricao").Value)
                           Exit Sub
                        End If
                  End If
            End If

            TabTemp.MoveNext

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF_TRANSP.PromptInclude = False

         If Trim(txtNOTA.Text) = "" Then
            txtNOTA.Text = "" & GERA_NUMERO_NFe_N
            txtNOTA.Refresh
         End If

GRAVA_NOTA txtNOTA.Text, _
           txtSerie.Text, _
           txtMODELO.Text, _
           TIPO_NFe_GERAR, _
           TxtQuantidadeRodapeNota.Text, _
           TxtPesoBruto.Text, _
           TxtPesoLiquido.Text, _
           cmbPresencaAUX.Text, _
           cmbLocalAUX.Text, _
           cmbCFOPAux.Text, _
           Trim(txtCNPJCPF_TRANSP.Text)

         IMPRESSAO_NF
      End If
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERAR_NFE_old"
End Sub

Private Sub MONTA_NOTA_DV()
'On Error GoTo ERRO_TRATA

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      'passou do PEDIDO, checar na tabela nf agora
      SQL = "select * from NF WITH (NOLOCK)"
      SQL = SQL & " where nf_id = " & NF_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         cmbTipoOperaAUX.Text = 0
         cmbTipoOpera.Text = "0-Entrada"

   cmbPresencaAUX.Text = 1
   cmbPresenca.Text = "1-Operação presencial"

         txtNOTA.Text = "" & TabNOTA!NUMR_NOTA
         txtSerie.Text = "" & TabNOTA!SERIE_NOTA
         txtMODELO.Text = "" & TabNOTA.Fields("modelo_doc").Value
         txtDtEmis.Text = "" & Format(TabNOTA!DT_EMISSAO, "dd/mm/yyyy")
         txtDtSaida.Text = "" & Format(TabNOTA!DT_ENTRASAI, "dd/mm/yyyy")

         If Not IsNull(TabNOTA!TRANSP_ID) Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select cnpjcpf,descricao from vwTRANSPORTADORA WITH (NOLOCK)"
            SQL = SQL & " where cnpjcpf = '" & TabNOTA!TRANSP_ID & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp.Fields(0).Value) Then
                  cmbCNPJCPF_TRANSP.Text = "" & Trim(TabTemp!CNPJCPF) & " - " & Trim(TabTemp!DESCRICAO)
                  txtCNPJCPF_TRANSP.Text = "" & Trim(TabTemp.Fields(0).Value)
                  'Volumes
                  TxtQuantidadeRodapeNota.Text = "" & TabNOTA!Qtd_Volume
                  TxtEspecie.Text = "UN"
                  TxtPesoBruto.Text = "" & TabNOTA!Peso_Bruto
                  TxtPesoLiquido.Text = "" & TabNOTA!PESO_LIQUIDO
               End If
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

'=====================================================
         If txtDtEmis.Text = "" Then _
            txtDtEmis.Text = Format(Date, "dd/mm/yyyy")
         If txtDtSaida.Text = "" Then _
            txtDtSaida.Text = Format(Date, "dd/mm/yyyy")

         If txtSerie.Text = "" Then
            txtSerie.Text = 1
            txtSerie.Refresh
         End If

         cmbEmail.Text = ""
         INDR_PRI = False

         If TabEmail.State = 1 Then _
            TabEmail.Close

         SQL = "select EMAIL.EMAIL from PESSOA WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN CLIENTE WITH (NOLOCK)"
         SQL = SQL & " ON PESSOA.PESSOA_ID = CLIENTE.PESSOA_ID "
         SQL = SQL & " INNER JOIN EMAIL WITH (NOLOCK)"
         SQL = SQL & " ON PESSOA.PESSOA_ID = EMAIL.PESSOA_ID"
         SQL = SQL & " where email.pessoa_id = " & PESSOA_ID_N
         TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabEmail.EOF
            cmbEmail.AddItem "" & Trim(TabEmail.Fields(0).Value)
            cmbEmail.Text = "" & Trim(TabEmail.Fields(0).Value)
            TabEmail.MoveNext
         Wend
         If TabEmail.State = 1 Then _
            TabEmail.Close

      Dim i
      For i = 1 To Len(cmbEmail.Text)
         If Mid(cmbEmail.Text, i, 1) <> " " Then
            If Mid(cmbEmail.Text, i, 1) = "@" Then
               INDR_PRI = False
               Exit For
            End If
            Else
               Exit For
         End If
      Next

      cmbEmail.ForeColor = vbRed
      cmbEmail.Refresh
      '================================
      cmbIE.Clear
      If TabEmail.State = 1 Then _
         TabEmail.Close

      SQL = "select numr_ie from IE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEmail.EOF Then
         If Trim(TabEmail.Fields("numr_ie").Value) <> "" Then
            cmbIE.Text = Trim(TabEmail.Fields("numr_ie").Value)
            Else: MsgBox "Inscrição Estadual inválida !!!"
         End If
      End If
      If TabEmail.State = 1 Then _
         TabEmail.Close

      SQL = "select * from FONE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero <> ''"
      TabEmail.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEmail.EOF Then
         txtFone.Text = "" & Trim(Left(TabEmail.Fields("numero").Value, 10))
         Else: MsgBox "Fone de cliente não encontrado !!!"
      End If
      If TabEmail.State = 1 Then _
         TabEmail.Close
'=============================

      MOSTRA_CLIENTE

      If Trim(cmbCFOPAux.Text) <> "" Then
         NaturezaOperacao_A = "" & Trim(TRAZ_CFOP(Trim(cmbCFOPAux.Text)))
         Else: MsgBox "Problemas no CFOP."
      End If

      TOTAIS_NOTA_DV

      SETA_GRID_DV

   txtMSG.Text = " NFe DEV.REF: " & txtNOTA.Text

'=====================================================

         If TabCabeca.State = 1 Then _
            TabCabeca.Close
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         fraNota.Enabled = False
         fraEmitente.Enabled = False
         Frame3.Enabled = False
         Frame4.Enabled = True
         Frame5.Enabled = True
         Frame8.Enabled = False
         Exit Sub
         Else
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            MsgBox "Registro de venda não encontrado."
            fraNota.Enabled = False
            fraEmitente.Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = True
            Frame5.Enabled = True
            Exit Sub
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_NOTA_DV"
End Sub

Private Sub TOTAIS_NOTA_DV()
'On Error GoTo ERRO_TRATA

   Dim VALOR_TOTAL_PRODUTO_N  As Double

   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_PRODUTO_N = 0

   txtBaseCalculo.Text = "" & Format(0, strFormatacao2Digitos)
   txtValorICMS.Text = "" & Format(0, strFormatacao2Digitos)
   txtBaseIcmsSub.Text = "" & Format(0, strFormatacao2Digitos)
   txtFrete.Text = "" & Format(0, strFormatacao2Digitos)
   txtValorIPI.Text = "" & Format(0, strFormatacao2Digitos)
   txtValorProdutos.Text = "" & Format(0, strFormatacao2Digitos)
   txtDesconto.Text = "" & Format(0, strFormatacao2Digitos)
   txtVlrIcmsSub.Text = "" & Format(0, strFormatacao2Digitos)
   txtValorOutros.Text = "" & Format(0, strFormatacao2Digitos)
   txtValorTotalNota.Text = "" & Format(0, strFormatacao2Digitos)

  'valor de desconto na cabeça
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_TOTAL_PRODUTO_N = 0

   SQL = "select sum(valor*qtdE) from NFITEM WITH (NOLOCK)"
   SQL = SQL & " where nf_id = " & NF_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         VALOR_TOTAL_PRODUTO_N = TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   VALOR_TOTAL_N = VALOR_TOTAL_PRODUTO_N - VALOR_DESCONTO_N

   txtValorTotalNota.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
   txtValorProdutos.Text = Format(VALOR_TOTAL_PRODUTO_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOTAIS_NOTA_DV"
End Sub

Private Sub SETA_GRID_DV()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear
   MSFlexGrid1.Visible = False
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select codg_produto as Código,descricao as Descrição, "
   SQL = SQL & " qtde as Qtde,Valor as PreçoVenda, (qtde*Valor) as Total,"
   SQL = SQL & " cfop_id as CFOP, stributaria as ST, PercIcms as ICMS,"
   SQL = SQL & " codg_ncm as NCM,Unidade_Medida as UN, NFITEM.produto_id,nf_id,seq_id"

   SQL = SQL & " from NFITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where nf_id = " & NF_ID_N
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabProduto.Fields.Count

      ReDim largura_coluna(0 To TabProduto.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabProduto.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabProduto.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabProduto.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabProduto.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabProduto.Fields.Count - 1
            If Coluna = 2 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabProduto.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna = 3 Or Coluna = 4 Or Coluna = 7 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabProduto.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabProduto.Fields(Coluna).Value)
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabProduto.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabProduto.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabProduto.MoveNext
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

'Codigo Produto
      MSFlexGrid1.ColWidth(0) = 1000
      MSFlexGrid1.ColAlignment(0) = 0

'Descrição Produto
      MSFlexGrid1.ColWidth(1) = 4000
      MSFlexGrid1.ColAlignment(1) = 0

'QTDE
      MSFlexGrid1.ColWidth(2) = 1500
      MSFlexGrid1.ColAlignment(2) = 7

'Valor Item
      MSFlexGrid1.ColWidth(3) = 1500
      MSFlexGrid1.ColAlignment(3) = 7

'Total Item
      MSFlexGrid1.ColWidth(4) = 1500
      MSFlexGrid1.ColAlignment(4) = 7

'cfop
      MSFlexGrid1.ColWidth(5) = 1000
      MSFlexGrid1.ColAlignment(5) = 7

'SITUAÇÃO TRIBUTARIA PRODUTO
      MSFlexGrid1.ColWidth(6) = 500
      MSFlexGrid1.ColAlignment(6) = 0

'ALIQUOTA ICMS
      MSFlexGrid1.ColWidth(7) = 1000
      MSFlexGrid1.ColAlignment(7) = 7

'NCM
      MSFlexGrid1.ColWidth(8) = 1000
      MSFlexGrid1.ColAlignment(8) = 0

'UN
      MSFlexGrid1.ColWidth(9) = 500
      MSFlexGrid1.ColAlignment(9) = 0

'produto_id
      MSFlexGrid1.ColWidth(10) = 0
      MSFlexGrid1.ColAlignment(10) = 0

'nf_id
      MSFlexGrid1.ColWidth(11) = 0
      MSFlexGrid1.ColAlignment(11) = 0

'seq_id
      MSFlexGrid1.ColWidth(12) = 0
      MSFlexGrid1.ColAlignment(12) = 0
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

   MSFlexGrid1.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_DV"
End Sub
