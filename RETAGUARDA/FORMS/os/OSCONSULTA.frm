VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSConsulta 
   Caption         =   "Consulta Ordem de Serviço"
   ClientHeight    =   7785
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OSCONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleMode       =   0  'User
   ScaleWidth      =   41067.56
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   11895
      Begin PVMaskEditLib.PVMaskEdit txtPlaca 
         Height          =   360
         Left            =   3720
         TabIndex        =   0
         ToolTipText     =   "Informe a Placa do Veículo"
         Top             =   240
         Width           =   1335
         _Version        =   524288
         _ExtentX        =   2355
         _ExtentY        =   635
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optVeiculo 
            Caption         =   "&Veículo"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optEqp 
            Caption         =   "&Equipamento"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtEqp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   37
         ToolTipText     =   "Informe Kilometragem atual"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbSituacaoAUX 
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
         Left            =   10320
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmblstOSAUX 
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
         Left            =   10920
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optDtFechamento 
         Caption         =   "D&ata Fechamento"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   2640
         Width           =   2055
      End
      Begin VB.OptionButton optDtAbertura 
         Caption         =   "D&ata Abertura"
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
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cmbTipoOSAUX 
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
         Left            =   10320
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbTIPOOS 
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
         Left            =   10245
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbProdutoAUX 
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
         Left            =   7320
         TabIndex        =   30
         Top             =   2070
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbProduto 
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
         Left            =   6840
         TabIndex        =   9
         Top             =   2070
         Width           =   4935
      End
      Begin VB.TextBox txtOS 
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
         Left            =   5970
         MaxLength       =   100
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbServicoAUX 
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
         Left            =   960
         TabIndex        =   26
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbServico 
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
         Left            =   960
         TabIndex        =   8
         Top             =   2040
         Width           =   4935
      End
      Begin VB.ComboBox cmbMecanicoAUX 
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
         Left            =   5490
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbMECANICO 
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
         Left            =   5520
         TabIndex        =   7
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cmbConsultorAUX 
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
         Left            =   1920
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtNome 
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
         Left            =   3720
         MaxLength       =   100
         TabIndex        =   15
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtCHASSI 
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
         Left            =   8760
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox cmbConsultor 
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cmbSituacao 
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
         Left            =   10200
         TabIndex        =   5
         Top             =   870
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   360
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   375
         Left            =   8400
         TabIndex        =   13
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
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
         Left            =   7320
         TabIndex        =   34
         Top             =   2640
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
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
         Left            =   4080
         TabIndex        =   33
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo O.S.:"
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
         Left            =   9240
         TabIndex        =   31
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   6000
         TabIndex        =   29
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O.S.:"
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
         Left            =   5190
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbltipo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3060
         TabIndex        =   27
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serviço:"
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
         Left            =   135
         TabIndex        =   25
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mecânico:"
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
         Left            =   4485
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblNome 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   8040
         TabIndex        =   20
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   150
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consultor Técnico:"
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
         TabIndex        =   18
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label lblRep 
         AutoSize        =   -1  'True
         Caption         =   "Situação:"
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
         Left            =   9240
         TabIndex        =   17
         Top             =   840
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   960
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1693
      ButtonWidth     =   1535
      ButtonHeight    =   1535
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
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
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3360
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":941B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":AC2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":C007
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":D317
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTA.frx":E422
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView lstOS 
      Height          =   3540
      Left            =   0
      TabIndex        =   14
      Top             =   4200
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6244
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ILTw"
      BorderStyle     =   1
      Appearance      =   1
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11940
      DesignHeight    =   7785
   End
End
Attribute VB_Name = "frmOSConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim VALOR_TOTAL_PRODUTO_N As Double
   Dim VALOR_TOTAL_SERVICO_N As Double
   Dim QTD_COTAS             As Long

Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   LIMPA_CONSULTA

   CARREGA_COMBOS

   If optEqp.Value = True Then
      txtEqp.Visible = True
      txtPlaca.Visible = False
      lbltipo.Caption = "Equip.:"
      Label2.Caption = "Técnico:"
      lblNome.Visible = False
      txtCHASSI.Visible = False
      SETA_GRID_EQP
   End If
   If optVeiculo.Value = True Then
      txtEqp.Visible = False
      txtPlaca.Visible = True
      lbltipo.Caption = "Placa:"
      SETA_GRID_VEICULO
   End If
   lbltipo.Refresh

   optDtAbertura.Value = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub optEqp_Click()
   If optEqp.Value = True Then
      txtEqp.Visible = True
      txtPlaca.Visible = False
      lbltipo.Caption = "Equip.:"
      Label2.Caption = "Técnico:"
      lblNome.Visible = False
      txtCHASSI.Visible = False
      Else
         txtEqp.Visible = False
         txtPlaca.Visible = True
         lbltipo.Caption = "Placa:"
   End If
      
   SETA_GRID_EQP
End Sub

Private Sub optVeiculo_Click()
   If optEqp.Value = True Then
      txtEqp.Visible = True
      txtPlaca.Visible = False
      lbltipo.Caption = "Equip.:"
      Label2.Caption = "Técnico:"
      lblNome.Visible = False
      txtCHASSI.Visible = False
      Else
         txtEqp.Visible = False
         txtPlaca.Visible = True
         lbltipo.Caption = "Placa:"
   End If
SETA_GRID_VEICULO
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "limpar"
         LIMPA_CONSULTA
      Case "consultar"
         If optEqp.Value = True Then
            SETA_GRID_EQP
            Else: SETA_GRID_VEICULO
         End If
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub LSTOS_DblClick()
On Error Resume Next

   PEDIDO_ID_N = 0
   SQL3 = ""
   cmblstOSAUX.ListIndex = (lstOS.Nodes(lstOS.SelectedItem.key).Index) - 1
   If Not IsNull(cmblstOSAUX.Text) Then
      If cmblstOSAUX.Text <> "" Then
         If IsNumeric(cmblstOSAUX.Text) Then
            SQL3 = cmblstOSAUX.Text
            Unload Me
         End If
      End If
   End If
End Sub

Private Sub cmbProduto_Click()
On Error Resume Next

   cmbProdutoAUX.ListIndex = cmbProduto.ListIndex
End Sub

Private Sub cmbConsultor_Click()
On Error Resume Next

   cmbConsultorAUX.ListIndex = cmbConsultor.ListIndex
End Sub

Private Sub cmbmecanico_Click()
On Error Resume Next

   cmbMecanicoAUX.ListIndex = cmbMecanico.ListIndex
End Sub

Private Sub cmbServico_Click()
On Error Resume Next

   cmbServicoAUX.ListIndex = cmbServico.ListIndex
End Sub

Private Sub cmbtipoos_Click()
On Error Resume Next

   cmbTipoOSAUX.ListIndex = cmbTipoOS.ListIndex
End Sub

Private Sub cmbSituacao_Click()
On Error Resume Next

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex
End Sub

Private Sub cmbSituacao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbSituacao_KeyPress"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF.Mask)

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

   If Trim(txtCHASSI.Text) <> "" Then
      cmbSituacao.SetFocus
      Exit Sub
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If CNPJCPF_A <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
         CNPJCPF_A = ""
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbSituacao.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcnpjcpf_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "" Then

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE "
      SQL = SQL & " where cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         txtNome.Text = "" & Trim(TabCliente!NOME)
         PESSOA_ID_N = TabCliente.Fields("pessoa_id").Value
         Else
            txtCNPJCPF.SelStart = 0
            txtCNPJCPF.SelLength = Len(txtOs)
            txtCNPJCPF.SetFocus
            MsgBox "Cliente não Cadastrado."
            txtCNPJCPF.SetFocus
            Exit Sub
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

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
   End If
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_LostFocus"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
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

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
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

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPlaca.Text <> "" Then
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select placa from OSVEICULO "
         SQL = SQL & " where placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabAUX.EOF Then
            MsgBox "Placa não cadastrado."
            txtPlaca.SetFocus
            Exit Sub
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close
      End If
      cmbSituacao.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPlaca_KeyPress"
End Sub

Private Sub SETA_GRID_VEICULO()
'On Error GoTo ERRO_TRATA

   Dim VACA_VEIA

   cmblstOSAUX.Clear
   lstOS.Nodes.Clear

   CONT_N = 0
   VACA_VEIA = "Endereço"
   QTD_COTAS = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS"
   SQL = SQL & " where placa <> '' "

   If Trim(txtPlaca.Text) <> "" Then _
      SQL = SQL & " and placa = '" & Replace(Trim(txtPlaca.Text), "-", "") & "'"

   If Trim(txtOs.Text) <> "" Then _
      If IsNumeric(txtOs.Text) Then _
         SQL = SQL & " and os_id = " & txtOs.Text

   If Trim(txtCHASSI.Text) <> "" Then _
      SQL = SQL & " and chassi = '" & Trim(txtCHASSI.Text) & "'"

   If Trim(cmbTipoOSAUX.Text) <> "" Then _
      If IsNumeric(cmbTipoOSAUX.Text) Then _
         SQL = SQL & " and tipo_os = " & Trim(cmbTipoOSAUX.Text)

   txtCNPJCPF.PromptInclude = False
   If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and OSEQUIPAMENTO.pessoa_id = " & PESSOA_ID_N

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao_os = " & cmbSituacaoAUX.Text

   If Trim(cmbConsultorAUX.Text) <> "" Then _
      If IsNumeric(cmbConsultorAUX.Text) Then _
         SQL = SQL & " and ct_id = " & cmbConsultorAUX.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optDtAbertura.Value = True Then
         SQL = SQL & " and dt_os >= '" & txtDtIni.Text & "'"
         SQL = SQL & " and dt_os <= '" & txtDtFim.Text & "'"
      End If
      If optDtFechamento.Value = True Then
         SQL = SQL & " and dt_fecha >= '" & txtDtIni.Text & "'"
         SQL = SQL & " and dt_fecha <= '" & txtDtFim.Text & "'"
      End If
   End If

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If TabTemp.EOF Then _
      MsgBox "Não existe O.S. para essa pesquisa."

   While Not TabTemp.EOF
'=====================================
      'totais serviço para cabeça do grid
      VALOR_TOTAL_SERVICO_N = 0
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum(valor_servico-desconto_servico) from OSSERVICO "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         VALOR_TOTAL_SERVICO_N = 0 & TabConsulta.Fields(0).Value

      'totais produto para cabeça do grid
      VALOR_TOTAL_PRODUTO_N = 0
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum((valor_item-desconto_produto) * qtde) from OSPECA "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         VALOR_TOTAL_PRODUTO_N = 0 & TabConsulta.Fields(0).Value
'====================================
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = TabTemp.Fields("cnpjcpf").Value
      If Len(Trim(TabTemp.Fields("cnpjcpf").Value)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If
      txtCNPJCPF.PromptInclude = True

      NOME_A = "" & Trim(TabTemp.Fields("cliente").Value)

      cmblstOSAUX.AddItem TabTemp.Fields("OS_ID").Value
      cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

      Set Nodx = lstOS.Nodes.Add(, , VACA_VEIA & QTD_COTAS, "O.S. = " & Trim(TabTemp.Fields("OS_ID").Value) & _
      " ; Total O.S. = " & Format(VALOR_TOTAL_PRODUTO_N + VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos) & _
      " ; Cliente: " & Trim(txtCNPJCPF.Text) & " - " & Trim(NOME_A) & " ; Placa: " & Trim(TabTemp.Fields("placa").Value))

      lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlack

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      'SERVIÇO
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from OSSERVICO "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value

      If Trim(cmbMecanicoAUX.Text) <> "" Then _
         If IsNumeric(cmbMecanicoAUX.Text) Then _
            SQL = SQL & " and mecanico_id = " & cmbMecanicoAUX.Text

      If Trim(cmbServico.Text) <> "" Then _
         SQL = SQL & " and descricao_serviço like " & cmbServicoAUX.Text & "%"

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         cmblstOSAUX.AddItem TabConsulta.Fields("OS_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         Set Nodx = lstOS.Nodes.Add(VACA_VEIA & QTD_COTAS, tvwChild, "Tarefa" & CONT_N, "Servico(s) = R$ " & Format(VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos))

         lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlue
      End If

      VALOR_TOTAL_SERVICO_N = 0

      While Not TabConsulta.EOF
         'TAREFAS
         NOME_A = "" & Trim(TabConsulta.Fields("descricao").Value)

         Set Nodx = lstOS.Nodes.Add("Tarefa" & CONT_N, tvwChild, , Trim(TabConsulta.Fields("osservico_id").Value) _
         & " -  " & Trim(NOME_A) & " ; Valor = " & Format(TabConsulta.Fields("valor_servico").Value, strFormatacao2Digitos) & _
         " ; Desconto = " & Format(TabConsulta.Fields("desconto_servico").Value, strFormatacao2Digitos) & " ; Total = " & _
         Format(TabConsulta.Fields("valor_servico").Value - TabConsulta.Fields("desconto_servico").Value, strFormatacao2Digitos))

         cmblstOSAUX.AddItem TabConsulta.Fields("OSSERVICO_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         VALOR_TOTAL_SERVICO_N = VALOR_TOTAL_SERVICO_N + _
                                 TabConsulta.Fields("valor_servico").Value - _
                                 TabConsulta.Fields("desconto_servico").Value

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TOTAL PEÇAS
      SQL = "select OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
      SQL = SQL & " from OSPECA "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value

      If Trim(cmbProdutoAUX.Text) <> "" Then _
         SQL = SQL & " and produto_id = " & cmbProdutoAUX

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         cmblstOSAUX.AddItem TabTemp.Fields("OS_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         Set Nodx = lstOS.Nodes.Add(VACA_VEIA & QTD_COTAS, tvwChild, "Produtos" & CONT_N, "Produto(s) = R$ " & Format(VALOR_TOTAL_PRODUTO_N, strFormatacao2Digitos))

         lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbRed
      End If

      VALOR_TOTAL_PRODUTO_N = 0

      While Not TabConsulta.EOF
         'PEÇAS
         NOME_A = "" & Trim(TabConsulta.Fields("descricao").Value)

         Set Nodx = lstOS.Nodes.Add("Produtos" & CONT_N, tvwChild, , Trim(TabConsulta.Fields("produto_id").Value) _
         & " -  " & Trim(NOME_A) & " ; Valor = " & Format(TabConsulta.Fields("VALOR_ITEM").Value * TabConsulta!QTDE, strFormatacao2Digitos) & _
         " ; Desconto = " & Format(TabConsulta.Fields("desconto_produto").Value, strFormatacao2Digitos) & " ; Total = " & _
         Format((TabConsulta.Fields("VALOR_ITEM").Value - TabConsulta!DESCONTO_PRODUTO) * TabConsulta!QTDE, strFormatacao2Digitos))

         cmblstOSAUX.AddItem TabConsulta.Fields("OSpeca_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         VALOR_TOTAL_PRODUTO_N = (TabConsulta.Fields("VALOR_ITEM").Value - TabConsulta!DESCONTO_PRODUTO) * TabConsulta!QTDE

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      QTD_COTAS = QTD_COTAS + 1
      CONT_N = CONT_N + 1

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstOS.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_VEICULO"
End Sub

Private Sub SETA_GRID_EQP()
'On Error GoTo ERRO_TRATA

   Dim TabOS      As New ADODB.Recordset
   Dim TabServico As New ADODB.Recordset
   Dim TabPeca    As New ADODB.Recordset
   Dim DtFecha_D  As Date

   cmblstOSAUX.Clear
   lstOS.Nodes.Clear

   CONT_N = 0
   SQL3 = "Endereço"
   QTD_COTAS = 0

   If TabOS.State = 1 Then _
      TabOS.Close

   SQL = "select * from vwOS "

   SQL = SQL & " where os_id > 0 "

   If Trim(txtEqp.Text) <> "" Then _
      SQL = SQL & " and equipamento_id = " & txtEqp.Text

   If Trim(txtOs.Text) <> "" Then _
      If IsNumeric(txtOs.Text) Then _
         SQL = SQL & " and os_id = " & txtOs.Text

   If Trim(txtCHASSI.Text) <> "" Then _
      SQL = SQL & " and chassi = '" & Trim(txtCHASSI.Text) & "'"

   If Trim(cmbTipoOSAUX.Text) <> "" Then _
      If IsNumeric(cmbTipoOSAUX.Text) Then _
         SQL = SQL & " and tipo_os = " & Trim(cmbTipoOSAUX.Text)

   txtCNPJCPF.PromptInclude = False
   If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao_os = " & cmbSituacaoAUX.Text

   If Trim(cmbConsultorAUX.Text) <> "" Then _
      If IsNumeric(cmbConsultorAUX.Text) Then _
         SQL = SQL & " and ct_id = " & cmbConsultorAUX.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optDtAbertura.Value = True Then
         SQL = SQL & " and dt_os >= '" & (txtDtIni.Text) & "'"
         SQL = SQL & " and dt_os <= '" & (txtDtFim.Text) & "'"
      End If
      If optDtFechamento.Value = True Then
         SQL = SQL & " and dt_fecha >= '" & (txtDtIni.Text) & "'"
         SQL = SQL & " and dt_fecha <= '" & (txtDtFim.Text) & "'"
      End If
   End If

   SQL = SQL & " order by OS_ID desc"

   TabOS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If TabOS.EOF Then _
      MsgBox "Não existe O.S. para essa pesquisa."

   While Not TabOS.EOF
      If PEDIDO_ID_N <> TabOS.Fields("OS_ID").Value Then
         '====================================cliente
         txtCNPJCPF.PromptInclude = False

         'txtCNPJCPF.Text = tabos.Fields("cnpjcpf").Value
         If Len(Trim(TabOS.Fields("cnpjcpf").Value)) <= 11 Then
            txtCNPJCPF.Mask = "###.###.###-##"
            Else: txtCNPJCPF.Mask = "##.###.###/####-##"
         End If

         txtCNPJCPF.Text = "" & TabOS.Fields("cnpjcpf").Value
         NOME_A = "" & Trim(TabOS.Fields("cliente").Value)
         PESSOA_ID_N = 0 & TabOS.Fields("pessoa_id").Value

         '======================

         'SQL = "select descricao,cnpjcpf,pessoa_id from PESSOA "
         'SQL = SQL & " where pessoa_id = " & TabOS.Fields("PESSOA_ID").Value
         'TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         'If Not TabConsulta.EOF Then
         '   txtCNPJCPF.Text = "" & Trim(TabConsulta.Fields("cnpjcpf").Value)
         '   NOME_A = "" & Trim(TabConsulta.Fields("descricao").Value)
         '   PESSOA_ID_N = 0 & Trim(TabConsulta.Fields("pessoa_id").Value)
         'End If
         'If TabConsulta.State = 1 Then _
            TabConsulta.Close

         txtCNPJCPF.PromptInclude = True

         'If Trim(TabOS.Fields("cliente").Value) <> "" Then _
            NOME_A = "" & Trim(TabOS.Fields("cliente").Value)
'==============
         cmblstOSAUX.AddItem TabOS.Fields("OS_ID").Value
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         '========================      'totais serviço para cabeça do grid
         VALOR_TOTAL_SERVICO_N = 0
         If TabServico.State = 1 Then _
            TabServico.Close

         SQL = "select sum(valor_servico-desconto_servico) from OSSERVICO "
         SQL = SQL & " where os_id = " & TabOS.Fields("os_id").Value
         TabServico.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabServico.EOF Then _
            VALOR_TOTAL_SERVICO_N = 0 & TabServico.Fields(0).Value
         If TabServico.State = 1 Then _
            TabServico.Close

         'totais produto para cabeça do grid
         VALOR_TOTAL_PRODUTO_N = 0
         If TabPeca.State = 1 Then _
            TabPeca.Close
   
         SQL = "select sum((valor_item-desconto_produto) * qtde) from OSPECA "
         SQL = SQL & " where os_id = " & TabOS.Fields("os_id").Value
         TabPeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPeca.EOF Then _
            VALOR_TOTAL_PRODUTO_N = 0 & TabPeca.Fields(0).Value
         If TabPeca.State = 1 Then _
            TabPeca.Close
'======================
         SqL2 = "" & TRAZ_DESCRITOR("Z", TabOS.Fields("situacao_os").Value)
         
         QTD_COTAS = QTD_COTAS + 1
         Set Nodx = lstOS.Nodes.Add(, , SQL3 & QTD_COTAS, "O.S. = " & Trim(TabOS.Fields("OS_ID").Value) & _
         " ; Total O.S. = " & Format(VALOR_TOTAL_PRODUTO_N + VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos) & _
         " ; Cliente: " & Trim(txtCNPJCPF.Text) & " - " & Trim(NOME_A) & " ; Equipamento: " & Trim(TabOS.Fields("equipamento_id").Value) & _
         " ; Situação: " & SqL2)

         PEDIDO_ID_N = TabOS.Fields("OS_ID").Value

         lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlack

         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""

         '==============SERVIÇO
         SQL = "select * from OSSERVICO "
         SQL = SQL & " where os_id = " & TabOS.Fields("os_id").Value

         If Trim(cmbMecanicoAUX.Text) <> "" Then _
            If IsNumeric(cmbMecanicoAUX.Text) Then _
               SQL = SQL & " and mecanico_id = " & cmbMecanicoAUX.Text

         If Trim(cmbServico.Text) <> "" Then _
            SQL = SQL & " and descricao_serviço like " & cmbServicoAUX.Text & "%"

         TabServico.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabServico.EOF Then
            cmblstOSAUX.AddItem TabServico.Fields("OS_ID")
            cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

            Set Nodx = lstOS.Nodes.Add(SQL3 & QTD_COTAS, tvwChild, "Tarefa" & CONT_N, "Servico(s) = R$ " & Format(VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos))

            lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlue
         End If

         VALOR_TOTAL_SERVICO_N = 0

         While Not TabServico.EOF
            'TAREFAS
            NOME_A = "" & Trim(TabServico.Fields("descricao").Value)
            DtFecha_D = 0
            If Not IsNull(TabServico.Fields("dt_FIM").Value) Then _
               DtFecha_D = TabServico.Fields("dt_FIM").Value

            Set Nodx = lstOS.Nodes.Add("Tarefa" & CONT_N, tvwChild, , _
            Trim(TabServico.Fields("osservico_id").Value) _
            & " -  " & Trim(NOME_A) & " ; Valor = " & Format(TabServico.Fields("valor_servico").Value, strFormatacao2Digitos) & _
            " ; Desconto = " & Format(TabServico.Fields("desconto_servico").Value, strFormatacao2Digitos) & " ; Total = " & _
            Format(TabServico.Fields("valor_servico").Value - TabServico.Fields("desconto_servico").Value, strFormatacao2Digitos) & _
            " ; DtEncerra = " & DtFecha_D _
            )

            cmblstOSAUX.AddItem TabServico.Fields("OSSERVICO_ID")
            cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

            VALOR_TOTAL_SERVICO_N = VALOR_TOTAL_SERVICO_N + _
                                    TabServico.Fields("valor_servico").Value - _
                                    TabServico.Fields("desconto_servico").Value

            lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbRed
            If Not IsNull(TabServico.Fields("DT_INICIO").Value) Then _
               lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlue

            TabServico.MoveNext
         Wend
         If TabServico.State = 1 Then _
            TabServico.Close

         '====================TOTAL PEÇAS

         SQL = "select OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
         SQL = SQL & " from OSPECA "
         SQL = SQL & " INNER JOIN PRODUTO "
         SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID "
         SQL = SQL & " where os_id = " & TabOS.Fields("os_id").Value

         If Trim(cmbProdutoAUX.Text) <> "" Then _
            SQL = SQL & " and produto_id = " & cmbProdutoAUX

         TabPeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPeca.EOF Then
            cmblstOSAUX.AddItem TabOS.Fields("OS_ID")
            cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

            Set Nodx = lstOS.Nodes.Add(SQL3 & QTD_COTAS, tvwChild, "Produtos" & CONT_N, "Produto(s) = R$ " & Format(VALOR_TOTAL_PRODUTO_N, strFormatacao2Digitos))

            lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlue
         End If
   
         VALOR_TOTAL_PRODUTO_N = 0

         While Not TabPeca.EOF
            'PEÇAS
            NOME_A = "" & Trim(TabPeca.Fields("descricao").Value)
   
            Set Nodx = lstOS.Nodes.Add("Produtos" & CONT_N, tvwChild, , Trim(TabPeca.Fields("produto_id").Value) _
            & " -  " & Trim(NOME_A) & " ; Valor = " & Format(TabPeca.Fields("VALOR_ITEM").Value * TabPeca!QTDE, strFormatacao2Digitos) & _
            " ; Desconto = " & Format(TabPeca.Fields("desconto_produto").Value, strFormatacao2Digitos) & " ; Total = " & _
            Format((TabPeca.Fields("VALOR_ITEM").Value - TabPeca!DESCONTO_PRODUTO) * TabPeca!QTDE, strFormatacao2Digitos))

            cmblstOSAUX.AddItem TabPeca.Fields("OSpeca_ID")
            cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

            VALOR_TOTAL_PRODUTO_N = (TabPeca.Fields("VALOR_ITEM").Value - TabPeca!DESCONTO_PRODUTO) * TabPeca!QTDE

            TabPeca.MoveNext
         Wend
         If TabPeca.State = 1 Then _
            TabPeca.Close

         'QTD_COTAS = QTD_COTAS + 1
         CONT_N = CONT_N + 1

      End If
'=========================
      TabOS.MoveNext
   Wend
   If TabOS.State = 1 Then _
      TabOS.Close

   lstOS.Refresh
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_Eqp"
End Sub

Private Sub LIMPA_CONSULTA()
   PESSOA_ID_N = 0
   txtEqp.Text = "   "
   cmbProduto.Text = ""
   cmbProdutoAUX.Text = ""
   lstOS.Nodes.Clear
   txtCHASSI.Text = ""
   cmbTipoOS.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   cmbConsultor.Text = ""
   cmbConsultorAUX.Text = ""
   cmbSituacao.Text = ""
   cmbSituacaoAUX.Text = ""
   cmbMecanico.Text = ""
   cmbMecanicoAUX.Text = ""
   cmbServico.Text = ""
   cmbServicoAUX.Text = ""
   optDtAbertura.Value = False
   optDtFechamento.Value = False

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False

   txtDtIni.Text = ""
   txtDtFim.Text = ""
   txtPlaca.Text = ""
End Sub

Sub CARREGA_COMBOS()
'On Error GoTo ERRO_TRATA

'parametros combos x tabela descr
'8 = consultor tecnico
'9 = mecanico

   cmbTipoOSAUX.Clear
   cmbTipoOS.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'H' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipoOS.AddItem Trim(TabDESCR!DESCRICAO)
      cmbTipoOSAUX.AddItem TabDESCR!Codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbSituacao.Clear
   cmbSituacaoAUX.Clear

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'Z' "
   SQL = SQL & "order by descricao"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbSituacao.AddItem Trim(TabDESCR!DESCRICAO)
      cmbSituacaoAUX.AddItem Trim(TabDESCR.Fields("CODIGO").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

cmbSituacao.Text = "Aberta"
cmbSituacaoAUX.Text = 1

   cmbConsultorAUX.Clear
   cmbConsultor.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 8 "   'consultor tecnico
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbConsultorAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbConsultor.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbMecanicoAUX.Clear
   cmbMecanico.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 9 "   'mecanico
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbMecanicoAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbMecanico.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   'cmbVendedorAUX.Clear
   'cmbVendedor.Clear

   'If TabDESCR.State = 1 Then _
      TabDESCR.Close

   'SQL = "select vendedor_id, descricao from vwVendedor "
   'SQL = SQL & " where status = 'A' "   'vendedor
   'TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'While Not TabDESCR.EOF

   '   cmbVendedorAUX.AddItem TabDESCR.Fields("vendedor_id").Value
   '   cmbVendedor.AddItem Trim(TabDESCR.Fields("descricao").Value) & "-" & Trim(TabDESCR.Fields("vendedor_id").Value)

   '   TabDESCR.MoveNext
   'Wend


   cmbProduto.Clear
   cmbProdutoAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select OSPECA.PRODUTO_ID, PRODUTO.Descricao "
   SQL = SQL & " from OS "
   SQL = SQL & " INNER JOIN OSPECA "
   SQL = SQL & " ON OS.OS_ID = OSPECA.OS_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbProdutoAUX.AddItem TabDESCR.Fields("PRODUTO_id").Value
      cmbProduto.AddItem Trim(TabDESCR.Fields("DESCRICAO").Value) & "-" & Trim(TabDESCR.Fields("PRODUTO_id").Value)

      TabDESCR.MoveNext
   Wend

   cmbServico.Clear
   cmbServicoAUX.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select OSSERVICO.DESCRICAO, OSSERVICO.OSSERVICO_ID"
   SQL = SQL & " from OS "
   SQL = SQL & " INNER JOIN OSSERVICO "
   SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbServicoAUX.AddItem TabDESCR.Fields("OSSERVICO_id").Value
      cmbServico.AddItem Trim(TabDESCR.Fields("DESCRICAO").Value) & "-" & Trim(TabDESCR.Fields("OSSERVICO_id").Value)

      TabDESCR.MoveNext
   Wend

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBOS"
End Sub

Private Sub SETA_GRID_EQPold()
'On Error GoTo ERRO_TRATA

   Dim VACA_VEIA

   cmblstOSAUX.Clear
   lstOS.Nodes.Clear

   CONT_N = 0
   VACA_VEIA = "Endereço"
   QTD_COTAS = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOSServico "

   SQL = SQL & " where equipamento_id > 0 "

   If Trim(txtEqp.Text) <> "" Then _
      SQL = SQL & " and equipamento_id = " & txtEqp.Text

   If Trim(txtOs.Text) <> "" Then _
      If IsNumeric(txtOs.Text) Then _
         SQL = SQL & " and os_id = " & txtOs.Text

   If Trim(txtCHASSI.Text) <> "" Then _
      SQL = SQL & " and chassi = '" & Trim(txtCHASSI.Text) & "'"

   If Trim(cmbTipoOSAUX.Text) <> "" Then _
      If IsNumeric(cmbTipoOSAUX.Text) Then _
         SQL = SQL & " and tipo_os = " & Trim(cmbTipoOSAUX.Text)

   txtCNPJCPF.PromptInclude = False
   If PESSOA_ID_N > 0 Then _
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N

   If Trim(cmbSituacao.Text) <> "" Then _
      SQL = SQL & " and situacao_os = " & cmbSituacaoAUX.Text

   If Trim(cmbConsultorAUX.Text) <> "" Then _
      If IsNumeric(cmbConsultorAUX.Text) Then _
         SQL = SQL & " and ct_id = " & cmbConsultorAUX.Text

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optDtAbertura.Value = True Then
         SQL = SQL & " and dt_os >= '" & Trim(txtDtIni.Text) & "'"
         SQL = SQL & " and dt_os <= '" & Trim(txtDtFim.Text) & "'"
      End If
      If optDtFechamento.Value = True Then
         SQL = SQL & " and dt_fecha >= '" & Trim(txtDtIni.Text) & "'"
         SQL = SQL & " and dt_fecha <= '" & Trim(txtDtFim.Text) & "'"
      End If
   End If

   SQL = SQL & " order by OS_ID desc"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If TabTemp.EOF Then _
      MsgBox "Não existe O.S. para essa pesquisa."

   While Not TabTemp.EOF
'=====================================
      'totais serviço para cabeça do grid
      VALOR_TOTAL_SERVICO_N = 0
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum(valor_servico-desconto_servico) from OSSERVICO "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         VALOR_TOTAL_SERVICO_N = 0 & TabConsulta.Fields(0).Value

      'totais produto para cabeça do grid
      VALOR_TOTAL_PRODUTO_N = 0
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum((valor_item-desconto_produto) * qtde) from OSPECA "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         VALOR_TOTAL_PRODUTO_N = 0 & TabConsulta.Fields(0).Value
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

'====================================
      txtCNPJCPF.PromptInclude = False
      'NOME_A = "" & Trim(TabTemp.Fields("cliente").Value)

      SQL = "select descricao,cnpjcpf from PESSOA "
      SQL = SQL & " where pessoa_id = " & TabTemp.Fields("PESSOA_ID").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtCNPJCPF.Text = "" & Trim(TabConsulta.Fields("cnpjcpf").Value)
         NOME_A = "" & Trim(TabConsulta.Fields("descricao").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'txtCNPJCPF.Text = TabTemp.Fields("cnpjcpf").Value
      If Len(Trim(TabTemp.Fields("cnpjcpf").Value)) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If
      txtCNPJCPF.PromptInclude = True
'======================

      cmblstOSAUX.AddItem TabTemp.Fields("OS_ID").Value
      cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

      If PEDIDO_ID_N <> TabTemp.Fields("OS_ID").Value Then
         QTD_COTAS = QTD_COTAS + 1
         Set Nodx = lstOS.Nodes.Add(, , VACA_VEIA & QTD_COTAS, "O.S. = " & Trim(TabTemp.Fields("OS_ID").Value) & _
         " ; Total O.S. = " & Format(VALOR_TOTAL_PRODUTO_N + VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos) & _
         " ; Cliente: " & Trim(txtCNPJCPF.Text) & " - " & Trim(NOME_A) & " ; Equipamento: " & Trim(TabTemp.Fields("equipamento_id").Value))

         PEDIDO_ID_N = TabTemp.Fields("OS_ID").Value
      End If

      lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlack

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""

      'SERVIÇO
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from OSSERVICO "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value

      If Trim(cmbMecanicoAUX.Text) <> "" Then _
         If IsNumeric(cmbMecanicoAUX.Text) Then _
            SQL = SQL & " and mecanico_id = " & cmbMecanicoAUX.Text

      If Trim(cmbServico.Text) <> "" Then _
         SQL = SQL & " and descricao_serviço like " & cmbServicoAUX.Text & "%"

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         cmblstOSAUX.AddItem TabConsulta.Fields("OS_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         Set Nodx = lstOS.Nodes.Add(VACA_VEIA & QTD_COTAS, tvwChild, "Tarefa" & CONT_N, "Servico(s) = R$ " & Format(VALOR_TOTAL_SERVICO_N, strFormatacao2Digitos))

         lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbBlue
      End If

      VALOR_TOTAL_SERVICO_N = 0

      While Not TabConsulta.EOF
         'TAREFAS
         NOME_A = "" & Trim(TabConsulta.Fields("descricao").Value)

         Set Nodx = lstOS.Nodes.Add("Tarefa" & CONT_N, tvwChild, , Trim(TabConsulta.Fields("osservico_id").Value) _
         & " -  " & Trim(NOME_A) & " ; Valor = " & Format(TabConsulta.Fields("valor_servico").Value, strFormatacao2Digitos) & _
         " ; Desconto = " & Format(TabConsulta.Fields("desconto_servico").Value, strFormatacao2Digitos) & " ; Total = " & _
         Format(TabConsulta.Fields("valor_servico").Value - TabConsulta.Fields("desconto_servico").Value, strFormatacao2Digitos))

         cmblstOSAUX.AddItem TabConsulta.Fields("OSSERVICO_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         VALOR_TOTAL_SERVICO_N = VALOR_TOTAL_SERVICO_N + _
                                 TabConsulta.Fields("valor_servico").Value - _
                                 TabConsulta.Fields("desconto_servico").Value

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TOTAL PEÇAS
      SQL = "select OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
      SQL = SQL & " from OSPECA "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " where os_id = " & TabTemp.Fields("os_id").Value

      If Trim(cmbProdutoAUX.Text) <> "" Then _
         SQL = SQL & " and produto_id = " & cmbProdutoAUX

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         cmblstOSAUX.AddItem TabTemp.Fields("OS_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         Set Nodx = lstOS.Nodes.Add(VACA_VEIA & QTD_COTAS, tvwChild, "Produtos" & CONT_N, "Produto(s) = R$ " & Format(VALOR_TOTAL_PRODUTO_N, strFormatacao2Digitos))

         lstOS.Nodes(lstOS.Nodes.Count).ForeColor = vbRed
      End If

      VALOR_TOTAL_PRODUTO_N = 0

      While Not TabConsulta.EOF
         'PEÇAS
         NOME_A = "" & Trim(TabConsulta.Fields("descricao").Value)

         Set Nodx = lstOS.Nodes.Add("Produtos" & CONT_N, tvwChild, , Trim(TabConsulta.Fields("produto_id").Value) _
         & " -  " & Trim(NOME_A) & " ; Valor = " & Format(TabConsulta.Fields("VALOR_ITEM").Value * TabConsulta!QTDE, strFormatacao2Digitos) & _
         " ; Desconto = " & Format(TabConsulta.Fields("desconto_produto").Value, strFormatacao2Digitos) & " ; Total = " & _
         Format((TabConsulta.Fields("VALOR_ITEM").Value - TabConsulta!DESCONTO_PRODUTO) * TabConsulta!QTDE, strFormatacao2Digitos))

         cmblstOSAUX.AddItem TabConsulta.Fields("OSpeca_ID")
         cmblstOSAUX.ItemData(cmblstOSAUX.ListCount - 1) = lstOS.Nodes.Count

         VALOR_TOTAL_PRODUTO_N = (TabConsulta.Fields("VALOR_ITEM").Value - TabConsulta!DESCONTO_PRODUTO) * TabConsulta!QTDE

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'QTD_COTAS = QTD_COTAS + 1
      CONT_N = CONT_N + 1

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstOS.Refresh
   PESSOA_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_Eqp"
End Sub

Function EQP_VEICULO() As Boolean
'On Error GoTo ERRO_TRATA

   EQP_VEICULO = True
   If optEqp.Value = True Then
      txtPlaca.Visible = False
      txtEqp.Visible = True
      lbltipo.Caption = "Eqp: "
      Else  'AQUI É VEÍCULO
         EQP_VEICULO = False
         txtPlaca.Visible = True
         txtEqp.Visible = False
         lbltipo.Caption = "Placa: "
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EQP_VEICULO"
End Function

