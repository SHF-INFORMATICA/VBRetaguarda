VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCHEQUECADASTRO 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Cheque"
   ClientHeight    =   7680
   ClientLeft      =   3855
   ClientTop       =   2355
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CHEQUECADASTRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   7740
   Begin VB.TextBox txtCMC7 
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   800
      Width           =   6255
   End
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
      Height          =   4455
      Left            =   50
      TabIndex        =   17
      Top             =   3150
      Width           =   7635
      Begin VB.CommandButton cmdRepasse 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2640
         Picture         =   "CHEQUECADASTRO.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3960
         Width           =   405
      End
      Begin VB.TextBox txtRepasse 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3120
         MaxLength       =   100
         TabIndex        =   15
         Top             =   3960
         Width           =   4455
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3120
         MaxLength       =   100
         TabIndex        =   13
         Top             =   3360
         Width           =   4455
      End
      Begin VB.CommandButton cmdCli 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2640
         Picture         =   "CHEQUECADASTRO.frx":6614
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3360
         Width           =   405
      End
      Begin VB.ComboBox cmbContaAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   6840
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbAgenciaAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   6840
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbBancoAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   6840
         TabIndex        =   37
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbConta 
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox cmbAgencia 
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
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox cmbBanco 
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
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   2895
      End
      Begin MSMask.MaskEdBox txtPORTADOR 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   192
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
      Begin MSComctlLib.Toolbar BarConta 
         Height          =   330
         Index           =   0
         Left            =   3135
         TabIndex        =   29
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ILBT16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "abrir"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "matar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar BarConta 
         Height          =   330
         Index           =   1
         Left            =   3135
         TabIndex        =   30
         Top             =   1230
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ILBT16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "abrir"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "matar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar BarConta 
         Height          =   330
         Index           =   2
         Left            =   3135
         TabIndex        =   31
         Top             =   510
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ILBT16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "abrir"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "matar"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txtPROP 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   12582912
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
      Begin MSComctlLib.ImageList ILBT16 
         Left            =   6600
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
               Picture         =   "CHEQUECADASTRO.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":BA0A
               Key             =   "verde"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":BFB2
               Key             =   "amarelo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":C55A
               Key             =   "vermelho"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":CB02
               Key             =   "azul"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":D0AA
               Key             =   "cinza"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":D64E
               Key             =   "preto"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCNPJCPF_REPASSE 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   192
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repasse:"
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
         TabIndex        =   48
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lblProp 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
         TabIndex        =   35
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblAgencia 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblConta 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3480
         TabIndex        =   23
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblBanco 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável/Portador:"
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
         TabIndex        =   18
         Top             =   3120
         Width           =   2145
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":DBF2
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":E046
            Key             =   "relampago"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":E362
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":E7B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":EC0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":EF2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":F37E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":F69E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":FAF2
            Key             =   "fechar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECADASTRO.frx":11286
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame framedados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   60
      TabIndex        =   24
      Top             =   1275
      Width           =   7635
      Begin VB.TextBox txtPraça 
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
         Left            =   6120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2090
         Picture         =   "CHEQUECADASTRO.frx":12F92
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtPedido 
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
         Left            =   3480
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtSERIE 
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
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCheque 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtValor 
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtDtDep 
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSMask.MaskEdBox txtDtEmis 
         Height          =   375
         Left            =   6120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   192
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
      Begin MSMask.MaskEdBox txtDtCompensa 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   192
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
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Praça:"
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
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   2640
         TabIndex        =   42
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
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
         Left            =   2760
         TabIndex        =   41
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
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
         TabIndex        =   40
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série:"
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
         TabIndex        =   36
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cheque"
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
         TabIndex        =   28
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbldtEmis 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Cadastro"
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
         Left            =   6270
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbldtDep 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Depósito:"
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
         Left            =   4740
         TabIndex        =   26
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   4560
         TabIndex        =   25
         Top             =   240
         Width           =   510
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList3"
      DisabledImageList=   "ImageList3"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Imprimir"
            Key             =   "imp"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   6360
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
               Picture         =   "CHEQUECADASTRO.frx":13994
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":14B2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":15BBD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":16E25
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECADASTRO.frx":18057
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CMC7:"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   240
      TabIndex        =   45
      Top             =   840
      Width           =   780
   End
End
Attribute VB_Name = "frmCHEQUECADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim PESSOA_REPASSE_ID_N    As Long
   Dim PESSOA_PORTADOR_ID_N   As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmCHEQUECADASTRO)

   LIMPA_TUDO

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "imp"
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
      Case "gravar"
         GRAVA_CHEQUE
         LIMPA_TUDO
         If INDR_PRI = True Then _
            Unload Me
      Case "print"
      Case "matar"
         If Trim(lblID.Caption) <> "" Then
            If IsNumeric(lblID.Caption) Then
               SQL = "delete from CHEQUE "
               SQL = SQL & " where cheque_id = " & lblID.Caption
               CONECTA_RETAGUARDA.Execute SQL
               LIMPA_TUDO
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   frmCHEQUECONSULTA.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_click"
End Sub

Private Sub txtCMC7_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCMC7.Text) <> "" Then
         LER_CMC7

         MOSTRA_CHEQUE

         txtValor.SetFocus
         Else: txtCheque.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCHEQUE_KeyPress"
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCHEQUE_KeyPress"
End Sub

Private Sub txtCNPJCPF_REPASSE_GotFocus()
   txtCNPJCPF_REPASSE.PromptInclude = False
      If Trim(txtCNPJCPF_REPASSE.Text) = "" Then _
         txtCNPJCPF_REPASSE.Text = "99999999999"
   txtCNPJCPF_REPASSE.PromptInclude = True
End Sub

Private Sub txtCNPJCPF_REPASSE_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF_REPASSE.PromptInclude = False
               txtCNPJCPF_REPASSE.Text = CNPJCPF_A
            txtCNPJCPF_REPASSE.PromptInclude = True
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_REPASSE_KeyDown"
End Sub

Private Sub txtCNPJCPF_REPASSE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCNPJCPF_REPASSE.PromptInclude = False

      If Trim(txtCNPJCPF_REPASSE.Text) <> "" Then _
         txtRepasse.Text = PROCURA_REPASSE(Trim(txtCNPJCPF_REPASSE.Text))

      txtCNPJCPF_REPASSE.PromptInclude = True
      txtRepasse.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_REPASSE_KeyPress"
End Sub

Private Sub txtCNPJCPF_REPASSE_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_REPASSE.PromptInclude = False

   If Trim(txtCNPJCPF_REPASSE.Text) <> "" Then _
      txtRepasse.Text = PROCURA_REPASSE(Trim(txtCNPJCPF_REPASSE.Text))

   txtCNPJCPF_REPASSE.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrepasse_LostFocus"
End Sub

Private Sub txtDtCompensa_GotFocus()
'On Error GoTo ERRO_TRATA

   'txtDtCompensa.PromptInclude = False
   'If Trim(txtDtCompensa.Text) = "" Then _
      txtDtCompensa.Text = Date
   'txtDtCompensa.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtCompensa_GotFocus"
End Sub

Private Sub txtDtCompensa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtDep.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtCompensa_KeyPress"
End Sub

Private Sub txtDtDep_GotFocus()
'On Error GoTo ERRO_TRATA

   'txtDtDep.PromptInclude = False
   'If Trim(txtDtDep.Text) = "" Then _
      txtDtDep.Text = Date
   'txtDtDep.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtDep_GotFocus"
End Sub

Private Sub txtDtDep_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPORTADOR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtDep_KeyPress"
End Sub

Private Sub txtDtEmis_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEmis.PromptInclude = False
   If Trim(txtDtEmis.Text) = "" Then _
      txtDtEmis.Text = Date
   txtDtEmis.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_GotFocus"
End Sub

Private Sub txtDTEMIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSerie.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEmis_KeyPress"
End Sub

Private Sub txtPORTADOR_GotFocus()
   txtPORTADOR.PromptInclude = False
      If Trim(txtPORTADOR.Text) = "" Then _
         txtPORTADOR.Text = "99999999999"
   txtPORTADOR.PromptInclude = True
End Sub

Private Sub txtPORTADOR_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtNome.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPORTADOR_KeyPress"
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF_REPASSE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnome_KeyPress"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNome.Text) <> "" Then
      txtNome.SelStart = 0
      txtNome.SelLength = Len(txtNome)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtrepasse_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtRepasse.Text) <> "" Then
      txtRepasse.SelStart = 0
      txtRepasse.SelLength = Len(txtRepasse)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrepasse_GotFocus"
End Sub

Private Sub cmdCli_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtPORTADOR.PromptInclude = False
      txtPORTADOR.Text = CNPJCPF_A

      If Trim(txtPORTADOR.Text) <> "" Then _
         txtNome.Text = "" & PROCURA_PORTADOR(Trim(txtPORTADOR.Text))

   End If
   CNPJCPF_A = ""
   txtPORTADOR.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCli_Click"
End Sub

Private Sub txtPORTADOR_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtPORTADOR.PromptInclude = False
               txtPORTADOR.Text = CNPJCPF_A
            txtPORTADOR.PromptInclude = True
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPORTADOR_KeyDown"
End Sub

Private Sub txtPORTADOR_LostFocus()
'On Error GoTo ERRO_TRATA

   txtPORTADOR.PromptInclude = False

   If Trim(txtPORTADOR.Text) <> "" Then _
      txtNome.Text = "" & PROCURA_PORTADOR(Trim(txtPORTADOR.Text))

   txtPORTADOR.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPORTADOR_LostFocus"
End Sub

Private Sub cmdRepasse_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF_REPASSE.PromptInclude = False
      txtCNPJCPF_REPASSE.Text = CNPJCPF_A

      If Trim(txtCNPJCPF_REPASSE.Text) <> "" Then _
         txtRepasse.Text = PROCURA_REPASSE(Trim(txtCNPJCPF_REPASSE.Text))

   End If
   CNPJCPF_A = ""
   txtCNPJCPF_REPASSE.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCli_Click"
End Sub

Private Sub txtPROP_GotFocus()
   SendKeys "{tab}"
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPedido.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSERIE_KeyPress"
End Sub

Private Sub txtpedido_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPraça.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPEDIDO_KeyPress"
End Sub

Private Sub txtpraça_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtCompensa.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpraça_KeyPress"
End Sub

Private Sub txtSERIE_LostFocus()
   'VALOR_ITEM_N = 0 & txtSERIE.Text
   'txtSERIE.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtEmis.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_KeyPress"
End Sub

Private Sub txtTXTDTEMIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTTXTDTEMIS_KeyPress"
End Sub

Private Sub txtTXTDTdep_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTTXTDTdep_KeyPress"
End Sub

Private Sub txtTXTPORTADOR_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTTXTPORTADOR_KeyPress"
End Sub

Private Sub cmbbanco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      'SendKeys "{tab}"
      cmbAgencia.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CMBBANCO_KeyPress"
End Sub

Private Sub cmbBanco_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbBanco.Text) = "" Then _
      CARREGA_COMBO_BANCO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_GotFocus"
End Sub

Private Sub cmbBanco_Click()
'On Error GoTo ERRO_TRATA

   If Trim(cmbBanco.Text) <> "" Then
      cmbBancoAux.ListIndex = cmbBanco.ListIndex

      If Trim(cmbBancoAux.Text) <> "" Then _
         MOSTRA_BANCO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_Click"
End Sub

Private Sub cmbBanco_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbBanco.Text) <> "" Then _
      MOSTRA_BANCO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_LostFocus"
End Sub

Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbConta.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAgencia_KeyPress"
End Sub

Private Sub cmbAgencia_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbAgencia.Text) = "" Then _
      CARREGA_COMBO_AGENCIA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAgencia_GotFocus"
End Sub

Private Sub cmbAgencia_Click()
'On Error GoTo ERRO_TRATA

   If Trim(cmbAgencia.Text) <> "" Then
      cmbAgenciaAUX.ListIndex = cmbAgencia.ListIndex

      If Trim(cmbAgenciaAUX.Text) <> "" Then _
         MOSTRA_AGENCIA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAGENCIA_Click"
End Sub

Private Sub cmbAgencia_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbAgencia.Text) <> "" Then _
      MOSTRA_AGENCIA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAgencia_LostFocus"
End Sub

Private Sub cmbConta_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPORTADOR.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CMBCONTA_KeyPress"
End Sub

Private Sub cmbConta_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbConta.Text) = "" Then _
      CARREGA_COMBO_CONTA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbConta_GotFocus"
End Sub

Private Sub cmbConta_Click()
'On Error GoTo ERRO_TRATA

   If Trim(cmbConta.Text) <> "" Then
      cmbContaAUX.ListIndex = cmbConta.ListIndex

      If Trim(cmbContaAUX.Text) <> "" Then _
         MOSTRA_CONTA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCONTA_Click"
End Sub

Private Sub cmbCONTA_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbConta.Text) <> "" Then _
      MOSTRA_CONTA

   If Trim(cmbConta.Text) <> "" And Trim(txtCheque.Text) <> "" Then _
      MOSTRA_CHEQUE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCONTA_LostFocus"
End Sub

Private Sub txtTXTPROP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTTXTPROP_KeyPress"
End Sub
'===================================
Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_REPASSE.PromptInclude = False
   txtCNPJCPF_REPASSE.Text = ""
   txtRepasse.Text = ""
   txtCMC7.Text = ""
   txtPraça.Text = ""
   lblID.Caption = ""
   txtPedido.Text = ""
   txtCheque.Text = ""
   txtSerie.Text = ""
   txtValor.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtEmis.Text = ""
   txtDtDep.PromptInclude = False
   txtDtDep.Text = ""

   txtPORTADOR.Text = ""
   txtNome.Text = ""
   cmbBanco.Text = ""
   cmbBancoAux.Text = ""
   lblBanco.Caption = ""
   cmbAgencia.Text = ""
   cmbAgenciaAUX.Text = ""
   lblAgencia.Caption = ""
   cmbConta.Text = ""
   cmbContaAUX.Text = ""
   lblConta.Caption = ""
   txtProp.Text = ""
   lblProp.Caption = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub CARREGA_COMBO_BANCO()
'On Error GoTo ERRO_TRATA

   cmbBanco.Clear
   cmbBancoAux.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from Banco "
   SQL = SQL & " order by nome_banco"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbBanco.AddItem Trim(TabTemp.Fields("codg_banco").Value)
      cmbBancoAux.AddItem TabTemp.Fields("banco_id").Value
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO_BANCO"
End Sub

Sub CARREGA_COMBO_AGENCIA()
'On Error GoTo ERRO_TRATA

   cmbAgencia.Clear
   cmbAgenciaAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbBancoAux.Text) <> "" Then
      SQL = "select * from AGENCIA "
      SQL = SQL & " where codg_banco  = '" & Trim(cmbBancoAux.Text) & "'"
      SQL = SQL & " order by nome_AGENCIA"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         cmbAgencia.AddItem Trim(TabTemp.Fields("numr_Agencia").Value)
         cmbAgenciaAUX.AddItem Trim(TabTemp.Fields("agencia_id").Value)
         TabTemp.MoveNext
      Wend
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO_AGENCIA"
End Sub

Sub CARREGA_COMBO_CONTA()
'On Error GoTo ERRO_TRATA

   cmbConta.Clear
   cmbContaAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select CONTA.CONTA_ID, CONTA.AGENCIA_ID, CONTA.PESSOA_ID, "
   SQL = SQL & " CONTA.NUMR_CONTA, CONTA.DESC_CONTA, CONTA.DT_Cadastro"
   SQL = SQL & " from BANCO "
   SQL = SQL & " INNER JOIN AGENCIA "
   SQL = SQL & " ON BANCO.BANCO_ID = AGENCIA.BANCO_ID "
   SQL = SQL & " RIGHT OUTER JOIN CONTA "
   SQL = SQL & " ON AGENCIA.AGENCIA_ID = CONTA.AGENCIA_ID"

   SQL = SQL & " where agencia.banco_id = '" & Trim(cmbBancoAux.Text) & "'"
   SQL = SQL & " and conta.agencia_id  = '" & Trim(cmbAgenciaAUX.Text) & "'"
   SQL = SQL & " order by numr_CONTA"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbConta.AddItem Trim(TabTemp.Fields("numr_CONTA").Value)
      cmbContaAUX.AddItem Trim(TabTemp.Fields("CONTA_id").Value)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO_CONTA"
End Sub

Sub MOSTRA_BANCO()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbBanco.Text) <> "" Then
      SQL = "select * from Banco "

      If IsNumeric(cmbBanco.Text) Then
         SQL = SQL & " where codg_banco = '" & Trim(cmbBanco.Text) & "'"
         Else: SQL = SQL & " where nome_banco = '" & Trim(cmbBanco.Text) & "'"
      End If

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         lblBanco.Caption = Trim(TabTemp!Nome_Banco) & " - " & TabTemp.Fields("CODG_banco").Value
         lblBanco.Refresh
         cmbBanco.Text = Trim(TabTemp.Fields("codg_banco").Value)
         cmbBanco.Refresh
         cmbBancoAux.Text = Trim(TabTemp.Fields("banco_ID").Value)
         cmbBancoAux.Refresh
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_BANCO"
End Sub

Sub MOSTRA_AGENCIA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbAgencia.Text) <> "" Then
      SQL = "select * from AGENCIA "

      If IsNumeric(cmbAgencia.Text) Then
         SQL = SQL & " where numr_AGENCIA = '" & Trim(cmbAgencia.Text) & "'"
         Else: SQL = SQL & " where nome_agencia = '" & Trim(cmbAgencia.Text) & "'"
      End If

SQL = SQL & " and banco_id = " & Trim(cmbBancoAux.Text)

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         lblAgencia.Caption = Trim(TabTemp!nome_agencia)
         cmbAgencia.Text = Trim(TabTemp.Fields("numr_AGENCIA").Value)
         cmbAgenciaAUX.Text = TabTemp.Fields("AGENCIA_id").Value
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_AGENCIA"
End Sub

Sub MOSTRA_CONTA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbConta.Text) <> "" Then
      SQL = "select * from CONTA "
      SQL = SQL & " where agencia_ID  = '" & Trim(cmbAgenciaAUX.Text) & "'"
      SQL = SQL & " and numr_CONTA = '" & Trim(cmbConta.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from CONTA "
         SQL = SQL & " where agencia_ID  = '" & Trim(cmbAgenciaAUX.Text) & "'"
         SQL = SQL & " and desc_conta = '" & Trim(cmbConta.Text) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      End If
      If Not TabTemp.EOF Then
         lblConta.Caption = Trim(TabTemp.Fields("desc_CONTA").Value) & " - " & TabTemp.Fields("numr_CONTA").Value
         cmbConta.Text = Trim(TabTemp.Fields("numr_CONTA").Value)
         cmbContaAUX.Text = TabTemp.Fields("CONTA_id").Value

         If Not IsNull(TabTemp.Fields("pessoa_id").Value) Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from PESSOA "
            SQL = SQL & " where pessoa_id = " & TabTemp.Fields("pessoa_id").Value
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then

               txtProp.PromptInclude = False
               txtProp.Text = Trim(TabConsulta.Fields("CNPJCPF").Value)
               txtProp.PromptInclude = True
               lblProp.Caption = Trim(TabConsulta.Fields("DESCRICAO").Value)

            End If

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CONTA"
End Sub

Function PROCURA_PORTADOR(CNPJCPF_A As String) As String
'On Error GoTo ERRO_TRATA

   PROCURA_PORTADOR = ""
   PESSOA_PORTADOR_ID_N = 0

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select * from PESSOA"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then
      PROCURA_PORTADOR = "" & TabPessoa.Fields("descricao").Value
      PESSOA_PORTADOR_ID_N = TabPessoa.Fields("pessoa_id").Value
      Else
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         MsgBox "CNPJ/CPF não encontrado"
         Exit Function
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PORTADOR"
End Function

Function PROCURA_REPASSE(CNPJCPF_A As String) As String
'On Error GoTo ERRO_TRATA

   PROCURA_REPASSE = ""
   PESSOA_REPASSE_ID_N = 0

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select descricao,pessoa_id from PESSOA"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then
      PROCURA_REPASSE = "" & Trim(TabPessoa.Fields("descricao").Value)
      PESSOA_REPASSE_ID_N = TabPessoa.Fields("pessoa_id").Value
      Else
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         MsgBox "CNPJ/CPF não encontrado"
         Exit Function
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_REPASSE"
End Function

Sub GRAVA_CHEQUE()
'On Error GoTo ERRO_TRATA

   Dim SITUAÇÃO_A    As String
   Dim Dt_Emis_D     As String
   Dim Dt_Dep_D      As String
   Dim Dt_Compensa_D As String

   Dt_Emis_D = "Null"
   Dt_Dep_D = "Null"
   Dt_Compensa_D = "Null"

   If Trim(cmbContaAUX.Text) = "" Then _
      cmbContaAUX.Text = "null"

   If Not IsNumeric(cmbContaAUX.Text) Then _
      cmbContaAUX.Text = "null"

   txtDtEmis.PromptInclude = True
   If IsDate(txtDtEmis.Text) Then _
      Dt_Emis_D = txtDtEmis.Text

   txtDtDep.PromptInclude = True
   If IsDate(txtDtDep.Text) Then _
      Dt_Dep_D = DMA(txtDtDep.Text)

   txtDtCompensa.PromptInclude = True
   If IsDate(txtDtCompensa.Text) Then _
      Dt_Compensa_D = DMA(txtDtCompensa.Text)

   If Trim(txtCheque.Text) = "" Then
      MsgBox "Informe número do cheque !!!"
      txtCheque.SetFocus
      Exit Sub
   End If

   txtDtEmis.PromptInclude = False
   If Trim(txtDtEmis.Text) = "" Then
      txtDtEmis.Text = Date
      SITUAÇÃO_A = "E"  'cadastrado
   End If
   txtDtEmis.PromptInclude = True
   If IsDate(txtDtEmis.Text) Then _
      SITUAÇÃO_A = "E"  'cadastrado

   txtDtCompensa.PromptInclude = False
   If Trim(txtDtCompensa.Text) = "" Then
      txtDtCompensa.Text = Format(0, "mm,dd,yyyy")
      SITUAÇÃO_A = "P"  'processado, compensado
   End If
   txtDtCompensa.PromptInclude = True
   If IsDate(txtDtCompensa.Text) Then _
      SITUAÇÃO_A = "P"  'processado, compensado

   txtDtDep.PromptInclude = True
   If IsDate(txtDtDep.Text) Then _
      SITUAÇÃO_A = "D"  'cadastrado

   txtPORTADOR.PromptInclude = False
   If Trim(txtPORTADOR.Text) <> "" Then
      txtNome.Text = "" & PROCURA_PORTADOR(Trim(txtPORTADOR.Text))
      Else
         MsgBox "Portador não informado."
         Exit Sub
   End If

   If Trim(lblID.Caption) = "" Then
      lblID.Caption = MAX_ID("cheque_id", "cheque", "", "", "", "")
      Else
         If Not IsNumeric(lblID.Caption) Then _
            lblID.Caption = MAX_ID("cheque_id", "cheque", "", "", "", "")
   End If

GRAVA_BANCO
GRAVA_AGENCIA
GRAVA_CONTA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select  * from CHEQUE"
   SQL = SQL & " where cheque_id = " & lblID.Caption
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      SQL = "update CHEQUE set "
         SQL = SQL & " NUMR_CHEQUE = '" & Trim(txtCheque.Text) & "'"
         SQL = SQL & " ,SERIE_CHEQUE = '" & Trim(txtSerie.Text) & "'"
         SQL = SQL & " ,conta_id = " & Trim(cmbContaAUX.Text)
         SQL = SQL & " ,VALOR = " & tpMOEDA(txtValor.Text)
         SQL = SQL & " ,DT_EMISSAO = '" & DMA(Dt_Emis_D) & "'"
         If Dt_Dep_D = "Null" Then
            SQL = SQL & " ,DT_DEPOSITO = null "
            Else: SQL = SQL & " ,DT_DEPOSITO = '" & DMA(Dt_Dep_D) & "'"
         End If
         SQL = SQL & " ,DT_COMPENSA = '" & DMA(Dt_Compensa_D) & "'"
         SQL = SQL & " ,STATUS = '" & Trim(SITUAÇÃO_A) & "'"
         SQL = SQL & " ,RESP_ID = " & PESSOA_PORTADOR_ID_N
         SQL = SQL & " ,numr_doc = '" & Trim(txtPedido.Text) & "'"

         SQL = SQL & " ,CMC7 = '" & Trim(txtCMC7.Text) & "'"                     'CMC7
         SQL = SQL & " ,PRAÇA = '" & Trim(txtPraça.Text) & "'"                   'PRAÇA
         SQL = SQL & " ,responsavel = '" & Trim(txtNome.Text) & "'"              'RESPONSAVEL

         SQL = SQL & " ,repasse_ID = " & PESSOA_REPASSE_ID_N                     'repasse_id
         SQL = SQL & " ,repasse = '" & Trim(txtRepasse.Text) & "'"               'repasse

      SQL = SQL & " where cheque_id = " & TabConsulta.Fields("cheque_id").Value
      Else
         SQL = "insert into CHEQUE "
         SQL = SQL & " ("
            SQL = SQL & " CHEQUE_ID,NUMR_CHEQUE,SERIE_CHEQUE,conta_id,VALOR,"
            SQL = SQL & " DT_EMISSAO,DT_DEPOSITO,DT_COMPENSA,STATUS,RESP_ID,NUMR_DOC,"
            SQL = SQL & " CMC7,PRAÇA,RESPONSAVEL,repasse_id,repasse,estabelecimento_id"
         SQL = SQL & " )"
         SQL = SQL & " values("
            SQL = SQL & MAX_ID("cheque_id", "cheque", "", "", "", "")   'CHEQUE_ID
            SQL = SQL & " ,'" & Trim(txtCheque.Text) & "'"              'NUMR_CHEQUE
            SQL = SQL & " ,'" & Trim(txtSerie.Text) & "'"               'SERIE_CHEQUE
            SQL = SQL & " ," & Trim(cmbContaAUX.Text)                   'conta_id
            SQL = SQL & " ," & tpMOEDA(txtValor.Text)                   'VALOR
            SQL = SQL & " ,'" & DMA(Dt_Emis_D) & "'"                    'DT_EMISSAO

            If Dt_Dep_D = "Null" Then
               SQL = SQL & " ,null"                                     'DT_DEPOSITO
               Else: SQL = SQL & " ,'" & DMA(Dt_Dep_D) & "'"            'DT_DEPOSITO
            End If

            SQL = SQL & " ,'" & DMA(Dt_Compensa_D) & "'"                'DT_COMPENSA
            SQL = SQL & " ,'" & SITUAÇÃO_A & "'"                        'STATUS
            SQL = SQL & " ," & PESSOA_PORTADOR_ID_N                     'RESP_ID
            SQL = SQL & " ,'" & txtPedido.Text & "'"                    'NUMR_DOC
            SQL = SQL & " ,'" & Trim(txtCMC7.Text) & "'"                'CMC7
            SQL = SQL & " ,'" & Trim(txtPraça.Text) & "'"               'PRAÇA
            SQL = SQL & " ,'" & Trim(txtNome.Text) & "'"                'RESPONSAVEL

            SQL = SQL & " ," & PESSOA_REPASSE_ID_N                      'repasse_id
            SQL = SQL & " ,'" & Trim(txtRepasse.Text) & "'"             'repasse
            SQL = SQL & " ," & ESTABELECIMENTO_ID_N                     'estabelecimento_id

         SQL = SQL & " )"
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   CONECTA_RETAGUARDA.Execute SQL

   MsgBox "Operação realizada com sucesso."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CHEQUE"
End Sub

Sub MOSTRA_CHEQUE()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from vwRelCheque "
   SQL = SQL & " where numr_cheque = '" & Trim(txtCheque.Text) & "'"
   SQL = SQL & " and serie_cheque = '" & Trim(txtSerie.Text) & "'"
   SQL = SQL & " and ESTABELECIMENTO_ID = " & ESTABELECIMENTO_ID_N
   'SQL = SQL & " and numr_conta = '" & Trim(cmbContaAUX.Text) & "'"

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      txtPedido.Text = "" & TabConsulta.Fields("numr_doc").Value
      txtValor.Text = "" & TabConsulta.Fields("valor").Value

      txtDtEmis.PromptInclude = False
      txtDtEmis.Text = "" & TabConsulta.Fields("dt_emissao").Value
      txtDtEmis.PromptInclude = True

      txtDtDep.PromptInclude = False
      txtDtDep.Text = "" & TabConsulta.Fields("dt_deposito").Value
      txtDtDep.PromptInclude = True

      txtPORTADOR.PromptInclude = False
      txtPORTADOR.Text = "" & TabConsulta.Fields("CNPJCPF_Terc").Value

      txtNome.Text = "" & TabConsulta.Fields("nome_Terc").Value
      cmbBanco.Text = "" & Trim(TabConsulta.Fields("codg_banco").Value)
      cmbBancoAux.Text = "" & TabConsulta.Fields("BANCO_id").Value
      lblBanco.Caption = "" & TabConsulta.Fields("nome_BANCO").Value
      cmbAgencia.Text = "" & TabConsulta.Fields("numr_agencia").Value
      cmbAgenciaAUX.Text = "" & TabConsulta.Fields("agencia_id").Value
      lblAgencia.Caption = ""
      cmbConta.Text = "" & TabConsulta.Fields("numr_conta").Value
      cmbContaAUX.Text = "" & TabConsulta.Fields("conta_id").Value
      lblConta.Caption = ""
      txtProp.Text = "" & TabConsulta.Fields("CNPJCPF_Prop").Value
      lblProp.Caption = "" & TabConsulta.Fields("nome_Prop").Value

      If TabPessoa.State = 1 Then _
         TabPessoa.Close

      If Not IsNull(TabConsulta.Fields("repasse_id").Value) Then
         SQL = "select * from PESSOA "
         SQL = SQL & " where pessoa_id = " & TabConsulta.Fields("repasse_id").Value
         TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabPessoa.EOF Then
            txtCNPJCPF_REPASSE.PromptInclude = False
               txtCNPJCPF_REPASSE.Text = Trim(TabPessoa.Fields("CNPJCPF").Value)
            txtCNPJCPF_REPASSE.PromptInclude = True
            txtRepasse.Text = Trim(TabPessoa.Fields("DESCRICAO").Value)
         End If

         If TabPessoa.State = 1 Then _
            TabPessoa.Close
      End If
   End If

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CHEQUE"
End Sub

Sub LER_CMC7()
'On Error GoTo ERRO_TRATA

'Campo: 1
'Descrição: Caracter de início de leitura, tamanho: 1
'início: 1
'fim: 1
'Conteúdo: “<”

'Campo: 2
'Descrição: Código do Banco
'tamanho: 3
'início: 2
'fim: 4
'Conteúdo: Código do Banco
   SQL3 = Mid(Trim(txtCMC7.Text), 2, 3)
   cmbBanco.Text = Trim(SQL3)
   If Trim(SQL3) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from Banco "
      SQL = SQL & " where codg_banco = '" & Trim(SQL3) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         lblBanco.Caption = Trim(TabTemp!Nome_Banco) & " - " & Trim(TabTemp.Fields("CODG_banco").Value)
         cmbBanco.Text = Trim(TabTemp.Fields("codg_banco").Value)
         cmbBancoAux.Text = Trim(TabTemp.Fields("banco_ID").Value)
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

'Campo: 3
'Descricão: Agência do Cheque
'tamanho: 4
'início: 5
'fim: 8
'Conteúdo: Agência do Cheque
   SQL3 = Mid(Trim(txtCMC7.Text), 5, 4)
   cmbAgencia.Text = Trim(SQL3)
   If Trim(SQL3) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from AGENCIA "
      SQL = SQL & " where numr_AGENCIA = '" & Trim(SQL3) & "'"
      SQL = SQL & " and banco_id = " & Trim(cmbBancoAux.Text)
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         lblAgencia.Caption = Trim(TabTemp!nome_agencia) & " - " & Trim(TabTemp.Fields("numr_AGENCIA").Value)
         cmbAgencia.Text = Trim(TabTemp.Fields("numr_AGENCIA").Value)
         cmbAgenciaAUX.Text = TabTemp.Fields("agencia_id").Value
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

'Campo: 4
'Descrição: Dígito Verificador DV2
'tamanho: 1
'início: 9
'fim: 9
'Conteúdo: Módulo 10 sobre os campos 6 a 8

'Campo: 5
'Descrição: Caracter usado pelo CMC7
'tamanho: 1
'início: 10
'fim: 10
'Conteúdo: Fixo “ < ”

'Campo: 6
'Descrição: Praça de Compensação
'tamanho: 3
'início: 11
'fim: 13
'Conteúdo: Praça de Compensação
   SQL3 = Mid(Trim(txtCMC7.Text), 11, 3)
   txtPraça.Text = Trim(SQL3)

'Campo: 7
'Descrição: Número do Cheque
'tamanho: 6
'início: 14
'fim: 19
'Conteúdo: Número do Cheque
   SQL3 = Mid(Trim(txtCMC7.Text), 14, 6)
   txtCheque.Text = Trim(SQL3)

'Campo: 8
'Descrição: Campo Fixo
'tamanho: 1
'início: 20
'fim: 20
'Conteúdo: “5?

'Campo: 9
'Descrição: Caracter usado pelo CMC7
'tamanho: 1
'início: 21
'fim: 21
'Conteúdo: Fixo “ < ”

'Campo: 10
'Descrição: Dígito verificador DV1
'início: 22
'fim: 22
'Conteúdo: Módulo 10 sobre os campos 2 a 3

'Campo: 11
'Descrição: Número da conta corrente
'tamanho: 10
'início: 23
'fim: 32
'Conteúdo: Número da Conta Corrente
   SQL3 = Mid(Trim(txtCMC7.Text), 23, 10)
   cmbConta.Text = Trim(SQL3)

   If Trim(SQL3) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CONTA "
      SQL = SQL & " where agencia_ID  = '" & Trim(cmbAgenciaAUX.Text) & "'"
      SQL = SQL & " and numr_CONTA = '" & Trim(SQL3) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         lblConta.Caption = Trim(TabTemp.Fields("desc_CONTA").Value) & " - " & Trim(TabTemp.Fields("numr_CONTA").Value)
         cmbConta.Text = Trim(TabTemp.Fields("numr_CONTA").Value)
         cmbContaAUX.Text = TabTemp.Fields("conta_id").Value

         If Not IsNull(TabTemp.Fields("pessoa_id").Value) Then
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from PESSOA "
            SQL = SQL & " where pessoa_id = " & TabTemp.Fields("pessoa_id").Value
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtProp.PromptInclude = False
               txtProp.Text = Trim(TabConsulta.Fields("CNPJCPF").Value)
               txtProp.PromptInclude = True
               lblProp.Caption = Trim(TabConsulta.Fields("DESCRICAO").Value)
            End If

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

'Campo: 12
'Descrição: Dígito Verificador – DV3
'tamanho: 1
'início: 33
'fim: 33
'Conteúdo: Módulo 10 sobre os campos 12

'Campo: 13
'Descrição: Caracter de fim de leitura
'tamanho:1
'início:34
'fim:34
'Conteúdo: “: ”

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LER_CMC7"
End Sub

Sub GRAVA_CONTA()
'On Error GoTo ERRO_TRATA

   If Trim(cmbAgencia.Text) <> "" And PESSOA_PORTADOR_ID_N > 0 Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CONTA "
      SQL = SQL & " where agencia_ID  = '" & Trim(cmbAgenciaAUX.Text) & "'"
      SQL = SQL & " and numr_CONTA = '" & Trim(cmbConta.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         cmbContaAUX.Text = MAX_ID("conta_id", "conta", "", "", "", "")

         SQL = "INSERT INTO CONTA "
            SQL = SQL & " (conta_id,agencia_ID,pessoa_id,numr_conta,desc_conta,dt_cadastro)"
         SQL = SQL & " VALUES ("
            SQL = SQL & cmbContaAUX.Text                       'conta_id
            SQL = SQL & "," & Trim(cmbAgenciaAUX.Text)         'agencia_ID
            SQL = SQL & "," & PESSOA_PORTADOR_ID_N             'pessoa_id
            SQL = SQL & ",'" & Trim(cmbConta.Text) & "'"       'numr_conta
            SQL = SQL & ",'" & Trim(lblConta.Caption) & "'"    'desc_conta
            SQL = SQL & ",'" & Now & "'"                 'dt_cadastro
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
         Else: cmbContaAUX.Text = TabTemp.Fields("conta_id").Value
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CONTA"
End Sub

Sub GRAVA_AGENCIA()
'On Error GoTo ERRO_TRATA

   If Trim(cmbAgencia.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from AGENCIA "
      SQL = SQL & " where numr_AGENCIA = '" & Trim(cmbAgencia.Text) & "'"
      SQL = SQL & " and banco_id = " & Trim(cmbBancoAux.Text)
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "INSERT INTO AGENCIA"
            SQL = SQL & " (AGENCIA_id,BANCO_ID,NUMR_AGENCIA,codg_banco, Nome_AGENCIA)"
         SQL = SQL & " VALUES ("
            SQL = SQL & MAX_ID("agencia_id", "agencia", "", "", "", "") 'AGENCIA_id
            SQL = SQL & "," & Trim(cmbBancoAux.Text)                    'BANCO_ID
            SQL = SQL & ",'" & Trim(cmbAgencia.Text) & "'"              'NUMR_AGENCIA
            SQL = SQL & "," & Trim(cmbBanco.Text)                       'codg_banco
            SQL = SQL & ",'" & Trim(lblAgencia.Caption) & "'"           'Nome_AGENCIA
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_AGENCIA"
End Sub

Sub GRAVA_BANCO()
'On Error GoTo ERRO_TRATA

   If Trim(cmbBanco.Text) <> "" Then
      If TabBANCO.State = 1 Then _
         TabBANCO.Close

      SP_PROC_BANCO cmbBanco.Text

      If TabBANCO.EOF Then
         SQL = "INSERT INTO BANCO "
            SQL = SQL & " (banco_id,codg_banco, Nome_banco)"
         SQL = SQL & " VALUES ("
            SQL = SQL & MAX_ID("banco_id", "banco", "", "", "", "")
            SQL = SQL & "'" & Trim(cmbBanco.Text) & "'"
            SQL = SQL & ",'" & Trim(lblBanco.Caption) & "'"
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabBANCO.State = 1 Then _
         TabBANCO.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_BANCO"
End Sub
