VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmABREOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abertura de Ordem de Serviço"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   11910
   Icon            =   "Abreos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   0
      TabIndex        =   59
      Top             =   4560
      Width           =   11895
      Begin VB.ComboBox cmbAuxVendedor 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   10080
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtDESCPRODUTO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtPRODUTO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MaxLength       =   30
         TabIndex        =   15
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtVALOR_PEÇA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtDESCONTO_PEÇA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtPERC_PEÇA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbVendedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10080
         TabIndex        =   17
         Top             =   195
         Width           =   1695
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtTOTAL_PEÇA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peça:"
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
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor = "
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
         Left            =   7200
         TabIndex        =   65
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto="
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
         Left            =   3000
         TabIndex        =   64
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   6165
         TabIndex        =   63
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   9000
         TabIndex        =   62
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Qtd.="
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
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total ="
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
         Left            =   9600
         TabIndex        =   60
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.TextBox txtDESCONTOPRODUTO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtTOTALPRODUTO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10545
      TabIndex        =   46
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtDESCONTOSERVIÇO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtTOTALSERVIÇO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   43
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   32
      Top             =   1800
      Width           =   11895
      Begin VB.TextBox txtValor_Total_Tarefa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbAuxMecanico 
         BackColor       =   &H80000004&
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
         Left            =   9240
         TabIndex        =   52
         Top             =   195
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbMecanico 
         BackColor       =   &H00FFFFFF&
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
         Left            =   9240
         TabIndex        =   10
         Top             =   195
         Width           =   2535
      End
      Begin VB.TextBox txtPERC_TAREFA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtDESCONTO_TAREFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtVALOR_TAREFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtDesc_Tarefa 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtCODG_TAREFA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderWidth     =   10
         X1              =   0
         X2              =   11880
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perc. Desconto="
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
         Left            =   3120
         TabIndex        =   56
         Top             =   720
         Width           =   1725
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total Tarefa="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8760
         TabIndex        =   55
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mecânico:"
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
         Left            =   8160
         TabIndex        =   51
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   5640
         TabIndex        =   50
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Desconto ="
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
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Tarefa="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   42
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblTarefa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarefa:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   29
      Top             =   -120
      Width           =   11895
      Begin VB.TextBox txtKM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin MSComctlLib.ListView LISTACT 
         Height          =   975
         Left            =   1560
         TabIndex        =   54
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1720
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nome"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Situação"
            Object.Width           =   882
         EndProperty
      End
      Begin VB.TextBox txtPLACA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   50
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbSTATUS 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         TabIndex        =   5
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtTOTALDESCONTOOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtTOTALOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cmbAUX 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbTipoOS 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtCli 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   26
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox txtNomeCt 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   480
         Width           =   2535
      End
      Begin MSMask.MaskEdBox txtCGCCPF 
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   10440
         TabIndex        =   27
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
      Begin VB.TextBox txtOs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km Atual:"
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
         Left            =   3480
         TabIndex        =   58
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label21 
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
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Desconto = "
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
         Left            =   6360
         TabIndex        =   41
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
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
         Left            =   9840
         TabIndex        =   39
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total O.S. = "
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
         Left            =   9240
         TabIndex        =   38
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo O.S.:"
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
         Left            =   5280
         TabIndex        =   35
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   3480
         TabIndex        =   34
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblCt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consultor"
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
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lblOs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número O.S."
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
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView LISTASERVIÇO 
      Height          =   1305
      Left            =   0
      TabIndex        =   40
      Top             =   3120
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2302
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
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Códg."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tarefa"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Descont."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mecanico"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCGC 
      Height          =   375
      Left            =   0
      TabIndex        =   57
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##.###.###/####-##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView LISTAPEÇA 
      Height          =   1545
      Left            =   0
      TabIndex        =   68
      Top             =   5800
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2725
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
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Códg."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtd."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Referência"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   10
      X1              =   0
      X2              =   11880
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   6000
      X2              =   6000
      Y1              =   7320
      Y2              =   7800
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc.Peças ="
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
      Left            =   6120
      TabIndex        =   48
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Peças ="
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
      Left            =   9000
      TabIndex        =   47
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc.Serviço ="
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
      Left            =   0
      TabIndex        =   45
      Top             =   7440
      Width           =   1590
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Serviços ="
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
      Left            =   2760
      TabIndex        =   44
      Top             =   7440
      Width           =   1710
   End
End
Attribute VB_Name = "frmABREOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim TOTAL_DESCONTO_SERVIÇO_N As Double, TOTAL_DESCONTO_PEÇAS_N As Double
   Dim IMPRESSORA As Printer, CONT_CURSO As Long
   Dim CONT_LINHAS As Long, PAGINA As Long
   Dim PaginaInicial, Paginafinal, NumeroDeCopias, i

Private Sub Form_Activate()
   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
   If INDR_PRI = True Then
       cmbAUX.Clear
       cmbTipoOS.Clear
       SQL = "select * from DESCR "
       SQL = SQL & "where tipo_a = 'H' "
       SQL = SQL & "order by desc_a"
       Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
       While Not TABDESCR.EOF
          cmbTipoOS.AddItem Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
          cmbAUX.AddItem TABDESCR!Codigo
          TABDESCR.MoveNext
       Wend
       TABDESCR.Close
    
       cmbStatus.Clear
       cmbStatus.AddItem "A - ATIVA"
       'cmbSTATUS.AddItem "B - BAIXADA"
       'cmbSTATUS.AddItem "C - CANCELADA"
       cmbStatus.AddItem "D - NEGOCIAÇÂO"
       'cmbSTATUS.AddItem "E - EXECUSÃO"
       'cmbSTATUS.AddItem "F - FECHADA"
    
       cmbAuxMecanico.Clear
       cmbMecanico.Clear
       SQL = "select * from USUARIO "
       SQL = SQL & " where tipo = 8 "
       SQL = SQL & " and empresa_id = " & EMPRESA_ID
       SQL = SQL & " order by nome "
       Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
       While Not TABDESCR.EOF
          cmbMecanico.AddItem TABDESCR!Nome & " - " & TABDESCR!Codigo
          cmbAuxMecanico.AddItem TABDESCR!Codigo
          TABDESCR.MoveNext
       Wend
       TABDESCR.Close
    
       cmbVENDEDOR.Clear
       cmbAuxVendedor.Clear
       SQL = "select * from VENDEDOR "
       SQL = SQL & "where status='A' " 'vendedores
       SQL = SQL & " order by nome_vend "
       Set TABDESCR = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
       While Not TABDESCR.EOF
          cmbVENDEDOR.AddItem TABDESCR!nome_vend & "-" & TABDESCR!codg_vend
          cmbAuxVendedor.AddItem TABDESCR!codg_vend
          TABDESCR.MoveNext
       Wend
       TABDESCR.Close
      INDR_PRI = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape
         If txtOs.Text <> "" And INDR_GRAVA = True Then
            Msg = "Deseja sair sem gravar?"
            Style = vbYesNo + vbCritical
            Title = "Atenção !!!"
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then
               DBARQAUX.Execute "delete * from ITEMOS where numr_os =" & txtOs.Text
               Unload Me
            End If
            Else: Unload Me
         End If
      'Case vbKeyF6: EXCLUIR_ITEM
      Case vbKeyF9
         LIMPA_TUDO
         txtOs.SetFocus
      Case vbKeyF10
         If txtOs.Text <> "" Then _
            IMPRIMIR_OS
   End Select
End Sub

Private Sub Form_Load()
   Call CentralizaJanela(frmABREOS)
   frmABREOS.Top = 0
End Sub

Private Sub TXTCGCCPF_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Clientes."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmCONSULTACHASSI.Show 1
         If SQL3 <> "" Then _
            txtPlaca.Text = SQL3
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCGCCPF.PromptInclude = False
      If txtCGCCPF.Text <> "" Then
         SQL = "select nome from CLIENTE "
         SQL = SQL & "where cgccpf = '" & txtCGCCPF.Text & "'"
         Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABCLI.EOF Then
            txtCli.Text = TABCLI!Nome
            Else
               txtCGCCPF.SelStart = 0
               txtCGCCPF.SelLength = Len(txtOs)
               MsgBox "Cliente não Cadastrado."
               txtCGCCPF.SetFocus
               Exit Sub
         End If
         Else
            txtPlaca.SetFocus
            Exit Sub
      End If
      cmbStatus.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtCGCCPF_LostFocus()
   txtCGCCPF.PromptInclude = False
   If txtCGCCPF.Text <> "" Then
      If Len(txtCGCCPF.Text) > 0 Then
         Select Case Len(txtCGCCPF.Text)
            Case Is = 11
               If Not CALCULACPF(txtCGCCPF.Text) Then
                  MsgBox "CPF com DV incorreto !!!"
                  txtCGCCPF.PromptInclude = False
                  txtCGCCPF = ""
                  txtCGCCPF.SetFocus
                  Exit Sub
               End If
            Case Is = 14
               If Not VALIDACGC(txtCGCCPF.Text) Then
                  MsgBox "CNPJ com DV incorreto !!! "
                  txtCGCCPF.PromptInclude = False
                  txtCGCCPF = ""
                  txtCGCCPF.SetFocus
                  Exit Sub
               End If
            Case Is > 14
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCGCCPF = ""
               txtCGCCPF.SetFocus
               Exit Sub
            Case Is < 11
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCGCCPF = ""
               txtCGCCPF.SetFocus
               Exit Sub
         End Select
         Else
            MsgBox "CNPJ/CPF com DV incorreto !!! "
            txtCGCCPF = ""
            txtCGCCPF.SetFocus
            Exit Sub
      End If
   End If
   txtCGCCPF.PromptInclude = True
End Sub

Private Sub cmbvendedor_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Vendedor"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   cmbVENDEDOR.Text = "Oficina"
   cmbAuxVendedor.Text = "9999"
End Sub

Private Sub txtOs_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta O.S."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtOs_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         NUMR_OS = 0
         frmCONSULTAOS.Show 1
         If Not IsNull(NUMR_OS) Then _
            If NUMR_OS > 0 Then _
               txtOs.Text = NUMR_OS
   End Select
End Sub

Private Sub txtOS_KeyPress(KeyAscii As Integer)
   On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      NUMR_SEQ_N = 0
      If txtOs.Text = "" Then
         NUMR_OS = 0
         SQL = "select * from EMPRESA "
         SQL = SQL & "where empresa_id = " & EMPRESA_ID
         Set TABEMP = DBARQEMP.OpenRecordset(SQL)
         TABEMP.Edit
            TABEMP!seq_reqorc = TABEMP!seq_reqorc + 1
         TABEMP.Update
         NUMR_OS = TABEMP!seq_reqorc
         txtOs.Text = NUMR_OS
         TABEMP.Close
         Else: NUMR_OS = txtOs.Text
      End If

      ABRE_BANCO_AUXILIAR

      SQL = "select * from CABECAOS "
      SQL = SQL & "where numr_os = " & NUMR_OS
      Set TABCABECA = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TABCABECA.EOF Then
         TRATA_OS
         If TABCABECA!Status = "F" Then
            txtOs.SelStart = 0
            txtOs.SelLength = Len(txtOs)
            txtOs.SetFocus
            MsgBox "O.S. Fechada.", vbOKOnly, "Atenção !!!"
            LIMPA_TUDO
            Exit Sub
         End If
         If TABCABECA!Status = "C" Then
            txtOs.SelStart = 0
            txtOs.SelLength = Len(txtOs)
            txtOs.SetFocus
            MsgBox "O.S. Cancelada.", vbOKOnly, "Atenção !!!"
            LIMPA_TUDO
            Exit Sub
         End If
         If TABCABECA!Status = "B" Then
            txtOs.SelStart = 0
            txtOs.SelLength = Len(txtOs)
            txtOs.SetFocus
            MsgBox "O.S. Baixada.", vbOKOnly, "Atenção !!!"
            LIMPA_TUDO
            Exit Sub
         End If
      End If
      TABCABECA.Close
      NUMR_OS = txtOs.Text
      txtDtIni.Text = Date
      txtCt.SetFocus
      'DBARQAUX.Close
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Me.Name, txtOs.Name
End Sub

Private Sub txtCli_gotfocus()
   'txtCGCCPF.SetFocus
End Sub

Private Sub txtCt_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta CT."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   LISTACT.Visible = True
   LISTACT.ListItems.Clear
   SQL = "select * from USUARIO "
   SQL = SQL & " where tipo = 6 "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABUSU = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TABUSU.EOF
      Set ITEM = LISTACT.ListItems.Add(, "seq." & TABUSU!Codigo, TABUSU!Codigo)
      ITEM.SubItems(1) = TABUSU!Nome
      ITEM.SubItems(2) = TABUSU!Status
      TABUSU.MoveNext
   Wend
   TABUSU.Close
End Sub

Private Sub txtCT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCt.Text = "" Then
         txtCt.Text = "9999"
         txtNomeCt.Text = "Consultor Geral"
         Else
            SQL = "select * from USUARIO "
            SQL = SQL & " where codigo = " & txtCt.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABUSU = DBARQEMP.OpenRecordset(SQL, 4)
            If TABUSU.EOF Then
               MsgBox "Consultor não cadastrado Não Cadastrado.", vbOKOnly, "ERRO !!!"
               txtCt.SelStart = 0
               txtCt.SelLength = Len(txtCt)
               txtCt.SetFocus
               Exit Sub
               Else
                  If TABUSU!Tipo = 6 Then
                     txtNomeCt.Text = TABUSU!Nome
                     'txtCGCCPF.SetFocus
                     Else
                        txtCt.SelStart = 0
                        txtCt.SelLength = Len(txtCt)
                        txtCt.SetFocus
                        MsgBox "Usuário não é Consultor Técnico."
                        txtCt.SetFocus
                        Exit Sub
                  End If
            End If
      End If
      cmbTipoOS.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtCt_LostFocus()
   LISTACT.Visible = False
End Sub

Private Sub txtDesc_Tarefa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtVALOR_TAREFA.SetFocus
   End If
End Sub

Private Sub txtDesc_TarefaS_GotFocus()
   cmbTipoOS.SetFocus
End Sub

Private Sub cmbmecanico_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione Mecânico"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
End Sub

Private Sub cmbmecanico_Click()
   cmbAuxMecanico.ListIndex = cmbMecanico.ListIndex
End Sub
   
Private Sub cmbmecanico_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbMecanico.Text = "" Then
         cmbMecanico.Text = "Oficina"
         cmbAuxMecanico.Text = "9999"
         Else
            SQL = "select * from USUARIO "
            SQL = SQL & " where codigo = " & cmbAuxMecanico.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then
            If TABDESCR.EOF Then
               If TABDESCR!Tipo <> 8 Then
                  TABDESCR.Close
                  MsgBox "Permitido somente mecanico."
                  Exit Sub
               End If
            End If
            End If
            TABDESCR.Close
      End If
      KeyAscii = 0
      txtDESCONTO_TAREFA.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtDESCONTO_TAREFA_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Valor de Desconto da tarefa"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtDESCONTO_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDESCONTO_TAREFA.Text <> "" Then
         If txtVALOR_TAREFA.Text <> "" Then
            VALOR_DESCONTO_N = txtDESCONTO_TAREFA.Text
            If VALOR_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_TAREFA.Text
               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               txtValor_Total_Tarefa.Text = Format(VALOR_TOTAL_N, "fixed")
               txtVALOR_TAREFA.Refresh

               txtPERC_TAREFA.Text = Format(((VALOR_DESCONTO_N * VALOR_ITEM_N) / 100), "fixed")
               txtPERC_TAREFA.Refresh
               GRAVA_ITEM_OS
               LIMPA_BODY_SERVIÇO
               Else: txtPERC_TAREFA.SetFocus
            End If
            Else: txtCODG_TAREFA.SetFocus
         End If
         Else: txtPERC_TAREFA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtPERC_TAREFA_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Percentual de Desconto da tarefa"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtPERC_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtVALOR_TAREFA.Text <> "" Then
         If txtPERC_TAREFA.Text <> "" Then
            PERC_DESCONTO_N = txtPERC_TAREFA.Text
            If PERC_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_TAREFA.Text

               VALOR_DESCONTO_N = ((PERC_DESCONTO_N * VALOR_ITEM_N) / 100)
               txtDESCONTO_TAREFA.Text = VALOR_DESCONTO_N

               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               txtValor_Total_Tarefa.Text = Format(VALOR_TOTAL_N, "fixed")
               txtValor_Total_Tarefa.Refresh

               GRAVA_ITEM_OS
               LIMPA_BODY_SERVIÇO
               Else
                  txtVALOR_TAREFA.Enabled = True
                  txtVALOR_TAREFA.SetFocus
            End If
            Else
               txtVALOR_TAREFA.Enabled = True
               txtVALOR_TAREFA.SetFocus
         End If
         Else
            txtVALOR_TAREFA.Enabled = True
            txtVALOR_TAREFA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtVALOR_TAREFA_GotFocus()
   txtVALOR_TAREFA.SelStart = 0
   txtVALOR_TAREFA.SelLength = Len(txtVALOR_TAREFA)
End Sub

Private Sub txtVALOR_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_ITEM_OS
      LIMPA_BODY_SERVIÇO
      txtCODG_TAREFA.SetFocus
   End If
End Sub

Private Sub cmbSTATUS_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione tipo de Ordem se Serviço"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbStatus.Text = "" Then _
         cmbStatus.Text = "A - Ativa"
      KeyAscii = 0
      txtKm.SetFocus
   End If
End Sub

Private Sub txtkm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbStatus.Text = "" Then _
         cmbStatus.Text = "A - Ativa"
      KeyAscii = 0
      txtCODG_TAREFA.SetFocus
   End If
End Sub

Private Sub txtCODG_TAREFA_GotFocus()
   txtVALOR_TAREFA.Enabled = False
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Tarefas."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F6 - Exclui Tarefas."
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (5)
   frmINICIO.BARI.Panels(5).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(5).AutoSize = sbrContents
   
End Sub

Private Sub txtCODG_TAREFA_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtCODG_TAREFA.Text) <> "" Then
            ABRE_BANCO_AUXILIAR
            SQL = "select * from ITEMOS "
            SQL = SQL & "where numr_os = " & NUMR_OS
            Set TABAUX = DBARQAUX.OpenRecordset(SQL)
            If Not TABAUX.EOF Then
               Msg = "Confirma exclusão dessa tarefa ?"
               PERGUNTA
               If RESPOSTA = vbYes Then
                  TABAUX.Delete
                  SETA_GRID_SERVIÇO
                  ATUALIZA_TOTAL_OS
                  txtCODG_TAREFA.SetFocus
               End If
            End If
            'TABAUX.Close
            'DBARQAUX.Close
         End If
      Case vbKeyF7
         CODG_PROD_A = ""
         frmCONSULTATAREFA.Show 1
         If CODG_PROD_A <> "" Then _
            txtCODG_TAREFA.Text = CODG_PROD_A
         CODG_PROD_A = ""
         txtCODG_TAREFA.SetFocus
   End Select
End Sub

Private Sub txtCODG_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtCODG_TAREFA.Text) = "" Then
         MsgBox "Tarefa deve ser informada."
         Exit Sub
         Else
            ABRE_BANCO_AUXILIAR
            SQL = "select * from TAREFA "
            SQL = SQL & "where codg_tarefa = '" & Trim(txtCODG_TAREFA.Text) & "'"
            Set TABTEMP = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TABTEMP.EOF Then
               txtDesc_Tarefa.Text = TABTEMP!Descricao
               txtVALOR_TAREFA.Text = Format(TABTEMP!valor_tarefa, "fixed")

               SQL = "select * from ITEMOS "
               SQL = SQL & "where numr_os = " & NUMR_OS
               SQL = SQL & " and codg_tarefa = '" & TABTEMP!Codg_tarefa & "'"
               Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
               If Not TABAUX.EOF Then
                  txtDESCONTO_TAREFA.Text = Format(TABAUX!valor_desc_tarefa, "fixed")
                  txtVALOR_TAREFA.Text = Format(TABAUX!valor_tarefa, "fixed")
                  cmbAuxMecanico.Text = TABAUX!codg_mecanico

                  SQL = "select * from USUARIO "
                  SQL = SQL & " where codigo = " & TABAUX!codg_mecanico
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID
                  Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
                  If Not TABDESCR.EOF Then _
                     cmbMecanico.Text = TABDESCR!Nome & " - " & TABDESCR!Codigo
                  TABDESCR.Close
               End If
               TABAUX.Close
               Else
                  TABTEMP.Close
                  MsgBox "Tarefa não cadastrada, verifique."
                  Exit Sub
            End If
            TABTEMP.Close
            DBARQAUX.Close
      End If
      cmbMecanico.SetFocus
   End If
End Sub

Private Sub cmbTIPOOS_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione Tipos O.S.."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub cmbTipoOS_Click()
   cmbAUX.ListIndex = cmbTipoOS.ListIndex
End Sub
   
Private Sub cmbTipoOS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbTipoOS.Text = "" Then
         cmbTipoOS.Text = "Normal"
         cmbAUX.Text = 1
      End If
      KeyAscii = 0
      txtPlaca.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtCHASSI_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Chassi"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPlaca.Text = "" Then
         'MsgBox "Chassi deve ser informado."
         'txtplaca.SetFocus
         'Exit Sub
         txtPlaca.SetFocus
         Exit Sub
         Else
            ABRE_BANCO_AUXILIAR
            PROCURA_PLACA
            DBARQAUX.Close
      End If
      cmbStatus.SetFocus
   End If
End Sub

Private Sub txtCHASSI_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmCONSULTACHASSI.Show 1
         If SQL3 <> "" Then _
            txtPlaca.Text = SQL3
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub
'==========
Private Sub txtPlaca_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Placa Veículo"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtplaca_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPlaca.Text = "" Then
         MsgBox "Placa deve ser informado."
         txtPlaca.SetFocus
         Exit Sub
         Else
            ABRE_BANCO_AUXILIAR
            SQL = "select placa from CHASSI "
            SQL = SQL & " where placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
            Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
            If TABAUX.EOF Then
               MsgBox "Placa não cadastrado."
               txtPlaca.SetFocus
               Exit Sub
               Else: PROCURA_PLACA
            End If
      End If
      cmbStatus.SetFocus
      Else
         If KeyAscii <> 8 Then
            CRITERIO = txtPlaca.Text
            If Len(CRITERIO) = 3 Then
               txtPlaca.Text = CRITERIO & "-"
               txtPlaca.SelStart = 4
               txtPlaca.Refresh
            End If
        End If
   End If
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmCONSULTACHASSI.Show 1
         If SQL3 <> "" Then
            ABRE_BANCO_AUXILIAR
            SQL = "select placa from CHASSI "
            SQL = SQL & "where nr_chassi = '" & SQL3 & "'"
            Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TABAUX.EOF Then
               txtPlaca.Text = Left(TABAUX!placa, 3) & "-" & Right(TABAUX!placa, 5)
            End If
            TABAUX.Clone
            DBARQAUX.Close
         End If
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub
'============================ PEÇAS
Private Sub txtproduto_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F6 - Excluir Peça"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Produtos"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (5)
   frmINICIO.BARI.Panels(5).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(5).AutoSize = sbrContents
   
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF6
         If txtOs.Text <> "" And txtPRODUTO.Text <> "" Then _
            MATA_SEQ
      Case vbKeyF7
         frmCONSULTAPRODUTO.Show 1
         If CODG_PROD_A <> "" Then
            txtPRODUTO.Text = CODG_PROD_A
            txtPRODUTO.SetFocus
         End If
   End Select
End Sub

Private Sub txtproduto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_prod = '" & txtPRODUTO.Text & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID
      Set TABPRODUTO = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
      If TABPRODUTO.EOF Then
         'MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção !!!"
         'txtPRODUTO.SelStart = 0
         'txtPRODUTO.SelLength = Len(txtPRODUTO)
         'txtPRODUTO.SetFocus
         'Exit Sub
         cmbVENDEDOR.SetFocus
         Exit Sub
         Else
            txtDESCPRODUTO.Text = TABPRODUTO!Descricao
            'frmINICIO.BARI.Panels(3).Text = "Quantidade em Estoque = " & _
               TABPRODUTO!qtd - TABPRODUTO!qtd_balcao
            'frmINICIO.BARI.Panels(3).AutoSize = sbrContents

            QTD_ESTOQUE = TABPRODUTO!QTD - TABPRODUTO!qtd_balcao
            If Not IsNull(TABPRODUTO!PRECO_VENDA) Then
               'frmINICIO.BARI.Panels(4).Text = Format(TABPRODUTO!PRECO_venda, "fixed")
               txtVALOR_PEÇA.Text = Format(TABPRODUTO!PRECO_VENDA, "fixed")
            End If
            If txtOs.Text = "" Or txtPRODUTO.Text = "" Then _
               Exit Sub
            SQL = "select * from ITEMREQ "
            SQL = SQL & "where codg_prod = '" & txtPRODUTO.Text & "'"
            SQL = SQL & " and numr_req = " & txtOs.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABREQITEM = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
            If Not TABREQITEM.EOF Then
               'txtSeq.Text = tabreqitem!seq
               txtVALOR_PEÇA.Text = TABREQITEM!Valor_Item
               txtDESCONTO_PEÇA.Text = Format(TABREQITEM!PERC_DESCONTO, "fixed")
               txtQtd.Text = TABREQITEM!qtd_pedida
               QTD_PEDIDO = TABREQITEM!seq
               QTD_EXTORNO_BALCAO = TABREQITEM!qtd_pedida
               VALOR_ITEM_N = TABREQITEM!Valor_Item
               VALOR_DIFERENCA_N = TABREQITEM!VALOR_TOTAL_ITEM
               MsgBox "Produto já consta nesse O.S. seqüência = " & TABREQITEM!seq
               QTD_ESTOQUE = TABPRODUTO!QTD + QTD_EXTORNO_BALCAO - TABPRODUTO!qtd_balcao
            End If
            TABPRODUTO.Close
            TABREQITEM.Close
      End If
      cmbVENDEDOR.SetFocus
   End If
End Sub

Private Sub cmbvendedor_Click()
   cmbAuxVendedor.ListIndex = cmbVENDEDOR.ListIndex
End Sub
   
Private Sub cmbvendedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbVENDEDOR.Text = "" Then
         cmbVENDEDOR.Text = "Oficina"
         cmbAuxVendedor.Text = "9999"
         Else
            SQL = "select * from VENDEDOR "
            SQL = SQL & "where codg_vend = " & cmbAuxVendedor.Text
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
            If TABDESCR.EOF Then
               MsgBox "Vendedor não cadastrado."
               cmbVENDEDOR.SetFocus
               Exit Sub
            End If
            TABDESCR.Close
      End If
      KeyAscii = 0
      txtQtd.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtQtd_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Quantidade"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Quantidade Disponível = " & QTD_ESTOQUE - QTD_PEDIDO
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (5)
   frmINICIO.BARI.Panels(5).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(5).AutoSize = sbrContents
   
   If txtPRODUTO.Text = Empty Then
      MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro !!!"
      txtPRODUTO.Text = 99999999
      txtPRODUTO.SetFocus
      Exit Sub
   End If
   If txtQtd.Text <> "" Then
      txtQtd.SelStart = 0
      txtQtd.SelLength = Len(txtQtd)
   End If
End Sub

Private Sub txtqtd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbAuxVendedor.Text = "" Then
         Beep
         MsgBox "Selecione Vendedor.", vbOKOnly, "Atenção !!!"
         cmbVENDEDOR.SetFocus
         Exit Sub
      End If
      If txtOs.Text = "" Then
         Beep
         MsgBox "Número de O.S. Inválido.", vbOKOnly, "Atenção !!!"
         txtOs.SetFocus
         Exit Sub
      End If
      txtCGCCPF.PromptInclude = False
      If txtCGCCPF.Text = "" Then
         Beep
         MsgBox "Informe Cliente.", vbOKOnly, "Atenção !!!"
         txtCGCCPF.SetFocus
         Exit Sub
      End If
      If txtPRODUTO.Text = "" Then
         MsgBox "Seqüência sem codigo de Produto.", vbOKOnly, "Atenção !!!"
         txtPRODUTO.SetFocus
         Exit Sub
      End If
      VALOR_ITEM_N = 0
      VALOR_ITEM_N = 0 & txtVALOR_PEÇA.Text
      If Not IsNull(VALOR_ITEM_N) Then
         If VALOR_ITEM_N <= 0 Then
            MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção !!!"
            txtPRODUTO.SetFocus
            Exit Sub
         End If
      End If
      KeyAscii = 0
      If txtQtd.Text = "" Then
         Beep
         MsgBox "Informe a quantidade.", vbOKOnly, "Atenção !!!"
         txtQtd.SetFocus
         Exit Sub
         Else
            'quantidade pedida
            QTD_PEDIDO = txtQtd.Text
            txtQtd.Text = Format(txtQtd.Text, "###.000")
            If INDR_CONTROLA_ESTOQUE = True Then
               If QTD_ESTOQUE < QTD_PEDIDO Then
                  Beep
                  MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção !!!"
                  txtQtd.SetFocus
                  Exit Sub
               End If
            End If
            If QTD_PEDIDO <= 0 Then
               Beep
               MsgBox "Quantidade pedida não permitido, deve ser maior que 0.", vbOKOnly, "Atenção !!!"
               txtQtd.SetFocus
               Exit Sub
            End If
      End If
      txtDESCONTO_PEÇA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtDESCONTO_peça_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Valor de Desconto"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtDESCONTO_peça_keypress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDESCONTO_PEÇA.Text <> "" Then
         If txtVALOR_PEÇA.Text <> "" Then
            VALOR_DESCONTO_N = txtDESCONTO_PEÇA.Text
            If VALOR_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_PEÇA.Text
               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               'txtTOTALPRODUTO.Text = Format(VALOR_TOTAL_N, "fixed")
               'txtTOTALPRODUTO.Refresh

               txtPERC_PEÇA.Text = Format(((VALOR_DESCONTO_N * VALOR_ITEM_N) / 100), "fixed")
               txtPERC_PEÇA.Refresh
               GRAVA_CABECA
               LIMPA_BODY_PEÇA
               Else: txtPERC_PEÇA.SetFocus
            End If
            Else: txtPRODUTO.SetFocus
         End If
         Else: txtPERC_PEÇA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtPERC_PEÇA_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Percentual de Desconto"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtPERC_peça_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtVALOR_PEÇA.Text <> "" Then
         If txtPERC_PEÇA.Text <> "" Then
            PERC_DESCONTO_N = txtPERC_PEÇA.Text
            If PERC_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_PEÇA.Text

               VALOR_DESCONTO_N = ((PERC_DESCONTO_N * VALOR_ITEM_N) / 100)
               txtDESCONTO_PEÇA.Text = VALOR_DESCONTO_N

               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               'txtTOTALPRODUTO.Text = Format(VALOR_TOTAL_N, "fixed")
               'txtTOTALPRODUTO.Refresh

               GRAVA_CABECA
               LIMPA_BODY_PEÇA
               Else
                  txtVALOR_PEÇA.Enabled = True
                  txtVALOR_PEÇA.SetFocus
            End If
            Else
               txtVALOR_PEÇA.Enabled = True
               txtVALOR_PEÇA.SetFocus
         End If
         Else
            txtVALOR_PEÇA.Enabled = True
            txtVALOR_PEÇA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtVALOR_peça_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_CABECA
      LIMPA_BODY_PEÇA
   End If
End Sub

Private Sub txtvalor_peça_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Valor Peça"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents


   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   txtVALOR_PEÇA.SelStart = 0
   txtVALOR_PEÇA.SelLength = Len(txtVALOR_PEÇA)
End Sub

'=================================================================
Private Sub GRAVA_CABECA()
   ABRE_BANCO_AUXILIAR
   GRAVA_CABECA_OS
   SQL = "select * from CABECAREQ "
   SQL = SQL & " where numr_req = " & txtOs.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABCABECA = DBARQEMP.OpenRecordset(SQL)
   If TABCABECA.EOF Then
      TABCABECA.AddNew
      Else: TABCABECA.Edit
   End If
   GRAVA_PEÇA
   TABCABECA.Close
End Sub

Private Sub GRAVA_PEÇA()
   txtCGCCPF.PromptInclude = False
   SQL = "select cgccpf from CLIENTE "
   SQL = SQL & "where cgccpf = '" & txtCGCCPF.Text & "'"
   Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
   If TABCLI.EOF Then
      TABCLI.Close
      MsgBox "Cliente não cadastrado, verifique."
      Exit Sub
   End If
   TABCLI.Close
   TABCABECA!numr_req = txtOs.Text
   txtCGCCPF.PromptInclude = False
   TABCABECA!CGCCPF = txtCGCCPF.Text
   If cmbAuxVendedor.Text = "" Then
      cmbVENDEDOR.Text = "Oficina"
      cmbAuxVendedor.Text = "9999"
   End If
   TABCABECA!vendedor = cmbAuxVendedor.Text
   TABCABECA!dt_req = Date
   TABCABECA!valor_total = VALOR_TOTAL_N
   TABCABECA!TIPOvenda_id = 1
   'AGORA TODAS VENDAS a vista vai para emitir cupom ou nota
   TABCABECA!Status = 2
   If txtDESCONTO_PEÇA.Text = "" Then
      TABCABECA!Valor_Desconto = 0
      TABCABECA!PERC_desc = 0
      Else
         If txtDESCONTOPRODUTO.Text <> "" Then
            VALOR_TOTAL_DESCONTO_N = txtDESCONTOPRODUTO.Text
            TABCABECA!Valor_Desconto = txtDESCONTOPRODUTO.Text
         End If
         TABCABECA!PERC_desc = (VALOR_TOTAL_DESCONTO_N / 100)
   End If
   TABCABECA!TIPO_REGISTRO = "S"
   TABCABECA!USUARIO_LIBERA_VENDA = USUARIO_LIBERA_VENDA
   If Not IsNull(CODG_USU_N) Then _
      TABCABECA!CODG_USU = CODG_USU_N
   TABCABECA!EMPRESA_ID = EMPRESA_ID
   TABCABECA.Update
   GRAVA_PEÇA_ITEM
End Sub

Private Sub GRAVA_PEÇA_ITEM()
   QTD_PEDIDO = txtQtd.Text
   VALOR_ITEM_N = txtVALOR_PEÇA.Text
   PERC_DESCONTO_N = 0
   If txtPERC_PEÇA.Text <> "" Then _
      PERC_DESCONTO_N = txtPERC_PEÇA.Text
   If txtDESCONTO_PEÇA.Text <> "" Then _
      VALOR_DESCONTO_N = txtDESCONTO_PEÇA.Text
   NUMR_SEQ_N = 1
   SQL = "select max(seq) from ITEMREQ "
   SQL = SQL & "where numr_req = " & txtOs.Text
   Set TABREQITEM = DBARQEMP.OpenRecordset(SQL)
   If Not TABREQITEM.EOF Then _
      If Not IsNull(TABREQITEM.Fields(0).Value) Then _
         NUMR_SEQ_N = 1 + TABREQITEM.Fields(0).Value
   TABREQITEM.Close

   SQL = "select * from ITEMREQ "
   SQL = SQL & " where codg_prod = '" & txtPRODUTO.Text & "'"
   SQL = SQL & " and numr_req = " & txtOs.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABREQITEM = DBARQEMP.OpenRecordset(SQL)
   If TABREQITEM.EOF Then
      TABREQITEM.AddNew
        TABREQITEM!numr_req = txtOs.Text
        TABREQITEM!seq = NUMR_SEQ_N
        TABREQITEM!Codg_Prod = txtPRODUTO.Text
      Else: TABREQITEM.Edit
   End If
   TABREQITEM!qtd_pedida = QTD_PEDIDO
   TABREQITEM!Valor_Item = VALOR_ITEM_N
   TABREQITEM!PERC_desc = PERC_DESCONTO_N
   TABREQITEM!EMPRESA_ID = EMPRESA_ID
   'TABREQITEM!VALOR_TOTAL_ITEM = (VALOR_ITEM_N * QTD_PEDIDO) - VALOR_DESCONTO_N
   TABREQITEM.Update
   TABREQITEM.Close
   SETA_GRID_PEÇA
   ATUALIZA_TOTAL_OS
End Sub

Private Sub MATA_SEQ()
   SQL = "select * from ITEMREQ "
   SQL = SQL & " where codg_prod = '" & txtPRODUTO.Text & "'"
   SQL = SQL & " and numr_req = " & txtOs.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABTEMP = DBARQEMP.OpenRecordset(SQL)
   If Not TABTEMP.EOF Then
      Msg = "Deseja Excluir Esse Item?"
      Style = vbYesNo + 32
      Title = "Atenção !!!"
      Help = "DEMO.HLP"
      Ctxt = 1000
      RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
      If RESPOSTA = vbYes Then
         If INDR_ATUALIZA_ESTOQUE = False Then
            SQL = "select * from PRODUTO "
            SQL = SQL & " where codg_prod = '" & TABTEMP!Codg_Prod & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABPRODUTO = DBARQEMP.OpenRecordset(SQL)
            If Not TABPRODUTO.EOF Then
               TABPRODUTO.Edit
                  TABPRODUTO!qtd_balcao = TABPRODUTO!qtd_balcao - TABTEMP!qtd_pedida
               TABPRODUTO.Update
            End If
            TABPRODUTO.Close
         End If
         TABTEMP.Delete
         TABTEMP.Close
         LIMPA_BODY_PEÇA
         SETA_GRID_PEÇA
         ATUALIZA_TOTAL_OS
         Else: TABTEMP.Close
      End If
   End If
   txtPRODUTO.SetFocus
End Sub

Private Sub LIMPA_TUDO()
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   TOTAL_SERVIÇO_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0

   txtPlaca.Text = ""
   LISTASERVIÇO.ListItems.Clear
   txtKm.Text = ""
   LISTAPEÇA.ListItems.Clear
   TOTAL_SERVIÇO_N = 0
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   txtOs.Text = ""
   txtCt.Text = ""
   txtNomeCt.Text = ""
   cmbAUX.Text = ""
   cmbTipoOS.Text = ""
   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   txtCli.Text = ""
   cmbStatus.Text = ""
   txtTOTALDESCONTOOS.Text = ""
   txtTOTALOS.Text = ""
   txtDESCONTOSERVIÇO.Text = ""
   txtTOTALSERVIÇO.Text = ""
   txtDESCONTOPRODUTO.Text = ""
   txtTOTALPRODUTO.Text = ""
   LIMPA_BODY_SERVIÇO
   LIMPA_BODY_PEÇA
End Sub

Private Sub LIMPA_BODY_SERVIÇO()
   txtCODG_TAREFA.Text = ""
   txtDesc_Tarefa.Text = ""
   txtDESCONTO_TAREFA.Text = ""
   txtPERC_TAREFA.Text = ""
   txtVALOR_TAREFA.Text = ""
   txtValor_Total_Tarefa.Text = ""
End Sub

Private Sub LIMPA_BODY_PEÇA()
   txtPRODUTO.Text = ""
   txtDESCPRODUTO.Text = ""
   cmbVENDEDOR.Text = ""
   cmbAuxVendedor.Text = ""
   txtQtd.Text = ""
   txtDESCONTO_PEÇA.Text = ""
   txtPERC_PEÇA.Text = ""
   txtVALOR_PEÇA.Text = ""
   txtTOTAL_PEÇA.Text = ""
   txtPRODUTO.SetFocus
End Sub

Private Sub TRATA_OS()
   MOSTRA_OS
   SETA_GRID_SERVIÇO
   SETA_GRID_PEÇA
   ATUALIZA_TOTAL_OS
End Sub

Private Sub MOSTRA_OS()
   txtOs.Text = TABCABECA!NUMR_OS
   txtCt.Text = TABCABECA!ct
   If Not IsNull(TABCABECA!km_atual) Then
      txtKm.Text = TABCABECA!km_atual
      Else: txtKm.Text = ""
   End If
   SQL = "select nome from USUARIO "
   SQL = SQL & " where codigo = " & TABCABECA!ct
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABUSU = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TABUSU.EOF Then _
      txtNomeCt.Text = TABUSU!Nome
   TABUSU.Close

   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'H' "
   SQL = SQL & "and codigo = " & TABCABECA!tipo_os
   Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TABDESCR.EOF Then
      cmbTipoOS.Text = Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
      cmbAUX.Text = TABDESCR!Codigo
   End If
   TABDESCR.Close

   txtDtIni.PromptInclude = False
      txtDtIni.Text = TABCABECA!dt_abertura
   txtDtIni.PromptInclude = True

   SQL = "select * from CHASSI "
   SQL = SQL & "where placa = '" & Trim(TABCABECA!placa) & "'"
   Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
   If Not TABAUX.EOF Then
      txtPlaca.Text = TABAUX!placa
      txtCGCCPF.PromptInclude = False
      txtCGCCPF.Text = TABAUX!CGCCPF
      SQL = "select nome from CLIENTE "
      SQL = SQL & "where cgccpf = '" & TABAUX!CGCCPF & "'"
      Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
      If Not TABCLI.EOF Then
         txtCli.Text = TABCLI!Nome
         Else: MsgBox "Cliente não cadastrado, verifique."
      End If
   End If
   TABAUX.Close

   If TABCABECA!Status = "A" Then _
      cmbStatus.Text = "Aberta"
   If TABCABECA!Status = "B" Then _
      cmbStatus.Text = "Baixada"
   If TABCABECA!Status = "C" Then _
      cmbStatus.Text = "Cancelada"
   If TABCABECA!Status = "D" Then _
      cmbStatus.Text = "Em Negociação"
   If TABCABECA!Status = "E" Then _
      cmbStatus.Text = "Em Execução"
   If TABCABECA!Status = "F" Then _
      cmbStatus.Text = "Fechada"
End Sub

Private Sub SETA_GRID_SERVIÇO()
   LISTASERVIÇO.ListItems.Clear
   VALOR_TOTAL_N = 0
   TOTAL_SERVIÇO_N = 0
   VALOR_DESCONTO_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0
   NUMR_SEQ_N = 1
   SQL = "select * from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   SQL = SQL & " order by hora_inicio"
   Set TABTEMP = DBARQAUX.OpenRecordset(SQL, 4)
   While Not TABTEMP.EOF
      NUMR_SEQ_N = 1 + NUMR_SEQ_N
      Set ITEM = LISTASERVIÇO.ListItems.Add(, "seq." & NUMR_SEQ_N, TABTEMP!Codg_tarefa)
      SQL = "select * from TAREFA "
      SQL = SQL & "where codg_tarefa = '" & TABTEMP!Codg_tarefa & "'"
      Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TABAUX.EOF Then _
         ITEM.SubItems(1) = TABAUX!Descricao
      TABAUX.Close
      TOTAL_SERVIÇO_N = TOTAL_SERVIÇO_N + TABTEMP!valor_tarefa
      TOTAL_DESCONTO_SERVIÇO_N = TOTAL_DESCONTO_SERVIÇO_N + TABTEMP!valor_desc_tarefa

      ITEM.SubItems(2) = Format(TABTEMP!valor_tarefa, "fixed")
      ITEM.SubItems(3) = Format(TABTEMP!valor_desc_tarefa, "fixed")
      ITEM.SubItems(4) = Format(TABTEMP!valor_tarefa - TABTEMP!valor_desc_tarefa, "fixed")
      If TABTEMP!Status = "A" Then _
         ITEM.SubItems(5) = "Ativo"
      If TABTEMP!Status = "B" Then _
         ITEM.SubItems(5) = "Baixado"
      If TABTEMP!Status = "C" Then _
         ITEM.SubItems(5) = "Cancelado"
      If TABTEMP!Status = "E" Then _
         ITEM.SubItems(5) = "Execução"
      SQL = "select * from USUARIO "
      SQL = SQL & "where codigo = " & TABTEMP!codg_mecanico
      Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
      If Not TABDESCR.EOF Then _
         ITEM.SubItems(6) = TABDESCR!Nome & " - " & TABDESCR!Codigo
      TABDESCR.Close
      TABTEMP.MoveNext
   Wend
   TABTEMP.Close
   txtTOTALSERVIÇO.Text = Format(TOTAL_SERVIÇO_N - TOTAL_DESCONTO_SERVIÇO_N, "fixed")
   txtTOTALSERVIÇO.Refresh
   txtDESCONTOSERVIÇO.Text = Format(TOTAL_DESCONTO_SERVIÇO_N, "fixed")
   txtDESCONTOSERVIÇO.Refresh
End Sub

Private Sub SETA_GRID_PEÇA()
   LISTAPEÇA.ListItems.Clear
   NUMR_SEQ_N = 0
   VALOR_DESCONTO_N = 0
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   SQL = "select * from ITEMREQ "
   SQL = SQL & "where numr_req = " & txtOs.Text
   If NUMR_SEQ_N < 10 Then
      SQL = SQL & " order by seq asc"
      Else: SQL = SQL & " order by seq desc"
   End If
   Set TABREQITEM = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
   While Not TABREQITEM.EOF
      TOTAL_PEÇAS_N = TOTAL_PEÇAS_N + (TABREQITEM!Valor_Item * TABREQITEM!qtd_pedida)
      TOTAL_DESCONTO_PEÇAS_N = TOTAL_DESCONTO_PEÇAS_N + (TABREQITEM!PERC_desc * TABREQITEM!Valor_Item / 100)

      Set ITEM = LISTAPEÇA.ListItems.Add(, "seq." & TABREQITEM!Codg_Prod, TABREQITEM!Codg_Prod)
      SQL = "select descricao,referencia from PRODUTO "
      SQL = SQL & "where codg_prod = '" & TABREQITEM!Codg_Prod & "'"
      Set TABTEMP = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
      If Not TABTEMP.EOF Then
         ITEM.SubItems(1) = TABTEMP!Descricao
         If Not IsNull(TABTEMP!Referencia) Then
            ITEM.SubItems(6) = TABTEMP!Referencia
         End If
      End If
      TABTEMP.Close
      ITEM.SubItems(2) = TABREQITEM!qtd_pedida
      ITEM.SubItems(3) = Format(TABREQITEM!Valor_Item, "fixed")
      ITEM.SubItems(4) = Format(TABREQITEM!Valor_Item * TABREQITEM!PERC_desc / 100, "fixed")
      ITEM.SubItems(5) = Format(TABREQITEM!Valor_Item * TABREQITEM!qtd_pedida - (TABREQITEM!Valor_Item * TABREQITEM!PERC_desc / 100), "fixed")
      TABREQITEM.MoveNext
   Wend
   TABREQITEM.Close
   txtTOTALPRODUTO.Text = Format(TOTAL_PEÇAS_N - TOTAL_DESCONTO_PEÇAS_N, "fixed")
   txtTOTALPRODUTO.Refresh
   txtDESCONTOPRODUTO.Text = Format(TOTAL_DESCONTO_PEÇAS_N, "fixed")
   txtDESCONTOPRODUTO.Refresh
End Sub

Private Sub GRAVA_CABECA_OS()
   SQL = "select * from CABECAOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TABCABECA = DBARQAUX.OpenRecordset(SQL)
   If Not TABCABECA.EOF Then
      TABCABECA.Edit
      Else
         TABCABECA.AddNew
            TABCABECA!dt_abertura = Now
   End If
   TABCABECA!NUMR_OS = NUMR_OS
   TABCABECA!placa = Replace(txtPlaca.Text, "-", "")
   If cmbAUX.Text <> "" Then
      TABCABECA!tipo_os = cmbAUX.Text
   End If
   If txtCt.Text <> "" Then
      TABCABECA!ct = txtCt.Text
   End If
   If cmbStatus.Text <> "" Then
      TABCABECA!Status = Left(cmbStatus.Text, 1)
   End If
   TABCABECA!km_atual = txtKm.Text
   'If txtTOTALDESCONTOOS.Text = "" Then
   '   TABCABECA!valor_desconto = 0
   '   Else: TABCABECA!valor_desconto = txtTOTALDESCONTOOS.Text
   'End If
   TABCABECA.Update
   TABCABECA.Close
End Sub

Private Sub GRAVA_ITEM_OS()
   If txtOs.Text = "" Then
      MsgBox "Informe número de O.S."
      txtOs.SetFocus
      Exit Sub
   End If
   If txtCt.Text = "" Then
      MsgBox "Informe Consultor Técnico"
      txtCt.SetFocus
      Exit Sub
   End If
   If cmbTipoOS.Text = "" Then
      MsgBox "Informe Tipo de O.S."
      cmbTipoOS.SetFocus
      Exit Sub
   End If
   If cmbStatus.Text = "" Then
      MsgBox "Informe status da O.S."
      cmbStatus.SetFocus
      Exit Sub
   End If
   If Trim(txtCODG_TAREFA.Text) = "" Then
      MsgBox "Informe Código da tarefa da O.S."
      txtCODG_TAREFA.SetFocus
      Exit Sub
   End If
   If cmbMecanico.Text = "" Then
      MsgBox "Informe mecanico dessa tarefa da O.S."
      cmbMecanico.SetFocus
      Exit Sub
   End If
   If txtVALOR_TAREFA.Text = "" Then
      MsgBox "Informe valor da tarefa dessa O.S."
      txtVALOR_TAREFA.SetFocus
      Exit Sub
   End If

   ABRE_BANCO_AUXILIAR

   GRAVA_CABECA_OS

   SQL = "select * from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   SQL = SQL & " and codg_tarefa = '" & Trim(txtCODG_TAREFA.Text) & "'"
   Set TABAUX = DBARQAUX.OpenRecordset(SQL)
   If TABAUX.EOF Then
      TABAUX.AddNew
         TABAUX!HORA_INICIO = Now
      Else: TABAUX.Edit
   End If
   TABAUX!NUMR_OS = NUMR_OS
   TABAUX!Codg_tarefa = Trim(txtCODG_TAREFA.Text)
   TABAUX!valor_tarefa = txtVALOR_TAREFA.Text
   If txtDESCONTO_TAREFA.Text <> "" Then
      TABAUX!valor_desc_tarefa = txtDESCONTO_TAREFA.Text
      Else: TABAUX!valor_desc_tarefa = 0
   End If
   TABAUX!Status = "A"
   TABAUX!codg_mecanico = cmbAuxMecanico.Text
   TABAUX.Update
   TABAUX.Close
   SETA_GRID_SERVIÇO
   txtCODG_TAREFA.SetFocus
   DBARQAUX.Close

   ATUALIZA_TOTAL_OS
End Sub

Private Sub PROCURA_PLACA()
   SQL = "select * from CHASSI "
   If txtPlaca.Text <> "" Then _
      SQL = SQL & " where placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
   Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
   If TABAUX.EOF Then
      MsgBox "Placa não cadastrado."
      txtPlaca.SetFocus
      Exit Sub
      Else
         txtCGCCPF.PromptInclude = False
         txtCGCCPF.Text = TABAUX!CGCCPF
         'txtPLACA.Text = Left(TABAUX!placa, 3) & "-" & Right(TABAUX!placa, 5)

         SQL = "select nome from CLIENTE "
         SQL = SQL & "where cgccpf = '" & txtCGCCPF.Text & "'"
         Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABCLI.EOF Then
            txtCli.Text = TABCLI!Nome
            txtCGCCPF.PromptInclude = False
            If txtCGCCPF.Text <> "" Then
               If Len(txtCGCCPF.Text) > 0 Then
                  Select Case Len(txtCGCCPF.Text)
                     Case Is = 11
                        If Not CALCULACPF(txtCGCCPF.Text) Then
                           MsgBox "CPF com DV incorreto !!!"
                           txtCGCCPF.PromptInclude = False
                           'txtCGCCPF = ""
                           'txtCGCCPF.SetFocus
                           Exit Sub
                        End If
                     Case Is = 14
                        If Not VALIDACGC(txtCGCCPF.Text) Then
                           MsgBox "CNPJ com DV incorreto !!! "
                           txtCGCCPF.PromptInclude = False
                           'txtCGCCPF = ""
                           'txtCGCCPF.SetFocus
                           Exit Sub
                        End If
                     Case Is > 14
                        MsgBox "CNPJ/CPF com DV incorreto !!! "
                        'txtCGCCPF = ""
                        'txtCGCCPF.SetFocus
                        Exit Sub
                     Case Is < 11
                        MsgBox "CNPJ/CPF com DV incorreto !!! "
                        'txtCGCCPF = ""
                        'txtCGCCPF.SetFocus
                        Exit Sub
                  End Select
                  Else
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     'txtCGCCPF = ""
                     'txtCGCCPF.SetFocus
                     Exit Sub
               End If
            End If
            txtCGCCPF.PromptInclude = True
         End If
   End If
   TABAUX.Close
End Sub

Private Sub ATUALIZA_TOTAL_OS()
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   TOTAL_SERVIÇO_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0

   SQL = "select sum(valor_item*qtd_pedida) from ITEMREQ "
   SQL = SQL & "where numr_req = " & NUMR_OS
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABCONSULTA = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TABCONSULTA.EOF Then _
      If Not IsNull(TABCONSULTA.Fields(0).Value) Then _
         TOTAL_PEÇAS_N = TABCONSULTA.Fields(0).Value
   TABCONSULTA.Close

   SQL = "select sum(perc_desc) from ITEMREQ "
   SQL = SQL & " where numr_req = " & NUMR_OS
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABCONSULTA = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TABCONSULTA.EOF Then _
      If Not IsNull(TABCONSULTA.Fields(0).Value) Then _
         TOTAL_DESCONTO_PEÇAS_N = TOTAL_PEÇAS_N * TABCONSULTA.Fields(0).Value / 100
   TABCONSULTA.Close

   ABRE_BANCO_AUXILIAR

   SQL = "select sum(valor_tarefa) from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TABCONSULTA = DBARQAUX.OpenRecordset(SQL, 4)
   If Not TABCONSULTA.EOF Then _
      If Not IsNull(TABCONSULTA.Fields(0).Value) Then _
         TOTAL_SERVIÇO_N = TABCONSULTA.Fields(0).Value
   TABCONSULTA.Close

   SQL = "select sum(valor_desc_tarefa) from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TABCONSULTA = DBARQAUX.OpenRecordset(SQL, 4)
   If Not TABCONSULTA.EOF Then _
      If Not IsNull(TABCONSULTA.Fields(0).Value) Then _
         TOTAL_DESCONTO_SERVIÇO_N = TABCONSULTA.Fields(0).Value
   TABCONSULTA.Close

   VALOR_DESCONTO_N = 0
   SQL = "select valor_desconto from CABECAOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TABUSU = DBARQAUX.OpenRecordset(SQL, 4)
   If Not IsNull(TABUSU!Valor_Desconto) Then _
      VALOR_DESCONTO_N = TABUSU!Valor_Desconto

   txtTOTALOS.Text = Format((TOTAL_SERVIÇO_N + TOTAL_PEÇAS_N) - (TOTAL_DESCONTO_PEÇAS_N + TOTAL_DESCONTO_SERVIÇO_N + VALOR_DESCONTO_N), "fixed")
   txtTOTALOS.Refresh
   txtTOTALDESCONTOOS.Text = Format((TOTAL_DESCONTO_PEÇAS_N + TOTAL_DESCONTO_SERVIÇO_N + VALOR_DESCONTO_N), "fixed")
   txtTOTALDESCONTOOS.Refresh

   DBARQAUX.Close
End Sub
'================
Private Sub IMPRIMIR_OS()
   If txtOs.Text = "" Then _
      Exit Sub
   OBS_A = ""
   If txtTOTALOS.Text <> "" Then
      VALOR_TOTAL_N = txtTOTALOS.Text
      VALOR_ITEM_N = txtTOTALOS.Text
   End If
   VALOR_TOTAL_DESCONTO_N = 0
   frmDESCONTO.Show 1

   ABRE_BANCO_AUXILIAR
   SQL = "update CABECAOS set "
   SQL = SQL & " valor_desconto = '" & VALOR_TOTAL_DESCONTO_N & "'"
   SQL = SQL & " where numr_os = " & NUMR_OS
   'SQL = SQL & " and empresa_id = " & EMPRESA_ID
   DBARQAUX.Execute SQL

   SQL = "select * from CABECAOS "
   SQL = SQL & " where numr_os = " & NUMR_OS
   'SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABCABECA = DBARQAUX.OpenRecordset(SQL)
   If Not TABCABECA.EOF Then
      TABCABECA.Edit
         TABCABECA!Valor_Desconto = VALOR_TOTAL_DESCONTO_N
         'TABCABECA!EMPRESA_ID = EMPRESA_ID
      TABCABECA.Update
   End If
   TABCABECA.Close

   SQL = "update RELOS set "
   SQL = SQL & "valor_desconto = '" & Replace(VALOR_TOTAL_DESCONTO_N, ",", ".") & "'"
   SQL = SQL & ", obs = '" & OBS_A & "'"
   SQL = SQL & " where numr_os = " & NUMR_OS
   DBARQAUX.Execute SQL

   If txtDESCONTOPRODUTO.Text <> "" Or txtTOTALPRODUTO.Text <> "" Then
      SQL = "update CABECAREQ set "
      'SQL = SQL & "valor_desconto = '" & Replace(txtDESCONTOPRODUTO.Text, ",", ".") & "'"
      SQL = SQL & " valor_total = '" & Replace(txtTOTALPRODUTO.Text, ",", ".") & "'"
      SQL = SQL & " where numr_req = " & NUMR_OS
      SQL = SQL & " and empresa_id = " & EMPRESA_ID
      DBARQEMP.Execute SQL
   End If

   SQL = "select * from RELOS "
   SQL = SQL & " where numr_os = " & NUMR_OS
   Set TABCABECA = DBARQAUX.OpenRecordset(SQL)
   If Not TABCABECA.EOF Then
      TABCABECA.Edit
         TABCABECA!Valor_Desconto = VALOR_TOTAL_DESCONTO_N
      TABCABECA.Update
   End If
   TABCABECA.Close

   SQL = "update OBS set "
   SQL = SQL & " obs = '" & OBS_A & "'"
   SQL = SQL & " ,prop = " & NUMR_OS
   SQL = SQL & " ,seq = 1 "
   SQL = SQL & " where prop = " & NUMR_OS
   SQL = SQL & " and seq = 1 "
   DBARQEMP.Execute SQL

   SQL = "select * from OBS "
   SQL = SQL & " where prop = " & NUMR_OS
   Set TABCABECA = DBARQEMP.OpenRecordset(SQL)
   If Not TABCABECA.EOF Then
      TABCABECA.Edit
         TABCABECA!obs = OBS_A
         TABCABECA!prop = NUMR_OS
      TABCABECA.Update
   End If
   TABCABECA.Close

   SQL = "select cgc,razao_social,nome_fant,ie from EMPRESA "
   SQL = SQL & "where empresa_id = " & EMPRESA_ID
   Set TABEMP = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TABEMP.EOF Then
      txtCGC.PromptInclude = False
      txtCGCCPF.PromptInclude = False
         txtCGC.Text = TABEMP!CGC
      txtCGC.PromptInclude = True

      ABRE_BANCO_AUXILIAR
      
      SQL = "create table RELOS "
      SQL = SQL & "("
         SQL = SQL & "cgc text(20), razao_social text(80),nome_fant text(80),ie "
         SQL = SQL & "text(40),end_emp text(40),bairro_emp text(40),cep_emp text(10),"
         SQL = SQL & "cidade_uf_emp text(40), nome_cli text(60),end_cli text(80),"
         SQL = SQL & "descricao_veiculo text(40), cor text(10),ano_modelo text(10),motor text(10),"
         SQL = SQL & "chassi text(80),combustivel text(10),placa text(10),dt_abre text(20),"
         SQL = SQL & "km text(10),dt_fecha text(20),consultor text(30),obs  text(255), tipo_os text(30),"
         SQL = SQL & "valor_desconto double,numr_os long not null,status text(1), fone text(50)"
         'SQL = SQL & " constraint numr_os unique (CHAVE_numr_os)"
      SQL = SQL & ")"
'MsgBox SQL
      'DBARQAUX.Execute SQL

      SQL = "create table RELOSITEM "
      SQL = SQL & "("
         SQL = SQL & "numr_os long not null,codg_item text(10),desc_item text(50),valor_item double"
         SQL = SQL & ",desconto_item double,qtd long,tipo_item text(1)"
      SQL = SQL & ")"
      'DBARQAUX.Execute SQL

      SQL = "delete * from RELOS "
      SQL = SQL & "where numr_os = " & txtOs.Text
      DBARQAUX.Execute SQL

      GRAVA_CABECA_OS

      SQL = "select * from CABECAOS "
      SQL = SQL & "where numr_os = " & txtOs.Text
      Set TABTEMP = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TABTEMP.EOF Then
         SQL = "select * from RELOS "
         SQL = SQL & "where numr_os = " & TABTEMP!NUMR_OS
         Set TABAUX = DBARQAUX.OpenRecordset(SQL)
         If Not TABAUX.EOF Then
            TABAUX.Edit
            Else: TABAUX.AddNew
         End If
         TABAUX!CGC = txtCGC.Text
         TABAUX!razao_social = TABEMP!razao_social
         TABAUX!Nome_Fant = TABEMP!Nome_Fant
         TABAUX!IE = TABEMP!IE
         TABAUX!dt_abre = TABTEMP!dt_abertura
         TABAUX!dt_fecha = TABTEMP!dt_fechamento
         TABAUX!Valor_Desconto = TABTEMP!Valor_Desconto
         TABAUX!NUMR_OS = TABTEMP!NUMR_OS
         TABAUX!Status = TABTEMP!Status
         'TIPO OS
         SQL = "select * from DESCR "
         SQL = SQL & "where tipo_a = 'H' "
         SQL = SQL & "and codigo = " & TABTEMP!tipo_os
         Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABDESCR.EOF Then _
            TABAUX!tipo_os = Trim(TABDESCR!desc_a)
         TABDESCR.Close

         SQL = "select nome from USUARIO "
         SQL = SQL & "where codigo = " & TABTEMP!ct
         Set TABUSU = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABUSU.EOF Then _
            TABAUX!consultor = TABUSU!Nome
         TABUSU.Close

         TABAUX!obs = OBS_A

         CRITERIO = ""
         SQL = "select * from FONE "
         SQL = SQL & "where prop = '" & TABEMP!CGC & "'"
         Set TABEND = DBARQEMP.OpenRecordset(SQL, 4)
         While Not TABEND.EOF
            If Not IsNull(TABEND!ddd) Then _
               CRITERIO = CRITERIO & "   (" & TABEND!ddd & ") "
            If Not IsNull(TABEND!Numero) Then _
               CRITERIO = CRITERIO & TABEND!Numero
            TABEND.MoveNext
         Wend
         TABEND.Close
         TABAUX!FONE = Trim(CRITERIO)

         'ENDEREÇO EMPRESA
         SQL = "select * from ENDERECO "
         SQL = SQL & "where prop = '" & TABEMP!CGC & "'"
         SQL = SQL & " and tipo = 'C' "
         Set TABEND = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABEND.EOF Then
            If Not IsNull(TABEND!Rua) Then _
               TABAUX!end_emp = TABEND!Rua
            If Not IsNull(TABAUX!end_emp) Then
               If Not IsNull(TABEND!Complemento) Then _
                  TABAUX!end_emp = TABAUX!end_emp & " ; " & TABEND!Complemento
               Else
                  If Not IsNull(TABEND!Complemento) Then _
                     TABAUX!end_emp = TABEND!Complemento
            End If
            If Not IsNull(TABEND!Bairro) Then _
               TABAUX!bairro_emp = TABEND!Bairro
            If Not IsNull(TABEND!Cep) Then _
               TABAUX!cep_emp = TABEND!Cep
            If Not IsNull(TABEND!Cep) Then
               If TABEND!Cep <> "" Then
                  SQL = "select * from CEP "
                  SQL = SQL & "where cep = " & TABEND!Cep
                  Set TABCEP = DBARQEMP.OpenRecordset(SQL, 4)
                  If Not IsNull(TABCEP!Cidade) Then _
                     TABAUX!cidade_uf_emp = TABCEP!Cidade & " - " & TABCEP!uf
                  TABCEP.Close
               End If
            End If
         End If
         TABEND.Close

         'CLIENTE
         SQL = "select nome,cgccpf,razao_social from CLIENTE "
         SQL = SQL & "where cgccpf = '" & txtCGCCPF.Text & "'"
         Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABCLI.EOF Then
            CRITERIO = ""
            SQL = "select * from FONE "
            SQL = SQL & "where prop = '" & TABCLI!CGCCPF & "'"
            Set TABEND = DBARQEMP.OpenRecordset(SQL, 4)
            While Not TABEND.EOF
               If Not IsNull(TABEND!ddd) Then _
                  CRITERIO = CRITERIO & "   (" & TABEND!ddd & ") "
               If Not IsNull(TABEND!Numero) Then _
                  CRITERIO = CRITERIO & TABEND!Numero
               TABEND.MoveNext
            Wend
            TABEND.Close

            TABAUX!nome_cli = TABCLI!Nome
            If Not IsNull(TABCLI!razao_social) Then _
               If TABCLI!razao_social <> "" Then _
                  TABAUX!nome_cli = TABCLI!razao_social

            'ENDEREÇO CLIENTE
            SQL = "select * from ENDERECO "
            SQL = SQL & " where prop = '" & txtCGCCPF.Text & "'"
            If Len(TABCLI!CGCCPF) <= 11 Then
               SQL = SQL & " and tipo = 'R' "
               Else: SQL = SQL & " and tipo = 'C' "
            End If
             Set TABEND = DBARQEMP.OpenRecordset(SQL, 4)
             If Not TABEND.EOF Then
                If Not IsNull(TABEND!Rua) Then _
                   TABAUX!end_cli = TABEND!Rua
                If Not IsNull(TABEND!Complemento) Then _
                   TABAUX!end_cli = TABAUX!end_cli & " ; " & TABEND!Complemento
                If Not IsNull(TABEND!Bairro) Then _
                   TABAUX!end_cli = TABAUX!end_cli & " ; " & TABEND!Bairro
                If Not IsNull(TABEND!Cep) Then
                   If TABEND!Cep <> "" Then
                      TABAUX!end_cli = TABAUX!end_cli & " ; " & TABEND!Cep
                      SQL = "select * from CEP "
                      SQL = SQL & "where cep = " & TABEND!Cep
                      Set TABCEP = DBARQEMP.OpenRecordset(SQL, 4)
                      If Not TABCEP.EOF Then
                         If Not IsNull(TABCEP!Cidade) Then _
                            TABAUX!end_cli = TABAUX!end_cli & " ; " & TABCEP!Cidade & " - " & TABCEP!uf
                      End If
                      TABCEP.Close
                   End If
                End If
             End If
             TABEND.Close
         End If
         TABCLI.Close

         'VEICULO
         SQL = "select * from CHASSI "
         SQL = SQL & "where cgccpf = '" & txtCGCCPF.Text & "'"
         SQL = SQL & " and placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
         Set TABCONSULTA = DBARQAUX.OpenRecordset(SQL, 4)
         If Not TABCONSULTA.EOF Then
            TABAUX!descricao_veiculo = TABCONSULTA!Descricao
            TABAUX!ano_modelo = TABCONSULTA!ano & "/" & TABCONSULTA!Modelo
            TABAUX!motor = Left(TABCONSULTA!motor, 30)
            TABAUX!chassi = TABCONSULTA!nr_chassi
            TABAUX!placa = TABCONSULTA!placa
            TABAUX!km = txtKm.Text
            'TABAUX!km = TABCONSULTA!km_atual
            'COR
            SQL = "select * from DESCR "
            SQL = SQL & "where tipo_a = 'Q' "
            SQL = SQL & "and codigo = " & TABCONSULTA!cor
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then _
               TABAUX!cor = Trim(TABDESCR!desc_a)
            TABDESCR.Close
            'COMBUSTIVEL
            SQL = "select * from DESCR "
            SQL = SQL & "where tipo_a = 'S' "
            SQL = SQL & "and codigo = " & TABCONSULTA!combustivel
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then _
               TABAUX!combustivel = Trim(TABDESCR!desc_a)
            TABDESCR.Close
         End If
         TABCONSULTA.Close
         TABAUX!FONE_cli = Trim(Left(CRITERIO, 50))
         TABAUX.Update
         TABAUX.Close
         
         'item serviço
         SQL = "select * from ITEMOS "
         SQL = SQL & "where numr_os = " & NUMR_OS
         Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
         While Not TABAUX.EOF
            SQL = "select * from RELOSITEM "
            SQL = SQL & "where numr_os = " & NUMR_OS
            SQL = SQL & " and tipo_item = 'S'" 'serviço
            SQL = SQL & " and codg_item = '" & TABAUX!Codg_tarefa & "'"
            Set TABCONSULTA = DBARQAUX.OpenRecordset(SQL)
            If Not TABCONSULTA.EOF Then
               TABCONSULTA.Edit
               Else: TABCONSULTA.AddNew
            End If

            TABCONSULTA!NUMR_OS = NUMR_OS
            TABCONSULTA!codg_item = TABAUX!Codg_tarefa

            SQL = "select descricao from TAREFA "
            SQL = SQL & "where codg_tarefa = '" & TABAUX!Codg_tarefa & "'"
            Set TABDESCR = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then _
               If Not IsNull(TABDESCR!Descricao) Then _
                  TABCONSULTA!desc_item = Left(TABDESCR!Descricao, 50)
            TABDESCR.Close
            
            SQL = "select nome from USUARIO "
            SQL = SQL & " where codigo = " & TABAUX!codg_mecanico
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABUSU = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABUSU.EOF Then _
               If Not IsNull(TABUSU!Nome) Then _
                  TABCONSULTA!mecanico = TABUSU!Nome
            TABUSU.Close

            TABCONSULTA!Valor_Item = TABAUX!valor_tarefa
            TABCONSULTA!desconto_item = TABAUX!valor_desc_tarefa
            TABCONSULTA!QTD = 1
            TABCONSULTA!tipo_item = "S"

            TABCONSULTA.Update
            TABCONSULTA.Close
            TABAUX.MoveNext
         Wend
         TABAUX.Close

            SQL = "select * from CABECAreq "
            SQL = SQL & " where numr_req = " & NUMR_OS
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABCABECA = DBARQEMP.OpenRecordset(SQL)
            If Not TABCABECA.EOF Then
               TOTAL_PEÇAS_N = 0
               SQL = "select sum(valor_item*qtd_pedida) from ITEMREQ "
               SQL = SQL & " where numr_req = " & NUMR_OS
               SQL = SQL & " and empresa_id = " & EMPRESA_ID
               Set TABCONSULTA = DBARQEMP.OpenRecordset(SQL, 4)
               If Not TABCONSULTA.EOF Then _
                  If Not IsNull(TABCONSULTA.Fields(0).Value) Then _
                     TOTAL_PEÇAS_N = TABCONSULTA.Fields(0).Value
               TABCONSULTA.Close

               TABCABECA.Edit
                  TABCABECA!Valor_Desconto = VALOR_TOTAL_DESCONTO_N
                  TABCABECA!valor_total = TOTAL_PEÇAS_N
               TABCABECA.Update
            End If
            TABCABECA.Close
         
         'item peça
         SQL = "select * from ITEMREQ "
         SQL = SQL & " where numr_req = " & NUMR_OS
         SQL = SQL & " and empresa_id = " & EMPRESA_ID
         Set TABAUX = DBARQEMP.OpenRecordset(SQL, 4)
         While Not TABAUX.EOF
            SQL = "select * from RELOSITEM "
            SQL = SQL & "where numr_os = " & NUMR_OS
            SQL = SQL & " and tipo_item = 'P'" 'peças
            SQL = SQL & " and codg_item = '" & TABAUX!Codg_Prod & "'"
            Set TABCONSULTA = DBARQAUX.OpenRecordset(SQL)
            If Not TABCONSULTA.EOF Then
               TABCONSULTA.Edit
               Else: TABCONSULTA.AddNew
            End If

            TABCONSULTA!NUMR_OS = NUMR_OS
            TABCONSULTA!codg_item = TABAUX!Codg_Prod

            SQL = "select descricao from PRODUTO "
            SQL = SQL & " where codg_prod = '" & TABAUX!Codg_Prod & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then _
               If Not IsNull(TABDESCR!Descricao) Then _
                  TABCONSULTA!desc_item = Left(TABDESCR!Descricao, 50)
            TABDESCR.Close

            TABCONSULTA!Valor_Item = TABAUX!Valor_Item
            TABCONSULTA!desconto_item = (TABAUX!Valor_Item * TABAUX!qtd_pedida) * TABAUX!PERC_desc / 100
            TABCONSULTA!QTD = TABAUX!qtd_pedida
            TABCONSULTA!tipo_item = "P"

            TABCONSULTA.Update
            TABCONSULTA.Close
            TABAUX.MoveNext
         Wend
         TABAUX.Close
      End If
      TABTEMP.Close
      DBARQAUX.Close

      frmINICIO.DIALOGO.CancelError = True
      On Error GoTo TRATAERRO4
      'mostra a janela para impressora
      frmINICIO.DIALOGO.ShowPrinter

      'IMPRIMIR RELATÓRIO
      frmINICIO.RELOS.SelectionFormula = ""
      frmINICIO.RELOS.Destination = 0
      frmINICIO.RELOS.SelectionFormula = "{RELOS.numr_os} = " & txtOs.Text
      frmINICIO.RELOS.ReportFileName = PATH_REL & "rel_abre_os.rpt"
      frmINICIO.RELOS.Action = 1

      SQL = "delete * from RELOS "
      SQL = SQL & "where numr_os = " & txtOs.Text
      DBARQAUX.Execute SQL

TRATAERRO4:
      Else: MsgBox "Empresa não cadastrada."
   End If
   TABEMP.Close
End Sub

Private Sub CABEÇALHO_IMPRESSÃO()
   Printer.Font = frmINICIO.DIALOGO.FontSize
   Print #1, "----------------------------------------------------------------------------------------------"
   Print #1, TABEMP!Nome_Fant
   frmABREOS.txtCGC.PromptInclude = False
      frmABREOS.txtCGC.Text = TABEMP!CGC
   frmABREOS.txtCGC.PromptInclude = True
   CRITERIO = "Insc.Estadual: " & Trim(TABEMP!IE)

   Print #1, Trim(TABEMP!razao_social); " - "; "CNPJ: "; Trim(frmABREOS.txtCGC.Text); " - "; Trim(CRITERIO)
   Print #1, "----------------------------------------------------------------------------------------------"
   'procura endereço empresa
   SQL = "select * from ENDERECO e, CEP p "
   SQL = SQL & "where e.prop = '" & TABEMP!CGC & "'"
   SQL = SQL & " and e.cep = p.cep "
   SQL = SQL & " e.tipo = 'C' "
   Set TABEND = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TABEND.EOF Then
      Print #1, Trim(TABEND!Rua); ", "; Trim(TABEND!Complemento); ", "; Trim(TABEND!Bairro); ", "; _
      Trim(TABEND.Fields("e.cep")); " "; Trim(TABEND!Cidade); " - "; Trim(TABEND!uf)
   End If
   TABEND.Close

   CRITERIO = ""
   SQL = "select * from FONE "
   SQL = SQL & "where prop = '" & TABEMP!CGC & "'"
   Set TABEND = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TABEND.EOF
      If Not IsNull(TABEND!Numero) Then _
         CRITERIO = "(" & TABEND!ddd & ") "
      If Not IsNull(TABEND!Numero) Then _
         CRITERIO = CRITERIO & TABEND!Numero
      TABEND.MoveNext
   Wend
   TABEND.Close
   CRITERIO = "Telefax: " & CRITERIO
   Print #1,
      Printer.Font = frmINICIO.DIALOGO.FontSize + 2
      Print #1, Spc(20); CRITERIO
   Print #1, "----------------------------------------------------------------------------------------------"
   Print #1,
   Print #1, Spc(10); "      LANTERNAGEM - PINTURA - ELETRICIDADE - INJEÇÃO"
   Print #1, Spc(10); "MECÂNICA EM GERAL - SUSPENSÃO - ALINHAMENTO - BALANCEAMENTO"
   Close #1
   LoadEXE ("C:\Arquivos de programas\Acessórios\WORDPAD.EXE c:\texte.txt")
End Sub
