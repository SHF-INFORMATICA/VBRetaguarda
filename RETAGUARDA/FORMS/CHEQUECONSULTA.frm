VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCHEQUECONSULTA 
   Caption         =   "Consulta Cheque"
   ClientHeight    =   8010
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CHEQUECONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbEstabAUX 
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
      Left            =   4200
      TabIndex        =   48
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbEstab 
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
      Left            =   4200
      TabIndex        =   46
      Top             =   7320
      Width           =   3015
   End
   Begin VB.ComboBox cmbCampo 
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
      Left            =   6600
      TabIndex        =   45
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox cmbTipoOrdem 
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
      Left            =   5280
      TabIndex        =   44
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdProp 
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
      Left            =   8280
      Picture         =   "CHEQUECONSULTA.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2280
      Width           =   405
   End
   Begin VB.CommandButton cmdTerc 
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
      Left            =   2520
      Picture         =   "CHEQUECONSULTA.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2280
      Width           =   405
   End
   Begin Threed.SSCommand cmdPagamento 
      Height          =   735
      Left            =   120
      TabIndex        =   35
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      _Version        =   262144
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CHEQUECONSULTA.frx":7016
      Caption         =   "Pagamentos"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin VB.ComboBox cmbContaAUX 
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
      Left            =   10080
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbAgenciaAUX 
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
      Left            =   9840
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
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
      Height          =   1215
      Left            =   4440
      TabIndex        =   24
      Top             =   600
      Width           =   4335
      Begin VB.OptionButton optDtCad 
         Caption         =   "Por Dt.Cad."
         Height          =   240
         Left            =   2760
         TabIndex        =   34
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optResp 
         Caption         =   "Por Resp."
         Height          =   240
         Left            =   1440
         TabIndex        =   33
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   360
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   360
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Relat:"
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   1035
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   4320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbldtEmis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial:"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbldtDep 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final:"
         Height          =   240
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   540
      End
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
      Left            =   10320
      TabIndex        =   22
      Top             =   840
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
      Left            =   9600
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
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
      Left            =   9600
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
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
      Left            =   9600
      TabIndex        =   0
      Top             =   840
      Width           =   1935
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
      Height          =   1455
      Left            =   50
      TabIndex        =   10
      Top             =   600
      Width           =   4335
      Begin VB.ComboBox cmbStatusAux 
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
         Left            =   1440
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbStatus 
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
         Left            =   1200
         TabIndex        =   30
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtValor 
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
         Left            =   3000
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtCheque 
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtSERIE 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Situação:"
         Height          =   240
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   240
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cheque:"
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série:"
         Height          =   240
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   570
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   480
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
            Picture         =   "CHEQUECONSULTA.frx":7330
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":7784
            Key             =   "relampago"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":7AA0
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":7EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":8348
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":8668
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":8ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":8DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":9230
            Key             =   "fechar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHEQUECONSULTA.frx":A9C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Incluir"
            Key             =   "incluir"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "imp"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   9600
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9840
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECONSULTA.frx":C6D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECONSULTA.frx":D86A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECONSULTA.frx":E8F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECONSULTA.frx":FA04
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECONSULTA.frx":109B9
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CHEQUECONSULTA.frx":12830
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF_PROP 
      Height          =   345
      Left            =   6480
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      _Version        =   393216
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
   Begin MSMask.MaskEdBox txtCPFCNPJ_TERCEIROS 
      Height          =   345
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      _Version        =   393216
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   10920
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11535
      DesignHeight    =   8010
   End
   Begin MSComctlLib.ListView lstCheque 
      Height          =   4215
      Left            =   0
      TabIndex        =   23
      Top             =   2760
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N.Cheque"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Série"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dt.Emissão"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dt.Depósito"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dt.Vencimento"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Responsável"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Repasse"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Proprietário"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Banco"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Agencia"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Conta"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "codbanco"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "cnpjcpfprop"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "cnpjcpfter"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Praça"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "CMC7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "repasse_id"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Estabelecimento"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Estabelecimento:"
      Height          =   240
      Left            =   2520
      TabIndex        =   47
      Top             =   7320
      Width           =   1635
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Ordem:"
      Height          =   240
      Left            =   4440
      TabIndex        =   43
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor = "
      Height          =   240
      Left            =   9120
      TabIndex        =   41
      Top             =   7560
      Width           =   750
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastrado(s) ="
      Height          =   240
      Left            =   8280
      TabIndex        =   40
      Top             =   7080
      Width           =   1470
   End
   Begin VB.Label lblTotReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   39
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label lblTotValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   38
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      X1              =   0
      X2              =   11640
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label lblResp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3000
      TabIndex        =   21
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terc.:"
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   255
      Left            =   8880
      TabIndex        =   19
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Agenc:"
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   255
      Left            =   8880
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prop.:"
      Height          =   240
      Left            =   5760
      TabIndex        =   16
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label lblProp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmCHEQUECONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim LinX          As Integer
   Dim TamX          As Double
   Dim TamY          As Double
   Dim strImagem     As String
   Dim PESSOA_ID_N   As Long
   Dim QTD_COTAS     As Integer

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA
   
   Call CentralizaJanela(frmCHEQUECONSULTA)

   cmbBanco.Clear
   cmbBancoAux.Clear

   If TabBANCO.State = 1 Then _
      TabBANCO.Close

   SQL = "select * from Banco "
   SQL = SQL & " order by nome_banco"
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabBANCO.EOF
      cmbBanco.AddItem Trim(TabBANCO!Nome_Banco)
      cmbBancoAux.AddItem TabBANCO.Fields("banco_id").Value
      TabBANCO.MoveNext
   Wend
   If TabBANCO.State = 1 Then _
      TabBANCO.Close

'==========================================
   cmbSTATUS.Clear

   cmbSTATUS.AddItem "Todos"
   cmbStatusAux.AddItem ""   'todos

   cmbSTATUS.AddItem "Cadastrado(s)"
   cmbStatusAux.AddItem "E"   'cadastrado

   cmbSTATUS.AddItem "Depositado(s)"
   cmbStatusAux.AddItem "D"   'depositado

   cmbSTATUS.AddItem "Compensado(s)"
   cmbStatusAux.AddItem "P"   'processado, compensado

   cmbSTATUS.Text = "Cadastrado(s)"
   cmbStatusAux.Text = "E"

   cmbTipoOrdem.Clear
   cmbTipoOrdem.AddItem "Crescente"
   cmbTipoOrdem.AddItem "Decrescente"
   cmbTipoOrdem.Text = "Crescente"

   cmbCampo.Clear
   cmbCampo.AddItem "Vencimento"
   cmbCampo.AddItem "Cadastro"
   cmbCampo.AddItem "Depósito"
   cmbCampo.AddItem "Valor"
   cmbCampo.Text = "Vencimento"

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True
'==========================================
   cmbEstabAUX.Clear
   cmbEstab.Clear
   cmbEstab.AddItem "Todos"
   cmbEstabAUX.AddItem ""

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO "
   SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbEstab.AddItem Trim(TabDESCR!DESCRICAO) & "-" & Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      cmbEstabAUX.AddItem Trim(TabDESCR.Fields("ESTABELECIMENTO_id").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstab.Enabled = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
      cmbEstab.Enabled = True

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

Private Sub lstcheque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   OrdenaListView lstCheque, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         Dim i As Integer

         INDR_PRI = True

         For i = lstCheque.ListItems.Count To 1 Step -1

            If lstCheque.ListItems(i).Checked = True Then

               If INDR_PRI = True Then
                  INDR_PRI = False

                  Msg = "Confirma exclusão dos cheque(s) selecionado(s) ?"
                  PERGUNTA Msg, vbYesNo + 32, "Cheque", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     SQL = "delete from CHEQUE "
                     SQL = SQL & " where cheque_id = " & lstCheque.ListItems(i).SubItems(16)
                     CONECTA_RETAGUARDA.Execute SQL
                  End If

               End If

            End If

         Next i
         If INDR_PRI = False Then _
            MONTA_CONSULTA
      Case "incluir"
         frmCHEQUECADASTRO.Show 1
         MONTA_CONSULTA
      Case "imp"
         MONTA_REL
      Case "voltar"
         Unload Me
      Case "limpar"
         Limpar_Tudo
         txtCPFCNPJ_TERCEIROS.SetFocus
      Case "print"
      Case "consultar"
         MONTA_CONSULTA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub LSTCHEQUE_DblClick()
   If Not IsNull(lstCheque.SelectedItem.Text) Then _
      CHAMA_TELA_CADASTRO
End Sub

Private Sub cmbestab_Click()
On Error Resume Next

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex
End Sub

Private Sub TXTcnpjcpf_prop_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTcnpjcpf_prop_KeyPress"
End Sub

Private Sub cmdProp_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF_PROP.PromptInclude = False
         txtCNPJCPF_PROP.Text = CNPJCPF_A
      txtCNPJCPF_PROP.PromptInclude = True
   End If
   txtCNPJCPF_PROP.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdProp_Click"
End Sub

Private Sub TXTcnpjcpf_prop_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_PROP.PromptInclude = False
   CRITERIO_A = txtCNPJCPF_PROP.Text
   If txtCNPJCPF_PROP.Text <> "" Then
      CRITERIO_A = txtCNPJCPF_PROP.Text
      If Not IsNull(txtCNPJCPF_PROP.Text) Then
         If Len(txtCNPJCPF_PROP.Text) <= 11 Then
            txtCNPJCPF_PROP.Mask = "###.###.###-##"
            Else: txtCNPJCPF_PROP.Mask = "##.###.###/####-##"
         End If
      End If
      txtCNPJCPF_PROP.Text = CRITERIO_A
      txtCNPJCPF_PROP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTcnpjcpf_prop_GotFocus"
End Sub

Private Sub txtCPFCNPJ_TERCEIROS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCPFCNPJ_TERCEIROS.PromptInclude = False
      CRITERIO_A = txtCPFCNPJ_TERCEIROS.Text
      If txtCPFCNPJ_TERCEIROS.Text <> "" Then
         CRITERIO_A = txtCPFCNPJ_TERCEIROS.Text
         If Not IsNull(txtCPFCNPJ_TERCEIROS.Text) Then
            If Len(txtCPFCNPJ_TERCEIROS.Text) <= 11 Then
               txtCPFCNPJ_TERCEIROS.Mask = "###.###.###-##"
               Else: txtCPFCNPJ_TERCEIROS.Mask = "##.###.###/####-##"
            End If
         End If
         txtCPFCNPJ_TERCEIROS.Text = CRITERIO_A
      
         txtCPFCNPJ_TERCEIROS.PromptInclude = False
         If Trim(txtCPFCNPJ_TERCEIROS.Text) <> "" Then
            PROCURA_PORTADOR Trim(txtCPFCNPJ_TERCEIROS.Text)
            lblResp.Caption = NOME_A
         End If
         txtCPFCNPJ_TERCEIROS.PromptInclude = True
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCPFCNPJ_TERCEIROS_KeyPress"
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If
Exit Sub

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCheque_KeyPress"
End Sub

Private Sub txtCNPJCPF_PROP_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF_PROP.PromptInclude = False
               txtCNPJCPF_PROP.Text = CNPJCPF_A
            txtCNPJCPF_PROP.PromptInclude = True
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_PROP_KeyDown"
End Sub

Private Sub txtCNPJCPF_PROP_LostFocus()
   txtCNPJCPF_PROP.PromptInclude = False
   If Trim(txtCNPJCPF_PROP.Text) <> "" Then
      PROCURA_PORTADOR Trim(txtCNPJCPF_PROP.Text)
      lblProp.Caption = NOME_A
   End If
   txtCNPJCPF_PROP.PromptInclude = True
End Sub

Private Sub cmdTerc_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCPFCNPJ_TERCEIROS.PromptInclude = False
         txtCPFCNPJ_TERCEIROS.Text = CNPJCPF_A
      txtCPFCNPJ_TERCEIROS.PromptInclude = True
   End If
   txtCPFCNPJ_TERCEIROS.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdTerc_Click"
End Sub

Private Sub txtCPFCNPJ_TERCEIROS_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "CLIENTE"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCPFCNPJ_TERCEIROS.PromptInclude = False
               txtCPFCNPJ_TERCEIROS.Text = CNPJCPF_A
            txtCPFCNPJ_TERCEIROS.PromptInclude = True
         End If
         txtCPFCNPJ_TERCEIROS.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCPFCNPJ_TERCEIROS_KeyDown"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtfim_GotFocus"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCheque.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtfim_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtini_GotFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtini_KeyPress"
End Sub

Private Sub lblResp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lblResp_KeyPress"
End Sub

Private Sub txtDtIni_LostFocus()
'On Error GoTo ERRO_TRATA

   If Not IsDate(txtDtIni.Text) Then
      MsgBox "Data Informada Inválida !!!"
      txtDtIni.PromptInclude = False
         txtDtIni.Text = Date
      txtDtIni.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtini_LostFocus"
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyPress"
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtserie_KeyPress"
End Sub

Private Sub Procura_Cliente()
'On Error GoTo ERRO_TRATA

   txtCPFCNPJ_TERCEIROS.PromptInclude = False

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from CLIENTE "
   SQL = SQL & " where cgccpf='" & txtCPFCNPJ_TERCEIROS.Text & "'"
   SQL = SQL & " and status = 'A'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      lblResp.Caption = "" & TabCliente!NOME
      cmbBanco.SetFocus
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Procura_Cliente"
End Sub

Private Sub cmbAgencia_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(cmbAgencia.Text) Then _
      If Trim(cmbAgencia.Text) <> "" Then _
         cmbAgenciaAUX.ListIndex = cmbAgencia.ListIndex

   cmbConta.Clear

   SQL = "select * from CONTA "
   SQL = SQL & " where agencia_id = " & cmbAgenciaAUX.Text
   TabCONTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCONTA.EOF
      cmbConta.AddItem Trim(TabCONTA!NUMR_CONTA)
      cmbContaAUX.AddItem TabCONTA.Fields("conta_id").Value
      TabCONTA.MoveNext
   Wend
   If TabCONTA.State = 1 Then _
      TabCONTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAgencia_Click"
End Sub

Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbagencia_KeyPress"
End Sub

Private Sub cmbBanco_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(cmbBanco.Text) Then _
      If Trim(cmbBanco.Text) <> "" Then _
         cmbBancoAux.ListIndex = cmbBanco.ListIndex

   If TabBANCO.State = 1 Then _
      TabBANCO.Close

   cmbAgencia.Clear
   cmbConta.Clear

   SQL = "select * from Agencia "
   SQL = SQL & " where banco_id = " & cmbBancoAux.Text
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabBANCO.EOF
      cmbAgencia.AddItem Trim(TabBANCO!NUMR_AGENCIA) & " - " & Trim(TabBANCO.Fields("nome_agencia").Value)
      cmbAgenciaAUX.AddItem TabBANCO.Fields("agencia_id").Value
      TabBANCO.MoveNext
   Wend
   If TabBANCO.State = 1 Then _
      TabBANCO.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_Click"
End Sub

Private Sub cmbbanco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbbanco_KeyPress"
End Sub

Private Sub cmbConta_Change()
'On Error GoTo ERRO_TRATA

   cmbConta_Click

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbConta_Change"
End Sub

Private Sub cmbConta_Click()
'On Error GoTo ERRO_TRATA

   If TabCONTA.State = 1 Then _
      TabCONTA.Close

   SQL = "select * from CONTA "
   SQL = SQL & " where NUMR_CONTA='" & cmbConta.Text & "'"
   TabCONTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCONTA.EOF Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select Nome, cgccpf, Codigo from CLIENTE "
      SQL = SQL & " where cliente_id = " & TabCONTA.Fields("CLIENTE_ID").Value
      SQL = SQL & " and status = 'A'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         txtCNPJCPF_PROP.PromptInclude = False
            txtCNPJCPF_PROP.Text = TabCliente!CGCCPF
         txtCNPJCPF_PROP.PromptInclude = True
         lblProp.Caption = TabCliente!NOME & ""
         'lblCGCCPF.Caption = TabCliente!CGCCPF & ""
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close
      Else
         txtCNPJCPF_PROP.PromptInclude = False
            txtCNPJCPF_PROP.Text = ""
         txtCNPJCPF_PROP.PromptInclude = True
         lblProp.Caption = ""
   End If
   If TabCONTA.State = 1 Then _
      TabCONTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbConta_Click"
End Sub

Private Sub cmbConta_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbConta_KeyPress"
End Sub

Private Sub cmbStatus_Click()
   If Trim(cmbSTATUS.Text) <> "" Then _
      cmbStatusAux.ListIndex = cmbSTATUS.ListIndex
End Sub

Private Sub cmdPagamento_Click()
'On Error GoTo ERRO_TRATA

   Dim i                As Integer
   Dim Selecao_Cheque   As String
   Dim INDR_VAI         As Boolean

   Selecao_Cheque = ""
   CRITERIO_A = ""
   INDR_VAI = False
   frmCHEQUEPAGTO.lstCheque.ListItems.Clear

   For i = lstCheque.ListItems.Count To 1 Step -1
      If lstCheque.ListItems(i).Checked = True Then
         Set item = frmCHEQUEPAGTO.lstCheque.ListItems.Add(, "a" & lstCheque.ListItems(i).SubItems(16), Trim(lstCheque.ListItems(i).Text))   'N.Cheque
         item.SubItems(1) = "" & lstCheque.ListItems(i).SubItems(2)  'Valor
         item.SubItems(2) = "" & lstCheque.ListItems(i).SubItems(5)  'Dt.Vencimento
         item.SubItems(3) = "" & lstCheque.ListItems(i).SubItems(9)  'banco
         item.SubItems(4) = "" & lstCheque.ListItems(i).SubItems(7)  'repasse
         item.SubItems(5) = "" & lstCheque.ListItems(i).SubItems(11) 'status
         item.SubItems(6) = "" & lstCheque.ListItems(i).SubItems(15) 'cnpjcpfter
         item.SubItems(7) = "" & lstCheque.ListItems(i).SubItems(16) 'CHEQUE_ID
         item.SubItems(8) = "" & lstCheque.ListItems(i).SubItems(19) 'repasse_ID
         item.SubItems(9) = "" & lstCheque.ListItems(i).SubItems(20) 'estabelecimento_id

         item.Checked = lstCheque.ListItems(i).Checked
         INDR_VAI = True
      End If
   Next i

   If INDR_VAI = True Then
      frmCHEQUEPAGTO.Show 1
      MONTA_CONSULTA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdPagamento_Click"
End Sub
'===========
Private Sub Limpar_Tudo()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstab.Text = ESTABELECIMENTO_ID_N
   lblTotReg.Caption = ""
   lblTotValor.Caption = ""
   cmbSTATUS.Text = ""
   cmbStatusAux.Text = ""
   lblResp.Caption = ""
   lstCheque.ListItems.Clear
   cmbBanco.Text = ""
   cmbAgencia.Text = ""
   cmbAgencia.Text = ""
   cmbAgenciaAUX.Text = ""
   cmbContaAUX.Text = ""
   cmbBancoAux.Text = ""
   cmbConta.Text = ""
   cmbConta.Text = ""
   txtCNPJCPF_PROP.PromptInclude = False
      txtCNPJCPF_PROP.Text = ""
      txtCNPJCPF_PROP.Mask = "##############"
   txtCNPJCPF_PROP.PromptInclude = True
   lblProp.Caption = ""
   txtCPFCNPJ_TERCEIROS.PromptInclude = False
      txtCPFCNPJ_TERCEIROS.Text = ""
      txtCPFCNPJ_TERCEIROS.Mask = "##############"
   txtCPFCNPJ_TERCEIROS.PromptInclude = True

   txtCheque.Text = ""
   txtSerie.Text = ""
   txtValor.Text = ""
   txtDtIni.PromptInclude = False
      txtDtIni.Text = ""
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = False
      txtDtFim.Text = ""
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Limpar_Tudo"
End Sub
            
Private Sub MONTA_CONSULTA()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "Aguarde, Pesquisando ...", "", "", "", ""

   HORA_INI = Time
   PESSOA_ID_N = 0
   VALOR_TOTAL_N = 0
   CONT_N = 0

   SQL = "select * from vwRelCheque "
   SQL = SQL & " where numr_cheque <> ''"

   CRITERIO_A = "{vwRelCheque.numr_cheque} <> '' "

   If Trim(cmbEstabAUX.Text) <> "" Then
      SQL = SQL & " and ESTABELECIMENTO_ID = " & ESTABELECIMENTO_ID_N
      CRITERIO_A = "{vwRelCheque.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   End If

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   If IsDate(txtDtIni.Text) Then
      If IsDate(txtDtFim.Text) Then

         If cmbStatusAux.Text = "E" Then
            SQL = SQL & " and dt_emissao >= '" & DMA(txtDtIni.Text) & "'"
            SQL = SQL & " and dt_emissao <= '" & DMA(txtDtFim.Text) & "'"

            CRITERIO_A = CRITERIO_A & " and {vwRelCheque.DT_EMISSAO} >= date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
            CRITERIO_A = CRITERIO_A & " and {vwRelCheque.DT_EMISSAO} <= date  (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
         End If

         If cmbStatusAux.Text = "D" Then
            SQL = SQL & " and dt_deposito >= '" & DMA(txtDtIni.Text) & "'"
            SQL = SQL & " and dt_deposito <= '" & DMA(txtDtFim.Text) & "'"

            CRITERIO_A = CRITERIO_A & " and {vwRelCheque.DT_deposito} >= date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
            CRITERIO_A = CRITERIO_A & " and {vwRelCheque.DT_deposito} <= date  (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
         End If

         If cmbStatusAux.Text = "P" Then
            SQL = SQL & " and dt_compensa >= '" & DMA(txtDtIni.Text) & "'"
            SQL = SQL & " and dt_compensa <= '" & DMA(txtDtFim.Text) & "'"

            CRITERIO_A = CRITERIO_A & " and {vwRelCheque.DT_compensa} >= date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
            CRITERIO_A = CRITERIO_A & " and {vwRelCheque.DT_compensa} <= date  (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
         End If
         Else
            If Trim(cmbStatusAux.Text) <> "" Then
               SQL = SQL & " and status = '" & Trim(cmbStatusAux.Text) & "'"
               CRITERIO_A = CRITERIO_A & " and {vwRelCheque.status} = '" & Trim(cmbStatusAux.Text) & "'"
            End If
      End If
   End If

   If Trim(cmbBancoAux.Text) <> "" Then
      SQL = SQL & " and banco_id = " & cmbBancoAux.Text
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.BANCO_id} = " & cmbBancoAux.Text
   End If

   If Trim(cmbAgenciaAUX.Text) <> "" Then
      SQL = SQL & " and agencia_id = " & cmbAgenciaAUX.Text
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.agencia_id} = '" & cmbAgenciaAUX.Text & "'"
   End If

   If Trim(cmbContaAUX.Text) <> "" Then
      SQL = SQL & " and conta_id = " & cmbContaAUX.Text
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.conta_id} = '" & cmbContaAUX.Text & "'"
   End If

   txtCNPJCPF_PROP.PromptInclude = False
   If txtCNPJCPF_PROP.Text <> "" Then
      SQL = SQL & " and CNPJCPF_Prop = '" & Trim(txtCNPJCPF_PROP.Text) & "'"
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.CNPJCPF_Prop} = '" & Trim(txtCNPJCPF_PROP.Text) & "'"
   End If

   txtCPFCNPJ_TERCEIROS.PromptInclude = False
   If txtCPFCNPJ_TERCEIROS.Text <> "" Then
      SQL = SQL & " and CNPJCPF_TERC = '" & Trim(txtCPFCNPJ_TERCEIROS.Text) & "'"
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.CNPJCPF_TERC} = '" & Trim(txtCPFCNPJ_TERCEIROS.Text) & "'"
   End If

   If Trim(txtCheque.Text) <> "" Then
      SQL = SQL & " and numr_cheque = '" & txtCheque.Text & "'"
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.numr_cheque} = '" & txtCheque.Text & "'"
   End If

   If Trim(txtSerie.Text) <> "" Then
      SQL = SQL & " and serie_cheque = '" & txtSerie.Text & "'"
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.serie_cheque} = '" & txtSerie.Text & "'"
   End If

   If txtValor <> "" Then
      VALOR_ITEM_N = txtValor.Text
      SQL = SQL & " and valor >= " & tpMOEDA(VALOR_ITEM_N)
      SQL = SQL & " and valor <= " & tpMOEDA(VALOR_ITEM_N)

      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.valor} >= " & tpMOEDA(VALOR_ITEM_N)
      CRITERIO_A = CRITERIO_A & " and {vwRelCheque.valor} <= " & tpMOEDA(VALOR_ITEM_N)
   End If

   SQL3 = ""
   If Trim(UCase(cmbTipoOrdem.Text)) = UCase("Decrescente") Then _
      SQL3 = "desc"

   If Trim(UCase(cmbCampo.Text)) = UCase("Vencimento") Then _
      SqL2 = "dt_compensa"
   If Trim(UCase(cmbCampo.Text)) = UCase("Cadastro") Then _
      SqL2 = "dt_emissao"
   If Trim(UCase(cmbCampo.Text)) = UCase("Depósito") Then _
      SqL2 = "dt_deposito"
   If Trim(UCase(cmbCampo.Text)) = UCase("Valor") Then _
      SqL2 = "valor"

   SQL = SQL & " order by " & SqL2 & " " & SQL3

   CONT_N = 0
   lstCheque.ListItems.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      CONT_N = CONT_N + 1
      Set item = lstCheque.ListItems.Add(, "seq." & CONT_N, Trim(TabConsulta!NUMR_CHEQUE))

      item.SubItems(1) = "" & Trim(TabConsulta!SERIE_CHEQUE)
      item.SubItems(2) = "" & Format(TabConsulta!VALOR, strFormatacao2Digitos)

      If TabConsulta!DT_EMISSAO = "01/01/1900" Then
         item.SubItems(3) = ""
         Else: item.SubItems(3) = "" & TabConsulta!DT_EMISSAO
      End If
      If TabConsulta.Fields("DT_DEPOSITO").Value = "01/01/1900" Then
         item.SubItems(4) = ""                                                   'Dt.Depósito
         Else: item.SubItems(4) = "" & TabConsulta.Fields("DT_DEPOSITO").Value   'Dt.Depósito
      End If
      If TabConsulta.Fields("DT_COMPENSA").Value = "01/01/1900" Then
         item.SubItems(5) = ""                                                   'Dt.Compensa
         Else: item.SubItems(5) = "" & TabConsulta.Fields("DT_COMPENSA").Value   'Dt.Compensa
      End If

      item.SubItems(6) = "" & TabConsulta.Fields("NOME_TERC").Value              'Responsável

      If Not IsNull(TabConsulta.Fields("responsavel").Value) Then _
         If Trim(TabConsulta.Fields("responsavel").Value) <> "" Then _
            item.SubItems(6) = "" & TabConsulta.Fields("responsavel").Value      'Responsável

      item.SubItems(7) = "" & TabConsulta.Fields("repasse").Value                'repasse

      item.SubItems(8) = "" & TabConsulta.Fields("NOME_PROP").Value              'Proprietário

      item.SubItems(9) = "" & TabConsulta.Fields("NOME_BANCO").Value             'Banco
      item.SubItems(10) = "" & TabConsulta.Fields("NUMR_AGENCIA").Value          'Agencia
      item.SubItems(11) = "" & TabConsulta.Fields("NUMR_CONTA").Value            'Conta

      item.SubItems(12) = "" & TabConsulta.Fields("status").Value                'status
      item.SubItems(13) = "" & TabConsulta.Fields("NOME_BANCO").Value            'codgbanco
      item.SubItems(14) = "" & TabConsulta.Fields("CNPJCPF_PROP").Value          'CNPJCPF_Prop
      item.SubItems(15) = "" & TabConsulta.Fields("CNPJCPF_Terc").Value          'CNPJCPF_TERC
      item.SubItems(16) = "" & TabConsulta.Fields("CHEQUE_ID").Value             'CHEQUE_ID
      item.SubItems(17) = "" & TabConsulta.Fields("praça").Value                 'praça
      item.SubItems(18) = "" & TabConsulta.Fields("cmc7").Value                  'cmc7

      item.SubItems(19) = "" & TabConsulta.Fields("repasse_id").Value            'repasse_id
      item.SubItems(19) = "" & TabConsulta.Fields("estabelecimento_id").Value    'estabelecimento_id

      VALOR_TOTAL_N = VALOR_TOTAL_N + TabConsulta!VALOR

      lblTotReg.Caption = CONT_N
      lblTotValor.Caption = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   HORA_FIM = Time

   MOSTRA_RODAPE "ESC - Sair", "Duplo click para selecionar", "Duração da consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss"), "Total de Registros Encontrados = " & NUMR_CONSULTA_N, ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub

Private Sub MONTA_REL()
'On Error GoTo ERRO_TRATA

   HORA_INI = Time
   PESSOA_ID_N = 0

   MOSTRA_RODAPE "Aguarde, Pesquisando ...", "", "", "", ""

   FORMULA_REL = ""

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   FORMULA_REL = "{vwRelCheque.numr_cheque} <> '' "

   If Trim(cmbEstabAUX.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.estabelecimento_id} = " & ESTABELECIMENTO_ID_N

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If cmbStatusAux.Text = "E" Then
         FORMULA_REL = FORMULA_REL & " and {vwRelCheque.DT_EMISSAO} >= date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
         FORMULA_REL = FORMULA_REL & " and {vwRelCheque.DT_EMISSAO} <= date  (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
      End If

      If cmbStatusAux.Text = "D" Then
         FORMULA_REL = FORMULA_REL & " and {vwRelCheque.DT_deposito} >= date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
         FORMULA_REL = FORMULA_REL & " and {vwRelCheque.DT_deposito} <= date  (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
      End If

      If cmbStatusAux.Text = "P" Then
         FORMULA_REL = FORMULA_REL & " and {vwRelCheque.DT_compensa} >= date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ")"
         FORMULA_REL = FORMULA_REL & " and {vwRelCheque.DT_compensa} <= date  (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
      End If
      Else
         If Trim(cmbStatusAux.Text) <> "" Then
            FORMULA_REL = FORMULA_REL & " and {vwRelCheque.status} = '" & Trim(cmbStatusAux.Text) & "'"
         End If
   End If

   If Trim(cmbBancoAux.Text) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.BANCO_id} = " & cmbBancoAux.Text
   End If

   If Trim(cmbAgenciaAUX.Text) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.agencia_id} = '" & cmbAgenciaAUX.Text & "'"
   End If

   If Trim(cmbContaAUX.Text) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.conta_id} = '" & cmbContaAUX.Text & "'"
   End If

   txtCNPJCPF_PROP.PromptInclude = False
   If txtCNPJCPF_PROP.Text <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.CNPJCPF_Prop} = '" & Trim(txtCNPJCPF_PROP.Text) & "'"
   End If

   txtCPFCNPJ_TERCEIROS.PromptInclude = False
   If txtCPFCNPJ_TERCEIROS.Text <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.CNPJCPF_Terc} = '" & Trim(txtCPFCNPJ_TERCEIROS.Text) & "'"
   End If

   If Trim(txtCheque.Text) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.numr_cheque} = '" & txtCheque.Text & "'"
   End If

   If Trim(txtSerie.Text) <> "" Then
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.serie_cheque} = '" & txtSerie.Text & "'"
   End If

   If txtValor <> "" Then
      VALOR_ITEM_N = txtValor.Text
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.valor} >= " & tpMOEDA(VALOR_ITEM_N)
      FORMULA_REL = FORMULA_REL & " and {vwRelCheque.valor} <= " & tpMOEDA(VALOR_ITEM_N)
   End If

   If optResp.Value = True Then
      Nome_Relatorio = "Cheque_resp.rpt"
      Else: Nome_Relatorio = "Cheque_data.rpt"
   End If

   frmRELATORIO10.Show 1

   HORA_FIM = Time

   MOSTRA_RODAPE "ESC - Sair", "Duplo click para selecionar", "Duração da consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss"), "Total de Registros Encontrados = " & NUMR_CONSULTA_N, ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_REL"
End Sub

Sub CHAMA_TELA_CADASTRO()
On Error Resume Next

   NUMR_ID_N = 0

   frmCHEQUECADASTRO.txtCheque.Text = "" & lstCheque.SelectedItem.Text
   frmCHEQUECADASTRO.txtSerie.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(1).Text
   frmCHEQUECADASTRO.txtValor.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(2).Text

   frmCHEQUECADASTRO.txtDtEmis.PromptInclude = False
      frmCHEQUECADASTRO.txtDtEmis.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(3).Text
   frmCHEQUECADASTRO.txtDtEmis.PromptInclude = True

   frmCHEQUECADASTRO.txtDtDep.PromptInclude = False
      frmCHEQUECADASTRO.txtDtDep.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(4).Text
   frmCHEQUECADASTRO.txtDtDep.PromptInclude = True

   frmCHEQUECADASTRO.txtDtCompensa.PromptInclude = False
      frmCHEQUECADASTRO.txtDtCompensa.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(5).Text
   frmCHEQUECADASTRO.txtDtCompensa.PromptInclude = True

   frmCHEQUECADASTRO.txtProp.PromptInclude = False
      frmCHEQUECADASTRO.txtProp.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(14).Text
      frmCHEQUECADASTRO.lblProp.Caption = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(8).Text
   frmCHEQUECADASTRO.txtProp.PromptInclude = True

   frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = False
      frmCHEQUECADASTRO.txtPORTADOR.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(15).Text
      frmCHEQUECADASTRO.txtNome.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(6).Text
   frmCHEQUECADASTRO.txtPORTADOR.PromptInclude = True

   'frmCHEQUECADASTRO.txtCNPJCPF_REPASSE.PromptInclude = False
   '   If TabPessoa.State = 1 Then _
         TabPessoa.Close

   '   SQL = "select cnpjcpf from PESSOA"
   '   SQL = SQL & " where razao = '" & Trim(lstCheque.ListItems(lstCheque.selectedItem.Index).ListSubItems(7).Text) & "'"
   '   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   '   If Not TabPessoa.EOF Then _
         frmCHEQUECADASTRO.txtCNPJCPF_REPASSE.Text = "" & Trim(TabPessoa.Fields("cnpjcpf").Value)

   '   If TabPessoa.State = 1 Then _
         TabPessoa.Close

   '   frmCHEQUECADASTRO.txtRepasse.Text = "" & lstCheque.ListItems(lstCheque.selectedItem.Index).ListSubItems(7).Text
   'frmCHEQUECADASTRO.txtCNPJCPF_REPASSE.PromptInclude = True

   If Trim(lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(15).Text) <> "" Then
      frmCHEQUECADASTRO.txtCNPJCPF_REPASSE.PromptInclude = False

      If Trim(lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(15).Text) <> "" Then _
         frmCHEQUECADASTRO.txtRepasse.Text = frmCHEQUECADASTRO.PROCURA_REPASSE(Trim(Trim(lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(15).Text)))

      frmCHEQUECADASTRO.txtCNPJCPF_REPASSE.PromptInclude = True
   End If

   frmCHEQUECADASTRO.cmbBancoAux.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(13).Text
   frmCHEQUECADASTRO.cmbBanco.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(9).Text

   frmCHEQUECADASTRO.cmbAgenciaAUX.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(7).Text
   frmCHEQUECADASTRO.cmbAgencia.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(10).Text

   frmCHEQUECADASTRO.cmbContaAUX.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(11).Text
   frmCHEQUECADASTRO.cmbConta.Text = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(11).Text

   frmCHEQUECADASTRO.lblID.Caption = "" & lstCheque.ListItems(lstCheque.SelectedItem.Index).ListSubItems(16).Text

   Call frmCHEQUECADASTRO.MOSTRA_CHEQUE

   frmCHEQUECADASTRO.Show 1

   NUMR_ID_N = 0

   MONTA_CONSULTA
End Sub

Function PROCURA_PORTADOR(CNPJCPF_A As String) As Boolean
   PROCURA_PORTADOR = False
   NOME_A = ""
    '= 0

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select * from PESSOA"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then
      NOME_A = TabPessoa.Fields("descricao").Value
'       = TabPessoa.Fields("pessoa_id").Value
      Else
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         MsgBox "CNPJ/CPF não encontrado"
         Exit Function
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   PROCURA_PORTADOR = True
End Function
