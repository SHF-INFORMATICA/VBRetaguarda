VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNOTAENTRADA 
   Caption         =   "Nota Fiscal Entrada"
   ClientHeight    =   7935
   ClientLeft      =   1965
   ClientTop       =   2250
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NOTAENTRADA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtValorDig 
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   75
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FraSeq 
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   50
      TabIndex        =   61
      Top             =   4080
      Width           =   12060
      Begin VB.TextBox txtRef 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDescontoItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   25
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   720
         MaxLength       =   8
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbCFOPAUX 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4080
         TabIndex        =   80
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCST 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "000"
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox cmbCFOP 
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
         Height          =   360
         Left            =   4080
         TabIndex        =   21
         Text            =   "-- Selecione --"
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton cmdCadProd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6840
         Picture         =   "NOTAENTRADA.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Cadastro Produto"
         Top             =   240
         Width           =   405
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   6360
         Picture         =   "NOTAENTRADA.frx":B214
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtUN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   9360
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtBarras 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         Left            =   7800
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNCM 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   720
         MaxLength       =   8
         TabIndex        =   19
         Text            =   "00"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4560
         MaxLength       =   30
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   7320
         MaxLength       =   100
         TabIndex        =   32
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   11040
         MaxLength       =   9
         TabIndex        =   23
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPrecoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   24
         Text            =   "0"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtICMS_SUBST_Item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   9360
         MaxLength       =   6
         TabIndex        =   28
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtIPIItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   7680
         MaxLength       =   6
         TabIndex        =   27
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtPercFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   11040
         MaxLength       =   6
         TabIndex        =   29
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtICMSItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6120
         MaxLength       =   6
         TabIndex        =   26
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref.:"
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
         Index           =   5
         Left            =   1500
         TabIndex        =   84
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Desconto:"
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
         Left            =   3255
         TabIndex        =   83
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label23 
         Caption         =   "Item:"
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
         Index           =   4
         Left            =   90
         TabIndex        =   81
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "U.N.:"
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
         Index           =   3
         Left            =   8880
         TabIndex        =   79
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "CST:"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   78
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NCM:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "CFOP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   76
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%Frete:"
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
         Left            =   10335
         TabIndex        =   68
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SUB%:"
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
         TabIndex        =   67
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Preço Custo:"
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
         TabIndex        =   66
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   65
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Qtde:"
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
         Left            =   10530
         TabIndex        =   64
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%IPI:"
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
         Left            =   7200
         TabIndex        =   63
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "%ICMS:"
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
         Left            =   5310
         TabIndex        =   62
         Top             =   1200
         Width           =   705
      End
   End
   Begin VB.Frame fraXML 
      Caption         =   "Importação de XML - NotaFiscal - Modelo 55"
      Height          =   735
      Left            =   4200
      TabIndex        =   56
      Top             =   6000
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtPathXML 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   59
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmd_explorer 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmd_ler_xml 
         Caption         =   "Preencher dados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   200
         Width           =   1155
      End
      Begin VB.Label lbl_caminho_xml_demo 
         Alignment       =   2  'Center
         Caption         =   "Caminho XML:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraCalculos 
      Height          =   1455
      Left            =   50
      TabIndex        =   46
      Top             =   2520
      Width           =   12060
      Begin VB.TextBox txtBaseCalculoICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPercICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtBaseICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValorICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPercICMSSubst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   12
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtValorOutras 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor  Frete:"
         Height          =   240
         Left            =   9255
         TabIndex        =   55
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Base Cálculo do ICMS:"
         Height          =   240
         Left            =   150
         TabIndex        =   54
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS:"
         Height          =   240
         Left            =   855
         TabIndex        =   53
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percentual do ICMS:"
         Height          =   240
         Left            =   375
         TabIndex        =   52
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Base Cálculo do ICMS Substituto:"
         Height          =   240
         Left            =   4245
         TabIndex        =   51
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS Substituto:"
         Height          =   240
         Left            =   4950
         TabIndex        =   50
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percentual do ICMS Substituto:"
         Height          =   240
         Left            =   4470
         TabIndex        =   49
         Top             =   960
         Width           =   2745
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Outras:"
         Height          =   240
         Left            =   9180
         TabIndex        =   48
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor do IPI:"
         Height          =   240
         Left            =   9270
         TabIndex        =   47
         Top             =   960
         Width           =   1065
      End
   End
   Begin VB.Frame fraCabeça 
      Height          =   1695
      Left            =   45
      TabIndex        =   35
      Top             =   720
      Width           =   12060
      Begin VB.TextBox txtUF 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4320
         TabIndex        =   82
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdFornec 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   7800
         Picture         =   "NOTAENTRADA.frx":BC16
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtModeloNF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   4080
         MaxLength       =   9
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7680
         MaxLength       =   12
         TabIndex        =   30
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtValorTotalNota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   10560
         MaxLength       =   12
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbTransAux 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   5880
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   360
         Left            =   5880
         TabIndex        =   6
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtNota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   8280
         TabIndex        =   36
         Top             =   240
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtEntrada 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtDtEmissao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor do Desconto:"
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
         Left            =   5865
         TabIndex        =   45
         Top             =   1245
         Width           =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Nota:"
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
         Left            =   9510
         TabIndex        =   44
         Top             =   1245
         Width           =   1035
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Transportadora:"
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
         Left            =   4260
         TabIndex        =   42
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "S. NF:"
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
         Left            =   2820
         TabIndex        =   41
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº.NF:"
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
         Left            =   750
         TabIndex        =   40
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Entrada:"
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
         Left            =   285
         TabIndex        =   39
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Emissão:"
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
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fornec:"
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
         Left            =   5055
         TabIndex        =   37
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1270
      ButtonWidth     =   3096
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Confirma Nota"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir lançamento nota"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Relatório"
            Key             =   "print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fornecedor"
            Key             =   "CadFornec"
            Object.ToolTipText     =   "Devolucao de Mercadoria"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "XML"
            Key             =   "xml"
            ImageIndex      =   10
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
         Left            =   12120
         TabIndex        =   70
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7440
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":C618
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":DA40
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":EACF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":FD37
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":11434
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":126D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":137DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":14BD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":15D72
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NOTAENTRADA.frx":16FA4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.StatusBar stBarReq 
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   7560
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   8
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2222
            MinWidth        =   2222
            Text            =   "Diponível:"
            TextSave        =   "Diponível:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "disponivel"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3152
            MinWidth        =   3152
            Text            =   "Preço Venda:"
            TextSave        =   "Preço Venda:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Key             =   "unitario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3510
            MinWidth        =   3510
            Text            =   "Preço Custo:"
            TextSave        =   "Preço Custo:"
            Key             =   "descvalr_unit"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Key             =   "desconto"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2294
            Text            =   "Valor Total:"
            TextSave        =   "Valor Total:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   3134
            MinWidth        =   3134
            Key             =   "total"
            Object.Tag             =   ""
         EndProperty
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
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12120
      DesignHeight    =   7935
   End
   Begin MSComDlg.CommonDialog buscaxml 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   45
      TabIndex        =   74
      Top             =   5880
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   2778
      _Version        =   393216
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
End
Attribute VB_Name = "frmNOTAENTRADA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
   Dim Valor_Tot_n            As Double
   Dim VLR_FRETE_N            As Currency
   Dim VLR_OUTROS_N           As Currency
   Dim VALOR_TAXA_VAREJO_N    As Currency
   Dim VALOR_TAXA_ATACADO_N   As Currency
   Dim INDR_PedidoCompra      As Boolean
   Private LastRow            As Long ' Ultima linha em que se editou
   Private LastCol            As Long ' ultima coluna em que se editou
   Private ControlVisible     As Boolean
   Private gb_Recordset       As New ADODB.Recordset ' yuri 11/05/2012
   Dim INDR_GRAVA             As Boolean

   'Importar XML
   Private cNotaEntrada As New cNotaEntrada ' yuri 11/05/2012

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INDR_GRAVA = False
   FORNEC_ID_N = 0
   PRODUTO_ID_N = 0
   TIPO_ENTRADA_N = 1
   Me.Caption = Me.Caption & Me.Name

   txtDtEmissao.PromptInclude = False
      txtDtEmissao.Text = Date
      txtDtEntrada.Text = Date
   txtDtEmissao.PromptInclude = True
   stBarReq.Refresh

   CARREGA_COMBO

   PEGA_DADOS_EMPRESA 'ficar em memoria os cfops padroes e outros dados

   If Trim(cmbTransAux.Text) = "" Then _
      cmbTransAux.Text = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Activate()

   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   SQL = "select * from TRANSPORTADORA"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabConsulta.EOF Then
      MsgBox "Realizar cadastro de transportadora primeiro."
      Unload Me
      Exit Sub
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF9
         PEDIDO_ID_N = 0
         LIMPA_TUDO
      Case vbKeyF10
         'GRAVA_TUDO_NOTA
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "xml"
         'Call cmd_explorer_Click

         If Trim(txtPathXML.Text) <> "" Then
            Msg = "Carregar XML ?"
            PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            'If RESPOSTA = vbYes Then _
               Call cmd_ler_xml_Click
         End If
      Case "consultar"
         txtNOTA.Enabled = True
         txtNOTA.SetFocus
         CRITERIO_A = ""
         frmNOTACONSULTA.Show 1
         If Trim(txtNOTA.Text) <> "" Then
            MOSTRA_NOTA_ENTRADA
            txtCNPJCPF.SetFocus
            Call txtCNPJCPF_LostFocus
         End If
         CRITERIO_A = ""
      Case "nota"
         txtNOTA.Enabled = True
         txtNOTA.SetFocus
         If Trim(txtNOTA.Text) <> "" Then
            If IsNumeric(txtNOTA.Text) Then
               FORMULA_REL = ""
               FORMULA_REL = "{vwRel_Nf_Entrada.NUMR_NOTA} = " & txtNOTA.Text

               If chkImp.Value = 1 Then _
                  ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

               Nome_Relatorio = "rel_nf_entrada.rpt"
               frmRELATORIO10.Show 1
            End If
         End If
      Case "print"
         txtNOTA.Enabled = True
         txtNOTA.SetFocus
         FORMULA_REL = ""
         If txtNOTA.Text <> "" Then _
            FORMULA_REL = "{NOTAENTRADA.numr_nota} = " & txtNOTA.Text

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Entrada.rpt"
         frmRELATORIO10.Show 1
      Case "matar"
         EXCLUIR_NOTA
      Case "voltar"
         Unload Me
      Case "gravar"
         If INDR_GRAVA = False Then _
            Exit Sub
         Msg = "Esta operação realizará o controle de estoque. Esta nota não poderá ser manipulada após atualização, confirma ?"
         PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            GRAVA_CABEÇA_NOTA_ENTRADA Trim(txtNOTA.Text), Trim(txtSerie.Text), FORNEC_ID_N, "E"

            If INDR_CONTROLA_ESTOQUE = True Then
               GRAVA_ESTOQUE
               FINANCEIRO_FORM
            End If

            LIMPA_TUDO
            txtNOTA.SetFocus
         End If
      Case "limpar"
         PEDIDO_ID_N = 0
         LIMPA_TUDO
         txtNOTA.Enabled = True
         txtNOTA.SetFocus
      Case "CadFornec"
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaCadastro.Show 1
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsProd_Click()
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
End Sub

Private Sub cmdCadProd_Click()
   If TIPO_USUARIO = 3 Or TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Then _
      frmCADASTROPRODUTO.Show 1
End Sub

Private Sub cmdFornec_Click()
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   txtCNPJCPF.PromptInclude = False
   If Trim(CNPJCPF_A) <> "" Then _
      txtCNPJCPF.Text = CNPJCPF_A
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus
End Sub

Private Sub txtBaseCalculoICMS_GotFocus()
'On Error GoTo ERRO_TRATA

    txtBaseCalculoICMS.SelStart = 0
    txtBaseCalculoICMS.SelLength = Len(txtBaseCalculoICMS.Text)
    txtBaseCalculoICMS.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseCalculoICMS_GotFocus"
End Sub

Private Sub txtBaseICMSSubst_GotFocus()
   txtBaseICMSSubst.SelStart = 0
   txtBaseICMSSubst.SelLength = Len(txtBaseICMSSubst)
   txtBaseICMSSubst.BackColor = &HC0FFFF
End Sub

'==================CNPJcpf
Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.SelLength = Len(txtCNPJCPF)

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Fornecedores", "", "", ""
   
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Mask = "##############"

   txtCNPJCPF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         txtCNPJCPF.PromptInclude = False
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   Dim strTemp As String
   Dim dblTemp As String

   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then _
      Exit Sub

   If VALIDA_CNPJCPF(Trim(txtCNPJCPF.Text)) = False Then
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   FORNEC_ID_N = 0

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select descricao,fornecedor_id,status from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabFornecedor.EOF Then
      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      Beep
      MsgBox "CNPJ/CPF não Cadastrado.", vbOKOnly, "Atenção."
      txtCNPJCPF.SetFocus
      
      Exit Sub
      Else
         txtNome.Text = Trim(TabFornecedor.Fields("descricao").Value)
         FORNEC_ID_N = Trim(TabFornecedor.Fields("fornecedor_id").Value)

         If Not IsNull(TabFornecedor!STATUS) Then
            If TabFornecedor!STATUS <> "A" Then
               If TabFornecedor.State = 1 Then _
                  TabFornecedor.Close
               MsgBox "Fornecedor Desativado, Favor Atualizar Cadastro!"
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from ENDERECO WITH (NOLOCK)"
         SQL = SQL & " Where pessoa_id = " & PESSOA_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            'Pegou o CEP do cliente
            If Not IsNull(TabTemp!CEP_ID) Then
               dblTemp = TabTemp!CEP_ID
               Else 'Não tem cadastrado CEP_id, impossivel fazer tributacao sem a uf
                  If TabFornecedor.State = 1 Then _
                     TabFornecedor.Close
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
                  MsgBox "O Cadastro do Fornecedor não está completo. Verique os dados (CEP_id, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close

            'Pegar a uf do cliente
            SQL = "select * from CEP WITH (NOLOCK) "
            SQL = SQL & " where cep_ID = " & dblTemp
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               If Not IsNull(TabTemp!UF) Then
                  txtUF.Text = TabTemp!UF
                  Else 'UF nao localizada
                     If TabTemp.State = 1 Then _
                        TabTemp.Close
                     If TabFornecedor.State = 1 Then _
                        TabFornecedor.Close

                     MsgBox "O Cadastro do fornecedor não está completo. Verique os dados (CEP_id, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                     txtCNPJCPF.SetFocus
                     Exit Sub
               End If
               Else
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
                  If TabFornecedor.State = 1 Then _
                     TabFornecedor.Close

                  MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais incompletos. Verique-os, principalmente o Estado(UF) da empresa"
                  txtCNPJCPF.SetFocus
                  Exit Sub
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If

         If txtUF.Text = "GO" Then
            cmbCFOPAux.Text = CFOP_ENTRADA_DE
            Else: cmbCFOPAux.Text = CFOP_ENTRADA_FE
         End If
         cmbCFOP.Text = cmbCFOPAux.Text & "-" & TRAZ_CFOP(cmbCFOPAux.Text)
   End If   'If TabFornecedor.EOF Then
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   txtCNPJCPF.PromptInclude = False
   CRITERIO_A = txtCNPJCPF.Text

   If Trim(txtNOTA.Text) <> "" And Trim(txtSerie.Text) <> "" And Trim(CRITERIO_A) <> "" Then _
      MOSTRA_NOTA_ENTRADA

   If Trim(CRITERIO_A) <> "" Then
      If Len(CRITERIO_A) <= 11 Then
         txtCNPJCPF.Mask = "###.###.###-##"
         Else: txtCNPJCPF.Mask = "##.###.###/####-##"
      End If
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CRITERIO_A
   End If
   txtCNPJCPF.PromptInclude = True
   txtCNPJCPF.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_LostFocus"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtEmissao.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub txtCST_GotFocus()
   txtCST.SelStart = 0
   txtCST.SelLength = Len(txtCST)
   txtCST.BackColor = &HC0FFFF
End Sub

Private Sub TXTCST_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCFOP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCST_KeyPress"
End Sub

Private Sub txtCST_LostFocus()
   txtCST.BackColor = &HFFFFFF
End Sub

Private Sub txtDesconto_GotFocus()
'On Error GoTo ERRO_TRATA

    txtDesconto.SelStart = 0
    txtDesconto.SelLength = Len(txtDesconto.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesconto_GotFocus"
End Sub

Private Sub txtDtEmissao_GotFocus()
   txtDtEmissao.SelStart = 0
   txtDtEmissao.SelLength = Len(txtDtEmissao)
   txtDtEmissao.BackColor = &HC0FFFF
End Sub

Private Sub txtDTEMISSAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtEntrada.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTEMISSAO_KeyPress"
End Sub

Private Sub txtdtemissao_LostFocus()
'On Error GoTo ERRO_TRATA

   If Not IsDate(txtDtEmissao.Text) Then
      txtDtEmissao.PromptInclude = False
         txtDtEmissao.Text = Date
      txtDtEmissao.PromptInclude = True
   End If
   txtDtEmissao.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtemissao_LostFocus"
End Sub

Private Sub txtDtEntrada_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtEntrada.SelStart = 0
   txtDtEntrada.SelLength = Len(txtDtEntrada)
   txtDtEntrada.PromptInclude = True
   txtDtEntrada.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEntrada_GotFocus"
End Sub

Private Sub txtdtentrada_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBaseCalculoICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdtentrada_KeyPress"
End Sub

Private Sub cmbTrans_GotFocus()
   cmbTrans.BackColor = &HC0FFFF
End Sub

Private Sub cmbTrans_Click()
'On Error GoTo ERRO_TRATA

   cmbTransAux.ListIndex = cmbTrans.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTrans_Click"
End Sub

Private Sub cmbtrans_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtrans_KeyPress"
End Sub

Private Sub txtBaseCalculoICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseCalculoICMS_KeyPress"
End Sub

Private Sub txtBaseCalculoICMS_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtBaseCalculoICMS.Text = "" Then _
      txtBaseCalculoICMS.Text = 0
   If IsNumeric(txtBaseCalculoICMS.Text) Then
      VALOR_ITEM_N = txtBaseCalculoICMS.Text
      txtBaseCalculoICMS.Text = VALOR_ITEM_N
   End If
   txtBaseCalculoICMS.BackColor = &HFFFFFF
   txtBaseCalculoICMS.Text = Format(txtBaseCalculoICMS.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseCalculoICMS_LostFocus"
End Sub

Private Sub txtFrete_GotFocus()
   txtFrete.SelStart = 0
   txtFrete.SelLength = Len(txtFrete)
   txtFrete.BackColor = &HC0FFFF
End Sub

Private Sub txtICMSItem_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtICMSItem.Text = "" Then _
      txtICMSItem.Text = 0

   txtICMSItem.SelStart = 0
   txtICMSItem.SelLength = Len(txtICMSItem.Text)
   txtICMSItem.BackColor = &HC0FFFF
            
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMSItem_GotFocus"
End Sub

Private Sub txtICMSItem_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtICMSItem.Text = "" Then _
      txtICMSItem.Text = 0

   VALOR_ITEM_N = txtICMSItem.Text
   txtICMSItem.BackColor = &HFFFFFF
   txtICMSItem.Text = Format(txtICMSItem.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMSItem_LostFocus"
End Sub

Private Sub txtIPIItem_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtIPIItem.Text = "" Then _
      txtIPIItem.Text = 0

   txtIPIItem.SelStart = 0
   txtIPIItem.SelLength = Len(txtIPIItem.Text)
   txtIPIItem.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIPIItem_GotFocus"
End Sub

Private Sub txtIPIItem_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtIPIItem.Text = "" Then _
      txtIPIItem.Text = 0

   VALOR_ITEM_N = txtIPIItem.Text
   txtIPIItem.BackColor = &HFFFFFF
   txtIPIItem.Text = Format(txtIPIItem.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIPIItem_LostFocus"
End Sub

Private Sub txtICMS_SUBST_Item_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtICMS_SUBST_Item.Text = "" Then _
      txtICMS_SUBST_Item.Text = 0

   txtICMS_SUBST_Item.SelStart = 0
   txtICMS_SUBST_Item.SelLength = Len(txtICMS_SUBST_Item.Text)
   txtICMS_SUBST_Item.BackColor = &HC0FFFF
            
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_SUBST_Item_GotFocus"
End Sub

Private Sub txtICMS_SUBST_Item_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtICMS_SUBST_Item.Text = "" Then _
      txtICMS_SUBST_Item.Text = 0

   VALOR_ITEM_N = txtICMS_SUBST_Item.Text
   txtICMS_SUBST_Item.BackColor = &HFFFFFF
   txtICMS_SUBST_Item.Text = Format(txtICMS_SUBST_Item.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_SUBST_Item_LostFocus"
End Sub

Private Sub txtModeloNF_GotFocus()
   txtModeloNF.SelStart = 0
   txtModeloNF.SelLength = Len(txtModeloNF)
   txtModeloNF.BackColor = &HC0FFFF
End Sub

Private Sub txtmodelonf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtmodelonf_KeyPress"
   'Resume
End Sub

Private Sub txtModeloNF_LostFocus()
   txtModeloNF.BackColor = &HFFFFFF
End Sub

Private Sub txtNCM_GotFocus()
   txtNCM.SelStart = 0
   txtNCM.SelLength = Len(txtNCM)
   txtNCM.BackColor = &HC0FFFF
End Sub

Private Sub TXTNCM_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCST.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTNCM_KeyPress"
End Sub

Private Sub txtNCM_LostFocus()
   txtNCM.BackColor = &HFFFFFF
   If Trim(txtNCM.Text) = "" Then _
      txtNCM.Text = "00"
End Sub

Private Sub txtNota_GotFocus()
   txtNOTA.SelStart = 0
   txtNOTA.SelLength = Len(txtNOTA)
   txtNOTA.BackColor = &HC0FFFF
End Sub

Private Sub txtNOTA_LostFocus()
   txtNOTA.BackColor = &HFFFFFF
End Sub

Private Sub TXTPERCFRETE_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtPercFrete.Text = "" Then _
      txtPercFrete.Text = 0

   txtPercFrete.SelStart = 0
   txtPercFrete.SelLength = Len(txtPercFrete.Text)
   txtPercFrete.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPERCFRETE_GotFocus"
End Sub

Private Sub TXTPERCFRETE_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtPercFrete.Text = "" Then _
      txtPercFrete.Text = 0

   VALOR_ITEM_N = txtPercFrete.Text
   txtPercFrete.BackColor = &HFFFFFF
   txtPercFrete.Text = Format(txtPercFrete.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPERCFRETE_LostFocus"
End Sub

Private Sub txtPercICMS_GotFocus()
'On Error GoTo ERRO_TRATA

   txtPercICMS.SelStart = 0
   txtPercICMS.SelLength = Len(txtPercICMS.Text)
   txtPercICMS.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercICMS_GotFocus"
End Sub

Private Sub txtPercICMSSubst_GotFocus()
   txtPercICMSSubst.SelStart = 0
   txtPercICMSSubst.SelLength = Len(txtPercICMSSubst)
   txtPercICMSSubst.BackColor = &HC0FFFF
End Sub

Private Sub txtPrecoCusto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtPrecoCusto.SelStart = 0
   txtPrecoCusto.SelLength = Len(txtPrecoCusto.Text)
   txtPrecoCusto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPrecoCusto_GotFocus"
End Sub

Private Sub txtdescontoitem_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtDescontoItem.Text = "" Then _
      txtDescontoItem.Text = 0

   txtDescontoItem.SelStart = 0
   txtDescontoItem.SelLength = Len(txtDescontoItem.Text)
   txtDescontoItem.BackColor = &HC0FFFF
            
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdescontoitem_GotFocus"
End Sub

Private Sub txtdescontoitem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtICMSItem.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdescontoitem_KeyPress"
End Sub

Private Sub txtdescontoitem_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtDescontoItem.Text = "" Then _
      txtDescontoItem.Text = 0
   txtDescontoItem.BackColor = &HFFFFFF
   txtDescontoItem.Text = Format(txtDescontoItem.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdescontoitem_LostFocus"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   If txtQTDE.Text = "" Then _
      txtQTDE.Text = 0

   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE.Text)
   txtQTDE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_GotFocus"
End Sub

Private Sub txtseq_LostFocus()
'On Error GoTo ERRO_TRATA

   txtSeq.BackColor = &HFFFFFF

   If Trim(txtSeq.Text) <> "" Then
      MOSTRA_ITEM_NOTA_ENTRADA NOTAENTRADA_ID_N, txtSeq.Text
      Else
         SEQ_ID_N = MAX_ID("seq_id", "NOTAENTRADAITEM", "entrada_id", Trim(NOTAENTRADA_ID_N), "", "")
         txtSeq.Text = SEQ_ID_N
   End If
   If Trim(txtSeq.Text) = "" Then _
      txtSeq.Text = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtSeq_LostFocus"
End Sub

Private Sub txtSerie_GotFocus()
   txtSerie.SelStart = 0
   txtSerie.SelLength = Len(txtSerie)
   txtSerie.BackColor = &HC0FFFF
End Sub

Private Sub txtSERIE_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtSerie.Text) = "" Then _
      txtSerie.Text = 1

   txtSerie.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtSERIE_LostFocus"
End Sub

Private Sub txtUN_GotFocus()
   txtUN.SelStart = 0
   txtUN.SelLength = Len(txtUN)
   txtUN.BackColor = &HC0FFFF
End Sub

Private Sub txtun_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtQTDE.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUN_KeyPress"
End Sub

Private Sub txtUN_LostFocus()
   txtUN.BackColor = &HFFFFFF
End Sub

Private Sub txtValorDig_LostFocus()
   txtValorDig.BackColor = &HFFFFFF
End Sub

Private Sub txtValorICMS_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorICMS.SelStart = 0
   txtValorICMS.SelLength = Len(txtValorICMS.Text)
   txtValorICMS.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMS_GotFocus"
End Sub

Private Sub txtValorICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMS_KeyPress"
End Sub

Private Sub txtValorICMS_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorICMS.Text = "" Then _
      txtValorICMS.Text = 0
   If IsNumeric(txtValorICMS.Text) Then
      VALOR_ITEM_N = txtValorICMS.Text
      txtValorICMS.Text = VALOR_ITEM_N
   End If
   txtValorICMS.BackColor = &HFFFFFF
   txtValorICMS.Text = Format(txtValorICMS.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMS_LostFocus"
End Sub

Private Sub txtpercICMS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBaseICMSSubst.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercICMS_KeyPress"
End Sub

Private Sub txtpercICMS_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtPercICMS.Text = "" Then _
      txtPercICMS.Text = 0
   If IsNumeric(txtPercICMS.Text) Then _
      VALOR_ITEM_N = txtPercICMS.Text
   txtPercICMS.BackColor = &HFFFFFF
   txtPercICMS.Text = Format(txtPercICMS.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercICMS_LostFocus"
End Sub

Private Sub txtBaseICMSSubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorICMSSubst.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseICMSSubst_KeyPress"
End Sub

Private Sub txtBaseICMSSubst_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtBaseICMSSubst.Text = "" Then _
      txtBaseICMSSubst.Text = 0
   If IsNumeric(txtBaseICMSSubst.Text) Then _
      VALOR_ITEM_N = txtBaseICMSSubst.Text
   txtBaseICMSSubst.BackColor = &HFFFFFF
   txtBaseICMSSubst.Text = Format(txtBaseICMSSubst.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBaseICMSSubst_LostFocus"
End Sub

Private Sub txtValorICMSSubst_GotFocus()
   txtValorICMSSubst.SelStart = 0
   txtValorICMSSubst.SelLength = Len(txtValorICMSSubst)
   txtValorICMSSubst.BackColor = &HC0FFFF
End Sub

Private Sub txtValorICMSSubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercICMSSubst.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMSSubst_KeyPress"
End Sub

Private Sub txtValorICMSSubst_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorICMSSubst.Text = "" Then _
      txtValorICMSSubst.Text = 0
   If IsNumeric(txtValorICMSSubst.Text) Then _
      VALOR_ITEM_N = txtValorICMSSubst.Text
   txtValorICMSSubst.BackColor = &HFFFFFF
   txtValorICMSSubst.Text = Format(txtValorICMSSubst.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorICMSSubst_LostFocus"
End Sub

Private Sub txtPercICMSSubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorOutras.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercICMSSubst_KeyPress"
End Sub

Private Sub txtPercICMSSubst_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtPercICMSSubst.Text = "" Then _
      txtPercICMSSubst.Text = 0
   If IsNumeric(txtPercICMSSubst.Text) Then _
      VALOR_ITEM_N = txtPercICMSSubst.Text
   txtPercICMSSubst.BackColor = &HFFFFFF
   txtPercICMSSubst.Text = Format(txtPercICMSSubst.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPercICMSSubst_LostFocus"
End Sub

Private Sub txtValorIPI_GotFocus()
   txtValorIPI.SelStart = 0
   txtValorIPI.SelLength = Len(txtValorIPI)
   txtValorIPI.BackColor = &HC0FFFF
End Sub

Private Sub txtValorOutras_GotFocus()
   txtValorOutras.SelStart = 0
   txtValorOutras.SelLength = Len(txtValorOutras)
   txtValorOutras.BackColor = &HC0FFFF
End Sub

Private Sub txtValorOutras_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtFrete.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorOutras_KeyPress"
End Sub

Private Sub txtValorOutras_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorOutras.Text = "" Then _
      txtValorOutras.Text = 0
   If IsNumeric(txtValorOutras.Text) Then _
      VALOR_ITEM_N = txtValorOutras.Text
   txtValorOutras.BackColor = &HFFFFFF
   txtValorOutras.Text = Format(txtValorOutras.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorOutras_LostFocus"
End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorIPI.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFrete_KeyPress"
End Sub

Private Sub txtFrete_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtFrete.Text = "" Then _
      txtFrete.Text = 0
   If IsNumeric(txtFrete.Text) Then _
      VALOR_ITEM_N = txtFrete.Text
   txtFrete.BackColor = &HFFFFFF
   txtFrete.Text = Format(txtFrete.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFrete_LostFocus"
End Sub

Private Sub txtvaloripi_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorIPI_KeyPress"
End Sub

Private Sub txtValorIPI_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorIPI.Text = "" Then _
      txtValorIPI.Text = 0
   If IsNumeric(txtValorIPI.Text) Then _
      VALOR_ITEM_N = txtValorIPI.Text
   txtValorIPI.BackColor = &HFFFFFF
   txtValorIPI.Text = Format(txtValorIPI.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorIPI_LostFocus"
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCFOP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_KeyPress"
End Sub

Private Sub txtDesconto_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtDesconto.Text = "" Then _
      txtDesconto.Text = 0
   If IsNumeric(txtDesconto.Text) Then _
      VALOR_ITEM_N = txtDesconto.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesconto_LostFocus"
End Sub

Private Sub txtDtEntrada_LostFocus()
'On Error GoTo ERRO_TRATA

   If Not IsDate(txtDtEntrada.Text) Then
      txtDtEntrada.PromptInclude = False
         txtDtEntrada.Text = Date
      txtDtEntrada.PromptInclude = True
   End If
   txtDtEntrada.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtEntrada_LostFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtNOTA.Text) = "" Then _
         Exit Sub
      txtSerie.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnota_KeyPress"
   'Resume
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtModeloNF.Enabled = True
      txtModeloNF.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtserie_KeyPress"
   'Resume
End Sub


Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   txtSeq.SelStart = 0
   txtSeq.SelLength = Len(txtSeq)
   txtSeq.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_GotFocus"
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         EXCLUIR_ITEM_NOTA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyDown"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtRef.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtseq_KeyPress"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   MOSTRA_RODAPE "F3 - Consulta Produtos", "F6 - Excluir Item", "", "", ""
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         EXCLUIR_ITEM_NOTA
      Case vbKeyF7
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtProduto.Text) <> "" Then
         If FORNEC_ID_N <= 0 Then _
            txtCNPJCPF_LostFocus

         PROCESSA_DADOS_PRODUTOS
         Else
            txtProduto.SetFocus
            Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtRef_GotFocus()
'On Error GoTo ERRO_TRATA

   txtRef.SelStart = 0
   txtRef.SelLength = Len(txtRef)
   txtRef.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtref_GotFocus"
End Sub

Private Sub txtref_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         'EXCLUIR_ITEM_NOTA
      Case vbKeyF7
         'frmrefConsulta.Show 1
         'If SQL3 <> "" Then
         '   txtRef.Text = SQL3
         '   txtRef.SetFocus
         'End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtref_KeyDown"
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtRef.Text) <> "" Then
         If FORNEC_ID_N <= 0 Then _
            txtCNPJCPF_LostFocus

         LE_REFERENCIA
         txtProduto.SetFocus
         Else
            txtProduto.SetFocus
            Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtref_KeyPress"
End Sub

Private Sub txtref_LostFocus()
   txtRef.BackColor = &HFFFFFF
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtQTDE.Text = "" Then
         MsgBox "Quantidade não pode ser 0."
         txtQTDE.SetFocus
         Exit Sub
      End If
      QTDE_PEDIDO = txtQTDE.Text
      txtPrecoCusto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtQTDE.Text) = "" Then _
      txtQTDE.Text = 0

   QTDE_PEDIDO = 0 & txtQTDE.Text

   If QTDE_PEDIDO <= 0 Then
      MsgBox "Qtde Não Informada !!!"
      txtQTDE.SetFocus
      Exit Sub
   End If
   txtQTDE.BackColor = &HFFFFFF
   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_LostFocus"
End Sub

Private Sub txtIPIItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtIPIItem.Text = "" Then _
         txtIPIItem.Text = 0

      If txtValorIPI.Text = "" Then _
         txtValorIPI.Text = 0

      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      txtICMS_SUBST_Item.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIPIItem_KeyPress"
End Sub

Private Sub txtICMSItem_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtICMSItem.Text = "" Then _
         txtICMSItem.Text = 0

      If txtValorICMS.Text = "" Then _
         txtValorICMS.Text = 0

      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      txtIPIItem.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMSItem_KeyPress"
End Sub

Private Sub txtICMS_SUBST_item_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtICMS_SUBST_Item.Text = "" Then _
         txtICMS_SUBST_Item.Text = 0

      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      txtPercFrete.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_SUBST_item_KeyPress"
End Sub

Private Sub TXTPERCFRETE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtPercFrete.Text = "" Then _
         txtPercFrete.Text = 0

      VALOR_ITEM_N = 0
      Valor_Tot_n = 0
      VALOR_DIFERENCA_N = 0

      VALIDA_DADOS_NOTA_ENTRADA
      LIMPA_BODY

      txtSeq.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTPERCFRETE_KeyPress"
End Sub

Private Sub txtprecocusto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescontoItem.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   
   TRATA_ERROS Err.Description, Me.Name, "txtPrecoCusto_KeyPress"
End Sub

Private Sub txtPrecoCusto_LostFocus()
'On Error GoTo ERRO_TRATA

   VALOR_ITEM_N = 0 & txtPrecoCusto.Text
   QTDE_PEDIDO = 0 & txtQTDE.Text

   If VALOR_ITEM_N <= 0 Then
      MsgBox "Valor informado incorreto !!!"
      txtPrecoCusto.SetFocus
      Exit Sub
   End If
   txtPrecoCusto.BackColor = &HFFFFFF
   txtPrecoCusto.Text = Format(txtPrecoCusto.Text, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPrecoCusto_LostFocus"
End Sub

Private Sub cmbCFOP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(cmbCFOP.Text) = "" Then
         MsgBox "Selecione CFOP"
         cmbCFOP.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      txtUN.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfop_KeyPress"
End Sub

Private Sub cmbcfop_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbCFOPAux.Text) <> "" Then _
      cmbCFOP.Text = cmbCFOPAux.Text & "-" & TRAZ_CFOP(cmbCFOPAux.Text)

   cmbCFOP.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCFOP_LostFocus"
End Sub

Private Sub cmbCFOP_Click()
'On Error GoTo ERRO_TRATA

   cmbCFOPAux.ListIndex = cmbCFOP.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfop_Click"
End Sub

Private Sub cmbCFOP_GotFocus()
'On Error GoTo ERRO_TRATA

   If cmbCFOP.Text = "" Then 'nao escolheu cfop, coloca o default
      If txtUF.Text = "GO" Then
         cmbCFOPAux.Text = CFOP_ENTRADA_DE
         Else: cmbCFOPAux.Text = CFOP_ENTRADA_FE
      End If
      cmbCFOP.Text = cmbCFOPAux.Text & "-" & TRAZ_CFOP(cmbCFOPAux.Text)
   End If
   cmbCFOP.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCFOP_GotFocus"
End Sub

Private Sub cmbTrans_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbTrans.Text) = "" Then
      cmbTransAux.Text = 1

      If TabTemp.State = 1 Then _
         TabTemp.Close
      SQL = "select descricao,cnpjcpf from vwTRANSPORTADORA WITH (NOLOCK)"
      SQL = SQL & " where transp_id = " & cmbTransAux.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         cmbTrans.Text = Trim(TabTemp!DESCRICAO) & "-" & Trim(TabTemp!CNPJCPF)
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   cmbTrans.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTrans_LostFocus"
End Sub

' FIM DAS FUNÇÕES PARA A IMPORTAÇÃO DO XML, CONFORME O CAMINHO DADO PARA O USUARIO
Private Sub cmd_explorer_Click()
   buscaxml.InitDir = "C:\"
   buscaxml.ShowOpen
   txtPathXML.Text = buscaxml.FileName
   buscaxml.FileName = ""
End Sub

Private Sub cmd_ler_xml_Click()
   ' yuri 11/05/2012
   Call PREENCHE_CABEÇA_XML

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from NOTAENTRADA WITH (NOLOCK)"
   SQL = SQL & " where numr_nota = " & txtNOTA.Text
   SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      PEDIDO_ID_N = 0
      LIMPA_TUDO
      MOSTRA_NOTA_ENTRADA

      If TabNOTA!STATUS = "C" Then
          If TabNOTA.State = 1 Then _
             TabNOTA.Close

         MsgBox "Nota fiscal cancelada, impossível alterar."
         LIMPA_TUDO
         txtNOTA.SetFocus
         Exit Sub
      End If

      If TabNOTA!STATUS = "E" Then
          If TabNOTA.State = 1 Then _
             TabNOTA.Close

         MsgBox "Nota fiscal já atualizada, impossível alterar."
         LIMPA_TUDO
         Exit Sub
      End If
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   Call PREENCHE_TRANSPORTE_XML
   Call PREENCHE_PRODUTO_XML(txtPathXML)
End Sub
'================================
Sub MOSTRA_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   NOTAENTRADA_ID_N = 0

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      MsgBox "Informe fornecedor."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If

   If Trim(txtNOTA.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtNOTA.Text) Then _
      Exit Sub
   If Trim(txtSerie.Text) = "" Then _
      txtSerie.Text = 1

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select * from vwNotaEntrada WITH (NOLOCK)"

   SQL = SQL & " where numr_nota = " & txtNOTA.Text
   SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      LIMPA_TUDO

      NOTAENTRADA_ID_N = 0 & TabNOTA.Fields("entrada_id").Value
      txtNOTA.Text = "" & TabNOTA!NUMR_NOTA
      txtSerie.Text = "" & TabNOTA!SERIE_NOTA
      FORNEC_ID_N = TabNOTA!FORNECEDOR_ID
      txtNome.Text = Trim(TabNOTA.Fields("nomefornecedor").Value)
      txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = Trim(TabNOTA.Fields("cnpjcpf").Value)
      txtDtEntrada.PromptInclude = False
         txtDtEntrada.Text = TabNOTA!DT_ENTRADA
      txtDtEntrada.PromptInclude = True
      txtDtEmissao.PromptInclude = False
         txtDtEmissao.Text = TabNOTA!DT_EMISSAO
      txtDtEmissao.PromptInclude = True
      txtBaseCalculoICMS.Text = Format(TabNOTA!BASE_CALC_ICMS, strFormatacao2Digitos)
      txtValorICMS.Text = Format(TabNOTA!VALOR_ICMS, strFormatacao2Digitos)
      txtPercICMS.Text = Format(TabNOTA!PERC_ICMS, strFormatacao2Digitos)
      txtBaseICMSSubst.Text = TabNOTA!BASE_CALC_ICMS_SUBST
      txtValorICMSSubst.Text = TabNOTA!VALOR_ICMS_SUBST
      txtPercICMSSubst.Text = TabNOTA!PERC_ICMS_SUBST
      txtValorOutras.Text = TabNOTA!VALOR_OUTRAS
      txtFrete.Text = TabNOTA!VALOR_FRETE
      txtValorIPI.Text = TabNOTA!VALOR_IPI
      txtDesconto.Text = TabNOTA!Valor_Desconto

      If Not IsNull(TabNOTA!TRANSP_ID) Then
         cmbTransAux.Text = TabNOTA!TRANSP_ID

         If TabTemp.State = 1 Then _
            TabTemp.Close
         SQL = "select * from vwTRANSPORTADORA WITH (NOLOCK)"
         SQL = SQL & " where transp_id = " & cmbTransAux.Text
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            cmbTrans.Text = Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("cnpjcpf").Value)
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      'cmbCFOPAux.Text = Trim(TabNOTA!CFOP)
      'cmbCFOP.Text = cmbCFOPAux.Text & "-" & TRAZ_CFOP(cmbCFOPAux.Text)

      SETA_GRID

      stBarReq.Refresh

      If TabNOTA!STATUS = "C" Then
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         MsgBox "Nota fiscal cancelada, impossível alterar."

         LIMPA_TUDO

         txtNOTA.SetFocus
         Exit Sub
      End If
      If TabNOTA!STATUS = "E" Then
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         MsgBox "Nota fiscal já atualizada, impossível alterar."

         LIMPA_TUDO
         txtNOTA.SetFocus
         Exit Sub
      End If
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_NOTA_ENTRADA"
End Sub

Sub EXCLUIR_NOTA()
'On Error GoTo ERRO_TRATA

   txtNOTA.Enabled = True
   txtNOTA.SetFocus
   If Trim(txtNOTA.Text) <> "" And Trim(txtSerie.Text) <> "" And FORNEC_ID_N > 0 Then
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "select entrada_id,status from NOTAENTRADA WITH (NOLOCK)"
      SQL = SQL & " where numr_nota = " & txtNOTA.Text
      SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
         If Trim(TabNOTA.Fields("status")) = "A" Then
            Msg = "Confirma exclusão de nota ?"
            PERGUNTA Msg & txtNOTA.Text, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
   
               SQL = "delete from NotaEntradaitem "
               SQL = SQL & " where entrada_id = " & TabNOTA.Fields(0).Value
               CONECTA_RETAGUARDA.Execute SQL
   
               SQL = "delete from NotaEntrada "
               SQL = SQL & " where entrada_id = " & TabNOTA.Fields(0).Value
               CONECTA_RETAGUARDA.Execute SQL
   
               PEDIDO_ID_N = 0
               LIMPA_TUDO
            End If
            MsgBox "Nota fiscal atualizada."
         End If
         Else: MsgBox "Nota Fiscal inexistente."
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close
      Else: MsgBox "Informe Dados para Pesquisa."
   End If
   txtNOTA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_NOTA"
End Sub

Sub EXCLUIR_ITEM_NOTA()
'On Error GoTo ERRO_TRATA

   If Trim(txtSeq.Text) <> "" And Trim(txtNOTA.Text) <> "" And Trim(txtSerie.Text) <> "" And FORNEC_ID_N > 0 Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select NOTAENTRADAITEM.ENTRADA_ID from NOTAENTRADA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN NOTAENTRADAITEM WITH (NOLOCK)"
      SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
      SQL = SQL & " AND NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"

      SQL = SQL & " where seq_id = '" & Trim(txtSeq.Text) & "'"
      SQL = SQL & " and numr_nota = " & txtNOTA.Text
      SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Confirma exclusão desse seq ?"
         PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            SQL = "delete from NOTAENTRADAITEM "
            SQL = SQL & " where entrada_id = " & TabTemp.Fields(0).Value
            SQL = SQL & " and seq_id = " & Trim(txtSeq.Text)
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_BODY
            MOSTRA_TOTAL_NOTA
         End If
      End If
      txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM_NOTA"
End Sub

Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   FORNEC_ID_N = 0
   INDR_GRAVA = False
   txtNOTA.Text = ""
   txtSerie.Text = ""
   txtModeloNF.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   txtDtEmissao.PromptInclude = False
   txtDtEmissao.Text = ""
   txtDtEntrada.PromptInclude = False
   txtDtEntrada.Text = ""
   txtUF.Text = ""
   cmbTrans.Text = ""
   cmbTransAux.Text = ""
   txtDesconto.Text = ""
   txtValorTotalNota.Text = ""
   txtBaseCalculoICMS.Text = 0
   txtValorICMS.Text = 0
   txtPercICMS.Text = 0
   txtBaseICMSSubst.Text = 0
   txtValorICMSSubst.Text = 0
   txtPercICMSSubst.Text = 0
   txtValorOutras.Text = 0
   txtFrete.Text = 0
   txtValorIPI.Text = 0

   NOTAENTRADA_ID_N = 0
   FORNEC_ID_N = 0
   INDR_RECEITA = 0
   CRITERIO_A = ""
   VALOR_TOTAL_N = 0
   VALOR_DESCONTO_N = 0
   Valor_Tot_n = 0

   stBarReq.Panels(8).Text = ""
   stBarReq.Refresh

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtRef.Text = ""
   txtBarras.Text = ""
   txtSeq.Text = ""
   txtProduto.Text = ""
   txtDesc.Text = ""
   txtNCM.Text = ""
   txtCST.Text = ""
   cmbCFOP.Text = ""
   cmbCFOPAux.Text = ""
   txtUN.Text = ""
   txtQTDE.Text = ""
   txtPrecoCusto.Text = 0
   txtDescontoItem.Text = 0
   txtICMSItem.Text = 0
   txtIPIItem.Text = 0
   txtICMS_SUBST_Item.Text = 0
   txtPercFrete.Text = 0

   stBarReq.Panels(2).Text = ""
   stBarReq.Panels(4).Text = ""
   stBarReq.Panels(6).Text = ""

   VALOR_ITEM_N = 0
   VALOR_DIFERENCA_N = 0
   QTDE_PEDIDO = 0
   PRODUTO_ID_N = 0

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub MOSTRA_ITEM_NOTA_ENTRADA(NOTA_ENTRADA_ID_N As Long, SEQ_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If NOTA_ENTRADA_ID_N > 0 Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select NOTAENTRADAITEM.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
      SQL = SQL & " from NOTAENTRADAITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON  NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where entrada_id = " & NOTA_ENTRADA_ID_N
      SQL = SQL & " and seq_id = " & SEQ_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtProduto.Text = "" & Trim(TabConsulta.Fields("codg_produto").Value)
         txtDesc.Text = "" & Trim(TabConsulta.Fields("descricao").Value)
         txtNCM.Text = "" & Trim(TabConsulta.Fields("ncm").Value)
         txtCST.Text = "" & Trim(TabConsulta.Fields("cst").Value)
         cmbCFOPAux.Text = "" & Trim(TabConsulta.Fields("cfop_id").Value)
         cmbCFOP.Text = cmbCFOPAux.Text & "-" & TRAZ_CFOP(cmbCFOPAux.Text)
         txtUN.Text = "" & Trim(TabConsulta.Fields("un").Value)
         txtQTDE.Text = "" & Format(TabConsulta.Fields("qtde_entrada").Value, strFormatacao3Digitos)
         txtICMSItem.Text = "" & Format(TabConsulta.Fields("PERC_ICMS").Value, strFormatacao2Digitos)
         txtIPIItem.Text = "" & Format(TabConsulta.Fields("PERC_IPI").Value, strFormatacao2Digitos)
         txtICMS_SUBST_Item.Text = "" & Format(TabConsulta.Fields("PERC_ICMS_SUB").Value, strFormatacao2Digitos)
         txtPercFrete.Text = "" & Format(TabConsulta.Fields("perc_frete").Value, strFormatacao2Digitos)

         QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, TabConsulta.Fields("produto_id").Value)
         stBarReq.Panels(2).Text = Format(QTDE_ESTOQUE_N, strFormatacao3Digitos)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ITEM_NOTA_ENTRADA"
End Sub

Private Sub VALIDA_DADOS_NOTA_ENTRADA()
'On Error GoTo ERRO_TRATA

   If Trim(txtNOTA.Text) = "" Then
      txtNOTA.SetFocus
      Exit Sub
   End If
   If Trim(txtSerie.Text) = "" Then
      txtSerie.SetFocus
      Exit Sub
   End If
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Or FORNEC_ID_N <= 0 Then
      MsgBox "Informe Fornecedor."
      txtCNPJCPF.SetFocus
      Exit Sub
   End If
   txtDtEntrada.PromptInclude = True
   If Not IsDate(txtDtEntrada.Text) Then
      MsgBox "Informe data de entrada."
      txtDtEntrada.SetFocus
      Exit Sub
   End If
   txtDtEmissao.PromptInclude = True
   If Not IsDate(txtDtEmissao.Text) Then
      MsgBox "Informe data de emissão."
      txtDtEmissao.SetFocus
      Exit Sub
   End If
   If Trim(cmbCFOPAux.Text) = "" Then
      MsgBox "Selecione CFOP"
      cmbCFOP.SetFocus
      Exit Sub
   End If

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Seqüência sem codigo de Produto.", vbOKOnly, "Atenção !!!"
      txtProduto.SetFocus
      Exit Sub
   End If

   If Not IsNull(txtPrecoCusto.Text) Then
      If txtPrecoCusto.Text <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção !!!"
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtQTDE.Text) = "" Then
      MsgBox "Quantidade inválida.", vbOKOnly, "Atenção !!!"
      txtQTDE.SetFocus
      Exit Sub
   End If

   If IsNumeric(txtQTDE.Text) Then _
      QTDE_PEDIDO = txtQTDE.Text

   If Trim(txtPrecoCusto.Text) = "" Then
      MsgBox "Valor total inválido."
      txtPrecoCusto.SetFocus
      Exit Sub
      Else
         VALOR_ITEM_N = 0 & txtPrecoCusto.Text
         VALOR_TOTAL_N = (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N + VALOR_TOTAL_N
   End If

   GRAVA_CABEÇA_NOTA_ENTRADA Trim(txtNOTA.Text), Trim(txtSerie.Text), FORNEC_ID_N, "A"
   GRAVA_ITEM_NOTA_ENTRADA NOTAENTRADA_ID_N, _
                           txtSeq.Text, _
                           PRODUTO_ID_N, _
                           txtPrecoCusto.Text, _
                           txtQTDE.Text, _
                           "A", _
                           cmbCFOPAux.Text, _
                           txtIPIItem.Text, _
                           txtICMSItem.Text, _
                           txtDescontoItem.Text, _
                           txtICMS_SUBST_Item.Text, _
                           txtPercFrete.Text, _
                           txtNCM.Text, _
                           txtCST.Text, _
                           txtUN.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VALIDA_DADOS_NOTA_ENTRADA"
End Sub

Private Sub MOSTRA_TOTAL_NOTA()
'On Error GoTo ERRO_TRATA
'SET ARRUMAR SUBROTINAS
   VALOR_TOTAL_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum((PRECO_CUSTO * qtde_entrada)-isnull(NOTAENTRADAITEM.VALOR_DESCONTO,0)) as ValorTotal "
   SQL = SQL & " from NOTAENTRADA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM WITH (NOLOCK)"
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID "
   SQL = SQL & " AND NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"

   SQL = SQL & " where NOTAENTRADA.entrada_id = NOTAENTRADAITEM.entrada_id "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      SQL = SQL & " and numr_nota = " & txtNOTA.Text
      SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         VALOR_TOTAL_N = TabTemp!ValorTotal
   If TabTemp.State = 1 Then _
      TabTemp.Close

   VALOR_DESCONTO_N = 0

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select isnull(valor_desconto,0) from NOTAENTRADA WITH (NOLOCK)"
   
      SQL = SQL & " where numr_nota = " & txtNOTA.Text
      SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVENDEDOR.EOF Then _
      VALOR_DESCONTO_N = 0 & TabVENDEDOR.Fields(0).Value + VALOR_DESCONTO_N
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   txtValorTotalNota.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
   stBarReq.Panels(8).Text = txtValorTotalNota.Text
   stBarReq.Refresh

   'txtValorTotalNota.Text = Format(VALOR_TOTAL_N - VALOR_DESCONTO_N, strFormatacao2Digitos)
   txtValorTotalNota.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TOTAL_NOTA"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   txtValorDig.Visible = False
   MSFlexGrid1.Clear
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   If Trim(txtNOTA.Text) = "" Then _
      Exit Sub

   Dim Coluna, Linha, Largura_Campo

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select "
   SQL = SQL & " NOTAENTRADAITEM.SEQ_id as Seq, "              '1
   SQL = SQL & " PRODUTO.CODG_PRODUTO as Código, "             '2
   SQL = SQL & " PRODUTO.DESCRICAO as Descrição,"              '3
   SQL = SQL & " NOTAENTRADAITEM.CFOP_ID as CFOP, "            '4

   SQL = SQL & " NOTAENTRADAITEM.QTDE_ENTRADA as Qtde, "       '5
   SQL = SQL & " NOTAENTRADAITEM.PRECO_CUSTO as PreçoCusto,"   '6

   SQL = SQL & " (NOTAENTRADAITEM.QTDE_ENTRADA*"
   SQL = SQL & "  NOTAENTRADAITEM.PRECO_CUSTO) AS TotCusto, "  '7

   SQL = SQL & " NOTAENTRADAITEM.PERC_IPI as IPI,"             '8
   SQL = SQL & " NOTAENTRADAITEM.PERC_ICMS as ICMS, "          '9
   SQL = SQL & " NOTAENTRADAITEM.VALOR_DESCONTO as Desconto, " '10
   SQL = SQL & " NOTAENTRADAITEM.PERC_ICMS_SUB as ICMS_SUB, "  '11
   SQL = SQL & " NOTAENTRADAITEM.PERC_FRETE as Frete,"         '12

   SQL = SQL & " NOTAENTRADAITEM.NCM, "                        '13
   SQL = SQL & " NOTAENTRADAITEM.CST, "                        '14
   SQL = SQL & " NOTAENTRADAITEM.UN, "                         '15

   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID,"                   '16
   SQL = SQL & " NOTAENTRADAITEM.ENTRADA_ID, "                 '17
   SQL = SQL & " NOTAENTRADAITEM.PRODUTO_ID, "                 '18
   SQL = SQL & " NOTAENTRADAITEM.STATUS "                      '19

   SQL = SQL & " from NOTAENTRADAITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where entrada_id =  " & NOTAENTRADA_ID_N

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabConsulta.EOF Then
      INDR_GRAVA = True
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabConsulta.Fields.Count

      ReDim largura_coluna(0 To TabConsulta.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabConsulta.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabConsulta.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabConsulta.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabConsulta.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabConsulta.Fields.Count - 1
            If Coluna = 4 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabConsulta.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna > 4 And Coluna <= 12 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabConsulta.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabConsulta.Fields(Coluna).Value)
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabConsulta.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabConsulta.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabConsulta.MoveNext
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

'SEQ_ID
      MSFlexGrid1.ColWidth(0) = 1000
      MSFlexGrid1.ColAlignment(0) = 0

'CODG_PRODUTO
      MSFlexGrid1.ColWidth(1) = 2000
      MSFlexGrid1.ColAlignment(1) = 0

'DESCRICAO
      MSFlexGrid1.ColWidth(2) = 6000
      MSFlexGrid1.ColAlignment(2) = 0

'CFOP_ID
      MSFlexGrid1.ColWidth(3) = 900
      MSFlexGrid1.ColAlignment(3) = 0

'QTDE_ENTRADA
      MSFlexGrid1.ColWidth(4) = 2000
      MSFlexGrid1.ColAlignment(4) = 7

'PRECO_CUSTO
      MSFlexGrid1.ColWidth(5) = 2000
      MSFlexGrid1.ColAlignment(5) = 7

'TOTAL_ITEM_CUSTO
      MSFlexGrid1.ColWidth(6) = 2000
      MSFlexGrid1.ColAlignment(6) = 7

'PERC_IPI
      MSFlexGrid1.ColWidth(7) = 1500
      MSFlexGrid1.ColAlignment(7) = 0

'PERC_ICMS
      MSFlexGrid1.ColWidth(8) = 1500
      MSFlexGrid1.ColAlignment(8) = 0

'VALOR_DESCONTO
      MSFlexGrid1.ColWidth(9) = 1500
      MSFlexGrid1.ColAlignment(9) = 0

'PERC_ICMS_SUB
      MSFlexGrid1.ColWidth(10) = 1500
      MSFlexGrid1.ColAlignment(10) = 0

'PERC_FRETE
      MSFlexGrid1.ColWidth(11) = 1500
      MSFlexGrid1.ColAlignment(11) = 0

'NCM
      MSFlexGrid1.ColWidth(12) = 600
      MSFlexGrid1.ColAlignment(12) = 0

'CST
      MSFlexGrid1.ColWidth(13) = 600
      MSFlexGrid1.ColAlignment(13) = 0

'UN
      MSFlexGrid1.ColWidth(14) = 600
      MSFlexGrid1.ColAlignment(14) = 0

'FAMILIAPRODUTO_ID
      MSFlexGrid1.ColWidth(15) = 0

'ENTRADA_ID
      MSFlexGrid1.ColWidth(16) = 0

'PRODUTO_ID
      MSFlexGrid1.ColWidth(17) = 0

'STATUS
      MSFlexGrid1.ColWidth(18) = 0
   End If

   ' fecha o recordset e a conexao
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   MOSTRA_TOTAL_NOTA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub CARREGA_COMBO()
'On Error GoTo ERRO_TRATA

   cmbTrans.Clear
   cmbTransAux.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from VWTRANSPORTADORA WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbTrans.AddItem Trim(TabTemp.Fields("descricao").Value) & "-" & Trim(TabTemp.Fields("cnpjcpf").Value)
      cmbTransAux.AddItem Trim(TabTemp!TRANSP_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
   cmbCFOP.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close
   SQL = "select * from CFOP WITH (NOLOCK)"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then
      TabDESCR.MoveFirst
      Do Until TabDESCR.EOF
         DoEvents
         cmbCFOPAux.AddItem TabDESCR.Fields("cfop_id").Value
         cmbCFOP.AddItem TabDESCR.Fields("cfop_id").Value & "-" & TabDESCR!DESCRICAO
         TabDESCR.MoveNext
      Loop
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_COMBO"
End Sub

Private Sub FINANCEIRO_FORM()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = txtCNPJCPF.Text

   If Trim(txtNOTA.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtNOTA.Text) Then _
      Exit Sub
   If Trim(txtSerie.Text) = "" Then _
      txtSerie.Text = 1

   INDR_RECEITA = 2
   TIPO_ENTRADA_N = 1 'com nota fiscal

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select entrada_id from NOTAENTRADA WITH (NOLOCK)"
   SQL = SQL & " where numr_nota = " & txtNOTA.Text
   SQL = SQL & " and serie_nota = '" & Trim(txtSerie.Text) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      NOTAENTRADA_ID_N = 0 & TabNOTA.Fields(0).Value
      frmNOTAENTRADAFINANC.Show 1
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FINANCEIRO_FORM"
End Sub

Private Sub GRAVA_CABEÇA_NOTA_ENTRADA(NUMERO_NOTA_N As Long, _
                                      SERIE_NOTA_A As String, _
                                      FORNEC_ID_N As Long, _
                                      STATUS_NOTA_ENTRADA As String)
'On Error GoTo ERRO_TRATA

   If NUMERO_NOTA_N > 0 And Trim(SERIE_NOTA_A) <> "" And FORNEC_ID_N > 0 Then
      txtCNPJCPF.PromptInclude = False
      txtDtEntrada.PromptInclude = True
      txtDtEmissao.PromptInclude = True
      NOTAENTRADA_ID_N = 0

      If Not IsDate(txtDtEntrada.Text) Then _
         txtDtEntrada.Text = Date

      If Trim(cmbTransAux.Text) = "" Then _
         cmbTransAux.Text = 1

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "select entrada_id from NOTAENTRADA WITH (NOLOCK)"
      SQL = SQL & " where numr_nota = " & NUMERO_NOTA_N
      SQL = SQL & " and serie_nota = '" & Trim(SERIE_NOTA_A) & "'"
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabNOTA.EOF Then
      
         NOTAENTRADA_ID_N = TabNOTA!ENTRADA_ID

         SQL = "UPDATE NOTAENTRADA SET "
         SQL = SQL & " TRANSP_ID = " & Trim(cmbTransAux.Text)
         SQL = SQL & ", TIPOENTRADA_id = " & TIPO_ENTRADA_N
         SQL = SQL & ", usuario_id = " & USUARIO_ID_N
         SQL = SQL & ", NUMR_NOTA = " & NUMERO_NOTA_N
         SQL = SQL & ", serie_nota = '" & Trim(SERIE_NOTA_A) & "'"
         SQL = SQL & ", pedidocompra_id = " & 0
         SQL = SQL & ", fornecedor_id = " & FORNEC_ID_N
         SQL = SQL & ", dt_entrada = '" & DMA(txtDtEntrada.Text) & "'"
         SQL = SQL & ", dt_Emissao = '" & DMA(txtDtEmissao.Text) & "'"
         SQL = SQL & ", Status = '" & STATUS_NOTA_ENTRADA & "'"
         SQL = SQL & ", BASE_CALC_ICMS = " & tpMOEDA(txtBaseCalculoICMS.Text)
         SQL = SQL & ", VALOR_ICMS = " & tpMOEDA(txtValorICMS.Text)
         SQL = SQL & ", PERC_ICMS = " & tpMOEDA(txtPercICMS.Text)
         SQL = SQL & ", BASE_CALC_ICMS_SUBST = " & tpMOEDA(txtBaseICMSSubst.Text)
         SQL = SQL & ", VALOR_ICMS_SUBST = " & tpMOEDA(txtValorICMSSubst.Text)
         SQL = SQL & ", PERC_ICMS_SUBST = " & tpMOEDA(txtPercICMSSubst.Text)
         SQL = SQL & ", VALOR_OUTRAS = " & tpMOEDA(txtValorOutras.Text)
         SQL = SQL & ", valor_frete = " & tpMOEDA(txtFrete.Text)
         SQL = SQL & ", VALOR_IPI = " & tpMOEDA(txtValorIPI.Text)
         SQL = SQL & ", Valor_Desconto = " & tpMOEDA(txtDesconto.Text)
         SQL = SQL & " Where entrada_id = " & NOTAENTRADA_ID_N
         Else
            NOTAENTRADA_ID_N = MAX_ID("entrada_id", "notaentrada", "", "", "", "")

            SQL = "INSERT INTO NOTAENTRADA "
            SQL = SQL & " ("
               SQL = SQL & " estabelecimento_id , ENTRADA_ID, pedidocompra_id, NUMR_NOTA, "
               SQL = SQL & " usuario_id, SERIE_NOTA, fornecedor_id, dt_entrada, dt_Emissao, "
               SQL = SQL & " Status, BASE_CALC_ICMS, VALOR_ICMS , PERC_ICMS, BASE_CALC_ICMS_SUBST, "
               SQL = SQL & " VALOR_ICMS_SUBST, PERC_ICMS_SUBST,  VALOR_OUTRAS, valor_frete, "
               SQL = SQL & " VALOR_IPI, Valor_Desconto, TIPOENTRADA_ID, TRANSP_ID"
            SQL = SQL & ") "
            SQL = SQL & " VALUES ("
               SQL = SQL & ESTABELECIMENTO_ID_N                         'ESTABELECIMENTO_ID
               SQL = SQL & "," & NOTAENTRADA_ID_N                           'ENTRADA_ID
               SQL = SQL & ",0" & 0                                     'pedidocompra_id
               SQL = SQL & "," & NUMERO_NOTA_N
               SQL = SQL & "," & USUARIO_ID_N
               SQL = SQL & ",'" & SERIE_NOTA_A & "'"
               SQL = SQL & "," & FORNEC_ID_N
               SQL = SQL & ",'" & DMA(txtDtEntrada.Text) & "'"
               SQL = SQL & ",'" & Now & "'"
               SQL = SQL & ",'" & STATUS_NOTA_ENTRADA & "'"
               SQL = SQL & "," & tpMOEDA(txtBaseCalculoICMS.Text)
               SQL = SQL & "," & tpMOEDA(txtValorICMS.Text)
               SQL = SQL & "," & tpMOEDA(txtPercICMS.Text)
               SQL = SQL & "," & tpMOEDA(txtBaseICMSSubst.Text)
               SQL = SQL & "," & tpMOEDA(txtValorICMSSubst.Text)
               SQL = SQL & "," & tpMOEDA(txtPercICMSSubst.Text)
               SQL = SQL & "," & tpMOEDA(txtValorOutras.Text)
               SQL = SQL & "," & tpMOEDA(txtFrete.Text)
               SQL = SQL & "," & tpMOEDA(txtValorIPI.Text)
               SQL = SQL & "," & tpMOEDA(txtDesconto.Text)
               SQL = SQL & "," & TIPO_ENTRADA_N
               SQL = SQL & "," & Trim(cmbTransAux.Text)
            SQL = SQL & ")"
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      CONECTA_RETAGUARDA.Execute SQL

      Else: MsgBox "Informe corretamente os dados."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABEÇA_NOTA_ENTRADA"
End Sub

Sub GRAVA_ITEM_NOTA_ENTRADA(NOTA_ENTRADA_ID_N As Long, _
                            SEQ_ID_N As Long, _
                            PROD_ID_N As Long, _
                            PRECO_CUSTO_N As Double, _
                            QTDE_N As Double, _
                            SIT_A As String, _
                            CFOP_N As Integer, _
                            PERC_IPI_N As Double, _
                            PERC_ICMS_N As Double, _
                            VALOR_DESCONTO_N As Double, _
                            PERC_ICMS_SUB_N As Double, _
                            PERC_FRETE_N As Double, _
                            NCM_N As String, _
                            CST_A As String, _
                            UN_A As String)
'On Error GoTo ERRO_TRATA

   If NOTA_ENTRADA_ID_N <= 0 Then _
      Exit Sub
   If PRECO_CUSTO_N <= 0 Then _
      Exit Sub
   If PROD_ID_N <= 0 Then _
      Exit Sub
   If QTDE_N <= 0 Then _
      Exit Sub
   If Trim(SIT_A) = "" Then _
      Exit Sub
   If CFOP_N <= 0 Then _
      Exit Sub
   If Trim(CST_A) = "" Then _
      CST_A = 0
   If Trim(UN_A) = "" Then _
      UN_A = "UN"

   Dim TabItemNota   As New ADODB.Recordset

   If TabItemNota.State = 1 Then _
      TabItemNota.Close

   SQL = "select seq_id from NOTAENTRADAITEM WITH (NOLOCK)"
   SQL = SQL & " where entrada_id = " & NOTA_ENTRADA_ID_N
   SQL = SQL & " and seq_id = " & SEQ_ID_N
   TabItemNota.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabItemNota.EOF Then
      SEQ_ID_N = MAX_ID("seq_id", "NOTAENTRADAITEM", "entrada_id", Trim(NOTA_ENTRADA_ID_N), "", "")

      SQL = "INSERT INTO NOTAENTRADAITEM "
      SQL = SQL & " ("
         SQL = SQL & "ENTRADA_ID,SEQ_id,PRODUTO_ID,PRECO_CUSTO,QTDE_ENTRADA,"
         SQL = SQL & "STATUS,CFOP_ID,PERC_IPI,PERC_ICMS,VALOR_DESCONTO,PERC_ICMS_SUB,PERC_FRETE,"
         SQL = SQL & "NCM,CST,UN"
      SQL = SQL & ")"
      SQL = SQL & " VALUES ("
         SQL = SQL & NOTA_ENTRADA_ID_N                'ENTRADA_ID
         SQL = SQL & "," & SEQ_ID_N                   'SEQ
         SQL = SQL & "," & PROD_ID_N                  'PRODUTO_ID
         SQL = SQL & "," & tpMOEDA(PRECO_CUSTO_N)     'PRECO_CUSTO
         SQL = SQL & "," & tpMOEDA(QTDE_N)            'QTDE_ENTRADA
         SQL = SQL & ",'" & Trim(SIT_A) & "'"         'Status
         SQL = SQL & "," & CFOP_N                     'CFOP_ID
         SQL = SQL & "," & tpMOEDA(PERC_IPI_N)        'PERC_IPI
         SQL = SQL & "," & tpMOEDA(PERC_ICMS_N)       'PERC_ICMS
         SQL = SQL & "," & tpMOEDA(VALOR_DESCONTO_N)  'Valor_Desconto
         SQL = SQL & "," & tpMOEDA(PERC_ICMS_SUB_N)   'PERC_ICMS_SUB
         SQL = SQL & "," & tpMOEDA(PERC_FRETE_N)      'PERC_FRETE
         SQL = SQL & "," & Trim(NCM_N)                'NCM
         SQL = SQL & "," & Trim(CST_A)                'CST
         SQL = SQL & ",'" & Trim(UN_A) & "'"          'UN
      SQL = SQL & ")"
      Else
         SEQ_ID_N = TabItemNota.Fields(0).Value

         SQL = "UPDATE NOTAENTRADAITEM SET "

            SQL = SQL & "PRODUTO_ID = " & PROD_ID_N                        'PRODUTO_ID
            SQL = SQL & ",PRECO_CUSTO = " & tpMOEDA(PRECO_CUSTO_N)         'PRECO_CUSTO
            SQL = SQL & ",QTDE_ENTRADA = " & tpMOEDA(QTDE_N)               'QTDE_ENTRADA
            SQL = SQL & ",STATUS = '" & Trim(SIT_A) & "'"                  'Status
            SQL = SQL & ",CFOP_ID = " & CFOP_N                             'CFOP_ID
            SQL = SQL & ",PERC_IPI = " & tpMOEDA(PERC_IPI_N)               'PERC_IPI
            SQL = SQL & ",PERC_ICMS = " & tpMOEDA(PERC_ICMS_N)             'PERC_ICMS
            SQL = SQL & ",VALOR_DESCONTO = " & tpMOEDA(VALOR_DESCONTO_N)   'Valor_Desconto
            SQL = SQL & ",PERC_ICMS_SUB = " & tpMOEDA(PERC_ICMS_SUB_N)     'PERC_ICMS_SUB
            SQL = SQL & ",PERC_FRETE = " & tpMOEDA(PERC_FRETE_N)           'PERC_FRETE
            SQL = SQL & ",NCM = " & Trim(NCM_N)                            'NCM
            SQL = SQL & ",CST = '" & Trim(CST_A) & "'"                     'CST
            SQL = SQL & ",UN = '" & Trim(UN_A) & "'"                       'UN

         SQL = SQL & " Where entrada_id = " & NOTA_ENTRADA_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
   End If
   If TabItemNota.State = 1 Then _
      TabItemNota.Close
INDR_GRAVA = True
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITEM_NOTA_ENTRADA"
End Sub
' Inicio FUNÇÕES PARA A IMPORTAÇÃO DO XML, CONFORME O CAMINHO DADO PARA O USUARIO
' yuri 11/05/2012
'*****************************************************************************
'Criação: Yuri Grandinetti                                    Data: 11/05/2012
'Propósito: Preenche todos os campos do cabeçalho da nota.
'           A busca do fornecedor é feita pelo CNPJ da mesma. O fornecedor tem
'           que estar cadastrado no sistema com o CNPJ da NF-e!
' Substituir os textbox,etc pelo nossos no megasim, se nao tiver cria como invisivel
'*****************************************************************************
Private Sub PREENCHE_CABEÇA_XML()
'On Error GoTo ERRO_TRATA

   Call BUSCA_FORNEC_XML(0, cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "emit//CNPJ"))                                'Fornecedor
   txtNOTA.Text = Val(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "ide//nNF"))                                       'Número
   txtSerie.Text = Val(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "ide//serie"))                                    'Série
   txtModeloNF.Text = Val(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "ide//mod"))                                   'Modelo
   'lbl_forma.Caption = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "ide//natOp")                                    'Natureza
   txtDtEmissao.PromptInclude = False
      txtDtEmissao.Text = Format(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "ide//dEmi"), "dd/mm/yyyy")             'Data da emissão
   txtDtEmissao.PromptInclude = True
   'txt_chave_acesso.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "infProt//chNFe")                            'Chave de acesso
   txtValorTotalNota.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vNF"), ".", ",")    'Total da nota
   txtBaseCalculoICMS.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vBC"), ".", ",")   'Base Calculo ICMS
   txtValorICMS.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vICMS"), ".", ",")       'Valor ICMS
   txtBaseICMSSubst.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vBCST"), ".", ",")   'Base Subs. Trib.
   txtValorICMSSubst.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vST"), ".", ",")    'Valor ST
   txtValorIPI.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vIPI"), ".", ",")         'Valor IPI
   txtFrete.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vFrete"), ".", ",")          'Valor Frete
   'txt_seguro.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vSeg"), ".", ",")         'Valor Seguro
   txtDesconto.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vDesc"), ".", ",")        'Valor Desconto
   txtValorOutras.Text = Replace(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "total//ICMSTot//vOutro"), ".", ",")    'Valor Outro

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREENCHE_CABEÇA_XML"
End Sub
' Yuri 11/05/2012
' coloque aqui os componentes  relacionado com a tela do megasim
Private Sub LIMPA_FORNEC_XML()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
    'txt_codigo_Fornecedor = ""
    'txt_fornecedor = ""
    'lbl_cgc = ""
    'lbl_inscricao = ""
    'lbl_endereco = ""
    'lbl_cidade = ""
    'lbl_uf = ""
    'lbl_pais = ""
    'x_perc_icms_subst = 0
    'txt_porc_red_icms = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_FORNEC_XML"
End Sub
' 11/05/2012 yuri
'*****************************************************************************
'Criação: Yuri Grandinetti                                         Data: 11/05/2012
'Propósito: Preenche os campos da transportadora com os dados no XML.
'           O sistema busca os dados pelo CNPJ da transp., onde o mesmo tem
'           que estar cadastrado.
'
'*****************************************************************************
Private Sub PREENCHE_TRANSPORTE_XML()
'On Error GoTo ERRO_TRATA

   Call BUSCA_TRANSP_XML(0, cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//transporta//CNPJ"))

    'txt_volume.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//vol//qVol")
    'txt_especie.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//vol//esp")
    'txt_peso_liquido.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//vol//pesoL")
    'txt_peso_bruto.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//vol//pesoB")
    'txt_transportadora_placa.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//veicTransp//placa")
    'txt_transportadora_placa_uf.Text = cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//veicTransp//UF")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREENCHE_TRANSPORTE_XML"
End Sub
'*****************************************************************************
'Criação: Yuri Grandinetti                                     Data: 11/05/2012
'Propósito: Preenche dados da transportadora
' Sugestão caso nao tenha alguns textbox cria como invisivel para o usuario
'*****************************************************************************
Private Sub PREENCHE_TRANSP_XML()
'On Error GoTo ERRO_TRATA

    With gb_Recordset
        'txt_codigo_transportadora = !codigo
        'txt_transportadora_nome = !NOME
        'txt_transportadora_endereco = !Endereco
        'txt_transportadora_bairro = !Bairro
        'txt_transportadora_cidade = !Cidade
        'txt_transportadora_uf = !UF
        'txt_transportadora_cnpj = !CGC
        'txt_transportadora_inscricao_estadual = !inscricao_estadual
    End With

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREENCHE_TRANSP_XML"
End Sub
' YURI 11/05/2012
' aqui voce vai precisar adaptar para o listview do megasim, pois aqui a grade faz referencia ao
' MSHFLEXGRID
'xmlElem.selectSingleNode SÃO OS atributos do xml da notafiscal eletronica
' pega uma nota e voce entendera
' se nao tiver algum elemento no listview, neste caso voce precisara mudar o listview
' OBS .: Acho que voce vai precisar infelizmente trocar o listview pelo componente
' mshflexgrid, pois não tem como sabermos o codigo do produto, pois esse código e do nosso fornecedor
' como ja existe muitos produtos cadastrados nos nossos clientes não temos como amarrar isso, pois não vem no
' xml do cliente
' SUGESTÃO : SE VOCE NÃO CONSEGUIR UMA MANEIRA DE ALTERAR O PRODUTO QUE VAI VIR NO LISTVIEW, INFELIZMENTE VOCE TERA QUE
' TRABALHAR COM ESSE COMPONENTE
' UMA COISA IMPORTANTE CASO QUEIRA TRABALHAR E NÃO RECOMENDO ISSO PELO TRABALHO QUE VAI LHE DAR EU TE PASSO
' A ROTNA DA NOTA FISCAL DE ENTRADA PARA VOCE TER UMA IDÉIA DE COMO FAZER, EU ACHO UM POUCO COMPLEXO, PARA FAZER ISSO
' A TOQUE DE CAIXA
Public Sub PREENCHE_PRODUTO_XML(strCaminhoXML As String)
On Error Resume Next

   Dim lngItem          As Long
   Dim lngLinha         As Long
   Dim XML              As DOMDocument
   Dim xmlElem          As IXMLDOMNode
   Dim bolFimProdutos   As Boolean

   Set XML = New DOMDocument

   lngItem = 0
   lngLinha = 1
   XML.async = False
   bolFimProdutos = False
   If XML.Load(strCaminhoXML) Then

      Do Until bolFimProdutos
         On Error Resume Next
         Set xmlElem = XML.selectNodes("/nfeProc/NFe/infNFe/det").item(lngItem).firstChild

         '*** Verfica se tem algum valor no código do produto, senão tiver finaliza o loop ***
         If xmlElem.selectSingleNode("cProd").Text = "" Then bolFimProdutos = True: Exit Do

'grade1.TextMatrix(lngLinha, 2) = xmlElem.selectSingleNode("cProd").Text                              'Codigo produto
         txtProduto.Text = "" & Trim(xmlElem.selectSingleNode("cProd").Text)                          'Codigo produto

'grade1.TextMatrix(lngLinha, 3) = xmlElem.selectSingleNode("xProd").Text                              'Descrição produto
         txtDesc.Text = "" & Trim(xmlElem.selectSingleNode("xProd").Text)                             'Descrição produto

'grade1.TextMatrix(lngLinha, 5) = Trim$(Format(xmlElem.selectSingleNode("qCom").Text, strFormatacao2Digitos)) 'Quantidade Produto
         txtQTDE.Text = Trim(Replace(xmlElem.selectSingleNode("qCom").Text, ".", ","))   'Quantidade Produto

'grade1.TextMatrix(lngLinha, 6) = Replace(xmlElem.selectSingleNode("vUnCom").Text, ".", ",")          'Valor unitário Produto
         txtPrecoCusto.Text = "" & Trim(Replace(xmlElem.selectSingleNode("vUnCom").Text, ".", ","))   'Valor unitário Produto

         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select produto_id from PRODUTO "
         SQL = SQL & " where descricao = '" & Trim(txtDesc.Text) & "'"
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabProduto.EOF Then
'grade1.TextMatrix(lngLinha, 7) = Replace(xmlElem.selectSingleNode("vDesc").Text, ".", ",")        'Valor Desconto Produto
'grade1.TextMatrix(lngLinha, 8) = Replace(xmlElem.selectSingleNode("vUnCom").Text, ".", ",")       'Valor unitário Produto
'grade1.TextMatrix(lngLinha, 11) = Replace(xmlElem.selectSingleNode("vProd").Text, ".", ",")       'Valor Total Produto

'grade1.TextMatrix(lngLinha, 17) = "NF" ' sergio aqui é so para referenciar que é uma nota fiscal e nao um cupom fiscal
'grade1.TextMatrix(lngLinha, 28) = xmlElem.selectSingleNode("CFOP_id").Text                           'CFOP Produto
'grade1.TextMatrix(lngLinha, 38) = xmlElem.selectSingleNode("cEAN").Text                           'Cód. barras produto
            txtBarras.Text = "" & Trim(xmlElem.selectSingleNode("cEAN").Text)                         'Cód. barras produto

'grade1.TextMatrix(lngLinha, 40) = xmlElem.selectSingleNode("NCM").Text                            'Código NCM Produto
            txtNCM.Text = Trim(xmlElem.selectSingleNode("NCM").Text)                          'Código NCM Produto

'grade1.TextMatrix(lngLinha, 4) = xmlElem.selectSingleNode("uCom").Text                            'Unidade produto
            txtUN.Text = "" & Trim(xmlElem.selectSingleNode("uCom").Text)                          'Unidade produto

            GRAVA_PRODUTO_XML
            Else: PRODUTO_ID_N = 0 & TabProduto.Fields(0).Value
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

txtProduto.Text = PRODUTO_ID_N
Call txtprecocusto_KeyPress(13)

         xmlElem.selectSingleNode("cProd").Text = "" 'Limpa o objeto para setar um novo.

         'LIMPA_BODY

         lngItem = lngItem + 1
         lngLinha = lngLinha + 1
         Err.Clear

      Loop
      Else: MsgBox "Não foi possível abrir o arquivo XML da NFe especificada para Leitura.", vbCritical, "Erro."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREENCHE_PRODUTO_XML"
End Sub


'Yuri 11/05/2012
' CASO NAO EXISTE O FORNECEDOR PELA PESQUISA
' VOCE DEIXA A O CAMPO TXTCNPJCPF EM BRANCO PARA O USUARIO SELECIONAR
' VOCE PODE AUTOMATIZAR ESSE PROCESSO FAZENDO A INCLUSÃO DOS DADOS DO FORNECEDOR QUE ESTA NO XML
' SE NÃO VOCE TERA QUE CONSTRUIR UMA PESQUISA E DA PESQUISAR FAZER O USUARIO CADASTRAR O FORNECEDOR
' POIS AQUI NÃO FOI IMPLEMENTADO
' NÃO LEMBRO SE TEM ESSE VALIDADOR vALIDAeRROS SE NAO TIVER USA O SEU
Private Sub BUSCA_FORNEC_XML(xcodigo As String, Optional strCNPJ As String)
'On Error GoTo ERRO_TRATA

   Dim strCondicao As String

   If strCNPJ <> "" Then
      strCondicao = " cnpjcpf = '" & Trim(strCNPJ) & "'"
      Else: strCondicao = " fornecedor_id = '" & xcodigo & "'"
   End If

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select cnpjcpf,descricao from vwFornecedor WITH (NOLOCK) WHERE " & strCondicao
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFornecedor.EOF Then
      LIMPA_FORNEC_XML
      'AtualTelaFornecedor

      txtCNPJCPF.Text = TabFornecedor.Fields("cnpjcpf").Value
      txtNome.Text = Trim(TabFornecedor.Fields("descricao").Value)

      Call txtCNPJCPF_LostFocus
      Else
         If TabFornecedor.State = 1 Then _
            TabFornecedor.Close

         MsgBox "Fornecedor não cadastrado."

         'frmCADASTROFORNECEDOR.txtCNPJCPF.PromptInclude = False
         '   frmCADASTROFORNECEDOR.txtCNPJCPF.Text = Trim(strCNPJ)
         'frmCADASTROFORNECEDOR.txtCNPJCPF.PromptInclude = True

         'frmCADASTROFORNECEDOR.txtRazao.Text = Trim(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "emit//xnome"))
         'frmCADASTROFORNECEDOR.txtFant.Text = Trim(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "emit//xfant"))

         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaCadastro.Show 1
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   If txtUF.Text = "GO" Then
      cmbCFOPAux.Text = CFOP_ENTRADA_DE
      Else: cmbCFOPAux.Text = CFOP_ENTRADA_FE
   End If
   cmbCFOP.Text = cmbCFOPAux.Text & "-" & TRAZ_CFOP(cmbCFOPAux.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_FORNEC_XML"
End Sub


' 11/05/2012
' aqui voce substitui pelo nosso cadastro de transportadora
Private Sub BUSCA_TRANSP_XML(strCodigo As String, Optional strCNPJ As String)
'On Error GoTo ERRO_TRATA

   Dim strCondicao As String

   If strCNPJ <> "" Then
      strCondicao = " cnpjcpf = '" & Trim(strCNPJ) & "'"
      Else: strCondicao = " transp_id = '" & strCodigo & "'"
   End If

   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

   SQL = "select cnpjcpf,descricao,transp_id from vwTRANSPORTADORA WITH (NOLOCK) WHERE " & strCondicao
   TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Not TabFornecedor.EOF Then
      cmbTransAux.Text = "" & TabFornecedor.Fields("transp_id").Value
      cmbTrans.Text = Trim(TabFornecedor.Fields("descricao").Value)
      Else
         If TabFornecedor.State = 1 Then _
            TabFornecedor.Close

         MsgBox "Transportadora não cadastrada."

         frmPessoaCadastro.txtCNPJCPF.PromptInclude = False
            frmPessoaCadastro.txtCNPJCPF.Text = Trim(strCNPJ)
         frmPessoaCadastro.txtCNPJCPF.PromptInclude = True

         frmPessoaCadastro.txtRazao.Text = Trim(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//transporta//xNome"))
         frmPessoaCadastro.txtNome.Text = Trim(cNotaEntrada.RetornaTagXML((Trim(txtPathXML)), "transp//transporta//xNome"))

         frmPessoaCadastro.Show 1
   End If
   If TabFornecedor.State = 1 Then _
      TabFornecedor.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_TRANSP_XML"
End Sub

Private Sub GRAVA_PRODUTO_XML()
'On Error GoTo ERRO_TRATA

   If txtProduto.Text = "" Then
      MsgBox "Código do produto deve ser informado."
      txtProduto.SetFocus
      Exit Sub
   End If
   If txtDesc.Text = "" Then
      MsgBox "Descrição do produto deve ser informado."
      txtDesc.SetFocus
      Exit Sub
   End If
   If txtPrecoCusto.Text = "" Then
      MsgBox "Preço de custo do produto deve ser informado."
      txtPrecoCusto.SetFocus
      Exit Sub
   End If
   If txtQTDE.Text = "" Then
      MsgBox "Qtde informada inválida."
      txtQTDE.SetFocus
      Exit Sub
   End If

   PRODUTO_ID_N = MAX_ID("produto_id", "produto", "", "", "", "")

   SQL = "insert into PRODUTO "
   SQL = SQL & "("
      SQL = SQL & "EMPRESA_ID,PRODUTO_ID,CODG_PRODUTO,DESCRICAO,FAMILIAPRODUTO_ID,"
      SQL = SQL & "UNIDADE_MEDIDA,CODG_BARRA,SITUACAO,QTDE,SITUACAO_TRIBUTARIA,"
      SQL = SQL & "ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,REFERENCIA,CODG_NCM,"
      SQL = SQL & "COMP_TRIBUTARIA,fornecedor_id,PRECO_CUSTO_ANTERIOR,qtd_ped_anterior,PRECO_CUSTO,"
      SQL = SQL & "PRECO_ATACADO,PRECO_Venda,PERCIVA,DT_CADASTRO,PERC_COMIS,PATH_IMAGEM,ORIGEM_MERCADO,"
      SQL = SQL & "LOCACAO,PRECO_VAREJO_ANTERIOR,PRECO_ATACADO_ANTERIOR,EMBALAGEM,USUARIO_ID,"
      SQL = SQL & "QTD_MINIMO,QTD_MAXIMO,DT_ULT_COMPRA "
   SQL = SQL & ")"
   SQL = SQL & " values ("
      SQL = SQL & EMPRESA_ID_N                                                'EMPRESA_ID
      SQL = SQL & ",0" & PRODUTO_ID_N                                         'PRODUTO_ID
      SQL = SQL & ",'" & Trim(PRODUTO_ID_N) & "'"                             'CODG_PRODUTO
      SQL = SQL & ",'" & Trim(txtDesc.Text) & "'"                             'DESCRICAO
      SQL = SQL & ",0"                                                        'FAMILIAPRODUTO_ID
      SQL = SQL & ",'" & Trim(txtUN.Text) & "'"                               'UNIDADE_MEDIDA
      SQL = SQL & ",'" & Trim(txtBarras.Text) & "'"                           'CODG_BARRA
      SQL = SQL & ",'A'"                                                      'SITUACAO
      SQL = SQL & "," & tpMOEDA(0)                                            'QTDE
      SQL = SQL & ",'00'"                                                     'SITUACAO_TRIBUTARIA
      SQL = SQL & ",17"                                                       'ALIQUOTA_ICMS_NORMAL_DENTRO_UF
      SQL = SQL & ",0"                                                        'PERC_DESCONTO
      SQL = SQL & ",1"                                                        'TIPO_PROD
      SQL = SQL & ",''"                                                       'REFERENCIA
      SQL = SQL & ",'" & Trim(txtNCM.Text) & "'"                              'CODG_NCM
      SQL = SQL & ",0"                                                        'COMP_TRIBUTARIA
      SQL = SQL & ",0" & FORNEC_ID_N                                          'fornecedor_id
      SQL = SQL & ",0"                                                        'PRECO_CUSTO_ANTERIOR
      SQL = SQL & ",0"                                                        'qtd_ped_anterior
      SQL = SQL & "," & tpMOEDA(txtPrecoCusto.Text)                           'PRECO_CUSTO
      SQL = SQL & "," & tpMOEDA(txtPrecoCusto.Text)                           'PRECO_ATACADO
      SQL = SQL & "," & tpMOEDA(txtPrecoCusto.Text)                           'PRECO_Venda
      SQL = SQL & ",0"                                                        'PERCIVA
      SQL = SQL & ",'" & Now & "'"                                      'DT_CADASTRO
      SQL = SQL & "," & tpMOEDA(0)                                            'PERC_COMIS
      SQL = SQL & ",'" & Trim("") & "'"                                       'PATH_IMAGEM
      SQL = SQL & ",0"                                                        'ORIGEM_MERCADO
      SQL = SQL & ",'" & Trim("") & "'"                                       'LOCACAO
      SQL = SQL & ",0"                                                        'PRECO_VAREJO_ANTERIOR
      SQL = SQL & ",0"                                                        'PRECO_ATACADO_ANTERIOR
      SQL = SQL & ",0"                                                        'EMBALAGEM
      SQL = SQL & ",0" & USUARIO_ID_N                                           'USUARIO_ID
      SQL = SQL & "," & tpMOEDA(0)                                            'QTD_MINIMO
      SQL = SQL & "," & tpMOEDA(0)                                            'QTD_MAXIMO
      SQL = SQL & ",'" & Now & "'"                                      'DT_ULT_COMPRA
   SQL = SQL & ")"

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PRODUTO_XML"
   Err.Clear
End Sub

Private Sub GRAVA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   Dim VALOR_IPI_N         As Double
   Dim VALOR_ICMS_N        As Double
   Dim VALOR_DIF_ICMS_SUB  As Currency
   Dim VALOR_ACUMULADO_SUB As Currency
   Dim PERC_ICMS_SUB       As Currency
   Dim VALOR_ICMS_SUB_ITEM As Currency
   Dim VALOR_SUB_CALCULADO As Currency
   Dim VALOR_ICMS_NORMAL   As Currency
   Dim VALOR_NOTA_SEM_SUB  As Currency
   Dim ValorICMSSubst      As Double
   Dim ValorIPI            As Double
   Dim TOTAL_PRECO_CUSTO_N As Double
   Dim TabGravaEstoque         As New ADODB.Recordset

   VALOR_ICMS_N = 0
   VALOR_IPI_N = 0
   VALOR_DESCONTO_N = 0 & txtDesconto.Text
   VALOR_TOTAL_N = 0 & txtValorTotalNota.Text
   ValorICMSSubst = 0 & txtValorICMSSubst.Text
   ValorIPI = 0 & txtValorIPI.Text
   TOTAL_PRECO_CUSTO_N = 0
   VALOR_NOTA_SEM_SUB = VALOR_TOTAL_N - VALOR_DESCONTO_N
   txtValorTotalNota.Text = (VALOR_NOTA_SEM_SUB + ValorICMSSubst + ValorIPI)

  'Pegando Calculo Valor Icms Substituicao se tiver
   If IsNumeric(txtBaseICMSSubst.Text) Then
      If txtBaseICMSSubst.Text > 0 Then
         If VALOR_NOTA_SEM_SUB <= txtValorTotalNota.Text Then
            VALOR_DIF_ICMS_SUB = (txtValorTotalNota.Text - txtValorIPI.Text - VALOR_NOTA_SEM_SUB)
            Else: VALOR_DIF_ICMS_SUB = (VALOR_NOTA_SEM_SUB - txtValorIPI.Text - txtBaseICMSSubst.Text)
         End If
      End If
   End If
   If VALOR_DIF_ICMS_SUB > 0 Then
      VALOR_ACUMULADO_SUB = (VALOR_DIF_ICMS_SUB * 100)
      'Percentual correspondente a valor agregado da substituicao do valor total da compra conforme formula indicada
      PERC_ICMS_SUB = (VALOR_ACUMULADO_SUB / VALOR_NOTA_SEM_SUB)
   End If

   If INDR_AT_VENDA_MKP = True Then
      Msg = "Deseja atualizar preço de custo dos itens no estoque ?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   End If

   If TabGravaEstoque.State = 1 Then _
      TabGravaEstoque.Close

   SQL = "select NOTAENTRADAITEM.ENTRADA_ID, NOTAENTRADAITEM.SEQ_id, "
   SQL = SQL & " NOTAENTRADAITEM.PRODUTO_ID, "
   SQL = SQL & " NOTAENTRADAITEM.PRECO_CUSTO as Custo_Item, "
   SQL = SQL & " NOTAENTRADAITEM.qtde_entrada, NOTAENTRADAITEM.STATUS,"
   SQL = SQL & " NOTAENTRADAITEM.CFOP_id, NOTAENTRADAITEM.PERC_IPI, "
   SQL = SQL & " NOTAENTRADAITEM.PERC_ICMS, NOTAENTRADAITEM.VALOR_DESCONTO,"
   SQL = SQL & " NOTAENTRADAITEM.PERC_ICMS_SUB, NOTAENTRADAITEM.PERC_FRETE, "

   SQL = SQL & " PRODUTO.EMPRESA_ID, PRODUTO.DESCRICAO, "
   SQL = SQL & " produto.preco_custo as Custo_Produto, "
   SQL = SQL & " produto.preco_atacado, produto.preco_venda"

   SQL = SQL & " from NOTAENTRADAITEM WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " and entrada_id = " & NOTAENTRADA_ID_N

   TabGravaEstoque.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabGravaEstoque.EOF
      'rotina para criar o registro na tabela estoque
      RODA_AT_ESTOQUE Trim(TabGravaEstoque.Fields("produto_id").Value), ESTABELECIMENTO_ID_N

      SQL = "update PRODUTO set "
      SQL = SQL & " Dt_Ult_Compra = '" & Now & "'"
      SQL = SQL & " where produto_id = " & Trim(TabGravaEstoque.Fields("produto_id").Value)
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "update ESTOQUE set "
      SQL = SQL & " qtde_estoque = qtde_estoque + " & tpMOEDA(TabGravaEstoque.Fields("qtde_entrada").Value)
      SQL = SQL & " where produto_id = " & Trim(TabGravaEstoque.Fields("produto_id").Value)
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      If RESPOSTA = vbYes And INDR_AT_VENDA_MKP = True Then
         'Calculando Preco de Custo da Mercadoria , Falta Frete
         'VALOR_ICMS_N = TabGravaEstoque!preco_custo * TabGravaEstoque!PERC_ICMS / 100
         VALOR_ICMS_N = TabGravaEstoque.Fields("Custo_item").Value * txtPercICMS.Text / 100
         VALOR_IPI_N = TabGravaEstoque.Fields("custo_item").Value * TabGravaEstoque!PERC_IPI / 100

         'Calculo Preco Custo Mercadoria

         'Calculo Valor Substituicao por item automatico
         If PERC_ICMS_SUB > 0 Then
            VALOR_ICMS_SUB_ITEM = (((TabGravaEstoque.Fields("custo_item").Value + VALOR_IPI_N) * PERC_ICMS_SUB) / 100)
            'VALOR_SUB_CALCULADO = (VALOR_ICMS_SUB_ITEM + (TabGravaEstoque!Preco_Custo + VALOR_IPI_N))
            'VALOR_SUB_CALCULADO = ((VALOR_SUB_CALCULADO * 17) / 100) 'pegar sempre aliquota interna
            'VALOR_ICMS_NORMAL = ((TabGravaEstoque!Preco_Custo * ALIQUOTA_FORNEC) / 100) ' pegar a aliquota externa da nota
            'VALOR_ICMS_SUB_ITEM = (VALOR_SUB_CALCULADO - VALOR_ICMS_NORMAL)
         End If

         'Calculo para achar valor de frete e outros cobre os itens da nota de entrada
         If txtFrete.Text > 0 Then
            VLR_FRETE_N = VALOR_NOTA_SEM_SUB - txtFrete.Text
            VLR_FRETE_N = (VLR_FRETE_N * 100)
            VLR_FRETE_N = (VLR_FRETE_N / VALOR_NOTA_SEM_SUB)
            VLR_FRETE_N = 100 - VLR_FRETE_N
         End If

         If txtValorOutras.Text > 0 Then
            VLR_OUTROS_N = VALOR_NOTA_SEM_SUB - txtValorOutras.Text
            VLR_OUTROS_N = (VLR_OUTROS_N * 100)
            VLR_OUTROS_N = (VLR_OUTROS_N / VALOR_NOTA_SEM_SUB)
            VLR_OUTROS_N = 100 - VLR_OUTROS_N
         End If

         If VLR_FRETE_N > 0 Then _
            VLR_FRETE_N = TabGravaEstoque.Fields("custo_item").Value * VLR_FRETE_N / 100

         If VLR_OUTROS_N > 0 Then _
            VLR_OUTROS_N = TabGravaEstoque.Fields("custo_item").Value * VLR_OUTROS_N / 100

         'If TabGravaEstoque!preco_custo <> TabGravaEstoque!PRECO_CUSTO_ANTERIOR Then
         '   If TabGravaEstoque!PRECO_CUSTO_ANTERIOR > TabGravaEstoque!preco_custo Then
         '      intRetorno_2 = MsgBox("Custo Anterior Maior que Custo Atual,Deseja atualizar preço de custo Pela media de Custo?", vbQuestion + vbYesNo + vbDefaultButton2)
         '      Else
         '         intRetorno_2 = MsgBox("Custo Anterior Menor que Custo Atual,Deseja atualizar preço de custo Pela media de Custo?", vbQuestion + vbYesNo + vbDefaultButton2)
         '   End If
         '   If intRetorno_2 = vbYes Then
         '      TabGravaEstoque!preco_custo = Format(TabGravaEstoque!PRECO_CUSTO_ANTERIOR + TabGravaEstoque!preco_custo / TabGravaEstoque!qtd, strFormatacao2Digitos)
         '   End If
         'End If

         Dim VALOR_TAXA_MARC_VAREJO_N As Double
         Dim VALOR_TAXA_MARC_ATACADO_N As Double

         VALOR_TAXA_VAREJO_N = 0
         VALOR_TAXA_ATACADO_N = 0

         'Busca Taxa Marcacao Varejo
         '1  VAREJO
         If TabTemp.State = 1 Then _
            TabTemp.Close
         SQL = "select sum(perc_taxa) from TAXAMARKUP WITH (NOLOCK)"
         SQL = SQL & " where produto_id = " & Trim(TabGravaEstoque.Fields("produto_id").Value)
         SQL = SQL & "  and tipomercado_id = 1"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               VALOR_TAXA_MARC_VAREJO_N = (100 - TabTemp.Fields(0).Value) / 100  'Calculo de Valores da Taxa de Marcacao
         If TabTemp.State = 1 Then _
            TabTemp.Close

         VALOR_TOTAL_N = 0 ' Zerando para fazer marcacao de atacado

         'Busca Taxa Marcacao atacado
         '2  ATACADO
         If TabTemp.State = 1 Then _
            TabTemp.Close
         SQL = "select sum(perc_taxa) from TAXAMARKUP WITH (NOLOCK)"
         SQL = SQL & " where produto_id = " & Trim(TabGravaEstoque.Fields("produto_id").Value)
         SQL = SQL & " and tipomercado_id = 2"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               VALOR_TAXA_MARC_ATACADO_N = (100 - TabTemp.Fields(0).Value) / 100 'Calculo de Valores da Taxa de Marcacao
         If TabTemp.State = 1 Then _
            TabTemp.Close

'NO CADASTRO DE TAXA DE MARCAÇÃO, O DIFERENCIAL DO PERCENTUAL SE DÁ
'PELO CADASTRO DAS OCORRENCIAS, OU SEJA, TAXAS DISTINTAS PARA PREÇO DE MERCADO

         TOTAL_PRECO_CUSTO_N = TabGravaEstoque.Fields("custo_item").Value + VALOR_IPI_N + VALOR_ICMS_SUB_ITEM + VLR_FRETE_N + VLR_OUTROS_N

         SQL = "UPDATE Produto SET "
         SQL = SQL & " preco_custo = " & tpMOEDA(VALOR_TOTAL_N)

         SQL = SQL & ",PRECO_venda = " & tpMOEDA(TOTAL_PRECO_CUSTO_N / VALOR_TAXA_MARC_VAREJO_N)
         SQL = SQL & ",PRECO_atacado = " & tpMOEDA(TOTAL_PRECO_CUSTO_N / VALOR_TAXA_MARC_ATACADO_N)

         SQL = SQL & ",PRECO_CUSTO_ANTERIOR = " & tpMOEDA(TabGravaEstoque.Fields("Custo_Produto").Value)
         SQL = SQL & ",PRECO_ATACADO_ANTERIOR = " & tpMOEDA(TabGravaEstoque!PRECO_ATACADO)
         SQL = SQL & ",PRECO_VAREJO_ANTERIOR = " & tpMOEDA(TabGravaEstoque!PRECO_Venda)

         SQL = SQL & " Where produto_id = " & TabGravaEstoque.Fields("produto_id").Value
         CONECTA_RETAGUARDA.Execute SQL
      End If

      TabGravaEstoque.MoveNext
   Wend
   If TabGravaEstoque.State = 1 Then _
      TabGravaEstoque.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ESTOQUE"
End Sub

Sub LE_REFERENCIA()
'On Error GoTo ERRO_TRATA

   Dim TabProdFornec As New ADODB.Recordset
   PRODUTO_ID_N = 0

   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

   SQL = "select PRODUTOFORNECEDOR.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO"
   SQL = SQL & " from PRODUTOFORNECEDOR "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PRODUTOFORNECEDOR.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where codg_prod_fornec = '" & Trim(txtRef.Text) & "'"
   SQL = SQL & " and PRODUTOFORNECEDOR.fornecedor_id = " & FORNEC_ID_N
   TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProdFornec.EOF Then
      txtProduto.Text = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
      PRODUTO_ID_N = 0 & Trim(TabProdFornec.Fields("produto_id").Value)
      Else
         If TabProdFornec.State = 1 Then _
            TabProdFornec.Close

         SQL = "select codg_produto,descricao,produto_id,preco_custo,codg_ncm,unidade_medida,aliquota_icms,situacao_tributaria "
         SQL = SQL & " from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where referencia = '" & Trim(txtRef.Text) & "'"
         SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N
         SQL = SQL & " and situacao <> 'C' "
         TabProdFornec.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProdFornec.EOF Then
            txtProduto.Text = "" & Trim(TabProdFornec.Fields("codg_produto").Value)
            PRODUTO_ID_N = 0 & Trim(TabProdFornec.Fields("produto_id").Value)
         End If
   End If
   If TabProdFornec.State = 1 Then _
      TabProdFornec.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_REFERENCIA"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = "" Then _
      Exit Sub

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then _
      Exit Sub

   txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtProduto.Text = Trim(CODG_PRODUTO_A)
   txtDesc.Text = DESC_PRODUTO_A

   txtUN.Text = "" & UNIDADE_MEDIDA_A
   txtICMSItem.Text = "" & ALIQUOTA_ICMS_N
   txtCST.Text = "" & SITUACAO_TRIBUT_A
   txtNCM.Text = "" & CODG_NCM_A
   stBarReq.Panels(4).Text = "" & PR_CUSTO_PRODUTO_N
   txtPrecoCusto.Text = "" & PR_CUSTO_PRODUTO_N

   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = MAX_ID("seq_id", "NOTAENTRADAITEM", "entrada_id", Trim(NOTAENTRADA_ID_N), "", "")
      txtSeq.Text = SEQ_ID_N
   End If

   MOSTRA_ITEM_NOTA_ENTRADA NOTAENTRADA_ID_N, txtSeq.Text
   txtNCM.SetFocus

   If TabProduto.State = 1 Then _
      TabProduto.Close

   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub
