VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A09ABE67-EA60-4BCF-892E-35160344BEC0}#1.0#0"; "SGI.ocx"
Begin VB.Form frmConsignacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processo de Consignação"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   1155
      Left            =   1920
      TabIndex        =   17
      Top             =   0
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "' Cliente/Empresa"
      ShadowStyle     =   1
      Begin VB.CommandButton cmdEmpresa 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   555
         Width           =   255
      End
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   5
         Top             =   510
         Width           =   975
      End
      Begin VB.CommandButton cmdPesquisaCliente 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   3
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "&2 - Empresa.:"
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   555
         Width           =   945
      End
      Begin VB.Label lblEmitente 
         Alignment       =   1  'Right Justify
         Caption         =   "&1 - Cliente.:"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   255
         Width           =   795
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1155
      Left            =   30
      TabIndex        =   18
      Top             =   0
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  Período  "
      ShadowStyle     =   1
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   315
         Left            =   570
         TabIndex        =   0
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   78512129
         CurrentDate     =   64228
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   315
         Left            =   570
         TabIndex        =   1
         Top             =   690
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   78512129
         CurrentDate     =   37565
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "&Inicial:"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "&Final:"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   750
         Width           =   405
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1155
      Left            =   4080
      TabIndex        =   22
      Top             =   0
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Opções da Nota"
      ShadowStyle     =   1
      Begin VB.TextBox txtNota 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   8
         Top             =   825
         Width           =   975
      End
      Begin VB.TextBox txtSequencia 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   7
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox txtPedido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "&5 - Nota.:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "&4 - Seq.:"
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   555
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&3 - Pedido.:"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   255
         Width           =   825
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1155
      Left            =   9270
      TabIndex        =   26
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tipo Consulta"
      ShadowStyle     =   1
      Begin Threed.SSOption optSintetico 
         Height          =   240
         Left            =   210
         TabIndex        =   9
         Top             =   330
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         _Version        =   262144
         Caption         =   "Sintetico"
      End
      Begin Threed.SSOption optAnalitico 
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   690
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   262144
         ForeColor       =   0
         Caption         =   "Analitico"
         Value           =   -1
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   1155
      Left            =   7500
      TabIndex        =   27
      Top             =   0
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Situação"
      ShadowStyle     =   1
      Begin VB.TextBox txtPrazo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   12
         Text            =   "60"
         Top             =   540
         Width           =   315
      End
      Begin Threed.SSOption SSOption3 
         Height          =   240
         Left            =   1020
         TabIndex        =   28
         Top             =   2700
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         _Version        =   262144
         Caption         =   "Sintetico"
      End
      Begin Threed.SSOption optPrazo 
         Height          =   225
         Left            =   90
         TabIndex        =   11
         Top             =   540
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   262144
         Caption         =   "Acima de"
      End
      Begin Threed.SSOption optMostra 
         Height          =   270
         Left            =   90
         TabIndex        =   47
         Top             =   840
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   262144
         Caption         =   "Mostra Todos"
      End
      Begin Threed.SSOption optDiferenca 
         Height          =   270
         Left            =   90
         TabIndex        =   48
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         _Version        =   262144
         Caption         =   "Diferença  >  0"
         Value           =   -1
      End
      Begin VB.Label Label6 
         Caption         =   "Dias"
         Height          =   195
         Left            =   1380
         TabIndex        =   29
         Top             =   570
         Width           =   315
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   1155
      Left            =   10710
      TabIndex        =   30
      Top             =   0
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Executar Processos"
      ShadowStyle     =   1
      Begin VB.CommandButton cmdLimpar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Limpar"
         Height          =   615
         Left            =   885
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   210
         Width           =   735
      End
      Begin VB.CommandButton cmdAjuda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ajuda"
         Height          =   615
         Left            =   3705
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   210
         Width           =   735
      End
      Begin VB.CheckBox chkFatura 
         Caption         =   "Fatura Tudo"
         Height          =   195
         Left            =   1470
         TabIndex        =   46
         Top             =   900
         Width           =   1335
      End
      Begin VB.CheckBox chkDevolve 
         Caption         =   "Devolve Tudo"
         Height          =   225
         Left            =   60
         TabIndex        =   45
         Top             =   900
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   210
         Width           =   735
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   795
      End
      Begin VB.CommandButton cmdDevolver 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Devolve"
         Height          =   615
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   210
         Width           =   735
      End
      Begin VB.CommandButton cmdFaturar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Faturar"
         Height          =   615
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   735
      End
   End
   Begin Threed.SSFrame SSFrame7 
      Height          =   795
      Left            =   -480
      TabIndex        =   32
      Top             =   7650
      Width           =   15690
      _ExtentX        =   27675
      _ExtentY        =   1402
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.Label lblvlrTotItem 
         Height          =   15
         Left            =   13080
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblvlrLiqItem 
         Height          =   135
         Left            =   10320
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblvlrUniItem 
         Height          =   15
         Left            =   8040
         TabIndex        =   67
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblBICMS 
         Height          =   15
         Left            =   3120
         TabIndex        =   66
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblICMS 
         Height          =   135
         Left            =   1080
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblIPI 
         Height          =   15
         Left            =   600
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblvlrfaturarIPI 
         Height          =   135
         Left            =   5310
         TabIndex        =   63
         Top             =   330
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblvlrfaturar1 
         Height          =   30
         Left            =   8070
         TabIndex        =   62
         Top             =   390
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblVlrConsig 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1710
         TabIndex        =   61
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblVlrFaturar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4020
         TabIndex        =   60
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblVlrDevolver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   9030
         TabIndex        =   59
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblVlrConsigDesc 
         Caption         =   "Vlr.  Consig.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   58
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label Label21 
         Caption         =   "Vlr. Faturar.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   57
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Vlr. Devolver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7710
         TabIndex        =   56
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label19 
         Caption         =   "Vlr. Faturado.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5220
         TabIndex        =   55
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblVlrFaturado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   6510
         TabIndex        =   54
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label Label17 
         Caption         =   "Vlr. Devolvida.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   53
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label lblVlrDevolvido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   11640
         TabIndex        =   52
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label label52 
         Caption         =   "Vlr. Diferença.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12840
         TabIndex        =   51
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label lblVlrDiferenca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   14160
         TabIndex        =   50
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label lblQtdDifer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   14160
         TabIndex        =   44
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label Label11 
         Caption         =   "Qtd Diferença:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12810
         TabIndex        =   43
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblQtdDevolvida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   11640
         TabIndex        =   42
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   "Qtd Devolvida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10170
         TabIndex        =   41
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblQtdFaturada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   6510
         TabIndex        =   40
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label Label12 
         Caption         =   "Qtd Faturada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5190
         TabIndex        =   39
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Qtd  Devolver:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   38
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "Qtd  Faturar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblQtdConsigDesc 
         Caption         =   "Qtd  Consig.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   36
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblQtdDevolver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   9030
         TabIndex        =   35
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblQtdFaturar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4020
         TabIndex        =   34
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblQtdConsig 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E9D6C7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1710
         TabIndex        =   33
         Top             =   60
         Width           =   1095
      End
   End
   Begin UltraGrid.SSUltraGrid GridVaca 
      Height          =   6045
      Left            =   0
      TabIndex        =   49
      Top             =   1530
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10663
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108868
      RowConnectorColor=   12632256
      Caption         =   "Consulta Consignacao"
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   8
      ResizeFonts     =   0   'False
      DesignWidth     =   15240
      DesignHeight    =   8460
   End
   Begin Threed.SSFrame SSFrame8 
      Height          =   1155
      Left            =   6000
      TabIndex        =   71
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2037
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tipo Consulta"
      ShadowStyle     =   1
      Begin Threed.SSFrame SSFrame9 
         Height          =   1155
         Left            =   90
         TabIndex        =   72
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2037
         _Version        =   262144
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tipo Consulta"
         ShadowStyle     =   1
         Begin Threed.SSOption optTeste 
            Height          =   210
            Left            =   60
            TabIndex        =   73
            Top             =   210
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   370
            _Version        =   262144
            Caption         =   "Teste"
         End
         Begin Threed.SSOption optConsignado 
            Height          =   285
            Left            =   60
            TabIndex        =   74
            Top             =   405
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            _Version        =   262144
            ForeColor       =   0
            Caption         =   "Consignado"
            Value           =   -1
         End
         Begin Threed.SSOption optEmprestimo 
            Height          =   225
            Left            =   60
            TabIndex        =   75
            Top             =   660
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   262144
            ForeColor       =   0
            Caption         =   "Emprestimo"
         End
         Begin Threed.SSOption optTodos 
            Height          =   225
            Left            =   60
            TabIndex        =   76
            Top             =   900
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   397
            _Version        =   262144
            ForeColor       =   0
            Caption         =   "Todos"
         End
      End
   End
   Begin Threed.SSFrame SSFrame10 
      Height          =   375
      Left            =   0
      TabIndex        =   77
      Top             =   1170
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   661
      _Version        =   262144
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.CommandButton cmdProduto 
         BackColor       =   &H00EBC8AB&
         Height          =   255
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   60
         Width           =   255
      End
      Begin SGI.Pesquisa Vendedor 
         Height          =   315
         Left            =   -210
         TabIndex        =   79
         Top             =   30
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   556
         Caption         =   "Vendedor:"
         enabled         =   -1  'True
      End
      Begin MSAdodcLib.Adodc adoProduto 
         Height          =   330
         Left            =   7770
         Top             =   30
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcProduto 
         Height          =   315
         Left            =   6180
         TabIndex        =   80
         Top             =   30
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descricao"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.TextBox txtProduto 
         Height          =   315
         Left            =   5040
         TabIndex        =   81
         Top             =   30
         Width           =   1125
      End
      Begin VB.Label lblEmprestimo 
         Caption         =   "78-Emprestimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12960
         TabIndex        =   85
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label lblteste 
         Caption         =   "71-Teste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12030
         TabIndex        =   84
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Produto:"
         Height          =   165
         Left            =   4200
         TabIndex        =   83
         Top             =   90
         Width           =   765
      End
      Begin VB.Label lblConsig 
         Caption         =   "70-Consig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10920
         TabIndex        =   82
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmConsignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGrid As New ADODB.Recordset
Dim rsGridTeste As New ADODB.Recordset
Dim EstoqueSalva As Boolean
Dim IntTimTesteSaida As Integer
Dim IntTimConsigSaida As Integer
Dim IntTimTesteVenda As Integer
Dim IntTimConsigVenda As Integer
Dim IntTimTesteDevolve As Integer
Dim IntTimConsigDevolve As Integer
Dim IntTimEmprestimoSaida As Integer
Dim IntTimEmprestimoDevolve As Integer
Dim IntTimEmprestimoVenda As Integer
Dim strTipoMovimento As String
Public intSequencia As Long
Dim intSequenciaItem As Long
Dim Devolver As Boolean
Dim DevolverTeste As Boolean, DevolverEmprestimo As Boolean
Dim Faturar As Boolean
Dim FaturarTeste As Boolean, FaturarEmprestimo As Boolean
Dim VlrIPI As Double
Dim VlrICMS As Double
Dim IntNotaAnterior As Long
Dim VlrUnitario As Double, VlrFaturar As Double, VlrFaturarIPI As Double, VlrDevolver As Double, VlrDevolverIPI As Double, VlrDevolverICMS As Double
Dim VlrUnitarioDevolucao As Double
Dim VlrIPITotal As Double
Dim VlrICMSTotal As Double
Dim BaseICMS As Double
Dim qtdfaturar As Long
Dim qtdDevolver As Long
Dim IntClienteAnterior As Long, SequenciaAnterior As Long
Dim intCodigoProduto As Long
Dim QtdFaturaItem As Long
Dim strNumeroNota As String
Dim strMsgNotaImposto As String
Dim strMsgNota As String
Dim aCell As SSCell
Dim CodigoDepositoTIM As Integer, EstoqueTIM As String
Dim ValorIPIItem As Double, ValorICMSITEM As Double
Dim booCalculaIPI As Boolean
Dim booVendedor As Boolean
Dim booContagem As Boolean 'Para verificar se vai utilizar contagem na consignacao
Dim TipoConsignacao As String
Dim ValorComDesconto As Double, ValorDesconto As Double, ValorComDescontoAcumulado As Double
Dim ValorUnitarioComDesconto As Double
Dim ValorDescontoAcumulado As Double
Dim ValorLimiteVendaConsignacao As Integer
Dim booConsultaCodigo As Boolean
Dim rs2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim intSequenciaProduto As Integer
Dim booQtdEstoque As Boolean
Dim CodigoProdutoAnterior As Long

Dim booAtualizaProduto As Boolean

Sub VerificaBandoDados()
'Emanoel - 19/04/2010
    VerificaParametro "TIM_CONSIG_SAIDA", "Saida de Consignacao", "N", "69", ""
    VerificaParametro "TIM_CONSIG_VENDA", "Venda de Consignacao", "N", "77", ""
    VerificaParametro "TIM_CONSIG_DEVOLVE", "Devolucao de Consignacao", "N", "81", ""
    VerificaParametro "TIM_TESTE_SAIDA", "Saida de Teste", "N", "70", ""
    VerificaParametro "TIM_TESTE_VENDA", "Venda de Teste", "N", "63", ""
    VerificaParametro "TIM_TESTE_DEVOLVE", "Devolucao de Teste", "N", "65", ""
    VerificaParametro "TIM_EMPRESTIMO_SAIDA", "Saida de Emprestimo", "N", "90", ""
    VerificaParametro "TIM_EMPRESTIMO_VENDA", "Venda de Emprestimo", "N", "89", ""
    VerificaParametro "TIM_EMPRESTIMO_DEVOLVE", "Devolucao de Emprestimo", "N", "91", ""
    VerificaParametro "CALCULA_IPI_CONSIGNACAO", "Faturar com IPI SIM ou NAO", "B", "True", ""
    VerificaParametro "LIMITE_CONSIGNACAO_VENDA", "Limite de Valor Para Faturar Consignacao", "N", "100"
    VerificaParametro "CONTAGEM_CONSIG", "Usa Contagem na Consignação SIM ou NAO", "B", "False", ""
End Sub

Private Sub chkTeste_Click()
   cmdPesquisar_Click
End Sub

Private Sub chkTeste_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub chkConsignacao_Click()
   cmdPesquisar_Click
End Sub

Private Sub chkConsignacao_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub chkDevolve_Click()
On Error GoTo TrataErros
    If chkDevolve.Value = 1 Then
        If rsGrid.RecordCount > 0 Then
           rsGrid.MoveFirst
           Do While Not rsGrid.EOF
              DoEvents:
              db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & rsGrid("QtdDifer") & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
              rsGrid.MoveNext
           Loop
        End If
    Else
        If rsGrid.RecordCount > 0 Then
           rsGrid.MoveFirst
           Do While Not rsGrid.EOF
              DoEvents:
              db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
              rsGrid.MoveNext
           Loop
        End If
    End If
   cmdPesquisar_Click
   
   Exit Sub
TrataErros:
    Me.MousePointer = 0
    If Err.Number = -2147217842 Then
    
    ElseIf Err.Number <> 0 Then
        ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
        Exit Sub
    End If
End Sub

Private Sub chkFatura_Click()
On Error GoTo TrataErros
   If chkFatura.Value = 1 Then
      If rsGrid.RecordCount > 0 Then
         rsGrid.MoveFirst
         Do While Not rsGrid.EOF
            DoEvents:
            db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & rsGrid("QtdDifer") & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
            rsGrid.MoveNext
         Loop
         cmdPesquisar_Click
      End If
      Else
          If rsGrid.RecordCount > 0 Then
             rsGrid.MoveFirst
             Do While Not rsGrid.EOF
                DoEvents:
                db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
             rsGrid.MoveNext
             Loop
             cmdPesquisar_Click
          End If
   End If
   
   Exit Sub
TrataErros:
    Me.MousePointer = 0
    If Err.Number = -2147217842 Then
    
    ElseIf Err.Number <> 0 Then
        ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
        Exit Sub
    End If
End Sub

Private Sub cmdAjuda_Click()
    strRetorno = "AjudaConsignacao"
    frmAjudaPDF.Show
End Sub

Private Sub cmdDevolver_Click()
    On Error GoTo TrataErros
    Dim strMensagem As String
    ConsultaEstoque
    If booQtdEstoque = True Then
        If rsGrid.RecordCount > 0 Then
            Faturar = False
            FaturarTeste = False
            FaturarEmprestimo = False
            
            If optTeste.Value = True Then
               DevolverTeste = True
               Devolver = False
               DevolverEmprestimo = False
            ElseIf optConsignado.Value = True Then
                  DevolverTeste = False
                  DevolverEmprestimo = False
                  Devolver = True
            ElseIf optEmprestimo.Value = True Then
                  DevolverTeste = False
                  DevolverEmprestimo = True
                  Devolver = False
            End If
            Else
                MsgBox "Nao tem nenhum registro para Devolver, Verifique!"
                Exit Sub
        End If
        strTipoMovimento = "DV"
        
        If MsgBox("Ao clicar em SIM o Sistema ira Devolver Quantidades Solicitadas, CONFIRMA? ", 36, Me.Caption) = vbNo Then
          Exit Sub
          Else
             Screen.MousePointer = 0
             strRetorno = "Consignacao"
             frmMotivos.Show 1
             
             If Gravar = True Then
                'Cancelar
             End If
        End If
        MsgBox "Devolucao Efetuada com Sucesso...", 48, Me.Caption
        Me.MousePointer = 0
        cmdPesquisar_Click
    End If
    cmdPesquisar_Click
    Exit Sub
TrataErros:
    Me.MousePointer = 0
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
End Sub

Private Sub cmdFaturar_Click()
   On Error GoTo TrataErros
   ConsultaEstoque
   If booQtdEstoque = True Then
        If rsGrid.RecordCount > 0 Then
             Devolver = False
             DevolverTeste = False
             DevolverEmprestimo = False
             
             If optTeste.Value = True Then
                FaturarTeste = True
                FaturarEmprestimo = False
                Faturar = False
             ElseIf optConsignado.Value = True Then
                   FaturarTeste = False
                   FaturarEmprestimo = False
                   Faturar = True
             ElseIf optEmprestimo.Value = True Then
                   FaturarTeste = False
                   FaturarEmprestimo = True
                   Faturar = False
             End If
        Else
            MsgBox "Nao tem nenhum Registro Para Faturar, Verifique!"
            Exit Sub
        End If
        
        strTipoMovimento = "PV"
        
        If MsgBox("Ao clicar em SIM o sistema ira Faturar Quantidades Solicitada, CONFIRMA? ", 36, Me.Caption) = vbNo Then
           Exit Sub
        Else
            Screen.MousePointer = 0
            strRetorno = "Consignacao"
            frmMotivos.Show 1
            
            If Gravar = False Then
               MsgBox "Confirme Valor a Faturar...", 48, Me.Caption
               Exit Sub
            End If
        End If
        
        MsgBox "Quantidade Faturada com sucesso...", 48, Me.Caption
        Me.MousePointer = 0
        cmdPesquisar_Click
   End If
   cmdPesquisar_Click
   Exit Sub

TrataErros:
    Me.MousePointer = 0
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo TrataErros
    Dim strInicial As String
    Dim strFinal As String
    Dim strRelatorio As String
    Dim strNomeRelatorio As String
    Dim strPrazo As String
    Dim strimpressoa As String
    Dim strParametro As String
    Dim strNomeRelatorioParametro As String
    If rsGrid.RecordCount <= 0 Then
       MsgBox "Nao tem Nenhum Registro para Ser Impresso, Verifique!"
       Exit Sub
    End If
    If strimpressoa = "" Then
        strRetorno1 = strNomeRelatorio
        frmImpressora.Show 1
    End If
    
    If strRetorno = "" Then Exit Sub
    strimpressoa = strRetorno
    strNomeRelatorio = App.Path & "\Relatorio\RelConsignacao.rpt"
    
    If optTeste.Value = True Then
       strNomeRelatorioParametro = "Relatorio Teste"
    ElseIf optConsignado.Value = True Then
       strNomeRelatorioParametro = "Relatorio Consignacao"
    ElseIf optEmprestimo.Value = True Then
       strNomeRelatorioParametro = "Relatorio Emprestimos"
    ElseIf optTodos.Value = True Then
       strNomeRelatorioParametro = "Relatorio Todos"
    End If
    strParametro = "Período de   " & DMA(dtpInicial.Value) & " à " & DMA(dtpFinal.Value)
    
    Set crxReport = crxApplication.OpenReport(strNomeRelatorio)
    crxReport.DiscardSavedData
     
    If frmImpressora.chkConexao.Value = 1 Then
        GS_Define_Usuario_Senha
    End If
    If optSintetico.Value Then
        strNomeRelatorioParametro = strNomeRelatorioParametro & "-Sintetico"
    ElseIf optAnalitico.Value = True Then
        strNomeRelatorioParametro = strNomeRelatorioParametro & "-Analitico"
    End If

    LS_Envia_Formula "{@NomeRelatorio}", "'" & strNomeRelatorioParametro & "'"
    LS_Envia_Formula "{@Parametro}", "'Período de " & DMA(dtpInicial.Value) & " a " & DMA(dtpFinal.Value) & "'"
    
    
    LS_Envia_Parametro "{?@CodigoEmpresa}", "" & intEmpresa
    LS_Envia_Parametro "{?@CodigoFilial}", IIf(IsNumeric(txtEmpresa.Text) = True, txtEmpresa.Text, "0")
    LS_Envia_Parametro "{?@DataInicial}", mdaI(dtpInicial.Value)
    LS_Envia_Parametro "{?@DataFinal}", mdaF(dtpFinal.Value)
    LS_Envia_Parametro "{?@CodigoCliente}", IIf(IsNumeric(txtCliente.Text) = True, txtCliente.Text, "0")
    LS_Envia_Parametro "{?@Sequencia}", IIf(IsNumeric(txtSequencia.Text) = True, txtSequencia.Text, "0")
    LS_Envia_Parametro "{?@Pedido}", IIf(IsNumeric(txtPedido.Text) = True, txtPedido.Text, "0")
    LS_Envia_Parametro "{?@Nota}", IIf(IsNumeric(txtNota.Text) = True, txtNota.Text, "0")
    LS_Envia_Parametro "{?@Situacao}", IIf(optDiferenca.Value = True, "D", IIf(optPrazo.Value = True, "P", "T"))
    LS_Envia_Parametro "{?@Dias}", IIf(IsNumeric(txtPrazo.Text) = True, txtPrazo.Text, "0")
    LS_Envia_Parametro "{?@TipoSinteticoAnalitico}", IIf(optSintetico.Value = True, "S", "A")
    LS_Envia_Parametro "{?@Vendedor}", IIf(IsNumeric(Vendedor.Text) = True, Vendedor.Text, "0")
    LS_Envia_Parametro "{?@Produto}", IIf(IsNumeric(txtProduto.Text) = True, txtProduto.Text, "0")
    If optTeste.Value = True Then
       LS_Envia_Parametro "{?@RelatorioConsignadoTeste}", "T"
    ElseIf optConsignado.Value = True Then
       LS_Envia_Parametro "{?@RelatorioConsignadoTeste}", "C"
    ElseIf optEmprestimo.Value = True Then
       LS_Envia_Parametro "{?@RelatorioConsignadoTeste}", "E"
    ElseIf optTodos.Value = True Then
       LS_Envia_Parametro "{?@RelatorioConsignadoTeste}", "D"
    End If
    
    If strimpressoa = "VIDEO" Then
        Dim frmRel As New frmRelatorio
    
        frmRel.crvRelatorio.ReportSource = crxReport
        frmRel.Caption = "Relatório de Consignacao"
        frmRel.crvRelatorio.ViewReport
        frmRel.crvRelatorio.Zoom 105
        frmRel.crvRelatorio.DisplayGroupTree = True
        frmRel.Show
    Else
        crxReport.SelectPrinter strimpressoa, strimpressoa, ""
        crxReport.PrintOut False
    End If
    
    Exit Sub


TrataErros:
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
End Sub

Private Sub cmdLimpar_Click()
    cmdLimpar.Enabled = False
    If chkFatura.Value = 1 Then chkFatura.Value = 0
    If chkDevolve.Value = 1 Then chkDevolve.Value = 0
    chkFatura_Click
    chkDevolve_Click
    cmdLimpar.Enabled = True
End Sub

Private Sub cmdProduto_Click()
    strRetorno = ""
    frmConsultaPrecoEstoque.Show 1
    If strRetorno <> "" Then
        booConsultaCodigo = True
        txtProduto.Text = strRetorno
        txtProduto_KeyPress 13
    End If
End Sub
Private Sub dtcProduto_Click(Area As Integer)
    If Not IsNumeric(dtcProduto.BoundText) Then
        Exit Sub
    Else
        If dtcProduto.BoundText <> 0 Then
            txtProduto.Text = dtcProduto.BoundText
            booConsultaCodigo = True
        End If
    End If
End Sub

Private Sub dtcProduto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo TrataErros
    If KeyCode = 114 Then
        Screen.MousePointer = 11
        PreencheObjetos adoProduto, "SELECT Produto.CodigoProduto as Codigo, Produto.CodigoProdutoTransferencia + '-' + Produto.Descricao as Descricao FROM Produto INNER JOIN TabelaPrecoProduto ON Produto.CodigoEmpresa = TabelaPrecoProduto.CodigoEmpresa AND Produto.CodigoProduto = TabelaPrecoProduto.CodigoProduto WHERE Produto.CodigoProduto <> 99 and TabelaPrecoProduto.CodigoEmpresa = " & intEmpresa & " and Produto.Descricao like '%" & dtcProduto.Text & "%' AND (TabelaPrecoProduto.CodigoTabelaPreco = " & 0 & ") ORDER BY Produto.Descricao", dtcProduto
        Screen.MousePointer = 0
    End If
    Exit Sub
    
TrataErros:
    Screen.MousePointer = 0
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
End Sub

Private Sub dtcProduto_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub dtcProduto_LostFocus()
    'txtProdutoAcabado_KeyPress 13
End Sub


Private Sub GRID_Click()
   For i = 0 To Grid.Bands(0).Columns.Count - 1
      'grid.Bands(0).Columns(1).Key
   Next i
End Sub


Private Sub optTodos_Click(Value As Integer)
    lblQtdConsigDesc.Caption = "Qtd. Todos.:"
    lblVlrConsigDesc.Caption = "Vlr. Todos.:"
    cmdDevolver.Enabled = False
    cmdFaturar.Enabled = False
    cmdPesquisar_Click
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 80 Or KeyCode = 112 Then
       cmdProduto_Click
    End If
End Sub

Private Sub Form_Load()
   Dim DataInicial As Date
    DataInicial = "01/01/2008"
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 20
    VerificaBandoDados
    
    dtpInicial.Value = DMA(DateAdd("d", 0, DataInicial))
    dtpFinal.Value = DMA(Date)
    
    'Parametros de Emprestimo
    IntTimTesteSaida = PegaValorParametro("TIM_TESTE_SAIDA") ' 70
    IntTimConsigSaida = PegaValorParametro("TIM_CONSIG_SAIDA") ' 71
    IntTimEmprestimoSaida = PegaValorParametro("TIM_EMPRESTIMO_SAIDA") ' 78
    'Parametros de Venda
    IntTimTesteVenda = PegaValorParametro("TIM_TESTE_VENDA") ' 71
    IntTimConsigVenda = PegaValorParametro("TIM_CONSIG_VENDA") '77
    IntTimEmprestimoVenda = PegaValorParametro("TIM_EMPRESTIMO_VENDA") '90
    'Parametros de Devolucao
    IntTimTesteDevolve = PegaValorParametro("TIM_TESTE_DEVOLVE") ' 78
    IntTimConsigDevolve = PegaValorParametro("TIM_CONSIG_DEVOLVE") ' 81
    IntTimEmprestimoDevolve = PegaValorParametro("TIM_EMPRESTIMO_DEVOLVE") ' 91
    
    booCalculaIPI = PegaValorParametro("CALCULA_IPI_CONSIGNACAO") 'Verifica se Calula ou nao o IPI
    booContagem = PegaValorParametro("CONTAGEM_CONSIG") 'Verifica se vai usar Contagem na Consginacao ou nao
    
    
    PreencheObjetos adoProduto, "SELECT Produto.CodigoProduto as Codigo, Produto.CodigoProdutoTransferencia + '-' + Produto.Descricao as Descricao FROM Produto INNER JOIN TabelaPrecoProduto ON Produto.CodigoEmpresa = TabelaPrecoProduto.CodigoEmpresa AND Produto.CodigoProduto = TabelaPrecoProduto.CodigoProduto WHERE Produto.CodigoProduto <> 99  and TabelaPrecoProduto.CodigoEmpresa = " & intEmpresa & " and CodigoTipoUso <> 4 and codigotipouso <> 6 and codigotipouso <> 3 and Produto.Descricao like '%" & dtcProduto.Text & "%' AND (TabelaPrecoProduto.CodigoTabelaPreco = " & 0 & ")  ORDER BY Produto.Descricao", dtcProduto
    Vendedor.CarregaGrafico "Funcionario", "SELECT CodigoFuncionario as Codigo, Nome as Descricao FROM Funcionario WHERE (CodigoEmpresa = " & intEmpresa & ") AND (CodigoSituacao = 1) AND (Vendedor = 1)", "CodigoFuncionario", "Nome", strconexao, intEmpresa, intUsuario, uid, pwd
    
    ValorLimiteVendaConsignacao = PegaValorParametro("LIMITE_CONSIGNACAO_VENDA")
    
    
    lblConsig.Caption = IntTimConsigSaida & "-Consig"
    lblteste.Caption = IntTimTesteSaida & "-Teste"
    lblEmprestimo.Caption = IntTimEmprestimoSaida & "-Emprestimo"
    
    'Emanoel a pedido da Mitsu
    'Rotina Para Verificar se Usuario e um Vendedor Comum se for nao vai Alterar e nem pesquisar venda de outros vendedores
    booVendedor = False
    sSQL = "Select * from Funcionario where CodigoEmpresa = " & intEmpresa & " and CodigoFuncionario = " & intUsuario
    rs.Open sSQL, db, , , adCmdText
    If Not rs.EOF Then
       If rs!Vendedor = True And rs!Administrativo = False And rs!Gerente = False = rs!Supervisor = False Then
          Vendedor.CarregaGrafico "Funcionario", "SELECT CodigoFuncionario as Codigo, Nome as Descricao FROM Funcionario WHERE (CodigoEmpresa = " & intEmpresa & ") AND (CodigoSituacao = 1) AND (CodigoFuncionario = " & rs!CodigoFuncionario & ") AND (Vendedor = 1)", "CodigoFuncionario", "Nome", strconexao, intEmpresa, intUsuario, uid, pwd
          Vendedor.Text = rs!CodigoFuncionario
          Vendedor.Enabled = False
          booVendedor = True
          cmdDevolver.Enabled = False
          cmdFaturar.Enabled = False
       End If
    End If
    rs.Close
    
    sSQL = "Update CapaItemRelacionaCapaItem"
    sSQL = sSQL & " Set SequenciaProduto = CapaItem.SequenciaProduto"
    sSQL = sSQL & " FROM CapaItemRelacionaCapaItem INNER JOIN Capa ON CapaItemRelacionaCapaItem.CodigoEmpresa = Capa.CodigoEmpresa AND CapaItemRelacionaCapaItem.Sequencia = Capa.Sequencia INNER JOIN CapaItem ON CapaItemRelacionaCapaItem.CodigoEmpresa = CapaItem.CodigoEmpresa AND CapaItemRelacionaCapaItem.Sequencia = CapaItem.Sequencia AND CapaItemRelacionaCapaItem.CodigoProduto = CapaItem.CodigoProduto AND CapaItemRelacionaCapaItem.SequenciaProduto <> CapaItem.SequenciaProduto"
    sSQL = sSQL & " WHERE (CapaItemRelacionaCapaItem.CodigoEmpresa = " & intEmpresa & ") AND (Capa.Tipo = 'PV') AND (Capa.CodigoSituacao <> 99)"
    db.Execute sSQL
    
    booAtualizaProduto = True
    cmdPesquisar_Click
End Sub

Private Sub dtpFinal_GotFocus()
    SelecionaCampo Me, dtpFinal
End Sub

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub dtpInicial_GotFocus()
    SelecionaCampo Me, dtpInicial
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub GRID_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
On Error GoTo TrataErros

    If KeyAscii = 13 Then
        'Testar somente na coluna de edicao
        lblQtdDevolver.Caption = 0
        lblVlrDevolver.Caption = 0
        lblQtdFaturar.Caption = 0
        lblVlrFaturar.Caption = 0
        If optSintetico.Value = False Or optMostra.Value = False Then
            If optAnalitico.Value = True And optMostra.Value = False Then
                If rsGrid.RecordCount > 0 Then
                   Zera_Variaveis

                    rsGrid.MoveFirst
                    
                   Do While Not rsGrid.EOF
                   'Liberando o sistema
                   DoEvents:
                      'Testando Quantidade a Fatura e Devolver
                       If booContagem = False Then
                          If rsGrid("QtdeAfaturar") > 0 Or rsGrid("QtdeADevolver") > 0 Then
                             If (rsGrid("QtdeAfaturar") + rsGrid("QtdeADevolver")) > rsGrid("Qtddifer") Then
                                MsgBox "Quantidade a Faturar ou Devolver, Maior ou Igual a Quantidade Disponivel, Qtd Disponivel = " & rsGrid("Qtddifer"), vbCritical, "SAID"
                                db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                                db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                                cmdPesquisar_Click
                                Exit Sub
                             End If
                          End If
                       Else
                          If rsGrid("QtdeContagem") >= 0 Or rsGrid("QtdeADevolver") > 0 Then
                             If (rsGrid("QtdeContagem") + rsGrid("QtdeADevolver")) >= rsGrid("Qtddifer") Then
                                MsgBox "Quantidade a Faturar ou Devolver, Maior ou Igual a Quantidade Disponivel, Qtd Disponivel = " & rsGrid("Qtddifer"), vbCritical, "SAID"
                                db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                                db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                                cmdPesquisar_Click
                                Exit Sub
                             End If
                          End If
                       End If

                      If booContagem = False Then
                         If Not IsNull(rsGrid!QtdeAfaturar) Then
                            If rsGrid("QtdeAfaturar") > 0 Then
                               db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & rsGrid!QtdeAfaturar & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                            Else
                               If rsGrid!QtdeAfaturar = 0 Then 'so vai zerar se qtdeafaturar for = 0
                                  db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                               End If
                            End If
                            'Calculando Valor a Faturar
                            lblQtdFaturar.Caption = CCur(lblQtdFaturar.Caption) + rsGrid!QtdeAfaturar
                            If rsGrid!QtdeAfaturar > 0 Then
                               VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                               lblVlrFaturar.Caption = lblVlrFaturar.Caption + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
                            End If
                         End If
                      Else
                         If Not IsNull(rsGrid!qtdecontagem) Then
                            If rsGrid("QtdeContagem") >= 0 Then
                               db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & (rsGrid!qtddifer - rsGrid!qtdecontagem) & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                            Else
                               If rsGrid!QtdeAfaturar = 0 Then 'so vai zerar se qtdeafaturar for = 0
                                  db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                               End If
                            End If
                            lblQtdFaturar.Caption = CCur(lblQtdFaturar.Caption) + (rsGrid!qtddifer - rsGrid!qtdecontagem)
                            If rsGrid!qtdecontagem >= 0 Then 'Se for 0 e porque vai faturar tudo!
                               VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                               lblVlrFaturar.Caption = lblVlrFaturar.Caption + Arredonda((rsGrid!qtddifer - rsGrid!qtdecontagem) * VlrUnitario, 2)
                            End If
                         End If
                      End If
                      
                       'Pegando Valor IPI de Acordo com o Percentual do Produto
                       If booCalculaIPI = True Then
                          If rsGrid!AliquotaIPI > 0 Then
                             VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                             If booContagem = False Then
                                VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                             Else
                                If rsGrid!qtdecontagem = 0 Then
                                   VlrIPI = Arredonda(VlrIPI * (rsGrid!qtddifer - rsGrid!qtdecontagem), 2)
                                End If
                             End If
                             lblVlrFaturar.Caption = Arredonda(lblVlrFaturar.Caption + VlrIPI, 2)
                          End If
                      End If
                       If rsGrid("QtdeADevolver") > 0 Then
                          lblQtdDevolver.Caption = lblQtdDevolver.Caption + rsGrid!QtdeADevolver
                          db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & rsGrid("QtdeADevolver") & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                          VlrUnitarioDevolucao = Arredonda(rsGrid!VlSaida / rsGrid!qtdsaida, 2)
                          lblVlrDevolver.Caption = lblVlrDevolver.Caption + Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
                          If booCalculaIPI = True Then
                             If rsGrid!AliquotaIPI > 0 Then
                                VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                                VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                                lblVlrDevolver.Caption = Arredonda(lblVlrDevolver.Caption + VlrIPI, 2)
                             End If
                          End If
                          Else
                             db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                       End If
                       
                        'If rsGrid.RecordCount > 1 Then
                            rsGrid.MoveNext
                        'Else
                        '    GoTo Passa
                        'End If
                   Loop
'Passa:
                   cmdPesquisar_Click
                End If
            End If
        End If
    End If
    
    Exit Sub
TrataErros:
    If Err.Number <> 0 Then
        ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
        Exit Sub
    End If
    Resume
End Sub

Private Sub optConsignado_Click(Value As Integer)
    lblQtdConsigDesc.Caption = "Qtd. Consig.:"
    lblVlrConsigDesc.Caption = "Vlr. Consig.:"
    If booVendedor = False Then
        cmdDevolver.Enabled = True
        cmdFaturar.Enabled = True
    End If
    cmdPesquisar_Click
End Sub
    
Private Sub optEmprestimo_Click(Value As Integer)
    lblQtdConsigDesc.Caption = "Qtd. Empr.:"
    lblVlrConsigDesc.Caption = "Vlr. Empr.:"
    If booVendedor = False Then
       cmdDevolver.Enabled = True
       cmdFaturar.Enabled = True
    End If
    cmdPesquisar_Click
End Sub

Private Sub optDiferenca_Click(Value As Integer)
    txtPrazo.Text = 60
    txtPrazo.Enabled = False
    If booVendedor = False Then
       cmdDevolver.Enabled = True
       cmdFaturar.Enabled = True
       chkDevolve.Enabled = True
       chkFatura.Enabled = True
    End If
    cmdPesquisar_Click
End Sub

Private Sub optTeste_Click(Value As Integer)
    lblQtdConsigDesc.Caption = "Qtd. Teste.:"
    lblVlrConsigDesc.Caption = "Vlr. Teste.:"
    If booVendedor = False Then
       cmdDevolver.Enabled = True
       cmdFaturar.Enabled = True
    End If
    cmdPesquisar_Click
End Sub

Private Sub txtCliente_GotFocus()
    SelecionaCampo Me, txtCliente
End Sub
Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub txtCliente_LostFocus()
    On Error Resume Next
    
    If Len(txtCliente) = 0 Then Exit Sub
    
    If Len(txtCliente.Text) = 14 Then
        txtCliente.Text = Format(txtCliente.Text, "&&.&&&.&&&/&&&&-&&")
    ElseIf Len(txtCliente.Text) = 11 Then
        txtCliente.Text = Format(txtCliente.Text, "&&&.&&&.&&&-&&")
    End If
    
    sSQL = "SELECT Emitente.CodigoEmitente, Emitente.Nome, Emitente.CGC FROM Emitente Where Emitente.CodigoEmpresa=" & intEmpresa
    If Len(txtCliente) < 14 Then 'Usuário digitou o Código...
        sSQL = sSQL & " AND Emitente.CodigoEmitente=" & Val(txtCliente)
    Else  'Usuário digitou o CGC...
        sSQL = sSQL & " AND CGC='" & txtCliente.Text & "'"
    End If

    rs2.Open sSQL, db, , , adCmdText
    If Not rs2.EOF Then
        txtCliente.Text = rs2!CodigoEmitente
        txtCliente.Tag = rs2!CodigoEmitente
    Else
        MsgBox "CGC ou Código do Emitente não foi encontrado !", 48, "Sistema de Vendas"
        txtCliente = ""
        rs2.Close
        Exit Sub
    End If
    rs2.Close
End Sub

Private Sub txtPedido_GotFocus()
    SelecionaCampo Me, txtPedido
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = SomenteNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdPesquisar_Click
    End If
End Sub

Private Sub txtProduto_GotFocus()
   SelecionaCampo Me, txtProduto
End Sub
Private Sub txtProduto_KeyPress(KeyAscii As Integer)
On Error GoTo TrataErros
    If KeyAscii = 13 Then
        If txtProduto.Text <> "" Then
             sSQL = "SELECT  Produto.CodigoProduto, Produto.Descricao "
             sSQL = sSQL & " FROM Produto "
             sSQL = sSQL & " WHERE Produto.CodigoEmpresa=" & intEmpresa & " AND Produto.CodigoProduto=" & txtProduto.Text
             rs.Open sSQL, db, , , adCmdText
             If Not rs.EOF Then
                dtcProduto.Text = rs!Descricao
                Else 'Caso nao Exista
                    MsgBox "Atenção!!! Codigo de Produto Invalido", 48, Me.Caption
                    txtProduto.Text = ""
                    txtProduto.SetFocus
                    rs.Close
                    Exit Sub
             End If
             rs.Close
             Enter KeyAscii
             Else
                Enter KeyAscii
        End If
        Exit Sub
    End If
TrataErros:
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
    'Resume
End Sub

Private Sub txtSequencia_GotFocus()
    SelecionaCampo Me, txtSequencia
End Sub

Private Sub txtSequencia_KeyPress(KeyAscii As Integer)
    KeyAscii = SomenteNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdPesquisar_Click
    End If
End Sub

Private Sub txtNota_GotFocus()
    SelecionaCampo Me, txtNota
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
    KeyAscii = SomenteNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdPesquisar_Click
    End If
End Sub


Private Sub txtEmpresa_GotFocus()
    SelecionaCampo Me, txtEmpresa
End Sub

Private Sub txtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 80 Or KeyCode = 112 Then
        cmdEmpresa_Click
    End If
End Sub
Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
    KeyAscii = SomenteNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdPesquisar_Click
    End If
End Sub


Private Sub cmdEmpresa_Click()
    Screen.MousePointer = vbHourglass
    varaux(0) = "CodigoFilial"
    varaux(1) = "Descricao"
    varaux(2) = "SELECT Filial.CodigoFilial as Codigo, Filial.Descricao FROM Filial INNER JOIN AutorizacaoFilial ON Filial.CodigoEmpresa = AutorizacaoFilial.CodigoEmpresa AND Filial.CodigoFilial = AutorizacaoFilial.CodigoFilial WHERE AutorizacaoFilial.CodigoUsuario  = " & intUsuario & " and (Filial.CodigoEmpresa = " & intEmpresa & ") AND (Filial.CodigoSituacao <> 99)"
    varaux(3) = db
    varaux(4) = "Filial"
    varaux(5) = ""
    frmConsultaPadrao.Show 1
    If Len(strRetorno) > 0 Then
         txtEmpresa.Text = ZerosEsquerda(INTRETORNO, 4)
         txtEmpresa_KeyPress (13)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPesquisaCliente_Click()
    Screen.MousePointer = vbHourglass
    strRetorno = ""
    frmConsultaEmitente.Show 1
    If Len(strRetorno) > 0 Then
         txtCliente.Text = INTRETORNO
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub optSintetico_Click(Value As Integer)
    cmdDevolver.Enabled = False
    cmdFaturar.Enabled = False
    chkDevolve.Enabled = False
    chkFatura.Enabled = False
    
    cmdPesquisar_Click
End Sub

Private Sub optSintetico_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub optAnalitico_Click(Value As Integer)
    If booVendedor = False Then
       cmdDevolver.Enabled = True
       cmdFaturar.Enabled = True
       chkDevolve.Enabled = True
       chkFatura.Enabled = True
    End If
    cmdPesquisar_Click
End Sub

Private Sub optAnalitico_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub optMostra_Click(Value As Integer)
    txtPrazo.Text = 60
    txtPrazo.Enabled = False
    cmdDevolver.Enabled = False
    cmdFaturar.Enabled = False
    chkDevolve.Enabled = False
    chkFatura.Enabled = False
    cmdPesquisar_Click
End Sub

Private Sub optMostra_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub optPrazo_Click(Value As Integer)
    txtPrazo.Enabled = True
    If booVendedor = False Then
       cmdDevolver.Enabled = True
       cmdFaturar.Enabled = True
       chkDevolve.Enabled = True
       chkFatura.Enabled = True
    End If
    cmdPesquisar_Click
    'txtPrazo.SetFocus
End Sub

Private Sub optPrazo_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub cmdPesquisar_Click()
    On Error GoTo TrataErros
    Dim strSQL As String
    Dim Prazo As Integer
    Dim strTIN As String
    'Zerando Variaveis que acumula totais
    lblQtdConsig.Caption = 0
    lblQtdFaturar.Caption = 0
    lblQtdDevolver.Caption = 0
    lblQtdFaturada.Caption = 0
    lblQtdDevolvida.Caption = 0
    lblQtdDifer.Caption = 0
    
    lblVlrConsig.Caption = 0
    lblVlrFaturar.Caption = 0
    lblVlrDevolver.Caption = 0
    lblVlrFaturado.Caption = 0
    lblVlrDevolvido.Caption = 0
    lblVlrDiferenca.Caption = 0
    
    
    If Not IsDate(dtpInicial.Value) Then
        MsgBox "Informe a data inicial...", 48, Me.Caption
        dtpInicial.SetFocus
        Exit Sub
    End If

    If Not IsDate(dtpFinal.Value) Then
        MsgBox "Informe a data final...", 48, Me.Caption
        dtpFinal.SetFocus
        Exit Sub
    End If
    
    If optTeste.Value = True Then
        Msg "Pesquisando Teste..."
    Else
        Msg "Pesquisando Consignações..."
    End If
    
    Zera_Variaveis
    
    If optConsignado.Value = True Then
       TipoConsignacao = "C"
    ElseIf optTeste.Value = True Then
       TipoConsignacao = "T"
    ElseIf optEmprestimo.Value = True Then
       TipoConsignacao = "E"
    ElseIf optTodos.Value = True Then
       TipoConsignacao = "D"
    End If
    
'    If booAtualizaProduto = True Then
'        strTIN = PegaValorParametro("TIM_CONSIG_SAIDA")
'        strTIN = strTIN & "," & PegaValorParametro("TIM_TESTE_SAIDA")
'        strTIN = strTIN & "," & PegaValorParametro("TIM_EMPRESTIMO_SAIDA")
'
'        strRetorno = "SELECT Capa.CodigoEmitente FROM Capa INNER JOIN CapaItem ON Capa.CodigoEmpresa = CapaItem.CodigoEmpresa AND Capa.Sequencia = CapaItem.Sequencia WHERE Capa.CodigoEmpresa = " & intEmpresa & " AND Capa.CodigoFilial = " & intFilial & ""
'
'        If IsNumeric(txtCliente.Text) = True Then
'            strRetorno = strRetorno & " AND Capa.CodigoEmitente = " & txtCliente.Text & ""
'        End If
'
'        strRetorno = strRetorno & " AND Capa.CodigoTIM IN (" & strTIN & ") AND Capa.DataCadastro BETWEEN '" & mdaI(dtpInicial.Value) & "' and '" & mdaF(dtpFinal.Value) & "' AND Capa.CodigoSituacao <> 99 "
'
'        If IsNumeric(txtProduto.Text) = True Then
'            strRetorno = strRetorno & " AND CapaItem.CodigoProduto = " & txtProduto.Text & ""
'        End If
'
'        strRetorno = strRetorno & " GROUP BY Capa.CodigoEmitente"
'
'        rs.Open strRetorno, db, , , adCmdText
'        If Not rs.EOF Then
'            Do While Not rs.EOF
'                db.Execute "EXEC spProcessaConsignacao  1,0,'',''," & rs!CodigoEmitente & ",0,0,0,0," & IIf(IsNumeric(txtProduto.Text) = True, txtProduto.Text, 0) & ",'D',60,'A','C'"
'
'                rs.MoveNext
'            Loop
'        End If
'        rs.Close
'    End If

    sSQL = "EXEC spRelatorioConsignacaoTeste " & intEmpresa & "," & IIf(IsNumeric(txtEmpresa.Text) = True, txtEmpresa.Text, 0) & ",'" & mdaI(dtpInicial.Value) & "','" & mdaF(dtpFinal.Value) & "'," & IIf(IsNumeric(txtCliente.Text) = True, txtCliente.Text, 0) & "," & IIf(IsNumeric(txtSequencia.Text) = True, txtSequencia.Text, 0) & "," & IIf(IsNumeric(txtPedido.Text) = True, txtPedido.Text, 0) & "," & IIf(IsNumeric(txtNota.Text) = True, txtNota.Text, 0) & "," & IIf(IsNumeric(Vendedor.Text) = True, Vendedor.Text, 0) & "," & IIf(IsNumeric(txtProduto.Text) = True, txtProduto.Text, 0) & ",'" & IIf(optDiferenca.Value = True, "D", IIf(optPrazo.Value = True, "P", "T")) & "'," & IIf(IsNumeric(txtPrazo.Text) = True, txtPrazo.Text, 0) & ",'" & IIf(optSintetico.Value = True, "S", "A") & "','" & TipoConsignacao & "'"
    GridVaca.ValueLists.Clear
    rsGrid.Open sSQL, strConexaoGrid, adOpenKeyset, adLockOptimistic
    Set GridVaca.DataSource = rsGrid
    'Ordenando
    GridVaca.ViewStyleBand = ssViewStyleBandVertical
    GridVaca.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
    GridVaca.Override.FetchRows = ssFetchRowsPreloadWithParent
    GridVaca.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    GridVaca.TabNavigation = ssTabNavigationNextCell
    
    If txtPrazo.Text <= 60 Then
        GridVaca.Bands(0).Columns(0).Width = 750 'Codigo Produto
        GridVaca.Bands(0).Columns(0).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(0).LockedWidth = False
        GridVaca.Bands(0).Columns(0).Activation = ssActivationActivateNoEdit
        
        If optSintetico.Value = True Then
           GridVaca.Bands(0).Columns(1).Width = 5150 'Descricao Produto
        ElseIf optAnalitico.Value = True Then
           GridVaca.Bands(0).Columns(1).Width = 2000 'Descricao Produto
        End If
        
        GridVaca.Bands(0).Columns(1).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(1).LockedWidth = False
        GridVaca.Bands(0).Columns(1).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(2).Width = 300 'Embalagem
        GridVaca.Bands(0).Columns(2).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(2).LockedWidth = True
        GridVaca.Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
    
        GridVaca.Bands(0).Columns(3).Width = 4000 'Nome Cliente
        GridVaca.Bands(0).Columns(3).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(3).LockedWidth = False
        GridVaca.Bands(0).Columns(3).Activation = ssActivationActivateNoEdit
        
        If txtCliente.Text = "" Then
            If optSintetico.Value = True Then
                GridVaca.Bands(0).Columns(3).Hidden = True 'Deixa Desabilitado
            ElseIf optAnalitico.Value = True Then
                GridVaca.Bands(0).Columns(3).Hidden = False 'Deixa o Campo Habilitado
            End If
        Else
            GridVaca.Bands(0).Columns(3).Hidden = True 'Desabilita o campo
        End If
        
        GridVaca.Bands(0).Columns(4).Width = 700 'Numero Nota
        GridVaca.Bands(0).Columns(4).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(4).LockedWidth = False
        GridVaca.Bands(0).Columns(4).Activation = ssActivationActivateNoEdit
        
        If optSintetico.Value = True Or optMostra.Value = True Then
            GridVaca.Bands(0).Columns(4).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = True Then
            GridVaca.Bands(0).Columns(4).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = False Then
            GridVaca.Bands(0).Columns(4).Hidden = False 'Deixa o Campo Habilitado
        End If
    
        GridVaca.Bands(0).Columns(5).Width = 1000 'Data Consignacao
        GridVaca.Bands(0).Columns(5).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(5).LockedWidth = True
        GridVaca.Bands(0).Columns(5).Activation = ssActivationActivateNoEdit
        
        If optSintetico.Value = True Or optMostra.Value = True Then
            GridVaca.Bands(0).Columns(5).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = True Then
            GridVaca.Bands(0).Columns(5).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = False Then
            GridVaca.Bands(0).Columns(5).Hidden = False 'Deixa o Campo Habilitado
        End If
    
        GridVaca.Bands(0).Columns(6).Width = 750 'Qtd Atual Consignacao
        GridVaca.Bands(0).Columns(6).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(6).LockedWidth = True
        GridVaca.Bands(0).Columns(6).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(7).Width = 900 'Valor Atual Consignacao
        GridVaca.Bands(0).Columns(7).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(7).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(7).Format = "standard" 'Coloca Padrao com Duas casas Decimais
        GridVaca.Bands(0).Columns(7).Activation = ssActivationActivateNoEdit
        
        If booContagem = False Then
            GridVaca.Bands(0).Columns(8).Width = 1200 'Qtd A faturar
            GridVaca.Bands(0).Columns(8).LockedWidth = True
            
            If optSintetico.Value = True Or optMostra.Value = True Then
                GridVaca.Bands(0).Columns(8).CellAppearance.BackColor = &H8000000F
                GridVaca.Bands(0).Columns(8).Activation = ssActivationActivateNoEdit
            Else
                If booVendedor = False Then
                    GridVaca.Bands(0).Columns(8).Activation = ssActivationAllowEdit
                    GridVaca.UpdateMode = ssUpdateOnCellChange
                Else
                    GridVaca.Bands(0).Columns(8).CellAppearance.BackColor = &H8000000F
                    GridVaca.Bands(0).Columns(8).Activation = ssActivationActivateNoEdit
                End If
            End If
           
            GridVaca.Bands(0).Columns(9).Width = 700 'Qtd Contagem
            GridVaca.Bands(0).Columns(9).Hidden = True 'Desabilita o campo
        Else
            GridVaca.Bands(0).Columns(8).Width = 900 'Qtd a Faturar
            GridVaca.Bands(0).Columns(8).CellAppearance.BackColor = &H8000000F
            GridVaca.Bands(0).Columns(8).LockedWidth = True
            GridVaca.Bands(0).Columns(8).Activation = ssActivationActivateNoEdit
            
            GridVaca.Bands(0).Columns(9).Width = 1200 'Qtd Contagem
            GridVaca.Bands(0).Columns(9).LockedWidth = True
            
            If optSintetico.Value = True Or optMostra.Value = True Then
                GridVaca.Bands(0).Columns(9).CellAppearance.BackColor = &H8000000F
                GridVaca.Bands(0).Columns(9).Activation = ssActivationActivateNoEdit
            Else
                If booVendedor = False Then
                    GridVaca.Bands(0).Columns(9).Activation = ssActivationAllowEdit
                    GridVaca.UpdateMode = ssUpdateOnCellChange
                Else
                    GridVaca.Bands(0).Columns(9).CellAppearance.BackColor = &H8000000F
                    GridVaca.Bands(0).Columns(9).Activation = ssActivationActivateNoEdit
                End If
            End If
        End If
        
        GridVaca.Bands(0).Columns(10).Width = 1000 'Qtd Faturada Consignacao
        GridVaca.Bands(0).Columns(10).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(10).LockedWidth = True
        GridVaca.Bands(0).Columns(10).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(11).Width = 850 'Valor Faturada Consignacao
        GridVaca.Bands(0).Columns(11).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(11).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(11).Format = "standard"
        GridVaca.Bands(0).Columns(11).Activation = ssActivationActivateNoEdit
    
        GridVaca.Bands(0).Columns(12).Width = 1200 'Qtd a Devolver
        GridVaca.Bands(0).Columns(12).LockedWidth = True
        
        If optSintetico.Value = True Or optMostra.Value = True Then
            GridVaca.Bands(0).Columns(12).CellAppearance.BackColor = &H8000000F
            GridVaca.Bands(0).Columns(12).Activation = ssActivationActivateNoEdit
        Else
            If booVendedor = False Then
                GridVaca.Bands(0).Columns(12).Activation = ssActivationAllowEdit
                GridVaca.UpdateMode = ssUpdateOnCellChange
            Else
                GridVaca.Bands(0).Columns(12).CellAppearance.BackColor = &H8000000F
                GridVaca.Bands(0).Columns(12).Activation = ssActivationActivateNoEdit
            End If
        End If
        
        GridVaca.Bands(0).Columns(13).Width = 1100 'Qtd Devolucao
        GridVaca.Bands(0).Columns(13).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(13).LockedWidth = True
        GridVaca.Bands(0).Columns(13).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(14).Width = 1000 'Valor Devolucao
        GridVaca.Bands(0).Columns(14).CellAppearance.BackColor = &H8000000F
        GridVaca.Bands(0).Columns(14).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(14).Format = "standard"
        GridVaca.Bands(0).Columns(14).Activation = ssActivationActivateNoEdit
                
        GridVaca.Bands(0).Columns(15).Width = 700 'Qtd Diferença
        GridVaca.Bands(0).Columns(15).CellAppearance.BackColor = &H80C0FF
        GridVaca.Bands(0).Columns(15).LockedWidth = True
        GridVaca.Bands(0).Columns(15).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(16).Width = 850 'Valor Diferença
        GridVaca.Bands(0).Columns(16).CellAppearance.BackColor = &H80C0FF
        GridVaca.Bands(0).Columns(16).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(16).Format = "standard"
        GridVaca.Bands(0).Columns(16).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(17).Width = 700 'Valor Diferença
        GridVaca.Bands(0).Columns(17).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(18).Width = 700 'Sequencia
        GridVaca.Bands(0).Columns(18).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(19).Width = 700 'Valor IPI
        GridVaca.Bands(0).Columns(19).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(20).Width = 700 'Valor ICMS
        GridVaca.Bands(0).Columns(20).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(21).Width = 700 'Codigo Deposito
        GridVaca.Bands(0).Columns(21).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(22).Width = 700 'Codigo Cliente
        GridVaca.Bands(0).Columns(22).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(23).Width = 700 'Sequencia Produto Original
        GridVaca.Bands(0).Columns(23).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(24).Width = 700 'Codigo Unidade
        GridVaca.Bands(0).Columns(24).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(25).Width = 700 'Valor Liquido Unitario Sem IPI
        GridVaca.Bands(0).Columns(25).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(26).Width = 700 'Valor Liquido Unitario Com IPI
        GridVaca.Bands(0).Columns(26).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(27).Width = 700 'Valor Desconto
        GridVaca.Bands(0).Columns(27).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(28).Width = 700 'Valor Desconto
        GridVaca.Bands(0).Columns(28).Hidden = True 'Desabilita o campo
        
        If optTodos.Value = True Then
            GridVaca.Bands(0).Columns(29).Width = 1200 'Tipo
            GridVaca.Bands(0).Columns(29).CellAppearance.BackColor = &H80C0FF
            GridVaca.Bands(0).Columns(29).LockedWidth = True
            GridVaca.Bands(0).Columns(29).Activation = ssActivationActivateNoEdit
        Else
            GridVaca.Bands(0).Columns(29).Width = 700
            GridVaca.Bands(0).Columns(29).Hidden = True
        End If
    Else 'se for maior que 60 dias fica vermelho
        GridVaca.Bands(0).Columns(0).Width = 750 'Codigo Produto
        GridVaca.Bands(0).Columns(0).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(0).LockedWidth = False
        GridVaca.Bands(0).Columns(0).Activation = ssActivationActivateNoEdit
                
        If optSintetico.Value = True Then
           GridVaca.Bands(0).Columns(1).Width = 5150 'Descricao Produto
        ElseIf optAnalitico.Value = True Then
           GridVaca.Bands(0).Columns(1).Width = 2000 'Descricao Produto
        End If
        
        GridVaca.Bands(0).Columns(1).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(1).LockedWidth = False
        GridVaca.Bands(0).Columns(1).Activation = ssActivationActivateNoEdit
    
        GridVaca.Bands(0).Columns(2).Width = 300 'Embalagem
        GridVaca.Bands(0).Columns(2).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(2).LockedWidth = True
        GridVaca.Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
    
        GridVaca.Bands(0).Columns(3).Width = 1450 'Nome Cliente
        GridVaca.Bands(0).Columns(3).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(3).LockedWidth = False
        GridVaca.Bands(0).Columns(3).Activation = ssActivationActivateNoEdit
        
        If txtCliente.Text = "" Then
            If optSintetico.Value = True Then
                GridVaca.Bands(0).Columns(3).Hidden = True 'Deixa Desabilitado
            ElseIf optAnalitico.Value = True Then
                GridVaca.Bands(0).Columns(3).Hidden = False 'Deixa o Campo Habilitado
            End If
        Else
            GridVaca.Bands(0).Columns(3).Hidden = True 'Desabilita o campo
        End If
    
        GridVaca.Bands(0).Columns(4).Width = 700 'Numero Nota
        GridVaca.Bands(0).Columns(4).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(4).LockedWidth = False
        GridVaca.Bands(0).Columns(4).Activation = ssActivationActivateNoEdit
        
        If optSintetico.Value = True Or optMostra.Value = True Then
            GridVaca.Bands(0).Columns(4).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = True Then
            GridVaca.Bands(0).Columns(4).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = False Then
            GridVaca.Bands(0).Columns(4).Hidden = False 'Deixa o Campo Habilitado
        End If
        
        GridVaca.Bands(0).Columns(5).Width = 1000 'Data Consignacao
        GridVaca.Bands(0).Columns(5).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(5).LockedWidth = True
        GridVaca.Bands(0).Columns(5).Activation = ssActivationActivateNoEdit
        
        If optSintetico.Value = True Or optMostra.Value = True Then
            GridVaca.Bands(0).Columns(5).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = True Then
            GridVaca.Bands(0).Columns(5).Hidden = True 'Deixa Desabilitado
        ElseIf optAnalitico.Value = True And optMostra.Value = False Then
            GridVaca.Bands(0).Columns(5).Hidden = False 'Deixa o Campo Habilitado
        End If
        
        GridVaca.Bands(0).Columns(6).Width = 750 'Qtd Atual Consignacao
        GridVaca.Bands(0).Columns(6).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(6).LockedWidth = True
        GridVaca.Bands(0).Columns(6).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(7).Width = 900 'Valor Atual Consignacao
        GridVaca.Bands(0).Columns(7).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(7).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(7).Format = "standard"
        GridVaca.Bands(0).Columns(7).Activation = ssActivationActivateNoEdit
        
        If booContagem = False Then
            GridVaca.Bands(0).Columns(8).Width = 1200 'Qtd A faturar
            GridVaca.Bands(0).Columns(8).LockedWidth = True
            
            If optSintetico.Value = True Or optMostra.Value = True Then
                GridVaca.Bands(0).Columns(8).CellAppearance.BackColor = &HFF&
                GridVaca.Bands(0).Columns(8).Activation = ssActivationActivateNoEdit
            Else
                If booVendedor = False Then
                    GridVaca.Bands(0).Columns(8).Activation = ssActivationAllowEdit
                    GridVaca.UpdateMode = ssUpdateOnCellChange
                Else
                    GridVaca.Bands(0).Columns(8).CellAppearance.BackColor = &HFF&
                    GridVaca.Bands(0).Columns(8).Activation = ssActivationActivateNoEdit
                End If
            End If
            
            GridVaca.Bands(0).Columns(9).Width = 700 'Qtd Contagem
            GridVaca.Bands(0).Columns(9).Hidden = True 'Desabilita o campo
        Else
           GridVaca.Bands(0).Columns(8).Width = 900 'Qtd a Faturar
           GridVaca.Bands(0).Columns(8).CellAppearance.BackColor = &HFF&
           GridVaca.Bands(0).Columns(8).LockedWidth = True
           GridVaca.Bands(0).Columns(8).Activation = ssActivationActivateNoEdit
           
           GridVaca.Bands(0).Columns(9).Width = 1200 'Qtd Contagem
           GridVaca.Bands(0).Columns(9).LockedWidth = True
           
            If optSintetico.Value = True Or optMostra.Value = True Then
                GridVaca.Bands(0).Columns(9).CellAppearance.BackColor = &HFF&
                GridVaca.Bands(0).Columns(9).Activation = ssActivationActivateNoEdit
            Else
                If booVendedor = False Then
                    GridVaca.Bands(0).Columns(9).Activation = ssActivationAllowEdit
                    GridVaca.UpdateMode = ssUpdateOnCellChange
                Else
                    GridVaca.Bands(0).Columns(9).CellAppearance.BackColor = &HFF&
                    GridVaca.Bands(0).Columns(9).Activation = ssActivationActivateNoEdit
                End If
            End If
        End If
        
        GridVaca.Bands(0).Columns(10).Width = 1000 'Qtd Faturada Consignacao
        GridVaca.Bands(0).Columns(10).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(10).LockedWidth = True
        GridVaca.Bands(0).Columns(10).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(11).Width = 850 'Valor Faturada Consignacao
        GridVaca.Bands(0).Columns(11).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(11).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(11).Format = "standard"
        GridVaca.Bands(0).Columns(11).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(12).Width = 1000 'Qtd a Devolver
        GridVaca.Bands(0).Columns(12).LockedWidth = True
        
        If optSintetico.Value = True Or optMostra.Value = True Then
            GridVaca.Bands(0).Columns(12).CellAppearance.BackColor = &HFF&
            GridVaca.Bands(0).Columns(12).Activation = ssActivationActivateNoEdit
        Else
            If booVendedor = False Then
                GridVaca.Bands(0).Columns(12).Activation = ssActivationAllowEdit
                GridVaca.UpdateMode = ssUpdateOnCellChange
            Else
                GridVaca.Bands(0).Columns(12).Activation = ssActivationActivateNoEdit
            End If
        End If
        
        GridVaca.Bands(0).Columns(13).Width = 1100 'Qtd Devolucao
        GridVaca.Bands(0).Columns(13).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(13).LockedWidth = True
        GridVaca.Bands(0).Columns(13).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(14).Width = 1000 'Valor Devolucao
        GridVaca.Bands(0).Columns(14).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(14).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(14).Format = "standard"
        GridVaca.Bands(0).Columns(14).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(15).Width = 700 'Qtd Diferença
        GridVaca.Bands(0).Columns(15).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(15).LockedWidth = True
        GridVaca.Bands(0).Columns(15).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(16).Width = 850 'Valor Diferença
        GridVaca.Bands(0).Columns(16).CellAppearance.BackColor = &HFF&
        GridVaca.Bands(0).Columns(16).LockedWidth = True
        GridVaca.Layout.Bands(0).Columns(16).Format = "standard"
        GridVaca.Bands(0).Columns(16).Activation = ssActivationActivateNoEdit
        
        GridVaca.Bands(0).Columns(17).Width = 700 'Filial
        GridVaca.Bands(0).Columns(17).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(18).Width = 700 'Sequencia
        GridVaca.Bands(0).Columns(18).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(19).Width = 700 'Valor IPI
        GridVaca.Bands(0).Columns(19).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(20).Width = 700 'Valor ICMS
        GridVaca.Bands(0).Columns(20).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(21).Width = 700 'Codigo Deposito
        GridVaca.Bands(0).Columns(21).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(22).Width = 700 'Codigo Cliente
        GridVaca.Bands(0).Columns(22).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(23).Width = 700 'Sequencia Produto Original
        GridVaca.Bands(0).Columns(23).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(24).Width = 700 'Codigo Unidade
        GridVaca.Bands(0).Columns(24).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(25).Width = 700 'Valor Liquido Unitario sem IPI
        GridVaca.Bands(0).Columns(25).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(26).Width = 700 'Valor Liquido Unitario Com IPI
        GridVaca.Bands(0).Columns(26).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(27).Width = 700 'Valor Desconto
        GridVaca.Bands(0).Columns(27).Hidden = True 'Desabilita o campo
        
        GridVaca.Bands(0).Columns(28).Width = 700 'Valor Desconto
        GridVaca.Bands(0).Columns(28).Hidden = True 'Desabilita o campo
        
        If optTodos.Value = True Then
            GridVaca.Bands(0).Columns(29).Width = 1200 'Tipo
            GridVaca.Bands(0).Columns(29).CellAppearance.BackColor = &HFF&
            GridVaca.Bands(0).Columns(29).LockedWidth = True
            GridVaca.Bands(0).Columns(29).Activation = ssActivationActivateNoEdit
        Else
            GridVaca.Bands(0).Columns(29).Width = 700
            GridVaca.Bands(0).Columns(29).Hidden = True
        End If
    End If
    
    If Not rsGrid.EOF Then
       While Not rsGrid.EOF
             'Totalizacao de Valores
                lblVlrDevolvido.Caption = Arredonda(lblVlrDevolvido.Caption + rsGrid!VlDevolvido, 2)
                lblVlrFaturado.Caption = Arredonda(lblVlrFaturado.Caption + rsGrid!VlFaturado, 2)
                lblVlrConsig.Caption = Arredonda(lblVlrConsig.Caption + rsGrid!VlSaida, 2)
                lblVlrDiferenca.Caption = Arredonda(lblVlrDiferenca.Caption + rsGrid!vldifer, 2)
             
                lblQtdConsig.Caption = CCur(lblQtdConsig.Caption) + rsGrid!qtdsaida
                lblQtdFaturar.Caption = CCur(lblQtdFaturar.Caption) + rsGrid!QtdeAfaturar
                
                'Calculando Valor a Faturar
                If rsGrid!QtdeAfaturar > 0 Then
                    VlrUnitario = Arredonda(rsGrid!VlSaida / rsGrid!qtdsaida, 2)
                    lblVlrFaturar.Caption = lblVlrFaturar.Caption + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
                End If
                
                'Pegando Valor IPI de Acordo com o Percentual do Produto
                If rsGrid!QtdeAfaturar > 0 Then
                   If rsGrid!AliquotaIPI > 0 Then
                         VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                         VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                         lblVlrFaturar.Caption = Arredonda(lblVlrFaturar.Caption + VlrIPI, 2)
                   End If
                End If
                
                lblQtdFaturada.Caption = CCur(lblQtdFaturada.Caption) + rsGrid!QtdFaturado
                lblQtdDevolver.Caption = CCur(lblQtdDevolver.Caption) + rsGrid!QtdeADevolver
                
                'Calculando valores de ICMS e IPI para Devolucao
                If rsGrid!QtdeADevolver > 0 Then
                    VlrUnitarioDevolucao = Arredonda(rsGrid!VlSaida / rsGrid!qtdsaida, 2)
                    lblVlrDevolver.Caption = lblVlrDevolver.Caption + Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
                    
                    If rsGrid!AliquotaIPI > 0 Then
                        VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                        VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                        VlrIPITotal = Arredonda(VlrIPITotal + VlrIPI, 2)
                        lblVlrDevolver.Caption = Arredonda(lblVlrDevolver.Caption + VlrIPI, 2)
                    End If
                    
                    If rsGrid!AliquotaICMS > 0 Then
                        VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                        VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                        VlrICMSTotal = Arredonda(VlrICMSTotal + VlrICMS, 2)
                        BaseICMS = lblVlrDevolver.Caption
                    End If
                End If
                
                lblQtdDevolvida.Caption = CCur(lblQtdDevolvida.Caption) + rsGrid!QtdDevolvida
                lblQtdDifer.Caption = CCur(lblQtdDifer.Caption) + rsGrid!qtddifer
             rsGrid.MoveNext
        Wend
        rsGrid.MoveFirst
    End If
    
    If optTeste.Value = True Then
          Grid.Caption = "Pesquisa Testes: " & rsGrid.RecordCount
    ElseIf optEmprestimo.Value = True Then
          Grid.Caption = "Pesquisa Emprestimo: " & rsGrid.RecordCount
    ElseIf optConsignado.Value = True Then
          Grid.Caption = "Pesquisa Consignação: " & rsGrid.RecordCount
    ElseIf optTodos.Value = True Then
          Grid.Caption = "Pesquisa Todos: " & rsGrid.RecordCount
    End If
    
'    If CCur(lblVlrConsig.Caption) > 0 And CCur(lblQtdConsig.Caption) > 0 Then
'        lblVlrConsig.tag = CCur(lblVlrConsig.Caption) / CCur(lblQtdConsig.Caption)
'        lblVlrFaturar.Caption = Moeda(CCur(lblQtdFaturar.Caption) * CCur(lblVlrConsig.tag))
'        lblVlrFaturado.Caption = Moeda(CCur(lblQtdFaturada.Caption) * CCur(lblVlrConsig.tag))
'        lblVlrDevolver.Caption = Moeda(CCur(lblQtdDevolver.Caption) * CCur(lblVlrConsig.tag))
'        lblVlrDevolvido.Caption = Moeda(CCur(lblQtdDevolvida.Caption) * CCur(lblVlrConsig.tag))
'
'        lblQtdDifer.Caption = CCur(lblQtdFaturar.Caption) - CCur(lblQtdFaturada.Caption) - CCur(lblQtdDevolvida.Caption)
'    End If
    
fim:
    
    Msg ""
    Screen.MousePointer = 0
    Exit Sub

TrataErros:
    
    If Err.Number = 3705 Then 'objeto aberto
       rsGrid.Close
       Resume
    ElseIf Err.Number = 3704 Then 'objeto fechado
        Resume Next
    End If
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption

End Sub

Function Gravar() As Boolean
On Error GoTo TrataErros
    Dim varsql As String
    
    'TotalizaGeral
    Gravar = False
    'Se o Usuario nao der Enter ele faz o Calculo da Quantidade a Faturar aqui!
    If booContagem = True Then
        If Not rsGrid.EOF Then
            rsGrid.MoveFirst
            While Not rsGrid.EOF
                If (rsGrid("QtdeContagem") + rsGrid("QtdeADevolver")) >= rsGrid("Qtddifer") Then
                    MsgBox "Quantidade a Faturar ou Devolver, Maior ou Igual a Quantidade Disponivel, Qtd Disponivel = " & rsGrid("Qtddifer"), vbCritical, "SAID"
                    Gravar = False
                    Exit Function
                Else
                    If rsGrid!qtdecontagem >= 0 Then
                        db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & (rsGrid!qtddifer - rsGrid!qtdecontagem) & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                    End If
                End If
                rsGrid.MoveNext
            Wend
            rsGrid.MoveFirst
            cmdPesquisar_Click
        End If
    End If
    
    If Not rsGrid.EOF Then
        rsGrid.MoveFirst
        While Not rsGrid.EOF
           lblQtdFaturar.Caption = CCur(lblQtdFaturar.Caption) + rsGrid!QtdeAfaturar
           'Calculando Valor a Faturar
           If rsGrid!QtdeAfaturar > 0 Then
              VlrUnitario = Arredonda(rsGrid!VlSaida / rsGrid!qtdsaida, 2)
              lblVlrFaturar.Caption = lblVlrFaturar.Caption + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
           End If
           'Pegando Valor IPI de Acordo com o Percentual do Produto
           If rsGrid!QtdeAfaturar > 0 Then
              If rsGrid!AliquotaIPI > 0 Then
                    VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                    VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                    lblVlrFaturar.Caption = Arredonda(lblVlrFaturar.Caption + VlrIPI, 2)
              End If
           End If
           rsGrid.MoveNext
        Wend
        rsGrid.MoveFirst
    End If
    
    If Faturar = True Then
       If ValorLimiteVendaConsignacao > lblVlrFaturar.Tag Then
          MsgBox "Valor do Faturamento da Consignação Inferior ao Limite Estabelecido, Valor Limite Venda.: R$ " & ValorLimiteVendaConsignacao & ",00" & " Valor a Faturar.: R$ " & lblVlrFaturar.Tag & " !", 48, "Sistema de Vendas"
          Gravar = False
          Exit Function
       End If
    End If
    
    db.BeginTrans 'Inicializando Tranzacao
    intSequenciaProduto = 0
    If Faturar = True Then
        
        'Pegando o Deposito e o Tipo de Estoque da Tim
        sSQL = "SELECT CodigoDeposito, Estoque "
        sSQL = sSQL & "From TIM "
        sSQL = sSQL & "Where Codigo = " & IntTimConsigVenda
        rs.Open sSQL, db, , , adCmdText
        If Not rs.EOF Then
            If strRetorno = "" Then
                CodigoDepositoTIM = rs!CodigoDeposito
            Else
                CodigoDepositoTIM = Str(Left(strRetorno, 1))
            End If
           
            If rs!Estoque = 0 Then
                EstoqueTIM = "N" 'Nenhum
            ElseIf rs!Estoque = 1 Then
                EstoqueTIM = "S" 'Saida
            ElseIf rs!Estoque = 2 Then
                EstoqueTIM = "E" 'Entrada
            ElseIf rs!Estoque = 3 Then
                EstoqueTIM = "R" 'Reserva
            End If
        End If
        rs.Close
        
        ValorIPIItem = 0
        ValorICMSITEM = 0
        intSequencia = 0
        intSequenciaItem = 0
        
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            DoEvents:
            If rsGrid("QtdeAfaturar") > 0 Then
               varsql = "SELECT  Capa.CodigoEmitente, Capa.Quantidade, Capa.ValorTotal, Capa.ValorLiquido "
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE Capa.CodigoEmpresa=" & intEmpresa & " AND Capa.CodigoEmitente = " & rsGrid("CodigoCliente") & " AND Capa.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                    'Acumulando Valores Caso seja mesmo Cliente
                    qtdfaturar = qtdfaturar + rsGrid!QtdeAfaturar
                    VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                    VlrFaturar = VlrFaturar + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)

                    'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                       End If
                    End If

                    VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
                    lblvlrfaturar1.Caption = Replace(VlrFaturar, ",", ".")
                    lblvlrfaturarIPI.Caption = Replace(VlrFaturarIPI, ",", ".")

                   'Atualizando Cabeca do pedido
                   db.Execute "UPDATE Capa SET Quantidade = " & qtdfaturar & ", ValorTotal = " & lblvlrfaturarIPI.Caption & ", ValorLiquido = " & lblvlrfaturar1.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo PulaCliente
               End If
               rs.Close
               
               'Vai appendando numeros de Nota da Consignacao para Controle da Empresa
               
               'Calculando Valor a Faturar
               qtdfaturar = rsGrid!QtdeAfaturar
               VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
               VlrFaturar = Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
               
               'Pegando Valor IPI de Acordo com o Percentual do Produto
               If booCalculaIPI = True Then
                  If rsGrid!AliquotaIPI > 0 Then
                     VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                     VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                     VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
                  End If
               End If
               
               VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
               lblvlrfaturar1.Caption = Replace(VlrFaturar, ",", ".")
               lblvlrfaturarIPI.Caption = Replace(VlrFaturarIPI, ",", ".")
               
               intSequencia = BuscaCodigo("SELECT isnull(max(Sequencia),0) + 1 as Total FROM Capa WHERE CodigoEmpresa = " & intEmpresa, db, "Total")
               'Copiando Remessa Consignacao mudando algumas particularidades como sequencia e valores e tim devolucao
               varsql = "INSERT INTO Capa (CodigoEmpresa, Sequencia, CodigoFilial, Tipo, CodigoSituacao, CodigoVendedor, CodigoTabelaPreco, CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, Requisicao, Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia, Quantidade, ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca, ValorLiquido, ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor,  ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, TipoFrete, BaseICMS, ValorICMS, BaseIPI, ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, "
               varsql = varsql & " DataCadastro, DataFaturamento, DataEntradaSaida, DataContabilizacao, DataMovimentacao, CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, CIFFOB, CodigoCentroCusto, CodigoNatureza, CodigoNaturezaOperacao, Confirmada, CodigoMotivo, AtualizaFinanceiro, AtualizaEstoque, CodigoRequisitante, CodigoAutorizacao, CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, DataCancelamento, CodigoUsuarioCancelamento, DataAlteracao, CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, EmiteNF, EmiteCupom, Acertado, CodigoContaCorrente, "
               varsql = varsql & " TipoVenda, TipoComissao, PercentualComissao, Placa, UFPlaca, ParametroFaturamento) "
               varsql = varsql & " SELECT CodigoEmpresa, " & intSequencia & " AS Sequencia, CodigoFilial, '" & strTipoMovimento & "' as Tipo, 1 as CodigoSituacao, CodigoVendedor, CodigoTabelaPreco," & IntTimConsigVenda & " AS CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, " & intSequencia & " as Requisicao, " & 0 & " AS Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia," & rsGrid("QtdeAFaturar") & " as  Quantidade," & lblvlrfaturarIPI.Caption & " as ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, " & 0 & " as ValorDescontoItem, ValorBonificacao, ValorTroca," & lblvlrfaturar1.Caption & " as ValorLiquido," & 0 & " as ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor, ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, "
               varsql = varsql & " TipoFrete, " & 0 & " as  BaseICMS, " & 0 & " as  ValorICMS, " & 0 & " as  BaseIPI, " & 0 & " as  ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, getdate() as DataCadastro, '" & strDataBase & "' as DataFaturamento, '" & strDataBase & "' as DataEntradaSaida, '" & strDataBase & "' as DataContabilizacao, '" & strDataBase & "' as DataMovimentacao, " & intUsuario & " as CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, '" & "C" & "' as CIFFOB, CodigoCentroCusto, CodigoNatureza, " & 0 & " as CodigoNaturezaOperacao, Confirmada, CodigoMotivo, '" & "N" & "' AS AtualizaFinanceiro, '" & EstoqueTIM & "' AS AtualizaEstoque, "
               varsql = varsql & " CodigoRequisitante, CodigoAutorizacao, " & 0 & " as CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, null as DataCancelamento, null as CodigoUsuarioCancelamento, null as DataAlteracao, null as CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, 0 as EmiteNF, 0 as EmiteCupom, Acertado, CodigoContaCorrente, TipoVenda, TipoComissao, PercentualComissao, '" & 0 & "' as Placa, UFPlaca , '" & 1 & "' as ParametroFaturamento"
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & rsGrid("Sequencia")
               db.Execute varsql
               
               'Verificando se tem alguma coisa e deleta antes de iniciar processo
               db.Execute "DELETE FROM CapaItem WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
               
    'Rotina Caso O Cliente Seja O mesmo
PulaCliente:
               
               If IntNotaAnterior <> rsGrid!Nota Then
                    'Lendo tabela Capaitem Para pegar Valores IPI e ICMS para jogar na Msgs de Faturamento
                    varsql = "SELECT  CapaItem.ValorIPI, CapaItem.ValorICMS, CapaItem.Quantidade "
                    varsql = varsql & " FROM CapaItem "
                    varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid!Sequencia & " and SequenciaProduto = " & rsGrid!SequenciaOriginal
                    rs.Open varsql, db, , , adCmdText
                    If Not rs.EOF Then
                        If rs!ValorIPI > 0 Then
                            VlrIPI = Arredondar((rs!ValorIPI / rs!Quantidade) * rsGrid("QtdeAFaturar"), 2)
                        Else
                            VlrIPI = 0
                        End If
                       
                        If rs!ValorICMS > 0 Then
                            VlrICMS = Arredondar((rs!ValorICMS / rs!Quantidade) * rsGrid("QtdeAFaturar"), 2)
                        Else
                            VlrICMS = 0
                        End If
                    End If
                    rs.Close
                    
                    If SequenciaAnterior > 0 Then
                        If intSequencia = SequenciaAnterior Then
                            If strNumeroNota <> "" Then
                                If Len(strNumeroNota) <= 500 Then
                                    strNumeroNota = strNumeroNota & "," & rsGrid!Nota
                                    ValorIPIItem = ValorIPIItem + VlrIPI
                                    ValorICMSITEM = ValorICMSITEM + VlrICMS
                                End If
                            Else
                                strNumeroNota = rsGrid!Nota
                                ValorIPIItem = ValorIPIItem + VlrIPI
                                ValorICMSITEM = ValorICMSITEM + VlrICMS
                            End If
                        End If
                    Else
                        strNumeroNota = rsGrid!Nota
                        ValorIPIItem = ValorIPIItem + VlrIPI
                        ValorICMSITEM = ValorICMSITEM + VlrICMS
                    End If
               End If
               
               If SequenciaAnterior > 0 Then
                    If intSequencia <> SequenciaAnterior Then
                       Dim spaco As String
                       strMsgNota = "Simples faturamento de merc. em consig. mercantil ref. NFs: "
                       strMsgNota = strMsgNota & strNumeroNota
                       strMsgNotaImposto = ValorIPIItem & " e ICMS R$ " & ValorICMSITEM
                       strMsgNota = strMsgNota & " Os impostos foram pagos através das citadas NFs: IPI R$ " & strMsgNotaImposto
                       
                       db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & SequenciaAnterior
                       strNumeroNota = rsGrid!Nota
                       ValorIPIItem = VlrIPI
                       ValorICMSITEM = VlrICMS
                    End If
               End If
               
               'Relacionado Capa ao caparelacinaCapa
               varsql = "SELECT  CapaRelacionaCapa.Sequencia, CapaRelacionaCapa.Sequencia "
               varsql = varsql & " FROM CapaRelacionaCapa "
               varsql = varsql & " WHERE CapaRelacionaCapa.CodigoEmpresa=" & intEmpresa & " AND CapaRelacionaCapa.sequencia = " & intSequencia & " AND CapaRelacionaCapa.SequenciaOriginal = " & rsGrid("Sequencia")
               rs.Open varsql, db, , , adCmdText
               If rs.EOF Then
                  db.Execute "INSERT INTO CapaRelacionaCapa (CodigoEmpresa, Sequencia, SequenciaOriginal) VALUES (" & intEmpresa & "," & intSequencia & "," & rsGrid("Sequencia") & ")"
               End If
               rs.Close
               
               varsql = "SELECT  Capaitem.CodigoProduto, Capaitem.Quantidade, Capaitem.ValorUnitario, Capaitem.ValorLiquido "
               varsql = varsql & " FROM Capaitem "
               varsql = varsql & " WHERE Capaitem.CodigoEmpresa=" & intEmpresa & " AND Capaitem.CodigoProduto = " & rsGrid("Codigo") & " AND Capaitem.Sequencia = " & intSequencia & " and CapaItem.SequenciaProduto = " & intSequenciaItem
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                  'Acumulando Valores Caso seja o Mesmo Produto mesmo Cliente
                   QtdFaturaItem = rs!Quantidade + rsGrid("QtdeAFaturar")

                   'Pegando Valor unitario do que vem para atualizar
                   lblvlrUniItem.Caption = (rsGrid("VlSaida") / rsGrid("qtdSaida"))

                   'Calculando o valor unitario com a quantidade solicitada
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption * rsGrid("qtdeAfaturar")

                   'Calculando total Quantidade anterior com quantidade atual
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption + (rs!Quantidade * rs!ValorUnitario)

                   'Por Fim pegando o valor Unitario pela divisao das quantidades e valores
                   lblvlrUniItem.Caption = (lblvlrUniItem.Caption / QtdFaturaItem)
                   lblvlrLiqItem.Caption = lblvlrUniItem.Caption * QtdFaturaItem

                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If rsGrid!AliquotaIPI > 0 Then
                       VlrIPI = Arredonda(((lblvlrUniItem.Caption * rsGrid!AliquotaIPI) / 100), 2)
                       VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                    End If

                    lblvlrTotItem.Caption = Arredonda(lblvlrLiqItem.Caption + VlrIPI, 2)

                    lblvlrUniItem.Caption = Replace(lblvlrUniItem.Caption, ",", ".")
                    lblvlrLiqItem.Caption = Replace(lblvlrLiqItem.Caption, ",", ".")
                    lblvlrTotItem.Caption = Replace(lblvlrTotItem.Caption, ",", ".")
                   'Atualizando Itens  do pedido
                   db.Execute "UPDATE CapaItem SET Quantidade = " & QtdFaturaItem & ", ValorUnitario = " & lblvlrUniItem.Caption & ", ValorLiquido = " & lblvlrLiqItem.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & intSequencia & " and SequenciaProduto = " & intSequenciaItem
                   rs.Close
                   GoTo pula
               End If
               rs.Close
                           
                'intSequenciaProduto = intSequenciaProduto + 1
                intSequenciaItem = intSequenciaItem + 1
                sSQL = " DECLARE @CodigoEmpresa tinyint"
                sSQL = sSQL & " DECLARE @Sequencia int"
                sSQL = sSQL & " DECLARE @SequenciaProduto tinyint"
                sSQL = sSQL & " DECLARE @CodigoProduto int"
                sSQL = sSQL & " DECLARE @Quantidade money"
                sSQL = sSQL & " DECLARE @Desconto decimal(19,4)"
                sSQL = sSQL & " DECLARE @ValorUnitario decimal(19,4)"
                sSQL = sSQL & " DECLARE @CodigoVendedor smallint"
                sSQL = sSQL & " DECLARE @AliquotaIPI money"
                sSQL = sSQL & " DECLARE @AliquotaICMS money"
                sSQL = sSQL & " DECLARE @AliquotaReducao money"
                sSQL = sSQL & " DECLARE @IVA money"
                sSQL = sSQL & " DECLARE @ICMSSubstituicao money"
                sSQL = sSQL & " DECLARE @CodigoClassificacaoFiscal money"
                sSQL = sSQL & " DECLARE @CodigoNaturezaOperacao int"
                sSQL = sSQL & " DECLARE @ValorTrocaItem money"
                sSQL = sSQL & " DECLARE @ValorBonificacaoItem money"
                sSQL = sSQL & " DECLARE @QuantidadeTroca money"
                sSQL = sSQL & " DECLARE @CodigoUnidade money"
                sSQL = sSQL & " DECLARE @FatorConversao smallint"
                sSQL = sSQL & " DECLARE @Controle varchar(20)"
                sSQL = sSQL & " DECLARE @CodigoDeposito smallint"
                sSQL = sSQL & " DECLARE @ValorLiquido decimal(19,4)"
                
                sSQL = sSQL & " SELECT @CodigoEmpresa = " & intEmpresa
                sSQL = sSQL & " SELECT @Sequencia = " & intSequencia
                sSQL = sSQL & " SELECT @SequenciaProduto = " & 0
                sSQL = sSQL & " SELECT @CodigoProduto = " & tpLNG(rsGrid("Codigo"))
                sSQL = sSQL & " SELECT @Quantidade = " & tpMOEDA(rsGrid("qtdeAfaturar"))
                sSQL = sSQL & " SELECT @Desconto = " & 0 'tpMoeda(rsGrid("Desconto"))
                sSQL = sSQL & " SELECT @ValorUnitario = " & Str(rsGrid("VlSaida") / rsGrid("qtdSaida"))
                sSQL = sSQL & " SELECT @CodigoVendedor = " & 0
                sSQL = sSQL & " SELECT @AliquotaIPI = " & tpMOEDA(rsGrid!AliquotaIPI)
                sSQL = sSQL & " SELECT @AliquotaICMS = " & 0
                sSQL = sSQL & " SELECT @AliquotaReducao = " & 0
                sSQL = sSQL & " SELECT @IVA = " & 0
                sSQL = sSQL & " SELECT @ICMSSubstituicao = " & 0
                sSQL = sSQL & " SELECT @CodigoClassificacaoFiscal = " & 0
                sSQL = sSQL & " SELECT @CodigoNaturezaOperacao = " & 0
                sSQL = sSQL & " SELECT @ValorTrocaItem = 0"
                sSQL = sSQL & " SELECT @QuantidadeTroca = 0"
                sSQL = sSQL & " SELECT @ValorBonificacaoItem = 0"
                sSQL = sSQL & " SELECT @CodigoUnidade = " & rsGrid!CodigoUnidade
                sSQL = sSQL & " SELECT @FatorConversao = " & 0
                sSQL = sSQL & " SELECT @Controle = ''"
                sSQL = sSQL & " SELECT @CodigoDeposito = " & CodigoDepositoTIM
                VlrUnitario = Arredonda(rsGrid!VlSaida / rsGrid!qtdsaida, 2)
                sSQL = sSQL & " SELECT @ValorLiquido = " & Str(rsGrid("qtdeAfaturar") * VlrUnitario)
                
                'Armazenando Quantidade para usar no caso se for o mesmo produto para atualizar no capaitem
                sSQL = sSQL & " EXEC spGravaITEM @CodigoEmpresa, @Sequencia, @SequenciaProduto, @CodigoProduto, @Quantidade, @Desconto, @ValorUnitario, @CodigoVendedor, @AliquotaIPI, @AliquotaICMS, @AliquotaReducao, @IVA, @ICMSSubstituicao, @CodigoClassificacaoFiscal, 1, @CodigoNaturezaOperacao, @ValorTrocaItem, @ValorBonificacaoItem, @QuantidadeTroca, @CodigoUnidade, @FatorConversao, @CodigoDeposito, @Controle, ValorLiquido"
                db.Execute sSQL
                If rsGrid!Desconto > 0 Then
                    'ValorDesconto = Arredondar(((rsGrid("QtdeAFaturar") * rsGrid!ValorUnitario) * rsGrid!Desconto) / 100, 2)
                    'db.Execute "UPDATE CapaItem SET Desconto = " & tpMoeda(rsGrid!Desconto) & ", ValorUnitario = " & tpMoeda(rsGrid!ValorUnitario) & ", ValorDescontoItem = " & tpMoeda(ValorDesconto) & "    WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia & " and CodigoProduto = " & rsGrid!Codigo
                    'db.Execute "UPDATE CapaItem SET Desconto = " & tpMoeda(rsGrid!Desconto) & ", ValorUnitario = " & tpMoeda(rsGrid!ValorUnitario) & ", ValorDescontoItem = " & tpMoeda(ValorDesconto) & "    WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia & " and CodigoProduto = " & rsGrid!Codigo
                End If
    'Rotina Caso O Item Seja o Mesmo
pula:
                'Rotina para zerar quantidade a Faturar
                'db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                sSQL = "UPDATE CapaItem SET QuantidadeConferida = QuantidadeConferida - " & tpMOEDA(rsGrid("qtdeAfaturar")) & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia") & " and SequenciaProduto = " & rsGrid!SequenciaOriginal
                db.Execute sSQL
                
                sSQL = "INSERT INTO CapaItemRelacionaCapaItem (CodigoEmpresa, Sequencia, SequenciaProduto, SequenciaOriginal, SequenciaProdutoOriginal, CodigoProduto, CodigoUnidade) "
                sSQL = sSQL & " VALUES (" & intEmpresa & "," & intSequencia & "," & intSequenciaItem & "," & rsGrid!Sequencia & "," & rsGrid!SequenciaOriginal & "," & rsGrid!codigo & "," & rsGrid!CodigoUnidade & ")"
                db.Execute sSQL
                Complemento 'Chama uma rotina que grava numero do pedido e observacoes no Itens quando faz o agrupamento de produtos, Nao foi possivel implementar a rotina aqio pois estava dando um erro de procedure too large ou seja estorou o tamanho da procedure Emanoel
                
                IntNotaAnterior = rsGrid!Nota
                SequenciaAnterior = intSequencia
            End If
            DoEvents:
            
            sSQL = "EXEC spProcessaConsignacao  1,0,'',''," & rsGrid("CodigoCliente") & ",0,0,0,0," & tpLNG(rsGrid("Codigo")) & ",'D',60,'A','C'"
            db.Execute sSQL
            
            rsGrid.MoveNext
        Loop
        
        strMsgNota = "Simples faturamento de merc. em consig. mercantil ref. NFs: "
        strMsgNota = strMsgNota & strNumeroNota
        strMsgNotaImposto = ValorIPIItem & " e ICMS R$ " & ValorICMSITEM
        strMsgNota = strMsgNota & " Os impostos foram pagos através das citadas NFs: IPI R$ " & strMsgNotaImposto
        db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "' WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
        
        db.Execute "exec spAtualizaItem " & intEmpresa & ", '" & intSequencia & "','Capa'"
        db.Execute "exec spAtualizaCapa " & intEmpresa & ", '" & intSequencia & "','Capa'"
        
        strNumeroNota = ""
        ValorComDescontoAcumulado = 0
        ValorComDesconto = 0
        ValorDesconto = 0
        strMsgNota = ""
        VlrIPITotal = 0
        VlrICMSTotal = 0
        intSequencia = 0
        IntNotaAnterior = 0
        SequenciaAnterior = 0
        ValorIPIItem = 0
        ValorICMSITEM = 0
    End If
    
    If FaturarTeste = True Then
       'Pegando o Deposito e o Tipo de Estoque da Tim
        
        sSQL = "SELECT CodigoDeposito, Estoque "
        sSQL = sSQL & "From TIM "
        sSQL = sSQL & "Where Codigo = " & IntTimTesteVenda
        rs.Open sSQL, db, , , adCmdText
        If Not rs.EOF Then
           If strRetorno = "" Then
              CodigoDepositoTIM = rs!CodigoDeposito
              Else
                 CodigoDepositoTIM = Str(Left(strRetorno, 1))
           End If
           If rs!Estoque = 0 Then
              EstoqueTIM = "N" 'Nenhum
           ElseIf rs!Estoque = 1 Then
              EstoqueTIM = "S" 'Saida
           ElseIf rs!Estoque = 2 Then
              EstoqueTIM = "E" 'Entrada
           ElseIf rs!Estoque = 3 Then
              EstoqueTIM = "R" 'Reserva
           End If
        End If
        rs.Close

        intSequencia = 0
        intSequenciaItem = 0
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            DoEvents:
            If rsGrid("QtdeAfaturar") > 0 Then
               
               intSequenciaProduto = intSequenciaProduto + 1
               intSequenciaItem = intSequenciaItem + 1
               varsql = "SELECT  Capa.CodigoEmitente, Capa.Quantidade, Capa.ValorTotal, Capa.ValorLiquido "
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE Capa.CodigoEmpresa=" & intEmpresa & " AND Capa.CodigoEmitente = " & rsGrid("CodigoCliente") & " AND Capa.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                  'Acumulando Valores Caso seja mesmo Cliente
                    qtdfaturar = qtdfaturar + rsGrid!QtdeAfaturar
                    VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                    VlrFaturar = VlrFaturar + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
                    
                    'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                       End If
                    End If
                    VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
                    lblvlrfaturar1.Caption = Replace(VlrFaturar, ",", ".")
                    lblvlrfaturarIPI.Caption = Replace(VlrFaturarIPI, ",", ".")
                    
                   'Atualizando Cabeca do pedido
                   db.Execute "UPDATE Capa SET Quantidade = " & qtdfaturar & ", ValorTotal = " & lblvlrfaturarIPI.Caption & ", ValorLiquido = " & lblvlrfaturar1.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo PulaClienteFteste
               End If
               rs.Close
               
               'Calculando Valor a Faturar
               qtdfaturar = rsGrid!QtdeAfaturar
               VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
               VlrFaturar = Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
               'Pegando Valor IPI de Acordo com o Percentual do Produto
               If booCalculaIPI = True Then
                  If rsGrid!AliquotaIPI > 0 Then
                     VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                     VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                     VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
                  End If
               End If
               VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
               lblvlrfaturar1.Caption = Replace(VlrFaturar, ",", ".")
               lblvlrfaturarIPI.Caption = Replace(VlrFaturarIPI, ",", ".")
               
               intSequencia = BuscaCodigo("SELECT isnull(max(Sequencia),0) + 1 as Total FROM Capa WHERE CodigoEmpresa = " & intEmpresa, db, "Total")
               'Copiando Remessa Consignacao mudando algumas particularidades como sequencia e valores e tim devolucao
               varsql = "INSERT INTO Capa (CodigoEmpresa, Sequencia, CodigoFilial, Tipo, CodigoSituacao, CodigoVendedor, CodigoTabelaPreco, CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, Requisicao, Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia, Quantidade, ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca, ValorLiquido, ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor,  ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, TipoFrete, BaseICMS, ValorICMS, BaseIPI, ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, "
               varsql = varsql & " DataCadastro, DataFaturamento, DataEntradaSaida, DataContabilizacao, DataMovimentacao, CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, CIFFOB, CodigoCentroCusto, CodigoNatureza, CodigoNaturezaOperacao, Confirmada, CodigoMotivo, AtualizaFinanceiro, AtualizaEstoque, CodigoRequisitante, CodigoAutorizacao, CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, DataCancelamento, CodigoUsuarioCancelamento, DataAlteracao, CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, EmiteNF, EmiteCupom, Acertado, CodigoContaCorrente, "
               varsql = varsql & " TipoVenda, TipoComissao, PercentualComissao, Placa, UFPlaca)"
               varsql = varsql & " SELECT CodigoEmpresa, " & intSequencia & " AS Sequencia, CodigoFilial, '" & strTipoMovimento & "' as Tipo, 1 as CodigoSituacao, CodigoVendedor, CodigoTabelaPreco," & IntTimTesteVenda & " AS CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, " & intSequencia & " as Requisicao, " & 0 & " AS Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia," & rsGrid("QtdeAFaturar") & " as  Quantidade," & lblvlrfaturarIPI.Caption & " as ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca," & lblvlrfaturar1.Caption & " as ValorLiquido," & 0 & " as ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor, ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, "
               varsql = varsql & " TipoFrete, " & 0 & " as  BaseICMS, " & 0 & " as  ValorICMS, " & 0 & " as  BaseIPI, " & 0 & " as  ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, getdate() as DataCadastro, '" & strDataBase & "' as DataFaturamento, '" & strDataBase & "' as DataEntradaSaida, '" & strDataBase & "' as DataContabilizacao, '" & strDataBase & "' as DataMovimentacao, " & intUsuario & " as CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, '" & "C" & "' as CIFFOB, CodigoCentroCusto, CodigoNatureza, " & 0 & " as CodigoNaturezaOperacao, Confirmada, CodigoMotivo, '" & "N" & "' AS AtualizaFinanceiro, '" & EstoqueTIM & "' AS AtualizaEstoque, "
               varsql = varsql & " CodigoRequisitante, CodigoAutorizacao, " & 0 & " as CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, null as DataCancelamento, null as CodigoUsuarioCancelamento, null as DataAlteracao, null as CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, 0 as EmiteNF, 0 as EmiteCupom, Acertado, CodigoContaCorrente, TipoVenda, TipoComissao, PercentualComissao, '" & 0 & "' as Placa, UFPlaca"
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & rsGrid("Sequencia")
               db.Execute varsql
               
               'Verificando se tem alguma coisa e deleta antes de iniciar processo
               db.Execute "DELETE FROM CapaItem WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
               
    'Rotina Caso O Cliente Seja O mesmo
PulaClienteFteste:
               
               If SequenciaAnterior > 0 Then
                    If intSequencia <> SequenciaAnterior Then
                       strMsgNota = "Mercadoria de nossa  propriedade que segue para teste, devendo retornar em 30 dias."
                       db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & SequenciaAnterior
                    End If
               End If
               
               'Relacionado Capa ao caparelacinaCapa
               varsql = "SELECT  CapaRelacionaCapa.Sequencia, CapaRelacionaCapa.Sequencia "
               varsql = varsql & " FROM CapaRelacionaCapa "
               varsql = varsql & " WHERE CapaRelacionaCapa.CodigoEmpresa=" & intEmpresa & " AND CapaRelacionaCapa.sequencia = " & intSequencia & " AND CapaRelacionaCapa.SequenciaOriginal = " & rsGrid("Sequencia")
               rs.Open varsql, db, , , adCmdText
               If rs.EOF Then
                  db.Execute "INSERT INTO CapaRelacionaCapa (CodigoEmpresa, Sequencia, SequenciaOriginal) VALUES (" & intEmpresa & "," & intSequencia & "," & rsGrid("Sequencia") & ")"
               End If
               rs.Close
               
               varsql = "SELECT  Capaitem.CodigoProduto, Capaitem.Quantidade, Capaitem.ValorUnitario, Capaitem.ValorLiquido "
               varsql = varsql & " FROM Capaitem "
               varsql = varsql & " WHERE Capaitem.CodigoEmpresa=" & intEmpresa & " AND Capaitem.CodigoProduto = " & rsGrid("Codigo") & " AND Capaitem.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                   'Acumulando Valores Caso seja o Mesmo Produto mesmo Cliente
                   QtdFaturaItem = rs!Quantidade + rsGrid("QtdeAFaturar")
                   'Pegando Valor unitario do que vem para atualizar
                   lblvlrUniItem.Caption = (rsGrid("VlSaida") / rsGrid("qtdSaida"))
                   'Calculando o valor unitario com a quantidade solicitada
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption * rsGrid("qtdeAfaturar")
                   'Calculando total Quantidade anterior com quantidade atual
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption + (rs!Quantidade * rs!ValorUnitario)
                   'Por Fim pegando o valor Unitario pela divisao das quantidades e valores
                   lblvlrUniItem.Caption = (lblvlrUniItem.Caption / QtdFaturaItem)
                   lblvlrLiqItem.Caption = lblvlrUniItem.Caption * QtdFaturaItem
                   
                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((lblvlrUniItem.Caption * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                       End If
                    End If
                    lblvlrTotItem.Caption = Arredonda(lblvlrLiqItem.Caption + VlrIPI, 2)
                   
                    lblvlrUniItem.Caption = Replace(lblvlrUniItem.Caption, ",", ".")
                    lblvlrLiqItem.Caption = Replace(lblvlrLiqItem.Caption, ",", ".")
                    lblvlrTotItem.Caption = Replace(lblvlrTotItem.Caption, ",", ".")
                   'Atualizando Itens  do pedido
                   db.Execute "UPDATE CapaItem SET Quantidade = " & QtdFaturaItem & ", ValorUnitario = " & lblvlrUniItem.Caption & ", ValorLiquido = " & lblvlrLiqItem.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo pulaFTeste
               End If
               rs.Close
               
               
                sSQL = " DECLARE @CodigoEmpresa tinyint"
                sSQL = sSQL & " DECLARE @Sequencia int"
                sSQL = sSQL & " DECLARE @SequenciaProduto tinyint"
                sSQL = sSQL & " DECLARE @CodigoProduto int"
                sSQL = sSQL & " DECLARE @Quantidade money"
                sSQL = sSQL & " DECLARE @Desconto decimal(19,4)"
                sSQL = sSQL & " DECLARE @ValorUnitario decimal(19,4)"
                sSQL = sSQL & " DECLARE @CodigoVendedor smallint"
                sSQL = sSQL & " DECLARE @AliquotaIPI money"
                sSQL = sSQL & " DECLARE @AliquotaICMS money"
                sSQL = sSQL & " DECLARE @AliquotaReducao money"
                sSQL = sSQL & " DECLARE @IVA money"
                sSQL = sSQL & " DECLARE @ICMSSubstituicao money"
                sSQL = sSQL & " DECLARE @CodigoClassificacaoFiscal money"
                sSQL = sSQL & " DECLARE @CodigoNaturezaOperacao int"
                sSQL = sSQL & " DECLARE @ValorTrocaItem money"
                sSQL = sSQL & " DECLARE @ValorBonificacaoItem money"
                sSQL = sSQL & " DECLARE @QuantidadeTroca money"
                sSQL = sSQL & " DECLARE @CodigoUnidade money"
                sSQL = sSQL & " DECLARE @FatorConversao smallint"
                sSQL = sSQL & " DECLARE @Controle varchar(20)"
                sSQL = sSQL & " DECLARE @CodigoDeposito smallint"
                sSQL = sSQL & " DECLARE @ValorLiquido decimal(19,4)"
                
                sSQL = sSQL & " SELECT @CodigoEmpresa = " & intEmpresa
                sSQL = sSQL & " SELECT @Sequencia = " & intSequencia
                sSQL = sSQL & " SELECT @SequenciaProduto = " & 0
                sSQL = sSQL & " SELECT @CodigoProduto = " & tpLNG(rsGrid("Codigo"))
                sSQL = sSQL & " SELECT @Quantidade = " & tpMOEDA(rsGrid("qtdeAfaturar"))
                sSQL = sSQL & " SELECT @Desconto = " & 0
                sSQL = sSQL & " SELECT @ValorUnitario = " & Str(rsGrid("VlSaida") / rsGrid("qtdSaida"))
                sSQL = sSQL & " SELECT @CodigoVendedor = " & 0
                sSQL = sSQL & " SELECT @AliquotaIPI = " & 0
                sSQL = sSQL & " SELECT @AliquotaICMS = " & 0
                sSQL = sSQL & " SELECT @AliquotaReducao = " & 0
                sSQL = sSQL & " SELECT @IVA = " & 0
                sSQL = sSQL & " SELECT @ICMSSubstituicao = " & 0
                sSQL = sSQL & " SELECT @CodigoClassificacaoFiscal = " & 0
                sSQL = sSQL & " SELECT @CodigoNaturezaOperacao = " & 0
                sSQL = sSQL & " SELECT @ValorTrocaItem = 0"
                sSQL = sSQL & " SELECT @QuantidadeTroca = 0"
                sSQL = sSQL & " SELECT @ValorBonificacaoItem = 0"
                sSQL = sSQL & " SELECT @CodigoUnidade = " & rsGrid!CodigoUnidade
                sSQL = sSQL & " SELECT @FatorConversao = " & 0
                sSQL = sSQL & " SELECT @Controle = ''"
                sSQL = sSQL & " SELECT @CodigoDeposito = " & CodigoDepositoTIM
                sSQL = sSQL & " SELECT @ValorLiquido = " & lblvlrfaturarIPI.Caption
                'Armazenando Quantidade para usar no caso se for o mesmo produto para atualizar no capaitem
                sSQL = sSQL & " EXEC spGravaITEM @CodigoEmpresa, @Sequencia, @SequenciaProduto, @CodigoProduto, @Quantidade, @Desconto, @ValorUnitario, @CodigoVendedor, @AliquotaIPI, @AliquotaICMS, @AliquotaReducao, @IVA, @ICMSSubstituicao, @CodigoClassificacaoFiscal, 1, @CodigoNaturezaOperacao, @ValorTrocaItem, @ValorBonificacaoItem, @QuantidadeTroca, @CodigoUnidade, @FatorConversao, @CodigoDeposito, @Controle, ValorLiquido"
                db.Execute sSQL
                
                If rsGrid!Desconto > 0 Then
                    ValorDesconto = Arredondar(((rsGrid("QtdeAFaturar") * rsGrid!ValorUnitario) * rsGrid!Desconto) / 100, 2)
                    db.Execute "UPDATE CapaItem SET Desconto = " & tpMOEDA(rsGrid!Desconto) & ", ValorUnitario = " & tpMOEDA(rsGrid!ValorUnitario) & ", ValorDescontoItem = " & tpMOEDA(ValorDesconto) & "    WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia & " and CodigoProduto = " & rsGrid!codigo
                End If
                
                
                
    'Rotina Caso O Item Seja o Mesmo
pulaFTeste:
                
                
                'Rotina para zerar quantidade a Faturar
                db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                
                db.Execute "INSERT INTO CapaItemRelacionaCapaItem (CodigoEmpresa, Sequencia, SequenciaProduto, SequenciaOriginal, SequenciaProdutoOriginal, CodigoProduto, CodigoUnidade) VALUES (" & intEmpresa & "," & intSequencia & "," & intSequenciaItem & "," & rsGrid!Sequencia & "," & rsGrid!SequenciaOriginal & "," & rsGrid!codigo & "," & rsGrid!CodigoUnidade & ")"
                
                Complemento 'Chama uma rotina que grava numero do pedido e observacoes no Itens quando faz o agrupamento de produtos, Nao foi possivel implementar a rotina aqio pois estava dando um erro de procedure too large ou seja estorou o tamanho da procedure Emanoel
            End If
            
            sSQL = "EXEC spProcessaConsignacao  1,0,'',''," & rsGrid("CodigoCliente") & ",0,0,0,0," & tpLNG(rsGrid("Codigo")) & ",'D',60,'A','C'"
            db.Execute sSQL
            
            rsGrid.MoveNext
        Loop
        
        strMsgNota = "Mercadoria de nossa  propriedade que segue para teste, devendo retornar em 30 dias."
        db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
        db.Execute "exec spAtualizaItem " & intEmpresa & ", '" & intSequencia & "','Capa'"
        db.Execute "exec spAtualizaCapa " & intEmpresa & ", '" & intSequencia & "','Capa'"
        
    End If
    
    If FaturarEmprestimo = True Then
        
        'Pegando o Deposito e o Tipo de Estoque da Tim
        
        sSQL = "SELECT CodigoDeposito, Estoque "
        sSQL = sSQL & "From TIM "
        sSQL = sSQL & "Where Codigo = " & IntTimEmprestimoVenda
        rs.Open sSQL, db, , , adCmdText
        If Not rs.EOF Then
           If strRetorno = "" Then
              CodigoDepositoTIM = rs!CodigoDeposito
              Else
                 CodigoDepositoTIM = Str(Left(strRetorno, 1))
           End If
           If rs!Estoque = 0 Then
              EstoqueTIM = "N" 'Nenhum
           ElseIf rs!Estoque = 1 Then
              EstoqueTIM = "S" 'Saida
           ElseIf rs!Estoque = 2 Then
              EstoqueTIM = "E" 'Entrada
           ElseIf rs!Estoque = 3 Then
              EstoqueTIM = "R" 'Reserva
           End If
        End If
        rs.Close
        
        intSequencia = 0
        intSequenciaItem = 0
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            DoEvents:
            If rsGrid("QtdeAfaturar") > 0 Then
               
               intSequenciaProduto = intSequenciaProduto + 1
               intSequenciaItem = intSequenciaItem + 1
               varsql = "SELECT  Capa.CodigoEmitente, Capa.Quantidade, Capa.ValorTotal, Capa.ValorLiquido "
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE Capa.CodigoEmpresa=" & intEmpresa & " AND Capa.CodigoEmitente = " & rsGrid("CodigoCliente") & " AND Capa.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                    'Acumulando Valores Caso seja mesmo Cliente
                    qtdfaturar = qtdfaturar + rsGrid!QtdeAfaturar
                    VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                    VlrFaturar = VlrFaturar + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
                    
                    'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                       End If
                    End If
                    VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
                    lblvlrfaturar1.Caption = Replace(VlrFaturar, ",", ".")
                    lblvlrfaturarIPI.Caption = Replace(VlrFaturarIPI, ",", ".")
                    
                    'Atualizando Cabeca do pedido
                   db.Execute "UPDATE Capa SET Quantidade = " & qtdfaturar & ", ValorTotal = " & lblvlrfaturarIPI.Caption & ", ValorLiquido = " & lblvlrfaturar1.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
                   rs.Close
                 GoTo PulaClienteFEmp
               End If
               rs.Close
               
               'Calculando Valor a Faturar
               qtdfaturar = rsGrid!QtdeAfaturar
               VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
               VlrFaturar = Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
               'Pegando Valor IPI de Acordo com o Percentual do Produto
               If booCalculaIPI = True Then
                  If rsGrid!AliquotaIPI > 0 Then
                     VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                     VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                     VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
                  End If
               End If
               VlrFaturarIPI = Arredonda(VlrFaturar + VlrIPI, 2)
               lblvlrfaturar1.Caption = Replace(VlrFaturar, ",", ".")
               lblvlrfaturarIPI.Caption = Replace(VlrFaturarIPI, ",", ".")
               
               intSequencia = BuscaCodigo("SELECT isnull(max(Sequencia),0) + 1 as Total FROM Capa WHERE CodigoEmpresa = " & intEmpresa, db, "Total")
               'Copiando Remessa Consignacao mudando algumas particularidades como sequencia e valores e tim devolucao
               varsql = "INSERT INTO Capa (CodigoEmpresa, Sequencia, CodigoFilial, Tipo, CodigoSituacao, CodigoVendedor, CodigoTabelaPreco, CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, Requisicao, Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia, Quantidade, ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca, ValorLiquido, ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor,  ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, TipoFrete, BaseICMS, ValorICMS, BaseIPI, ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, "
               varsql = varsql & " DataCadastro, DataFaturamento, DataEntradaSaida, DataContabilizacao, DataMovimentacao, CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, CIFFOB, CodigoCentroCusto, CodigoNatureza, CodigoNaturezaOperacao, Confirmada, CodigoMotivo, AtualizaFinanceiro, AtualizaEstoque, CodigoRequisitante, CodigoAutorizacao, CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, DataCancelamento, CodigoUsuarioCancelamento, DataAlteracao, CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, EmiteNF, EmiteCupom, Acertado, CodigoContaCorrente, "
               varsql = varsql & " TipoVenda, TipoComissao, PercentualComissao, Placa, UFPlaca) "
               varsql = varsql & " SELECT CodigoEmpresa, " & intSequencia & " AS Sequencia, CodigoFilial, '" & strTipoMovimento & "' as Tipo, 1 as CodigoSituacao, CodigoVendedor, CodigoTabelaPreco," & IntTimEmprestimoVenda & " AS CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, " & intSequencia & " as Requisicao, " & 0 & " AS Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia," & rsGrid("QtdeAFaturar") & " as  Quantidade," & lblvlrfaturarIPI.Caption & " as ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca," & lblvlrfaturar1.Caption & " as ValorLiquido," & 0 & " as ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor, ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, "
               varsql = varsql & " TipoFrete, " & 0 & " as  BaseICMS, " & 0 & " as  ValorICMS, " & 0 & " as  BaseIPI, " & 0 & " as  ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, getdate() as DataCadastro, '" & strDataBase & "' as DataFaturamento, '" & strDataBase & "' as DataEntradaSaida, '" & strDataBase & "' as DataContabilizacao, '" & strDataBase & "' as DataMovimentacao, " & intUsuario & " as CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, '" & "C" & "' as CIFFOB, CodigoCentroCusto, CodigoNatureza, " & 0 & " as CodigoNaturezaOperacao, Confirmada, CodigoMotivo, '" & "N" & "' AS AtualizaFinanceiro, '" & EstoqueTIM & "' AS AtualizaEstoque, "
               varsql = varsql & " CodigoRequisitante, CodigoAutorizacao, " & 0 & " as CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, null as DataCancelamento, null as CodigoUsuarioCancelamento, null as DataAlteracao, null as CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, 0 as EmiteNF, 0 as EmiteCupom, Acertado, CodigoContaCorrente, TipoVenda, TipoComissao, PercentualComissao, '" & 0 & "' as Placa, UFPlaca"
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & rsGrid("Sequencia")
               db.Execute varsql
               
               'Verificando se tem alguma coisa e deleta antes de iniciar processo
               db.Execute "DELETE FROM CapaItem WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
               
    'Rotina Caso O Cliente Seja O mesmo
PulaClienteFEmp:
               
               If SequenciaAnterior > 0 Then
                  If intSequencia <> SequenciaAnterior Then
                     strMsgNota = "Mercadoria de nossa  propriedade que segue para Emprestimo, devendo retornar em 30 dias."
                     db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & SequenciaAnterior
                  End If
               End If
               
               'Relacionado Capa ao caparelacinaCapa
               varsql = "SELECT  CapaRelacionaCapa.Sequencia, CapaRelacionaCapa.Sequencia "
               varsql = varsql & " FROM CapaRelacionaCapa "
               varsql = varsql & " WHERE CapaRelacionaCapa.CodigoEmpresa=" & intEmpresa & " AND CapaRelacionaCapa.sequencia = " & intSequencia & " AND CapaRelacionaCapa.SequenciaOriginal = " & rsGrid("Sequencia")
               rs.Open varsql, db, , , adCmdText
               If rs.EOF Then
                  db.Execute "INSERT INTO CapaRelacionaCapa (CodigoEmpresa, Sequencia, SequenciaOriginal) VALUES (" & intEmpresa & "," & intSequencia & "," & rsGrid("Sequencia") & ")"
               End If
               rs.Close
               
               varsql = "SELECT  Capaitem.CodigoProduto, Capaitem.Quantidade, Capaitem.ValorUnitario, Capaitem.ValorLiquido "
               varsql = varsql & " FROM Capaitem "
               varsql = varsql & " WHERE Capaitem.CodigoEmpresa=" & intEmpresa & " AND Capaitem.CodigoProduto = " & rsGrid("Codigo") & " AND Capaitem.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                   'Acumulando Valores Caso seja o Mesmo Produto mesmo Cliente
                   QtdFaturaItem = rs!Quantidade + rsGrid("QtdeAFaturar")
                   'Pegando Valor unitario do que vem para atualizar
                   lblvlrUniItem.Caption = (rsGrid("VlSaida") / rsGrid("qtdSaida"))
                   'Calculando o valor unitario com a quantidade solicitada
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption * rsGrid("qtdeAfaturar")
                   'Calculando total Quantidade anterior com quantidade atual
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption + (rs!Quantidade * rs!ValorUnitario)
                   'Por Fim pegando o valor Unitario pela divisao das quantidades e valores
                   lblvlrUniItem.Caption = (lblvlrUniItem.Caption / QtdFaturaItem)
                   lblvlrLiqItem.Caption = lblvlrUniItem.Caption * QtdFaturaItem
                   
                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((lblvlrUniItem.Caption * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                       End If
                    End If
                    lblvlrTotItem.Caption = Arredonda(lblvlrLiqItem.Caption + VlrIPI, 2)
                   
                    lblvlrUniItem.Caption = Replace(lblvlrUniItem.Caption, ",", ".")
                    lblvlrLiqItem.Caption = Replace(lblvlrLiqItem.Caption, ",", ".")
                    lblvlrTotItem.Caption = Replace(lblvlrTotItem.Caption, ",", ".")
                   'Atualizando Itens  do pedido
                   db.Execute "UPDATE CapaItem SET Quantidade = " & QtdFaturaItem & ", ValorUnitario = " & lblvlrUniItem.Caption & ", ValorLiquido = " & lblvlrLiqItem.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo pulaFEmp
               End If
               rs.Close
               
                sSQL = " DECLARE @CodigoEmpresa tinyint"
                sSQL = sSQL & " DECLARE @Sequencia int"
                sSQL = sSQL & " DECLARE @SequenciaProduto tinyint"
                sSQL = sSQL & " DECLARE @CodigoProduto int"
                sSQL = sSQL & " DECLARE @Quantidade money"
                sSQL = sSQL & " DECLARE @Desconto decimal(19,4)"
                sSQL = sSQL & " DECLARE @ValorUnitario decimal(19,4)"
                sSQL = sSQL & " DECLARE @CodigoVendedor smallint"
                sSQL = sSQL & " DECLARE @AliquotaIPI money"
                sSQL = sSQL & " DECLARE @AliquotaICMS money"
                sSQL = sSQL & " DECLARE @AliquotaReducao money"
                sSQL = sSQL & " DECLARE @IVA money"
                sSQL = sSQL & " DECLARE @ICMSSubstituicao money"
                sSQL = sSQL & " DECLARE @CodigoClassificacaoFiscal money"
                sSQL = sSQL & " DECLARE @CodigoNaturezaOperacao int"
                sSQL = sSQL & " DECLARE @ValorTrocaItem money"
                sSQL = sSQL & " DECLARE @ValorBonificacaoItem money"
                sSQL = sSQL & " DECLARE @QuantidadeTroca money"
                sSQL = sSQL & " DECLARE @CodigoUnidade money"
                sSQL = sSQL & " DECLARE @FatorConversao smallint"
                sSQL = sSQL & " DECLARE @Controle varchar(20)"
                sSQL = sSQL & " DECLARE @CodigoDeposito smallint"
                sSQL = sSQL & " DECLARE @ValorLiquido decimal(19,4)"
                
                sSQL = sSQL & " SELECT @CodigoEmpresa = " & intEmpresa
                sSQL = sSQL & " SELECT @Sequencia = " & intSequencia
                sSQL = sSQL & " SELECT @SequenciaProduto = " & 0
                sSQL = sSQL & " SELECT @CodigoProduto = " & tpLNG(rsGrid("Codigo"))
                sSQL = sSQL & " SELECT @Quantidade = " & tpMOEDA(rsGrid("qtdeAfaturar"))
                sSQL = sSQL & " SELECT @Desconto = " & 0
                sSQL = sSQL & " SELECT @ValorUnitario = " & Str(rsGrid("VlSaida") / rsGrid("qtdSaida"))
                sSQL = sSQL & " SELECT @CodigoVendedor = " & 0
                sSQL = sSQL & " SELECT @AliquotaIPI = " & 0
                sSQL = sSQL & " SELECT @AliquotaICMS = " & 0
                sSQL = sSQL & " SELECT @AliquotaReducao = " & 0
                sSQL = sSQL & " SELECT @IVA = " & 0
                sSQL = sSQL & " SELECT @ICMSSubstituicao = " & 0
                sSQL = sSQL & " SELECT @CodigoClassificacaoFiscal = " & 0
                sSQL = sSQL & " SELECT @CodigoNaturezaOperacao = " & 0
                sSQL = sSQL & " SELECT @ValorTrocaItem = 0"
                sSQL = sSQL & " SELECT @QuantidadeTroca = 0"
                sSQL = sSQL & " SELECT @ValorBonificacaoItem = 0"
                sSQL = sSQL & " SELECT @CodigoUnidade = " & rsGrid!CodigoUnidade
                sSQL = sSQL & " SELECT @FatorConversao = " & 0
                sSQL = sSQL & " SELECT @Controle = ''"
                sSQL = sSQL & " SELECT @CodigoDeposito = " & CodigoDepositoTIM
                sSQL = sSQL & " SELECT @ValorLiquido = " & lblvlrfaturarIPI.Caption
                'Armazenando Quantidade para usar no caso se for o mesmo produto para atualizar no capaitem
                sSQL = sSQL & " EXEC spGravaITEM @CodigoEmpresa, @Sequencia, @SequenciaProduto, @CodigoProduto, @Quantidade, @Desconto, @ValorUnitario, @CodigoVendedor, @AliquotaIPI, @AliquotaICMS, @AliquotaReducao, @IVA, @ICMSSubstituicao, @CodigoClassificacaoFiscal, 1, @CodigoNaturezaOperacao, @ValorTrocaItem, @ValorBonificacaoItem, @QuantidadeTroca, @CodigoUnidade, @FatorConversao, @CodigoDeposito, @Controle, ValorLiquido"
                db.Execute sSQL
                
                If rsGrid!Desconto > 0 Then
                    ValorDesconto = Arredondar(((rsGrid("QtdeAFaturar") * rsGrid!ValorUnitario) * rsGrid!Desconto) / 100, 2)
                    db.Execute "UPDATE CapaItem SET Desconto = " & tpMOEDA(rsGrid!Desconto) & ", ValorUnitario = " & tpMOEDA(rsGrid!ValorUnitario) & ", ValorDescontoItem = " & tpMOEDA(ValorDesconto) & "    WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia & " and CodigoProduto = " & rsGrid!codigo
                End If
    
                
    'Rotina Caso O Item Seja o Mesmo
pulaFEmp:
                'Rotina para zerar quantidade a Faturar
                db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                
                db.Execute "INSERT INTO CapaItemRelacionaCapaItem (CodigoEmpresa, Sequencia, SequenciaProduto, SequenciaOriginal, SequenciaProdutoOriginal, CodigoProduto, CodigoUnidade) VALUES (" & intEmpresa & "," & intSequencia & "," & intSequenciaItem & "," & rsGrid!Sequencia & "," & rsGrid!SequenciaOriginal & "," & rsGrid!codigo & "," & rsGrid!CodigoUnidade & ")"
                
                Complemento 'Chama uma rotina que grava numero do pedido e observacoes no Itens quando faz o agrupamento de produtos, Nao foi possivel implementar a rotina aqio pois estava dando um erro de procedure too large ou seja estorou o tamanho da procedure Emanoel
            End If
            
            sSQL = "EXEC spProcessaConsignacao  1,0,'',''," & rsGrid("CodigoCliente") & ",0,0,0,0," & tpLNG(rsGrid("Codigo")) & ",'D',60,'A','C'"
            db.Execute sSQL
            
            rsGrid.MoveNext
        Loop
        
        strMsgNota = "Mercadoria de nossa  propriedade que segue para Emprestimo, devendo retornar em 30 dias."
        db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
        
        db.Execute "exec spAtualizaItem " & intEmpresa & ", '" & intSequencia & "','Capa'"
        db.Execute "exec spAtualizaCapa " & intEmpresa & ", '" & intSequencia & "','Capa'"
        
    End If
    
    If DevolverTeste = True Then
        intSequencia = 0
        intSequenciaItem = 0
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            DoEvents:
            If rsGrid("QtdeADevolver") > 0 Then
                intSequenciaItem = intSequenciaItem + 1
                varsql = "SELECT  Capa.CodigoEmitente, Capa.Quantidade, Capa.ValorTotal, Capa.ValorLiquido "
                varsql = varsql & " FROM Capa"
                varsql = varsql & " WHERE Capa.CodigoEmpresa=" & intEmpresa & " AND Capa.CodigoEmitente = " & rsGrid("CodigoCliente") & " AND Capa.Sequencia = " & intSequencia
                rs.Open varsql, db, , , adCmdText
                If Not rs.EOF Then
                    'Acumulando Valores Caso seja mesma Nota
                    qtdDevolver = qtdDevolver + rsGrid!QtdeADevolver
                    VlrUnitarioDevolucao = rsGrid!VlSaida / rsGrid!qtdsaida
                    VlrDevolver = VlrDevolver + Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
                    'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                        If rsGrid!AliquotaIPI > 0 Then
                            VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                            VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                        End If
                    End If
                    
                    VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
                    lblvlrfaturar1.Caption = Replace(VlrDevolver, ",", ".")
                    lblvlrfaturarIPI.Caption = Replace(VlrDevolverIPI, ",", ".")
                    
                    If rsGrid!AliquotaICMS > 0 Then
                        VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                        VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                        BaseICMS = BaseICMS + lblvlrfaturar1.Caption
                    End If
                    
                    lblIPI.Caption = Replace(VlrIPI, ",", ".")
                    lblICMS.Caption = Replace(VlrICMS, ",", ".")
                    lblBICMS.Caption = Replace(BaseICMS, ",", ".")
                    
                    'Atualizando Cabeca do pedido
                    db.Execute "UPDATE Capa SET Quantidade = " & qtdDevolver & ", ValorTotal = " & lblvlrfaturarIPI.Caption & ", ValorLiquido = " & lblvlrfaturar1.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
                    rs.Close
                    
                    GoTo PulaClienteDTeste
                End If
                rs.Close
                
                'Pegando o Deposito e o Tipo de Estoque da Tim
                sSQL = "SELECT TIM.TIMContraPartida"
                sSQL = sSQL & " FROM Capa INNER JOIN TIM ON Capa.CodigoEmpresa = TIM.CodigoEmpresa AND Capa.CodigoTIM = TIM.Codigo"
                sSQL = sSQL & " WHERE Capa.CodigoEmpresa = " & intEmpresa & " AND Sequencia = " & rsGrid("Sequencia")
                rs.Open sSQL, db, , , adCmdText
                If Not rs.EOF Then
                    IntTimTesteDevolve = rs!TIMContrapartida
                End If
                rs.Close
                 
                sSQL = "SELECT CodigoDeposito, Estoque "
                sSQL = sSQL & "From TIM "
                sSQL = sSQL & "Where Codigo = " & IntTimTesteDevolve
                rs.Open sSQL, db, , , adCmdText
                If Not rs.EOF Then
                    If strRetorno = "" Then
                        CodigoDepositoTIM = rs!CodigoDeposito
                    Else
                        CodigoDepositoTIM = Str(Left(strRetorno, 1))
                    End If
                    
                    If rs!Estoque = 0 Then
                        EstoqueTIM = "N" 'Nenhum
                    ElseIf rs!Estoque = 1 Then
                        EstoqueTIM = "S" 'Saida
                    ElseIf rs!Estoque = 2 Then
                        EstoqueTIM = "E" 'Entrada
                    ElseIf rs!Estoque = 3 Then
                        EstoqueTIM = "R" 'Reserva
                    End If
                End If
                rs.Close
                
               'Calculando Valor a Faturar
               qtdDevolver = rsGrid("qtdeAdevolver")
               VlrUnitarioDevolucao = rsGrid!VlSaida / rsGrid!qtdsaida
               VlrDevolver = Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
               'Pegando Valor IPI de Acordo com o Percentual do Produto
               If booCalculaIPI = True Then
                  If rsGrid!AliquotaIPI > 0 Then
                     VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                     VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                     VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
                  End If
               End If
               VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
               lblvlrfaturar1.Caption = Replace(VlrDevolver, ",", ".")
               lblvlrfaturarIPI.Caption = Replace(VlrDevolverIPI, ",", ".")
               lblIPI.Caption = Replace(VlrIPI, ",", ".")
               
               If rsGrid!AliquotaICMS > 0 Then
                  VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                  VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                  BaseICMS = lblvlrfaturar1.Caption
               End If
               
               lblICMS.Caption = Replace(VlrICMS, ",", ".")
               lblBICMS.Caption = Replace(BaseICMS, ",", ".")
                  
               intSequencia = BuscaCodigo("SELECT isnull(max(Sequencia),0) + 1 as Total FROM Capa WHERE CodigoEmpresa = " & intEmpresa, db, "Total")
               'Copiando Remessa Consignacao mudando algumas particularidades como sequencia e valores e tim devolucao
               varsql = "INSERT INTO Capa (CodigoEmpresa, Sequencia, CodigoFilial, Tipo, CodigoSituacao, CodigoVendedor, CodigoTabelaPreco, CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, Requisicao, Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia, Quantidade, ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca, ValorLiquido, ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor,  ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, TipoFrete, BaseICMS, ValorICMS, BaseIPI, ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, "
               varsql = varsql & " DataCadastro, DataFaturamento, DataEntradaSaida, DataContabilizacao, DataMovimentacao, CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, CIFFOB, CodigoCentroCusto, CodigoNatureza, CodigoNaturezaOperacao, Confirmada, CodigoMotivo, AtualizaFinanceiro, AtualizaEstoque, CodigoRequisitante, CodigoAutorizacao, CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, DataCancelamento, CodigoUsuarioCancelamento, DataAlteracao, CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, EmiteNF, EmiteCupom, Acertado, CodigoContaCorrente, "
               varsql = varsql & " TipoVenda, TipoComissao, PercentualComissao, Placa, UFPlaca) "
               varsql = varsql & " SELECT CodigoEmpresa, " & intSequencia & " AS Sequencia, CodigoFilial, '" & strTipoMovimento & "' as Tipo, 9 as CodigoSituacao, CodigoVendedor, CodigoTabelaPreco," & IntTimTesteDevolve & " AS CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, " & intSequencia & " as Requisicao, " & 0 & " AS Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia," & rsGrid("QtdeaDevolver") & " as  Quantidade," & lblvlrfaturarIPI.Caption & " as ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca," & lblvlrfaturar1.Caption & " as ValorLiquido," & 0 & " as ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor, ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, "
               varsql = varsql & " TipoFrete, " & lblBICMS.Caption & " as  BaseICMS, " & lblICMS.Caption & " as  ValorICMS, " & 0 & " as  BaseIPI, " & lblIPI.Caption & " as  ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, getdate() as DataCadastro, '" & strDataBase & "' as DataFaturamento, '" & strDataBase & "' as DataEntradaSaida, '" & strDataBase & "' as DataContabilizacao, '" & strDataBase & "' as DataMovimentacao, " & intUsuario & " as CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, '" & "C" & "' as CIFFOB, CodigoCentroCusto, CodigoNatureza, " & 0 & " as CodigoNaturezaOperacao, Confirmada, CodigoMotivo, '" & "N" & "' AS AtualizaFinanceiro, '" & EstoqueTIM & "' AS AtualizaEstoque, "
               varsql = varsql & " CodigoRequisitante, CodigoAutorizacao, " & 0 & " as CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, null as DataCancelamento, null as CodigoUsuarioCancelamento, null as DataAlteracao, null as CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, 0 as EmiteNF, 0 as EmiteCupom, Acertado, CodigoContaCorrente, TipoVenda, TipoComissao, PercentualComissao, '" & 0 & "' as Placa, UFPlaca"
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & rsGrid("Sequencia")
               db.Execute varsql
               
               'Verificando se tem alguma coisa e deleta antes de iniciar processo
               db.Execute "DELETE FROM CapaItem WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
               
    'Rotina Caso O Cliente Seja O mesmo
PulaClienteDTeste:
               
               If IntNotaAnterior <> rsGrid!Nota Then
                   If SequenciaAnterior > 0 Then
                        If intSequencia = SequenciaAnterior Then
                           If strNumeroNota <> "" Then
                              If Len(strNumeroNota) <= 500 Then
                                 strNumeroNota = strNumeroNota & "," & rsGrid!Nota
                              End If
                              Else
                                  strNumeroNota = rsGrid!Nota
                           End If
                        End If
                        Else
                           strNumeroNota = rsGrid!Nota
                   End If
               End If
               If SequenciaAnterior > 0 Then
                    If intSequencia <> SequenciaAnterior Then
                       strMsgNota = "Devolucao Teste Ref. NFs: "
                       strMsgNota = strMsgNota & strNumeroNota
                       db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & SequenciaAnterior
                       strNumeroNota = rsGrid!Nota
                    End If
               End If
               
               'Relacionado Capa ao caparelacinaCapa
               varsql = "SELECT  CapaRelacionaCapa.Sequencia, CapaRelacionaCapa.Sequencia "
               varsql = varsql & " FROM CapaRelacionaCapa "
               varsql = varsql & " WHERE CapaRelacionaCapa.CodigoEmpresa=" & intEmpresa & " AND CapaRelacionaCapa.sequencia = " & intSequencia & " AND CapaRelacionaCapa.SequenciaOriginal = " & rsGrid("Sequencia")
               rs.Open varsql, db, , , adCmdText
               If rs.EOF Then
                  db.Execute "INSERT INTO CapaRelacionaCapa (CodigoEmpresa, Sequencia, SequenciaOriginal) VALUES (" & intEmpresa & "," & intSequencia & "," & rsGrid("Sequencia") & ")"
               End If
               rs.Close
               
               varsql = "SELECT  Capaitem.CodigoProduto, Capaitem.Quantidade, Capaitem.ValorUnitario, Capaitem.ValorLiquido "
               varsql = varsql & " FROM Capaitem "
               varsql = varsql & " WHERE Capaitem.CodigoEmpresa=" & intEmpresa & " AND Capaitem.CodigoProduto = " & rsGrid("Codigo") & " AND Capaitem.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                   'Acumulando Valores Caso seja o Mesmo Produto mesmo Cliente
                   QtdFaturaItem = rs!Quantidade + rsGrid("QtdeADevolver")
                   'Pegando Valor unitario do que vem para atualizar
                   lblvlrUniItem.Caption = (rsGrid("VlSaida") / rsGrid("qtdSaida"))
                   'Calculando o valor unitario com a quantidade solicitada
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption * rsGrid("QtdeADevolver")
                   'Calculando total Quantidade anterior com quantidade atual
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption + (rs!Quantidade * rs!ValorUnitario)
                   'Por Fim pegando o valor Unitario pela divisao das quantidades e valores
                   lblvlrUniItem.Caption = (lblvlrUniItem.Caption / QtdFaturaItem)
                   lblvlrLiqItem.Caption = lblvlrUniItem.Caption * QtdFaturaItem
                   
                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((lblvlrUniItem.Caption * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid("QtdeADevolver"), 2)
                       End If
                    End If
                    lblvlrTotItem.Caption = Arredonda(lblvlrLiqItem.Caption + VlrIPI, 2)
                   
                    lblvlrUniItem.Caption = Replace(lblvlrUniItem.Caption, ",", ".")
                    lblvlrLiqItem.Caption = Replace(lblvlrLiqItem.Caption, ",", ".")
                    lblvlrTotItem.Caption = Replace(lblvlrTotItem.Caption, ",", ".")
                   'Atualizando Itens  do pedido
                   db.Execute "UPDATE CapaItem SET Quantidade = " & QtdFaturaItem & ", ValorUnitario = " & lblvlrUniItem.Caption & ", ValorLiquido = " & lblvlrLiqItem.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo pulaDTeste
               End If
               rs.Close
                intSequenciaProduto = intSequenciaProduto + 1
                sSQL = " DECLARE @CodigoEmpresa tinyint"
                sSQL = sSQL & " DECLARE @Sequencia int"
                sSQL = sSQL & " DECLARE @SequenciaProduto tinyint"
                sSQL = sSQL & " DECLARE @CodigoProduto int"
                sSQL = sSQL & " DECLARE @Quantidade money"
                sSQL = sSQL & " DECLARE @Desconto decimal(19,4)"
                sSQL = sSQL & " DECLARE @ValorUnitario decimal(19,4)"
                sSQL = sSQL & " DECLARE @CodigoVendedor smallint"
                sSQL = sSQL & " DECLARE @AliquotaIPI money"
                sSQL = sSQL & " DECLARE @AliquotaICMS money"
                sSQL = sSQL & " DECLARE @AliquotaReducao money"
                sSQL = sSQL & " DECLARE @IVA money"
                sSQL = sSQL & " DECLARE @ICMSSubstituicao money"
                sSQL = sSQL & " DECLARE @CodigoClassificacaoFiscal money"
                sSQL = sSQL & " DECLARE @CodigoNaturezaOperacao int"
                sSQL = sSQL & " DECLARE @ValorTrocaItem money"
                sSQL = sSQL & " DECLARE @ValorBonificacaoItem money"
                sSQL = sSQL & " DECLARE @QuantidadeTroca money"
                sSQL = sSQL & " DECLARE @CodigoUnidade money"
                sSQL = sSQL & " DECLARE @FatorConversao smallint"
                sSQL = sSQL & " DECLARE @Controle varchar(20)"
                sSQL = sSQL & " DECLARE @CodigoDeposito smallint"
                sSQL = sSQL & " DECLARE @QuantidadeAvarias smallint"
                sSQL = sSQL & " DECLARE @ValorLiquido decimal(19,4)"
                
                sSQL = sSQL & " SELECT @CodigoEmpresa = " & intEmpresa
                sSQL = sSQL & " SELECT @Sequencia = " & intSequencia
                sSQL = sSQL & " SELECT @SequenciaProduto = " & 0
                sSQL = sSQL & " SELECT @CodigoProduto = " & tpLNG(rsGrid("Codigo"))
                sSQL = sSQL & " SELECT @Quantidade = " & tpMOEDA(rsGrid("qtdeADevolver"))
                sSQL = sSQL & " SELECT @Desconto = " & 0
                sSQL = sSQL & " SELECT @ValorUnitario = " & tpMOEDA(rsGrid("VlSaida") / rsGrid("qtdSaida"))
                sSQL = sSQL & " SELECT @CodigoVendedor = " & 0
                sSQL = sSQL & " SELECT @AliquotaIPI = " & 0
                sSQL = sSQL & " SELECT @AliquotaICMS = " & 0
                sSQL = sSQL & " SELECT @AliquotaReducao = " & 0
                sSQL = sSQL & " SELECT @IVA = " & 0
                sSQL = sSQL & " SELECT @ICMSSubstituicao = " & 0
                sSQL = sSQL & " SELECT @CodigoClassificacaoFiscal = " & 0
                sSQL = sSQL & " SELECT @CodigoNaturezaOperacao = 0 "
                sSQL = sSQL & " SELECT @ValorTrocaItem = 0"
                sSQL = sSQL & " SELECT @QuantidadeTroca = 0"
                sSQL = sSQL & " SELECT @ValorBonificacaoItem = 0"
                sSQL = sSQL & " SELECT @CodigoUnidade = " & rsGrid!CodigoUnidade 'ele passa o parâmetro 'um' para o Gravaitem"
                sSQL = sSQL & " SELECT @FatorConversao = 0"
                sSQL = sSQL & " SELECT @Controle = ''"
                sSQL = sSQL & " SELECT @CodigoDeposito = " & CodigoDepositoTIM
                sSQL = sSQL & " SELECT @ValorLiquido = " & lblvlrfaturarIPI.Caption
                sSQL = sSQL & " EXEC spGravaITEM @CodigoEmpresa, @Sequencia, @SequenciaProduto, @CodigoProduto, @Quantidade, @Desconto, @ValorUnitario, @CodigoVendedor, @AliquotaIPI, @AliquotaICMS, @AliquotaReducao, @IVA, @ICMSSubstituicao, @CodigoClassificacaoFiscal, 1, @CodigoNaturezaOperacao, @ValorTrocaItem, @ValorBonificacaoItem, @QuantidadeTroca, @CodigoUnidade, @FatorConversao, @CodigoDeposito, @Controle, ValorLiquido"
                db.Execute sSQL
    'Rotina Caso O Item Seja o Mesmo
pulaDTeste:
                'Rotina para zerar quantidade a Faturar
                db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                
                'Rotina Para Relacionar Itens da Devolucao na Tabela CapaItemRelacionaCapaItem
                db.Execute "INSERT INTO CapaItemRelacionaCapaItem (CodigoEmpresa, Sequencia, SequenciaProduto, SequenciaOriginal, SequenciaProdutoOriginal, CodigoProduto, CodigoUnidade) VALUES (" & intEmpresa & "," & intSequencia & "," & intSequenciaItem & "," & rsGrid!Sequencia & "," & rsGrid!SequenciaOriginal & "," & rsGrid!codigo & "," & rsGrid!CodigoUnidade & ")"
                
                Complemento 'Chama uma rotina que grava numero do pedido e observacoes no Itens quando faz o agrupamento de produtos, Nao foi possivel implementar a rotina aqio pois estava dando um erro de procedure too large ou seja estorou o tamanho da procedure Emanoel
                
            End If
            
            IntNotaAnterior = rsGrid!Nota
            SequenciaAnterior = intSequencia
            
            sSQL = "EXEC spProcessaConsignacao  1,0,'',''," & rsGrid("CodigoCliente") & ",0,0,0,0," & tpLNG(rsGrid("Codigo")) & ",'D',60,'A','C'"
            db.Execute sSQL
            
            rsGrid.MoveNext
        Loop
        
        strMsgNota = "Devolucao Teste Ref. NFs: "
        strMsgNota = strMsgNota & strNumeroNota
        db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
        strNumeroNota = 0
        
        db.Execute "exec spAtualizaItem " & intEmpresa & ", '" & intSequencia & "','Capa'"
        db.Execute "exec spAtualizaCapa " & intEmpresa & ", '" & intSequencia & "','Capa'"
        
    End If
    
    If Devolver = True Then
        intSequencia = 0
        intSequenciaItem = 0
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            DoEvents:
            If rsGrid("QtdeADevolver") > 0 Then
               intSequenciaItem = intSequenciaItem + 1
               varsql = "SELECT  Capa.CodigoEmitente, Capa.Quantidade, Capa.ValorTotal, Capa.ValorLiquido "
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE Capa.CodigoEmpresa=" & intEmpresa & " AND Capa.CodigoEmitente = " & rsGrid("CodigoCliente") & " AND Capa.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                     'Acumulando Valores Caso seja mesma Nota
                    qtdDevolver = qtdDevolver + rsGrid!QtdeADevolver
                    VlrUnitarioDevolucao = rsGrid!VlSaida / rsGrid!qtdsaida
                    VlrDevolver = VlrDevolver + Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
                    'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                       End If
                    End If
                    
                    VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
                    lblvlrfaturar1.Caption = Replace(VlrDevolver, ",", ".")
                    lblvlrfaturarIPI.Caption = Replace(VlrDevolverIPI, ",", ".")
                    lblIPI.Caption = Replace(VlrIPI, ",", ".")
                    
                    If rsGrid!AliquotaIPI > 0 Then
                       VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                       VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                       BaseICMS = BaseICMS + lblvlrfaturar1.Caption
                    End If
               
                    lblICMS.Caption = Replace(VlrICMS, ",", ".")
                    lblBICMS.Caption = Replace(BaseICMS, ",", ".")
                   
                   'Atualizando Cabeca do pedido
                   db.Execute "UPDATE Capa SET Quantidade = " & qtdDevolver & ", ValorTotal = " & lblvlrfaturarIPI.Caption & ", ValorLiquido = " & lblvlrfaturar1.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
                   rs.Close
                    
                    GoTo PulaClienteDConfig
               End If
               rs.Close
                   
                'Pegando o Deposito e o Tipo de Estoque da Tim
                sSQL = "SELECT TIM.TIMContraPartida"
                sSQL = sSQL & " FROM Capa INNER JOIN TIM ON Capa.CodigoEmpresa = TIM.CodigoEmpresa AND Capa.CodigoTIM = TIM.Codigo"
                sSQL = sSQL & " WHERE Capa.CodigoEmpresa = " & intEmpresa & " AND Sequencia = " & rsGrid("Sequencia")
                rs.Open sSQL, db, , , adCmdText
                If Not rs.EOF Then
                    IntTimConsigDevolve = rs!TIMContrapartida
                End If
                rs.Close
                 
                sSQL = "SELECT CodigoDeposito, Estoque "
                sSQL = sSQL & "From TIM "
                sSQL = sSQL & "Where Codigo = " & IntTimConsigDevolve
                rs.Open sSQL, db, , , adCmdText
                If Not rs.EOF Then
                    If strRetorno = "" Then
                        CodigoDepositoTIM = rs!CodigoDeposito
                    Else
                        CodigoDepositoTIM = Str(Left(strRetorno, 1))
                    End If
                    
                    If rs!Estoque = 0 Then
                        EstoqueTIM = "N" 'Nenhum
                    ElseIf rs!Estoque = 1 Then
                        EstoqueTIM = "S" 'Saida
                    ElseIf rs!Estoque = 2 Then
                        EstoqueTIM = "E" 'Entrada
                    ElseIf rs!Estoque = 3 Then
                        EstoqueTIM = "R" 'Reserva
                    End If
                End If
                rs.Close
               
               'Calculando Valor a Faturar
               qtdDevolver = rsGrid("qtdeAdevolver")
               VlrUnitarioDevolucao = rsGrid!VlSaida / rsGrid!qtdsaida
               VlrDevolver = Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
               'Pegando Valor IPI de Acordo com o Percentual do Produto
               If booCalculaIPI = True Then
                  If rsGrid!AliquotaIPI > 0 Then
                     VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                     VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                     VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
                  End If
               End If
               
               VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
               lblvlrfaturar1.Caption = Replace(VlrDevolver, ",", ".")
               lblvlrfaturarIPI.Caption = Replace(VlrDevolverIPI, ",", ".")
               lblIPI.Caption = Replace(VlrIPI, ",", ".")
               
               If rsGrid!AliquotaICMS > 0 Then
                  VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                  VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                  BaseICMS = lblVlrDevolver.Caption
               End If
               lblICMS.Caption = Replace(VlrICMS, ",", ".")
               lblBICMS.Caption = Replace(BaseICMS, ",", ".")
                  
               intSequencia = BuscaCodigo("SELECT isnull(max(Sequencia),0) + 1 as Total FROM Capa WHERE CodigoEmpresa = " & intEmpresa, db, "Total")
               
               'Copiando Remessa Consignacao mudando algumas particularidades como sequencia e valores e tim devolucao
               varsql = "INSERT INTO Capa (CodigoEmpresa, Sequencia, CodigoFilial, Tipo, CodigoSituacao, CodigoVendedor, CodigoTabelaPreco, CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, Requisicao, Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia, Quantidade, ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca, ValorLiquido, ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor,  ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, TipoFrete, BaseICMS, ValorICMS, BaseIPI, ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, "
               varsql = varsql & " DataCadastro, DataFaturamento, DataEntradaSaida, DataContabilizacao, DataMovimentacao, CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, CIFFOB, CodigoCentroCusto, CodigoNatureza, CodigoNaturezaOperacao, Confirmada, CodigoMotivo, AtualizaFinanceiro, AtualizaEstoque, CodigoRequisitante, CodigoAutorizacao, CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, DataCancelamento, CodigoUsuarioCancelamento, DataAlteracao, CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, EmiteNF, EmiteCupom, Acertado, CodigoContaCorrente, "
               varsql = varsql & " TipoVenda, TipoComissao, PercentualComissao, Placa, UFPlaca) "
               varsql = varsql & " SELECT CodigoEmpresa, " & intSequencia & " AS Sequencia, CodigoFilial, '" & strTipoMovimento & "' as Tipo, 9 as CodigoSituacao, CodigoVendedor, CodigoTabelaPreco," & IntTimConsigDevolve & " AS CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, " & intSequencia & " as Requisicao, " & 0 & " AS Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia," & rsGrid("QtdeaDevolver") & " as  Quantidade," & lblvlrfaturarIPI.Caption & " as ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca," & lblvlrfaturar1.Caption & " as ValorLiquido," & 0 & " as ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor, ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, "
               varsql = varsql & " TipoFrete, " & lblBICMS.Caption & " as  BaseICMS, " & lblICMS.Caption & " as  ValorICMS, " & 0 & " as  BaseIPI, " & lblIPI.Caption & " as  ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, getdate() as DataCadastro, '" & strDataBase & "' as DataFaturamento, '" & strDataBase & "' as DataEntradaSaida, '" & strDataBase & "' as DataContabilizacao, '" & strDataBase & "' as DataMovimentacao, " & intUsuario & " as CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, '" & "C" & "' as CIFFOB, CodigoCentroCusto, CodigoNatureza, " & 0 & " as CodigoNaturezaOperacao, Confirmada, CodigoMotivo, '" & "N" & "' AS AtualizaFinanceiro, '" & EstoqueTIM & "' AS AtualizaEstoque, "
               varsql = varsql & " CodigoRequisitante, CodigoAutorizacao, " & 0 & " as CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, null as DataCancelamento, null as CodigoUsuarioCancelamento, null as DataAlteracao, null as CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, 0 as EmiteNF, 0 as EmiteCupom, Acertado, CodigoContaCorrente, TipoVenda, TipoComissao, PercentualComissao, '" & 0 & "' as Placa, UFPlaca"
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & rsGrid("Sequencia")
               db.Execute varsql
               
               'Verificando se tem alguma coisa e deleta antes de iniciar processo
               db.Execute "DELETE FROM CapaItem WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
               
    'Rotina Caso O Cliente Seja O mesmo
PulaClienteDConfig:
               If IntNotaAnterior <> rsGrid!Nota Then
                   If SequenciaAnterior > 0 Then
                        If intSequencia = SequenciaAnterior Then
                           If strNumeroNota <> "" Then
                              If Len(strNumeroNota) <= 500 Then
                                 strNumeroNota = strNumeroNota & "," & rsGrid!Nota
                              End If
                              Else
                                  strNumeroNota = rsGrid!Nota
                           End If
                        End If
                        Else
                           strNumeroNota = rsGrid!Nota
                   End If
               End If
               If SequenciaAnterior > 0 Then
                    If intSequencia <> SequenciaAnterior Then
                       strMsgNota = "Devolucao Consignação Ref. NFs: "
                       strMsgNota = strMsgNota & strNumeroNota
                       db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & SequenciaAnterior
                       strNumeroNota = rsGrid!Nota
                    End If
               End If
    
               'Relacionado Capa ao caparelacinaCapa
               varsql = "SELECT  CapaRelacionaCapa.Sequencia, CapaRelacionaCapa.Sequencia "
               varsql = varsql & " FROM CapaRelacionaCapa "
               varsql = varsql & " WHERE CapaRelacionaCapa.CodigoEmpresa=" & intEmpresa & " AND CapaRelacionaCapa.sequencia = " & intSequencia & " AND CapaRelacionaCapa.SequenciaOriginal = " & rsGrid("Sequencia")
               rs.Open varsql, db, , , adCmdText
               If rs.EOF Then
                  db.Execute "INSERT INTO CapaRelacionaCapa (CodigoEmpresa, Sequencia, SequenciaOriginal) VALUES (" & intEmpresa & "," & intSequencia & "," & rsGrid("Sequencia") & ")"
               End If
               rs.Close
               
               varsql = "SELECT  Capaitem.CodigoProduto, Capaitem.Quantidade, Capaitem.ValorUnitario, Capaitem.ValorLiquido "
               varsql = varsql & " FROM Capaitem "
               varsql = varsql & " WHERE Capaitem.CodigoEmpresa=" & intEmpresa & " AND Capaitem.CodigoProduto = " & rsGrid("Codigo") & " AND Capaitem.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                   'Acumulando Valores Caso seja o Mesmo Produto mesmo Cliente
                   QtdFaturaItem = rs!Quantidade + rsGrid("QtdeADevolver")
                   'Pegando Valor unitario do que vem para atualizar
                   lblvlrUniItem.Caption = (rsGrid("VlSaida") / rsGrid("qtdSaida"))
                   'Calculando o valor unitario com a quantidade solicitada
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption * rsGrid("QtdeADevolver")
                   'Calculando total Quantidade anterior com quantidade atual
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption + (rs!Quantidade * rs!ValorUnitario)
                   'Por Fim pegando o valor Unitario pela divisao das quantidades e valores
                   lblvlrUniItem.Caption = (lblvlrUniItem.Caption / QtdFaturaItem)
                   lblvlrLiqItem.Caption = lblvlrUniItem.Caption * QtdFaturaItem
                   
                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((lblvlrUniItem.Caption * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid("QtdeADevolver"), 2)
                       End If
                    End If
                    lblvlrTotItem.Caption = Arredonda(lblvlrLiqItem.Caption + VlrIPI, 2)
                   
                    lblvlrUniItem.Caption = Replace(lblvlrUniItem.Caption, ",", ".")
                    lblvlrLiqItem.Caption = Replace(lblvlrLiqItem.Caption, ",", ".")
                    lblvlrTotItem.Caption = Replace(lblvlrTotItem.Caption, ",", ".")
                   'Atualizando Itens  do pedido
                   db.Execute "UPDATE CapaItem SET Quantidade = " & QtdFaturaItem & ", ValorUnitario = " & lblvlrUniItem.Caption & ", ValorLiquido = " & lblvlrLiqItem.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo pulaDConsig
               End If
               rs.Close
                
                intSequenciaProduto = intSequenciaProduto + 1
                sSQL = " DECLARE @CodigoEmpresa tinyint"
                sSQL = sSQL & " DECLARE @Sequencia int"
                sSQL = sSQL & " DECLARE @SequenciaProduto tinyint"
                sSQL = sSQL & " DECLARE @CodigoProduto int"
                sSQL = sSQL & " DECLARE @Quantidade money"
                sSQL = sSQL & " DECLARE @Desconto decimal(19,4)"
                sSQL = sSQL & " DECLARE @ValorUnitario decimal(19,4)"
                sSQL = sSQL & " DECLARE @CodigoVendedor smallint"
                sSQL = sSQL & " DECLARE @AliquotaIPI money"
                sSQL = sSQL & " DECLARE @AliquotaICMS money"
                sSQL = sSQL & " DECLARE @AliquotaReducao money"
                sSQL = sSQL & " DECLARE @IVA money"
                sSQL = sSQL & " DECLARE @ICMSSubstituicao money"
                sSQL = sSQL & " DECLARE @CodigoClassificacaoFiscal money"
                sSQL = sSQL & " DECLARE @CodigoNaturezaOperacao int"
                sSQL = sSQL & " DECLARE @ValorTrocaItem money"
                sSQL = sSQL & " DECLARE @ValorBonificacaoItem money"
                sSQL = sSQL & " DECLARE @QuantidadeTroca money"
                sSQL = sSQL & " DECLARE @CodigoUnidade money"
                sSQL = sSQL & " DECLARE @FatorConversao smallint"
                sSQL = sSQL & " DECLARE @Controle varchar(20)"
                sSQL = sSQL & " DECLARE @CodigoDeposito smallint"
                sSQL = sSQL & " DECLARE @QuantidadeAvarias smallint"
                sSQL = sSQL & " DECLARE @ValorLiquido decimal(19,4)"
                
                sSQL = sSQL & " SELECT @CodigoEmpresa = " & intEmpresa
                sSQL = sSQL & " SELECT @Sequencia = " & intSequencia
                sSQL = sSQL & " SELECT @SequenciaProduto = " & 0
                sSQL = sSQL & " SELECT @CodigoProduto = " & tpLNG(rsGrid("Codigo"))
                sSQL = sSQL & " SELECT @Quantidade = " & tpMOEDA(rsGrid("qtdeADevolver"))
                sSQL = sSQL & " SELECT @Desconto = " & 0
                sSQL = sSQL & " SELECT @ValorUnitario = " & tpMOEDA(rsGrid("VlSaida") / rsGrid("qtdSaida"))
                sSQL = sSQL & " SELECT @CodigoVendedor = " & 0
                sSQL = sSQL & " SELECT @AliquotaIPI = " & 0
                sSQL = sSQL & " SELECT @AliquotaICMS = " & 0
                sSQL = sSQL & " SELECT @AliquotaReducao = " & 0
                sSQL = sSQL & " SELECT @IVA = " & 0
                sSQL = sSQL & " SELECT @ICMSSubstituicao = " & 0
                sSQL = sSQL & " SELECT @CodigoClassificacaoFiscal = " & 0
                sSQL = sSQL & " SELECT @CodigoNaturezaOperacao = 0 "
                sSQL = sSQL & " SELECT @ValorTrocaItem = 0"
                sSQL = sSQL & " SELECT @QuantidadeTroca = 0"
                sSQL = sSQL & " SELECT @ValorBonificacaoItem = 0"
                sSQL = sSQL & " SELECT @CodigoUnidade = " & rsGrid!CodigoUnidade
                sSQL = sSQL & " SELECT @FatorConversao = 0"
                sSQL = sSQL & " SELECT @Controle = ''"
                sSQL = sSQL & " SELECT @CodigoDeposito = " & CodigoDepositoTIM
                sSQL = sSQL & " SELECT @QuantidadeAvarias = 0" 'Faturamento ou seja a venda
                sSQL = sSQL & " SELECT @ValorLiquido = " & lblvlrfaturarIPI.Caption
                sSQL = sSQL & " EXEC spGravaITEM @CodigoEmpresa, @Sequencia, @SequenciaProduto, @CodigoProduto, @Quantidade, @Desconto, @ValorUnitario, @CodigoVendedor, @AliquotaIPI, @AliquotaICMS, @AliquotaReducao, @IVA, @ICMSSubstituicao, @CodigoClassificacaoFiscal, 1, @CodigoNaturezaOperacao, @ValorTrocaItem, @ValorBonificacaoItem, @QuantidadeTroca, @CodigoUnidade, @FatorConversao, @CodigoDeposito, @Controle"
                db.Execute sSQL
pulaDConsig:
                'Rotina para zerar quantidade a Faturar
                db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                'Rotina Para Relacionar Itens da Devolucao na Tabela CapaItemRelacionaCapaItem
                db.Execute "INSERT INTO CapaItemRelacionaCapaItem (CodigoEmpresa, Sequencia, SequenciaProduto, SequenciaOriginal, SequenciaProdutoOriginal, CodigoProduto, CodigoUnidade) VALUES (" & intEmpresa & "," & intSequencia & "," & intSequenciaItem & "," & rsGrid!Sequencia & "," & rsGrid!SequenciaOriginal & "," & rsGrid!codigo & "," & rsGrid!CodigoUnidade & ")"
                
                Complemento 'Chama uma rotina que grava numero do pedido e observacoes no Itens quando faz o agrupamento de produtos, Nao foi possivel implementar a rotina aqio pois estava dando um erro de procedure too large ou seja estorou o tamanho da procedure Emanoel
                
            End If
            
            sSQL = "EXEC spProcessaConsignacao  1,0,'',''," & rsGrid("CodigoCliente") & ",0,0,0,0," & tpLNG(rsGrid("Codigo")) & ",'D',60,'A','C'"
            db.Execute sSQL
            
            rsGrid.MoveNext
        Loop
        
        strMsgNota = "Devolucao Consignação Ref. NFs: "
        strMsgNota = strMsgNota & strNumeroNota
        db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
        strNumeroNota = 0
        
        db.Execute "exec spAtualizaItem " & intEmpresa & ", '" & intSequencia & "','Capa'"
        db.Execute "exec spAtualizaCapa " & intEmpresa & ", '" & intSequencia & "','Capa'"
    End If
    
    If DevolverEmprestimo = True Then
        intSequencia = 0
        intSequenciaItem = 0
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            DoEvents:
            If rsGrid("QtdeADevolver") > 0 Then
               intSequenciaItem = intSequenciaItem + 1
               varsql = "SELECT  Capa.CodigoEmitente, Capa.Quantidade, Capa.ValorTotal, Capa.ValorLiquido "
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE Capa.CodigoEmpresa=" & intEmpresa & " AND Capa.CodigoEmitente = " & rsGrid("CodigoCliente") & " AND Capa.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                     'Acumulando Valores Caso seja mesma Nota
                    qtdDevolver = qtdDevolver + rsGrid!QtdeADevolver
                    VlrUnitarioDevolucao = rsGrid!VlSaida / rsGrid!qtdsaida
                    VlrDevolver = VlrDevolver + Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
                    'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                       End If
                    End If
                    VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
                    lblvlrfaturar1.Caption = Replace(VlrDevolver, ",", ".")
                    lblvlrfaturarIPI.Caption = Replace(VlrDevolverIPI, ",", ".")
                    lblIPI.Caption = Replace(VlrIPI, ",", ".")
    
                    If rsGrid!AliquotaIPI > 0 Then
                       VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                       VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                       BaseICMS = BaseICMS + lblvlrfaturar1.Caption
                    End If
                    lblICMS.Caption = Replace(VlrICMS, ",", ".")
                    lblBICMS.Caption = Replace(BaseICMS, ",", ".")
    
                   'Atualizando Cabeca do pedido
                   db.Execute "UPDATE Capa SET Quantidade = " & qtdDevolver & ", ValorTotal = " & lblvlrfaturarIPI.Caption & ", ValorLiquido = " & lblvlrfaturar1.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
                   rs.Close
                    
                    GoTo PulaClienteDEmp
               End If
               rs.Close
               
               'Pegando o Deposito e o Tipo de Estoque da Tim
                sSQL = "SELECT TIM.TIMContraPartida"
                sSQL = sSQL & " FROM Capa INNER JOIN TIM ON Capa.CodigoEmpresa = TIM.CodigoEmpresa AND Capa.CodigoTIM = TIM.Codigo"
                sSQL = sSQL & " WHERE Capa.CodigoEmpresa = " & intEmpresa & " AND Sequencia = " & rsGrid("Sequencia")
                rs.Open sSQL, db, , , adCmdText
                If Not rs.EOF Then
                    IntTimEmprestimoDevolve = rs!TIMContrapartida
                End If
                rs.Close
        
                sSQL = "SELECT CodigoDeposito, Estoque "
                sSQL = sSQL & "From TIM "
                sSQL = sSQL & "Where Codigo = " & IntTimEmprestimoDevolve
                rs.Open sSQL, db, , , adCmdText
                If Not rs.EOF Then
                    If strRetorno = "" Then
                       CodigoDepositoTIM = rs!CodigoDeposito
                       Else
                          CodigoDepositoTIM = Str(Left(strRetorno, 1))
                    End If
                    
                    If rs!Estoque = 0 Then
                       EstoqueTIM = "N" 'Nenhum
                    ElseIf rs!Estoque = 1 Then
                       EstoqueTIM = "S" 'Saida
                    ElseIf rs!Estoque = 2 Then
                       EstoqueTIM = "E" 'Entrada
                    ElseIf rs!Estoque = 3 Then
                       EstoqueTIM = "R" 'Reserva
                    End If
                End If
                rs.Close
    
               'Calculando Valor a Faturar
               qtdDevolver = rsGrid("qtdeAdevolver")
               VlrUnitarioDevolucao = rsGrid!VlSaida / rsGrid!qtdsaida
               VlrDevolver = Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
               'Pegando Valor IPI de Acordo com o Percentual do Produto
               If booCalculaIPI = True Then
                  If rsGrid!AliquotaIPI > 0 Then
                     VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                     VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                     VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
                  End If
               End If
               VlrDevolverIPI = Arredonda(VlrDevolver + VlrIPI, 2)
               lblvlrfaturar1.Caption = Replace(VlrDevolver, ",", ".")
               lblvlrfaturarIPI.Caption = Replace(VlrDevolverIPI, ",", ".")
               lblIPI.Caption = Replace(VlrIPI, ",", ".")
    
               If rsGrid!AliquotaICMS > 0 Then
                  VlrICMS = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaICMS) / 100), 2)
                  VlrICMS = Arredonda(VlrICMS * rsGrid!QtdeADevolver, 2)
                  BaseICMS = lblVlrDevolver.Caption
               End If
               lblICMS.Caption = Replace(VlrICMS, ",", ".")
               lblBICMS.Caption = Replace(BaseICMS, ",", ".")
    
               intSequencia = BuscaCodigo("SELECT isnull(max(Sequencia),0) + 1 as Total FROM Capa WHERE CodigoEmpresa = " & intEmpresa, db, "Total")
               'Copiando Remessa Consignacao mudando algumas particularidades como sequencia e valores e tim devolucao
               varsql = "INSERT INTO Capa (CodigoEmpresa, Sequencia, CodigoFilial, Tipo, CodigoSituacao, CodigoVendedor, CodigoTabelaPreco, CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, Requisicao, Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia, Quantidade, ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca, ValorLiquido, ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor,  ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, TipoFrete, BaseICMS, ValorICMS, BaseIPI, ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, "
               varsql = varsql & " DataCadastro, DataFaturamento, DataEntradaSaida, DataContabilizacao, DataMovimentacao, CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, CIFFOB, CodigoCentroCusto, CodigoNatureza, CodigoNaturezaOperacao, Confirmada, CodigoMotivo, AtualizaFinanceiro, AtualizaEstoque, CodigoRequisitante, CodigoAutorizacao, CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, DataCancelamento, CodigoUsuarioCancelamento, DataAlteracao, CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, EmiteNF, EmiteCupom, Acertado, CodigoContaCorrente, "
               varsql = varsql & " TipoVenda, TipoComissao, PercentualComissao, Placa, UFPlaca) "
               varsql = varsql & " SELECT CodigoEmpresa, " & intSequencia & " AS Sequencia, CodigoFilial, '" & strTipoMovimento & "' as Tipo, 9 as CodigoSituacao, CodigoVendedor, CodigoTabelaPreco," & IntTimEmprestimoDevolve & " AS CodigoTIM, PercentualDescontoTabela, CodigoEmitente, PercentualDescontoEmitente, CodigoTipoNegociacao, PercentualDescontoTipoNegociacao, CodigoFormaPagamento, " & intSequencia & " as Requisicao, " & 0 & " AS Nota, Serie, OrdemCarga, SequenciaCarga, CodigoUrgencia," & rsGrid("QtdeaDevolver") & " as  Quantidade," & lblvlrfaturarIPI.Caption & " as ValorTotal, ValorDescontoTabela, ValorDescontoCondicao, ValorDescontoEmitente, ValorDescontoItem, ValorBonificacao, ValorTroca," & lblvlrfaturar1.Caption & " as ValorLiquido," & 0 & " as ValorFaturado, CMV, CustoMedio, ValorComissaoVendedor, ValorComissaoEmpresa, ValorSeguro, ICMSSeguro, TipoIPIEmbalagem, IPIEmbalagem, ValorEmbalagem, ICMSEmbalagem, BaseFrete, ValorFrete, ICMSFrete, DataVencimentoFrete, "
               varsql = varsql & " TipoFrete, " & lblBICMS.Caption & " as  BaseICMS, " & lblICMS.Caption & " as  ValorICMS, " & 0 & " as  BaseIPI, " & lblIPI.Caption & " as  ValorIPI, BaseICMSSubstituicao, ValorICMSSubstituicao, BaseISS, ValorISS, ISSRetido, BaseINSS, ValorINSS, ValorIRF, IRFRetido, ValorEncargos, ValorOutras, ValorJuro, PesoBruto, PesoLiquido, Volume, ObservacaoFaturamento, ObservacaoNota, getdate() as DataCadastro, '" & strDataBase & "' as DataFaturamento, '" & strDataBase & "' as DataEntradaSaida, '" & strDataBase & "' as DataContabilizacao, '" & strDataBase & "' as DataMovimentacao, " & intUsuario & " as CodigoUsuario, CodigoTransportador, CodigoRota, CodigoSubRota, CodigoSupervisor, CodigoRegiao, CodigoMoeda, CodigoEntregador, '" & "C" & "' as CIFFOB, CodigoCentroCusto, CodigoNatureza, " & 0 & " as CodigoNaturezaOperacao, Confirmada, CodigoMotivo, '" & "N" & "' AS AtualizaFinanceiro, '" & EstoqueTIM & "' AS AtualizaEstoque, "
               varsql = varsql & " CodigoRequisitante, CodigoAutorizacao, " & 0 & " as CodigoVeiculo, Lacre, TipoIndenizacao, TipoBonificacao, CodigoContrato, CodigoRemetente, CodigoDestinatario, CodigoDespacho, LocalColeta, LocalEntrega, CodigoFilialNegociacao, SequenciaAnterior, null as DataCancelamento, null as CodigoUsuarioCancelamento, null as DataAlteracao, null as CodigoUsuarioAlteracao, DataEntrega, DataValidade, ValorEntrada, TaxaFinanciamento, QuantidadeParcelas, ValorParcelas, ValorReceber, LicitacaoTipo, LicitacaoNumero, LicitacaoData, LicitacaoHora, LicitacaoReferencia, 0 as EmiteNF, 0 as EmiteCupom, Acertado, CodigoContaCorrente, TipoVenda, TipoComissao, PercentualComissao, '" & 0 & "' as Placa, UFPlaca"
               varsql = varsql & " FROM Capa "
               varsql = varsql & " WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & rsGrid("Sequencia")
               db.Execute varsql
    
    
               'Verificando se tem alguma coisa e deleta antes de iniciar processo
               db.Execute "DELETE FROM CapaItem WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
    
    'Rotina Caso O Cliente Seja O mesmo
PulaClienteDEmp:
               If IntNotaAnterior <> rsGrid!Nota Then
                   If SequenciaAnterior > 0 Then
                        If intSequencia = SequenciaAnterior Then
                           If strNumeroNota <> "" Then
                                 If Len(strNumeroNota) <= 500 Then
                                    strNumeroNota = strNumeroNota & "," & rsGrid!Nota
                                 End If
                              Else
                                  strNumeroNota = rsGrid!Nota
                           End If
                        End If
                        Else
                           strNumeroNota = rsGrid!Nota
                   End If
               End If
               If SequenciaAnterior > 0 Then
                    If intSequencia <> SequenciaAnterior Then
                       strMsgNota = "Devolucao Emprestimo Ref. NFs: "
                       strMsgNota = strMsgNota & strNumeroNota
                       db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & SequenciaAnterior
                       strNumeroNota = rsGrid!Nota
                    End If
               End If
    
               'Relacionado Capa ao caparelacinaCapa
               varsql = "SELECT  CapaRelacionaCapa.Sequencia, CapaRelacionaCapa.Sequencia "
               varsql = varsql & " FROM CapaRelacionaCapa "
               varsql = varsql & " WHERE CapaRelacionaCapa.CodigoEmpresa=" & intEmpresa & " AND CapaRelacionaCapa.sequencia = " & intSequencia & " AND CapaRelacionaCapa.SequenciaOriginal = " & rsGrid("Sequencia")
               rs.Open varsql, db, , , adCmdText
               If rs.EOF Then
                  db.Execute "INSERT INTO CapaRelacionaCapa (CodigoEmpresa, Sequencia, SequenciaOriginal) VALUES (" & intEmpresa & "," & intSequencia & "," & rsGrid("Sequencia") & ")"
               End If
               rs.Close
    
               varsql = "SELECT  Capaitem.CodigoProduto, Capaitem.Quantidade, Capaitem.ValorUnitario, Capaitem.ValorLiquido "
               varsql = varsql & " FROM Capaitem "
               varsql = varsql & " WHERE Capaitem.CodigoEmpresa=" & intEmpresa & " AND Capaitem.CodigoProduto = " & rsGrid("Codigo") & " AND Capaitem.Sequencia = " & intSequencia
               rs.Open varsql, db, , , adCmdText
               If Not rs.EOF Then
                  'Acumulando Valores Caso seja o Mesmo Produto mesmo Cliente
                   QtdFaturaItem = rs!Quantidade + rsGrid("QtdeADevolver")
                   'Pegando Valor unitario do que vem para atualizar
                   lblvlrUniItem.Caption = (rsGrid("VlSaida") / rsGrid("qtdSaida"))
                   'Calculando o valor unitario com a quantidade solicitada
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption * rsGrid("QtdeADevolver")
                   'Calculando total Quantidade anterior com quantidade atual
                   lblvlrUniItem.Caption = lblvlrUniItem.Caption + (rs!Quantidade * rs!ValorUnitario)
                   'Por Fim pegando o valor Unitario pela divisao das quantidades e valores
                   lblvlrUniItem.Caption = (lblvlrUniItem.Caption / QtdFaturaItem)
                   lblvlrLiqItem.Caption = lblvlrUniItem.Caption * QtdFaturaItem
    
                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                    If booCalculaIPI = True Then
                       If rsGrid!AliquotaIPI > 0 Then
                          VlrIPI = Arredonda(((lblvlrUniItem.Caption * rsGrid!AliquotaIPI) / 100), 2)
                          VlrIPI = Arredonda(VlrIPI * rsGrid("QtdeADevolver"), 2)
                       End If
                    End If
                    lblvlrTotItem.Caption = Arredonda(lblvlrLiqItem.Caption + VlrIPI, 2)
    
                    lblvlrUniItem.Caption = Replace(lblvlrUniItem.Caption, ",", ".")
                    lblvlrLiqItem.Caption = Replace(lblvlrLiqItem.Caption, ",", ".")
                    lblvlrTotItem.Caption = Replace(lblvlrTotItem.Caption, ",", ".")
                   'Atualizando Itens  do pedido
                   db.Execute "UPDATE CapaItem SET Quantidade = " & QtdFaturaItem & ", ValorUnitario = " & lblvlrUniItem.Caption & ", ValorLiquido = " & lblvlrLiqItem.Caption & "  WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & intSequencia
                   rs.Close
                   GoTo pulaDEmp
               End If
               rs.Close
                intSequenciaProduto = intSequenciaProduto + 1
                sSQL = " DECLARE @CodigoEmpresa tinyint"
                sSQL = sSQL & " DECLARE @Sequencia int"
                sSQL = sSQL & " DECLARE @SequenciaProduto tinyint"
                sSQL = sSQL & " DECLARE @CodigoProduto int"
                sSQL = sSQL & " DECLARE @Quantidade money"
                sSQL = sSQL & " DECLARE @Desconto decimal(19,4)"
                sSQL = sSQL & " DECLARE @ValorUnitario decimal(19,4)"
                sSQL = sSQL & " DECLARE @CodigoVendedor smallint"
                sSQL = sSQL & " DECLARE @AliquotaIPI money"
                sSQL = sSQL & " DECLARE @AliquotaICMS money"
                sSQL = sSQL & " DECLARE @AliquotaReducao money"
                sSQL = sSQL & " DECLARE @IVA money"
                sSQL = sSQL & " DECLARE @ICMSSubstituicao money"
                sSQL = sSQL & " DECLARE @CodigoClassificacaoFiscal money"
                sSQL = sSQL & " DECLARE @CodigoNaturezaOperacao int"
                sSQL = sSQL & " DECLARE @ValorTrocaItem money"
                sSQL = sSQL & " DECLARE @ValorBonificacaoItem money"
                sSQL = sSQL & " DECLARE @QuantidadeTroca money"
                sSQL = sSQL & " DECLARE @CodigoUnidade money"
                sSQL = sSQL & " DECLARE @FatorConversao smallint"
                sSQL = sSQL & " DECLARE @Controle varchar(20)"
                sSQL = sSQL & " DECLARE @CodigoDeposito smallint"
                sSQL = sSQL & " DECLARE @QuantidadeAvarias smallint"
                sSQL = sSQL & " DECLARE @ValorLiquido decimal(19,4)"
    
                sSQL = sSQL & " SELECT @CodigoEmpresa = " & intEmpresa
                sSQL = sSQL & " SELECT @Sequencia = " & intSequencia
                sSQL = sSQL & " SELECT @SequenciaProduto = " & 0
                sSQL = sSQL & " SELECT @CodigoProduto = " & tpLNG(rsGrid("Codigo"))
                sSQL = sSQL & " SELECT @Quantidade = " & tpMOEDA(rsGrid("qtdeADevolver"))
                sSQL = sSQL & " SELECT @Desconto = " & 0
                sSQL = sSQL & " SELECT @ValorUnitario = " & Str(rsGrid("VlSaida") / rsGrid("qtdSaida"))
                sSQL = sSQL & " SELECT @CodigoVendedor = " & 0
                sSQL = sSQL & " SELECT @AliquotaIPI = " & 0
                sSQL = sSQL & " SELECT @AliquotaICMS = " & 0
                sSQL = sSQL & " SELECT @AliquotaReducao = " & 0
                sSQL = sSQL & " SELECT @IVA = " & 0
                sSQL = sSQL & " SELECT @ICMSSubstituicao = " & 0
                sSQL = sSQL & " SELECT @CodigoClassificacaoFiscal = " & 0
                sSQL = sSQL & " SELECT @CodigoNaturezaOperacao = 0 "
                sSQL = sSQL & " SELECT @ValorTrocaItem = 0"
                sSQL = sSQL & " SELECT @QuantidadeTroca = 0"
                sSQL = sSQL & " SELECT @ValorBonificacaoItem = 0"
                sSQL = sSQL & " SELECT @CodigoUnidade = " & rsGrid!CodigoUnidade
                sSQL = sSQL & " SELECT @FatorConversao = 0"
                sSQL = sSQL & " SELECT @Controle = ''"
                sSQL = sSQL & " SELECT @CodigoDeposito = " & CodigoDepositoTIM
                sSQL = sSQL & " SELECT @QuantidadeAvarias = 0" 'Faturamento ou seja a venda
                sSQL = sSQL & " SELECT @ValorLiquido = " & lblvlrfaturarIPI.Caption
                sSQL = sSQL & " EXEC spGravaITEM @CodigoEmpresa, @Sequencia, @SequenciaProduto, @CodigoProduto, @Quantidade, @Desconto, @ValorUnitario, @CodigoVendedor, @AliquotaIPI, @AliquotaICMS, @AliquotaReducao, @IVA, @ICMSSubstituicao, @CodigoClassificacaoFiscal, 1, @CodigoNaturezaOperacao, @ValorTrocaItem, @ValorBonificacaoItem, @QuantidadeTroca, @CodigoUnidade, @FatorConversao, @CodigoDeposito, @Controle"
                db.Execute sSQL
pulaDEmp:
                'Rotina para zerar quantidade a Faturar
                db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                'Rotina Para Relacionar Itens da Devolucao na Tabela CapaItemRelacionaCapaItem
                db.Execute "INSERT INTO CapaItemRelacionaCapaItem (CodigoEmpresa, Sequencia, SequenciaProduto, SequenciaOriginal, SequenciaProdutoOriginal, CodigoProduto, CodigoUnidade) VALUES (" & intEmpresa & "," & intSequencia & "," & intSequenciaItem & "," & rsGrid!Sequencia & "," & rsGrid!SequenciaOriginal & "," & rsGrid!codigo & "," & rsGrid!CodigoUnidade & ")"
                
                Complemento 'Chama uma rotina que grava numero do pedido e observacoes no Itens quando faz o agrupamento de produtos, Nao foi possivel implementar a rotina aqio pois estava dando um erro de procedure too large ou seja estorou o tamanho da procedure Emanoel
                
            End If
            
            sSQL = "EXEC spProcessaConsignacao  1,0,'',''," & rsGrid("CodigoCliente") & ",0,0,0,0," & tpLNG(rsGrid("Codigo")) & ",'D',60,'A','C'"
            db.Execute sSQL
            
            rsGrid.MoveNext
        Loop
        
        strMsgNota = "Devolucao Emprestimo Ref. NFs: "
        strMsgNota = strMsgNota & strNumeroNota
        db.Execute "UPDATE Capa SET ObservacaoNota = '" & strMsgNota & "'  WHERE CodigoEmpresa = " & intEmpresa & " and Sequencia = " & intSequencia
        strNumeroNota = 0
    End If
    
    db.CommitTrans 'Finalizando Transacao
    Gravar = True
    Exit Function
    
TrataErros:
    Gravar = False
    db.RollbackTrans
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
    Resume 'Tem que Tirar
End Function

Private Function Arredonda(ByVal Value As Double, ByVal _
    digits As Integer) As Double
Dim Shift As Double

    Shift = 10 ^ digits
    Arredonda = CLng(Value * Shift) / Shift
End Function

Private Sub Zera_Variaveis()
   VlrIPI = 0
   VlrICMS = 0
   VlrUnitario = 0
   VlrUnitarioDevolucao = 0
   VlrICMSTotal = 0
   VlrIPITotal = 0
   BaseICMS = 0
End Sub

Private Sub Complemento()
    sSQL = "Select * From CapaItem where CodigoEmpresa = " & intEmpresa & " and Codigoproduto = " & rsGrid!codigo & " and Sequencia = " & rsGrid!Sequencia & " and SequenciaProduto = " & rsGrid!SequenciaOriginal
    rs.Open sSQL, db, , , adCmdText
    If Not rs.EOF Then
       If Not IsNull(rs!NumeroPedido) Then
          db.Execute "UPDATE CapaItem SET NumeroPedido = '" & rs!NumeroPedido & "' WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rs!CodigoProduto & " and Sequencia = " & intSequencia & " and SequenciaProduto = " & intSequenciaProduto
       End If
       If Not IsNull(rs!Observacao) Then
          db.Execute "UPDATE CapaItem SET Observacao = '" & rs!Observacao & "' WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rs!CodigoProduto & " and Sequencia = " & intSequencia & " and SequenciaProduto = " & intSequenciaProduto
       End If
       If Not IsNull(rs!ObservacaoFaturamento) Then
          db.Execute "UPDATE CapaItem SET ObservacaoFaturamento = '" & rs!ObservacaoFaturamento & "' WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rs!CodigoProduto & " and Sequencia = " & intSequencia & " and SequenciaProduto = " & intSequenciaProduto
       End If
    End If
    rs.Close
End Sub

Private Sub ConsultaEstoque()
On Error GoTo TrataErros
    lblQtdDevolver.Tag = lblQtdDevolver.Caption
    lblQtdDevolver.Caption = 0
    lblVlrDevolver.Tag = lblVlrDevolver.Caption
    lblVlrDevolver.Caption = 0
    lblQtdFaturar.Tag = lblQtdFaturar.Caption
    lblQtdFaturar.Caption = 0
    lblVlrFaturar.Tag = lblVlrFaturar.Caption
    lblVlrFaturar.Caption = 0
    booQtdEstoque = True
    If optSintetico.Value = False Or optMostra.Value = False Then
        If optAnalitico.Value = True And optMostra.Value = False Then
            If rsGrid.RecordCount > 0 Then
               Zera_Variaveis
               rsGrid.MoveFirst
               Do While Not rsGrid.EOF
               'Liberando o sistema
               DoEvents:
                  'Testando Quantidade a Fatura e Devolver
                   If booContagem = False Then
                      If rsGrid("QtdeAfaturar") > 0 Or rsGrid("QtdeADevolver") > 0 Then
                         If (rsGrid("QtdeAfaturar") + rsGrid("QtdeADevolver")) > rsGrid("Qtddifer") Then
                            MsgBox "Quantidade a Faturar ou Devolver, Maior ou Igual a Quantidade Disponivel, Qtd Disponivel = " & rsGrid("Qtddifer"), vbCritical, "SAID"
                            db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                            db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                            booQtdEstoque = False
                            rsGrid.MoveFirst
                            Exit Sub
                         End If
                      End If
                   Else
                      If rsGrid("QtdeContagem") >= 0 Or rsGrid("QtdeADevolver") > 0 Then
                         If (rsGrid("QtdeContagem") + rsGrid("QtdeADevolver")) >= rsGrid("Qtddifer") Then
                            MsgBox "Quantidade a Faturar ou Devolver, Maior ou Igual a Quantidade Disponivel, Qtd Disponivel = " & rsGrid("Qtddifer"), vbCritical, "SAID"
                            db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                            db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                            booQtdEstoque = False
                            Exit Sub
                         End If
                      End If
                   End If

                  If booContagem = False Then
                     If Not IsNull(rsGrid!QtdeAfaturar) Then
                        If rsGrid("QtdeAfaturar") > 0 Then
                           db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & rsGrid!QtdeAfaturar & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                        Else
                           If rsGrid!QtdeAfaturar = 0 Then 'so vai zerar se qtdeafaturar for = 0
                              db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                           End If
                        End If
'                        'Calculando Valor a Faturar
                        lblQtdFaturar.Caption = CCur(lblQtdFaturar.Caption) + rsGrid!QtdeAfaturar
                        If rsGrid!QtdeAfaturar > 0 Then
                           VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                           lblVlrFaturar.Caption = lblVlrFaturar.Caption + Arredonda(rsGrid!QtdeAfaturar * VlrUnitario, 2)
                        End If
                     End If
                  Else
                     If Not IsNull(rsGrid!qtdecontagem) Then
                        If rsGrid("QtdeContagem") >= 0 Then
                           db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & (rsGrid!qtddifer - rsGrid!qtdecontagem) & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                        Else
                           If rsGrid!QtdeAfaturar = 0 Then 'so vai zerar se qtdeafaturar for = 0
                              db.Execute "UPDATE CapaItem SET QuantidadeConferida = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                           End If
                        End If
                        lblQtdFaturar.Caption = CCur(lblQtdFaturar.Caption) + (rsGrid!qtddifer - rsGrid!qtdecontagem)
                        If rsGrid!qtdecontagem >= 0 Then 'Se for 0 e porque vai faturar tudo!
                           VlrUnitario = rsGrid!VlSaida / rsGrid!qtdsaida
                           lblVlrFaturar.Caption = lblVlrFaturar.Caption + Arredonda((rsGrid!qtddifer - rsGrid!qtdecontagem) * VlrUnitario, 2)
                        End If
                     End If
                  End If

                   'Pegando Valor IPI de Acordo com o Percentual do Produto
                   If booCalculaIPI = True Then
                      If rsGrid!AliquotaIPI > 0 Then
                         VlrIPI = Arredonda(((VlrUnitario * rsGrid!AliquotaIPI) / 100), 2)
                         If booContagem = False Then
                            VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeAfaturar, 2)
                         Else
                            If rsGrid!qtdecontagem = 0 Then
                               VlrIPI = Arredonda(VlrIPI * (rsGrid!qtddifer - rsGrid!qtdecontagem), 2)
                            End If
                         End If
                         lblVlrFaturar.Caption = Arredonda(lblVlrFaturar.Caption + VlrIPI, 2)
                      End If
                  End If
                   If rsGrid("QtdeADevolver") > 0 Then
                      lblQtdDevolver.Caption = lblQtdDevolver.Caption + rsGrid!QtdeADevolver
                      db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & rsGrid("QtdeADevolver") & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                      VlrUnitarioDevolucao = Arredonda(rsGrid!VlSaida / rsGrid!qtdsaida, 2)
                      lblVlrDevolver.Caption = lblVlrDevolver.Caption + Arredonda(rsGrid!QtdeADevolver * VlrUnitarioDevolucao, 2)
                      If booCalculaIPI = True Then
                         If rsGrid!AliquotaIPI > 0 Then
                            VlrIPI = Arredonda(((VlrUnitarioDevolucao * rsGrid!AliquotaIPI) / 100), 2)
                            VlrIPI = Arredonda(VlrIPI * rsGrid!QtdeADevolver, 2)
                            lblVlrDevolver.Caption = Arredonda(lblVlrDevolver.Caption + VlrIPI, 2)
                         End If
                      End If
                      Else
                         db.Execute "UPDATE CapaItem SET QuantidadeAvarias = " & 0 & " WHERE CodigoEmpresa = " & intEmpresa & " and CodigoProduto = " & rsGrid("Codigo") & " and Sequencia = " & rsGrid("Sequencia")
                   End If
                  rsGrid.MoveNext
               Loop
'               cmdPesquisar_Click
            End If
        End If
    End If
    Exit Sub
TrataErros:
    ControleErros Err.Number, Err.Description, Err.Source, Me.Caption
End Sub
