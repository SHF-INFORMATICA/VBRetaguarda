VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVENDACOMISSAO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rotina de Comissão Vendas"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VENDACOMISSAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Comissão"
      TabPicture(0)   =   "VENDACOMISSAO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbllivro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line3(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblvendedor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblinicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblfinal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line2(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line3(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSituacao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblProc"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDtfim"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDtIni"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbAuxVendedor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbvend"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "optFinanceiro"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Parametros"
      TabPicture(1)   =   "VENDACOMISSAO.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "Line2(3)"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Line3(2)"
      Tab(1).Control(6)=   "Line4"
      Tab(1).Control(7)=   "Line2(2)"
      Tab(1).ControlCount=   8
      Begin VB.CheckBox optFinanceiro 
         Caption         =   "Considerar Título Baixado?"
         Height          =   495
         Left            =   3480
         TabIndex        =   42
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Frame Frame6 
         Caption         =   "Por Venda"
         Height          =   855
         Left            =   -74520
         TabIndex        =   37
         Top             =   4920
         Width           =   4695
         Begin VB.TextBox txtVenda 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   3120
            TabIndex        =   39
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox optVenda 
            Caption         =   "Valor Venda"
            Height          =   240
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   4200
            TabIndex        =   40
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   3120
         TabIndex        =   30
         Top             =   2640
         Width           =   2895
         Begin VB.OptionButton optVendedor 
            Caption         =   "Por Vendedo&r"
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Percentual informado no cadastro do vendedor"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton optValorProduto 
            Caption         =   "Por Valor &Produto"
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Baseado no preço por produto"
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optValorVenda 
            Caption         =   "Por Valor &Venda"
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Baseado no valor de venda do pedido"
            Top             =   120
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Atacado"
         Height          =   1335
         Left            =   -74520
         TabIndex        =   21
         Top             =   3240
         Width           =   4695
         Begin VB.CheckBox optAtacado 
            Caption         =   "Abaixo Valor de Atacado"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox optAtacado 
            Caption         =   "Acima Valor de Atacado"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtAtacado 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   0
            Left            =   3120
            TabIndex        =   10
            ToolTipText     =   "Informe Percentual de Comissão"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtAtacado 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   1
            Left            =   3120
            TabIndex        =   11
            ToolTipText     =   "Informe Percentual de Comissão"
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4200
            TabIndex        =   29
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4200
            TabIndex        =   28
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Varejo"
         Height          =   1815
         Left            =   -74520
         TabIndex        =   20
         Top             =   1200
         Width           =   4695
         Begin VB.TextBox txtVarejo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   2
            Left            =   3120
            TabIndex        =   44
            ToolTipText     =   "Informe Percentual de Comissão"
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox optVarejo 
            Caption         =   "Valor de Varejo"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox optVarejo 
            Caption         =   "Abaixo Valor de Varejo"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   2775
         End
         Begin VB.CheckBox optVarejo 
            Caption         =   "Acima Valor de Varejo"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtVarejo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   1
            Left            =   3120
            TabIndex        =   9
            ToolTipText     =   "Informe Percentual de Comissão"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtVarejo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   0
            Left            =   3120
            TabIndex        =   8
            ToolTipText     =   "Informe Percentual de Comissão"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   4200
            TabIndex        =   45
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4200
            TabIndex        =   27
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   4200
            TabIndex        =   26
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.ComboBox cmbvend 
         Height          =   360
         Left            =   1680
         TabIndex        =   1
         Top             =   1080
         Width           =   4095
      End
      Begin VB.ComboBox cmbAuxVendedor 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
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
         TabIndex        =   15
         Top             =   2640
         Width           =   2895
         Begin VB.OptionButton optFaturado 
            Caption         =   "&Faturados"
            Height          =   375
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Comissão sobre somente pedidos faturados"
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optNota 
            Caption         =   "Somente &Nota"
            Height          =   375
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Comissão sobre nota emitida"
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton optFatNota 
            Caption         =   "Fa&turados/Nota"
            Height          =   375
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Comissão Geral"
            Top             =   1080
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   4935
         Begin VB.OptionButton optSintetico 
            Caption         =   "&Sintetico"
            Height          =   375
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   4215
         End
         Begin VB.OptionButton optAnalitico 
            Caption         =   "&Analitico"
            Height          =   375
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   4215
         End
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtDtfim 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label lblProc 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   195
         TabIndex        =   41
         Top             =   5640
         Width           =   60
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   3
         X1              =   -75000
         X2              =   -69000
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label lblSituacao 
         AutoSize        =   -1  'True
         Caption         =   "Situação"
         Height          =   240
         Left            =   3375
         TabIndex        =   35
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Parametrização"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   520
         Left            =   -74970
         TabIndex        =   34
         Top             =   530
         Width           =   5970
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   2
         X1              =   -75000
         X2              =   -69000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         X1              =   -75000
         X2              =   -69000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   2
         X1              =   -75000
         X2              =   -69000
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   6000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   6000
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label lblfinal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblinicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblvendedor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   240
         Left            =   585
         TabIndex        =   17
         Top             =   1080
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         X1              =   0
         X2              =   6000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   6000
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000002&
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   6000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lbllivro 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Comissão Vendas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   30
         TabIndex        =   13
         Top             =   360
         Width           =   5970
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar"
            Key             =   "gravar"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   3960
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8760
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
               Picture         =   "VENDACOMISSAO.frx":5C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACOMISSAO.frx":6DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACOMISSAO.frx":7E73
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACOMISSAO.frx":8E28
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACOMISSAO.frx":9F33
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VENDACOMISSAO.frx":BF15
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmVENDACOMISSAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim TABCOMISSAO         As New ADODB.Recordset
   Dim TABitemCOMISSAO     As New ADODB.Recordset
   Dim COMISSAO_ID_N       As Integer
   Dim Valr_Comissao_N     As Double
   Dim Perc_Comissao_N     As Double
   Dim Valr_Desconto_Item  As Double
   Dim Valr_Faturado_Item  As Double

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmVENDACOMISSAO)
   Me.Caption = Me.Caption & " - " & Me.Name
   preencheCombovend cmbVend

   CRITERIO_A = Month(Date)
   If Len(CRITERIO_A) = 1 Then _
      CRITERIO_A = "0" & CRITERIO_A

   txtDtIni.PromptInclude = False
   txtDtIni.Text = "01/" & CRITERIO_A & "/" & Year(Date)
   txtDtIni.PromptInclude = True

   txtDtFim.PromptInclude = False
   
   'strData = Format(strData, "yyyymmdd")
   CRITERIO_A = FimDoMes(txtDtIni.Text, False)
   CRITERIO_A = Right(CRITERIO_A, 2) & "/" & Mid(CRITERIO_A, 5, 2) & "/" & Left(CRITERIO_A, 4)
   txtDtFim.Text = CRITERIO_A
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub optVarejo_Click(Index As Integer)
'On Error GoTo ERRO_TRATA

   HABILITA_CAMPOS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optVarejo_Click"
End Sub

Private Sub optatacado_Click(Index As Integer)
'On Error GoTo ERRO_TRATA

   HABILITA_CAMPOS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optatacado_Click"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   'Toolbar1.Buttons(3).Visible = True
   Toolbar1.Buttons(4).Caption = "Gerar"
   Toolbar1.Buttons(2).Visible = True

   If SSTab1.Tab = 1 Then
      Toolbar1.Buttons(2).Visible = False
      Toolbar1.Buttons(3).Visible = False
      Toolbar1.Buttons(4).Caption = "Gravar"

      MOSTRA_PARAMETROS
   End If

   Toolbar1.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSTab1_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         If SSTab1.Tab = 1 Then
            GRAVA_PARAMETROS
            Else: GERA_REL
         End If
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
      Case "print"
         CHAMA_REL
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub optVenda_Click()
   If optVenda.Value = 0 Then
      txtVenda.Text = ""
      txtVenda.Enabled = False
      Else
         txtVenda.Enabled = True
         txtVenda.SetFocus
   End If
End Sub

Private Sub txtDtFim_LostFocus()
'On Error GoTo ERRO_TRATA

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then

      If TABCOMISSAO.State = 1 Then _
         TABCOMISSAO.Close

      SQL = "select * from COMISSAO"
      SQL = SQL & " where dtini >= '" & DMA(txtDtIni.Text) & "'"
      SQL = SQL & " and dtfim <= '" & DMA(txtDtFim.Text) & "'"
      TABCOMISSAO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TABCOMISSAO.EOF Then _
         lblSituacao.Caption = Trim(TABCOMISSAO.Fields("situacao").Value)

      If TABCOMISSAO.State = 1 Then _
         TABCOMISSAO.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtfim_LostFocus"
End Sub

Private Sub txtVarejo_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAtacado_KeyPress"
End Sub

Private Sub txtAtacado_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAtacado_KeyPress"
End Sub

Private Sub cmbvend_Click()
On Error Resume Next

   cmbAuxVendedor.ListIndex = cmbVend.ListIndex
End Sub

Private Sub cmbvend_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtDtIni.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_KeyPress"
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
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
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Function RetornaDescricaoVEND(vendId As String) As String
'On Error GoTo ERRO_TRATA

   RetornaDescricaoVEND = ""

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select descricao from vwVendedor "
   SQL = SQL & " where vendedor_id = " & vendId
   Set TabVENDEDOR = CONECTA_RETAGUARDA.OpenRecordset(SQL, 4)
   If Not TabVENDEDOR.EOF Then _
      RetornaDescricaoVEND = TabVENDEDOR!DESCRICAO
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RetornaDescricaoVEND"
End Function

Private Sub preencheCombovend(NomeCombo As ComboBox)
'On Error GoTo ERRO_TRATA

   NomeCombo.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select descricao,vendedor_id from vwVendedor "
   SQL = SQL & " where status = 'A'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by descricao"
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVENDEDOR.EOF Then
      'Mundando o ponteiro do mouse, para mostrar para o usuario que esta processando...
      Screen.MousePointer = vbHourglass

      TabVENDEDOR.MoveFirst
      Do Until TabVENDEDOR.EOF
         'Importantissimo
         DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"

         NomeCombo.AddItem Trim(TabVENDEDOR!DESCRICAO) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
         cmbAuxVendedor.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
         TabVENDEDOR.MoveNext
      Loop
   End If
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   'Voltando o ponteiro do mouse para o tipo default, ponteiro.
   Screen.MousePointer = vbDefault

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "preencheCombovend"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   cmbVend.Text = ""
   txtDtIni.PromptInclude = False
   txtDtIni.Text = ""
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   optSintetico.Value = False
   optAnalitico.Value = False
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   optValorProduto.Value = False
   optValorVenda.Value = False
   optFaturado.Value = False
   optFinanceiro.Value = 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub GRAVA_PARAMETROS()
'On Error GoTo ERRO_TRATA

   CONT_N = 0
   While CONT_N < 3
      If txtVarejo(CONT_N).Text <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select COMISSAOPARAM_ID,TIPO,VALOR from COMISSAOPARAM"
         SQL = SQL & " where tipo = '" & Trim(optVarejo(CONT_N).Caption) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            NUMR_ID_N = MAX_ID("comissaoparam_id", "comissaoparam", "", "", "", "")
            SQL = "insert into COMISSAOPARAM values("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & ",'" & Trim(optVarejo(CONT_N).Caption) & "'"
               SQL = SQL & "," & tpMOEDA(txtVarejo(CONT_N).Text)
            SQL = SQL & ")"
            Else
               SQL = "update COMISSAOPARAM set "
                  SQL = SQL & " valor = " & tpMOEDA(txtVarejo(CONT_N).Text)
               SQL = SQL & " where tipo = '" & Trim(optVarejo(CONT_N).Caption) & "'"
         End If
         CONECTA_RETAGUARDA.Execute SQL
      End If

      CONT_N = CONT_N + 1
   Wend


   CONT_N = 0
   While CONT_N < 2
      If txtAtacado(CONT_N).Text <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select COMISSAOPARAM_ID,TIPO,VALOR from COMISSAOPARAM"
         SQL = SQL & " where tipo = '" & Trim(optAtacado(CONT_N).Caption) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            NUMR_ID_N = MAX_ID("comissaoparam_id", "comissaoparam", "", "", "", "")
            SQL = "insert into COMISSAOPARAM values("
               SQL = SQL & NUMR_ID_N
               SQL = SQL & ",'" & Trim(optAtacado(CONT_N).Caption) & "'"
               SQL = SQL & "," & tpMOEDA(txtAtacado(CONT_N).Text)
            SQL = SQL & ")"
            Else
               SQL = "update COMISSAOPARAM set "
                  SQL = SQL & " valor = " & tpMOEDA(txtAtacado(CONT_N).Text)
               SQL = SQL & " where tipo = '" & Trim(optAtacado(CONT_N).Caption) & "'"
         End If
         CONECTA_RETAGUARDA.Execute SQL
      End If

      CONT_N = CONT_N + 1
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select COMISSAOPARAM_ID,TIPO,VALOR from COMISSAOPARAM"
   SQL = SQL & " where tipo = '" & Trim(optVenda.Caption) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      NUMR_ID_N = MAX_ID("comissaoparam_id", "comissaoparam", "", "", "", "")
      SQL = "insert into COMISSAOPARAM values("
         SQL = SQL & NUMR_ID_N
         SQL = SQL & ",'" & Trim(optVenda.Caption) & "'"
         SQL = SQL & "," & tpMOEDA(txtVenda.Text)
      SQL = SQL & ")"
      Else
         SQL = "update COMISSAOPARAM set "
            SQL = SQL & " valor = " & tpMOEDA(txtVenda.Text)
         SQL = SQL & " where tipo = '" & Trim(optVenda.Caption) & "'"
   End If
   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   MsgBox "Operação realizada com sucesso."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PARAMETROS"
End Sub

Sub HABILITA_CAMPOS()
'On Error GoTo ERRO_TRATA

   If optVarejo(0).Value = 0 Then
      txtVarejo(0).Text = ""
      txtVarejo(0).Enabled = False
      Else
         txtVarejo(0).Enabled = True
         txtVarejo(0).SetFocus
   End If

   If optVarejo(1).Value = 0 Then
      txtVarejo(1).Text = ""
      txtVarejo(1).Enabled = False
      Else
         txtVarejo(1).Enabled = True
         txtVarejo(1).SetFocus
   End If

   If optVarejo(2).Value = 0 Then
      txtVarejo(2).Text = ""
      txtVarejo(2).Enabled = False
      Else
         txtVarejo(2).Enabled = True
         txtVarejo(2).SetFocus
   End If

   If optAtacado(0).Value = 0 Then
      txtAtacado(0).Text = ""
      txtAtacado(0).Enabled = False
      Else
         txtAtacado(0).Enabled = True
         txtAtacado(0).SetFocus
   End If

   If optAtacado(1).Value = 0 Then
      txtAtacado(1).Text = ""
      txtAtacado(1).Enabled = False
      Else
         txtAtacado(1).Enabled = True
         txtAtacado(1).SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "HABILITA_CAMPOS"
End Sub

Sub MOSTRA_PARAMETROS()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select COMISSAOPARAM_ID,TIPO,VALOR from COMISSAOPARAM"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If Trim(TabTemp.Fields("tipo").Value) = Trim(optAtacado(0).Caption) Then
         txtAtacado(0).Text = tpMOEDA(TabTemp.Fields("valor").Value)
         optAtacado(0).Value = 1
         Else
            If Trim(TabTemp.Fields("tipo").Value) = Trim(optAtacado(1).Caption) Then
               txtAtacado(1).Text = tpMOEDA(TabTemp.Fields("valor").Value)
               optAtacado(1).Value = 1
            End If
      End If

      If Trim(TabTemp.Fields("tipo").Value) = Trim(optVarejo(0).Caption) Then
         txtVarejo(0).Text = tpMOEDA(TabTemp.Fields("valor").Value)
         optVarejo(0).Value = 1
         'Else
      End If

      If Trim(TabTemp.Fields("tipo").Value) = Trim(optVarejo(1).Caption) Then
         txtVarejo(1).Text = tpMOEDA(TabTemp.Fields("valor").Value)
         optVarejo(1).Value = 1
      End If

      If Trim(TabTemp.Fields("tipo").Value) = Trim(optVarejo(2).Caption) Then
         txtVarejo(2).Text = tpMOEDA(TabTemp.Fields("valor").Value)
         optVarejo(2).Value = 1
      End If

      If Trim(TabTemp.Fields("tipo").Value) = Trim(optVenda.Caption) Then
         txtVenda.Text = tpMOEDA(TabTemp.Fields("valor").Value)
         optVenda.Value = 1
      End If

      TabTemp.MoveNext
   Wend

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PARAMETROS"
End Sub

Sub DESABILITA_CAMPOS()
   txtVarejo(0).Enabled = False
   txtVarejo(1).Enabled = False

   txtAtacado(0).Enabled = False
   txtAtacado(1).Enabled = False
End Sub

Sub GERA_REL()
'On Error GoTo ERRO_TRATA

   Call ATUALIZA_TABELA_COMISSAO

   CRITERIO_A = ""
   NUMR_SEQ_N = 0

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      Toolbar1.Enabled = False
      lblSituacao.Caption = "Aberta"

      GRAVA_CABEÇA_COMISSAO

      If TabCabeca.State = 1 Then _
         TabCabeca.Close

'=============================================================
      SQL = "select PEDIDO.EMPRESA_ID, PEDIDO.CGCCPF, "
      SQL = SQL & " PEDIDO.vendedor_id, PEDIDO.DT_REQ, PEDIDO.STATUS, "
      SQL = SQL & " PEDIDO.VALOR_DESCONTO, PEDIDO.NOME_CLIENTE, PEDIDO.PEDIDO_ID, "
      SQL = SQL & " PEDIDO.VALOR_TOTAL, PEDIDO.cliente_id, "

SQL = SQL & " ITEMLANCAMENTO.SEQ , ITEMLANCAMENTO.formapagto_id, ITEMLANCAMENTO.VALOR_ITEM, "
SQL = SQL & " ITEMLANCAMENTO.STATUS AS status_titulo, ITEMLANCAMENTO.DT_VENCIMENTO, "
SQL = SQL & " ITEMLANCAMENTO.DT_BAIXA, ITEMLANCAMENTO.DT_CANCELA, ITEMLANCAMENTO.DT_CAD, "
SQL = SQL & " ITEMLANCAMENTO.VALOR_DESCONTO AS vlr_desconto_titulo, ITEMLANCAMENTO.NUMR_DP "

      SQL = SQL & " from PEDIDO "
      SQL = SQL & " INNER JOIN LANCAMENTO "
      SQL = SQL & " ON PEDIDO.estabelecimento_ID = LANCAMENTO.estabelecimento_ID "
      SQL = SQL & " AND PEDIDO.pessoa_id = LANCAMENTO.pessoa_id "
      SQL = SQL & " AND PEDIDO.pedido_id = LANCAMENTO.NUMR_DOC "
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
'=============================================================
      SQL = SQL & " where pedido.estabelecimento_id = " & ESTABELECIMENTO_ID_N

SQL = SQL & " and ITEMLANCAMENTO.DT_BAIXA >= '" & DMA(txtDtIni.Text) & "'"
SQL = SQL & " and ITEMLANCAMENTO.DT_BAIXA <= '" & DMA(txtDtFim.Text) & "'"

SQL = SQL & " and year(ITEMLANCAMENTO.DT_BAIXA) >= " & Year(txtDtFim.Text)

      SQL = SQL & " and PEDIDO.tipo_registro = 'R' "
      SQL = SQL & " and PEDIDO.status in (3,5,7) "
      'SQL = SQL & " and ITEMLANCAMENTO.status = 'B'"

      If Trim(cmbAuxVendedor.Text) <> "" Then _
         SQL = SQL & " and PEDIDO.vendedor_id = " & Trim(cmbAuxVendedor.Text)

      If optFinanceiro.Value = 1 Then _
         SQL = SQL & " and ITEMLANCAMENTO.Status = 'B'"

      SQL = SQL & " ORDER BY PEDIDO.vendedor_id, ITEMLANCAMENTO.SEQ"

      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabCabeca.EOF
         NUMR_SEQ_N = NUMR_SEQ_N + 1
         lblProc.Caption = "Processando ... " & NUMR_SEQ_N
         DoEvents

         If optFatNota.Value = True Then
            GERA_ITEM_REQUISIÇÃO
            Else
               If optNota.Value = True Then
                  'GERA_NOTA
                  Else
                     'If optFaturado.Value = True Then _
                        gera_somente_faturado
               End If
         End If

         TabCabeca.MoveNext
      Wend

      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      CALCULA_COMISSÃO

      CHAMA_REL

      Toolbar1.Enabled = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_REL"
End Sub

Sub GERA_ITEM_REQUISIÇÃO()
'On Error GoTo ERRO_TRATA

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select PEDIDO.CLIENTE_ID, PEDIDO.EMPRESA_ID, PEDIDO.VENDEDOR_ID, PEDIDO.PEDIDO_ID, "
   SQL = SQL & " PEDIDO.CGCCPF AS CNPJCPF_CLI, PEDIDO.usuario_id, PEDIDO.DT_REQ, PEDIDO.STATUS AS StatusPedido,"
   SQL = SQL & " PEDIDO.TIPO_REGISTRO, PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_RECEBIDO, PEDIDO.VALOR_TOTAL, PEDIDOITEM.SEQ_ID,"
   SQL = SQL & " PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.PERC_DESC, PEDIDOITEM.VALOR_DESCONTO, "
   SQL = SQL & " PEDIDOITEM.STATUS AS StatusPedidoItem, PEDIDOITEM.PRECO_CUSTO AS PrecoCustoItem, PEDIDOITEM.TIPO_REG, PRODUTO.CODG_PRODUTO,"
   SQL = SQL & " PRODUTO.DESCRICAO, PRODUTO.TIPO_PROD, PRODUTO.PRECO_CUSTO AS PrecoCustoProduto,"
   SQL = SQL & " Produto.PRECO_ATACADO , Produto.PRECO_Venda, Produto.PERC_COMIS"

   SQL = SQL & " from PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.pedido_id = " & TabCabeca.Fields("pedido_id").Value
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and PEDIDOITEM.produto_id = PRODUTO.produto_id"
   SQL = SQL & " and PEDIDOITEM.tipo_reg = 'PC' "
   SQL = SQL & " and PEDIDO.status in (3,5,7) "
   SQL = SQL & " and pedidoitem.status <> 'C' "

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabPedidoItem.EOF

      GRAVA_ITEM_COMISSAO

      TabPedidoItem.MoveNext
   Wend
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_ITEM_REQUISIÇÃO"
End Sub

Sub GRAVA_CABEÇA_COMISSAO()
'On Error GoTo ERRO_TRATA

   COMISSAO_ID_N = 0

   If TABCOMISSAO.State = 1 Then _
      TABCOMISSAO.Close

   SQL = "select * from COMISSAO"
   SQL = SQL & " where dtini >= '" & DMA(txtDtIni.Text) & "'"
   SQL = SQL & " and dtfim <= '" & DMA(txtDtFim.Text) & "'"
   TABCOMISSAO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TABCOMISSAO.EOF Then _
      COMISSAO_ID_N = TABCOMISSAO.Fields("comissao_id").Value

   If TABCOMISSAO.State = 1 Then _
      TABCOMISSAO.Close

   SQL = "delete from COMISSAOITEM "
   'SQL = SQL & " where comissao_id = " & COMISSAO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from COMISSAO "
   'SQL = SQL & " where comissao_id = " & COMISSAO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   COMISSAO_ID_N = MAX_ID("comissao_id", "comissao", "", "", "", "")
   SQL = "insert into COMISSAO values("
      SQL = SQL & COMISSAO_ID_N
      SQL = SQL & "," & EMPRESA_ID_N
      SQL = SQL & ",'" & DMA(txtDtIni.Text) & "'"
      SQL = SQL & ",'" & DMA(txtDtFim.Text) & "'"
      SQL = SQL & ",'" & Trim(lblSituacao.Caption) & "'"
   SQL = SQL & " )"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABEÇA_COMISSAO"
End Sub

Sub GRAVA_ITEM_COMISSAO()
'On Error GoTo ERRO_TRATA

   VALOR_DESCONTO_N = 0 & TabPedidoItem.Fields("valor_desconto").Value

   Valr_Desconto_Item = 0
   Valr_Faturado_Item = 0

   If TabCabeca.Fields("status_titulo").Value = "B" Then
      'Valr_Desconto_Item = 0 & Round(TabCABECA.Fields("vlr_desconto_titulo").Value)
      'Valr_Faturado_Item = 0 & Round(TabCABECA.Fields("valor_item").Value)

      Valr_Desconto_Item = 0 & TabCabeca.Fields("vlr_desconto_titulo").Value
      Valr_Faturado_Item = 0 & TabCabeca.Fields("valor_item").Value
   End If

   If TABitemCOMISSAO.State = 1 Then _
      TABitemCOMISSAO.Close

   SQL = "select EMPRESA.EMPRESA_ID, EMPRESA.CGC, COMISSAO.COMISSAO_ID, "
   SQL = SQL & " COMISSAO.DTINI, COMISSAO.DTFIM, COMISSAO.SITUACAO, COMISSAOITEM.PEDIDO_ID, "
   SQL = SQL & " COMISSAOITEM.VENDEDOR_ID, COMISSAOITEM.PRODUTO_ID, COMISSAOITEM.DESC_PROD,"
   SQL = SQL & " COMISSAOITEM.CLIENTE_ID, COMISSAOITEM.CNPJCPF, COMISSAOITEM.NOME_CLI, "
   SQL = SQL & " COMISSAOITEM.NUMR_NFE, COMISSAOITEM.NUMR_CUPOM, COMISSAOITEM.PR_ITEM_VENDA, "
   SQL = SQL & " COMISSAOITEM.PR_ITEM_VAREJO, COMISSAOITEM.PR_ITEM_ATACADO, COMISSAOITEM.VALR_COMIS_PROD,"
   SQL = SQL & " COMISSAOITEM.VALR_COMIS_TOT , COMISSAOITEM.QTDE_VENDIDA, COMISSAOITEM.PERC_COMIS, "
   SQL = SQL & " COMISSAOITEM.VALOR_FATURADO "
   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN COMISSAO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = COMISSAO.EMPRESA_ID "
   SQL = SQL & " INNER JOIN COMISSAOITEM "
   SQL = SQL & " ON COMISSAO.COMISSAO_ID = COMISSAOITEM.COMISSAO_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID "
   SQL = SQL & " AND COMISSAO.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"

   SQL = SQL & " where COMISSAOITEM.comissao_id = " & COMISSAO_ID_N
   SQL = SQL & " and COMISSAOITEM.vendedor_id = " & Trim(TabCabeca.Fields("vendedor_id").Value)
   SQL = SQL & " and COMISSAOITEM.PRODUTO_ID = " & Trim(TabPedidoItem.Fields("PRODUTO_ID").Value)
   SQL = SQL & " and COMISSAOITEM.PEDIDO_ID = " & Trim(TabPedidoItem.Fields("PEDIDO_ID").Value)
   SQL = SQL & " and estabelecimento.empresa_ID = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento.estabelecimento_ID = " & ESTABELECIMENTO_ID_N

   TABitemCOMISSAO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TABitemCOMISSAO.EOF Then
      SQL = "insert into COMISSAOITEM values("
         SQL = SQL & COMISSAO_ID_N                                                              'COMISSAO_ID
         SQL = SQL & "," & tpMOEDA(TabPedidoItem.Fields("PEDIDO_ID").Value)                     'PEDIDO_ID
         SQL = SQL & "," & TabPedidoItem.Fields("vendedor_id").Value                            'VENDEDOR_ID

         SQL = SQL & "," & Trim(TabPedidoItem.Fields("PRODUTO_ID").Value)                       'PRODUTO_ID
         SQL = SQL & ",'" & Trim(Left(TabPedidoItem.Fields("DESCRICAO").Value, 80)) & "'"       'DESCRICAO PRODUTO

         SQL = SQL & "," & Trim(TabPedidoItem.Fields("CLIENTE_ID").Value)                       'CLIENTE_ID
         SQL = SQL & ",'" & Trim(TabCabeca.Fields("CGCCPF").Value) & "'"                        'CNPJCPF
         SQL = SQL & ",'" & Trim(Left(TabCabeca.Fields("NOME_CLIENTE").Value, 30)) & "'"        'NOME_CLIENTE

         SQL = SQL & ",0" & Trim(TabPedidoItem.Fields("NUMR_DOC").Value)                        'NUMR_DOC
         SQL = SQL & ",0" & Trim(TabPedidoItem.Fields("NUMR_CUPOM").Value)                      'NUMR_CUPOM

         SQL = SQL & "," & tpMOEDA(TabPedidoItem.Fields("valor_item").Value - VALOR_DESCONTO_N) 'PR_ITEM_VENDA
         SQL = SQL & "," & tpMOEDA(TabPedidoItem.Fields("PRECO_VENDA").Value)                   'PR_ITEM_VAREJO
         SQL = SQL & "," & tpMOEDA(TabPedidoItem.Fields("PRECO_ATACADO").Value)                 'PR_ITEM_ATACADO

         SQL = SQL & "," & tpMOEDA(0)                                                           'VALR_COMIS_PROD
         SQL = SQL & "," & tpMOEDA(0)                                                           'VALR_COMIS_TOT
         SQL = SQL & "," & tpMOEDA(TabPedidoItem.Fields("QTD_pedida").Value)                    'QTDE_VENDIDA

         SQL = SQL & ",0"                                                                       'PERC_COMIS
         'SQL = SQL & "," & tpMOEDA(Valr_Faturado_Item - Valr_Desconto_Item)              'VALOR_FATURADO
         SQL = SQL & ",0"              'VALOR_FATURADO
      SQL = SQL & " )"
      Else
         SQL = "update COMISSAOITEM set "

            SQL = SQL & "PR_ITEM_VENDA = " & tpMOEDA(TabPedidoItem.Fields("valor_item").Value - VALOR_DESCONTO_N) 'PR_ITEM_VENDA

            SQL = SQL & ",PR_ITEM_VAREJO = " & tpMOEDA(TabPedidoItem.Fields("PRECO_VENDA").Value)    'PR_ITEM_VAREJO
            SQL = SQL & ",PR_ITEM_ATACADO = " & tpMOEDA(TabPedidoItem.Fields("PRECO_ATACADO").Value) 'PR_ITEM_ATACADO

            SQL = SQL & ",VALR_COMIS_PROD = " & tpMOEDA(0)                                                  'VALR_COMIS_PROD
            SQL = SQL & ",VALR_COMIS_TOT = VALR_COMIS_TOT + " & tpMOEDA(0)                                  'VALR_COMIS_TOT
            SQL = SQL & ",QTDE_VENDIDA = QTDE_VENDIDA + " & tpMOEDA(TabPedidoItem.Fields("QTD_PEDIDA").Value)              'QTDE_VENDIDA
            SQL = SQL & ",CNPJCPF = '" & Trim(TabCabeca.Fields("CGCCPF").Value) & "'"                       'CNPJCPF
            SQL = SQL & ",NOME_CLI = '" & Trim(Left(TabCabeca.Fields("NOME_CLIENTE").Value, 30)) & "'"      'NOME_CLI
            SQL = SQL & ",DESC_PROD = '" & Trim(Left(TabPedidoItem.Fields("DESCRICAO").Value, 80)) & "'"    'DESCRICAO
            SQL = SQL & ",PERC_COMIS = 0"                                                                   'PERC_COMIS
            'SQL = SQL & ",VALOR_FATURADO = VALOR_FATURADO + " & tpMOEDA(Valr_Faturado_Item - Valr_Desconto_Item)      'VALOR_FATURADO

         SQL = SQL & " where comissao_id = " & COMISSAO_ID_N
         SQL = SQL & " and vendedor_id = " & Trim(TabCabeca.Fields("vendedor_id").Value)
         SQL = SQL & " and PRODUTO_ID = " & Trim(TabPedidoItem.Fields("PRODUTO_ID").Value)
         SQL = SQL & " and PEDIDO_ID = " & Trim(TabPedidoItem.Fields("PEDIDO_ID").Value)
   End If
   If TABitemCOMISSAO.State = 1 Then _
      TABitemCOMISSAO.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ITEM_COMISSAO"
End Sub

Sub CALCULA_COMISSÃO()
'On Error GoTo ERRO_TRATA

'ATENÇÃO O DESCONTO DADO NA CABEÇA DA VENDA NÃO ESTA SENDO LEVADO EM CONSIDERAÇÃO
'OS CALCULOS REALIZADOS ABAIXO LEVA EM CONTA VALORES DE VENDA DOS PRODUTOS SOMENTE

   Dim Conta_Vaca As Long

   COMISSAO_ID_N = 0
   Valr_Comissao_N = 0

   If TABCOMISSAO.State = 1 Then _
      TABCOMISSAO.Close

   SQL = "select * from COMISSAO"
   SQL = SQL & " where dtini >= '" & DMA(txtDtIni.Text) & "'"
   SQL = SQL & " and dtfim <= '" & DMA(txtDtFim.Text) & "'"
   TABCOMISSAO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TABCOMISSAO.EOF Then
      COMISSAO_ID_N = TABCOMISSAO.Fields("comissao_id").Value

      If TABitemCOMISSAO.State = 1 Then _
         TABitemCOMISSAO.Close

      SQL = "select * from COMISSAOITEM "
      SQL = SQL & " where comissao_id = " & COMISSAO_ID_N
      TABitemCOMISSAO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TABitemCOMISSAO.EOF
         Valr_Comissao_N = 0
         Perc_Comissao_N = 0

         'pega valor de venda o item, e procura percentual conforme parametrização
         If optValorProduto.Value = True Then
            'menor que valor atacado
            If TABitemCOMISSAO.Fields("PR_ITEM_VENDA").Value < TABitemCOMISSAO.Fields("PR_ITEM_atacado").Value Then
               Perc_Comissao_N = 0
               BUSCA_PERCENTUAL_COMISSAO optAtacado(1).Caption
               Else
                  'menor que valor varejo
                  If TABitemCOMISSAO.Fields("PR_ITEM_VENDA").Value < TABitemCOMISSAO.Fields("PR_ITEM_varejo").Value Then
                     Perc_Comissao_N = 0
                     BUSCA_PERCENTUAL_COMISSAO optVarejo(1).Caption
                     Else
                        'igual que valor varejo
'MsgBox TABitemCOMISSAO.Fields("PR_ITEM_varejo").Value
                        If TABitemCOMISSAO.Fields("PR_ITEM_VENDA").Value = TABitemCOMISSAO.Fields("PR_ITEM_varejo").Value Then
                           Perc_Comissao_N = 0
                           BUSCA_PERCENTUAL_COMISSAO optVarejo(2).Caption
                           Else
                              'maior que valor varejo
                              If TABitemCOMISSAO.Fields("PR_ITEM_VENDA").Value > TABitemCOMISSAO.Fields("PR_ITEM_varejo").Value Then
                                 Perc_Comissao_N = 0
                                 BUSCA_PERCENTUAL_COMISSAO optVarejo(0).Caption
                              End If
                        End If
                  End If
            End If
            Else
               'pega direto da tabela PEDIDO
               If optValorVenda.Value = True Then
                  Perc_Comissao_N = 0
                  BUSCA_PERCENTUAL_COMISSAO Trim(optVenda.Caption)
                  Else
                     'faz calculo baseano no percentual atribuido no cadastro do vendedor
                     If optVendedor.Value = True Then
                        Perc_Comissao_N = 0

                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close

                        SQL = "select PERC_COMISSAO from vwVENDEDOR"
                        SQL = SQL & " where vendedor_id = " & TABitemCOMISSAO.Fields("vendedor_id").Value
                        TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If Not TabConsulta.EOF Then _
                           If Not IsNull(TabConsulta.Fields(0).Value) Then _
                              If TabConsulta.Fields(0).Value > 0 Then _
                                 Perc_Comissao_N = TabConsulta.Fields(0).Value

                        If TabConsulta.State = 1 Then _
                           TabConsulta.Close
                     End If
               End If
         End If

'Valr_Comissao_N = ((TABitemCOMISSAO.Fields("valor_faturado").Value) * Perc_Comissao_N / 100)

         Valr_Comissao_N = (TABitemCOMISSAO.Fields("PR_ITEM_VENDA").Value * _
                            TABitemCOMISSAO.Fields("qtde_vendida").Value) * _
                            Perc_Comissao_N / 100

         'If optFinanceiro.Value = 1 Then
            'If TABCABECA.Fields("status_titulo").Value = "B" Then
            'Valr_Comissao_N = (TABitemCOMISSAO.Fields("valor_faturado").Value * _
                               TABitemCOMISSAO.Fields("qtde_vendida").Value) * _
                               Perc_Comissao_N / 100
         'End If

         SQL = "update COMISSAOITEM set "
            SQL = SQL & "VALR_COMIS_PROD = " & tpMOEDA(Valr_Comissao_N)                       'VALR_COMIS_PROD
            SQL = SQL & ",PERC_COMIS = " & tpMOEDA(Perc_Comissao_N)                           'PERC_COMIS

If Conta_Vaca <> TABitemCOMISSAO.Fields("pedido_id").Value Then
   SQL = SQL & ",valor_faturado = " & tpMOEDA(BUSCA_FATURAMENTO_QUITADO(TABitemCOMISSAO.Fields("pedido_id").Value))                              'valor_faturado
   Conta_Vaca = TABitemCOMISSAO.Fields("pedido_id").Value
End If

         SQL = SQL & " where comissao_id = " & Trim(TABitemCOMISSAO.Fields("comissao_id").Value)
         SQL = SQL & " and vendedor_id = " & Trim(TABitemCOMISSAO.Fields("vendedor_id").Value)
         SQL = SQL & " and PRODUTO_ID = '" & Trim(TABitemCOMISSAO.Fields("PRODUTO_ID").Value) & "'"
         SQL = SQL & " and PEDIDO_ID = " & Trim(TABitemCOMISSAO.Fields("PEDIDO_ID").Value)
         CONECTA_RETAGUARDA.Execute SQL

         TABitemCOMISSAO.MoveNext
      Wend
      If TABitemCOMISSAO.State = 1 Then _
         TABitemCOMISSAO.Close
   End If

   If TABCOMISSAO.State = 1 Then _
      TABCOMISSAO.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CALCULA_COMISSÃO"
End Sub

Sub CHAMA_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{COMISSAO.empresa_id} = " & EMPRESA_ID_N
   FORMULA_REL = FORMULA_REL & " and {COMISSAO.DTINI} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

   If Trim(cmbAuxVendedor.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {COMISSAOitem.vendedor_id} = " & Trim(cmbAuxVendedor.Text)

   If optAnalitico.Value = True Then
      Nome_Relatorio = "comissao.rpt"
      Else
         Nome_Relatorio = "comissao_sintetico.rpt"
         Nome_Relatorio = "comis_fat.rpt"
   End If

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHAMA_REL"
End Sub

Sub BUSCA_PERCENTUAL_COMISSAO(TIPO As String)
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select VALOR from COMISSAOPARAM"
   SQL = SQL & " where tipo = '" & Trim(TIPO) & "'"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         If TabConsulta.Fields(0).Value > 0 Then _
            Perc_Comissao_N = TabConsulta.Fields(0).Value

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_PERCENTUAL_COMISSAO"
End Sub

Function BUSCA_FATURAMENTO_QUITADO(NUMR_PEDIDO_N As Long) As Double
'On Error GoTo ERRO_TRATA

   BUSCA_FATURAMENTO_QUITADO = 0

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SqL2 = "select sum(ITEMLANCAMENTO.VALOR_ITEM) from ITEMLANCAMENTO "
   SqL2 = SqL2 & " where numr_doc = " & NUMR_PEDIDO_N
   SqL2 = SqL2 & " and status = 'B' "
   TabLancamento.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then _
      If Not IsNull(TabLancamento.Fields(0).Value) Then _
         BUSCA_FATURAMENTO_QUITADO = 0 & TabLancamento.Fields(0).Value
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_FATURAMENTO_QUITADO"
End Function
