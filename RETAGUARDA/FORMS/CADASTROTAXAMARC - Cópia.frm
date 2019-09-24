VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCADASTROTAXAMARC 
   Caption         =   "Cadastro de Taxa de Marcação"
   ClientHeight    =   7140
   ClientLeft      =   1635
   ClientTop       =   2910
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROTAXAMARC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   12015
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   6
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
            Caption         =   "&Atualizar"
            Key             =   "gravar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Gravar"
            Key             =   "markup"
            Description     =   "Grava Markup Fornecedor"
            Object.ToolTipText     =   "Gravar Markup do Forncedor"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9480
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
               Picture         =   "CADASTROTAXAMARC.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTAXAMARC.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTAXAMARC.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTAXAMARC.frx":BFCF
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTAXAMARC.frx":D237
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CADASTROTAXAMARC.frx":E59A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   50
      TabIndex        =   18
      Top             =   840
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Cadastro de taxa Markup"
      TabPicture(0)   =   "CADASTROTAXAMARC.frx":F734
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblvlrtaxa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblperc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbliten"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbltipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblfornec"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label19"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label20"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line3(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label25"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line3(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblEstab"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCgcCpf"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lstMarkup"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDtMovimento"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtSoma"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtVlrTaxa"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtPerc"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtDescItem"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSeq"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbTipoMercado"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtNome"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtMark_Fornec_Ata"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtMark_Fornec_Var"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdConsProd"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtProduto"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDescProd"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdMata"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmbTipoMercadoAux"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "&Alteração de Preços Avulço"
      TabPicture(1)   =   "CADASTROTAXAMARC.frx":F750
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbprod"
      Tab(1).Control(1)=   "txtprodalt"
      Tab(1).Control(2)=   "txtdescalt"
      Tab(1).Control(3)=   "opvalor"
      Tab(1).Control(4)=   "opperc"
      Tab(1).Control(5)=   "txtValor"
      Tab(1).Control(6)=   "cmbPreco"
      Tab(1).Control(7)=   "LISTA_ALT"
      Tab(1).Control(8)=   "lstProdutoPreço"
      Tab(1).Control(9)=   "Line2"
      Tab(1).Control(10)=   "lbltipoprod"
      Tab(1).Control(11)=   "lblprod"
      Tab(1).Control(12)=   "lblpreco"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Atualização Preço Por &Familia"
      TabPicture(2)   =   "CADASTROTAXAMARC.frx":F76C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdZera"
      Tab(2).Control(1)=   "txtRef"
      Tab(2).Control(2)=   "txtValorAcerto"
      Tab(2).Control(3)=   "chkCusto"
      Tab(2).Control(4)=   "chkAtacado"
      Tab(2).Control(5)=   "chkVenda"
      Tab(2).Control(6)=   "cmbFamiliaAUX"
      Tab(2).Control(7)=   "optAcre"
      Tab(2).Control(8)=   "optDesc"
      Tab(2).Control(9)=   "txtPercVlr"
      Tab(2).Control(10)=   "cmbFamilia"
      Tab(2).Control(11)=   "LISTA_PROD_GRP"
      Tab(2).Control(12)=   "Label24"
      Tab(2).Control(13)=   "Label23"
      Tab(2).Control(14)=   "lblpercvlr"
      Tab(2).Control(15)=   "lblgrp"
      Tab(2).ControlCount=   16
      Begin VB.ComboBox cmbTipoMercadoAux 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdMata 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   2800
         Picture         =   "CADASTROTAXAMARC.frx":F788
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1440
         Width           =   405
      End
      Begin VB.TextBox txtDescProd 
         Enabled         =   0   'False
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
         Left            =   4620
         MaxLength       =   100
         TabIndex        =   3
         Top             =   960
         Width           =   5145
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   4155
         Picture         =   "CADASTROTAXAMARC.frx":105C9
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   960
         Width           =   405
      End
      Begin VB.CommandButton cmdZera 
         Caption         =   "Zerar Preços"
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
         Left            =   -64800
         TabIndex        =   79
         Top             =   1000
         Width           =   1455
      End
      Begin VB.TextBox txtRef 
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
         Height          =   360
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   61
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtValorAcerto 
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
         Height          =   360
         Left            =   -69000
         TabIndex        =   60
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkCusto 
         Caption         =   "Preço Custo"
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
         Left            =   -66960
         TabIndex        =   76
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkAtacado 
         Caption         =   "Preço Atacado"
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
         Left            =   -66960
         TabIndex        =   75
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkVenda 
         Caption         =   "Preço Venda"
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
         Left            =   -66960
         TabIndex        =   74
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbFamiliaAUX 
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
         Left            =   -71160
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMark_Fornec_Var 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   10560
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtMark_Fornec_Ata 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   9120
         TabIndex        =   13
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         DataField       =   "Nome"
         Enabled         =   0   'False
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
         MaxLength       =   80
         TabIndex        =   12
         Top             =   2880
         Width           =   4695
      End
      Begin VB.OptionButton optAcre 
         Caption         =   "Acrescimo"
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
         Left            =   -64920
         TabIndex        =   66
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Decrescimo"
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
         Left            =   -64920
         TabIndex        =   65
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtPercVlr 
         Alignment       =   2  'Center
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
         Left            =   -69000
         MaxLength       =   5
         TabIndex        =   62
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbFamilia 
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
         Left            =   -73920
         TabIndex        =   59
         Top             =   465
         Width           =   3975
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
         Height          =   1935
         Left            =   9840
         TabIndex        =   53
         Top             =   360
         Width           =   1935
         Begin VB.TextBox txtVarejo 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtAtacado 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblvar 
            AutoSize        =   -1  'True
            Caption         =   "Taxa Varejo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label lblatc 
            AutoSize        =   -1  'True
            Caption         =   "Taxa Atacado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   1560
            TabIndex        =   55
            Top             =   1320
            Width           =   150
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   1560
            TabIndex        =   54
            Top             =   480
            Width           =   150
         End
      End
      Begin VB.ComboBox cmbprod 
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
         Left            =   -74640
         TabIndex        =   49
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtprodalt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   -68520
         TabIndex        =   48
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtdescalt 
         Enabled         =   0   'False
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
         Left            =   -69360
         TabIndex        =   47
         Top             =   960
         Width           =   6135
      End
      Begin VB.OptionButton opvalor 
         Caption         =   "R$"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64920
         TabIndex        =   46
         Top             =   650
         Width           =   615
      End
      Begin VB.OptionButton opperc 
         Caption         =   "%"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64920
         TabIndex        =   45
         Top             =   450
         Width           =   495
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   -64080
         MaxLength       =   5
         TabIndex        =   44
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cmbPreco 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   -72120
         TabIndex        =   43
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbTipoMercado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtDescItem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3240
         TabIndex        =   5
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox txtPerc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8280
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtVlrTaxa 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtSoma 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8280
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin MSComctlLib.ListView LISTAVENDA 
         Height          =   3825
         Left            =   -74880
         TabIndex        =   19
         Top             =   2100
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   6747
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Parcela"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Modalidade"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Prazo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Juros"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView LISTAPAGTO 
         Height          =   4065
         Left            =   -74160
         TabIndex        =   20
         Top             =   1860
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   7170
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   11289
         EndProperty
      End
      Begin MSComctlLib.ListView LISTAENTRADA 
         Height          =   4065
         Left            =   -74400
         TabIndex        =   21
         Top             =   1800
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   7170
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   11289
         EndProperty
      End
      Begin MSComctlLib.ListView LISTACFOP 
         Height          =   3225
         Left            =   -74880
         TabIndex        =   22
         Top             =   2640
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   5689
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   11289
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Instrução Fisco"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDtMovimento 
         Height          =   375
         Left            =   8280
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstMarkup 
         Height          =   2625
         Left            =   15
         TabIndex        =   15
         Top             =   3480
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   4630
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Códg."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Percentual"
            Object.Width           =   2999
         EndProperty
      End
      Begin MSComctlLib.ListView LISTA_PROD_GRP 
         Height          =   4665
         Left            =   -74950
         TabIndex        =   64
         Top             =   1440
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   8229
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Produto"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Pr.Varejo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Pr.Atacado"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Pr.Custo"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSMask.MaskEdBox txtCgcCpf 
         Height          =   345
         Left            =   720
         TabIndex        =   11
         Top             =   2880
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
      Begin MSComctlLib.ListView LISTA_ALT 
         Height          =   4785
         Left            =   -69480
         TabIndex        =   73
         Top             =   1320
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   8440
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Produto"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Vl.Atual"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "% Ajuste"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Vl.Corrigido"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "codg_prod"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lstProdutoPreço 
         Height          =   4785
         Left            =   -74950
         TabIndex        =   82
         Top             =   1320
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   8440
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Produto555"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Pr.Varejo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Pr.Atacado"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Pr.Custo"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblEstab 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   85
         Top             =   480
         Width           =   120
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   11880
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   375
         Left            =   1080
         TabIndex        =   81
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Referência:"
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
         Left            =   -74745
         TabIndex        =   78
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor = "
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
         Left            =   -69720
         TabIndex        =   77
         Top             =   480
         Width           =   750
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   11880
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Varejo"
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
         Left            =   10560
         TabIndex        =   71
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Atacado"
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
         Left            =   9120
         TabIndex        =   70
         Top             =   2640
         Width           =   1320
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
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
         TabIndex        =   69
         Top             =   2880
         Width           =   570
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social:"
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
         TabIndex        =   68
         Top             =   2880
         Width           =   1320
      End
      Begin VB.Label lblfornec 
         AutoSize        =   -1  'True
         Caption         =   "Atualização por Fornecedor"
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
         Left            =   4320
         TabIndex        =   67
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label lblpercvlr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "% ="
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
         Left            =   -69360
         TabIndex        =   63
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lblgrp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Família:"
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
         Left            =   -74760
         TabIndex        =   58
         Top             =   480
         Width           =   780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   -69480
         X2              =   -69480
         Y1              =   360
         Y2              =   6120
      End
      Begin VB.Label lbltipoprod 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Produto"
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
         Left            =   -74640
         TabIndex        =   52
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label lblprod 
         AutoSize        =   -1  'True
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
         Left            =   -69360
         TabIndex        =   51
         Top             =   480
         Width           =   810
      End
      Begin VB.Label lblpreco 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Preço"
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
         Left            =   -72120
         TabIndex        =   50
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lbltipo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Preço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Movimentação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6120
         TabIndex        =   41
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label lbliten 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sequência:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   40
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblperc 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9480
         TabIndex        =   39
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label lblvlrtaxa 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Taxa Markup = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   38
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor da Soma ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6225
         TabIndex        =   37
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Juros"
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
         Left            =   -68400
         TabIndex        =   36
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67440
         TabIndex        =   35
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Msg.Fisco:"
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
         Left            =   -74880
         TabIndex        =   34
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
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
         Left            =   -73650
         TabIndex        =   33
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
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
         Left            =   -74745
         TabIndex        =   32
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Prazo(dias)"
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
         Left            =   -66960
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblNrCon 
         AutoSize        =   -1  'True
         Caption         =   "Modalidade : "
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
         Index           =   2
         Left            =   -73560
         TabIndex        =   30
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Parcela(s)"
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
         Left            =   -74880
         TabIndex        =   29
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Tipo Entrada:"
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
         Left            =   -74040
         TabIndex        =   28
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descrição Tipo Entrada:"
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
         Left            =   -74400
         TabIndex        =   27
         Top             =   1320
         Width           =   2550
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descrição Forma Pagto:"
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
         Left            =   -74160
         TabIndex        =   26
         Top             =   1380
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Forma Pagto:"
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
         Left            =   -74640
         TabIndex        =   25
         Top             =   780
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição : "
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
         Left            =   -71760
         TabIndex        =   24
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
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
         Left            =   -73680
         TabIndex        =   23
         Top             =   720
         Width           =   885
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   240
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
            Picture         =   "CADASTROTAXAMARC.frx":10FCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":1141F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":1173B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":11B8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":11FE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":12303
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":12757
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":12A77
            Key             =   ""
         EndProperty
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
      DesignWidth     =   12015
      DesignHeight    =   7140
   End
   Begin VB.Label Label2 
      Caption         =   "Valor Taxa de Marcação"
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
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmCADASTROTAXAMARC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim CodigoFornec        As Integer
   Dim VALOR_PRECO_CUSTO   As Currency
   Dim VALOR_PRECO_VENDA   As Currency
   Dim VALOR_PRECO_atacado As Currency
   Dim Perc_Markup_Varejo  As Currency
   Dim PERC_MARKUP_ATACADO As Currency

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "", "", "", ""
   lblEstab.Caption = ESTABELECIMENTO_ID_N
   CRIA_TABELA
   
   Call CentralizaJanela(frmCADASTROTAXAMARC)
   txtDtMovimento.Text = Date

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select nome from usuario"
   SQL = SQL & " where usuario_id = " & CODG_USU_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      frmCADASTROTAXAMARC.Caption = frmCADASTROTAXAMARC.Caption & " ; " & Trim(TabUSU!NOME)
   If TabUSU.State = 1 Then _
      TabUSU.Close

   cmbTipoMercado.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo_a = 'M'"   'tipo mercado
   SQL = SQL & " order by codigo "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbTipoMercado.AddItem Trim(TabDESCR!desc_a)
      cmbTipoMercadoAux.AddItem TabDESCR!codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   cmbTipoMercado.Text = "Todos"
   cmbTipoMercadoAux.Text = 0

   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   SQL = "select indr_industria from ESTABELECIMENTO "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmpresa.EOF Then
      If TabEmpresa!INDR_INDUSTRIA = "S" Then
         INDR_TIPO_PROD = 0
         Else: INDR_TIPO_PROD = 1
      End If
   End If
   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close
   
   PreencheComboGrp cmbFamilia
   
   'totalizando varejo
   txtSoma.Text = ""
   txtVlrTaxa.Text = ""
   TOTALIZA_VARIAVEIS 1, ESTABELECIMENTO_ID_N
   txtVarejo.Text = txtVlrTaxa.Text

   'totalizando atacado
   txtSoma.Text = ""
   txtVlrTaxa.Text = ""
   TOTALIZA_VARIAVEIS 2, ESTABELECIMENTO_ID_N
   txtAtacado.Text = txtVlrTaxa.Text
   txtSoma.Text = ""
   txtVlrTaxa.Text = ""
   lstProdutoPreço.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub cmbTipoMercado_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "ESC - Sair", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoMercado_GotFocus"
End Sub

Private Sub cmbProd_Click()
'On Error GoTo ERRO_TRATA

   If cmbprod.Text <> "" Then
      SETA_GRID_PRODUTO
      cmbPreco.Enabled = True
      cmbPreco.Clear
      If Left(cmbprod.Text, 1) = 0 Then
         cmbPreco.AddItem "0 - Preço Custo "
         Else
            cmbPreco.AddItem "1 - Preço Atacado "
            cmbPreco.AddItem "2 - Preço Varejo "
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbProd_Click"
End Sub
Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   If Trim(cmbFamilia.Text) <> "" Then

      cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

      SETA_GRID_FAMILIA

      cmbPreco.Enabled = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbfamilia_Click"
End Sub

Private Sub txtPercVlr_Change()
   txtValorAcerto.Text = ""
End Sub

Private Sub cmdConsProd_Click()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtProduto.Text = SQL3
            txtProduto.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      PROCURA_PRODUTO
      Else: txtDescProd.Text = "Todos"
   End If
End Sub

Private Sub txtSeq_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtSeq.Text) = "" Then _
      If IsNumeric(txtSeq.Text) Then _
         NUMR_SEQ_N = txtSeq.Text

   If NUMR_SEQ_N <= 0 Then _
      NUMR_SEQ_N = MAX_ID("taxamarkup_id", "taxamarkup", "", "", "", "")

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "Select * from vwMARKUP "
   SQL = SQL & " where estabelecimento_ID = " & lblEstab.Caption

   If Trim(cmbTipoMercadoAux.Text) <> "" Then _
      SQL = SQL & " and tipomercado_ID = " & cmbTipoMercadoAux

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and produto_ID = " & PRODUTO_ID_N

   SQL = SQL & " and codg_taxa = " & txtSeq.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtDescItem.Text = Trim(TabTemp!Descricao)
      txtPerc.Text = Format(TabTemp!PERC_TAXA, strFormatacao2Digitos)
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

End Sub

Private Sub txtValorAcerto_Change()
   txtPercVlr.Text = ""
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      UCase (txtRef.Text)
      KeyAscii = 0

      SETA_GRID_FAMILIA
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "txtRef_KeyPress"
End Sub

Private Sub txtValorAcerto_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtValorAcerto_KeyPress"
End Sub

Private Sub txtPercVlr_KeyPress(KeyAscii As Integer)
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
   TRATA_ERROS Err.Description, Me.Name, "txtPercVlr_KeyPress"
End Sub

Private Sub cmbpreco_Click()
'On Error GoTo ERRO_TRATA

   If cmbPreco.Text <> "" Then
      txtprodalt.Enabled = True
      opperc.Enabled = True
      opvalor.Enabled = True
      LISTA_ALT.Enabled = True

      If Left(cmbprod.Text, 1) = 0 Then 'Se for Materia prima Altera Custo senao , nao altera
         If cmbPreco.ListIndex = 0 Then _
            cmbPreco.Tag = "PRECO_CUSTO"
         Else
            If cmbPreco.ListIndex = 1 Then _
               cmbPreco.Tag = "PRECO_ATACADO"
            If cmbPreco.ListIndex = 2 Then _
               cmbPreco.Tag = "PRECO_VENDA"
      End If
      txtprodalt.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbpreco_Click"
End Sub

Private Sub cmbprod_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbprod_KeyPress"
End Sub

Private Sub cmbTipoMercado_Click()
On Error Resume Next

   cmbTipoMercadoAux.ListIndex = cmbTipoMercado.ListIndex
   txtProduto.SetFocus

   txtSeq.SetFocus
End Sub

Private Sub cmbTipoMercado_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoMercado_KeyPress"
End Sub

Private Sub cmbTipoMercado_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbTipoMercado.Text) <> "" Then _
      SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbTipoMercado_LostFocus"
End Sub

Private Sub lstMarkup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView lstMarkup, ColumnHeader
End Sub

Private Sub TXTCGCCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCgcCpf.Mask = "##############"
   txtCgcCpf.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_GotFocus"
End Sub
Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_FORNEC
      txtCgcCpf.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_KeyPress"
End Sub
Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCgcCpf.PromptInclude = False
            txtCgcCpf.Text = CNPJCPF_A
            txtCgcCpf.PromptInclude = True

            If TabFORNECEDOR.State = 1 Then _
               TabFORNECEDOR.Close

            SQL = "select nome,razao_social,markup_atacado,markup_varejo, Fornecedor_id from FORNECEDOR "
            SQL = SQL & " where CGCCPF = '" & CNPJCPF_A & "'"
            TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFORNECEDOR.EOF Then
               If Trim(TabFORNECEDOR!NOME) = "" Then
                  txtNome.Text = Trim(TabFORNECEDOR!razao_social)
                  Else: txtNome.Text = Trim(TabFORNECEDOR!NOME)
               End If

               If Not IsNull(TabFORNECEDOR!markup_atacado) Then _
                  txtMark_Fornec_Ata.Text = TabFORNECEDOR!markup_atacado

               If Not IsNull(TabFORNECEDOR!markup_varejo) Then _
                  txtMark_Fornec_Var.Text = TabFORNECEDOR!markup_varejo

               CodigoFornec = TabFORNECEDOR!FORNECEDOR_ID
            End If
            If TabFORNECEDOR.State = 1 Then _
               TabFORNECEDOR.Close
         End If
         CNPJCPF_A = ""
         txtCgcCpf.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_KeyDown"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   lstProdutoPreço.Visible = False
   If SSTab1.Tab = 0 Then
      cmbTipoMercado.SetFocus
   End If
   If SSTab1.Tab = 1 Then
      cmbprod.SetFocus
      cmbprod.Clear
      cmbprod.AddItem "0 - Materia Prima "
      cmbprod.AddItem "1 - Produto Acabado "

      cmbPreco.Clear
      cmbPreco.AddItem "0 - Preço Custo "
      cmbPreco.AddItem "1 - Preço Atacado "
      cmbPreco.AddItem "2 - Preço Varejo "
      lstProdutoPreço.Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSTab1_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         If SSTab1.Tab = 0 Then _
            ATUALIZA_MARKAP
         If SSTab1.Tab = 1 Then _
            If Trim(LISTA_ALT.ListItems(1).Text) <> "" Then _
               ATUALIZA_PRECO_PRODUTOS
         If SSTab1.Tab = 2 Then _
              ATUALIZA_PRECO_GRP_PRODUTO
      Case "limpar"
         lstProdutoPreço.Visible = False
         cmbTipoMercado.Text = ""
         txtDtMovimento.Text = ""
         txtSeq.Text = ""
         txtDescItem.Text = ""
         txtPerc.Text = ""
         txtNome.Text = ""
         txtMark_Fornec_Ata.Text = ""
         txtMark_Fornec_Var.Text = ""
         txtCgcCpf.PromptInclude = False
         txtCgcCpf.Text = ""
         lstMarkup.ListItems.Clear
         LISTA_ALT.ListItems.Clear
         cmbTipoMercado.SetFocus
      Case "sair"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdMata_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtSeq.Text) <> "" Then
      If IsNumeric(txtSeq.Text) Then
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

         SQL = "select lancamento_id from LANCAMENTO "

         SQL = SQL & " where numr_doc = " & NUMR_SEQ_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and tipo_lancamento = " & SINAL_INDICADOR_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then
            If Not IsNull(TabLancamento.Fields(0).Value) Then

               Msg = "Confirma Exclusão do Item =  ?" & txtSeq.Text
               Style = vbYesNo + 32
               Title = "Atenção !!!"
               Help = "DEMO.HLP"
               Ctxt = 1000
               RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
               If RESPOSTA = vbYes Then

                  SQL = "delete from ITEMLANCAMENTO "
                  SQL = SQL & " where lancamento_id = " & TabLancamento.Fields(0).Value
                  SQL = SQL & " and seq = " & txtSeq.Text
                  CONECTA_RETAGUARDA.Execute SQL

                  SETA_GRID
               End If
            End If
         End If
         If TabLancamento.State = 1 Then _
            TabLancamento.Close

'         BUSCA_LANCAMENTO
      End If
   End If
   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMata_Click"
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         MATA_SEQUENCIA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEQ_KeyDown"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescItem.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEQ_KeyPress"
End Sub

Private Sub txtDescIteM_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPerc.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescIteM_KeyPress"
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_TUDO
      cmbTipoMercado.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPerc_KeyPress"
End Sub

Private Sub txtprodalt_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.optPA = True
         frmProdutoConsulta.optMP = False
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtprodalt.Text = SQL3
            txtprodalt.SetFocus
         End If
         SQL3 = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprodalt_KeyDown"
End Sub

Private Sub txtprodalt_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
       If txtprodalt.Text <> "" Then
         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "Select * from PRODUTO"
         SQL = SQL & " where codg_prod = '" & txtprodalt.Text & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         If cmbprod.Text <> "" Then
            SQL = SQL & " and tipo_prod = " & Left(cmbprod.Text, 1)
            Else: SQL = SQL & " and tipo_prod = " & Left(cmbFamilia.Text, 1)
         End If
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            txtdescalt.Text = Trim(TabProduto!Descricao)
            Else
               MsgBox "Produto não cadastrado."
               txtprodalt.SetFocus
               Exit Sub
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

         opperc.SetFocus
       End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtprodalt_KeyPress"
End Sub

Private Sub opperc_Click()
'On Error GoTo ERRO_TRATA

   txtValor.Enabled = True
   txtValor.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "opperc_Click"
End Sub

Private Sub opvalor_Click()
'On Error GoTo ERRO_TRATA

   txtValor.Enabled = True
   txtValor.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "opvalor_Click"
End Sub

Private Sub optDesc_Click()
'On Error GoTo ERRO_TRATA

   txtPercVlr.Enabled = True
   txtPercVlr.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optdesc_Click"
End Sub

Private Sub optacre_Click()
'On Error GoTo ERRO_TRATA

   txtPercVlr.Enabled = True
   txtPercVlr.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "optacre_Click"
End Sub

Private Sub opPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "opPerc_KeyPress"
End Sub

Private Sub opvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "opvalor_KeyPress"
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   Dim VALOR_CUSTO_PRODUTO_N As Double
   Dim VALOR_ATACADO_PRODUTO_N As Double
   Dim VALOR_VAREJO_PRODUTO_N As Double
   Dim VALOR_DIGITADO_N As Double
   Dim PERC_CUSTO_N As Double
   Dim PERC_ATACADO_N As Double
   Dim PERC_VAREJO_N As Double
   
   VALOR_CUSTO_PRODUTO_N = 0
   VALOR_ATACADO_PRODUTO_N = 0
   VALOR_VAREJO_PRODUTO_N = 0
   VALOR_DIGITADO_N = 0
   PERC_CUSTO_N = 0
   PERC_ATACADO_N = 0
   PERC_VAREJO_N = 0
   
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtValor.Text <> "" Then
         If TabProduto.State = 1 Then _
            TabProduto.Close

         If INDR_TIPO_PROD = 1 Then
            SQL = "select * from PRODUTO "
            SQL = SQL & " where codg_prod = '" & txtprodalt.Text & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            Else
               SQL = "select * from PRODUTO,MP "
               SQL = SQL & " where codg_mp = codg_prod "
               SQL = SQL & " and codg_pa = '" & txtprodalt.Text & "'"
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         End If
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            VALOR_CUSTO_PRODUTO_N = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
            VALOR_ATACADO_PRODUTO_N = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
            VALOR_VAREJO_PRODUTO_N = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
            VALOR_DIGITADO_N = txtValor.Text
            
            If Left(cmbPreco.Text, 1) = 0 Then _
               PERC_CUSTO_N = Format(VALOR_CUSTO_PRODUTO_N * (VALOR_DIGITADO_N / 100), strFormatacao2Digitos)
            If Left(cmbPreco.Text, 1) = 1 Then _
               PERC_ATACADO_N = Format(VALOR_ATACADO_PRODUTO_N * (VALOR_DIGITADO_N / 100), strFormatacao2Digitos)
            If Left(cmbPreco.Text, 1) = 2 Then _
               PERC_VAREJO_N = Format(VALOR_VAREJO_PRODUTO_N * (VALOR_DIGITADO_N / 100), strFormatacao2Digitos)
            
            If INDR_TIPO_PROD = 1 Then
               SQL = "update PRODUTO set "
               If opvalor.Value = True Then
                  If Left(cmbPreco.Text, 1) = 0 Then _
                     SQL = SQL & "PRECO_CUSTO = " & Replace(VALOR_DIGITADO_N + VALOR_CUSTO_PRODUTO_N, ",", ".")
                  If Left(cmbPreco.Text, 1) = 1 Then _
                     SQL = SQL & "PRECO_ATACADO = " & Replace(VALOR_DIGITADO_N + VALOR_ATACADO_PRODUTO_N, ",", ".")
                  If Left(cmbPreco.Text, 1) = 2 Then _
                     SQL = SQL & "PRECO_VENDA = " & Replace(VALOR_DIGITADO_N + VALOR_VAREJO_PRODUTO_N, ",", ".")
               End If
               If opperc.Value = True Then
                  If Left(cmbPreco.Text, 1) = 0 Then _
                     SQL = SQL & "PRECO_CUSTO = " & Replace(PERC_CUSTO_N + VALOR_CUSTO_PRODUTO_N, ",", ".")
                  If Left(cmbPreco.Text, 1) = 1 Then _
                     SQL = SQL & "PRECO_ATACADO = " & Replace(PERC_ATACADO_N + VALOR_ATACADO_PRODUTO_N, ",", ".")
                  If Left(cmbPreco.Text, 1) = 2 Then _
                     SQL = SQL & "PRECO_VENDA = " & Replace(PERC_VAREJO_N + VALOR_VAREJO_PRODUTO_N, ",", ".")
               End If
               SQL = SQL & " where codg_prod = '" & txtprodalt.Text & "'"
               SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
               CONECTA_RETAGUARDA.Execute SQL
               Else
                  SQL = "update PRODUTO,MP set "
                  If opvalor.Value = True Then
                     If Left(cmbPreco.Text, 1) = 0 Then _
                        SQL = SQL & "PRECO_CUSTO = " & Replace(VALOR_DIGITADO_N, ",", ".")
                     If Left(cmbPreco.Text, 1) = 1 Then _
                        SQL = SQL & "PRECO_ATACADO = " & Replace(VALOR_DIGITADO_N, ",", ".")
                     If Left(cmbPreco.Text, 1) = 2 Then _
                        SQL = SQL & "PRECO_VENDA = " & Replace(VALOR_DIGITADO_N, ",", ".")
                  End If
                  If opperc.Value = True Then
                     If Left(cmbPreco.Text, 1) = 0 Then _
                        SQL = SQL & "PRECO_CUSTO = " & Replace(PERC_CUSTO_N + VALOR_CUSTO_PRODUTO_N, ",", ".")
                     If Left(cmbPreco.Text, 1) = 1 Then _
                        SQL = SQL & "PRECO_ATACADO = " & Replace(PERC_ATACADO_N + VALOR_ATACADO_PRODUTO_N, ",", ".")
                     If Left(cmbPreco.Text, 1) = 2 Then _
                        SQL = SQL & "PRECO_VENDA = " & Replace(PERC_VAREJO_N + VALOR_VAREJO_PRODUTO_N, ",", ".")
                  End If
                  SQL = SQL & " where codg_mp = codg_prod "
                  SQL = SQL & " and codg_pa = '" & txtprodalt.Text & "'"
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
            End If
            
         End If
         If TabProduto.State = 1 Then _
            TabProduto.Close

         SETA_GRID_alteração

         txtprodalt.Text = ""
         txtdescalt.Text = ""
         txtValor.Text = ""
         txtprodalt.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyPress"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstMarkup.ListItems.Clear
   CONT_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwMARKUP"
   SQL = SQL & " where ESTABELECIMENTO_id = " & ESTABELECIMENTO_ID_N

   If Trim(cmbTipoMercadoAux.Text) <> "" Then _
      If IsNumeric(cmbTipoMercadoAux.Text) Then _
         SQL = SQL & " and tipomercado_id = " & cmbTipoMercadoAux.Text

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CONT_N = CONT_N + 1
      Set Item = lstMarkup.ListItems.Add(, "seq." & CONT_N, TabTemp!CODG_TAXA)
      Item.SubItems(1) = Trim(TabTemp!Descricao)
      Item.SubItems(2) = Format(TabTemp!PERC_TAXA, strFormatacao2Digitos)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   TOTALIZA_VARIAVEIS cmbTipoMercadoAux.Text, ESTABELECIMENTO_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub SETA_GRID_PRODUTO()
'On Error GoTo ERRO_TRATA

   lstProdutoPreço.ListItems.Clear
   NUMR_SEQ_N = 0

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO"
   SQL = SQL & " where tipo_prod = " & Left(cmbprod.Text, 1)
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set Item = lstProdutoPreço.ListItems.Add(, "seq." & NUMR_SEQ_N, TabProduto!CODG_PRODUTO & "-" & Trim(TabProduto!Descricao))
      Item.SubItems(1) = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
      Item.SubItems(2) = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      Item.SubItems(3) = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PRODUTO"
End Sub

Private Sub SETA_GRID_FAMILIA()
'On Error GoTo ERRO_TRATA

   LISTA_PROD_GRP.ListItems.Clear
   LISTA_PROD_GRP.Refresh

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N

   If Trim(cmbFamilia.Text) <> "" Then _
      SQL = SQL & " and FAMILIAPRODUTO_ID = " & cmbFamiliaAUX.Text

   If Trim(txtRef.Text) <> "" Then _
      SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      Set Item = LISTA_PROD_GRP.ListItems.Add(, "seq." & TabProduto.Fields("produto_id").Value, TabProduto!CODG_PRODUTO & "-" & Trim(TabProduto!Descricao))
      Item.SubItems(1) = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
      Item.SubItems(2) = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      Item.SubItems(3) = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_FAMILIA"
End Sub

Private Sub SETA_GRID_alteração()
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO"
   SQL = SQL & " where codg_prod = '" & txtprodalt.Text & "'"
   If cmbprod.Text <> "" Then
      SQL = SQL & " and tipo_prod = " & Left(cmbprod.Text, 1)
      Else: SQL = SQL & " and tipo_prod = " & Left(cmbFamilia.Text, 1)
   End If
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set Item = LISTA_ALT.ListItems.Add(, "seq." & NUMR_SEQ_N, TabProduto!Codg_Prod & " - " & Trim(TabProduto!Descricao))

      VALOR_ITEM_N = 0 & txtValor.Text
      Item.SubItems(2) = Format(VALOR_ITEM_N, strFormatacao2Digitos)

      If Left(cmbPreco.Text, 1) = 0 Then
         Item.SubItems(1) = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
         Item.SubItems(3) = Format((TabProduto!PRECO_CUSTO * VALOR_ITEM_N / 100) + TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
      End If
      If Left(cmbPreco.Text, 1) = 1 Then
         Item.SubItems(1) = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
         Item.SubItems(3) = Format((TabProduto!PRECO_ATACADO * VALOR_ITEM_N / 100) + TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      End If
      If Left(cmbPreco.Text, 1) = 2 Then
         Item.SubItems(1) = Format(TabProduto!PRECO_VENDA, strFormatacao2Digitos)
         Item.SubItems(3) = Format((TabProduto!PRECO_VENDA * VALOR_ITEM_N / 100) + TabProduto!PRECO_VENDA, strFormatacao2Digitos)
      End If
      Item.SubItems(4) = TabProduto!Codg_Prod
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_alteração("
End Sub

Private Sub TOTALIZA_VARIAVEIS(TIPO_MERC As Long, ESTAB_ID As Long)
'On Error GoTo ERRO_TRATA

   Dim VALOR_TAXA_MARC_N As Double
   VALOR_TOTAL_N = 0
   VALOR_TAXA_MARC_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select sum(perc_taxa) from vwMARKUP"
   SQL = SQL & " where ESTABELECIMENTO_ID = " & ESTAB_ID

   If TIPO_MERC > 0 Then _
      SQL = SQL & " and TAXAMARKUP_ID = " & TIPO_MERC

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields(0).Value) Then
         'Calculo de Valores da Taxa de Marcacao
         VALOR_TOTAL_N = 100 - TabTemp.Fields(0).Value
         VALOR_TAXA_MARC_N = 100 / VALOR_TOTAL_N
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtSoma.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtVlrTaxa.Text = Format(VALOR_TAXA_MARC_N, strFormatacao2Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TOTALIZA_VARIAVEIS"
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   If txtPerc.Text <> "" And txtSeq.Text <> "" And cmbTipoMercado.Text <> "" Then
      'Adciona Itens na compsicao de taxa da marcacao
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from vwMARKUP"
      SQL = SQL & " where codg_taxa = " & txtSeq.Text
      SQL = SQL & " and TAXAMARKUP_ID = " & Left(cmbTipoMercado.Text, 1)
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         SQL = "update TAXAMARKUP set "
         SQL = SQL & "TAXAMARKUP_ID = " & Left(cmbTipoMercado.Text, 1)
         SQL = SQL & ",codg_taxa = " & txtSeq.Text
         SQL = SQL & ",descricao = '" & txtDescItem.Text & "'"
         SQL = SQL & ",perc_taxa = " & tpMOEDA(txtPerc.Text)
         SQL = SQL & " where codg_taxa = " & txtSeq.Text & " and TAXAMARKUP_ID = " & Left(cmbTipoMercado.Text, 1)
         Else
            SQL = "insert into TAXAMARKUP values ("
            SQL = SQL & ESTABELECIMENTO_ID_N
            SQL = SQL & "," & Left(cmbTipoMercado.Text, 1)
            SQL = SQL & "," & txtSeq.Text
            SQL = SQL & "," & "'" & txtDescItem.Text & "'"
            SQL = SQL & "," & tpMOEDA(txtPerc.Text)
            SQL = SQL & ")"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
      CONECTA_RETAGUARDA.Execute SQL

      txtSeq.Text = ""
      txtDescItem.Text = ""
      txtPerc.Text = ""

      SETA_GRID

      'totalizando varejo
      txtSoma.Text = ""
      txtVlrTaxa.Text = ""
      TOTALIZA_VARIAVEIS 1
      txtVarejo.Text = txtVlrTaxa.Text
   
      'totalizando atacado
      txtSoma.Text = ""
      txtVlrTaxa.Text = ""
      TOTALIZA_VARIAVEIS 2
      txtAtacado.Text = txtVlrTaxa.Text
      txtSoma.Text = ""
      txtVlrTaxa.Text = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub ATUALIZA_MARKAP()
'On Error GoTo ERRO_TRATA

   Dim VALOR_CUSTO_MATERIA    As Double  'Custo da Materia Prima
   Dim VALOR_CUSTO_PRODUTO    As Double  'Custo do produto Acabado
   Dim VALOR_PRECO_VENDA      As Double    'Preco de Venda Produto
   Dim VLR_TAXA_MARKUP        As Double    'Valor Markup
   Dim Cont                   As Double    'Contador
   Dim VALOR_PRECO_atacado    As Double
   Dim PERC_MARKUP_ATACADO    As Double
   Dim Perc_Markup_Varejo     As Double

   VALOR_CUSTO_MATERIA = 0
   VALOR_CUSTO_PRODUTO = 0
   VALOR_PRECO_VENDA = 0
   VLR_TAXA_MARKUP = 0
   Cont = 0
   VALOR_PRECO_atacado = 0
   PERC_MARKUP_ATACADO = 0
   Perc_Markup_Varejo = 0
   Cont = 0

   'inicio
   'Fazer Calculos do preco para gravar no arquivo de produtos
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select distinct(codg_produto) from PRODUTO "
   SQL = SQL & " where situacao <> 'C' "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " order by codg_produto"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
       
      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_produto = '" & TabTemp!CODG_PRODUTO & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N

      If CodigoFornec <> 0 Then _
         SQL = SQL & " and fornecedor_id = " & CodigoFornec

      SQL = SQL & " and situacao <> 'C' "
      SQL = SQL & " and tipo_prod = 1   "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"
         If CodigoFornec <> 0 Then

            If txtMark_Fornec_Var.Text = "" Then _
               txtMark_Fornec_Var.Text = txtVarejo.Text

            If txtMark_Fornec_Ata.Text = "" Then _
               txtMark_Fornec_Ata.Text = txtAtacado.Text

            Perc_Markup_Varejo = txtMark_Fornec_Var.Text
            PERC_MARKUP_ATACADO = txtMark_Fornec_Ata.Text
            VALOR_PRECO_VENDA = Format(TabConsulta!PRECO_CUSTO * Perc_Markup_Varejo, strFormatacao2Digitos)
            VALOR_PRECO_atacado = Format(TabConsulta!PRECO_CUSTO * PERC_MARKUP_ATACADO, strFormatacao2Digitos)
            Else
               Perc_Markup_Varejo = txtVarejo.Text
               PERC_MARKUP_ATACADO = txtAtacado.Text
               VALOR_PRECO_VENDA = Format(TabConsulta!PRECO_CUSTO * Perc_Markup_Varejo, strFormatacao2Digitos)
               VALOR_PRECO_atacado = Format(TabConsulta!PRECO_CUSTO * PERC_MARKUP_ATACADO, strFormatacao2Digitos)
         End If

         SQL = "update PRODUTO set "
         SQL = SQL & "preco_venda = " & tpMOEDA(VALOR_PRECO_VENDA)
         SQL = SQL & ",preco_atacado = " & tpMOEDA(VALOR_PRECO_atacado)
         SQL = SQL & " where codg_produto = '" & Trim(TabTemp!CODG_PRODUTO) & "'"
         SQL = SQL & " and tipo_prod = 1   "
         CONECTA_RETAGUARDA.Execute SQL
      End If
      TabTemp.MoveNext

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MsgBox "Preços Atualizados com Sucesso"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_MARKAP"
End Sub

Private Sub Grava_Markup_Fornecedor()
'On Error GoTo ERRO_TRATA

   If TabFORNECEDOR.State = 1 Then _
      TabFORNECEDOR.Close

   SQL = "select * from FORNECEDOR "
   SQL = SQL & " where cgccpf = '" & txtCgcCpf.Text & "'"
   TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFORNECEDOR.EOF Then
      SQL = "update FORNECEDOR set "
      SQL = SQL & "markup_varejo  = " & Replace(txtVarejo.Text, ",", ".")
      SQL = SQL & ",markup_atacado = " & Replace(txtAtacado.Text, ",", ".")
      SQL = SQL & " where cgccpf = '" & txtCgcCpf.Text & "'"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabFORNECEDOR.State = 1 Then _
      TabFORNECEDOR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "grava_markup_fornecedor"
End Sub

Private Sub ATUALIZA_MARKAP_MATERIA()
'On Error GoTo ERRO_TRATA

   Dim VALOR_CUSTO_MATERIA As Double  'Custo da Materia Prima
   Dim VALOR_CUSTO_PRODUTO As Double  'Custo do produto Acabado
   Dim VALOR_PRECO_VENDA   As Double    'Preco de Venda Produto
   Dim VLR_TAXA_MARKUP     As Double    'Valor Markup
   Dim Cont                As Double    'Contador
   Dim VALOR_PRECO_atacado As Double
   Dim PERC_MARKUP_ATACADO As Double
   Dim Perc_Markup_Varejo  As Double
   Dim rstPreco            As Recordset

   VALOR_CUSTO_MATERIA = 0
   VALOR_CUSTO_PRODUTO = 0
   VALOR_PRECO_VENDA = 0
   VLR_TAXA_MARKUP = 0
   Cont = 0
   VALOR_PRECO_atacado = 0
   PERC_MARKUP_ATACADO = 0
   Perc_Markup_Varejo = 0
   Cont = 0

   'inicio
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select distinct(codg_pa) from MP "
   SQL = SQL & " order by codg_pa"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      'Fazer Calculos do preco para gravar no arquivo de produtos
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select sum(valor*qtde) from PRODUTO p, MP m "
      SQL = SQL & " where p.codg_produto = m.codg_pa "
      SQL = SQL & " and p.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and m.codg_pa = '" & TabTemp!codg_pa & "'"
      SQL = SQL & " and p.tipo_prod = 1   "
      SQL = SQL & " and p.situacao <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         If Not IsNull(TabConsulta.Fields(0).Value) Then
            
            'DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"
            
            Perc_Markup_Varejo = txtVarejo.Text
            PERC_MARKUP_ATACADO = txtAtacado.Text
                                    
'            SQL = "select Sum(MP.QTDE * MP.VALOR)  from MP ,PRODUTO"
'            SQL = SQL & " where codg_mp = codg_prod "
'            SQL = SQL & " and codg_pa = '" & TABTEMP!codg_pa & "'"
'            Set rstPreco = CONECTA_RETAGUARDA.OpenRecordset(SQL, 4)
'            If Not rstPreco.EOF Then
'               If Not IsNull(rstPreco.Fields(0).Value) Then
'                  VALOR_PRECO_VENDA = Format(rstPreco.Fields(0).Value * perc_markup_varejo, strFormatacao2Digitos)
'                  VALOR_PRECO_atacado = Format(rstPreco.Fields(0).Value * PERC_MARKUP_ATACADO, strFormatacao2Digitos)
'               End If
'            End If
'            rstPreco.Clone
            
            SQL = "update PRODUTO set "
            SQL = SQL & "preco_venda = " & Replace(VALOR_PRECO_VENDA, ",", ".")
            SQL = SQL & ",preco_atacado = " & Replace(VALOR_PRECO_atacado, ",", ".")
            SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and codg_produto = '" & TabTemp!codg_pa & "'"
            CONECTA_RETAGUARDA.Execute SQL
         End If
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MsgBox "Preços Atualizados com Sucesso"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_MARKAP_MATERIA"
End Sub

Private Sub ATUALIZA_PRECO_PRODUTOS()
'On Error GoTo ERRO_TRATA

   Dim PRECO_CUSTO_ANTERIOR As Double

   PRECO_CUSTO_ANTERIOR = 0
   CONT_N = 1

   While Trim(LISTA_ALT.ListItems(CONT_N).Text) <> ""
      If cmbPreco.Tag = "PRECO_CUSTO" Then
         SQL = "update PRODUTO set "
         SQL = SQL & " preco_custo_anterior  = '" & LISTA_ALT.ListItems.Item(CONT_N).SubItems(1) & "'"
         SQL = SQL & ", preco_custo = " & Replace(LISTA_ALT.ListItems.Item(CONT_N).SubItems(3), ",", ".")
         SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and codg_produto = '" & LISTA_ALT.ListItems.Item(CONT_N).SubItems(4) & "'"
         CONECTA_RETAGUARDA.Execute SQL
         Else
            SQL = "update PRODUTO set "
            SQL = SQL & cmbPreco.Tag & " = " & Replace(LISTA_ALT.ListItems.Item(CONT_N).SubItems(3), ",", ".")
            SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and codg_produto = '" & LISTA_ALT.ListItems.Item(CONT_N).SubItems(4) & "'"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      If Left(cmbprod.Text, 1) = 0 Then 'Alterando Valores da materia prima no arquivo de materia prima
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from MP "
         SQL = SQL & " where codg_mp = '" & LISTA_ALT.ListItems.Item(CONT_N).SubItems(4) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            SQL = "update MP set "
            SQL = SQL & " valor = " & Replace(LISTA_ALT.ListItems.Item(CONT_N).SubItems(3), ",", ".")
            SQL = SQL & " where codg_mp = '" & LISTA_ALT.ListItems.Item(CONT_N).SubItems(4) & "'"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      
         'Quando preco da materia prima for alterado,
         'sera atualizado o preco de custo de todos produtos relacionado
         'aquela materia prima!
         
         SQL = "select distinct(codg_pa) from MP "
         SQL = SQL & " where codg_mp = '" & LISTA_ALT.ListItems.Item(CONT_N).SubItems(4) & "'"
         SQL = SQL & " order by codg_pa"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            'Fazer Calculos do preco para gravar no arquivo de produtos
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select sum(valor*qtde) from PRODUTO p, MP m "
            SQL = SQL & " where p.codg_prod = m.codg_pa "
            SQL = SQL & " and p.empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and m.codg_pa = '" & TabTemp!codg_pa & "'"
            SQL = SQL & " and p.tipo_prod = 1   "
            SQL = SQL & " and p.situacao <> 'C' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               If Not IsNull(TabConsulta.Fields(0).Value) Then

'                  SQL = "select Sum(MP.QTDE * MP.VALOR)  from MP ,PRODUTO"
'                  SQL = SQL & " where codg_mp = codg_prod "
'                  SQL = SQL & " and codg_pa = '" & TABTEMP!codg_pa & "'"
'                  Set rstPreco = CONECTA_RETAGUARDA.OpenRecordset(SQL, 4)
'                  If Not rstPreco.EOF Then
'                     If Not IsNull(rstPreco.Fields(0).Value) Then
'                        VALOR_PRECO_CUSTO = Format(rstPreco.Fields(0).Value, strFormatacao2Digitos)
'                     End If
'                  End If
'                  rstPreco.Clone

                  'Buscar Preco Custo Anterior do produto acabado
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select produto.preco_custo from PRODUTO"
                  SQL = SQL & " where codg_prod = '" & TabTemp!codg_pa & "'"
                  SQL = SQL & " and situacao <> 'C' "
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabProduto.EOF Then _
                     PRECO_CUSTO_ANTERIOR = TabProduto!PRECO_CUSTO
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "update PRODUTO set "
                  SQL = SQL & "preco_custo_anterior = " & Replace(PRECO_CUSTO_ANTERIOR, ",", ".")
                  SQL = SQL & ", preco_custo = " & Replace(VALOR_PRECO_CUSTO, ",", ".")
                  SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
                  SQL = SQL & " and codg_produto = '" & TabTemp!codg_pa & "'"
                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
            If TabProduto.State = 1 Then _
               TabConsulta.Close

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      CONT_N = CONT_N + 1
   Wend

   MsgBox "Atualização de preço " & cmbPreco.Text & " realizada com sucesso."
   cmbPreco.Enabled = False
   txtprodalt.Enabled = False
   opperc.Enabled = False
   opvalor.Enabled = False
   LISTA_ALT.Enabled = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_PRECO_PRODUTOS"
End Sub

Private Sub PreencheComboGrp(NomeCombo As ComboBox)
'On Error GoTo ERRO_TRATA

    Dim rstGRP As New ADODB.Recordset

    SQL = "select * from FAMILIAPRODUTO order by DESCRICAO"
    rstGRP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

    NomeCombo.Clear
    If Not rstGRP.EOF Then
       'Mundando o ponteiro do mouse, para mostrar para o usuario que esta processando...
        Screen.MousePointer = vbHourglass

        rstGRP.MoveFirst
        Do Until rstGRP.EOF
            'Importantissimo
            DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"
            
            NomeCombo.AddItem Trim(rstGRP!CODG_FAMILIA) & "-" & Trim(rstGRP!Descricao)
            cmbFamiliaAUX.AddItem Trim(rstGRP.Fields("familiaproduto_id").Value)
            rstGRP.MoveNext
        Loop
    End If

    'Voltando o ponteiro do mouse para o tipo default, ponteiro.
    Screen.MousePointer = vbDefault
    rstGRP.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "preencheCombogrp"
End Sub

Private Sub cmdZera_Click()
'On Error GoTo ERRO_TRATA

   If chkVenda.Value = 1 Or chkAtacado.Value = 1 Or chkCusto.Value = 1 Then
      Else
         MsgBox "Escolha Venda/Atacado/Custo !!!"
         Exit Sub
   End If

   Msg = "Essa rotina vai zerar os preços que foram selecionados ao lado. Confirma Operação ?"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO  "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N

      If Trim(cmbFamilia.Text) <> "" Then _
         SQL = SQL & " and FAMILIAPRODUTO_ID = " & cmbFamiliaAUX.Text

      If Trim(txtRef.Text) <> "" Then _
         SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabProduto.EOF
         DoEvents

         SQL = "update PRODUTO set "
         SQL = SQL & " empresa_id = " & EMPRESA_ID_N

         If chkVenda.Value = 1 Then _
            SQL = SQL & " ,preco_venda = 0"
         If chkAtacado.Value = 1 Then _
            SQL = SQL & " ,preco_atacado = 0"
         If chkCusto.Value = 1 Then _
            SQL = SQL & " ,preco_custo = 0"

         SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         TabProduto.MoveNext
      Wend
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SETA_GRID_FAMILIA
      MsgBox "Preços Zerados com Sucesso"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdZera_Click"
End Sub

Private Sub ATUALIZA_PRECO_GRP_PRODUTO()
'On Error GoTo ERRO_TRATA

   If optDesc.Value = False And optAcre.Value = False Then
      MsgBox "Escolha <<Acressimo>> ou <<Desconto>> !!!"
      Exit Sub
   End If

   If chkVenda.Value = 1 Or chkAtacado.Value = 1 Or chkCusto.Value = 1 Then
      Else
         MsgBox "Escolha Venda/Atacado/Custo !!!"
         Exit Sub
   End If

   Dim VALOR_CUSTO_PRODUTO          As Double  'Custo do produto Acabado
   Dim PRECO_CUSTO_ANTERIOR_MATERIA As Double
   Dim PRECO_CUSTO_ANTERIOR_ACABADO As Double
   Dim PERC_VLR_N                   As Double
   Dim PERC_VLR_ATA_N               As Double
   Dim Valor_Atualiza_N             As Double
   Dim Perc_Atualiza_N              As Double

   VALOR_CUSTO_PRODUTO = 0
   VALOR_PRECO_VENDA = 0
   VALOR_PRECO_atacado = 0
   PRECO_CUSTO_ANTERIOR_MATERIA = 0
   PRECO_CUSTO_ANTERIOR_ACABADO = 0
   CONT_N = 0
   PERC_VLR_N = 0
   PERC_VLR_ATA_N = 0
   Valor_Atualiza_N = 0
   Perc_Atualiza_N = 0
   VALOR_ITEM_N = 0

   If Trim(txtValorAcerto.Text) <> "" Then _
      If IsNumeric(txtValorAcerto.Text) Then _
         Valor_Atualiza_N = txtValorAcerto.Text

   If Trim(txtPercVlr.Text) <> "" Then _
      If IsNumeric(txtPercVlr.Text) Then _
         Perc_Atualiza_N = txtPercVlr.Text

   If Valor_Atualiza_N > 0 Then
      VALOR_ITEM_N = Valor_Atualiza_N
      Else
         If Perc_Atualiza_N > 0 Then _
            VALOR_ITEM_N = Perc_Atualiza_N
   End If

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from PRODUTO  "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N

   If Trim(cmbFamilia.Text) <> "" Then _
      SQL = SQL & " and FAMILIAPRODUTO_ID = " & cmbFamiliaAUX.Text

   If Trim(txtRef.Text) <> "" Then _
      SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      DoEvents

      If optDesc.Value = True Then 'BAIXANDO O PREÇO
         If Valor_Atualiza_N > 0 Then
            SQL = "update PRODUTO set "
            SQL = SQL & " empresa_id = " & EMPRESA_ID_N

            If chkVenda.Value = 1 Then
               SQL = SQL & " ,preco_varejo_anterior = preco_venda "
               SQL = SQL & " ,preco_venda = preco_venda - " & tpMOEDA(Valor_Atualiza_N)
            End If
            If chkAtacado.Value = 1 Then
               SQL = SQL & " ,preco_atacado_anterior = preco_atacado "
               SQL = SQL & " ,preco_atacado = preco_atacado - " & tpMOEDA(Valor_Atualiza_N)
            End If
            If chkCusto.Value = 1 Then
               SQL = SQL & " ,preco_custo_anterior = preco_custo "
               SQL = SQL & " ,preco_custo = preco_custo - " & tpMOEDA(Valor_Atualiza_N)
            End If

            SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            CONECTA_RETAGUARDA.Execute SQL
            Else
               If Perc_Atualiza_N > 0 Then
                  SQL = "update PRODUTO set "
                  SQL = SQL & " empresa_id = " & EMPRESA_ID_N

                  If chkVenda.Value = 1 Then
                     SQL = SQL & " ,preco_varejo_anterior = preco_venda "
                     SQL = SQL & " ,preco_venda = preco_venda - " & tpMOEDA(TabProduto.Fields("preco_venda").Value * Perc_Atualiza_N / 100)
                  End If
                  If chkAtacado.Value = 1 Then
                     SQL = SQL & " ,preco_atacado_anterior = preco_atacado "
                     SQL = SQL & " ,preco_atacado = preco_atacado - " & tpMOEDA(TabProduto.Fields("preco_atacado").Value * Perc_Atualiza_N / 100)
                  End If
                  If chkCusto.Value = 1 Then
                     SQL = SQL & " ,preco_custo_anterior = preco_custo "
                     SQL = SQL & " ,preco_custo = preco_custo  - " & tpMOEDA(TabProduto.Fields("preco_custo").Value * Perc_Atualiza_N / 100)
                  End If

                  SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
         End If
      End If
      If optAcre.Value = True Then 'Esta aumentando
         If Valor_Atualiza_N > 0 Then
            SQL = "update PRODUTO set "
            SQL = SQL & " empresa_id = " & EMPRESA_ID_N

            If chkVenda.Value = 1 Then
               SQL = SQL & " ,preco_varejo_anterior = preco_venda "
               SQL = SQL & " ,preco_venda = preco_venda + " & tpMOEDA(Valor_Atualiza_N)
            End If
            If chkAtacado.Value = 1 Then
               SQL = SQL & " ,preco_atacado_anterior = preco_atacado "
               SQL = SQL & " ,preco_atacado = preco_atacado + " & tpMOEDA(Valor_Atualiza_N)
            End If
            If chkCusto.Value = 1 Then
               SQL = SQL & " ,preco_custo_anterior = preco_custo "
               SQL = SQL & " ,preco_custo = preco_custo + " & tpMOEDA(Valor_Atualiza_N)
            End If

            SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            CONECTA_RETAGUARDA.Execute SQL
            Else
               If Perc_Atualiza_N > 0 Then
                  SQL = "update PRODUTO set "
                  SQL = SQL & " empresa_id = " & EMPRESA_ID_N

                  If chkVenda.Value = 1 Then
                     SQL = SQL & " ,preco_varejo_anterior = preco_venda "
                     SQL = SQL & " ,preco_venda = preco_venda + " & tpMOEDA(TabProduto.Fields("preco_venda").Value * Perc_Atualiza_N / 100)
                  End If
                  If chkAtacado.Value = 1 Then
                     SQL = SQL & " ,preco_atacado_anterior = preco_atacado "
                     SQL = SQL & " ,preco_atacado = preco_atacado + " & tpMOEDA(TabProduto.Fields("preco_atacado").Value * Perc_Atualiza_N / 100)
                  End If
                  If chkCusto.Value = 1 Then
                     SQL = SQL & " ,preco_custo_anterior = preco_custo "
                     SQL = SQL & " ,preco_custo = preco_custo  + " & tpMOEDA(TabProduto.Fields("preco_custo").Value * Perc_Atualiza_N / 100)
                  End If
      
                  SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
                  CONECTA_RETAGUARDA.Execute SQL
               End If
         End If
      End If

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SETA_GRID_FAMILIA

   MsgBox "Preços Atualizados com Sucesso"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_PRECO_GRP_PRODUTO"
End Sub

Private Sub PROCURA_FORNEC()
'On Error GoTo ERRO_TRATA

   txtCgcCpf.PromptInclude = False

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select * from FORNECEDOR "
   SQL = SQL & " where CGCCPF = '" & txtCgcCpf.Text & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      txtNome.Text = TabCliente!NOME
      CodigoFornec = TabCliente!FORNECEDOR_ID

      If Not IsNull(TabCliente!markup_atacado) Then _
         txtMark_Fornec_Ata.Text = TabCliente!markup_atacado

      If Not IsNull(TabCliente!markup_varejo) Then _
         txtMark_Fornec_Var.Text = TabCliente!markup_varejo
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_FORNEC"
End Sub

Sub PROCURA_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select produto_id,descricao from PRODUTO "
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         txtDescProd.Text = TabConsulta.Fields("descricao").Value
         PRODUTO_ID_N = TabConsulta.Fields("produto_id").Value
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_PRODUTO"
End Sub

Sub MATA_SEQUENCIA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwMARKUP"
   SQL = SQL & " where codg_taxa = " & txtSeq.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      MsgBox "Registro não encontrado.", vbOKOnly, "Atenção !!!"
      txtSeq.SetFocus
      Exit Sub
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   Msg = "Confirma Exclusão?"
   Style = vbYesNo + 32
   Title = "Atenção !!!"
   Help = "DEMO.HLP"
   Ctxt = 1000
   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
   If RESPOSTA = vbYes Then
      SQL = "delete  from vwMARKUP"
      SQL = SQL & " where codg_taxa = " & txtSeq.Text
      txtSeq.Text = ""
      txtDescItem.Text = ""
      txtPerc.Text = ""
      CONECTA_RETAGUARDA.Execute SQL
      SETA_GRID
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MATA_SEQUENCIA"
End Sub

Sub CRIA_TABELA()
'On Error GoTo ERRO_TRATA

   If ExisteTabela("TAXAMARKUP", "U") = False Then
      SQL = "CREATE TABLE [dbo].[TAXAMARKUP]("
      SQL = SQL & " [TAXAMARKUP_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [TIPOMERCADO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_TAXAMARKUP] PRIMARY KEY CLUSTERED([TAXAMARKUP_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TAXAMARKUP]  WITH CHECK ADD  CONSTRAINT [FK_TAXAMARKUP_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TAXAMARKUP] CHECK CONSTRAINT [FK_TAXAMARKUP_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TAXAMARKUP]  WITH CHECK ADD  CONSTRAINT [FK_TAXAMARKUP_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TAXAMARKUP] CHECK CONSTRAINT [FK_TAXAMARKUP_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("TAXAMARKUPITEM", "U") = False Then
      SQL = "CREATE TABLE [dbo].[TAXAMARKUPITEM]("
      SQL = SQL & " [TAXAMARKUP_ID] [bigint] NOT NULL,"
      SQL = SQL & " [TAXAMARKUPITEM_ID] [bigint] NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](200) NOT NULL,"
      SQL = SQL & " [PERC_TAXA] [Float] not null"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TAXAMARKUPITEM]  WITH CHECK ADD  CONSTRAINT [FK_TAXAMARKUPITEM_TAXAMARKUP] FOREIGN KEY([TAXAMARKUP_ID])"
      SQL = SQL & " References [dbo].[TAXAMARKUP]([TAXAMARKUP_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[TAXAMARKUPITEM] CHECK CONSTRAINT [FK_TAXAMARKUPITEM_TAXAMARKUP]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("vwMARKUP", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP VIEW [dbo].[vwMARKUP]"

   SQL = "CREATE VIEW [dbo].[vwMARKUP]"
   SQL = SQL & " AS "
   SQL = SQL & " SELECT TAXAMARKUP.*, TAXAMARKUPITEM.TAXAMARKUPITEM_ID, TAXAMARKUPITEM.DESCRICAO, "
   SQL = SQL & " TAXAMARKUPITEM.PERC_TAXA "
   SQL = SQL & " from vwMARKUP "
   SQL = SQL & " INNER JOIN TAXAMARKUPITEM "
   SQL = SQL & " ON TAXAMARKUP.TAXAMARKUP_ID = TAXAMARKUPITEM.TAXAMARKUP_ID"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CRIA_TABELA"
End Sub
