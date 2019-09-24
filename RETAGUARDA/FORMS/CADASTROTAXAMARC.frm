VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROTAXAMARC 
   Caption         =   "Cadastro de Taxa de Marcação"
   ClientHeight    =   7725
   ClientLeft      =   1635
   ClientTop       =   2910
   ClientWidth     =   11970
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
   ScaleHeight     =   7725
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   1270
      ButtonWidth     =   2540
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
            Key             =   "at"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
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
      Height          =   6975
      Left            =   45
      TabIndex        =   13
      Top             =   750
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   12303
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
      Tab(0).Control(0)=   "lblperc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbltipo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblfornec"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label21"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line3(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label25"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line3(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblEstab"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbliten"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label26"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblatc"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblvar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label18"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "MSFlexGrid1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lstFamilia"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lstMercado"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCNPJCPF_FORNEC"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDtMovimento"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPerc"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtOcorrencia"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtNome"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtMark_Fornec_Ata"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtMark_Fornec_Var"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdConsProd"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtProduto"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtDescProd"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtValorDig"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtAtacado"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtVarejo"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtBKPAtacado"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtBKPVarejo"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtCusto"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "&Alteração de Preços Avulço"
      TabPicture(1)   =   "CADASTROTAXAMARC.frx":F750
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblpreco"
      Tab(1).Control(1)=   "lblprod"
      Tab(1).Control(2)=   "lbltipoprod"
      Tab(1).Control(3)=   "Line2"
      Tab(1).Control(4)=   "lstProdutoPreço"
      Tab(1).Control(5)=   "LISTA_ALT"
      Tab(1).Control(6)=   "cmbPreco"
      Tab(1).Control(7)=   "txtValor"
      Tab(1).Control(8)=   "opperc"
      Tab(1).Control(9)=   "opvalor"
      Tab(1).Control(10)=   "txtdescalt"
      Tab(1).Control(11)=   "txtprodalt"
      Tab(1).Control(12)=   "cmbprod"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Atualização Preço Por &Familia"
      TabPicture(2)   =   "CADASTROTAXAMARC.frx":F76C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblgrp"
      Tab(2).Control(1)=   "lblpercvlr"
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(3)=   "Label24"
      Tab(2).Control(4)=   "LISTA_PROD_GRP"
      Tab(2).Control(5)=   "cmbFamilia"
      Tab(2).Control(6)=   "txtPercVlr"
      Tab(2).Control(7)=   "optDesc"
      Tab(2).Control(8)=   "optAcre"
      Tab(2).Control(9)=   "cmbFamiliaAUX"
      Tab(2).Control(10)=   "chkVenda"
      Tab(2).Control(11)=   "chkAtacado"
      Tab(2).Control(12)=   "chkCusto"
      Tab(2).Control(13)=   "txtValorAcerto"
      Tab(2).Control(14)=   "txtRef"
      Tab(2).Control(15)=   "cmdZera"
      Tab(2).ControlCount=   16
      Begin VB.TextBox txtCusto 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   10320
         TabIndex        =   83
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtBKPVarejo 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   10320
         TabIndex        =   81
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtBKPAtacado 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   7320
         TabIndex        =   79
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtVarejo 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   4080
         TabIndex        =   78
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtAtacado 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   1560
         TabIndex        =   75
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   9960
         TabIndex        =   74
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDescProd 
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
         Height          =   375
         Left            =   3960
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txtProduto 
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
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdConsProd 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3435
         Picture         =   "CADASTROTAXAMARC.frx":F788
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1560
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
         TabIndex        =   66
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   59
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   10
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtMark_Fornec_Ata 
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
         Height          =   375
         Left            =   9120
         TabIndex        =   9
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         DataField       =   "Nome"
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
         Height          =   375
         Left            =   4200
         MaxLength       =   80
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   49
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
         TabIndex        =   46
         Top             =   465
         Width           =   3975
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtOcorrencia 
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
         Left            =   1560
         TabIndex        =   3
         Top             =   2040
         Width           =   6015
      End
      Begin VB.TextBox txtPerc 
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
         Left            =   7680
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin MSComctlLib.ListView LISTAVENDA 
         Height          =   3825
         Left            =   -74880
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         Left            =   10320
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin MSComctlLib.ListView LISTA_PROD_GRP 
         Height          =   4665
         Left            =   -74950
         TabIndex        =   51
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
      Begin MSMask.MaskEdBox txtCNPJCPF_FORNEC 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
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
      Begin MSComctlLib.ListView LISTA_ALT 
         Height          =   4785
         Left            =   -69480
         TabIndex        =   60
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
         TabIndex        =   69
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
      Begin MSComctlLib.ListView lstMercado 
         Height          =   735
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1296
         View            =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   176
         EndProperty
      End
      Begin MSComctlLib.ListView lstFamilia 
         Height          =   735
         Left            =   3240
         TabIndex        =   1
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   1296
         View            =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Forma Pagto."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   176
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3735
         Left            =   45
         TabIndex        =   73
         Top             =   3120
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6588
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Vlr.Custo="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9300
         TabIndex        =   84
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "MKP Varejo = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9000
         TabIndex        =   82
         Top             =   2520
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "MKP Atacado = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5760
         TabIndex        =   80
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label lblvar 
         Alignment       =   1  'Right Justify
         Caption         =   "% Varejo = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3000
         TabIndex        =   77
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label lblatc 
         Alignment       =   1  'Right Justify
         Caption         =   "% Atacado = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   330
         TabIndex        =   76
         Top             =   2520
         Width           =   1230
      End
      Begin VB.Label Label26 
         Caption         =   "Familia Produto"
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
         Left            =   3240
         TabIndex        =   72
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lbliten 
         Alignment       =   1  'Right Justify
         Caption         =   "Ocorrência:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   71
         Top             =   2040
         Width           =   1245
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
         TabIndex        =   70
         Top             =   120
         Width           =   120
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   11880
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   435
         TabIndex        =   68
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   65
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   64
         Top             =   480
         Width           =   750
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   11880
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label22 
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
         TabIndex        =   58
         Top             =   3120
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label21 
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
         TabIndex        =   57
         Top             =   3120
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label20 
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
         TabIndex        =   56
         Top             =   3360
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label19 
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
         TabIndex        =   55
         Top             =   3360
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblfornec 
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4320
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblpercvlr 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   50
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lblgrp 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label lblprod 
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
         TabIndex        =   43
         Top             =   480
         Width           =   810
      End
      Begin VB.Label lblpreco 
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
         TabIndex        =   42
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lbltipo 
         Caption         =   "Tipo Mercado"
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
         Left            =   165
         TabIndex        =   34
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Dt.Mov.:"
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
         Left            =   9360
         TabIndex        =   33
         Top             =   1560
         Width           =   885
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
         Left            =   8880
         TabIndex        =   32
         Top             =   2040
         Width           =   195
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
            Picture         =   "CADASTROTAXAMARC.frx":1018A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":105DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":108FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":10D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":111A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":114C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":11916
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAXAMARC.frx":11C36
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
      DesignWidth     =   11970
      DesignHeight    =   7725
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
      TabIndex        =   12
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
   Dim VALOR_PRECO_CUSTO   As Currency
   Dim VALOR_PRECO_VENDA   As Currency
   Dim VALOR_PRECO_atacado As Currency
   Private LastRow         As Long ' Ultima linha em que se editou
   Private LastCol         As Long ' ultima coluna em que se editou
   Dim i, Selecao_Mercado, Selecao_Familia

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "", "", "", ""
   lblEstab.Caption = ESTABELECIMENTO_ID_N

   CRIA_TABELA
   LIMPA_TUDO
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub lstFamilia_Click()
'On Error GoTo ERRO_TRATA

   Selecao_Familia = ""
   INDR_PRI = True

   If lstFamilia.ListItems.Count > 0 Then
      For i = lstFamilia.ListItems.Count To 1 Step -1
         If lstFamilia.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               Selecao_Familia = lstFamilia.ListItems(i).SubItems(1)
               Else: Selecao_Familia = Selecao_Familia & "," & lstFamilia.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstFamilia_Click"
End Sub

Private Sub lstMercado_Click()
'On Error GoTo ERRO_TRATA

   Selecao_Mercado = ""
   INDR_PRI = True

   If lstMercado.ListItems.Count > 0 Then
      For i = lstMercado.ListItems.Count To 1 Step -1
         If lstMercado.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               Selecao_Mercado = lstMercado.ListItems(i).SubItems(1)
               Else: Selecao_Mercado = Selecao_Mercado & "," & lstMercado.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstMercado_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         If SSTab1.Tab = 0 Then _
            GRAVA_TUDO
      Case "at"
         If SSTab1.Tab = 0 Then _
            SETA_GRID
         If SSTab1.Tab = 1 Then _
            If Trim(LISTA_ALT.ListItems(1).Text) <> "" Then _
               ATUALIZA_PRECO_PRODUTOS
         If SSTab1.Tab = 2 Then _
              ATUALIZA_PRECO_GRP_PRODUTO
      Case "limpar"
         If SSTab1.Tab = 0 Then _
            LIMPA_TUDO
      Case "sair"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   lstProdutoPreço.Visible = False
   If SSTab1.Tab = 0 Then
      
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

Private Sub txtOcorrencia_GotFocus()
   txtOcorrencia.SelStart = 0
   txtOcorrencia.SelLength = Len(txtOcorrencia.Text)
End Sub

Private Sub txtPerc_GotFocus()
   txtPerc.SelStart = 0
   txtPerc.SelLength = Len(txtPerc.Text)
End Sub

Private Sub txtPerc_LostFocus()
   If Trim(txtPerc.Text) <> "" Then _
      txtPerc.Text = "" & Format(txtPerc.Text, strFormatacao2Digitos)
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

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Private Sub txtProduto_GotFocus()
   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto.Text)
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtOcorrencia.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyPress"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TXTPRODUTO_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) <> "" Then
      MOSTRA_PRODUTO
      MOSTRA_MKP_ATACADO
      MOSTRA_MKP_VAREJO
      Else: txtDescProd.Text = "Todos"
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_LostFocus"
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

Private Sub TXTCNPJCPF_FORNEC_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF_FORNEC.Mask = "##############"
   txtCNPJCPF_FORNEC.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_FORNEC_GotFocus"
End Sub
Private Sub TXTCNPJCPF_FORNEC_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_FORNEC
      txtCNPJCPF_FORNEC.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_FORNEC_KeyPress"
End Sub
Private Sub TXTCNPJCPF_FORNEC_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF_FORNEC.PromptInclude = False
            txtCNPJCPF_FORNEC.Text = CNPJCPF_A
            txtCNPJCPF_FORNEC.PromptInclude = True

            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

            SQL = "select * from vwFornecedor WITH (NOLOCK)"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
            TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFornecedor.EOF Then
               If Trim(TabFornecedor!NOME) = "" Then
                  txtNome.Text = Trim(TabFornecedor!RAZAO_SOCIAL)
                  Else: txtNome.Text = Trim(TabFornecedor!NOME)
               End If
               FORNEC_ID_N = TabFornecedor!FORNECEDOR_ID
            End If
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close
         End If
         CNPJCPF_A = ""
         txtCNPJCPF_FORNEC.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_FORNEC_KeyDown"
End Sub

Private Sub TXTOCORRENCIA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPerc.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTOCORRENCIA_KeyPress"
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_TUDO
      txtProduto.SetFocus
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

         SQL = "select * from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where CODG_PRODUTO = '" & txtprodalt.Text & "'"
         If cmbprod.Text <> "" Then
            SQL = SQL & " and tipo_prod = " & Left(cmbprod.Text, 1)
            Else: SQL = SQL & " and tipo_prod = " & Left(cmbFamilia.Text, 1)
         End If
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            txtdescalt.Text = Trim(TabProduto!DESCRICAO)
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

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
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

         SQL = "select * from PRODUTO WITH (NOLOCK)"
         SQL = SQL & " where CODG_PRODUTO = '" & txtprodalt.Text & "'"
         SQL = SQL & " and situacao <> 'C' "
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then
            VALOR_CUSTO_PRODUTO_N = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
            VALOR_ATACADO_PRODUTO_N = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
            VALOR_VAREJO_PRODUTO_N = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
            VALOR_DIGITADO_N = txtValor.Text
            
            If Left(cmbPreco.Text, 1) = 0 Then _
               PERC_CUSTO_N = Format(VALOR_CUSTO_PRODUTO_N * (VALOR_DIGITADO_N / 100), strFormatacao2Digitos)
            If Left(cmbPreco.Text, 1) = 1 Then _
               PERC_ATACADO_N = Format(VALOR_ATACADO_PRODUTO_N * (VALOR_DIGITADO_N / 100), strFormatacao2Digitos)
            If Left(cmbPreco.Text, 1) = 2 Then _
               PERC_VAREJO_N = Format(VALOR_VAREJO_PRODUTO_N * (VALOR_DIGITADO_N / 100), strFormatacao2Digitos)
            
            'If INDR_TIPO_PROD = 1 Then
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
               SQL = SQL & " where CODG_PRODUTO = '" & txtprodalt.Text & "'"
               CONECTA_RETAGUARDA.Execute SQL
            '   Else
            '      SQL = "update PRODUTO,MP set "
            '      If opvalor.Value = True Then
            '         If Left(cmbPreco.Text, 1) = 0 Then _
            '            SQL = SQL & "PRECO_CUSTO = " & Replace(VALOR_DIGITADO_N, ",", ".")
            '         If Left(cmbPreco.Text, 1) = 1 Then _
            '            SQL = SQL & "PRECO_ATACADO = " & Replace(VALOR_DIGITADO_N, ",", ".")
            '         If Left(cmbPreco.Text, 1) = 2 Then _
            '            SQL = SQL & "PRECO_VENDA = " & Replace(VALOR_DIGITADO_N, ",", ".")
            '      End If
            '      If opperc.Value = True Then
            '         If Left(cmbPreco.Text, 1) = 0 Then _
            '            SQL = SQL & "PRECO_CUSTO = " & Replace(PERC_CUSTO_N + VALOR_CUSTO_PRODUTO_N, ",", ".")
            '         If Left(cmbPreco.Text, 1) = 1 Then _
            '            SQL = SQL & "PRECO_ATACADO = " & Replace(PERC_ATACADO_N + VALOR_ATACADO_PRODUTO_N, ",", ".")
            '         If Left(cmbPreco.Text, 1) = 2 Then _
            '            SQL = SQL & "PRECO_VENDA = " & Replace(PERC_VAREJO_N + VALOR_VAREJO_PRODUTO_N, ",", ".")
            '      End If
            '      SQL = SQL & " where codg_mp = CODG_PRODUTO "
            '      SQL = SQL & " and codg_pa = '" & txtprodalt.Text & "'"
            '      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            '      CONECTA_RETAGUARDA.Execute SQL
            'End If
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

   If SSTab1.Tab = 0 Then
      CONT_N = 0

      Dim Coluna, Linha, Largura_Campo

      MSFlexGrid1.Clear

      MSFlexGrid1.Gridlines = flexGridFlat
      MSFlexGrid1.FixedRows = 1
      MSFlexGrid1.FixedCols = 1
      MSFlexGrid1.ScrollBars = flexScrollBarBoth
      MSFlexGrid1.AllowUserResizing = flexResizeColumns

      'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
      'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      MONTA_SQL

      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
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
               If Coluna = 5 Then
                  MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabConsulta.Fields(Coluna).Value, strFormatacao2Digitos)
                  Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabConsulta.Fields(Coluna).Value)
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

         'TAXAMARKUP.TAXAMARKUP_ID
            MSFlexGrid1.ColWidth(0) = 0

         'TAXAMARKUP.ESTABELECIMENTO_ID
            MSFlexGrid1.ColWidth(1) = 0

         'TAXAMARKUP.TIPOMERCADO_ID
            MSFlexGrid1.ColWidth(2) = 1000

         'TAXAMARKUP.PRODUTO_ID
            MSFlexGrid1.ColWidth(3) = 0

         'TAXAMARKUP.OCORRENCIA
            MSFlexGrid1.ColWidth(4) = 6000

         'TAXAMARKUP.PERC_TAXA
            MSFlexGrid1.ColWidth(5) = 1000

         'PRODUTO.CODG_PRODUTO
            MSFlexGrid1.ColWidth(6) = 1500

         'PRODUTO.FORNECEDOR_ID
            MSFlexGrid1.ColWidth(7) = 0

         'PRODUTO.DESCRICAO as DescProduto
            MSFlexGrid1.ColWidth(8) = 6000
            MSFlexGrid1.ColAlignment(8) = 0

         'PRODUTO.FAMILIAPRODUTO_ID
            MSFlexGrid1.ColWidth(9) = 0
            MSFlexGrid1.ColAlignment(9) = 0

         'PRODUTO.PRECO_CUSTO
            MSFlexGrid1.ColWidth(10) = 0
            MSFlexGrid1.ColAlignment(10) = 7

         'PRODUTO.PRECO_ATACADO
            MSFlexGrid1.ColWidth(11) = 0
            MSFlexGrid1.ColAlignment(11) = 7

         'PRODUTO.PRECO_Venda
            MSFlexGrid1.ColWidth(12) = 0

         'FAMILIAPRODUTO.CODG_FAMILIA
            MSFlexGrid1.ColWidth(13) = 0

         'FAMILIAPRODUTO.DESCRICAO AS DescFornec
            MSFlexGrid1.ColWidth(14) = 0

         'FORNECEDOR.CNPJCPF
            MSFlexGrid1.ColWidth(15) = 0

         'FORNECEDOR.NOME
            MSFlexGrid1.ColWidth(16) = 0

         'FORNECEDOR.MARKUP_VAREJO
            MSFlexGrid1.ColWidth(17) = 0

         'FORNECEDOR.MARKUP_ATACADO
            MSFlexGrid1.ColWidth(18) = 0
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

   'MOSTRA_TOTAIS

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

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where tipo_prod = " & Left(cmbprod.Text, 1)
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set item = lstProdutoPreço.ListItems.Add(, "seq." & NUMR_SEQ_N, TabProduto!Codg_Produto & "-" & Trim(TabProduto!DESCRICAO))
      item.SubItems(1) = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
      item.SubItems(2) = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      item.SubItems(3) = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
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

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where situacao <> 'C' "
   If Trim(cmbFamilia.Text) <> "" Then _
      SQL = SQL & " and FAMILIAPRODUTO_ID = " & cmbFamiliaAUX.Text

   If Trim(txtRef.Text) <> "" Then _
      SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      Set item = LISTA_PROD_GRP.ListItems.Add(, "seq." & TabProduto.Fields("produto_id").Value, TabProduto!Codg_Produto & "-" & Trim(TabProduto!DESCRICAO))
      item.SubItems(1) = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
      item.SubItems(2) = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      item.SubItems(3) = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
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

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where CODG_PRODUTO = '" & txtprodalt.Text & "'"
   If cmbprod.Text <> "" Then
      SQL = SQL & " and tipo_prod = " & Left(cmbprod.Text, 1)
      Else: SQL = SQL & " and tipo_prod = " & Left(cmbFamilia.Text, 1)
   End If
   SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set item = LISTA_ALT.ListItems.Add(, "seq." & NUMR_SEQ_N, TabProduto!Codg_Produto & " - " & Trim(TabProduto!DESCRICAO))

      VALOR_ITEM_N = 0 & txtValor.Text
      item.SubItems(2) = Format(VALOR_ITEM_N, strFormatacao2Digitos)

      If Left(cmbPreco.Text, 1) = 0 Then
         item.SubItems(1) = Format(TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
         item.SubItems(3) = Format((TabProduto!PRECO_CUSTO * VALOR_ITEM_N / 100) + TabProduto!PRECO_CUSTO, strFormatacao2Digitos)
      End If
      If Left(cmbPreco.Text, 1) = 1 Then
         item.SubItems(1) = Format(TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
         item.SubItems(3) = Format((TabProduto!PRECO_ATACADO * VALOR_ITEM_N / 100) + TabProduto!PRECO_ATACADO, strFormatacao2Digitos)
      End If
      If Left(cmbPreco.Text, 1) = 2 Then
         item.SubItems(1) = Format(TabProduto!PRECO_Venda, strFormatacao2Digitos)
         item.SubItems(3) = Format((TabProduto!PRECO_Venda * VALOR_ITEM_N / 100) + TabProduto!PRECO_Venda, strFormatacao2Digitos)
      End If
      item.SubItems(4) = TabProduto!Codg_Produto
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_alteração("
End Sub

Private Sub GRAVA_TUDO()
'On Error GoTo ERRO_TRATA

   If Trim(Selecao_Mercado) = "" Then
      MsgBox "Selecione tipo mercado."
      lstMercado.SetFocus
      Exit Sub
   End If

   If Trim(Selecao_Familia) = "" Then
      If Trim(txtProduto.Text) = "" Then
         MsgBox "Selecione tipo mercado ou produto."
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtOcorrencia.Text) = "" Then
      MsgBox "Informe a ocorrencia da taxa."
      txtOcorrencia.SetFocus
      Exit Sub
   End If

   Dim Perc_Taxa As Double
   Perc_Taxa = 0

   If Trim(txtPerc.Text) <> "" Then _
      If IsNumeric(txtPerc.Text) Then _
         Perc_Taxa = txtPerc.Text

   If Perc_Taxa <= 0 Then
      MsgBox "Informe a taxa."
      txtPerc.SetFocus
      Exit Sub
   End If

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PRODUTO.PRODUTO_ID from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " FULL OUTER JOIN FORNECEDOR WITH (NOLOCK)"
   'set errado SQL = SQL & " ON PRODUTO.CODIGO_FABRICA = FORNECEDOR.FORNECEDOR_ID"
   SQL = SQL & " Where codg_produto Is Not Null"

   If Trim(Selecao_Familia) <> "" Then _
      SQL = SQL & " and familiaproduto_id in (" & Selecao_Familia & ")"

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and codg_produto = '" & Trim(txtProduto.Text) & "'"

   If FORNEC_ID_N > 0 Then _
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID_N

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      If lstMercado.ListItems.Count > 0 Then
         NUMR_ID_N = 0

         For i = lstMercado.ListItems.Count To 1 Step -1
            If lstMercado.ListItems(i).Checked = True Then
               NUMR_ID_N = MAX_ID("TAXAMARKUP_ID", "TAXAMARKUP", "", "", "", "")

               If TabTemp.State = 1 Then _
                  TabTemp.Close
   
               SQL = "select TAXAMARKUP_ID from TAXAMARKUP WITH (NOLOCK)"
               SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
               SQL = SQL & " and tipomercado_id = " & lstMercado.ListItems(i).SubItems(1)
               SQL = SQL & " and produto_id = " & TabConsulta.Fields("produto_id").Value
               SQL = SQL & " and TAXAMARKUP_ID = " & NUMR_ID_N
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabTemp.EOF Then
                  SQL = "insert into TAXAMARKUP "
                     SQL = SQL & "(TAXAMARKUP_ID,ESTABELECIMENTO_ID,TIPOMERCADO_ID,PRODUTO_ID,OCORRENCIA,PERC_TAXA)"
                  SQL = SQL & " values("
                     SQL = SQL & NUMR_ID_N                                       'TAXAMARKUP_ID
                     SQL = SQL & "," & ESTABELECIMENTO_ID_N                      'ESTABELECIMENTO_ID
                     SQL = SQL & "," & lstMercado.ListItems(i).SubItems(1)       'TIPOMERCADO_ID
                     SQL = SQL & "," & TabConsulta.Fields("produto_id").Value    'PRODUTO_ID
                     SQL = SQL & ",'" & Trim(txtOcorrencia.Text) & "'"           'OCORRENCIA
                     SQL = SQL & "," & tpMOEDA(txtPerc.Text)                     'Perc_Taxa
                  SQL = SQL & ")"
                  Else
                     SQL = "update TAXAMARKUP set"
                        SQL = SQL & " OCORRENCIA = '" & Trim(txtOcorrencia.Text) & "'" 'OCORRENCIA
                        SQL = SQL & ",Perc_Taxa = " & tpMOEDA(txtPerc.Text)            'Perc_Taxa
                     SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
                     SQL = SQL & " and tipomercado_id = " & lstMercado.ListItems(i).SubItems(1)
                     SQL = SQL & " and produto_id = " & TabConsulta.Fields("produto_id").Value
                     SQL = SQL & " and TAXAMARKUP_ID = " & NUMR_ID_N
               End If
               If TabTemp.State = 1 Then _
                  TabTemp.Close

               CONECTA_RETAGUARDA.Execute SQL
            End If
         Next i
      End If

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SETA_GRID
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO"
End Sub

Private Sub ATUALIZA_PRECO_PRODUTOS()
'On Error GoTo ERRO_TRATA

   Dim PRECO_CUSTO_ANTERIOR As Double

   PRECO_CUSTO_ANTERIOR = 0
   CONT_N = 1

   While Trim(LISTA_ALT.ListItems(CONT_N).Text) <> ""
      If cmbPreco.Tag = "PRECO_CUSTO" Then
         SQL = "update PRODUTO set "
         SQL = SQL & " preco_custo_anterior  = '" & LISTA_ALT.ListItems.item(CONT_N).SubItems(1) & "'"
         SQL = SQL & ", preco_custo = " & Replace(LISTA_ALT.ListItems.item(CONT_N).SubItems(3), ",", ".")
         SQL = SQL & " where codg_produto = '" & LISTA_ALT.ListItems.item(CONT_N).SubItems(4) & "'"
         CONECTA_RETAGUARDA.Execute SQL
         Else
            SQL = "update PRODUTO set "
            SQL = SQL & cmbPreco.Tag & " = " & Replace(LISTA_ALT.ListItems.item(CONT_N).SubItems(3), ",", ".")
            SQL = SQL & " where codg_produto = '" & LISTA_ALT.ListItems.item(CONT_N).SubItems(4) & "'"
            CONECTA_RETAGUARDA.Execute SQL
      End If
      If Left(cmbprod.Text, 1) = 0 Then 'Alterando Valores da materia prima no arquivo de materia prima
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from MP WITH (NOLOCK)"
         SQL = SQL & " where codg_mp = '" & LISTA_ALT.ListItems.item(CONT_N).SubItems(4) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            SQL = "update MP set "
            SQL = SQL & " valor = " & Replace(LISTA_ALT.ListItems.item(CONT_N).SubItems(3), ",", ".")
            SQL = SQL & " where codg_mp = '" & LISTA_ALT.ListItems.item(CONT_N).SubItems(4) & "'"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      
         'Quando preco da materia prima for alterado,
         'sera atualizado o preco de custo de todos produtos relacionado
         'aquela materia prima!
         
         SQL = "select distinct(codg_pa) from MP WITH (NOLOCK)"
         SQL = SQL & " where codg_mp = '" & LISTA_ALT.ListItems.item(CONT_N).SubItems(4) & "'"
         SQL = SQL & " order by codg_pa"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            'Fazer Calculos do preco para gravar no arquivo de produtos
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select sum(valor*qtde) from PRODUTO p, MP m WITH (NOLOCK)"
            SQL = SQL & " where p.CODG_PRODUTO = m.codg_pa "
            SQL = SQL & " and m.codg_pa = '" & TabTemp!codg_pa & "'"
            SQL = SQL & " and p.tipo_prod = 1   "
            SQL = SQL & " and p.situacao <> 'C' "
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               If Not IsNull(TabConsulta.Fields(0).Value) Then

'                  SQL = "select Sum(MP.QTDE * MP.VALOR)  from MP ,PRODUTO"
'                  SQL = SQL & " where codg_mp = CODG_PRODUTO "
'                  SQL = SQL & " and codg_pa = '" & TABTEMP!codg_pa & "'"
'                  Set rstPreco = CONECTA_RETAGUARDA.OpenRecordset(SQL, 4)
'                  If Not rstPreco.EOF Then
'                     If Not IsNull(rstPreco.Fields(0).Value) Then
'                        VALOR_PRECO_CUSTO = Format(rstPreco.Fields(0).Value, strFormatacao2Digitos)
'                     End If
'                  End If
'                  rstPreco.close

                  'Buscar Preco Custo Anterior do produto acabado
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select produto.preco_custo from PRODUTO WITH (NOLOCK)"
                  SQL = SQL & " where CODG_PRODUTO = '" & TabTemp!codg_pa & "'"
                  SQL = SQL & " and situacao <> 'C' "
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabProduto.EOF Then _
                     PRECO_CUSTO_ANTERIOR = TabProduto!PRECO_CUSTO
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "update PRODUTO set "
                  SQL = SQL & "preco_custo_anterior = " & Replace(PRECO_CUSTO_ANTERIOR, ",", ".")
                  SQL = SQL & ", preco_custo = " & Replace(VALOR_PRECO_CUSTO, ",", ".")
                  SQL = SQL & " where codg_produto = '" & TabTemp!codg_pa & "'"
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

      SQL = "select * from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where situacao <> 'C' "

      If Trim(cmbFamilia.Text) <> "" Then _
         SQL = SQL & " and FAMILIAPRODUTO_ID = " & cmbFamiliaAUX.Text

      If Trim(txtRef.Text) <> "" Then _
         SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabProduto.EOF
         DoEvents

         SQL = "update PRODUTO set "
         SQL = SQL & " produto_id = " & TabProduto.Fields("produto_id").Value

         If chkVenda.Value = 1 Then _
            SQL = SQL & " ,preco_venda = 0"
         If chkAtacado.Value = 1 Then _
            SQL = SQL & " ,preco_atacado = 0"
         If chkCusto.Value = 1 Then _
            SQL = SQL & " ,preco_custo = 0"

         SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
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

   SQL = "select * from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where situacao <> 'C' "

   If Trim(cmbFamilia.Text) <> "" Then _
      SQL = SQL & " and FAMILIAPRODUTO_ID = " & cmbFamiliaAUX.Text

   If Trim(txtRef.Text) <> "" Then _
      SQL = SQL & " and referencia like '" & Trim(txtRef.Text) & "%" & "'"

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      DoEvents

      If optDesc.Value = True Then 'BAIXANDO O PREÇO
         If Valor_Atualiza_N > 0 Then
            SQL = "update PRODUTO set "
            SQL = SQL & " produto_id = " & TabProduto.Fields("produto_id").Value

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
            CONECTA_RETAGUARDA.Execute SQL
            Else
               If Perc_Atualiza_N > 0 Then
                  SQL = "update PRODUTO set "
                  SQL = SQL & " produto_id = " & TabProduto.Fields("produto_id").Value

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
                  CONECTA_RETAGUARDA.Execute SQL
               End If
         End If
      End If
      If optAcre.Value = True Then 'Esta aumentando
         If Valor_Atualiza_N > 0 Then
            SQL = "update PRODUTO set "
            SQL = SQL & " produto_id = " & TabProduto.Fields("produto_id").Value

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
            CONECTA_RETAGUARDA.Execute SQL
            Else
               If Perc_Atualiza_N > 0 Then
                  SQL = "update PRODUTO set "
                  SQL = SQL & " produto_id = " & TabProduto.Fields("produto_id").Value

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

   txtCNPJCPF_FORNEC.PromptInclude = False

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = "select Descricao,fornecedor_id from vwFornecedor WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF_FORNEC.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      txtNome.Text = Trim(TabCliente.Fields("descricao").Value)
      FORNEC_ID_N = TabCliente!FORNECEDOR_ID
   End If
   If TabCliente.State = 1 Then _
      TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_FORNEC"
End Sub

Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0

   If Trim(txtProduto.Text) <> "" Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select produto_id,descricao,preco_custo,preco_atacado,preco_venda from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         txtDescProd.Text = TabProduto.Fields("descricao").Value
         PRODUTO_ID_N = TabProduto.Fields("produto_id").Value
         txtCusto.Text = Format(TabProduto.Fields("preco_custo").Value, strFormatacao2Digitos)
      End If
      If TabProduto.State = 1 Then _
         TabProduto.Close
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub CRIA_TABELA()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TAXAMARKUP", "U") = False Then
      SQL = "CREATE TABLE [dbo].[TAXAMARKUP]("
      SQL = SQL & " [TAXAMARKUP_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [TIPOMERCADO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [OCORRENCIA] [nvarchar](max) NOT NULL,"
      SQL = SQL & " [PERC_TAXA] [float] NOT NULL,"

      SQL = SQL & " CONSTRAINT [PK_TAXAMARKUP_1] PRIMARY KEY CLUSTERED "
      SQL = SQL & " ([TAXAMARKUP_ID] ASC,[ESTABELECIMENTO_ID] ASC,[TIPOMERCADO_ID] ASC,   [PRODUTO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
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

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CRIA_TABELA"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtDtMovimento.Text = Date
   PRODUTO_ID_N = 0
   FORNEC_ID_N = 0
   i = 0
   CRITERIO_A = ""
   SQL = ""
   SqL2 = ""
   SQL3 = ""
   Selecao_Mercado = ""
   Selecao_Familia = ""

   CARREGA_COMBO

   'If INDR_INDUSTRIA_B = True Then
   '   INDR_TIPO_PROD = 0
   '   Else: INDR_TIPO_PROD = 1
   'End If

   MSFlexGrid1.Clear
   txtValorDig.Text = ""
   
   txtNome.Text = ""
   txtMark_Fornec_Ata.Text = ""
   txtMark_Fornec_Var.Text = ""
   txtCNPJCPF_FORNEC.PromptInclude = False
   txtCNPJCPF_FORNEC.Text = ""
   LISTA_ALT.ListItems.Clear
   lstProdutoPreço.Visible = False

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtProduto.Text = ""
   txtDescProd.Text = ""
   txtOcorrencia.Text = ""
   txtPerc.Text = "" & Format(0, strFormatacao2Digitos)
   lblEstab.Caption = ESTABELECIMENTO_ID_N
   PRODUTO_ID_N = 0
   NUMR_ID_N = 0
   SQL = ""
   SqL2 = ""
   SQL3 = ""
   CRITERIO_A = ""
   txtAtacado.Text = ""
   txtVarejo.Text = ""
   txtBKPAtacado.Text = ""
   txtBKPVarejo.Text = ""
   txtCusto.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub CARREGA_COMBO()
'On Error GoTo ERRO_TRATA

   lstFamilia.ListItems.Clear
   cmbFamilia.Clear
   cmbFamiliaAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select familiaproduto_id,descricao from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " order by DESCRICAO"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbFamilia.AddItem Trim(TabTemp!FAMILIAPRODUTO_ID) & "-" & Trim(TabTemp!DESCRICAO)
      cmbFamiliaAUX.AddItem Trim(TabTemp.Fields("familiaproduto_id").Value)

      Set item = lstFamilia.ListItems.Add(, "seq." & Trim(TabTemp.Fields("familiaproduto_id").Value), _
                                                     Trim(TabTemp.Fields("descricao").Value))
      item.SubItems(1) = "" & Trim(TabTemp.Fields("familiaproduto_id").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstMercado.ListItems.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'M'"   'tipo mercado
   SQL = SQL & " order by codigo "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      Set item = lstMercado.ListItems.Add(, "seq." & Trim(TabDESCR.Fields("codigo").Value), _
                                                     Trim(TabDESCR.Fields("descricao").Value))
      item.SubItems(1) = "" & Trim(TabDESCR.Fields("codigo").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    'OcultarControles

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

   txtProduto.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 0)
   'txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 11)
   'txtPesoItem.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 2)

   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         MATA_REGISTRO
      Case vbKeyF12
         'frmobs.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
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
         AtribuiValorCelula
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

            If txtValorDig.Visible Then
               txtValorDig.ZOrder
               txtValorDig.SetFocus
            End If
      End Select

      OK = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ExibirCelula"
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
      If LastCol > 3 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      Dim QTDE_RETIDO_ESTORNO As Double

      QTDE_RETIDO_ESTORNO = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)

      AtribuiValorCelula
      'ProximaCelula
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
      txtValorDig.Text = ""
      MSFlexGrid1.SetFocus
      
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
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
      Case 3 To 5
         texto = txtValorDig.Text

         If LastCol = 3 Then
            MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
            Else: MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao2Digitos)
         End If

         VALOR_VAREJO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
         VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, LastCol)

'&H80C0FF = LARANJA
'&H8000000F = CINZA
'&HFF& = VERMELHO
'vbBlack 0x0
'vbRed 0xFF
'vbGreen 0xFF00
'vbYellow 0xFFFF
'vbBlue 0xFF0000
'vbMagenta 0xFF00FF
'vbCyan 0xFFFF00
'vbWhite 0xFFFFFF

         If VALOR_ITEM_N < VALOR_VAREJO_N Then
            MSFlexGrid1.CellForeColor = vbRed
            MSFlexGrid1.CellFontBold = True
            MSFlexGrid1.CellBackColor = &H8000000F
            Else
               If VALOR_ITEM_N = VALOR_VAREJO_N Then
                  MSFlexGrid1.CellForeColor = vbBlack
                  MSFlexGrid1.CellFontBold = True
                  MSFlexGrid1.CellBackColor = vbCyan
                  Else
                     MSFlexGrid1.CellForeColor = vbBlue
                     MSFlexGrid1.CellFontBold = True
                     MSFlexGrid1.CellBackColor = vbWhite
               End If
         End If
      Case Else
         'texto = txtValorDig.Text
         'MSFlexGrid1.TextMatrix(LastRow, LastCol) = texto
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   txtValorDig.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub

Sub MONTA_SQL()
'On Error GoTo ERRO_TRATA

   SQL = "select TAXAMARKUP.TAXAMARKUP_ID, TAXAMARKUP.ESTABELECIMENTO_ID, "
   SQL = SQL & " TAXAMARKUP.TIPOMERCADO_ID as Mercado, TAXAMARKUP.PRODUTO_ID, "
   SQL = SQL & " TAXAMARKUP.OCORRENCIA, TAXAMARKUP.PERC_TAXA as '%',"

   SQL = SQL & " PRODUTO.CODG_PRODUTO as Código, PRODUTO.FORNECEDOR_ID, PRODUTO.DESCRICAO as DescProduto, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.PRECO_CUSTO, "
   SQL = SQL & " PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, "

   SQL = SQL & " FAMILIAPRODUTO.CODG_FAMILIA, FAMILIAPRODUTO.DESCRICAO AS DescFornec"

   SQL = SQL & " from TAXAMARKUP WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON TAXAMARKUP.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID "
   SQL = SQL & " left JOIN FORNECEDOR WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"

   SQL = SQL & " where TAXAMARKUP.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(Selecao_Mercado) <> "" Then _
      SQL = SQL & " and TAXAMARKUP.tipomercado_id in (" & Selecao_Mercado & ")"

   If Trim(Selecao_Familia) <> "" Then _
      SQL = SQL & " and PRODUTO.familiaproduto_id in (" & Selecao_Familia & ")"

   If Trim(txtProduto.Text) <> "" Then _
      SQL = SQL & " and PRODUTO.codg_produto = '" & Trim(txtProduto.Text) & "'"

   If FORNEC_ID_N > 0 Then _
      SQL = SQL & " and FORNECEDOR.fornecedor_id = " & FORNEC_ID_N

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_SQL"
End Sub
'=============================
Sub MATA_REGISTRO()
'On Error GoTo ERRO_TRATA

   If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then          'TAXAMARKUP.TAXAMARKUP_ID
      If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) Then       'TAXAMARKUP.ESTABELECIMENTO_ID
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) Then    'TAXAMARKUP.TIPOMERCADO_ID
            If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)) Then 'TAXAMARKUP.PRODUTO_ID
               SQL = "delete TAXAMARKUP "
               SQL = SQL & " where estabelecimento_id = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
               SQL = SQL & " and tipomercado_id = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
               SQL = SQL & " and produto_id = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
               SQL = SQL & " and TAXAMARKUP_ID = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
               CONECTA_RETAGUARDA.Execute SQL
               SETA_GRID
               txtProduto.SetFocus
            End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MATA_REGISTRO"
End Sub

Sub MOSTRA_MKP_ATACADO()
'On Error GoTo ERRO_TRATA

   Dim MKP_ATACADO          As Double
   Dim PRECO_CUSTO_N        As Double
   Dim MOSTRA_MKP_ATACADO_N As Double

   MKP_ATACADO = 0
   MOSTRA_MKP_ATACADO_N = 0
   PRECO_CUSTO_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select TAXAMARKUP.PERC_TAXA, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda"
   SQL = SQL & " from TAXAMARKUP WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON TAXAMARKUP.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where TAXAMARKUP.produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and TAXAMARKUP.estabelecimento_id = " & EMPRESA_ID_N
   SQL = SQL & " and TAXAMARKUP.tipomercado_id = 2 "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      MKP_ATACADO = MKP_ATACADO + TabTemp.Fields("PERC_TAXA").Value
      PRECO_CUSTO_N = 0 & TabTemp.Fields("PRECO_CUSTO").Value

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If MKP_ATACADO > 0 Then
      MOSTRA_MKP_ATACADO_N = (100 - MKP_ATACADO)
      MOSTRA_MKP_ATACADO_N = MOSTRA_MKP_ATACADO_N / 100
      txtAtacado.Text = "" & Format(MOSTRA_MKP_ATACADO_N, strFormatacao2Digitos)

      MOSTRA_MKP_ATACADO_N = PRECO_CUSTO_N / MOSTRA_MKP_ATACADO_N
      txtBKPAtacado.Text = "" & Format(MOSTRA_MKP_ATACADO_N * PRECO_CUSTO_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_MKP_ATACADO"
End Sub

Sub MOSTRA_MKP_VAREJO()
'On Error GoTo ERRO_TRATA

   Dim MKP_VAREJO          As Double
   Dim PRECO_CUSTO_N       As Double
   Dim MOSTRA_MKP_VAREJO_N As Double

   MKP_VAREJO = 0
   MOSTRA_MKP_VAREJO_N = 0
   PRECO_CUSTO_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select TAXAMARKUP.PERC_TAXA, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_Venda"
   SQL = SQL & " from TAXAMARKUP WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON TAXAMARKUP.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where TAXAMARKUP.produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and TAXAMARKUP.estabelecimento_id = " & EMPRESA_ID_N
   SQL = SQL & " and TAXAMARKUP.tipomercado_id = 1 "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      MKP_VAREJO = MKP_VAREJO + TabTemp.Fields("PERC_TAXA").Value
      PRECO_CUSTO_N = 0 & TabTemp.Fields("PRECO_CUSTO").Value

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If MKP_VAREJO > 0 Then
      MOSTRA_MKP_VAREJO_N = (100 - MKP_VAREJO)
      MOSTRA_MKP_VAREJO_N = MOSTRA_MKP_VAREJO_N / 100
      txtVarejo.Text = "" & Format(MOSTRA_MKP_VAREJO_N, strFormatacao2Digitos)

      MOSTRA_MKP_VAREJO_N = PRECO_CUSTO_N / MOSTRA_MKP_VAREJO_N
      txtBKPVarejo.Text = "" & Format(MOSTRA_MKP_VAREJO_N * PRECO_CUSTO_N, strFormatacao2Digitos)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_MKP_VAREJO"
End Sub
